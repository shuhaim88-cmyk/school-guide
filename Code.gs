/**
 * دليل المدارس — الخادم (Google Apps Script)
 *
 * يحفظ كل البيانات في Google Drive الخاص بالحساب الذي ينشر التطبيق:
 *   مجلد «دليل المدارس»
 *     ├─ جدول «دليل المدارس - البيانات» (المدارس، الأقسام، المعلومات، المستخدمون)
 *     └─ مجلد «المرفقات» ← مجلد لكل مدرسة يحوي الصور والمستندات
 *
 * يُنشأ المجلد والجدول تلقائياً عند أول تشغيل.
 */

const APP_TITLE = 'دليل المدارس';
const MAX_FILE_MB = 10;            // أقصى حجم للمرفق الواحد
const MAX_TEXT = 45000;            // أقصى طول للنص (حد خلية Google Sheets 50 ألف حرف)

const SH = { schools: 'المدارس', sections: 'الأقسام', info: 'المعلومات', users: 'المستخدمون', notes: 'الملاحظات' };
const HEAD = {
  schools: ['الولاية', 'رمز المدرسة', 'اسم المدرسة', 'المرحلة', 'الجنس', 'الفترة'],
  sections: ['المعرف', 'القسم الأب', 'اسم القسم', 'الأيقونة', 'اللون', 'الترتيب'],
  info: ['رمز المدرسة', 'اسم المدرسة', 'معرف القسم', 'مسار القسم', 'معرف المعلومة', 'النوع', 'العنوان', 'النص',
    'اسم الملف', 'معرف الملف', 'رابط الملف', 'الحجم', 'نوع الملف', 'أضافها', 'التاريخ'],
  users: ['المعرف', 'الاسم', 'الدور', 'بصمة الرقم السري', 'الأقسام المستثناة', 'تاريخ الإضافة'],
  notes: ['المعرف', 'رمز المدرسة', 'اسم المدرسة', 'معرف القسم', 'مسار القسم', 'معرف المعلومة', 'عنوان المعلومة',
    'نص الملاحظة', 'كاتبها', 'معرف الكاتب', 'دور الكاتب', 'التاريخ', 'قرأها']
};
// أعمدة جدول الملاحظات (تبدأ من 0)
const N = { id: 0, code: 1, sname: 2, leaf: 3, path: 4, item: 5, ititle: 6, text: 7, by: 8, byId: 9, role: 10, at: 11, readBy: 12 };
const TEXT_TYPES = { text: true, mindmap: true };
// أعمدة جدول المعلومات (تبدأ من 0)
const I = { code: 0, sname: 1, leaf: 2, path: 3, id: 4, type: 5, title: 6, text: 7, fname: 8, fid: 9, furl: 10, size: 11, mime: 12, by: 13, at: 14 };

/* ===================== صفحة التطبيق ===================== */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/* ===================== التخزين ===================== */
const props_ = () => PropertiesService.getScriptProperties();
let SS_ = null;

function rootFolder_() {
  const p = props_(); const id = p.getProperty('folderId');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const f = DriveApp.createFolder(APP_TITLE);
  p.setProperty('folderId', f.getId());
  return f;
}
function childFolder_(parent, name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}
function ss_() {
  if (SS_) return SS_;
  const p = props_(); const id = p.getProperty('ssId');
  if (id) { try { SS_ = SpreadsheetApp.openById(id); return SS_; } catch (e) {} }
  const ss = SpreadsheetApp.create(APP_TITLE + ' - البيانات');
  DriveApp.getFileById(ss.getId()).moveTo(rootFolder_());
  ['schools', 'sections', 'info', 'users'].forEach((k, i) => {
    const sh = i === 0 ? ss.getSheets()[0].setName(SH[k]) : ss.insertSheet(SH[k]);
    sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold').setBackground('#D5EDE8');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  });
  ss.getSheetByName(SH.schools).getRange('B:B').setNumberFormat('@');
  ss.getSheetByName(SH.info).getRange('A:A').setNumberFormat('@');
  writeSections_(ss, [
    { id: 'buildings', name: 'الأبنية', icon: 'building', tone: 'bld', children: [
      { id: 'hasr', name: 'استمارة الحصر', icon: 'clipboard', children: [] },
      { id: 'rawafid', name: 'مرجع الروافد', icon: 'map', children: [] },
      { id: 'shuab', name: 'تشكيلة الشعب', icon: 'table', children: [] }] },
    { id: 'hr', name: 'المورد البشري', icon: 'people', tone: 'hr', children: [] }
  ]);
  p.setProperty('ssId', ss.getId());
  SS_ = ss;
  return ss;
}
function sheet_(k) {
  const ss = ss_();
  let sh = ss.getSheetByName(SH[k]);
  if (!sh) {
    sh = ss.insertSheet(SH[k]);
    sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold').setBackground('#D5EDE8');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
    if (k === 'notes') sh.getRange('B:B').setNumberFormat('@');
  }
  return sh;
}
function rows_(k, display) {
  const sh = sheet_(k); const n = sh.getLastRow();
  if (n < 2) return [];
  const r = sh.getRange(2, 1, n - 1, HEAD[k].length);
  return display ? r.getDisplayValues() : r.getValues();
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function bump_() { const v = String(Date.now()); props_().setProperty('ver', v); return v; }
const ver_ = () => props_().getProperty('ver') || '0';
const newId_ = () => Utilities.getUuid().replace(/-/g, '').slice(0, 14);

/* ---------- الأقسام ---------- */
function readSections_() {
  const rows = rows_('sections', true).filter(r => r[0]);
  const map = {};
  rows.forEach(r => { map[r[0]] = { id: String(r[0]), parent: String(r[1] || ''), name: String(r[2]), icon: String(r[3] || 'folder'), tone: String(r[4] || ''), order: Number(r[5]) || 0, children: [] }; });
  const roots = [];
  Object.values(map).forEach(n => { (n.parent && map[n.parent] ? map[n.parent].children : roots).push(n); });
  const sort = list => { list.sort((a, b) => a.order - b.order); list.forEach(n => { sort(n.children); delete n.order; delete n.parent; if (!n.tone) delete n.tone; }); };
  sort(roots);
  const tones = ['bld', 'hr', 't3', 't4', 't5', 't6'];
  roots.forEach((r, i) => { if (!r.tone) r.tone = tones[i % tones.length]; });
  return roots;
}
function writeSections_(ss, roots) {
  const out = [];
  const walk = (list, parent, isRoot) => list.forEach((n, i) => {
    out.push([String(n.id), parent, String(n.name || '').slice(0, 80), String(n.icon || 'folder'), isRoot ? String(n.tone || '') : '', i]);
    walk(n.children || [], String(n.id), false);
  });
  walk(roots, '', true);
  const sh = ss.getSheetByName(SH.sections);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.sections.length).clearContent();
  if (out.length) sh.getRange(2, 1, out.length, HEAD.sections.length).setValues(out);
}
function pathTo_(roots, leafId) {
  let found = null;
  const walk = (list, trail) => list.some(n => { const t = trail.concat(n); if (n.id === leafId) { found = t; return true; } return walk(n.children || [], t); });
  walk(roots, []);
  return found;
}

/* ---------- المستخدمون والجلسات ---------- */
function readUsers_() {
  return rows_('users').filter(r => r[0]).map(r => ({
    id: String(r[0]), name: String(r[1]), role: String(r[2]), h: String(r[3]),
    hidden: String(r[4] || '').split(',').map(s => s.trim()).filter(String), at: r[5] instanceof Date ? r[5].getTime() : 0
  }));
}
function writeUsers_(users) {
  const sh = sheet_('users');
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.users.length).clearContent();
  if (users.length) sh.getRange(2, 1, users.length, HEAD.users.length).setValues(users.map(u => [u.id, u.name, u.role, u.h, (u.hidden || []).join(','), new Date(u.at || Date.now())]));
}
const publicUser_ = u => ({ id: u.id, name: u.name, role: u.role, hidden: u.hidden || [] });
function hash_(pin) {
  const p = props_(); let salt = p.getProperty('pinSalt');
  if (!salt) { salt = Utilities.getUuid(); p.setProperty('pinSalt', salt); }
  const d = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + pin, Utilities.Charset.UTF_8);
  return d.map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
function checkPin_(pin) { if (!/^\d{4,8}$/.test(String(pin))) throw new Error('BAD_PIN'); }
function newSession_(u) {
  const token = Utilities.getUuid() + Utilities.getUuid();
  props_().setProperty('t:' + token, JSON.stringify({ uid: u.id, h: u.h }));
  return token;
}
function auth_(token, need) {
  const raw = token && props_().getProperty('t:' + token);
  if (!raw) throw new Error('AUTH');
  const s = JSON.parse(raw);
  const u = readUsers_().find(x => x.id === s.uid);
  if (!u || u.h !== s.h) { props_().deleteProperty('t:' + token); throw new Error('AUTH'); }
  if (need === 'edit' && u.role === 'viewer') throw new Error('FORBIDDEN');
  if (need === 'admin' && u.role !== 'admin') throw new Error('FORBIDDEN');
  return u;
}

/* ---------- ما يُسمح للمستخدم برؤيته ---------- */
function visibleRoots_(u, roots) {
  if (u.role !== 'viewer') return roots;
  const H = {}; (u.hidden || []).forEach(id => { H[id] = true; });
  const filter = list => list.map(n => {
    if (H[n.id]) return null;
    if (!n.children.length) return n;
    const ch = filter(n.children);
    return ch.length ? Object.assign({}, n, { children: ch }) : null;
  }).filter(Boolean);
  return filter(roots);
}
function allowedLeaves_(u, roots) {
  const ids = {};
  const walk = list => list.forEach(n => { ids[n.id] = true; walk(n.children || []); });
  walk(visibleRoots_(u, roots));
  return ids;
}

/* ---------- المعلومات ---------- */
function rowToItem_(r) {
  const it = { id: String(r[I.id]), type: String(r[I.type]), title: String(r[I.title] || ''), by: String(r[I.by] || ''),
    at: r[I.at] instanceof Date ? r[I.at].getTime() : Number(r[I.at]) || 0 };
  if (TEXT_TYPES[it.type]) it.text = String(r[I.text] || '');
  else { it.name = String(r[I.fname]); it.fileId = String(r[I.fid]); it.url = String(r[I.furl]); it.size = Number(r[I.size]) || 0; it.mime = String(r[I.mime]); }
  return it;
}
function infoFor_(code, allowed) {
  const out = {};
  rows_('info').forEach(r => {
    if (String(r[I.code]) !== String(code) || !r[I.id]) return;
    const leaf = String(r[I.leaf]);
    if (allowed && !allowed[leaf]) return;
    (out[leaf] = out[leaf] || { items: [] }).items.push(rowToItem_(r));
  });
  Object.keys(out).forEach(k => out[k].items.sort((a, b) => b.at - a.at));
  return out;
}
function findInfoRow_(itemId) {
  const sh = sheet_('info');
  if (sh.getLastRow() < 2) return null;
  const hit = sh.getRange(2, I.id + 1, sh.getLastRow() - 1, 1).createTextFinder(String(itemId)).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : null;
}
function schoolName_(code) {
  const r = rows_('schools', true).find(x => String(x[1]) === String(code));
  return r ? String(r[2]) : '';
}
function newInfoRow_(u, code, leaf, item) {
  const path = (pathTo_(readSections_(), leaf) || []).map(n => n.name).join(' › ');
  const r = new Array(HEAD.info.length).fill('');
  r[I.code] = String(code); r[I.sname] = schoolName_(code); r[I.leaf] = leaf; r[I.path] = path;
  r[I.id] = item.id; r[I.type] = item.type; r[I.title] = item.title || ''; r[I.text] = item.text || '';
  r[I.fname] = item.name || ''; r[I.fid] = item.fileId || ''; r[I.furl] = item.url || ''; r[I.size] = item.size || '';
  r[I.mime] = item.mime || ''; r[I.by] = u.name; r[I.at] = new Date();
  return r;
}
function leafResult_(u, code, leaf) {
  const info = infoFor_(code, null);
  return { items: (info[leaf] && info[leaf].items) || [], ver: bump_() };
}

/* ===================== واجهة التطبيق (تُستدعى من الصفحة) ===================== */
function apiStatus() {
  return { needsSetup: readUsers_().length === 0 };
}
function apiSetup(name, pin) {
  return withLock_(() => {
    if (readUsers_().length) throw new Error('SETUP_DONE');
    checkPin_(pin);
    const u = { id: newId_(), name: String(name || 'المدير').slice(0, 60), role: 'admin', h: hash_(pin), hidden: [], at: Date.now() };
    writeUsers_([u]); bump_();
    return { token: newSession_(u) };
  });
}
function apiLogin(pin) {
  checkPin_(pin);
  const h = hash_(pin);
  const u = readUsers_().find(x => x.h === h);
  if (!u) { Utilities.sleep(700); throw new Error('WRONG_PIN'); }
  return { token: newSession_(u) };
}
function apiLogout(token) { if (token) props_().deleteProperty('t:' + token); return true; }

function apiBootstrap(token) {
  const u = auth_(token);
  const roots = readSections_();
  const schools = rows_('schools', true).filter(r => r[1] && r[2]).map(r => ({
    wilayah: String(r[0]).trim(), code: String(r[1]).trim(), name: String(r[2]).trim(), stage: String(r[3]).trim(), gender: String(r[4]).trim(), period: String(r[5]).trim()
  }));
  return {
    me: publicUser_(u),
    config: { roots: visibleRoots_(u, roots) },
    schools,
    users: u.role === 'admin' ? readUsers_().map(publicUser_) : null,
    ver: ver_(),
    unreadNotes: unreadCount_(u),
    driveUrl: u.role === 'admin' ? rootFolder_().getUrl() : ''
  };
}
function apiPoll(token) { auth_(token); return ver_(); }

/* ---------- الإحصائيات ---------- */
// يعيد [رمز المدرسة، معرف القسم، النوع] لكل معلومة يُسمح للمستخدم برؤيتها
function apiStats(token) {
  const u = auth_(token);
  const allowed = u.role === 'viewer' ? allowedLeaves_(u, readSections_()) : null;
  return rows_('info').filter(r => r[I.id] && (!allowed || allowed[String(r[I.leaf])]))
    .map(r => [String(r[I.code]), String(r[I.leaf]), String(r[I.type])]);
}

/* ---------- الملاحظات ---------- */
function notesRows_() { return sheet_('notes').getLastRow() < 2 ? [] : rows_('notes'); }
function unreadCount_(u) {
  if (u.role === 'viewer') return 0;
  return notesRows_().filter(r => r[N.id] && String(r[N.byId]) !== u.id && String(r[N.readBy] || '').split(',').indexOf(u.id) < 0).length;
}
function apiAddNote(token, code, leaf, itemId, text) {
  const u = auth_(token);
  text = String(text || '').trim();
  if (!text) throw new Error('NOT_FOUND');
  if (text.length > 5000) throw new Error('TEXT_TOO_LONG');
  if (u.role === 'viewer' && !allowedLeaves_(u, readSections_())[String(leaf)]) throw new Error('FORBIDDEN');
  return withLock_(() => {
    const row = findInfoRow_(itemId); if (!row) throw new Error('NOT_FOUND');
    const ir = sheet_('info').getRange(row, 1, 1, HEAD.info.length).getValues()[0];
    const label = String(ir[I.title] || ir[I.fname] || '') || ({ text: 'نص', image: 'صورة', file: 'مستند', mindmap: 'خريطة ذهنية' }[String(ir[I.type])] || 'معلومة');
    const out = new Array(HEAD.notes.length).fill('');
    out[N.id] = newId_(); out[N.code] = String(code); out[N.sname] = String(ir[I.sname] || schoolName_(code)); out[N.leaf] = String(leaf);
    out[N.path] = String(ir[I.path] || ''); out[N.item] = String(itemId); out[N.ititle] = label; out[N.text] = text;
    out[N.by] = u.name; out[N.byId] = u.id; out[N.role] = u.role; out[N.at] = new Date(); out[N.readBy] = u.role === 'viewer' ? '' : u.id;
    sheet_('notes').appendRow(out);
    return { ver: bump_() };
  });
}
function apiNotes(token) {
  const u = auth_(token, 'edit');
  return notesRows_().filter(r => r[N.id]).map(r => ({
    id: String(r[N.id]), code: String(r[N.code]), school: String(r[N.sname]), leaf: String(r[N.leaf]), path: String(r[N.path]),
    item: String(r[N.item]), itemTitle: String(r[N.ititle]), text: String(r[N.text]), by: String(r[N.by]), byId: String(r[N.byId]),
    role: String(r[N.role]), at: r[N.at] instanceof Date ? r[N.at].getTime() : Number(r[N.at]) || 0,
    read: String(r[N.byId]) === u.id || String(r[N.readBy] || '').split(',').indexOf(u.id) >= 0
  })).sort((a, b) => b.at - a.at).slice(0, 1000);
}
function apiMarkNotes(token, ids, read) {
  const u = auth_(token, 'edit');
  const want = {}; (ids || []).forEach(id => { want[String(id)] = true; });
  return withLock_(() => {
    const sh = sheet_('notes'); const n = sh.getLastRow();
    if (n < 2) return { unread: 0 };
    const rng = sh.getRange(2, N.readBy + 1, n - 1, 1), ids2 = sh.getRange(2, 1, n - 1, 1).getValues(), vals = rng.getValues();
    let changed = false;
    ids2.forEach((r, i) => {
      if (!want[String(r[0])]) return;
      const list = String(vals[i][0] || '').split(',').filter(String), has = list.indexOf(u.id) >= 0;
      if (read && !has) { list.push(u.id); changed = true; }
      if (!read && has) { list.splice(list.indexOf(u.id), 1); changed = true; }
      vals[i][0] = list.join(',');
    });
    if (changed) rng.setValues(vals);
    return { unread: unreadCount_(u) };
  });
}
function apiDeleteNote(token, id) {
  const u = auth_(token, 'edit');
  return withLock_(() => {
    const sh = sheet_('notes'); if (sh.getLastRow() < 2) throw new Error('NOT_FOUND');
    const hit = sh.getRange(2, 1, sh.getLastRow() - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
    if (!hit) throw new Error('NOT_FOUND');
    sh.deleteRow(hit.getRow());
    return { unread: unreadCount_(u), ver: bump_() };
  });
}

function apiGetInfo(token, code) {
  const u = auth_(token);
  return infoFor_(code, u.role === 'viewer' ? allowedLeaves_(u, readSections_()) : null);
}
function apiAddText(token, code, leaf, title, text, kind) {
  const u = auth_(token, 'edit');
  if (String(text).length > MAX_TEXT) throw new Error('TEXT_TOO_LONG');
  const type = kind === 'mindmap' ? 'mindmap' : 'text';
  return withLock_(() => {
    sheet_('info').appendRow(newInfoRow_(u, code, leaf, { id: newId_(), type, title: String(title || '').slice(0, 200), text: String(text) }));
    return leafResult_(u, code, leaf);
  });
}
function apiEditText(token, code, leaf, itemId, title, text) {
  const u = auth_(token, 'edit');
  if (String(text).length > MAX_TEXT) throw new Error('TEXT_TOO_LONG');
  return withLock_(() => {
    const row = findInfoRow_(itemId); if (!row) throw new Error('NOT_FOUND');
    const sh = sheet_('info');
    sh.getRange(row, I.title + 1, 1, 2).setValues([[String(title || '').slice(0, 200), String(text)]]);
    sh.getRange(row, I.by + 1, 1, 2).setValues([[u.name, new Date()]]);
    return leafResult_(u, code, leaf);
  });
}
function apiDeleteItem(token, code, leaf, itemId) {
  const u = auth_(token, 'edit');
  return withLock_(() => {
    const row = findInfoRow_(itemId); if (!row) throw new Error('NOT_FOUND');
    const sh = sheet_('info');
    const fid = sh.getRange(row, I.fid + 1).getValue();
    if (fid) { try { DriveApp.getFileById(String(fid)).setTrashed(true); } catch (e) {} }
    sh.deleteRow(row);
    return leafResult_(u, code, leaf);
  });
}
function apiUpload(token, code, leaf, file) {
  const u = auth_(token, 'edit');
  const bytes = Utilities.base64Decode(file.data);
  if (bytes.length > MAX_FILE_MB * 1048576) throw new Error('TOO_LARGE');
  const sname = schoolName_(code);
  const folder = childFolder_(childFolder_(rootFolder_(), 'المرفقات'), (String(code) + ' - ' + sname).trim());
  const f = folder.createFile(Utilities.newBlob(bytes, file.mime || 'application/octet-stream', file.name || 'ملف'));
  const item = { id: newId_(), type: /^image\//.test(file.mime) ? 'image' : 'file', title: '', name: f.getName(), fileId: f.getId(), url: f.getUrl(), size: bytes.length, mime: file.mime };
  return withLock_(() => {
    sheet_('info').appendRow(newInfoRow_(u, code, leaf, item));
    return leafResult_(u, code, leaf);
  });
}
function apiGetFile(token, fileId) {
  const u = auth_(token);
  const row = rows_('info').find(r => String(r[I.fid]) === String(fileId));
  if (!row) throw new Error('NOT_FOUND');
  if (u.role === 'viewer' && !allowedLeaves_(u, readSections_())[String(row[I.leaf])]) throw new Error('FORBIDDEN');
  const blob = DriveApp.getFileById(String(fileId)).getBlob();
  return { data: Utilities.base64Encode(blob.getBytes()), mime: blob.getContentType(), name: String(row[I.fname]) };
}

function apiSaveSchools(token, list) {
  auth_(token, 'edit');
  return withLock_(() => {
    const sh = sheet_('schools');
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.schools.length).clearContent();
    const rows = (list || []).filter(s => s.code && s.name).map(s => [s.wilayah || '', String(s.code), s.name, s.stage || '', s.gender || '', s.period || '']);
    sh.getRange('B:B').setNumberFormat('@');
    if (rows.length) sh.getRange(2, 1, rows.length, HEAD.schools.length).setValues(rows);
    return { ver: bump_() };
  });
}
function apiSaveConfig(token, roots) {
  auth_(token, 'admin');
  return withLock_(() => { writeSections_(ss_(), roots || []); return { ver: bump_() }; });
}

function apiAddUser(token, name, pin, role) {
  auth_(token, 'admin');
  checkPin_(pin);
  return withLock_(() => {
    const users = readUsers_(); const h = hash_(pin);
    if (users.some(x => x.h === h)) throw new Error('PIN_TAKEN');
    users.push({ id: newId_(), name: String(name).slice(0, 60), role: role === 'editor' ? 'editor' : 'viewer', h, hidden: [], at: Date.now() });
    writeUsers_(users);
    return { users: users.map(publicUser_), ver: bump_() };
  });
}
function apiUpdateUser(token, id, patch) {
  const me = auth_(token, 'admin');
  return withLock_(() => {
    const users = readUsers_(); const u = users.find(x => x.id === id);
    if (!u) throw new Error('NOT_FOUND');
    if (patch.role && u.id !== me.id && u.role !== 'admin') u.role = patch.role === 'editor' ? 'editor' : 'viewer';
    if (patch.hidden) u.hidden = patch.hidden.map(String);
    if (patch.pin) {
      checkPin_(patch.pin); const h = hash_(patch.pin);
      if (users.some(x => x.h === h && x.id !== id)) throw new Error('PIN_TAKEN');
      u.h = h;
    }
    writeUsers_(users);
    const out = { users: users.map(publicUser_), ver: bump_() };
    if (patch.pin && u.id === me.id) { apiLogout(token); out.token = newSession_(u); }
    return out;
  });
}
function apiDeleteUser(token, id) {
  const me = auth_(token, 'admin');
  return withLock_(() => {
    let users = readUsers_();
    const u = users.find(x => x.id === id);
    if (!u || u.id === me.id) throw new Error('FORBIDDEN');
    if (u.role === 'admin' && users.filter(x => x.role === 'admin').length < 2) throw new Error('FORBIDDEN');
    users = users.filter(x => x.id !== id);
    writeUsers_(users);
    return { users: users.map(publicUser_), ver: bump_() };
  });
}
function apiChangeMyPin(token, pin) {
  const me = auth_(token);
  checkPin_(pin);
  return withLock_(() => {
    const users = readUsers_(); const h = hash_(pin);
    if (users.some(x => x.h === h && x.id !== me.id)) throw new Error('PIN_TAKEN');
    const u = users.find(x => x.id === me.id); u.h = h;
    writeUsers_(users);
    apiLogout(token);
    return { token: newSession_(u), ver: bump_() };
  });
}

/** شغّل هذه الدالة مرة واحدة من المحرر لإنشاء المجلد والجدول ومنح الصلاحيات. */
function setup() {
  const ss = ss_();
  Logger.log('المجلد: ' + rootFolder_().getUrl());
  Logger.log('جدول البيانات: ' + ss.getUrl());
}

/**
 * إعادة تعيين الرقم السري للمدير عند نسيانه.
 * 1) اكتب الرقم الجديد (4 إلى 8 أرقام) بين علامتي التنصيص في السطر التالي.
 * 2) اختر الدالة resetAdminPin من القائمة أعلى المحرر واضغط «تشغيل».
 * 3) أعد السطر إلى '' بعد الاستخدام حتى لا يبقى الرقم مكتوباً في الكود.
 * لا يستطيع تشغيل هذه الدالة إلا صاحب حساب Google الذي يملك المشروع.
 */
function resetAdminPin() {
  const NEW_PIN = '';
  if (!/^\d{4,8}$/.test(NEW_PIN)) throw new Error('اكتب رقماً سرياً من 4 إلى 8 أرقام في NEW_PIN ثم شغّل الدالة مرة أخرى.');
  withLock_(() => {
    const users = readUsers_();
    const h = hash_(NEW_PIN);
    if (users.some(u => u.h === h && u.role !== 'admin')) throw new Error('هذا الرقم مستخدم لحساب آخر. اختر رقماً مختلفاً.');
    let admin = users.find(u => u.role === 'admin');
    if (!admin) { admin = { id: newId_(), name: 'المدير', role: 'admin', hidden: [], at: Date.now() }; users.push(admin); }
    admin.h = h;
    writeUsers_(users);
    bump_();
    Logger.log('تم تعيين الرقم السري الجديد للمدير «' + admin.name + '». ادخل به إلى التطبيق، ثم أعد NEW_PIN إلى \'\'.');
  });
}
