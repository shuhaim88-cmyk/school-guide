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

const SH = { schools: 'المدارس', sections: 'الأقسام', info: 'المعلومات', users: 'المستخدمون', notes: 'الملاحظات', accounts: 'حسابات المدارس' };
const HEAD = {
  schools: ['الولاية', 'رمز المدرسة', 'اسم المدرسة', 'المرحلة', 'الجنس', 'الفترة'],
  sections: ['المعرف', 'القسم الأب', 'اسم القسم', 'الأيقونة', 'اللون', 'الترتيب'],
  info: ['رمز المدرسة', 'اسم المدرسة', 'معرف القسم', 'مسار القسم', 'معرف المعلومة', 'النوع', 'العنوان', 'النص',
    'اسم الملف', 'معرف الملف', 'رابط الملف', 'الحجم', 'نوع الملف', 'أضافها', 'التاريخ', 'مرئية للمدارس', 'المدارس المستهدفة'],
  users: ['المعرف', 'الاسم', 'الدور', 'بصمة الرقم السري', 'الأقسام المستثناة', 'تاريخ الإضافة'],
  notes: ['المعرف', 'رمز المدرسة', 'اسم المدرسة', 'معرف القسم', 'مسار القسم', 'معرف المعلومة', 'عنوان المعلومة',
    'نص الملاحظة', 'كاتبها', 'معرف الكاتب', 'دور الكاتب', 'التاريخ', 'قرأها'],
  accounts: ['رمز المدرسة', 'اسم المدرسة', 'بصمة كلمة المرور', 'حالة كلمة المرور', 'حالة الحساب', 'آخر دخول', 'تاريخ تغيير كلمة المرور', 'أقسام مضافة لها', 'أقسام مستبعدة عنها']
};
// أعمدة جدول حسابات المدارس (تبدأ من 0)
const A = { code: 0, sname: 1, h: 2, pw: 3, status: 4, last: 5, changed: 6, plus: 7, minus: 8 };
const PW = { def: 'افتراضية', changed: 'تم تغييرها', reset: 'أُعيد تعيينها' };
const ST = { active: 'مفعّل', disabled: 'موقوف' };
// من يرى المعلومة من حسابات المدارس
const AUD = { all: 'جميع المدارس', list: 'مدارس محددة' };
// أعمدة جدول الملاحظات (تبدأ من 0)
const N = { id: 0, code: 1, sname: 2, leaf: 3, path: 4, item: 5, ititle: 6, text: 7, by: 8, byId: 9, role: 10, at: 11, readBy: 12 };
const TEXT_TYPES = { text: true, mindmap: true, table: true };
// أعمدة جدول المعلومات (تبدأ من 0)
const I = { code: 0, sname: 1, leaf: 2, path: 3, id: 4, type: 5, title: 6, text: 7, fname: 8, fid: 9, furl: 10, size: 11, mime: 12, by: 13, at: 14, aud: 15, tgt: 16 };

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
  while (it.hasNext()) { const f = it.next(); if (!f.isTrashed()) return f; }
  return parent.createFolder(name);
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
    if (k === 'notes' || k === 'accounts') sh.getRange(k === 'notes' ? 'B:B' : 'A:A').setNumberFormat('@');
  }
  if (k === 'accounts' && !ACC_HEAD_OK_) {
    const h = sh.getRange(1, A.plus + 1, 1, 2);
    if (String(h.getValues()[0][0]) !== HEAD.accounts[A.plus]) h.setValues([[HEAD.accounts[A.plus], HEAD.accounts[A.minus]]]).setFontWeight('bold').setBackground('#D5EDE8');
    ACC_HEAD_OK_ = true;
  }
  if (k === 'info' && !INFO_HEAD_OK_) {
    const h = sh.getRange(1, I.aud + 1, 1, 2);
    if (String(h.getValues()[0][0]) !== HEAD.info[I.aud]) { h.setValues([[HEAD.info[I.aud], HEAD.info[I.tgt]]]).setFontWeight('bold').setBackground('#D5EDE8'); sh.getRange(2, I.tgt + 1, sh.getMaxRows() - 1, 1).setNumberFormat('@'); }
    INFO_HEAD_OK_ = true;
  }
  return sh;
}
let INFO_HEAD_OK_ = false, ACC_HEAD_OK_ = false;
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
function auth_(token, need, allowMustChange) {
  const raw = token && props_().getProperty('t:' + token);
  if (!raw) throw new Error('AUTH');
  const s = JSON.parse(raw);
  if (s.school) {
    const su = schoolUser_(s.school);
    if (!su || su.disabled || su.h !== s.h) { props_().deleteProperty('t:' + token); throw new Error('AUTH'); }
    if (su.mustChange && !allowMustChange) throw new Error('MUST_CHANGE');
    if (need) throw new Error('FORBIDDEN');   // حساب المدرسة لا يعدّل شيئاً
    return su;
  }
  const u = readUsers_().find(x => x.id === s.uid);
  if (!u || u.h !== s.h) { props_().deleteProperty('t:' + token); throw new Error('AUTH'); }
  if (need === 'edit' && u.role !== 'admin' && u.role !== 'editor') throw new Error('FORBIDDEN');
  if (need === 'admin' && u.role !== 'admin') throw new Error('FORBIDDEN');
  return u;
}

/* ---------- حسابات المدارس ----------
 * لكل مدرسة في ورقة «المدارس» حساب تلقائي: اسم المستخدم = رمز المدرسة.
 * ما دامت كلمة المرور «افتراضية» أو «أُعيد تعيينها» فكلمة المرور المؤقتة هي رمز المدرسة، ويُطلب تغييرها عند الدخول.
 * لا تُحفظ كلمة المرور نفسها، بل بصمتها فقط. */
function schoolRows_() {
  return rows_('schools', true).filter(r => String(r[1]).trim() && String(r[2]).trim())
    .map(r => ({ wilayah: String(r[0]).trim(), code: String(r[1]).trim(), name: String(r[2]).trim(), stage: String(r[3]).trim(), gender: String(r[4]).trim(), period: String(r[5]).trim() }));
}
function readAccounts_() {
  const out = {};
  if (sheet_('accounts').getLastRow() < 2) return out;
  rows_('accounts').forEach((r, i) => { const c = String(r[A.code]).trim(); if (c) out[c] = { row: i + 2, h: String(r[A.h] || ''), pw: String(r[A.pw] || PW.def), status: String(r[A.status] || ST.active), last: r[A.last] instanceof Date ? r[A.last].getTime() : 0, plus: idList_(r[A.plus]), minus: idList_(r[A.minus]) }; });
  return out;
}
function saveAccount_(code, patch) {
  const sh = sheet_('accounts'), acc = readAccounts_()[code];
  const cur = acc ? sh.getRange(acc.row, 1, 1, HEAD.accounts.length).getValues()[0] : [code, schoolName_(code), '', PW.def, ST.active, '', '', '', ''];
  Object.keys(patch).forEach(k => { cur[A[k]] = patch[k]; });
  cur[A.code] = String(code); cur[A.sname] = schoolName_(code) || cur[A.sname];
  if (acc) sh.getRange(acc.row, 1, 1, HEAD.accounts.length).setValues([cur]); else sh.appendRow(cur);
}
// مفتاح الجلسة يتغير مع كل تغيير أو إعادة تعيين لكلمة المرور، فتنتهي الجلسات القديمة تلقائياً
function schoolUser_(code) {
  code = String(code || '').trim();
  const s = schoolRows_().find(x => x.code === code); if (!s) return null;
  const acc = readAccounts_()[code];
  const changed = acc && acc.pw === PW.changed;
  return { id: 'school:' + code, name: s.name, role: 'school', school: code, hidden: [],
    h: acc ? (acc.h || 'default') : 'default', mustChange: !changed, disabled: !!acc && acc.status === ST.disabled };
}
const schoolPwHash_ = (code, pw) => hash_('school:' + code + ':' + pw);
function apiSchoolLogin(code, password) {
  code = String(code || '').trim(); password = String(password || '');
  const cache = CacheService.getScriptCache(), key = 'sfail:' + code;
  if (Number(cache.get(key) || 0) >= 8) throw new Error('LOCKED');
  const su = code && schoolUser_(code);
  const acc = su && readAccounts_()[code];
  const ok = su && (su.mustChange ? password === code : schoolPwHash_(code, password) === acc.h);
  if (!ok) { cache.put(key, String(Number(cache.get(key) || 0) + 1), 600); Utilities.sleep(700); throw new Error('WRONG_PASSWORD'); }
  if (su.disabled) throw new Error('DISABLED');
  cache.remove(key);
  withLock_(() => saveAccount_(code, { last: new Date() }));
  const token = Utilities.getUuid() + Utilities.getUuid();
  props_().setProperty('t:' + token, JSON.stringify({ school: code, h: schoolUser_(code).h }));
  return { token };
}
function apiSchoolChangePassword(token, newPw, currentPw) {
  const su = auth_(token, null, true);
  if (su.role !== 'school') throw new Error('FORBIDDEN');
  newPw = String(newPw || '');
  if (newPw.length < 6 || newPw === su.school) throw new Error('WEAK_PASSWORD');
  if (!su.mustChange && schoolPwHash_(su.school, String(currentPw || '')) !== readAccounts_()[su.school].h) throw new Error('WRONG_PASSWORD');
  return withLock_(() => {
    const h = schoolPwHash_(su.school, newPw);
    saveAccount_(su.school, { h, pw: PW.changed, changed: new Date() });
    apiLogout(token);
    const t = Utilities.getUuid() + Utilities.getUuid();
    props_().setProperty('t:' + t, JSON.stringify({ school: su.school, h }));
    return { token: t, ver: bump_() };
  });
}
function apiSchoolAccounts(token) {
  auth_(token, 'admin');
  const accs = readAccounts_(), roots = readSections_();
  const seen = {}; // عدد المعلومات التي تراها كل مدرسة
  const leavesOf = {};
  const schools = schoolRows_();
  schools.forEach(x => { leavesOf[x.code] = schoolLeaves_(x.code, roots, accs[x.code] || null); });
  rows_('info').forEach(r => { const c = String(r[I.code]); if (r[I.id] && leavesOf[c] && leavesOf[c][String(r[I.leaf])]) seen[c] = (seen[c] || 0) + 1; });
  return { defaults: schoolDefault_(), schools: schools.map(x => { const a = accs[x.code];
    return { code: x.code, name: x.name, wilayah: x.wilayah, status: a && a.status === ST.disabled ? 'disabled' : 'active',
      pw: !a || a.pw === PW.def ? 'default' : a.pw === PW.reset ? 'reset' : 'changed', last: a ? a.last : 0,
      plus: a ? a.plus : [], minus: a ? a.minus : [], sections: Object.keys(leavesOf[x.code]).length, seen: seen[x.code] || 0 }; }) };
}
function apiResetSchoolPassword(token, code) {
  auth_(token, 'admin');
  if (!schoolUser_(code)) throw new Error('NOT_FOUND');
  return withLock_(() => { saveAccount_(String(code).trim(), { h: 'reset:' + Date.now(), pw: PW.reset }); return { ok: true }; });
}
function apiSetSchoolStatus(token, code, active) {
  auth_(token, 'admin');
  if (!schoolUser_(code)) throw new Error('NOT_FOUND');
  return withLock_(() => { saveAccount_(String(code).trim(), { status: active ? ST.active : ST.disabled }); return { ok: true }; });
}

/* ---------- ما تراه حسابات المدارس ----------
 * كل مدرسة ترى معلومات مدرستها فقط، وفي الأقسام المسموحة لها فقط.
 * الأقسام المسموحة = التحديد الجماعي لكل المدارس + أقسام مضافة لهذه المدرسة − أقسام مستبعدة عنها.
 * تحديد قسم رئيسي أو فرعي يشمل كل ما تحته، ومنه الأقسام التي تُضاف لاحقاً. */
const idList_ = v => String(v || '').split(',').map(x => x.trim()).filter(String);
const targets_ = r => idList_(r[I.tgt]);   // عمود المشاركة القديم، للقراءة فقط
const schoolDefault_ = () => idList_(props_().getProperty('schoolSections'));
// يعيد دالة تقول هل القسم (أو أحد آبائه) ضمن مجموعة معرفات
function coverage_(roots) {
  const parent = {};
  const walk = (list, p) => list.forEach(n => { parent[n.id] = p; walk(n.children || [], n.id); });
  walk(roots, '');
  return (set, id) => { for (let x = id; x; x = parent[x]) if (set[x]) return true; return false; };
}
const asSet_ = list => { const o = {}; (list || []).forEach(x => { o[x] = true; }); return o; };
// معرفات الأقسام الأخيرة (التي تحمل المعلومات) المسموحة لهذه المدرسة
function schoolLeaves_(code, roots, acc) {
  const covered = coverage_(roots), def = asSet_(schoolDefault_());
  acc = acc === undefined ? readAccounts_()[String(code)] : acc;
  const plus = asSet_(acc ? acc.plus : []), minus = asSet_(acc ? acc.minus : []);
  const out = {};
  const walk = list => list.forEach(n => { if ((n.children || []).length) walk(n.children); else if ((covered(def, n.id) || covered(plus, n.id)) && !covered(minus, n.id)) out[n.id] = true; });
  walk(roots);
  return out;
}
// تستطيع المدرسة رؤية الصف إذا كان لمدرستها وفي قسم مسموح لها
const schoolSees_ = (r, code, leaves) => String(r[I.code]) === String(code) && !!leaves[String(r[I.leaf])];
function schoolRoots_(code, roots) {
  const leaves = schoolLeaves_(code, roots);
  const filter = list => list.map(n => {
    if (!(n.children || []).length) return leaves[n.id] ? n : null;
    const ch = filter(n.children);
    return ch.length ? Object.assign({}, n, { children: ch }) : null;
  }).filter(Boolean);
  return filter(roots);
}
function apiSetSchoolDefault(token, ids) {
  auth_(token, 'admin');
  return withLock_(() => { props_().setProperty('schoolSections', (ids || []).map(String).join(',')); return { ver: bump_() }; });
}
function apiSetSchoolOverride(token, code, plus, minus) {
  auth_(token, 'admin');
  if (!schoolUser_(code)) throw new Error('NOT_FOUND');
  return withLock_(() => { saveAccount_(String(code).trim(), { plus: (plus || []).map(String).join(','), minus: (minus || []).map(String).join(',') }); return { ver: bump_() }; });
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
  const a = String(r[I.aud] || '');
  it.aud = a === AUD.all ? 'all' : a === AUD.list ? 'list' : '';
  it.targets = targets_(r); it.home = String(r[I.code]); it.homeName = String(r[I.sname] || '');
  return it;
}
// u: المستخدم الحالي (لحساب المدرسة تُعرض المعلومات المشاركة معها فقط، وتُحذف بيانات المشاركة الداخلية)
function infoFor_(code, allowed, u) {
  const out = {}, school = u && u.role === 'school';
  const leaves = school ? schoolLeaves_(code, readSections_()) : null;
  rows_('info').forEach(r => {
    if (!r[I.id] || !(school ? schoolSees_(r, code, leaves) : String(r[I.code]) === String(code))) return;
    const leaf = String(r[I.leaf]);
    if (allowed && !allowed[leaf]) return;
    const it = rowToItem_(r);
    if (school) { delete it.url; delete it.targets; delete it.aud; delete it.home; delete it.homeName; }
    (out[leaf] = out[leaf] || { items: [] }).items.push(it);
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
  if (item.share) { r[I.aud] = AUD.list; r[I.tgt] = String(code); }
  return r;
}
function leafResult_(u, code, leaf) {
  const info = infoFor_(code, null, u);
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
  const u = auth_(token, null, true);
  const roots = readSections_();
  const schools = schoolRows_();
  if (u.role === 'school') return {
    me: { id: u.id, name: u.name, role: 'school', school: u.school, mustChange: u.mustChange, hidden: [] },
    config: { roots: u.mustChange ? [] : schoolRoots_(u.school, roots) },
    schools: u.mustChange ? [] : schools.filter(x => x.code === u.school),
    users: null, ver: ver_(), unreadNotes: 0, driveUrl: ''
  };
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
function apiPoll(token) { auth_(token, null, true); return ver_(); }

/* ---------- الإحصائيات ---------- */
// يعيد [رمز المدرسة، معرف القسم، النوع] لكل معلومة يُسمح للمستخدم برؤيتها
function apiStats(token) {
  const u = auth_(token);
  if (u.role === 'school') throw new Error('FORBIDDEN');
  const allowed = u.role === 'viewer' ? allowedLeaves_(u, readSections_()) : null;
  return rows_('info').filter(r => r[I.id] && (!allowed || allowed[String(r[I.leaf])]))
    .map(r => [String(r[I.code]), String(r[I.leaf]), String(r[I.type])]);
}

/* ---------- الملاحظات ---------- */
function notesRows_() { return sheet_('notes').getLastRow() < 2 ? [] : rows_('notes'); }
function unreadCount_(u) {
  if (u.role === 'viewer' || u.role === 'school') return 0;
  return notesRows_().filter(r => r[N.id] && String(r[N.byId]) !== u.id && String(r[N.readBy] || '').split(',').indexOf(u.id) < 0).length;
}
function apiAddNote(token, code, leaf, itemId, text) {
  const u = auth_(token);
  text = String(text || '').trim();
  if (!text) throw new Error('NOT_FOUND');
  if (text.length > 5000) throw new Error('TEXT_TOO_LONG');
  if (u.role === 'viewer' && !allowedLeaves_(u, readSections_())[String(leaf)]) throw new Error('FORBIDDEN');
  if (u.role === 'school' && String(code) !== u.school) throw new Error('FORBIDDEN');
  return withLock_(() => {
    const row = findInfoRow_(itemId); if (!row) throw new Error('NOT_FOUND');
    const ir = sheet_('info').getRange(row, 1, 1, HEAD.info.length).getValues()[0];
    if (u.role === 'school' && !schoolSees_(ir, u.school, schoolLeaves_(u.school, readSections_()))) throw new Error('FORBIDDEN');
    const label = String(ir[I.title] || ir[I.fname] || '') || ({ text: 'نص', image: 'صورة', file: 'مستند', mindmap: 'خريطة ذهنية', table: 'جدول' }[String(ir[I.type])] || 'معلومة');
    const out = new Array(HEAD.notes.length).fill('');
    out[N.id] = newId_(); out[N.code] = String(code); out[N.sname] = String(ir[I.sname] || schoolName_(code)); out[N.leaf] = String(leaf);
    out[N.path] = String(ir[I.path] || ''); out[N.item] = String(itemId); out[N.ititle] = label; out[N.text] = text;
    out[N.by] = u.name; out[N.byId] = u.id; out[N.role] = u.role; out[N.at] = new Date(); out[N.readBy] = u.role === 'viewer' || u.role === 'school' ? '' : u.id;
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
  if (u.role === 'school' && String(code) !== u.school) throw new Error('FORBIDDEN');
  return infoFor_(code, u.role === 'viewer' ? allowedLeaves_(u, readSections_()) : null, u);
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
// opts (اختياري، للإدخال الجماعي): { title, share: إظهارها في حساب المدرسة, bulk: لا تُعِد قائمة القسم }
function apiUpload(token, code, leaf, file, opts) {
  const u = auth_(token, 'edit');
  opts = opts || {};
  if (opts.bulk && !schoolName_(code)) throw new Error('NOT_FOUND');
  const bytes = Utilities.base64Decode(file.data);
  if (bytes.length > MAX_FILE_MB * 1048576) throw new Error('TOO_LARGE');
  const sname = schoolName_(code);
  const folder = childFolder_(childFolder_(rootFolder_(), 'المرفقات'), (String(code) + ' - ' + sname).trim());
  const f = folder.createFile(Utilities.newBlob(bytes, file.mime || 'application/octet-stream', file.name || 'ملف'));
  const item = { id: newId_(), type: /^image\//.test(file.mime) ? 'image' : 'file', title: String(opts.title || '').slice(0, 200), name: f.getName(), fileId: f.getId(), url: f.getUrl(), size: bytes.length, mime: file.mime, share: !!opts.share };
  return withLock_(() => {
    sheet_('info').appendRow(newInfoRow_(u, code, leaf, item));
    return opts.bulk ? { ver: bump_() } : leafResult_(u, code, leaf);
  });
}
/* ---------- الإدخال الجماعي ----------
 * items: [{ code, type: 'text' | 'table', title, text }] — كل عنصر يُضاف لمدرسته في القسم نفسه.
 * share: تُشارك كل معلومة مع حساب مدرستها فقط. */
function apiBulkAdd(token, leaf, items, share) {
  const u = auth_(token, 'edit');
  const path = pathTo_(readSections_(), String(leaf));
  if (!path) throw new Error('NOT_FOUND');
  const pathName = path.map(n => n.name).join(' › ');
  const names = {}; schoolRows_().forEach(x => { names[x.code] = x.name; });
  const rows = [], skipped = [];
  (items || []).forEach(it => {
    const c = String(it.code || '').trim(), text = String(it.text || '');
    if (!names[c] || !text.trim() || text.length > MAX_TEXT) { skipped.push(c); return; }
    const r = new Array(HEAD.info.length).fill('');
    r[I.code] = c; r[I.sname] = names[c]; r[I.leaf] = String(leaf); r[I.path] = pathName; r[I.id] = newId_();
    r[I.type] = it.type === 'table' ? 'table' : 'text'; r[I.title] = String(it.title || '').slice(0, 200); r[I.text] = text;
    r[I.by] = u.name; r[I.at] = new Date();
    if (share) { r[I.aud] = AUD.list; r[I.tgt] = c; }
    rows.push(r);
  });
  return withLock_(() => {
    if (rows.length) {
      const sh = sheet_('info'), start = sh.getLastRow() + 1;
      sh.getRange(start, I.code + 1, rows.length, 1).setNumberFormat('@');
      sh.getRange(start, I.title + 1, rows.length, 2).setNumberFormat('@');   // نص كما هو، حتى لو بدأ بعلامة =
      sh.getRange(start, I.tgt + 1, rows.length, 1).setNumberFormat('@');
      sh.getRange(start, 1, rows.length, HEAD.info.length).setValues(rows);
    }
    return { added: rows.length, skipped, ver: bump_() };
  });
}

function apiGetFile(token, fileId) {
  const u = auth_(token);
  const row = rows_('info').find(r => String(r[I.fid]) === String(fileId));
  if (!row) throw new Error('NOT_FOUND');
  if (u.role === 'viewer' && !allowedLeaves_(u, readSections_())[String(row[I.leaf])]) throw new Error('FORBIDDEN');
  if (u.role === 'school' && !schoolSees_(row, u.school, schoolLeaves_(u.school, readSections_()))) throw new Error('FORBIDDEN');
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

// يحذف كل المعلومات من كل الأقسام لكل المدارس (للمدير فقط، ويشترط كتابة كلمة «حذف»).
// تُنقل المرفقات إلى سلة المهملات في Drive (قابلة للاسترجاع 30 يوماً). تبقى المدارس والأقسام والمستخدمون والملاحظات.
function apiResetInfo(token, word) {
  auth_(token, 'admin');
  if (String(word || '').trim() !== 'حذف') throw new Error('CONFIRM_WORD');
  return withLock_(() => {
    const sh = sheet_('info'); const n = sh.getLastRow();
    if (n > 1) sh.getRange(2, 1, n - 1, sh.getMaxColumns()).clearContent();
    const it = rootFolder_().getFoldersByName('المرفقات');
    const stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    while (it.hasNext()) { try { const f = it.next(); if (!f.isTrashed()) { f.setName('المرفقات (محذوفة ' + stamp + ')'); f.setTrashed(true); } } catch (e) {} }
    return { deleted: Math.max(0, n - 1), ver: bump_() };
  });
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
  if (me.role === 'school') throw new Error('FORBIDDEN');
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
