#!/usr/bin/env python3
"""
Builds the deployable Index.html from src/Index.source.html.

Google Apps Script's HTML service rewrites inline JavaScript and can cut parts of
some lines (seen with URLs and template literals). To keep the app code intact, the
main script is shipped as base64url text (letters, digits, '-' and '_' only) and a
small loader decodes and runs it in the browser.

Usage:  python3 build.py
"""
import base64, pathlib, re, sys

root = pathlib.Path(__file__).parent
src = (root / 'src' / 'Index.source.html').read_text(encoding='utf-8')

start = src.index('<script>\n(() => {')
end = src.index('</script>', start)
code = src[start + len('<script>\n'):end]

payload = base64.urlsafe_b64encode(code.encode('utf-8')).decode('ascii').rstrip('=')
lines = '\n'.join(payload[i:i + 120] for i in range(0, len(payload), 120))

loader = (
    '<script type="text/plain" id="appCode">\n' + lines + '\n</script>\n'
    '<script>\n'
    '/* Decodes and runs the app code (see build.py for why it is encoded). */\n'
    '(function () {\n'
    "  var t = document.getElementById('appCode').textContent;\n"
    "  t = t.split('\\n').join('').split('\\r').join('').split(' ').join('');\n"
    "  t = t.split('-').join('+').split('_').join('/');\n"
    "  while (t.length % 4) t += '=';\n"
    '  var b = atob(t), u = new Uint8Array(b.length);\n'
    '  for (var i = 0; i < b.length; i++) u[i] = b.charCodeAt(i);\n'
    "  var s = document.createElement('script');\n"
    '  s.text = new TextDecoder().decode(u);\n'
    '  document.body.appendChild(s);\n'
    '})();\n'
    '</script>'
)
out = src[:start] + loader + src[end + len('</script>'):]
(root / 'Index.html').write_text(out, encoding='utf-8')

# the deployable file must not contain anything a JS comment-stripper could cut
inline = re.findall(r'<script>(.*?)</script>', out, re.S)
bad = [i for i, s in enumerate(inline) if '//' in s]
print(f'Index.html: {out.count(chr(10)) + 1} lines, {len(out):,} chars; inline scripts with "//": {bad or "none"}')
sys.exit(1 if bad else 0)
