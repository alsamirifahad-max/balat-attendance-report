#!/usr/bin/env python3
"""Reproducible build: src/ + pinned npm packages (tooling/) -> site/ (fully self-hosted, no third-party requests)."""
import hashlib, os, re, shutil, sys

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(ROOT, 'src')
OUT = os.path.join(ROOT, 'site')
NM = os.path.join(ROOT, 'tooling', 'node_modules')

tpl = open(os.path.join(SRC, 'page_template.html'), encoding='utf-8').read()
parts = [open(os.path.join(SRC, f), encoding='utf-8').read() for f in ('core_parts.js', 'core_tail.js', 'core_tail2.js')]
core = ('const Core=(function(){\n' + parts[0] + '\n' + parts[1] + '\n' + parts[2] +
        '\nreturn {makeExtractor,parseTeachers,parseStudents,buildModel,fillDocx,guessSubject,pct,docxContent,docxKind,'
        'parseStudentsDocx,parseTeachersDocx,finishTeachers,formInfo,makeListDocx,parseDates};\n})();\n')

shutil.rmtree(OUT, ignore_errors=True)
for d in ('assets/fonts', 'vendor'):
    os.makedirs(os.path.join(OUT, d), exist_ok=True)

def sha(path):
    return hashlib.sha256(open(path, 'rb').read()).hexdigest()

# ---- vendor libraries (pinned, copied from npm packages) ----
VENDOR = {
    'pdf-lib.min.js': 'pdf-lib/dist/pdf-lib.min.js',
    'jszip.min.js': 'jszip/dist/jszip.min.js',
    'html2canvas.min.js': 'html2canvas/dist/html2canvas.min.js',
    'jspdf.umd.min.js': 'jspdf/dist/jspdf.umd.min.js',
}
for dst, src in VENDOR.items():
    data = open(os.path.join(NM, src), 'rb').read()
    # strip sourceMappingURL comments (maps are not shipped)
    data = re.sub(rb'\n?//# sourceMappingURL=\S+\s*$', b'\n', data)
    open(os.path.join(OUT, 'vendor', dst), 'wb').write(data)

# ---- fonts (self-hosted, no third-party requests) ----
FONTS = []
for w in (400, 500, 600, 700):
    for sub in ('arabic', 'latin'):
        f = f'ibm-plex-sans-arabic-{sub}-{w}-normal.woff2'
        shutil.copy(os.path.join(NM, '@fontsource/ibm-plex-sans-arabic/files', f), os.path.join(OUT, 'assets/fonts', f))
        FONTS.append(('IBM Plex Sans Arabic', w, sub, f))
for w in (500, 600, 700):
    for sub in ('arabic', 'latin'):
        f = f'readex-pro-{sub}-{w}-normal.woff2'
        shutil.copy(os.path.join(NM, '@fontsource/readex-pro/files', f), os.path.join(OUT, 'assets/fonts', f))
        FONTS.append(('Readex Pro', w, sub, f))
RANGE = {'arabic': 'U+0600-06FF,U+0750-077F,U+0870-088E,U+0890-0891,U+0897-08E1,U+08E3-08FF,U+200C-200E,U+2010-2011,U+204F,U+2E41,U+FB50-FDFF,U+FE70-FEFC',
         'latin': 'U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+0304,U+0308,U+0329,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD'}
fontcss = ''.join(
    f'@font-face{{font-family:"{fam}";font-style:normal;font-weight:{w};font-display:swap;src:url(fonts/{f}) format("woff2");unicode-range:{RANGE[sub]}}}\n'
    for fam, w, sub, f in FONTS)

# ---- default form ----
shutil.copy(os.path.join(SRC, 'form.docx'), os.path.join(OUT, 'assets/form.docx'))

# ---- split template ----
m_style = re.search(r'<style>(.*?)</style>', tpl, re.S)
css = fontcss + m_style.group(1).strip() + '\n'
scripts = re.findall(r'<script>(.*?)</script>', tpl, re.S)
assert len(scripts) == 2, len(scripts)
app = scripts[1]

# load the built-in form from a same-origin file instead of an inline base64 blob
app = app.replace('const DEFAULT_FORM_B64="__FORM__";\n', '')
app = app.replace(
    "files.form={name:'الاستمارة المحفوظة (بلاط الشهداء)',bytes:b64ToBytes(DEFAULT_FORM_B64),builtin:true};",
    """fetch('assets/form.docx',{credentials:'omit',cache:'no-cache'}).then(r=>{ if(!r.ok) throw new Error(r.status); return r.arrayBuffer() })
  .then(b=>{ if(!files.form){ files.form={name:'الاستمارة المحفوظة (بلاط الشهداء)',bytes:new Uint8Array(b),builtin:true}; refreshSlots() } })
  .catch(()=>setStatus('تعذّر تحميل الاستمارة المحفوظة. ارفع الاستمارة بصيغة Word.',true));""")
assert 'DEFAULT_FORM_B64' not in app or 'b64ToBytes(' in app

# ---- security hardening inside the app ----
SEC = r"""
/* ---- security guards ---- */
const MAX_FILE_BYTES=25*1024*1024;            // 25 MB per uploaded file
const MAX_PDF_PAGES=400;
const MAX_UNZIPPED_BYTES=80*1024*1024;        // zip-bomb guard for .docx
function sniff(bytes){ if(bytes.length>4&&bytes[0]===0x25&&bytes[1]===0x50&&bytes[2]===0x44&&bytes[3]===0x46) return 'pdf'; if(bytes.length>4&&bytes[0]===0x50&&bytes[1]===0x4b&&bytes[2]===0x03&&bytes[3]===0x04) return 'zip'; return null }
async function guardZip(bytes){ const z=await JSZip.loadAsync(bytes.slice(0)); let total=0, n=0;
  z.forEach((p,f)=>{ n++; const u=f._data&&f._data.uncompressedSize; total+=(typeof u==='number'?u:0) });
  if(n>2000||total>MAX_UNZIPPED_BYTES) throw new Error('ملف Word غير صالح أو كبير جدًا.');
  if(!z.file('word/document.xml')) throw new Error('الملف ليس مستند Word صالحًا.'); }
// neutralise spreadsheet formula injection (=, +, -, @, tab, CR at cell start)
function csvCell(v){ let s=String(v??''); if(/^[=+\-@\t\r]/.test(s)) s="'"+s; return '"'+s.replace(/"/g,'""')+'"' }
"""
app = app.replace("function b64ToBytes(b)", SEC + "\nfunction b64ToBytes(b)", 1)

# size / type checks before any parsing
old_classify = "async function classify(file){\n  const bytes=new Uint8Array(await file.arrayBuffer());\n  const isZip=bytes[0]===0x50&&bytes[1]===0x4b;\n  if(isZip){"
assert old_classify in app
app = app.replace(old_classify, """async function classify(file){
  if(file.size>MAX_FILE_BYTES) throw new Error('الملف «'+file.name+'» أكبر من 25 ميجابايت.');
  const bytes=new Uint8Array(await file.arrayBuffer());
  const kind0=sniff(bytes);
  if(!kind0) throw new Error('الملف «'+file.name+'» ليس PDF أو Word.');
  const isZip=kind0==='zip';
  if(isZip){ await guardZip(bytes);""")
old_pdf = "  const {extract}=Core.makeExtractor(PDFLib); const pages=await extract(bytes.slice(0));\n  const txt="
assert old_pdf in app
app = app.replace(old_pdf, "  const {extract}=Core.makeExtractor(PDFLib); const pages=await extract(bytes.slice(0));\n  if(pages.length>MAX_PDF_PAGES) throw new Error('ملف PDF طويل جدًا.');\n  const txt=")
old_catch = "    }catch(e){ console.error(e); err='تعذّرت قراءة «'+f.name+'». تأكد أن الملف PDF أو Word سليم.' }"
assert old_catch in app
app = app.replace(old_catch, "    }catch(e){ console.error(e); err=(e&&/[\\u0600-\\u06FF]/.test(e.message||''))?e.message:'تعذّرت قراءة «'+f.name+'». تأكد أن الملف PDF أو Word سليم.' }")
# CSV through the sanitiser
old_csv = """'\ufeff'+rows.map(r=>r.map(v=>'"'+String(v).replace(/"/g,'""')+'"').join(',')).join('\\r\\n')"""
assert old_csv in app, 'csv'
app = app.replace(old_csv, """'\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\\r\\n')""")
# downloads: plain browser download on the public site
old_dl = "function getDownloads(){ if(!downloadsP) downloadsP=(window.claude&&window.claude.use)?window.claude.use('downloads').catch(()=>null):Promise.resolve(null); return downloadsP }"
assert old_dl in app
app = app.replace(old_dl, "function getDownloads(){ return Promise.resolve(null) }")
# file names: keep only safe characters
old_fn = "function fname(base,ext){ const d=curDate(); return"
assert old_fn in app
app = app.replace(old_fn, "function fname(base,ext){ return safeName(fname0(base,ext)) }\nfunction safeName(s){ return s.replace(/[\\\\/:*?\"<>|\\u0000-\\u001f]/g,'_').slice(0,150) }\nfunction fname0(base,ext){ const d=curDate(); return")
open(os.path.join(OUT, 'assets/app.css'), 'w', encoding='utf-8').write(css)
open(os.path.join(OUT, 'assets/core.js'), 'w', encoding='utf-8').write("'use strict';\n" + core)
open(os.path.join(OUT, 'assets/app.js'), 'w', encoding='utf-8').write("'use strict';\n(function(){\n" + app + "\n})();\n")

# ---- body markup ----
body = tpl[m_style.end():]
body = re.sub(r'<script src="https://cdnjs[^"]+"></script>\s*', '', body)
body = re.sub(r'<script>.*?</script>\s*', '', body, flags=re.S).strip()
assert 'https://' not in body, re.findall(r'https://\S+', body)

def v(rel):
    return f'{rel}?v={sha(os.path.join(OUT, rel))[:12]}'

scripts_html = '\n'.join(f'<script src="{v(p)}"></script>' for p in (
    'vendor/pdf-lib.min.js', 'vendor/jszip.min.js', 'vendor/html2canvas.min.js', 'vendor/jspdf.umd.min.js', 'assets/core.js', 'assets/app.js'))
html = f"""<!doctype html>
<html lang="ar" dir="rtl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="robots" content="noindex,nofollow,noarchive">
<meta name="referrer" content="no-referrer">
<meta name="color-scheme" content="light dark">
<title>تقرير الحضور والتحضير</title>
<link rel="icon" href="data:,">
<link rel="stylesheet" href="{v('assets/app.css')}">
</head>
<body>
{body}
{scripts_html}
</body>
</html>
"""
open(os.path.join(OUT, 'index.html'), 'w', encoding='utf-8').write(html)

open(os.path.join(OUT, '404.html'), 'w', encoding='utf-8').write("""<!doctype html><html lang="ar" dir="rtl"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="robots" content="noindex"><title>الصفحة غير موجودة</title></head>
<body><main><h1>الصفحة غير موجودة</h1><p><a href="/">العودة إلى تقرير الحضور والتحضير</a></p></main></body></html>
""")
open(os.path.join(OUT, 'robots.txt'), 'w').write('User-agent: *\nDisallow: /\n')

# integrity manifest for every shipped file (checked in CI)
lines = []
for root, _, files in os.walk(OUT):
    for fn in sorted(files):
        p = os.path.join(root, fn)
        rel = os.path.relpath(p, OUT)
        if rel == 'SHA256SUMS':
            continue
        lines.append(f'{sha(p)}  {rel}')
open(os.path.join(OUT, 'SHA256SUMS'), 'w').write('\n'.join(sorted(lines, key=lambda s: s.split('  ')[1])) + '\n')
print('built', OUT, len(lines), 'files')
