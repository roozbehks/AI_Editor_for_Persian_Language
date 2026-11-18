#!/usr/bin/env python3
"""صفحهٔ «ویراستار هوشواره» را برای GitHub Pages می‌سازد.

    pip install markdown
    python3 tools/build_site.py

از persian-writing/ می‌خواند و این‌ها را در docs/ می‌نویسد: index.html، kit.json (همهٔ پرونده‌های بسته،
برای بارگیری و راهنماها)، fa_lint.py (همان ابزار، برای صفحه و افزونهٔ ورد) و word/manifest.xml.
پس از هر تغییر در persian-writing/ یا word-addin/manifest.xml دوباره اجرا کنید.
"""
import json
import os
import re
import shutil
import sys

import markdown

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PKG = ROOT + '/persian-writing'
OUT = ROOT + '/docs'
REPO = 'https://github.com/roozbehks/AI_Editor_for_Persian_Language'
sys.path.insert(0, PKG + '/scripts')
import fa_lint as F  # noqa: E402

os.makedirs(OUT, exist_ok=True)

# ── kit.json: همهٔ پرونده‌های بستهٔ عمومی ──
files = {}
for base, _, names in os.walk(PKG):
    if '__pycache__' in base:
        continue
    for n in sorted(names):
        if n.endswith('.pyc'):
            continue
        p = os.path.join(base, n)
        files[os.path.relpath(p, PKG)] = open(p, encoding='utf-8').read()
guides = [
    ('references/rasm-ol-khat.md', 'رسم‌الخط', 'فاصله و نیم‌فاصله، وندها، واژهٔ مرکب، همزه، تنوین، املاهای دوگانه'),
    ('references/jomle.md', 'جمله', 'حذف فعل، وجه وصفی، «را»، مجهول، مطابقت، ابهام'),
    ('references/vazhe.md', 'واژه', 'حشو، درازنویسی، گرته‌برداری، جمع، غلط‌های رایج، وام‌واژه'),
    ('references/neshanegozari.md', 'نشانه‌گذاری', 'ویرگول و گیومه، عدد و تاریخ، عنوان و فهرست، متن دوزبانه'),
    ('references/tarjome.md', 'از انگلیسی به فارسی', 'عادت‌های نحوی، واژه‌های دام، تصویر فرهنگی، لحن'),
    ('references/manabe.md', 'منابع', 'قاعده‌ها از کجا آمده است و منابع کجا اختلاف دارند'),
]
json.dump({'version': F.__version__, 'files': files}, open(OUT + '/kit.json', 'w', encoding='utf-8'), ensure_ascii=False)

# ── تکه‌های ایستا ──
skill = files['SKILL.md']
a = skill.index('## قاعده‌های هسته')
b = skill.index('## کدام مرجع را کی بخوانم')
core_md = skill[a:b].replace('## قاعده‌های هسته\n', '')
core_md = re.sub(r'\(جدول در `references/tarjome.md`\)', '(جدول در راهنمای «از انگلیسی به فارسی»)', core_md)
core_html = markdown.markdown(core_md, extensions=['tables'])
core_html = core_html.replace('<h2>هفت پرسش پایانی</h2>', '<h3>هفت پرسش پایانی</h3>')
prompt = re.search(r'```text\n(.*?)\n```', files['PROMPT.md'], flags=re.S).group(1)

SAMPLE = ('لازم به ذکر است که پیشنهادات مربوطه توسط کمیته مورد بررسی قرار می گیرد و نتیجه آن متعاقبا به اطلاع کارمندان محترم رسانده می شود. '
          'اگرچه این فرایند زمان بر می باشد , اما ما معتقدیم که کیفیت کار ها را بهتر میکند.')

PY_GLUE = r'''
import json, fa_lint
def run_check(text, strict=False):
    first = fa_lint.lint(text, {}, strict)
    fixed, rest = fa_lint.fix_text(text, {}, strict)
    nfix = sum(1 for f in first if f[5] is not None and (f[1] in 'EW' or f[0] in fa_lint.AUTOFIX_C))
    out = []
    for f in rest:
        a, b = fixed[max(0, f[2] - 30):f[2]], fixed[f[3]:f[3] + 30]
        if f[2] > 30 and ' ' in a:
            a = a[a.index(' ') + 1:]
        if f[3] + 30 < len(fixed) and ' ' in b:
            b = b[:b.rindex(' ')]
        out.append({'code': f[0], 'level': f[1], 'msg': f[4], 'a': a.replace('\n', ' '), 'hit': fixed[f[2]:f[3]], 'b': b.replace('\n', ' ')})
    return json.dumps({'fixed': fixed, 'nfix': nfix, 'rest': out}, ensure_ascii=False)
'''
ns = {}
exec(PY_GLUE, ns)
sample_result = ns['run_check'](SAMPLE)


def esc(s):
    return s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')


guide_items = '\n'.join(
    '<li><button type="button" class="guide" data-path="{0}"><span class="g-title">{1}</span><span class="g-desc">{2}</span></button></li>'.format(p, t, d)
    for p, t, d in guides)

html = open(ROOT + '/tools/site/template.html', encoding='utf-8').read()
for k, v in {
    '{{CORE}}': core_html, '{{PROMPT}}': esc(prompt), '{{SAMPLE}}': esc(SAMPLE), '{{GUIDES}}': guide_items,
    '{{SAMPLE_RESULT}}': sample_result.replace('</', '<\\/'), '{{PY_GLUE}}': json.dumps(PY_GLUE),
    '{{VERSION}}': F.__version__.translate(str.maketrans('0123456789.', '۰۱۲۳۴۵۶۷۸۹٫')),
    '{{REPO}}': REPO, '{{WORD_GUIDE}}': REPO + '/tree/main/word-addin#readme', '{{WORD_MANIFEST}}': 'word/manifest.xml',
}.items():
    assert k in html, k
    html = html.replace(k, v)
assert '{{' not in html, re.findall(r'\{\{[A-Z_]+\}\}', html)
open(OUT + '/index.html', 'w', encoding='utf-8').write(html)
shutil.copy(PKG + '/scripts/fa_lint.py', OUT + '/fa_lint.py')
os.makedirs(OUT + '/word', exist_ok=True)
shutil.copy(ROOT + '/word-addin/manifest.xml', OUT + '/word/manifest.xml')
print('index.html', len(html), 'chars | kit.json', os.path.getsize(OUT + '/kit.json'), 'bytes |', len(files), 'files')
