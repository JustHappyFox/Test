# Правки шаблонов .docx на месте (24.11), остальное в файле не трогается; слово может быть разбито Word'ом на куски - тоже:
#  1) «Минскэнерго» -> «МИНСКЭНЕРГО» (всегда большими);
#  2) подпись начальника: «Начальник ГРП» в начале абзаца -> поле «НАЧДОЛЖН» (должность на выбор: Начальник ГРП / ОПСУ),
#     в доп. соглашении БС «Начальник ГРП  ОПСУ» -> «НАЧДОЛЖНДС».
# python3 tpl_fix.py dogovor/*.docx   (после пересоздания шаблонов bs_templates.py / udalenie_template.py - запускать снова)
import sys, zipfile, re, shutil
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

def runs(p): return [t for t in p.iter(W + 't')]

def replace(ts, start, end, new):
    pos, done = 0, False
    for t in ts:
        s = t.text or ''
        a, b = max(start, pos), min(end, pos + len(s))
        if a < b:
            t.text = s[:a - pos] + ('' if done else new) + s[b - pos:]
            done = True
            if t.text and (t.text[0] == ' ' or t.text[-1] == ' '): t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        pos += len(s)

def fix(xml):
    root = etree.fromstring(xml)
    n = 0
    for p in root.iter(W + 'p'):
        ts = runs(p)
        text = ''.join(t.text or '' for t in ts)
        for m in list(re.finditer('минскэнерго', text, re.I)):
            if m.group(0) != 'МИНСКЭНЕРГО':
                replace(ts, m.start(), m.end(), 'МИНСКЭНЕРГО'); n += 1
        text = ''.join(t.text or '' for t in ts)
        m = re.match(r'\s*(Начальник ГРП\s+ОПСУ)', text)
        if m: replace(ts, m.start(1), m.end(1), '«НАЧДОЛЖНДС»'); n += 1; continue
        m = re.match(r'\s*(Начальник ГРП)', text)
        if m: replace(ts, m.start(1), m.end(1), '«НАЧДОЛЖН»'); n += 1
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True), n

for f in sys.argv[1:]:
    zin = zipfile.ZipFile(f)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    total, out = 0, []
    for i, data in items:
        if re.match(r'word/(document|header\d*|footer\d*)\.xml$', i.filename):
            data2, n = fix(data)
            if n: data = data2; total += n
        out.append((i, data))
    if total:
        with zipfile.ZipFile(f + '.tmp', 'w') as z:
            for i, data in out: z.writestr(i, data, compress_type=i.compress_type)
        shutil.move(f + '.tmp', f)
    print(f, total)
