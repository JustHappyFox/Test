# Шаблон «Справка юр» (jur_spravka.docx) из шаблона справки пакета юр лица (bs_spravka.docx) - по фото 26.11:
# «Адрес:» и причина - жирным; «Данные приборов учета «ДАТАПУ»» по центру; таблица мощности убрана - вместо неё
# «К справке прилагается:» (по центру, жирным, подчёркнуто) и «ПРИЛОЖЕНИЕ» (строки); должность инженера - поле «ИНЖДОЛЖН».
# python3 jur_spravka_template.py dogovor/bs_spravka.docx dogovor/jur_spravka.docx
import sys, zipfile, copy
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
src, dst = sys.argv[1], sys.argv[2]
zin = zipfile.ZipFile(src)
items = [(i, zin.read(i.filename)) for i in zin.infolist()]
root = etree.fromstring(dict((i.filename, d) for i, d in items)['word/document.xml'])
body = root.find(W + 'body')

def text(p): return ''.join(t.text or '' for t in p.iter(W + 't'))
def set_text(p, s, bold=None, center=None, underline=None, size=None):
    rs = p.findall(W + 'r')
    keep = rs[0] if rs else etree.SubElement(p, W + 'r')
    for r in rs[1:]: p.remove(r)
    for t in keep.findall(W + 't'): keep.remove(t)
    for t in keep.findall(W + 'tab'): keep.remove(t)
    t = etree.SubElement(keep, W + 't'); t.text = s; t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    rpr = keep.find(W + 'rPr')
    if rpr is None: rpr = etree.Element(W + 'rPr'); keep.insert(0, rpr)
    for tag, on in (('b', bold), ('u', underline)):
        if on is None: continue
        e = rpr.find(W + tag)
        if on and e is None: e = etree.SubElement(rpr, W + tag); e.set(W + 'val', 'single') if tag == 'u' else None
        if not on and e is not None: rpr.remove(e)
    if size:
        for tag in ('sz', 'szCs'):
            e = rpr.find(W + tag)
            if e is None: e = etree.SubElement(rpr, W + tag)
            e.set(W + 'val', str(size))
    if center is not None:
        ppr = p.find(W + 'pPr')
        if ppr is None: ppr = etree.Element(W + 'pPr'); p.insert(0, ppr)
        jc = ppr.find(W + 'jc')
        if jc is None: jc = etree.SubElement(ppr, W + 'jc')
        jc.set(W + 'val', center if isinstance(center, str) else 'center' if center else 'both')

kids = list(body)
paras = [k for k in kids if k.tag == W + 'p']
for p in paras:
    t = text(p)
    if t == 'Адрес «АДРЕС».': set_text(p, 'Адрес: «АДРЕС»', bold=True)
    elif t.startswith('В связи с допуском прошу ввести'): set_text(p, '«ПРИЧИНА»', bold=True)
    elif t.startswith('Данные приборов учета для производства расчетов'): set_text(p, 'Данные приборов учета «ДАТАПУ»', bold=True, center=True, size=26)
    elif t.startswith('Инженер по РПЭ'):
        set_text(p, '«ИНЖДОЛЖН» ____________________________________/«ИНЖЕНЕР»', bold=True)
# таблица мощности и её заголовок - убрать; на их место - «К справке прилагается:» и приложения
kids = list(body)
cap = [k for k in kids if k.tag == W + 'p' and text(k).startswith('Данные для производства расчетов по установленной мощности')][0]
tbl = cap.getnext()
assert tbl.tag == W + 'tbl'
att = copy.deepcopy(cap)
set_text(att, 'К справке прилагается:', bold=True, center=True, underline=True)
lst = copy.deepcopy(cap)
set_text(lst, '«ПРИЛОЖЕНИЕ»', bold=False, center='left', underline=False)   # строки - влево (по ширине растягивает)
cap.addprevious(att); cap.addprevious(lst)
body.remove(cap); body.remove(tbl)
data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as z:
    for i, d in items: z.writestr(i, data if i.filename == 'word/document.xml' else d)
print(dst)
