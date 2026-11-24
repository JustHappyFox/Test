# 24.11: шаблон «Потерь» (bs_poteri.docx): после схемы питания - одно поле «ПОТЕРИРАСЧЕТ», куда программа вставляет элементы схемы,
# сведения для расчёта и порядок расчёта с формулами (razrabotka/src/poteri.cs). Шапка, схема и подписи - как в образце.
# Старая часть (фиксированные формулы-рисунки для линии 0,4 кВ) убирается. Повторный запуск ничего не меняет.
# python3 poteri_template.py dogovor/bs_poteri.docx
import sys, zipfile, re, shutil
from lxml import etree
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
f = sys.argv[1]
zin = zipfile.ZipFile(f)
items = [(i, zin.read(i.filename)) for i in zin.infolist()]
zin.close()
doc = dict((i.filename, d) for i, d in items)['word/document.xml']
root = etree.fromstring(doc)
body = root.find(W + 'body')
els = list(body)
text = lambda e: ''.join(t.text or '' for t in e.iter(W + 't'))
if any('ПОТЕРИРАСЧЕТ' in text(e) for e in els):
    print(f, 'уже'); sys.exit(0)
img = [k for k, e in enumerate(els) if e.find('.//' + WP + 'inline') is not None][0]
end = [k for k, e in enumerate(els) if e.tag == W + 'tbl' and 'Энергоснабжающая' in text(e)][0]
for e in els[img + 1:end]: body.remove(e)
p = etree.Element(W + 'p'); r = etree.SubElement(p, W + 'r'); t = etree.SubElement(r, W + 't'); t.text = '«ПОТЕРИРАСЧЕТ»'
els[img].addnext(p)
data = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
with zipfile.ZipFile(f + '.tmp', 'w') as z:
    for i, d in items: z.writestr(i, data if i.filename == 'word/document.xml' else d, compress_type=i.compress_type)
shutil.move(f + '.tmp', f)
print(f, 'убрано элементов:', end - img - 1)
