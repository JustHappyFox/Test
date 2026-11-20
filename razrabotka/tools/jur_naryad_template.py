# Шаблон юр наряда (A4) по фото пользователя 20.11 (фото в репозиторий не кладётся): поля «ИМЯ» - для программы.
# python3 jur_naryad_template.py <dogovor/jur_naryad.docx>
import sys
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_UNDERLINE
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(2.5), Cm(1.5), Cm(1.5), Cm(1.5)
st = doc.styles['Normal']
st.font.name = 'Times New Roman'; st.font.size = Pt(14)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
st.paragraph_format.space_after = Pt(0); st.paragraph_format.space_before = Pt(0); st.paragraph_format.line_spacing = 1.0
W = 17.0   # ширина текста, см

def p(align=WD_ALIGN_PARAGRAPH.LEFT, before=0, after=0):
    para = doc.add_paragraph(); para.alignment = align
    para.paragraph_format.space_before = Pt(before); para.paragraph_format.space_after = Pt(after)
    return para

def r(para, text, bold=False, italic=False, under=False, size=None):
    run = para.add_run(text); run.bold = bold; run.italic = italic
    if under: run.underline = True
    if size: run.font.size = Pt(size)
    return run

def line_value(text, bold=True, italic=False, size=None):
    # значение по центру строки, подчёркнуто на всю ширину (табуляции с подчёркиванием)
    para = p(WD_ALIGN_PARAGRAPH.LEFT)
    ts = para.paragraph_format.tab_stops
    ts.add_tab_stop(Cm(W / 2), WD_TAB_ALIGNMENT.CENTER); ts.add_tab_stop(Cm(W), WD_TAB_ALIGNMENT.RIGHT)
    r(para, '\t', under=True)
    r(para, text, bold=bold, italic=italic, under=True, size=size)
    r(para, '\t', under=True)
    return para

def caption(text):
    para = p(WD_ALIGN_PARAGRAPH.CENTER); r(para, text, bold=True, italic=True, size=11); return para

para = p(); r(para, '«ДАТАН»', under=True)
para = p(WD_ALIGN_PARAGRAPH.CENTER, before=6)
r(para, 'НАРЯД № '); r(para, '«НОМЕР»', bold=True, under=True)
line_value('«РЭС»')
para = p(WD_ALIGN_PARAGRAPH.CENTER)
r(para, 'на ', under=True); r(para, '«ВИД» к', bold=True, under=True); r(para, ' электросети электроустановок', under=True)
line_value('«ПОТРЕБИТЕЛЬ»')
caption('(наименование предприятия потребителя)')
line_value('«МЕСТО»')
caption('( место «ВИДА»)')
line_value('«ПРИЧИНА»')
caption('( причина «ВИДА» )')
para = p(before=4); r(para, 'Адрес: ', under=True); r(para, '«АДРЕС»', under=True)
para = p(WD_ALIGN_PARAGRAPH.JUSTIFY)
r(para, 'следующих объектов: '); r(para, '«ОБЪЕКТЫ»', bold=True)
para = p(before=8)
ts = para.paragraph_format.tab_stops; ts.add_tab_stop(Cm(W), WD_TAB_ALIGNMENT.RIGHT)
r(para, 'Подписи должностных лиц, выдавших наряд:')
para = p(); r(para, 'филиал «Энергосбыт»')

def signer(dolzh, fio):
    para = p(before=2)
    ts = para.paragraph_format.tab_stops
    ts.add_tab_stop(Cm(10.5), WD_TAB_ALIGNMENT.LEFT); ts.add_tab_stop(Cm(W), WD_TAB_ALIGNMENT.RIGHT)
    r(para, dolzh, under=True); r(para, '\t', under=True); r(para, '\t' + fio, under=True)
    para = p()
    ts = para.paragraph_format.tab_stops
    ts.add_tab_stop(Cm(8.5), WD_TAB_ALIGNMENT.CENTER); ts.add_tab_stop(Cm(14.5), WD_TAB_ALIGNMENT.CENTER)
    r(para, '      Должность', italic=True, size=10); r(para, '\tПодпись', italic=True, size=10); r(para, '\tФ.И.О.', italic=True, size=10)

signer('«НАЧДОЛЖН»', '«НАЧАЛЬНИК»')
signer('«ИНЖДОЛЖН»', '«ИНЖЕНЕР»')
# линия отреза - пунктир через всю строку
para = p(WD_ALIGN_PARAGRAPH.CENTER, before=10, after=6)
pPr = para._p.get_or_add_pPr(); bdr = OxmlElement('w:pBdr'); b = OxmlElement('w:bottom')
for k, v in (('w:val', 'dashed'), ('w:sz', '8'), ('w:space', '1'), ('w:color', 'auto')): b.set(qn(k), v)
bdr.append(b); pPr.append(bdr)
r(para, 'Линия отреза', bold=True, italic=True, size=11)
para = p(before=4); r(para, 'Отрывной корешок наряда № '); r(para, '«НОМЕР»', under=True)
para = p(); r(para, 'Подлежит возврату в Энергосбыт, в 5-дневный срок', under=True)
para = p(); r(para, 'Согласно наряду № '); r(para, '«НОМЕР»', under=True); r(para, ' от '); r(para, '«ДАТАНК»', under=True)
para = p(); r(para, 'произведено '); r(para, '«ВИДП»', under=True); r(para, ' ________________________ г.')
para = p(); r(para, 'Дата '); r(para, '«ВИДА2»', under=True); r(para, ': «____» _______________ 20____ г.')
para = p(before=6)
ts = para.paragraph_format.tab_stops; ts.add_tab_stop(Cm(W), WD_TAB_ALIGNMENT.RIGHT)
r(para, 'Начальник (мастер) участка РЭС ________________ / ____________________')
para = p(); ts = para.paragraph_format.tab_stops; ts.add_tab_stop(Cm(14), WD_TAB_ALIGNMENT.CENTER)
r(para, '\tФИО', italic=True, size=10)
cp = doc.core_properties; cp.author = ''; cp.title = 'Шаблон «Оформление договора»: юр наряд'; cp.last_modified_by = ''
doc.save(sys.argv[1])
print('ok', sys.argv[1])
