# Шаблон «Справка в расчетную группу» (удаление ПУ из юр базы) по фото пользователя 24.11 (фото в репозиторий не кладётся).
# Два экземпляра на листе делает программа (DocxTemplate.TwoCopies), как у справки. python3 udalenie_template.py dogovor/udalenie.docx
import sys
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn

doc = Document()
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21), Cm(29.7)
sec.left_margin, sec.right_margin, sec.top_margin, sec.bottom_margin = Cm(2.0), Cm(1.5), Cm(1.0), Cm(1.0)
st = doc.styles['Normal']
st.font.name = 'Times New Roman'; st.font.size = Pt(14)
st.element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
pf = st.paragraph_format; pf.space_after = Pt(0); pf.space_before = Pt(0); pf.line_spacing = 1.0
W = 17.5

def p(align=WD_ALIGN_PARAGRAPH.LEFT, before=0):
    para = doc.add_paragraph(); para.alignment = align; para.paragraph_format.space_before = Pt(before); return para

def r(para, text, bold=False, under=False):
    run = para.add_run(text); run.bold = bold; run.underline = under or None; return run

r(p(), '«ДОГКАРТ»', bold=True, under=True)
r(p(WD_ALIGN_PARAGRAPH.CENTER, before=10), 'Справка в расчетную группу', bold=True, under=True)
para = p(WD_ALIGN_PARAGRAPH.JUSTIFY, before=10)
para.paragraph_format.first_line_indent = Cm(0)
r(para, 'В связи с продажей «ВИД» (с собственником заключен договор на оказание услуги электроснабжения №«ЛС») по договору №«ДОГНОМ» «ДОГНАИМ» прошу исключить из расчетов следующий ПУ:')
r(p(), 'Карточка №«КАРТОЧКА»', bold=True, under=True)
r(p(), 'Объект: «ОБЪЕКТ»', bold=True, under=True)
r(p(), 'Расположенный по адресу: «АДРЕС»', bold=True, under=True)
r(p(), 'Данные ПУ на момент продажи')
r(p(), '«ПУ»')
p()

def sign(left, right, before):
    para = p(before=before)
    para.paragraph_format.tab_stops.add_tab_stop(Cm(W), WD_TAB_ALIGNMENT.RIGHT)
    r(para, left + '\t' + right)
    return para

sign('Инженер по работе с потребителями', '«ИНЖЕНЕР»', 12)
r(p(), '«ДАТА»')
sign('Начальник ГРП', '«НАЧАЛЬНИК»', 18)
r(p(before=18), 'Справку получила ________________ « ___ » ____________ «ГОД» ____________')
cp = doc.core_properties; cp.author = ''; cp.last_modified_by = ''; cp.title = 'Шаблон «Оформление договора»: справка в расчетную группу'
doc.save(sys.argv[1]); print('ok', sys.argv[1])
