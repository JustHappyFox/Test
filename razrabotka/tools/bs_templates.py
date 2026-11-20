# Шаблоны документов БС (юр лицо) из образцов пользователя: места, выделенные красным, -> поля «ИМЯ» программы.
# Образцы с данными абонента в репозиторий НЕ кладутся; готовые шаблоны - проверяются на остатки данных (STOP).
# python3 bs_templates.py <папка с образцами .docx (.doc - заранее в .docx через LibreOffice)> <папка dogovor>
import sys, os, re, zipfile, io
from lxml import etree as ET     # lxml: пространства имён остаются как в файле (иначе Word - «файл повреждён»)
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
q = lambda t: '{%s}%s' % (W, t)

# замены красных мест по порядку в документе; «»-поля - имена полей программы
PU = ['«ПУН»', '«ПУТОЧКА»', '«ПУТИП»', '№ «ПУНОМЕР»', '«ПУВ»', '«ПУА»', '«ПУПОКАЗ»', '«ПУКВ»', ' «ПУГОД» г.  ', '«ПУТТ»', '«ПУК»']
MOSHN = ['«МН»', '«МНАГР»', '«ММОЩН»', '«МКИ»', '«МВРЕМЯ»', '«МВЫХ»']
DOCS = {
  'bs_spravka.docx': ('Справка .docx',
      ['№ «ДОГНОМ»', 'Абонент: «АБОНЕНТ»', '«ОБЪЕКТ»', 'Адрес «АДРЕС».', '«ДАТАПУ»'] + PU + MOSHN + ['«ДАТАПОДП» г.'] * 3,
      [('Черницкий Н.В.', '«ИНЖЕНЕР»'), ('Зенченко Д.И.', '«НАЧАЛЬНИК»')]),
  'bs_akt.docx': ('акт приема.docx',
      ['«ДАТААКТ»', '«АБОНЕНТ»', '«ОБЪЕКТ»', '«АДРЕС».', '«ТЕЛЕФОН»', '«ИНЖЕНЕРФИО»', '«АДРЕС».', '«ПИТАНИЕ»', '«ДАТАРАЗГР»', '«ЛИНИЯ»']
      + PU + ['«¶МИНУС»'] + MOSHN + ['«ПУТИПНОМ1»', '«ПЛОМБЫ»', '«ИНЖЕНЕР»'],
      []),
  'bs_dopsogl.docx': ('доп. соглашение с объектом на возобновление.docx',
      ['«ДОГНОМ»', '«ДОГДАТА»', '«АБОНЕНТ»', 'Абонент: «АБОНЕНТ»', 'Объект: «ОБЪЕКТ»', 'Адрес: «АДРЕС».',
       '«ОБЪЕКТАДРЕС»', '«ПУТОЧКА»', '«ПУТИПНОМ»', '«ПУКВ»', ' «ПУГОД» г.  ', '«ПУТТ»', '«ПУТН»', '«ПУК»', '«ММОЩН»,', '«МРЕЖИМ»',
       '«ТАРИФГР»', '«ТРЭТ»', '«ПРИМ»', '«ИНЖЕНЕР»'],
      [('Дополнить Приложение № 3 к Договору   следующим объектом и местом установки средств расчетного учета, по которому производится расчет за отпущенную электроэнергию:', '«ПУНКТ1»'),
       ('Зенченко Д.И.', '«НАЧАЛЬНИК»')]),
  'bs_naryad.docx': ('наряд.docx',
      ['«ВИД»', '«ПОТРЕБИТЕЛЬ»', '«ПРИЧИНА»', '«МЕСТО»', '«ОБЪЕКТКР»', '«АДРЕСНАР»', '«ВИД»', '«ОБЪЕКТКР»', '«АДРЕСКР»', '«ИНЖЕНЕР»'],
      [('Зенченко Д.И.', '«НАЧАЛЬНИК»')]),
  'bs_poteri.docx': ('потери.docx',
      ['«ДОГНОМ»', '«АБОНЕНТ»', '«ОБЪЕКТ»', '«АДРЕС».', '«ДАТАРАЗГР»', '«ЛН»', '«ЛU»', '«ЛМАРКА»', '«ЛДЛИНА»', '«ЛR0»', '«ЛR»', '«ИНЖЕНЕР»', '«НАЧАЛЬНИК»'],
      []),
}
# чего в готовых шаблонах быть не должно (данные абонента из образцов)
STOP = ['Мобильн', 'ТелеСистем', 'МТС', '6966', 'Копище', 'Авиационн', '2602715520', '11098', '4948093', '4726', '80255191', '80247937',
        '510-11-25', 'Gavrosh', 'IMG_2026']

def red(r):
    rp = r.find(q('rPr'))
    if rp is None: return False
    h = rp.find(q('highlight')); c = rp.find(q('color'))
    return (h is not None and h.get(q('val')) == 'red') or (c is not None and c.get(q('val')) == 'FF0000')

def unred(r):
    rp = r.find(q('rPr'))
    for tag in ('highlight',):
        for e in rp.findall(q(tag)): rp.remove(e)
    c = rp.find(q('color'))
    if c is not None and c.get(q('val')) == 'FF0000': rp.remove(c)

def set_text(r, text):
    for e in list(r):
        if e.tag in (q('t'), q('tab'), q('br'), q('cr')): r.remove(e)
    t = ET.SubElement(r, q('t')); t.text = text
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def process(xml, repl, plain):
    root = ET.fromstring(xml)
    groups = []
    for p in root.iter(q('p')):
        cur = []
        kids = [k for k in p if k.tag == q('r')]
        for r in kids:
            has_text = any(e.tag in (q('t'), q('tab'), q('br')) for e in r)
            if red(r) and has_text: cur.append(r)
            elif has_text:
                if cur: groups.append((p, cur)); cur = []
        if cur: groups.append((p, cur))
    return root, groups

def main(src, dst):
    for out, (name, repl, plain) in DOCS.items():
        z = zipfile.ZipFile(os.path.join(src, name))
        raw = z.read('word/document.xml').decode('utf-8')
        root, groups = process(raw.encode('utf-8'), repl, plain)
        if len(groups) != len(repl):
            for i, (p, g) in enumerate(groups): print(i, ''.join(''.join(e.text or '' for e in r.iter(q('t'))) for r in g))
            raise SystemExit('%s: красных мест %d, замен %d' % (name, len(groups), len(repl)))
        for (p, g), new in zip(groups, repl):
            set_text(g[0], new); unred(g[0])
            for r in g[1:]: p.remove(r)
        xml = ET.tostring(root, encoding='unicode')
        assert 'xmlns:w14=' in xml or 'w14' not in raw[:3000], 'потерялись пространства имён'
        # не красное, но меняется (подписи): заменить текст в узлах w:t
        for old, new in plain:
            n = xml.count('>' + old + '<')
            if n == 0:
                # текст разбит на несколько w:t - собрать абзац заново не будем, ищем целиком в одном узле
                raise SystemExit('%s: не нашлось «%s» одним куском' % (out, old))
            xml = xml.replace('>' + old + '<', '>' + new + '<')
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as o:
            for i in z.infolist():
                data = z.read(i.filename)
                if i.filename == 'word/document.xml': data = xml.encode('utf-8')
                elif i.filename == 'docProps/core.xml':
                    data = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
                            'xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
                            '<dc:title>Шаблон «Оформление договора»: БС</dc:title></cp:coreProperties>').encode('utf-8')
                elif i.filename == 'docProps/app.xml':
                    data = re.sub(rb'<Company>[^<]*</Company>', b'<Company></Company>', data)
                    data = re.sub(rb'<Manager>[^<]*</Manager>', b'', data)
                elif i.filename == 'docProps/custom.xml':
                    continue
                elif i.filename.startswith('word/media/') and i.filename.endswith('.jpeg'):
                    data = placeholder()
                o.writestr(i, data)
        result = buf.getvalue()
        # остатки данных абонента
        zz = zipfile.ZipFile(io.BytesIO(result))
        text = ''
        for n in zz.namelist():
            if n.endswith('.xml') or n.endswith('.rels'): text += zz.read(n).decode('utf-8', 'replace')
        text = re.sub(r'descr="[^"]*"', 'descr=""', text)
        bad = [s for s in STOP if s in text]
        if bad: raise SystemExit('%s: остались данные образца: %s' % (out, bad))
        if 'descr="C:' in zz.read('word/document.xml').decode('utf-8'):
            result = result  # путь к фото у пользователя - убираем ниже
        open(os.path.join(dst, out), 'wb').write(strip_descr(result))
        print(out, 'ок, полей', len(set(re.findall(r'«[□¶]?([A-ZА-ЯЁ0-9_]+)»', xml))))

def strip_descr(data):
    zin = zipfile.ZipFile(io.BytesIO(data)); buf = io.BytesIO()
    with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as o:
        for i in zin.infolist():
            d = zin.read(i.filename)
            if i.filename == 'word/document.xml': d = re.sub(rb'descr="[^"]*"', b'descr=""', d)
            o.writestr(i, d)
    return buf.getvalue()

def placeholder():
    from PIL import Image, ImageDraw, ImageFont
    im = Image.new('RGB', (930, 1000), (238, 238, 238)); d = ImageDraw.Draw(im)
    d.rectangle([6, 6, 923, 993], outline=(150, 150, 150), width=4)
    try: f = ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf', 40)
    except Exception: f = None
    d.text((80, 440), 'Схема питания электроустановки', fill=(90, 90, 90), font=f)
    d.text((80, 500), '(вставляется программой из файла)', fill=(120, 120, 120), font=f)
    b = io.BytesIO(); im.save(b, 'JPEG', quality=80); return b.getvalue()

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
