# Правки шаблонов .docx на месте (24.11), остальное в файле не трогается; слово может быть разбито Word'ом на куски - тоже:
#  1) «Минскэнерго» -> «МИНСКЭНЕРГО» (всегда большими);
#  3) акт приемки БС, п. 4: «...используется для электроснабжения базовой станции» -> поле «ДЕЯТЕЛЬНОСТЬ»;
#  8) таблицы шире текста страницы (потери, доп. соглашение) - столбцы ужимаются до ширины текста;
#  7) наряд БС: таблица шире листа - наряд целиком на листе, строка инженера - табуляцией;
#  6) подписи «ЭНЕРГОСНАБЖАЮЩАЯ ОРГАНИЗАЦИЯ ... МП» из нижнего колонтитула - в конец текста документа;
#  5) «ПУГОД» г. -> «ПУГОДГ» (год поверки с «г.», без года - пусто);
#  4) потери: формула в ячейке, плавающая поверх схемы, - в строку ячейки;
#  9) 26.11 по правкам пользователя: справка - абзац «В связи с допуском прошу ввести в расчёты ... в договор «ДОГНОМ» от «ДОГДАТА».»;
#     справка, потери - «Инженер ГРП» -> «Инженер по РПЭ»; наряд - причина и место включения подчёркнуты до конца строки
#     (табуляция с заполнением «_»), место - посередине линии; акт приемки, п. 3 - после линии поле «АКТДОП» (для ЭЗС: тип зарядок,
#     колонки, коннекторы);
#  2) подпись начальника: «Начальник ГРП» в начале абзаца -> поле «НАЧДОЛЖН» (должность на выбор: Начальник ГРП / ОПСУ),
#     в доп. соглашении БС «Начальник ГРП  ОПСУ» -> «НАЧДОЛЖНДС».
# python3 tpl_fix.py dogovor/*.docx   (после пересоздания шаблонов bs_templates.py / udalenie_template.py - запускать снова)
import sys, zipfile, re, shutil, copy
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

V = '{urn:schemas-microsoft-com:vml}'
def naryad_fit(root):
    # 7) наряд БС (альбомный лист): таблица 8364+8208+8208 шире страницы - правый край наряда (средняя колонка) уходил за поле
    #    и обрезался при печати («Допуск в эксплуатаци...»). Пустая левая колонка уже, пустая правая - узкая: наряд целиком на листе.
    #    Строка «Инженер группы ... «ИНЖЕНЕР»» (ставилась пробелами, фамилия переносилась) - отступ под наряд и табуляция вправо.
    n = 0
    for tbl in root.iter(W + 'tbl'):
        cols = tbl.findall(W + 'tblGrid/' + W + 'gridCol')
        if [c.get(W + 'w') for c in cols] != ['8364', '8208', '8208']: continue
        new = ['7564', '8208', '300']
        for c, w in zip(cols, new): c.set(W + 'w', w)
        for tr in tbl.findall(W + 'tr'):
            for tc, w in zip(tr.findall(W + 'tc'), new):
                tw = tc.find(W + 'tcPr/' + W + 'tcW')
                if tw is not None: tw.set(W + 'w', w)
        tw = tbl.find(W + 'tblPr/' + W + 'tblW')
        if tw is not None: tw.set(W + 'w', str(sum(int(x) for x in new)))
        n += 1
        for p in root.iter(W + 'p'):
            ts = runs(p)
            text = ''.join(x.text or '' for x in ts)
            if text.startswith('«ПРИЧИНА»'):      # длинная причина («Допуск в эксплуатацию») - без растягивания по ширине
                jc = p.find(W + 'pPr/' + W + 'jc')
                if jc is not None: jc.set(W + 'val', 'left')
                for x in ts:
                    if (x.text or '').strip().strip('_') == '' and '_' in (x.text or ''): x.text = ' '
                continue
            if 'Инженер группы' not in text or '«ИНЖЕНЕР»' not in text or not text.startswith('    '): continue
            ppr = p.find(W + 'pPr')
            tabs = ppr.find(W + 'tabs')
            for tb in list(tabs): tabs.remove(tb)
            tab = etree.SubElement(tabs, W + 'tab'); tab.set(W + 'val', 'right'); tab.set(W + 'pos', '15700')
            ind = etree.SubElement(ppr, W + 'ind'); ind.set(W + 'left', '7780')
            for x in ts:
                s = x.text or ''
                if s.strip() == '': x.text = ''
                elif 'по работе с потребителями' in s: x.text = s.rstrip(); 
                elif s == '«ИНЖЕНЕР»':
                    r = x.getparent(); tr = etree.Element(W + 'r'); etree.SubElement(tr, W + 'tab'); r.addprevious(tr)
            n += 1
    return n

def tables_fit(root):
    # 8) таблица шире текста страницы (в потерях - до края листа: при печати обрезалась) - столбцы ужать пропорционально
    sect = root.find('.//' + W + 'sectPr')
    if sect is None: return 0
    pg, mar = sect.find(W + 'pgSz'), sect.find(W + 'pgMar')
    if pg is None or mar is None: return 0
    text = int(pg.get(W + 'w')) - int(mar.get(W + 'left')) - int(mar.get(W + 'right'))
    n = 0
    for tbl in root.iter(W + 'tbl'):
        if any(a.tag == W + 'tbl' for a in tbl.iterancestors()): continue
        cols = tbl.findall(W + 'tblGrid/' + W + 'gridCol')
        if not cols: continue
        ind = tbl.find(W + 'tblPr/' + W + 'tblInd')
        i = int(ind.get(W + 'w')) if ind is not None and ind.get(W + 'type', 'dxa') == 'dxa' else 0
        g = sum(int(c.get(W + 'w')) for c in cols)
        room = text - max(i, 0)
        if g <= room + 20: continue
        k = room / g
        for c in cols: c.set(W + 'w', str(int(int(c.get(W + 'w')) * k)))
        for tw in tbl.iter(W + 'tcW'):
            if tw.get(W + 'type') == 'dxa': tw.set(W + 'w', str(int(int(tw.get(W + 'w')) * k)))
        w = tbl.find(W + 'tblPr/' + W + 'tblW')
        if w is not None and w.get(W + 'type') == 'dxa': w.set(W + 'w', str(int(int(w.get(W + 'w')) * k)))
        n += 1
    return n

def text_of(p): return ''.join(t.text or '' for t in runs(p))

def user_26(root):
    n = 0
    for p in list(root.iter(W + 'p')):
        ts = runs(p)
        text = text_of(p)
        # справка: после адреса - просьба ввести приборы учёта в расчёты по договору
        if text == 'Адрес «АДРЕС».' and FILE.endswith('bs_spravka.docx'):
            nxt = p.getnext()
            if nxt is None or not text_of(nxt).startswith('В связи с допуском'):
                q = copy.deepcopy(p)
                qt = runs(q)
                qt[0].text = 'В связи с допуском прошу ввести в расчёты следующие приборы учета, в договор «ДОГНОМ» от «ДОГДАТА».'
                for t in qt[1:]: t.text = ''
                p.addnext(q)
                n += 1
            continue
        for a, b in (('Инженер ГРП _____________________________________/', 'Инженер по РПЭ ____________________________________/'),
                     ('Инженер ГРП      __________     ', 'Инженер по РПЭ   _____________ ')):
            k = text.find(a)
            if k >= 0: replace(ts, k, k + len(a), b); n += 1; text = text_of(p)
        # наряд: причина - подчёркивание до конца строки
        if text.strip().startswith('«ПРИЧИНА»') and p.find(W + 'pPr/' + W + 'tabs') is not None:
            rs = p.findall(W + 'r')
            idx = [i for i, r in enumerate(rs) if ''.join(t.text or '' for t in r.iter(W + 't')) == '«ПРИЧИНА»']
            if idx and not any(r.find(W + 'tab') is not None for r in rs[idx[0] + 1:]):
                for r in rs[idx[0] + 1:]:
                    if ''.join(t.text or '' for t in r.iter(W + 't')).strip() == '' and r.find(W + 'tab') is None: p.remove(r)
                tr = copy.deepcopy(rs[idx[0]])
                for t in tr.findall(W + 't'): tr.remove(t)
                etree.SubElement(tr, W + 'tab')
                rs[idx[0]].addnext(tr)
                n += 1
        # наряд: место включения - посередине линии, линия с обеих сторон до краёв
        if '«МЕСТО»' in text and text.replace('_', '').strip() == '«МЕСТО»':
            ppr = p.find(W + 'pPr')
            tabs = ppr.find(W + 'tabs') if ppr is not None else None
            if tabs is not None and not any(t.get(W + 'val') == 'center' for t in tabs):
                for t in list(tabs): tabs.remove(t)
                for val, pos in (('center', '3818'), ('right', '7636')):
                    t = etree.SubElement(tabs, W + 'tab'); t.set(W + 'val', val); t.set(W + 'leader', 'underscore'); t.set(W + 'pos', pos)
                jc = ppr.find(W + 'jc')
                if jc is not None: jc.set(W + 'val', 'left')
                rs = p.findall(W + 'r')
                m = None
                for r in rs:
                    tx = ''.join(t.text or '' for t in r.iter(W + 't'))
                    if tx == '«МЕСТО»': m = r
                    elif tx.strip('_') == '' and tx: p.remove(r)
                tr = copy.deepcopy(m)
                for t in tr.findall(W + 't'): tr.remove(t)
                etree.SubElement(tr, W + 'tab')
                m.addprevious(tr)
                t = m.find(W + 't'); t.text = ' «МЕСТО» '; t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                n += 1
        # подчёркивание в конце строки (табуляция в конце абзаца) Word и LibreOffice не рисуют - последним знаком ставится «_»
        if ('«ПРИЧИНА»' in text or '«МЕСТО»' in text) and text.replace('_', '').strip() in ('«ПРИЧИНА»', '«МЕСТО»'):
            rs = p.findall(W + 'r')
            if rs and rs[-1].find(W + 'tab') is not None:
                q = copy.deepcopy(rs[-1])
                for t in q.findall(W + 'tab'): q.remove(t)
                etree.SubElement(q, W + 't').text = '_'
                rs[-1].addnext(q)
                n += 1
            for t in p.iter(W + 'tab'):
                if t.get(W + 'val') == 'right' and t.get(W + 'pos') == '7636': t.set(W + 'pos', '7520'); n += 1
        # акт приемки, п. 3: дополнение после линии
        if 'прибор учета электроэнергии, отходящая' in text and '«ЛИНИЯ»' in text and '«АКТДОП»' not in text:
            for r in p.findall(W + 'r'):
                if ''.join(t.text or '' for t in r.iter(W + 't')) == '«ЛИНИЯ»':
                    q = copy.deepcopy(r); q.find(W + 't').text = '«АКТДОП»'; r.addnext(q); n += 1
                    break
    return n

def fix(xml):
    root = etree.fromstring(xml)
    n = naryad_fit(root) + tables_fit(root) + user_26(root)
    # 4) формула-рисунок в ячейке таблицы, плавающая с отрицательным отступом (в потерях «Wал=Wпу» налезала на схему) - в строку ячейки
    for sh in root.iter(V + 'shape'):
        st = sh.get('style') or ''
        if 'position:absolute' in st and 'margin-top:-' in st and any(a.tag == W + 'tc' for a in sh.iterancestors()):
            keep = [p for p in st.split(';') if p.split(':')[0] in ('width', 'height')]
            sh.set('style', ';'.join(keep))
            for w in [c for c in sh if c.tag.endswith('}wrap')]: sh.remove(w)
            for r in sh.iterancestors(W + 'r'):
                for hl in r.iter(W + 'highlight'): hl.getparent().remove(hl)
                break
            n += 1
    for p in root.iter(W + 'p'):
        ts = runs(p)
        text = ''.join(t.text or '' for t in ts)
        for m in list(re.finditer('минскэнерго', text, re.I)):
            if m.group(0) != 'МИНСКЭНЕРГО':
                replace(ts, m.start(), m.end(), 'МИНСКЭНЕРГО'); n += 1
        text = ''.join(t.text or '' for t in ts)
        k = text.find('«ПУГОД» г.')
        if k >= 0:                                              # 5) год поверки: нет года - нет и «г.»
            replace(ts, k, k + len('«ПУГОД» г.'), '«ПУГОДГ»'); n += 1
            text = ''.join(t.text or '' for t in ts)
        k = text.find('электрическая энергия используется для электроснабжения базовой станции')
        if k >= 0 and 'характер деятельности' in text:          # акт приемки, п. 4 - поле «ДЕЯТЕЛЬНОСТЬ»
            replace(ts, k, k + len('электрическая энергия используется для электроснабжения базовой станции'), '«ДЕЯТЕЛЬНОСТЬ»'); n += 1; continue
        m = re.match(r'\s*(Начальник ГРП\s+ОПСУ)', text)
        if m: replace(ts, m.start(1), m.end(1), '«НАЧДОЛЖНДС»'); n += 1; continue
        m = re.match(r'\s*(Начальник ГРП)', text)
        if m: replace(ts, m.start(1), m.end(1), '«НАЧДОЛЖН»'); n += 1
    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True), n

def footer_to_body(items):
    # 6) подписи сторон с «МП» лежали в нижнем колонтитуле (доп. соглашение) - переносятся в конец текста, колонтитул отключается
    names = [i.filename for i, d in items]
    if 'word/document.xml' not in names or 'word/_rels/document.xml.rels' not in names: return items, 0
    doc = etree.fromstring(dict((i.filename, d) for i, d in items)['word/document.xml'])
    rels = etree.fromstring(dict((i.filename, d) for i, d in items)['word/_rels/document.xml.rels'])
    body = doc.find(W + 'body'); sect = body.find(W + 'sectPr')
    if sect is None: return items, 0
    R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    n = 0
    for ref in list(sect.findall(W + 'footerReference')):
        target = [r.get('Target') for r in rels if r.get('Id') == ref.get(R)]
        if not target: continue
        ftr = etree.fromstring(dict((i.filename, d) for i, d in items).get('word/' + target[0], b'<x/>'))
        if 'МП' not in ''.join(t.text or '' for t in ftr.iter(W + 't')): continue
        gap = etree.SubElement(body, W + 'p'); body.remove(gap); sect.addprevious(gap)
        for el in ftr:
            if el.tag == W + 'tbl':
                c = copy.deepcopy(el)
                for tr in c.iter(W + 'tr'):
                    pr = tr.find(W + 'trPr')
                    if pr is None: pr = etree.SubElement(tr, W + 'trPr'); tr.remove(pr); tr.insert(0, pr)
                    etree.SubElement(pr, W + 'cantSplit')
                sect.addprevious(c)
        sect.remove(ref)
        n += 1
    if not n: return items, 0
    data = etree.tostring(doc, xml_declaration=True, encoding='UTF-8', standalone=True)
    return [(i, data if i.filename == 'word/document.xml' else d) for i, d in items], n

FILE = ''
for f in sys.argv[1:]:
    FILE = f
    zin = zipfile.ZipFile(f)
    items = [(i, zin.read(i.filename)) for i in zin.infolist()]
    zin.close()
    items, total = footer_to_body(items)
    out = []
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
