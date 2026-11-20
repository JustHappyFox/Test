# Текст docx с отметкой выделенных красным мест: [[...]]. Таблицы - построчно, ячейки через " | ".
import sys, zipfile, re
from xml.etree import ElementTree as ET
W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
def runs_text(p):
    out=[]; cur=False
    for r in p.iter(W+'r'):
        rp=r.find(W+'rPr'); red=False
        if rp is not None:
            h=rp.find(W+'highlight'); c=rp.find(W+'color')
            red=(h is not None and h.get(W+'val')=='red') or (c is not None and c.get(W+'val')=='FF0000')
        t=''.join((x.text or '') if x.tag==W+'t' else ('\t' if x.tag==W+'tab' else ('\n' if x.tag in (W+'br',W+'cr') else '')) for x in r)
        if not t: continue
        if red and not cur: out.append('[['); cur=True
        if not red and cur: out.append(']]'); cur=False
        out.append(t)
    if cur: out.append(']]')
    return ''.join(out)
def walk(el, depth=0):
    for ch in el:
        if ch.tag==W+'p':
            s=runs_text(ch)
            if s.strip(): print(s)
        elif ch.tag==W+'tbl':
            print('<<ТАБЛИЦА')
            for tr in ch.iter(W+'tr'):
                cells=[' / '.join(x for x in (runs_text(p) for p in tc.iter(W+'p')) if x.strip()) for tc in tr.findall(W+'tc')]
                print('  | ' + ' | '.join(cells))
            print('ТАБЛИЦА>>')
        elif ch.tag in (W+'sdt', W+'sdtContent', W+'body', W+'txbxContent'):
            walk(ch, depth+1)
root=ET.fromstring(zipfile.ZipFile(sys.argv[1]).read('word/document.xml'))
walk(root.find(W+'body'))
