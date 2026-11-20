# Вставить общий блок выгрузок (razrabotka/shared/vygruzka.cs) в программы: menu, prosmotr-bazy, dogovor.
# Блок уже есть - заменяется между метками; нет - ставится после строк using. Добавляет using System.Xml и ссылку на System.Xml.
# python3 inject.py
import os, sys
sys.path.insert(0, os.path.dirname(__file__))
from pyedit import Doc
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
BEGIN = '// ==== ОБЩИЙ БЛОК ВЫГРУЗОК'
END = '// ==== КОНЕЦ ОБЩЕГО БЛОКА ВЫГРУЗОК ===='
block = open(os.path.join(ROOT, 'razrabotka', 'shared', 'vygruzka.cs'), encoding='utf-8-sig').read().replace('\r\n', '\n').strip('\n')
assert block.startswith(BEGIN) and block.endswith(END)
assert "\n'@" not in block, "в блоке не должно быть строки, начинающейся с '@ (конец here-string PowerShell)"
for prog in ('menu/menu.ps1', 'prosmotr-bazy/prosmotr.ps1', 'dogovor/dogovor.ps1'):
    d = Doc(os.path.join(ROOT, prog))
    t = d.t
    start = t.find("$code = @'\n")
    assert start >= 0
    if BEGIN in t:
        a = t.index(BEGIN); b = t.index(END) + len(END)
        t = t[:a] + block + t[b:]
    else:
        lines = t[start:].split('\n')
        i = 1
        while lines[i].startswith('using ') or lines[i].strip() == '': i += 1
        pos = start + len('\n'.join(lines[:i])) + 1
        t = t[:pos] + block + '\n\n' + t[pos:]
    if 'using System.Xml;' not in t[start:start + 3000]:
        t = t.replace("$code = @'\n", "$code = @'\nusing System.Xml;\n", 1)
    if '[System.Xml.XmlDocument].Assembly.Location' not in t:
        t = t.replace('[System.Drawing.Point].Assembly.Location)', '[System.Drawing.Point].Assembly.Location, [System.Xml.XmlDocument].Assembly.Location)', 1)
    d.t = t
    d.save()
    print(prog, 'ок')

# блок формы БС (только «Договор»): razrabotka/src/bs_form.cs, перед классом SealRow
B2, E2 = '// ==== БЛОК ФОРМЫ БС', '// ==== КОНЕЦ БЛОКА ФОРМЫ БС ===='
bs = open(os.path.join(ROOT, 'razrabotka', 'src', 'bs_form.cs'), encoding='utf-8-sig').read().replace('\r\n', '\n').strip('\n')
assert "\n'@" not in bs
d = Doc(os.path.join(ROOT, 'dogovor/dogovor.ps1'))
if B2 in d.t:
    a = d.t.index(B2); b = d.t.index(E2) + len(E2)
    d.t = d.t[:a] + bs + d.t[b:]
else:
    d.rep('public class SealRow {', bs + '\n\npublic class SealRow {')
d.save()
print('dogovor/dogovor.ps1 БС ок')
