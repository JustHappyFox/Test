# Вымышленная юр база - таблица FTU (DBF, кодировка 866, как у DOS-баз): дом «Макетино, Полевая 34» (есть в gen_byt.py) и другой дом
# python3 gen_ftu.py out.dbf
import sys, struct, datetime
F = [('KP','N',5),('NP','C',40),('NAB','N',3),('NAB_NAIM','C',40),('NAIM','C',30),('NST','C',20),('TIPSCH','C',20),('VID','C',14),
     ('FAZA','N',1),('GODEN','C',6),('AMPER','C',8),('VOLT','N',3),('SHKALA','C',2)]
rows = []
for kv in (5, 12, 40):
    rows.append([11118, 'ООО "Тестстрой", г.Минск', 220, 'Нераспр.кв. Макетино, ул.Полевая 34', '%d квартира' % kv, '7000%04d' % kv, 'СЕ 318 BY', 'Электронный', 3, '2025-3', '5-60', 400, '6'])
rows.append([11118, 'ООО "Тестстрой", г.Минск', 221, 'Офис ул.Лесная 7', 'ЩУЭ офиса', '71110001', 'СЕ 208 BY', 'Электронный', 1, '2024-1', '5-60', 230, '6'])
rows.append([22222, 'ЧУП "Пример"', 1, 'Магазин Тестовка, ул.Садовая 3', 'ЩУ магазина', '72220001', 'СЕ 318 BY', 'Электронный', 3, '2023-2', '10-100', 400, '6'])
reclen = 1 + sum(l for _, _, l in F); hl = 32 + 32 * len(F) + 1
h = bytearray(32); h[0] = 3; d = datetime.date.today(); h[1:4] = bytes([d.year % 100, d.month, d.day])
h[4:8] = struct.pack('<I', len(rows)); h[8:10] = struct.pack('<H', hl); h[10:12] = struct.pack('<H', reclen); h[29] = 0x26
out = bytearray(h); off = 1
for n, t, l in F:
    fd = bytearray(32); fd[0:len(n)] = n.encode(); fd[11] = ord(t); fd[12:16] = struct.pack('<I', off); fd[16] = l; out += fd; off += l
out += b'\r'
for r in rows:
    rec = bytearray(b' ')
    for (n, t, l), v in zip(F, r):
        s = str(v); rec += (s.rjust(l) if t == 'N' else s.ljust(l)).encode('cp866')[:l]
    out += rec
out += b'\x1a'
open(sys.argv[1], 'wb').write(out); print(sys.argv[1], len(rows))
