# Вымышленная выгрузка быта в .dbf - как TEMP.DBF старой программы базы (DOS, 866): из .xlsx (gen_byt.py).
# python3 gen_dbf.py byt.xlsx out.dbf [метка кодировки: 0 - нет (по буквам), 0x26 - 866, 0xC9 - 1251]
# Даты - поле D, числа - N, остальное - C; одна удалённая запись (копия первой с «*») - должна пропускаться.
import sys, struct, datetime, openpyxl
src, out = sys.argv[1], sys.argv[2]
ldid = int(sys.argv[3], 0) if len(sys.argv) > 3 else 0
enc = 'cp1251' if ldid == 0xC9 else 'cp866'
ws = openpyxl.load_workbook(src, read_only=True).active
rows = [list(r) for r in ws.iter_rows(values_only=True)]
head, data = rows[0], rows[1:]
def kind(c):
    vals = [r[c] for r in data if c < len(r) and r[c] not in (None, '')]
    if vals and all(isinstance(v, datetime.datetime) for v in vals): return 'D', 8, 0
    if vals and all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in vals):
        dec = 2 if any(isinstance(v, float) and v != int(v) for v in vals) else 0
        return 'N', 15, dec
    return 'C', max([len(str(v).encode(enc, 'replace')) for v in vals] + [1]), 0
fields = [(str(h),) + kind(c) for c, h in enumerate(head) if h]
def cell(v, t, ln, dec):
    if v in (None, ''): return b' ' * ln
    if t == 'D': return v.strftime('%Y%m%d').encode()
    if t == 'N': return (('%.' + str(dec) + 'f') % float(v)).rjust(ln).encode()
    return str(v).encode(enc, 'replace')[:ln].ljust(ln)
recs = []
for r in data:
    b = b' ' + b''.join(cell(r[i] if i < len(r) else None, t, ln, dec) for i, (n, t, ln, dec) in enumerate(fields))
    recs.append(b)
recs.insert(1, b'*' + recs[0][1:])          # удалённая запись
reclen = 1 + sum(f[2] for f in fields)
hlen = 32 + 32 * len(fields) + 1
d = datetime.date.today()
h = struct.pack('<BBBBIHH', 0x03, d.year % 100, d.month, d.day, len(recs), hlen, reclen) + b'\0' * 17 + bytes([ldid]) + b'\0\0'
for n, t, ln, dec in fields:
    nm = n.encode(enc)[:10].ljust(11, b'\0')
    h += nm + t.encode() + b'\0' * 4 + bytes([ln, dec]) + b'\0' * 14
h += b'\r'
open(out, 'wb').write(h + b''.join(recs) + b'\x1a')
print(out, len(fields), 'полей', len(recs), 'записей (1 удалённая)', 'метка', hex(ldid), enc)
