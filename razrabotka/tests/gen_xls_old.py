# Вымышленная выгрузка в старых видах .xls - из .xlsx (gen_byt.py): Excel 5.0/95 (BIFF5 в OLE, 1251), Excel 2.1 как пишет FoxPro
# (BIFF2 без OLE, 866, без записи кодировки), XML-таблица Excel 2003; и .xlsx из двух листов (для перевода в .xls из двух листов).
# python3 gen_xls_old.py byt.xlsx папка
import sys, os, struct, datetime, openpyxl
src, out = sys.argv[1], sys.argv[2]
ws = openpyxl.load_workbook(src, read_only=True).active
rows = [list(r) for r in ws.iter_rows(values_only=True)]
def serial(v):
    if isinstance(v, datetime.datetime): return (v - datetime.datetime(1899, 12, 30)).days + 0.0
    return v
def rec(t, data): return struct.pack('<HH', t, len(data)) + data

def biff5_stream():
    cells = b''
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            v = serial(v)
            if v is None or v == '': continue
            if isinstance(v, (int, float)): cells += rec(0x0203, struct.pack('<HHHd', r, c, 15, float(v)))
            else:
                b = str(v).encode('cp1251')[:255]
                cells += rec(0x0204, struct.pack('<HHHH', r, c, 15, len(b)) + b)
    glob_wo_bs = rec(0x0809, struct.pack('<HHHH', 0x0500, 0x0005, 0, 0)) + rec(0x0042, struct.pack('<H', 1251))
    name = 'TEMP'.encode('cp1251')
    bs_len = 4 + 4 + 2 + 1 + len(name)
    eof = rec(0x000A, b'')
    sheet_off = len(glob_wo_bs) + bs_len + len(eof)
    glob = glob_wo_bs + rec(0x0085, struct.pack('<IBBB', sheet_off, 0, 0, len(name)) + name) + eof
    sheet = rec(0x0809, struct.pack('<HHHH', 0x0500, 0x0010, 0, 0)) + cells + eof
    return glob + sheet

def ole(stream, name):
    ss = 512
    data = stream + b'\0' * (-len(stream) % ss)
    if len(data) < 4096: data += b'\0' * (4096 - len(data))
    n = len(data) // ss
    # сектора: [поток n][каталог 1][FAT k]
    k = 1
    while (n + 1 + k) > k * 128: k += 1
    fat = []
    for i in range(n): fat.append(i + 1 if i + 1 < n else 0xFFFFFFFE)
    fat.append(0xFFFFFFFE)                       # каталог
    fat += [0xFFFFFFFD] * k                      # сами FAT
    fat += [0xFFFFFFFF] * (k * 128 - len(fat))
    def entry(nm, typ, start, size, child=0xFFFFFFFF):
        e = nm.encode('utf-16-le') + b'\0\0'
        e = e + b'\0' * (64 - len(e))
        e += struct.pack('<HBB', len(nm) * 2 + 2, typ, 1) + struct.pack('<III', 0xFFFFFFFF, 0xFFFFFFFF, child)
        e += b'\0' * 16 + b'\0' * 4 + b'\0' * 16 + struct.pack('<III', start, size, 0)
        return e
    d = entry('Root Entry', 5, 0xFFFFFFFE, 0, 1) + entry(name, 2, 0, len(stream))
    d += b'\0' * (ss - len(d))
    h = b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1' + b'\0' * 16 + struct.pack('<HHHH', 0x3E, 3, 0xFFFE, 9) + struct.pack('<H', 6) + b'\0' * 6
    h += struct.pack('<IIIIIIIII', 0, k, n, 0, 4096, 0xFFFFFFFE, 0, 0xFFFFFFFE, 0)
    difat = [n + 1 + i for i in range(k)] + [0xFFFFFFFF] * (109 - k)
    h += struct.pack('<109I', *difat)
    return h + data + d + struct.pack('<%dI' % len(fat), *fat)

def biff2():
    out = rec(0x0009, struct.pack('<HH', 0x0002, 0x0010))
    for r, row in enumerate(rows):
        for c, v in enumerate(row):
            v = serial(v)
            if v is None or v == '': continue
            if isinstance(v, (int, float)): out += rec(0x0003, struct.pack('<HH', r, c) + b'\0\0\0' + struct.pack('<d', float(v)))
            else:
                b = str(v).encode('cp866')[:255]
                out += rec(0x0004, struct.pack('<HH', r, c) + b'\0\0\0' + struct.pack('<B', len(b)) + b)
    return out + rec(0x000A, b'')

def xml2003():
    from xml.sax.saxutils import escape
    o = ['<?xml version="1.0"?>\n<?mso-application progid="Excel.Sheet"?>\n<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet" '
         'xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">\n<Worksheet ss:Name="TEMP"><Table>']
    for row in rows:
        o.append('<Row>')
        skip = False
        for c, v in enumerate(row):
            if v is None or v == '': skip = True; continue
            idx = ' ss:Index="%d"' % (c + 1) if skip else ''
            skip = False
            if isinstance(v, datetime.datetime): o.append('<Cell%s><Data ss:Type="DateTime">%s.000</Data></Cell>' % (idx, v.strftime('%Y-%m-%dT%H:%M:%S')))
            elif isinstance(v, (int, float)): o.append('<Cell%s><Data ss:Type="Number">%s</Data></Cell>' % (idx, v))
            else: o.append('<Cell%s><Data ss:Type="String">%s</Data></Cell>' % (idx, escape(str(v))))
        o.append('</Row>\n')
    o.append('</Table></Worksheet></Workbook>\n')
    return ''.join(o).encode('utf-8')

open(os.path.join(out, 'b5.xls'), 'wb').write(ole(biff5_stream(), 'Book'))
open(os.path.join(out, 'b2_866.xls'), 'wb').write(biff2())
open(os.path.join(out, 'x2003.xls'), 'wb').write(xml2003())
wb = openpyxl.Workbook(); a = wb.active; a.title = 'TEMP'; b = wb.create_sheet('TEMP2')
for i, r in enumerate(rows):
    if i == 0 or i <= 600: a.append(r)
    if i == 0 or i > 600: b.append(r)
for s in (a, b):
    for c in s['V'][1:]: c.number_format = 'dd.mm.yyyy'
wb.save(os.path.join(out, 'two.xlsx'))
print('ok', len(rows))
