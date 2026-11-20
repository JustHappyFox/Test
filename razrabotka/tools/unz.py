import zipfile, sys, os
src, dst = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(src) as z:
    for i in z.infolist():
        n = i.filename
        if not (i.flag_bits & 0x800):
            n = n.encode('cp437').decode('cp866')
        p = os.path.normpath(os.path.join(dst, n))
        assert p.startswith(os.path.abspath(dst))
        if n.endswith('/'): os.makedirs(p, exist_ok=True); continue
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, 'wb') as f: f.write(z.read(i))
        t = i.date_time; import time; ts = time.mktime(t + (0,0,-1)); os.utime(p, (ts, ts))
