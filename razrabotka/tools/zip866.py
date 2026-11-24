# Архив с русскими именами файлов для Windows 7: встроенная распаковка Windows не понимает имена в UTF-8 (выйдут
# «кракозябры») - имена пишутся в cp866, как делает сама Windows. python3 zip866.py архив.zip папка (содержимое папки - в корень архива)
import sys, os, zipfile
class Info866(zipfile.ZipInfo):
    def _encodeFilenameFlags(self):
        try: return self.filename.encode('cp866'), self.flag_bits & ~0x800
        except UnicodeEncodeError: return self.filename.encode('utf-8'), self.flag_bits | 0x800
out, src = sys.argv[1], sys.argv[2]
with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
    for root, dirs, files in os.walk(src):
        dirs.sort()
        for f in sorted(files):
            p = os.path.join(root, f)
            info = Info866.from_file(p, os.path.relpath(p, src))
            info.compress_type = zipfile.ZIP_DEFLATED
            with open(p, 'rb') as fh: z.writestr(info, fh.read())
print(out)
