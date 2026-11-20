# Правка файлов программ с сохранением BOM и CRLF (Python в текстовом режиме теряет CRLF!).
# from pyedit import Doc; d = Doc('dogovor/dogovor.ps1'); d.rep(old, new); d.save()
# old/new пишутся с обычными \n - перевод в CRLF при сохранении.
class Doc:
    def __init__(self, path):
        self.path = path
        raw = open(path, 'rb').read()
        self.bom = raw.startswith(b'\xef\xbb\xbf')
        self.crlf = b'\r\n' in raw
        self.t = raw.decode('utf-8-sig').replace('\r\n', '\n')
    def rep(self, old, new, count=1):
        n = self.t.count(old)
        assert n == count, 'нашлось %d раз (нужно %d): %r' % (n, count, old[:80])
        self.t = self.t.replace(old, new)
    def save(self):
        t = self.t.replace('\n', '\r\n') if self.crlf else self.t
        open(self.path, 'wb').write((b'\xef\xbb\xbf' if self.bom else b'') + t.encode('utf-8'))
