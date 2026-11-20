# Архивы для пользователя.
#   python3 pack.py program dogovor        - пересобрать dogovor.zip в корне репозитория (файлы папки без журнала)
#   python3 pack.py update DD.MM file...   - obnovlenie-DD.MM.zip в scratchpad: только изменённые файлы с путями по папкам
#   python3 pack.py all                    - programmy.zip (menu, prosmotr-bazy, dogovor, foto) в scratchpad
import sys, os, zipfile
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
SP = '/tmp/claude-0/-home-user-Test/ecddb7b9-ba2d-55d0-b8b6-3d6ae08ebda9/scratchpad'
SKIP = {'TREBOVANIYA.md'}
def files(folder):
    d = os.path.join(ROOT, folder)
    return sorted(f for f in os.listdir(d) if os.path.isfile(os.path.join(d, f)) and f not in SKIP)
def write(zpath, pairs):
    with zipfile.ZipFile(zpath, 'w', zipfile.ZIP_DEFLATED) as z:
        for src, name in pairs: z.write(src, name)
    print(zpath, len(pairs), 'файлов')
cmd = sys.argv[1]
if cmd == 'program':
    f = sys.argv[2]
    write(os.path.join(ROOT, f + '.zip'), [(os.path.join(ROOT, f, n), n) for n in files(f)])
elif cmd == 'update':
    write(os.path.join(SP, 'obnovlenie-%s.zip' % sys.argv[2]), [(os.path.join(ROOT, p), p) for p in sys.argv[3:]])
elif cmd == 'all':
    write(os.path.join(SP, 'programmy.zip'), [(os.path.join(ROOT, f, n), f + '/' + n) for f in ('menu', 'prosmotr-bazy', 'dogovor', 'foto') for n in files(f)])
