"""Каналы, где розыгрыш ЗАКОНЧИЛСЯ за последние N дней (идея владельца 06.11: «узнать, какие завершались
за 7 дней, и проверить те группы»). Без аккаунта, по веб-странице t.me/s.

Итоги розыгрыша Telegram публикует в канале сам (messageMediaGiveawayResults). На сайте это заглушка,
которая ОТВЕЧАЕТ на исходный пост-розыгрыш, и текст в ответе пустой (у модуля текста нет):
  <a class="tgme_widget_message_reply" href="https://t.me/канал/23"> … js-message_reply_text"></div>
Проверено на rzavayuui/39 (→ 23), energyplusmedia/2122, 2088.
"""
import argparse
import asyncio
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tg_web_preview as wp  # noqa: E402

REPLY = re.compile(r'class="tgme_widget_message_reply[^"]*" href="https?://t\.me/([A-Za-z0-9_]+)/(\d+)"')
REPLY_TEXT = re.compile(r'js-message_reply_text"[^>]*>(.*?)</div>', re.S)
POST = re.compile(r'data-post="([^"/]+)/(\d+)"')
DATE = re.compile(r'<time datetime="([^"]+)"')


def results_in(html: str):
    out = []
    for blk in html.split('class="tgme_widget_message_wrap')[1:]:
        if wp.STUB_CLASS not in blk or "tgme_widget_message_reply" not in blk:
            continue
        r, rt, p, d = REPLY.search(blk), REPLY_TEXT.search(blk), POST.search(blk), DATE.search(blk)
        if not (r and rt and p) or re.sub(r"<[^>]+>|\s", "", rt.group(1)):
            continue  # ответ на пост с текстом — не итоги розыгрыша
        own = blk.split("</a>", 1)  # собственного текста у поста тоже нет (кроме цитаты ответа)
        out.append((int(p.group(2)), int(r.group(2)), datetime.fromisoformat(d.group(1)) if d else None))
    return out


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("names", nargs="+", help="файлы со списками @username (берутся по очереди)")
    ap.add_argument("--days", type=int, default=7)
    ap.add_argument("--pages", type=int, default=3)
    ap.add_argument("--stop-after", type=int, default=10)
    ap.add_argument("--parallel", type=int, default=8)
    a = ap.parse_args()
    names = []
    for f in a.names:
        for ln in Path(f).read_text(encoding="utf-8").splitlines():
            n = ln.split("#", 1)[0].strip().lstrip("@")
            if n and n not in names:
                names.append(n)
    since = datetime.now(timezone.utc) - timedelta(days=a.days)
    w = wp.WebPreview(parallel=a.parallel, max_limited=10 ** 9)
    found, checked, t0 = [], 0, time.time()

    async def one(name):
        nonlocal checked
        url = w.BASE + name
        hits = []
        for _ in range(a.pages):
            html = await w.fetch(url)
            posts = wp.parse_page(html) if html else None
            if not posts:
                break
            hits += [h for h in results_in(html) if h[2] is None or h[2] >= since]
            oldest = min(posts, key=lambda p: p.id)
            if (oldest.date and oldest.date < since) or oldest.id <= 1:
                break
            url = f"{w.BASE}{name}?before={oldest.id}"
        checked += 1
        if hits:
            found.append((name, hits))
            for pid, rid, d in hits:
                print(f"  НАЙДЕН https://t.me/{name}/{pid} — итоги {d:%Y-%m-%d}, розыгрыш был https://t.me/{name}/{rid}",
                      flush=True)

    for i in range(0, len(names), 100):
        await asyncio.gather(*(one(n) for n in names[i:i + 100]))
        print(f"  проверено {checked}/{len(names)}, каналов с итогами за {a.days} дн.: {len(found)}, "
              f"страниц {w.pages}, 429: {w.limited}, {time.time() - t0:.0f} с", flush=True)
        if len(found) >= a.stop_after:
            break
    await w.close()
    print(f"Итог: каналов {checked}, с итогами розыгрыша за {a.days} дн.: {len(found)}")


if __name__ == "__main__":
    asyncio.run(main())
