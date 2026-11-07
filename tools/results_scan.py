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


OWN_CONTENT = re.compile(r"js-message_text|tgme_widget_message_photo|tgme_widget_message_video|media_supported_cont|"
                         r"tgme_widget_message_document|tgme_widget_message_poll|tgme_widget_message_sticker")


def results_in(html: str):
    """Итоги розыгрыша: пустая заглушка (без своего текста, фото, видео), которая отвечает на пост с пустой
    цитатой. 07.11 владелец проверил: ответы-посты с фото/видео и текстом (profitbase/2190, ecoplatform/2742)
    — не итоги; пустые заглушки-ответы (betinjabet/730, rzavayuui/39, energyplusmedia/2122) — итоги."""
    out = []
    for blk in html.split('class="tgme_widget_message_wrap')[1:]:
        if "tgme_widget_message_reply" not in blk:
            continue
        r, rt, p, d = REPLY.search(blk), REPLY_TEXT.search(blk), POST.search(blk), DATE.search(blk)
        if not (r and rt and p) or re.sub(r"<[^>]+>|\s", "", rt.group(1)):
            continue  # ответ на пост с текстом — не итоги розыгрыша
        own = re.sub(r'<a class="tgme_widget_message_reply.*?</a>', "", blk, count=1, flags=re.S)
        own = own.split("tgme_widget_message_footer", 1)[0]
        if "message_media_not_supported_wrap" not in own or OWN_CONTENT.search(own):
            continue  # у поста свой текст/фото/видео — обычный ответ, не итоги
        out.append((int(p.group(2)), int(r.group(2)), datetime.fromisoformat(d.group(1)) if d else None))
    return out


MIN_GIVEAWAY = timedelta(hours=1)


def target_is_module(html: str, post_id: int, results_date: datetime | None = None) -> bool:
    """Исходный пост (на который отвечают итоги) — тоже пустая заглушка без текста и медиа, сам ни на что
    не отвечает, и между ним и итогами прошёл хотя бы час. 07.11: у новостного @nikvesti пустые заглушки
    отвечали друг на друга цепочкой через минуту (84912 → 84913 → 84915) — это не розыгрыши."""
    for blk in (html or "").split('class="tgme_widget_message_wrap')[1:]:
        p = POST.search(blk)
        if p and int(p.group(2)) == post_id:
            body = blk.split("tgme_widget_message_footer", 1)[0]
            if "message_media_not_supported_wrap" not in body or OWN_CONTENT.search(body):
                return False
            if "tgme_widget_message_reply" in body:
                return False
            d = DATE.search(blk)
            if results_date and d and results_date - datetime.fromisoformat(d.group(1)) < MIN_GIVEAWAY:
                return False
            return True
    return False


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
            for h in results_in(html):
                if h[2] is not None and h[2] < since:
                    continue
                target = await w.fetch(f"{w.BASE}{name}?before={h[1] + 1}")
                if target_is_module(target, h[1], h[2]):
                    hits.append(h)
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
