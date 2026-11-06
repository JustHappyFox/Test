"""Полный код поста на сайте t.me — чтобы сравнить модуль-розыгрыш с другими «заглушками».

Для каждого поста канал/номер выкачиваем три вида страницы и печатаем их целиком:
  1. лента   https://t.me/s/канал?before=N+1   — блок этого поста (так читает наш поиск);
  2. виджет  https://t.me/канал/N?embed=1&mode=tme — отдельный пост, как его встраивают на сайты;
  3. пост    https://t.me/канал/N                — страница поста: мета-теги og:/twitter: из <head>.

--stubs канал …  — найти в ленте канала посты-заглушки (без текста и медиа) и выкачать их так же:
это каналы, где Dolli уже открыла заглушки и розыгрыша там не оказалось (для сравнения).

Пример: python tools/web_dump.py --posts rzavayuui/43 korchaginnn/28886 --stubs samara dva_majors
"""
import argparse
import asyncio
import re
import sys
from pathlib import Path

import aiohttp

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tg_web_preview as wp  # noqa: E402

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"}
WRAP = '<div class="tgme_widget_message_wrap'


async def get(s: aiohttp.ClientSession, url: str) -> str:
    for i in range(3):
        try:
            async with s.get(url, headers=UA, timeout=aiohttp.ClientTimeout(total=30)) as r:
                if r.status == 200:
                    return await r.text()
                print(f"  {url}: HTTP {r.status}")
        except Exception as e:  # сеть — повторить
            print(f"  {url}: {e}")
        await asyncio.sleep(2 * (i + 1))
    return ""


def block_of(html: str, post: str) -> str:
    for part in html.split(WRAP)[1:]:
        if f'data-post="{post}"' in part:
            return WRAP + part
    return ""


def head_meta(html: str) -> str:
    head = html.split("</head>", 1)[0]
    return "\n".join(re.findall(r"<meta [^>]*>", head))


def body_of(html: str) -> str:
    m = re.search(r"<body[^>]*>(.*)</body>", html, re.S)
    return (m.group(1) if m else html).strip()


async def dump(s: aiohttp.ClientSession, post: str, label: str) -> None:
    chan, num = post.split("/")
    print(f"\n{'=' * 100}\n### {label}: {post}\n{'=' * 100}")
    feed = await get(s, f"https://t.me/s/{chan}?before={int(num) + 1}")
    print(f"--- 1. лента (блок поста и предыдущий):")
    print(block_of(feed, f"{chan}/{int(num) - 1}") or "  (предыдущего на странице нет)")
    print(block_of(feed, post) or "  (поста на странице нет)")
    embed = await get(s, f"https://t.me/{chan}/{num}?embed=1&mode=tme")
    print(f"--- 2. виджет ?embed=1:")
    print(body_of(embed))
    page = await get(s, f"https://t.me/{chan}/{num}")
    print(f"--- 3. страница поста, мета-теги:")
    print(head_meta(page))


async def stub_posts(s: aiohttp.ClientSession, chan: str, limit: int) -> list[str]:
    html = await get(s, f"https://t.me/s/{chan}")
    posts = wp.parse_page(html) or []
    return [f"{chan}/{p.id}" for p in posts if p.stub and not p.forwarded][-limit:]


async def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--posts", nargs="*", default=[], help="известные розыгрыши: канал/номер")
    ap.add_argument("--stubs", nargs="*", default=[], help="каналы, где заглушки — не розыгрыши")
    ap.add_argument("--per-channel", type=int, default=2)
    a = ap.parse_args()
    async with aiohttp.ClientSession() as s:
        for p in a.posts:
            await dump(s, p, "РОЗЫГРЫШ")
        for ch in a.stubs:
            found = await stub_posts(s, ch, a.per_channel)
            if not found:
                print(f"\n### {ch}: заглушек в ленте сейчас нет")
            for p in found:
                await dump(s, p, "НЕ РОЗЫГРЫШ (заглушка)")


if __name__ == "__main__":
    asyncio.run(main())
