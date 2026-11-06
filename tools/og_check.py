"""Отличает ли страница поста t.me/канал/N модуль-розыгрыш от других «заглушек» (идея владельца 06.11).

У модуля-розыгрыша текст пустой, поэтому мета-тег og:description страницы поста пустой. У других
заглушек (видео/фото с защитой, неподдерживаемые вложения) сайт прячет пост в ленте, но в
og:description кладёт его текст. Если пост недоступен виджету, в og:description — описание канала.

  --posts канал/N …   известные розыгрыши: у всех ли пусто
  --gap-file FILE     список кандидатов: каналы «второго уровня» (последнее число в строке — 0)
  --sample N          сколько таких каналов взять (с начала списка)

Итог: у скольких заглушек второго уровня есть текст (это точно не розыгрыш — Dolli их не открывает).
"""
import argparse
import asyncio
import html as htmllib
import re
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

import aiohttp

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import tg_web_preview as wp  # noqa: E402

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"}
OG = re.compile(r'<meta property="og:description" content="([^"]*)"')


STATUS: dict = {}


async def get(s, url):
    for i in range(5):
        try:
            async with s.get(url, headers=UA, timeout=aiohttp.ClientTimeout(total=30)) as r:
                STATUS[r.status] = STATUS.get(r.status, 0) + 1
                if r.status == 200:
                    return (await r.read()).decode("utf-8", errors="replace")
                if r.status == 429:
                    await asyncio.sleep(float(r.headers.get("Retry-After") or 5 * (i + 1)))
                    continue
                return ""
        except Exception as e:
            STATUS[type(e).__name__] = STATUS.get(type(e).__name__, 0) + 1
            await asyncio.sleep(2 * (i + 1))
    return ""


async def og(s, url):
    m = OG.search(await get(s, url))
    return htmllib.unescape(m.group(1)).strip() if m else None


async def classify(s, chan, num, chan_desc):
    d = await og(s, f"https://t.me/{chan}/{num}")
    if d is None:
        return "нет страницы", d  # страница не загрузилась или в ней нет og:description
    if not d:
        return "пусто", d
    if chan_desc is not None and d == chan_desc:
        return "описание канала", d
    return "текст", d


async def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--posts", nargs="*", default=[])
    ap.add_argument("--gap-file")
    ap.add_argument("--sample", type=int, default=300)
    ap.add_argument("--days", type=int, default=7)
    ap.add_argument("--parallel", type=int, default=2)
    a = ap.parse_args()
    async with aiohttp.ClientSession() as s:
        print("=== Известные розыгрыши ===")
        stat = {}
        for p in a.posts:
            chan, num = p.split("/")
            cd = await og(s, f"https://t.me/{chan}")
            kind, d = await classify(s, chan, num, cd)
            stat[kind] = stat.get(kind, 0) + 1
            print(f"  {p}: {kind}" + (f" — {d[:80]!r}" if d else ""))
        print("Итог по розыгрышам:", stat)

        if not a.gap_file:
            return
        names = []
        for ln in open(a.gap_file, encoding="utf-8"):
            parts = ln.split()
            if parts and parts[-1] == "0":
                names.append(parts[0].lstrip("@"))
        names = names[:a.sample]
        w = wp.WebPreview(parallel=8)
        since = datetime.now(timezone.utc) - timedelta(days=a.days)
        sem = asyncio.Semaphore(a.parallel)
        chan_stat = {"все с текстом": 0, "есть без текста": 0, "нет заглушек сейчас": 0, "нет страницы": 0}
        stub_stat = {}
        keep = []

        async def one(name):
            async with sem:
                try:
                    ids = await w.candidates(name, since, max_pages=5, strict=True, gaps=True)
                except Exception:
                    ids = None
                if ids is None:
                    chan_stat["нет страницы"] += 1
                    return
                if not ids:
                    chan_stat["нет заглушек сейчас"] += 1
                    return
                cd = await og(s, f"https://t.me/{name}")
                kinds = []
                for i in ids[:10]:
                    kind, d = await classify(s, name, i, cd)
                    kinds.append(kind)
                    stub_stat[kind] = stub_stat.get(kind, 0) + 1
                if all(k == "текст" for k in kinds):
                    chan_stat["все с текстом"] += 1
                else:
                    chan_stat["есть без текста"] += 1
                    keep.append(f"@{name} {list(zip(ids[:10], kinds))}")

        await asyncio.gather(*(one(n) for n in names))
        await w.close()
        print(f"\n=== Каналы второго уровня: {len(names)} ===")
        print("Заглушки:", stub_stat)
        print("Ответы сайта:", STATUS)
        print("Каналы:", chan_stat)
        print("Остаются для Dolli (есть заглушка без текста):")
        for k in keep:
            print("  ", k)


if __name__ == "__main__":
    asyncio.run(main())
