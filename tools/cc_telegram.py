#!/usr/bin/env python3
"""
Список публичных @username из индекса Common Crawl (общедоступный архив интернета).

Идея из статьи «Как я сделал глобальный семантический поиск для Telegram» (habr.com/ru/publications/1015876):
в архиве Common Crawl есть миллионы ссылок на t.me. Сервер поиска по индексу (index.commoncrawl.org)
часто перегружен (504), поэтому идём напрямую по файлам индекса:

  1. cluster.idx — оглавление индекса: первая запись каждого блока (~3000 адресов) и где блок лежит.
     Записи отсортированы по «SURT»-ключу, у t.me он начинается с «me,t)/», значит все адреса t.me
     лежат подряд в нескольких блоках.
  2. Эти блоки скачиваем частичными запросами (HTTP Range) из data.commoncrawl.org, распаковываем
     и вытаскиваем @username из адресов вида t.me/<имя>, t.me/s/<имя>, t.me/<имя>/<пост>.

Пример:  python tools/cc_telegram.py --crawl CC-MAIN-2026-43 --out cc_usernames.txt --max-blocks 400
"""
from __future__ import annotations

import argparse
import gzip
import json
import random
import re
import sys
import time
import urllib.request

BASE = "https://data.commoncrawl.org/"
SURT_PREFIX = "me,t)/"
# пути t.me, которые не являются именами каналов
RESERVED = {
    "s", "c", "joinchat", "addstickers", "addemoji", "addtheme", "addlist", "share", "proxy", "socks",
    "iv", "login", "confirmphone", "setlanguage", "bg", "invoice", "giftcode", "boost", "contact",
    "m", "nft", "auction", "web", "k", "z", "a", "dl", "apps", "blog", "faq", "privacy", "tos",
    "addstyle", "call", "auth", "i", "v", "embed", "telegram", "botfather",
}
NAME_RE = re.compile(r"^[A-Za-z][A-Za-z0-9_]{3,31}$")


def get(url: str, rng: tuple[int, int] | None = None, tries: int = 6) -> bytes:
    headers = {"User-Agent": "giveaway-finder (Common Crawl index reader)"}
    if rng:
        headers["Range"] = f"bytes={rng[0]}-{rng[1]}"
    for attempt in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=headers), timeout=60) as r:
                return r.read()
        except Exception as e:  # 503/504/сеть — подождать и повторить
            if attempt == tries - 1:
                raise
            print(f"  повтор {url} ({e})", file=sys.stderr)
            time.sleep(5 * (attempt + 1))  # 503 Slow Down — сервер просит сбавить темп
    return b""


def tme_blocks(crawl: str) -> list[tuple[str, int, int]]:
    """Блоки индекса, где могут быть адреса t.me: (файл cdx, смещение, длина)."""
    url = f"{BASE}cc-index/collections/{crawl}/indexes/cluster.idx"
    req = urllib.request.Request(url, headers={"User-Agent": "giveaway-finder"})
    blocks: list[tuple[str, int, int]] = []
    prev = None
    with urllib.request.urlopen(req, timeout=120) as r:
        for raw in r:
            line = raw.decode("utf-8", "replace").rstrip("\n")
            key, _, rest = line.partition(" ")
            parts = rest.split("\t")
            if len(parts) < 4:
                continue
            entry = (parts[1], int(parts[2]), int(parts[3]))
            if key.startswith(SURT_PREFIX):
                if not blocks and prev is not None:
                    blocks.append(prev)  # первые t.me могут лежать в конце предыдущего блока
                blocks.append(entry)
            elif blocks:
                break  # блоки t.me кончились — дальше читать оглавление незачем
            prev = entry
    return blocks


def usernames_from_block(crawl: str, block: tuple[str, int, int]) -> set[str]:
    name, offset, length = block
    data = gzip.decompress(get(f"{BASE}cc-index/collections/{crawl}/indexes/{name}", (offset, offset + length - 1)))
    found: set[str] = set()
    for line in data.decode("utf-8", "replace").splitlines():
        if not line.startswith(SURT_PREFIX):
            continue
        try:
            url = json.loads(line.split(" ", 2)[2])["url"]
        except (IndexError, ValueError, KeyError):
            continue
        path = url.split("://", 1)[-1].split("?", 1)[0].split("#", 1)[0]
        segs = [s for s in path.split("/")[1:] if s]
        if segs and segs[0] == "s":
            segs = segs[1:]
        if not segs:
            continue
        cand = segs[0].lstrip("@")
        if cand.lower() not in RESERVED and NAME_RE.match(cand) and not cand.lower().endswith("bot"):
            found.add(cand)
    return found


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--crawl", nargs="+", default=["CC-MAIN-2026-43", "CC-MAIN-2026-39", "CC-MAIN-2026-34"],
                    help="снимки Common Crawl (список — index.commoncrawl.org/collinfo.json)")
    ap.add_argument("--out", required=True, help="куда записать @username, по одному в строке")
    ap.add_argument("--max-blocks", type=int, default=0, help="взять не больше N блоков (равномерно по всему t.me); 0 — все")
    args = ap.parse_args()

    t0 = time.time()
    names: set[str] = set()
    total = 0
    for crawl in args.crawl:
        try:
            blocks = tme_blocks(crawl)
        except Exception as e:
            print(f"{crawl}: оглавление не прочитано ({e}) — пропускаю", file=sys.stderr)
            continue
        if args.max_blocks and len(blocks) > args.max_blocks:
            step = len(blocks) / args.max_blocks
            blocks = [blocks[int(i * step)] for i in range(args.max_blocks)]
        before = len(names)
        for b in blocks:
            try:
                names |= usernames_from_block(crawl, b)
            except Exception as e:
                print(f"  {crawl}: блок {b[0]}@{b[1]} пропущен: {e}", file=sys.stderr)
            time.sleep(0.5)  # бережём сервер Common Crawl
        total += len(blocks)
        print(f"{crawl}: блоков с t.me {len(blocks)}, новых имён {len(names) - before}, всего {len(names)}")
    blocks = range(total)
    out = sorted(names, key=str.lower)
    random.Random(1).shuffle(out)  # вперемешку: при частичной проверке — равномерная выборка
    with open(args.out, "w", encoding="utf-8") as f:
        f.write("\n".join("@" + n for n in out) + "\n")
    print(f"Готово: {len(out)} имён из {len(blocks)} блоков за {time.time() - t0:.0f} с → {args.out}")


if __name__ == "__main__":
    main()
