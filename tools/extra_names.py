"""Новые источники @username каналов (владелец 07.11: «надо расширить список каналов, не только русские»).

  wayback  — индекс Wayback Machine (archive.org CDX) по адресам t.me/…: сохранённые за годы страницы
             каналов со всего мира. Берём страницы индекса по очереди (--pages), из адресов — имена.
  wikidata — свойство P3789 «Telegram username»: официальные аккаунты компаний, СМИ, брендов всех стран.

Пишем только имена, которых ещё нет в --known. Аккаунт Telegram не нужен.
"""
import argparse
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

UA = {"User-Agent": "giveaway-finder-research/1.0 (public channel discovery)"}
NAME = re.compile(r"^(?:https?://)?(?:www\.)?(?:t|telegram)\.me(?::\d+)?/(?:s/)?([A-Za-z][A-Za-z0-9_]{3,31})(?:[/?#]|$)", re.I)
SKIP = {"joinchat", "addstickers", "addemoji", "share", "proxy", "socks", "iv", "login", "addlist", "boost",
        "giftcode", "invoice", "setlanguage", "confirmphone", "contact", "addtheme", "bg", "c"}


def get(url, tries=4, timeout=120):
    for i in range(tries):
        try:
            with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=timeout) as r:
                return r.read().decode("utf-8", "replace")
        except Exception as e:
            print(f"  повтор {url[:100]}: {e}", flush=True)
            time.sleep(10 * (i + 1))
    return ""


def wayback(pages: int, start: int, limit: int = 50000, pause: float = 15.0, on_letter=None):
    """Индекс Wayback Machine по t.me/. Первые сотни тысяч адресов — приглашения «t.me/+…» (они идут первыми
    по алфавиту), поэтому идём по буквам: t.me/a…, t.me/b…, и берём только страницы каналов
    (t.me/имя и t.me/s/имя) — фильтр на стороне archive.org. Медленно: при частых запросах 503."""
    names = set()
    letters = "abcdefghijklmnopqrstuvwxyz"[start:]
    for ch in letters:
        resume = ""
        for i in range(pages):
            # страницы t.me/s/<канал> — у каждого канала одна (плюс ?before=…), фильтр не нужен:
            # фильтр на стороне archive.org по всему t.me/ падает с 504
            url = (f"https://web.archive.org/cdx/search/cdx?url=t.me/s/{ch}&matchType=prefix&fl=original&collapse=urlkey" +
                   f"&limit={limit}&showResumeKey=true" + (f"&resumeKey={urllib.parse.quote(resume)}" if resume else ""))
            txt = get(url, tries=6, timeout=300)
            lines = txt.splitlines()
            resume = lines[-1].strip() if len(lines) >= 2 and not lines[-2].strip() else ""
            before = len(names)
            for ln in lines:
                m = NAME.match(ln.strip())
                if m and m.group(1).lower() not in SKIP:
                    names.add(m.group(1))
            print(f"  t.me/s/{ch}… кусок {i}: адресов {len(lines)}, новых имён {len(names) - before}, всего {len(names)}",
                  flush=True)
            time.sleep(pause)
            if not resume:
                break
        if on_letter:
            on_letter(names)  # промежуточное сохранение: задание могут прервать по времени
    return names


def wikidata():
    q = "SELECT ?u WHERE { ?i wdt:P3789 ?u }"
    txt = get("https://query.wikidata.org/sparql?format=json&query=" + urllib.parse.quote(q), timeout=300)
    names = set()
    try:
        for b in json.loads(txt)["results"]["bindings"]:
            u = b["u"]["value"].strip().lstrip("@")
            if re.fullmatch(r"[A-Za-z][A-Za-z0-9_]{3,31}", u):
                names.add(u)
    except Exception as e:
        print("Wikidata: не разобрал ответ:", e, txt[:300])
    print(f"Wikidata P3789: имён {len(names)}", flush=True)
    return names


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("source", choices=["wayback", "wikidata", "both"])
    ap.add_argument("--known", default="state/cc_usernames.txt")
    ap.add_argument("--pages", type=int, default=5)
    ap.add_argument("--start", type=int, default=0)
    ap.add_argument("--out", required=True)
    a = ap.parse_args()
    known = {ln.strip().lstrip("@").lower() for ln in Path(a.known).read_text(encoding="utf-8").splitlines() if ln.strip()}
    names = set()
    if a.source in ("wikidata", "both"):
        names |= wikidata()
    if a.source in ("wayback", "both"):
        def save_partial(part):
            new = sorted(n for n in names | part if n.lower() not in known)
            Path(a.out).write_text("".join(f"@{n}\n" for n in new), encoding="utf-8")
        names |= wayback(a.pages, a.start, on_letter=save_partial)
    new = sorted(n for n in names if n.lower() not in known)
    Path(a.out).write_text("".join(f"@{n}\n" for n in new), encoding="utf-8")
    print(f"Итог: имён {len(names)}, новых (нет в нашем списке из {len(known)}) — {len(new)} → {a.out}")
    print("Примеры новых:", " ".join(new[:: max(1, len(new) // 40)][:40]))


if __name__ == "__main__":
    main()
