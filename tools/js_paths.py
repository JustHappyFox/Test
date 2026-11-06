"""Куда обращаются скрипты страницы t.me (вопрос владельца 06.11: «данные розыгрыша должны вставляться
по какому-то пути — путь должен быть в коде страницы»). Скачиваем скрипты как текст (ничего не
запускаем) и выписываем все обращения к сети: ajax/XMLHttpRequest/fetch, адреса, имена методов API.
"""
import re
import sys
import urllib.request

UA = {"User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"}
SCRIPTS = ["https://telegram.org/js/telegram-web.js?14", "https://telegram.org/js/widget-frame.js?68",
           "https://oauth.tg.dev/js/telegram-widget.js?24", "https://telegram.org/js/tgsticker.js?32"]
NET = re.compile(r"(\$\.ajax\(|\.ajax\(|XMLHttpRequest|fetch\(|\.open\(|apiRequest|api_url|upload_url|method\s*[:=]|"
                 r"['\"]/[a-z_]+/?['\"]|https?://[^'\"\s)]+)")


def get(url):
    with urllib.request.urlopen(urllib.request.Request(url, headers=UA), timeout=30) as r:
        return r.read().decode("utf-8", "replace")


def main():
    for url in SCRIPTS + sys.argv[1:]:
        try:
            src = get(url)
        except Exception as e:
            print(f"\n=== {url}: не скачался ({e})")
            continue
        print(f"\n=== {url}: {len(src)} символов")
        seen = set()
        for m in NET.finditer(src):
            a, b = max(0, m.start() - 90), min(len(src), m.end() + 140)
            ctx = re.sub(r"\s+", " ", src[a:b])
            key = ctx[60:200]
            if key in seen:
                continue
            seen.add(key)
            print("  …", ctx)
        methods = sorted(set(re.findall(r"method\s*:\s*['\"]([A-Za-z_.]+)['\"]", src)) |
                         set(re.findall(r"apiRequest\(\s*['\"]([A-Za-z_.]+)['\"]", src)))
        print("  методы API, упомянутые в скрипте:", methods or "нет")


if __name__ == "__main__":
    main()
