"""Конвейер квиз-ролика с озвучкой клиента (владелец 12.11: «построй скрипт последовательности действий,
чтобы не терять время и ресурсы»).

Порядок: текст на озвучку → картинки рисуются, пока клиент записывает (≈1 ч) → голосовые распознаются
и раскладываются по пунктам сами (дубли — берём последний) → сборка → отправка.

  python quiz_flow.py script  JOB.json                      — пункты для озвучки по порядку (1. … 12. …)
  python quiz_flow.py collect JOB.json --media DIR --after ID --out VOICE_DIR
        Распознаёт голосовые из DIR (имена — номер сообщения: 000114.oga), новее сообщения ID,
        сопоставляет с пунктами и кладёт в VOICE_DIR как 01.oga … 12.oga.
        Код выхода: 0 — все пункты есть; 3 — каких-то не хватает; 4 — все есть, но пришло что-то неопознанное.
        --other JOB2.json … — темы, которые записываются в те же дни: их записи пропускаются.
  python quiz_flow.py build   JOB.json --voices VOICE_DIR --out ROLIK.mp4 --work DIR --assets DIR [-- аргументы make_quiz]
"""
import argparse
import difflib
import json
import os
import re
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
AUDIO_EXT = {".ogg", ".oga", ".opus", ".mp3", ".m4a", ".wav", ".aac", ".flac", ".mp4", ".webm"}
OUTRO = "Сколько угадал? Пиши в комментариях!"


def fragments(quiz: dict) -> list[tuple[str, str]]:
    """Пункты озвучки по порядку: вступление, вопрос и пояснение по каждому вопросу, концовка."""
    items = [("intro", quiz.get("intro_voice") or f"Квиз про {quiz['topic']}! Проверь себя!")]
    for i, q in enumerate(quiz["questions"], 1):
        items += [(f"q{i}", q.get("q_voice") or q["q"]), (f"e{i}", q.get("explain_voice") or q["explain"])]
    items.append(("outro", quiz.get("outro_voice") or OUTRO))
    return items


def norm(text: str) -> str:
    text = text.lower().replace("ё", "е")
    text = re.sub(r"[^\w\s]", " ", text)
    return " ".join(text.split())


def transcribe(files: list[Path], cache: Path, model_dir: Path) -> dict[str, str]:
    """Whisper small (локально); результаты кешируем, чтобы не распознавать одно и то же по кругу."""
    done = json.loads(cache.read_text(encoding="utf-8")) if cache.exists() else {}
    todo = [f for f in files if f"{f.name}|{f.stat().st_size}" not in done]
    if todo:
        os.environ.setdefault("HF_HUB_DISABLE_XET", "1")  # xet-сервер HF закрыт сетью — качаем по обычной ссылке
        import numpy as np
        from faster_whisper import WhisperModel
        m = WhisperModel("small", device="cpu", compute_type="int8", download_root=str(model_dir))
        for f in todo:
            raw = subprocess.run(["ffmpeg", "-v", "error", "-i", str(f), "-ac", "1", "-ar", "16000", "-f", "f32le", "-"],
                                 capture_output=True, check=True).stdout
            segs, _ = m.transcribe(np.frombuffer(raw, np.float32), language="ru", beam_size=5)
            done[f"{f.name}|{f.stat().st_size}"] = " ".join(s.text.strip() for s in segs)
        cache.write_text(json.dumps(done, ensure_ascii=False, indent=1), encoding="utf-8")
    return {f.name: done[f"{f.name}|{f.stat().st_size}"] for f in files}


def collect(quiz: dict, media: Path, after: int, out: Path, model_dir: Path, min_score: float = 0.45,
            others: list[dict] = ()) -> int:
    """others — другие темы, которые клиент озвучивает в те же дни (до 3 роликов в день, владелец 13.11):
    запись, которая больше похожа на пункт другой темы, тихо пропускаем — она для другого ролика."""
    items = fragments(quiz)
    foreign = [t for o in others for _, t in fragments(o)]
    files = sorted((f for f in media.iterdir() if f.suffix.lower() in AUDIO_EXT and f.stem.isdigit()
                    and int(f.stem) > after), key=lambda f: int(f.stem))
    if not files:
        print("Записей пока нет")
        return 3
    texts = transcribe(files, media / ".transcripts.json", model_dir)
    best: dict[str, tuple[int, Path, float]] = {}
    unknown = []
    for f in files:
        heard = norm(texts[f.name])
        scores = [(difflib.SequenceMatcher(None, heard, norm(t)).ratio(), name) for name, t in items]
        score, name = max(scores)
        if foreign and max(difflib.SequenceMatcher(None, heard, norm(t)).ratio() for t in foreign) > score:
            continue  # это пункт другой темы
        if score < min_score:
            unknown.append(f"{f.name}: «{texts[f.name]}» (похоже на {name} лишь на {score:.2f})")
            continue
        if name not in best or int(f.stem) > best[name][0]:  # дубль — берём последний
            best[name] = (int(f.stem), f, score)
    if out.exists():
        shutil.rmtree(out)
    out.mkdir(parents=True)
    missing = []
    for n, (name, text) in enumerate(items, 1):
        if name in best:
            _, f, score = best[name]
            shutil.copy(f, out / f"{n:02d}{f.suffix}")
            print(f"{n:02d} {name:5} ← {f.name} ({score:.2f}): {texts[f.name]}")
        else:
            missing.append(f"{n}. {text}")
    for u in unknown:
        print("НЕ ОПОЗНАНА", u)
    for m in missing:
        print("НЕТ ЗАПИСИ", m)
    if missing:
        print(f"Не хватает {len(missing)}, не опознано {len(unknown)}")
        return 3
    if unknown:  # все пункты есть, но пришло что-то лишнее (комментарий голосом?) — пусть посмотрит человек
        print("Все пункты есть, но есть неопознанные записи — проверь")
        return 4
    print("ВСЕ ЗАПИСИ НА МЕСТЕ")
    return 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("cmd", choices=["script", "collect", "build"])
    ap.add_argument("job")
    ap.add_argument("--media")
    ap.add_argument("--after", type=int, default=0)
    ap.add_argument("--out")
    ap.add_argument("--voices")
    ap.add_argument("--work")
    ap.add_argument("--assets")
    ap.add_argument("--whisper", help="папка модели Whisper", default=str(Path.home() / ".cache" / "whisper"))
    ap.add_argument("--other", nargs="*", default=[], help="задания других тем, которые записываются параллельно")
    a, rest = ap.parse_known_args()
    quiz = json.loads(Path(a.job).read_text(encoding="utf-8"))
    if a.cmd == "script":
        print("\n".join(f"{n}. {t}" for n, (_, t) in enumerate(fragments(quiz), 1)))
    elif a.cmd == "collect":
        others = [json.loads(Path(o).read_text(encoding="utf-8")) for o in a.other if Path(o) != Path(a.job)]
        sys.exit(collect(quiz, Path(a.media), a.after, Path(a.out), Path(a.whisper), others=others))
    else:
        cmd = [sys.executable, str(HERE / "make_quiz.py"), a.job, "--out", a.out, "--work", a.work,
               "--assets", a.assets, "--voice-files", a.voices] + [x for x in rest if x != "--"]
        sys.exit(subprocess.run(cmd).returncode)


if __name__ == "__main__":
    main()
