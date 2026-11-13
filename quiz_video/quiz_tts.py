"""Озвучка квиза клоном голоса клиента (владелец 13.11: «можем на основе её голоса сами дальше озвучивать?»;
«она сама и просила, чтобы не записывать каждый раз»).

  python quiz_tts.py JOB.json --tts ПАПКА --out ПАПКА_ФРАЗ --work РАБОЧАЯ [--whisper ПАПКА_МОДЕЛИ]
  → ПАПКА_ФРАЗ/01.wav … (по порядку: вступление, вопрос и пояснение по каждому вопросу, концовка);
    дальше: make_quiz.py … --voice-files ПАПКА_ФРАЗ --voice-ready

ПАПКА (--tts) с файлом voice.json (пути относительно неё):
  {"python": окружение с qwen-tts, "model": модель Qwen3-TTS Base, "ref_audio": образец голоса в одобренной
   обработке, "ref_text": его точный текст, "ref_set": шаблон файлов образца для сверки сходства,
   "spk_model": onnx диктора (см. voice_qa.py), "post_eq": эквалайзер после синтеза}

Как это работает и почему:
  * синтез (tts_clone.py) с первого раза иногда даёт брак, особенно на коротких фразах → по несколько вариантов
    каждой фразы (разные сиды) и автоотбор по измерениям: слова совпали с текстом (Whisper), голос похож на
    образец, качество DNSMOS, нет срывов высоты (размах > 12 полутонов);
  * слабым фразам после первого круга добавляются ещё варианты;
  * отобранное: эквалайзер (синтез глуше сверху, чем образец) и громкость как у живых фраз (voice_fx.level).
Интонацию и ударения измерить нечем — их проверяют на слух владелец и клиент.
"""
import argparse
import difflib
import glob
import json
import os
import re
import subprocess
import sys
import warnings
from pathlib import Path

import numpy as np

warnings.filterwarnings("ignore")
HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
import quiz_flow  # noqa: E402
import voice_fx  # noqa: E402
import voice_qa  # noqa: E402

GOOD = {"txt": 0.90, "sim": 0.70, "ovrl": 2.9}     # ниже — фраза «слабая», добавляем варианты


def norm(t: str) -> str:
    return " ".join(re.sub(r"[^\w\s]", " ", t.lower().replace("ё", "е")).split())


def shown_texts(quiz: dict) -> dict[str, str]:
    """Текст «как на экране» (с цифрами) — с ним сверяем распознанное: Whisper пишет числа цифрами."""
    d = {"intro": quiz.get("intro_voice") or f"Квиз про {quiz['topic']}! Проверь себя!"}
    for i, q in enumerate(quiz["questions"], 1):
        d[f"q{i}"], d[f"e{i}"] = q["q"], q["explain"]
    d["outro"] = quiz.get("outro_voice") or quiz_flow.OUTRO
    return d


def f0_range(f: str) -> float:
    import parselmouth
    p = parselmouth.Sound(f).to_pitch(time_step=0.01, pitch_floor=100, pitch_ceiling=500).selected_array["frequency"]
    v = p[p > 0]
    if len(v) < 10:
        return 99.0
    st = 12 * np.log2(v / np.median(v))
    return float(np.percentile(st, 95) - np.percentile(st, 5))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("job")
    ap.add_argument("--tts", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--work", required=True)
    ap.add_argument("--whisper", default=str(Path.home() / ".cache" / "whisper"))
    ap.add_argument("--seeds", default="7,11,23")
    ap.add_argument("--more-seeds", default="31,47,59", help="добавка для слабых фраз")
    a = ap.parse_args()

    quiz = json.loads(Path(a.job).read_text(encoding="utf-8"))
    tts = Path(a.tts).resolve()
    v = json.loads((tts / "voice.json").read_text(encoding="utf-8"))
    cand = Path(a.work).resolve() / "tts"
    cand.mkdir(parents=True, exist_ok=True)
    items = quiz_flow.fragments(quiz)               # [(имя, текст «как читать»)]
    shown = shown_texts(quiz)
    spk = str(tts / v["spk_model"])
    center, _ = voice_qa.centroid(sorted(glob.glob(str(tts / v["ref_set"]))), spk)

    def synth(jobs):
        cfg = {"model": str(tts / v["model"]), "out": str(cand), "ref_audio": str(tts / v["ref_audio"]),
               "ref_text": v["ref_text"], "jobs": jobs}
        (cand / "job.json").write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        subprocess.run([str(tts / v["python"]), str(HERE / "tts_clone.py"), str(cand / "job.json")], check=True)

    os.environ.setdefault("HF_HUB_DISABLE_XET", "1")
    from faster_whisper import WhisperModel
    wm = WhisperModel("small", device="cpu", compute_type="int8", download_root=a.whisper)
    cache_path = cand / "scores.json"
    cache = json.loads(cache_path.read_text()) if cache_path.exists() else {}

    def measure(f: str, name: str, text: str) -> dict:
        if f not in cache:
            segs, _ = wm.transcribe(voice_qa.dec(f).astype(np.float32), language="ru", beam_size=5)
            heard = " ".join(s.text.strip() for s in segs)
            txt = max(difflib.SequenceMatcher(None, norm(heard), norm(t)).ratio() for t in (shown[name], text))
            m = voice_qa.score([f])[0]
            r = dict(file=f, heard=heard, txt=txt, sim=float(voice_qa.embed(f, spk) @ center), ovrl=float(m[0]),
                     bak=float(m[2]), rng=f0_range(f))
            # слова важнее всего; затем сходство и качество; срывы высоты штрафуем
            r["total"] = 4 * r["txt"] + 2 * r["sim"] + 0.5 * r["ovrl"] - (0.5 if r["rng"] > 12 else 0)
            cache[f] = r
            cache_path.write_text(json.dumps(cache, ensure_ascii=False, indent=1))
        return cache[f]

    def pick() -> dict:
        best = {}
        for name, text in items:
            rows = [measure(f, name, text) for f in sorted(glob.glob(str(cand / f"{name}_s*.wav")))]
            best[name] = max(rows, key=lambda r: r["total"])
        return best

    weak = lambda r: r["txt"] < GOOD["txt"] or r["sim"] < GOOD["sim"] or r["ovrl"] < GOOD["ovrl"] or r["rng"] > 12
    synth([[n, t, int(s)] for n, t in items for s in a.seeds.split(",")])
    best = pick()
    again = [(n, t) for n, t in items if weak(best[n])]
    if again:
        print("слабые фразы, добавляю варианты:", ", ".join(n for n, _ in again), flush=True)
        synth([[n, t, int(s)] for n, t in again for s in a.more_seeds.split(",")])
        best = pick()

    import soundfile as sf
    out = Path(a.out)
    out.mkdir(parents=True, exist_ok=True)
    for old in out.glob("*.wav"):
        old.unlink()
    sr = 44100
    for k, (name, _) in enumerate(items, 1):
        r = best[name]
        x, src_sr = sf.read(r["file"], dtype="float32")
        y = voice_fx.ffilter(x, "highpass=f=80," + v["post_eq"], src_sr, sr)
        sf.write(out / f"{k:02d}.wav", voice_fx.level(y, sr), sr, subtype="PCM_16")
        flag = "  ПРОВЕРИТЬ НА СЛУХ" if weak(r) else ""
        print(f"{k:02d} {name:5} {Path(r['file']).name:14} слова {r['txt']:.2f} сходство {r['sim']:.2f} "
              f"OVRL {r['ovrl']:.2f} размах {r['rng']:4.1f}пт{flag} | {r['heard'][:60]}")
    rows = list(best.values())
    print(f"среднее: слова {np.mean([r['txt'] for r in rows]):.2f}, сходство {np.mean([r['sim'] for r in rows]):.2f}, "
          f"OVRL {np.mean([r['ovrl'] for r in rows]):.2f}; слабых фраз: {sum(weak(r) for r in rows)}")
    (out / "report.json").write_text(json.dumps(best, ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
