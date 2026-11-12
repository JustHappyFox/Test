"""Квиз-ролик по теме: картинки (Z-Image, локально), озвучка (Piper, локально), кадры (Chromium), сборка (ffmpeg).

  python make_quiz.py quiz.json --out ролик.mp4 --assets ПАПКА_С_МОДЕЛЯМИ

quiz.json:
  {"topic": "Луна", "hook": "…", "cover_prompt": "англ. описание картинки",
   "icons": ["🍫", "🍬"],
   "questions": [{"q": "…", "options": ["…", "…", "…"], "answer": 0,
                  "explain": "пояснение на обороте карточки (история: кто открыл, почему так)", "icons": ["🌱"],
                  "image_prompt": "…", "q_voice"/"explain_voice": "как читать вслух (числа словами)"}],
   "outro_title": "…", "outro_sub": "…", "outro_voice": "…", "palette": 0}
Правила владельца (11.11): 5 вопросов, сложность растёт от «очень легко» до «очень сложно»; 10 с на размышление
(таймер идёт, пока голос читает вопрос); после ответа карточка переворачивается и голос читает пояснение;
ролик от 62 с; 90 с — ориентир (владелец 12.11): главное, чтобы каждый текст успели прочитать.

В ПАПКЕ_С_МОДЕЛЯМИ: models/z_image_turbo-Q4_0.gguf, models/qwen_3_4b-Q8_0.gguf, models/ae-f16.gguf,
voices/ru_RU-<голос>-medium.onnx (+ .onnx.json). Модели и голоса в репозиторий не кладём — они большие.
"""
import argparse
import hashlib
import json
import math
import subprocess
import sys
import time
import wave
from pathlib import Path

import numpy as np

HERE = Path(__file__).resolve().parent
SR = 44100  # частота звуковой дорожки (Piper даёт 22050 — пересчитываем)

PALETTES = [
    {"bg1": "#ff8a3d", "bg2": "#ffc145", "accent": "#1fb6e8"},   # оранжевая
    {"bg1": "#6a3df0", "bg2": "#c04bf4", "accent": "#ff9f1c"},   # фиолетовая
    {"bg1": "#0f9b8e", "bg2": "#46e08a", "accent": "#ff5e7e"},   # бирюзовая
    {"bg1": "#2448a8", "bg2": "#3fa9f5", "accent": "#ffb703"},   # синяя
    {"bg1": "#ff5f9e", "bg2": "#ffa58b", "accent": "#4d5bff"},   # розовая
]
# Сложность растёт плавно: подписи по номеру вопроса (5 ступеней), цвет от зелёного к красному.
LEVELS = [("ОЧЕНЬ ЛЕГКО", "#22c55e"), ("ЛЕГКО", "#84cc16"), ("СРЕДНЕ", "#f59e0b"),
          ("СЛОЖНО", "#f97316"), ("ОЧЕНЬ СЛОЖНО", "#ef4444")]
STYLE = ("simple flat vector cartoon drawing, children's book illustration, thick dark outlines, solid bright colors, "
         "clean simple background, not a photo, no photorealism, no text, no letters, no words")


def log(*a):
    print(time.strftime("%H:%M:%S"), *a, flush=True)


# ---------- картинки ----------
class Painter:
    def __init__(self, assets: Path, threads: int):
        self.assets, self.threads, self.sd = assets, threads, None

    def _load(self):
        from stable_diffusion_cpp import StableDiffusion
        m = self.assets / "models"
        log("загружаю модель картинок…")
        self.sd = StableDiffusion(diffusion_model_path=str(m / "z_image_turbo-Q4_0.gguf"),
                                  llm_path=str(m / "qwen_3_4b-Q8_0.gguf"), vae_path=str(m / "ae-f16.gguf"),
                                  n_threads=self.threads, diffusion_flash_attn=True, verbose=False)

    def paint(self, prompt: str, w: int, h: int, out: Path, seed: int, steps: int) -> Path:
        if out.exists():
            return out
        if self.sd is None:
            self._load()
        t = time.time()
        img = self.sd.generate_image(prompt=f"{prompt}, {STYLE}", width=w, height=h, cfg_scale=1.0,
                                     sample_steps=steps, seed=seed)[0]
        img.save(out)
        log(f"картинка {out.name}: {time.time() - t:.0f} с")
        return out


def placeholder(out: Path, w: int, h: int, color: str) -> Path:
    """Если модели нет — однотонная заглушка (ролик всё равно соберётся)."""
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-f", "lavfi", "-i", f"color=c={color}:s={w}x{h}",
                    "-frames:v", "1", str(out)], check=True)
    return out


# ---------- голос ----------
def say(text: str, voice: Path, out: Path, speed: float) -> np.ndarray:
    if not out.exists():
        raw = out.with_name(out.stem + "_raw.wav")
        subprocess.run([sys.executable, "-m", "piper", "-m", str(voice), "-f", str(raw),
                        "--length-scale", str(speed), "--sentence-silence", "0.15"],
                       input=text.encode(), check=True, capture_output=True)
        subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", str(raw), "-ar", str(SR), str(out)], check=True)
    return read_voice(out)


def read_voice(path: Path, peak_norm: bool = True) -> np.ndarray:
    with wave.open(str(path)) as w:
        assert w.getframerate() == SR, w.getframerate()
        a = np.frombuffer(w.readframes(w.getnframes()), dtype=np.int16).astype(np.float32) / 32768
    if not peak_norm:
        return a
    peak = float(np.abs(a).max()) or 1.0
    return a * (0.85 / peak)


AUDIO_EXT = {".ogg", ".oga", ".opus", ".mp3", ".m4a", ".wav", ".aac", ".flac", ".mp4", ".webm"}


def own_voice(src: Path, work: Path, name: str, opts: dict) -> np.ndarray:
    """Живой голос из записи (голосовое Telegram и т.п.) — обработка voice_fx.py: шумы, тишина по краям,
    лёгкий автотюн, немного другой голос (высота и тембр), ровная громкость."""
    import voice_fx
    code = hashlib.md5(Path(voice_fx.__file__).read_bytes()).hexdigest()[:6]  # правки обработки — пересобрать
    key = hashlib.md5(f"{src.resolve()}|{src.stat().st_size}|{sorted(opts.items())}|{code}".encode()).hexdigest()[:8]
    out = work / f"{name}_own_{key}.wav"
    if not out.exists():
        voice_fx.process(src, out, SR, **opts)
    return read_voice(out, peak_norm=False)  # громкость уже выровнена по RMS


# ---------- звуки ----------
def tone(freqs, dur, decay, vol):
    t = np.arange(int(SR * dur)) / SR
    s = sum(np.sin(2 * np.pi * f * t) for f in freqs) / len(freqs)
    return (s * np.exp(-t * decay) * vol).astype(np.float32)


def tick(high=False):
    return tone([2200 if high else 1600], 0.05, 70, 0.5 if high else 0.28)


def ding():
    return tone([1046.5, 1318.5, 1568.0], 1.1, 4.5, 0.42)


def pop():
    t = np.arange(int(SR * 0.09)) / SR
    f = 380 + 900 * t / 0.09
    return (np.sin(2 * np.pi * np.cumsum(f) / SR) * np.exp(-t * 35) * 0.22).astype(np.float32)


def whoosh(seed):
    n = int(SR * 0.35)
    r = np.random.default_rng(seed).standard_normal(n).astype(np.float32)
    r = np.convolve(r, np.ones(24) / 24, mode="same")  # мягкий шум
    env = np.sin(np.linspace(0, np.pi, n)) ** 2
    return r * env * 0.35


def music(total: float) -> np.ndarray:
    """Негромкая «музыкальная шкатулка»: арпеджио по аккордам C–Am–F–G, 104 удара в минуту."""
    bpm, out = 104, np.zeros(int(SR * (total + 2)), np.float32)
    chords = [[60, 64, 67, 72], [57, 60, 64, 69], [53, 57, 60, 65], [55, 59, 62, 67]]
    step = 60 / bpm / 2  # восьмые
    pattern = [0, 1, 2, 3, 2, 1, 2, 3]
    i = 0
    while i * step < total + 1:
        ch = chords[(i // 8) % 4]
        n = ch[pattern[i % 8]] + 12
        f = 440 * 2 ** ((n - 69) / 12)
        t = np.arange(int(SR * 0.6)) / SR
        note = (np.sin(2 * np.pi * f * t) + 0.3 * np.sin(4 * np.pi * f * t)) * np.exp(-t * 7) * 0.06
        if i % 8 == 0:  # бас
            fb = 440 * 2 ** ((ch[0] - 12 - 69) / 12)
            note = note + np.sin(2 * np.pi * fb * t) * np.exp(-t * 3) * 0.07
        a = int(i * step * SR)
        out[a:a + len(note)] += note.astype(np.float32)
        i += 1
    fade = int(SR * 1.5)
    end = int(SR * total)
    out[end - fade:end] *= np.linspace(1, 0, fade)
    out[end:] = 0
    return out


def music_calm(total: float) -> np.ndarray:
    """Спокойная тихая подложка (владелец 12.11): мягкие аккорды Cmaj7–Am7–Fmaj7–G6 по 4 с и редкие
    «капельки» на сильные доли. Без ударных, без резких атак."""
    n = int(SR * (total + 2))
    out = np.zeros(n, np.float32)
    chords = [[48, 55, 59, 64], [45, 52, 55, 60], [41, 48, 52, 57], [43, 50, 52, 59]]
    bar = 4.0
    rng = np.random.default_rng(7)
    k = 0
    while k * bar < total + 1:
        ch = chords[k % 4]
        dur = bar + 1.5  # с перекрытием — аккорды перетекают
        t = np.arange(int(SR * dur)) / SR
        env = np.minimum(1, t / 0.9) * np.minimum(1, np.maximum(0, (dur - t) / 1.4))
        pad = np.zeros_like(t)
        for m in ch:
            f = 440 * 2 ** ((m - 69) / 12)
            for det in (-0.0015, 0.0015):  # лёгкая расстройка — «тёплый» пэд
                ph = rng.uniform(0, 2 * np.pi)
                pad += np.sin(2 * np.pi * f * (1 + det) * t + ph) + 0.12 * np.sin(4 * np.pi * f * (1 + det) * t)
        a = int(k * bar * SR)
        seg = (pad * env / (len(ch) * 2)).astype(np.float32)
        out[a:a + len(seg)] += seg[:max(0, n - a)]
        for beat in (0, 2):  # «капельки»: тоны аккорда октавой выше, тихо
            m = ch[(k + beat) % len(ch)] + 24
            f = 440 * 2 ** ((m - 69) / 12)
            tt = np.arange(int(SR * 1.6)) / SR
            drop = (np.sin(2 * np.pi * f * tt) * np.exp(-tt * 3.2) * np.minimum(1, tt / 0.01) * 0.35).astype(np.float32)
            b = int((k * bar + beat) * SR)
            out[b:b + len(drop)] += drop[:max(0, n - b)]
        k += 1
    end = int(SR * total)
    fade = int(SR * 2.0)
    out[end - fade:end] *= np.linspace(1, 0, fade)
    out[end:] = 0
    out[:int(SR * 1.0)] *= np.linspace(0, 1, int(SR * 1.0))
    return out


def duck(music: np.ndarray, voice: np.ndarray, depth_db: float) -> np.ndarray:
    """Музыка тише, пока звучит голос (плавно: атака ~80 мс, отпускание ~500 мс)."""
    w = int(SR * 0.02)
    n = len(voice) // w
    rms = np.sqrt(np.mean(voice[:n * w].reshape(n, w) ** 2, axis=1))
    on = (rms > 10 ** (-40 / 20)).astype(np.float32)
    env = np.zeros_like(on)
    up, down = 1 - np.exp(-1 / 4), 1 - np.exp(-1 / 25)  # в окнах по 20 мс
    v = 0.0
    for i, x in enumerate(on):
        v += (x - v) * (up if x > v else down)
        env[i] = v
    g = np.repeat(10 ** (-depth_db * env / 20), w)
    g = np.concatenate([g, np.full(max(0, len(music) - len(g)), g[-1] if len(g) else 1.0)])[:len(music)]
    return music * g.astype(np.float32)


def mix(track: np.ndarray, clip: np.ndarray, at: float):
    a = int(at * SR)
    b = min(len(track), a + len(clip))
    track[a:b] += clip[:b - a]


# ---------- сборка ----------
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("quiz")
    ap.add_argument("--out", required=True)
    ap.add_argument("--assets", required=True, help="папка с models/ и voices/")
    ap.add_argument("--work", help="рабочая папка (по умолчанию рядом с --out)")
    ap.add_argument("--voice", default="dmitri")
    ap.add_argument("--speed", type=float, default=0.9, help="темп речи Piper: меньше — быстрее")
    ap.add_argument("--voice-files", help="папка с записями живого голоса по порядку: вступление, вопрос 1, "
                                          "пояснение 1, …, вопрос N, пояснение N, концовка (клиент 12.11)")
    ap.add_argument("--voice-pitch", type=float, default=-1.5, help="сдвиг живого голоса, полутона (0 — без изменений)")
    ap.add_argument("--voice-tune", type=float, default=0.4, help="«автотюн»: доля подтяжки к нотам (0 — выкл., 1 — робот)")
    ap.add_argument("--voice-formant", type=float, default=1.0, help="тембр живого голоса (<1 — ниже; сдвиг тембра хуже для разборчивости — проверка 12.11)")
    ap.add_argument("--voice-tempo", type=float, default=1.0, help="скорость пояснений и концовки живым голосом (1.1 — на 10%% быстрее)")
    ap.add_argument("--timer", type=float, default=10.0, help="секунд на размышление")
    ap.add_argument("--min-duration", type=float, default=62.0, help="ролик не короче, с")
    ap.add_argument("--max-duration", type=float, default=100.0,
                    help="ориентир длины, с: длиннее — только предупреждение (владелец 12.11: 90 с — ориентировочно, "
                         "можно 95–100, главное — чтобы текст успели прочитать)")
    ap.add_argument("--read-cps", type=float, default=15.0,
                    help="скорость чтения зрителя, знаков в секунду: пояснение висит, пока его можно прочитать")
    ap.add_argument("--steps", type=int, default=6)
    ap.add_argument("--image-size", default="512x384", help="картинки вопросов (рисунки, не фото — владелец 11.11)")
    ap.add_argument("--cover-size", default="512x512")
    ap.add_argument("--threads", type=int, default=4)
    ap.add_argument("--no-images", action="store_true", help="без нейросети: заглушки вместо картинок")
    ap.add_argument("--no-music", action="store_true")
    ap.add_argument("--music", choices=["calm", "box"], default="calm",
                    help="calm — спокойная подложка (владелец 12.11), box — шкатулка-арпеджио")
    ap.add_argument("--music-db", type=float, default=-30, help="громкость музыки без голоса, дБ (RMS)")
    ap.add_argument("--duck-db", type=float, default=9, help="на сколько дБ музыка тише под голосом")
    ap.add_argument("--loudness", type=float, default=-16, help="итоговая громкость, LUFS (−16 — голос не «кричит»)")
    ap.add_argument("--preview", help="только кадры PNG в эти секунды, через запятую")
    ap.add_argument("--audio-only", action="store_true", help="только звуковая дорожка (audio.wav и audio.mp3)")
    ap.add_argument("--images-only", action="store_true",
                    help="только нарисовать картинки (пока клиент записывает голос — экономим время)")
    a = ap.parse_args()

    quiz = json.loads(Path(a.quiz).read_text(encoding="utf-8"))
    out = Path(a.out).resolve()
    work = Path(a.work) if a.work else out.with_suffix("")
    work.mkdir(parents=True, exist_ok=True)
    assets = Path(a.assets)
    seed = int(hashlib.md5(quiz["topic"].encode()).hexdigest()[:6], 16)
    pal = dict(PALETTES[quiz.get("palette", seed) % len(PALETTES)])
    pal.update({"accent-dark": pal["accent"]})
    qs = quiz["questions"]

    # картинки: пути заранее (рисуем после проверки длины — это долго)
    painter = Painter(assets, a.threads)
    model_ok = not a.no_images and (assets / "models" / "z_image_turbo-Q4_0.gguf").exists()
    jobs = []

    def picture(prompt, w, h, name, s):
        p = work / f"{name}_{hashlib.md5(f'{prompt}|{w}x{h}|{a.steps}|{s}'.encode()).hexdigest()[:8]}.png"
        jobs.append((prompt, w, h, p, s))
        return p

    cw, ch = map(int, a.cover_size.split("x"))
    qw, qh = map(int, a.image_size.split("x"))
    cover = picture(quiz["cover_prompt"], cw, ch, "cover", seed)
    images = [picture(q["image_prompt"], qw, qh, f"q{i + 1}", seed + q.get("seed_off", i + 1))
              for i, q in enumerate(qs)]

    if a.images_only:
        for prompt, w, h, p, sd_seed in jobs:
            if model_ok:
                painter.paint(prompt, w, h, p, sd_seed, a.steps)
        log("картинки готовы:", ", ".join(p.name for *_, p, _s in jobs))
        return

    def read_time(text: str) -> float:  # сколько секунд нужно, чтобы спокойно прочитать текст
        return len(text) / a.read_cps + 1.0

    # голос и раскладка по времени
    voice = assets / "voices" / f"ru_RU-{a.voice}-medium.onnx"
    own = {}
    if a.voice_files:  # живой голос вместо Piper
        files = sorted(p for p in Path(a.voice_files).iterdir() if p.suffix.lower() in AUDIO_EXT)
        names = ["intro"] + [f"{k}{i + 1}" for i in range(len(qs)) for k in ("q", "e")] + ["outro"]
        if len(files) != len(names):
            sys.exit(f"ОШИБКА: записей {len(files)}, а нужно {len(names)} (вступление, вопрос и пояснение ×{len(qs)}, "
                     f"концовка)")
        own = dict(zip(names, files))
        log("живой голос:", ", ".join(f"{n}={p.name}" for n, p in own.items()))

    def vo(text, name):
        if own:
            # ускоряем только пояснения и концовку: вопросы на длину ролика не влияют (таймер 10 с)
            tempo = a.voice_tempo if name[0] in "eo" else 1.0
            return own_voice(own[name], work, name, {"pitch": a.voice_pitch, "tune": a.voice_tune,
                                                     "formant": a.voice_formant, "tempo": tempo})
        h = hashlib.md5(f"{text}|{a.speed}".encode()).hexdigest()[:8]
        return say(text, voice, work / f"{name}_{a.voice}_{h}.wav", a.speed)
    intro_voice = vo(quiz.get("intro_voice") or f"Квиз про {quiz['topic']}! Проверь себя!", "intro")
    clips, vclips, scenes = [], [], []  # vclips — голос (под него приглушаем музыку)

    d = len(intro_voice) / SR
    scenes.append({"type": "intro", "start": 0.0, "end": max(0.3 + d + 0.4, 2.5, 0.3 + read_time(quiz["hook"]) * 0.6),
                   "icons": quiz.get("icons") or []})
    clips += [(whoosh(1), 0.05), (pop(), 0.25)]
    vclips += [(intro_voice, 0.3)]
    t = scenes[-1]["end"]

    report = []
    for i, q in enumerate(qs):
        k = round(i * (len(LEVELS) - 1) / max(len(qs) - 1, 1))
        level_name, level_color = q.get("level") or LEVELS[k][0], LEVELS[k][1]
        qv = vo(q.get("q_voice") or q["q"], f"q{i + 1}")
        ev = vo(q.get("explain_voice") or q["explain"], f"e{i + 1}")
        dq, de = len(qv) / SR, len(ev) / SR
        # таймер идёт с появления вопроса, пока голос его читает (владелец 11.11: 10 с, 5 вопросов, ≤ 90 с)
        t_timer = 0.3
        t_reveal = max(t_timer + a.timer, 0.3 + dq + 1.5)
        t_flip = t_reveal + 0.5
        t_opts = [0.6 + j * 0.25 for j in range(len(q["options"]))]
        s = {"type": "q", "idx": i, "total": len(qs), "q": q["q"], "options": q["options"], "answer": q["answer"],
             "level": level_name, "level_n": k + 1, "level_max": len(LEVELS), "level_color": level_color,
             "explain": q["explain"], "explain_title": q.get("explain_title", "Пояснение"),
             "icons": q.get("icons") or quiz.get("icons") or [],
             "image": images[i].resolve().as_uri(), "timer": t_reveal - t_timer,
             "t_opts": t_opts, "t_timer": t_timer, "t_reveal": t_reveal, "t_flip": t_flip,
             "start": t, "end": t + t_flip + 0.35 + max(de + 0.3, read_time(q["explain"]))}
        scenes.append(s)
        report.append(f"вопрос {i + 1}: читается {dq:.1f} с, пояснение {de:.1f} с (чтение {read_time(q['explain']):.1f} с), "
                      f"всего {s['end'] - t:.1f} с")
        clips += [(whoosh(i + 2), t)]
        vclips += [(qv, t + 0.3), (ev, t + t_flip + 0.35)]
        clips += [(pop(), t + x) for x in t_opts]
        n = int(round(t_reveal - t_timer))
        clips += [(tick(j >= n - 3), t + t_reveal - n + j) for j in range(n)]
        clips += [(ding(), t + t_reveal), (whoosh(50 + i), t + t_flip)]
        t = s["end"]

    ov = vo(quiz.get("outro_voice") or "Сколько угадал? Пиши в комментариях!", "outro")
    end = max(t + 0.3 + len(ov) / SR + 0.7, a.min_duration)
    scenes.append({"type": "outro", "start": t, "end": end, "icons": quiz.get("icons") or []})
    clips += [(whoosh(99), t), (ding(), t + 0.3)]
    vclips += [(ov, t + 0.3)]
    total = scenes[-1]["end"]
    log(f"длина ролика {total:.1f} с; " + "; ".join(report))
    if total > a.max_duration:
        log(f"ВНИМАНИЕ: ролик {total:.1f} с длиннее ориентира {a.max_duration:.0f} с (не ошибка — владелец 12.11)")

    for j, (prompt, w, h, p, sd_seed) in enumerate(jobs):
        if model_ok and not a.preview and not a.audio_only:
            painter.paint(prompt, w, h, p, sd_seed, a.steps)
        elif not p.exists():  # заглушка под другим именем, чтобы не принять её потом за готовую картинку
            stub = p.with_name(p.stem + "_stub.png")
            placeholder(stub, w, h, pal["accent"])
            if j == 0:
                cover = stub
            else:
                images[j - 1] = stub
    for sc in scenes:
        if sc["type"] == "q":
            sc["image"] = images[sc["idx"]].resolve().as_uri()

    tl = {"fps": 30, "duration": total, "seed": seed % 1000, "palette": pal, "topic": quiz["topic"],
          "hook": quiz["hook"], "cover": cover.resolve().as_uri(), "scenes": scenes,
          "outro_title": quiz.get("outro_title", "Сколько угадал?"),
          "outro_score": quiz.get("outro_score", f"?/{len(qs)}"),
          "outro_sub": quiz.get("outro_sub", "Пиши результат в комментариях 👇")}
    tl_path = work / "timeline.json"
    tl_path.write_text(json.dumps(tl, ensure_ascii=False, indent=1), encoding="utf-8")

    if a.preview:
        subprocess.run(["node", str(HERE / "render.js"), str(tl_path), str(work / "preview"), "--preview", a.preview],
                       check=True)
        log("кадры:", work)
        return

    track = np.zeros(int(SR * (total + 1)), np.float32)
    vtrack = np.zeros_like(track)
    for clip, at in vclips:
        mix(vtrack, clip, at)
    if not a.no_music:
        m = (music_calm if a.music == "calm" else music)(total)[:len(track)]
        rms = float(np.sqrt(np.mean(m[m != 0] ** 2))) or 1.0
        m = m * (10 ** (a.music_db / 20) / rms)        # музыка тихая и одинаковая по громкости
        track += duck(m, vtrack, a.duck_db)            # и ещё тише, пока говорит голос
    for clip, at in clips:
        mix(track, clip, at)
    track += vtrack
    track = np.clip(track, -1, 1)
    audio = work / "audio.wav"
    with wave.open(str(audio), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(SR)
        w.writeframes((track * 32767).astype(np.int16).tobytes())

    if a.audio_only:
        subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", str(audio), "-af",
                        f"loudnorm=I={a.loudness}:TP=-1.5:LRA=11", "-ar", "44100", "-b:a", "160k",
                        str(work / "audio.mp3")], check=True)
        log("звук:", work / "audio.mp3")
        return

    video = work / "video.mp4"
    log(f"рендер кадров: {total:.1f} с ролика…")
    subprocess.run(["node", str(HERE / "render.js"), str(tl_path), str(video)], check=True)
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", str(video), "-i", str(audio), "-map", "0:v", "-map", "1:a",
                    "-c:v", "copy", "-af", f"loudnorm=I={a.loudness}:TP=-1.5:LRA=11", "-ar", "44100", "-c:a", "aac",
                    "-b:a", "160k", "-shortest", "-movflags", "+faststart", str(out)], check=True)
    log(f"готово: {out} ({out.stat().st_size / 1e6:.1f} МБ, {total:.0f} с)")


if __name__ == "__main__":
    main()
