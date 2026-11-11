"""Квиз-ролик по теме: картинки (Z-Image, локально), озвучка (Piper, локально), кадры (Chromium), сборка (ffmpeg).

  python make_quiz.py quiz.json --out ролик.mp4 --assets ПАПКА_С_МОДЕЛЯМИ

quiz.json:
  {"topic": "Луна", "hook": "…", "cover_prompt": "англ. описание картинки",
   "icons": ["🍫", "🍬"],
   "questions": [{"level": "Лёгкий|Средний|Сложный", "q": "…", "options": ["…", "…", "…"], "answer": 0,
                  "explain": "пояснение на обороте карточки", "icons": ["🌱"], "image_prompt": "…",
                  "q_voice"/"answer_voice"/"explain_voice": "как читать вслух (числа словами)"}],
   "outro_title": "…", "outro_sub": "…", "outro_voice": "…", "palette": 0}
Сложность растёт: лёгкий → средний → сложный; 10 с на размышление; после ответа карточка переворачивается
и голос читает пояснение; ролик не короче 62 с (пожелания владельца 11.11).

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
SR = 22050  # частота голосов Piper medium

PALETTES = [
    {"bg1": "#ff8a3d", "bg2": "#ffc145", "accent": "#1fb6e8"},   # оранжевая
    {"bg1": "#6a3df0", "bg2": "#c04bf4", "accent": "#ff9f1c"},   # фиолетовая
    {"bg1": "#0f9b8e", "bg2": "#46e08a", "accent": "#ff5e7e"},   # бирюзовая
    {"bg1": "#2448a8", "bg2": "#3fa9f5", "accent": "#ffb703"},   # синяя
    {"bg1": "#ff5f9e", "bg2": "#ffa58b", "accent": "#4d5bff"},   # розовая
]
LEVELS = {
    "Лёгкий": {"n": 1, "color": "#22c55e", "say": "Лёгкий вопрос."},
    "Средний": {"n": 2, "color": "#f59e0b", "say": "Вопрос посложнее."},
    "Сложный": {"n": 3, "color": "#ef4444", "say": "А теперь сложный вопрос."},
}
LEVEL_ORDER = ["Лёгкий", "Средний", "Сложный"]
STYLE = ("vibrant flat vector cartoon illustration, bold clean outlines, playful bright colors, "
         "simple background, centered composition, high quality, no text, no letters, no words")


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
        subprocess.run([sys.executable, "-m", "piper", "-m", str(voice), "-f", str(out),
                        "--length-scale", str(speed), "--sentence-silence", "0.25"],
                       input=text.encode(), check=True, capture_output=True)
    with wave.open(str(out)) as w:
        assert w.getframerate() == SR, w.getframerate()
        a = np.frombuffer(w.readframes(w.getnframes()), dtype=np.int16).astype(np.float32) / 32768
    peak = float(np.abs(a).max()) or 1.0
    return a * (0.85 / peak)


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
    ap.add_argument("--speed", type=float, default=0.92, help="темп речи Piper: меньше — быстрее")
    ap.add_argument("--timer", type=float, default=10.0, help="секунд на размышление")
    ap.add_argument("--min-duration", type=float, default=62.0, help="ролик не короче, с")
    ap.add_argument("--steps", type=int, default=8)
    ap.add_argument("--threads", type=int, default=4)
    ap.add_argument("--no-images", action="store_true", help="без нейросети: заглушки вместо картинок")
    ap.add_argument("--no-music", action="store_true")
    ap.add_argument("--preview", help="только кадры PNG в эти секунды, через запятую")
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

    # картинки
    painter = Painter(assets, a.threads)
    model_ok = not a.no_images and (assets / "models" / "z_image_turbo-Q4_0.gguf").exists()

    def picture(prompt, w, h, name, s):
        p = work / f"{name}_{hashlib.md5(f'{prompt}|{w}x{h}|{a.steps}|{s}'.encode()).hexdigest()[:8]}.png"
        if model_ok:
            return painter.paint(prompt, w, h, p, s, a.steps)
        return placeholder(p, w, h, pal["accent"])

    cover = picture(quiz["cover_prompt"], 768, 768, "cover", seed)
    images = [picture(q["image_prompt"], 768, 576, f"q{i + 1}", seed + i + 1) for i, q in enumerate(qs)]

    # голос
    voice = assets / "voices" / f"ru_RU-{a.voice}-medium.onnx"
    vo = lambda text, name: say(text, voice, work / f"{name}_{a.voice}_{hashlib.md5(text.encode()).hexdigest()[:8]}.wav", a.speed)
    intro_voice = vo(quiz.get("intro_voice") or f"Квиз. {quiz['topic']}. {quiz['hook']}", "intro")
    letters = "АБВГ"
    clips, scenes, t = [], [], 0.0

    d = len(intro_voice) / SR
    scenes.append({"type": "intro", "start": 0.0, "end": max(0.6 + d + 0.7, 4.5)})
    clips += [(whoosh(1), 0.05), (pop(), 0.25), (intro_voice, 0.6)]
    t = scenes[-1]["end"]

    for i, q in enumerate(qs):
        level_name = q.get("level") or LEVEL_ORDER[min(i * len(LEVEL_ORDER) // len(qs), len(LEVEL_ORDER) - 1)]
        lvl = LEVELS[level_name]
        qv = vo(f"{lvl['say']} {q.get('q_voice') or q['q']}", f"q{i + 1}")
        right = q["options"][q["answer"]]
        av = vo(q.get("answer_voice") or f"Правильный ответ: {right}.", f"a{i + 1}")
        ev = vo(q.get("explain_voice") or q["explain"], f"e{i + 1}")
        dq, da, de = len(qv) / SR, len(av) / SR, len(ev) / SR
        t_timer = 0.5 + dq + 0.35
        t_reveal = t_timer + a.timer
        t_flip = t_reveal + 0.45 + da + 0.3
        t_opts = [min(0.8 + j * 0.3, t_timer - 0.6 + j * 0.12) for j in range(len(q["options"]))]
        s = {"type": "q", "idx": i, "total": len(qs), "q": q["q"], "options": q["options"], "answer": q["answer"],
             "level": level_name, "level_n": lvl["n"], "level_color": lvl["color"],
             "explain": q["explain"], "explain_title": q.get("explain_title", "Пояснение"),
             "icons": q.get("icons") or quiz.get("icons") or [],
             "image": images[i].resolve().as_uri(), "timer": a.timer,
             "t_opts": t_opts, "t_timer": t_timer, "t_reveal": t_reveal, "t_flip": t_flip,
             "start": t, "end": t + t_flip + 0.8 + de + 1.3}
        scenes.append(s)
        clips += [(whoosh(i + 2), t), (qv, t + 0.5)]
        clips += [(pop(), t + x) for x in t_opts]
        n = int(a.timer)
        clips += [(tick(k >= n - 3), t + t_timer + k) for k in range(n)]
        clips += [(ding(), t + t_reveal), (av, t + t_reveal + 0.45)]
        clips += [(whoosh(50 + i), t + t_flip), (ev, t + t_flip + 0.8)]
        t = s["end"]

    ov = vo(quiz.get("outro_voice") or "Сколько правильных ответов у тебя? Напиши в комментариях!", "outro")
    end = max(t + 0.4 + len(ov) / SR + 1.8, a.min_duration)
    scenes.append({"type": "outro", "start": t, "end": end, "icons": quiz.get("icons") or []})
    clips += [(whoosh(99), t), (ding(), t + 0.5), (ov, t + 0.4)]
    total = scenes[-1]["end"]
    scenes[0]["icons"] = quiz.get("icons") or []

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
    if not a.no_music:
        m = music(total)
        track += m[:len(track)]
    for clip, at in clips:
        mix(track, clip, at)
    track = np.clip(track, -1, 1)
    audio = work / "audio.wav"
    with wave.open(str(audio), "wb") as w:
        w.setnchannels(1)
        w.setsampwidth(2)
        w.setframerate(SR)
        w.writeframes((track * 32767).astype(np.int16).tobytes())

    video = work / "video.mp4"
    log(f"рендер кадров: {total:.1f} с ролика…")
    subprocess.run(["node", str(HERE / "render.js"), str(tl_path), str(video)], check=True)
    subprocess.run(["ffmpeg", "-y", "-loglevel", "error", "-i", str(video), "-i", str(audio), "-map", "0:v", "-map", "1:a",
                    "-c:v", "copy", "-af", "loudnorm=I=-14:TP=-1.5:LRA=11", "-ar", "44100", "-c:a", "aac",
                    "-b:a", "160k", "-shortest", "-movflags", "+faststart", str(out)], check=True)
    log(f"готово: {out} ({out.stat().st_size / 1e6:.1f} МБ, {total:.0f} с)")


if __name__ == "__main__":
    main()
