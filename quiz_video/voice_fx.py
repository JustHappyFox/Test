"""Обработка живого голоса для квиз-роликов: чистый, близкий и звонкий голос из записи на телефон.
Рецепт одобрен владельцем 13.11 — подробно, с запретами и проверкой: VOICE.md рядом с этим файлом.

  python voice_fx.py ВХОД ВЫХОД.wav [--pitch 1.0] [--tune 0] [--tempo 1.0] [--deep-filter ПУТЬ]

Владелец 13.11 про «Машину»: «с какими-то шумами, как будто в ванной, а должен быть чистый и звонкий. Как будто
говорит человек в комнате со мной, а там как будто разговаривает из другой комнаты».
Что показали измерения (voice_qa.py — DNSMOS, оценки 1–5; слух у меня заменяют они):
  запись Крис как есть      общее 2,47  речь 3,22  фон 3,05
  прежняя обработка         общее 2,00  речь 3,17  фон 2,17   — стало ХУЖЕ записи
  эта цепочка               общее 3,13  речь 3,47  фон 3,92   (чистый синтез Piper: 3,30 / 3,56 / 4,06)
Причины и что сделано:
  * шум и «ванная»: компрессор поднимал хвосты эха между слогами, afftdn/noisereduce давали призвук →
    нейросеть DeepFilterNet3 (убирает шум и хвосты эха), компрессора нет вовсе;
  * «из другой комнаты»: в записи почти нет верхов (1,3–10 кГц ниже нормы на 10–15 дБ), а я ещё понижал голос →
    эквалайзер подтягивает спектр к норме женской речи; «немного изменить голос» (просьба клиента) теперь
    сдвигом ВВЕРХ на 1 полутон вместе с тембром (простой пересчёт частоты — без вокодера, качество не падает);
  * автотюн (PSOLA) стоит ~0,1 балла — по умолчанию выключен (--tune 0.3 включает).
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

import numpy as np

FS = 48000
HERE = Path(__file__).resolve().parent
# Спектр к норме женской речи (после DeepFilterNet3 и сдвига +1): по полосам в пределах ~3 дБ от нормы.
EQ = ["highpass=f=90:r=f64",
      "equalizer=f=200:t=q:w=0.8:g=-1:r=f64",
      "equalizer=f=650:t=q:w=1.0:g=-3:r=f64",      # «гнусавость»
      "equalizer=f=1000:t=q:w=1.4:g=-4:r=f64",     # «зажатый нос», «коробка»
      "equalizer=f=1900:t=q:w=0.9:g=8:r=f64",      # открыть голос: в записи тут провал
      "equalizer=f=3300:t=q:w=1.0:g=1:r=f64",
      "highshelf=f=5000:g=6:r=f64"]                # звонкость и «воздух»


def decode(src: Path, sr: int = FS) -> np.ndarray:
    raw = subprocess.run(["ffmpeg", "-v", "error", "-i", str(src), "-ac", "1", "-ar", str(sr), "-f", "f32le", "-"],
                         capture_output=True, check=True).stdout
    return np.frombuffer(raw, np.float32).copy()


def ffilter(x: np.ndarray, af: str, sr: int = FS, out_sr: int | None = None) -> np.ndarray:
    """Пропустить сигнал через фильтры ffmpeg."""
    p = subprocess.run(["ffmpeg", "-v", "error", "-f", "f32le", "-ar", str(sr), "-ac", "1", "-i", "-",
                        "-af", af, "-ar", str(out_sr or sr), "-ac", "1", "-f", "f32le", "-"],
                       input=x.astype(np.float32).tobytes(), capture_output=True, check=True).stdout
    return np.frombuffer(p, np.float32).copy()


def trim(x: np.ndarray, sr: int = FS, thr_db: float = -38, pad: float = 0.06) -> np.ndarray:
    """Обрезать тишину и шорохи по краям (по уровню громкости в окнах 20 мс)."""
    w = int(sr * 0.02)
    n = len(x) // w
    if n == 0:
        return x
    rms = np.sqrt(np.mean(x[:n * w].reshape(n, w) ** 2, axis=1) + 1e-12)
    db = 20 * np.log10(rms / (rms.max() + 1e-12))
    on = np.where(db > thr_db)[0]
    if len(on) == 0:
        return x
    a = max(0, on[0] * w - int(pad * sr))
    b = min(len(x), (on[-1] + 1) * w + int(pad * 1.5 * sr))
    y = x[a:b].copy()
    f = int(0.01 * sr)
    y[:f] *= np.linspace(0, 1, f)
    y[-f:] *= np.linspace(1, 0, f)
    return y


def find_deep_filter(path: str | None = None) -> str | None:
    """Программа deep-filter (DeepFilterNet3, модель внутри): аргумент, $DEEP_FILTER, PATH или tools/ рядом."""
    for c in (path, os.environ.get("DEEP_FILTER"), shutil.which("deep-filter"), HERE / "tools" / "deep-filter"):
        if c and Path(c).is_file():
            return str(c)
    return None


def deep_filter(x: np.ndarray, binary: str) -> np.ndarray:
    """Шум и хвосты эха — нейросетью DeepFilterNet3 (48 кГц). Скачанную программу запускаем без сети."""
    import soundfile as sf
    with tempfile.TemporaryDirectory() as d:
        sf.write(f"{d}/a.wav", x, FS, subtype="PCM_16")
        cmd = [binary, "-D", "-o", f"{d}/o", f"{d}/a.wav"]
        if shutil.which("unshare"):
            cmd = ["unshare", "-r", "-n"] + cmd
        subprocess.run(cmd, check=True, capture_output=True)
        y, _ = sf.read(f"{d}/o/a.wav", dtype="float32")
    return y


def autotune(x: np.ndarray, sr: int, tune: float) -> np.ndarray:
    """Лёгкий «автотюн» через Praat (PSOLA): высота — к ближайшей ноте на долю tune. Стоит ~0,1 балла качества."""
    import parselmouth
    from parselmouth.praat import call
    snd = parselmouth.Sound(x.astype(np.float64), sampling_frequency=sr)
    man = call(snd, "To Manipulation", 0.01, 75, 600)
    tier = call(man, "Extract pitch tier")
    n = int(call(tier, "Get number of points"))
    if not n:
        return x
    ts = [call(tier, "Get time from index", i + 1) for i in range(n)]
    m = 12 * np.log2(np.array([call(tier, "Get value at index", i + 1) for i in range(n)]) / 440.0)
    m = m + tune * (np.round(m) - m)
    call(tier, "Remove points between", 0, snd.duration + 1)
    for t, v in zip(ts, m):
        call(tier, "Add point", t, 440.0 * 2 ** (v / 12))
    call([tier, man], "Replace pitch tier")
    return np.asarray(call(man, "Get resynthesis (overlap-add)").values[0], np.float32)


def process(src: Path, out: Path, out_sr: int, pitch: float = 1.0, tune: float = 0.0, tempo: float = 1.0,
            deepfilter: str | None = None) -> Path:
    """pitch — сдвиг высоты вместе с тембром, полутона (+ звонче, − глуше); tune — автотюн 0…1; tempo — скорость."""
    import soundfile as sf
    x = ffilter(decode(src), "highpass=f=70")
    binary = find_deep_filter(deepfilter)
    if binary:
        x = deep_filter(x, binary)
    else:  # запасной вариант заметно хуже — предупреждаем
        print("ВНИМАНИЕ: deep-filter не найден — шумы чистит простой фильтр, звук будет хуже", file=sys.stderr)
        x = ffilter(x, "afftdn=nr=10:nf=-50:tn=1")
    x = trim(x)
    x = x / (np.abs(x).max() + 1e-9) * 0.5   # запас по громкости: подъёмы эквалайзера без перегруза
    if tune:
        x = autotune(x, FS, tune)
    chain = []
    ratio = 2 ** (pitch / 12)
    if pitch:      # высота и тембр вместе: пересчёт частоты, длительность возвращаем
        chain += [f"asetrate={FS * ratio:.0f}", f"aresample={FS}"]
    if pitch or tempo != 1.0:
        chain.append(f"atempo={tempo / ratio:.5f}")
    y = ffilter(x, ",".join(chain + EQ), FS, out_sr)   # компрессора нет намеренно: он поднимает эхо и шум
    y = trim(y, out_sr, thr_db=-45, pad=0.04)
    # ровная громкость клипов: по RMS звонкой части, пики — мягким ограничителем
    w = int(out_sr * 0.05)
    n = max(1, len(y) // w)
    rms = np.sqrt(np.mean(y[:n * w].reshape(n, w) ** 2, axis=1) + 1e-12)
    loud = rms[rms > rms.max() * 0.25]
    y = y * (10 ** (-18 / 20) / (np.sqrt(np.mean(loud ** 2)) + 1e-9))
    y = ffilter(y, "alimiter=limit=0.89:attack=5:release=60:level=disabled", out_sr)
    sf.write(out, np.clip(y, -1, 1), out_sr, subtype="PCM_16")
    return out


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("src")
    ap.add_argument("out")
    ap.add_argument("--sr", type=int, default=44100)
    ap.add_argument("--pitch", type=float, default=1.0, help="сдвиг высоты и тембра, полутона")
    ap.add_argument("--tune", type=float, default=0.0, help="автотюн 0…1 (0 — выключен)")
    ap.add_argument("--tempo", type=float, default=1.0)
    ap.add_argument("--deep-filter", help="путь к программе deep-filter (DeepFilterNet)")
    a = ap.parse_args()
    process(Path(a.src), Path(a.out), a.sr, a.pitch, a.tune, a.tempo, a.deep_filter)
