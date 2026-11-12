"""Обработка живого голоса для квиз-роликов (владелец 12.11: «нормальная интонация, отфильтровать шумы и
посторонние звуки, немного заавтотюнить»; клиент: «немного изменишь голос»).

  python voice_fx.py ВХОД ВЫХОД.wav [--pitch -1.5] [--tune 0.4] [--formant 1.0] [--tempo 1.0]

Цепочка: срез низа, щелчки → мягкое шумоподавление (afftdn) → обрезка тишины по краям →
Praat (PSOLA): лёгкая подтяжка высоты к ближайшей ноте («автотюн»), сдвиг высоты, при нужде — ускорение →
эквалайзер (подъём 1,3–3,5 кГц против «зажатого носа»), мягкий компрессор → ровная громкость (RMS −18 дБ).
Проверка 12.11 распознаванием речи: WORLD-вокодер и сдвиг тембра (formant) портят разборчивость,
Praat с formant=1.0 — нет, поэтому по умолчанию Praat без сдвига тембра.
"""
import argparse
import subprocess
from pathlib import Path

import numpy as np

FS = 48000


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


def retune(x: np.ndarray, sr: int, pitch: float, tune: float, formant: float, expand: float) -> np.ndarray:
    """WORLD: высота голоса — к ближайшей ноте на долю tune, сдвиг на pitch полутонов, тембр × formant."""
    import pyworld as pw
    xd = x.astype(np.float64)
    f0, t = pw.harvest(xd, sr, f0_floor=70, f0_ceil=700, frame_period=5.0)
    sp = pw.cheaptrick(xd, f0, t, sr)
    ap = pw.d4c(xd, f0, t, sr)
    voiced = f0 > 0
    if voiced.any():
        m = 12 * np.log2(f0[voiced] / 440.0)
        center = np.median(m)
        m = center + (m - center) * expand          # интонация: чуть живее, если чтение монотонное
        m = m + tune * (np.round(m) - m)            # «автотюн»: подтянуть к ноте, но не до робота
        m = m + pitch
        f0n = f0.copy()
        f0n[voiced] = 440.0 * 2 ** (m / 12)
        # сгладить ступеньки между кадрами (медиана по 5 кадрам внутри звонких участков)
        from scipy.ndimage import median_filter
        sm = median_filter(f0n, size=5)
        f0n = np.where(voiced & (sm > 0), sm, f0n)
    else:
        f0n = f0
    if formant != 1.0:  # сдвиг формант: растягиваем огибающую спектра по частоте
        k = np.arange(sp.shape[1])
        src = np.clip(k / formant, 0, sp.shape[1] - 1)
        sp = np.stack([np.interp(src, k, row) for row in sp])
    y = pw.synthesize(f0n, sp, ap, sr, frame_period=5.0)
    return y.astype(np.float32)


def retune_praat(x: np.ndarray, sr: int, pitch: float, tune: float, formant: float, expand: float,
                 tempo: float = 1.0) -> np.ndarray:
    """То же через Praat (PSOLA): согласные не трогаются, речь разборчивее, чем через WORLD (проверка 12.11)."""
    import parselmouth
    from parselmouth.praat import call
    snd = parselmouth.Sound(x.astype(np.float64), sampling_frequency=sr)
    man = call(snd, "To Manipulation", 0.01, 75, 600)
    tier = call(man, "Extract pitch tier")
    n = int(call(tier, "Get number of points"))
    if n and (pitch or tune or expand != 1.0):
        ts = [call(tier, "Get time from index", i + 1) for i in range(n)]
        fs = np.array([call(tier, "Get value at index", i + 1) for i in range(n)])
        m = 12 * np.log2(fs / 440.0)
        c = np.median(m)
        m = c + (m - c) * expand                      # интонация: чуть живее, если чтение монотонное
        m = m + tune * (np.round(m) - m)              # «автотюн»: подтянуть к ноте, но не до робота
        m = m + pitch
        call(tier, "Remove points between", 0, snd.duration + 1)
        for t, v in zip(ts, m):
            call(tier, "Add point", t, 440.0 * 2 ** (v / 12))
        call([tier, man], "Replace pitch tier")
    if tempo != 1.0:  # ускорение тем же PSOLA — разборчивее, чем atempo
        dt = call("Create DurationTier", "dur", 0, snd.duration)
        call(dt, "Add point", 0, 1 / tempo)
        call([man, dt], "Replace duration tier")
    y = call(man, "Get resynthesis (overlap-add)")
    if formant != 1.0:  # тембр: формантами вниз/вверх, высоту и длительность не меняем
        y = call(y, "Change gender", 75, 600, formant, 0, 1, 1)
    return np.asarray(y.values[0], np.float32)


def process(src: Path, out: Path, out_sr: int, pitch: float = -1.5, tune: float = 0.4, formant: float = 1.0,
            expand: float = 1.0, tempo: float = 1.0, denoise: str = "afftdn", engine: str = "praat") -> Path:
    """Владелец 12.11 после «Машины»: «голос приглушённый, как будто с зажатым носом». Разбор спектра:
    в её записях (микрофон телефона) провал 1,3–2 кГц на ~10 дБ против чистого голоса и много низов;
    noisereduce добавлял «подводный» призвук. Поэтому: мягкий шумодав ffmpeg, подъём 1,3–3,5 кГц,
    меньше низов, фильтры с двойной точностью, громкость — мягким ограничителем, без обрезки."""
    import soundfile as sf
    x = decode(src)
    x = ffilter(x, "highpass=f=80,adeclick")
    if denoise == "afftdn":    # мягко: гул и шипение, голос не трогаем
        x = ffilter(x, "afftdn=nr=10:nf=-50:tn=1")
    elif denoise == "nr":      # старый вариант (noisereduce) — даёт призвук, не по умолчанию
        import noisereduce as nr
        x = nr.reduce_noise(y=x, sr=FS, stationary=True, prop_decrease=0.7, n_fft=2048).astype(np.float32)
    x = trim(x)
    x = x / (np.abs(x).max() + 1e-9) * 0.5   # запас по громкости: подъёмы эквалайзера без перегруза
    if engine == "praat":
        x = retune_praat(x, FS, pitch, tune, formant, expand, tempo)
    elif pitch or tune or formant != 1.0 or expand != 1.0:
        x = retune(x, FS, pitch, tune, formant, expand)
    eq = ["highpass=f=90:r=f64",
          "equalizer=f=200:t=q:w=0.8:g=-3:r=f64",     # меньше «бубнения» и гула
          "equalizer=f=700:t=q:w=1.2:g=-3:r=f64",     # «гнусавость»: избыток 500–800 Гц
          "equalizer=f=1600:t=q:w=0.8:g=8:r=f64",     # открыть голос: провал 1,3–2 кГц — тот самый «нос»
          "equalizer=f=3200:t=q:w=1.0:g=3:r=f64",     # разборчивость согласных
          "highshelf=f=7000:g=2:r=f64",               # немного воздуха
          "acompressor=threshold=-26dB:ratio=2.5:attack=10:release=150"]
    y = ffilter(x, ",".join(eq), FS, out_sr)
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
    ap.add_argument("--pitch", type=float, default=-1.5)
    ap.add_argument("--tune", type=float, default=0.4)
    ap.add_argument("--formant", type=float, default=1.0)
    ap.add_argument("--expand", type=float, default=1.0)
    ap.add_argument("--tempo", type=float, default=1.0)
    ap.add_argument("--denoise", choices=["afftdn", "nr", "none"], default="afftdn")
    a = ap.parse_args()
    process(Path(a.src), Path(a.out), a.sr, a.pitch, a.tune, a.formant, a.expand, a.tempo, a.denoise)
