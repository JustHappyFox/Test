"""Объективная оценка речи (я звук не слышу): DNSMOS P.835 (SIG — сама речь, BAK — фон/шум, OVRL — общее, 1–5)
и SRMR (сухость: выше — меньше эха). Использование: voice_qa.py МЕТКА=шаблон_файлов …  (нужны speechmos, librosa, scipy)"""
import glob
import subprocess
import sys
import warnings

import numpy as np
from scipy import signal

warnings.filterwarnings("ignore")
from speechmos import dnsmos  # noqa: E402


def dec(f, sr=16000):
    raw = subprocess.run(["ffmpeg", "-v", "error", "-i", f, "-ac", "1", "-ar", str(sr), "-f", "f32le", "-"],
                         capture_output=True, check=True).stdout
    return np.frombuffer(raw, np.float32).astype(np.float64)


def srmr(x, sr=16000):
    """Упрощённый SRMR (Falk 2010): доля медленных модуляций огибающей (4–18 Гц, речь) к быстрым (29–128 Гц, эхо/шум)."""
    erb = lambda f: 24.7 * (4.37 * f / 1000 + 1)
    cfs, f = [], 125.0
    while f < 0.45 * sr and len(cfs) < 23:
        cfs.append(f); f += erb(f) * 1.0
    mod_cf = 4.0 * (128 / 4.0) ** (np.arange(8) / 7)
    E = np.zeros((len(cfs), 8))
    for i, cf in enumerate(cfs):
        b, a = signal.gammatone(cf, "iir", fs=sr)
        env = np.abs(signal.hilbert(signal.lfilter(b, a, x)))
        for k, mf in enumerate(mod_cf):
            bw = mf / 2
            sos = signal.butter(1, [max(0.5, mf - bw / 2), mf + bw / 2], "bandpass", fs=sr, output="sos")
            E[i, k] = np.mean(signal.sosfilt(sos, env) ** 2)
    return E[:, :4].sum() / (E[:, 4:].sum() + 1e-12)


def score(files):
    rows = []
    for f in files:
        x = dec(f)
        x = x / (np.abs(x).max() + 1e-9) * 0.7
        d = dnsmos.run(x.astype(np.float32), sr=16000)
        rows.append([d["ovrl_mos"], d["sig_mos"], d["bak_mos"], d["p808_mos"], srmr(x)])
    return np.array(rows)


if __name__ == "__main__":
    print(f"{'':26} OVRL   SIG   BAK  P808  SRMR   (n)")
    for arg in sys.argv[1:]:
        label, pat = arg.split("=", 1)
        files = sorted(glob.glob(pat))
        if not files:
            print(f"{label:26} нет файлов"); continue
        r = score(files)
        m = r.mean(0)
        print(f"{label:26} {m[0]:4.2f}  {m[1]:4.2f}  {m[2]:4.2f}  {m[3]:4.2f}  {m[4]:5.2f}   ({len(files)}, мин OVRL {r[:,0].min():.2f})")
