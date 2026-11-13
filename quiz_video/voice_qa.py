"""Объективная оценка голоса (я звук не слышу — вместо слуха измерения).

  python voice_qa.py МЕТКА=шаблон_файлов …            — DNSMOS и SRMR по группам файлов
  python voice_qa.py --spk МОДЕЛЬ.onnx ОБРАЗЕЦ_шаблон МЕТКА=шаблон …   — сходство голоса с образцом

DNSMOS P.835, шкала 1–5: SIG — сама речь, BAK — фон/шум, OVRL — общее. SRMR — сухость (выше — меньше эха).
Сходство голоса — косинус эмбеддингов диктора WeSpeaker ResNet34 (onnx с HuggingFace
Wespeaker/wespeaker-voxceleb-resnet34-LM): записи Крис между собой ≈ 0,82, чужой голос ≈ 0,18.
Нужны: speechmos, librosa, scipy; для сходства — onnxruntime, kaldi-native-fbank.
"""
import glob
import subprocess
import sys
import warnings

import numpy as np

warnings.filterwarnings("ignore")


def dec(f, sr=16000):
    raw = subprocess.run(["ffmpeg", "-v", "error", "-i", str(f), "-ac", "1", "-ar", str(sr), "-f", "f32le", "-"],
                         capture_output=True, check=True).stdout
    return np.frombuffer(raw, np.float32).astype(np.float64)


def srmr(x, sr=16000):
    """Упрощённый SRMR (Falk 2010): доля медленных модуляций огибающей (4–18 Гц, речь) к быстрым (29–128 Гц, эхо/шум)."""
    from scipy import signal
    erb = lambda f: 24.7 * (4.37 * f / 1000 + 1)
    cfs, f = [], 125.0
    while f < 0.45 * sr and len(cfs) < 23:
        cfs.append(f)
        f += erb(f)
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
    """По файлу: [OVRL, SIG, BAK, P808, SRMR]."""
    from speechmos import dnsmos
    rows = []
    for f in files:
        x = dec(f)
        x = x / (np.abs(x).max() + 1e-9) * 0.7
        d = dnsmos.run(x.astype(np.float32), sr=16000)
        rows.append([d["ovrl_mos"], d["sig_mos"], d["bak_mos"], d["p808_mos"], srmr(x)])
    return np.array(rows)


_spk = {}


def embed(f, model):
    """Эмбеддинг диктора (единичной длины)."""
    import kaldi_native_fbank as knf
    import onnxruntime as ort
    if model not in _spk:
        _spk[model] = ort.InferenceSession(str(model), providers=["CPUExecutionProvider"])
    sess = _spk[model]
    x = dec(f)
    o = knf.FbankOptions()
    o.frame_opts.dither = 0.0
    o.frame_opts.samp_freq = 16000
    o.mel_opts.num_bins = 80
    fb = knf.OnlineFbank(o)
    fb.accept_waveform(16000, (x * 32768).tolist())
    fb.input_finished()
    feats = np.stack([fb.get_frame(i) for i in range(fb.num_frames_ready)]).astype(np.float32)
    feats -= feats.mean(0, keepdims=True)
    e = sess.run(None, {sess.get_inputs()[0].name: feats[None]})[0][0]
    return e / np.linalg.norm(e)


def centroid(files, model):
    E = np.stack([embed(f, model) for f in files])
    c = E.mean(0)
    return c / np.linalg.norm(c), E


if __name__ == "__main__":
    args = sys.argv[1:]
    if args and args[0] == "--spk":
        model, ref = args[1], sorted(glob.glob(args[2]))
        c, E = centroid(ref, model)
        loo = [float(E[i] @ (lambda m: m / np.linalg.norm(m))(np.delete(E, i, 0).mean(0))) for i in range(len(E))]
        print(f"{'образец: свой–свой':30} {np.mean(loo):.3f} (мин {np.min(loo):.3f}, n={len(E)})")
        for arg in args[3:]:
            label, pat = arg.split("=", 1)
            sims = [float(embed(f, model) @ c) for f in sorted(glob.glob(pat))]
            print(f"{label:30} {np.mean(sims):.3f} (мин {np.min(sims):.3f}, n={len(sims)})" if sims else f"{label:30} нет файлов")
        sys.exit(0)
    print(f"{'':26} OVRL   SIG   BAK  P808  SRMR   (n)")
    for arg in args:
        label, pat = arg.split("=", 1)
        files = sorted(glob.glob(pat))
        if not files:
            print(f"{label:26} нет файлов")
            continue
        r = score(files)
        m = r.mean(0)
        print(f"{label:26} {m[0]:4.2f}  {m[1]:4.2f}  {m[2]:4.2f}  {m[3]:4.2f}  {m[4]:5.2f}   ({len(files)}, мин OVRL {r[:, 0].min():.2f})")
