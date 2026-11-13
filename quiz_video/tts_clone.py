"""Синтез фраз клоном голоса: Qwen3-TTS Base (Apache-2.0), на процессоре ≈ 6 с счёта на 1 с речи.

Запускается СВОИМ Python (отдельное окружение с torch и пакетом qwen-tts), не тем, что остальная сборка:
  ОКРУЖЕНИЕ/bin/python tts_clone.py задание.json
Задание: {"model": папка модели, "out": папка, "ref_audio": образец.wav, "ref_text": его точный текст,
          "jobs": [[имя, текст, сид], …]}  →  out/имя_sСИД.wav (24 кГц). Готовые файлы пропускаются.
Образец — голос в уже одобренной обработке: клон повторяет его тембр. Текст — «как читать» (числа словами).
"""
import json
import sys
import time
from pathlib import Path

import numpy as np
import soundfile as sf
import torch
from qwen_tts import Qwen3TTSModel

cfg = json.load(open(sys.argv[1], encoding="utf-8"))
out = Path(cfg["out"])
out.mkdir(parents=True, exist_ok=True)
jobs = [(n, t, s) for n, t, s in cfg["jobs"] if not (out / f"{n}_s{s}.wav").exists()]
if not jobs:
    sys.exit(0)
torch.set_num_threads(cfg.get("threads", 4))
t0 = time.time()
model = Qwen3TTSModel.from_pretrained(cfg["model"], device_map="cpu", dtype=torch.float32)
prompt = model.create_voice_clone_prompt(ref_audio=cfg["ref_audio"], ref_text=cfg["ref_text"])
print(f"модель готова за {time.time() - t0:.0f} с, фраз к синтезу: {len(jobs)}", flush=True)
for name, text, seed in jobs:
    t = time.time()
    torch.manual_seed(seed)
    wavs, sr = model.generate_voice_clone(text=text, language="Russian", voice_clone_prompt=prompt)
    y = np.asarray(wavs[0], np.float32)
    sf.write(out / f"{name}_s{seed}.wav", y, sr)
    print(f"{name} s{seed}: {len(y) / sr:.1f} с звука за {time.time() - t:.0f} с", flush=True)
