// Рендер кадров шаблона в видео: node render.js timeline.json out.mp4 [--preview t1,t2,...]
// Каждый кадр — снимок страницы после window.render(t); кадры идут в ffmpeg через pipe.
const fs = require('fs');
const path = require('path');
const { once } = require('events');
const { spawn } = require('child_process');
const { chromium } = require('playwright');

(async () => {
  const [tlPath, out, flag, list] = process.argv.slice(2);
  const tl = JSON.parse(fs.readFileSync(tlPath, 'utf8'));
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1080, height: 1920 }, deviceScaleFactor: 1 });
  await page.goto('file://' + path.join(__dirname, 'template.html'));
  await page.evaluate(t => window.setup(t), tl);

  if (flag === '--preview') {  // отдельные кадры в PNG для проверки
    for (const t of list.split(',').map(Number)) {
      await page.evaluate(x => window.render(x), t);
      await page.screenshot({ path: `${out}_${t.toFixed(2)}.png` });
    }
    await browser.close();
    return;
  }

  const fps = tl.fps || 30, frames = Math.ceil(tl.duration * fps);
  const ff = spawn('ffmpeg', ['-y', '-loglevel', 'error', '-f', 'image2pipe', '-framerate', String(fps),
    '-c:v', 'mjpeg', '-i', '-', '-c:v', 'libx264', '-preset', 'medium', '-crf', '19',
    '-pix_fmt', 'yuv420p', out], { stdio: ['pipe', 'inherit', 'inherit'] });
  const t0 = Date.now();
  for (let i = 0; i < frames; i++) {
    await page.evaluate(x => window.render(x), i / fps);
    const buf = await page.screenshot({ type: 'jpeg', quality: 93 });
    if (!ff.stdin.write(buf)) await once(ff.stdin, 'drain');
    if (i % (fps * 10) === 0) console.log(`кадр ${i}/${frames}, ${((Date.now() - t0) / 1000).toFixed(0)} с`);
  }
  ff.stdin.end();
  const [code] = await once(ff, 'close');
  await browser.close();
  if (code !== 0) { console.error('ffmpeg завершился с ошибкой', code); process.exit(1); }
  console.log(`видео готово: ${out}, кадров ${frames}, ${((Date.now() - t0) / 1000).toFixed(0)} с`);
})().catch(e => { console.error(e); process.exit(1); });
