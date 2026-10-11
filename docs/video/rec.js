const { chromium } = require('playwright');
(async () => {
  const b = await chromium.launch();
  const p = await b.newPage({ viewport: { width: 1080, height: 1920 } });
  await p.goto('file://' + process.cwd() + '/video.html'); await p.waitForTimeout(800);
  const fps = 30, total = await p.evaluate(() => window.TOTAL);
  const n = Math.round(total * fps);
  for (let i = 0; i < n; i++) {
    await p.evaluate(t => window.render(t), i / fps);
    await p.screenshot({ path: 'frames/f' + String(i).padStart(5, '0') + '.jpg', type: 'jpeg', quality: 90 });
    if (i % 300 === 0) console.log('frame', i, '/', n);
  }
  await b.close();
})();
