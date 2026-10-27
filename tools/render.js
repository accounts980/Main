// Render helper (Playwright / Chromium).
//   node render.js shots <dir-with-html> <out-dir>   -> PNG of #shot for every *.html
//   node render.js pdf <input.html> <output.pdf>
//   node render.js pages <input.pdf-dir-unused>        (reserved)
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

(async () => {
  const [mode, a, b] = process.argv.slice(2);
  const browser = await chromium.launch();
  if (mode === 'shots') {
    const page = await browser.newPage({ viewport: { width: 2200, height: 1200 }, deviceScaleFactor: 2 });
    for (const f of fs.readdirSync(a).filter(x => x.endsWith('.html'))) {
      await page.goto('file://' + path.resolve(a, f));
      await page.evaluate(() => document.fonts.ready);
      const el = await page.$('#shot');
      const out = path.join(b, f.replace('.html', '.png'));
      await el.screenshot({ path: out });
      console.log('shot', out);
    }
  } else if (mode === 'svg') {
    // rasterise an SVG for visual checking
    const page = await browser.newPage({ deviceScaleFactor: 1.5 });
    await page.goto('file://' + path.resolve(a));
    const el = await page.$('svg');
    await el.screenshot({ path: b });
  } else if (mode === 'pdf') {
    const page = await browser.newPage();
    await page.goto('file://' + path.resolve(a), { waitUntil: 'load' });
    await page.evaluate(() => document.fonts.ready);
    const foot = `<div style="font-family:Inter,'DejaVu Sans',Arial,sans-serif;font-size:7.5px;color:#5A6B82;width:100%;padding:0 14mm;display:flex;justify-content:space-between;">
      <span>ISKO Excel ERP Rebuild Plan - 100,000-Row Resilience Blueprint - v1.0 - 27 Oct 2026 - planning document, not a live ERP</span>
      <span>Page <span class="pageNumber"></span> of <span class="totalPages"></span></span></div>`;
    await page.pdf({ path: b, preferCSSPageSize: true, printBackground: true, displayHeaderFooter: true,
                     headerTemplate: '<div></div>', footerTemplate: foot });
    console.log('pdf', b);
  }
  await browser.close();
})();
