const { chromium } = require('playwright');
const path = require('path');

(async () => {
  const browser = await chromium.launch({
    executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome',
    args: ['--no-sandbox']
  });
  const page = await browser.newPage({ viewport: { width: 1280, height: 900 } });

  const filePath = path.resolve(__dirname, 'cv.html');
  await page.goto('file://' + filePath, { waitUntil: 'networkidle', timeout: 30000 });

  await page.waitForTimeout(3000);

  // Force all reveal animations to show
  await page.evaluate(() => {
    document.querySelectorAll('.reveal').forEach(el => {
      el.classList.add('visible');
    });
  });

  await page.waitForTimeout(1000);

  // Scroll through the entire page to ensure everything renders
  await page.evaluate(async () => {
    const delay = ms => new Promise(r => setTimeout(r, ms));
    const h = document.body.scrollHeight;
    for (let i = 0; i < h; i += 400) {
      window.scrollTo(0, i);
      await delay(100);
    }
    window.scrollTo(0, 0);
  });

  await page.waitForTimeout(500);

  await page.screenshot({
    path: path.resolve(__dirname, 'cv-preview.png'),
    fullPage: true
  });

  await page.pdf({
    path: path.resolve(__dirname, 'Reman_Badawood_CV.pdf'),
    format: 'A4',
    printBackground: true,
    margin: { top: '0', right: '0', bottom: '0', left: '0' }
  });

  await browser.close();
  console.log('Done');
})();
