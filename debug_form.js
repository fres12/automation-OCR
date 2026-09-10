const { chromium } = require('playwright');

(async () => {
  const browser = await chromium.launch({
    channel: 'chrome',
    headless: false,
    args: ['--no-first-run', '--no-default-browser-check'],
  });
  const page = await browser.newPage({ viewport: { width: 1440, height: 1600 } });

  await page.goto('https://forms.gle/9oteZBmCz87DBRrt9', { waitUntil: 'domcontentloaded', timeout: 30000 });
  await page.waitForTimeout(10000);

  console.log('URL:', page.url());
  console.log('TITLE:', await page.title());

  const bodyText = await page.locator('body').innerText();
  console.log('BODY SAMPLE:\n', bodyText.slice(0, 4000));

  const selectors = [
    'input',
    'textarea',
    'select',
    '[role="radio"]',
    '[role="checkbox"]',
    '[role="combobox"]',
    '[role="button"]',
    '[data-params]',
    '[jsname]'
  ];

  for (const selector of selectors) {
    const count = await page.locator(selector).count();
    console.log(`COUNT ${selector}:`, count);
    for (let i = 0; i < Math.min(count, 10); i++) {
      const el = page.locator(selector).nth(i);
      const tag = await el.evaluate((node) => node.tagName);
      const aria = await el.getAttribute('aria-label');
      const name = await el.getAttribute('name');
      const placeholder = await el.getAttribute('placeholder');
      const text = await el.innerText().catch(() => '');
      console.log({ i, selector, tag, aria, name, placeholder, text: String(text).slice(0, 200) });
    }
  }

  console.log('done');
  await page.screenshot({ path: 'debug-form.png', fullPage: true });
  console.log('saved debug-form.png');
  await browser.close();
})();
