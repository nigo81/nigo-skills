async page => {
  const keys = [];
  await page.route(url => url.href.includes('athena/forward'), async route => {
    keys.push(route.request().headers()['athenaappkey'] || '');
    await route.continue();
  });
  await page.goto('https://www.gov.cn/zhengce/xxgk/gjgzk/index.htm', { waitUntil: 'networkidle' });
  await page.waitForTimeout(2000);
  await page.unroute(url => url.href.includes('athena/forward'));
  if (keys.length > 0 && keys[0]) {
    await page.evaluate(k => { document.title = 'GOVKEY:' + k; }, keys[0]);
  }
}
