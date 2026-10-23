// Trip.com hotel list near a landmark, cheapest first (headless Chromium; works from the cloud, ~1 min per run).
// Usage: NODE_PATH=$(npm root -g) node scripts/trip_hotel_list.js CHECKIN CHECKOUT ADULTS ROOMS OUT.json [scrolls] [extraFilters]
// Only the 10 cheapest cards load (no paging without login). Landmark = NEU Nanhu (Trip.com id 1803379).
// Filter codes: 1 double bed 4~1*4*1, 2 beds 4~2*4*2, family room 81~1188*81*1188, price band 15~Range*15*150~250 (the list may ignore them).
const { chromium } = require('playwright');
(async () => {
  const [ci, co, ad, rm, out, scrolls, extra] = process.argv.slice(2);
  const lf = `17~3*17*3*2,29~1*29*1~${ad}*2` + (extra ? ',' + extra : '');
  const url = `https://www.trip.com/hotels/list?cityId=451&searchType=LM&optionId=1803379&searchValue=13%7C1803379*13*41.7655311%7C123.4181971%7CNortheastern%20University%7C1803379%7C2&checkin=${ci}&checkout=${co}&crn=${rm}&adult=${ad}&listFilters=${encodeURIComponent(lf)}&curr=CNY&locale=en-XX`;
  const browser = await chromium.launch({ headless: true, proxy: { server: process.env.HTTPS_PROXY }, args: ['--disable-blink-features=AutomationControlled', '--disable-dev-shm-usage'] });
  const ctx = await browser.newContext({ locale: 'zh-CN', userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36', viewport: { width: 1400, height: 1000 } });
  const page = await ctx.newPage();
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(e => console.log('goto', e.message));
  await page.waitForTimeout(12000);
  for (let i = 0; i < parseInt(scrolls || '10'); i++) {
    try {
      await page.evaluate(() => window.scrollTo(0, document.body.scrollHeight - 2500));
      await page.waitForTimeout(1200);
      await page.mouse.wheel(0, 1500);
      await page.waitForTimeout(3500);
    } catch (e) { console.log('scroll', e.message.split('\n')[0]); break; }
  }
  const cards = await page.evaluate(() => [...document.querySelectorAll('[data-offline-hotelid]')].map(e => ({ id: e.getAttribute('data-offline-hotelid'), t: e.innerText.replace(/\n+/g, ' | ') })));
  const seen = new Set(), rows = [];
  for (const c of cards) {
    if (seen.has(c.id)) continue; seen.add(c.id);
    const t = c.t, parts = t.split(' | ');
    const g = re => { const m = t.match(re); return m ? m[1] : null; };
    rows.push({ id: c.id, name: parts[0] === 'Hi China!' ? parts[1] : parts[0], score: g(/([\d.]+)\/10/), reviews: g(/(\d[\d,]*) reviews/),
      dist: g(/([\d.]+ k?m (?:walk|drive)) from/), total: g(/Total price: CNY ([\d,]+)/), units: g(/(\d+ rooms? × \d+ nights?)/), text: t });
  }
  require('fs').writeFileSync(out, JSON.stringify({ url: page.url(), rows }, null, 1));
  console.log(`${ci}->${co} ${ad}a/${rm}r: ${rows.length} hotels`);
  for (const r of rows) console.log(`${r.total ? 'CNY ' + r.total : '?'} | ${r.units} | ${r.name} | ${r.score} (${r.reviews}) | ${r.dist} | ${(r.text.match(/Show on Map \| (.*?) \|/) || [, ''])[1]}`);
  await browser.close();
})();
