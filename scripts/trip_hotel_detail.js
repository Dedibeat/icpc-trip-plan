// Trip.com hotel page: every room offer + "Room Details" popups with bed widths (e.g. "2 queen beds (1.8m wide)").
// Usage: NODE_PATH=$(npm root -g) node scripts/trip_hotel_detail.js HOTELID CHECKIN CHECKOUT ADULTS ROOMS OUT.txt SLUG
// SLUG must match the hotel (e.g. sunshine-select-homestay) or Trip.com shows a sign-in wall; NOPOP=1 skips the popups.
// Parse OUT.txt with scripts/parse_trip_detail.py.
const { chromium } = require('playwright');
(async () => {
  const [id, ci, co, ad, rm, out, slug] = process.argv.slice(2);
  const url = `https://www.trip.com/hotels/shenyang-hotel-detail-${id}/${slug || "x"}/?checkIn=${ci}&checkOut=${co}&adult=${ad}&crn=${rm}&children=0&curr=CNY&locale=en-XX`;
  const browser = await chromium.launch({ headless: true, proxy: { server: process.env.HTTPS_PROXY }, args: ['--disable-blink-features=AutomationControlled', '--disable-dev-shm-usage'] });
  const ctx = await browser.newContext({ locale: 'zh-CN', userAgent: 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0.0.0 Safari/537.36', viewport: { width: 1400, height: 1000 } });
  const page = await ctx.newPage();
  ctx.on('page', p => p.close().catch(() => {}));
  await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 60000 }).catch(e => console.log('goto', e.message));
  await page.waitForTimeout(14000);
  for (let i = 0; i < 4; i++) { try { await page.mouse.wheel(0, 1200); await page.waitForTimeout(1200); } catch (e) { break; } }
  let res = '';
  try {
    const head = await page.evaluate(() => { const t = document.body.innerText; return t.slice(0, 1500); });
    res += 'HEAD: ' + head.replace(/\n+/g, ' / ').slice(0, 900) + '\n';
    const body = await page.evaluate(() => document.body.innerText);
    res += 'BODY: ' + body.replace(/\n+/g, ' / ') + '\n';
    const links = page.getByText('Room Details', { exact: true });
    const n = process.env.NOPOP ? 0 : Math.min(await links.count(), 6);
    for (let i = 0; i < n; i++) {
      try {
        await links.nth(i).click({ timeout: 4000 }); await page.waitForTimeout(2500);
        const pop = await page.evaluate(() => { const els = [...document.querySelectorAll('[class*=modal], [class*=Modal], [class*=popup], [class*=dialog], [role=dialog]')].filter(e => e.offsetParent !== null || getComputedStyle(e).position === 'fixed'); return els.map(e => e.innerText).sort((a, b) => b.length - a.length)[0] || ''; });
        res += `POPUP${i}: ` + pop.replace(/\n+/g, ' / ').slice(0, 1500) + '\n';
        await page.evaluate(() => { const els = [...document.querySelectorAll('[class*=modal], [class*=Modal], [class*=popup], [class*=dialog], [role=dialog]')].filter(e => e.offsetParent !== null || getComputedStyle(e).position === 'fixed');
          for (const m of els) { const c = m.querySelector('[class*=close], [class*=Close], [aria-label*=lose]'); if (c) { c.click(); } } });
        await page.keyboard.press('Escape');
        await page.waitForTimeout(600);
      } catch (e) { res += `POPUP${i} ERR ${e.message.split('\n')[0]}\n`; }
    }
  } catch (e) { res += 'ERR ' + e.message; }
  require('fs').writeFileSync(out, res);
  console.log(id, 'done', res.length);
  await browser.close();
})();
