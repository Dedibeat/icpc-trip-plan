# ICPC 2026–27 trip plan – project details

Budget workbooks for a Mongolian university team (3 contestants, sometimes + 1 coach) travelling from
Ulaanbaatar to ICPC Asia EC contests. The university funds **₮15,000,000 per regional**. It pays for flights,
trains, hotels, meals and local transport. The team pays the contest fee and "Other" costs (UB airport taxi,
insurance, eSIM, visa).

## Files

| File | What it is |
|---|---|
| `ICPC_Single_Contest_Plans.xlsx` | **Main plan.** Compare sheet + 2 variants each for Shenyang, Hong Kong and EC Final Hangzhou, costed for 3 and 4 people. Separate trips with home in between. |
| `ICPC_2026_Nanjing_Shenyang_Budget (2).xlsx` | Older combined Nanjing + Shenyang plan for 3 people: A Balanced / B Budget / C Home between, plus an Itineraries sheet. |
| `ICPC_flight_fares_30Sep2026.xlsx` | Trip.com fares found via the Bright Data scraper, compared against the fares the plans used before. Rows 11–22 checked 30 Sep, rows 38–63 checked 1 Oct 2026, rows 65–73 Hohhot gateway + other flight sources. Sheets: "Connector notes", "Trains & hotels" (Ctrip trains and Trip.com hotels checked 1 Oct). |
| `scripts/recalc_check.py` | Recalculates every workbook in Python, with no Excel or LibreOffice needed. Prints the key totals and exits 1 on any formula error. |
| `requirements.txt` | Python deps: `openpyxl` (edit workbooks), `formulas` (recalculate). |

## Setup

```bash
python -m venv .venv
.venv/Scripts/python -m pip install -r requirements.txt
.venv/Scripts/python scripts/recalc_check.py
```

- Installing `formulas` into the global Python 3.11 failed with a WinError 2 file lock on `schedula.exe`. Use the project `.venv`.
- LibreOffice is not installed on this machine. `scripts/recalc_check.py` replaces the skill's `recalc.py`. Workbooks have `fullCalcOnLoad=True`, so Excel recalculates them on open.
- Set `PYTHONIOENCODING=utf-8` when printing cell text. The console is cp1252 and crashes on `→`, `₮` and `–`.

## Workbook conventions (keep them when editing)

- Blue font = inputs, black = formulas, green = cross-sheet links. Yellow fill = assumption cells, used only when the source note says "ASSUMPTION".
- Every hard-coded price has a source note saying which tool, the date checked, and the total seen.
- Single-contest sheets: row 10 = outbound flight, row 11 = return flight (EC Fly: row 10 Aero Mongolia UB → Hohhot, row 11 Xiamen Hohhot → Hangzhou, row 12 return through ticket). Travel times are stored as `timedelta` with format `[h]" h "mm" min"`.
- The Compare sheet's "Recommended picks" text (B23–B25) contains **hard-coded ₮ figures**. Update it whenever totals change.
- Sheet names contain en dashes (`Shenyang – Budget`). Quote them in references.

## Fare knowledge (Trip.com, 30 Sep – 1 Oct 2026, per person, one way)

Trip.com search URL that works through the Bright Data `scrape_as_markdown` / `scrape_batch` tools:
`https://www.trip.com/flights/showfarefirst?dcity=ubn&acity=bjs&ddate=YYYY-MM-DD&triptype=ow&class=y&quantity=1&curr=USD`
(use `triptype=rt&rdate=…` for a round trip). Scrape at most 2 pages per batch; the pages are big.

| Leg | Date | Best fare | Notes |
|---|---|---|---|
| UB → Beijing PEK | 6, 13 Nov | MIAT 17:15–19:30 **$150** | SlickTrip had $154 |
| UB → Beijing PEK | 12 Nov | MIAT 07:30–09:45 **$151** | |
| UB → Beijing PEK | 25 Jan | MIAT 08:15–10:30 **$150** | Air China $363 that day |
| Beijing PEK → UB | 16 Nov | **Air China 12:55–15:05 $171** | MIAT 11:30 is $194–196. Now used in all 16 Nov returns |
| Beijing PEK → UB | 10 Nov | MIAT 20:45 $219; Air China 08:10 $172 | Air China needs an extra Beijing night |
| Beijing PEK → UB | 30 Jan | MIAT 11:30–13:50 **$219** | Air China 08:10 is $416 |
| UB ↔ Hong Kong | 8–11 Jan | MIAT nonstop **$583 return** | SlickTrip had $590.50 |
| Nanjing → Shenyang | 13 Nov | **China Eastern 07:40–09:50 $121, bag incl.** | Replaced Spring 9C7758 19:55–23:50 ($110.67 + ~100 CNY bag). Shenzhen Air 19:25 $129 is the backup. Cheapest days are 7/10/12/14 Nov at $88 |
| Beijing → Hangzhou | 25 Jan | ¥2,265 (CA 21:30, lands 00:05); ¥2,265–2,540 otherwise | Spring Festival rush, southbound. **The old ¥700 assumption is not on sale.** 26 Jan ¥1,286 is Shandong 18:35 (too late for day 1) |
| Hangzhou → Beijing | 29 Jan | ¥1,809 (Air China, 07:00–10:00 hourly) | |
| **UB → Hangzhou through** | 25 Jan | **Air China ¥3,265**, UB 11:50 → PEK (2 h 35 min) → HGH 18:45 | Same price as separate MIAT + domestic. One ticket, bags through. Now the backup to the Hohhot route (¥2,752). Also MIAT + Hong Kong Airlines via HKG ¥2,721, lands 23:40 |
| **Hangzhou → UB through** | 29 Jan | **Air China ¥2,250**, HGH 08:00 → PEK (2 h 35 min) → UB 15:05 | Cheaper than domestic + Beijing night + MIAT 30 Jan (~¥3,280 + hotel). Home a day earlier |
| Shenyang → UB through | 16 Nov | Air China $286, SHE 09:00 → UB 15:05 | ~$50 pp dearer than train + Beijing night + Air China. Use only if the closing runs late |
| UB → Nanjing / Nanjing → UB through | 6 / 10 Nov | $353 / $360–394 via Beijing | Dearer than MIAT + fast train. Not used |

| **UB → Hohhot (HET)** | 25 Jan | **Aero Mongolia 06:30–08:00 ¥892** | Also Air China 13:00 ¥1,932. Not on 6 Nov (MIAT 22:25 ¥962 only) |
| **Hohhot → Hangzhou** | 25 Jan | **Xiamen Airlines 11:15–13:40 ¥1,860** | Air China 19:40–22:20 ¥1,710. With Aero Mongolia: ¥2,752 vs ¥3,265 through ticket – **now the EC Fly outbound** (self-transfer at HET, 3 h 15 min) |
| Hangzhou → Hohhot → UB | 29 Jan | ¥1,710 + Aero Mongolia 23:30 ¥998 = ¥2,708 | Dearer than the ¥2,250 through ticket. Not used |
| UB ⇄ Hohhot | 13 / 16 Nov | Air China $157 out; Aero Mongolia 16:55 $150 back | Useless for Shenyang/Nanjing (no fast rail from Hohhot) |

- Air China UB↔Beijing is often cheaper than MIAT on the return leg. SlickTrip shows Air China as $0 (unpriced), so check it on Trip.com.
- Trip.com offers "Student tickets" on some routes. Check them at booking.
- Domestic Chinese fares 4 months out show near-full fares on Trip.com (CNY). Use `curr=CNY`. Domestic routes redirect to `/chinaflights/ShowFareFirst`, and the date strip shows the cheapest fare for each day.
- **Through tickets** (one Air China ticket UB ⇄ Chinese city via Beijing Capital) are worth checking whenever a plan has a Beijing night + domestic leg. They won for Hangzhou (EC Final), but not for Nanjing or Shenyang, where high-speed rail is cheap.
- Skyscanner China route pages have no data this far out (Jan showed "查找价格"). Ctrip's own API (`flights.ctrip.com/itinerary/api/12808/lowestPrice`) is blocked by Bright Data without KYC.
- Other flight sources (1 Oct): Google Flights (scrape `google.com/travel/flights?q=Flights from UBN to PEK on YYYY-MM-DD one way&curr=USD`) shows MIAT only ($154/$196), no Air China prices. SlickTrip calendar matches. Expedia's MCP tool is geo-blocked for Mongolia. Trip.com stays the best flight source.

## Train source: Ctrip timetables (built-in browser, no login)

`https://trains.ctrip.com/webapp/train/list?ticketType=0&dStation=沈阳&aStation=北京&dDate=YYYY-MM-DD` (URL-encode the
Chinese city names) lists every train with fares for any date, even before sales open (15 days ahead). Read it with the
built-in browser's page text; Bright Data/Trip.com route pages only give "from $X".

- **Overnight sleepers replace a G-train + Beijing hotel night** for Shenyang: K53 Beijing 22:35 → Shenyang North 07:00 (13 Nov) and K54 Shenyang North 22:00 → Beijing 06:58 (15 Nov), hard sleeper ¥171 / soft ¥263. Reach Beijing station from PEK by Airport Express + metro line 2 (~1 h). Used in all Shenyang plans.
- Beijing → Nanjing sleepers (D5 21:21, D11 21:22, ¥377–401) leave too soon after MIAT lands at 19:30, so the 6 Nov Beijing night stays.
- Shenyang → Beijing G-train fares vary ¥268–383 by train; the plans keep ¥339 where used.

## Hotel source: Trip.com list pages (built-in browser)

- Search the venue in the Trip.com hotel search box once to get its landmark id, then reuse the list URL:
  `https://www.trip.com/hotels/list?cityId=<city>&searchType=LM&optionId=<id>&searchValue=13|<id>*13*<lat>|<lon>|<name>|<id>|2&checkin=YYYY-MM-DD&checkout=YYYY-MM-DD&crn=1&adult=<n>&listFilters=29~1*29*1~<n>*2&curr=CNY&locale=en-XX`.
  Add `17~3*17*3*2,` to the front of `listFilters` for lowest price, `17~5*17*5*2,` for distance. A single hotel: `searchType=H&optionId=<hotelId>&searchValue=31|<hotelId>*31*<hotelId>`.
- Landmark ids: NEU Shenyang `451 / 1803379 / 41.7655311,123.4181971`; NUAA Jiangjun Rd `12 / 49661385 / 31.9356703,118.7899533`; HKU `58 / 4193832 / 22.2804093,114.1416149`; HZNU Cangqian `17 / 15665541 / 30.2887424,120.0068028`.
- Fast extraction: `fetch()` the list URL from inside a trip.com page and parse the server-rendered HTML (`[data-offline-hotelId]` cards: `.hotelName`, `.score`, `.comment-num`, `.position-desc`, `.room-name`; the last `CNY n` in the card is the total incl. taxes). Only the first 5 cards are server-rendered.
- Searching with 3 or 4 adults in **1 room** finds family rooms that fit the whole team (e.g. RiCH Family Room ¥976/3 nights for 3 or 4). These beat Booking.com's 2-room prices for 4 people.
- Hong Kong: Trip.com is not cheaper than Booking.com near HKU; cheap Tsim Sha Tsui guesthouses there are rated < 7.
- Tripadvisor's MCP tool timed out; SlickTrip hotel search is Google Hotels data and misses most Chinese budget hotels.

- If a workbook is open in Excel, saving fails with PermissionError (look for `~$*.xlsx` lock files). Ask the user to close it.

## Current results (after 1 Oct train/hotel/Hohhot update, incl. 10% contingency on the university part)

| Plan | Univ.-paid 3 ppl | Univ.-paid 4 ppl | Out-of-pocket pp (3 / 4) |
|---|---|---|---|
| Shenyang – Budget (K53/K54 hard sleepers, Bestay) | ₮5.05M | ₮6.75M | ₮401k / ₮317k |
| Shenyang – Comfort (RiCH family room, K54 soft sleeper back) – recommended | ₮6.34M | ₮8.26M | ₮401k / ₮317k |
| Hong Kong – Budget (recommended) | ₮9.11M | ₮12.17M | ₮409k / ₮323k |
| Hong Kong – Comfort | ₮9.53M | ₮13.10M | ₮409k / ₮323k |
| EC Final – Train (rail out, Air China through ticket home 29 Jan, HanTing) – recommended | ₮8.65M | ₮11.76M | ₮401k / ₮317k |
| EC Final – Fly (Aero Mongolia + Xiamen via Hohhot out, through ticket home, Mehood Theater) | ₮10.96M | ₮14.82M | ₮401k / ₮317k |

Combined Nanjing + Shenyang (3 ppl, total incl. fees and contingency): A Balanced ₮12.22M, B Budget ₮11.11M,
C Home between ₮15.61M (subtotal fits; ~₮0.6M over with contingency).

## Open items / next optimization ideas

- Home-between plan: Air China 08:10 on 10 Nov ($172 vs $219) saves ~₮507k for 3 people but needs a Beijing night on 9 Nov (~300 CNY). Not applied yet.
- Hotels re-checked on Trip.com 1 Oct (see "Trains & hotels" sheet). Bestay's low price is real (8.2, 1,062 reviews). Hong Kong unchanged.
- K53 on 13 Nov leaves ~1 h 20 min after MIAT immigration; if MIAT is late, use the one free change to a 14 Nov G-train. Buy sleepers when sales open (29 / 31 Oct).
- EC Fly outbound is two separate tickets via Hohhot: a late Aero Mongolia flight is not protected (fallback Air China 19:40 HET → HGH, or the ¥3,265 through ticket).
- Confirm the 2026 Shenyang fee and the HK and EC Final hosts and venues. All contest dates are provisional.
- EC Final: buy the 25 Jan Beijing → Hangzhou train the day sales open (~10 Jan). Re-check separate MIAT + domestic tickets ~40 days ahead (mid-Dec) in case discount fares appear.
- PEK international ⇄ domestic transfer: bags may have to be collected for customs and re-dropped. Keep connections ≥ 2 h 30 min.
