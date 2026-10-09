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
| `ICPC_flight_fares_30Sep2026.xlsx` | Trip.com fares found via the Bright Data scraper, compared against the fares the plans used before. Rows 11–22 checked 30 Sep, rows 38–63 checked 1 Oct 2026, rows 65–73 Hohhot gateway + other flight sources. Sheets: "Connector notes", "Trains & hotels" (Ctrip trains and Trip.com hotels checked 1 Oct), "12306 & ChinaTicketOnline" (official fares, sale times and reseller mark-ups, checked 3 Oct). |
| `scripts/recalc_check.py` | Recalculates every workbook in Python, with no Excel or LibreOffice needed. Prints the key totals and exits 1 on any formula error. |
| `scripts/q12306.py` | Lists every 12306 train on a route/date with official fares, over plain HTTP (no browser, no login). Works from the cloud container. |
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

## Official train source: 12306 (built-in browser, no login)

Open `https://kyfw.12306.cn/otn/leftTicket/init?linktypeid=dc&fs=北京,BJP&ts=沈阳北,SBT&date=YYYY-MM-DD&flag=N,N,Y` once,
then run `fetch()` from inside that page (same origin, no login needed):
- Train list: ``/otn/${CLeftTicketUrl}?leftTicketDTO.train_date=D&leftTicketDTO.from_station=BJP&leftTicketDTO.to_station=SBT&purpose_codes=ADULT``
  (`CLeftTicketUrl` is a page global, currently `leftTicket/queryG`). Split each `data.result` row on `|`: [2] train_no,
  [3] code, [6]/[7] station codes (names in `data.map`), [8] dep, [9] arr, [16]/[17] from/to station no, [35] seat types.
- Fares: `/otn/leftTicket/queryTicketPrice?train_no=…&from_station_no=…&to_station_no=…&seat_types=…&train_date=YYYY-MM-DD`.
  Keys: A1 hard seat, A3 hard sleeper, A4 soft sleeper, A6 deluxe soft, O 2nd class, M 1st, A9 business, WZ standing.
  D-train sleepers come back as I (1st-class sleeper) / J (2nd-class sleeper) **in tenths of a yuan** (`4910` = ¥491).
- Station sale times: POST `https://www.12306.cn/index/otn/index12306/queryAllCacheSaleTime` (from a www.12306.cn page).
- Only the next 15 days are queryable, counting today (on 3 Oct: up to 17 Oct). So a train on date D goes on sale on
  **D − 14 days**. For later plan dates, query the same train on a date inside the window. Normal-train fares are
  fixed. Ctrip's 25 Jan connecting-train fares matched 12306's 15 Oct fares, so high-speed fares barely float.
- Station codes: Beijing BJP, Beijing South VNP, Beijing West BXP, Fengtai FTP, Chaoyang IFP, Shenyang North SBT,
  Shenyang SYT, Nanjing NJH, Hangzhou HZH, Hong Kong West Kowloon XJA.
- Sale times (Beijing time): Beijing 10:00, Beijing South 12:45, Beijing West / Fengtai 08:00, Chaoyang 10:00,
  Shenyang North 09:00, Shenyang 09:30, Nanjing / Nanjing South 08:15, Hangzhou / Hangzhou East 10:45, Hohhot East 14:00,
  HK West Kowloon 08:00.
- Plan sale dates: K53 (13 Nov) **30 Oct 10:00**, K54 (15 Nov) **1 Nov 09:00**, Z366 (12 Nov) **29 Oct 08:15**,
  D17 (25 Jan) **11 Jan 10:00**. The older "29 / 31 Oct" dates were one day off.

| Train | Route | Times | Official fare (CNY) | Use |
|---|---|---|---|---|
| K53 / K54 | Beijing ⇄ Shenyang North | 22:35–07:00 / 22:00–06:58 | hard 171, soft 263 | All Shenyang plans |
| K341 | Beijing → Shenyang | 23:00–08:47 | hard 194, soft 301 | Backup if K53 sells out |
| **Z366** | Nanjing → Shenyang North | 16:58–09:14 | hard 337, soft 530 | **Combined A (soft) / B (hard), 12 Nov**. Backups Z516 18:38, Z176 18:32 (same fares) |
| **D17** | Beijing → Hangzhou | 19:10–09:15 | 2nd-class sleeper 491, 1st 670 | **EC Final – Train, 25 Jan** |
| D11 | Beijing South → Hangzhou | 21:22–11:20 | 2nd-class sleeper 564, 1st 714 | Backup |
| Z281 | Beijing Fengtai → Hangzhou | 19:10–10:36 | hard 328, soft 515 | Cheapest backup |
| G-trains | Beijing South → Hangzhou East | 4.5–6 h | 2nd class 564–748 (afternoon 601–644) | Old plan |
| G381 | Beijing West → HK West Kowloon | 10:00–18:12 | 2nd class 1,248 | Not used (MIAT nonstop is cheaper) |

- Sleeper fares are the cheapest berth. Lower berths cost a bit more (ChinaTicketOnline quotes K53 hard at ¥182).
- 12306 needs an account with passport details to buy. Trip.com / Ctrip resell at about the official fare.

## ChinaTicketOnline (chinaticketonline.com)

- A pre-booking agent: it accepts orders any time and buys when 12306 sales open. It does not guarantee the ticket.
- API: from a chinaticketonline.com page, POST `/wp-admin/admin-ajax.php` with `action=getTrainListAjax&from=北京&to=沈阳&date=YYYY-MM-DD&trainCat=`
  (Chinese city names). It returns every train with per-seat `price` (USD), `serviceFee` (USD), `priceCNY` and
  `preSaleTime`. It works for any date (Jan 2027 too), so it is a handy timetable + sale-date lookup.
- **Prices are 30–70% above 12306** once the $8.50–20 per-ticket service fee is added (K53 hard $29 + $8.50 ≈ ¥252
  vs ¥171; D17 sleeper $110 + $12 ≈ ¥819 vs ¥491). Card transaction fee extra. Not used in any plan.
- Its Beijing ⇄ Ulaanbaatar / Moscow international trains page still says "Suspended" (stale COVID page). No Mongolia
  train booking there.

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

## Current results (after the 3 Oct 12306 update, incl. 10% contingency on the university part)

| Plan | Univ.-paid 3 ppl | Univ.-paid 4 ppl | Out-of-pocket pp (3 / 4) |
|---|---|---|---|
| Shenyang – Budget (K53/K54 hard sleepers, Bestay) | ₮5.05M | ₮6.75M | ₮401k / ₮317k |
| Shenyang – Comfort (RiCH family room, K54 soft sleeper back) – recommended | ₮6.34M | ₮8.26M | ₮401k / ₮317k |
| Hong Kong – Budget (recommended) | ₮9.11M | ₮12.17M | ₮409k / ₮323k |
| Hong Kong – Comfort | ₮9.53M | ₮13.10M | ₮409k / ₮323k |
| EC Final – Train (D17 sleeper out, Air China through ticket home 29 Jan, HanTing 3 nights) – recommended | ₮8.21M | ₮11.12M | ₮401k / ₮317k |
| EC Final – Fly (Aero Mongolia + Xiamen via Hohhot out, through ticket home, Mehood Theater) | ₮10.96M | ₮14.82M | ₮401k / ₮317k |

Combined Nanjing + Shenyang (3 ppl, total incl. fees and contingency): A Balanced ₮11.57M, B Budget ₮10.21M,
C Home between ₮15.61M (subtotal fits; ~₮0.6M over with contingency).

3 Oct changes:
- EC Train: D17 sleeper replaces the G-train and the 25 Jan hotel night (−₮440k / −₮641k). Also fixed the
  Hangzhou East taxi, which had been counted twice (the 29 Jan trip is the airport taxi row).
- Combined A / B: Z366 replaces China Eastern $121 and one Nanjing night (−₮647k / −₮905k). The workbook rows moved:
  16 Air China, 17 "Trains", 18 G-train, 19 Z366, 20 K54, 21 blank. Bag row removed.

## Cloud session notes (9 Oct 2026)

- **Trip.com is closed to automated access now.** The Bright Data scraper refuses trip.com ("not available for immediate
  residential (no KYC) access … robots.txt"), and headless Chromium gets a slider captcha. So no Air China fares from a cloud
  session. SlickTrip still prices MIAT, Spring, China Southern.
- 12306 answers plain HTTP from the container: `python scripts/q12306.py BJP SHH 2026-10-20 300`.
- The cloud container has LibreOffice (`soffice`) but not the `formulas` package.
- Booking.com does not resolve railway stations as a destination; use coordinates (Beijing Chaoyang station ≈ 39.9455, 116.5126).

## Shenyang without sleeper trains (9 Oct 2026, 3 ppl, incl. 10% contingency)

12306 fares (queried for 20 Oct; fixed fares): Beijing Chaoyang → Shenyang North G101 08:00–10:29 / G103 08:05–10:34 ¥355.
Shenyang North → Beijing Chaoyang evening: G122 18:19–20:49 / G126 19:00–21:39 ¥344, **G142 19:51–22:30 ¥268**,
G156 (Shenyang stn) 19:05–21:44 ¥268. Earliest morning train back is G3602 Shenyang 07:02 → Beijing Chaoyang 10:28 (¥369),
too tight for the 12:55 Air China flight. Beijing night: **Dequan Railway Station Hotel** (~1 km from Beijing Chaoyang stn,
8.7, 65 reviews) ¥304 for 3 ppl / 1 room, 13 and 15 Nov (Booking.com 9 Oct).

| Option | Univ.-paid 3 ppl | vs current |
|---|---|---|
| Comfort, but G142 back + Dequan night (rest day, RiCH 3 nights) | ₮6.53M | +₮188k vs ₮6.34M |
| Same with G126 19:00 (¥344) | ₮6.66M | +₮322k |
| Budget: MIAT 13 Nov 17:15, Dequan, G101, Bestay 14 Nov, G142, Dequan, Air China 16 Nov | ₮5.91M | +₮856k vs ₮5.05M |
| Same with RiCH (near NEU) instead of Bestay | ₮6.05M | |
| Comfort + 4th RiCH night + Air China SHE 09:00 → UB $286 | ₮7.43M | +₮1.09M |

Planes & trains: ~9 h 50 min for the no-sleeper plans (current Comfort 16 h 07 min with K54). Not put in the workbook yet.

## Shanghai regional (5–6 Dec 2026, provisional) – researched 9 Oct, not in the workbook yet

- Host Shanghai University; 2024 and 2025 editions at the **Baoshan campus gym** (≈31.3165, 121.3925; metro line 7 Shangda Rd).
  2025 invitation: fee **¥1,500/team**; day 1 registration 09:00–14:00, opening 14:00–15:00, warm-up 15:00–17:00; day 2
  contest 09:00–14:00, analysis 14:30–15:30, awards 15:30–17:00 (Weichang Building). Wildcard (外卡) requests by email to
  shu_icpc@163.com, first round closed 17 Oct in 2025 – a Mongolian team needs one.
- **Spring Airlines nonstop UB ⇄ Pudong**: 9C6520 UB 13:00 → PVG 17:00 ($164, Tue/Sat), 9C6519 PVG 08:00 → UB 12:00
  ($153, Tue 8 Dec) on SlickTrip. Hand baggage only. Sat 5 Dec lands after registration closes, so out Tue 1 Dec.
  MIAT OM265/266 UB ⇄ PVG is seasonal (last winter 17 Dec – 17 Jan, 20:40 → 00:10); none found for early Dec.
  UB → SHA via Beijing on China Southern: $480. MIAT UB → PEK 3 Dec 07:30 / 4 Dec 17:15 $154; PEK → UB 7 Dec 11:30 $197.
- 12306 Beijing ⇄ Shanghai: G-trains 4 h 18 – 4 h 54, 2nd class ¥598–672 (G17 Beijing South 13:00 → Shanghai 17:35 ¥667).
  Overnight 2nd-class sleepers: D7 Beijing 19:18 → Shanghai 07:25 ¥440, D5 21:21 → 09:27 ¥474; back D8 Shanghai 19:08 →
  Beijing 07:17 ¥440, D6 21:15 → 09:25 ¥474 (too late for MIAT 11:30). Z281/Z282 soft sleeper ¥476.5.
- Hotels (Booking.com 9 Oct, totals): Jenny's Apartment ~1 km, 10/10 (19 reviews): 3–6 Dec ¥1,375 (3 ppl) / ¥1,726 (4 ppl,
  2 rooms); 1–8 Dec ¥2,984 / ¥3,766. Atour Shangda Rd ~400 m, 9.3: ¥1,493 / ¥2,833 (3 nights). Mrs Li's Home ~1 km, 9.4:
  ¥1,288 / ¥2,444. Budget: Holiday Inn Express Gongkang ~4.7 km, 8.5: ¥919 / ¥1,838 (3 nights), ¥2,103 (7 nights, 3 ppl).
- Costs (3 ppl, univ.-paid incl. contingency): **A** MIAT 3 Dec + G17, Jenny's 3 nights, D8 sleeper + MIAT 7 Dec:
  ₮7.82M (5 days away). **B** Spring Tue 1 – Tue 8 Dec, Jenny's 7 nights, assumed ¥300 bag each way, ¥220 PVG taxis: ₮7.50M
  (8 days away). Personal ₮400.6k pp. A no-sleeper version of A still needs working out.

## Nanchang regional (19–20 Dec 2026, provisional) – started 9 Oct

- Host Jiangxi Normal University (on the 2026–27 host list). The 2019 Nanchang regional was at its **Yaohu campus**
  (紫阳大道 99 号) gym; metro line 1 to the campus. 2019 schedule: day 1 opening 14:30, warm-up 15:00–17:00; day 2 contest
  09:00–14:00, closing 15:00–16:30. Trains, flights and hotels not researched yet.

## Open items / next optimization ideas

- Home-between plan: Air China 08:10 on 10 Nov ($172 vs $219) saves ~₮507k for 3 people but needs a Beijing night on 9 Nov (~300 CNY). Not applied yet.
- Hotels re-checked on Trip.com 1 Oct (see "Trains & hotels" sheet). Bestay's low price is real (8.2, 1,062 reviews). Hong Kong unchanged.
- K53 on 13 Nov leaves ~1 h 20 min after MIAT immigration; if MIAT is late, use the one free change to a 14 Nov G-train (or K341 23:00). Buy sleepers when sales open: Z366 29 Oct 08:15, K53 30 Oct 10:00, K54 1 Nov 09:00.
- EC Fly outbound is two separate tickets via Hohhot: a late Aero Mongolia flight is not protected (fallback Air China 19:40 HET → HGH, or the ¥3,265 through ticket).
- Confirm the 2026 Shenyang fee and the HK and EC Final hosts and venues. All contest dates are provisional.
- EC Final: buy D17 on 12306 at 10:00 on 11 Jan (Spring Festival rush; use the 候补 waitlist if it sells out). The China timetable often changes in January, so re-check D17 in early Jan. It leaves ~8 h after MIAT lands, which gives plenty of buffer. Re-check separate MIAT + domestic tickets ~40 days ahead (mid-Dec) in case discount fares appear.
- PEK international ⇄ domestic transfer: bags may have to be collected for customs and re-dropped. Keep connections ≥ 2 h 30 min.
