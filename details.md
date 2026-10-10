# ICPC 2026–27 trip plan – project details

Budget workbooks for a Mongolian university team (3 contestants, sometimes + 1 coach) travelling from
Ulaanbaatar to ICPC Asia EC contests.

**Funding model (user, 10 Oct 2026):** one regional is **funded with ₮12,000,000 and the coach travels** (4 people);
the other regional is **unfunded: the 3 contestants pay everything** themselves. Which regional is which is not decided yet.
In `ICPC_Single_Contest_Plans.xlsx` every "3 ppl" column / H column = the UNFUNDED regional (`Compare!B12` = ₮0, so the
"out-of-pocket per person" is the whole cost per contestant) and every "4 ppl" column / K column = the FUNDED regional
(`Compare!B5` = ₮12M). Funding covers flights, trains, hotels, meals and local transport; the contest fee and "Other" costs
(UB airport taxi, insurance, eSIM, visa) are paid by the team. The older combined workbook still assumes ₮15M for 3 people.

## Files

| File | What it is |
|---|---|
| `ICPC_Single_Contest_Plans.xlsx` | **Main plan.** Compare sheet + 2–3 variants each for Shenyang (4), Shanghai (3), Nanchang (3), Hong Kong (3) and EC Final Hangzhou (2), costed for 3 and 4 people. Separate trips with home in between. |
| `ICPC_2026_Nanjing_Shenyang_Budget (2).xlsx` | Older combined Nanjing + Shenyang plan for 3 people: A Balanced / B Budget / C Home between, plus an Itineraries sheet. |
| `ICPC_flight_fares_30Sep2026.xlsx` | Trip.com fares found via the Bright Data scraper, compared against the fares the plans used before. Rows 11–22 checked 30 Sep, rows 38–63 checked 1 Oct 2026, rows 65–73 Hohhot gateway + other flight sources. Sheets: "Connector notes", "Trains & hotels" (Ctrip trains and Trip.com hotels checked 1 Oct), "12306 & ChinaTicketOnline" (official fares, sale times and reseller mark-ups, checked 3 Oct), "10 Oct – Dec & Jan checks" (Trip.com / SlickTrip / 12306 / CTO figures behind the Shanghai, Nanchang, Shenyang G-train and Hong Kong sheets). |
| `scripts/recalc_check.py` | Recalculates every workbook in Python, with no Excel or LibreOffice needed. Prints the key totals and exits 1 on any formula error. |
| `scripts/q12306.py` | Lists every 12306 train on a route/date with official fares, over plain HTTP (no browser, no login). Works from the cloud container. On Windows set `Q12306_INSECURE=1` (12306's CA is not in the Windows store); it is slow (~1 s per train), so the built-in browser snippet below is faster. |
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

| Plan | Travel + stay 3 ppl (unfunded) | Travel + stay 4 ppl (₮12M funded) | Fee + other pp (3 / 4) |
|---|---|---|---|
| Shenyang – Budget (K53 soft sleeper before day 1, K54 hard after, Bestay) | ₮5.21M | ₮6.96M | ₮401k / ₮317k |
| Shenyang – Comfort (RiCH family room, K54 hard sleeper back) – recommended | ₮6.18M | ₮8.04M | ₮401k / ₮317k |
| Hong Kong – Budget (recommended) | ₮9.11M | ₮12.17M | ₮409k / ₮323k |
| Hong Kong – Comfort | ₮9.53M | ₮13.10M | ₮409k / ₮323k |
| EC Final – Train (D17 sleeper out, Air China through ticket home 29 Jan, HanTing 3 nights) – recommended | ₮8.21M | ₮11.12M | ₮401k / ₮317k |
| EC Final – Fly (Aero Mongolia + Xiamen via Hohhot out, through ticket home, Mehood Theater) | ₮10.96M | ₮14.82M | ₮401k / ₮317k |
| Shenyang – Comfort (G-train) / Budget (G-train) | ₮6.72M / ₮6.05M | ₮8.84M / ₮8.27M | ₮401k / ₮317k |
| Shanghai – Sleeper (D8 seat back; recommended) / Spring nonstop / Train (no sleeper) | ₮7.48M / ₮8.21M / ₮8.11M | ₮9.83M / ₮10.73M / ₮10.78M | ₮401k / ₮317k |
| Nanchang – Sleeper (D136 seat back; recommended) / Comfort / Fly in | ₮6.31M / ₮7.96M / ₮8.41M | ₮8.19M / ₮10.32M / ₮10.95M | ₮401k / ₮317k |
| Hong Kong – Budget (₮9.08M / ₮12.12M), Value (₮9.08M / ₮12.38M), Comfort (₮9.56M / ₮13.15M), **Sleep first** (2 rooms in Sheung Wan, ₮11.63M / ₮14.45M) – updated 10 Oct | | | ₮409k / ₮323k |

Combined Nanjing + Shenyang (3 ppl, total incl. fees and contingency): A Balanced ₮11.57M, B Budget ₮10.21M,
C Home between ₮15.61M (subtotal fits; ~₮0.6M over with contingency).

3 Oct changes:
- EC Train: D17 sleeper replaces the G-train and the 25 Jan hotel night (−₮440k / −₮641k). Also fixed the
  Hangzhou East taxi, which had been counted twice (the 29 Jan trip is the airport taxi row).
- Combined A / B: Z366 replaces China Eastern $121 and one Nanjing night (−₮647k / −₮905k). The workbook rows moved:
  16 Air China, 17 "Trains", 18 G-train, 19 Z366, 20 K54, 21 blank. Bag row removed.

## Cloud session notes (9 Oct 2026)

- **Trip.com is closed to automated access from the cloud** (on Windows the built-in browser still reads it, see the 10 Oct section). The Bright Data scraper refuses trip.com ("not available for immediate
  residential (no KYC) access … robots.txt"), and headless Chromium gets a slider captcha. So no Air China fares from a cloud
  session. SlickTrip still prices MIAT, Spring, China Southern.
- 12306 answers plain HTTP from the container: `python scripts/q12306.py BJP SHH 2026-10-20 300`.
- The cloud container has LibreOffice (`soffice`) but not the `formulas` package.
- Booking.com does not resolve railway stations as a destination; use coordinates (Beijing Chaoyang station ≈ 39.9455, 116.5126).

## Sleep principle (user rule, 10 Oct 2026)

Losing sleep **before** a contest hurts performance; **after** the contest anything goes. Rules used in the plans and the `Sleep check` sheet:
the two nights before the contest day are in a quiet bed near the venue (≥7.5 h in bed, wake ≥06:00 on contest morning); no overnight train or
flight on the contest eve; a sleeper before day 1 only as a soft sleeper that arrives by midday with a bed night in between; flights before 08:00
only ≥2 nights before; after the contest use the cheapest option (hard sleeper, seat, red-eye, early flight). Applied: Shenyang Budget K53 hard →
soft (pre), Shenyang K54 soft → hard, Shanghai D8 sleeper → seat, Nanchang D136 sleeper → seat, older workbook K54 soft → hard (post),
new `Hong Kong – Sleep first` (2 rooms, Sheung Wan). The funding is ₮15M per regional and plans use ₮6–9M, so spend the headroom on sleep.
Audit table: `Sleep check` sheet of `ICPC_Single_Contest_Plans.xlsx`. Weak spots still flagged: Hong Kong – Budget (noisy shared room, 35 min
commute), EC Final – Train (D17 sleeper is the night before the check-in day), Shenyang – Budget (K53 before day 1).

## 10 Oct 2026 session: Shenyang G-train, Shanghai, Nanchang, Hong Kong (now all in the workbook)

Sources used (as the user asked): the built-in browser with **Trip.com** (flights, hotels), **12306** (fares) and **ChinaTicketOnline** (timetables / sale dates), plus SlickTrip.

### Method notes

- **Trip.com flights work in the built-in browser** (Bright Data is still refused). Use the mobile URL
  `https://www.trip.com/m/flights/ubn-to-bjs/tickets-ubn-bjs/?dcitycode=ubn&acitycode=bjs&ddate=YYYY-MM-DD&triptype=0&classtype=0&adult=1&curr=USD&locale=en-XX`
  (round trip: `&rdate=…&triptype=1`; Shanghai city code `sha` covers PVG + SHA; Beijing `bjs`). The page needs **~40 s and must be the fronted
  tab** (`tabs_select`) to finish loading; background tabs stay on "Finalizing search results". Read it with
  `document.body.innerText` after `Average one-way`. It lists Air China fares that SlickTrip leaves unpriced.
- Trip.com hotel landmark ids (use the list-page recipe above): **JXNU Yaohu campus** `cityId=21959 optionId=6687249 28.6789462|116.0316786`
  (listed as "Nanchang County"); **Shanghai University Baoshan** `cityId=2 optionId=9534399 31.3159716|121.3935523`.
  Booking.com does not geocode these campuses; give it `latitude=…&longitude=…` instead (Nanchang gym 28.6829, 116.0323; HKU 22.2828, 114.1371).
- Nominatim (`nominatim.openstreetmap.org/search?q=紫阳大道99号+南昌&format=json`) finds Chinese addresses when English names fail.
- **12306 from the browser** (fast): on a kyfw.12306.cn page define `q(fromCode,toCode,date)` = fetch `/otn/leftTicket/queryG?...` and
  `p(train)` = fetch `/otn/leftTicket/queryTicketPrice?...` (see "Official train source"). Calling many price requests in one script times
  out at 45 s: fire them without `await` and store results on `window`, then read them in a second call.
- **ChinaTicketOnline works for the real plan dates** (Dec / Jan): it confirms the train runs and gives the exact sale date. Its fares are
  30–70% above 12306 (e.g. D8 ¥625 vs ¥440).
- Sale-date rule again: D − 14 days at the station's sale time. 17 Dec trains from Fengtai / Beijing West: **3 Dec 08:00**; D136 20 Dec from
  Nanchang: **6 Dec 09:45**; G17 3 Dec: **19 Nov 12:45**; D8 / G32 6 Dec: **22 Nov 14:45 / 13:45**.

### Shenyang without sleeper trains (workbook sheets "… (G-train)")

12306 fares (20 Oct query; fixed): Beijing Chaoyang → Shenyang North G101 / G103 / G105 08:00–08:55 ¥355, G149 12:19 ¥369, G3539 11:43 ¥355,
slow G3503 / G3509 ¥296, G3537 ¥317. Shenyang North → Chaoyang G142 19:51 ¥268, G156 (Shenyang stn) 19:05 ¥268, G3522 19:21 ¥268,
G122 / G126 ¥344. The old ¥339 for a 12 Nov train is not on the 12306 list (¥355–369). Beijing night: **Dequan Railway Station Hotel**
(~1 km from Chaoyang station, 8.7, 65 reviews) ¥304 for 3 ppl / 1 room (Booking.com 9 Oct).

| Option | Univ.-paid 3 / 4 ppl | vs sleeper plan |
|---|---|---|
| Shenyang – Comfort (G-train): MIAT 12 Nov, G149, RiCH 3 nights, G142 back, Dequan night, Air China 16 Nov | ₮6.72M / ₮8.84M | +₮0.54M vs ₮6.18M |
| Shenyang – Budget (G-train): MIAT 13 Nov 17:15, Dequan ×2, G101, Bestay 1 night, G142 | ₮6.05M / ₮8.27M | +₮0.84M vs ₮5.21M |

### Shanghai regional (5–6 Dec 2026, provisional)

- Host Shanghai University, Baoshan campus gym (31.3165, 121.3925; metro line 7 Shangda Rd). 2025 invitation: fee **¥1,500/team**; day 1
  registration 09:00–14:00, opening 14:00–15:00, warm-up 15:00–17:00; day 2 contest 09:00–14:00, analysis 14:30–15:30, awards 15:30–17:00.
  **Wildcard (外卡) by email to shu_icpc@163.com** (first round closed 17 Oct in 2025) – a Mongolian team needs one.
- Flights (Trip.com 10 Oct / SlickTrip): MIAT UB → PEK 3 Dec 07:30 **$151** / $154; Air China PEK → UB 7 Dec 12:55 **$172** (08:25 also $172;
  MIAT 11:30 $193); **Spring nonstop** UB → PVG Tue 1 Dec 13:00 $173 / $164, PVG → UB Tue 8 Dec 08:00 $159 / $153 (sale fares: usual round trip
  $465–700; hand baggage only; flies Tue / Sat). Air China 11:50 UB → PEK $200.
- 12306: Beijing South → Shanghai G17 13:00 → 17:35 **¥667** (G15 12:00 → Hongqiao ¥661 is too tight after MIAT, G1 06:30 ¥598 too early);
  Shanghai → Beijing D8 19:08 → 07:17 (Fengtai 06:58) 2nd-class sleeper **¥440**, D6 21:15 → 09:25 ¥474; day trains from Hongqiao G32 19:25 →
  23:49 **¥598**, G30 18:52 ¥626, G808 / G810 17:25 / 17:46 ¥498.
- Hotels: Booking.com 9 Oct Jenny's Apartment ~1 km (10/10, 19 reviews) 3–6 Dec ¥1,375 / ¥1,726 (3 / 4 ppl), 1–8 Dec ¥2,984 / ¥3,766.
  Trip.com 10 Oct (3 nights, 3 / 4 adults): Zsmart Zhishang 740 m 8.3 ¥1,317 (family room) / ¥1,414; Netfish e-sports 870 m 9.8 (755 reviews)
  ¥1,544 (triple) / ¥1,805 (5-person); Longyang Business 920 m 8.4 ¥1,044; Atour Shangda Rd 830 m 9.7 ¥2,604; 7 nights: Zsmart ¥3,073,
  Netfish ¥3,488 / ¥4,097. No hotel is clearly cheaper than Jenny's near the campus.
- Results (3 / 4 ppl, incl. contingency): **Sleeper** ₮7.48M / ₮9.83M (MIAT + G17, D8 seat back) – recommended; **Spring nonstop** ₮8.21M / ₮10.73M
  (7 nights; ₮7.1M with hand baggage only); **Train (no sleeper)** ₮8.11M / ₮10.78M (G32 + Beijing night).

### Nanchang regional (19–20 Dec 2026, provisional)

- Host Jiangxi Normal University, **Yaohu campus** (紫阳大道 99 号, gym 28.6829 N 116.0323 E). The 2019 handbook (acm.zju.edu.cn
  `南昌2019ICPC参赛手册.pdf`) shows: registration 13:00–18:00 the day before and 08:00–12:00 on day 1, opening 14:30 in the gym, warm-up
  15:00–17:00, contest day 2 09:00–14:00, closing 15:00–16:30; metro line 1 to **Aoti Zhongxin exit 2** (~100 m); from Nanchang station 53 min
  (line 2 + 1, taxi ¥40), from Nanchang West 1 h 8 min (taxi ¥105), from the airport 1 h 50 min. Fee not known (assumed ¥1,500).
- Flights (Trip.com 10 Oct): UB → PEK 17 Dec **Air China CA956 16:10–18:20 $157** (cheapest), MIAT 07:30 $193, Air China 11:50 $363;
  PEK → UB 21 Dec Air China 12:55 / 08:25 **$172**, MIAT 11:30 $193. SlickTrip: Air China through ticket UB → KHN $356 (CA902 11:50 + CA1581
  19:20, lands 21:40); KHN → UB 21 Dec has no same-day connection (Beijing night, $294). MIAT 18 Dec 17:15 only $196.
- 12306 (20 Oct query): Beijing → Nanchang **Z111 Fengtai 23:06 → 11:52 hard sleeper ¥296.5** (soft ¥464.5), K105 Beijing West 23:31 → 16:00
  ¥304.5, **D135 Fengtai 19:48 → Nanchang 07:49 2nd-class sleeper ¥382**, D137 20:00 → 08:01 ¥382, D133 18:12 → Nanchang West 06:30 ¥383;
  daytime G333 08:05 / G335 12:00 / G337 16:55 → Nanchang West ¥742 / ¥745 / ¥732 (≈6 h). Back: **D136 19:22 → Fengtai 07:31 ¥382**,
  D138 19:38 ¥382, D140 Nanchang West 19:18 ¥383, D134 20:58 → 09:01 ¥383, G338 14:42 → Beijing West 20:54 ¥742. D trains arrive at
  Beijing Fengtai (~1 h 15 min from PEK).
- Hotels (Trip.com 10 Oct, 3 adults / 4 adults): **Orange Hotel (Yaohu West Subway Station, Normal University branch)** 730 m, 9.6 (662 reviews):
  2 nights ¥792 / ¥828, 3 nights ¥1,160 / ¥1,214 (Deluxe Family Room for 4); Lavande Aixi Lake 9.4 (2,069) 1.3 km ¥1,216 / 2 nights;
  Xana Hotel Taizidian Station 9.0 (754) 1.3 km ¥1,358 / 3 nights; Meimei Apartment 6.8 ¥105/night; Jiangxi Bailu Hotel 8.8 ¥1,280 / 2 nights.
  Booking.com has almost nothing near the campus (nearest Atour Aixihu 3.6 km ¥2,154 / 3 nights).
- Results (3 / 4 ppl): **Sleeper** (Air China $157 + Z111, 2 nights, D136 seat, Air China $172) ₮6.31M / ₮8.19M – recommended; **Comfort** (MIAT + G335,
  3 nights) ₮7.96M / ₮10.32M; **Fly in** (Air China through ticket $356) ₮8.41M / ₮10.95M.

### Hong Kong (9–10 Jan 2027)

- Flights stay MIAT nonstop: Trip.com 10 Oct round trip 8–11 Jan **$586** (was $583 on 30 Sep; Cathay codeshare $626; via Seoul / Hong Kong
  Airlines $783+), SlickTrip $593, which it flags as low against the usual $708–900. The 2-month SlickTrip calendar has nothing cheaper
  around 8 Jan ($593; 21 Jan $593, 29 Jan $568). Alternatives checked and rejected: UB → Beijing + Cathay PEK → HKG $206 (≈$350–400 each way),
  Aero Mongolia UB → Hohhot ¥892 + Shenzhen Airlines HET → SZX $214 (≈$340), Korean Air via Seoul $515 with an overnight stop.
- Hotels (Booking.com 10 Oct, 3 nights, near HKU map point): Premium Lounge (8.1, 1,242 reviews) triple room **HK$1,177** (3 ppl) / HK$1,611
  (4 ppl: triple + 1 dorm bed) – cheaper than the HK$1,308 / 1,790 of 30 Sep; Kusa Inn (9.6, 15 reviews) HK$1,194 / HK$2,122; Mochi Inn
  (9.1, 17 reviews) 4 ppl HK$2,212; Good Fortune Inn (8.6, 715) HK$1,759 / HK$2,889; Hi Backpackers dorms ~HK$1,950 for 3 beds.
- Results (3 / 4 ppl): Budget ₮9.08M / ₮12.12M, Value (Kusa Inn) ₮9.08M / ₮12.38M, Comfort ₮9.56M / ₮13.15M.

## Funded vs unfunded (10 Oct 2026, incl. 10% contingency)

User answer (10 Oct): funded regional is probably **Hong Kong**, unfunded is likely **Shenyang**; the team will choose. Unfunded 3 ppl =
whole cost per person (travel + stay + fee + other); funded 4 ppl = travel + stay against the ₮12M (`Compare` columns L / M show
"fits incl. contingency?" and "left before contingency"):

| Plan | Unfunded per person | Funded 4 ppl total | Left of ₮12M before contingency |
|---|---|---|---|
| Shenyang – Budget | ₮2.14M | ₮6.96M | ₮5.7M |
| Shenyang – Comfort | ₮2.46M | ₮8.04M | ₮4.7M |
| Nanchang – Sleeper | ₮2.50M | ₮8.19M | ₮4.6M |
| Shanghai – Sleeper | ₮2.89M | ₮9.83M | ₮3.1M |
| Shanghai – Spring nonstop / Train | ₮3.14M / ₮3.10M | ₮10.73M / ₮10.78M | ₮2.2M |
| Nanchang – Comfort / Fly in | ₮3.05M / ₮3.20M | ₮10.32M / ₮10.95M | ₮2.6M / ₮2.0M |
| EC Final – Train | ₮3.14M | ₮11.12M | ₮1.9M |
| Hong Kong – Budget | ₮3.43M | ₮12.12M | ₮1.0M (over with contingency) |
| Hong Kong – Value | ₮3.44M | ₮12.38M | ₮0.7M |
| **Hong Kong – Funded 12M** (BW Plus 2 rooms, airport bus) | ₮3.53M | ₮12.87M | **₮0.3M** (₮11.7M) |
| **Hong Kong – Sleep priority** (Good Fortune Inn, 4 twin beds, airport bus) | ₮3.46M | ₮12.47M | **₮0.67M** (₮11.33M) |
| Hong Kong – Comfort | ₮3.60M | ₮13.15M | ₮0.04M |
| Hong Kong – Sleep first | ₮4.28M | ₮14.45M | over |
| EC Final – Fly | ₮4.05M | ₮14.82M | over |

- Hong Kong flights are 70% of the 4-person cost (₮8.4M; MIAT nonstop $586 return ≈ $293 each way; one-ways are dearer: $343 + $335).
  Savings found: airport bus (~HK$45) instead of the HK$115 Airport Express ≈ ₮0.26M for 4; contingency is only a buffer – book the flights early.
- Shenyang unfunded per person (Budget): flights ₮1.16M (58%), trains ₮0.23M, hotel ₮0.02M, meals + local ₮0.18M, fee ₮0.27M, UB taxi ₮0.07M,
  insurance ₮0.03M, eSIM ₮0.04M, contingency ₮0.16M. Flights are already the cheapest found: MIAT 13 Nov 17:15 $151 (Air China 16:10 $157,
  Hunnu Air UB → Daxing 20:45 $158), Air China PEK → UB 16 Nov 12:55 $172 (MIAT $193, **Hunnu Air Daxing PKX → UB 00:05 $151**).
  Hunnu Air (Mongolian low-cost, Embraer 190, Daxing airport) is new: the 00:05 flight would need G122 18:19 (¥344) from Shenyang to be at Daxing
  by ~22:15, so it saves nothing net versus K54 + Air China.

## Hong Kong: sleep over distance (10 Oct 2026)

Asked what changes if sleep is prioritised over distance to HKU. Findings (Booking.com, 8–11 Jan, 4 adults / 2 rooms, HKD, 3 nights):
Good Fortune Inn **8.6 (715 reviews), 2 rooms with 4 twin beds + private bathrooms, HK$2,889** (cheapest real option with a bed each);
South Nest 8.0 (609) 2 twin rooms HK$3,941; ALVA Hotel by Royal, Tsuen Wan 8.5 (1,038) two queen beds HK$3,890 (~50 min to HKU);
Sleep Inn 8.4 HK$3,401 but one full bed per room (two people share); Kusa Inn 9.6 (15 reviews) double + triple HK$2,122; Mochi Inn 9.1 (17)
HK$2,025; BW Plus (near HKU) HK$3,640 (7.4 on Booking, 8.2 on Trip.com HK$4,243); Premium Lounge triple + dorm bed HK$1,611 (a full bed is
shared in a triple). Trip.com HKU-area results were no better. The 2026 host is not certain (2025 = HKUST), so a campus-adjacent hotel is a bet;
Kowloon guesthouses sit on the MTR for any campus. Chosen: **Hong Kong – Sleep priority** = Good Fortune Inn + A21 airport bus ≈ ₮11.33M before
contingency (₮0.67M spare); commute ~35–40 min each way (wake ~06:15, lights out 22:30 = 7.75 h). Booking.com's price filter is
`nflt=price%3DHKD-min-1450-1%3Breview_score%3D80` (per-night, both rooms).

## Open items / next optimization ideas

- **Book soon (10 Oct):** Spring UB ⇄ PVG fares and MIAT UB–HKG are far below their usual levels (SlickTrip "low"); Air China UB → PEK 17 Dec $157.
- Trains to buy on the sale date: Z366 29 Oct 08:15, K53 30 Oct 10:00, K54 1 Nov 09:00, G17 19 Nov 12:45, D8 22 Nov 14:45,
  Z111 3 Dec 08:00, D136 6 Dec 09:45, D17 11 Jan 10:00. 12306 needs a registered account with passport details.
- Wildcard for Shanghai (shu_icpc@163.com) – ask before booking anything non-refundable. Nanchang fee, schedule and hotel still unconfirmed.
- Not yet researched: Beijing–Shenyang low-cost flights; Trip.com Hong Kong hotels near HKU (earlier check: not cheaper than Booking.com).

- Home-between plan: Air China 08:10 on 10 Nov ($172 vs $219) saves ~₮507k for 3 people but needs a Beijing night on 9 Nov (~300 CNY). Not applied yet.
- Hotels re-checked on Trip.com 1 Oct (see "Trains & hotels" sheet). Bestay's low price is real (8.2, 1,062 reviews). Hong Kong unchanged.
- K53 on 13 Nov leaves ~1 h 20 min after MIAT immigration; if MIAT is late, use the one free change to a 14 Nov G-train (or K341 23:00). Buy sleepers when sales open: Z366 29 Oct 08:15, K53 30 Oct 10:00, K54 1 Nov 09:00.
- EC Fly outbound is two separate tickets via Hohhot: a late Aero Mongolia flight is not protected (fallback Air China 19:40 HET → HGH, or the ¥3,265 through ticket).
- Confirm the 2026 Shenyang fee and the HK and EC Final hosts and venues. All contest dates are provisional.
- EC Final: buy D17 on 12306 at 10:00 on 11 Jan (Spring Festival rush; use the 候补 waitlist if it sells out). The China timetable often changes in January, so re-check D17 in early Jan. It leaves ~8 h after MIAT lands, which gives plenty of buffer. Re-check separate MIAT + domestic tickets ~40 days ahead (mid-Dec) in case discount fares appear.
- PEK international ⇄ domestic transfer: bags may have to be collected for customs and re-dropped. Keep connections ≥ 2 h 30 min.
