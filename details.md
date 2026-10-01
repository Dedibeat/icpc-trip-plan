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
| `ICPC_flight_fares_30Sep2026.xlsx` | Trip.com fares found via the Bright Data scraper, compared against the fares the plans used before. Rows 11–22 checked 30 Sep, rows 38–45 checked 1 Oct 2026. Has a "Connector notes" sheet. |
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
- Single-contest sheets: row 10 = outbound flight, row 11 = return flight (EC Fly: rows 10–13). Travel times are stored as `timedelta` with format `[h]" h "mm" min"`.
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
| **UB → Hangzhou through** | 25 Jan | **Air China ¥3,265**, UB 11:50 → PEK (2 h 35 min) → HGH 18:45 | Same price as separate MIAT + domestic. One ticket, bags through. Backup: MIAT + Hong Kong Airlines via HKG ¥2,721, lands 23:40 |
| **Hangzhou → UB through** | 29 Jan | **Air China ¥2,250**, HGH 08:00 → PEK (2 h 35 min) → UB 15:05 | Cheaper than domestic + Beijing night + MIAT 30 Jan (~¥3,280 + hotel). Home a day earlier |
| Shenyang → UB through | 16 Nov | Air China $286, SHE 09:00 → UB 15:05 | ~$50 pp dearer than train + Beijing night + Air China. Use only if the closing runs late |
| UB → Nanjing / Nanjing → UB through | 6 / 10 Nov | $353 / $360–394 via Beijing | Dearer than MIAT + fast train. Not used |

- Air China UB↔Beijing is often cheaper than MIAT on the return leg. SlickTrip shows Air China as $0 (unpriced), so check it on Trip.com.
- Trip.com offers "Student tickets" on some routes. Check them at booking.
- Domestic Chinese fares 4 months out show near-full fares on Trip.com (CNY). Use `curr=CNY`. Domestic routes redirect to `/chinaflights/ShowFareFirst`, and the date strip shows the cheapest fare for each day.
- **Through tickets** (one Air China ticket UB ⇄ Chinese city via Beijing Capital) are worth checking whenever a plan has a Beijing night + domestic leg. They won for Hangzhou (EC Final), but not for Nanjing or Shenyang, where high-speed rail is cheap.
- Skyscanner China route pages have no data this far out (Jan showed "查找价格"). Ctrip's own API (`flights.ctrip.com/itinerary/api/12808/lowestPrice`) is blocked by Bright Data without KYC.
- If a workbook is open in Excel, saving fails with PermissionError (look for `~$*.xlsx` lock files). Ask the user to close it.

## Current results (after Trip.com update, incl. 10% contingency on the university part)

| Plan | Univ.-paid 3 ppl | Univ.-paid 4 ppl | Out-of-pocket pp (3 / 4) |
|---|---|---|---|
| Shenyang – Budget | ₮5.91M | ₮8.41M | ₮401k / ₮317k |
| Shenyang – Comfort (recommended) | ₮7.01M | ₮9.57M | ₮401k / ₮317k |
| Hong Kong – Budget (recommended) | ₮9.11M | ₮12.17M | ₮409k / ₮323k |
| Hong Kong – Comfort | ₮9.53M | ₮13.10M | ₮409k / ₮323k |
| EC Final – Train (rail out, Air China through ticket home 29 Jan) – recommended | ₮8.87M | ₮11.99M | ₮401k / ₮317k |
| EC Final – Fly (Air China through tickets both ways) | ₮12.07M | ₮16.34M (₮1.34M over) | ₮401k / ₮653k |

Combined Nanjing + Shenyang (3 ppl, total incl. fees and contingency): A Balanced ₮14.00M, B Budget ₮11.67M,
C Home between ₮16.97M (over budget).

## Open items / next optimization ideas

- Home-between plan: Air China 08:10 on 10 Nov ($172 vs $219) saves ~₮507k for 3 people but needs a Beijing night on 9 Nov (~300 CNY). Not applied yet.
- Re-check hotels (Bestay Shenyang at 86 CNY/night looks too cheap). Also price 4-person plans as a triple + single room.
- Confirm the 2026 Shenyang fee and the HK and EC Final hosts and venues. All contest dates are provisional.
- EC Final: buy the 25 Jan Beijing → Hangzhou train the day sales open (~10 Jan). Re-check separate MIAT + domestic tickets ~40 days ahead (mid-Dec) in case discount fares appear.
- PEK international ⇄ domestic transfer: bags may have to be collected for customs and re-dropped. Keep connections ≥ 2 h 30 min.
