"""Snapshot Ctrip's low-price calendar (cheapest fare per departure date, ~1 year ahead) and log it to CSV.

Usage:  python scripts/ctrip_calendar.py [ROUTE ...] [--from YYYY-MM-DD --to YYYY-MM-DD]
  ROUTE is FROM-TO in Ctrip city codes (ULN Ulaanbaatar, HET Hohhot, SHE Shenyang, SHA Shanghai, BJS Beijing).
  Default routes: ULN-HET HET-SHE ULN-SHA SHA-SHE ULN-SHE.
  Appends one row per route and departure date to data/ctrip_calendar.csv (a re-run on the same day replaces
  that day's rows), then prints the --from/--to window (default 2026-11-10 to 2026-11-14) for every snapshot.
  price = fare before tax (domestic) / base fare (international); total = incl. taxes. 0 = no fare listed.
  The calendar is the cheapest fare of ANY airline that day, not one flight. Run it daily to build a history.
"""
import csv, datetime as dt, json, os, re, sys, urllib.request

API = "https://m.ctrip.com/restapi/soa2/15380/bjjson/FlightIntlAndInlandLowestPriceSearch"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
CSV = os.path.join(os.path.dirname(__file__), "..", "data", "ctrip_calendar.csv")
CN = dt.timezone(dt.timedelta(hours=8))

def calendar(fr, to, start):
    body = {"departNewCityCode": fr, "arriveNewCityCode": to, "startDate": start, "grade": 3, "flag": 0,
            "channelName": "FlightIntlOnline", "searchType": 1,
            "passengerList": [{"passengercount": 1, "passengertype": "Adult"}],
            "calendarSelections": [{"selectionType": 8, "selectionContent": ["3"]}]}
    req = urllib.request.Request(API, json.dumps(body).encode(), {"content-type": "application/json", "User-Agent": UA})
    j = json.loads(urllib.request.urlopen(req, timeout=30).read())
    for p in j.get("priceList") or []:
        ms = int(re.search(r"\d+", p["departDate"]).group())
        yield dt.datetime.fromtimestamp(ms / 1000, CN).date().isoformat(), round(p["price"]), round(p["totalPrice"])

args = sys.argv[1:]
win = ["2026-11-10", "2026-11-14"]
for i, k in enumerate(("--from", "--to")):
    if k in args:
        j = args.index(k); win[i] = args[j + 1]; del args[j:j + 2]
routes = args or ["ULN-HET", "HET-SHE", "ULN-SHA", "SHA-SHE", "ULN-SHE"]
today = dt.datetime.now(CN).date().isoformat()

rows = []
if os.path.exists(CSV):
    with open(CSV, newline="", encoding="utf-8") as f:
        rows = [r for r in csv.DictReader(f) if not (r["snapshot"] == today and r["route"] in routes)]
for rt in routes:
    fr, to = rt.split("-")
    rows += [dict(snapshot=today, route=rt, depart=d, price=p, total=t) for d, p, t in calendar(fr, to, today)]
os.makedirs(os.path.dirname(CSV), exist_ok=True)
with open(CSV, "w", newline="", encoding="utf-8") as f:
    w = csv.DictWriter(f, ["snapshot", "route", "depart", "price", "total"]); w.writeheader(); w.writerows(rows)

days = sorted({r["depart"] for r in rows if win[0] <= r["depart"] <= win[1]})
for rt in routes:
    print(f"\n{rt}  total incl. tax (CNY), departures {win[0]} .. {win[1]}")
    print("snapshot    " + " ".join(d[5:] for d in days))
    for snap in sorted({r["snapshot"] for r in rows if r["route"] == rt}):
        got = {r["depart"]: r["total"] for r in rows if r["route"] == rt and r["snapshot"] == snap}
        print(snap + "  " + " ".join(f"{got.get(d, '-'):>5}" for d in days))
