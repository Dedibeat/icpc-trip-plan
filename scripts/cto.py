"""List ChinaTicketOnline trains with official fares, CTO prices and sale times (plain HTTP, works from the cloud).

Usage:  python scripts/cto.py FROM_CN TO_CN YYYY-MM-DD [TRAIN,TRAIN,...]
  FROM_CN/TO_CN are Chinese station or city names (沈阳, 北京, 呼和浩特, 集宁南, 二连 – not 二连浩特).
  Works for any date, also before 12306 sales open. Each seat shows: official fare ¥ (CTO price $ + CTO fee $, seats left).
  "seats left" is a placeholder until 12306 sales open.
"""
import json, sys, urllib.parse, urllib.request

fr, to, date = sys.argv[1:4]
want = sys.argv[4].split(",") if len(sys.argv) > 4 else None
data = urllib.parse.urlencode({"action": "getTrainListAjax", "from": fr, "to": to, "date": date, "trainCat": ""}).encode()
req = urllib.request.Request("https://www.chinaticketonline.com/wp-admin/admin-ajax.php", data=data, headers={
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) Chrome/124.0", "Referer": "https://www.chinaticketonline.com/",
    "X-Requested-With": "XMLHttpRequest"})
rows = json.loads(urllib.request.urlopen(req, timeout=60).read().decode())
seats = {"hardSeat": "seat", "hardSleeper": "hardsl", "softSleeper": "softsl", "secondClassSeat": "2nd", "firstClassSeat": "1st",
         "secondClassSleeper": "2ndsl", "firstClassSleeper": "1stsl", "standing": "stand"}
print(f"{fr} -> {to} {date}: {len(rows)} trains")
for r in rows:
    if want and r["trainCode"] not in want:
        continue
    fares = [f'{n} ¥{r[k]["priceCNY"]} (${r[k]["price"]}+${r[k]["serviceFee"]}, left {r[k]["remainNum"]})' for k, n in seats.items() if r.get(k)]
    print(f'{r["trainCode"]:6} {r["startCityCN"]}{r["startTime"]} -> {r["endCityCN"]}{r["endTime"]} ({r["costTime"]}) '
          f'sale {r["preSaleTime"]} | ' + "; ".join(fares))
