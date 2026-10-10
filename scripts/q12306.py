"""Query 12306 train list + official fares for one route and date (no browser, no login).

Usage:  python scripts/q12306.py FROM TO YYYY-MM-DD [MAX_MINUTES|all]
  FROM/TO are 12306 telegraph codes (BJP Beijing, SHH Shanghai, SBT Shenyang North, NCG Nanchang ...).
  The date must be inside the 15-day sale window; fares are fixed, so query a near date for later trips.
  On Windows set Q12306_INSECURE=1 (12306's CA is not in the Windows certificate store).
  MAX_MINUTES keeps G/D trains up to that duration plus every D/Z/T/K train leaving after 17:00.
  Sleeper fares for D trains (1stsl/2ndsl) are in tenths of a yuan (4400 = ¥440).
"""
import http.cookiejar, json, os, ssl, sys, time, urllib.parse, urllib.request

UA = "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0 Safari/537.36"
cj = http.cookiejar.CookieJar()
handlers = [urllib.request.HTTPCookieProcessor(cj)]
if os.environ.get("Q12306_INSECURE"):  # 12306 uses a CA missing from the Windows store; data is public, read-only
    handlers.append(urllib.request.HTTPSHandler(context=ssl._create_unverified_context()))
op = urllib.request.build_opener(*handlers)
op.addheaders = [("User-Agent", UA), ("Referer", "https://kyfw.12306.cn/otn/leftTicket/init")]

def get(url):
    for i in range(3):
        try:
            return op.open(url, timeout=30).read().decode("utf-8")
        except Exception as e:
            err = e; time.sleep(2)
    raise err

get("https://kyfw.12306.cn/otn/leftTicket/init?linktypeid=dc")
fr, to, date = sys.argv[1:4]
mode = sys.argv[4] if len(sys.argv) > 4 else "all"
q = get(f"https://kyfw.12306.cn/otn/leftTicket/queryG?leftTicketDTO.train_date={date}&leftTicketDTO.from_station={fr}&leftTicketDTO.to_station={to}&purpose_codes=ADULT")
d = json.loads(q)["data"]
names = d["map"]
rows = []
for r in d["result"]:
    f = r.split("|")
    rows.append(dict(no=f[2], code=f[3], fs=names.get(f[6], f[6]), ts=names.get(f[7], f[7]), dep=f[8], arr=f[9], dur=f[10],
                     fsn=f[16], tsn=f[17], seats=f[35]))
def mins(s):
    h, m = s.split(":"); return int(h) * 60 + int(m)
sel = []
for r in rows:
    over = mins(r["dep"]) >= 17 * 60 and r["code"][0] in "DZTK"
    if mode == "all" or over or (r["code"][0] in "GD" and mins(r["dur"]) <= int(mode)):
        sel.append(r)
print(f"{len(rows)} trains; showing {len(sel)}")
for r in sel:
    p = get("https://kyfw.12306.cn/otn/leftTicket/queryTicketPrice?" + urllib.parse.urlencode(dict(
        train_no=r["no"], from_station_no=r["fsn"], to_station_no=r["tsn"], seat_types=r["seats"], train_date=date)))
    try:
        pd = json.loads(p)["data"]
    except Exception:
        pd = {}
    keys = {"O": "2nd", "M": "1st", "A9": "biz", "A1": "hardseat", "A3": "hardsl", "A4": "softsl", "I": "1stsl", "J": "2ndsl", "F": "dongwo"}
    fares = {keys[k]: v for k, v in pd.items() if k in keys}
    print(f'{r["code"]:7} {r["fs"]}->{r["ts"]} {r["dep"]}-{r["arr"]} ({r["dur"]}) {fares}')
    time.sleep(0.6)
