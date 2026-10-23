"""Print every room offer in pages saved by scripts/trip_hotel_detail.js.

Usage:  python scripts/parse_trip_detail.py OUT.txt [OUT2.txt ...]
  One line per offer: file | team total CNY | room | beds | adults | cancellation policy.
  Multi-night pages say "Total price: CNY n", one-night pages "Total (incl. taxes & fees): CNY n".
  Room names can come out as a photo count ("8"); the popup lines (POPUPn) in the file give the full name and bed widths.
"""
import re, sys

bed = re.compile(r"^(\d+ (queen|double|single|king|twin) beds?|Bed type randomly)")
for f in sys.argv[1:]:
    t = open(f, encoding="utf-8").read()
    m = re.search(r"^BODY: (.*)$", t, re.M)
    toks = [x.strip() for x in m.group(1).split(" / ")] if m else []
    if not toks or "Membership rewards" in toks[:40]:
        print(f, "| no rooms (sign-in wall: check the slug)")
        continue
    start = next((i for i, x in enumerate(toks) if x in ("Rooms with windows", "Select multiple rooms to meet your requirements:")), None)
    if start is None:
        start = next((i for i, x in enumerate(toks) if x == "View on map"), 0)
    room = beds = pol = ad = None
    seen = set()
    for i in range(start, len(toks)):
        x = toks[i]
        if x in ("Guest Reviews", "What guests say", "More Room Types"):
            break
        if i + 1 < len(toks) and bed.match(toks[i + 1]) and not bed.match(x):
            room, beds, pol, ad = x, toks[i + 1], None, None
        if "ancel" in x or "refund" in x:
            pol = x
        if re.fullmatch(r"\d+ adults?", x):
            ad = x
        mm = re.match(r"Total (?:price|\(incl\. taxes & fees\)): CNY ([\d,]+)", x)
        if mm and room and (room, mm.group(1)) not in seen:
            seen.add((room, mm.group(1)))
            print(f"{f} | {mm.group(1)} | {room[:60]} | {beds} | {ad or ''} | {pol or ''}")
