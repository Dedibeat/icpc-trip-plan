"""Recalculate the ICPC budget workbooks without Excel/LibreOffice and print key totals.

Usage:  .venv/Scripts/python scripts/recalc_check.py [workbook.xlsx ...]
Fails (exit 1) if any formula evaluates to an Excel error (#REF!, #VALUE!, ...).
"""
import glob
import sys

import formulas

TOTAL_LABELS = (
    "University-paid total incl. contingency",
    "Funding left (negative = over)",
    "Out-of-pocket per person (personal + any overrun)",
    "TOTAL incl. contingency",
    "Funding left after subtotal (negative = over)",
    "Per person incl. contingency",
)


def recalc(path):
    model = formulas.ExcelModel().loads(path).finish()
    sol = model.calculate()
    cells = {}
    for key, rng in sol.items():
        # key looks like "'[file.xlsx]SHEET NAME'!A1" (sheet names upper-cased by formulas)
        try:
            sheet, ref = key.rsplit("!", 1)
        except ValueError:
            continue
        sheet = sheet.strip("'").split("]", 1)[-1]
        val = rng.value[0, 0] if hasattr(rng, "value") else rng
        cells[(sheet, ref)] = val
    return cells


def main(paths):
    bad = 0
    for path in paths:
        print(f"== {path}")
        cells = recalc(path)
        errors = [(s, r, v) for (s, r), v in cells.items() if isinstance(v, str) and v.startswith("#")]
        errors += [(s, r, v) for (s, r), v in cells.items() if type(v).__name__ == "XlError"]
        for s, r, v in errors:
            print(f"   ERROR {s}!{r}: {v}")
        bad += len(errors)
        # print every labelled total: label in column G (single plans) or G (combined) -> value in H and K
        labels = {(s, r): v for (s, r), v in cells.items() if isinstance(v, str) and v in TOTAL_LABELS}
        for (s, r), label in sorted(labels.items()):
            row = r.lstrip("ABCDEFGHIJKLMNOPQRSTUVWXYZ")
            vals = [cells.get((s, f"{c}{row}")) for c in "HK"]
            shown = "  ".join(f"{v:>14,.0f}" for v in vals if isinstance(v, (int, float)))
            print(f"   {s:<28} {label:<50} {shown}")
    print("errors:", bad)
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:] or sorted(glob.glob("*.xlsx"))))
