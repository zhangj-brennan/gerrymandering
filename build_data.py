"""Split the Excel workbook into one CSV per sheet, then rewrite the data
embedded in each chart page from those CSVs.

Run after updating the workbook:  python3 build_data.py
"""
import csv
import json
import re
from collections import Counter
from pathlib import Path

import openpyxl

ROOT = Path(__file__).parent
DATA = ROOT / "10142026_data"
XLSX = DATA / "Tennessee county population change 1900-1960.xlsx"
PREFIX = "Tennessee county population change 1900-1960 - "
YEARS = [1900, 1910, 1920, 1930, 1940, 1950, 1960]


def num(v):
    """Whole-number floats as ints, so CSVs read 23615 rather than 23615.0."""
    return int(v) if isinstance(v, float) and v.is_integer() else v


# ---------- workbook -> CSVs ----------
wb = openpyxl.load_workbook(XLSX, data_only=True)
for ws in wb:
    rows = [[("" if v is None else num(v)) for v in row] for row in ws.iter_rows(values_only=True)]
    while rows and not any(str(v).strip() for v in rows[-1]):
        rows.pop()
    with open(DATA / f"{PREFIX}{ws.title}.csv", "w", newline="") as f:
        csv.writer(f).writerows(rows)


def read(name):
    with open(DATA / f"{PREFIX}{name}.csv", newline="") as f:
        return list(csv.reader(f))


POP_COLS = [2, 5, 8, 11, 14, 17, 20]   # 1900 ... 1960 in the district sheets

# Statewide totals from the counties sheet.
counties = [r for r in read("Tennessee counties")[1:] if r[0].strip()]
TOTALS = [sum(int(r[i + 1]) for r in counties) for i in range(len(YEARS))]

# House: one row per district. For a county with more than 1 seat, the sheet
# gives population per seat; multiply back to the whole county.
house = [r for r in read("House districts")[1:] if r[0].strip() and r[1].strip()]
HOUSE = [{"name": r[0], "seats": int(r[1]),
          "pop": [round(float(r[c]) * int(r[1])) for c in POP_COLS]} for r in house]
N_HOUSE = sum(d["seats"] for d in HOUSE)
HOUSE_IDEAL = [t / N_HOUSE for t in TOTALS]

# Senate: a county with more than 1 seat is listed once per seat, each row
# carrying the whole county; divide it equally.
senate = [r for r in read("Senate districts")[1:] if r[0].strip()]
repeats = Counter(r[1] for r in senate)
SENATE = []
for r in senate:
    n = repeats[r[1]]
    pop = [round(float(r[c]) / n) for c in POP_COLS]
    SENATE.append({"name": r[1], "seat": int(r[0]), "seats": n, "v": pop[0], "pop": pop})
SENATE_IDEAL = [t / len(SENATE) for t in TOTALS]


def js(v):
    return json.dumps(v, separators=(",", ":"))


def js_list(v):
    return json.dumps(v)


def replace_line(text, start, new, path):
    """Replace the single line beginning with `start`."""
    pattern = re.compile("^" + re.escape(start) + ".*$", re.M)
    if len(pattern.findall(text)) != 1:
        raise SystemExit(f"{path.name}: expected one line starting {start!r}")
    return pattern.sub(lambda _: new, text)


def patch(name, *edits):
    path = ROOT / name
    text = path.read_text()
    for start, new in edits:
        text = replace_line(text, start, new, path)
    path.write_text(text)


# Pages with population per seat (multi-seat counties get one circle per seat).
per_seat = [{"name": d["name"], "v": round(d["pop"][0] / d["seats"]),
             "pop": [round(p / d["seats"]) for p in d["pop"]]} for d in HOUSE]
seats = {d["name"]: d["seats"] for d in HOUSE if d["seats"] > 1}
for page in ["ideal_population_balance.html", "ideal_population_circles.html",
             "ideal_population_trails.html", "ideal_population_box.html"]:
    patch(page,
          ("const YEAR_TOTAL = ", f"const YEAR_TOTAL = {js_list(TOTALS)};"),
          ("const YEAR_IDEAL = ", f"const YEAR_IDEAL = {js_list([round(v, 1) for v in HOUSE_IDEAL])};"),
          ("const COUNTIES = ", f"const COUNTIES = {js(per_seat)};"),
          ("const SEATS = ", f"const SEATS = {js(seats)};"))

# Pages with whole-district totals and a seat count.
patch("ideal_population_line.html",
      ("const DATA = ", "const DATA = " + js({"years": YEARS, "ideal": HOUSE_IDEAL, "districts": HOUSE}) + ";"))
patch("house_districts_chart.html",
      ("const DATA = ", "const DATA = " + js({"years": YEARS, "ideal": [round(v) for v in HOUSE_IDEAL],
                                               "districts": HOUSE}) + ";"))

patch("ideal_population_trails_senate.html",
      ("const YEAR_TOTAL = ", f"const YEAR_TOTAL = {js_list(TOTALS)};"),
      ("const YEAR_IDEAL = ", f"const YEAR_IDEAL = {js_list([round(v, 1) for v in SENATE_IDEAL])};"),
      ("const DISTRICTS = ", f"const DISTRICTS = {js(SENATE)};"))

print(f"House: {len(HOUSE)} districts, {N_HOUSE} seats. Senate: {len(SENATE)} seats.")
print("Totals:", TOTALS)
