#!/usr/bin/env python3
"""
Builds the Girls Inc. Outputs and Outcomes report (index.html) from two spreadsheets.

  data/master_query*.xlsx  -> metric results (Clients Served, Youth Opportunity)
  data/master_file*.xlsx   -> client-level records (Demographics, Service Map)

If several files match, the newest one (by file name, which includes the export
timestamp) is used. Usage:

    pip install openpyxl
    python build.py                      # writes index.html
    python build.py --out _site/index.html
"""
import argparse
import glob
import json
import os
import re
import sys
import warnings
from collections import Counter, defaultdict
from datetime import datetime

import openpyxl

warnings.filterwarnings("ignore")

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "data")
QUARTERS = ["Q1", "Q2", "Q3", "Q4", "Total"]


# ---------------------------------------------------------------- helpers
def newest(pattern):
    files = [
        f for f in sorted(glob.glob(os.path.join(DATA_DIR, pattern)))
        if not os.path.basename(f).startswith("~$")
    ]
    if not files:
        sys.exit(f"ERROR: no file matching data/{pattern} found.")
    return files[-1]


def read_table(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h).strip() if h is not None else "" for h in rows[0]]
    out = []
    for r in rows[1:]:
        if r is None or all(v is None or str(v).strip() == "" for v in r):
            continue
        out.append({header[i]: r[i] for i in range(len(header))})
    return header, out


def clean(v):
    """Trim/collapse whitespace; blank -> None."""
    if v is None:
        return None
    if isinstance(v, str):
        v = re.sub(r"\s+", " ", v).strip()
        return v or None
    return v


def need(header, cols, name):
    missing = [c for c in cols if c not in header]
    if missing:
        sys.exit(f"ERROR: {name} is missing column(s): {', '.join(missing)}")


def stamp_from_name(path):
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})T", os.path.basename(path))
    if m:
        return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3)))
    return datetime.fromtimestamp(os.path.getmtime(path))


# ---------------------------------------------------------------- metrics
def build_metrics(path):
    header, rows = read_table(path)
    need(header, ["Section", "MetricCode", "Metric", "Period"], os.path.basename(path))

    overview = defaultdict(dict)
    youth_raw = defaultdict(dict)
    youth_names = {}
    for r in rows:
        section, code, period = clean(r["Section"]), clean(r["MetricCode"]), clean(r["Period"])
        if not (section and code and period):
            continue
        if section == "Overview":
            if r.get("Value") is not None:
                overview[code][period] = float(r["Value"])
        elif section == "Youth Opportunity":
            served = r.get("Total Served")
            succ = r.get("Successful")
            pct = r.get("Success Percent")
            youth_raw[code][period] = [
                float(served) if served is not None else 0.0,
                float(succ) if succ is not None else 0.0,
                float(pct) if pct is not None else None,
            ]
            youth_names[code] = clean(r.get("Metric")) or code

    # Q1 new clients = everyone served in Q1 (there is no earlier quarter)
    if "Q1" in overview.get("CLIENTS", {}):
        overview["NEW"].setdefault("Q1", overview["CLIENTS"]["Q1"])

    with open(os.path.join(ROOT, "config", "youth_labels.json")) as f:
        labels = json.load(f)
    known = {l["code"] for l in labels}
    labels += [
        {"code": c, "name": youth_names[c], "served_label": "Served", "success_label": "Successful"}
        for c in youth_raw if c not in known
    ]
    youth = []
    for l in labels:
        if l["code"] not in youth_raw:
            continue
        q = {p: youth_raw[l["code"]].get(p, [0.0, 0.0, None]) for p in QUARTERS}
        youth.append({**l, "q": q})
    return dict(overview), youth


# ---------------------------------------------------------------- clients
LABELS = {
    "NotHispanicOrLatino": "Not Hispanic or Latino",
    "HispanicOrLatino": "Hispanic or Latino",
    "BlackAfricanAmerican": "Black / African American",
}


def pretty(token):
    if token in LABELS:
        return LABELS[token]
    return re.sub(r"(?<!^)(?=[A-Z])", " ", token)


def multi_label(v):
    """'A|B|A|' -> 'A & B' (duplicates and empty parts removed)."""
    v = clean(v)
    if v is None:
        return "Unknown"
    seen = []
    for part in str(v).split("|"):
        part = part.strip()
        if part:
            lbl = pretty(part)
            if lbl not in seen:
                seen.append(lbl)
    return " & ".join(seen) if seen else "Unknown"


def plain(v):
    v = clean(v)
    return "Unknown" if v is None else str(v)


def count_list(values, key=None):
    c = Counter(values)
    items = [i for i in c.items() if i[0] != "Unknown"]
    items.sort(key=key or (lambda x: (-x[1], x[0])))
    if "Unknown" in c:
        items.append(("Unknown", c["Unknown"]))
    return [list(i) for i in items]


def income_key(item):
    label = item[0]
    if label.lower().startswith("less than"):
        return (0, 0, label)
    nums = re.findall(r"\d[\d,]*", label)
    return (1, int(nums[0].replace(",", "")) if nums else 10**12, label)


GRADE_ORDER = ["PK", "GK"] + [f"G{i}" for i in range(1, 13)]


def grade_sort(g):
    if g in GRADE_ORDER:
        return GRADE_ORDER.index(g)
    m = re.match(r"PS(\d+)", g)
    if m:
        return 100 + int(m.group(1))
    return 999


def state_from_zip(z):
    p = int(z[:3])
    if 300 <= p <= 319:
        return "GA"
    if 350 <= p <= 369:
        return "AL"
    return None


def build_clients(path):
    header, rows = read_table(path)
    need(
        header,
        ["Ethnicity", "Race", "County", "UW Income Range", "ALICE Household Type",
         "ZipCode", "School Name", "Grade Level"],
        os.path.basename(path),
    )
    total = len(rows)

    demographics = {
        "ethnicity": count_list([multi_label(r["Ethnicity"]) for r in rows]),
        "race": count_list([multi_label(r["Race"]) for r in rows]),
        "income": count_list([plain(r["UW Income Range"]) for r in rows], key=income_key),
        "household": count_list([plain(r["ALICE Household Type"]) for r in rows]),
    }

    # ---- zip / county / grade
    zip_votes = defaultdict(Counter)
    zip_grades = defaultdict(Counter)
    unmapped = 0
    for r in rows:
        zraw = clean(r["ZipCode"])
        m = re.match(r"\d{5}", str(zraw)) if zraw is not None else None
        grade = plain(r["Grade Level"])
        if not m or state_from_zip(m.group(0)) is None:
            unmapped += 1
            continue
        z = m.group(0)
        county = plain(r["County"])
        if county != "Unknown":
            zip_votes[z][county] += 1
        zip_grades[z][grade] += 1

    zips = {}
    for z, gr in zip_grades.items():
        county = zip_votes[z].most_common(1)[0][0] if zip_votes[z] else "Unknown"
        zips[z] = {
            "county": county,
            "state": state_from_zip(z),
            "total": sum(gr.values()),
            "grades": dict(sorted(gr.items(), key=lambda kv: grade_sort(kv[0]))),
        }

    county_totals = defaultdict(lambda: defaultdict(int))
    for v in zips.values():
        county_totals[v["state"]][v["county"]] += v["total"]
    counties = {
        st: [c for c, _ in sorted(county_totals[st].items(), key=lambda kv: (-kv[1], kv[0]))]
        for st in ("GA", "AL")
    }

    # ---- grade table + buckets (all records, including unmapped zips)
    gc = Counter(plain(r["Grade Level"]) for r in rows)
    grade_table = [[g, gc[g]] for g in sorted((g for g in gc if g != "Unknown"), key=grade_sort)]
    if gc.get("Unknown"):
        grade_table.append(["Unknown", gc["Unknown"]])

    buckets = {"Elementary (K-5)": 0, "Middle (6-8)": 0, "High School (9-12)": 0,
               "College (PS)": 0, "Unknown": 0}
    for g, c in gc.items():
        if g in ("GK", "PK") or re.fullmatch(r"G[1-5]", g):
            buckets["Elementary (K-5)"] += c
        elif re.fullmatch(r"G[6-8]", g):
            buckets["Middle (6-8)"] += c
        elif re.fullmatch(r"G(9|1[0-2])", g):
            buckets["High School (9-12)"] += c
        elif g.startswith("PS"):
            buckets["College (PS)"] += c
        else:
            buckets["Unknown"] += c

    # ---- schools: fewest participants first (ascending), Unknown last
    schools = count_list(
        [plain(r["School Name"]) for r in rows], key=lambda x: (x[1], x[0].lower())
    )

    return {
        "demographics": demographics,
        "map": {"zips": zips, "counties": counties, "unmapped": unmapped},
        "grade_table": grade_table,
        "grade_buckets": [[k, v] for k, v in buckets.items()],
        "schools": schools,
        "total": total,
    }


# ---------------------------------------------------------------- main
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.join(ROOT, "index.html"))
    args = ap.parse_args()

    q_path = newest("master_query*.xlsx")
    f_path = newest("master_file*.xlsx")
    print("Metrics file :", os.path.relpath(q_path, ROOT))
    print("Clients file :", os.path.relpath(f_path, ROOT))

    overview, youth = build_metrics(q_path)
    demo = build_clients(f_path)
    asof = stamp_from_name(q_path)
    data = {
        "overview": overview,
        "youth": youth,
        "demo": demo,
        "asof": f"{asof:%B} {asof.day}, {asof.year}",
    }

    with open(os.path.join(ROOT, "template.html"), encoding="utf-8") as f:
        html = f.read()
    if "/*__DATA__*/{}" not in html:
        sys.exit("ERROR: template.html is missing the /*__DATA__*/{} placeholder.")
    payload = json.dumps(data, separators=(",", ":")).replace("</", "<\\/")
    html = html.replace("/*__DATA__*/{}", payload)

    os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
    with open(args.out, "w", encoding="utf-8") as f:
        f.write(html)
    print(f"Built {os.path.relpath(args.out, ROOT)}  "
          f"({demo['total']} clients, {len(demo['schools'])} schools, "
          f"{len(demo['map']['zips'])} zip codes, data as of {data['asof']})")


if __name__ == "__main__":
    main()
