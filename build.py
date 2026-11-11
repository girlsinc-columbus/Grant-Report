#!/usr/bin/env python3
"""
Builds the Girls Inc. Outputs and Outcomes report (index.html) from three spreadsheets.

  data/master_query*.xlsx     -> metric results (Clients Served, Youth Opportunity)
  data/master_file*.xlsx      -> client-level records (Demographics, Service Map)
  data/GeocodeResults*.xlsx   -> census tract for each master_file record, used for
                                  the Service Map. Produced by uploading master_file's
                                  addresses, IN ROW ORDER, to the Census Bureau batch
                                  geocoder (geocoding.geo.census.gov/geocoder, "Address
                                  Batch" -> "Find Matching Geographies"). The geocoder's
                                  RECORD ID NUMBER column must equal each address's row
                                  number in master_file (1 = first data row), which is
                                  what the geocoder returns when the input ID column is
                                  built the same way.

If several files match a pattern, the newest one (by last-modified time) is used.
Only aggregated counts are ever written to index.html -- never names, addresses,
coordinates, or any other per-person detail. Usage:

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
def newest(pattern, required=True):
    files = [
        f for f in glob.glob(os.path.join(DATA_DIR, pattern))
        if not os.path.basename(f).startswith("~$")
    ]
    if not files:
        if required:
            sys.exit(f"ERROR: no file matching data/{pattern} found.")
        return None
    def when(f):
        # Prefer the export timestamp in the file name (reliable on GitHub, where
        # every file's modified time is just the moment of checkout); fall back
        # to modified time for names without one.
        m = re.search(r"(\d{4}-\d{2}-\d{2})T(\d{2})_(\d{2})_(\d{2})", os.path.basename(f))
        return (1, m.group(0)) if m else (0, "%020.3f" % os.path.getmtime(f))
    files.sort(key=when)
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


STATE_FIPS = {"13": "GA", "01": "AL"}
# Known counties, so the map can label them by name instead of a FIPS code.
# Add an entry here if a new county starts showing up in the geocode results.
COUNTY_FIPS = {
    ("13", "215"): "Muscogee County",
    ("13", "145"): "Harris County",
    ("01", "113"): "Russell County",
    ("01", "081"): "Lee County",
}


def clean_tract(tract6):
    """'001000' -> '10', '003301' -> '33.01', '010106' -> '101.06'."""
    base, suffix = tract6[:4], tract6[4:]
    label = str(int(base))
    if suffix and suffix != "00":
        label += "." + suffix
    return label


def read_geocode_tracts(path):
    """
    Parses GeocodeResults' RECORD ID NUMBER column, which echoes back whatever
    was in the "Unique ID" column of the batch uploaded to the Census
    geocoder. Returns {key: tract info} for matched addresses, where key is
    either:
      - a 0-based row INDEX into whatever master_file the batch was built
        from, if the Unique ID was a plain row number (the old convention,
        kept for any already-geocoded batch using it), or
      - the client's Constituent ID directly, as a string, if the Unique ID
        was built from Constituent ID (the current convention -- see
        tools/build_geocoder_upload.py). This case needs no row-position
        matching at all.
    """
    if not path:
        return {}
    header, rows = read_table(path)
    need(
        header,
        ["RECORD ID NUMBER", "TIGER MATCH TYPE", "STATE CODE", "COUNTY CODE", "TRACT CODE"],
        os.path.basename(path),
    )
    out = {}
    for r in rows:
        if clean(r.get("TIGER MATCH TYPE")) is None:
            continue  # no match / tie -- no tract to use
        rid = clean(r.get("RECORD ID NUMBER"))
        if rid is None:
            continue
        try:
            key = int(rid) - 1  # legacy: plain row number -> 0-based row index
        except ValueError:
            key = str(rid)  # current: Constituent ID, used as-is
        state, county, tract = clean(r.get("STATE CODE")), clean(r.get("COUNTY CODE")), clean(r.get("TRACT CODE"))
        if not (state and county and tract):
            continue
        out[key] = {
            "state": STATE_FIPS.get(state, state),
            "county": COUNTY_FIPS.get((state, county), f"County {state}{county}"),
            "geoid": f"{state}{county}{tract}",
            "tract": clean_tract(str(tract).zfill(6)),
        }
    return out


ID_COLS = ("Constituent ID", "ConstituentId")


def get_id(row):
    for c in ID_COLS:
        v = row.get(c)
        if v not in (None, ""):
            return clean(v)
    return None


def load_geocode_cache(cache_path):
    if os.path.exists(cache_path):
        with open(cache_path) as f:
            return json.load(f)
    return {"source_file": None, "source_mtime": None, "records": {}}


def save_geocode_cache(cache_path, cache):
    with open(cache_path, "w") as f:
        json.dump(cache, f, separators=(",", ":"))


def geocode_by_id(rows, geocode_paths, cache_path):
    """
    Returns {Constituent ID: tract info} for the Service Map.

    GeocodeResults files carry whatever was in the batch's "Unique ID" column.
    Current convention (tools/build_geocoder_upload.py): the Constituent ID
    itself, so results are keyed by ID with no position matching at all.
    Legacy convention: a plain row number in whichever master_file the batch
    was built from. Row numbers are only meaningful against that exact file,
    so each legacy file is matched by position ONCE, then recorded in the
    cache as consumed and never re-matched against a later master_file.

    Every GeocodeResults file found in data/ is read on every build (oldest
    first, so newer results win), which means ID-keyed results keep working
    even when this runs on a fresh GitHub Actions checkout. Keep those files
    in data/ -- they are the permanent record.
    """
    cache = load_geocode_cache(cache_path)
    consumed = set(cache.get("position_files_consumed", []))
    if cache.get("source_file"):
        consumed.add(cache["source_file"])  # caches written by the earlier version
    for gp in geocode_paths:
        name = os.path.basename(gp)
        merged = 0
        for key, g in read_geocode_tracts(gp).items():
            if isinstance(key, str):
                cache["records"][key] = g
                merged += 1
            elif name not in consumed and 0 <= key < len(rows):
                cid = get_id(rows[key])
                if cid is not None:
                    cache["records"][cid] = g
                    merged += 1
        if merged:
            print(f"Geocode cache  : {name} -> {merged} record(s) applied")
        consumed.add(name)
    cache["position_files_consumed"] = sorted(consumed)
    cache.pop("source_file", None)
    cache.pop("source_mtime", None)
    save_geocode_cache(cache_path, cache)
    return cache["records"]


def build_clients(path, geocode_paths):
    header, rows = read_table(path)
    need(
        header,
        ["Ethnicity", "Race", "County", "UW Income Range", "ALICE Household Type",
         "School Name", "Grade Level"],
        os.path.basename(path),
    )
    total = len(rows)

    demographics = {
        "ethnicity": count_list([multi_label(r["Ethnicity"]) for r in rows]),
        "race": count_list([multi_label(r["Race"]) for r in rows]),
        "income": count_list([plain(r["UW Income Range"]) for r in rows], key=income_key),
        "household": count_list([plain(r["ALICE Household Type"]) for r in rows]),
    }

    # ---- census tract / county / grade (Service Map). Tract comes from the
    # ID-keyed geocode cache (see geocode_by_id) -- no address, name, or
    # coordinate from either spreadsheet is kept past this point.
    cache_path = os.path.join(DATA_DIR, "geocode_cache.json")
    by_id = geocode_by_id(rows, geocode_paths, cache_path)
    if not any(get_id(r) for r in rows):
        print("WARNING: master_file has no Constituent ID column -- Service Map will be empty.")
    tract_info = {}
    tract_grades = defaultdict(Counter)
    for r in rows:
        cid = get_id(r)
        g = by_id.get(cid) if cid else None
        if g is None:
            continue  # never geocoded, or no match -- quietly left off the map
        grade = plain(r["Grade Level"])
        tract_info[g["geoid"]] = g
        tract_grades[g["geoid"]][grade] += 1

    tracts = {
        geoid: {
            "tract": tract_info[geoid]["tract"],
            "county": tract_info[geoid]["county"],
            "state": tract_info[geoid]["state"],
            "total": sum(gr.values()),
            "grades": dict(sorted(gr.items(), key=lambda kv: grade_sort(kv[0]))),
        }
        for geoid, gr in tract_grades.items()
    }

    county_totals = defaultdict(lambda: defaultdict(int))
    for v in tracts.values():
        county_totals[v["state"]][v["county"]] += v["total"]
    counties = {
        st: [c for c, _ in sorted(county_totals[st].items(), key=lambda kv: (-kv[1], kv[0]))]
        for st in county_totals
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

    # ---- schools: most participants first (descending), Unknown last
    schools = count_list(
        [plain(r["School Name"]) for r in rows], key=lambda x: (-x[1], x[0].lower())
    )

    boundaries_path = os.path.join(DATA_DIR, "tract_boundaries.json")
    boundaries = {}
    if os.path.exists(boundaries_path):
        with open(boundaries_path) as f:
            boundaries = json.load(f)
    else:
        print("WARNING: data/tract_boundaries.json is missing -- the Service Map will have no tract outlines.")

    return {
        "demographics": demographics,
        "map": {"tracts": tracts, "counties": counties, "boundaries": boundaries},
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
    g_paths = sorted(
        (f for f in glob.glob(os.path.join(DATA_DIR, "GeocodeResults*.xlsx"))
         if not os.path.basename(f).startswith("~$")),
        key=os.path.getmtime)
    print("Metrics file :", os.path.relpath(q_path, ROOT))
    print("Clients file :", os.path.relpath(f_path, ROOT))
    print("Geocode files:", ", ".join(os.path.basename(g) for g in g_paths) or "(none found -- cached tracts only)")

    overview, youth = build_metrics(q_path)
    demo = build_clients(f_path, g_paths)
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
          f"{len(demo['map']['tracts'])} census tracts with data, "
          f"{len(demo['map']['boundaries'])} tract outlines, data as of {data['asof']})")


if __name__ == "__main__":
    main()
