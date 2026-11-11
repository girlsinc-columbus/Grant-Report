#!/usr/bin/env python3
"""
Builds the CSV to upload to the Census Bureau's batch geocoder
(geocoding.geo.census.gov/geocoder -> Address Batch -> Find Matching
Geographies), so the Service Map can place new clients by census tract.

Only clients NOT already in data/geocode_cache.json are included -- there's
no need to re-geocode anyone whose tract is already known. Run this after
uploading a new master_file, whenever new clients have shown up.

The "Unique ID" column is each client's Constituent ID (not a row number).
build.py reads it back that way and caches tracts by ID directly, so there's
no row-position matching involved at all, now or on any future master_file
update. (A GeocodeResults file already using the old row-number convention
still works too -- see read_geocode_tracts in build.py.)

Usage:
    pip install openpyxl
    python3 tools/build_geocoder_upload.py
    # -> writes geocoder_upload.csv in the repo root

Then: upload geocoder_upload.csv at the link above, "Current" benchmark,
default vintage, download the results, and drop that file into data/ as
GeocodeResults....xlsx (or .csv).
"""
import csv
import glob
import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from build import DATA_DIR, ROOT, clean, get_id, load_geocode_cache, newest, read_table  # noqa: E402


def main():
    f_path = newest("master_file*.xlsx")
    cache = load_geocode_cache(os.path.join(DATA_DIR, "geocode_cache.json"))
    known_ids = set(cache["records"].keys())

    header, rows = read_table(f_path)
    need_cols = ["AddressLine1", "City", "State", "ZipCode"]
    missing = [c for c in need_cols if c not in header]
    if missing:
        sys.exit(f"ERROR: {os.path.basename(f_path)} is missing column(s): {', '.join(missing)}")

    out_rows = []
    skipped_no_id, skipped_no_address = 0, 0
    for r in rows:
        cid = get_id(r)
        if cid is None:
            skipped_no_id += 1
            continue
        if cid in known_ids:
            continue  # already geocoded -- no need to resubmit
        street = clean(r.get("AddressLine1"))
        if not street:
            skipped_no_address += 1
            continue
        city, state, zipc = clean(r.get("City")) or "", clean(r.get("State")) or "", clean(r.get("ZipCode")) or ""
        out_rows.append([cid, street, city, state, zipc])

    out_path = os.path.join(ROOT, "geocoder_upload.csv")
    with open(out_path, "w", newline="") as f:
        w = csv.writer(f)
        w.writerows(out_rows)  # no header row -- the Census batch geocoder expects none

    print(f"{os.path.basename(f_path)}: {len(rows)} client(s) total")
    print(f"Already geocoded (cached)  : {len(rows) - len(out_rows) - skipped_no_id - skipped_no_address}")
    if skipped_no_id:
        print(f"Skipped, no Constituent ID: {skipped_no_id}")
    if skipped_no_address:
        print(f"Skipped, no address on file: {skipped_no_address}")
    print(f"Wrote {len(out_rows)} record(s) needing geocoding to {os.path.relpath(out_path, ROOT)}")


if __name__ == "__main__":
    main()
