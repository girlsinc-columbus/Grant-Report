# Girls Inc. Outputs and Outcomes

A single-page report for Girls Inc. of Columbus & Phenix-Russell. The report is **generated from three spreadsheet exports**, so updating the numbers is just a matter of replacing the spreadsheets. The published page never shows names, addresses, or coordinates -- only the aggregated counts below.

Tabs: Clients Served, Youth Opportunity, Demographics, Service Map.

Colors: Girls Inc Red `#ED1849`, Girls Inc Grey `#949CA1`, Black `#000000`, White `#FFFFFF`.

## The three spreadsheets

All three live in the `data/` folder:

| File (name starts with) | What it feeds |
|---|---|
| `master_query` | Clients Served and Youth Opportunity results (Section, MetricCode, Metric, Period, Value, Total Served, Successful, Success Percent) |
| `master_file` | Client-level records for Demographics and Service Map (Constituent ID, Ethnicity, Race, County, UW Income Range, ALICE Household Type, School Name, Grade Level, plus the address columns used only to build the geocoder upload -- see below) |
| `GeocodeResults` | Census tract for new/ungeocoded `master_file` records, from the Census Bureau's batch geocoder (Service Map only -- see "The map's tract outlines" and "data/geocode_cache.json" below for why you usually don't need a new one of these) |

You can keep the original export names, but **the file must start with `master_file`** -- if your reporting tool names the download something else (e.g. `query_result_...xlsx`), rename it on upload so `build.py` can find it; if it isn't renamed, the build fails with a clear error rather than silently using the wrong file. If several files match a pattern, the **newest by file** is used, so new exports automatically win over old ones. Column headings must stay the same as in the current exports.

### Getting the GeocodeResults file

The Service Map is built on census tracts, not zip codes, so each record needs its tract. That comes from the Census Bureau's free batch geocoder -- there's no reliable way to go from a zip code alone to a tract. You only need to do this for clients who don't already have a tract cached (new clients, mainly) -- not the whole roster every time.

1. After uploading the latest `master_file`, run:
   ```bash
   pip install openpyxl
   python3 tools/build_geocoder_upload.py
   ```
   This writes `geocoder_upload.csv` in the repo root, containing only the clients not already in `data/geocode_cache.json`, with each client's Constituent ID as the Unique ID.
2. Go to **[geocoding.geo.census.gov/geocoder](https://geocoding.geo.census.gov/geocoder)** → **Address Batch** → **Find Matching Geographies**, and upload `geocoder_upload.csv` as-is. Choose the **Current** benchmark and leave the vintage as default.
3. Download the results, open and save them as an Excel file (`.xlsx`), and upload that to `data/` with a name starting `GeocodeResults` (for example `GeocodeResults_2026-10-16.xlsx`).
4. Every build reads **all** `GeocodeResults*.xlsx` files in `data/` and applies them by Constituent ID, so **keep each one in `data/`** -- they are the permanent record (the GitHub build starts fresh every time and can't remember anything on its own).

If no `GeocodeResults` or `geocode_cache.json` file is present at all, the build still runs -- Clients Served, Youth Opportunity, and Demographics are unaffected, and the Service Map is simply empty.

## Updating the report

**On GitHub (no software needed)**
1. Open the `data` folder in the repository.
2. Choose **Add file → Upload files** and drop in the new `master_query`, `master_file`, and `GeocodeResults` exports.
3. Optionally delete the older exports, then **Commit changes**.
4. The **Build and Deploy** action runs (about a minute) and the published site updates. Check progress under the **Actions** tab.

**On your computer**
```bash
pip install openpyxl
python build.py          # writes index.html; open it in a browser
```

## One-time GitHub Pages setup
In **Settings → Pages**, set **Source** to **GitHub Actions**.

## What the build does automatically
- **Data as of** date is taken from the timestamp in the `master_query` file name.
- Client Movement: Q1 New Clients equals Q1 Clients Served. Exits for quarters that have not started yet display as 0.
- Blank values are shown as **Unknown** and counted.
- Race entries with repeated or empty values (for example `BlackAfricanAmerican|`) are cleaned up; multi-race entries show as "A & B".
- Income Range is ordered lowest to highest, Unknown last. Schools Served is ordered most to fewest participants, Unknown last.
- Service Map: each record's census tract, county, and state come from `GeocodeResults`. The very first time a given `GeocodeResults` file is seen, it's matched to `master_file` by row position (that's the only information the geocoder gives back) and immediately saved into `data/geocode_cache.json`, keyed by each client's **Constituent ID** instead. Every build after that looks tracts up by ID, so uploading a new `master_file` -- with rows added, removed, or reordered -- can no longer shift anyone's tract. A client only needs a fresh `GeocodeResults` run if they're new and were never geocoded before; `data/geocode_cache.json` should not be deleted, since it's what carries everyone else's tract forward. Tract codes have their leading zeros stripped and the decimal suffix restored (e.g. `010106` -> tract `101.06`). New tracts and counties appear automatically as they show up in the data. Records that never matched an address, or whose Constituent ID has no cached tract yet, are simply left off the map, with no count or note displayed.
- Grade groups: Elementary = GK-G5, Middle = G6-G8, High School = G9-G12, College = PS1, PS2, etc.

## The map's tract outlines
The Service Map draws each county's real census tract shapes (not boxes), shaded darker for more participants served and grey for tracts with none on file. The outlines come from the Census Bureau's TIGER/Line tract shapefiles, which are a one-time addition to this repo:

- `tools/tiger/tl_2025_13_tract.zip` (Georgia) and `tl_2025_01_tract.zip` (Alabama) -- the original Census downloads.
- `tools/extract_tracts.py` -- a small pure-Python script (no GIS library required) that reads those shapefiles, keeps only Muscogee, Harris, Russell, and Lee counties, simplifies the shapes, and writes `data/tract_boundaries.json`. This file is what the report actually embeds; `build.py` reads it automatically if present.

You don't need to run this again for a normal spreadsheet update. Re-run it only if:
- a new county starts appearing in `GeocodeResults` -- add it to the `want` dict near the top of `tools/extract_tracts.py` and to `COUNTY_FIPS` in `build.py`, or
- you want to refresh to a newer TIGER/Line release (tract boundaries rarely change -- mainly at the once-a-decade redistricting, plus the occasional small correction). Download the new state files from the Census Bureau's TIGER page, replace the zips in `tools/tiger/`, and run:
  ```bash
  python3 tools/extract_tracts.py
  ```

## Keeping client data off the published page
`build.py` only ever writes aggregated counts into `index.html` -- by tract, county, grade, school, and so on. It never writes names, addresses, or coordinates from `master_file` or `GeocodeResults`; those are read only in memory to compute the counts above. Tract boundary shapes are public Census geography, not client data.

## Files
- `build.py`: reads the spreadsheets and writes the report
- `template.html`: the page design and layout (edit this to change the look)
- `config/youth_labels.json`: display names and wording for each Youth Opportunity measure (add an entry when a new MetricCode is introduced; unknown codes still show up with generic labels)
- `index.html`: the built report (regenerated by `build.py`)
- `data/geocode_cache.json`: Constituent ID -> census tract, built up automatically (see Service Map above). Holds the older row-number-based batch; don't delete it.
- `tools/build_geocoder_upload.py`: generates the CSV to upload to the Census batch geocoder for clients not yet in the cache (see "Getting the GeocodeResults file" above)
- `.github/workflows/pages.yml`: rebuilds and publishes on every push
