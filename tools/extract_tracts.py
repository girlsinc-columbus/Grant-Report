#!/usr/bin/env python3
"""
One-time / maintenance step for the Service Map's real tract outlines.

Reads the state-wide TIGER/Line Census Tract shapefiles in tools/tiger/
(tl_2025_01_tract.zip = Alabama, tl_2025_13_tract.zip = Georgia -- download
current ones from https://www2.census.gov/geo/tiger/TIGER2025/TRACT/ if you
need to refresh these), keeps only the tracts in the counties listed in
`want` below, simplifies the polygons, and writes data/tract_boundaries.json.

No third-party geo library is used (none were installable in the sandbox
this was built in) -- just a small pure-Python .shp/.dbf reader.

Run this again only when:
  - a new county starts showing up in GeocodeResults (add it to `want`
    below, and to COUNTY_FIPS in build.py), or
  - the Census Bureau updates tract boundaries (redistricting; roughly
    once per decade, plus the occasional small correction).

    python3 tools/extract_tracts.py

Does NOT run as part of the normal spreadsheet -> report build.
"""
import struct
import json
import sys

def read_dbf(path):
    with open(path, 'rb') as f:
        data = f.read()
    num_records = struct.unpack('<I', data[4:8])[0]
    header_size = struct.unpack('<H', data[8:10])[0]
    record_size = struct.unpack('<H', data[10:12])[0]
    fields = []
    pos = 32
    while data[pos] != 0x0D:
        name = data[pos:pos+11].split(b'\x00')[0].decode('ascii')
        typ = chr(data[pos+11])
        length = data[pos+16]
        fields.append((name, typ, length))
        pos += 32
    records = []
    rec_start = header_size
    for i in range(num_records):
        off = rec_start + i * record_size
        rec = data[off:off+record_size]
        if rec[0:1] == b'*':
            records.append(None)
            continue
        fpos = 1
        row = {}
        for name, typ, length in fields:
            raw = rec[fpos:fpos+length].decode('ascii', errors='replace').strip()
            row[name] = raw
            fpos += length
        records.append(row)
    return records


def read_shp(path):
    with open(path, 'rb') as f:
        data = f.read()
    shapes = []
    pos = 100
    n = len(data)
    while pos < n:
        rec_num = struct.unpack('>I', data[pos:pos+4])[0]
        content_len_words = struct.unpack('>I', data[pos+4:pos+8])[0]
        content_len = content_len_words * 2
        rpos = pos + 8
        shape_type = struct.unpack('<I', data[rpos:rpos+4])[0]
        if shape_type == 0:
            shapes.append(None)
        else:
            # Polygon (5) or PolygonZ (15): box(32) + numParts(4) + numPoints(4) + parts + points
            o = rpos + 4 + 32
            num_parts = struct.unpack('<i', data[o:o+4])[0]
            num_points = struct.unpack('<i', data[o+4:o+8])[0]
            o += 8
            parts = struct.unpack('<%di' % num_parts, data[o:o+4*num_parts])
            o += 4 * num_parts
            pts = []
            for _ in range(num_points):
                x, y = struct.unpack('<dd', data[o:o+16])
                pts.append((x, y))
                o += 16
            rings = []
            for i, start in enumerate(parts):
                end = parts[i+1] if i+1 < num_parts else num_points
                rings.append(pts[start:end])
            shapes.append(rings)
        pos += 8 + content_len
    return shapes


def perp_dist(pt, a, b):
    (x, y), (ax, ay), (bx, by) = pt, a, b
    dx, dy = bx - ax, by - ay
    if dx == 0 and dy == 0:
        return ((x-ax)**2 + (y-ay)**2) ** 0.5
    t = ((x-ax)*dx + (y-ay)*dy) / (dx*dx + dy*dy)
    t = max(0, min(1, t))
    px, py = ax + t*dx, ay + t*dy
    return ((x-px)**2 + (y-py)**2) ** 0.5


def rdp(points, eps):
    if len(points) < 3:
        return points
    dmax, idx = 0, 0
    for i in range(1, len(points)-1):
        d = perp_dist(points[i], points[0], points[-1])
        if d > dmax:
            dmax, idx = d, i
    if dmax > eps:
        left = rdp(points[:idx+1], eps)
        right = rdp(points[idx:], eps)
        return left[:-1] + right
    return [points[0], points[-1]]


def simplify_ring(ring, eps):
    if len(ring) <= 4:
        return ring
    out = rdp(ring, eps)
    if len(out) < 4:
        return ring[:1] + ring[len(ring)//2:len(ring)//2+1] + ring[-1:]
    return out


import os
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TIGER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'tiger')
import zipfile, tempfile

def unzip_once(zip_name, out_dir):
    if os.path.isdir(out_dir):
        return
    with zipfile.ZipFile(os.path.join(TIGER, zip_name)) as z:
        z.extractall(out_dir)

def clean_tract(tract6):
    """'001000' -> '10', '003301' -> '33.01', '010106' -> '101.06' (matches build.py)."""
    base, suffix = tract6[:4], tract6[4:]
    label = str(int(base))
    if suffix and suffix != "00":
        label += "." + suffix
    return label


def main():
    unzip_once('tl_2025_13_tract.zip', os.path.join(TIGER, 'ga'))
    unzip_once('tl_2025_01_tract.zip', os.path.join(TIGER, 'al'))
    want = {
        ('13', '215'): 'Muscogee County',
        ('13', '145'): 'Harris County',
        ('01', '113'): 'Russell County',
        ('01', '081'): 'Lee County',
    }
    out = {}
    counts = {}
    for state_fips, shp_dir, shp_name in [
        ('13', 'ga', 'tl_2025_13_tract'),
        ('01', 'al', 'tl_2025_01_tract'),
    ]:
        recs = read_dbf(os.path.join(TIGER, shp_dir, shp_name + '.dbf'))
        shapes = read_shp(os.path.join(TIGER, shp_dir, shp_name + '.shp'))
        assert len(recs) == len(shapes), (len(recs), len(shapes))
        for rec, rings in zip(recs, shapes):
            if rec is None or rings is None:
                continue
            key = (rec['STATEFP'], rec['COUNTYFP'])
            if key not in want:
                continue
            geoid = rec['GEOID']
            simplified = [simplify_ring(ring, 0.0003) for ring in rings]
            out[geoid] = {
                'state': 'GA' if state_fips == '13' else 'AL',
                'county': want[key],
                'tract': clean_tract(rec['TRACTCE']),
                'rings': [[[round(x, 5), round(y, 5)] for x, y in ring] for ring in simplified],
            }
            counts[key] = counts.get(key, 0) + 1

    for k, v in counts.items():
        print(k, want[k], v, 'tracts')
    total_pts = sum(len(r) for t in out.values() for r in t['rings'])
    print('total tracts', len(out), 'total points after simplify', total_pts)

    out_path = os.path.join(ROOT, 'data', 'tract_boundaries.json')
    with open(out_path, 'w') as f:
        json.dump(out, f, separators=(',', ':'))
    print('wrote', os.path.relpath(out_path, ROOT))


if __name__ == '__main__':
    main()
