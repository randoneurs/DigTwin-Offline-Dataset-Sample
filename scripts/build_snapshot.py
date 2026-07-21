#!/usr/bin/env python3
"""Convert META market-share CSV exports into the JSON snapshot index.html loads
via "Load snapshot (.json)". Run this whenever new all_city_operator_*.csv /
region_*.csv files land, then drop the resulting file into the dashboard.

Usage:
  python3 scripts/build_snapshot.py \
      --date 20260628 --all-city all_city_operator_20260628.csv --region region_20260628.csv \
      --wow-date 20260621 --wow-all-city all_city_operator_20260621.csv --wow-region region_20260621.csv \
      --prev-date 20260524 --prev-all-city all_city_operator_20260524.csv --prev-region region_20260524.csv \
      --yoy-date 20250629 --yoy-all-city all_city_operator_20250629.csv --yoy-region region_20250629.csv \
      --out snapshot_20260628.json

--wow-*, --prev-*, and --yoy-* are each optional; when given, real week-over-
week (WoW, exactly 7 days back) / month-over-month (MoM) / year-over-year
(YoY) deltas and a chronological sparkline are computed instead of being left
null. Every operator's own share (not just tsel's) is carried forward as
shareByOperatorWow/Prev/Yoy so the dashboard can show each competitor's own
change, not just Telkomsel's.
"""
import argparse
import csv
import json
import re
import sys
from datetime import date

AREA_MAP = {
    "Area 1": ["SUMBAGUT", "SUMBAGTENG", "SUMBAGSEL"],
    "Area 2": ["JAKARTA-BANTEN", "EASTERN JABOTABEK", "JABAR"],
    "Area 3": ["JATENG-DIY", "JATIM", "BALINUSRA"],
    "Area 4": ["KALIMANTAN", "SULAWESI", "PUMA"],
}
REGION_TO_AREA = {region: area for area, regions in AREA_MAP.items() for region in regions}

# region_*.csv's "Region" column (any case) -> this dashboard's REGION_LABELS key.
REGION_CSV_TO_KEY = {
    "sumbagut": "SUMBAGUT",
    "sumbagteng": "SUMBAGTENG",
    "sumbagsel": "SUMBAGSEL",
    "jakarta banten": "JAKARTA-BANTEN",
    "eastern jabotabek": "EASTERN JABOTABEK",
    "jabar": "JABAR",
    "jateng": "JATENG-DIY",
    "jatim": "JATIM",
    "balinusra": "BALINUSRA",
    "kalimantan": "KALIMANTAN",
    "sulawesi": "SULAWESI",
    "puma": "PUMA",
}

# The dashboard's curated CITY_MAP (index.html) — only these cities resolve
# through the fArea/fRegion/fCity filters, so only these are pulled from the
# 510-city CSV. Everything else in the CSV is real data but isn't reachable
# from a dropdown yet.
CITY_MAP = {
    "SUMBAGUT": ["Kota Medan", "Kota Banda Aceh", "Kabupaten Deli Serdang", "Kota Pematangsiantar", "Kota Binjai", "Kabupaten Langkat", "Kota Lhokseumawe", "Kabupaten Simalungun"],
    "SUMBAGTENG": ["Kota Pekanbaru", "Kota Padang", "Kota Batam", "Kota Jambi", "Kota Dumai", "Kabupaten Kampar", "Kota Tanjungpinang", "Kabupaten Bungo"],
    "SUMBAGSEL": ["Kota Palembang", "Kota Bandar Lampung", "Kota Bengkulu", "Kota Pangkalpinang", "Kabupaten Ogan Ilir", "Kota Prabumulih", "Kabupaten Lampung Selatan", "Kota Metro"],
    "JAKARTA-BANTEN": ["Kota Jakarta Selatan", "Kota Jakarta Pusat", "Kota Tangerang", "Kota Tangerang Selatan", "Kota Serang", "Kabupaten Tangerang", "Kota Cilegon", "Kota Jakarta Barat"],
    "EASTERN JABOTABEK": ["Kota Bekasi", "Kabupaten Bekasi", "Kota Depok", "Kota Bogor", "Kabupaten Bogor", "Kota Jakarta Timur", "Kabupaten Karawang", "Kota Jakarta Utara"],
    "JABAR": ["Kota Bandung", "Kota Cirebon", "Kabupaten Bandung", "Kota Sukabumi", "Kota Tasikmalaya", "Kabupaten Garut", "Kota Cimahi", "Kabupaten Sumedang"],
    "JATENG-DIY": ["Kota Semarang", "Kota Yogyakarta", "Kota Surakarta", "Kabupaten Sleman", "Kota Magelang", "Kota Tegal", "Kabupaten Banyumas", "Kota Pekalongan"],
    "JATIM": ["Kota Surabaya", "Kota Malang", "Kota Kediri", "Kabupaten Sidoarjo", "Kota Madiun", "Kabupaten Jember", "Kota Mojokerto", "Kabupaten Banyuwangi"],
    "BALINUSRA": ["Kota Denpasar", "Kabupaten Badung", "Kota Mataram", "Kota Kupang", "Kabupaten Sumbawa", "Kabupaten Buleleng", "Kabupaten Lombok Barat", "Kabupaten Ende"],
    "KALIMANTAN": ["Kota Balikpapan", "Kota Samarinda", "Kota Banjarmasin", "Kota Pontianak", "Kota Palangka Raya", "Kabupaten Kutai Kartanegara", "Kota Tarakan", "Kota Banjarbaru"],
    "SULAWESI": ["Kota Makassar", "Kota Manado", "Kota Palu", "Kota Kendari", "Kota Gorontalo", "Kabupaten Bone", "Kota Pare-Pare", "Kabupaten Minahasa"],
    "PUMA": ["Kota Jayapura", "Kota Ambon", "Kabupaten Merauke", "Kota Sorong", "Kabupaten Mimika", "Kota Ternate", "Kabupaten Jayawijaya", "Kota Manokwari"],
}


# Known spots where the CSV's spelling doesn't follow the generic "Kota X" ->
# "KOTA X" / "Kabupaten X" -> "X" rule (checked against a real export).
CITY_NAME_OVERRIDES = {
    "Kota Pematangsiantar": "KOTA PEMATANG SIANTAR",
    "Kota Tanjungpinang": "KOTA TANJUNG PINANG",
    "Kota Jakarta Selatan": "JAKARTA SELATAN",
    "Kota Jakarta Pusat": "JAKARTA PUSAT",
    "Kota Jakarta Barat": "JAKARTA BARAT",
    "Kota Jakarta Timur": "JAKARTA TIMUR",
    "Kota Jakarta Utara": "JAKARTA UTARA",
    "Kota Palangka Raya": "KOTA PALANGKARAYA",
    "Kota Banjarbaru": "KOTA BANJAR BARU",
    "Kota Manokwari": "MANOKWARI",
}


def city_csv_key(dashboard_name):
    if dashboard_name in CITY_NAME_OVERRIDES:
        return CITY_NAME_OVERRIDES[dashboard_name]
    if dashboard_name.startswith("Kota "):
        return "KOTA " + dashboard_name[len("Kota "):].upper()
    if dashboard_name.startswith("Kabupaten "):
        return dashboard_name[len("Kabupaten "):].upper()
    return dashboard_name.upper()


CITY_CSV_TO_DASHBOARD = {
    city_csv_key(city): (region, city)
    for region, cities in CITY_MAP.items()
    for city in cities
}

OPERATOR_FIELDS = ["tsel", "xl", "isat", "three", "smartfren", "xl+", "ioh"]


def to_float(value):
    value = (value or "").strip()
    return float(value) if value else None


def share_by_operator(row, keymap):
    return {
        "tsel": to_float(row[keymap["tsel"]]),
        "xlLegacy": to_float(row[keymap["xl"]]),
        "isat": to_float(row[keymap["isat"]]),
        "three": to_float(row[keymap["three"]]),
        "smartfren": to_float(row[keymap["smartfren"]]),
        # XL+ = XL + Smartfren merged brand ("XLSmart") — this is the dashboard's "xl" operator identity.
        "xl": to_float(row[keymap["xl+"]]),
        # IOH = Indosat + Tri merged brand — the dashboard's "ioh" operator identity.
        "ioh": to_float(row[keymap["ioh"]]),
    }


def header_keymap(fieldnames):
    lower = {h.lower(): h for h in fieldnames}
    return {op: lower[op] for op in OPERATOR_FIELDS}


def iso_week_label(d):
    iso_year, iso_week, _ = d.isocalendar()
    return f"W{iso_week:02d}-{iso_year}"


def parse_yyyymmdd(s):
    return date(int(s[0:4]), int(s[4:6]), int(s[6:8]))


def read_all_city(path):
    """Returns dict: (region_key, dashboard_city_name) -> shareByOperator, plus count of unmatched CSV rows."""
    out = {}
    skipped = 0
    total = 0
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        keymap = header_keymap(reader.fieldnames)
        for r in reader:
            total += 1
            match = CITY_CSV_TO_DASHBOARD.get(r["CITY"].strip().upper())
            if not match:
                skipped += 1
                continue
            out[match] = share_by_operator(r, keymap)
    return out, skipped, total


def read_region(path):
    """Returns list of (level, key_tuple, shareByOperator) where key_tuple is
    (area, region, city, island, archetype) with 'ALL' for the axes that don't apply."""
    out = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        keymap = header_keymap(reader.fieldnames)
        for r in reader:
            level = r["level"].strip()
            name = r["Region"].strip()
            share = share_by_operator(r, keymap)
            if level == "Nationwide":
                out.append((level, ("ALL", "ALL", "ALL", None, None), share))
            elif level == "Area Group":
                out.append((level, ("ALL", "ALL", "ALL", name, None), share))
            elif level == "Area":
                out.append((level, (name, "ALL", "ALL", None, None), share))
            elif level == "Region":
                region_key = REGION_CSV_TO_KEY.get(name.lower())
                if not region_key:
                    print(f"warning: unmapped region '{name}', skipping", file=sys.stderr)
                    continue
                out.append((level, (REGION_TO_AREA[region_key], region_key, "ALL", None, None), share))
            elif level == "Archetype":
                out.append((level, ("ALL", "ALL", "ALL", None, name), share))
            elif level == "Archetype Area":
                m = re.match(r"^(.*) AREA (\d)$", name, re.IGNORECASE)
                if not m:
                    print(f"warning: unrecognized Archetype Area '{name}', skipping", file=sys.stderr)
                    continue
                out.append((level, (f"Area {m.group(2)}", "ALL", "ALL", None, m.group(1).strip()), share))
            else:
                print(f"warning: unknown level '{level}' for '{name}', skipping", file=sys.stderr)
    return out


def build_share_lookup(all_city_path, region_path):
    """Full per-operator share dict for every (level, area, region, city, island, archetype)
    key in one period's CSVs — lets callers pull tsel's own share (for the MoM/YoY delta) or
    a competitor's (xl/ioh) share, all from the same lookup."""
    lookup = {}
    all_city, _, _ = read_all_city(all_city_path)
    for (region_key, city_name), share in all_city.items():
        lookup[("City", REGION_TO_AREA[region_key], region_key, city_name, None, None)] = share
    for level, (area, region, city, island, archetype), share in read_region(region_path):
        lookup[(level, area, region, city, island, archetype)] = share
    return lookup


def build_history(yoy_val, prev_val, wow_val, current_val):
    """Chronological [yoy, prev(MoM), wow, current], dropping whichever are missing."""
    pts = [v for v in (yoy_val, prev_val, wow_val) if v is not None]
    pts.append(current_val)
    return pts if len(pts) >= 2 else None


def build(args):
    d = parse_yyyymmdd(args.date)
    week = iso_week_label(d)

    wow_share_by_key = build_share_lookup(args.wow_all_city, args.wow_region) if args.wow_date else {}
    prev_share_by_key = build_share_lookup(args.prev_all_city, args.prev_region) if args.prev_date else {}
    yoy_share_by_key = build_share_lookup(args.yoy_all_city, args.yoy_region) if args.yoy_date else {}

    def add_comparisons(row, key, share):
        wow_share = wow_share_by_key.get(key)
        prev_share, yoy_share = prev_share_by_key.get(key), yoy_share_by_key.get(key)
        wow = (wow_share or {}).get("tsel")
        prev, yoy = (prev_share or {}).get("tsel"), (yoy_share or {}).get("tsel")
        if wow is not None:
            row["metaMarketSharePrevWowPct"] = wow
        if prev is not None:
            row["metaMarketSharePrevPct"] = prev
        if yoy is not None:
            row["metaMarketSharePrevYoyPct"] = yoy
        if wow_share is not None:
            row["shareByOperatorWow"] = wow_share
        if prev_share is not None:
            row["shareByOperatorPrev"] = prev_share
        if yoy_share is not None:
            row["shareByOperatorYoy"] = yoy_share
        history = build_history(yoy, prev, wow, share["tsel"])
        if history:
            row["metaShareHistory"] = history

    rows = []

    all_city, skipped, total = read_all_city(args.all_city)
    for (region_key, city_name), share in all_city.items():
        area_key = REGION_TO_AREA[region_key]
        row = {
            "level": "City", "area": area_key, "region": region_key, "city": city_name, "week": week,
            "metaMarketSharePct": share["tsel"], "shareByOperator": share,
        }
        add_comparisons(row, ("City", area_key, region_key, city_name, None, None), share)
        rows.append(row)
    if skipped:
        print(f"note: {skipped} of {total} city rows in {args.all_city} aren't in this dashboard's "
              f"curated CITY_MAP and were skipped — not an error, just not reachable from a filter dropdown yet.",
              file=sys.stderr)

    for level, (area, region, city, island, archetype), share in read_region(args.region):
        row = {
            "level": level, "area": area, "region": region, "city": city, "week": week,
            "metaMarketSharePct": share["tsel"], "shareByOperator": share,
        }
        if island:
            row["island"] = island
        if archetype:
            row["archetype"] = archetype
        add_comparisons(row, (level, area, region, city, island, archetype), share)
        rows.append(row)

    snapshot = {
        "generatedAt": d.isoformat() + "T00:00:00+07:00",
        "note": "Generated from META market-share exports via scripts/build_snapshot.py — do not hand-edit.",
        "rows": rows,
    }
    if args.wow_date:
        snapshot["comparedToDateWow"] = parse_yyyymmdd(args.wow_date).isoformat() + "T00:00:00+07:00"
    if args.prev_date:
        snapshot["comparedToDate"] = parse_yyyymmdd(args.prev_date).isoformat() + "T00:00:00+07:00"
    if args.yoy_date:
        snapshot["comparedToDateYoy"] = parse_yyyymmdd(args.yoy_date).isoformat() + "T00:00:00+07:00"
    return snapshot


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--date", required=True, help="YYYYMMDD, matches the current CSVs' filename suffix")
    ap.add_argument("--all-city", required=True, help="path to all_city_operator_YYYYMMDD.csv")
    ap.add_argument("--region", required=True, help="path to region_YYYYMMDD.csv")
    ap.add_argument("--wow-date", help="YYYYMMDD exactly 7 days prior, to compute real WoW deltas")
    ap.add_argument("--wow-all-city", help="path to the prior week's all_city_operator CSV")
    ap.add_argument("--wow-region", help="path to the prior week's region CSV")
    ap.add_argument("--prev-date", help="YYYYMMDD for the prior period, to compute real MoM deltas")
    ap.add_argument("--prev-all-city", help="path to the prior period's all_city_operator CSV")
    ap.add_argument("--prev-region", help="path to the prior period's region CSV")
    ap.add_argument("--yoy-date", help="YYYYMMDD for the same period one year prior, to compute real YoY deltas")
    ap.add_argument("--yoy-all-city", help="path to the year-ago all_city_operator CSV")
    ap.add_argument("--yoy-region", help="path to the year-ago region CSV")
    ap.add_argument("--out", required=True, help="output snapshot JSON path (dated archive copy)")
    ap.add_argument("--latest-copy", default="snapshot_latest.json",
                    help="also write here — this is the stable filename index.html auto-fetches on load "
                         "(default: snapshot_latest.json; pass '' to skip)")
    args = ap.parse_args()

    if args.wow_date and not (args.wow_all_city and args.wow_region):
        ap.error("--wow-date requires --wow-all-city and --wow-region")
    if args.prev_date and not (args.prev_all_city and args.prev_region):
        ap.error("--prev-date requires --prev-all-city and --prev-region")
    if args.yoy_date and not (args.yoy_all_city and args.yoy_region):
        ap.error("--yoy-date requires --yoy-all-city and --yoy-region")

    snapshot = build(args)
    with open(args.out, "w") as f:
        json.dump(snapshot, f, indent=2)
    print(f"wrote {args.out}: {len(snapshot['rows'])} rows")

    if args.latest_copy:
        with open(args.latest_copy, "w") as f:
            json.dump(snapshot, f, indent=2)
        print(f"wrote {args.latest_copy} (auto-fetched by index.html)")


if __name__ == "__main__":
    main()
