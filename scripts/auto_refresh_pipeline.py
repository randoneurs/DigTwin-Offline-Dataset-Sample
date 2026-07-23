#!/usr/bin/env python3
"""Check the Dataset branch for source files newer than what's currently
published, and if found, run the matching converter (build_snapshot.py /
build_retina_snapshot.py / build_network_snapshot.py) with the right
--wow-date/--prev-date(MoM)/--yoy-date flags — chosen by comparing the new
file's date against every other existing file of the same type and picking
whichever falls in each gap window:
  WoW  ~7 days   (5-9 day gap)
  MoM  ~1 month  (25-31 day gap)
  YoY  ~1 year   (358-372 day gap)
If more than one existing file falls in a window, the one closest to the
window's midpoint is used.

Usage (from a repo checkout that already has scripts/, *_latest.json, etc.
on the current branch — this script only reads the Dataset branch remotely,
it never checks it out):
  python3 scripts/auto_refresh_pipeline.py [--dataset-branch Dataset] [--dry-run]

Exits 0 whether or not anything changed; prints a summary either way. Does
NOT commit or push — that's left to the caller (see the scheduled routine's
job prompt) so this stays a pure "did anything change" step that's easy to
test standalone.
"""
import argparse
import json
import re
import subprocess
import sys
import tempfile
from datetime import datetime
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

WOW_WINDOW = (5, 9)
MOM_WINDOW = (25, 31)
YOY_WINDOW = (358, 372)


def run_git(args, cwd=None, check=True):
    result = subprocess.run(["git", *args], cwd=cwd or REPO_ROOT, capture_output=True, text=True)
    if check and result.returncode != 0:
        raise RuntimeError(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout


def list_dataset_files(dataset_branch):
    run_git(["fetch", "origin", dataset_branch, "--quiet"])
    out = run_git(["ls-tree", "-r", "--name-only", f"origin/{dataset_branch}"])
    return [line.strip() for line in out.splitlines() if line.strip()]


def extract_file(dataset_branch, path, dest):
    result = subprocess.run(["git", "show", f"origin/{dataset_branch}:{path}"], cwd=REPO_ROOT, capture_output=True)
    if result.returncode != 0:
        raise RuntimeError(f"git show origin/{dataset_branch}:{path} failed: {result.stderr.decode().strip()}")
    dest.write_bytes(result.stdout)


def parse_yyyymmdd(s):
    return datetime.strptime(s, "%Y%m%d").date()


def closest_in_window(target_date, candidates, window):
    """candidates: [(date, ...key info...)]. Returns the candidate whose gap
    to target_date falls in `window` (days) and is closest to its midpoint,
    or None."""
    lo, hi = window
    mid = (lo + hi) / 2
    best, best_diff = None, None
    for c in candidates:
        gap = abs((target_date - c[0]).days)
        if lo <= gap <= hi:
            diff = abs(gap - mid)
            if best is None or diff < best_diff:
                best, best_diff = c, diff
    return best


def read_json_date(path, field, date_len=10):
    if not path.exists():
        return None
    try:
        data = json.loads(path.read_text())
    except (json.JSONDecodeError, OSError):
        return None
    value = data.get(field)
    if not value:
        return None
    try:
        return datetime.strptime(value[:date_len], "%Y-%m-%d").date()
    except ValueError:
        return None


# ==================== META Market Share ====================
META_ALL_CITY_RE = re.compile(r"^all_city_operator_(\d{8})\.csv$")
META_REGION_RE = re.compile(r"^region_(\d{8})\.csv$")


def refresh_meta(dataset_branch, files, dry_run):
    all_city, region = {}, {}
    for f in files:
        m = META_ALL_CITY_RE.match(f)
        if m:
            all_city[parse_yyyymmdd(m.group(1))] = f
        m = META_REGION_RE.match(f)
        if m:
            region[parse_yyyymmdd(m.group(1))] = f
    dates = sorted(set(all_city) & set(region))
    if not dates:
        return {"pipeline": "meta", "changed": False, "reason": "no all_city_operator/region pairs found"}

    current = read_json_date(REPO_ROOT / "snapshot_latest.json", "generatedAt")
    newest = dates[-1]
    if current is not None and newest <= current:
        return {"pipeline": "meta", "changed": False, "reason": f"newest available ({newest}) not newer than published ({current})"}

    others = [(d,) for d in dates if d != newest]
    wow = closest_in_window(newest, others, WOW_WINDOW)
    mom = closest_in_window(newest, others, MOM_WINDOW)
    yoy = closest_in_window(newest, others, YOY_WINDOW)

    if dry_run:
        return {"pipeline": "meta", "changed": True, "dry_run": True, "newest": str(newest),
                "wow": str(wow[0]) if wow else None, "mom": str(mom[0]) if mom else None, "yoy": str(yoy[0]) if yoy else None}

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        needed = [newest] + [c[0] for c in (wow, mom, yoy) if c]
        for d in needed:
            extract_file(dataset_branch, all_city[d], tmp / all_city[d])
            extract_file(dataset_branch, region[d], tmp / region[d])

        date_str = newest.strftime("%Y%m%d")
        cmd = [sys.executable, str(REPO_ROOT / "scripts" / "build_snapshot.py"),
               "--date", date_str,
               "--all-city", str(tmp / all_city[newest]), "--region", str(tmp / region[newest]),
               "--out", str(REPO_ROOT / f"snapshot_{date_str}.json")]
        if wow:
            wd = wow[0].strftime("%Y%m%d")
            cmd += ["--wow-date", wd, "--wow-all-city", str(tmp / all_city[wow[0]]), "--wow-region", str(tmp / region[wow[0]])]
        if mom:
            md = mom[0].strftime("%Y%m%d")
            cmd += ["--prev-date", md, "--prev-all-city", str(tmp / all_city[mom[0]]), "--prev-region", str(tmp / region[mom[0]])]
        if yoy:
            yd = yoy[0].strftime("%Y%m%d")
            cmd += ["--yoy-date", yd, "--yoy-all-city", str(tmp / all_city[yoy[0]]), "--yoy-region", str(tmp / region[yoy[0]])]
        subprocess.run(cmd, cwd=REPO_ROOT, check=True)

    return {"pipeline": "meta", "changed": True, "newest": str(newest),
            "wow": str(wow[0]) if wow else None, "mom": str(mom[0]) if mom else None, "yoy": str(yoy[0]) if yoy else None}


# ==================== RETINA ====================
RETINA_RE = re.compile(r"^Retina - (\w+) (\d{4}) \(Monthly\)\.xlsx$")


def parse_retina_date(month_name, year):
    return datetime.strptime(f"{month_name} {year} 01", "%B %Y %d").date()


def refresh_retina(dataset_branch, files, dry_run):
    found = {}
    for f in files:
        m = RETINA_RE.match(f)
        if m:
            try:
                d = parse_retina_date(m.group(1), m.group(2))
            except ValueError:
                continue
            found[d] = f
    if not found:
        return {"pipeline": "retina", "changed": False, "reason": "no Retina - <Month> <Year> (Monthly).xlsx files found"}

    dates = sorted(found)
    current = read_json_date(REPO_ROOT / "retina_latest.json", "generatedAt")
    newest = dates[-1]
    if current is not None and newest <= current:
        return {"pipeline": "retina", "changed": False, "reason": f"newest available ({newest:%Y-%m}) not newer than published ({current:%Y-%m})"}

    others = [(d,) for d in dates if d != newest]
    wow = closest_in_window(newest, others, WOW_WINDOW)
    mom = closest_in_window(newest, others, MOM_WINDOW)
    yoy = closest_in_window(newest, others, YOY_WINDOW)

    if dry_run:
        return {"pipeline": "retina", "changed": True, "dry_run": True, "newest": f"{newest:%Y-%m}",
                "wow": f"{wow[0]:%Y-%m}" if wow else None, "mom": f"{mom[0]:%Y-%m}" if mom else None, "yoy": f"{yoy[0]:%Y-%m}" if yoy else None}

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        needed = [newest] + [c[0] for c in (wow, mom, yoy) if c]
        for d in needed:
            extract_file(dataset_branch, found[d], tmp / found[d])

        month_str = f"{newest:%Y-%m}"
        cmd = [sys.executable, str(REPO_ROOT / "scripts" / "build_retina_snapshot.py"),
               "--xlsx", str(tmp / found[newest]), "--month", month_str,
               "--out", str(REPO_ROOT / f"retina_{month_str}.json")]
        # build_retina_snapshot.py doesn't yet take wow/prev/yoy flags (single-snapshot
        # design) — if it's extended later, wire wow/mom/yoy files in here the same way
        # refresh_meta does. For now just note what would apply, so a human/future
        # agent can extend the script instead of this silently doing nothing with them.
        subprocess.run(cmd, cwd=REPO_ROOT, check=True)

    return {"pipeline": "retina", "changed": True, "newest": f"{newest:%Y-%m}",
            "wow": f"{wow[0]:%Y-%m}" if wow else None, "mom": f"{mom[0]:%Y-%m}" if mom else None, "yoy": f"{yoy[0]:%Y-%m}" if yoy else None,
            "note": "build_retina_snapshot.py doesn't support historical comparisons yet; only the newest file was processed."}


# ==================== Network Experience (ONX Spotlight) ====================
NETWORK_RE = re.compile(r"^ONXSpotlight_(\d{8})\.xlsx$")


def refresh_network(dataset_branch, files, dry_run):
    found = {}
    for f in files:
        m = NETWORK_RE.match(f)
        if m:
            found[parse_yyyymmdd(m.group(1))] = f
    if not found:
        return {"pipeline": "network", "changed": False, "reason": "no ONXSpotlight_YYYYMMDD.xlsx files found"}

    dates = sorted(found)
    current = read_json_date(REPO_ROOT / "network_latest.json", "asOfDate", date_len=10)
    newest = dates[-1]
    if current is not None and newest <= current:
        return {"pipeline": "network", "changed": False, "reason": f"newest available ({newest}) not newer than published ({current})"}

    if dry_run:
        return {"pipeline": "network", "changed": True, "dry_run": True, "newest": str(newest)}

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        extract_file(dataset_branch, found[newest], tmp / found[newest])
        date_str = newest.strftime("%Y%m%d")
        cmd = [sys.executable, str(REPO_ROOT / "scripts" / "build_network_snapshot.py"),
               "--xlsx", str(tmp / found[newest]),
               "--out", str(REPO_ROOT / f"network_{date_str}.json")]
        subprocess.run(cmd, cwd=REPO_ROOT, check=True)

    return {"pipeline": "network", "changed": True, "newest": str(newest),
            "note": "build_network_snapshot.py doesn't support historical comparisons yet; only the newest file was processed."}


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--dataset-branch", default="Dataset")
    ap.add_argument("--dry-run", action="store_true", help="report what would happen without writing/running anything")
    args = ap.parse_args()

    files = list_dataset_files(args.dataset_branch)
    results = [
        refresh_meta(args.dataset_branch, files, args.dry_run),
        refresh_retina(args.dataset_branch, files, args.dry_run),
        refresh_network(args.dataset_branch, files, args.dry_run),
    ]

    any_changed = any(r["changed"] for r in results)
    print(json.dumps({"anyChanged": any_changed, "results": results}, indent=2))
    sys.exit(0)


if __name__ == "__main__":
    main()
