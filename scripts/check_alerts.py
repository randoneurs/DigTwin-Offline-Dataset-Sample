#!/usr/bin/env python3
"""Evaluate real threshold-based alert rules against the published nationwide
meta market-share snapshot (snapshot_latest.json) and print any triggered
alerts as JSON.

This is intentionally separate from the dashboard's "Active alerts" card,
which currently renders fixed illustrative demo content (see index.html's
buildAlerts()) rather than deriving alerts from real thresholds. This script
is the real, data-driven counterpart, meant to be run by the daily scheduled
routine after auto_refresh_pipeline.py publishes new data.

Usage:
  python3 scripts/check_alerts.py [--snapshot snapshot_latest.json]

Exits 0 always; prints {"alerts": [...]} to stdout (empty list if nothing
triggered). Each alert has: id, severity ("critical"|"warning"|"info"),
title, detail.
"""
import argparse
import json
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parent.parent

# Thresholds (percentage points of national meta market share)
TSEL_WOW_DROP_WARN = 0.5
TSEL_WOW_DROP_CRIT = 1.0
TSEL_YOY_DROP_WARN = 1.5
TSEL_YOY_DROP_CRIT = 3.0
COMPETITOR_WOW_GAIN_WARN = 0.5

COMPETITOR_NAMES = {
    'ioh': 'Indosat Ooredoo Hutchison',
    'isat': 'Indosat (ISAT)',
    'xl': 'XL Axiata',
    'three': '3 (Tri)',
    'smartfren': 'Smartfren',
}


def find_nationwide_row(rows):
    for r in rows:
        if r.get('level') == 'Nationwide':
            return r
    return None


def check_meta_alerts(snapshot_path):
    alerts = []
    if not snapshot_path.exists():
        return alerts
    try:
        data = json.loads(snapshot_path.read_text())
    except (json.JSONDecodeError, OSError):
        return alerts

    row = find_nationwide_row(data.get('rows', []))
    if not row:
        return alerts

    week = row.get('week', '')
    tsel = row.get('metaMarketSharePct')
    tsel_wow_prev = row.get('metaMarketSharePrevWowPct')
    tsel_yoy_prev = row.get('metaMarketSharePrevYoyPct')

    if tsel is not None and tsel_wow_prev is not None:
        wow_delta = round(tsel - tsel_wow_prev, 2)
        if -wow_delta >= TSEL_WOW_DROP_CRIT:
            alerts.append({
                'id': 'tsel_wow_drop_critical', 'severity': 'critical',
                'title': f'TSEL national market share dropped {-wow_delta:.2f}pp week-over-week',
                'detail': f'{week}: {tsel:.2f}% vs {tsel_wow_prev:.2f}% last week (compared to {data.get("comparedToDateWow", "")[:10]}).',
            })
        elif -wow_delta >= TSEL_WOW_DROP_WARN:
            alerts.append({
                'id': 'tsel_wow_drop_warning', 'severity': 'warning',
                'title': f'TSEL national market share down {-wow_delta:.2f}pp week-over-week',
                'detail': f'{week}: {tsel:.2f}% vs {tsel_wow_prev:.2f}% last week (compared to {data.get("comparedToDateWow", "")[:10]}).',
            })

    if tsel is not None and tsel_yoy_prev is not None:
        yoy_delta = round(tsel - tsel_yoy_prev, 2)
        if -yoy_delta >= TSEL_YOY_DROP_CRIT:
            alerts.append({
                'id': 'tsel_yoy_drop_critical', 'severity': 'critical',
                'title': f'TSEL national market share down {-yoy_delta:.2f}pp year-over-year',
                'detail': f'{week}: {tsel:.2f}% vs {tsel_yoy_prev:.2f}% a year ago (compared to {data.get("comparedToDateYoy", "")[:10]}).',
            })
        elif -yoy_delta >= TSEL_YOY_DROP_WARN:
            alerts.append({
                'id': 'tsel_yoy_drop_warning', 'severity': 'warning',
                'title': f'TSEL national market share down {-yoy_delta:.2f}pp year-over-year',
                'detail': f'{week}: {tsel:.2f}% vs {tsel_yoy_prev:.2f}% a year ago (compared to {data.get("comparedToDateYoy", "")[:10]}).',
            })

    share = row.get('shareByOperator') or {}
    share_wow = row.get('shareByOperatorWow') or {}
    for key, label in COMPETITOR_NAMES.items():
        cur = share.get(key)
        prev = share_wow.get(key)
        if cur is None or prev is None:
            continue
        gain = round(cur - prev, 2)
        if gain >= COMPETITOR_WOW_GAIN_WARN:
            alerts.append({
                'id': f'{key}_wow_gain_warning', 'severity': 'warning',
                'title': f'{label} gained {gain:.2f}pp national share week-over-week',
                'detail': f'{week}: {cur:.2f}% vs {prev:.2f}% last week.',
            })

    return alerts


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--snapshot', default=str(REPO_ROOT / 'snapshot_latest.json'))
    args = ap.parse_args()

    alerts = check_meta_alerts(Path(args.snapshot))
    print(json.dumps({'alerts': alerts}, indent=2))
    sys.exit(0)


if __name__ == '__main__':
    main()
