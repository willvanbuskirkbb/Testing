#!/usr/bin/env python3
"""Recompute FY and Grand Opening Quarter from Grand Opening Date.

Nike's fiscal year runs June through May:
  FY27 = 2026-06-01 through 2027-05-31
  Q1 = Jun-Aug, Q2 = Sep-Nov, Q3 = Dec-Feb, Q4 = Mar-May

Targets the two live sheets:
  - NVS List (NEW)  sheet 8700079041892228
  - NSO List (NEW)  sheet 328848479571844

For each row:
  * Grand Opening Date set -> compute FY (e.g. "FY27") and GQ (e.g. "Q1");
    update only the cells whose current value differs.
  * Grand Opening Date blank -> leave both columns untouched (preserves the
    existing "Missing Grand Opening Date" / "TBD" markers).

Idempotent. --dry-run prints the diff without sending PUTs.

Auth: SMARTSHEET_ACCESS_TOKEN env var, or AWS Secrets Manager via
SMARTSHEET_AWS_SECRET_ID (default "smartsheet/access_token"), AWS_REGION
(default us-west-2). In this org the standing secret is
"sync-monitors/prod/smartsheet-api-token".
"""

from __future__ import annotations

import argparse
import json
import os
import random
import re
import sys
import time
from datetime import date, datetime
from typing import Any

import requests

API_BASE = "https://api.smartsheet.com/2.0"

SHEETS = [
    {
        "label": "NVS List (NEW)",
        "sheet_id": 8700079041892228,
        "go_date_col": 3109902451691396,
        "fy_col":      5335313986310020,
        "gq_col":      3083514172624772,
    },
    {
        "label": "NSO List (NEW)",
        "sheet_id": 328848479571844,
        "go_date_col": 3406040446750596,
        "fy_col":      6027276167368580,
        "gq_col":      3775476353683332,
    },
]

PUT_CHUNK_SIZE = 400
DATE_RE = re.compile(r"^(\d{4})-(\d{2})-(\d{2})")


def get_token() -> str:
    token = os.environ.get("SMARTSHEET_ACCESS_TOKEN")
    if token:
        return token.strip()
    secret_id = os.environ.get("SMARTSHEET_AWS_SECRET_ID", "smartsheet/access_token")
    region = os.environ.get("AWS_REGION", "us-west-2")
    import boto3  # type: ignore
    raw = boto3.client("secretsmanager", region_name=region).get_secret_value(
        SecretId=secret_id)["SecretString"]
    s = raw.strip()
    if s.startswith("{"):
        try:
            obj = json.loads(s)
        except json.JSONDecodeError:
            return s
        for key in ("token", "SMARTSHEET_ACCESS_TOKEN", "smartsheet_access_token",
                    "api_token", "access_token", "value"):
            if obj.get(key):
                return str(obj[key]).strip()
        str_vals = [v for v in obj.values() if isinstance(v, str) and v]
        if len(str_vals) == 1:
            return str_vals[0].strip()
        sys.exit(f"Could not find a token key in secret {secret_id}")
    return s


def parse_go_date(raw: Any) -> date | None:
    if raw is None:
        return None
    if isinstance(raw, date):
        return raw
    if isinstance(raw, datetime):
        return raw.date()
    s = str(raw).strip()
    if not s:
        return None
    m = DATE_RE.match(s)
    if not m:
        return None
    y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return date(y, mo, d)
    except ValueError:
        return None


def fy_and_quarter(go: date) -> tuple[str, str]:
    """Nike fiscal year: June -> May. FY label is e.g. 'FY27' for Jun-2026 onward."""
    if go.month >= 6:
        fy_year = go.year + 1
    else:
        fy_year = go.year
    fy = f"FY{fy_year % 100:02d}"
    # Quarter from calendar month (Jun=Q1 ... May=Q4):
    quarter_by_month = {
        6: "Q1", 7: "Q1", 8: "Q1",
        9: "Q2", 10: "Q2", 11: "Q2",
        12: "Q3", 1: "Q3", 2: "Q3",
        3: "Q4", 4: "Q4", 5: "Q4",
    }
    return fy, quarter_by_month[go.month]


class SmartsheetClient:
    def __init__(self, token: str):
        self.s = requests.Session()
        self.s.headers.update({
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
            "Accept": "application/json",
        })

    def _request(self, method: str, path: str, **kw) -> dict:
        url = f"{API_BASE}{path}"
        for attempt in range(1, 7):
            try:
                r = self.s.request(method, url, timeout=60, **kw)
            except requests.RequestException as e:
                if attempt == 6:
                    raise
                wait = min(2 ** attempt, 30) + random.uniform(0, 1)
                print(f"[retry] {method} {path}: {e!r}; sleeping {wait:.1f}s")
                time.sleep(wait)
                continue
            if r.status_code == 429 or 500 <= r.status_code < 600:
                if attempt == 6:
                    r.raise_for_status()
                wait = min(2 ** attempt, 30) + random.uniform(0, 1)
                print(f"[retry] {method} {path}: HTTP {r.status_code}; "
                      f"sleeping {wait:.1f}s")
                time.sleep(wait)
                continue
            if not r.ok:
                raise requests.HTTPError(
                    f"{method} {path} -> {r.status_code}: {r.text[:500]}",
                    response=r,
                )
            return r.json() if r.content else {}
        raise RuntimeError("unreachable")

    def get_sheet(self, sheet_id: int) -> dict:
        all_rows: list[dict] = []
        page = 1
        cols = None
        name = None
        while True:
            data = self._request(
                "GET", f"/sheets/{sheet_id}",
                params={"pageSize": 5000, "page": page, "include": "objectValue"},
            )
            if cols is None:
                cols = data.get("columns", [])
                name = data.get("name")
            all_rows.extend(data.get("rows", []))
            if page >= data.get("totalPages", 1):
                break
            page += 1
        return {"name": name, "columns": cols or [], "rows": all_rows}

    def update_rows(self, sheet_id: int, updates: list[dict]) -> int:
        total = 0
        for i in range(0, len(updates), PUT_CHUNK_SIZE):
            chunk = updates[i:i + PUT_CHUNK_SIZE]
            self._request("PUT", f"/sheets/{sheet_id}/rows", data=json.dumps(chunk))
            total += len(chunk)
        return total


def process_sheet(client: SmartsheetClient, cfg: dict, dry_run: bool,
                  sample_limit: int = 6) -> dict:
    label = cfg["label"]
    sheet_id = cfg["sheet_id"]
    sheet = client.get_sheet(sheet_id)
    rows = sheet["rows"]
    print(f"\n=== {label} (id {sheet_id}) — {len(rows)} rows ===")

    fy_col = cfg["fy_col"]
    gq_col = cfg["gq_col"]
    gd_col = cfg["go_date_col"]

    no_date = 0
    bad_date = 0
    fy_changes = gq_changes = 0
    already_ok = 0
    updates: list[dict] = []
    samples: list[str] = []

    for row in rows:
        cells = {c["columnId"]: c for c in row.get("cells", [])}
        gd_cell = cells.get(gd_col, {})
        fy_cell = cells.get(fy_col, {})
        gq_cell = cells.get(gq_col, {})
        go = parse_go_date(gd_cell.get("value") or gd_cell.get("displayValue"))
        if go is None:
            if gd_cell.get("value") or gd_cell.get("displayValue"):
                bad_date += 1
            else:
                no_date += 1
            continue

        fy_want, gq_want = fy_and_quarter(go)
        fy_have = fy_cell.get("value")
        gq_have = gq_cell.get("value")
        fy_have_s = None if fy_have is None else str(fy_have).strip()
        gq_have_s = None if gq_have is None else str(gq_have).strip()

        new_cells = []
        if fy_have_s != fy_want:
            new_cells.append({"columnId": fy_col, "value": fy_want})
            fy_changes += 1
        if gq_have_s != gq_want:
            new_cells.append({"columnId": gq_col, "value": gq_want})
            gq_changes += 1

        if not new_cells:
            already_ok += 1
            continue

        if len(samples) < sample_limit:
            samples.append(
                f"  row {row['id']}: GO={go.isoformat()}  "
                f"FY: {fy_have_s!r} -> {fy_want!r}  GQ: {gq_have_s!r} -> {gq_want!r}"
            )
        updates.append({"id": row["id"], "cells": new_cells})

    print(f"  rows with no Grand Opening Date (untouched): {no_date}")
    if bad_date:
        print(f"  rows with unparseable date (untouched):      {bad_date}")
    print(f"  already correct (skipped):                    {already_ok}")
    print(f"  to update:                                    {len(updates)}")
    print(f"    of which FY changes: {fy_changes}, GQ changes: {gq_changes}")
    if samples:
        print("  sample diffs:")
        for s in samples:
            print(s)

    if not updates:
        print("  -> no changes")
        return {"updates": 0, "fy": fy_changes, "gq": gq_changes}

    if dry_run:
        print("  -> DRY RUN: not sending PUT")
        return {"would_update": len(updates), "fy": fy_changes, "gq": gq_changes}

    n = client.update_rows(sheet_id, updates)
    print(f"  -> updated {n} rows")
    return {"updates": n, "fy": fy_changes, "gq": gq_changes}


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--dry-run", action="store_true")
    return p.parse_args()


def main() -> int:
    args = parse_args()
    client = SmartsheetClient(get_token())
    results = [(c["label"], process_sheet(client, c, args.dry_run)) for c in SHEETS]
    print("\n=== summary ===")
    for label, r in results:
        print(f"  {label}: {r}")
    print(f"  mode: {'DRY RUN' if args.dry_run else 'LIVE WRITE'}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
