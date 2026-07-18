#!/usr/bin/env python3
"""Build the compact on-device Medicinpriser dataset (skive 2.5).

Two modes:
  --inspect   Download the source, print sheet names + column headers + a few sample rows.
              RUN THIS FIRST to verify the real field names, then fill in COLUMN_MAP below.
  --build     Transform the source into the compact JSON contract the app decodes, gzipped,
              with one (latest) record per varenummer. Asserts the result is < 3 MB.

Privacy note: this produces the dataset that ships to the phone. All lookups happen on-device;
no varenummer is ever sent to a server. See README.md.
"""
from __future__ import annotations

import argparse
import gzip
import io
import json
import sys
from datetime import date

import requests
from openpyxl import load_workbook

# Source: eSundhed/Medicinpriser (moved to sundhedsdatabank.dk). Verify/refresh this URL — the
# path is dated. See README "Åbne spørgsmål".
SOURCE_URL = (
    "https://cdn1.gopublic.dk/sundhedsdatastyrelsen/Media/639177217206502773/"
    "26_00015-13-medicinpriser-udgivet-22062026-4609666_1_0.xlsx"
)
SOURCE_LABEL = "Lægemiddelstyrelsen / Medicinpriser"
MAX_BYTES = 3 * 1024 * 1024  # spec acceptkriterie 5

# TODO: fill in the real column headers from `--inspect` output. Keys are our compact fields;
# Verified against the real "Medicinpriser data" sheet (--inspect, 2026-06-22). Values are the
# source column headers; price/date columns are ignored. 3 rows per product (AIP/AUP/DDD price
# indicators) share the same metadata → deduped to one record per varenummer.
SHEET_NAME = "Medicinpriser data"
COLUMN_MAP = {
    "vnr": "Varenummer",
    "navn": "Lægemiddel",
    "aktivstof": "Indholdsstof",
    "atc": "ATC",
    "form": "Form",
    "styrke": "Styrke",
    "pakning": "Pakning",
}


def download() -> bytes:
    print(f"Henter {SOURCE_URL} …", file=sys.stderr)
    resp = requests.get(SOURCE_URL, timeout=300)
    resp.raise_for_status()
    print(f"  {len(resp.content):,} bytes", file=sys.stderr)
    return resp.content


def inspect(data: bytes) -> None:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    for sheet in wb.sheetnames:
        ws = wb[sheet]
        print(f"\n=== Ark: {sheet} ===")
        for i, row in enumerate(ws.iter_rows(values_only=True)):
            if i == 0:
                print("Kolonner:", [str(c) for c in row])
            elif i <= 3:
                print(f"  Række {i}:", row)
            else:
                break


def build(data: bytes) -> bytes:
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    if SHEET_NAME not in wb.sheetnames:
        raise SystemExit(f"Fandt ikke arket '{SHEET_NAME}'. Ark: {wb.sheetnames}")
    ws = wb[SHEET_NAME]

    rows = ws.iter_rows(values_only=True)
    header = [str(c).strip() if c is not None else "" for c in next(rows)]
    idx = {field: header.index(col) for field, col in COLUMN_MAP.items() if col in header}
    if "vnr" not in idx:
        raise SystemExit(
            f"Fandt ikke varenummer-kolonnen '{COLUMN_MAP['vnr']}'. Kør --inspect og ret COLUMN_MAP.\n"
            f"Kolonner i arket: {header}"
        )

    records: dict[str, dict] = {}
    for row in rows:
        vnr_raw = row[idx["vnr"]]
        if vnr_raw is None:
            continue
        vnr = "".join(ch for ch in str(vnr_raw) if ch.isdigit())
        if len(vnr) != 6:
            continue
        record = {}
        for field, col_idx in idx.items():
            if field == "vnr":
                continue
            value = row[col_idx]
            if value is not None and str(value).strip():
                record[field] = str(value).strip()
        # Latest row wins (the sheet is chronological); keep overwriting.
        records[vnr] = record

    dataset = {"version": date.today().isoformat(), "source": SOURCE_LABEL, "records": records}
    raw = json.dumps(dataset, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    packed = gzip.compress(raw, compresslevel=9)
    print(f"{len(records):,} varenumre · {len(raw):,} bytes JSON · {len(packed):,} bytes gzip",
          file=sys.stderr)
    if len(packed) > MAX_BYTES:
        raise SystemExit(f"Datasæt {len(packed):,} B > loft {MAX_BYTES:,} B — stram feltvalg/dedup.")
    return packed


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--inspect", action="store_true", help="print sheets/columns/sample rows")
    ap.add_argument("--build", action="store_true", help="build the compact gzipped dataset")
    ap.add_argument("--out", default="medicinpriser.json.gz")
    args = ap.parse_args()

    if not (args.inspect or args.build):
        ap.error("angiv --inspect eller --build")

    data = download()
    if args.inspect:
        inspect(data)
    if args.build:
        packed = build(data)
        with open(args.out, "wb") as f:
            f.write(packed)
        print(f"Skrev {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
