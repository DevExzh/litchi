#!/usr/bin/env python3
"""Independent size and repeated-query pricing probe for source-backed XLS.

This probe does not import litchi-xls.  It uses olefile only to obtain a
Workbook/Book stream, then walks BIFF8 frames arithmetically.  The
wire facts correspond to the current source owner:

* ``SheetEntry`` is bounded by each BoundSheet8 offset and the next offset;
* ``WorksheetScan`` visits every frame through EOF;
* ``CellSlot`` candidates are one slot per recognized stored-cell occurrence,
  expanding MulRk/MulBlank exactly as the 0668 measure-only path does;
* ``SharedStringSstScan`` retains one segment locator per SST/Continue record
  and one entry locator per unique string.

The output is JSONL so a reviewer can recompute all aggregates without a Rust
build.  olefile's successful stream extraction does not prove that the
litchi-xls source-backed owner admits the fixture; malformed, encrypted, and
unsupported fixtures remain labelled as wire-level probe data.  It is
evidence about retained logical bytes and frame work, not a claim about
allocator overhead or runtime latency.
"""

import argparse
import hashlib
import json
import os
import struct
from pathlib import Path

import olefile


EOF = 0x000A
BOF = 0x0809
BOUND_SHEET = 0x0085
SST = 0x00FC
CONTINUE = 0x003C

# The current frozen design's honest Rust layout: u64 stream offset plus five
# u16 fields, rounded to 8-byte alignment.  The wrapper/lookup charge is kept
# separate and deliberately conservative.
CELL_SLOT_BYTES = 24
SST_SEGMENT_BYTES = 24  # u64 source_offset + usize logical_offset + usize len
SST_ENTRY_BYTES = 16  # usize start + usize end
INDEX_ENTRY_OVERHEAD = 64

CELL_KINDS = {
    0x0006: "formula",
    0x0201: "blank",
    0x0203: "number",
    0x0204: "label",
    0x0205: "boolerr",
    0x027E: "rk",
    0x00BD: "mulrk",
    0x00BE: "mulblank",
    0x00FD: "labelsst",
}


def read_workbook(path: Path) -> bytes:
    with olefile.OleFileIO(str(path), write_mode=False) as ole:
        for name in ("Workbook", "Book"):
            if ole.exists(name):
                return ole.openstream(name).read()
    raise ValueError("Workbook/Book stream is missing")


def frames(data: bytes):
    result = []
    offset = 0
    while offset + 4 <= len(data):
        kind, length = struct.unpack_from("<HH", data, offset)
        end = offset + 4 + length
        if end > len(data):
            raise ValueError(f"truncated BIFF frame at {offset}")
        result.append((offset, kind, length))
        offset = end
        if kind == EOF:
            break
    return result


def bound_sheets(data: bytes, recs):
    sheets = []
    for offset, kind, length in recs:
        if kind != BOUND_SHEET or length < 8:
            continue
        position = struct.unpack_from("<I", data, offset + 4)[0]
        sheet_type = data[offset + 9]
        if sheet_type != 0:
            continue
        char_count = data[offset + 10]
        flags = data[offset + 11]
        name_bytes = char_count * (2 if flags & 1 else 1)
        if 12 + name_bytes > offset + 4 + length:
            name = "<malformed-name>"
        elif flags & 1:
            name = data[offset + 12 : offset + 12 + name_bytes].decode("utf-16le", "replace")
        else:
            name = data[offset + 12 : offset + 12 + name_bytes].decode("cp1252", "replace")
        sheets.append({"start": position, "name": name})
    sheets.sort(key=lambda sheet: sheet["start"])
    for index, sheet in enumerate(sheets):
        sheet["end"] = sheets[index + 1]["start"] if index + 1 < len(sheets) else len(data)
    return sheets


def sst_shape(data: bytes, recs):
    groups = []
    index = 0
    while index < len(recs):
        offset, kind, length = recs[index]
        if kind != SST:
            index += 1
            continue
        group = [recs[index]]
        index += 1
        while index < len(recs) and recs[index][1] == CONTINUE:
            group.append(recs[index])
            index += 1
        groups.append(group)
    if not groups:
        return {"segments": 0, "entries": 0, "payload_bytes": 0, "wire_bytes": 0}
    if len(groups) != 1:
        raise ValueError("multiple SST records")
    group = groups[0]
    first_payload = group[0][0] + 4
    last_end = group[-1][0] + 4 + group[-1][2]
    payload_bytes = sum(length for _offset, _kind, length in group)
    if payload_bytes < 8:
        raise ValueError("short SST header")
    total, unique = struct.unpack_from("<II", data, first_payload)
    # Keep the same cheap structural admissions as the current open-time
    # scan.  A hostile encrypted fixture can carry a giant count in an 8-byte
    # SST payload; charging that declaration as retained memory would price a
    # structure the owner refuses before publication.
    valid = (
        total <= 1_000_000
        and unique <= 1_000_000
        and total >= unique
        and unique <= max(0, payload_bytes - 8) // 3
    )
    return {
        "segments": len(group),
        "entries": unique,
        "total_entries": total,
        "payload_bytes": payload_bytes,
        "wire_bytes": last_end - group[0][0],
        "admissible": valid,
        "invalid_reason": None if valid else "open-time SST count/length limit",
    }


def worksheet_shape(data: bytes, sheet):
    start = sheet["start"]
    end = min(sheet["end"], len(data))
    if start >= end:
        raise ValueError("empty worksheet region")
    rows = []
    offset = start
    cell_records = 0
    packed_cells = 0
    unique_positions = set()
    while offset + 4 <= end:
        kind, length = struct.unpack_from("<HH", data, offset)
        frame_end = offset + 4 + length
        if frame_end > end:
            raise ValueError(f"worksheet frame crosses boundary at {offset}")
        payload = data[offset + 4 : frame_end]
        rows.append((offset, kind, length))
        if kind in CELL_KINDS:
            if kind == 0x00BD and len(payload) >= 6:
                row, first_col = struct.unpack_from("<HH", payload, 0)
                last_col = struct.unpack_from("<H", payload, len(payload) - 2)[0]
                count = max(0, last_col - first_col + 1)
                packed_cells += count
                cell_records += count
                unique_positions.update((row, first_col + i) for i in range(count))
            elif kind == 0x00BE and len(payload) >= 6:
                row, first_col = struct.unpack_from("<HH", payload, 0)
                last_col = struct.unpack_from("<H", payload, len(payload) - 2)[0]
                count = max(0, last_col - first_col + 1)
                packed_cells += count
                cell_records += count
                unique_positions.update((row, first_col + i) for i in range(count))
            elif len(payload) >= 4:
                row, column = struct.unpack_from("<HH", payload, 0)
                cell_records += 1
                unique_positions.add((row, column))
        offset = frame_end
        if kind == EOF:
            break
    if not rows or rows[0][1] != BOF or rows[-1][1] != EOF:
        raise ValueError("worksheet does not have BOF..EOF")
    return {
        "name": sheet["name"],
        "start": start,
        "end": end,
        "span_bytes": end - start,
        "frames": len(rows),
        "framed_bytes": sum(4 + length for _offset, _kind, length in rows),
        "cell_records": cell_records,
        "unique_positions": len(unique_positions),
        "packed_cells": packed_cells,
        # Keep every occurrence.  A prior invalid duplicate at the requested
        # coordinate can refuse before a later valid duplicate wins, so a
        # latest-only index would change TargetCell's error semantics.
        "slot_weight_bytes": cell_records * CELL_SLOT_BYTES + INDEX_ENTRY_OVERHEAD,
    }


def analyze(path: Path):
    data = read_workbook(path)
    recs = frames(data)
    sheets = bound_sheets(data, recs)
    worksheets = [worksheet_shape(data, sheet) for sheet in sheets]
    sst = sst_shape(data, recs)
    sst["logical_weight_bytes"] = (
        sst["segments"] * SST_SEGMENT_BYTES
        + sst["entries"] * SST_ENTRY_BYTES
        + INDEX_ENTRY_OVERHEAD
        if sst.get("admissible") and (sst["segments"] or sst["entries"])
        else 0
    )
    return {
        "fixture": str(path),
        "fixture_sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "workbook_bytes": len(data),
        "worksheets": worksheets,
        "worksheet_count": len(worksheets),
        "worksheet_slot_occurrences": sum(item["cell_records"] for item in worksheets),
        "worksheet_distinct_positions": sum(item["unique_positions"] for item in worksheets),
        "worksheet_weight_bytes": sum(item["slot_weight_bytes"] for item in worksheets),
        "sst": sst,
    }


def percentile(values, fraction):
    values = sorted(values)
    if not values:
        return 0
    index = min(len(values) - 1, int(round((len(values) - 1) * fraction)))
    return values[index]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("root", nargs="?", default="test-data")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    paths = sorted(
        path
        for path in Path(args.root).rglob("*")
        if path.is_file() and path.suffix.lower() in {".xls", ".xlt"}
    )
    rows = []
    refused = []
    for path in paths:
        try:
            rows.append(analyze(path))
        except Exception as error:  # noqa: BLE001 - probe records refusals
            refused.append({"fixture": str(path), "error": str(error)})
    output = args.output.open("w", encoding="utf-8") if args.output else None
    stream = output or os.sys.stdout
    for row in rows:
        stream.write(json.dumps(row, sort_keys=True) + "\n")
    if output:
        output.close()

    sheets = [sheet for row in rows for sheet in row["worksheets"]]
    ssts = [
        row["sst"]
        for row in rows
        if row["sst"]["segments"] and row["sst"].get("admissible")
    ]
    summary = {
        "fixtures_seen": len(paths),
        "fixtures_analyzed": len(rows),
        "fixtures_refused": refused,
        "worksheet_substreams": len(sheets),
        "worksheet_slot_occurrences": sum(sheet["cell_records"] for sheet in sheets),
        "worksheet_distinct_positions": sum(sheet["unique_positions"] for sheet in sheets),
        "worksheet_weight_bytes": sum(sheet["slot_weight_bytes"] for sheet in sheets),
        "worksheet_weight_quantiles": {
            "p50": percentile([sheet["slot_weight_bytes"] for sheet in sheets], 0.50),
            "p95": percentile([sheet["slot_weight_bytes"] for sheet in sheets], 0.95),
            "max": max((sheet["slot_weight_bytes"] for sheet in sheets), default=0),
        },
        "max_sheet_slots": max((sheet["unique_positions"] for sheet in sheets), default=0),
        "sst_workbooks": len(ssts),
        "sst_logical_weight_bytes": sum(item["logical_weight_bytes"] for item in ssts),
        "sst_weight_quantiles": {
            "p50": percentile([item["logical_weight_bytes"] for item in ssts], 0.50),
            "p95": percentile([item["logical_weight_bytes"] for item in ssts], 0.95),
            "max": max((item["logical_weight_bytes"] for item in ssts), default=0),
        },
    }
    print(json.dumps({"summary": summary}, sort_keys=True), file=os.sys.stderr)


if __name__ == "__main__":
    main()
