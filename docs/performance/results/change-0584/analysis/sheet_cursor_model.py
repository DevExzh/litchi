#!/usr/bin/env python3
"""Model the per-sheet cursor chain walk of a source-backed XLS scan.

`WorksheetScan::new` builds one `SharedOleStreamCursor` per worksheet, seeking to
that sheet's declared `BoundSheet8` start. `stream_cursor_at` reaches the offset
by walking the `Workbook` stream's FAT chain from its **first** sector, so the
walk is repeated in full for every sheet.

This script prices that repetition directly from the file's bytes. It parses the
CFB header, FAT, DIFAT and directory, extracts the `Workbook` (or `Book`) stream,
walks the BIFF record sequence for `BoundSheet8` (0x0085) records, converts each
declared start offset to a sector ordinal, and sums two walks:

    links_today   -- every sheet walks from ordinal 0
    links_hinted  -- a sheet at or after the previous sheet's ordinal resumes
                     from it; a sheet before it restarts cold

It runs no repository code and needs no build, so it is an independent check on
the profile's own caller attribution rather than a restatement of it.

Bounds: FAT chains only (a `Workbook` stream short enough to live in the MiniFAT
is not modelled), and the resumable position is the cursor's construction ordinal
rather than a position advanced by the cursor's own later reads.

Usage:
    python3 sheet_cursor_model.py <file.xls> [<file.xls> ...]
"""

import os
import struct
import sys

from sst_walk import cfb_workbook

BOUNDSHEET8 = 0x0085


def sheet_start_ordinals(workbook: bytes, sector_size: int) -> list[int]:
    """Sector ordinal of every BoundSheet8-declared worksheet start."""
    ordinals = []
    offset = 0
    while offset + 4 <= len(workbook):
        record_id, record_len = struct.unpack_from("<HH", workbook, offset)
        if record_id == BOUNDSHEET8 and record_len >= 6:
            start = struct.unpack_from("<I", workbook, offset + 4)[0]
            ordinals.append(start // sector_size)
        offset += 4 + record_len
    return ordinals


def price(ordinals: list[int]) -> tuple[int, int]:
    """Return (links_today, links_hinted) for this sheet ordering."""
    today = sum(ordinals)
    hinted = 0
    previous = None
    for ordinal in ordinals:
        if previous is None or ordinal < previous:
            hinted += ordinal
        else:
            hinted += ordinal - previous
        previous = ordinal
    return today, hinted


def main(paths: list[str]) -> None:
    for path in paths:
        try:
            workbook, sector_size = cfb_workbook(path)
        except Exception as error:  # noqa: BLE001 - a corpus sweep must not stop
            print(f"{os.path.basename(path):40s} SKIP {type(error).__name__}: {error}")
            continue
        ordinals = sheet_start_ordinals(workbook, sector_size)
        if not ordinals:
            continue
        today, hinted = price(ordinals)
        saving = 100.0 * (today - hinted) / max(today, 1)
        print(
            f"{os.path.basename(path):40s} "
            f"sheets={len(ordinals):3d} "
            f"ordinals={min(ordinals)}..{max(ordinals)} "
            f"links_today={today:7d} links_hinted={hinted:6d} saving={saving:5.1f}%"
        )


if __name__ == "__main__":
    main(sys.argv[1:])
