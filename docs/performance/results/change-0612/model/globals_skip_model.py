#!/usr/bin/env python3
"""Price skipping the BIFF8 globals payloads the semantic pass never interprets.

Change 0574 opportunity 2 modelled the *closure* — how many globals bytes a
source-backed XLS open would have to read if it fetched only the four-byte
headers and the payloads of the twelve consumed record kinds.  It did not model
the *schedule*: how many read requests either shape actually issues.  Without
that the trade the opportunity makes — fewer bytes for more requests — cannot be
priced, and change 0587 ranked the item behind exactly that question.

This model carries three things:

1.  a byte-exact replay of **today's** `parse_globals` fill schedule
    (`crates/litchi-xls/src/workbook/source.rs`: the four-record exact prologue,
    the coupled next-header prefetch, then 512-byte windows doubling to 64 KiB,
    clamped by the stream length, by `max_global_bytes` and by the running
    minimum `BoundSheet8` position);
2.  the same replay for a **candidate** scan that keeps a sliding window like
    `WorksheetScan`, seeks past a payload the semantic pass will not read, and
    gates the seek on a density rule in the shape change 0568 built for the
    worksheet scan;
3.  the **gate census**: on how many real fixtures the gate fires, what it
    saves, and what it costs.

Deterministic.  Standard library plus `olefile`; no repository code executes,
so the replay is an independent oracle for the counted evidence rather than a
restatement of it.

Usage:
    python3 globals_skip_model.py --repo /home/zhuhe/code/litchi --out out.json
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from collections import Counter

try:
    import olefile
except ImportError:  # pragma: no cover
    sys.exit("olefile required")

# ---------------------------------------------------------------------------
# Constants, read from crates/litchi-xls/src/workbook/source.rs at 1e4198321.
# ---------------------------------------------------------------------------

BOF = 0x0809
EOF = 0x000A
CODEPAGE = 0x0042
BOUND_SHEET = 0x0085
FILEPASS = 0x002F
SST = 0x00FC
CONTINUE = 0x003C

GLOBALS_EXACT_PROLOGUE_RECORDS = 4
GLOBALS_FIRST_WINDOW_BYTES = 512
GLOBALS_MAX_WINDOW_BYTES = 64 * 1024
DEFAULT_GLOBAL_BYTES = 128 * 1024 * 1024
MAX_RECORD_BYTES = 8224

# Payload-consuming kinds.  `parse_globals` itself consumes FILEPASS, CODEPAGE,
# BOUND_SHEET, SST and the CONTINUE run that follows an SST; it also parses the
# first record's payload as BOF and requires the last to be an empty EOF.
# `Formatting::parse_globals` (number_format/codec.rs) consumes Date1904,
# Format, XF, XFCRC, XFExt and DXF.  Verified against the tree at 1e4198321:
#   number_format/mod.rs  DATE1904_RECORD 0x0022  FORMAT_RECORD 0x041e
#                         XF_RECORD 0x00e0        XFCRC_RECORD 0x087c
#   xf_ext.rs             XF_EXT_RECORD_TYPE 0x087d
#   differential_format/mod.rs DXF_RECORD_TYPE 0x088d
CONSUMED = {
    0x0809: "BOF",
    0x000A: "EOF",
    0x002F: "FilePass",
    0x0042: "CodePage",
    0x0085: "BoundSheet8",
    0x00FC: "SST",
    0x0022: "Date1904",
    0x041E: "Format",
    0x00E0: "XF",
    0x087C: "XFCRC",
    0x087D: "XFExt",
    0x088D: "DXF",
}

NAMES = dict(CONSUMED)
NAMES.update(
    {
        0x003C: "Continue",
        0x00EB: "MsoDrawingGroup",
        0x00FF: "ExtSST",
        0x0293: "Style",
        0x0892: "StyleExt",
        0x0031: "Font",
        0x005C: "WriteAccess",
        0x0018: "Lbl",
        0x005A: "Crn",
        0x088E: "TableStyles",
        0x0899: "Theme",
        0x08C8: "PLV",
        0x00E2: "Interface",
        0x0160: "UsesElfs",
        0x01AF: "Prot4Rev",
        0x00DA: "BookBool",
        0x0161: "DSF",
        0x0050: "Dconn",
        0x00D3: "Obproj",
        0x0092: "Palette",
        0x0085: "BoundSheet8",
        0x0083: "HCenter",
        0x008C: "Country",
        0x0040: "Backup",
        0x008D: "HideObj",
        0x000C: "Calccount",
        0x000D: "Calcmode",
        0x000E: "Precision",
        0x000F: "Refmode",
        0x0010: "Delta",
        0x0011: "Iteration",
        0x0013: "Password",
        0x0019: "WinProtect",
        0x001C: "Note",
        0x0022: "Date1904",
        0x002A: "PrintRowCol",
        0x002B: "PrintGrid",
        0x003D: "Window1",
        0x005B: "FileSharing",
        0x0081: "WsBool",
        0x00BD: "MulRk",
        0x00C1: "HLink",
        0x00DC: "Param",
        0x00E1: "InterfaceHdr",
        0x01C0: "ExcelServerstuff",
        0x0200: "Dimensions",
        0x023E: "Window2",
        0x0862: "SheetExt",
        0x0863: "BookExt",
        0x0867: "SXAddl",
        0x0868: "CrErr",
        0x086C: "GUIDTypeLib",
        0x087B: "RichTextStream",
        0x089A: "MTRSettings",
        0x089B: "CompressPictures",
        0x089C: "ForceFullCalculation",
        0x08A3: "ForceFullCalculation2",
        0x08A4: "ShapePropsStream",
    }
)


def kind_name(kind: int) -> str:
    return NAMES.get(kind, f"0x{kind:04X}")


def u16(buf: bytes, off: int) -> int:
    return buf[off] | (buf[off + 1] << 8)


def u32(buf: bytes, off: int) -> int:
    return (
        buf[off]
        | (buf[off + 1] << 8)
        | (buf[off + 2] << 16)
        | (buf[off + 3] << 24)
    )


# ---------------------------------------------------------------------------
# Framing
# ---------------------------------------------------------------------------


def frame_globals(stream: bytes):
    """Frame the globals substream exactly as `parse_globals`' loop does.

    Returns `(records, globals_end, issue)` where each record is
    `(offset, kind, payload_len, consumed)`.  `consumed` is True when the
    semantic pass reads the payload: the twelve kinds above, plus a CONTINUE
    inside the run that follows an SST record.
    """
    records = []
    off = 0
    n = len(stream)
    prev_non_continue = None
    in_sst_run = False
    while off + 4 <= n:
        kind = u16(stream, off)
        length = u16(stream, off + 2)
        end = off + 4 + length
        if length > MAX_RECORD_BYTES:
            return records, off, "oversize-record"
        if end > n:
            return records, off, "truncated"
        if kind == SST:
            in_sst_run = True
        elif kind != CONTINUE:
            in_sst_run = False
        consumed = kind in CONSUMED or (kind == CONTINUE and in_sst_run)
        records.append((off, kind, length, consumed))
        prev_non_continue = kind if kind != CONTINUE else prev_non_continue
        off = end
        if kind == EOF:
            return records, off, None
    return records, off, "no-eof"


def continuation_owner(records):
    """Bytes of CONTINUE payload attributed to the record each run continues."""
    owned = Counter()
    previous = None
    for _off, kind, length, _consumed in records:
        if kind == CONTINUE:
            if previous is not None:
                owned[previous] += length
        else:
            previous = kind
    return owned


# ---------------------------------------------------------------------------
# Schedule replay: today
# ---------------------------------------------------------------------------


class TodaySchedule:
    """Byte-exact replay of `GlobalsBuffer` plus the `parse_globals` loop.

    One `read_stream_range_hinted` call per fill; one source observation per
    call.  The buffer is a prefix `[0, filled)`, so every byte below the last
    fill is read whether the semantic pass reads it or not.
    """

    def __init__(self, stream_len, max_global_bytes=DEFAULT_GLOBAL_BYTES):
        self.stream_len = stream_len
        self.max_global_bytes = max_global_bytes
        self.filled = 0
        self.window = GLOBALS_FIRST_WINDOW_BYTES
        self.min_sheet_start = None
        self.sheet_clamp_active = True
        self.reads = []          # (start, end)

    def fill_cap(self, need):
        if (
            self.sheet_clamp_active
            and self.min_sheet_start is not None
            and need > self.min_sheet_start
        ):
            self.sheet_clamp_active = False
        cap = min(self.stream_len, max(self.max_global_bytes, need))
        if self.sheet_clamp_active and self.min_sheet_start is not None:
            cap = min(cap, self.min_sheet_start)
        return max(cap, need)

    def ensure(self, need, exact):
        if need <= self.filled:
            return
        if exact:
            end = need
        else:
            cap = self.fill_cap(need)
            end = min(max(need, self.filled + self.window), cap)
            self.window = min(self.window * 2, GLOBALS_MAX_WINDOW_BYTES)
        self.reads.append((self.filled, end))
        self.filled = end

    def note_sheet_start(self, position):
        self.min_sheet_start = (
            position
            if self.min_sheet_start is None
            else min(self.min_sheet_start, position)
        )


def replay_today(stream, records, max_global_bytes=DEFAULT_GLOBAL_BYTES):
    sched = TodaySchedule(len(stream), max_global_bytes)
    for index, (off, kind, length, _consumed) in enumerate(records):
        exact = index < GLOBALS_EXACT_PROLOGUE_RECORDS
        sched.ensure(off + 4, exact)
        end = off + 4 + length
        if exact and kind != EOF:
            need = min(end + 4, sched.stream_len)
        else:
            need = end
        sched.ensure(need, exact)
        if kind == BOUND_SHEET and length >= 4:
            sched.note_sheet_start(u32(stream, off + 4))
    return sched


# ---------------------------------------------------------------------------
# Schedule replay: the candidate
# ---------------------------------------------------------------------------


class CandidateSchedule:
    """A sliding globals window that seeks past payloads nothing interprets.

    Shape, taken from `WorksheetScan` (change 0568): the window holds
    `[window_start, window_start + len)`; a fill drops the framed prefix and
    appends one read; a skipped payload that ends past the filled end is passed
    with `skip_forward`, which reads nothing and takes no observation.

    The gate is the difference from 0568, and it is exact rather than
    predictive.  0568 had to decide a fill size *before* framing the record the
    fill would cover, so its rule is a running mean over records already framed.
    Here the skip decision is taken *after* the header is framed, so the payload
    length is known exactly.  Two rules therefore compose:

      * `skip_min` — the per-record rule.  A payload is seeked past only when it
        is at least this many bytes, so a short skippable payload is read
        through rather than paid for with a request.
      * `dense_mean` — the fill-size rule, the running-mean form of 0568.
        `skipped_bytes_seen / records_framed > dense_mean` makes the next fill
        exact instead of windowed, so a window does not read ahead over
        payloads the scan is about to skip.  The mean is recomputed for every
        fill, so it is its own hysteresis and not a one-way latch.

      * `reset_on_skip` — the hysteresis.  A seek empties the window, so the
        next fill starts at the record after the skipped payload; if the fill
        target kept doubling, that fill would read straight back over the next
        skippable payload and the skip would fire at most once per window.
        Resetting the target to `GLOBALS_FIRST_WINDOW_BYTES` on a seek keeps a
        run of large skippable records on small fills, and lets the target
        double again as soon as consumed records resume.  This is state-free
        hysteresis: the schedule, not a counter, remembers.

    Setting `skip_min = 0`, `dense_mean = -1` and `reset_on_skip = False` gives
    the ungated shape, which is what the 0574 closure model priced.
    """

    def __init__(
        self,
        stream_len,
        skip_min,
        dense_mean,
        reset_on_skip=True,
        max_global_bytes=DEFAULT_GLOBAL_BYTES,
    ):
        self.stream_len = stream_len
        self.skip_min = skip_min
        self.dense_mean = dense_mean
        self.reset_on_skip = reset_on_skip
        self.max_global_bytes = max_global_bytes
        self.window_start = 0
        self.window_len = 0
        self.target = GLOBALS_FIRST_WINDOW_BYTES
        self.min_sheet_start = None
        self.sheet_clamp_active = True
        self.skipped_bytes = 0
        self.records_framed = 0
        self.reads = []          # (start, end)
        self.seeks = 0
        self.retained = 0        # bytes copied into the compacted buffer
        self.gate_fired = False  # the dense rule made at least one fill exact
        self.skips = 0

    def filled_end(self):
        return self.window_start + self.window_len

    def dense(self):
        if self.records_framed == 0:
            return False
        return self.skipped_bytes // self.records_framed > self.dense_mean

    def fill_cap(self, need):
        if (
            self.sheet_clamp_active
            and self.min_sheet_start is not None
            and need > self.min_sheet_start
        ):
            self.sheet_clamp_active = False
        cap = min(self.stream_len, max(self.max_global_bytes, need))
        if self.sheet_clamp_active and self.min_sheet_start is not None:
            cap = min(cap, self.min_sheet_start)
        return max(cap, need)

    def ensure(self, position, need_end, exact=False):
        if need_end <= self.filled_end():
            return
        # Drop the framed prefix, exactly as `WorksheetScan::fill` does.
        framed = min(max(position - self.window_start, 0), self.window_len)
        self.window_start += framed
        self.window_len -= framed
        filled_end = self.filled_end()
        if exact:
            end = need_end
        elif self.dense():
            self.gate_fired = True
            self.target = GLOBALS_FIRST_WINDOW_BYTES
            end = need_end
        else:
            cap = self.fill_cap(need_end)
            end = min(max(need_end, filled_end + self.target), cap)
            end = max(end, need_end)
            self.target = min(self.target * 2, GLOBALS_MAX_WINDOW_BYTES)
        self.reads.append((filled_end, end))
        self.window_len += end - filled_end

    def skip(self, position, end):
        """Pass a payload the semantic pass will not read."""
        filled_end = self.filled_end()
        if end <= filled_end:
            # Already resident: the window read it, so nothing is saved here.
            return False
        self.seeks += 1
        self.window_start = end
        self.window_len = 0
        if self.reset_on_skip:
            self.target = GLOBALS_FIRST_WINDOW_BYTES
        return True

    def note_sheet_start(self, position):
        self.min_sheet_start = (
            position
            if self.min_sheet_start is None
            else min(self.min_sheet_start, position)
        )


def replay_candidate(
    stream,
    records,
    skip_min,
    dense_mean,
    reset_on_skip=True,
    max_global_bytes=DEFAULT_GLOBAL_BYTES,
):
    sched = CandidateSchedule(
        len(stream), skip_min, dense_mean, reset_on_skip, max_global_bytes
    )
    position = 0
    for index, (off, kind, length, consumed) in enumerate(records):
        assert off == position
        # The exact prologue is unchanged: a FilePass in the first five record
        # positions is still refused before any payload byte is requested.
        exact = index < GLOBALS_EXACT_PROLOGUE_RECORDS
        sched.ensure(position, off + 4, exact)
        end = off + 4 + length
        sched.records_framed += 1
        take = consumed or length < skip_min
        if take:
            if exact and kind != EOF:
                need = min(end + 4, sched.stream_len)
            else:
                need = end
            sched.ensure(position, need, exact)
            if consumed:
                sched.retained += 4 + length
            else:
                # Read through a short skippable payload: its bytes land in the
                # window but are never copied into the compacted buffer.
                sched.retained += 4
                sched.skipped_bytes += length
        else:
            sched.retained += 4
            sched.skipped_bytes += length
            sched.skips += 1
            sched.skip(position, end)
        if kind == BOUND_SHEET and length >= 4:
            sched.note_sheet_start(u32(stream, off + 4))
        position = end
    return sched


# ---------------------------------------------------------------------------
# Per-fixture analysis
# ---------------------------------------------------------------------------


def analyse(path, gates):
    try:
        ole = olefile.OleFileIO(path)
    except Exception as exc:  # noqa: BLE001
        return {"path": path, "error": f"cfb: {exc}"}
    name = next((c for c in ("Workbook", "Book") if ole.exists(c)), None)
    if name is None:
        ole.close()
        return None
    stream = ole.openstream(name).read()
    ole.close()

    records, globals_end, issue = frame_globals(stream)
    if not records or records[0][1] != BOF:
        return {"path": path, "error": "no BOF"}
    if any(kind == FILEPASS for _o, kind, _l, _c in records):
        # Encrypted: the open refuses inside the exact prologue, so no schedule
        # exists past it.  Reported, not modelled.
        return {
            "path": path,
            "error": "encrypted (FilePass)",
            "globals_records": len(records),
            "globals_end": globals_end,
        }

    header_bytes = 4 * len(records)
    consumed_payload = sum(l for _o, _k, l, c in records if c)
    skipped_payload = sum(l for _o, _k, l, c in records if not c)
    skipped_hist = Counter()
    skipped_sizes = []
    for _off, kind, length, consumed in records:
        if not consumed and length:
            skipped_hist[kind] += length
            skipped_sizes.append(length)

    owned = continuation_owner(records)

    today = replay_today(stream, records)
    today_bytes = sum(e - s for s, e in today.reads)

    row = {
        "path": path,
        "file_size": os.path.getsize(path),
        "stream": name,
        "stream_len": len(stream),
        "globals_end": globals_end,
        "globals_records": len(records),
        "issue": issue,
        "header_bytes": header_bytes,
        "consumed_payload_bytes": consumed_payload,
        "closure_bytes": header_bytes + consumed_payload,
        "skipped_payload_bytes": skipped_payload,
        "closure_share_of_globals": (header_bytes + consumed_payload) / globals_end
        if globals_end
        else None,
        "mean_skipped_bytes_per_record": skipped_payload / len(records),
        "max_skipped_payload": max(skipped_sizes) if skipped_sizes else 0,
        "skipped_payloads_over_1k": sum(1 for s in skipped_sizes if s >= 1024),
        "skipped_bytes_in_payloads_over_1k": sum(s for s in skipped_sizes if s >= 1024),
        "skipped_payloads_over_4k": sum(1 for s in skipped_sizes if s >= 4096),
        "skipped_bytes_in_payloads_over_4k": sum(s for s in skipped_sizes if s >= 4096),
        "top_skipped": [
            [kind_name(k), v]
            for k, v in sorted(skipped_hist.items(), key=lambda kv: -kv[1])[:5]
        ],
        "continuation_bytes_by_owner": [
            [kind_name(k), v] for k, v in sorted(owned.items(), key=lambda kv: -kv[1])[:3]
        ],
        "today_reads": len(today.reads),
        "today_bytes": today_bytes,
        "today_overread_past_globals_end": max(0, today.filled - globals_end),
    }

    for label, skip_min, dense_mean, reset_on_skip in gates:
        cand = replay_candidate(
            stream, records, skip_min, dense_mean, reset_on_skip
        )
        cand_bytes = sum(e - s for s, e in cand.reads)
        row[f"{label}_reads"] = len(cand.reads)
        row[f"{label}_bytes"] = cand_bytes
        row[f"{label}_seeks"] = cand.seeks
        row[f"{label}_skips"] = cand.skips
        row[f"{label}_retained"] = cand.retained
        row[f"{label}_gate_fired"] = cand.gate_fired
        row[f"{label}_delta_reads"] = len(cand.reads) - len(today.reads)
        row[f"{label}_delta_bytes"] = cand_bytes - today_bytes
    return row


# ---------------------------------------------------------------------------


def fixtures(repo):
    files = []
    for dirpath, _dirnames, filenames in os.walk(os.path.join(repo, "test-data")):
        for fn in filenames:
            if fn.lower().endswith((".xls", ".xlt")):
                files.append(os.path.join(dirpath, fn))
    files.sort()
    return files


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--repo", default="/home/zhuhe/code/litchi")
    ap.add_argument("--out")
    args = ap.parse_args()

    # label, skip_min (per-record rule), dense_mean (0568's running mean),
    # reset_on_skip (the schedule's own hysteresis)
    gates = [
        # The shape change 0574 priced: skip every uninterpreted payload, no
        # per-record rule, every fill exact.
        ("ungated", 0, -1, True),
        # The per-record rule alone, with 0568's cumulative running mean as the
        # fill-size rule and no reset.  This is the naive port of 0568.
        ("mean1k", 1024, 1024, False),
        # The recommended shape: per-record rule plus target reset.
        ("g1k", 1024, 1 << 62, True),
        ("g2k", 2048, 1 << 62, True),
        ("g4k", 4096, 1 << 62, True),
        ("g8k", 8192, 1 << 62, True),
    ]

    rows = []
    for path in fixtures(args.repo):
        try:
            row = analyse(path, gates)
        except Exception as exc:  # noqa: BLE001
            row = {"path": path, "error": repr(exc)}
        if row is None:
            continue
        row["path"] = os.path.relpath(row["path"], args.repo)
        rows.append(row)

    ok = [r for r in rows if "error" not in r and r.get("globals_end")]

    def tot(key):
        return sum(r[key] for r in ok)

    summary = {
        "fixtures_scanned": len(rows),
        "fixtures_modelled": len(ok),
        "fixtures_refused": len(rows) - len(ok),
        "total_globals_bytes": tot("globals_end"),
        "total_header_bytes": tot("header_bytes"),
        "total_consumed_payload_bytes": tot("consumed_payload_bytes"),
        "total_closure_bytes": tot("closure_bytes"),
        "total_skipped_payload_bytes": tot("skipped_payload_bytes"),
        "closure_share": tot("closure_bytes") / tot("globals_end"),
        "total_today_reads": tot("today_reads"),
        "total_today_bytes": tot("today_bytes"),
        "fixtures_with_any_skipped_payload_over_1k": sum(
            1 for r in ok if r["skipped_payloads_over_1k"]
        ),
        "fixtures_with_any_skipped_payload_over_4k": sum(
            1 for r in ok if r["skipped_payloads_over_4k"]
        ),
        "fixtures_with_any_skipped_payload_over_8k": sum(
            1 for r in ok if r["max_skipped_payload"] >= 8192
        ),
        "fixtures_over_half_consumed": sum(
            1 for r in ok if r["closure_share_of_globals"] > 0.5
        ),
    }
    for label, _s, _d, _r in gates:
        summary[f"{label}_total_reads"] = tot(f"{label}_reads")
        summary[f"{label}_total_bytes"] = tot(f"{label}_bytes")
        summary[f"{label}_total_seeks"] = tot(f"{label}_seeks")
        summary[f"{label}_fixtures_gate_fired"] = sum(
            1 for r in ok if r[f"{label}_gate_fired"]
        )
        summary[f"{label}_fixtures_any_seek"] = sum(1 for r in ok if r[f"{label}_seeks"])
        summary[f"{label}_fixtures_more_reads"] = sum(
            1 for r in ok if r[f"{label}_delta_reads"] > 0
        )
        summary[f"{label}_fixtures_fewer_bytes"] = sum(
            1 for r in ok if r[f"{label}_delta_bytes"] < 0
        )
        # 116 ns per request (change 0564), 53.4 ns per KiB (change 0574).
        net = sum(
            r[f"{label}_delta_reads"] * 116.0 + r[f"{label}_delta_bytes"] / 1024.0 * 53.4
            for r in ok
        )
        summary[f"{label}_modelled_net_ns_file_source"] = round(net, 1)
        summary[f"{label}_fixtures_net_worse_file_source"] = sum(
            1
            for r in ok
            if r[f"{label}_delta_reads"] * 116.0
            + r[f"{label}_delta_bytes"] / 1024.0 * 53.4
            > 0
        )

    doc = {
        "summary": summary,
        "gates": [
            {
                "label": lbl,
                "skip_min_bytes": s,
                "dense_mean_bytes": d,
                "reset_target_on_skip": r,
            }
            for lbl, s, d, r in gates
        ],
        "consumed_kinds": {f"0x{k:04X}": v for k, v in CONSUMED.items()},
        "files": rows,
    }
    text = json.dumps(doc, indent=2, sort_keys=True)
    if args.out:
        with open(args.out, "w") as fh:
            fh.write(text + "\n")
    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
