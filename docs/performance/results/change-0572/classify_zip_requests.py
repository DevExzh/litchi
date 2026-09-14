#!/usr/bin/env python3
"""Classify recorded positional reads against a ZIP file's actual layout.

The probe binary records every ``(offset, length)`` a source-backed OOXML open
issues, in order.  This script parses the fixture's ZIP central directory and
local records directly from the bytes -- it does not ask the library where the
regions are -- and assigns every recorded byte to exactly one region:

  eocd                 the end-of-central-directory record (and any ZIP64
                       end-of-central-directory record/locator and comment)
  central_directory    the central directory block
  local_header         a member's fixed 30-byte local file header
  local_header_var     a member's variable header region (name + extra)
  payload              a member's stored/deflated payload bytes
  data_descriptor      a member's trailing data descriptor
  gap                  any byte belonging to none of the above

A request that spans several regions contributes bytes to each.  A request is
counted in ``requests_in_gap`` if any of its in-file bytes land in a gap.
Bytes past end-of-file are reported separately and are not gap bytes: a read
that asks for more than the file holds is a short read, not a misdirected one.

Only the Python standard library is used.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import struct
import sys
from pathlib import Path

EOCD_SIG = b"PK\x05\x06"
EOCD64_SIG = b"PK\x06\x06"
EOCD64_LOC_SIG = b"PK\x06\x07"
CD_SIG = b"PK\x01\x02"
LFH_SIG = b"PK\x03\x04"
DD_SIG = b"PK\x07\x08"

REGIONS = (
    "eocd",
    "central_directory",
    "local_header",
    "local_header_var",
    "payload",
    "data_descriptor",
    "gap",
)


def _u16(b, o):
    return struct.unpack_from("<H", b, o)[0]


def _u32(b, o):
    return struct.unpack_from("<I", b, o)[0]


def _u64(b, o):
    return struct.unpack_from("<Q", b, o)[0]


def parse_zip(path: Path):
    """Return (file_size, sha256, members, intervals) for one ZIP fixture.

    ``intervals`` is a sorted list of ``(start, end, region, member_index)``
    half-open ranges covering every byte the layout accounts for.
    """
    data = path.read_bytes()
    size = len(data)
    sha = hashlib.sha256(data).hexdigest()

    # Locate the end-of-central-directory record.
    tail_start = max(0, size - (0xFFFF + 22))
    idx = data.rfind(EOCD_SIG, tail_start)
    if idx < 0:
        raise ValueError(f"{path}: no end-of-central-directory record")
    eocd_off = idx
    comment_len = _u16(data, eocd_off + 20)
    eocd_end = eocd_off + 22 + comment_len
    entry_count = _u16(data, eocd_off + 10)
    cd_size = _u32(data, eocd_off + 12)
    cd_off = _u32(data, eocd_off + 16)

    eocd_region_start = eocd_off
    # ZIP64 locator sits immediately before the EOCD when present.
    if eocd_off >= 20 and data[eocd_off - 20 : eocd_off - 16] == EOCD64_LOC_SIG:
        loc_off = eocd_off - 20
        z64_off = _u64(data, loc_off + 8)
        if 0 <= z64_off < size and data[z64_off : z64_off + 4] == EOCD64_SIG:
            entry_count = _u64(data, z64_off + 32)
            cd_size = _u64(data, z64_off + 40)
            cd_off = _u64(data, z64_off + 48)
            eocd_region_start = z64_off
        else:
            eocd_region_start = loc_off

    # Walk the central directory.
    members = []
    p = cd_off
    for _ in range(entry_count):
        if data[p : p + 4] != CD_SIG:
            raise ValueError(f"{path}: bad central record at {p}")
        flags = _u16(data, p + 8)
        method = _u16(data, p + 10)
        csize = _u32(data, p + 20)
        usize = _u32(data, p + 24)
        n = _u16(data, p + 28)
        m = _u16(data, p + 30)
        k = _u16(data, p + 32)
        lho = _u32(data, p + 42)
        name = data[p + 46 : p + 46 + n].decode("utf-8", "replace")
        extra = data[p + 46 + n : p + 46 + n + m]
        # ZIP64 extended information overrides the 32-bit placeholders.
        if 0xFFFFFFFF in (csize, usize, lho):
            e = 0
            while e + 4 <= len(extra):
                hid = _u16(extra, e)
                hsz = _u16(extra, e + 2)
                if hid == 0x0001:
                    q = e + 4
                    if usize == 0xFFFFFFFF and q + 8 <= e + 4 + hsz:
                        usize = _u64(extra, q)
                        q += 8
                    if csize == 0xFFFFFFFF and q + 8 <= e + 4 + hsz:
                        csize = _u64(extra, q)
                        q += 8
                    if lho == 0xFFFFFFFF and q + 8 <= e + 4 + hsz:
                        lho = _u64(extra, q)
                        q += 8
                    break
                e += 4 + hsz
        members.append(
            {
                "index": len(members),
                "name": name,
                "method": method,
                "flags": flags,
                "has_descriptor": bool(flags & 0x08),
                "compressed_size": csize,
                "uncompressed_size": usize,
                "local_header_offset": lho,
                "central_record_offset": p,
            }
        )
        p += 46 + n + m + k
    cd_end = p

    intervals = []
    for mem in members:
        lho = mem["local_header_offset"]
        if data[lho : lho + 4] != LFH_SIG:
            raise ValueError(f"{path}: bad local header for {mem['name']} at {lho}")
        n = _u16(data, lho + 26)
        m = _u16(data, lho + 28)
        var_start = lho + 30
        data_start = var_start + n + m
        data_end = data_start + mem["compressed_size"]
        mem["local_name_len"] = n
        mem["local_extra_len"] = m
        mem["payload_offset"] = data_start
        mem["payload_end"] = data_end
        intervals.append((lho, var_start, "local_header", mem["index"]))
        if m + n:
            intervals.append((var_start, data_start, "local_header_var", mem["index"]))
        if data_end > data_start:
            intervals.append((data_start, data_end, "payload", mem["index"]))
        if mem["has_descriptor"]:
            sig = data[data_end : data_end + 4] == DD_SIG
            # ZIP64 descriptors carry 8-byte sizes.
            dd_len = (4 if sig else 0) + 4 + (16 if mem["uncompressed_size"] > 0xFFFFFFFF or mem["compressed_size"] > 0xFFFFFFFF else 8)
            mem["descriptor_offset"] = data_end
            mem["descriptor_len"] = dd_len
            intervals.append((data_end, data_end + dd_len, "data_descriptor", mem["index"]))
        else:
            mem["descriptor_offset"] = None
            mem["descriptor_len"] = 0

    intervals.append((cd_off, cd_end, "central_directory", None))
    intervals.append((eocd_region_start, max(eocd_end, size), "eocd", None))
    intervals.sort()

    # Verify the map does not overlap itself; overlapping regions would make
    # per-region byte totals meaningless.
    prev_end = 0
    for start, end, region, _mi in intervals:
        if start < prev_end:
            raise ValueError(f"{path}: overlapping layout regions near {start}")
        prev_end = max(prev_end, end)

    return {
        "path": str(path),
        "size": size,
        "sha256": sha,
        "member_count": len(members),
        "descriptor_members": sum(1 for m in members if m["has_descriptor"]),
        "central_directory_offset": cd_off,
        "central_directory_size": cd_end - cd_off,
        "eocd_offset": eocd_region_start,
        "members": members,
    }, intervals


def build_lookup(intervals, size):
    """Return a byte-classifier closure over the sorted interval list."""
    import bisect

    starts = [i[0] for i in intervals]

    def classify(start, end):
        """Yield ``(region, member_index, byte_count)`` for ``[start, end)``."""
        out = []
        pos = start
        i = bisect.bisect_right(starts, pos) - 1
        if i < 0:
            i = 0
        while pos < end:
            # Advance to the first interval that can contain or follow pos.
            while i < len(intervals) and intervals[i][1] <= pos:
                i += 1
            if i >= len(intervals):
                out.append(("gap", None, end - pos))
                break
            s, e, region, mi = intervals[i]
            if pos < s:
                take = min(s, end) - pos
                out.append(("gap", None, take))
                pos += take
                continue
            take = min(e, end) - pos
            out.append((region, mi, take))
            pos += take
        return out

    return classify


def classify_requests(layout, intervals, requests):
    """Classify one ordered request list.  ``requests`` is a list of dicts."""
    size = layout["size"]
    classify = build_lookup(intervals, size)
    header_offsets = {m["local_header_offset"] for m in layout["members"]}

    region_requests = {r: 0 for r in REGIONS}
    role_requests = {r: 0 for r in REGIONS}
    role_requests["local_record_probe"] = 0
    region_bytes = {r: 0 for r in REGIONS}
    beyond_eof_bytes = 0
    requests_beyond_eof = 0
    requests_in_gap = 0
    gap_examples = []
    classified = []

    for n, req in enumerate(requests):
        off = req["offset"]
        length = req["length"]
        end = off + length
        in_file_end = min(end, size)
        parts = classify(off, in_file_end) if in_file_end > off else []
        past = end - in_file_end
        if past > 0:
            beyond_eof_bytes += past
            requests_beyond_eof += 1
        agg = {}
        for region, mi, count in parts:
            agg[region] = agg.get(region, 0) + count
            region_bytes[region] += count
        if agg.get("gap"):
            requests_in_gap += 1
            if len(gap_examples) < 20:
                gap_examples.append({"index": n, "offset": off, "length": length, "gap_bytes": agg["gap"]})
        # A request's primary region is the one holding most of its in-file bytes.
        primary = max(agg.items(), key=lambda kv: (kv[1], kv[0]))[0] if agg else "eocd"
        for region in agg:
            pass
        region_requests[primary] += 1
        probe = off in header_offsets
        role_requests["local_record_probe" if probe else primary] += 1
        members = sorted({mi for _r, mi, _c in parts if mi is not None})
        classified.append(
            {
                "i": n,
                "offset": off,
                "length": length,
                "returned": req.get("returned"),
                "primary_region": primary,
                "local_record_probe": probe,
                "regions": agg,
                "members": members,
            }
        )

    return {
        "request_count": len(requests),
        "requested_bytes": sum(r["length"] for r in requests),
        "returned_bytes": sum(r.get("returned") or 0 for r in requests),
        "requests_by_primary_region": region_requests,
        "requests_by_role": role_requests,
        "bytes_by_region": region_bytes,
        "requests_in_gap": requests_in_gap,
        "gap_bytes": region_bytes["gap"],
        "gap_examples": gap_examples,
        "requests_beyond_eof": requests_beyond_eof,
        "bytes_beyond_eof": beyond_eof_bytes,
    }, classified


def run_length_summary(classified):
    """Collapse an ordered classified list into a compact run-length view."""
    runs = []
    for c in classified:
        key = (c["primary_region"],)
        if runs and runs[-1]["region"] == c["primary_region"]:
            runs[-1]["count"] += 1
            runs[-1]["bytes"] += c["length"]
            runs[-1]["last_offset"] = c["offset"]
        else:
            runs.append(
                {
                    "region": c["primary_region"],
                    "count": 1,
                    "bytes": c["length"],
                    "first_offset": c["offset"],
                    "last_offset": c["offset"],
                }
            )
    return runs


FRAMING = frozenset(("local_header", "local_header_var", "data_descriptor"))


def find_proof_sweep(layout, classified):
    """Locate the strict-layout proof: the contiguous local-record sweep.

    Change 0567 named the proof as a whole-archive sweep that visits every
    member in layout order before the scenario's own payload read. Measured, a
    sweep request is one of two shapes:

      * a **probe** -- a request beginning exactly at a member's local header
        offset, sized to cover the local record's framing (and, incidentally,
        payload bytes the proof does not use); or
      * a **framing read** -- a request whose bytes lie entirely inside local
        header, variable header or data-descriptor regions, which is how a
        descriptor-bearing member and a clamped final member are finished.

    The proof is the longest run of consecutive requests of those two shapes.
    A request that reads member payload and does not start at a local header
    offset ends the run, which is what separates the sweep from the
    scenario's own reads.
    """
    by_offset = sorted(layout["members"], key=lambda m: m["local_header_offset"])
    order = {m["local_header_offset"]: pos for pos, m in enumerate(by_offset)}

    def shape(c):
        if c["offset"] in order:
            return "probe"
        if c["regions"] and all(r in FRAMING for r in c["regions"]):
            return "framing"
        return None

    runs = []
    i = 0
    n = len(classified)
    while i < n:
        if shape(classified[i]) is None:
            i += 1
            continue
        j = i
        while j < n and shape(classified[j]) is not None:
            j += 1
        window = classified[i:j]
        probes = [c for c in window if shape(c) == "probe"]
        positions = [order[c["offset"]] for c in probes]
        runs.append(
            {
                "start_index": i,
                "end_index": j - 1,
                "request_count": j - i,
                "probe_count": len(probes),
                "framing_count": (j - i) - len(probes),
                "descriptor_count": sum(
                    1 for c in window if set(c["regions"]) == {"data_descriptor"}
                ),
                "distinct_members_probed": len(set(positions)),
                "layout_order_ascending": all(b > a for a, b in zip(positions, positions[1:])),
                "requested_bytes": sum(c["length"] for c in window),
                "probe_lengths": sorted({c["length"] for c in probes}),
            }
        )
        i = j

    longest = max(runs, key=lambda r: r["request_count"]) if runs else None
    after = None
    if longest is not None and longest["end_index"] + 1 < n:
        nxt = classified[longest["end_index"] + 1]
        after = {
            "index": nxt["i"],
            "offset": nxt["offset"],
            "length": nxt["length"],
            "primary_region": nxt["primary_region"],
            "members": nxt["members"],
        }
    return {
        "runs": runs,
        "longest": longest,
        "first_request_after_longest_run": after,
        "member_count": layout["member_count"],
        "descriptor_members": layout["descriptor_members"],
    }


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--capture", required=True, help="probe capture JSON")
    ap.add_argument("--out", required=True, help="analysed output JSON")
    ap.add_argument("--repo", default=".", help="repository root for fixture paths")
    ap.add_argument(
        "--raw-sequence-for",
        action="append",
        default=[],
        help="scenario:policy pattern whose full raw sequence is retained",
    )
    args = ap.parse_args(argv)

    capture = json.loads(Path(args.capture).read_text())
    repo = Path(args.repo)

    layouts = {}
    intervals_by_fixture = {}
    for fixture in capture["corpus"]:
        p = repo / fixture["path"]
        layout, intervals = parse_zip(p)
        layout["fixture_id"] = fixture["id"]
        if layout["sha256"] != fixture["sha256"]:
            raise SystemExit(
                f"fixture {fixture['id']} hash changed between capture and analysis"
            )
        layouts[fixture["id"]] = layout
        intervals_by_fixture[fixture["id"]] = intervals

    results = []
    determinism = []
    for run in capture["runs"]:
        fid = run["fixture"]
        layout = layouts[fid]
        intervals = intervals_by_fixture[fid]
        repeats = run["repeats"]
        # Determinism: the ordered (offset, length) sequence across repeats.
        seqs = [[(r["offset"], r["length"]) for r in rep["requests"]] for rep in repeats]
        identical = all(s == seqs[0] for s in seqs)
        digests = [
            hashlib.sha256(
                ";".join(f"{o}+{l}" for o, l in s).encode()
            ).hexdigest()[:16]
            for s in seqs
        ]
        determinism.append(
            {
                "scenario": run["scenario"],
                "fixture": fid,
                "policy": run["policy"],
                "route": run["route"],
                "transport": run["transport"],
                "repeats": len(seqs),
                "identical_across_repeats": identical,
                "sequence_digests": digests,
            }
        )
        summary, classified = classify_requests(layout, intervals, repeats[0]["requests"])
        proof = find_proof_sweep(layout, classified)
        entry = {
            "scenario": run["scenario"],
            "fixture": fid,
            "policy": run["policy"],
            "route": run["route"],
            "transport": run["transport"],
            "identical_across_repeats": identical,
            "sequence_digest": digests[0],
            "summary": summary,
            "run_length": run_length_summary(classified),
            "longest_header_sweep": proof,
            "elapsed_ns": sorted(rep["elapsed_ns"] for rep in repeats),
        }
        want_raw = any(
            pat.split(":")[0] in ("*", run["scenario"])
            and pat.split(":")[1] in ("*", run["policy"])
            for pat in args.raw_sequence_for
            if ":" in pat
        )
        if want_raw:
            entry["raw_requests"] = classified
        results.append(entry)

    out = {
        "schema_version": 1,
        "record_kind": "litchi-perf-0572-request-attribution",
        "capture": {k: v for k, v in capture.items() if k not in ("runs",)},
        "layouts": {
            fid: {
                k: v
                for k, v in layout.items()
                if k != "members"
            }
            | {
                "members": [
                    {
                        "index": m["index"],
                        "name": m["name"],
                        "local_header_offset": m["local_header_offset"],
                        "payload_offset": m["payload_offset"],
                        "compressed_size": m["compressed_size"],
                        "has_descriptor": m["has_descriptor"],
                        "descriptor_len": m["descriptor_len"],
                    }
                    for m in layout["members"]
                ]
            }
            for fid, layout in layouts.items()
        },
        "determinism": determinism,
        "results": results,
    }
    Path(args.out).write_text(json.dumps(out, indent=1, sort_keys=False) + "\n")

    gaps = sum(r["summary"]["requests_in_gap"] for r in results)
    nondet = [d for d in determinism if not d["identical_across_repeats"]]
    print(f"wrote {args.out}")
    print(f"arms={len(results)} requests_in_gap_total={gaps} non_deterministic_arms={len(nondet)}")
    for d in nondet:
        print("  NON-DETERMINISTIC:", d["scenario"], d["fixture"], d["policy"], d["transport"], d["sequence_digests"])
    return 0


if __name__ == "__main__":
    sys.exit(main())
