#!/usr/bin/env python3
"""Validate and summarize one or more diagnostic 0704 MCE trace files.

The report describes repeated byte-level observations.  It deliberately does
not contain timing, allocation, RSS, or throughput conclusions.  Pointer
fields are retained as diagnostics only: a pointer match is never treated as
proof of immutable ownership.
"""
from __future__ import annotations

import argparse
import json
import re
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

MCE_PREFIX = "LITCHI0704_MCE "
BOUNDARY_PREFIX = "LITCHI0704_BOUNDARY "
HEX64 = re.compile(r"^[0-9a-f]{64}$")
POINTER = re.compile(r"^0x[0-9a-f]+$")


def fields(line: str, prefix: str) -> dict[str, str]:
    result: dict[str, str] = {}
    for token in line[len(prefix) :].split():
        key, separator, value = token.partition("=")
        if not separator or key in result:
            raise ValueError(f"malformed trace token {token!r}")
        result[key] = value
    return result


def integer(row: dict[str, str], key: str) -> int:
    try:
        return int(row[key], 0) if key.endswith("ptr") else int(row[key])
    except (KeyError, ValueError) as error:
        raise ValueError(f"invalid integer field {key!r}: {row.get(key)!r}") from error


def stage(phase: str) -> str:
    return phase.rsplit(".", 1)[-1]


def parse_file(path: Path) -> tuple[list[dict[str, Any]], Counter[str]]:
    records: list[dict[str, Any]] = []
    boundaries: Counter[str] = Counter()
    for line_number, raw_line in enumerate(path.read_text().splitlines(), 1):
        line = raw_line.strip()
        if line.startswith(BOUNDARY_PREFIX):
            row = fields(line, BOUNDARY_PREFIX)
            if row.get("event") != "begin" or "phase" not in row:
                raise ValueError(f"{path}:{line_number}: invalid boundary row")
            boundaries[row["phase"]] += 1
            continue
        if not line.startswith(MCE_PREFIX):
            continue
        row = fields(line, MCE_PREFIX)
        required = {
            "call",
            "phase",
            "profile",
            "raw_ptr",
            "raw_len",
            "raw_sha256",
            "output_ptr",
            "output_len",
            "output_capacity",
            "ownership",
            "output_sha256",
            "limits_input",
            "limits_output",
            "limits_depth",
            "limits_bindings",
            "limits_directives",
            "limits_choices",
            "status",
        }
        if set(row) != required:
            raise ValueError(
                f"{path}:{line_number}: field census differs: "
                f"missing={sorted(required - set(row))} extra={sorted(set(row) - required)}"
            )
        call = integer(row, "call")
        raw_ptr = integer(row, "raw_ptr")
        output_ptr = integer(row, "output_ptr")
        raw_len = integer(row, "raw_len")
        output_len = integer(row, "output_len")
        output_capacity = integer(row, "output_capacity")
        if raw_ptr < 0 or output_ptr < 0 or raw_len < 0 or output_len < 0 or output_capacity < 0:
            raise ValueError(f"{path}:{line_number}: negative trace field")
        if not POINTER.fullmatch(row["raw_ptr"]) or not POINTER.fullmatch(row["output_ptr"]):
            raise ValueError(f"{path}:{line_number}: pointer is not hexadecimal")
        if not HEX64.fullmatch(row["raw_sha256"]):
            raise ValueError(f"{path}:{line_number}: raw digest is not SHA-256")
        if row["status"] == "ok":
            if row["ownership"] not in {"borrowed", "owned"}:
                raise ValueError(f"{path}:{line_number}: invalid successful ownership")
            if not HEX64.fullmatch(row["output_sha256"]):
                raise ValueError(f"{path}:{line_number}: output digest is not SHA-256")
            if row["ownership"] == "borrowed" and output_capacity != 0:
                raise ValueError(f"{path}:{line_number}: borrowed output has capacity")
            if row["ownership"] == "owned" and output_capacity < output_len:
                raise ValueError(f"{path}:{line_number}: owned capacity is below length")
        elif row["status"] == "error":
            if row["ownership"] != "error" or output_ptr != 0 or output_len != 0:
                raise ValueError(f"{path}:{line_number}: malformed error ownership")
            if row["output_sha256"] != "none" or output_capacity != 0:
                raise ValueError(f"{path}:{line_number}: malformed error output")
        else:
            raise ValueError(f"{path}:{line_number}: invalid status {row['status']!r}")
        limits = tuple(row[key] for key in (
            "profile",
            "limits_input",
            "limits_output",
            "limits_depth",
            "limits_bindings",
            "limits_directives",
            "limits_choices",
        ))
        records.append(
            {
                "source_file": str(path),
                "line": line_number,
                "call": call,
                "phase": row["phase"],
                "stage": stage(row["phase"]),
                "profile": row["profile"],
                "raw_ptr": row["raw_ptr"],
                "raw_len": raw_len,
                "raw_sha256": row["raw_sha256"],
                "output_ptr": row["output_ptr"],
                "output_len": output_len,
                "output_capacity": output_capacity,
                "ownership": row["ownership"],
                "output_sha256": row["output_sha256"],
                "status": row["status"],
                "options": limits,
            }
        )
    if records:
        calls = [row["call"] for row in records]
        if calls != list(range(len(calls))):
            raise ValueError(f"{path}: call sequence is not zero-based and contiguous: {calls}")
    return records, boundaries


def group_summary(records: list[dict[str, Any]], keys: tuple[str, ...]) -> list[dict[str, Any]]:
    groups: dict[tuple[Any, ...], list[dict[str, Any]]] = defaultdict(list)
    for row in records:
        groups[tuple(row[key] for key in keys)].append(row)
    result: list[dict[str, Any]] = []
    for key, rows in sorted(groups.items(), key=lambda item: (item[0], item[1][0]["call"])):
        result.append(
            {
                "key": list(key),
                "calls": [row["call"] for row in rows],
                "phases": [row["phase"] for row in rows],
                "stages": sorted({row["stage"] for row in rows}),
                "raw_pointers": sorted({row["raw_ptr"] for row in rows}),
                "output_pointers": sorted({row["output_ptr"] for row in rows}),
                "raw_len": rows[0]["raw_len"],
                "output_len": rows[0]["output_len"],
                "output_capacity_max": max(row["output_capacity"] for row in rows),
            }
        )
    return result


def phase_pair_groups(
    records: list[dict[str, Any]], keys: tuple[str, ...], left: str, right: str
) -> list[dict[str, Any]]:
    return [
        group
        for group in group_summary(records, keys)
        if left in group["stages"] and right in group["stages"]
    ]


def summarize(path: Path) -> dict[str, Any]:
    records, boundaries = parse_file(path)
    if not records:
        raise ValueError(f"{path}: no LITCHI0704_MCE rows found")
    options = sorted({row["options"] for row in records})
    raw_groups = group_summary(records, ("raw_sha256", "raw_len", "options"))
    projection_groups = group_summary(
        records, ("output_sha256", "output_len", "ownership", "options")
    )
    pair_groups = group_summary(
        records,
        ("raw_sha256", "raw_len", "output_sha256", "output_len", "ownership", "options"),
    )
    repeated_raw = [group for group in raw_groups if len(group["calls"]) > 1]
    repeated_projection = [group for group in projection_groups if len(group["calls"]) > 1]
    repeated_pair = [group for group in pair_groups if len(group["calls"]) > 1]
    cross_phase_capture_commit = phase_pair_groups(
        records,
        ("raw_sha256", "raw_len", "output_sha256", "output_len", "ownership", "options"),
        "capture",
        "commit",
    )
    cross_phase_capture_apply = phase_pair_groups(
        records,
        ("raw_sha256", "raw_len", "output_sha256", "output_len", "ownership", "options"),
        "capture",
        "apply",
    )
    unique_raw = {(row["raw_sha256"], row["raw_len"], row["options"]): row for row in records}
    unique_projection = {
        (row["output_sha256"], row["output_len"], row["ownership"], row["options"])
        for row in records
        if row["status"] == "ok"
    }
    owned_projection_capacity = sum(
        group["output_capacity_max"]
        for group in projection_groups
        if group["key"][2] == "owned"
    )
    owned_projection_length = sum(
        group["output_len"]
        for group in projection_groups
        if group["key"][2] == "owned"
    )
    borrowed_source_length = sum(
        row["raw_len"]
        for row in unique_raw.values()
        if row["status"] == "ok" and row["ownership"] == "borrowed"
    )
    return {
        "schema": "litchi-0704-mce-trace-summary-v1",
        "diagnostic_only": True,
        "timing_evidence": False,
        "trace_file": str(path),
        "calls": len(records),
        "status_counts": dict(Counter(row["status"] for row in records)),
        "ownership_counts": dict(Counter(row["ownership"] for row in records)),
        "phase_counts": dict(Counter(row["phase"] for row in records)),
        "stage_counts": dict(Counter(row["stage"] for row in records)),
        "boundary_counts": dict(boundaries),
        "options_profiles": [list(option) for option in options],
        "raw_unique_keys": len(unique_raw),
        "projection_unique_keys": len(unique_projection),
        "repeated_raw_calls_beyond_first": sum(len(group["calls"]) - 1 for group in repeated_raw),
        "repeated_projection_calls_beyond_first": sum(
            len(group["calls"]) - 1 for group in repeated_projection
        ),
        "repeated_exact_pair_calls_beyond_first": sum(
            len(group["calls"]) - 1 for group in repeated_pair
        ),
        "cross_phase_capture_commit_groups": len(cross_phase_capture_commit),
        "cross_phase_capture_apply_groups": len(cross_phase_capture_apply),
        "pointer_reuse_is_diagnostic_only": True,
        "same_raw_groups": repeated_raw,
        "same_projection_groups": repeated_projection,
        "same_raw_and_projection_groups": repeated_pair,
        "same_pair_capture_commit_groups": cross_phase_capture_commit,
        "same_pair_capture_apply_groups": cross_phase_capture_apply,
        "logical_retention_estimate": {
            "unique_raw_payload_bytes_if_source_owner_must_be_retained": sum(
                row["raw_len"] for row in unique_raw.values()
            ),
            "unique_borrowed_source_bytes": borrowed_source_length,
            "unique_owned_projection_payload_bytes": owned_projection_length,
            "unique_owned_projection_capacity_bytes": owned_projection_capacity,
            "interpretation": (
                "These are payload/capacity sums for one representative per exact "
                "digest/profile key. They are lower-bound accounting figures for a "
                "hypothetical bounded cache, not a measured allocator or RSS bound. "
                "Borrowed projections require a separately retained source owner "
                "only if the source package does not already keep it alive."
            ),
        },
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("trace", nargs="+", type=Path, help="stderr trace files")
    parser.add_argument("--json-out", type=Path)
    args = parser.parse_args()
    summaries = [summarize(path) for path in args.trace]
    result: dict[str, Any] = {
        "schema": "litchi-0704-mce-trace-report-v1",
        "diagnostic_only": True,
        "timing_evidence": False,
        "runs": summaries,
    }
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    if args.json_out:
        args.json_out.write_text(encoded)
    else:
        print(encoded, end="")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError) as error:
        raise SystemExit(f"trace validation failed: {error}")
