#!/usr/bin/env python3
"""Create the immutable 0895 heaptrack plan without launching anything."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from typing import Any

from common import (  # type: ignore[import-not-found]
    BASE_PLAN_PATH,
    CPU_AFFINITY,
    DEFAULT_PRINT_ARGS,
    PACKET,
    PRIMARY_CASE,
    ROLES,
    SAMPLES,
    SCHEMA,
    WARMUP,
    assert_rejected_0894,
    binary,
    descriptor,
    entry,
    evidence_descriptors,
    exact_0894_helper_hashes,
    absolute_executable,
    load_base_plan,
    require,
    same_descriptor,
    utc_now,
    write_json,
)


def _parse_args(raw: str | None) -> list[str]:
    if raw is None:
        return list(DEFAULT_PRINT_ARGS)
    value = json.loads(raw)
    require(isinstance(value, list) and all(isinstance(item, str) for item in value), "--print-args-json must be a JSON string array")
    require("{profile}" in value and "{stacks}" in value and "{histogram}" in value, "print argv must bind profile, stacks and histogram")
    return value


def _tool(path: Path, version: str, label: str) -> dict[str, Any]:
    actual = descriptor(path)
    require(actual["executable"] is True, f"{label} is not executable")
    require(version.strip(), f"{label} version is empty")
    return {"path": str(path.resolve()), "identity": actual, "version": version}


def freeze(
    base_plan_path: Path,
    output: Path,
    heaptrack: Path,
    heaptrack_version: str,
    heaptrack_print: Path,
    heaptrack_print_version: str,
    print_args: list[str],
) -> dict[str, Any]:
    base_plan_path = base_plan_path.resolve()
    base_plan = load_base_plan(base_plan_path)
    require(base_plan_path == BASE_PLAN_PATH.resolve(), "0895 must bind the retained 0894 plan")
    assert_rejected_0894(base_plan)
    helpers = exact_0894_helper_hashes(base_plan)
    evidence = evidence_descriptors(base_plan)
    packet_helpers = {
        name: descriptor(PACKET / name)
        for name in ("common.py", "freeze.py", "capture.py", "export.py", "reduce.py", "verify.py", "attribution.py")
    }
    baseline_binary = binary(base_plan, "baseline")
    candidate_binary = binary(base_plan, "candidate")
    same_descriptor(base_plan["binaries"]["before"]["native"], baseline_binary["identity"], "0894 baseline binary")
    same_descriptor(base_plan["binaries"]["after"]["native"], candidate_binary["identity"], "0894 candidate binary")

    rows = [
        {"ordinal": ordinal, "role": role, "case": PRIMARY_CASE, "warmup": WARMUP, "samples": SAMPLES, "pair": f"profile-{ordinal:02d}"}
        for ordinal, role in enumerate(ROLES)
    ]
    require([row["role"] for row in rows] == list(ROLES), "0895 role schedule changed")

    plan = {
        "schema": SCHEMA,
        "status": "frozen",
        "frozen_utc": utc_now(),
        "base_plan": descriptor(base_plan_path),
        "base_plan_schema": base_plan["schema"],
        "base_commit": base_plan["base_commit"],
        "candidate_adoption": False,
        "production_adoption": False,
        "cpu_affinity": CPU_AFFINITY,
        "cache": "warm",
        "case": PRIMARY_CASE,
        "xlsx_shape": "medium",
        "warmup": WARMUP,
        "samples": SAMPLES,
        "rows": rows,
        "total_rows": len(rows),
        "total_measured_samples": len(rows) * SAMPLES,
        "fresh_child": True,
        "fail_fast": True,
        "guard_contract": base_plan["guard_contract"],
        "binaries": {"baseline": baseline_binary, "candidate": candidate_binary},
        "evidence": evidence,
        # Keep descriptors and the seven hash values together.  The former
        # makes path/size/mode changes visible; the latter is a compact bind
        # for reviewers and downstream receipts.
        "helper_hashes": helpers,
        "helper_sha256": {name: item["sha256"] for name, item in helpers.items()},
        "packet_helpers": packet_helpers,
        "tools": {
            "taskset": "/usr/bin/taskset",
            "taskset_identity": descriptor(Path("/usr/bin/taskset")),
            "heaptrack": _tool(heaptrack, heaptrack_version, "heaptrack"),
            "heaptrack_print": _tool(heaptrack_print, heaptrack_print_version, "heaptrack_print"),
        },
        "record_argv_template": [
            "/usr/bin/taskset", "-c", CPU_AFFINITY,
            "{heaptrack}", "--record-only", "-o", "{profile}", "{binary}",
            "--case", PRIMARY_CASE, "--xlsx-shape", "medium", "--warmup", "0",
            "--samples", str(SAMPLES), "--filesystem-cache", "warm", "--json",
            "{report}", "--producer-evidence", "{producer}",
        ],
        "print_argv_template": print_args,
        # This host's pinned heaptrack emits zstd-compressed traces even when
        # the output basename has no suffix.  The suffix is frozen from the
        # tool contract so capture/verify cannot silently accept a different
        # profile format.
        "profile_suffix": ".zst",
        "owner_categories": ["legacycodec", "mce_stream", "selected_worksheet", "x14ac", "other"],
        "inclusive_categories_are_diagnostic": True,
        "performance_claim": None,
    }
    require(Path(plan["tools"]["taskset"]).is_file(), "taskset path is missing")
    require(len(plan["helper_hashes"]) == 7, "0894 helper set is not exactly seven entries")
    require(not output.exists() and not output.is_symlink(), f"plan output already exists or is symlinked: {output}")
    write_json(output, plan)
    return plan


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--base-plan", type=Path, default=BASE_PLAN_PATH)
    parser.add_argument("--output", type=Path, default=PACKET / "plan.json")
    parser.add_argument("--heaptrack", type=Path, required=True)
    parser.add_argument("--heaptrack-version", required=True)
    parser.add_argument("--heaptrack-print", type=Path, required=True)
    parser.add_argument("--heaptrack-print-version", required=True)
    parser.add_argument("--print-args-json")
    args = parser.parse_args()
    heaptrack = absolute_executable(str(args.heaptrack.resolve()), "heaptrack")
    heaptrack_print = absolute_executable(str(args.heaptrack_print.resolve()), "heaptrack_print")
    plan = freeze(
        args.base_plan,
        args.output.resolve(),
        heaptrack,
        args.heaptrack_version,
        heaptrack_print,
        args.heaptrack_print_version,
        _parse_args(args.print_args_json),
    )
    print(json.dumps(plan, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
