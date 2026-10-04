#!/usr/bin/env python3
"""Offline attribution of the retained 0895 baseline allocation stacks.

This helper intentionally reads retained receipts and text exports only.  It
uses read-only git identity queries (`show`, `rev-parse`, and an ancestry
check) to bind the source state; it does not invoke Cargo, the producer
binary, heaptrack, or any workload.  The stack files are whole-process
allocation-hook evidence; this report does not turn them into a latency, RSS,
or native timing claim.
"""

from __future__ import annotations

import argparse
from collections import Counter, defaultdict
import hashlib
import json
from pathlib import Path
import re
import stat
import subprocess
import sys
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
RETAINED = PACKET.parent / "change-0895"
BASE_COMMIT = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
OBSERVED_EVALUATION_HEAD = "c6d7ba22f525fbf82bbe7a7bcc1beb495169763b"

ROWS = (
    ("00-baseline", 0),
    ("03-baseline", 3),
)
EXPECTED_SAMPLES = 20
EXPECTED_CLOSURE_PER_SAMPLE = 25_933

CLOSURE_MARKER = (
    "litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan::"
    "_$u7b$$u7b$closure$u7d$$u7d$::"
)
GUARD_MARKER = (
    "litchi_perf_baseline::producer_shape::"
    "force_full_cell_worksheet_tolerance_reports::"
)
RUNNER_MARKER = "litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan::"
CODEC_MARKER = "litchi_ooxml_common::mce::codec::"
START_MARKER = "litchi_ooxml_common::mce::codec::start::"
UNRESOLVED = re.compile(r"(?:^0x[0-9a-f]+$|\?\?|<unknown>)", re.IGNORECASE)
INTEGER = re.compile(r"^[0-9]+$")

# These are deliberately presence rules.  A flamegraph frame is not a DWARF
# source-line attribution, and the retained heaptrack export has no reliable
# temporal or inline-call boundary.  Precedence makes the four start buckets
# disjoint when a stack contains both a high-level operation and allocator
# frames beneath it.
OWNED_NAME_TOKENS = (
    "ToOwned8to_owned",
    "alloc..borrow..ToOwned$u20$for$u20$$u5b$T$u5d$",
)
QUICK_XML_DUPLICATE_TOKEN = (
    "quick_xml::events::attributes::IterState::check_for_duplicates::",
)
RAW_RESERVE_TOKENS = (
    "litchi_ooxml_common::mce::codec::reserve_amortized::",
    "alloc::raw_vec::RawVecInner$LT$A$GT$::grow_",
    "alloc::raw_vec::RawVecInner$LT$A$GT$::try_reserve",
    "alloc::raw_vec::RawVecInner$LT$A$GT$::reserve",
    "alloc::raw_vec::RawVec$LT$T$C$A$GT$::try_reserve",
    "alloc::vec::Vec$LT$T$C$A$GT$::try_reserve",
)

SOURCE_FILES = {
    "mce_codec": ROOT / "crates/litchi-ooxml-common/src/mce/codec.rs",
    "mce_stream": ROOT / "crates/litchi-ooxml-common/src/mce/stream.rs",
    "producer_shape": ROOT / "tools/perf-baseline/src/producer_shape.rs",
    "perf_runner": ROOT / "tools/perf-baseline/src/lib.rs",
    "guard_tests": ROOT / "tools/perf-baseline/src/producer_shape/full_cell_guard_tests.rs",
}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    info = path.stat()
    require(stat.S_ISREG(info.st_mode), f"not a regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": sha_file(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def git(*arguments: str) -> bytes:
    result = subprocess.run(
        ["git", *arguments],
        cwd=ROOT,
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    return result.stdout


def git_text(*arguments: str) -> str:
    return git(*arguments).decode("utf-8").strip()


def is_ancestor(ancestor: str, descendant: str) -> bool:
    result = subprocess.run(
        ["git", "merge-base", "--is-ancestor", ancestor, descendant],
        cwd=ROOT,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    return result.returncode == 0


def source_descriptor(name: str, path: Path) -> dict[str, Any]:
    current = descriptor(path)
    base_bytes = git("show", f"{BASE_COMMIT}:{path.relative_to(ROOT)}")
    base_hash = sha_bytes(base_bytes)
    observed_bytes = git("show", f"{OBSERVED_EVALUATION_HEAD}:{path.relative_to(ROOT)}")
    observed_hash = sha_bytes(observed_bytes)
    return {
        "current": current,
        "base_commit": BASE_COMMIT,
        "base_blob_sha256": base_hash,
        "matches_base_blob": current["sha256"] == base_hash,
        "observed_evaluation_head": OBSERVED_EVALUATION_HEAD,
        "observed_evaluation_blob_sha256": observed_hash,
        "matches_observed_evaluation_blob": current["sha256"] == observed_hash,
    }


def source_line(path: Path, line_number: int) -> dict[str, Any]:
    lines = path.read_text(encoding="utf-8").splitlines()
    require(1 <= line_number <= len(lines), f"source line is outside file: {path}:{line_number}")
    text = lines[line_number - 1]
    return {
        "path": str(path.resolve()),
        "line": line_number,
        "text": text,
        "line_sha256": sha_bytes((text + "\n").encode("utf-8")),
    }


def expected_rows() -> list[dict[str, Any]]:
    return [
        {
            "case": "xlsx_producer_medium_source_full_cell_scan",
            "ordinal": ordinal,
            "pair": f"profile-{ordinal:02d}",
            "role": role,
            "samples": EXPECTED_SAMPLES,
            "warmup": 0,
        }
        for role, ordinal in (
            ("baseline", 0),
            ("candidate", 1),
            ("candidate", 2),
            ("baseline", 3),
        )
    ]


def validate_receipts() -> tuple[dict[str, Any], dict[str, Any]]:
    plan_path = RETAINED / "plan.json"
    capture_path = RETAINED / "capture-v1/capture.json"
    export_path = RETAINED / "export-v1/export.json"
    verification_path = RETAINED / "verification.json"
    attribution_path = RETAINED / "attribution.json"
    reduction_path = RETAINED / "reduction.json"
    for path in (
        plan_path,
        capture_path,
        export_path,
        verification_path,
        attribution_path,
        reduction_path,
    ):
        descriptor(path)

    plan = read_json(plan_path)
    require(plan.get("schema") == "litchi.xlsx-selected-metadata-heaptrack-0895.v1", "0895 plan schema changed")
    require(plan.get("status") == "frozen", "0895 plan is not frozen")
    require(plan.get("base_commit") == BASE_COMMIT, "0895 base commit changed")
    require(plan.get("case") == "xlsx_producer_medium_source_full_cell_scan", "0895 case changed")
    require(plan.get("cache") == "warm" and plan.get("cpu_affinity") == "12", "0895 environment changed")
    require(plan.get("samples") == EXPECTED_SAMPLES and plan.get("warmup") == 0, "0895 sample policy changed")
    require(plan.get("rows") == expected_rows(), "0895 plan rows changed")
    require(plan.get("total_rows") == 4 and plan.get("total_measured_samples") == 80, "0895 plan totals changed")

    capture = read_json(capture_path)
    require(capture.get("schema") == "litchi.xlsx-selected-metadata-heaptrack-capture-0895.v1", "capture schema changed")
    require(capture.get("status") == "pass", "capture is not passing")
    require(capture.get("children") == 4 and capture.get("measured_samples") == 80, "capture schedule changed")
    require(capture.get("rows") == [{"path": f"runs/{i:02d}-{role}/receipt.json", "row": row} for i, (role, _) in enumerate((("baseline", 0), ("candidate", 1), ("candidate", 2), ("baseline", 3))) for row in [expected_rows()[i]]], "capture rows changed")

    export = read_json(export_path)
    require(export.get("schema") == "litchi.xlsx-selected-metadata-heaptrack-export-0895.v1", "export schema changed")
    require(export.get("status") == "pass", "export is not passing")
    require(export.get("rows") == capture.get("rows"), "export rows changed")

    verification = read_json(verification_path)
    require(verification.get("schema") == "litchi.xlsx-selected-metadata-heaptrack-verification-0895.v1", "verification schema changed")
    require(verification.get("status") == "pass", "verification is not passing")
    require(verification.get("children") == 4 and verification.get("measured_samples") == 80, "verification schedule changed")
    require(verification.get("profiler_workload_launched") is False, "verification launched a profiler workload")
    require(verification.get("capture", {}).get("receipt", {}).get("sha256") == sha_file(capture_path), "verification capture binding changed")
    require(verification.get("export", {}).get("receipt", {}).get("sha256") == sha_file(export_path), "verification export binding changed")
    require(verification.get("reduction", {}).get("sha256") == sha_file(reduction_path), "verification reduction binding changed")

    evidence = {
        "plan": descriptor(plan_path),
        "capture": descriptor(capture_path),
        "export": descriptor(export_path),
        "verification": descriptor(verification_path),
        "attribution": descriptor(attribution_path),
        "reduction": descriptor(reduction_path),
        "0895_attribution_helper": descriptor(RETAINED / "attribution.py"),
        "0895_reduction_helper": descriptor(RETAINED / "reduce.py"),
    }
    return evidence, {
        "plan": plan,
        "capture": capture,
        "export": export,
        "verification": verification,
        "attribution": read_json(attribution_path),
        "reduction": read_json(reduction_path),
    }


def region_for(frames: list[str]) -> str:
    closure = any(CLOSURE_MARKER in frame for frame in frames)
    guard = any(GUARD_MARKER in frame for frame in frames)
    require(not (closure and guard), "a stack belongs to both closure and guard")
    if closure:
        return "scan_closure"
    if guard:
        return "forced_worksheet_guard"
    if any(RUNNER_MARKER in frame for frame in frames):
        return "runner_other"
    return "process_other"


def matching_frames(frames: Iterable[str], tokens: Iterable[str]) -> list[str]:
    return [frame for frame in frames if any(token in frame for token in tokens)]


def stack_hash(stack: str) -> str:
    # The trailing semicolon is the heaptrack_print flamegraph delimiter and
    # is retained in the source line, but the signature hash covers the exact
    # frame string after removing only that delimiter.
    require(stack.endswith(";"), "stack signature is missing its terminal delimiter")
    return sha_bytes(stack[:-1].encode("utf-8"))


def stack_record(
    stack: str,
    line_rows: list[dict[str, Any]],
    category: str,
    frames: list[str],
) -> dict[str, Any]:
    return {
        "category": category,
        "allocation_cost": sum(row["allocation_cost"] for row in line_rows),
        "line_ids": [row["line_id"] for row in line_rows],
        "line_costs": [
            {
                "line_id": row["line_id"],
                "allocation_cost": row["allocation_cost"],
                "raw_line_sha256": row["raw_line_sha256"],
            }
            for row in line_rows
        ],
        "stack_sha256": stack_hash(stack),
        "frame_count": len(frames),
        "matched_frames": matching_frames(frames, subgroup_tokens(category)),
        "codec_start_positions": [index for index, frame in enumerate(frames) if START_MARKER in frame],
    }


def subgroup_tokens(category: str) -> tuple[str, ...]:
    if category == "owned_to_owned_name_hypothesis":
        return OWNED_NAME_TOKENS
    if category == "quickxml_duplicate":
        return QUICK_XML_DUPLICATE_TOKEN
    if category == "rawreserve":
        return RAW_RESERVE_TOKENS
    return ()


def classify_start(frames: list[str]) -> str:
    require(any(START_MARKER in frame for frame in frames), "start classifier called without codec::start")
    if matching_frames(frames, OWNED_NAME_TOKENS):
        return "owned_to_owned_name_hypothesis"
    if matching_frames(frames, QUICK_XML_DUPLICATE_TOKEN):
        return "quickxml_duplicate"
    if matching_frames(frames, RAW_RESERVE_TOKENS):
        return "rawreserve"
    return "other"


def unresolved_frames(frames: list[str]) -> list[str]:
    return [frame for frame in frames if UNRESOLVED.search(frame)]


def parse_stack_file(path: Path) -> dict[str, Any]:
    region_totals: Counter[str] = Counter()
    region_lines: Counter[str] = Counter()
    region_stacks: dict[str, Counter[str]] = defaultdict(Counter)
    region_unresolved_cost: Counter[str] = Counter()
    region_unresolved_rows: dict[str, list[dict[str, Any]]] = defaultdict(list)
    closest_owner: dict[str, Counter[str]] = defaultdict(Counter)
    noncodec_owner: dict[str, Counter[str]] = defaultdict(Counter)
    noncodec_stacks: dict[str, Counter[str]] = defaultdict(Counter)
    codec_owner: dict[str, Counter[str]] = defaultdict(Counter)
    start_groups: Counter[str] = Counter()
    start_group_rows: dict[str, dict[str, list[dict[str, Any]]]] = defaultdict(lambda: defaultdict(list))
    start_group_frames: dict[str, dict[str, list[str]]] = defaultdict(dict)
    all_cost = 0
    all_rows: list[dict[str, Any]] = []
    all_stacks: set[str] = set()

    with path.open(encoding="utf-8", newline="") as handle:
        for line_id, raw in enumerate(handle, 1):
            text = raw[:-1] if raw.endswith("\n") else raw
            if text.endswith("\r"):
                text = text[:-1]
            require(text and not text.startswith("#"), f"empty/comment stack line {line_id}: {path}")
            try:
                stack, raw_cost = text.rsplit(None, 1)
            except ValueError as error:
                raise ValueError(f"stack line {line_id} has no cost: {path}") from error
            require(INTEGER.fullmatch(raw_cost) is not None, f"stack line {line_id} cost is not decimal")
            cost = int(raw_cost, 10)
            require(cost > 0, f"stack line {line_id} cost is not positive")
            fields = stack.split(";")
            require(fields and fields[-1] == "", f"stack line {line_id} lacks exactly one terminal delimiter")
            frames = fields[:-1]
            require(frames and all(frames), f"stack line {line_id} has an empty frame")
            region = region_for(frames)
            row = {
                "line_id": line_id,
                "allocation_cost": cost,
                "raw_line_sha256": sha_bytes(raw.encode("utf-8")),
            }
            all_rows.append(row)
            all_stacks.add(stack)
            all_cost += cost
            region_totals[region] += cost
            region_lines[region] += 1
            region_stacks[region][stack] += cost
            apps = [frame for frame in frames if frame.startswith("litchi_")]
            owner = apps[-1] if apps else "<no litchi application frame>"
            closest_owner[region][owner] += cost
            codec_frames = [frame for frame in frames if CODEC_MARKER in frame]
            if codec_frames:
                codec_owner[region][codec_frames[-1]] += cost
            else:
                noncodec_owner[region][owner] += cost
                noncodec_stacks[region][stack] += cost
            unknown = unresolved_frames(frames)
            if unknown:
                region_unresolved_cost[region] += cost
                region_unresolved_rows[region].append({**row, "frames": unknown, "stack_sha256": stack_hash(stack)})

            # Start subgroups are limited to stacks whose nearest (rightmost)
            # codec owner is `codec::start`.  A stack can contain an inlined
            # start frame below another codec owner; assigning that stack here
            # would violate the nearest-owner rule used by 0895.
            if region == "scan_closure" and codec_frames and START_MARKER in codec_frames[-1]:
                subgroup = classify_start(frames)
                start_groups[subgroup] += cost
                start_group_rows[subgroup][stack].append(row)
                start_group_frames[subgroup].setdefault(stack, frames)

    require(all_rows, f"stack file is empty: {path}")
    require(sum(region_totals.values()) == all_cost, "region totals do not cover all stack cost")

    def top_stack_records(region: str, limit: int = 10) -> list[dict[str, Any]]:
        result = []
        for stack, cost in noncodec_stacks[region].most_common(limit):
            rows = [row for row in all_rows_for_stack(all_rows, path, stack)]
            frames = stack[:-1].split(";")
            apps = [frame for frame in frames if frame.startswith("litchi_")]
            result.append(
                {
                    "allocation_cost": cost,
                    "line_ids": [row["line_id"] for row in rows],
                    "line_costs": [
                        {
                            "line_id": row["line_id"],
                            "allocation_cost": row["allocation_cost"],
                            "raw_line_sha256": row["raw_line_sha256"],
                        }
                        for row in rows
                    ],
                    "stack_sha256": stack_hash(stack),
                    "frame_count": len(frames),
                    "closest_litchi_owner": apps[-1] if apps else "<no litchi application frame>",
                }
            )
        return result

    def owner_map(counter: Counter[str]) -> list[dict[str, Any]]:
        return [
            {"owner": owner, "allocation_cost": cost}
            for owner, cost in sorted(counter.items(), key=lambda item: (-item[1], item[0]))
        ]

    groups: dict[str, dict[str, Any]] = {}
    for group in ("owned_to_owned_name_hypothesis", "quickxml_duplicate", "rawreserve", "other"):
        signatures = []
        for stack, rows in sorted(
            start_group_rows[group].items(),
            key=lambda item: (-sum(row["allocation_cost"] for row in item[1]), stack_hash(item[0])),
        ):
            signatures.append(stack_record(stack, rows, group, start_group_frames[group][stack]))
        groups[group] = {
            "allocation_cost": start_groups[group],
            "unique_stack_count": len(signatures),
            "line_count": sum(len(rows) for rows in start_group_rows[group].values()),
            "signatures": signatures,
        }

    regions: dict[str, Any] = {}
    for region in ("scan_closure", "forced_worksheet_guard", "runner_other", "process_other"):
        cost = region_totals[region]
        codec_cost = sum(codec_owner[region].values())
        noncodec_cost = sum(noncodec_stacks[region].values())
        require(codec_cost + noncodec_cost == cost, f"codec/noncodec split does not cover {region}")
        regions[region] = {
            "allocation_cost": cost,
            "share_of_process": cost / all_cost,
            "line_count": region_lines[region],
            "unique_stack_count": len(region_stacks[region]),
            "codec_frame_presence_cost": codec_cost,
            "noncodec_cost": noncodec_cost,
            "closest_litchi_application_owners": owner_map(closest_owner[region]),
            "noncodec_closest_litchi_application_owners": owner_map(noncodec_owner[region]),
            "noncodec_top_stacks": top_stack_records(region),
            "unresolved_cost": region_unresolved_cost[region],
            "unresolved_rows": region_unresolved_rows[region],
        }

    closure_codec = {
        "allocation_cost": sum(codec_owner["scan_closure"].values()),
        "nearest_codec_owner_symbols": [
            {"symbol": symbol, "allocation_cost": cost}
            for symbol, cost in sorted(
                codec_owner["scan_closure"].items(), key=lambda item: (-item[1], item[0])
            )
        ],
        "start_subgroups": groups,
        "start_subgroup_sum": sum(start_groups.values()),
        "non_start_codec_cost": sum(codec_owner["scan_closure"].values()) - sum(start_groups.values()),
    }
    require(closure_codec["start_subgroup_sum"] == 343_320, "unexpected retained start total")
    require(closure_codec["allocation_cost"] == 345_380, "unexpected retained closure codec total")
    require(regions["scan_closure"]["allocation_cost"] == 518_660, "unexpected retained closure total")
    require(regions["scan_closure"]["noncodec_cost"] == 173_280, "unexpected retained noncodec closure total")
    require(regions["process_other"]["allocation_cost"] == 154_156, "unexpected retained process total")
    require(regions["runner_other"]["allocation_cost"] == 847, "unexpected retained runner total")

    return {
        "input": descriptor(path),
        "stack_lines": len(all_rows),
        "unique_stacks": len(all_stacks),
        "allocation_cost": all_cost,
        "regions": regions,
        "closure_codec": closure_codec,
        "unresolved_allocation_cost": sum(region_unresolved_cost.values()),
        "unresolved_rows": [
            {"region": region, **row}
            for region in ("scan_closure", "forced_worksheet_guard", "runner_other", "process_other")
            for row in region_unresolved_rows[region]
        ],
    }


def all_rows_for_stack(all_rows: list[dict[str, Any]], path: Path, stack: str) -> list[dict[str, Any]]:
    # The stack text is small compared with the retained export.  Re-read only
    # for a selected top/signature row so the result can retain exact input
    # line IDs and raw-line hashes without embedding the full stack strings.
    rows = []
    with path.open(encoding="utf-8", newline="") as handle:
        for line_id, raw in enumerate(handle, 1):
            text = raw[:-1] if raw.endswith("\n") else raw
            if text.endswith("\r"):
                text = text[:-1]
            if text.rsplit(None, 1)[0] == stack:
                rows.append(
                    {
                        "line_id": line_id,
                        "allocation_cost": int(text.rsplit(None, 1)[1], 10),
                        "raw_line_sha256": sha_bytes(raw.encode("utf-8")),
                    }
                )
    require(rows, f"selected stack disappeared during offline read: {path}")
    return rows


def canonical_start_signatures(analysis: dict[str, Any]) -> dict[str, Any]:
    """Return repeatability data without depending on stack-file line order.

    The full profile keeps line IDs and raw-line hashes for custody.  The
    repeated-baseline check compares the stack signature, cost, and matching
    frame metadata so a harmless export reorder cannot look like a count
    change.
    """
    canonical: dict[str, Any] = {}
    for group, value in analysis["closure_codec"]["start_subgroups"].items():
        signatures = [
            {
                "stack_sha256": signature["stack_sha256"],
                "allocation_cost": signature["allocation_cost"],
                "line_count": len(signature["line_costs"]),
                "frame_count": signature["frame_count"],
                "matched_frames": signature["matched_frames"],
                "codec_start_positions": signature["codec_start_positions"],
            }
            for signature in value["signatures"]
        ]
        signatures.sort(key=lambda item: item["stack_sha256"])
        canonical[group] = {
            "allocation_cost": value["allocation_cost"],
            "unique_stack_count": value["unique_stack_count"],
            "line_count": value["line_count"],
            "signatures": signatures,
        }
    return canonical


def reduction_crosscheck(evidence_data: dict[str, Any], analyses: list[dict[str, Any]]) -> dict[str, Any]:
    attribution = evidence_data["attribution"]
    reduction = evidence_data["reduction"]
    expected_regions = {
        "scan_closure": 518_660,
        "forced_worksheet_guard": 0,
        "runner_other": 847,
        "process_other": 154_156,
    }
    expected_exclusive = {
        "legacycodec": 535_697,
        "mce_stream": 14_418,
        "other": 112_395,
        "selected_worksheet": 1_361,
        "x14ac": 9_792,
    }
    rows = []
    for analysis, (key, ordinal) in zip(analyses, ROWS):
        attr = attribution["profiles"].get(key)
        require(isinstance(attr, dict), f"0895 attribution row is missing: {key}")
        require(attr.get("total") == 673_663, f"0895 attribution total changed: {key}")
        observed_regions = dict(attr.get("regions", {}))
        observed_regions.setdefault("forced_worksheet_guard", 0)
        require(observed_regions == expected_regions, f"0895 attribution regions changed: {key}")
        reduction_profile = next(
            profile for profile in reduction["profiles"] if profile["row"].get("ordinal") == ordinal
        )
        require(reduction_profile["row"].get("role") == "baseline", f"0895 reduction role changed: {key}")
        require(reduction_profile["reconciliation"]["pass"] is True, f"0895 reduction reconciliation failed: {key}")
        require(reduction_profile["reconciliation"]["collapsed_stack_total"] == 673_663, f"0895 reduced total changed: {key}")
        observed_exclusive = {
            name: item["allocation_cost"]
            for name, item in reduction_profile["exclusive"].items()
        }
        require(observed_exclusive == expected_exclusive, f"0895 exclusive attribution changed: {key}")
        require(analysis["allocation_cost"] == 673_663, f"new parser total changed: {key}")
        require(analysis["regions"]["scan_closure"]["allocation_cost"] == EXPECTED_CLOSURE_PER_SAMPLE * EXPECTED_SAMPLES, f"closure repetition changed: {key}")
        rows.append(
            {
                "profile": key,
                "ordinal": ordinal,
                "stack_sha256": analysis["input"]["sha256"],
                "new_parser_allocation_cost": analysis["allocation_cost"],
                "0895_attribution_total": attr["total"],
                "0895_attribution_regions": observed_regions,
                "0895_reduction_reconciliation": reduction_profile["reconciliation"],
                "0895_reduction_exclusive": observed_exclusive,
                "start_signature_counts_and_hashes": canonical_start_signatures(analysis),
                "line_provenance_retained_in_profile": True,
                "exact_matches": True,
            }
        )
    require(canonical_start_signatures(analyses[0]) == canonical_start_signatures(analyses[1]), "baseline subgroup counts or signature hashes differ")
    require(analyses[0]["closure_codec"]["nearest_codec_owner_symbols"] == analyses[1]["closure_codec"]["nearest_codec_owner_symbols"], "baseline codec-owner counts differ")
    return {
        "rows": rows,
        "repeated_baseline_counts_identical": True,
        "scan_closure_counter": {
            "per_sample_allocation_cost": EXPECTED_CLOSURE_PER_SAMPLE,
            "samples": EXPECTED_SAMPLES,
            "expression": "25933 * 20 = 518660",
            "allocation_cost": EXPECTED_CLOSURE_PER_SAMPLE * EXPECTED_SAMPLES,
        },
        "candidate_rows_analyzed": 0,
        "candidate_comparison": None,
    }


def build_result() -> dict[str, Any]:
    evidence, evidence_data = validate_receipts()
    head = git_text("rev-parse", "HEAD")
    require(is_ancestor(OBSERVED_EVALUATION_HEAD, head), "current HEAD does not descend from the observed evaluation head")
    source = {name: source_descriptor(name, path) for name, path in SOURCE_FILES.items()}
    require(all(item["matches_base_blob"] for item in source.values()), "current production source differs from 506aa2 baseline")
    require(all(item["matches_observed_evaluation_blob"] for item in source.values()), "current source differs from the observed evaluation head")

    stack_paths = [
        RETAINED / "export-v1/runs/00-baseline/stacks.txt",
        RETAINED / "export-v1/runs/03-baseline/stacks.txt",
    ]
    analyses = [parse_stack_file(path) for path in stack_paths]
    for analysis in analyses:
        require(analysis["allocation_cost"] == 673_663, "unexpected baseline allocation total")
        require(analysis["unresolved_allocation_cost"] == 1, "unexpected unresolved baseline cost")
        require(sum(item["allocation_cost"] for item in analysis["regions"].values()) == analysis["allocation_cost"], "region accounting does not close")
    crosscheck = reduction_crosscheck(evidence_data, analyses)

    codec_path = SOURCE_FILES["mce_codec"]
    owned_group = analyses[0]["closure_codec"]["start_subgroups"]["owned_to_owned_name_hypothesis"]
    candidate_lines = [
        {
            "stack_sha256": signature["stack_sha256"],
            "line": line,
        }
        for signature in owned_group["signatures"]
        for line in signature["line_costs"]
        if line["line_id"] == 1
    ]
    require(len(candidate_lines) == 1, "retained line 1 is not uniquely in the owned-name hypothesis group")
    candidate_line = candidate_lines[0]["line"]
    candidate_signature = next(
        signature
        for signature in owned_group["signatures"]
        if signature["stack_sha256"] == candidate_lines[0]["stack_sha256"]
    )
    source_bindings = {
        "evaluation_head": OBSERVED_EVALUATION_HEAD,
        "current_head": head,
        "base_commit": BASE_COMMIT,
        "current_head_descends_from_evaluation_head": is_ancestor(OBSERVED_EVALUATION_HEAD, head),
        "files": source,
        "start_owned_name_expressions": {
            "possible_name_retention_lines": [
                source_line(codec_path, 950),
                source_line(codec_path, 1116),
                source_line(codec_path, 1132),
            ],
            "other_owned_expressions_in_start": [
                source_line(codec_path, 903),
                source_line(codec_path, 909),
                source_line(codec_path, 996),
                source_line(codec_path, 997),
            ],
        },
        "start_function": source_line(codec_path, 855),
        "checked_attributes_loop": source_line(codec_path, 877),
        "reserve_amortized": source_line(codec_path, 200),
        "expanded_attribute_duplicate_checker": source_line(codec_path, 1414),
    }

    return {
        "schema": "litchi.xlsx-baseline-allocation-attribution-0897.v1",
        "status": "pass",
        "scope": {
            "mode": "offline-only retained-artifact analysis",
            "roles_analyzed": ["baseline"],
            "candidate_rows_analyzed": 0,
            "whole_process_stack_cost": True,
            "allocation_counter_semantics": "heaptrack allocation-hook call cost from the retained 0895 flamegraph export",
            "not_a_latency_metric": True,
            "native_timing_claim": None,
            "performance_claim": None,
        },
        "tooling": {
            "helper": descriptor(Path(__file__)),
            "python": descriptor(Path(sys.executable)),
            "workload_launched_by_this_helper": False,
            "build_launched_by_this_helper": False,
            "profiler_launched_by_this_helper": False,
            "external_profiler_replayed": False,
        },
        "source_bindings": source_bindings,
        "evidence": evidence,
        "classification": {
            "region_precedence": ["scan_closure", "forced_worksheet_guard", "runner_other", "process_other"],
            "codec_owner_direction": "rightmost codec frame in the root-to-leaf heaptrack stack",
            "start_subgroup_precedence": ["owned_to_owned_name_hypothesis", "quickxml_duplicate", "rawreserve", "other"],
            "start_subgroup_rules": {
                "owned_to_owned_name_hypothesis": {
                    "tokens": list(OWNED_NAME_TOKENS),
                    "label": "owned str::to_owned/name mechanism hypothesis",
                    "meaning": "presence of retained ToOwned symbols only; generic slice ToOwned does not prove the Rust type or source call site",
                },
                "quickxml_duplicate": list(QUICK_XML_DUPLICATE_TOKEN),
                "rawreserve": {
                    "tokens": list(RAW_RESERVE_TOKENS),
                    "label": "generic growth/reserve stack (raw-reserve mechanism hypothesis)",
                    "meaning": "presence does not identify the allocated object or prove that the allocation belongs to the raw-attribute Vec",
                },
                "other": "no preceding explicit presence rule matched; retained in the total",
            },
            "inclusive_values_used": False,
            "unknown_or_unresolved_cost_discarded": 0,
            "stack_hash_definition": "SHA-256 of the exact flamegraph frame string after removing its terminal delimiter; raw input line hashes retain the stack, cost, and newline",
        },
        "profiles": [
            {
                "profile": key,
                "ordinal": ordinal,
                "analysis": analysis,
            }
            for (key, ordinal), analysis in zip(ROWS, analyses)
        ],
        "q_to_owned_inference": {
            "measured_subgroup": "owned_to_owned_name_hypothesis",
            "raw_measured_stack_is_not_source_line_proof": True,
            "possible_source_lines": [
                source_line(codec_path, 950),
                source_line(codec_path, 1116),
                source_line(codec_path, 1132),
            ],
            "other_start_owned_expression_lines": [
                source_line(codec_path, 903),
                source_line(codec_path, 909),
                source_line(codec_path, 996),
                source_line(codec_path, 997),
            ],
            "candidate_raw_line": {
                "line_id": candidate_line["line_id"],
                "raw_line_sha256": candidate_line["raw_line_sha256"],
                "stack_sha256": candidate_signature["stack_sha256"],
                "allocation_cost": candidate_line["allocation_cost"],
                "signature_aggregate_allocation_cost": candidate_signature["allocation_cost"],
                "signature_line_costs": candidate_signature["line_costs"],
            },
            "inference_basis": [
                "the exact retained stack contains Rust ToOwned symbols and codec::start",
                "codec.rs lines 950, 1116, and 1132 are all possible q.to_owned/name-retention expressions in start",
                "other owned expressions in start (lines 903, 909, 996, and 997) remain possible symbol sources",
                "heaptrack_print preserves a root-to-leaf stack with codec::start appearing after str::to_owned in this inlined signature",
            ],
            "claim": "The line-1 raw signature is consistent with an owned-name mechanism in codec::start, but the allocation counter cannot distinguish the three possible q.to_owned/name-retention lines or other owned expressions; the report records only a mechanism hypothesis.",
        },
        "baseline_repetition_validation": crosscheck,
        "analysis_limits": [
            "The retained profiles include setup, scan, serialization, teardown, and guards; stack presence has no temporal segmentation.",
            "Subgroups are disjoint stack-presence buckets, not exclusive source-line or byte-size attribution.",
            "The noncodec owner table is retained to prevent the codec subgroup from being mistaken for all end-to-end allocation cost.",
            "No candidate comparison, native timing claim, RSS causal claim, or production-adoption decision is made here.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=PACKET / "analysis.json")
    args = parser.parse_args()
    output = args.output.resolve()
    require(not output.exists() and not output.is_symlink(), f"output already exists or is symlinked: {output}")
    result = build_result()
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as handle:
        json.dump(result, handle, indent=2, sort_keys=True)
        handle.write("\n")
    print(json.dumps({"output": descriptor(output), "status": result["status"]}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
