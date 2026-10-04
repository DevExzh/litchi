#!/usr/bin/env python3
"""Reduce retained 0892/0895 diagnostics without launching a workload.

Callgrind rows and perf leaf periods are kept in separate namespaces.  The
Callgrind table is reconciled as exclusive guest instructions; its selected
inclusive rows are retained only as overlapping context.  The heaptrack stack
and perf leaf views are independent whole-process evidence and make no native
latency claim.
"""

from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import re
import stat
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
BASELINE_REVISION = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
EVALUATION_HEAD = "4f8f667983672776be8c747e89795195d15a3357"
TOTAL_ALLOCATION_COST = 673_663
CLOSURE_ALLOCATION_COST = 518_660
DUPLICATE_FN = "quick_xml::events::attributes::IterState::check_for_duplicates"
ITER_NEXT_FN = "quick_xml::events::attributes::IterState::next"
MCE_START_FN = "litchi_ooxml_common::mce::codec::start"
MCE_PROCESS_FN = "litchi_ooxml_common::mce::codec::process_markup_compatibility"
CODEC_MARKER = "litchi_ooxml_common::mce::codec::"
START_MARKER = "litchi_ooxml_common::mce::codec::start::"
CLOSURE_MARKER = (
    "litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan::"
    "_$u7b$$u7b$closure$u7d$$u7d$::"
)
GUARD_MARKER = (
    "litchi_perf_baseline::producer_shape::"
    "force_full_cell_worksheet_tolerance_reports::"
)
RUNNER_MARKER = "litchi_perf_baseline::producer_shape::run_xlsx_full_cell_scan::"

CALLGRIND_ROWS = (
    ("quick_xml_duplicate_check", DUPLICATE_FN, "/quick-xml-0.41.0/src/events/attributes.rs:"),
    ("quick_xml_iter_state_next", ITER_NEXT_FN, "/quick-xml-0.41.0/src/events/attributes.rs:"),
    ("mce_start", MCE_START_FN, "/crates/litchi-ooxml-common/src/mce/codec.rs:"),
    ("mce_process", MCE_PROCESS_FN, "/crates/litchi-ooxml-common/src/mce/codec.rs:"),
)
LEAF_SELECTORS = {
    "duplicate_check": DUPLICATE_FN,
    "bounded_output_reserve": "litchi_ooxml_common::mce::codec::BoundedOutput::reserve",
    "mce_esc": "litchi_ooxml_common::mce::codec::esc",
    "mce_start": MCE_START_FN,
    "quick_xml_iter_state_next": ITER_NEXT_FN,
    "mce_expand_parts": "litchi_ooxml_common::mce::codec::expand_parts",
    "is_qualified_name": "litchi_ooxml_common::xml_name::validation::is_qualified_name",
    "is_ncname": "litchi_ooxml_common::xml_name::validation::is_ncname",
}
CALLGRIND_EXCLUSIVE_SELECTORS = {
    "duplicate_check": DUPLICATE_FN,
    "bounded_output_reserve": "litchi_ooxml_common::mce::codec::BoundedOutput::reserve",
    "mce_esc": "litchi_ooxml_common::mce::codec::esc",
    "mce_start": MCE_START_FN,
    "quick_xml_iter_state_next": ITER_NEXT_FN,
    "mce_expand_parts": "litchi_ooxml_common::mce::codec::expand_parts",
    "is_qualified_name": "litchi_ooxml_common::xml_name::validation::is_qualified_name",
    "is_ncname": "litchi_ooxml_common::xml_name::validation::is_ncname",
}
EXPECTED_LEAF = {
    "child-00": {"sample_count": 3817, "total_period": 31_519_530_592},
    "child-01": {"sample_count": 3814, "total_period": 31_665_956_013},
    "child-02": {"sample_count": 3800, "total_period": 31_453_030_639},
}

ANNOTATION_ROW = re.compile(
    r"^\s*(?P<ir>[0-9][0-9,]*)\s+\(\s*(?P<percent>[0-9.]+)\s*%\)\s+(?P<location>.+?)\s*$"
)
ANNOTATION_HEADER = re.compile(r"^\s*Ir\s+file:function\s*$")
SEPARATOR = re.compile(r"^\s*-{8,}\s*$")
DOT_ROW = re.compile(r"^\s*\.\s+.+$")
LEAF_ROW = re.compile(
    r"^\d+/\d+\s+[\d.]+:\s+(?P<period>\d+)\s+"
    r"(?P<event>\S+):\s+(?P<ip>[0-9a-f]+)\s+(?P<symbol>.*?)\s+\((?P<dso>.*)\)$"
)


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
    mode = path.stat().st_mode
    require(stat.S_ISREG(mode), f"not a regular file: {path}")
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha_file(path),
        "mode": stat.S_IMODE(mode),
        "executable": bool(mode & stat.S_IXUSR),
    }


def read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def path(relative: str) -> Path:
    return ROOT / relative


def validate_source_identity() -> dict[str, Any]:
    identity_path = PACKET / "source-identity.json"
    identity = read_json(identity_path)
    require(identity.get("schema") == "litchi.duplicate-attribution-source-0899.v1", "source identity schema changed")
    require(identity.get("status") == "pass", "source identity is not passing")
    require(identity.get("evaluation_head") == EVALUATION_HEAD, "source evaluation HEAD changed")
    require(identity.get("baseline_revision") == BASELINE_REVISION, "source baseline revision changed")
    require(identity.get("verified_selected_files") == 10_771, "selected source census changed")
    require(identity.get("source_selection_unchanged") is True, "selected source set changed")
    require(identity.get("accepted_adrs_unchanged") is True, "accepted ADR set changed")
    require(identity.get("callgrind_0892_production_crates_unchanged") is True, "callgrind production source changed")
    require(identity.get("no_workload_or_build") is True, "source identity was not offline-only")

    sources = identity.get("sources")
    require(isinstance(sources, dict), "source identity has no selected sources")
    required_sources = {
        "crates/litchi-opc/src/xml_attributes.rs",
        "crates/litchi-opc/src/xml_attributes/tests.rs",
        "crates/litchi-ooxml-common/src/mce/codec.rs",
        "tools/perf-baseline/src/producer_shape.rs",
        "tools/perf-baseline/Cargo.lock",
    }
    require(required_sources.issubset(sources), "source identity omits a relevant owner/build path")
    checked = {}
    for name, expected in sources.items():
        require(isinstance(expected, dict), f"invalid source descriptor: {name}")
        current = descriptor(Path(expected["path"]))
        require(current["bytes"] == expected.get("bytes"), f"source byte count changed: {name}")
        require(current["sha256"] == expected.get("sha256"), f"source checksum changed: {name}")
        checked[name] = current

    quick_xml = identity.get("quick_xml")
    require(isinstance(quick_xml, dict) and quick_xml.get("version") == "0.41.0", "quick-xml version changed")
    require(quick_xml.get("cargo_checksum") == "e660451e55124f798a69a5af3f49ccfbefbd41910eefd25caf2393e1f3473ec1", "quick-xml checksum changed")
    for key in ("crate", "attributes_source"):
        expected = quick_xml.get(key)
        require(isinstance(expected, dict), f"quick-xml identity omits {key}")
        current = descriptor(Path(expected["path"]))
        require(current["bytes"] == expected.get("bytes"), f"quick-xml {key} byte count changed")
        require(current["sha256"] == expected.get("sha256"), f"quick-xml {key} checksum changed")
    require(quick_xml.get("registry_source_matches_checksum_verified_archive") is True, "quick-xml registry custody is not verified")
    return {
        "receipt": descriptor(identity_path),
        "evaluation_head": identity["evaluation_head"],
        "baseline_revision": identity["baseline_revision"],
        "verified_selected_files": identity["verified_selected_files"],
        "sources": checked,
        "quick_xml": quick_xml,
    }


def validate_prior_evidence() -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    verification_path = path("docs/performance/results/change-0895/verification.json")
    allocation_analysis_path = path("docs/performance/results/change-0897/analysis.json")
    callgrind_analysis_path = path("docs/performance/results/change-0892/callgrind-analysis.json")
    leaf_analysis_path = path("docs/performance/results/change-0892/self-analysis.json")
    verification = read_json(verification_path)
    allocation_analysis = read_json(allocation_analysis_path)
    callgrind_analysis = read_json(callgrind_analysis_path)
    leaf_analysis = read_json(leaf_analysis_path)
    allocation_custody_path = PACKET / "allocation-custody-replay.json"
    profile_custody_path = PACKET / "profile-content-replay.json"
    allocation_custody = read_json(allocation_custody_path)
    profile_custody = read_json(profile_custody_path)
    require(verification.get("schema") == "litchi.xlsx-selected-metadata-heaptrack-verification-0895.v1", "0895 verification schema changed")
    require(verification.get("status") == "pass", "0895 verification is not passing")
    require(verification.get("workload_launched") is False and verification.get("profiler_workload_launched") is False, "0895 verification launched work")
    require(verification.get("performance_claim") is None, "0895 verification contains a performance claim")
    require(all(allocation_analysis.get(key) == value for key, value in (
        ("schema", "litchi.xlsx-baseline-allocation-attribution-0897.v1"),
        ("status", "pass"),
    )), "0897 analysis identity changed")
    require(allocation_analysis.get("scope", {}).get("candidate_rows_analyzed") == 0, "0897 candidate rows were analyzed")
    require(allocation_analysis.get("scope", {}).get("not_a_latency_metric") is True, "0897 scope changed")
    require(allocation_analysis.get("scope", {}).get("native_timing_claim") is None, "0897 contains native timing")
    require(callgrind_analysis.get("schema") == "litchi.xlsx-scan-current-runner-ir-0892.v1", "0892 Callgrind analysis schema changed")
    require(callgrind_analysis.get("status") == "pass" and callgrind_analysis.get("performance_claim") is False, "0892 Callgrind analysis is not diagnostic-only")
    require("overlap" in callgrind_analysis.get("comparison_to_0886", {}), "0892 overlap declaration missing")
    require(leaf_analysis.get("schema") == "litchi.xlsx-scan-current-self-analysis-0892.v1", "0892 leaf analysis schema changed")
    require(leaf_analysis.get("status") == "pass" and leaf_analysis.get("performance_claim") is False, "0892 leaf analysis is not diagnostic-only")
    require(len(leaf_analysis.get("rows", [])) == 3, "0892 leaf analysis child count changed")
    require(allocation_custody.get("schema") == "litchi.xlsx-selected-metadata-heaptrack-verification-0895.v1" and allocation_custody.get("status") == "pass", "fresh 0895 allocation custody replay is not passing")
    require(profile_custody.get("schema") == "litchi.profile-content-custody-0899.v1" and profile_custody.get("status") == "pass", "fresh profile content custody replay is not passing")
    require(profile_custody.get("strict_historical_verifier_pass") is False, "profile custody replay unexpectedly claims strict historical verifier success")
    evidence_paths = {
        "0895_verification": verification_path,
        "0895_reduction_helper": path("docs/performance/results/change-0895/reduce.py"),
        "0895_attribution_helper": path("docs/performance/results/change-0895/attribution.py"),
        "0897_analysis": allocation_analysis_path,
        "0897_reduction_helper": path("docs/performance/results/change-0897/reduce.py"),
        "0892_callgrind_analysis": callgrind_analysis_path,
        "0892_callgrind_helper": path("docs/performance/results/change-0892/analyze_callgrind.py"),
        "0892_leaf_analysis": leaf_analysis_path,
        "0892_leaf_helper": path("docs/performance/results/change-0892/reduce.py"),
        "0899_allocation_custody_replay": allocation_custody_path,
        "0899_profile_content_replay": profile_custody_path,
    }
    return (
        {name: descriptor(item) for name, item in evidence_paths.items()},
        {"verification": verification, "allocation_analysis": allocation_analysis},
        {"callgrind_analysis": callgrind_analysis, "leaf_analysis": leaf_analysis},
    )


def validate_annotation_receipt(profile: str, text_path: Path, receipt_path: Path, stderr_path: Path) -> dict[str, Any]:
    receipt = read_json(receipt_path)
    raw_path = path(f"docs/performance/results/change-0892/callgrind/{profile}/callgrind.out")
    raw = descriptor(raw_path)
    text = descriptor(text_path)
    stderr = descriptor(stderr_path)
    require(receipt.get("status") == "pass" and receipt.get("exit_code") == 0, f"{profile} self annotation did not pass")
    argv = receipt.get("argv")
    require(isinstance(argv, list) and "--inclusive=no" in argv and "--threshold=100" in argv and "--auto=no" in argv and argv[-1] == raw["path"], f"{profile} annotation flags/input changed")
    require("--inclusive=yes" not in argv, f"{profile} annotation is not exclusive")
    require(receipt.get("scope") == "Offline annotation of existing raw profile; no workload or profiler capture launched", f"{profile} annotation scope changed")
    require(receipt.get("input", {}).get("path") == raw["path"], f"{profile} annotation input path changed")
    require(receipt.get("input", {}).get("sha256") == raw["sha256"], f"{profile} annotation input changed")
    require(receipt.get("output", {}).get("path") == text["path"], f"{profile} annotation output path changed")
    require(receipt.get("output", {}).get("bytes") == text["bytes"] and receipt.get("output", {}).get("sha256") == text["sha256"], f"{profile} annotation output custody failed")
    require(receipt.get("stderr", {}).get("path") == stderr["path"], f"{profile} annotation stderr path changed")
    require(receipt.get("stderr", {}).get("bytes") == 0 and stderr["bytes"] == 0, f"{profile} annotation emitted stderr")
    return {
        "receipt": descriptor(receipt_path),
        "raw": raw,
        "annotation": text,
        "stderr": stderr,
    }


def function_matches(location: str, selector: str) -> bool:
    marker = ":" + selector
    start = 0
    while True:
        found = location.find(marker, start)
        if found < 0:
            return False
        end = found + len(marker)
        suffix = location[end:]
        if not suffix or suffix.startswith(" ["):
            return True
        start = end


def parse_annotation(pathname: Path, reconcile: bool) -> dict[str, Any]:
    with pathname.open(encoding="utf-8", newline="") as handle:
        text = handle.read()
    lines = text.splitlines(keepends=True)
    totals = []
    for number, raw in enumerate(lines, 1):
        line = raw.rstrip("\r\n")
        match = ANNOTATION_ROW.fullmatch(line)
        if match and match.group("location") == "PROGRAM TOTALS":
            totals.append((number, int(match.group("ir").replace(",", ""))))
    require(len(totals) == 1, f"Callgrind annotation has no unique PROGRAM TOTALS row: {pathname}")
    headers = [number for number, raw in enumerate(lines, 1) if ANNOTATION_HEADER.fullmatch(raw.rstrip("\r\n"))]
    require(len(headers) == 1, f"Callgrind annotation has no unique function table: {pathname}")
    rows = []
    for number, raw in enumerate(lines[headers[0]:], headers[0] + 1):
        line = raw.rstrip("\r\n")
        if not line.strip() or SEPARATOR.fullmatch(line):
            continue
        match = ANNOTATION_ROW.fullmatch(line)
        if match:
            rows.append({
                "line": number,
                "self_or_inclusive_ir": int(match.group("ir").replace(",", "")),
                "reported_percent": float(match.group("percent")),
                "location": match.group("location"),
                "raw_line_sha256": sha_bytes(raw.encode("utf-8")),
            })
            continue
        require(DOT_ROW.fullmatch(line) is not None, f"unparsed Callgrind table line {pathname}:{number}")
    require(rows, f"Callgrind annotation has no numeric function rows: {pathname}")
    total = totals[0][1]
    summed = sum(row["self_or_inclusive_ir"] for row in rows)
    if reconcile:
        require(summed == total, f"exclusive Callgrind rows do not reconcile for {pathname}: {summed} != {total}")
    return {
        "input": descriptor(pathname),
        "total_ir": total,
        "numeric_row_count": len(rows),
        "summed_rows_ir": summed,
        "rows": rows,
        "reconciled": summed == total,
    }


def selected_row(annotation: dict[str, Any], selector: str, source_suffix: str) -> dict[str, Any]:
    matches = [
        row for row in annotation["rows"]
        if source_suffix in row["location"] and function_matches(row["location"], selector)
    ]
    require(len(matches) == 1, f"canonical inclusive row is missing or ambiguous: {selector}")
    row = dict(matches[0])
    row["selector"] = selector
    row["inclusive_ir"] = row.pop("self_or_inclusive_ir")
    row["percent_of_program"] = 100 * row["inclusive_ir"] / annotation["total_ir"]
    return row


def analyze_callgrind(profile: str, self_path: Path, self_receipt: Path, self_stderr: Path, inclusive_path: Path, prior_row: dict[str, Any]) -> dict[str, Any]:
    receipt = validate_annotation_receipt(profile, self_path, self_receipt, self_stderr)
    exclusive = parse_annotation(self_path, reconcile=True)
    inclusive = parse_annotation(inclusive_path, reconcile=False)
    canonical = {}
    for name, selector, suffix in CALLGRIND_ROWS:
        canonical[name] = selected_row(inclusive, selector, suffix)
    require(inclusive["total_ir"] == prior_row["total_ir"], f"0892 inclusive total changed: {profile}")
    expected = prior_row.get("selected_inclusive_rows", {})
    for key in ("mce_start", "mce_process"):
        require(canonical[key]["inclusive_ir"] == expected[key]["inclusive_ir"], f"0892 canonical {key} changed: {profile}")
    exclusive_groups = {}
    for name, selector in CALLGRIND_EXCLUSIVE_SELECTORS.items():
        matches = []
        for row in exclusive["rows"]:
            if function_matches(row["location"], selector):
                matches.append({
                    "line": row["line"],
                    "self_ir": row["self_or_inclusive_ir"],
                    "reported_percent": row["reported_percent"],
                    "percent_of_program": 100 * row["self_or_inclusive_ir"] / exclusive["total_ir"],
                    "location": row["location"],
                    "raw_line_sha256": row["raw_line_sha256"],
                    "selector": selector,
                })
        require(matches, f"exclusive function rows are missing: {profile}/{selector}")
        combined = sum(row["self_ir"] for row in matches)
        exclusive_groups[name] = {
            "selector": selector,
            "matching_rows": matches,
            "matching_row_count": len(matches),
            "combined_self_ir": combined,
            "combined_self_percent": 100 * combined / exclusive["total_ir"],
        }
    return {
        "profile": profile,
        "exclusive_annotation": receipt,
        "exclusive_reconciliation": {
            "total_ir": exclusive["total_ir"],
            "summed_numeric_rows_ir": exclusive["summed_rows_ir"],
            "numeric_row_count": exclusive["numeric_row_count"],
            "pass": exclusive["reconciled"],
        },
        "duplicate_check": {
            "selector": DUPLICATE_FN,
            "matching_rows": exclusive_groups["duplicate_check"]["matching_rows"],
            "matching_row_count": exclusive_groups["duplicate_check"]["matching_row_count"],
            "combined_self_ir": exclusive_groups["duplicate_check"]["combined_self_ir"],
            "combined_self_percent": exclusive_groups["duplicate_check"]["combined_self_percent"],
            "percent_definition": "sum of exclusive matching self Ir divided by PROGRAM TOTALS Ir; printed row percentages are not summed",
        },
        "exclusive_function_groups": exclusive_groups,
        "canonical_inclusive_rows": canonical,
        "inclusive_context": {
            "total_ir": inclusive["total_ir"],
            "annotation": inclusive["input"],
            "overlap": True,
            "native_time": None,
            "interpretation": "Selected inclusive guest-instruction rows are overlapping context and are never summed or treated as native latency.",
        },
    }


def classify_region(frames: list[str]) -> str:
    closure = any(CLOSURE_MARKER in frame for frame in frames)
    guard = any(GUARD_MARKER in frame for frame in frames)
    require(not (closure and guard), "stack belongs to both closure and guard")
    if closure:
        return "scan_closure"
    if guard:
        return "forced_worksheet_guard"
    if any(RUNNER_MARKER in frame for frame in frames):
        return "runner_other"
    return "process_other"


def parse_stack(pathname: Path, prior_profile: dict[str, Any]) -> dict[str, Any]:
    regions: Counter[str] = Counter()
    region_lines: Counter[str] = Counter()
    region_stacks: dict[str, set[str]] = {name: set() for name in ("scan_closure", "forced_worksheet_guard", "runner_other", "process_other")}
    duplicate_cost = Counter({"whole_process": 0, "scan_closure": 0})
    duplicate_lines = Counter({"whole_process": 0, "scan_closure": 0})
    duplicate_stacks: dict[str, set[str]] = {"whole_process": set(), "scan_closure": set()}
    owner_cost: Counter[str] = Counter()
    owner_lines: Counter[str] = Counter()
    owner_stacks: dict[str, set[str]] = {}
    owner_class_cost: Counter[str] = Counter()
    owner_class_lines: Counter[str] = Counter()
    all_cost = 0
    line_count = 0
    with pathname.open(encoding="utf-8", newline="") as handle:
        for line_number, raw in enumerate(handle, 1):
            text = raw[:-1] if raw.endswith("\n") else raw
            if text.endswith("\r"):
                text = text[:-1]
            require(text and not text.startswith("#"), f"invalid heaptrack stack line {pathname}:{line_number}")
            try:
                stack, cost_text = text.rsplit(None, 1)
            except ValueError as error:
                raise ValueError(f"heaptrack stack line has no cost: {pathname}:{line_number}") from error
            require(cost_text.isdecimal() and int(cost_text) > 0, f"invalid heaptrack stack cost: {pathname}:{line_number}")
            cost = int(cost_text)
            fields = stack.split(";")
            require(fields and fields[-1] == "" and all(fields[:-1]), f"invalid heaptrack stack framing: {pathname}:{line_number}")
            frames = fields[:-1]
            region = classify_region(frames)
            all_cost += cost
            line_count += 1
            regions[region] += cost
            region_lines[region] += 1
            region_stacks[region].add(stack)
            codec_frames = [frame for frame in frames if CODEC_MARKER in frame]
            nearest = codec_frames[-1] if codec_frames else "<no codec frame>"
            if region == "scan_closure":
                owner_cost[nearest] += cost
                owner_lines[nearest] += 1
                owner_stacks.setdefault(nearest, set()).add(stack)
                owner_class = "no_codec_owner" if not codec_frames else ("mce_start" if START_MARKER in nearest else "other_codec_owner")
                owner_class_cost[owner_class] += cost
                owner_class_lines[owner_class] += 1
            if any(DUPLICATE_FN in frame for frame in frames):
                duplicate_cost["whole_process"] += cost
                duplicate_lines["whole_process"] += 1
                duplicate_stacks["whole_process"].add(stack)
                if region == "scan_closure":
                    duplicate_cost["scan_closure"] += cost
                    duplicate_lines["scan_closure"] += 1
                    duplicate_stacks["scan_closure"].add(stack)

    require(all_cost == TOTAL_ALLOCATION_COST, f"heaptrack total changed: {pathname} {all_cost}")
    require(regions["scan_closure"] == CLOSURE_ALLOCATION_COST, f"heaptrack closure total changed: {pathname} {regions['scan_closure']}")
    require(sum(regions.values()) == all_cost, f"heaptrack region reconciliation failed: {pathname}")
    require(sum(owner_class_cost.values()) == CLOSURE_ALLOCATION_COST, f"heaptrack closure owner reconciliation failed: {pathname}")
    expected = prior_profile["analysis"]
    require(expected["allocation_cost"] == all_cost, f"0897 allocation total differs: {pathname}")
    for region in ("scan_closure", "forced_worksheet_guard", "runner_other", "process_other"):
        require(expected["regions"].get(region, {}).get("allocation_cost", 0) == regions[region], f"0897 region differs: {pathname}/{region}")

    owner_rows = []
    for owner, cost in sorted(owner_cost.items(), key=lambda item: (-item[1], item[0])):
        owner_rows.append({
            "nearest_codec_owner": owner,
            "stack_line_count": owner_lines[owner],
            "unique_stack_count": len(owner_stacks[owner]),
            "allocation_cost": cost,
            "percent_of_closure": 100 * cost / CLOSURE_ALLOCATION_COST,
        })
    class_rows = {}
    for name in ("mce_start", "other_codec_owner", "no_codec_owner"):
        class_rows[name] = {
            "stack_line_count": owner_class_lines[name],
            "allocation_cost": owner_class_cost[name],
            "percent_of_closure": 100 * owner_class_cost[name] / CLOSURE_ALLOCATION_COST,
        }
    region_rows = {}
    for name in ("scan_closure", "forced_worksheet_guard", "runner_other", "process_other"):
        region_rows[name] = {
            "stack_line_count": region_lines[name],
            "unique_stack_count": len(region_stacks[name]),
            "allocation_cost": regions[name],
            "percent_of_process": 100 * regions[name] / all_cost,
        }
    duplicate_rows = {}
    for name in ("whole_process", "scan_closure"):
        duplicate_rows[name] = {
            "stack_line_count": duplicate_lines[name],
            "unique_stack_count": len(duplicate_stacks[name]),
            "allocation_cost": duplicate_cost[name],
            "percent_of_process" if name == "whole_process" else "percent_of_closure": 100 * duplicate_cost[name] / (all_cost if name == "whole_process" else CLOSURE_ALLOCATION_COST),
        }
    return {
        "profile": pathname.parent.name,
        "input": descriptor(pathname),
        "stack_line_count": line_count,
        "whole_process_allocation_cost": all_cost,
        "regions": region_rows,
        "duplicate_check": {
            "selector": DUPLICATE_FN,
            "whole_process": duplicate_rows["whole_process"],
            "scan_closure": duplicate_rows["scan_closure"],
        },
        "nearest_closure_owners": owner_rows,
        "nearest_mce_start_vs_other_closure_owners": class_rows,
        "integer_reconciliation": {
            "whole_process_total": all_cost,
            "expected_whole_process_total": TOTAL_ALLOCATION_COST,
            "scan_closure_total": regions["scan_closure"],
            "expected_scan_closure_total": CLOSURE_ALLOCATION_COST,
            "sum_of_regions": sum(regions.values()),
            "pass": all_cost == TOTAL_ALLOCATION_COST and regions["scan_closure"] == CLOSURE_ALLOCATION_COST and sum(regions.values()) == all_cost,
        },
    }


def parse_leaf(pathname: Path, prior_row: dict[str, Any]) -> dict[str, Any]:
    counts: Counter[str] = Counter()
    periods: Counter[str] = Counter()
    total = 0
    sample_count = 0
    with pathname.open(encoding="utf-8", newline="") as handle:
        for line_number, raw in enumerate(handle, 1):
            text = raw[:-1] if raw.endswith("\n") else raw
            if text.endswith("\r"):
                text = text[:-1]
            match = LEAF_ROW.fullmatch(text)
            require(match is not None, f"unparsed perf leaf line {pathname}:{line_number}")
            period = int(match.group("period"))
            require(period > 0 and match.group("event") == "cycles:u", f"invalid perf leaf event/period: {pathname}:{line_number}")
            symbol = match.group("symbol").strip()
            group = next((name for name, selector in LEAF_SELECTORS.items() if symbol == selector), "other")
            counts[group] += 1
            periods[group] += period
            total += period
            sample_count += 1
    child = pathname.parent.name
    expected = EXPECTED_LEAF[child]
    require(sample_count == expected["sample_count"] and total == expected["total_period"], f"perf leaf population changed: {child}")
    require(prior_row["sample_count"] == sample_count and prior_row["total_period"] == total, f"0892 leaf receipt differs: {child}")
    require(prior_row["lost_samples"] == 0, f"0892 retained lost samples: {child}")
    require(prior_row.get("input_sha256", {}).get("leaf-samples.txt") == sha_file(pathname), f"0892 leaf input custody differs: {child}")
    groups = {}
    for name in (*LEAF_SELECTORS.keys(), "other"):
        groups[name] = {
            "sample_count": counts[name],
            "period": periods[name],
            "percent_of_total": 100 * periods[name] / total,
        }
    return {
        "child": child,
        "input": descriptor(pathname),
        "sample_count": sample_count,
        "total_period": total,
        "groups": groups,
        "reconciliation": {
            "expected_sample_count": expected["sample_count"],
            "expected_total_period": expected["total_period"],
            "sum_group_period": sum(periods.values()),
            "sum_group_sample_count": sum(counts.values()),
            "pass": sum(periods.values()) == total and sum(counts.values()) == sample_count,
        },
        "scope": "Disjoint sampled-IP leaf symbols from retained cycles:u profiles; native sample evidence only, no latency estimate.",
    }


def build_result() -> dict[str, Any]:
    source_identity = validate_source_identity()
    evidence, prior_alloc, prior_perf = validate_prior_evidence()
    inputs: dict[str, dict[str, Any]] = {
        "source_identity": source_identity["receipt"],
        "callgrind_00_self": descriptor(PACKET / "callgrind-00-self.txt"),
        "callgrind_01_self": descriptor(PACKET / "callgrind-01-self.txt"),
        "callgrind_00_self_command": descriptor(PACKET / "callgrind-00-self-command.json"),
        "callgrind_01_self_command": descriptor(PACKET / "callgrind-01-self-command.json"),
        "callgrind_00_self_stderr": descriptor(PACKET / "callgrind-00-self.stderr"),
        "callgrind_01_self_stderr": descriptor(PACKET / "callgrind-01-self.stderr"),
        "callgrind_00_inclusive": descriptor(path("docs/performance/results/change-0892/callgrind/callgrind-00/inclusive-canonical.txt")),
        "callgrind_01_inclusive": descriptor(path("docs/performance/results/change-0892/callgrind/callgrind-01/inclusive-canonical.txt")),
        "callgrind_00_raw": descriptor(path("docs/performance/results/change-0892/callgrind/callgrind-00/callgrind.out")),
        "callgrind_01_raw": descriptor(path("docs/performance/results/change-0892/callgrind/callgrind-01/callgrind.out")),
    }
    for ordinal in range(3):
        inputs[f"leaf_child_{ordinal:02d}"] = descriptor(path(f"docs/performance/results/change-0892/capture/child-{ordinal:02d}/leaf-samples.txt"))
        inputs[f"leaf_child_{ordinal:02d}_self_symbols"] = descriptor(path(f"docs/performance/results/change-0892/capture/child-{ordinal:02d}/self-symbols.txt"))
    for name, item in evidence.items():
        inputs[name] = item

    callgrind_prior_rows = {row["child"]: row for row in prior_perf["callgrind_analysis"]["rows"]}
    callgrind_profiles = []
    for profile, ordinal in (("callgrind-00", 0), ("callgrind-01", 1)):
        callgrind_profiles.append(analyze_callgrind(
            profile,
            PACKET / f"{profile}-self.txt",
            PACKET / f"{profile}-self-command.json",
            PACKET / f"{profile}-self.stderr",
            path(f"docs/performance/results/change-0892/callgrind/{profile}/inclusive-canonical.txt"),
            callgrind_prior_rows[profile],
        ))

    prior_leaf_rows = {row["child"]: row for row in prior_perf["leaf_analysis"]["rows"]}
    leaf_profiles = [
        parse_leaf(path(f"docs/performance/results/change-0892/capture/{child}/leaf-samples.txt"), prior_leaf_rows[child])
        for child in ("child-00", "child-01", "child-02")
    ]

    prior_profiles = {row["profile"]: row for row in prior_alloc["allocation_analysis"]["profiles"]}
    stack_profiles = [
        parse_stack(path(f"docs/performance/results/change-0895/export-v1/runs/{ordinal:02d}-baseline/stacks.txt"), prior_profiles[profile])
        for profile, ordinal in (("00-baseline", 0), ("03-baseline", 3))
    ]
    require(stack_profiles[0]["regions"] == stack_profiles[1]["regions"], "repeated baseline region accounting differs")
    return {
        "schema": "litchi.duplicate-attribution-analysis-0899.v1",
        "status": "pass",
        "scope": {
            "mode": "offline-only retained-artifact analysis",
            "source_evaluation_head": EVALUATION_HEAD,
            "source_baseline_revision": BASELINE_REVISION,
            "native_leaf_periods": True,
            "guest_instruction_callgrind": True,
            "heaptrack_whole_process_stacks": True,
            "native_latency_claim": None,
            "performance_claim": None,
            "adoption_decision": None,
            "interpretation": "Disjoint self/leaf rows are reconciled within their own profiler populations. Inclusive guest rows overlap and are context only; none is native time or end-to-end latency.",
        },
        "source_identity": source_identity,
        "inputs": inputs,
        "helpers": {
            "reducer": descriptor(Path(__file__)),
            "prior_helpers": {name: item for name, item in evidence.items() if name.endswith("helper")},
        },
        "callgrind": {
            "profiles": callgrind_profiles,
            "self_rows": {
                "disjoint": True,
                "unit": "Ir",
                "duplicate_selector": DUPLICATE_FN,
                "combined_percent_definition": "combined exclusive duplicate-checker self Ir / PROGRAM TOTALS Ir",
            },
        },
        "leaf_periods": {
            "profiles": leaf_profiles,
            "groups": list(LEAF_SELECTORS) + ["other"],
            "disjoint": True,
            "event": "cycles:u",
        },
        "heaptrack_stacks": {
            "profiles": stack_profiles,
            "region_semantics": "0897 region_for: scan_closure, forced_worksheet_guard, runner_other, process_other",
            "nearest_owner_semantics": "rightmost litchi_ooxml_common::mce::codec frame; mce_start is the exact start owner, other_codec_owner covers other codec frames, and no_codec_owner is retained separately",
            "integer_reconciliation": {
                "expected_whole_process_total": TOTAL_ALLOCATION_COST,
                "expected_scan_closure_total": CLOSURE_ALLOCATION_COST,
                "repeated_baseline_region_counts_identical": True,
            },
        },
        "prior_evidence": {
            "0895_verification": evidence["0895_verification"],
            "0897_analysis": evidence["0897_analysis"],
            "0892_callgrind_analysis": evidence["0892_callgrind_analysis"],
            "0892_leaf_analysis": evidence["0892_leaf_analysis"],
            "0899_allocation_custody_replay": evidence["0899_allocation_custody_replay"],
            "0899_profile_content_replay": evidence["0899_profile_content_replay"],
        },
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
