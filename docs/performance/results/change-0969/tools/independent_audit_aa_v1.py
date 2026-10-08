#!/usr/bin/env python3
"""Independent, read-only audit and reduction of the completed 0969 DOC A/A lane.

This helper deliberately does not import the frozen reducer.  It validates raw
packet sidecars, command identities, executable/fixture/source custody, the
twenty correctness gates and twelve-field semantic oracles, then recomputes
the sealed p50/p95 paired bootstrap from the elapsed samples.
"""
from __future__ import annotations

import hashlib
import json
import math
import random
import sys
from pathlib import Path
from typing import Any


ROOT = Path("/home/zhuhe/litchi-performance-0969")
PLAN_PATH = ROOT / "plans/plan-0969-aa-v6.json"
GO_PATH = ROOT / "receipts/root-go-aa-v1.json"
CAPTURE = ROOT / "capture/aa-v4"
AA_DIR = CAPTURE / "aa"
READINESS_PATH = CAPTURE / "aa-readiness-v4.json"
STAGE_PATH = CAPTURE / "capture-stage-v4.json"
TERMINAL_PATH = ROOT / "receipts/root-terminal-aa-v1.json"
OUT_PATH = ROOT / "analysis/aa-independent-audit-v1.json"

SEED = 969004
RESAMPLES = 10000
CI_LIMIT = 1.05
CASES = ("docnohf", "docfloat")
EDITS = (1, 2, 8, 32)
RETENTIONS = ("zero", "default", "release")
RETENTION_REPORT = {"zero": "zero", "default": "default_8mib", "release": "release_8mib"}
RETENTION_CEILING = {"zero": 0, "default": 8388608, "release": 8388608}
CORRECTNESS = (
    "all_stages_change_text", "exact_cross_route_outputs", "final_paragraph_semantic_readback",
    "final_public_reader_reopen", "exact_noop_source", "noop_shares_source_allocation",
    "forward_patch_exact_output", "inverse_patch_exact_source", "foreign_patch_refused",
    "stale_patch_refused", "same_value_preserves_retention", "operation_limit_refusal_atomic",
    "retained_capacities_within_ceiling", "zero_route_clears_retention",
    "release_route_clears_every_stage", "untouched_streams_preserved", "directory_metadata_preserved",
    "semantic_oracles_passed", "semantic_change_count_matches_edit_count",
    "allocator_retained_matches_output_capacity",
)
ORACLE = (
    "oracle_ok", "source_structurally_valid", "expected_structurally_valid",
    "output_structurally_valid", "semantic_reopen_ok", "semantic_directory_metadata_matches_expected",
    "unchanged_stream_bytes_match_expected", "unchanged_source_output_stream_bytes_match",
    "output_stream_paths_match_expected", "output_stream_bytes_match_expected",
    "root_clsid_preserved", "storage_clsids_preserved",
)
REPORT_KEYS = {
    "allocation", "allocator_instrumented", "allocator_timing_claim", "alternating_texts", "case",
    "change", "correctness", "edits", "elapsed_ns", "exact_output_equal_reference", "input",
    "output_bytes", "output_capacity", "output_sha256", "reference_output_bytes",
    "reference_output_sha256", "region_definition", "retention", "retention_ceiling_bytes",
    "route_outputs", "samples", "samples_requested", "schema", "schema_version", "selected_paragraph",
    "source_bytes", "source_sha256", "timing_claim", "warmups",
}
SAMPLE_BOOL = (
    "exact_output_equal_reference", "final_paragraph_semantic_readback", "final_public_reader_reopen",
    "retained_capacities_within_ceiling", "allocator_retained_matches_output_capacity",
)
ROUTE_KEYS = {
    "final_retained_before_commit", "oracle", "output_bytes", "output_capacity", "output_sha256",
    "retention", "retention_ceiling_bytes", "semantic_change_count", "stages",
}


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def percentile(values: list[float], quantile: float) -> float:
    ordered = sorted(float(value) for value in values)
    position = (len(ordered) - 1) * quantile
    low = math.floor(position)
    high = math.ceil(position)
    if low == high:
        return ordered[low]
    return ordered[low] + (ordered[high] - ordered[low]) * (position - low)


def bootstrap(values: list[float]) -> dict[str, float]:
    rng = random.Random(SEED)
    medians: list[float] = []
    for _ in range(RESAMPLES):
        sample = [values[rng.randrange(len(values))] for _ in values]
        medians.append(percentile(sample, 0.5))
    return {
        "median": percentile(values, 0.5),
        "lower": percentile(medians, 0.025),
        "upper": percentile(medians, 0.975),
    }


def expected_cells() -> list[str]:
    return [f"{case}/e{edits}/{retention}" for case in CASES for edits in EDITS for retention in RETENTIONS]


def file_set_digest(paths: list[Path], root: Path) -> str:
    """Digest sorted `sha256 relative-path` lines; this is this audit's custody seal."""
    rows = []
    for path in paths:
        rows.append(f"{sha256_file(path)}  {path.relative_to(root).as_posix()}\n")
    return hashlib.sha256("".join(sorted(rows)).encode()).hexdigest()


def main() -> int:
    errors: list[dict[str, str]] = []
    observations: dict[str, Any] = {}

    def fail(label: str, detail: str) -> None:
        errors.append({"label": label, "detail": detail})

    if not PLAN_PATH.is_file() or not GO_PATH.is_file():
        fail("sealed_inputs", "plan or root GO is missing")
        return write_result(errors, observations, {}, {}, {})
    plan = read_json(PLAN_PATH)
    go = read_json(GO_PATH)
    plan_sha = sha256_file(PLAN_PATH)
    go_sha = sha256_file(GO_PATH)
    observations["sealed"] = {
        "plan": {"path": str(PLAN_PATH), "sha256": plan_sha, "expected_sha256": "e853030180d7464c06cf1ed5d525957ce2d50514b8337a1b168392374d00b5df"},
        "root_go": {"path": str(GO_PATH), "sha256": go_sha, "expected_sha256": "f3584215eade930b7748e1e3b7ea5a9b17bf331c037d2cda35f43afe8c93c7fa"},
    }
    if plan_sha != observations["sealed"]["plan"]["expected_sha256"]:
        fail("plan_sha256", f"observed {plan_sha}")
    if go_sha != observations["sealed"]["root_go"]["expected_sha256"]:
        fail("root_go_sha256", f"observed {go_sha}")
    if not go.get("authorized") or go.get("authorized_stage") != "baseline_only_aa":
        fail("root_go_authorization", repr({"authorized": go.get("authorized"), "stage": go.get("authorized_stage")}))
    if go.get("plan", {}).get("path") != str(PLAN_PATH) or go.get("plan", {}).get("sha256") != plan_sha:
        fail("root_go_plan_binding", repr(go.get("plan")))
    if plan.get("no_build_or_workload") is not True or plan.get("status") != "aa_authorized":
        fail("plan_scope", repr({"no_build_or_workload": plan.get("no_build_or_workload"), "status": plan.get("status")}))

    # Recheck every path/hash descriptor carried by the sealed GO and plan.
    custody: dict[str, Any] = {"checked": [], "missing": [], "mismatch": []}

    def descriptor(label: str, row: Any, expected: Any = None) -> None:
        if not isinstance(row, dict) or not isinstance(row.get("path"), str):
            fail(label, "descriptor is missing a path")
            return
        path = Path(row["path"])
        if not path.is_file():
            custody["missing"].append(label)
            fail(label, f"missing {path}")
            return
        actual = sha256_file(path)
        item = {"label": label, "path": str(path), "sha256": actual, "expected_sha256": row.get("sha256")}
        custody["checked"].append(item)
        if row.get("sha256") != actual:
            custody["mismatch"].append(label)
            fail(label, f"SHA observed {actual}, expected {row.get('sha256')}")
        if expected and (row.get("path") != expected.get("path") or row.get("sha256") != expected.get("sha256")):
            fail(label + "_plan_binding", repr({"go": row, "plan": expected}))

    descriptor("protocol_document", plan.get("protocol_document"))
    descriptor("plan_baseline_reuse_review", plan.get("baseline_reuse_review"))
    descriptor("plan_design_review", plan.get("design_review"))
    for label, row in sorted((go.get("helper_bindings") or {}).items()):
        descriptor("go_helper." + label, row, (plan.get("helpers") or {}).get(label))
    for label, row in sorted((go.get("review_bindings") or {}).items()):
        descriptor("go_review." + label, row, None)
    for label in ("independent_protocol_review", "build_equivalence_audit"):
        descriptor("go_" + label, go.get(label))
    source_plan = plan.get("source") or {}
    source_go = go.get("source_bindings") or {}
    # The GO schema names these source descriptors `manifest`, whereas plan
    # source rows use `manifest`/`manifest_sha256`.  Candidate source is an
    # explicitly unexecuted historical placeholder; only its manifest custody
    # is checked here, while the baseline tree is checked below.
    for label in ("baseline", "candidate"):
        row = source_go.get(label) or {}
        source_row = source_plan.get(label) or {}
        if row.get("manifest") != source_row.get("manifest") or row.get("manifest_sha256") != source_row.get("manifest_sha256"):
            fail("go_source." + label + "_binding", repr({"go": row, "plan": source_row}))
        manifest_row = {"path": row.get("manifest"), "sha256": row.get("manifest_sha256")}
        descriptor("go_source." + label, manifest_row, {"path": source_row.get("manifest"), "sha256": source_row.get("manifest_sha256")})
    observations["custody_descriptors"] = custody

    # Executable and fixture custody, including actual bytes and SHA-256.
    binaries: dict[str, Any] = {}
    for label in ("native_baseline", "allocator_baseline"):
        row = (plan.get("binaries") or {}).get(label) or {}
        path = Path(row.get("path", ""))
        if not path.is_file():
            fail("binary." + label, f"missing {path}")
            continue
        actual = sha256_file(path)
        binaries[label] = {"path": str(path), "bytes": path.stat().st_size, "sha256": actual, "expected_sha256": row.get("sha256")}
        if actual != row.get("sha256"):
            fail("binary." + label, f"SHA observed {actual}, expected {row.get('sha256')}")
    observations["binaries"] = binaries
    fixtures: dict[str, Any] = {}
    for case, row in sorted((plan.get("fixtures") or {}).items()):
        path = Path(row.get("path", ""))
        if not path.is_file():
            fail("fixture." + case, f"missing {path}")
            continue
        actual = sha256_file(path)
        fixtures[case] = {"path": str(path), "bytes": path.stat().st_size, "sha256": actual, "expected_bytes": row.get("bytes"), "expected_sha256": row.get("sha256")}
        if path.stat().st_size != row.get("bytes") or actual != row.get("sha256"):
            fail("fixture." + case, repr(fixtures[case]))
    observations["fixtures"] = fixtures

    # Validate the immutable source export through its 924-entry file manifest.
    source_manifest_path = Path(source_plan.get("baseline", {}).get("manifest", ""))
    source_manifest = read_json(source_manifest_path) if source_manifest_path.is_file() else {}
    source_file_manifest_path = source_manifest_path.parent / "source-file-manifest-v1.json"
    source_file_manifest = read_json(source_file_manifest_path) if source_file_manifest_path.is_file() else {}
    source_audit: dict[str, Any] = {"manifest": {}, "file_manifest": {}, "roots": {}}
    if source_manifest_path.is_file():
        source_audit["manifest"] = {"path": str(source_manifest_path), "sha256": sha256_file(source_manifest_path), "expected_sha256": source_plan.get("baseline", {}).get("manifest_sha256"), "schema": source_manifest.get("schema")}
        if source_audit["manifest"]["sha256"] != source_audit["manifest"]["expected_sha256"]:
            fail("source_manifest_sha256", repr(source_audit["manifest"]))
    else:
        fail("source_manifest", f"missing {source_manifest_path}")
    entries = source_file_manifest.get("entries") if isinstance(source_file_manifest, dict) else None
    if not isinstance(entries, list) or len(entries) != 924:
        fail("source_file_manifest_entries", f"expected 924, observed {len(entries) if isinstance(entries, list) else None}")
        entries = []
    source_audit["file_manifest"] = {"path": str(source_file_manifest_path), "sha256": sha256_file(source_file_manifest_path) if source_file_manifest_path.is_file() else None, "schema": source_file_manifest.get("schema"), "entries": len(entries)}
    for role in ("baseline", "candidate"):
        binding = source_plan.get(role) or {}
        root = Path(binding.get("root", ""))
        expected_paths = {str(e.get("path")): e for e in entries if isinstance(e, dict) and isinstance(e.get("path"), str)}
        actual_paths = {p.relative_to(root).as_posix(): p for p in root.rglob("*") if p.is_file()} if root.is_dir() else {}
        if set(actual_paths) != set(expected_paths):
            fail("source." + role + ".path_set", f"expected {len(expected_paths)}, observed {len(actual_paths)}")
        mismatches = 0
        mismatch_paths: list[str] = []
        total = 0
        for rel, entry in expected_paths.items():
            path = actual_paths.get(rel)
            if path is None:
                mismatches += 1
                continue
            actual_size = path.stat().st_size
            actual_hash = sha256_file(path)
            expected_hash = entry.get("baseline_sha256" if role == "baseline" else "candidate_sha256")
            total += actual_size
            if actual_size != entry.get("bytes") or actual_hash != expected_hash:
                mismatches += 1
                mismatch_paths.append(rel)
        exp = source_manifest.get("exports", {})
        expected_bytes = exp.get(role + "_bytes")
        expected_count = exp.get(role + "_file_count")
        source_audit["roots"][role] = {"root": str(root), "file_count": len(actual_paths), "bytes": total, "expected_file_count": expected_count, "expected_bytes": expected_bytes, "tree_sha256": binding.get("tree_sha256"), "file_entry_mismatches": mismatches, "mismatch_paths": mismatch_paths}
        # The root GO explicitly makes no candidate execution/custody claim;
        # later candidate edits are retained as a custody observation, not an
        # A/A invalidation.  Baseline entries remain a hard admission gate.
        if role == "baseline" and (mismatches or len(actual_paths) != expected_count or total != expected_bytes):
            fail("source." + role + ".entries", repr(source_audit["roots"][role]))
    observations["source"] = source_audit

    # Raw packet and sidecar validation.
    packet_paths = sorted(AA_DIR.glob("*.json"))
    all_raw_paths = sorted(p for p in AA_DIR.iterdir() if p.is_file()) if AA_DIR.is_dir() else []
    packet_summary: dict[str, Any] = {
        "json_count": len(packet_paths), "raw_file_count": len(all_raw_paths),
        "raw_file_set_digest": file_set_digest(all_raw_paths, AA_DIR) if all_raw_paths else None,
        "json_bytes": sum(p.stat().st_size for p in packet_paths),
        "sidecar_bytes": 0, "stdout_bytes": 0, "stderr_bytes": 0, "rss_bytes": 0,
        "stdout_report_equal": 0, "stderr_empty": 0,
        "rss_matches": 0, "command_matches": 0, "valid_packets": 0,
    }
    expected_count = 24 * 12 * 2
    if len(packet_paths) != expected_count:
        fail("packet_count", f"expected {expected_count}, observed {len(packet_paths)}")
    packets: list[dict[str, Any]] = []
    groups: dict[tuple[str, str, int, str], dict[str, Any]] = {}
    expected_binary = (plan.get("binaries") or {}).get("native_baseline") or {}
    expected_source_sha = (source_plan.get("baseline") or {}).get("manifest_sha256")
    expected_cells_set = set(expected_cells())
    for path in packet_paths:
        try:
            packet = read_json(path)
        except Exception as exc:
            fail("packet_json." + path.name, repr(exc))
            continue
        packets.append(packet)
        pid = packet.get("packet_id")
        context = packet.get("context") or {}
        cell = context.get("cell") or {}
        case, edits, retention = cell.get("case"), cell.get("edits"), cell.get("retention")
        cell_id = cell.get("cell_id")
        pair, order, position = context.get("pair"), context.get("aa_order"), context.get("aa_position")
        expected_pid = f"aa-{case}-e{edits}-{retention}-p{pair:02d}-{position}" if isinstance(pair, int) and position in {"first", "second"} else None
        base_label = str(pid or path.name)
        if path.stem != pid or expected_pid != pid:
            fail("packet_identity." + base_label, repr({"file": path.stem, "packet_id": pid, "expected": expected_pid}))
        if packet.get("schema") != "litchi.performance.0969.runner.v4" or packet.get("role") != "aa_baseline":
            fail("packet_schema_role." + base_label, repr({"schema": packet.get("schema"), "role": packet.get("role")}))
        if packet.get("validation_error") is not None:
            fail("packet_validation_error." + base_label, repr(packet.get("validation_error")))
        terminal = packet.get("terminal") or {}
        if terminal.get("returncode") != 0 or terminal.get("timed_out") is not False or terminal.get("launch_error") is not None:
            fail("packet_terminal." + base_label, repr(terminal))
        if not (case in CASES and edits in EDITS and retention in RETENTIONS and cell_id == f"{case}/e{edits}/{retention}"):
            fail("packet_cell." + base_label, repr(cell))
        expected_order = "AB" if isinstance(pair, int) and 1 <= pair <= 6 else "BA" if isinstance(pair, int) and 7 <= pair <= 12 else None
        if expected_order != order or pair not in range(1, 13) or position not in {"first", "second"}:
            fail("packet_aa_context." + base_label, repr({"pair": pair, "order": order, "position": position}))
        if context.get("lane") != "aa" or context.get("variant") != "baseline" or context.get("binary_sha256") != expected_binary.get("sha256") or context.get("fixture_sha256") != (plan.get("fixtures", {}).get(case, {}) or {}).get("sha256") or context.get("source_manifest_sha256") != expected_source_sha:
            fail("packet_context_custody." + base_label, repr(context))
        binary = packet.get("binary") or {}
        if binary.get("path") != expected_binary.get("path") or binary.get("sha256") != expected_binary.get("sha256"):
            fail("packet_binary_custody." + base_label, repr(binary))

        # Sidecars are retained raw evidence and must be exactly the descriptors named by the packet.
        sidecars_ok = True
        for ext, section in (("stdout", "stdout"), ("stderr", "stderr"), ("rss", "rss")):
            desc = packet.get(section) or {}
            side = Path(desc.get("path", ""))
            expected_side = AA_DIR / f"{pid}.{ext}"
            if side.resolve() != expected_side.resolve() or not side.is_file():
                fail("sidecar_path." + base_label + "." + ext, repr(desc)); sidecars_ok = False; continue
            actual_bytes = side.stat().st_size
            packet_summary["sidecar_bytes"] += actual_bytes
            packet_summary[ext + "_bytes"] += actual_bytes
            if desc.get("bytes") != actual_bytes:
                fail("sidecar_bytes." + base_label + "." + ext, repr({"descriptor": desc.get("bytes"), "actual": actual_bytes})); sidecars_ok = False
            if ext == "rss":
                raw = side.read_text(encoding="utf-8").strip()
                if desc.get("kb") != packet.get("rss", {}).get("kb") or not raw.isdigit() or int(raw) != desc.get("kb"):
                    fail("rss_content." + base_label, repr({"raw": raw, "descriptor": desc})); sidecars_ok = False
            elif ext == "stderr":
                if side.read_bytes() != b"":
                    fail("stderr_nonempty." + base_label, "stderr is not empty"); sidecars_ok = False
                else:
                    packet_summary["stderr_empty"] += 1
        stdout_desc = packet.get("stdout") or {}
        stdout_path = Path(stdout_desc.get("path", ""))
        report = packet.get("report")
        try:
            stdout_report = json.loads(stdout_path.read_text(encoding="utf-8"))
        except Exception as exc:
            fail("stdout_json." + base_label, repr(exc)); stdout_report = None
        if stdout_report != report:
            fail("stdout_report_mismatch." + base_label, "stdout JSON differs from packet report")
        else:
            packet_summary["stdout_report_equal"] += 1
        if stdout_report is not None and stdout_report != report:
            sidecars_ok = False
        if (packet.get("rss") or {}).get("kb", 0) > 0:
            packet_summary["rss_matches"] += 1

        # Command line custody and report identity.
        cmd = packet.get("command")
        fixture = (plan.get("fixtures") or {}).get(case, {})
        expected_cmd = ["/usr/bin/time", "-f", "%M", "-o", str(AA_DIR / f"{pid}.rss"), str(expected_binary.get("path")), "--case", case, "--input", str(fixture.get("path")), "--edits", str(edits), "--retention", retention, "--warmup", "5", "--samples", "80"]
        if cmd == expected_cmd:
            packet_summary["command_matches"] += 1
        else:
            fail("command." + base_label, repr({"observed": cmd, "expected": expected_cmd}))

        # Report, all twenty correctness gates, all sample booleans, and all route oracles.
        valid_report = True
        if not isinstance(report, dict) or set(report) != REPORT_KEYS:
            fail("report_keys." + base_label, repr(sorted(set(report or {}) ^ REPORT_KEYS))); valid_report = False
        if valid_report:
            if report.get("case") != case or report.get("edits") != edits or report.get("retention") != RETENTION_REPORT.get(retention) or report.get("input") != fixture.get("path") or report.get("source_sha256") != fixture.get("sha256") or report.get("source_bytes") != fixture.get("bytes"):
                fail("report_identity." + base_label, "report identity/custody differs from packet context"); valid_report = False
            if report.get("schema") != "litchi.performance.0917.doc_retention.v1" or report.get("schema_version") != 1 or report.get("change") != "0917" or report.get("warmups") != 5 or report.get("samples_requested") != 80 or report.get("timing_claim") is not True or report.get("allocator_instrumented") is not False or report.get("allocator_timing_claim") is not False:
                fail("report_protocol." + base_label, "report protocol fields differ"); valid_report = False
            correctness = report.get("correctness")
            if set(correctness or {}) != set(CORRECTNESS) or any(type((correctness or {}).get(k)) is not bool or not (correctness or {}).get(k) for k in CORRECTNESS):
                fail("correctness_20." + base_label, repr(correctness)); valid_report = False
            elapsed = report.get("elapsed_ns")
            if not isinstance(elapsed, list) or len(elapsed) != 80 or any(type(v) is not int or v <= 0 for v in elapsed):
                fail("elapsed_samples." + base_label, "top elapsed_ns is not 80 positive integers"); valid_report = False
            samples = report.get("samples")
            if not isinstance(samples, list) or len(samples) != 80:
                fail("sample_count." + base_label, repr(len(samples) if isinstance(samples, list) else None)); valid_report = False
            else:
                for index, sample in enumerate(samples):
                    if not isinstance(sample, dict) or sample.get("index") != index or sample.get("elapsed_ns") != elapsed[index]:
                        fail("sample_identity." + base_label, f"sample {index} identity/elapsed mismatch"); valid_report = False; break
                    if any(type(sample.get(k)) is not bool or sample.get(k) is not True for k in SAMPLE_BOOL):
                        fail("sample_gates." + base_label, f"sample {index} has a false/nonboolean gate"); valid_report = False; break
                    if set((sample.get("oracle") or {})) != set(ORACLE) | {"failure_reasons"} or any(type((sample.get("oracle") or {}).get(k)) is not bool or not (sample.get("oracle") or {}).get(k) for k in ORACLE) or (sample.get("oracle") or {}).get("failure_reasons") != []:
                        fail("sample_oracle." + base_label, f"sample {index} oracle failure"); valid_report = False; break
                    if not isinstance(sample.get("stages"), list) or len(sample["stages"]) != 32:
                        fail("sample_stages." + base_label, f"sample {index} stage count"); valid_report = False; break
            routes = report.get("route_outputs")
            if not isinstance(routes, list) or len(routes) != 3:
                fail("route_count." + base_label, repr(len(routes) if isinstance(routes, list) else None)); valid_report = False
            else:
                seen_routes = []
                for route in routes:
                    if set(route) != ROUTE_KEYS:
                        fail("route_keys." + base_label, repr(sorted(set(route) ^ ROUTE_KEYS))); valid_report = False; continue
                    seen_routes.append(route.get("retention"))
                    if route.get("retention") not in {"zero", "default_8mib", "release_8mib"} or route.get("retention_ceiling_bytes") not in {0, 8388608} or route.get("semantic_change_count") != edits:
                        fail("route_identity." + base_label, repr(route)); valid_report = False
                    oracle = route.get("oracle") or {}
                    if set(oracle) != set(ORACLE) | {"failure_reasons"} or any(type(oracle.get(k)) is not bool or not oracle.get(k) for k in ORACLE) or oracle.get("failure_reasons") != []:
                        fail("route_oracle." + base_label, repr(oracle)); valid_report = False
                    if not isinstance(route.get("stages"), list) or len(route["stages"]) != 32:
                        fail("route_stages." + base_label, "route stages are not 32 entries"); valid_report = False
                if set(seen_routes) != {"zero", "default_8mib", "release_8mib"}:
                    fail("route_retention_coverage." + base_label, repr(seen_routes)); valid_report = False
        if valid_report:
            packet_summary["valid_packets"] += 1
        if cell_id in expected_cells_set and order in {"AB", "BA"} and pair in range(1, 13) and position in {"first", "second"}:
            key = (cell_id, order, pair, position)
            if key in groups:
                fail("duplicate_identity." + base_label, str(key))
            else:
                groups[key] = packet

    # Coverage and independent paired p50/p95 bootstrap reduction.
    expected_keys = {(cell, "AB" if pair <= 6 else "BA", pair, pos) for cell in expected_cells() for pair in range(1, 13) for pos in ("first", "second")}
    observed_keys = set(groups)
    missing = sorted(expected_keys - observed_keys)
    extra = sorted(observed_keys - expected_keys)
    if missing:
        fail("identity_coverage_missing", repr(missing[:8]) + (" ..." if len(missing) > 8 else ""))
    if extra:
        fail("identity_coverage_extra", repr(extra[:8]) + (" ..." if len(extra) > 8 else ""))
    cells_result: dict[str, Any] = {}
    fail_rows: list[dict[str, Any]] = []
    for cell in expected_cells():
        cell_result: dict[str, Any] = {}
        for order in ("AB", "BA"):
            order_result: dict[str, Any] = {}
            pairs = range(1, 7) if order == "AB" else range(7, 13)
            for metric, quantile in (("p50", 0.50), ("p95", 0.95)):
                ratios: list[float] = []
                complete = True
                for pair in pairs:
                    first = groups.get((cell, order, pair, "first"), {}).get("report", {}).get("elapsed_ns")
                    second = groups.get((cell, order, pair, "second"), {}).get("report", {}).get("elapsed_ns")
                    if not isinstance(first, list) or not isinstance(second, list) or len(first) != 80 or len(second) != 80:
                        complete = False; break
                    denominator = percentile(first, quantile)
                    numerator = percentile(second, quantile)
                    if denominator <= 0 or numerator <= 0:
                        complete = False; break
                    ratios.append(numerator / denominator)
                if not complete or len(ratios) != 6:
                    fail("reduction." + cell + "/" + order + "/" + metric, "paired samples unavailable")
                    continue
                interval = bootstrap(ratios)
                symmetric_upper = max(interval["upper"], 1.0 / interval["lower"] if interval["lower"] > 0 else math.inf)
                passed = symmetric_upper <= CI_LIMIT
                row = {**interval, "ratios": ratios, "symmetric_upper": symmetric_upper, "pass": passed}
                order_result[metric] = row
                if not passed:
                    fail_rows.append({"cell": cell, "order": order, "metric": metric, "symmetric_upper": symmetric_upper, "ratios": ratios})
            cell_result[order] = order_result
        cells_result[cell] = cell_result
    readiness_status = "incomplete" if errors or missing or extra else ("readiness_fail" if fail_rows else "pass")
    packet_summary["correctness_gate_instances"] = packet_summary["valid_packets"] * len(CORRECTNESS)
    packet_summary["sample_oracle_instances"] = packet_summary["valid_packets"] * 80
    packet_summary["route_oracle_instances"] = packet_summary["valid_packets"] * 3
    observations["packets"] = packet_summary
    observations["reduction"] = {"status": readiness_status, "interpretation": "uninterpretable" if readiness_status == "readiness_fail" else "no_decision" if readiness_status == "incomplete" else "ready_for_formal", "cells": cells_result, "fail_rows": fail_rows, "fail_cells": sorted({row["cell"] for row in fail_rows}), "bootstrap_seed": SEED, "bootstrap_resamples": RESAMPLES, "upper_ci_limit": CI_LIMIT, "pairs_per_cell": 12, "samples_per_child": 80, "warmups_per_child": 5}

    # Compare the independent reduction with the frozen receipt and capture-stage marker.
    frozen = read_json(READINESS_PATH) if READINESS_PATH.is_file() else {}
    stage = read_json(STAGE_PATH) if STAGE_PATH.is_file() else {}
    terminal = read_json(TERMINAL_PATH) if TERMINAL_PATH.is_file() else {}
    observations["frozen_outputs"] = {"readiness": {"path": str(READINESS_PATH), "sha256": sha256_file(READINESS_PATH) if READINESS_PATH.is_file() else None, "status": frozen.get("status")}, "capture_stage": {"path": str(STAGE_PATH), "sha256": sha256_file(STAGE_PATH) if STAGE_PATH.is_file() else None, "status": stage.get("status"), "aa_status": stage.get("aa_status"), "real_lanes_started": stage.get("real_lanes_started"), "formal_candidate_launch_permitted": stage.get("formal_candidate_launch_permitted")}}
    terminal_observation = {"path": str(TERMINAL_PATH), "sha256": sha256_file(TERMINAL_PATH) if TERMINAL_PATH.is_file() else None, "expected_sha256": "187494d87f4c0e3d7786dffcfb266094cea0faf7a625ac66d7d4e0acbaf71d9e", "schema": terminal.get("schema"), "status": terminal.get("status"), "exit_code": terminal.get("exit_code"), "packet_count": terminal.get("packet_count"), "runner_session_id": terminal.get("runner_session_id"), "no_retry": terminal.get("no_retry"), "candidate_native_children": terminal.get("candidate_native_children"), "candidate_allocator_children": terminal.get("candidate_allocator_children"), "xls_workload_children": terminal.get("xls_workload_children"), "binding_mismatches": []}
    for binding in terminal.get("bindings", []) if isinstance(terminal.get("bindings"), list) else []:
        path = Path(binding.get("path", ""))
        actual = sha256_file(path) if path.is_file() else None
        if actual != binding.get("sha256") or (path.stat().st_size if path.is_file() else None) != binding.get("bytes"):
            terminal_observation["binding_mismatches"].append({"path": str(path), "observed_sha256": actual, "observed_bytes": path.stat().st_size if path.is_file() else None, "binding": binding})
    observations["frozen_outputs"]["terminal"] = terminal_observation
    if terminal_observation["sha256"] != terminal_observation["expected_sha256"] or terminal.get("status") != "completed_aa_readiness_fail_stopped_before_candidate_formal" or terminal.get("exit_code") != 2 or terminal.get("packet_count") != expected_count or terminal.get("no_retry") is not True or terminal.get("candidate_native_children") != 0 or terminal.get("candidate_allocator_children") != 0 or terminal.get("xls_workload_children") != 0 or terminal_observation["binding_mismatches"]:
        fail("terminal_receipt", repr(terminal_observation))
    if frozen.get("status") != readiness_status and not errors:
        fail("frozen_status", repr({"independent": readiness_status, "frozen": frozen.get("status")}))
    if stage.get("aa_status") != frozen.get("status") or stage.get("real_lanes_started") is not False or stage.get("formal_candidate_launch_permitted") is not False:
        fail("capture_stage", repr(stage))
    frozen_mismatch = []
    for cell in expected_cells():
        for order in ("AB", "BA"):
            for metric in ("p50", "p95"):
                ours = cells_result.get(cell, {}).get(order, {}).get(metric, {})
                theirs = (frozen.get("cells", {}).get(cell, {}) or {}).get(order, {}).get(metric, {})
                for field in ("median", "lower", "upper", "symmetric_upper", "pass"):
                    if ours.get(field) != theirs.get(field):
                        frozen_mismatch.append({"cell": cell, "order": order, "metric": metric, "field": field, "independent": ours.get(field), "frozen": theirs.get(field)})
    observations["frozen_outputs"]["reduction_field_mismatches"] = frozen_mismatch
    if frozen_mismatch:
        fail("frozen_reduction_mismatch", repr(frozen_mismatch[:3]))
    return write_result(errors, observations, cells_result, fail_rows, packet_summary)


def write_result(errors: list[dict[str, str]], observations: dict[str, Any], cells: dict[str, Any], fail_rows: list[dict[str, Any]], packet_summary: dict[str, Any]) -> int:
    OUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    result = {
        "schema": "litchi.performance.0969.independent_aa_audit.v1",
        "helper": {"path": str(Path(__file__).resolve()), "sha256": sha256_file(Path(__file__).resolve())},
        "scope": {"read_only": True, "no_build_or_workload": True, "no_rerun": True, "raw_packet_root": str(AA_DIR)},
        "errors": errors,
        "error_count": len(errors),
        "complete_valid_packet_set": len(errors) == 0 and packet_summary.get("json_count") == 576 and packet_summary.get("valid_packets") == 576,
        "observations": observations,
    }
    OUT_PATH.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"path": str(OUT_PATH), "sha256": sha256_file(OUT_PATH), "error_count": len(errors), "status": (observations.get("reduction") or {}).get("status")}, sort_keys=True))
    return 0 if not errors else 2


if __name__ == "__main__":
    sys.exit(main())
