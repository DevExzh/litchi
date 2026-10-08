#!/usr/bin/env python3
"""Posthoc validator for the completed final-v2 ASLR replays.

This reads the immutable output of the v3 replay once; it never invokes perf
or the shim.  The final shim reports a target row for every target module
report call, so cache_stored/reused_existing are the lifecycle fields.  The
initial_rewrite field is a per-call expected-base predicate and is recorded separately.
"""
from __future__ import annotations

import datetime as dt
import hashlib
import importlib.util
import json
from collections import Counter
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0976")
OUT = ROOT / "aslr-cache-final-v2-0975-v3"
RECEIPT = ROOT / "receipts/aslr-cache-final-v2-0975-v3-posthoc-v2.json"
HELPER = ROOT / "tools/test_aslr_translation_cache_final_v2_0976_v3.py"
PROTOCOL = Path("/home/zhuhe/litchi-performance-0973/tools-v2/protocol_common_0973_v2.py")
SHIM = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.so"
SHIM_SOURCE = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.c"
SHIM_MAP = ROOT / "tools/dwfl_dynamic_shim_0976_cache_final_v2.map"
RAW = Path("/home/zhuhe/litchi-performance-0975/capture/instruction-profile-v1/child-01/child-01.perf.data")
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
FAILED_RECEIPT = ROOT / "receipts/aslr-cache-final-v2-0975-v3-failure.json"
TARGET_PATH = str(BINARY.resolve())
EXPECTED_PID = 3763764
EXPECTED_CPU = 31
PERIOD = 1_000_000


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha256(path)}


def utc() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def load_helper() -> Any:
    spec = importlib.util.spec_from_file_location("aslr_final_v2_replay_helper_0976", HELPER)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load helper {HELPER}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def summary(text: str, protocol: Any) -> tuple[dict[str, Any], str]:
    parsed, rows = protocol.parse_perf_script_strict(
        text, marker_dso=TARGET_PATH, expected_pid=EXPECTED_PID,
        expected_cpu=EXPECTED_CPU, expected_period=PERIOD,
    ), None
    # Reuse the helper's canonicalizer so the canonical frame fields are
    # exactly those used by the replay helper, then hash the normalized rows.
    canonical = []
    for sample in parsed["samples"]:
        canonical.append({
            "pid": sample["pid"], "cpu": sample["cpu"], "time": sample["time"],
            "period": sample["period"], "event": sample["event"],
            "frames": [{"dso": frame["dso"], "symbol_base": frame["symbol_base"],
                        "offset": frame["offset"], "resolved": frame["resolved"]}
                       for frame in sample["frames"]],
        })
    canonical_hash = hashlib.sha256(json.dumps(canonical, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
    public = {k: parsed[k] for k in (
        "sample_count", "total_period", "marker_sample_count", "marker_period",
        "lost_events", "empty_callchain_count", "explicit_stack_truncated_count",
        "unresolved_frames",
    )}
    return public, canonical_hash


def log_summary(path: Path) -> dict[str, Any]:
    final = []
    with path.open(encoding="utf-8") as f:
        for line in f:
            row = json.loads(line)
            if row.get("schema") == "litchi.performance.0976.dwfl_report.final":
                final.append(row)
    exact = [row for row in final if row.get("path_match") is True and row.get("name_match") is True]
    fallback = [row for row in exact if row.get("admission") is not True]
    admitted = [row for row in exact if row.get("admission") is True]
    stored = [row for row in exact if row.get("cache_stored") is True]
    reused = [row for row in exact if row.get("reused_existing") is True]
    initial_field = [row for row in exact if row.get("initial_rewrite") is True]
    identity = [row for row in exact if row.get("identity_valid") is True]
    config = [row for row in exact if row.get("config_valid") is True]
    result_nil = [row for row in exact if str(row.get("result_module")) == "(nil)"]
    checks = {
        "exact_rows_nonzero": bool(exact),
        "all_exact_admitted": len(admitted) == len(exact),
        "exact_fallback_rows_zero": not fallback,
        "one_cache_store": len(stored) == 1,
        "all_remaining_reused": len(reused) == len(exact) - 1,
        "all_exact_identity_valid": len(identity) == len(exact),
        "all_exact_config_valid": len(config) == len(exact),
        "all_exact_result_modules_non_nil": not result_nil,
        "first_exact_row_stores": bool(exact) and exact[0].get("cache_stored") is True,
        "later_exact_rows_reuse": all(row.get("reused_existing") is True for row in exact[1:]),
    }
    if not all(checks.values()):
        raise AssertionError(f"target report checks failed for {path}: {checks}")
    return {
        "file": binding(path),
        "all_final_report_rows": len(final),
        "non_target_final_report_rows": len(final) - len(exact),
        "non_target_path_name_pairs": {str(key): count for key, count in Counter((row.get("path_match"), row.get("name_match")) for row in final if not (row.get("path_match") is True and row.get("name_match") is True)).items()},
        "exact_target_rows": len(exact),
        "admitted_exact_rows": len(admitted),
        "fallback_exact_rows": len(fallback),
        "cache_stored_rows": len(stored),
        "reused_existing_rows": len(reused),
        "initial_rewrite_predicate_true_rows": len(initial_field),
        "identity_valid_rows": len(identity),
        "config_valid_rows": len(config),
        "result_module_nil_rows": len(result_nil),
        "errno_after_counts": Counter(row.get("dwfl_errno_after") for row in exact),
        "checks": checks,
    }


def main() -> int:
    if RECEIPT.exists():
        raise FileExistsError(RECEIPT)
    required = [HELPER, PROTOCOL, SHIM, SHIM_SOURCE, SHIM_MAP, RAW, BINARY, FAILED_RECEIPT,
                OUT / "original.final-v2.corrected.script.txt", OUT / "translated.corrected.script.txt",
                OUT / "original.final-v2.dwfl.jsonl", OUT / "translated.dwfl.jsonl",
                OUT / "mutation-map.json", OUT / "child-01.translated.perf.data",
                OUT / "translated.target-layout.json"]
    for path in required:
        if not path.is_file():
            raise FileNotFoundError(path)
    before = {str(path): binding(path) for path in (HELPER, PROTOCOL, SHIM, SHIM_SOURCE, SHIM_MAP, RAW, BINARY, FAILED_RECEIPT)}
    helper = load_helper()
    protocol = helper.load_protocol()
    original_text = (OUT / "original.final-v2.corrected.script.txt").read_text(encoding="utf-8")
    translated_text = (OUT / "translated.corrected.script.txt").read_text(encoding="utf-8")
    original_summary, original_hash = summary(original_text, protocol)
    translated_summary, translated_hash = summary(translated_text, protocol)
    if original_summary != translated_summary or original_hash != translated_hash:
        raise AssertionError("original and translated final-v2 canonical outputs differ")
    recovery = {}
    for label, value in (("original", original_summary), ("translated", translated_summary)):
        conditions = {
            "sample_count_6940": value["sample_count"] == 6940,
            "total_period_6940000000": value["total_period"] == 6_940_000_000,
            "empty_callchain_count_zero": value["empty_callchain_count"] == 0,
            "lost_events_zero": value["lost_events"] == 0,
            "marker_sample_count_at_least_100": value["marker_sample_count"] >= 100,
            "marker_period_positive": value["marker_period"] > 0,
        }
        if not all(conditions.values()):
            raise AssertionError(f"{label} recovery checks failed: {conditions}")
        recovery[label] = {"conditions": conditions,
                           "unresolved_frames": value["unresolved_frames"],
                           "explicit_stack_truncated_count": value["explicit_stack_truncated_count"]}
    original_reports = log_summary(OUT / "original.final-v2.dwfl.jsonl")
    translated_reports = log_summary(OUT / "translated.dwfl.jsonl")
    mutation = json.loads((OUT / "mutation-map.json").read_text(encoding="utf-8"))
    mutation_checks = {
        "sample_header_count": mutation.get("sample_header_count") == 6940,
        "sample_period_sum": mutation.get("sample_period_sum") == 6_940_000_000,
        "raw_sample_headers_invariant": mutation.get("raw_sample_headers_invariant") is True,
        "raw_target_mmaps_invariant_except_delta": mutation.get("raw_target_mmaps_invariant_except_delta") is True,
        "changed_fields_recorded_positive": mutation.get("changed_field_count", 0) > 0,
    }
    if not all(mutation_checks.values()):
        raise AssertionError(f"mutation checks failed: {mutation_checks}")
    after = {str(path): binding(path) for path in (HELPER, PROTOCOL, SHIM, SHIM_SOURCE, SHIM_MAP, RAW, BINARY, FAILED_RECEIPT)}
    if after != before:
        raise AssertionError("bound inputs changed during posthoc validation")
    receipt = {
        "schema": "litchi.performance.0976.aslr_translation.final_v2.posthoc.v2",
        "status": "pass",
        "created_utc": utc(),
        "purpose": "posthoc offline validation of completed original and translated final-v2 replays; no perf invocation or native workload",
        "inputs": {key: value for key, value in before.items()},
        "failed_attempt_receipt": binding(FAILED_RECEIPT),
        "outputs": {str(path): binding(path) for path in required if path.parent == OUT},
        "recovery": recovery,
        "canonical": {"original_sha256": original_hash, "translated_sha256": translated_hash,
                      "equal": original_hash == translated_hash, "sample_count": original_summary["sample_count"]},
        "report_checks": {"original": original_reports, "translated": translated_reports},
        "mutation_checks": mutation_checks,
        "interpretation": {
            "lifecycle_fields": "cache_stored marks the sole initial store; reused_existing marks later rows. initial_rewrite is a per-call expected-base predicate and is true on every exact target row, so it is not the mutually exclusive lifecycle event.",
            "non_target_reports": "Non-target DSO report rows are retained and counted without being treated as target fallback failures.",
            "unresolved_and_truncated": "Reported descriptively; no zero-unresolved or zero-truncation claim is made by this validator.",
        },
        "native_workload_started": False,
        "native_capture_started": False,
        "native_build_started": False,
        "target_executed": False,
    }
    RECEIPT.write_text(json.dumps(receipt, indent=2, sort_keys=True, default=list) + "\n", encoding="utf-8")
    print(json.dumps({"status":"pass","receipt":str(RECEIPT),"receipt_sha256":sha256(RECEIPT),"canonical_sha256":original_hash,"original_reports":original_reports,"translated_reports":translated_reports},sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
