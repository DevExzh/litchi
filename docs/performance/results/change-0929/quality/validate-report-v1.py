#!/usr/bin/env python3
import hashlib
import json
import re
from pathlib import Path

BASE = Path("/tmp/litchi-performance-0929")
CAP = BASE / "capture-fp"
OUT = BASE / "quality" / "report-validation-v1.json"
failures = []

def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()

def check(condition, message):
    if not condition:
        failures.append(message)

report_path = CAP / "generic-read-at-report.json"
terminal_path = CAP / "terminal-receipt.json"
report = json.loads(report_path.read_text())
terminal = json.loads(terminal_path.read_text())
expected_n = 1000
check(report["configuration"]["samples_per_case"] == expected_n, "samples_per_case")
check(report["configuration"]["warmup_iterations_per_case"] == 0, "warmups")
check(report["configuration"]["cases"] == ["docx_source_backed_existing_section_layout_edit_save"], "case configuration")
check(len(report["results"]) == 1, "result count")
result = report["results"][0]
check(result["case"] == "docx_source_backed_existing_section_layout_edit_save", "result case")
elapsed = result["elapsed_ns"]
elapsed_samples = elapsed["samples"]
check(len(elapsed_samples) == expected_n, "elapsed sample length")
check(sorted(elapsed["sample_order"]) == list(range(expected_n)), "elapsed sample order")
layout = result["source"]["docx_section_layout"]
sample_lists = [
    "preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns",
    "publication_ns", "source_read_calls", "source_read_bytes",
    "operation_end_source_read_calls", "operation_end_source_read_bytes",
    "ordinary_payload_materializations",
    "source_cache_before_publication_after_commit_hits",
    "source_cache_before_publication_after_commit_cold_loads",
    "source_cache_before_publication_after_commit_successful_loads",
    "source_cache_before_publication_after_commit_retained_entries",
    "source_cache_before_publication_after_commit_retained_bytes",
    "output_sha256",
]
for key in sample_lists:
    check(len(layout[key]) == expected_n, f"layout.{key} length")
for key in ["read_calls", "read_bytes", "ordinary_payload_read_calls",
            "ordinary_payload_read_bytes", "max_in_flight_reads",
            "ordinary_payload_materializations"]:
    check(len(result["source"][key]) == expected_n, f"source.{key} length")
for key in ["accepted_bytes", "write_calls", "largest_write"]:
    check(len(result["operation_metrics"]["sink"][key]["values"]) == expected_n,
          f"operation sink {key} length")
for key in ["bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
            "bytes_4097_to_16384", "bytes_16385_to_65536",
            "bytes_over_65536"]:
    check(len(result["operation_metrics"]["sink"]["write_size_buckets"][key]["values"]) == expected_n,
          f"operation sink bucket {key} length")
phase_keys = ["preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns"]
check(all(sum(layout[k][i] for k in phase_keys) == elapsed_samples[i]
          for i in range(expected_n)), "phase sums")
fixed_values = {
    "read_calls": [202] * expected_n,
    "read_bytes": [8409488] * expected_n,
    "ordinary_payload_read_calls": [138] * expected_n,
    "ordinary_payload_read_bytes": [8397632] * expected_n,
    "max_in_flight_reads": [1] * expected_n,
    "ordinary_payload_materializations": [1] * expected_n,
}
for key, expected in fixed_values.items():
    check(result["source"][key] == expected, f"source fixed values {key}")
for key, expected in {
    "operation_end_source_read_calls": [169] * expected_n,
    "operation_end_source_read_bytes": [6308061] * expected_n,
    "ordinary_payload_materializations": [1] * expected_n,
    "source_cache_before_publication_after_commit_hits": [0] * expected_n,
    "source_cache_before_publication_after_commit_cold_loads": [1] * expected_n,
    "source_cache_before_publication_after_commit_successful_loads": [1] * expected_n,
    "source_cache_before_publication_after_commit_retained_entries": [1] * expected_n,
    "source_cache_before_publication_after_commit_retained_bytes": [18855] * expected_n,
}.items():
    check(layout[key] == expected, f"layout fixed values {key}")
check(layout["owned_ingress_read_calls"] == [], "generic owned ingress call array")
check(layout["owned_ingress_read_bytes"] == [], "generic owned ingress bytes array")
check(layout["owned_ingress_owned_bytes"] == [], "generic owned ingress owned-byte array")
expected_source_sha = result["corpus"]["archive_sha256"]
expected_output_sha = result["output_sha256"]
check(layout["source_sha256"] == expected_source_sha, "source/layout corpus SHA")
check(layout["expected_output_sha256"] == expected_output_sha, "expected/top output SHA")
check(all(x == expected_output_sha for x in layout["output_sha256"]), "sample output SHAs")
check(result["corpus"]["archive_bytes"] == layout["source_bytes"], "archive/source byte count")
check(result["sink"]["write_calls"] == 75, "sink write count")
check(result["sink"]["accepted_bytes"] == 2101446, "sink accepted bytes")
check(result["sink"]["largest_write"] == 65536, "sink largest write")
check(sum(result["sink"]["write_size_buckets"].values()) == result["sink"]["write_calls"], "sink buckets")
flags = [
    "phase_sum_verified", "semantic_reopen_all_sections_verified",
    "table_cell_sectpr_excluded_verified", "raw_untouched_members_verified",
    "raw_member_order_and_comment_verified", "header_footer_relationships_verified",
    "source_immutability_verified", "patch_forward_inverse_verified",
    "publication_inverse_exact_source_verified",
    "stale_foreign_signed_noop_limits_partial_sink_verified",
    "cache_counters_verified", "sink_counters_verified",
]
check(all(layout[k] is True for k in flags), "correctness flags")
check(result["operation_metrics"]["sample_count"] == expected_n, "operation sample count")
check(result["operation_metrics"]["sample_indices"] == layout["sample_order"], "operation/sample alignment")
check(terminal["binary"]["sha256"] == report["binary_identity"]["binary_sha256"], "terminal/report binary SHA")
check(terminal["binary"]["bytes"] == report["binary_identity"]["binary_bytes"], "terminal/report binary bytes")
check(terminal["capture"]["samples"] == expected_n, "terminal sample count")
check(terminal["capture"]["warmups"] == 0, "terminal warmups")
check(terminal["capture"]["cpu"] == 12, "terminal CPU")
check(terminal["capture"]["route"] == "generic_read_at", "terminal route")
check(terminal["capture"]["perf"] == {"call_graph": "fp", "event": "cycles:u", "frequency_hz": 997}, "terminal perf")
check(terminal["tool_return"]["exit_code"] == 0, "terminal exit")
check(terminal["status"] == "terminal_success_persisted_before_offline_analysis", "terminal status")

perf_report_path = CAP / "perf-report-stdio.txt"
perf_self_path = CAP / "perf-report-self-symbol.txt"
perf_stack_path = CAP / "perf-script-stack-period.txt"
perf_stack_stderr_path = CAP / "perf-script-stack-period.stderr"
perf_leaf_path = CAP / "perf-script-leaf-period.txt"
perf_stderr_path = CAP / "perf-report-stdio.stderr"
perf_report = perf_report_path.read_text(errors="replace")
perf_stack = perf_stack_path.read_text(errors="replace")
stack_lines = perf_stack.splitlines()
header_re = re.compile(r"^\S+ \d+/\d+ \S+:\s+\d+ cycles:u:")
headers = [i for i, line in enumerate(stack_lines) if header_re.match(line)]
blocks = [stack_lines[headers[i]:(headers[i + 1] if i + 1 < len(headers) else len(stack_lines))]
  for i in range(len(headers))]
check(re.search(r"Total Lost Samples:\s*0\b", perf_report) is not None, "perf lost samples")
check(len(headers) == 15966, "stack header count")
check(sum(len(block) <= 1 for block in blocks) == 0, "stack blocks without frames")
check("truncated" not in perf_stack.lower(), "truncated stack marker")
check("PERF_RECORD" not in perf_stack, "perf record marker in stack export")
check("failed" not in perf_stack.lower(), "failed marker in stack export")
stack_headers = [stack_lines[i].split()[0] for i in headers]
check(stack_headers.count("litchi-perf-bas") == 15954, "target stack header count")
check(stack_headers.count("rustc") == 12, "other stack header count")
check(perf_stderr_path.read_text(errors="replace").count("could not read first record") == 11,
      "perf report addr2line diagnostic count")
summary = {
    "status": "pass" if not failures else "fail",
    "failures": failures,
    "report": {
        "path": str(report_path),
        "bytes": report_path.stat().st_size,
        "sha256": sha256(report_path),
        "case": result["case"],
        "samples": expected_n,
        "warmups": 0,
        "phase_sum_verified": True,
        "sample_array_lengths": expected_n,
        "source_archive_bytes": result["corpus"]["archive_bytes"],
        "source_archive_sha256": expected_source_sha,
        "source_observed_bytes": layout["source_bytes"],
        "source_read_calls": 202,
        "source_read_bytes": 8409488,
        "operation_end_source_read_calls": 169,
        "operation_end_source_read_bytes": 6308061,
        "ordinary_payload_read_calls": 138,
        "ordinary_payload_read_bytes": 8397632,
        "ordinary_payload_materializations": 1,
        "output_sha256": expected_output_sha,
        "all_sample_output_sha256_match": True,
        "correctness_flags": {k: layout[k] for k in flags},
    },
    "terminal_receipt": {
        "path": str(terminal_path),
        "bytes": terminal_path.stat().st_size,
        "sha256": sha256(terminal_path),
        "binary_sha256": terminal["binary"]["sha256"],
        "binary_bytes": terminal["binary"]["bytes"],
        "source": terminal["source"],
        "capture": terminal["capture"],
        "tool_return": terminal["tool_return"],
    },
    "perf": {
        "report_path": str(perf_report_path),
        "report_bytes": perf_report_path.stat().st_size,
        "report_sha256": sha256(perf_report_path),
        "lost_samples": 0,
        "stack_path": str(perf_stack_path),
        "stack_bytes": perf_stack_path.stat().st_size,
        "stack_sha256": sha256(perf_stack_path),
        "stack_header_count": len(headers),
        "stack_blocks_without_frames": sum(len(block) <= 1 for block in blocks),
        "stack_target_header_count": stack_headers.count("litchi-perf-bas"),
        "stack_other_process_headers": {"rustc": stack_headers.count("rustc")},
        "stack_unknown_frame_lines": sum("[unknown]" in line for line in stack_lines),
        "stack_truncated_marker_count": perf_stack.lower().count("truncated"),
        "stack_stderr_path": str(perf_stack_stderr_path),
        "stack_stderr_bytes": perf_stack_stderr_path.stat().st_size,
        "stack_stderr_sha256": sha256(perf_stack_stderr_path),
        "report_addr2line_error_count": perf_stderr_path.read_text(errors="replace").count("could not read first record"),
        "report_addr2line_stderr_path": str(perf_stderr_path),
        "report_addr2line_stderr_bytes": perf_stderr_path.stat().st_size,
        "report_addr2line_stderr_sha256": sha256(perf_stderr_path),
        "leaf_path": str(perf_leaf_path),
        "leaf_bytes": perf_leaf_path.stat().st_size,
        "leaf_sha256": sha256(perf_leaf_path),
    },
}
OUT.write_text(json.dumps(summary, indent=2) + "\n")
print(json.dumps(summary, indent=2))
raise SystemExit(0 if not failures else 1)
