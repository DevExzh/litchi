#!/usr/bin/env python3
"""Pure custody tests for the 0494 verified-cold evidence driver."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import cold_measure as cold  # noqa: E402
import measure as warm  # noqa: E402


ALIGNED_SHA = "a" * 64
OUTPUT_SHA = "b" * 64


def _limits() -> dict[str, object]:
    return {
        "read_limits": {
            "max_input_bytes": 512 * 1024 * 1024,
            "max_archive_members": 100_000,
            "max_archive_member_name_bytes": 4 * 1024,
            "max_archive_metadata_bytes": 64 * 1024 * 1024,
            "max_archive_compressed_bytes": 512 * 1024 * 1024,
            "max_archive_entry_bytes": 512 * 1024 * 1024,
            "max_archive_total_bytes": 2 * 1024 * 1024 * 1024,
            "max_parts": 100_000,
            "max_part_bytes": 512 * 1024 * 1024,
            "max_total_part_bytes": 512 * 1024 * 1024,
            "max_content_types_bytes": 8 * 1024 * 1024,
            "max_content_type_mappings": 100_000,
            "max_relationship_parts": 100_000,
            "max_relationship_xml_bytes": 8 * 1024 * 1024,
            "max_total_relationship_xml_bytes": 64 * 1024 * 1024,
            "max_relationships_per_part": 100_000,
            "max_total_relationships": 1_000_000,
            "max_relationship_graph_nodes": 100_000,
            "max_xml_events": 1_000_000,
            "max_total_relationship_xml_events": 8_000_000,
            "max_xml_depth": 256,
            "max_xml_attribute_bytes": 64 * 1024,
            "max_relationship_target_bytes": 4 * 1024,
        },
        "cache_max_bytes": warm.LIMITS["cache_max_bytes"],
        "cache_max_entries": warm.LIMITS["cache_max_entries"],
        "resource_budget": {
            "managed": False,
            "unmanaged_reason": cold.UNMANAGED_REASON,
            "memory_bytes": None,
            "input_bytes": None,
            "output_bytes": None,
            "objects": None,
            "depth": None,
            "work": None,
        },
        "max_tracked_ranges": warm.LIMITS["max_tracked_ranges"],
        "sink_max_write_bytes": warm.LIMITS["sink_max_write_bytes"],
    }


def _proof(status: str = "eligible") -> dict[str, object]:
    value = {
        "status": status,
        "filesystem_magic": None,
        "page_size_bytes": None,
        "source_bytes": None,
        "source_pages": None,
        "aligned_source_bytes": None,
        "aligned_source_sha256": None,
        "fsync_completed": None,
        "advice": None,
        "fincore_size_bytes": None,
        "resident_bytes": None,
        "dirty_bytes": None,
        "writeback_bytes": None,
        "fincore_tool": None,
        "fincore_sha256": None,
        "fincore_version": None,
        "fincore_stderr_sha256": None,
        "fincore_stderr_bytes": None,
        "fincore_version_stderr_sha256": None,
        "fincore_version_stderr_bytes": None,
        "fincore_method": None,
        "fincore_fallback": None,
        "read_bytes_before": None,
        "read_bytes_after": None,
        "read_bytes_delta": None,
    }
    if status == "eligible":
        value.update({
            "filesystem_magic": 61267, "page_size_bytes": 4096,
            "source_bytes": 16_793_600, "source_pages": 4100,
            "aligned_source_bytes": 16_793_600, "aligned_source_sha256": ALIGNED_SHA,
            "fsync_completed": True, "advice": cold.FINFORE_ADVICE,
            "fincore_size_bytes": 16_793_600, "resident_bytes": 0,
            "dirty_bytes": 0, "writeback_bytes": 0, "fincore_tool": "fincore",
            "fincore_sha256": "c" * 64, "fincore_version": "fincore from util-linux",
            "fincore_stderr_sha256": "d" * 64, "fincore_stderr_bytes": 0,
            "fincore_version_stderr_sha256": "e" * 64,
            "fincore_version_stderr_bytes": 0,
            "fincore_method": cold.FINFORE_METHOD, "fincore_fallback": cold.FINFORE_FALLBACK,
            "read_bytes_before": 1, "read_bytes_after": 2, "read_bytes_delta": 1,
        })
    return value


def _row(role: str = "normal", pid: int = 1234) -> dict[str, object]:
    value: dict[str, object] = {
        "sample_index": 0, "child_process_id": pid, "latency_ns": 1000,
        "output_bytes": 16_793_612, "output_sha256": OUTPUT_SHA,
        "materializations": 1, "commit_changed": True, "commit_operations": 1,
        "source_version_before": {"id": 3, "revision": 0},
        "source_version_after": {"id": 3, "revision": 0},
        "source_version_unchanged": True,
        "reads": {"source": {
            "availability": "available", "scope": cold.READ_SCOPE, "calls": 1,
            "empty_calls": 0, "requested_bytes": 1, "returned_bytes": 1,
            "short_reads": 0, "min_request_bytes": 1, "max_request_bytes": 1,
            "traced_ranges": 1, "requested_media_overlap_bytes": 0,
            "returned_media_overlap_bytes": 0,
            "ranges": [{"offset": 0, "requested": 1, "returned": 1}],
        }, "scope": cold.READ_EVIDENCE_SCOPE},
        "sink": {"accepted_bytes": 16_793_612, "write_calls": 2, "largest_write": 65_536},
        "cache": {"successful_loads": 1, "expected_successful_loads": 1,
                  "exactly_one_main_part_materialization": True},
        "process_metrics": {
            "rchar": 10, "wchar": 0, "read_bytes": 1, "write_bytes": 0,
            "cancelled_write_bytes": 0, "syscr": 1, "syscw": 0,
            "minor_faults": 1, "major_faults": 0, "user_cpu_ticks": 1,
            "system_cpu_ticks": 1, "clock_ticks_per_second": 100,
            "voluntary_context_switches": 0, "nonvoluntary_context_switches": 0,
            "rss_bytes": 100, "peak_rss_bytes": 200,
        },
        "process_metrics_scope": cold.PROCESS_METRICS_SCOPE,
        "rss_scope": cold.RSS_SCOPE,
        "oracles": {
            "commit_identity_verified": True, "output_exact_bytes": True,
            "semantic_reopen": True, "unchanged_media_preserved": True,
            "source_version_unchanged": True, "cache_load_count": True,
            "logical_source_reads_positive": True,
            "patch_oracles": {"scope": "untimed cold-child preflight commit patch oracles", "replay_forward": True,
                               "inverse_restores_source": True, "stale_target_refused": True,
                               "foreign_source_refused": True},
        },
    }
    if role == "allocator":
        value["allocation"] = {
            "status": "measured", "scope": warm.SAMPLE_ALLOCATION_SCOPE,
            "allocation_calls": 4, "deallocation_calls": 3, "reallocation_calls": 1,
            "failed_allocation_calls": 0, "allocated_bytes": 100,
            "deallocated_bytes": 90, "live_bytes_before": 10, "live_bytes_after": 20,
            "peak_live_bytes_before": 20, "peak_live_bytes_after": 40,
            "region_peak_live_bytes": 30,
        }
    return value


def _report(role: str = "normal", status: str = "eligible") -> dict[str, object]:
    value = {
        "limits": _limits(), "schema": cold.COLD_REPORT_SCHEMA,
        "benchmark": cold.BENCHMARK, "provider_scope": cold.PROVIDER_SCOPE,
        "timing_scope": cold.TIMING_SCOPE, "setup_scope": cold.SETUP_SCOPE,
        "cold_claim_scope": cold.COLD_CLAIM_SCOPE,
        "process_metrics_scope": cold.PROCESS_METRICS_SCOPE, "rss_scope": cold.RSS_SCOPE,
        "provider": "file-cold-verified", "cache_state": "cold-verified",
        "filesystem_root_selected": False,
        "source_archive_sha256": warm.CORPUS["archive_sha256"],
        "source_archive_bytes": warm.CORPUS["archive_bytes"],
        "aligned_source_sha256": ALIGNED_SHA, "aligned_source_bytes": 16_793_600,
        "expected_output_sha256": OUTPUT_SHA, "expected_output_bytes": 16_793_612,
        "source_revision": "a" * 40, "corpus": warm.CORPUS_MANIFEST,
        "preflight": {
            "expected_materializations": 1,
            "source_document_xml_sha256": "f" * 64,
            "candidate_document_xml_sha256": "0" * 64,
            "output_exact_source_changed": True, "semantic_reopen_verified": True,
            "unchanged_media_preserved": True, "replay_forward_verified": True,
            "inverse_restores_source_verified": True,
            "stale_target_refusal_verified": True,
            "foreign_source_refusal_verified": True,
        },
        "cold_verified_status": status, "cold_verified_samples": [_proof(status)],
        "cold_verified_fincore_command": cold.FINFORE_COMMAND,
        "warmup": 0, "samples": 1,
        "rows": [] if status != "eligible" else [_row(role)],
    }
    return value


class ColdMeasureTests(unittest.TestCase):
    def test_inventory_has_two_repeats_and_fresh_invocation_counts(self) -> None:
        formal = cold.formal_inventory()
        self.assertEqual(len(formal), 4)
        self.assertEqual(sum(item["samples"] for item in formal), 120)
        self.assertEqual(len(cold.formal_inventory(pilot=True)), 2)
        self.assertEqual(sum(item["samples"] for item in cold.formal_inventory(pilot=True)), 6)
        self.assertTrue(all(item["warmups"] == 0 for item in formal))

    def test_command_is_one_fresh_cold_parent_invocation(self) -> None:
        build = {"binary": {"path": "/bin/true"}}
        spec = cold.formal_inventory()[0]
        command = cold._command(spec, build, Path("private/report.json"),
                                Path("private/resource.txt"), "a" * 40)
        self.assertIn("docx-edit-provider-cold", command)
        self.assertEqual(command[command.index("--samples") + 1], "1")
        self.assertEqual(command[command.index("--warmup") + 1], "0")
        self.assertIn("file-cold-verified", command)
        self.assertIn("cold-verified", command)
        self.assertIn("a" * 40, command)

    def test_eligible_report_requires_alignment_fincore_and_positive_read_bytes(self) -> None:
        for role in cold.ROLES:
            value = _report(role)
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "report.json"
                path.write_text(json.dumps(value), encoding="utf-8")
                checked = cold.validate_report(path, role=role, source_revision="a" * 40)
            self.assertTrue(checked["eligible"])
            self.assertEqual(len(checked["rows"]), 1)

        bad = _report()
        bad["cold_verified_samples"][0]["resident_bytes"] = 1
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(bad), encoding="utf-8")
            with self.assertRaises(cold.ColdMeasureError):
                cold.validate_report(path, role="normal", source_revision="a" * 40)

    def test_ineligible_report_is_zero_row_and_not_a_formal_measurement(self) -> None:
        value = _report(status="ineligible_fincore_unavailable")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            checked = cold.validate_report(path, role="normal", source_revision="a" * 40)
        self.assertFalse(checked["eligible"])

        # Rust's serde envelope omits all unavailable optional proof fields on
        # an early failure, leaving only the explicit ineligible status.
        minimal = copy.deepcopy(value)
        minimal["cold_verified_samples"] = [{"status": "ineligible_source_page_size_unavailable"}]
        minimal["cold_verified_status"] = "ineligible_source_page_size_unavailable"
        self.assertEqual(checked["rows"], [])

        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "minimal.json"
            path.write_text(json.dumps(minimal), encoding="utf-8")
            checked = cold.validate_report(path, role="normal", source_revision="a" * 40)
        self.assertFalse(checked["eligible"])
        self.assertEqual(checked["rows"], [])

        early = copy.deepcopy(value)
        early["aligned_source_sha256"] = None
        early["aligned_source_bytes"] = None
        early["expected_output_sha256"] = None
        early["expected_output_bytes"] = None
        early["preflight"] = None
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "early.json"
            path.write_text(json.dumps(early), encoding="utf-8")
            checked = cold.validate_report(path, role="normal", source_revision="a" * 40)
        self.assertFalse(checked["eligible"])

        unknown = copy.deepcopy(value)
        unknown["cold_verified_status"] = "ineligible_future_reason"
        unknown["cold_verified_samples"][0]["status"] = "ineligible_future_reason"
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "unknown.json"
            path.write_text(json.dumps(unknown), encoding="utf-8")
            with self.assertRaises(cold.ColdMeasureError):
                cold.validate_report(path, role="normal", source_revision="a" * 40)

    def test_range_conservation_rejects_summary_without_raw_ranges(self) -> None:
        value = _report()
        del value["rows"][0]["reads"]["source"]["ranges"]
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(cold.ColdMeasureError):
                cold.validate_report(path, role="normal", source_revision="a" * 40)

        value = _report()
        value["rows"][0]["reads"]["source"]["ranges"][0]["offset"] = 16_793_600
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "out-of-bounds.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(cold.ColdMeasureError):
                cold.validate_report(path, role="normal", source_revision="a" * 40)

    def test_matching_two_materializations_cannot_claim_exactly_one(self) -> None:
        value = _report()
        value["preflight"]["expected_materializations"] = 2
        value["rows"][0]["materializations"] = 2
        value["rows"][0]["cache"]["successful_loads"] = 2
        value["rows"][0]["cache"]["expected_successful_loads"] = 2
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(cold.ColdMeasureError):
                cold.validate_report(path, role="normal", source_revision="a" * 40)

    def test_statistics_are_deterministic_and_repeat_only(self) -> None:
        values = [9, 2, 7, 4, 1]
        self.assertEqual(cold._percentiles(values), cold._percentiles(values))
        self.assertEqual(cold._bootstrap_median(values), cold._bootstrap_median(values))
        self.assertNotIn("warm", cold._bootstrap_median(values)["method"])

    def test_cold_protocol_has_separate_driver_bindings(self) -> None:
        value = cold.protocol_value()
        self.assertEqual(value["schema"], cold.COLD_PROTOCOL_SCHEMA)
        self.assertIsNone(value["warm_protocol"])
        self.assertEqual(set(value["drivers"]), {"cold_measure.py", "test_cold_measure.py"})
        self.assertEqual(value["expected_formal_invocations"], 120)


if __name__ == "__main__":
    unittest.main()
