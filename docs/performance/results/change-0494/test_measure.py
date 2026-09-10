#!/usr/bin/env python3
"""Pure custody and recomputation tests for the 0494 edit/provider driver.

These tests never build a Rust target, launch a benchmark, or freeze a
protocol. They exercise the fail-closed matrix, report oracles, statistical
summaries, and process custody in isolation.
"""

from __future__ import annotations

import copy
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure as matrix  # noqa: E402


def _source(scope: str, available: bool) -> dict[str, object]:
    if not available:
        return {"availability": "unavailable", "scope": scope,
                "calls": None, "empty_calls": None, "requested_bytes": None,
                "returned_bytes": None, "short_reads": None,
                "min_request_bytes": None, "max_request_bytes": None,
                "traced_ranges": None, "ranges": None,
                "requested_media_overlap_bytes": None,
                "returned_media_overlap_bytes": None}
    return {"availability": "available", "scope": scope, "calls": 3,
            "empty_calls": 0, "requested_bytes": 100, "returned_bytes": 100,
            "short_reads": 0, "min_request_bytes": 30, "max_request_bytes": 40,
            "traced_ranges": 3,
            "ranges": [{"offset": 0, "requested": 30, "returned": 30},
                       {"offset": 30, "requested": 30, "returned": 30},
                       {"offset": 60, "requested": 40, "returned": 40}],
            "requested_media_overlap_bytes": 0,
            "returned_media_overlap_bytes": 0}


def _range_adapter(available: bool, arm_name: str) -> dict[str, object]:
    if not available:
        return {"availability": "unavailable", "logical_calls": None,
                "requested_bytes": None, "returned_bytes": None,
                "short_reads": None, "delayed_calls": None,
                "transfer_paced_calls": None, "transfer_delay_ns": None}
    arm = matrix.ARM_BY_NAME[arm_name]
    paced = arm["transfer_bytes_per_second"] is not None
    transfer_delay = sum((returned * 1_000_000_000 + arm["transfer_bytes_per_second"] - 1)
                         // arm["transfer_bytes_per_second"]
                         for returned in (30, 30, 40)) if paced else 0
    return {"availability": "available", "logical_calls": 3,
            "requested_bytes": 100, "returned_bytes": 100, "short_reads": 0,
            "delayed_calls": 3 if arm["delay_us"] is not None else 0,
            "transfer_paced_calls": 3 if paced else 0,
            "transfer_delay_ns": transfer_delay}


def _sink() -> dict[str, object]:
    return {"accepted_bytes": matrix.EXPECTED_OUTPUT_BYTES, "write_calls": 339,
            "largest_write": 65536}


def _output() -> dict[str, object]:
    return {"bytes": matrix.EXPECTED_OUTPUT_BYTES, "sha256": matrix.EXPECTED_OUTPUT_SHA256}


def _row(index: int, arm_name: str, role: str = "normal") -> dict[str, object]:
    arm = matrix.ARM_BY_NAME[arm_name]
    available_logical = arm["provider"] in {"instrumented", "file", "short", "delayed"}
    available_physical = arm["provider"] in {"short", "delayed"}
    allocation: dict[str, object] = {}
    if role == "allocator":
        allocation = {"status": "measured", "scope": matrix.SAMPLE_ALLOCATION_SCOPE,
                      "allocation_calls": 4, "deallocation_calls": 3,
                      "reallocation_calls": 1, "failed_allocation_calls": 0,
                      "allocated_bytes": 100, "deallocated_bytes": 90,
                      "live_bytes_before": 10, "live_bytes_after": 20,
                      "peak_live_bytes_before": 20, "peak_live_bytes_after": 40,
                      "region_peak_live_bytes": 30}
    result = {
        "sample_index": index, "latency_ns": 100 + index,
        "source_version_before": {"id": 1, "revision": 0},
        "source_version_after": {"id": 1, "revision": 0},
        "source_version_unchanged": True, "output_bytes": matrix.EXPECTED_OUTPUT_BYTES,
        "output_sha256": matrix.EXPECTED_OUTPUT_SHA256, "output_exact_bytes": True,
        "materializations": 1, "commit_changed": True,
        "commit_operations": 1,
        "reads": {"logical": _source("caller-visible logical ReadAt calls", available_logical),
                  "physical": _source("underlying adapter ReadAt calls; transport model only",
                                      available_physical),
                  "range_adapter": _range_adapter(available_physical, arm_name),
                  "media_scope": "source compressed ranges for word/media members; exact output and OPC semantic checks additionally prove unchanged media payloads"},
        "sink": _sink(),
        "cache": {"successful_loads": 1, "expected_successful_loads": 1,
                   "exactly_one_main_part_materialization": True},
        "oracles": {"commit_identity_verified": True, "output_exact_bytes": True,
                     "semantic_reopen": True, "unchanged_media_preserved": True,
                     "source_version_unchanged": True, "cache_load_count": True,
                     "patch_oracles": {"scope": matrix.PATCH_ORACLE_SCOPE, "replay_forward": True,
                                        "inverse_restores_source": True,
                                        "stale_target_refused": True,
                                        "foreign_source_refused": True}},
    }
    if role == "allocator":
        result["allocation"] = allocation
    return result


def _report(arm_name: str, role: str = "normal", samples: int = 1,
            warmups: int = 0) -> dict[str, object]:
    arm = matrix.ARM_BY_NAME[arm_name]
    provider_kind = arm["kind"]
    return {
        "schema": matrix.SCHEMA, "version": matrix.VERSION,
        "case_name": matrix.CASE,
        "benchmark": "one DOCX paragraph replacement through source-backed open/commit/sequential publication",
        "provider_scope": matrix.PROVIDER_SCOPE,
        "timing_scope": matrix.TIMING_SCOPE, "setup_scope": matrix.SETUP_SCOPE,
        "allocation_scope": matrix.ALLOCATION_SCOPE,
        "physical_scope": matrix.PHYSICAL_SCOPE, "range_scope": matrix.RANGE_SCOPE,
        "zero_length_scope": matrix.ZERO_LENGTH_SCOPE,
        "file_scope": matrix.FILE_SCOPE,
        "corpus_version": matrix.CORPUS["version"],
        "corpus_generator": matrix.CORPUS["generator"],
        "source_archive_bytes": matrix.CORPUS["archive_bytes"],
        "source_archive_sha256": matrix.CORPUS["archive_sha256"],
        "source_bytes": matrix.CORPUS["archive_bytes"],
        "source_sha256": matrix.CORPUS["archive_sha256"],
        "expected_text_bytes": matrix.CORPUS["expected_text_bytes"],
        "expected_text_sha256": matrix.CORPUS["expected_text_sha256"],
        "expected_text_scope": "original deterministic corpus text identity before the selected paragraph replacement; it is not the changed output text identity",
        "expected_archive_members": matrix.CORPUS["archive_member_count"],
        "expected_output_bytes": matrix.EXPECTED_OUTPUT_BYTES,
        "expected_output_sha256": matrix.EXPECTED_OUTPUT_SHA256,
        "source_revision": "a" * 40, "requested_source_revision": "a" * 40,
        "limits": matrix.LIMITS,
        "binary_sha256": "c" * 64,
        "binary_bytes": 1, "current_exe": "/bin/true",
        "provider": {"name": arm["provider"], "provider": arm["provider"],
                     "kind": provider_kind, "trace_ranges": arm["trace_ranges"],
                     "short_read_bytes": arm["short_read_bytes"],
                     "zero_length_scope": matrix.ZERO_LENGTH_SCOPE,
                     "max_range_bytes": arm["max_range_bytes"],
                     "delay_us": arm["delay_us"],
                     "transfer_bytes_per_second": arm["transfer_bytes_per_second"],
                     "transfer_delay_policy": arm["transfer_delay_policy"],
                     "source_construction": "source adapters and FileSource handles are constructed before the operation clock",
                     "file_scope": matrix.FILE_SCOPE, "range_scope": matrix.RANGE_SCOPE},
        "corpus": matrix.CORPUS_MANIFEST,
        "preflight": {"expected_materializations": 1,
                       "source_document_xml_sha256": "d" * 64,
                       "candidate_document_xml_sha256": "e" * 64,
                       "output_exact_source_changed": True,
                       "semantic_reopen_verified": True,
                       "unchanged_media_preserved": True,
                       "replay_forward_verified": True,
                       "inverse_restores_source_verified": True,
                       "stale_target_refusal_verified": True,
                       "foreign_source_refusal_verified": True},
        "allocator": ("Rust system allocator" if role == "normal"
                       else "CountingSystemAllocator(std::alloc::System)"),
        "instrumentation": ("none" if role == "normal" else "system_allocator_operation_scoped"),
        "warmup": warmups, "samples": samples,
        "rows": [_row(warmups + index, arm_name, role) for index in range(samples)],
    }


class EditProviderMeasureTests(unittest.TestCase):
    def test_inventory_reverses_roles_and_provider_order(self) -> None:
        runs = matrix.formal_inventory()
        expected = len(matrix.ROLES) * len(matrix.REPEATS) * len(matrix.ARMS)
        self.assertEqual(len(runs), expected)
        self.assertEqual([item["role"] for item in runs[:len(matrix.ARMS)]], ["normal"] * len(matrix.ARMS))
        self.assertEqual([item["role"] for item in runs[len(matrix.ARMS):2 * len(matrix.ARMS)]], ["allocator"] * len(matrix.ARMS))
        pivot = 2 * len(matrix.ARMS)
        self.assertEqual([item["role"] for item in runs[pivot:pivot + len(matrix.ARMS)]], ["allocator"] * len(matrix.ARMS))
        self.assertEqual([item["arm"] for item in runs[pivot:pivot + len(matrix.ARMS)]],
                         [arm["name"] for arm in reversed(matrix.ARMS)])
        self.assertEqual(len({item["label"] for item in runs}), expected)
        self.assertEqual(len(matrix.formal_inventory(pilot=True)), len(matrix.ROLES) * len(matrix.ARMS))

    def test_command_binds_edit_provider_and_transport_controls(self) -> None:
        build = {"binary": {"path": "/bin/true"}}
        spec = next(item for item in matrix.formal_inventory() if item["arm"] == "delayed")
        command = matrix._command(spec, build, Path("report.json"), Path("resource.txt"), "a" * 40)
        self.assertIn("docx-edit-provider", command)
        self.assertIn("--provider", command)
        self.assertIn("delayed", command)
        self.assertIn("--samples", command)
        self.assertIn("--source-revision", command)
        self.assertIn("--transfer-delay-policy", command)

        short = next(item for item in matrix.formal_inventory() if item["arm"] == "short")
        short_command = matrix._command(short, build, Path("report.json"), Path("resource.txt"), "a" * 40)
        self.assertEqual(short_command.count("--max-range"), 1)
        self.assertNotIn("--short-read-bytes", short_command)

    def test_report_and_source_range_conservation_are_fail_closed(self) -> None:
        value = _report("instrumented")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="normal", arm_name="instrumented", samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["reads"]["logical"]["returned_bytes"] += 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="instrumented", samples=1, warmups=0)

    def test_edit_oracle_and_output_digest_are_required(self) -> None:
        value = _report("instrumented")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            altered = copy.deepcopy(value)
            altered["rows"][0]["oracles"]["patch_oracles"]["inverse_restores_source"] = False
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="instrumented", samples=1, warmups=0)

    def test_adapter_conservation_and_service_pacing_are_fail_closed(self) -> None:
        value = _report("delayed")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="normal", arm_name="delayed", samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["reads"]["range_adapter"]["logical_calls"] += 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="delayed", samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["reads"]["range_adapter"]["transfer_delay_ns"] -= 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="delayed", samples=1, warmups=0)

    def test_operation_types_allocation_conservation_and_patch_scope(self) -> None:
        value = _report("short", "allocator")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="allocator", arm_name="short", samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["commit_operations"] = True
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="allocator", arm_name="short", samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["allocation"]["live_bytes_after"] += 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="allocator", arm_name="short", samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["oracles"]["patch_oracles"]["scope"] = "test"
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="allocator", arm_name="short", samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["output_sha256"] = matrix.CORPUS["archive_sha256"]
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="instrumented", samples=1, warmups=0)

    def test_matching_two_materializations_is_rejected(self) -> None:
        value = _report("instrumented")
        value["preflight"]["expected_materializations"] = 2
        row = value["rows"][0]
        row["materializations"] = 2
        row["cache"]["successful_loads"] = 2
        row["cache"]["expected_successful_loads"] = 2
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="instrumented", samples=1, warmups=0)

    def test_allocator_envelope_and_normal_absence(self) -> None:
        for role in ("normal", "allocator"):
            value = _report("short", role)
            with tempfile.TemporaryDirectory() as directory:
                path = Path(directory) / "report.json"
                path.write_text(json.dumps(value), encoding="utf-8")
                matrix.validate_report(path, role=role, arm_name="short", samples=1, warmups=0)

    def test_reallocation_does_not_require_standalone_deallocation(self) -> None:
        value = _report("short", "allocator")
        value["rows"][0]["allocation"]["deallocation_calls"] = 0
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="allocator", arm_name="short", samples=1, warmups=0)

    def test_deterministic_bootstrap_and_tail_percentiles(self) -> None:
        values = [9, 1, 5, 3]
        self.assertEqual(matrix._percentiles(values)["p50"], 4)
        self.assertEqual(matrix._bootstrap_median(values), matrix._bootstrap_median(values))
        self.assertEqual(matrix._percentiles(values)["p99"], 9)

    def test_analysis_retains_source_summary_and_derives_request_histograms(self) -> None:
        report = _report("instrumented")
        entry = {"report": report, "resource": {"maximum_resident_set_size_(kbytes)": 12}}
        vectors = matrix._vectors(entry, "normal", matrix.ARM_BY_NAME["instrumented"])
        self.assertEqual(sum(vectors["source_logical_request_size_histogram"]), 3)
        summary = matrix._analysis_read_summary(report["rows"][0]["reads"])
        self.assertNotIn("ranges", summary["logical"])
        self.assertEqual(summary["logical"]["range_count"], 3)

        physical_report = _report("delayed")
        physical_entry = {"report": physical_report,
                          "resource": {"maximum_resident_set_size_(kbytes)": 12}}
        vectors = matrix._vectors(physical_entry, "normal", matrix.ARM_BY_NAME["delayed"])
        self.assertEqual(vectors["source_physical_requested_media_overlap_bytes"], [0])
        self.assertEqual(vectors["source_physical_returned_media_overlap_bytes"], [0])
        self.assertNotIn("cache_failed_loads", vectors)
        self.assertNotIn("cache_retained_bytes", vectors)
        self.assertNotIn("cache_retained_entries", vectors)
        self.assertEqual(vectors["cache_exactly_one_main_part_materialization"], [True])
        stats, bootstrap = matrix._summaries(vectors)
        self.assertIsNone(stats["cache_exactly_one_main_part_materialization"])
        self.assertIsNone(bootstrap["cache_exactly_one_main_part_materialization"])

    def test_timeout_kills_the_process_group(self) -> None:
        process = subprocess.Popen(["sleep", "30"], start_new_session=True)
        termination = matrix._kill_group(process)
        self.assertIn(termination, ("SIGTERM", "SIGKILL"))
        self.assertIsNotNone(process.returncode)

    def test_plan_has_no_capture_side_effects(self) -> None:
        plan = matrix.protocol_value(None)
        self.assertIsNone(plan["source"])
        self.assertIsNone(plan["builds"])
        self.assertFalse(plan["claim_authorized"])


if __name__ == "__main__":
    raise SystemExit(unittest.main())
