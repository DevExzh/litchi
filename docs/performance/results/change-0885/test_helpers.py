#!/usr/bin/env python3
"""Focused mocked tests for the 0885 capture and verification helpers."""

from __future__ import annotations

import json
from pathlib import Path
import stat
import tempfile
import unittest

import capture
import verify


def _sidecar() -> dict:
    parts = []
    numeric = [1024, 614, 1024, 1024]
    shared = [0, 410, 0, 0]
    for index, (number, strings) in enumerate(zip(numeric, shared), 1):
        parts.append(
            {
                "part": f"xl/worksheets/sheet{index}.xml",
                "numeric_cells": number,
                "shared_string_cells": strings,
                "markup_compatibility_namespace": True,
                "x14ac_namespace": True,
                "cols_block": True,
                "selected_cell_stream_ineligible_on_cols": True,
                "source_stream_eligible": False,
                "source_stream_ineligible_reasons": ["markup-compatibility-namespace", "x14ac-namespace"],
            }
        )
    parts.extend(
        [
            {"part": "xl/sharedStrings.xml", "numeric_cells": 0, "shared_string_cells": 0},
            {"part": "xl/workbook.xml", "numeric_cells": 0, "shared_string_cells": 0},
        ]
    )
    return {
        "schema": "litchi.perf-baseline.producer-shape-evidence.v1",
        "corpora": [
            {
                "generator": "litchi-xlsx-producer-shape-v1",
                "shape": "producer-medium-read",
                "variant": "read",
                "selected_sheet": 1,
                "selected_target": "Bench01!D17",
                "shared_string_part": "xl/sharedStrings.xml",
                "shared_string_unique_count": 64,
                "shared_string_cell_percent": 40,
                "parts": parts,
                "proven_refusals": [{"message": "admitted by the value-only editor"} for _ in range(6)],
            }
        ],
    }


def _report(mode: str, samples: int, nonidentity: bool = False) -> dict:
    values = list(range(1, samples + 1))
    order = list(reversed(range(samples))) if nonidentity else list(range(samples))
    tool = {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": "litchi-perf-baseline" if mode == "native" else "litchi-perf-baseline-alloc",
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": "none" if mode == "native" else "system_allocator_operation_scoped",
    }
    if mode == "allocator":
        tool["allocator_counter_revision"] = "serialized_region_peak_v3"
    scan = {
        "sample_indices": order,
        "archive_sha256": capture.ARCHIVE_SHA256,
        "archive_bytes": capture.ARCHIVE_BYTES,
        "expected_cell_count": 4096,
        "expected_numeric_count": 3686,
        "expected_text_count": 410,
        "observed_cell_counts": [4096] * samples,
        "observed_numeric_counts": [3686] * samples,
        "observed_text_counts": [410] * samples,
        "oracle_digest_sha256": capture.ORACLE_SHA256,
        "observed_digest_sha256": [capture.ORACLE_SHA256] * samples,
        "marker_refusal_census_sha256": capture.CENSUS_SHA256,
        "source_version_stable": [True] * samples,
        "physical_cold_eligible": False,
        "runtime_route_observed": False,
        "worksheet_route_eligibility": [
            {"worksheet": f"xl/worksheets/sheet{i}.xml", "source_stream_eligible": False} for i in range(1, 5)
        ],
    }
    metrics = {
        "sample_count": samples,
        "sample_indices": order,
        "alignment": "elapsed_ns.samples_by_elapsed_then_sample_index",
    }
    if mode == "native":
        metrics["process"] = {
            "peak_rss_bytes": {
                "values": [100] * samples,
                "status": "measured",
                "scope": "process_lifetime_high_water_after_not_operation_peak",
            }
        }
    if mode == "allocator":
        metrics["allocation"] = {
            "status": "measured",
            "scope": "operation_global_system_allocator",
            **{
                key: {
                    "values": [1] * samples if key in ("allocation_calls", "allocated_bytes", "region_peak_live_bytes") else [0] * samples,
                    "status": "measured",
                    "scope": "operation_global_system_allocator",
                }
                for key in ("allocation_calls", "allocated_bytes", "region_peak_live_bytes", "failed_allocation_calls", "deallocation_calls")
            },
        }
    return {
        "schema_version": 1,
        "tool": tool,
        "configuration": {
            "samples_per_case": samples,
            "warmup_iterations_per_case": 3,
            "filesystem_cache_states": ["warm"],
            "filesystem_fresh_child_per_sample": True,
            "filesystem_process_isolated": True,
            "cases": [capture.CASE],
        },
        "results": [
            {
                "case": capture.CASE,
                "elapsed_ns": {
                    "samples": values,
                    "sample_order": order,
                    **verify.raw_statistics(values),
                },
                "corpus": {"archive_sha256": capture.ARCHIVE_SHA256, "archive_bytes": capture.ARCHIVE_BYTES, "entry_count": 4096},
                "source": {"producer_full_cell_scan": scan},
                "operation_metrics": metrics,
            }
        ],
    }


class _FakeProcess:
    def __init__(self, argv: list[str], stdout, stderr) -> None:
        self.argv = argv
        self.stdout = stdout
        self.stderr = stderr

    def wait(self) -> int:
        self.stdout.write(b"mock child\n")
        mode = "allocator" if self.argv[3].endswith("allocator") else "native"
        samples = int(self.argv[9])
        Path(self.argv[13]).write_text(json.dumps(_report(mode, samples)) + "\n", encoding="utf-8")
        Path(self.argv[15]).write_text(json.dumps(_sidecar()) + "\n", encoding="utf-8")
        return 0


class HelpersTest(unittest.TestCase):
    def test_statistics_match_rust_midpoint_and_nearest_rank(self) -> None:
        self.assertEqual(verify.raw_statistics([1, 2, 3, 4]), {"min": 1, "p50": 2, "p95": 4, "p99": 4, "max": 4})

    def test_allocator_rejects_failed_calls_and_requires_positive_primary(self) -> None:
        allocation = _report("allocator", 15)["results"][0]["operation_metrics"]["allocation"]
        verify._validate_allocation(allocation, 15)
        allocation["failed_allocation_calls"]["values"][3] = 1
        with self.assertRaises(ValueError):
            verify._validate_allocation(allocation, 15)

    def test_sidecar_checks_full_four_sheet_census(self) -> None:
        corpus = verify._validate_sidecar(_sidecar())
        self.assertEqual(corpus["shape"], "producer-medium-read")
        broken = _sidecar()
        broken["corpora"][0]["parts"][1]["numeric_cells"] += 1
        with self.assertRaises(ValueError):
            verify._validate_sidecar(broken)

    def test_report_rejects_debug_profile(self) -> None:
        report = _report("native", 30)
        report["tool"]["profile"] = "debug"
        with self.assertRaises(ValueError):
            verify._validate_report(report, _sidecar(), "native", {"warmup": 3, "samples": 30})

    def test_report_validates_raw_elapsed_and_aligned_vectors(self) -> None:
        checked = verify._validate_report(_report("allocator", 15), _sidecar(), "allocator", {"warmup": 3, "samples": 15})
        self.assertEqual(checked["elapsed"]["p50"], 8)
        native = verify._validate_report(_report("native", 30, nonidentity=True), _sidecar(), "native", {"warmup": 3, "samples": 30})
        self.assertEqual(max(native["process_peak_rss"]), 100)

    def test_capture_mock_is_single_pass_and_retains_outputs(self) -> None:
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            packet = root / "packet"
            (packet / "binaries").mkdir(parents=True)
            paths = {}
            for mode in ("native", "allocator"):
                path = packet / "binaries" / mode
                path.write_bytes(b"mock executable")
                path.chmod(stat.S_IRUSR | stat.S_IXUSR)
                paths[mode] = capture.descriptor(path)
                paths[mode]["features"] = [] if mode == "native" else ["allocator-metrics"]
            build = {"status": "pass", "binaries": paths}
            (packet / "build.json").write_text(json.dumps(build), encoding="utf-8")
            (packet / "source.json").write_text("source\n", encoding="utf-8")
            (packet / "source.tar.gz").write_bytes(b"archive")
            (packet / "review.json").write_text("{}\n", encoding="utf-8")
            (packet / "environment.json").write_text("{}\n", encoding="utf-8")
            binding_paths = {
                "build": packet / "build.json",
                "source_manifest": packet / "source.json",
                "source_archive": packet / "source.tar.gz",
                "capture_helper": Path(capture.__file__).resolve(),
                "verify_helper": Path(verify.__file__).resolve(),
                "test_helper": Path(__file__).resolve(),
                "review": packet / "review.json",
                "environment": packet / "environment.json",
                "taskset": Path(capture.TASKSET),
            }
            plan = {
                "schema": capture.SCHEMA,
                "case": capture.CASE,
                "profile": "release",
                "cpu_affinity": "12",
                "cache": "warm",
                "fresh_child": True,
                "fail_fast": True,
                "schedule": ["native", "allocator"],
                "archive": {
                    "sha256": capture.ARCHIVE_SHA256,
                    "bytes": capture.ARCHIVE_BYTES,
                    "oracle_sha256": capture.ORACLE_SHA256,
                    "marker_refusal_census_sha256": capture.CENSUS_SHA256,
                    "entry_count": 4096,
                    "numeric_count": 3686,
                    "text_count": 410,
                },
                "binaries": {mode: paths[mode] for mode in paths},
                "bindings": {name: capture.descriptor(path) for name, path in binding_paths.items()},
                "lanes": {
                    "native": {"name": "native", "mode": "native", "children": 12, "warmup": 3, "samples": 30},
                    "allocator": {"name": "allocator", "mode": "allocator", "children": 6, "warmup": 3, "samples": 15},
                },
            }
            plan_path = packet / "plan.json"
            plan_path.write_text(json.dumps(plan), encoding="utf-8")
            output = root / "capture"
            argv = capture._command("native", str(packet / "binaries" / "native"), 3, 30, root / "r.json", root / "e.json")
            self.assertEqual(len(argv), 16)
            self.assertEqual(argv[0], "/usr/bin/taskset")
            self.assertEqual((argv[13], argv[15]), (str(root / "r.json"), str(root / "e.json")))
            capture.run_lane(plan_path, output, "native", popen=lambda argv, **kwargs: _FakeProcess(argv, kwargs["stdout"], kwargs["stderr"]))
            self.assertTrue((output / "native" / "child-11" / "report.json").is_file())
            with self.assertRaises(FileExistsError):
                capture.run_lane(plan_path, output, "native", popen=lambda argv, **kwargs: _FakeProcess(argv, kwargs["stdout"], kwargs["stderr"]))
            calls = []

            def bad_popen(argv, **kwargs):
                process = _FakeProcess(argv, kwargs["stdout"], kwargs["stderr"])
                original_wait = process.wait

                def wait():
                    code = original_wait()
                    calls.append(argv)
                    if len(calls) == 2:
                        report = Path(argv[13])
                        value = json.loads(report.read_text(encoding="utf-8"))
                        value["tool"]["profile"] = "debug"
                        report.write_text(json.dumps(value) + "\n", encoding="utf-8")
                    return code

                process.wait = wait
                return process

            with self.assertRaises(ValueError):
                capture.run_lane(plan_path, root / "bad-capture", "native", popen=bad_popen)
            self.assertEqual(len(calls), 2)
            self.assertFalse((root / "bad-capture" / "native" / "child-02").exists())


if __name__ == "__main__":
    unittest.main(verbosity=2)
