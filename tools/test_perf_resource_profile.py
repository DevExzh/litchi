"""Unit tests for the standard-library resource-profile parser."""

from __future__ import annotations

import argparse
import copy
import hashlib
import json
import tempfile
import unittest
from unittest import mock
from pathlib import Path

from tools import perf_resource_profile


# Independent retained-0251 corpus fixtures.  These deliberately do not read
# implementation constants: changing the validator's pin must break these
# tests until the retained evidence is consciously re-reviewed.
XLSX_TINY_CORPUS_FIXTURE = {
    "name": "xlsx-tiny",
    "generator": "litchi-xlsx-synthetic-v1",
    "package_format": "XLSX/OPC/ZIP",
    "shape": "tiny",
    "payload_kind": "deterministic-integer-grid",
    "compression": "deflate",
    "entry_count": 192,
    "archive_member_count": 8,
    "entry_bytes": 4,
    "uncompressed_payload_bytes": 768,
    "archive_bytes": 3561,
    "archive_sha256": "69ef199769a316eaa465a41ebf08f7a1b501f708775fabd7a084a90dc6a9b428",
    "target_entry": "Sheet1!A1",
    "target_payload_bytes": 1,
    "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
    "xlsx": {
        "sheet_count": 3,
        "rows_per_sheet": 8,
        "columns_per_sheet": 8,
        "one_percent_update_count": 2,
        "source_members": {
            "workbook": "xl/workbook.xml",
            "worksheets": [
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet3.xml",
            ],
            "shared_strings": None,
            "styles": "xl/styles.xml",
        },
    },
}
XLSX_MEDIUM_CORPUS_FIXTURE = {
    "name": "xlsx-cell-values-medium",
    "generator": "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
    "package_format": "XLSX/OPC/ZIP",
    "shape": "medium",
    "payload_kind": "deterministic-multi-sheet-scalar-grid-with-media",
    "compression": "deflate",
    "entry_count": 9216,
    "archive_member_count": 17,
    "entry_bytes": 4,
    "uncompressed_payload_bytes": 4231168,
    "archive_bytes": 4226429,
    "archive_sha256": "dfff7ec0c749d9e404091776f15a8fb690985af7f58efdfe659dbeaed7145036",
    "target_entry": "Sheet1!A1",
    "target_payload_bytes": 1,
    "target_payload_sha256": "5feceb66ffc86f38d952786c6d696c79c2dbc239dd4e91b46729d73a27fb57e9",
    "xlsx": {
        "sheet_count": 4,
        "rows_per_sheet": 48,
        "columns_per_sheet": 48,
        "one_percent_update_count": 93,
        "source_members": {
            "workbook": "xl/workbook.xml",
            "worksheets": [
                "xl/worksheets/sheet1.xml",
                "xl/worksheets/sheet2.xml",
                "xl/worksheets/sheet3.xml",
                "xl/worksheets/sheet4.xml",
            ],
            "shared_strings": None,
            "styles": "xl/styles.xml",
        },
    },
}
XLSX_CORPUS_CANONICAL_DIGESTS = {
    "tiny": "0f521b922f4c4a408a5cdaf87dd2dc84eefb28589fd38955c54ae99000351aed",
    "medium": "4cdb4fc4199a0604ea1431c044f3fe86c04e9822d1121044110ba44356f43efc",
}


class ResourceProfileParserTests(unittest.TestCase):
    def test_time_parser_extracts_rss_and_elapsed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "time.txt"
            path.write_text(
                """\
Maximum resident set size (kbytes): 12345
User time (seconds): 1.25
System time (seconds): 0.50
\tElapsed (wall clock) time (h:mm:ss or m:ss): 0:02.75
Voluntary context switches: 12
Involuntary context switches: 3
Major (requiring I/O) page faults: 1
Minor (reclaiming a frame) page faults: 34
""",
                encoding="utf-8",
            )
            parsed = perf_resource_profile.parse_time_report(path)
        self.assertEqual(parsed["status"], "ok")
        self.assertEqual(parsed["max_rss_kib"], 12345)
        self.assertEqual(parsed["user_seconds"], 1.25)
        self.assertEqual(parsed["elapsed_wall_seconds"], 2.75)
        self.assertEqual(parsed["voluntary_context_switches"], 12)

    def test_time_parser_rejects_partial_report_as_unavailable(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "partial-time.txt"
            path.write_text(
                "Maximum resident set size (kbytes): 12345\n"
                "Elapsed (wall clock) time (h:mm:ss or m:ss): 0:02.75\n",
                encoding="utf-8",
            )
            parsed = perf_resource_profile.parse_time_report(path)
        self.assertEqual(parsed["status"], "unavailable")
        self.assertIn("user_seconds", parsed["missing_fields"])
        self.assertNotEqual(parsed["status"], "ok")

    def test_time_parser_marks_malformed_numeric_fields_unparsed(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "malformed-time.txt"
            path.write_text(
                """\
Maximum resident set size (kbytes): 12345
User time (seconds): nope
""",
                encoding="utf-8",
            )
            parsed = perf_resource_profile.parse_time_report(path)
        self.assertEqual(parsed["status"], "unparsed")
        self.assertIsNone(parsed["user_seconds"])

    def test_strace_parser_buckets_successes_and_failures(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "strace.log"
            path.write_text(
                """\
1 read(3, \"x\", 10) = 10
1 read(3, \"\", 0) = 0
1 read(3, 0x0, 4) = -1 EIO (Input/output error)
1 write(1, \"x\", 4096) = 4096
1 write(1, \"x\", 70000) = 70000
""",
                encoding="utf-8",
            )
            parsed = perf_resource_profile.parse_strace(path, 0)
        self.assertEqual(parsed["status"], "ok")
        self.assertEqual(parsed["syscalls"]["read"]["calls"], 2)
        self.assertEqual(parsed["syscalls"]["read"]["failed_calls"], 1)
        self.assertEqual(parsed["syscalls"]["read"]["returned_bytes"], 10)
        self.assertEqual(parsed["syscalls"]["write"]["calls"], 2)
        self.assertEqual(parsed["syscalls"]["write"]["returned_bytes"], 74096)
        self.assertEqual(parsed["syscalls"]["write"]["size_buckets"]["4096-16383"], 1)
        self.assertEqual(parsed["syscalls"]["write"]["size_buckets"]["65536-262143"], 1)

    def test_perf_parser_fail_closed_with_null_counters(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "perf.txt"
            stderr = Path(directory) / "stderr.txt"
            path.write_text("<not counted>,,cycles\n", encoding="utf-8")
            stderr.write_text("No permission to enable cycles event.\n", encoding="utf-8")
            parsed = perf_resource_profile.parse_perf_stat(path, 255, stderr)
        self.assertEqual(parsed["status"], "unsupported")
        self.assertIsNone(parsed["counters"]["cycles"]["value"])
        self.assertFalse(parsed["counters"]["cycles"]["available"])

    def test_scaling_analysis_is_explicit_and_finite(self):
        measurements = [
            {
                "elapsed_ns": {"p50": 1000},
                "execution": {"worker_count": 1, "logical_tasks": 16, "logical_bytes": 100},
            },
            {
                "elapsed_ns": {"p50": 600},
                "execution": {"worker_count": 2, "logical_tasks": 16, "logical_bytes": 100},
            },
            {
                "elapsed_ns": {"p50": 400},
                "execution": {"worker_count": 4, "logical_tasks": 16, "logical_bytes": 100},
            },
        ]
        summary = perf_resource_profile.scaling_analysis(measurements)
        self.assertEqual(summary["status"], "observed")
        self.assertEqual(summary["baseline_workers"], 1)
        self.assertEqual(summary["rows"][1]["speedup_vs_baseline"], 1000 / 600)
        self.assertTrue(summary["rows"][1]["amdahl_serial_fraction_valid"])
        self.assertIsInstance(summary["rows"][1]["amdahl_serial_fraction"], float)
        self.assertTrue(summary["classification"])
        json.dumps(summary, allow_nan=False)

    def test_scaling_analysis_nulls_invalid_amdahl_fraction_before_classifying(self):
        measurements = [
            {"elapsed_ns": {"p50": 1000}, "execution": {"worker_count": 1}},
            {"elapsed_ns": {"p50": 2000}, "execution": {"worker_count": 2}},
        ]
        summary = perf_resource_profile.scaling_analysis(measurements)
        row = summary["rows"][1]
        self.assertIsNone(row["amdahl_serial_fraction"])
        self.assertGreater(row["amdahl_serial_fraction_raw"], 1.0)
        self.assertFalse(row["amdahl_serial_fraction_valid"])
        self.assertEqual(row["amdahl_serial_fraction_reason"], "outside_0_1")
        self.assertEqual(summary["classification"], "nonideal_or_measurement_noise")
        json.dumps(summary, allow_nan=False)

    def test_prebuilt_identity_is_hash_only_and_requires_clean_build_rerun(self):
        with tempfile.TemporaryDirectory() as directory:
            binary = Path(directory) / "litchi-perf-baseline"
            binary.write_bytes(b"prebuilt")
            identity = perf_resource_profile.build_identity(
                binary,
                build_command=["cargo", "build", "--locked"],
                build_result={"executed": False},
                source_identity={"revision": "head", "head_tree": "tree"},
                pre_build_source_identity=None,
                rerun_command=["python3", "tools/perf_resource_profile.py", "run", "--build"],
            )
        self.assertEqual(identity["provenance"]["status"], "prebuilt_binary_hash_only")
        self.assertTrue(identity["provenance"]["rerun_required"])
        self.assertEqual(identity["source_content_identity"]["revision"], "head")
        self.assertEqual(identity["build_result"]["executed"], False)
        json.dumps(identity, allow_nan=False)

    def test_untracked_content_identity_changes_with_file_content(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = root / "scratch.txt"
            path.write_text("first", encoding="utf-8")
            link = root / "scratch.link"
            try:
                link.symlink_to("scratch.txt")
            except (OSError, NotImplementedError):
                self.skipTest("symlinks unavailable")
            status = b"?? scratch.txt\0?? scratch.link\0"
            first = perf_resource_profile.untracked_content_identity(status, repo_root=root)
            path.write_text("second", encoding="utf-8")
            second = perf_resource_profile.untracked_content_identity(status, repo_root=root)
        self.assertEqual(first["status"], "ok")
        first_entries = {entry["path"]: entry for entry in first["entries"]}
        second_entries = {entry["path"]: entry for entry in second["entries"]}
        self.assertEqual(first_entries["scratch.txt"]["kind"], "file")
        self.assertEqual(first_entries["scratch.link"]["kind"], "symlink")
        self.assertEqual(first["total_bytes"], len(b"first") + len(b"scratch.txt"))
        self.assertNotEqual(first_entries["scratch.txt"]["sha256"], second_entries["scratch.txt"]["sha256"])

    def test_untracked_content_identity_refuses_symlinked_ancestor(self):
        with tempfile.TemporaryDirectory() as directory, tempfile.TemporaryDirectory() as outside:
            root = Path(directory)
            secret = Path(outside) / "secret.txt"
            secret.write_text("outside", encoding="utf-8")
            try:
                (root / "linkdir").symlink_to(Path(outside), target_is_directory=True)
            except (OSError, NotImplementedError):
                self.skipTest("directory symlinks unavailable")
            result = perf_resource_profile.untracked_content_identity(
                b"?? linkdir/secret.txt\0",
                repo_root=root,
            )
        self.assertEqual(result["status"], "error")
        self.assertIn("symlinked ancestor", result["error"])
        self.assertEqual(result["entries"], [])

    def test_build_status_requires_success_and_complete_clean_matching_snapshots(self):
        snapshot = {
            "revision": "head",
            "snapshot_status": "complete",
            "git_worktree_dirty": False,
            "untracked_content": {"status": "ok"},
        }
        with tempfile.TemporaryDirectory() as directory:
            binary = Path(directory) / "litchi-perf-baseline"
            binary.write_bytes(b"built")
            common = {
                "build_command": ["cargo", "build"],
                "source_identity": snapshot,
                "pre_build_source_identity": snapshot.copy(),
                "rerun_command": ["python3", "tools/perf_resource_profile.py", "run", "--build"],
            }
            success = perf_resource_profile.build_identity(
                binary,
                build_result={"executed": True, "returncode": 0, "timed_out": False},
                **common,
            )
            failed = perf_resource_profile.build_identity(
                binary,
                build_result={"executed": True, "returncode": 1, "timed_out": False},
                **common,
            )
            timed_out = perf_resource_profile.build_identity(
                binary,
                build_result={"executed": True, "returncode": 0, "timed_out": True},
                **common,
            )
            dirty = dict(snapshot, git_worktree_dirty=True)
            dirty_identity = perf_resource_profile.build_identity(
                binary,
                build_result={"executed": True, "returncode": 0, "timed_out": False},
                build_command=["cargo", "build"],
                source_identity=dirty,
                pre_build_source_identity=snapshot,
                rerun_command=common["rerun_command"],
            )
        self.assertEqual(success["provenance"]["status"], "build_succeeded_matching_source_snapshots")
        self.assertEqual(failed["provenance"]["status"], "build_failed")
        self.assertEqual(timed_out["provenance"]["status"], "build_failed")
        self.assertEqual(dirty_identity["provenance"]["status"], "build_succeeded_source_snapshot_only")
        self.assertTrue(failed["provenance"]["rerun_required"])
        self.assertTrue(dirty_identity["provenance"]["rerun_required"])
        json.dumps(success, allow_nan=False)
        json.dumps(failed, allow_nan=False)
        json.dumps(dirty_identity, allow_nan=False)

    def test_logical_measurements_retain_identity_and_counters(self):
        report = {
            "results": [
                {
                    "case": "rtf_streaming_create",
                    "corpus": {
                        "name": "rtf-medium",
                        "generator": "rtf-v1",
                        "archive_sha256": "a" * 64,
                        "archive_bytes": 123,
                    },
                    "elapsed_ns": {"unit": "ns", "p50": 10, "p95": 11, "p99": 12},
                    "sink": {"accepted_bytes": 123, "write_calls": 4},
                }
            ]
        }
        rows = perf_resource_profile.logical_measurements(report)
        self.assertEqual(rows[0]["case"], "rtf_streaming_create")
        self.assertEqual(rows[0]["corpus"]["archive_sha256"], "a" * 64)
        self.assertEqual(rows[0]["sink"]["write_calls"], 4)
        self.assertNotIn("target_payload_sha256", rows[0]["corpus"])

    def test_artifact_reports_exact_retained_hash_and_size(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "retained.bin"
            path.write_bytes(b"retained-artifact")
            retained = perf_resource_profile.artifact(path, retained=True)
            missing = perf_resource_profile.artifact(
                Path(directory) / "missing.bin", retained=True
            )
        self.assertTrue(retained["present"])
        self.assertTrue(retained["retained"])
        self.assertEqual(retained["bytes"], len(b"retained-artifact"))
        self.assertEqual(
            retained["sha256"], perf_resource_profile.sha256_bytes(b"retained-artifact")
        )
        self.assertFalse(missing["present"])
        self.assertIsNone(missing["sha256"])

    def test_missing_heaptrack_is_not_zero(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "missing.txt"
            parsed = perf_resource_profile.parse_heaptrack_print(path)
        self.assertEqual(parsed["status"], "missing")
        self.assertNotIn("peak_heap_bytes", parsed)

    def test_heaptrack_summary_and_histogram_are_parsed(self):
        with tempfile.TemporaryDirectory() as directory:
            summary = Path(directory) / "print.txt"
            histogram = Path(directory) / "hist.tsv"
            summary.write_text(
                """\
calls to allocation functions: 12 (3/s)
MOST TEMPORARY ALLOCATIONS
34 temporary allocations of 56 allocations in total (60.7%)
11 temporary allocations of 12 allocations in total (91.7%) from
some::other::stack::group
temporary memory allocations: 57 (4/s)
peak heap memory consumption: 2.50M
peak RSS (including heaptrack overhead): 4.00M
""",
                encoding="utf-8",
            )
            histogram.write_text("4\t3\n8\t2\n", encoding="utf-8")
            parsed = perf_resource_profile.parse_heaptrack_print(summary)
            allocated = perf_resource_profile.parse_heaptrack_histogram(histogram)
        self.assertEqual(parsed["allocation_calls"], 12)
        self.assertEqual(parsed["temporary_allocations"], 57)
        self.assertEqual(parsed["peak_heap_bytes"], int(2.5 * 1024 * 1024))
        self.assertEqual(parsed["peak_rss_bytes"], 4 * 1024 * 1024)
        self.assertEqual(allocated, 28)

    def test_heaptrack_parser_rejects_stack_group_without_process_total(self):
        with tempfile.TemporaryDirectory() as directory:
            summary = Path(directory) / "stack-only.txt"
            summary.write_text(
                "MOST TEMPORARY ALLOCATIONS\n"
                "34 temporary allocations of 56 allocations in total (60.7%)\n",
                encoding="utf-8",
            )
            parsed = perf_resource_profile.parse_heaptrack_print(summary)
        self.assertEqual(parsed["status"], "unparsed")
        self.assertIsNone(parsed["temporary_allocations"])

    def test_heaptrack_histogram_rejects_partial_rows_and_strict_evidence_rejects_null_total(self):
        malformed_rows = (
            "4\t3\nmalformed\n",
            "4 3\n",
            "-1\t3\n",
            "4\t-3\n",
            "4\tnot-a-number\n",
            "4\t3.5\n",
            "4\t3\textra\n",
            "\n\t \n",
        )
        with tempfile.TemporaryDirectory() as directory:
            for index, contents in enumerate(malformed_rows):
                path = Path(directory) / f"malformed-{index}.tsv"
                path.write_text(contents, encoding="utf-8")
                self.assertIsNone(perf_resource_profile.parse_heaptrack_histogram(path))

            legs = self._xlsx_xml_borrowed_legs(directory)
            for leg in legs:
                leg.update(self._strict_xlsx_resource_fields())
            parsed = Path(directory) / "partial.tsv"
            parsed.write_text("4\t3\nmalformed\n", encoding="utf-8")
            self.assertIsNone(perf_resource_profile.parse_heaptrack_histogram(parsed))
            legs[0]["heaptrack"]["print"]["parsed"]["histogram_artifact"] = (
                perf_resource_profile.artifact(parsed, retained=True)
            )
            legs[0]["heaptrack"]["print"]["parsed"]["allocated_bytes"] = None
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "allocated_bytes is missing or non-finite",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_resource_legs(legs)

    def test_heaptrack_bytes_token_fails_closed_on_huge_or_nonfinite_units(self):
        self.assertEqual(perf_resource_profile._bytes_token("2.50M"), 2_621_440)
        self.assertIsNone(perf_resource_profile._bytes_token("9" * 500 + "T"))
        self.assertIsNone(perf_resource_profile._bytes_token("1e999T"))
        self.assertIsNone(perf_resource_profile._bytes_token("nanM"))
        self.assertIsNone(perf_resource_profile._bytes_token("infG"))

    def test_retained_docx_heaptrack_report_is_reprocessed_from_verified_artifacts(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            artifacts = root / "artifacts"
            legs = []
            totals = {"A1": 101, "B1": 99, "B2": 97, "A2": 103}
            for label in perf_resource_profile.ABBA_LEG_ORDER:
                leg_dir = artifacts / label.lower()
                leg_dir.mkdir(parents=True)
                summary = leg_dir / "heaptrack-print.txt"
                histogram = leg_dir / "heaptrack-histogram.tsv"
                summary.write_text(
                    "7 temporary allocations of 8 allocations in total (87.5%)\n"
                    f"temporary memory allocations: {totals[label]} (4/s)\n",
                    encoding="utf-8",
                )
                histogram.write_text("4\t3\n", encoding="utf-8")
                legs.append(
                    {
                        "leg": label,
                        "artifact_directory": str(leg_dir),
                        "heaptrack": {
                            "print": {
                                "artifact": perf_resource_profile.artifact(
                                    summary, retained=True
                                ),
                                "parsed": {
                                    "histogram_artifact": perf_resource_profile.artifact(
                                        histogram, retained=True
                                    )
                                },
                            }
                        },
                    }
                )
            source = root / "source.json"
            output = root / "reprocessed.json"
            source.write_text(
                json.dumps(
                    {
                        "tool": {"mode": "docx-semantic-abba-resource-profile"},
                        "scope": {
                            "cases": list(perf_resource_profile.DOCX_SEMANTIC_CASES)
                        },
                        "artifact_directory": str(artifacts),
                        "legs": legs,
                    }
                ),
                encoding="utf-8",
            )
            result = perf_resource_profile.run_reprocess_docx_heaptrack(
                argparse.Namespace(input=source, output=output)
            )
            published = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(result, 0)
        observed = {
            leg["leg"]: leg["heaptrack"]["print"]["parsed"]["temporary_allocations"]
            for leg in published["legs"]
        }
        self.assertEqual(observed, totals)
        metric = published["statistics"]["metrics"]["heaptrack.temporary_allocations"]
        self.assertEqual(metric["values_by_leg"], totals)
        self.assertTrue(published["reprocessing"]["raw_heaptrack_artifacts_verified"])

    def test_filesystem_evidence_omits_raw_request_arrays(self):
        compact = perf_resource_profile.compact_filesystem_evidence(
            [
                {
                    "case": "cfb-save",
                    "samples": [
                        {
                            "logical_read_calls": 4,
                            "logical_read_request_sizes": [1, 2, 3, 4],
                            "logical_read_request_size_buckets": {"1-511": 4},
                        }
                    ],
                }
            ]
        )
        sample = compact[0]["samples"][0]
        self.assertNotIn("logical_read_request_sizes", sample)
        self.assertEqual(sample["logical_read_calls"], 4)
        self.assertEqual(sample["logical_read_request_size_buckets"]["1-511"], 4)

    @staticmethod
    def _abba_report(revision="1" * 40, *, corpus_shape="medium", samples=3):
        return {
            "schema_version": perf_resource_profile.SCHEMA_VERSION,
            "tool": {
                "name": "litchi-perf-baseline",
                "version": "0.1.0",
                "profile": "release",
                "target_os": "linux",
                "target_arch": "x86_64",
            },
            "environment": {
                "git_revision": revision,
                "git_worktree_dirty": False,
                "rustc_version": "rustc 1.95.0 (test)",
                "logical_cpus_available": 8,
                "allocator": "Rust system allocator",
                "rustflags": None,
                "cargo_build_target": None,
                "perf_event_paranoid": "4",
                "os": "linux",
                "kernel": "test-kernel",
                "cpu_model": "test-cpu",
                "total_memory_bytes": 1_000_000,
                "page_size_bytes": 4096,
                "filesystem_type": None,
                "source_destination_same_device": None,
                "cpu_affinity": None,
                "storage_identifier": None,
            },
            "configuration": {
                "samples_per_case": samples,
                "warmup_iterations_per_case": 1,
                "cases": [perf_resource_profile.XLSX_MANAGED_BATCH_CASE],
                "xlsx_cell_crud_shapes": [corpus_shape],
            },
            "results": [
                {
                    "case": perf_resource_profile.XLSX_MANAGED_BATCH_CASE,
                    "corpus": {
                        "name": "xlsx-test",
                        "shape": corpus_shape,
                        "archive_sha256": "a" * 64,
                    },
                    "elapsed_ns": {
                        "unit": "ns",
                        "samples": [100, 105, 110],
                        "min": 100,
                        "p50": 100,
                        "p95": 110,
                        "p99": 120,
                        "max": 120,
                        "mean": 105,
                        "standard_deviation": 5,
                        "confidence_interval_95": {
                            "method": "two-sided Student's t interval for the mean",
                            "lower": 90.0,
                            "upper": 120.0,
                        },
                    },
                }
            ],
        }

    @classmethod
    def _abba_legs(cls, root):
        binaries = {}
        for variant, content in (
            ("control", b"control-binary"),
            ("candidate", b"candidate-binary"),
        ):
            path = Path(root) / f"{variant}-binary"
            path.write_bytes(content)
            path.chmod(0o755)
            binaries[variant] = perf_resource_profile.binary_identity(path, label=variant)
        legs = []
        for leg in perf_resource_profile.ABBA_LEG_ORDER:
            variant = perf_resource_profile.ABBA_LEG_VARIANTS[leg]
            revision = "1" * 40 if variant == "control" else "2" * 40
            legs.append(
                {
                    "leg": leg,
                    "variant": variant,
                    "binary_identity": dict(binaries[variant]),
                    "harness_report": cls._abba_report(revision),
                }
            )
        return legs

    def test_abba_order_is_fixed_and_rejects_reordering(self):
        self.assertEqual(
            perf_resource_profile.validate_abba_order(["A1", "B1", "B2", "A2"]),
            ("A1", "B1", "B2", "A2"),
        )
        with self.assertRaisesRegex(
            perf_resource_profile.ResourceProfileInputError, "leg order"
        ):
            perf_resource_profile.validate_abba_order(["A1", "B2", "B1", "A2"])

    def test_abba_validation_rejects_dirty_identical_and_mismatched_inputs(self):
        with tempfile.TemporaryDirectory() as directory:
            legs = self._abba_legs(directory)
            validated = perf_resource_profile.validate_abba_inputs(
                legs,
                expected_configuration={
                    "samples_per_case": 3,
                    "warmup_iterations_per_case": 1,
                    "cases": [perf_resource_profile.XLSX_MANAGED_BATCH_CASE],
                    "xlsx_cell_crud_shapes": ["medium"],
                },
            )
            self.assertEqual(validated["status"], "validated")
            self.assertNotEqual(
                validated["control_binary_sha256"], validated["candidate_binary_sha256"]
            )
            uppercase = self._abba_legs(directory)
            uppercase[1]["binary_identity"]["binary_sha256"] = uppercase[1]["binary_identity"][
                "binary_sha256"
            ].upper()
            uppercase[2]["binary_identity"]["binary_sha256"] = uppercase[2]["binary_identity"][
                "binary_sha256"
            ].upper()
            self.assertEqual(perf_resource_profile.validate_abba_inputs(uppercase)["status"], "validated")
            dirty = self._abba_legs(directory)
            dirty[0]["harness_report"]["environment"]["git_worktree_dirty"] = True
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "dirty"
            ):
                perf_resource_profile.validate_abba_inputs(dirty)

            identical = self._abba_legs(directory)
            for key in ("path", "binary_sha256", "binary_bytes", "executable"):
                identical[1]["binary_identity"][key] = identical[0]["binary_identity"][key]
                identical[2]["binary_identity"][key] = identical[0]["binary_identity"][key]
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "identical"
            ):
                perf_resource_profile.validate_abba_inputs(identical)

            mismatched = self._abba_legs(directory)
            mismatched[2]["harness_report"]["results"][0]["corpus"]["archive_sha256"] = "b" * 64
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "corpora do not match"
            ):
                perf_resource_profile.validate_abba_inputs(mismatched)

            malformed = self._abba_legs(directory)
            malformed[1]["binary_identity"]["binary_sha256"] = "z" * 64
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "hexadecimal"
            ):
                perf_resource_profile.validate_abba_inputs(malformed)

            missing = self._abba_legs(directory)
            missing[1]["binary_identity"]["path"] = str(Path(directory) / "missing")
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "not readable"
            ):
                perf_resource_profile.validate_abba_inputs(missing)

            mode_mismatch = self._abba_legs(directory)
            Path(mode_mismatch[1]["binary_identity"]["path"]).chmod(0o700)
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "mode_bits"
            ):
                perf_resource_profile.validate_abba_inputs(mode_mismatch)

    def test_abba_validation_binds_tool_and_stable_environment_identity(self):
        mutations = (
            (
                "tool target",
                lambda report: report["tool"].update(target_arch="aarch64"),
                "tool identities do not match",
            ),
            (
                "environment CPU",
                lambda report: report["environment"].update(cpu_model="other-cpu"),
                "stable environment identities do not match",
            ),
        )
        for name, mutate, expected in mutations:
            with self.subTest(name=name), tempfile.TemporaryDirectory() as directory:
                legs = self._abba_legs(directory)
                mutate(legs[1]["harness_report"])
                with self.assertRaisesRegex(
                    perf_resource_profile.ResourceProfileInputError, expected
                ):
                    perf_resource_profile.validate_abba_inputs(legs)

    def test_abba_statistics_are_descriptive_and_preserve_not_measured_dimensions(self):
        with tempfile.TemporaryDirectory() as directory:
            legs = self._abba_legs(directory)
            resource_values = {
                "A1": {"harness.elapsed_ns.p50": 100, "time.max_rss_kib": 1000},
                "B1": {"harness.elapsed_ns.p50": 80, "time.max_rss_kib": 1100},
                "B2": {"harness.elapsed_ns.p50": 90, "time.max_rss_kib": 1050},
                "A2": {"harness.elapsed_ns.p50": 110, "time.max_rss_kib": 1000},
            }
            for leg in legs:
                leg["resource_metrics"] = resource_values[leg["leg"]]
            report = perf_resource_profile.abba_statistics(legs)
        elapsed = report["metrics"]["harness.elapsed_ns.p50"]
        self.assertEqual(elapsed["control"]["median"], 105.0)
        self.assertEqual(elapsed["candidate"]["median"], 85.0)
        self.assertAlmostEqual(
            elapsed["paired"]["A1_control_to_B1_candidate"]["relative_delta_percent"],
            -20.0,
        )
        self.assertIn("no automatic speedup claim", elapsed["claim"])
        self.assertEqual(report["not_measured"]["decompressed_bytes"].split(":", 1)[0], "not measured")
        self.assertIsNone(
            report["metrics"]["heaptrack.allocation_calls"]["control"]["median"]
        )
        json.dumps(report, allow_nan=False)

    def test_abba_statistics_null_overflowing_extreme_values(self):
        with tempfile.TemporaryDirectory() as directory:
            legs = self._abba_legs(directory)
            extreme_values = {
                "A1": {"harness.elapsed_ns.p50": 1e-308},
                "B1": {"harness.elapsed_ns.p50": 1e308},
                "B2": {"harness.elapsed_ns.p50": 1e308},
                "A2": {"harness.elapsed_ns.p50": 1e-308},
            }
            for leg in legs:
                leg["resource_metrics"] = extreme_values[leg["leg"]]
            report = perf_resource_profile.abba_statistics(legs)
        metric = report["metrics"]["harness.elapsed_ns.p50"]
        self.assertEqual(metric["control"]["status"], "observed")
        self.assertEqual(metric["control"]["mean"], 1e-308)
        self.assertEqual(metric["candidate"]["status"], "observed_with_overflow")
        self.assertIsNone(metric["candidate"]["mean"])
        self.assertIsNone(metric["candidate"]["median"])
        self.assertEqual(
            metric["paired"]["A1_control_to_B1_candidate"]["status"], "overflow"
        )
        self.assertIsNone(
            metric["paired"]["A1_control_to_B1_candidate"]["relative_delta_percent"]
        )
        json.dumps(report, allow_nan=False)

    def test_abba_reserves_fresh_output_and_artifact_paths(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            output = root / "result.json"
            artifacts = root / "artifacts"
            reserved_output, reserved_artifacts = perf_resource_profile.reserve_abba_paths(
                output, artifacts
            )
            self.assertEqual(reserved_output, output.resolve())
            self.assertEqual(reserved_artifacts, artifacts.resolve())
            self.assertTrue(artifacts.is_dir())

            output.write_text("existing", encoding="utf-8")
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "output path already exists"
            ):
                perf_resource_profile.reserve_abba_paths(output, root / "fresh-artifacts")

            stale = root / "stale-artifacts"
            stale.mkdir()
            (stale / "old-capture.zst").write_bytes(b"stale")
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "stale capture"
            ):
                perf_resource_profile.reserve_abba_paths(root / "fresh.json", stale)

            empty_existing = root / "empty-existing-artifacts"
            empty_existing.mkdir()
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "already exists"
            ):
                perf_resource_profile.reserve_abba_paths(root / "another.json", empty_existing)

    def test_abba_time_status_surfaces_missing_and_unparsed_reports(self):
        for contents, expected_status in ((None, "missing"), ("not time output\n", "unparsed")):
            with self.subTest(expected_status=expected_status), tempfile.TemporaryDirectory() as directory:
                root = Path(directory)
                binary = root / "control-binary"
                binary.write_bytes(b"control")
                binary.chmod(0o755)
                artifact_root = root / "artifacts"
                harness_dir = artifact_root / "a1"
                harness_dir.mkdir(parents=True)
                (harness_dir / "harness.json").write_text(
                    json.dumps(self._abba_report()), encoding="utf-8"
                )

                def fake_run(command, *, stdout_path, stderr_path, timeout_seconds):
                    del timeout_seconds
                    if contents is not None and "-o" in command:
                        time_path = Path(command[command.index("-o") + 1])
                        time_path.write_text(contents, encoding="utf-8")
                    return {
                        "command": list(command),
                        "returncode": 0,
                        "timed_out": False,
                        "wall_ns": 1,
                        "stdout": {},
                        "stderr": {},
                        "stderr_excerpt": None,
                    }

                descriptor = perf_resource_profile.binary_identity(binary, label="control")
                with mock.patch.object(
                    perf_resource_profile, "run_command", side_effect=fake_run
                ), mock.patch.object(
                    perf_resource_profile,
                    "_profile_abba_heaptrack",
                    return_value={"status": "unsupported"},
                ):
                    leg = perf_resource_profile.profile_xlsx_abba_leg(
                        leg="A1",
                        variant="control",
                        binary=binary,
                        binary_descriptor=descriptor,
                        artifact_root=artifact_root,
                        warmup=1,
                        samples=3,
                        tools={"time": {"available": True, "path": "/usr/bin/time"}},
                        timeout_seconds=1,
                    )
                self.assertEqual(leg["time"]["status"], expected_status)
                self.assertEqual(leg["time"]["parsed"]["status"], expected_status)

    def test_heaptrack_outer_status_surfaces_print_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)

            def fake_run(command, *, stdout_path, stderr_path, timeout_seconds):
                del timeout_seconds
                if "--record-only" in command:
                    (root / "heaptrack-profile.zst").write_bytes(b"capture")
                    returncode = 0
                else:
                    stdout_path.write_text("not heaptrack output\n", encoding="utf-8")
                    returncode = 1
                return {
                    "command": list(command),
                    "returncode": returncode,
                    "timed_out": False,
                    "wall_ns": 1,
                    "stdout": {},
                    "stderr": {},
                    "stderr_excerpt": None,
                }

            with mock.patch.object(perf_resource_profile, "run_command", side_effect=fake_run):
                result = perf_resource_profile._profile_abba_heaptrack(
                    root / "unused-binary",
                    perf_resource_profile.XLSX_MANAGED_BATCH_ARGS,
                    root,
                    {
                        "heaptrack": {"available": True, "path": "heaptrack"},
                        "heaptrack_print": {"available": True, "path": "heaptrack_print"},
                    },
                    warmup=1,
                    samples=3,
                    timeout_seconds=1,
                )
        self.assertEqual(result["status"], "failed")
        self.assertEqual(result["failure_stage"], "heaptrack_print")
        self.assertEqual(result["print"]["status"], "failed")

    def test_heaptrack_failed_run_is_not_relabelled_unsupported(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)

            def fake_run(command, *, stdout_path, stderr_path, timeout_seconds):
                del stdout_path, stderr_path, timeout_seconds
                (root / "heaptrack-profile.zst").write_bytes(b"partial-capture")
                return {
                    "command": list(command),
                    "returncode": 1,
                    "timed_out": False,
                    "wall_ns": 1,
                    "stdout": {},
                    "stderr": {},
                    "stderr_excerpt": "heaptrack failed",
                }

            with mock.patch.object(perf_resource_profile, "run_command", side_effect=fake_run):
                result = perf_resource_profile._profile_abba_heaptrack(
                    root / "unused-binary",
                    perf_resource_profile.XLSX_MANAGED_BATCH_ARGS,
                    root,
                    {
                        "heaptrack": {"available": True, "path": "heaptrack"},
                        "heaptrack_print": {"available": False},
                    },
                    warmup=1,
                    samples=3,
                    timeout_seconds=1,
                )
        self.assertEqual(result["status"], "failed")
        self.assertEqual(result["failure_stage"], "heaptrack")
        self.assertEqual(result["print"]["status"], "unsupported")

    def test_missing_optional_tools_remain_unsupported_without_zeroes(self):
        with tempfile.TemporaryDirectory() as directory:
            result = perf_resource_profile._profile_abba_heaptrack(
                Path(directory) / "unused-binary",
                perf_resource_profile.XLSX_MANAGED_BATCH_ARGS,
                Path(directory),
                {
                    "heaptrack": {"available": False},
                    "heaptrack_print": {"available": False},
                },
                warmup=1,
                samples=3,
                timeout_seconds=1,
            )
        self.assertEqual(result["status"], "unsupported")
        self.assertNotIn("allocation_calls", result)
        self.assertNotIn("peak_heap_bytes", result)
        self.assertNotIn("allocated_bytes", result)

    @staticmethod
    def _docx_report(
        revision="1" * 40,
        *,
        samples=3,
        shape="large",
        cases=perf_resource_profile.DOCX_SEMANTIC_CASES,
    ):
        cases = tuple(cases)
        configuration = {
            "samples_per_case": samples,
            "warmup_iterations_per_case": 1,
            "cases": list(cases),
            "semantic_shapes": [shape],
        }
        corpus = {
            "name": f"docx-semantic-{shape}",
            "generator": "litchi-docx-semantic-v1",
            "package_format": "DOCX/OPC/ZIP",
            "shape": shape,
            "payload_kind": "deterministic-semantic-text",
            "compression": "deflate",
            "entry_count": 10_000,
            "archive_member_count": 8,
            "entry_bytes": 49,
            "uncompressed_payload_bytes": 490_000,
            "archive_bytes": 50_000,
            "archive_sha256": "a" * 64,
            "target_entry": "paragraph:0",
            "target_payload_bytes": 49,
            "target_payload_sha256": "b" * 64,
            "rtf_variant": None,
            "xlsx": None,
        }
        results = []
        for offset, case in enumerate(cases):
            results.append(
                {
                    "case": case,
                    "corpus": dict(corpus),
                    "elapsed_ns": {
                        "unit": "ns",
                        "samples": [100 + offset * 10] * samples,
                        "min": 100 + offset * 10,
                        "p50": 100 + offset * 10,
                        "p95": 110 + offset * 10,
                        "p99": 120 + offset * 10,
                        "max": 120 + offset * 10,
                        "mean": 105 + offset * 10,
                        "standard_deviation": 5,
                        "confidence_interval_95": {
                            "method": "two-sided Student's t interval for the mean",
                            "lower": 90 + offset * 10,
                            "upper": 120 + offset * 10,
                        },
                    },
                }
            )
        return {
            "schema_version": perf_resource_profile.SCHEMA_VERSION,
            "tool": {
                "name": "litchi-perf-baseline",
                "version": "0.1.0",
                "profile": "release",
                "target_os": "linux",
                "target_arch": "x86_64",
            },
            "environment": {
                "git_revision": revision,
                "git_worktree_dirty": False,
                "rustc_version": "rustc 1.95.0 (test)",
                "logical_cpus_available": 8,
                "allocator": "Rust system allocator",
                "rustflags": None,
                "cargo_build_target": None,
                "perf_event_paranoid": "4",
                "os": "linux",
                "kernel": "test-kernel",
                "cpu_model": "test-cpu",
                "total_memory_bytes": 1_000_000,
                "page_size_bytes": 4096,
                "filesystem_type": None,
                "source_destination_same_device": None,
                "cpu_affinity": None,
                "storage_identifier": None,
            },
            "configuration": configuration,
            "results": results,
        }

    @classmethod
    def _docx_legs(
        cls,
        root,
        cases=perf_resource_profile.DOCX_SEMANTIC_CASES,
    ):
        binaries = {}
        for variant, content in (
            ("control", b"docx-control-binary"),
            ("candidate", b"docx-candidate-binary"),
        ):
            path = Path(root) / f"{variant}-docx-binary"
            path.write_bytes(content)
            path.chmod(0o755)
            binaries[variant] = perf_resource_profile.binary_identity(path, label=variant)
        legs = []
        for leg in perf_resource_profile.ABBA_LEG_ORDER:
            variant = perf_resource_profile.ABBA_LEG_VARIANTS[leg]
            revision = "1" * 40 if variant == "control" else "2" * 40
            legs.append(
                {
                    "leg": leg,
                    "variant": variant,
                    "binary_identity": dict(binaries[variant]),
                    "harness_report": cls._docx_report(revision, cases=cases),
                }
            )
        return legs

    @staticmethod
    def _xlsx_xml_borrowed_report(
        revision="1" * 40,
        *,
        samples=3,
        cases=perf_resource_profile.XLSX_XML_BORROWED_CASES,
    ):
        cases = tuple(cases)
        report = ResourceProfileParserTests._abba_report(
            revision, corpus_shape="tiny", samples=samples
        )
        report["configuration"] = {
            "samples_per_case": samples,
            "warmup_iterations_per_case": 1,
            "cases": list(cases),
            "xlsx_shapes": ["tiny"],
            "xlsx_cell_crud_shapes": ["medium"],
        }
        tiny_corpus = copy.deepcopy(XLSX_TINY_CORPUS_FIXTURE)
        medium_corpus = copy.deepcopy(XLSX_MEDIUM_CORPUS_FIXTURE)
        results = []
        for offset, case in enumerate(cases):
            result = {
                "case": case,
                "corpus": dict(
                    medium_corpus
                    if case
                    in {
                        "xlsx_eager_cell_values_one_edit_save",
                        "xlsx_source_backed_cell_values_one_edit_save",
                    }
                    else tiny_corpus
                ),
                "elapsed_ns": {
                    "unit": "ns",
                    "samples": [100 + offset] * samples,
                    "min": 100 + offset,
                    "p50": 100 + offset,
                    "p95": 110 + offset,
                    "p99": 120 + offset,
                    "max": 120 + offset,
                    "mean": 105 + offset,
                    "standard_deviation": 5,
                    "confidence_interval_95": {
                        "method": "two-sided Student's t interval for the mean",
                        "lower": 90 + offset,
                        "upper": 120 + offset,
                    },
                },
                "sink": None,
            }
            if case == "xlsx_source_first_cell":
                result["source"] = {"read_calls": [1] * samples, "read_bytes": [32] * samples}
            elif case == "xlsx_eager_cell_values_one_edit_save":
                result["sink"] = {
                    "accepted_bytes": 100,
                    "write_calls": 2,
                    "largest_write": 64,
                }
                result["output_sha256"] = "d" * 64
            elif case == "xlsx_source_backed_cell_values_one_edit_save":
                result["source"] = {
                    "read_calls": [2] * samples,
                    "read_bytes": [64] * samples,
                    "xlsx_cell_values": {
                        "open_ns": [10] * samples,
                        "plan_ns": [20] * samples,
                        "commit_ns": [30] * samples,
                        "publication_ns": [40] * samples,
                        "reopen_ns": [50] * samples,
                        "source_read_bytes": [64] * samples,
                    },
                }
                result["sink"] = {
                    "accepted_bytes": 120,
                    "write_calls": 3,
                    "largest_write": 64,
                }
                result["output_sha256"] = "e" * 64
            results.append(result)
        report["results"] = results
        return report

    @classmethod
    def _xlsx_xml_borrowed_legs(cls, root, *, samples=3):
        binaries = {}
        for variant, content in (
            ("control", b"xlsx-borrowed-control-binary"),
            ("candidate", b"xlsx-borrowed-candidate-binary"),
        ):
            path = Path(root) / f"{variant}-xlsx-borrowed-binary"
            path.write_bytes(content)
            path.chmod(0o755)
            binaries[variant] = perf_resource_profile.binary_identity(path, label=variant)
        legs = []
        for leg in perf_resource_profile.ABBA_LEG_ORDER:
            variant = perf_resource_profile.ABBA_LEG_VARIANTS[leg]
            revision = "1" * 40 if variant == "control" else "2" * 40
            legs.append(
                {
                    "leg": leg,
                    "variant": variant,
                    "binary_identity": dict(binaries[variant]),
                    "harness_report": cls._xlsx_xml_borrowed_report(
                        revision, samples=samples
                    ),
                }
            )
        return legs

    @staticmethod
    def _strict_xlsx_resource_fields(value=100, *, artifact_root=None):
        if artifact_root is None:
            retained = {
                "present": True,
                "retained": True,
                "sha256": "a" * 64,
                "bytes": 1,
            }
        else:
            artifact_root = Path(artifact_root)
            artifact_root.mkdir(parents=True, exist_ok=True)
            evidence_path = artifact_root / "strict-evidence.bin"
            evidence_path.write_bytes(b"strict-xlsx-evidence")
            retained = perf_resource_profile.artifact(evidence_path, retained=True)
        parsed_time = {
            "status": "ok",
            "max_rss_kib": value,
            "user_seconds": 1.0,
            "system_seconds": 1.0,
            "elapsed_wall_seconds": 1.0,
            "voluntary_context_switches": value,
            "involuntary_context_switches": value,
            "major_page_faults": value,
            "minor_page_faults": value,
            "artifact": dict(retained),
        }
        parsed_heaptrack = {
            "status": "ok",
            "allocation_calls": value,
            "allocated_bytes": value,
            "temporary_allocations": value,
            "peak_heap_bytes": value,
            "peak_rss_bytes": value,
            "histogram_artifact": dict(retained),
        }
        successful_run = {
            "returncode": 0,
            "timed_out": False,
            "stdout": dict(retained),
            "stderr": dict(retained),
        }
        return {
            "time": {
                "status": "ok",
                "run": dict(successful_run),
                "parsed": parsed_time,
            },
            "heaptrack": {
                "status": "ok",
                "harness_identity": {"status": "validated"},
                "run": dict(successful_run),
                "harness": dict(retained),
                "capture": dict(retained),
                "print": {
                    "status": "ok",
                    "run": dict(successful_run),
                    "artifact": dict(retained),
                    "parsed": parsed_heaptrack,
                },
            },
        }

    @staticmethod
    def _set_strict_xlsx_metric(leg, metric, value):
        if metric.startswith("time."):
            leg["time"]["parsed"][metric.split(".", 1)[1]] = value
        elif metric.startswith("heaptrack."):
            leg["heaptrack"]["print"]["parsed"][metric.split(".", 1)[1]] = value
        else:
            raise AssertionError(f"unexpected strict XLSX metric: {metric}")

    def test_docx_parser_exposes_explicit_comparison_cli(self):
        parsed = perf_resource_profile.build_parser().parse_args(
            [
                "compare-docx-semantic",
                "--control-binary",
                "/tmp/control",
                "--candidate-binary",
                "/tmp/candidate",
            ]
        )
        self.assertEqual(parsed.command, "compare-docx-semantic")
        self.assertIs(parsed.function, perf_resource_profile.run_docx_semantic_abba)
        optional = perf_resource_profile.build_parser().parse_args(
            [
                "compare-docx-semantic",
                "--control-binary",
                "/tmp/control",
                "--candidate-binary",
                "/tmp/candidate",
                "--include-one-paragraph-text",
            ]
        )
        self.assertTrue(optional.include_one_paragraph_text)
        run = perf_resource_profile.build_parser().parse_args(
            [
                "run",
                "--workload",
                perf_resource_profile.DOCX_SEMANTIC_ID,
                "--only",
                perf_resource_profile.DOCX_SEMANTIC_ID,
                "--control-binary",
                "/tmp/control",
                "--candidate-binary",
                "/tmp/candidate",
            ]
        )
        self.assertEqual(run.workload, perf_resource_profile.DOCX_SEMANTIC_ID)
        run_optional = perf_resource_profile.build_parser().parse_args(
            [
                "run",
                "--workload",
                perf_resource_profile.DOCX_SEMANTIC_ID,
                "--control-binary",
                "/tmp/control",
                "--candidate-binary",
                "/tmp/candidate",
                "--include-docx-one-paragraph-text",
            ]
        )
        self.assertTrue(run_optional.include_one_paragraph_text)
        alias = perf_resource_profile.build_parser().parse_args(
            [
                "run",
                "--workload",
                "docx-semantic-full-text",
                "--control-binary",
                "/tmp/control",
                "--candidate-binary",
                "/tmp/candidate",
            ]
        )
        self.assertEqual(alias.workload, "docx-semantic-full-text")

    def test_xlsx_xml_borrowed_parser_exposes_fixed_cli_and_defaults(self):
        parsed = perf_resource_profile.build_parser().parse_args(
            [
                "compare-xlsx-xml-borrowed",
                "--control-binary",
                "/tmp/control",
                "--candidate-binary",
                "/tmp/candidate",
            ]
        )
        self.assertEqual(parsed.command, "compare-xlsx-xml-borrowed")
        self.assertIs(parsed.function, perf_resource_profile.run_xlsx_xml_borrowed_abba)
        self.assertEqual(parsed.warmup, 30)
        self.assertEqual(parsed.samples, 500)
        run = perf_resource_profile.build_parser().parse_args(
            [
                "run",
                "--workload",
                perf_resource_profile.XLSX_XML_BORROWED_ID,
                "--control-binary",
                "/tmp/control",
                "--candidate-binary",
                "/tmp/candidate",
            ]
        )
        perf_resource_profile._apply_run_sampling_defaults(run)
        self.assertEqual(run.warmup, 30)
        self.assertEqual(run.samples, 500)
        self.assertEqual(
            perf_resource_profile.xlsx_xml_borrowed_args(),
            (
                "--case",
                ",".join(perf_resource_profile.XLSX_XML_BORROWED_CASES),
                "--xlsx-shape",
                "tiny",
                "--xlsx-cell-crud-shape",
                "medium",
            ),
        )

    def test_xlsx_xml_borrowed_validation_binds_four_rows_and_identity_channels(self):
        with tempfile.TemporaryDirectory() as directory:
            for shape, corpus in (
                ("tiny", XLSX_TINY_CORPUS_FIXTURE),
                ("medium", XLSX_MEDIUM_CORPUS_FIXTURE),
            ):
                digest = hashlib.sha256(
                    json.dumps(corpus, sort_keys=True, separators=(",", ":")).encode()
                ).hexdigest()
                self.assertEqual(digest, XLSX_CORPUS_CANONICAL_DIGESTS[shape])
            expected = {
                "samples_per_case": 3,
                "warmup_iterations_per_case": 1,
                "cases": list(perf_resource_profile.XLSX_XML_BORROWED_CASES),
                "xlsx_shapes": ["tiny"],
                "xlsx_cell_crud_shapes": ["medium"],
            }
            legs = self._xlsx_xml_borrowed_legs(directory)
            validated = perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                legs, expected_configuration=expected
            )
            self.assertEqual(validated["workload"], perf_resource_profile.XLSX_XML_BORROWED_ID)
            self.assertEqual(
                [item["case"] for item in validated["corpus_identities"]],
                list(perf_resource_profile.XLSX_XML_BORROWED_CASES),
            )
            self.assertEqual(
                [item["identity_sha256"] for item in validated["corpus_identities"]],
                [
                    XLSX_CORPUS_CANONICAL_DIGESTS["tiny"],
                    XLSX_CORPUS_CANONICAL_DIGESTS["tiny"],
                    XLSX_CORPUS_CANONICAL_DIGESTS["medium"],
                    XLSX_CORPUS_CANONICAL_DIGESTS["medium"],
                ],
            )
            self.assertEqual(
                [item["case"] for item in validated["result_identities"]],
                list(perf_resource_profile.XLSX_XML_BORROWED_CASES),
            )
            self.assertFalse(validated["result_identities"][0]["source_present"])
            self.assertTrue(validated["result_identities"][1]["source_present"])
            self.assertTrue(validated["result_identities"][3]["output_present"])

            reordered = self._xlsx_xml_borrowed_legs(directory)
            reordered[0]["harness_report"]["results"].reverse()
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "cases must be"
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(reordered)

            source_mismatch = self._xlsx_xml_borrowed_legs(directory)
            source_mismatch[1]["harness_report"]["results"][1]["source"]["read_calls"][
                0
            ] = 999
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "source/sink/output identities",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    source_mismatch
                )

            output_mismatch = self._xlsx_xml_borrowed_legs(directory)
            output_mismatch[2]["harness_report"]["results"][2]["output_sha256"] = "f" * 64
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "source/sink/output identities",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    output_mismatch
                )

            sink_mismatch = self._xlsx_xml_borrowed_legs(directory)
            sink_mismatch[0]["harness_report"]["results"][2]["sink"][
                "write_calls"
            ] = 99
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "source/sink/output identities",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    sink_mismatch
                )

            corpus_mismatch = self._xlsx_xml_borrowed_legs(directory)
            corpus_mismatch[3]["harness_report"]["results"][0]["corpus"][
                "archive_sha256"
            ] = "9" * 64
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "pinned",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    corpus_mismatch
                )

            altered_everywhere = self._xlsx_xml_borrowed_legs(directory)
            for leg in altered_everywhere:
                for result in leg["harness_report"]["results"]:
                    result["corpus"]["archive_sha256"] = "9" * 64
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "pinned",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    altered_everywhere
                )

            present_null_source = self._xlsx_xml_borrowed_legs(directory)
            present_null_source[0]["harness_report"]["results"][0]["source"] = None
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "source key presence",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    present_null_source
                )

            absent_source = self._xlsx_xml_borrowed_legs(directory)
            absent_source[0]["harness_report"]["results"][1].pop("source")
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "source key presence",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    absent_source
                )

            absent_sink = self._xlsx_xml_borrowed_legs(directory)
            absent_sink[0]["harness_report"]["results"][0].pop("sink")
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "sink key is required",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    absent_sink
                )

            present_null_output = self._xlsx_xml_borrowed_legs(directory)
            present_null_output[0]["harness_report"]["results"][2][
                "output_sha256"
            ] = None
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "output_sha256 value presence",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_abba_inputs(
                    present_null_output
                )

    def test_xlsx_xml_borrowed_metrics_are_case_specific_and_resource_only(self):
        report = self._xlsx_xml_borrowed_report()
        metrics = perf_resource_profile.instrumented_harness_metrics(
            report, perf_resource_profile.XLSX_XML_BORROWED_CASES
        )
        self.assertEqual(metrics["harness.xlsx_first_cell.elapsed_ns.p50"], 100)
        self.assertEqual(
            metrics[
                "harness.xlsx_source_backed_cell_values_one_edit_save.elapsed_ns.p50"
            ],
            103,
        )
        with tempfile.TemporaryDirectory() as directory:
            fallback_legs = self._xlsx_xml_borrowed_legs(directory)
            fallback_stats = perf_resource_profile.abba_statistics(
                fallback_legs,
                metric_specs=(
                    (
                        "harness.xlsx_source_first_cell.elapsed_ns.p50",
                        "instrumented source-first-cell elapsed; not latency evidence",
                    ),
                ),
            )
            self.assertEqual(
                fallback_stats["metrics"][
                    "harness.xlsx_source_first_cell.elapsed_ns.p50"
                ]["values_by_leg"]["A1"],
                101,
            )
            legs = self._xlsx_xml_borrowed_legs(directory)
            for leg in legs:
                leg["resource_metrics"] = {
                    "harness.xlsx_first_cell.elapsed_ns.p50": 100,
                    "time.max_rss_kib": 1024,
                    "heaptrack.allocation_calls": 200,
                }
            stats = perf_resource_profile.abba_statistics(
                legs,
                metric_specs=perf_resource_profile.XLSX_XML_BORROWED_RESOURCE_METRIC_SPECS,
            )
        self.assertIn("time.max_rss_kib", stats["metrics"])
        self.assertIn("heaptrack.allocation_calls", stats["metrics"])
        self.assertIn(
            "not latency evidence",
            stats["metrics"]["harness.xlsx_first_cell.elapsed_ns.p50"]["description"],
        )
        self.assertIn(
            "whole-process",
            stats["metrics"]["time.max_rss_kib"]["description"],
        )
        json.dumps(stats, allow_nan=False)

    def test_xlsx_xml_borrowed_resource_evidence_requires_complete_retained_process_totals(self):
        with tempfile.TemporaryDirectory() as directory:
            legs = self._xlsx_xml_borrowed_legs(directory)
            for leg in legs:
                leg.update(self._strict_xlsx_resource_fields())
            evidence = perf_resource_profile.validate_xlsx_xml_borrowed_resource_legs(legs)
            self.assertEqual(evidence["status"], "validated")
            self.assertEqual(len(evidence["legs"]), 4)

            failed_time = self._xlsx_xml_borrowed_legs(directory)
            for leg in failed_time:
                leg.update(self._strict_xlsx_resource_fields())
            failed_time[0]["time"]["run"]["returncode"] = 1
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "time.run must complete successfully",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_resource_legs(failed_time)

            missing_time = self._xlsx_xml_borrowed_legs(directory)
            for leg in missing_time:
                leg.update(self._strict_xlsx_resource_fields())
            missing_time[0]["time"]["parsed"].pop("user_seconds")
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "user_seconds is missing or non-finite",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_resource_legs(missing_time)

            null_heaptrack_total = self._xlsx_xml_borrowed_legs(directory)
            for leg in null_heaptrack_total:
                leg.update(self._strict_xlsx_resource_fields())
            null_heaptrack_total[1]["heaptrack"]["print"]["parsed"][
                "temporary_allocations"
            ] = None
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "temporary_allocations is missing or non-finite",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_resource_legs(
                    null_heaptrack_total
                )

            unretained_histogram = self._xlsx_xml_borrowed_legs(directory)
            for leg in unretained_histogram:
                leg.update(self._strict_xlsx_resource_fields())
            unretained_histogram[2]["heaptrack"]["print"]["parsed"][
                "histogram_artifact"
            ]["retained"] = False
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "histogram_artifact must be present and retained",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_resource_legs(
                    unretained_histogram
                )

    def test_xlsx_xml_borrowed_acceptance_enforces_paired_five_percent_boundary(self):
        required = perf_resource_profile.XLSX_XML_BORROWED_REQUIRED_RESOURCE_METRICS
        with tempfile.TemporaryDirectory() as directory:
            at_boundary = self._xlsx_xml_borrowed_legs(directory)
            for leg in at_boundary:
                leg.update(self._strict_xlsx_resource_fields(100))
            for index in (1, 2):
                for metric in required:
                    self._set_strict_xlsx_metric(at_boundary[index], metric, 105)
            at_boundary[0]["resource_metrics"] = {
                "harness.xlsx_first_cell.elapsed_ns.p50": 10**12
            }
            accepted = perf_resource_profile.validate_xlsx_xml_borrowed_acceptance(
                at_boundary
            )
            self.assertEqual(accepted["status"], "accepted")
            self.assertFalse(accepted["instrumented_elapsed_included"])
            self.assertEqual(
                accepted["metrics"][required[0]]["A1_to_B1"]["delta_percent"],
                5.0,
            )

            over_boundary = self._xlsx_xml_borrowed_legs(directory)
            for leg in over_boundary:
                leg.update(self._strict_xlsx_resource_fields(100))
            self._set_strict_xlsx_metric(over_boundary[1], required[0], 105.1)
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "exceeds 5.000000%",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_acceptance(over_boundary)

            missing = self._xlsx_xml_borrowed_legs(directory)
            for leg in missing:
                leg.update(self._strict_xlsx_resource_fields(100))
            self._set_strict_xlsx_metric(missing[0], required[0], None)
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "missing or non-finite",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_acceptance(missing)

            nonfinite = self._xlsx_xml_borrowed_legs(directory)
            for leg in nonfinite:
                leg.update(self._strict_xlsx_resource_fields(100))
            self._set_strict_xlsx_metric(nonfinite[2], required[1], float("nan"))
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "missing or non-finite",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_acceptance(nonfinite)

            mismatch = self._xlsx_xml_borrowed_legs(directory)
            for leg in mismatch:
                leg.update(self._strict_xlsx_resource_fields(100))
            mismatch[0]["resource_metrics"] = {required[0]: 101}
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "does not match parsed evidence",
            ):
                perf_resource_profile.validate_xlsx_xml_borrowed_acceptance(mismatch)

    def test_xlsx_xml_borrowed_leg_retains_time_heaptrack_and_harness_artifacts(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            binary = root / "control-xlsx-borrowed-binary"
            binary.write_bytes(b"control-xlsx-borrowed")
            binary.chmod(0o755)
            artifact_root = root / "artifacts"
            descriptor = perf_resource_profile.binary_identity(binary, label="control")

            def fake_run(command, *, stdout_path, stderr_path, timeout_seconds):
                del timeout_seconds
                command = list(command)
                if "--json" in command:
                    report_path = Path(command[command.index("--json") + 1])
                    report_path.write_text(
                        json.dumps(self._xlsx_xml_borrowed_report()),
                        encoding="utf-8",
                    )
                if "-o" in command and "--record-only" not in command:
                    time_path = Path(command[command.index("-o") + 1])
                    time_path.write_text(
                        """\
Maximum resident set size (kbytes): 123
User time (seconds): 1.25
System time (seconds): 0.50
Elapsed (wall clock) time (h:mm:ss or m:ss): 0:02.75
Voluntary context switches: 12
Involuntary context switches: 3
Major (requiring I/O) page faults: 1
Minor (reclaiming a frame) page faults: 34
""",
                        encoding="utf-8",
                    )
                if "--record-only" in command:
                    prefix = Path(command[command.index("-o") + 1])
                    prefix.with_suffix(".zst").write_bytes(b"capture")
                if "-H" in command:
                    histogram = Path(command[command.index("-H") + 1])
                    histogram.write_text("4\t3\n", encoding="utf-8")
                    stdout_path.write_text(
                        "calls to allocation functions: 1\n"
                        "MOST TEMPORARY ALLOCATIONS\n"
                        "1 temporary allocations of 2 allocations in total (50%)\n"
                        "temporary memory allocations: 2\n"
                        "peak heap memory consumption: 1K\n"
                        "peak RSS (including heaptrack overhead): 2K\n",
                        encoding="utf-8",
                    )
                return {
                    "command": command,
                    "returncode": 0,
                    "timed_out": False,
                    "wall_ns": 1,
                    "stdout": {},
                    "stderr": {},
                    "stderr_excerpt": None,
                }

            with mock.patch.object(
                perf_resource_profile, "run_command", side_effect=fake_run
            ):
                leg = perf_resource_profile.profile_xlsx_xml_borrowed_abba_leg(
                    leg="A1",
                    variant="control",
                    binary=binary,
                    binary_descriptor=descriptor,
                    artifact_root=artifact_root,
                    warmup=1,
                    samples=3,
                    tools={
                        "time": {"available": True, "path": "/usr/bin/time"},
                        "heaptrack": {"available": True, "path": "heaptrack"},
                        "heaptrack_print": {
                            "available": True,
                            "path": "heaptrack_print",
                        },
                    },
                    timeout_seconds=1,
                )
        self.assertEqual(leg["latency_evidence"]["status"], "not_measured")
        self.assertIn(
            "harness.xlsx_source_first_cell.elapsed_ns.p50",
            leg["resource_metrics"],
        )
        self.assertTrue(leg["harness"]["report"]["retained"])
        self.assertEqual(len(leg["harness"]["report"]["sha256"]), 64)
        self.assertTrue(leg["harness"]["run"]["stdout"]["retained"])
        self.assertTrue(leg["time"]["parsed"]["artifact"]["retained"])
        self.assertEqual(len(leg["time"]["parsed"]["artifact"]["sha256"]), 64)
        self.assertEqual(leg["heaptrack"]["status"], "ok")
        self.assertEqual(
            leg["heaptrack"]["harness_identity"]["status"], "validated"
        )
        self.assertTrue(leg["heaptrack"]["harness"]["retained"])
        self.assertEqual(len(leg["heaptrack"]["harness"]["sha256"]), 64)
        self.assertTrue(leg["heaptrack"]["run"]["stdout"]["retained"])
        self.assertTrue(leg["heaptrack"]["capture"]["retained"])
        self.assertTrue(leg["heaptrack"]["print"]["artifact"]["retained"])
        self.assertTrue(leg["heaptrack"]["print"]["run"]["stdout"]["retained"])
        self.assertTrue(
            leg["heaptrack"]["print"]["parsed"]["histogram_artifact"]["retained"]
        )
        command = leg["harness"]["command"]
        self.assertIn(",".join(perf_resource_profile.XLSX_XML_BORROWED_CASES), command)
        self.assertIn("tiny", command)
        self.assertIn("medium", command)

    def test_xlsx_xml_borrowed_heaptrack_harness_identity_mismatch_is_rejected(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            binary = root / "control-xlsx-borrowed-binary"
            binary.write_bytes(b"control-xlsx-borrowed")
            binary.chmod(0o755)
            expected_report = self._xlsx_xml_borrowed_report()
            expected_identity = (
                perf_resource_profile._xlsx_xml_borrowed_harness_identity(
                    expected_report, "timed.harness_report"
                )
            )
            heaptrack_workdir = root / "artifacts" / "a1"
            heaptrack_workdir.mkdir(parents=True)

            def fake_run(command, *, stdout_path, stderr_path, timeout_seconds):
                del timeout_seconds
                command = list(command)
                if "--json" in command:
                    report_path = Path(command[command.index("--json") + 1])
                    report = self._xlsx_xml_borrowed_report()
                    if "--record-only" in command:
                        report["results"][1]["source"]["read_bytes"][0] = 999
                    report_path.write_text(json.dumps(report), encoding="utf-8")
                if "--record-only" in command:
                    prefix = Path(command[command.index("-o") + 1])
                    prefix.with_suffix(".zst").write_bytes(b"capture")
                if "-H" in command:
                    histogram = Path(command[command.index("-H") + 1])
                    histogram.write_text("4\t3\n", encoding="utf-8")
                    stdout_path.write_text(
                        "calls to allocation functions: 1\n"
                        "MOST TEMPORARY ALLOCATIONS\n"
                        "1 temporary allocations of 2 allocations in total (50%)\n"
                        "temporary memory allocations: 2\n"
                        "peak heap memory consumption: 1K\n"
                        "peak RSS (including heaptrack overhead): 2K\n",
                        encoding="utf-8",
                    )
                return {
                    "command": command,
                    "returncode": 0,
                    "timed_out": False,
                    "wall_ns": 1,
                    "stdout": {},
                    "stderr": {},
                    "stderr_excerpt": None,
                }

            with mock.patch.object(
                perf_resource_profile, "run_command", side_effect=fake_run
            ):
                with self.assertRaisesRegex(
                    perf_resource_profile.ResourceProfileInputError,
                    "heaptrack harness identity does not match timed leg identity",
                ):
                    perf_resource_profile._profile_abba_heaptrack(
                        binary,
                        perf_resource_profile.XLSX_XML_BORROWED_ARGS,
                        heaptrack_workdir,
                        {
                            "heaptrack": {"available": True, "path": "heaptrack"},
                            "heaptrack_print": {
                                "available": True,
                                "path": "heaptrack_print",
                            },
                        },
                        warmup=1,
                        samples=3,
                        timeout_seconds=1,
                        expected_harness_identity=expected_identity,
                        harness_identity_validator=(
                            lambda value, location: perf_resource_profile._xlsx_xml_borrowed_harness_identity(
                                value, location
                            )
                        ),
                    )

    def test_xlsx_xml_borrowed_source_phase_timing_is_not_process_identity(self):
        timed = self._xlsx_xml_borrowed_report()
        heaptrack = self._xlsx_xml_borrowed_report()
        source = heaptrack["results"][3]["source"]["xlsx_cell_values"]
        for offset, key in enumerate(
            perf_resource_profile.XLSX_XML_BORROWED_SOURCE_TIMING_FIELDS, start=1
        ):
            source[key] = [offset * 1_000] * 3

        timed_identity = perf_resource_profile._xlsx_xml_borrowed_harness_identity(
            timed, "timed"
        )
        heaptrack_identity = perf_resource_profile._xlsx_xml_borrowed_harness_identity(
            heaptrack, "heaptrack"
        )

        self.assertEqual(timed_identity, heaptrack_identity)
        result_identity = timed_identity["result_identities"][3]
        self.assertEqual(len(result_identity["excluded_source_observations"]), 5)
        self.assertNotIn("commit_ns", result_identity["source"]["xlsx_cell_values"])
        self.assertEqual(
            result_identity["source"]["xlsx_cell_values"]["source_read_bytes"],
            [64, 64, 64],
        )

    def test_xlsx_xml_borrowed_unknown_source_timing_is_rejected(self):
        report = self._xlsx_xml_borrowed_report()
        report["results"][3]["source"]["xlsx_cell_values"]["mystery_ns"] = [1, 1, 1]

        with self.assertRaisesRegex(
            perf_resource_profile.ResourceProfileInputError,
            "unrecognized source timing observation",
        ):
            perf_resource_profile._xlsx_xml_borrowed_harness_identity(report, "report")

    def test_xlsx_xml_borrowed_abba_orchestration_publishes_identity_and_scope(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legs = self._xlsx_xml_borrowed_legs(root)
            output = root / "xlsx-xml-borrowed.json"
            artifacts = root / "xlsx-xml-borrowed-artifacts"

            def fake_leg(**kwargs):
                leg = next(item for item in legs if item["leg"] == kwargs["leg"])
                return {
                    **leg,
                    **self._strict_xlsx_resource_fields(
                        100, artifact_root=artifacts / kwargs["leg"].lower()
                    ),
                    "harness": {"logical_measurements": []},
                    "latency_evidence": dict(perf_resource_profile.LATENCY_SEPARATION),
                    "artifact_directory": str(artifacts / kwargs["leg"].lower()),
                }

            available = {
                "available": True,
                "path": "/fake/resource-tool",
                "version": "fake-resource-tool",
                "binary_sha256": "f" * 64,
                "returncode": 0,
            }
            arguments = argparse.Namespace(
                control_binary=root / "control-xlsx-borrowed-binary",
                candidate_binary=root / "candidate-xlsx-borrowed-binary",
                output=output,
                artifact_dir=artifacts,
                warmup=1,
                samples=3,
                timeout=1,
            )
            with mock.patch.object(
                perf_resource_profile,
                "profile_xlsx_xml_borrowed_abba_leg",
                side_effect=fake_leg,
            ), mock.patch.object(
                perf_resource_profile, "probe_tool", return_value=available
            ):
                self.assertEqual(
                    perf_resource_profile.run_xlsx_xml_borrowed_abba(arguments), 0
                )
            published = json.loads(output.read_text(encoding="utf-8"))
            artifact_files_exist = [
                (
                    Path(leg["artifact_directory"]) / "strict-evidence.bin"
                ).is_file()
                for leg in published["legs"]
            ]
        self.assertEqual(
            published["scope"]["workload"], perf_resource_profile.XLSX_XML_BORROWED_ID
        )
        self.assertEqual(
            published["scope"]["cases"],
            list(perf_resource_profile.XLSX_XML_BORROWED_CASES),
        )
        self.assertEqual(published["scope"]["xlsx_shape"], "tiny")
        self.assertEqual(published["scope"]["xlsx_cell_crud_shape"], "medium")
        self.assertEqual(published["tool"]["version"], "0.1.5")
        self.assertEqual(published["latency_evidence"]["status"], "not_measured")
        self.assertEqual(len(published["corpus_identities"]), 4)
        self.assertEqual(len(published["result_identities"]), 4)
        self.assertEqual(published["tools"]["heaptrack"]["available"], True)
        self.assertEqual(published["resource_evidence"]["status"], "validated")
        self.assertEqual(published["predeclared_acceptance"]["status"], "accepted")
        expected_artifact_sha256 = perf_resource_profile.sha256_bytes(
            b"strict-xlsx-evidence"
        )
        for leg in published["legs"]:
            self.assertNotIn("harness_report", leg)
            self.assertIn("harness_identity", leg)
            self.assertIn("result_identities", leg["harness_identity"])
            self.assertEqual(
                leg["time"]["parsed"]["artifact"]["sha256"],
                expected_artifact_sha256,
            )
            self.assertEqual(
                leg["heaptrack"]["print"]["parsed"]["histogram_artifact"]["sha256"],
                expected_artifact_sha256,
            )
            self.assertTrue(all(artifact_files_exist))

    def test_xlsx_xml_borrowed_requires_each_external_resource_tool(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            self._xlsx_xml_borrowed_legs(root)
            arguments = argparse.Namespace(
                control_binary=root / "control-xlsx-borrowed-binary",
                candidate_binary=root / "candidate-xlsx-borrowed-binary",
                output=root / "strict.json",
                artifact_dir=root / "strict-artifacts",
                warmup=1,
                samples=3,
                timeout=1,
            )

            def fake_probe(path, args):
                del args
                name = Path(path).name
                if name == missing:
                    return {
                        "available": False,
                        "path": None,
                        "version": None,
                        "binary_sha256": None,
                        "returncode": None,
                    }
                return {
                    "available": True,
                    "path": f"/fake/{name}",
                    "version": "fake",
                    "binary_sha256": "f" * 64,
                    "returncode": 0,
                }

            for missing in ("time", "heaptrack", "heaptrack_print"):
                with self.subTest(missing=missing), mock.patch.object(
                    perf_resource_profile, "probe_tool", side_effect=fake_probe
                ):
                    with self.assertRaisesRegex(
                        perf_resource_profile.ResourceProfileInputError,
                        rf"requires {missing}",
                    ):
                        perf_resource_profile.run_xlsx_xml_borrowed_abba(arguments)

    def test_docx_optional_flag_requires_explicit_docx_abba_dispatch(self):
        generic = perf_resource_profile.build_parser().parse_args(
            ["run", "--include-one-paragraph-text"]
        )
        with self.assertRaisesRegex(
            perf_resource_profile.ResourceProfileInputError,
            "explicit control and candidate binaries",
        ):
            generic.function(generic)

        xlsx = perf_resource_profile.build_parser().parse_args(
            [
                "run",
                "--workload",
                perf_resource_profile.XLSX_MANAGED_BATCH_ID,
                "--control-binary",
                "/tmp/control",
                "--candidate-binary",
                "/tmp/candidate",
                "--include-one-paragraph-text",
            ]
        )
        with self.assertRaisesRegex(
            perf_resource_profile.ResourceProfileInputError,
            "requires the DOCX semantic ABBA workload",
        ):
            xlsx.function(xlsx)

    def test_docx_validation_requires_matching_rows_and_full_corpus_identity(self):
        with tempfile.TemporaryDirectory() as directory:
            legs = self._docx_legs(directory)
            expected = {
                "samples_per_case": 3,
                "warmup_iterations_per_case": 1,
                "cases": list(perf_resource_profile.DOCX_SEMANTIC_CASES),
                "semantic_shapes": ["large"],
            }
            validated = perf_resource_profile.validate_docx_abba_inputs(
                legs, expected_configuration=expected
            )
            self.assertEqual(validated["workload"], perf_resource_profile.DOCX_SEMANTIC_ID)
            self.assertEqual(
                [item["case"] for item in validated["corpus_identities"]],
                list(perf_resource_profile.DOCX_SEMANTIC_CASES),
            )
            self.assertEqual(validated["corpus_identities"][0]["corpus"]["shape"], "large")
            self.assertEqual(
                len(validated["corpus_identities"][0]["identity_sha256"]), 64
            )

            reordered = self._docx_legs(directory)
            reordered[0]["harness_report"]["results"].reverse()
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "cases must be"
            ):
                perf_resource_profile.validate_docx_abba_inputs(reordered)

            within_report_mismatch = self._docx_legs(directory)
            within_report_mismatch[0]["harness_report"]["results"][1]["corpus"][
                "archive_sha256"
            ] = "d" * 64
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "corpus identities do not match",
            ):
                perf_resource_profile.validate_docx_abba_inputs(within_report_mismatch)

            missing_hash = self._docx_legs(directory)
            del missing_hash[1]["harness_report"]["results"][0]["corpus"][
                "target_payload_sha256"
            ]
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "target_payload_sha256"
            ):
                perf_resource_profile.validate_docx_abba_inputs(missing_hash)

            mismatched = self._docx_legs(directory)
            for result in mismatched[2]["harness_report"]["results"]:
                result["corpus"]["archive_sha256"] = "c" * 64
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError, "DOCX semantic corpora do not match"
            ):
                perf_resource_profile.validate_docx_abba_inputs(mismatched)

    def test_docx_optional_paragraph_text_case_keeps_exact_case_and_corpus_identity(self):
        cases = perf_resource_profile.DOCX_SEMANTIC_CASES_WITH_ONE_PARAGRAPH_TEXT
        with tempfile.TemporaryDirectory() as directory:
            legs = self._docx_legs(directory, cases=cases)
            validated = perf_resource_profile.validate_docx_abba_inputs(
                legs,
                expected_configuration={
                    "samples_per_case": 3,
                    "warmup_iterations_per_case": 1,
                    "cases": list(cases),
                    "semantic_shapes": ["large"],
                },
                cases=cases,
            )
        self.assertEqual(
            [item["case"] for item in validated["corpus_identities"]],
            list(cases),
        )
        self.assertEqual(len(validated["corpus_identities"]), 3)
        self.assertTrue(
            all(len(item["identity_sha256"]) == 64 for item in validated["corpus_identities"])
        )

    def test_docx_optional_case_is_rejected_if_not_selected_for_validation(self):
        cases = perf_resource_profile.DOCX_SEMANTIC_CASES_WITH_ONE_PARAGRAPH_TEXT
        with tempfile.TemporaryDirectory() as directory:
            legs = self._docx_legs(directory, cases=cases)
            with self.assertRaisesRegex(
                perf_resource_profile.ResourceProfileInputError,
                "selected DOCX semantic rows",
            ):
                perf_resource_profile.validate_docx_abba_inputs(legs)

    def test_docx_validation_fails_closed_on_harness_and_environment_identity(self):
        cases = (
            ("tool name", lambda report: report["tool"].update(name="other-harness"), "tool.name"),
            ("tool version", lambda report: report["tool"].update(version="9.9.9"), "tool.version"),
            ("schema", lambda report: report.update(schema_version=999), "schema_version"),
            ("short revision", lambda report: report["environment"].update(git_revision="a" * 39), "git_revision"),
            ("uppercase revision", lambda report: report["environment"].update(git_revision="A" * 40), "git_revision"),
            ("missing compiler", lambda report: report["environment"].pop("rustc_version"), "rustc_version"),
            ("changed CPU", lambda report: report["environment"].update(cpu_model="other-cpu"), "stable environment"),
        )
        for name, mutate, expected in cases:
            with self.subTest(name=name), tempfile.TemporaryDirectory() as directory:
                legs = self._docx_legs(directory)
                mutate(legs[1]["harness_report"])
                with self.assertRaisesRegex(
                    perf_resource_profile.ResourceProfileInputError, expected
                ):
                    perf_resource_profile.validate_docx_abba_inputs(legs)

    def test_docx_validation_requires_complete_requested_finite_elapsed_statistics(self):
        cases = (
            ("missing elapsed", lambda elapsed: None, "elapsed_ns is required"),
            ("wrong unit", lambda elapsed: elapsed.update(unit="ms"), "unit must be 'ns'"),
            ("short samples", lambda elapsed: elapsed.update(samples=[1]), "sample_count"),
            ("nonfinite mean", lambda elapsed: elapsed.update(mean=float("inf")), "mean"),
            (
                "nonfinite confidence",
                lambda elapsed: elapsed["confidence_interval_95"].update(upper=float("nan")),
                "confidence_interval_95.upper",
            ),
        )
        for name, mutate, expected in cases:
            with self.subTest(name=name), tempfile.TemporaryDirectory() as directory:
                legs = self._docx_legs(directory)
                elapsed = legs[0]["harness_report"]["results"][0]["elapsed_ns"]
                if name == "missing elapsed":
                    legs[0]["harness_report"]["results"][0].pop("elapsed_ns")
                else:
                    mutate(elapsed)
                with self.assertRaisesRegex(
                    perf_resource_profile.ResourceProfileInputError, expected
                ):
                    perf_resource_profile.validate_docx_abba_inputs(legs)

    def test_docx_resource_metrics_are_case_specific_and_not_latency_evidence(self):
        report = self._docx_report()
        metrics = perf_resource_profile.instrumented_harness_metrics(
            report, perf_resource_profile.DOCX_SEMANTIC_CASES
        )
        self.assertEqual(metrics["harness.docx_semantic_open.elapsed_ns.p50"], 100)
        self.assertEqual(metrics["harness.docx_semantic_full_text.elapsed_ns.p50"], 110)
        self.assertEqual(
            perf_resource_profile.LATENCY_SEPARATION["status"], "not_measured"
        )
        with tempfile.TemporaryDirectory() as directory:
            legs = self._docx_legs(directory)
            for leg in legs:
                leg["resource_metrics"] = {
                    "harness.docx_semantic_open.elapsed_ns.p50": 100,
                    "harness.docx_semantic_full_text.elapsed_ns.p50": 200,
                }
            statistics = perf_resource_profile.abba_statistics(
                legs, metric_specs=perf_resource_profile.DOCX_RESOURCE_METRIC_SPECS
            )
        self.assertIn(
            "harness.docx_semantic_full_text.elapsed_ns.p50",
            statistics["metrics"],
        )
        self.assertIn("not latency evidence", statistics["metrics"][
            "harness.docx_semantic_open.elapsed_ns.p50"
        ]["description"])
        json.dumps(statistics, allow_nan=False)

    def test_docx_optional_case_metrics_remain_instrumented_not_latency(self):
        cases = perf_resource_profile.DOCX_SEMANTIC_CASES_WITH_ONE_PARAGRAPH_TEXT
        report = self._docx_report(cases=cases)
        metrics = perf_resource_profile.instrumented_harness_metrics(report, cases)
        self.assertIn(
            "harness.docx_semantic_one_paragraph_text.elapsed_ns.p50",
            metrics,
        )
        specs = perf_resource_profile.docx_resource_metric_specs(cases)
        metric_names = {name for name, _description in specs}
        self.assertIn(
            "harness.docx_semantic_one_paragraph_text.elapsed_ns.p95",
            metric_names,
        )
        self.assertEqual(perf_resource_profile.LATENCY_SEPARATION["status"], "not_measured")

    def test_docx_leg_uses_time_and_heaptrack_without_collapsing_latency(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            cases = perf_resource_profile.DOCX_SEMANTIC_CASES_WITH_ONE_PARAGRAPH_TEXT
            binary = root / "control-docx-binary"
            binary.write_bytes(b"control-docx")
            binary.chmod(0o755)
            artifact_root = root / "artifacts"
            descriptor = perf_resource_profile.binary_identity(binary, label="control")

            def fake_run(command, *, stdout_path, stderr_path, timeout_seconds):
                del timeout_seconds
                command = list(command)
                if "--json" in command:
                    report_path = Path(command[command.index("--json") + 1])
                    report_path.write_text(
                        json.dumps(self._docx_report(cases=cases)), encoding="utf-8"
                    )
                if "-o" in command and "--record-only" not in command:
                    time_path = Path(command[command.index("-o") + 1])
                    time_path.write_text(
                        "Maximum resident set size (kbytes): 123\n",
                        encoding="utf-8",
                    )
                if "--record-only" in command:
                    (root / "artifacts" / "a1").mkdir(parents=True, exist_ok=True)
                    (root / "artifacts" / "a1" / "heaptrack-profile.zst").write_bytes(
                        b"capture"
                    )
                if "-H" in command:
                    histogram = Path(command[command.index("-H") + 1])
                    histogram.write_text("4\t3\n", encoding="utf-8")
                    stdout_path.write_text(
                        "calls to allocation functions: 1\n"
                        "peak heap memory consumption: 1K\n",
                        encoding="utf-8",
                    )
                return {
                    "command": command,
                    "returncode": 0,
                    "timed_out": False,
                    "wall_ns": 1,
                    "stdout": {},
                    "stderr": {},
                    "stderr_excerpt": None,
                }

            with mock.patch.object(
                perf_resource_profile, "run_command", side_effect=fake_run
            ):
                leg = perf_resource_profile.profile_docx_semantic_abba_leg(
                    leg="A1",
                    variant="control",
                    binary=binary,
                    binary_descriptor=descriptor,
                    artifact_root=artifact_root,
                    warmup=1,
                    samples=3,
                    tools={
                        "time": {"available": True, "path": "/usr/bin/time"},
                        "heaptrack": {"available": True, "path": "heaptrack"},
                        "heaptrack_print": {
                            "available": True,
                            "path": "heaptrack_print",
                        },
                    },
                    timeout_seconds=1,
                    cases=cases,
                )
        self.assertEqual(leg["latency_evidence"]["status"], "not_measured")
        self.assertIn("harness.docx_semantic_open.elapsed_ns.p50", leg["resource_metrics"])
        self.assertIn(
            "harness.docx_semantic_one_paragraph_text.elapsed_ns.p50",
            leg["resource_metrics"],
        )
        self.assertTrue(
            any(
                "docx_semantic_one_paragraph_text" in item
                for item in leg["harness"]["command"]
            )
        )
        self.assertIn("/usr/bin/time", leg["time"]["command"])
        self.assertEqual(leg["heaptrack"]["status"], "ok")

    def test_docx_abba_orchestration_publishes_identity_before_resource_statistics(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legs = self._docx_legs(root)
            output = root / "docx-resource.json"
            artifacts = root / "docx-artifacts"

            def fake_leg(**kwargs):
                leg = next(item for item in legs if item["leg"] == kwargs["leg"])
                return {
                    **leg,
                    "harness": {"logical_measurements": []},
                    "time": {"status": "unsupported"},
                    "heaptrack": {"status": "unsupported"},
                    "latency_evidence": dict(perf_resource_profile.LATENCY_SEPARATION),
                    "resource_metrics": {},
                    "artifact_directory": str(artifacts / kwargs["leg"].lower()),
                }

            unavailable = {
                "available": False,
                "path": None,
                "version": None,
                "binary_sha256": None,
                "returncode": None,
            }
            arguments = argparse.Namespace(
                control_binary=root / "control-docx-binary",
                candidate_binary=root / "candidate-docx-binary",
                output=output,
                artifact_dir=artifacts,
                warmup=1,
                samples=3,
                timeout=1,
            )
            with mock.patch.object(
                perf_resource_profile, "profile_docx_semantic_abba_leg", side_effect=fake_leg
            ), mock.patch.object(
                perf_resource_profile, "probe_tool", return_value=unavailable
            ):
                self.assertEqual(
                    perf_resource_profile.run_docx_semantic_abba(arguments), 0
                )
            published = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(published["scope"]["workload"], perf_resource_profile.DOCX_SEMANTIC_ID)
        self.assertEqual(
            published["corpus_identities"][0]["corpus"]["archive_sha256"], "a" * 64
        )
        self.assertEqual(published["latency_evidence"]["status"], "not_measured")
        self.assertEqual(published["tools"]["heaptrack"]["available"], False)
        self.assertEqual(published["statistics"]["not_measured"]["physical_cold_io"].split(":", 1)[0], "not measured")
        self.assertEqual(
            published["canonical_harness_identity"]["tool"]["name"],
            perf_resource_profile.HARNESS_TOOL_NAME,
        )
        self.assertEqual(
            published["canonical_harness_identity"]["environment"]["cpu_model"],
            "test-cpu",
        )
        self.assertEqual(
            published["canonical_harness_identity"]["leg_revisions"]["A1"],
            "1" * 40,
        )
        for leg in published["legs"]:
            self.assertNotIn("harness_report", leg)
            self.assertIn("harness_identity", leg)
            self.assertEqual(leg["harness_identity"]["environment"]["allocator"], "Rust system allocator")

    def test_docx_optional_abba_publishes_case_and_keeps_latency_unmeasured(self):
        cases = perf_resource_profile.DOCX_SEMANTIC_CASES_WITH_ONE_PARAGRAPH_TEXT
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            legs = self._docx_legs(root, cases=cases)
            output = root / "docx-resource-optional.json"
            artifacts = root / "docx-artifacts-optional"

            def fake_leg(**kwargs):
                leg = next(item for item in legs if item["leg"] == kwargs["leg"])
                return {
                    **leg,
                    "harness": {"logical_measurements": []},
                    "time": {"status": "unsupported"},
                    "heaptrack": {"status": "unsupported"},
                    "latency_evidence": dict(perf_resource_profile.LATENCY_SEPARATION),
                    "resource_metrics": {},
                    "artifact_directory": str(artifacts / kwargs["leg"].lower()),
                }

            unavailable = {
                "available": False,
                "path": None,
                "version": None,
                "binary_sha256": None,
                "returncode": None,
            }
            arguments = argparse.Namespace(
                control_binary=root / "control-docx-binary",
                candidate_binary=root / "candidate-docx-binary",
                output=output,
                artifact_dir=artifacts,
                warmup=1,
                samples=3,
                timeout=1,
                include_one_paragraph_text=True,
            )
            with mock.patch.object(
                perf_resource_profile,
                "profile_docx_semantic_abba_leg",
                side_effect=fake_leg,
            ), mock.patch.object(
                perf_resource_profile, "probe_tool", return_value=unavailable
            ):
                self.assertEqual(
                    perf_resource_profile.run_docx_semantic_abba(arguments), 0
                )
            published = json.loads(output.read_text(encoding="utf-8"))
        self.assertEqual(published["scope"]["cases"], list(cases))
        self.assertTrue(published["configuration"]["include_one_paragraph_text"])
        self.assertEqual(len(published["corpus_identities"]), 3)
        self.assertEqual(published["latency_evidence"]["status"], "not_measured")


if __name__ == "__main__":
    unittest.main()
