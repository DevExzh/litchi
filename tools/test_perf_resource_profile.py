"""Unit tests for the standard-library resource-profile parser."""

from __future__ import annotations

import argparse
import json
import tempfile
import unittest
from unittest import mock
from pathlib import Path

from tools import perf_resource_profile


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
34 temporary allocations of 56 allocations in total (60.7%)
peak heap memory consumption: 2.50M
peak RSS (including heaptrack overhead): 4.00M
""",
                encoding="utf-8",
            )
            histogram.write_text("4\t3\n8\t2\n", encoding="utf-8")
            parsed = perf_resource_profile.parse_heaptrack_print(summary)
            allocated = perf_resource_profile.parse_heaptrack_histogram(histogram)
        self.assertEqual(parsed["allocation_calls"], 12)
        self.assertEqual(parsed["temporary_allocations"], 34)
        self.assertEqual(parsed["peak_heap_bytes"], int(2.5 * 1024 * 1024))
        self.assertEqual(parsed["peak_rss_bytes"], 4 * 1024 * 1024)
        self.assertEqual(allocated, 28)

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
