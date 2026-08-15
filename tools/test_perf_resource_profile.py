"""Unit tests for the standard-library resource-profile parser."""

from __future__ import annotations

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
    def _abba_report(revision="control-revision", *, corpus_shape="medium", samples=3):
        return {
            "schema_version": perf_resource_profile.SCHEMA_VERSION,
            "tool": {
                "name": "litchi-perf-baseline",
                "version": "0.1.0",
                "profile": "release",
            },
            "environment": {
                "git_revision": revision,
                "git_worktree_dirty": False,
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
                        "p50": 100,
                        "p95": 110,
                        "p99": 120,
                        "mean": 105,
                        "standard_deviation": 5,
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
            revision = "control-revision" if variant == "control" else "candidate-revision"
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


if __name__ == "__main__":
    unittest.main()
