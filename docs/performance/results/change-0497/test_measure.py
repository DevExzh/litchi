#!/usr/bin/env python3
"""Focused custody, publication, matrix, timeout, and statistics tests."""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import stat
import subprocess
import tempfile
import unittest
from unittest import mock

HERE = Path(__file__).resolve().parent
SPEC = importlib.util.spec_from_file_location("change0497_measure_tested", HERE / "measure.py")
assert SPEC is not None and SPEC.loader is not None
measure = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(measure)


def _write(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, sort_keys=True) + "\n", encoding="utf-8")


def _meta(path: Path) -> dict[str, object]:
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


class InventoryTests(unittest.TestCase):
    def test_formal_inventory_counts_blocks_and_reverses_pairs(self) -> None:
        runs = measure.formal_inventory()
        self.assertEqual(len(runs), 288)
        self.assertEqual(sum(item["phase"] == "before" for item in runs), 72)
        self.assertEqual(sum(item["publication"] == "hashing_sink" for item in runs), 144)
        self.assertEqual(sum(item["publication"] == "counting_sink" for item in runs), 72)
        self.assertEqual(sum(item["publication"] == "atomic_path" for item in runs), 72)
        self.assertEqual(runs[0]["role"], "normal")
        self.assertEqual(runs[0]["phase"], "before")
        self.assertEqual(runs[1]["phase"], "after")
        self.assertEqual(runs[1]["publication"], "hashing_sink")
        self.assertEqual(runs[2]["role"], "allocator")
        self.assertEqual(runs[2]["phase"], "before")
        self.assertEqual(runs[3]["phase"], "after")
        self.assertEqual(runs[71]["arm"], measure.ARMS[-1]["id"])
        self.assertEqual(runs[71]["role"], "allocator")
        self.assertEqual(runs[72]["repeat"], 2)
        self.assertEqual(runs[72]["phase"], "after")
        self.assertEqual(runs[73]["phase"], "before")
        self.assertEqual(runs[143]["role"], "normal")
        self.assertEqual(runs[144]["publication"], "counting_sink")
        self.assertEqual(runs[145]["publication"], "atomic_path")
        self.assertEqual(runs[215]["publication"], "atomic_path")
        self.assertEqual(runs[216]["publication"], "atomic_path")
        self.assertEqual(runs[217]["publication"], "counting_sink")

    def test_pilot_is_separate_three_sample_lane(self) -> None:
        pilots = measure.pilot_inventory()
        self.assertEqual(len(pilots), 72)
        self.assertTrue(all(item["pilot"] and item["samples"] == 3
                            and item["warmups"] == 1 and item["role"] == "normal"
                            for item in pilots))
        self.assertEqual(len({item["arm"] for item in pilots}), 18)

    def test_pilot_cli_has_fixed_timeout_argument(self) -> None:
        args = measure._parser().parse_args(["pilot", "--attempt", "pilot1"])
        self.assertEqual(args.timeout, measure.DEFAULT_TIMEOUT_SECONDS)


class BuildCustodyTests(unittest.TestCase):
    def _builds(self, root: Path) -> Path:
        driver = root / "build.py"
        driver.write_text("driver\n", encoding="utf-8")
        source = root / "source.json"
        source_value = {
            name: {"path": f"/tmp/source/tools/perf-baseline/{name.rsplit('/', 1)[-1]}",
                   "bytes": 1, "sha256": "0" * 64}
            for name in measure.HARNESS_CUSTODY_FILES
        }
        for index, name in enumerate((
            "crates/litchi-docx/src/source_backed/tail_append_stream.rs",
            "crates/litchi-docx/tests/source_backed_tail_append_stream.rs",
            "tools/perf-baseline/src/docx_replayable_tail_append.rs",
            "tools/perf-baseline/src/docx_replayable_tail_append/route_smoke_tests.rs",
            "tools/perf-baseline/src/docx_replayable_tail_append/publication_route_tests.rs",
        ), start=1):
            source_value[name] = {"path": f"/tmp/source/{name}", "bytes": index,
                                  "sha256": f"{index:064x}"}
        _write(source, source_value)
        source_meta = _meta(source)
        builds: dict[str, object] = {}
        for phase in measure.PHASES:
            for role in measure.ROLES:
                binary = root / f"{phase}-{role}"
                binary.write_bytes(f"{phase}/{role}".encode())
                binary.chmod(binary.stat().st_mode | stat.S_IXUSR)
                gate = root / f"{phase}-{role}.gate.json"
                gate_stdout = root / f"{phase}-{role}.stdout"
                gate_stderr = root / f"{phase}-{role}.stderr"
                gate_stdout.write_bytes(b"")
                gate_stderr.write_bytes(b"")
                environment = {
                    "RUSTUP_TOOLCHAIN": "1.98.1", "CARGO_BUILD_JOBS": "4",
                    "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_RELEASE_DEBUG": "1",
                    "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
                    "CARGO_TARGET_DIR": str(root / "target"), "TMPDIR": str(root / "tmp"),
                    "DEBUGINFOD_URLS": "", "LC_ALL": "C", "RUSTDOCFLAGS": "-Dwarnings",
                }
                argv = ["cargo", "build", "--release", "--locked", "--offline",
                        "--manifest-path", "tools/perf-baseline/Cargo.toml"]
                if role == "allocator":
                    argv += ["--features", "allocator-metrics"]
                argv += ["--bin", "docx_replayable_tail_append"]
                gate_value = {
                    "argv": argv, "cwd": "/tmp/source", "environment": environment,
                    "source_manifest": source_meta, "started_ns": 1, "finished_ns": 2,
                    "driver": _meta(driver), "pid": 10, "exit_code": 0,
                    "source_unchanged": True, "stdout": _meta(gate_stdout),
                    "stderr": _meta(gate_stderr),
                }
                _write(gate, gate_value)
                builds[f"{phase}/{role}"] = {
                    "binary": measure._meta(binary, executable=True),
                    "git_revision": measure.EXPECTED_REVISIONS[phase],
                    "source_manifest": source_meta, "gate": _meta(gate),
                }
        path = root / "builds.json"
        _write(path, {"schema": measure.BUILDS_SCHEMA, "builds": builds})
        return path

    def test_gate_exit_and_retained_binary_are_authenticated(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = self._builds(root)
            candidate = {name: {"bytes": item["bytes"], "sha256": item["sha256"]}
                         for name, item in json.loads((root / "source.json").read_text()).items()}
            with mock.patch.object(measure, "_candidate_source_binding",
                                   return_value=({"manifest": {}, "patch": {}}, candidate)):
                builds = measure.load_builds(path)
            self.assertEqual(set(builds), {f"{phase}/{role}" for phase in measure.PHASES for role in measure.ROLES})
            self.assertEqual(builds["before/normal"]["binary"]["sha256"], _meta(root / "before-normal")["sha256"])
            gate = root / "after-normal.gate.json"
            value = json.loads(gate.read_text(encoding="utf-8"))
            value["exit_code"] = 1
            _write(gate.with_name("bad-gate.json"), value)
            data = json.loads(path.read_text(encoding="utf-8"))
            data["builds"]["after/normal"]["gate"] = _meta(gate.with_name("bad-gate.json"))
            bad = root / "bad-builds.json"
            _write(bad, data)
            with mock.patch.object(measure, "_candidate_source_binding",
                                   return_value=({"manifest": {}, "patch": {}}, candidate)):
                with self.assertRaises(measure.MeasureError):
                    measure.load_builds(bad)

    def test_source_manifest_change_outside_reviewed_overlay_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = self._builds(root)
            data = json.loads(path.read_text(encoding="utf-8"))
            source_path = root / "source.json"
            source = json.loads(source_path.read_text(encoding="utf-8"))
            after_source_path = root / "after-source.json"
            after_source = dict(source)
            after_source["unreviewed.rs"] = {"path": "/tmp/source/unreviewed.rs",
                                              "bytes": 1, "sha256": "f" * 64}
            _write(after_source_path, after_source)
            # Rebinding all four source-manifest metadata records simulates a
            # build that quietly carried an unrelated compilation input.
            source_meta = _meta(after_source_path)
            for key, record in data["builds"].items():
                if not key.startswith("after/"):
                    continue
                record["source_manifest"] = source_meta
                gate_path = Path(record["gate"]["path"])
                gate = json.loads(gate_path.read_text(encoding="utf-8"))
                gate["source_manifest"] = source_meta
                _write(gate_path.with_name("rebound-" + gate_path.name), gate)
                record["gate"] = _meta(gate_path.with_name("rebound-" + gate_path.name))
            rebound = root / "rebound.json"
            _write(rebound, data)
            candidate = {name: {"bytes": item["bytes"], "sha256": item["sha256"]}
                         for name, item in source.items()
                         if name != "unreviewed.rs"}
            with mock.patch.object(measure, "_candidate_source_binding",
                                   return_value=({"manifest": {}, "patch": {}}, candidate)):
                with self.assertRaises(measure.MeasureError):
                    measure.load_builds(rebound)


class ReceiptTests(unittest.TestCase):
    def test_terminal_rejects_exit_failure_and_timeout(self) -> None:
        base = {"status": "pass", "exit_code": 0, "timed_out": False,
                "termination": None, "launch_error": None, "validation_error": None,
                "missing_artifacts": []}
        measure._require_success_terminal(base)
        for field, value in (("exit_code", 1), ("timed_out", True),
                             ("validation_error", "bad")):
            tampered = dict(base)
            tampered[field] = value
            with self.assertRaises(measure.MeasureError):
                measure._require_success_terminal(tampered)

    def test_timeout_uses_process_group_termination(self) -> None:
        process = subprocess.Popen(
            [os.environ.get("PYTHON", "python3"), "-c", "import time; time.sleep(30)"],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
        try:
            with self.assertRaises(subprocess.TimeoutExpired):
                process.communicate(timeout=0.02)
            self.assertIn(measure._kill_group(process), ("SIGTERM", "SIGKILL", None))
            self.assertIsNotNone(process.returncode)
        finally:
            if process.poll() is None:
                measure._kill_group(process)

    def test_publication_rejects_fake_atomic_sink_counts(self) -> None:
        sample = {"publication": {
            "schema": "docx-replayable-tail-append-publication-v1",
            "route": "atomic_path",
            "timing_scope": "source_admission_prepare_atomic_write_data_sync_rename_parent_directory_sync_publication_drop",
            "timed_candidate_artifact_bytes": 4,
            "timed_candidate_artifact_sha256": "0" * 64,
            "timed_candidate_matches_oracle": True,
            "verification_scope": "timed_production_artifact_proof_plus_post_timer_path_oracle",
            "atomic": {},
        }, "sink": {"write_calls": 1}}
        with self.assertRaises(measure.MeasureError):
            measure._validate_publication_sample(sample, "atomic_path", 4, "0" * 64,
                                                  "sample", Path("/tmp/private/tmp"))

    def test_publication_accepts_unavailable_atomic_sink_nulls(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            private_tmpdir = Path(directory) / "tmp"
            private_parent = private_tmpdir / "atomic-parent"
            destination = private_parent / "published.docx"
            sample = {"publication": {
                "schema": "docx-replayable-tail-append-publication-v1",
                "route": "atomic_path",
                "timing_scope": "source_admission_prepare_atomic_write_data_sync_rename_parent_directory_sync_publication_drop",
                "timed_candidate_artifact_bytes": 4,
                "timed_candidate_artifact_sha256": "0" * 64,
                "timed_candidate_matches_oracle": True,
                "verification_scope": "timed_production_artifact_proof_plus_post_timer_path_oracle",
                "atomic": {
                    "destination_path": str(destination),
                    "private_parent_path": str(private_parent),
                    "before": {"exists": False, "regular_file": False, "bytes": None},
                    "after": {"exists": True, "regular_file": True, "bytes": 4},
                    "post_timer_archive_bytes": 4,
                    "post_timer_archive_sha256": "0" * 64,
                    "output_bytes_exact": True,
                    "output_sha256_exact": True,
                    "inverse_oracle_scope": (
                        "untimed_fixture_publication_inverse_exact; "
                        "timed_atomic_publication_inverse_not_reexecuted"),
                    "post_timer_oracle": {
                        "candidate_archive_bytes": 4,
                        "candidate_archive_sha256": "0" * 64,
                        "candidate_main_xml_bytes": 1,
                        "candidate_main_xml_sha256": "1" * 64,
                        "candidate_semantic": {
                            "paragraph_count": 1,
                            "order_sha256": "2" * 64,
                            "text_sha256": "3" * 64,
                            "text_bytes": 1,
                        },
                        "candidate_member_count": 1,
                        **{field: True for field in (
                            "candidate_xml_exact", "candidate_semantic_exact",
                            "untouched_member_metadata_exact", "untouched_raw_members_preserved",
                            "physical_order_exact", "opaque_member_exact", "source_unchanged",
                            "inverse_exact")},
                    },
                    "cleanup": {"destination_removed": True, "parent_removed": True},
                },
            }, "sink": {field: None for field in (
                "accepted_bytes", "write_calls", "largest_write", "histogram", "sha256")}}
            measure._validate_publication_sample(sample, "atomic_path", 4, "0" * 64,
                                                  "sample", private_tmpdir)


class AtomicConfigTests(unittest.TestCase):
    def test_atomic_config_binds_selected_provider_input_replay_and_compression(self) -> None:
        arm = measure.ARM_BY_ID["file_store-owned-s64-a64-short-c64"]
        route_spec = measure._route_spec(arm)
        replay_dir = Path("/tmp/0497-private/replay")
        expected = measure._expected_new_report_config(
            arm, route_spec, "atomic_path", measure.FORMAL_SAMPLES,
            measure.FORMAL_WARMUPS, replay_dir)
        measure._validate_new_report_config(
            expected, arm, route_spec, "atomic_path", measure.FORMAL_SAMPLES,
            measure.FORMAL_WARMUPS, replay_dir, "atomic.config")
        for field, value in (
            ("provider", "deterministic"),
            ("input_mode", "file"),
            ("replay_sync", "none"),
            ("compression", "store"),
            ("replay_dir", "/tmp/0497-private/other"),
        ):
            tampered = copy.deepcopy(expected)
            tampered[field] = value
            with self.assertRaises(measure.MeasureError):
                measure._validate_new_report_config(
                    tampered, arm, route_spec, "atomic_path", measure.FORMAL_SAMPLES,
                    measure.FORMAL_WARMUPS, replay_dir, "atomic.config")

    def test_failed_child_private_tree_is_preserved_with_hash_inventory(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory) / "goal"
            private = root / "runs" / "attempt" / "child"
            private.mkdir(parents=True)
            artifact = private / "committed-output.docx"
            artifact.write_bytes(b"retained failure output")
            with mock.patch.object(measure, "TEMP", root):
                cleanup = measure._cleanup_private(private, preserve_failure=True)
            self.assertEqual(cleanup["status"], "preserved_failure")
            self.assertEqual(cleanup["remaining"], [str(private)])
            self.assertEqual(cleanup["inventory"][0]["sha256"], _meta(artifact)["sha256"])
            self.assertTrue(artifact.exists())


class OrderingAndStatsTests(unittest.TestCase):
    def test_actual_receipt_order_is_not_replaced_by_ordinal_order(self) -> None:
        first, second = measure.formal_inventory()[:2]
        entries = [
            {"spec": dict(second, attempt="formal1"),
             "started_at": measure._timestamp("2026-01-01T00:00:00Z", "start"),
             "finished_at": measure._timestamp("2026-01-01T00:00:01Z", "finish")},
            {"spec": dict(first, attempt="formal1"),
             "started_at": measure._timestamp("2026-01-01T00:00:02Z", "start"),
             "finished_at": measure._timestamp("2026-01-01T00:00:03Z", "finish")},
        ]
        ordered = sorted(entries, key=lambda entry: entry["started_at"])
        with self.assertRaises(measure.MeasureError):
            for expected, entry in zip((first, second), ordered):
                if entry["spec"] != dict(expected, attempt="formal1"):
                    measure.fail("actual receipt chronology differs")

    def test_bootstrap_is_deterministic_and_uses_repeat_blocks(self) -> None:
        one = measure.paired_bootstrap([10, 20], repetitions=100, seed=497)
        two = measure.paired_bootstrap([10, 20], repetitions=100, seed=497)
        self.assertEqual(one, two)
        self.assertEqual(one["blocks"], 2)
        self.assertEqual(one["method"], "paired_repeat_block_bootstrap_median")

    def test_allocator_peak_increment_is_reported_and_normal_is_unavailable(self) -> None:
        rows = []
        for index in range(2):
            rows.append({"allocation": {
                "status": "measured", "scope": "operation_global_system_allocator",
                "allocation_calls": 5, "deallocation_calls": 5, "reallocation_calls": 0,
                "failed_allocation_calls": 0, "allocated_bytes": 10,
                "deallocated_bytes": 10, "live_bytes_before": 4,
                "live_bytes_after": 4, "peak_live_bytes_before": 4,
                "peak_live_bytes_after": 20, "region_peak_live_bytes": 12,
            }})
        entry = {"spec": {"role": "allocator"}, "directory": "entry"}
        metrics = measure._allocation_metrics(entry, rows)
        self.assertEqual(metrics["vectors"]["allocation_peak_increment_bytes"], [8, 8])
        normal = {"spec": {"role": "normal"}, "directory": "normal"}
        unavailable = measure._allocation_metrics(normal, [{"allocation": None}])
        self.assertEqual(unavailable["availability"], "unavailable")
        self.assertIsNone(unavailable["vectors"])

    def test_counting_projection_keeps_raw_publication_and_removes_only_extension(self) -> None:
        raw = {
            "config": {"publication": "counting_sink", "sink": "counting"},
            "cases": [{"samples": [{
                "publication": {"schema": "docx-replayable-tail-append-publication-v1",
                                 "route": "counting_sink", "timed_candidate_artifact_sha256": "0" * 64},
                "sink": {"accepted_bytes": 1, "write_calls": 1, "largest_write": 1,
                         "histogram": {}},
            }]}],
        }
        projection = copy.deepcopy(raw)
        projection["cases"][0]["samples"][0]["_publication_digest_for_projection"] = "0" * 64
        projected = measure._report_path_for_projection(projection, counting=True)
        self.assertNotIn("publication", projected["config"])
        self.assertNotIn("publication", projected["cases"][0]["samples"][0])
        self.assertEqual(raw["cases"][0]["samples"][0]["publication"]["route"], "counting_sink")


if __name__ == "__main__":
    unittest.main()
