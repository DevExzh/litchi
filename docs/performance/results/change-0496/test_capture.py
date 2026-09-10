#!/usr/bin/env python3
"""Focused custody and statistical tests for the 0496 phase driver."""

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
SPEC = importlib.util.spec_from_file_location("change0496_capture_tested", HERE / "capture.py")
assert SPEC is not None and SPEC.loader is not None
capture = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(capture)


def _write(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, sort_keys=True) + "\n", encoding="utf-8")


def _meta(path: Path) -> dict[str, object]:
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


class InventoryTests(unittest.TestCase):
    def test_formal_inventory_has_exact_matrix_and_reversed_repeat(self) -> None:
        runs = capture.formal_inventory()
        self.assertEqual(len(runs), 32)
        self.assertEqual(sum(item["phase"] == "before" for item in runs), 12)
        self.assertEqual(sum(item["phase"] == "after" and item["api"] == capture.UNMANAGED for item in runs), 12)
        self.assertEqual(sum(item["phase"] == "after" and item["api"] == capture.MANAGED for item in runs), 8)
        self.assertEqual([item["phase"] for item in runs[:6]], ["before"] * 6)
        self.assertEqual([item["phase"] for item in runs[6:16]], ["after"] * 10)
        self.assertEqual([item["phase"] for item in runs[16:26]], ["after"] * 10)
        self.assertEqual([item["phase"] for item in runs[26:]], ["before"] * 6)
        self.assertEqual(runs[0]["role"], "normal")
        self.assertEqual(runs[16]["role"], "allocator")
        self.assertEqual(runs[16]["arm"], "short")
        self.assertEqual(runs[26]["arm"], "short")
        self.assertFalse(any(item["phase"] == "before" and item["api"] == capture.MANAGED for item in runs))


class PhaseContractTests(unittest.TestCase):
    def _phase(self, latency: int = 100) -> dict[str, object]:
        values = {name: 10 for name in capture.PHASE_DURATION_FIELDS}
        values["phase_sum_ns"] = sum(values.values())
        values["lifecycle_residual_ns"] = latency - values["phase_sum_ns"]
        return {
            "schema": capture.PHASE_SCHEMA,
            "timing_scope": capture.PHASE_TIMING_SCOPE,
            **values,
            "residual_scope": capture.PHASE_RESIDUAL_SCOPE,
            "instrumentation_overhead_ns": None,
            "instrumentation_scope": capture.PHASE_INSTRUMENTATION_SCOPE,
            "allocation_scope": capture.PHASE_ALLOCATION_SCOPE,
        }

    def _raw(self) -> dict[str, object]:
        rows = [{"latency_ns": 100, "phase_diagnostics": self._phase()} for _ in range(capture.FORMAL_SAMPLES)]
        # The projection test does not need the full 0495 report; its exact
        # canonical field set is supplied by the patch below.
        return {"marker": "canonical", "rows": rows}

    def test_phase_projection_removes_only_explicit_row_field(self) -> None:
        raw = self._raw()
        with mock.patch.object(capture, "_canonical_fields", return_value=("marker", "rows")):
            projection, phases = capture._phase_projection(raw, "report")
        self.assertEqual(set(projection), {"marker", "rows"})
        self.assertEqual(len(phases), capture.FORMAL_SAMPLES)
        self.assertNotIn("phase_diagnostics", projection["rows"][0])
        self.assertIn("phase_diagnostics", raw["rows"][0])

    def test_phase_nested_extension_and_nonconserving_oracle_fail(self) -> None:
        raw = self._raw()
        raw["rows"][0]["phase_diagnostics"]["unexpected"] = 1
        with mock.patch.object(capture, "_canonical_fields", return_value=("marker", "rows")):
            with self.assertRaises(capture.PhaseDiagnosticError):
                capture._phase_projection(raw, "report")
        raw = self._raw()
        raw["rows"][0]["phase_diagnostics"]["phase_sum_ns"] += 1
        with mock.patch.object(capture, "_canonical_fields", return_value=("marker", "rows")):
            with self.assertRaises(capture.PhaseDiagnosticError):
                capture._phase_projection(raw, "report")

    def test_phase_sample_count_is_exact(self) -> None:
        raw = self._raw()
        raw["rows"] = raw["rows"][:-1]
        with mock.patch.object(capture, "_canonical_fields", return_value=("marker", "rows")):
            with self.assertRaises(capture.PhaseDiagnosticError):
                capture._phase_projection(raw, "report")


class BuildCustodyTests(unittest.TestCase):
    def _builds(self, root: Path) -> Path:
        source = root / "source.json"
        source_value = {
            "tools/perf-baseline/src/docx_managed_edit.rs": {"path": "/removed/before.rs", "bytes": 1, "sha256": "0" * 64},
            "tools/perf-baseline/Cargo.toml": {"path": "/removed/Cargo.toml", "bytes": 1, "sha256": "1" * 64},
            "tools/perf-baseline/src/lib.rs": {"path": "/removed/lib.rs", "bytes": 1, "sha256": "2" * 64},
            "tools/perf-baseline/src/main.rs": {"path": "/removed/main.rs", "bytes": 1, "sha256": "3" * 64},
            "tools/perf-baseline/src/bin/litchi-perf-baseline-alloc.rs": {"path": "/removed/alloc.rs", "bytes": 1, "sha256": "4" * 64},
        }
        _write(source, source_value)
        source_meta = _meta(source)
        builds: dict[str, object] = {}
        for phase in capture.PHASES:
            for role in capture.ROLES:
                binary = root / f"{phase}-{role}"
                binary.write_bytes(f"{phase}/{role}".encode())
                binary.chmod(binary.stat().st_mode | stat.S_IXUSR)
                gate = root / f"{phase}-{role}.gate.json"
                gate_stdout = root / f"{phase}-{role}.stdout"
                gate_stderr = root / f"{phase}-{role}.stderr"
                gate_stdout.write_bytes(b"")
                gate_stderr.write_bytes(b"")
                environment = {
                    "RUSTUP_TOOLCHAIN": "1.98.1", "CARGO_BUILD_JOBS": "4", "CARGO_INCREMENTAL": "0",
                    "CARGO_PROFILE_RELEASE_DEBUG": "1", "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
                    "CARGO_TARGET_DIR": str(root / "target"), "TMPDIR": str(root / "tmp"),
                    "DEBUGINFOD_URLS": "", "LC_ALL": "C", "RUSTDOCFLAGS": "-Dwarnings",
                }
                argv = ["cargo", "build", "--release", "--locked", "--offline", "--manifest-path",
                        "tools/perf-baseline/Cargo.toml"]
                if role == "normal":
                    argv += ["--bin", "litchi-perf-baseline"]
                else:
                    argv += ["--features", "allocator-metrics", "--bin", "litchi-perf-baseline-alloc"]
                gate_value = {
                    # The retained Cargo.toml path is /removed/Cargo.toml;
                    # its recorded checkout root is therefore /.
                    "argv": argv, "cwd": "/", "environment": environment,
                    "source_manifest": source_meta, "exit_code": 0, "source_unchanged": True,
                    "started_ns": 1, "finished_ns": 2, "pid": 10,
                    "stdout": _meta(gate_stdout), "stderr": _meta(gate_stderr),
                }
                _write(gate, gate_value)
                builds[f"{phase}/{role}"] = {
                    "binary": _meta(binary), "git_revision": capture.EXPECTED_REVISIONS[phase],
                    "source_manifest": source_meta, "gate": _meta(gate),
                }
        path = root / "builds.json"
        _write(path, {"schema": capture.BUILDS_SCHEMA, "builds": builds})
        return path

    def test_build_gate_and_binary_binding_are_authenticated(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            path = self._builds(root)
            builds = capture.load_builds(path)
            self.assertEqual(set(builds), {f"{phase}/{role}" for phase in capture.PHASES for role in capture.ROLES})
            self.assertEqual(builds["before/normal"]["binary"]["sha256"], _meta(root / "before-normal")["sha256"])

            # A gate receipt with a new authenticated file but a nonzero exit
            # must fail the original gate check; this cannot be repaired by
            # pointing the record at a different mutable cargo artifact.
            bad_gate = root / "after-normal.gate.json"
            bad_value = _read_json_for_test(bad_gate)
            bad_value["exit_code"] = 1
            _write(bad_gate, bad_value)
            data = json.loads(path.read_text(encoding="utf-8"))
            data["builds"]["after/normal"]["gate"] = _meta(bad_gate)
            bad_path = root / "bad-builds.json"
            _write(bad_path, data)
            with self.assertRaises(capture.PhaseDiagnosticError):
                capture.load_builds(bad_path)


def _read_json_for_test(path: Path) -> dict[str, object]:
    return json.loads(path.read_text(encoding="utf-8"))


class ReceiptAndOrderTests(unittest.TestCase):
    def test_wrong_terminal_exit_code_is_rejected(self) -> None:
        terminal = {
            "status": "pass", "exit_code": 1, "timed_out": False,
            "termination": None, "launch_error": None, "validation_error": None,
            "missing_artifacts": [],
        }
        with self.assertRaises(capture.PhaseDiagnosticError):
            capture._require_success_terminal(terminal)

    def test_timeout_kills_process_group(self) -> None:
        process = subprocess.Popen(
            [os.environ.get("PYTHON", "python3"), "-c", "import time; time.sleep(30)"],
            stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            start_new_session=True,
        )
        try:
            with self.assertRaises(subprocess.TimeoutExpired):
                process.communicate(timeout=0.02)
            termination = capture._kill_group(process)
            self.assertIn(termination, ("SIGTERM", "SIGKILL", None))
            self.assertIsNotNone(process.returncode)
        finally:
            if process.poll() is None:
                capture._kill_group(process)

    def test_timed_out_terminal_is_rejected(self) -> None:
        terminal = {
            "status": "pass", "exit_code": 0, "timed_out": True,
            "termination": "SIGTERM", "launch_error": None, "validation_error": None,
            "missing_artifacts": [],
        }
        with self.assertRaises(capture.PhaseDiagnosticError):
            capture._require_success_terminal(terminal)

    def test_actual_started_order_is_checked_after_label_inventory(self) -> None:
        # Exercise the ordering rule directly with two distinct formal runs;
        # sorting by protocol ordinal would incorrectly accept this reversal.
        first, second = capture.formal_inventory()[:2]
        entries = [
            {"spec": dict(second, attempt="formal1"), "started_at": capture._timestamp("2026-01-01T00:00:00Z", "start"),
             "finished_at": capture._timestamp("2026-01-01T00:00:01Z", "finish")},
            {"spec": dict(first, attempt="formal1"), "started_at": capture._timestamp("2026-01-01T00:00:02Z", "start"),
             "finished_at": capture._timestamp("2026-01-01T00:00:03Z", "finish")},
        ]
        with self.assertRaises(capture.PhaseDiagnosticError):
            capture._chronological_entries(entries, [first, second], "formal1")


class StatisticsTests(unittest.TestCase):
    def test_analysis_uses_retained_phase_rows_after_canonical_projection(self) -> None:
        phase = {name: 10 for name in capture.PHASE_DURATION_FIELDS}
        phase.update({"phase_sum_ns": 70, "lifecycle_residual_ns": 30})
        entry = {
            "report": {"rows": [{"latency_ns": 100} for _ in range(30)]},
            "phase_rows": [{
                "schema": capture.PHASE_SCHEMA,
                "timing_scope": capture.PHASE_TIMING_SCOPE,
                **phase,
                "residual_scope": capture.PHASE_RESIDUAL_SCOPE,
                "instrumentation_overhead_ns": None,
                "instrumentation_scope": capture.PHASE_INSTRUMENTATION_SCOPE,
                "allocation_scope": capture.PHASE_ALLOCATION_SCOPE,
            } for _ in range(30)],
            "resource": {"maximum_resident_set_size_(kbytes)": 1},
        }
        with mock.patch.object(capture, "_within_child_bootstrap", return_value={"median": 10}):
            metrics = capture._entry_metrics(entry)
        self.assertEqual(metrics["phases"]["open_ns"]["p50"], 10)
        self.assertEqual(metrics["latency_ns"]["p50"], 100)
        self.assertEqual(metrics["allocation"]["availability"], "unavailable")
        self.assertIsNone(metrics["allocation"]["vectors"])

    def test_allocator_vectors_percentiles_and_peak_increment_are_retained(self) -> None:
        phase = {name: 10 for name in capture.PHASE_DURATION_FIELDS}
        phase.update({"phase_sum_ns": 70, "lifecycle_residual_ns": 30})
        allocation_fields = {
            "allocation_calls": 11, "deallocation_calls": 9,
            "reallocation_calls": 2, "failed_allocation_calls": 0,
            "allocated_bytes": 100, "deallocated_bytes": 80,
            "live_bytes_before": 20, "live_bytes_after": 40,
            "peak_live_bytes_before": 20, "peak_live_bytes_after": 55,
            "region_peak_live_bytes": 50,
        }
        rows = []
        for _ in range(30):
            rows.append({
                "latency_ns": 100,
                "allocation": {"status": "measured", "scope": capture.SAMPLE_ALLOCATION_SCOPE,
                                **allocation_fields},
            })
        entry = {
            "spec": {"role": "allocator"},
            "report": {"rows": rows},
            "phase_rows": [{
                "schema": capture.PHASE_SCHEMA,
                "timing_scope": capture.PHASE_TIMING_SCOPE,
                **phase,
                "residual_scope": capture.PHASE_RESIDUAL_SCOPE,
                "instrumentation_overhead_ns": None,
                "instrumentation_scope": capture.PHASE_INSTRUMENTATION_SCOPE,
                "allocation_scope": capture.PHASE_ALLOCATION_SCOPE,
            } for _ in rows],
            "resource": {"maximum_resident_set_size_(kbytes)": 1},
        }
        with mock.patch.object(capture, "_within_child_bootstrap", return_value={"median": 10}):
            metrics = capture._entry_metrics(entry)
        allocation = metrics["allocation"]
        self.assertEqual(allocation["availability"], "measured")
        self.assertEqual(allocation["vectors"]["allocation_calls"], [11] * 30)
        self.assertEqual(allocation["vectors"]["allocation_peak_increment_bytes"], [30] * 30)
        self.assertEqual(allocation["percentiles"]["allocation_peak_increment_bytes"]["p50"], 30)
        self.assertIn("region_peak_live_bytes - live_bytes_before", allocation["derived_fields"]["allocation_peak_increment_bytes"])

    def test_paired_bootstrap_is_deterministic_and_block_sized(self) -> None:
        values = [10, 20]
        first = capture.paired_bootstrap(values)
        second = capture.paired_bootstrap(values)
        self.assertEqual(first, second)
        self.assertEqual(first["blocks"], 2)
        self.assertEqual(first["seed"], capture.BOOTSTRAP_SEED)
        self.assertEqual(first["resamples"], capture.BOOTSTRAP_REPETITIONS)
        self.assertEqual(first["median"], 15)

    def test_one_adverse_repeat_is_not_hidden_by_two_repeat_median(self) -> None:
        def metrics(value: int) -> dict[str, object]:
            percentile = {"n": 30, "min": value, "max": value, "mean": value,
                          "p50": value, "p95": value, "p99": value}
            phases = {name: percentile for name in capture.PHASE_DURATION_FIELDS +
                      ("phase_sum_ns", "lifecycle_residual_ns")}
            return {
                "latency_ns": percentile,
                "latency_median_ci": {"median": value},
                "rss_bytes": value,
                "phases": phases,
                "phase_median_ci": {name: {"median": value} for name in phases},
            }

        runs = capture.formal_inventory()
        entries = []
        for repeat in capture.REPEATS:
            for phase in capture.PHASES:
                spec = next(item for item in runs
                            if item["repeat"] == repeat and item["phase"] == phase
                            and item["role"] == "normal" and item["arm"] == "owned"
                            and item["api"] == capture.UNMANAGED)
                entry = {"spec": spec}
                if phase == "before":
                    entry["_metrics"] = metrics(100)
                else:
                    entry["_metrics"] = metrics(106 if repeat == 1 else 104)
                entries.append(entry)
        comparisons = capture._paired_comparisons(entries)
        latency = next(item for item in comparisons if item["metric"] == "latency_ns.p50")
        self.assertEqual(latency["individual_adverse_blocks"], [1])
        self.assertTrue(latency["adverse_flag"])
        self.assertFalse(latency["aggregate_adverse_flag"])

    def test_allocator_pairs_and_repeat_visibility_include_full_lifecycle_fields(self) -> None:
        def metrics(value: int) -> dict[str, object]:
            percentile = {"n": 30, "min": value, "max": value, "mean": value,
                          "p50": value, "p95": value, "p99": value}
            phases = {name: percentile for name in capture.PHASE_DURATION_FIELDS +
                      ("phase_sum_ns", "lifecycle_residual_ns")}
            allocation_percentiles = {name: percentile for name in capture.ALLOCATION_VECTOR_FIELDS}
            allocation_ci = {name: {"median": value} for name in capture.ALLOCATION_VECTOR_FIELDS}
            return {
                "latency_ns": percentile, "latency_median_ci": {"median": value},
                "rss_bytes": value, "phases": phases,
                "phase_median_ci": {name: {"median": value} for name in phases},
                "allocation": {
                    "availability": "measured", "scope": capture.SAMPLE_ALLOCATION_SCOPE,
                    "vectors": {name: [value] * 30 for name in capture.ALLOCATION_VECTOR_FIELDS},
                    "percentiles": allocation_percentiles, "median_ci": allocation_ci,
                },
            }

        runs = capture.formal_inventory()
        entries = []
        for repeat in capture.REPEATS:
            for phase_name in capture.PHASES:
                spec = next(item for item in runs
                            if item["repeat"] == repeat and item["phase"] == phase_name
                            and item["role"] == "allocator" and item["arm"] == "owned"
                            and item["api"] == capture.UNMANAGED)
                entries.append({"spec": spec, "_metrics": metrics(100 if phase_name == "before" else 110)})
        comparisons = capture._paired_comparisons(entries)
        names = {item["metric"] for item in comparisons}
        self.assertIn("allocation.allocation_calls.p50", names)
        self.assertIn("allocation.allocation_peak_increment_bytes.p50", names)
        peak = next(item for item in comparisons
                    if item["metric"] == "allocation.allocation_peak_increment_bytes.p50")
        self.assertEqual(peak["blocks"][0]["absolute_delta"], 10)
        repeats = capture._repeat_visibility(entries)
        self.assertEqual(len(repeats), 2)
        for repeat in repeats:
            self.assertEqual(repeat["allocation"]["availability"], "measured")
            self.assertIn("allocation_peak_increment_bytes", repeat["allocation"]["metrics"])


if __name__ == "__main__":
    unittest.main()
