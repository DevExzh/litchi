#!/usr/bin/env python3
"""Pure custody and recomputation tests for the 0493 capture driver.

These tests never freeze a protocol, build a binary, or launch a benchmark.
They exercise the invariants that make a later capture auditable.
"""

from __future__ import annotations

import copy
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure as matrix  # noqa: E402


def _ranges(arm: dict[str, object]) -> list[dict[str, int]]:
    source = matrix.LOGICAL_RANGES if arm["policy"] == "exact" else matrix.PHYSICAL_CANDIDATE
    return [{"offset": offset, "requested": size, "returned": size} for offset, size in source]


def _row(arm: dict[str, object], role: str = "normal") -> dict[str, object]:
    physical_ranges = _ranges(arm)
    physical_bytes = sum(item["requested"] for item in physical_ranges)
    counts = [0] * matrix.REQUEST_BUCKETS
    for item in physical_ranges:
        counts[matrix._bucket(item["requested"])] += 1
    transport = {"logical_calls": len(physical_ranges), "requested_bytes": physical_bytes,
                 "returned_bytes": physical_bytes, "min_request_bytes": min(item["requested"] for item in physical_ranges),
                 "max_request_bytes": max(item["requested"] for item in physical_ranges),
                 "short_reads": 0, "delayed_calls": len(physical_ranges),
                 "transfer_paced_calls": len(physical_ranges) if arm["transfer_bytes_per_second"] else 0,
                 "transfer_delay_ns": 0, "request_size_counts": counts}
    budget = {"memory_before": 0, "memory_live": 1024, "memory_after_drop": 0,
              "memory_released": 1024, "input_bytes_before": 0,
              "input_bytes_after_drop": physical_bytes, "input_bytes_delta": physical_bytes,
              "objects_before": 0, "objects_live": 1, "objects_after_drop": 0,
              "objects_released": 1, "managed": True,
              "memory_released_to_baseline": True,
              "input_bytes_match_physical_returned": True}
    allocation: dict[str, object] = {}
    if role == "allocator":
        sample = {"status": "measured", "scope": matrix.SAMPLE_ALLOCATION_SCOPE,
                  "allocation_calls": 4, "deallocation_calls": 2, "reallocation_calls": 1,
                  "failed_allocation_calls": 0, "allocated_bytes": 100,
                  "deallocated_bytes": 50, "live_bytes_before": 60, "live_bytes_after": 110,
                  "peak_live_bytes_before": 70, "peak_live_bytes_after": 120,
                  "region_peak_live_bytes": 100}
        allocation = {"sample": sample, "region_peak_increment_bytes": 40}
    source_read = None
    if arm["policy"] != "exact":
        source_read = {"enabled": True, "configured_window_bytes": matrix.WINDOW_BYTES,
                       "retained_window_bytes": matrix.WINDOW_BYTES, "requests": 19,
                       "hits": 16, "misses": 3, "fills": 3,
                       "requested_bytes": matrix.BASELINE_PHYSICAL_BYTES,
                       "returned_bytes": matrix.CANDIDATE_PHYSICAL_BYTES}
    return {"sample_index": 0, "elapsed_ns": 100, "actual_text_verified": True,
            "actual_text_bytes": matrix.CORPUS["expected_text_bytes"],
            "actual_text_sha256": matrix.CORPUS["expected_text_sha256"],
            "source_version_before": {"id": 188, "revision": 0},
            "source_version_after": {"id": 188, "revision": 0},
            "source_version_unchanged": True,
            "policy": {"name": arm["policy"], "enabled": arm["policy"] != "exact",
                        "configured_window_bytes": arm["window_bytes"]},
            "source_read": source_read,
            "physical": {"scope": matrix.PHYSICAL_SCOPE, "calls": len(physical_ranges),
                          "requested_bytes": physical_bytes, "returned_bytes": physical_bytes,
                          "short_reads": 0, "ranges": physical_ranges},
            "transport": transport, "budget": budget, "allocation": allocation,
            "cache": {"open_successful_loads": 0, "successful_loads": 1,
                      "failed_loads": 0, "retained_bytes": 1024,
                      "retained_entries": 1, "budget_managed": True}}


def _report(arm_name: str, role: str = "normal") -> dict[str, object]:
    arm = matrix.ARM_BY_NAME[arm_name]
    return {"schema": matrix.SCHEMA, "version": 1, "case_name": matrix.CASE,
            "provider_scope": matrix.PROVIDER_SCOPE, "timing_scope": matrix.TIMING_SCOPE,
            "setup_scope": matrix.SETUP_SCOPE, "allocation_scope": matrix.ALLOCATION_SCOPE,
            "corpus_version": "source-edit-media-v1", "corpus_generator": matrix.CORPUS["generator"],
            "corpus": {"generator": matrix.CORPUS["generator"], "archive_bytes": matrix.CORPUS["archive_bytes"],
                       "archive_sha256": matrix.CORPUS["archive_sha256"],
                       "archive_member_count": matrix.CORPUS["archive_member_count"]},
            "source_bytes": matrix.CORPUS["archive_bytes"], "source_sha256": matrix.CORPUS["archive_sha256"],
            "expected_text_bytes": matrix.CORPUS["expected_text_bytes"],
            "expected_text_sha256": matrix.CORPUS["expected_text_sha256"],
            "expected_archive_members": matrix.CORPUS["archive_member_count"],
            "media_ranges": [{"start": start, "end": end} for start, end in matrix.MEDIA_RANGES],
            "expected_exact_ranges": [{"offset": offset, "requested": size, "returned": size}
                                       for offset, size in matrix.LOGICAL_RANGES],
            "expected_exact_physical_calls": matrix.BASELINE_PHYSICAL_CALLS,
            "expected_exact_physical_bytes": matrix.BASELINE_PHYSICAL_BYTES,
            "requested_source_revision": "0" * 40, "limits": matrix.LIMITS,
            "provider": {"name": "PptxRangeSource", "max_range_bytes": arm["max_range_bytes"],
                         "delay_us": arm["delay_us"], "transfer_bytes_per_second": arm["transfer_bytes_per_second"],
                         "transfer_delay_policy": arm["transfer_delay_policy"],
                         "physical_scope": matrix.PHYSICAL_SCOPE, "range_scope": matrix.RANGE_SCOPE},
            "policy": {"name": arm["policy"], "enabled": arm["policy"] != "exact",
                        "configured_window_bytes": arm["window_bytes"]},
            "warmup": 0, "samples": 1, "rows": [_row(arm, role)]}


class ManagedMeasureTests(unittest.TestCase):
    def test_inventory_reverses_role_and_arm_order(self) -> None:
        runs = matrix.formal_inventory()
        self.assertEqual(len(runs), 16)
        self.assertEqual([item["role"] for item in runs[:4]], ["normal"] * 4)
        self.assertEqual([item["role"] for item in runs[4:8]], ["allocator"] * 4)
        self.assertEqual([item["role"] for item in runs[8:12]], ["allocator"] * 4)
        self.assertEqual([item["role"] for item in runs[12:]], ["normal"] * 4)
        self.assertEqual([item["arm"] for item in runs[8:12]], [arm["name"] for arm in reversed(matrix.ARMS)])
        self.assertEqual(len({item["label"] for item in runs}), 16)
        self.assertEqual(len(matrix.formal_inventory(pilot=True)), 8)

    def test_command_binds_managed_policy_and_delay_controls(self) -> None:
        build = {"binary": {"path": "/bin/true"}}
        spec = matrix.formal_inventory()[3]
        command = matrix._command(spec, build, Path("report.json"), Path("resource.txt"), "a" * 40)
        self.assertIn("docx-managed-read-ahead", command)
        self.assertIn("--window-bytes", command)
        self.assertIn("4096", command)
        delayed = matrix.formal_inventory()[2]
        delayed_command = matrix._command(delayed, build, Path("report.json"), Path("resource.txt"), "a" * 40)
        self.assertIn("--transfer-bytes-per-second", delayed_command)
        self.assertIn("minimum-service", delayed_command)

    def test_report_and_range_conservation_are_fail_closed(self) -> None:
        value = _report("managed-4096-0us")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            path.write_text(json.dumps(value), encoding="utf-8")
            matrix.validate_report(path, role="normal", arm_name="managed-4096-0us",
                                   samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["physical"]["ranges"][0]["returned"] += 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="managed-4096-0us",
                                       samples=1, warmups=0)

    def test_budget_rejects_negative_or_nonconserving_input(self) -> None:
        value = _report("exact-0us")
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "report.json"
            altered = copy.deepcopy(value)
            altered["rows"][0]["budget"]["input_bytes_delta"] = -1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="exact-0us", samples=1, warmups=0)
            altered = copy.deepcopy(value)
            altered["rows"][0]["budget"]["input_bytes_delta"] += 1
            path.write_text(json.dumps(altered), encoding="utf-8")
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix.validate_report(path, role="normal", arm_name="exact-0us", samples=1, warmups=0)

    def test_deterministic_bootstrap_and_percentiles(self) -> None:
        values = [9, 1, 5, 3]
        self.assertEqual(matrix._percentiles(values)["p50"], 4)
        self.assertEqual(matrix._bootstrap_median(values), matrix._bootstrap_median(values))

    def test_timeout_kills_the_process_group(self) -> None:
        process = subprocess.Popen(["sleep", "30"], start_new_session=True)
        termination = matrix._kill_group(process)
        self.assertIn(termination, ("SIGTERM", "SIGKILL"))
        self.assertIsNotNone(process.returncode)

    def _terminal_fixture(self, directory: Path) -> tuple[dict[str, object], dict[str, object], dict[str, object]]:
        """Build one valid custody envelope without starting a benchmark."""
        directory.mkdir(exist_ok=True)
        (directory / "stdout.txt").write_bytes(b"")
        (directory / "stderr.txt").write_bytes(b"")
        (directory / "resource.txt").write_text("Maximum resident set size (kbytes): 7\n", encoding="utf-8")
        (directory / "report.json").write_text("{}\n", encoding="utf-8")
        (directory / "replay-cleanup.json").write_text(
            json.dumps({"schema": "docx-managed-private-cleanup-v1", "status": "pass",
                        "root": "/tmp/managed-run", "tmpdir": "/tmp/managed-run/tmp",
                        "removed": [], "remaining": []}), encoding="utf-8")
        spec = matrix.formal_inventory(pilot=True)[0]
        spec = dict(spec, attempt="fixture")
        source = {"files": 1, "path": "/tmp/source.json", "sha256": "c" * 64}
        build = {"path": "/tmp/build-normal.json", "receipt_sha256": "a" * 64,
                 "git_revision": "b" * 40, "binary": {"path": "/bin/true"}, "source": source}
        tmpdir = "/tmp/managed-run/tmp"
        environment = {key: matrix.ENV[key] for key in matrix.ENV_KEYS}
        environment["TMPDIR"] = tmpdir
        started = {
            "schema": matrix.CAPTURE_SCHEMA, "version": 1, "status": "running",
            "attempt": "fixture", "run": {key: spec[key] for key in
                                             ("kind", "repeat", "role", "arm", "provider", "samples", "warmups", "label")},
            "protocol": {"path": "protocol.json", "sha256": "d" * 64},
            "build": {"path": build["path"], "sha256": build["receipt_sha256"], "source": source},
            "binary": build["binary"], "source": source, "argv": [], "cwd": str(matrix.REPO),
            "environment": environment, "environment_artifact": {"path": "environment.json", "sha256": "e" * 64},
            "driver_bindings": {name: matrix._json_hash(matrix.ROOT / name) for name in matrix.DRIVER_FILES},
            "driver_sha256": matrix._json_hash(matrix.ROOT / "measure.py"),
            "support_sha256": matrix._json_hash(matrix.ROOT / "support.py"),
            "started_utc": "2026-09-10T00:00:00+00:00", "timeout_seconds": 180,
            "tmpdir": tmpdir,
        }
        started["argv"] = matrix._command(spec, build, directory / "report.json", directory / "resource.txt", build["git_revision"])
        (directory / "started.json").write_text(json.dumps(started), encoding="utf-8")
        cleanup = json.loads((directory / "replay-cleanup.json").read_text(encoding="utf-8"))
        terminal = dict(started)
        terminal.update({"schema": matrix.TERMINAL_SCHEMA, "status": "pass", "exit_code": 0,
                         "timed_out": False, "termination": None,
                         "finished_utc": "2026-09-10T00:00:01+00:00",
                         "artifacts": {name: {"path": str(directory / name),
                                                **matrix.meta(directory / name)}
                                       for name in ("stdout.txt", "stderr.txt", "resource.txt", "report.json", "replay-cleanup.json")},
                         "missing_artifacts": [], "cleanup": cleanup,
                         "started_artifact": matrix._file_meta(directory / "started.json"),
                         "source_before": source, "source_after": source, "source_unchanged": True})
        (directory / "terminal.json").write_text(json.dumps(terminal), encoding="utf-8")
        return started, terminal, build

    def test_terminal_exit_timeout_argv_and_helper_tampering_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            directory = Path(raw) / "capture"
            started, terminal, build = self._terminal_fixture(directory)
            spec = dict(matrix.formal_inventory(pilot=True)[0], attempt="fixture")
            matrix._validate_terminal(started, terminal, spec, build, "d" * 64, directory)
            for field, replacement in (("exit_code", 9), ("timed_out", True), ("argv", terminal["argv"] + ["tampered"]),):
                altered = copy.deepcopy(terminal)
                altered[field] = replacement
                with self.subTest(field=field), self.assertRaises(matrix.ProviderMatrixError):
                    matrix._validate_terminal(started, altered, spec, build, "d" * 64, directory)
            altered_started = copy.deepcopy(started)
            altered_started["driver_bindings"]["measure.py"] = "f" * 64
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix._validate_terminal(altered_started, terminal, spec, build, "d" * 64, directory)

    def test_source_manifest_binding_rejects_changed_digest(self) -> None:
        with tempfile.TemporaryDirectory() as raw:
            manifest = Path(raw) / "manifest.json"
            manifest.write_text(json.dumps({"source.rs": "a" * 64}, sort_keys=True, indent=2) + "\n", encoding="utf-8")
            binding = {"files": 1, "path": str(manifest), "sha256": "0" * 64}
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix._source_binding(binding, "fixture.source")

    def _build_fixture(self, directory: Path) -> Path:
        """Create a source-bound successful build/gate receipt without Cargo."""
        directory.mkdir(exist_ok=True)
        source_path = directory / "source.json"
        source_path.write_text(json.dumps({"source.rs": "a" * 64}, sort_keys=True, indent=2) + "\n", encoding="utf-8")
        source = {"files": 1, "path": str(source_path), "sha256": matrix.sha(source_path)}
        gate_path = directory / "gate-normal.json"
        gate_stdout = gate_path.with_suffix(".stdout")
        gate_stderr = gate_path.with_suffix(".stderr")
        gate_stdout.write_bytes(b"ok\n")
        gate_stderr.write_bytes(b"")
        command = matrix._expected_build_command("normal")
        environment = {key: matrix.ENV[key] for key in matrix.ENV_KEYS}
        gate = {
            "schema": "docx-managed-read-ahead-gate-v1", "label": "gate-normal", "attempt": "fixture",
            "argv": command, "cwd": str(matrix.REPO), "environment": environment,
            "driver_sha256": matrix._json_hash(matrix.ROOT / "gate.py"),
            "common_sha256": matrix._json_hash(matrix.ROOT / "support.py"),
            "started_utc": "2026-09-10T00:00:00+00:00", "finished_utc": "2026-09-10T00:00:01+00:00",
            "source_before": source, "source_after": source, "source_unchanged": True,
            "timeout_seconds": 1800, "timed_out": False, "termination": None, "exit_code": 0,
            "artifacts": {"gate-normal.stdout": matrix.meta(gate_stdout), "gate-normal.stderr": matrix.meta(gate_stderr)},
        }
        gate_path.write_text(json.dumps(gate), encoding="utf-8")
        binary_path = directory / "litchi-perf-baseline"
        original_dir = directory / "original"
        original_dir.mkdir()
        original_path = original_dir / "litchi-perf-baseline"
        shutil.copy2("/bin/true", binary_path)
        shutil.copy2("/bin/true", original_path)
        binary = {"path": str(binary_path), "bytes": binary_path.stat().st_size,
                   "sha256": matrix.sha(binary_path), "executable": True}
        original = {"path": str(original_path), "bytes": original_path.stat().st_size,
                    "sha256": matrix.sha(original_path), "executable": True}
        receipt = {
            "schema": "docx-provider-lifecycle-build-v1", "version": 1, "role": "normal",
            "attempt": "fixture", "copied_utc": "2026-09-10T00:00:02+00:00", "command": command,
            "environment": environment, "gate": {"path": str(gate_path), "sha256": matrix.sha(gate_path)},
            "binary": binary, "original_binary": original, "source_before": source,
            "source_after": source, "source_unchanged": True, "git_revision": "b" * 40,
            "retainer_sha256": matrix._json_hash(matrix.ROOT / "retain_build.py"),
        }
        path = directory / "build-normal.json"
        path.write_text(json.dumps(receipt), encoding="utf-8")
        return path

    def test_build_gate_and_retainer_bindings_reject_tampering(self) -> None:
        # This fixture is deliberately only a receipt test; no executable is
        # launched and no source is compiled.
        with tempfile.TemporaryDirectory() as raw:
            path = self._build_fixture(Path(raw))
            matrix._build_from(matrix._json(path), "normal", path)
            altered = matrix._json(path)
            altered["retainer_sha256"] = "f" * 64
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix._build_from(altered, "normal", path)
            altered = matrix._json(path)
            altered["source_after"] = dict(altered["source_after"], sha256="e" * 64)
            with self.assertRaises(matrix.ProviderMatrixError):
                matrix._build_from(altered, "normal", path)

    def test_plan_has_no_capture_side_effects(self) -> None:
        plan = matrix.protocol_value(None)
        self.assertIsNone(plan["source"])
        self.assertIsNone(plan["builds"])
        self.assertFalse(plan["claim_authorized"])


if __name__ == "__main__":
    raise SystemExit(unittest.main())
