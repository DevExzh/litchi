#!/usr/bin/env python3
"""Unit tests for fail-closed operator gates and binary receipt checks."""

from __future__ import annotations

import sys
import json
import subprocess
import tempfile
import unittest
from contextlib import redirect_stdout
from io import StringIO
from pathlib import Path
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))

from operator_capture_support import (
    allowed_scenario_stages,
    binary_digests_match,
    output_path_without_symlink,
    receipt_digest,
)
import operator_capture


PROJECT_ROOT = Path(__file__).resolve().parents[4]
PERF_DIR = Path(__file__).resolve().parent


class OperatorCaptureSupportTests(unittest.TestCase):
    def test_binary_comparison_uses_digest_not_receipt_path(self) -> None:
        digest = "a" * 64
        self.assertEqual(receipt_digest(f"{digest}  /before/bin\n"), digest)
        self.assertTrue(binary_digests_match(f"{digest}  /before/bin\n", digest))
        self.assertTrue(binary_digests_match(f"{digest}  /different/path\n", digest))
        self.assertFalse(binary_digests_match(f"{digest}  /before/bin\n", "b" * 64))

    def test_malformed_receipts_fail_closed(self) -> None:
        with self.assertRaises(ValueError):
            receipt_digest("not-a-digest /bin")
        with self.assertRaises(ValueError):
            binary_digests_match(f"{'a' * 64} /bin\n", "not-a-digest")

    def test_failed_host_probe_suppresses_matrix_and_lanes(self) -> None:
        self.assertEqual(allowed_scenario_stages(1, None), ("host-probe",))

    def test_failed_matrix_suppresses_lanes(self) -> None:
        self.assertEqual(
            allowed_scenario_stages(0, 1),
            ("host-probe", "matrix-correctness"),
        )

    def test_success_allows_bounded_lane_stage(self) -> None:
        self.assertEqual(
            allowed_scenario_stages(0, 0),
            ("host-probe", "matrix-correctness", "lanes"),
        )

    def test_output_path_rejects_symlinked_parent(self) -> None:
        with tempfile.TemporaryDirectory(dir="/var/tmp") as temporary:
            root = Path(temporary)
            real = root / "real"
            real.mkdir()
            link = root / "link"
            link.symlink_to(real, target_is_directory=True)
            with self.assertRaises(ValueError):
                output_path_without_symlink(str(link / "results"))


class OperatorCaptureOrchestrationTests(unittest.TestCase):
    def _run_operator(
        self,
        *,
        guard_status: int = 0,
        manifest_status: int = 0,
        build_status: int = 0,
        host_status: int = 0,
        matrix_status: int = 0,
        remove_binary_before_final: bool = False,
    ) -> tuple[int, list[tuple[str, ...]], Path, dict[str, object]]:
        expected_head = subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=PROJECT_ROOT, text=True
        ).strip()
        calls: list[tuple[str, ...]] = []

        with tempfile.TemporaryDirectory(dir="/var/tmp") as temporary:
            temporary_root = Path(temporary)
            results = temporary_root / "results"
            target = temporary_root / "target"
            binary = target / "release/pptx-ink-actions-performance"
            metadata_calls = 0

            def fake_check_output(argv, *, cwd=None, text=False, **_kwargs):
                command = tuple(str(value) for value in argv)
                calls.append(command)
                if "status" in command:
                    return ""
                if "--verify" in command or command[-1] == "HEAD":
                    return expected_head + ("\n" if text else b"\n")
                raise AssertionError(f"unexpected check_output: {command}")

            def fake_run(argv, *, env=None, capture_output=False, text=False, **_kwargs):
                nonlocal metadata_calls
                command = tuple(str(value) for value in argv)
                calls.append(command)
                if command[0] == "file":
                    return subprocess.CompletedProcess(argv, 0, stdout="fake ELF\n", stderr="")
                if command[:2] == ("cargo", "metadata"):
                    metadata_calls += 1
                    status = 0
                    if remove_binary_before_final and metadata_calls == 2 and binary.is_file():
                        binary.unlink()
                elif command[:2] == ("cargo", "build"):
                    status = build_status
                    if status == 0:
                        binary.parent.mkdir(parents=True, exist_ok=True)
                        binary.write_bytes(b"operator-test-binary")
                elif command[0] == "rustc":
                    status = 0
                elif command[0] == "cargo":
                    status = 0
                elif command[0] == "rustfmt":
                    status = 0
                elif command[0] == "python3" and len(command) > 1 and Path(command[1]).name == "committed_inputs.py":
                    status = guard_status
                elif command[0] == "python3" and len(command) > 1 and Path(command[1]).name == "source_manifest.py":
                    status = manifest_status
                elif command[0] == str(binary) and "--host-probe" in command:
                    status = host_status
                elif command[0] == str(binary) and "--matrix-correctness" in command:
                    status = matrix_status
                elif command[0] == str(binary) and "--lane" in command:
                    status = 0
                else:
                    raise AssertionError(f"unexpected run: {command}")
                return subprocess.CompletedProcess(argv, status, stdout="", stderr="")

            with patch.object(operator_capture.subprocess, "check_output", side_effect=fake_check_output):
                with patch.object(operator_capture.subprocess, "run", side_effect=fake_run):
                    with redirect_stdout(StringIO()):
                        status = operator_capture.main(
                            [
                                "--expected-head",
                                expected_head,
                                str(PROJECT_ROOT),
                                str(PERF_DIR),
                                str(results),
                                str(target),
                            ]
                        )
            provenance = json.loads((results / "capture-provenance.json").read_text())
            return status, calls, target, provenance

    @staticmethod
    def _scenario_commands(calls: list[tuple[str, ...]], binary: Path) -> list[tuple[str, ...]]:
        return [command for command in calls if command and command[0] == str(binary)]

    def test_failed_preflight_skips_build_host_matrix_and_lanes(self) -> None:
        status, calls, target, _provenance = self._run_operator(guard_status=1)
        scenario = self._scenario_commands(calls, target / "release/pptx-ink-actions-performance")
        self.assertEqual(status, 1)
        self.assertEqual(scenario, [])
        self.assertFalse(any(command[:2] == ("cargo", "build") for command in calls))

    def test_failed_build_skips_host_matrix_and_lanes(self) -> None:
        status, calls, target, _provenance = self._run_operator(build_status=1)
        scenario = self._scenario_commands(calls, target / "release/pptx-ink-actions-performance")
        self.assertEqual(status, 1)
        self.assertEqual(scenario, [])

    def test_failed_host_probe_skips_matrix_and_lanes(self) -> None:
        status, calls, target, _provenance = self._run_operator(host_status=1)
        binary = target / "release/pptx-ink-actions-performance"
        scenario = self._scenario_commands(calls, binary)
        self.assertEqual(status, 1)
        self.assertEqual(len(scenario), 1)
        self.assertIn("--host-probe", scenario[0])

    def test_failed_matrix_skips_lanes(self) -> None:
        status, calls, target, _provenance = self._run_operator(matrix_status=1)
        binary = target / "release/pptx-ink-actions-performance"
        scenario = self._scenario_commands(calls, binary)
        self.assertEqual(status, 1)
        self.assertEqual(len(scenario), 2)
        self.assertIn("--host-probe", scenario[0])
        self.assertIn("--matrix-correctness", scenario[1])

    def test_missing_binary_after_scenarios_fails_capture(self) -> None:
        status, calls, target, provenance = self._run_operator(remove_binary_before_final=True)
        binary = target / "release/pptx-ink-actions-performance"
        scenario = self._scenario_commands(calls, binary)
        self.assertEqual(status, 1)
        self.assertEqual(len(scenario), 44)
        self.assertIn("built binary missing before final hash", provenance["failures"])


if __name__ == "__main__":
    unittest.main()
