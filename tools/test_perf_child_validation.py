"""Focused integration tests for ``perf_child_validation``.

The child is a short Python fixture, not an Office workload.  These tests
exercise the same subprocess -> terminal receipt -> report validator -> result
construction path that a future capture driver will use.
"""

from __future__ import annotations

import json
import hashlib
import importlib.util
import subprocess
import sys
import tempfile
import unittest
from dataclasses import replace
from pathlib import Path
from typing import Any, Mapping

from tools import perf_child_validation


DIGEST = "a" * 64
ROOT = Path(__file__).resolve().parents[1]
HISTORICAL_VALIDATOR = ROOT / "docs/performance/results/change-0937/comparison.py"
HISTORICAL_VALIDATOR_SHA256 = (
    "b70a2b38d098e3abd2b567d1bea70f40e9d27e62acaca8f80ff633dd2c596712"
)
HISTORICAL_REPORT = ROOT / "tools/testdata/perf_child_validation_0937_report.json"


class _FixtureValidator:
    def __init__(self, terminal_path: Path) -> None:
        self.terminal_path = terminal_path
        self.calls = 0

    def validate_report(
        self,
        report: dict[str, Any],
        route: str,
        *,
        warmups: int,
        samples: int,
        label: str,
        binary: Mapping[str, Any],
    ) -> Mapping[str, Any]:
        self.calls += 1
        # The runner must publish terminal state before this callback is
        # entered; this is the ordering assertion the old synthetic controls
        # failed to exercise.
        if not self.terminal_path.is_file():
            raise AssertionError("terminal receipt was not written first")
        return {"route": route, "elapsed_ns": [11, 13], "warmups": warmups}


def _child_command(report: Path, mode: str) -> tuple[str, ...]:
    code = (
        "import json, pathlib, subprocess, sys; "
        "path = pathlib.Path(sys.argv[1]); mode = sys.argv[2]; "
        "subprocess.Popen([sys.executable, '-c', 'import time; time.sleep(10)']) if mode == 'leak' else None; "
        "row = {}; "
        "digest = 'a' * 64 if mode in ('valid', 'leak') else 'b' * 64; "
        "row.update({'output_sha256': digest}) if mode in ('valid', 'leak', 'mutated') else None; "
        "value = {'results': [row]}; "
        "value.update({'output_sha256': 'a' * 64}) if mode == 'top_only' else None; "
        "path.write_text(json.dumps(value), encoding='utf-8')"
    )
    return (sys.executable, "-c", code, str(report), mode)


def _historical_child_command(report: Path, mode: str = "valid") -> tuple[str, ...]:
    code = (
        "import json, pathlib, sys; "
        "source = pathlib.Path(sys.argv[1]); destination = pathlib.Path(sys.argv[2]); "
        "mode = sys.argv[3]; value = json.loads(source.read_text(encoding='utf-8')); "
        "value['results'][0].pop('output_sha256', None) if mode == 'missing' else None; "
        "value['results'][0].update({'output_sha256': 'b' * 64}) if mode == 'mutated' else None; "
        "destination.write_text(json.dumps(value), encoding='utf-8')"
    )
    return (sys.executable, "-c", code, str(HISTORICAL_REPORT), str(report), mode)


def _load_historical_validator() -> Any:
    digest = hashlib.sha256(HISTORICAL_VALIDATOR.read_bytes()).hexdigest()
    if digest != HISTORICAL_VALIDATOR_SHA256:
        raise AssertionError("historical validator dependency changed")
    spec = importlib.util.spec_from_file_location(
        "comparison_0937_child_validation_fixture", HISTORICAL_VALIDATOR
    )
    if spec is None or spec.loader is None:
        raise AssertionError("historical validator could not be loaded")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ChildValidationIntegrationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        self.child = self.root / "child"
        self.child.mkdir()
        self.paths = perf_child_validation.ChildPaths(
            report=self.child / "report.json",
            stdout=self.child / "stdout.log",
            stderr=self.child / "stderr.log",
            terminal=self.child / "terminal.json",
            validation=self.child / "validation.json",
        )

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    def spec(self, mode: str) -> perf_child_validation.ChildSpec:
        return perf_child_validation.ChildSpec(
            command=_child_command(self.paths.report, mode),
            cwd=self.root,
            environment={
                "PATH": "/usr/bin:/bin",
                "LC_ALL": "C",
                "LANG": "C",
                "TZ": "UTC",
                "PYTHONUNBUFFERED": "1",
            },
            paths=self.paths,
            route="fixture_route",
            label="fixture-child",
            warmups=1,
            samples=2,
            expected_output_sha256=DIGEST,
            binary={"path": "fixture-binary", "bytes": 1, "sha256": "c" * 64},
            deadline_seconds=10.0,
        )

    def run_historical_cli(self, mode: str = "valid") -> subprocess.CompletedProcess[str]:
        validator = _load_historical_validator()
        return subprocess.run(
            [
                sys.executable,
                str(ROOT / "tools/perf_child_validation.py"),
                "--validator",
                str(HISTORICAL_VALIDATOR),
                "--validator-sha256",
                HISTORICAL_VALIDATOR_SHA256,
                "--cwd",
                str(self.root),
                "--report",
                str(self.paths.report),
                "--stdout",
                str(self.paths.stdout),
                "--stderr",
                str(self.paths.stderr),
                "--terminal",
                str(self.paths.terminal),
                "--validation",
                str(self.paths.validation),
                "--route",
                "generic_read_at",
                "--label",
                "historical-cli-fixture",
                "--warmups",
                "5",
                "--samples",
                "50",
                "--expected-output-sha256",
                validator.OUTPUT_SHA256,
                "--env",
                "PATH=/usr/bin:/bin",
                "--env",
                "LC_ALL=C",
                "--env",
                "LANG=C",
                "--env",
                "TZ=UTC",
                "--command",
                *_historical_child_command(self.paths.report, mode),
            ],
            cwd=ROOT,
            env={"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"},
            text=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=False,
        )

    def test_success_runs_child_terminal_validation_and_result_in_order(self) -> None:
        validator = _FixtureValidator(self.paths.terminal)
        outcome = perf_child_validation.run_child_and_validate(self.spec("valid"), validator)

        self.assertEqual(outcome["status"], "passed")
        self.assertEqual(validator.calls, 1)
        self.assertEqual(outcome["result"]["output_sha256"], DIGEST)
        self.assertEqual(outcome["result"]["sample_count"], 2)
        self.assertEqual(json.loads(self.paths.terminal.read_text())["status"], "terminal")
        validation = json.loads(self.paths.validation.read_text())
        self.assertEqual(validation["status"], "passed")
        self.assertEqual(validation["output_sha256"], DIGEST)

    def test_rejects_duplicate_output_paths_before_starting_child(self) -> None:
        duplicate_paths = replace(self.paths, stderr=self.paths.stdout)
        spec = replace(self.spec("valid"), paths=duplicate_paths)

        with self.assertRaisesRegex(
            perf_child_validation.ChildValidationError, "output paths must be distinct"
        ):
            perf_child_validation.run_child_and_validate(spec, _FixtureValidator(self.paths.terminal))
        self.assertFalse(self.paths.terminal.exists())

    def test_oversized_report_is_described_and_rejected_before_validation(self) -> None:
        validator = _FixtureValidator(self.paths.terminal)
        spec = replace(
            self.spec("valid"),
            command=_historical_child_command(self.paths.report),
            route="generic_read_at",
            warmups=5,
            samples=50,
            max_report_bytes=64,
        )
        outcome = perf_child_validation.run_child_and_validate(spec, validator)

        self.assertEqual(outcome["status"], "failed")
        self.assertEqual(validator.calls, 0)
        self.assertTrue(self.paths.terminal.is_file())
        terminal = json.loads(self.paths.terminal.read_text())
        self.assertGreater(terminal["report"]["bytes"], 64)
        self.assertIsNone(terminal["report"]["sha256"])
        self.assertIn("size limit", terminal["report_error"])

    def test_orphaned_descendant_is_cleaned_but_marks_terminal_failed(self) -> None:
        validator = _FixtureValidator(self.paths.terminal)
        outcome = perf_child_validation.run_child_and_validate(self.spec("leak"), validator)

        self.assertEqual(outcome["status"], "failed")
        self.assertEqual(validator.calls, 0)
        terminal = json.loads(self.paths.terminal.read_text())
        self.assertTrue(terminal["process_group_forced_cleanup"])
        self.assertTrue(terminal["process_group_closed"])
        self.assertIn("forced cleanup", outcome["validation"]["error"])

    def test_historical_validator_result_is_completed_by_real_child_adapter(self) -> None:
        validator = _load_historical_validator()
        report = json.loads(HISTORICAL_REPORT.read_text(encoding="utf-8"))
        historical_result = validator.validate_report(
            report,
            "generic_read_at",
            warmups=5,
            samples=50,
            label="historical-fixture",
            binary=None,
        )
        # This is the exact result contract that caused the 0938 KeyError.
        with self.assertRaises(KeyError):
            _ = historical_result["output_sha256"]

        process = self.run_historical_cli()
        self.assertEqual(process.returncode, 0, process.stderr)
        cli_result = json.loads(process.stdout)
        self.assertEqual(cli_result["status"], "passed")
        self.assertEqual(cli_result["validation"]["output_sha256"], validator.OUTPUT_SHA256)
        self.assertEqual(json.loads(self.paths.validation.read_text())["status"], "passed")

    def test_historical_cli_rejects_missing_nested_digest_after_terminal(self) -> None:
        process = self.run_historical_cli("missing")
        self.assertEqual(process.returncode, 1, process.stderr)
        self.assertEqual(json.loads(process.stdout)["status"], "failed")
        self.assertTrue(self.paths.terminal.is_file())
        self.assertEqual(json.loads(self.paths.terminal.read_text())["exit_code"], 0)
        self.assertIn("output", json.loads(self.paths.validation.read_text())["error"])

    def test_historical_cli_rejects_mutated_nested_digest_after_terminal(self) -> None:
        process = self.run_historical_cli("mutated")
        self.assertEqual(process.returncode, 1, process.stderr)
        self.assertEqual(json.loads(process.stdout)["status"], "failed")
        self.assertTrue(self.paths.terminal.is_file())
        self.assertIn("output", json.loads(self.paths.validation.read_text())["error"])


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
