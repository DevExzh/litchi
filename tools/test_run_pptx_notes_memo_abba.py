#!/usr/bin/env python3
"""Process-boundary tests for ``run_pptx_notes_memo_abba``."""

from __future__ import annotations

import json
import shutil
import stat
import sys
import tempfile
import unittest
from pathlib import Path

from tools import run_pptx_notes_memo_abba as runner
from tools.test_pptx_notes_memo_abba import report, statistics


CHILD = (
    "import pathlib, shutil, sys; "
    "shutil.copyfile(sys.argv[1], sys.argv[2])"
)


def _target(root: Path, name: str, marker: bytes) -> dict[str, object]:
    path = root / name
    shutil.copy2(sys.executable, path)
    path.write_bytes(path.read_bytes() + marker)
    path.chmod(stat.S_IMODE(path.stat().st_mode) | stat.S_IXUSR)
    return runner._file_identity(path, name)


def manifest(root: Path, *, unique_content: bool = True) -> tuple[Path, Path]:
    output_dir = root / "reports"
    receipt = root / "receipt.json"
    control = _target(root, "control-target", b"control-target-marker")
    candidate = _target(root, "candidate-target", b"candidate-target-marker")
    template_paths: dict[tuple[int, str], Path] = {}
    for round_index in range(1, runner.MIN_PROCESS_ROUNDS + 1):
        for label, revision, target in (
            ("a1", "a" * 40, control),
            ("b1", "b" * 40, candidate),
            ("b2", "b" * 40, candidate),
            ("a2", "a" * 40, control),
        ):
            template = root / f"{round_index}-{label}-template.json"
            data = report(revision, str(target["sha256"]))
            data["binary_identity"] = {
                "path": target["path"],
                "binary_sha256": target["sha256"],
                "binary_bytes": target["bytes"],
                "mode_bits": target["mode_bits"],
                "executable": True,
                "profile": "release",
            }
            offset = round_index * 1000 + {"a1": 0, "b1": 1, "b2": 2, "a2": 3}[label]
            for row in data["results"]:
                values = [observation["elapsed_ns"] + offset for observation in row["observations"]]
                for observation, value in zip(row["observations"], values):
                    observation["elapsed_ns"] = value
                row["elapsed_ns"] = statistics(values)
            template.write_text(json.dumps(data, allow_nan=False), encoding="utf-8")
            template_paths[(round_index, label)] = template
    rounds = []
    for round_index in range(1, runner.MIN_PROCESS_ROUNDS + 1):
        legs = []
        for label in runner.LEG_ORDER:
            source_label = label
            if not unique_content:
                source_label = "a1" if label.startswith("a") else "b1"
            template = template_paths[(round_index, source_label)]
            target = control if label.startswith("a") else candidate
            output = output_dir / f"round-{round_index}-{label}.json"
            command = [
                str(target["path"]),
                "-c",
                CHILD,
                str(template),
                str(output),
            ]
            legs.append(
                {
                    "label": label,
                    "command": command,
                    "report": str(output),
                    "source_revision": "a" * 40 if label.startswith("a") else "b" * 40,
                    "target_binary": target,
                }
            )
        rounds.append({"index": round_index, "legs": legs})
    manifest_path = root / "manifest.json"
    manifest_path.write_text(
        json.dumps(
            {
                "schema": runner.MANIFEST_SCHEMA,
                "mode": "native",
                "output_dir": str(output_dir),
                "timeout_seconds": 30.0,
                "rounds": rounds,
            },
            allow_nan=False,
        ),
        encoding="utf-8",
    )
    return manifest_path, receipt


class ProcessAbbaTests(unittest.TestCase):
    def test_four_explicit_rounds_produce_sixteen_fresh_processes(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            manifest_path, receipt_path = manifest(Path(directory))
            receipt = runner.run_manifest(manifest_path, receipt_path)
            self.assertEqual(receipt["process_round_count"], runner.MIN_PROCESS_ROUNDS)
            self.assertEqual(receipt["process_count"], runner.MIN_PROCESS_ROUNDS * 4)
            self.assertEqual(receipt["leg_order"], list(runner.LEG_ORDER))
            self.assertEqual(len(receipt["round_summaries"]), runner.MIN_PROCESS_ROUNDS)
            self.assertEqual(receipt["matrix"]["processes"], 16)
            self.assertEqual(receipt["matrix"]["sample_rounds_per_report"], 4)
            identities = {
                (process["pid"], process["start_time_ticks"])
                for process in receipt["processes"]
            }
            self.assertEqual(len(identities), 16)
            report_hashes = [process["report_sha256"] for process in receipt["processes"]]
            self.assertEqual(len(set(report_hashes)), 16)
            self.assertEqual(
                {process["round"] for process in receipt["processes"]},
                set(range(1, runner.MIN_PROCESS_ROUNDS + 1)),
            )
            for process in receipt["processes"]:
                self.assertEqual(process["process_executable"], process["target_binary"]["path"])
                self.assertTrue(process["target_binary"]["mode_bits"] & 0o111)
            self.assertEqual(
                receipt["manifest_sha256"],
                runner._sha256(manifest_path.read_bytes()),
            )
            self.assertTrue(receipt_path.is_file())

    def test_identical_report_content_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            manifest_path, receipt_path = manifest(Path(directory), unique_content=False)
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)

    def test_manifest_unknown_key_and_nonfinite_json_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest_path, receipt_path = manifest(root)
            value = json.loads(manifest_path.read_text(encoding="utf-8"))
            value["unexpected"] = True
            manifest_path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)

            value.pop("unexpected")
            manifest_path.write_text(
                json.dumps(value).replace('"timeout_seconds": 30.0', '"timeout_seconds": NaN'),
                encoding="utf-8",
            )
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)

    def test_manifest_requires_expected_target_identity_and_execute_bits(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest_path, receipt_path = manifest(root)
            value = json.loads(manifest_path.read_text(encoding="utf-8"))
            value["rounds"][0]["legs"][0]["target_binary"]["mode_bits"] = 0o644
            manifest_path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)

    def test_source_and_same_pid_target_bindings_are_fail_closed(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest_path, receipt_path = manifest(root)
            value = json.loads(manifest_path.read_text(encoding="utf-8"))
            value["rounds"][0]["legs"][0]["target_binary"]["sha256"] = "0" * 64
            manifest_path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest_path, receipt_path = manifest(root)
            value = json.loads(manifest_path.read_text(encoding="utf-8"))
            value["rounds"][0]["legs"][0]["command"][0] = sys.executable
            manifest_path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)

    def test_duplicate_manifest_member_is_rejected_before_execution(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest_path, receipt_path = manifest(root)
            text = manifest_path.read_text(encoding="utf-8")
            text = text.replace('"mode": "native"', '"mode": "native", "mode": "native"', 1)
            manifest_path.write_text(text, encoding="utf-8")
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)


if __name__ == "__main__":
    unittest.main()
