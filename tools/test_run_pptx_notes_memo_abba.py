#!/usr/bin/env python3
"""Process-boundary tests for ``run_pptx_notes_memo_abba``."""

from __future__ import annotations

import json
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


def manifest(root: Path, *, unique_content: bool = True) -> tuple[Path, Path]:
    output_dir = root / "reports"
    receipt = root / "receipt.json"
    template_paths: dict[str, Path] = {}
    legs = []
    for label, revision, binary_sha in (
        ("a1", "a" * 40, "1" * 64),
        ("b1", "b" * 40, "2" * 64),
        ("b2", "b" * 40, "2" * 64),
        ("a2", "a" * 40, "1" * 64),
    ):
        template = root / f"{label}-template.json"
        data = report(revision, binary_sha)
        pair = "control" if label.startswith("a") else "candidate"
        data["binary_identity"]["path"] = f"/tmp/{pair}-binary"
        offset = {"a1": 0, "b1": 1, "b2": 2, "a2": 3}[label]
        for row in data["results"]:
            values = [observation["elapsed_ns"] + offset for observation in row["observations"]]
            for observation, value in zip(row["observations"], values):
                observation["elapsed_ns"] = value
            row["elapsed_ns"] = statistics(values)
        template.write_text(json.dumps(data), encoding="utf-8")
        template_paths[label] = template
    for label in runner.LEG_ORDER:
        output = output_dir / f"{label}.json"
        if unique_content:
            command = [
                sys.executable,
                "-c",
                CHILD,
                str(template_paths[label]),
                str(output),
            ]
        else:
            command = [
                sys.executable,
                "-c",
                CHILD,
                str(template_paths["a1"]),
                str(output),
            ]
        legs.append({"label": label, "command": command, "report": str(output)})
    manifest_path = root / "manifest.json"
    manifest_path.write_text(
        json.dumps(
            {
                "schema": runner.MANIFEST_SCHEMA,
                "mode": "native",
                "output_dir": str(output_dir),
                "timeout_seconds": 30.0,
                "legs": legs,
            }
        ),
        encoding="utf-8",
    )
    return manifest_path, receipt


class ProcessAbbaTests(unittest.TestCase):
    def test_four_fresh_processes_produce_a_valid_finite_matrix(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            manifest_path, receipt_path = manifest(Path(directory))
            receipt = runner.run_manifest(manifest_path, receipt_path)
            self.assertEqual(receipt["process_count"], 4)
            self.assertEqual(receipt["leg_order"], list(runner.LEG_ORDER))
            self.assertEqual(receipt["summary"]["rounds"], 4)
            identities = {
                (process["pid"], process["start_time_ticks"])
                for process in receipt["processes"]
            }
            self.assertEqual(len(identities), 4)
            report_hashes = [process["report_sha256"] for process in receipt["processes"]]
            self.assertEqual(len(set(report_hashes)), 4)
            self.assertTrue(receipt_path.is_file())

    def test_identical_report_content_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            manifest_path, receipt_path = manifest(Path(directory), unique_content=False)
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)

    def test_manifest_unknown_key_and_nonfinite_timeout_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest_path, receipt_path = manifest(root)
            value = json.loads(manifest_path.read_text(encoding="utf-8"))
            value["unexpected"] = True
            manifest_path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)

            value.pop("unexpected")
            value["timeout_seconds"] = float("nan")
            manifest_path.write_text(json.dumps(value), encoding="utf-8")
            with self.assertRaises(runner.RunnerError):
                runner.run_manifest(manifest_path, receipt_path)


if __name__ == "__main__":
    unittest.main()
