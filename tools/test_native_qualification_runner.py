#!/usr/bin/env python3
"""Static/unit checks for the native qualification launcher.

These tests exercise manifest and guard logic only.  They do not invoke Cargo,
Rust, a benchmark binary, or a timing workload.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import re
import sys
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))
import native_qualification_runner as runner  # noqa: E402


class NativeQualificationRunnerTests(unittest.TestCase):
    def test_cargo_flags_are_added_once_before_rustc_separator(self) -> None:
        argv = runner.normalize_cargo_argv(
            ["cargo", "build", "--locked", "--", "-C", "link-arg=--offline"]
        )
        self.assertEqual(argv, ["cargo", "build", "--locked", "--offline", "--", "-C", "link-arg=--offline"])

    def test_python_literal_does_not_trigger_heavy_process_detection(self) -> None:
        python = {
            "comm": "python3",
            "exe": "/usr/bin/python3",
            "cwd": str(ROOT),
            "cmdline": ["python3", "-c", "print('cargo')"],
        }
        cargo = {
            "comm": "cargo",
            "exe": "/home/toolchain/bin/cargo",
            "cwd": "/home/zhuhe/code/litchi-worktrees/other",
            "cmdline": ["cargo", "build"],
        }
        unrelated = {
            "comm": "cargo",
            "exe": "/home/toolchain/bin/cargo",
            "cwd": "/tmp/unrelated-repository",
            "cmdline": ["cargo", "build"],
        }
        self.assertFalse(runner._is_heavy(python))
        self.assertTrue(runner._is_heavy(cargo))
        self.assertTrue(runner._in_litchi_scope(cargo))
        self.assertFalse(runner._in_litchi_scope(unrelated))

    def test_manifests_are_parameterized_and_keep_workload_order(self) -> None:
        paths = [
            ROOT / "tools/native_qualification/manifests/w4b-xlsb.json",
            ROOT / "tools/native_qualification/manifests/legacy-w5.json",
        ]
        for path in paths:
            text = path.read_text(encoding="utf-8")
            self.assertNotRegex(text, r"(?<![A-Za-z0-9])[0-9a-f]{40}(?![A-Za-z0-9])")
            data = json.loads(text)
            self.assertEqual(data["abba_order"], list(runner.ABBA_ORDER))
            self.assertEqual(
                [item["id"] for item in data["legs"]["before"]["programs"]],
                [item["id"] for item in data["legs"]["after"]["programs"]],
            )
            self.assertEqual(data["rounds"], 31)

    def test_manifest_shape_loads_without_source_ids(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-manifest-") as directory:
            target_root = Path(directory) / "targets"
            output = Path(directory) / "output"
            environment = {
                "NATIVE_QUAL_OUTPUT": str(output),
                "W4B_BEFORE_SOURCE": str(ROOT),
                "W4B_AFTER_SOURCE": str(Path(directory) / "w4b-after-source"),
                "W4B_TARGET_ROOT": str(target_root),
                "W4B_XLSB_FIXTURE": str(ROOT / "Cargo.lock"),
                "W5_BEFORE_SOURCE": str(ROOT),
                "W5_AFTER_SOURCE": str(Path(directory) / "w5-after-source"),
                "W5_TARGET_ROOT": str(target_root),
                "LITCHI_COORD_BIN": directory,
            }
            old = {key: os.environ.get(key) for key in environment}
            try:
                os.environ.update(environment)
                for name in ("w4b-xlsb.json", "legacy-w5.json"):
                    manifest_path = ROOT / "tools/native_qualification/manifests" / name
                    data = runner.load_manifest(manifest_path)
                    self.assertEqual(data["schema"], runner.SCHEMA)
            finally:
                for key, value in old.items():
                    if value is None:
                        os.environ.pop(key, None)
                    else:
                        os.environ[key] = value

    def test_failed_child_is_journaled_before_error(self) -> None:
        with tempfile.TemporaryDirectory(prefix="native-qual-child-") as directory:
            root = Path(directory)
            helper = root / "fail.py"
            helper.write_text("raise SystemExit(7)\n", encoding="utf-8")
            output = root / "output"
            target = root / "target"
            manifest = {
                "schema": runner.SCHEMA,
                "output_dir": str(output),
                "target_limit_bytes": 20 * 1024**3,
                "memory_limit_bytes": 1024**3,
                "abba_order": list(runner.ABBA_ORDER),
                "shared": {"probe_paths": [str(ROOT / "tools/native_qualification_runner.py")]},
                "legs": {
                    "before": {
                        "source_root": str(ROOT),
                        "target_dir": str(target),
                        "lock_paths": ["tools/native_qualification_runner.py"],
                        "allow_dirty": True,
                    },
                    "after": {
                        "source_root": str(ROOT),
                        "target_dir": str(target),
                        "lock_paths": ["tools/native_qualification_runner.py"],
                        "allow_dirty": True,
                    },
                },
            }
            safe = runner.SafeRunner(manifest, ROOT / "tools/native_qualification_runner.py", "build", None)
            # The real host currently has unrelated wave builders; this unit
            # test isolates child-journaling behavior from the host guard.
            safe.guard.require_quiet = lambda: None
            safe.guard.offenders = lambda _allowed_root=None: []
            identity = safe._identity("before")
            with self.assertRaises(runner.RunnerError):
                safe._run_command(
                    event_id="failed-child",
                    leg="before",
                    argv=[sys.executable, str(helper)],
                    cwd=ROOT,
                    env=os.environ.copy(),
                    identity=identity,
                )
            journal = output / "events-build.jsonl"
            self.assertTrue(journal.is_file())
            event = json.loads(journal.read_text(encoding="utf-8").splitlines()[0])
            self.assertEqual(event["returncode"], 7)
            self.assertEqual(event["event_id"], "failed-child")


if __name__ == "__main__":
    unittest.main()
