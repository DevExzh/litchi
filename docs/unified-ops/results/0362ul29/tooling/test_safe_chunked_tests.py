from __future__ import annotations

import hashlib
import importlib.util
import json
import os
from pathlib import Path
import stat
import subprocess
import sys
import tempfile
import textwrap
import unittest


MODULE_PATH = Path(__file__).with_name("safe_chunked_tests.py")
SPEC = importlib.util.spec_from_file_location("safe_chunked_tests_under_test", MODULE_PATH)
assert SPEC and SPEC.loader
safe = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(safe)


FAKE_COMMAND = textwrap.dedent(
    """
    import json
    from pathlib import Path
    import sys
    import time

    artifact = Path(sys.argv[1])
    marker = Path(sys.argv[2])
    artifact.parent.mkdir(parents=True, exist_ok=True)
    artifact.write_bytes(b"fake executable bytes")
    artifact.chmod(0o751)
    sidecar = artifact.with_suffix(".d")
    sidecar.write_bytes(b"fake private dependency sidecar")
    print(json.dumps({
        "reason": "compiler-artifact",
        "target": {"kind": ["example"]},
        "executable": str(artifact),
        "filenames": [str(artifact), str(sidecar)],
    }), flush=True)
    time.sleep(0.15)
    marker.write_text(str(artifact.exists()), encoding="utf-8")
    """
)


class SafeChunkedLifecycleTests(unittest.TestCase):
    def setUp(self) -> None:
        self.values = {
            "MAX_TARGET": safe.MAX_TARGET,
            "LIVE_TARGET": safe.LIVE_TARGET,
            "START_TARGET": safe.START_TARGET,
            "MIN_FREE": safe.MIN_FREE,
        }

    def tearDown(self) -> None:
        for name, value in self.values.items():
            setattr(safe, name, value)

    def _command(self, directory: Path, script: str) -> tuple[list[str], Path, Path, Path, Path, Path]:
        script_path = directory / "fake_command.py"
        script_path.write_text(script, encoding="utf-8")
        target = directory / "target"
        artifact = target / "debug" / "examples" / "fake-example"
        marker = directory / "marker"
        quarantine = directory / "quarantine"
        log = directory / "command.log"
        return [sys.executable, str(script_path), str(artifact), str(marker)], target, artifact, marker, quarantine, log

    def test_retirement_waits_for_child_and_preserves_hash_mode_and_manifest(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-chunked-lifecycle-") as raw:
            command, target, artifact, marker, quarantine, log = self._command(Path(raw), FAKE_COMMAND)
            safe.MIN_FREE = 0
            status, retired, _ = safe.execute(command, target, log, quarantine, 0)
            self.assertEqual(status, 0)
            self.assertEqual(marker.read_text(encoding="utf-8"), "True")
            self.assertEqual(len(retired), 2)
            record = retired[0]
            destination = Path(record["quarantine"])
            self.assertFalse(artifact.exists())
            self.assertFalse(artifact.with_suffix(".d").exists())
            self.assertEqual(destination.read_bytes(), b"fake executable bytes")
            self.assertEqual(record["sha256"], hashlib.sha256(destination.read_bytes()).hexdigest())
            self.assertEqual(record["mode"], "0751")
            self.assertEqual(stat.S_IMODE(destination.stat().st_mode), 0o751)
            manifest = quarantine / "manifest.jsonl"
            manifest_rows = [json.loads(line) for line in manifest.read_text(encoding="utf-8").splitlines()]
            self.assertEqual(manifest_rows[0]["quarantine"], str(destination))

    def test_nonzero_child_keeps_artifact_in_target(self) -> None:
        script = FAKE_COMMAND + "\nsys.exit(7)\n"
        with tempfile.TemporaryDirectory(prefix="safe-chunked-failure-") as raw:
            command, target, artifact, marker, quarantine, log = self._command(Path(raw), script)
            safe.MIN_FREE = 0
            status, retired, _ = safe.execute(command, target, log, quarantine, 0)
            self.assertEqual(status, 7)
            self.assertEqual(retired, [])
            self.assertTrue(artifact.exists())
            self.assertFalse((quarantine / "manifest.jsonl").exists())

    def test_descendant_finishes_before_retirement(self) -> None:
        script = textwrap.dedent(
            """
            import json
            from pathlib import Path
            import subprocess
            import sys

            artifact = Path(sys.argv[1])
            marker = Path(sys.argv[2])
            artifact.parent.mkdir(parents=True, exist_ok=True)
            artifact.write_bytes(b"descendant executable")
            artifact.chmod(0o751)
            subprocess.Popen(
                [sys.executable, "-c", "import pathlib,sys,time; time.sleep(.35); pathlib.Path(sys.argv[1]).write_text('done')", str(marker)],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            print(json.dumps({"reason": "compiler-artifact", "target": {"kind": ["test"]}, "executable": str(artifact)}), flush=True)
            """
        )
        with tempfile.TemporaryDirectory(prefix="safe-chunked-descendant-") as raw:
            command, target, artifact, marker, quarantine, log = self._command(Path(raw), script)
            safe.MIN_FREE = 0
            status, retired, _ = safe.execute(command, target, log, quarantine, 0)
            self.assertEqual(status, 0)
            self.assertEqual(marker.read_text(encoding="utf-8"), "done")
            self.assertEqual(len(retired), 1)
            self.assertFalse(artifact.exists())

    def test_live_overflow_stops_command_and_does_not_skip_or_retire_work(self) -> None:
        script = textwrap.dedent(
            """
            import json
            from pathlib import Path
            import sys
            import time

            artifact = Path(sys.argv[1])
            marker = Path(sys.argv[2])
            artifact.parent.mkdir(parents=True, exist_ok=True)
            artifact.write_bytes(b"not retired")
            artifact.chmod(0o751)
            print(json.dumps({"reason": "compiler-artifact", "target": {"kind": ["test"]}, "executable": str(artifact)}), flush=True)
            growth = artifact.parents[2] / "growth.bin"
            for _ in range(100):
                with growth.open("ab") as stream:
                    stream.write(b"x" * 8192)
                time.sleep(0.05)
            marker.write_text("after", encoding="utf-8")
            """
        )
        with tempfile.TemporaryDirectory(prefix="safe-chunked-overflow-") as raw:
            directory = Path(raw)
            command, target, artifact, marker, quarantine, log = self._command(directory, script)
            initial = safe.target_bytes(target)
            safe.MIN_FREE = 0
            safe.START_TARGET = initial + 1_000_000
            safe.LIVE_TARGET = initial + 32_000
            safe.MAX_TARGET = initial + 2_000_000
            with self.assertRaises(safe.TargetOverflow):
                safe.execute(command, target, log, quarantine, 0)
            self.assertFalse(marker.exists())
            self.assertTrue(artifact.exists())
            self.assertFalse((quarantine / "manifest.jsonl").exists())
            self.assertIn("safe_chunked_live_guard", log.read_text(encoding="utf-8"))

    def test_plan_partitions_every_test_and_example_without_skipping(self) -> None:
        metadata = {
            "packages": [
                {
                    "name": "fixture",
                    "targets": [
                        {"kind": ["lib"], "test": True, "doctest": True},
                        *({"kind": ["test"], "name": f"test_{i}", "test": True} for i in range(5)),
                        *({"kind": ["example"], "name": f"example_{i}", "test": False} for i in range(4)),
                    ],
                }
            ]
        }
        commands = safe.plan(metadata, {"fixture"}, [], 2)
        rendered = " ".join(" ".join(arguments) for _, arguments in commands)
        for name in [*(f"test_{i}" for i in range(5)), *(f"example_{i}" for i in range(4))]:
            self.assertEqual(rendered.count(name), 1)
        self.assertEqual(sum(1 for verb, _ in commands if verb == "test"), 5)
        self.assertEqual(sum(1 for verb, _ in commands if verb == "build"), 2)

    def test_symlinked_artifact_is_not_moved(self) -> None:
        with tempfile.TemporaryDirectory(prefix="safe-chunked-symlink-") as raw:
            directory = Path(raw)
            target = directory / "target"
            real = target / "debug" / "deps" / "real"
            real.parent.mkdir(parents=True)
            real.write_bytes(b"real")
            real.chmod(0o751)
            link = target / "debug" / "deps" / "link"
            link.symlink_to(real)
            quarantine = directory / "quarantine"
            manifest = directory / "manifest.jsonl"
            safe._retire_artifacts({link}, target, quarantine, manifest, ["fake"], 0, scan_private=False)
            self.assertTrue(link.is_symlink())
            self.assertFalse(manifest.exists())


if __name__ == "__main__":
    unittest.main()
