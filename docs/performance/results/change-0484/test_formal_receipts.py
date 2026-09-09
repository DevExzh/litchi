"""Formal capture failures must retain terminal, attributable receipts."""

from contextlib import ExitStack
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0, str(Path(__file__).resolve().parent))
import measure_routes as routes


class FailedCaptureTests(unittest.TestCase):
    def test_shared_validator_failure_keeps_route_and_axis_receipts(self):
        for axis in (False, True):
            with self.subTest(axis=axis), tempfile.TemporaryDirectory() as temporary:
                root = Path(temporary)
                build = root / "route-attempts" / "test" / "build-normal.json"
                build.parent.mkdir(parents=True)
                build.write_text("{}\n")
                binary = {"path": "/unused/test-binary", "sha256": "1" * 64, "bytes": 1}

                def child(argv, **kwargs):
                    # Preserve real directory creation, argv, artifact hashing,
                    # and receipt writing. Only process execution is simulated.
                    Path(argv[argv.index("--json") + 1]).write_text("{}\n")
                    Path(argv[argv.index("-o") + 1]).write_text("resource\n")
                    return subprocess.CompletedProcess(argv, 0)

                validator = "_check_axis_report" if axis else "check_route_report"
                with ExitStack() as stack:
                    stack.enter_context(patch.object(routes, "ROOT", root))
                    stack.enter_context(patch.object(routes, "load_protocol", return_value=({}, "2" * 64)))
                    stack.enter_context(patch.object(routes, "_load_builds", return_value={"normal": {"binary": binary}}))
                    stack.enter_context(patch.object(routes, "_machine_binding", return_value={"status": "ready"}))
                    stack.enter_context(patch.object(routes.subprocess, "run", side_effect=child))
                    stack.enter_context(patch.object(routes, validator, side_effect=routes.base.MeasureError("shared oracle rejected report")))
                    with self.assertRaisesRegex(routes.RouteMeasureError, "receipt retained"):
                        if axis:
                            routes.capture_axis_one("test", "axis-sink-4096-s64-a64-short-c64", "normal", pilot=False, repeat=1)
                        else:
                            routes.capture_one("test", "deterministic", "normal", "s64-a64-short-c64", pilot=False, repeat=1)
                receipts = list(root.glob("*-captures/test/*/receipt.json"))
                self.assertEqual(len(receipts), 1)
                receipt = json.loads(receipts[0].read_text())
                self.assertEqual(receipt["status"], "failed")
                self.assertEqual(receipt["exit_code"], 0)
                self.assertEqual(receipt["validation_error"], "shared oracle rejected report")
                self.assertEqual(receipt["missing_artifacts"], [])
                self.assertEqual(set(receipt["artifacts"]), {"report.json", "resource.txt", "stdout.txt", "stderr.txt"})
                for name, expected in receipt["artifacts"].items():
                    self.assertEqual(routes.meta(receipts[0].parent / name), expected)


if __name__ == "__main__":
    unittest.main()
