#!/usr/bin/env python3
"""Check the capture artifact empty-log boundary without running a workload."""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path
import tempfile


ROOT = Path(__file__).resolve().parent
CHECK = ROOT / "capture-selftest-check.json"


def load_capture_module():
    path = ROOT / "capture.py"
    spec = importlib.util.spec_from_file_location("change0419_capture", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> int:
    capture = load_capture_module()
    with tempfile.TemporaryDirectory(prefix="litchi-0419-capture-selftest-") as raw:
        root = Path(raw)
        empty_log = root / "stderr.txt"
        empty_log.touch()
        accepted = capture.artifact_record(empty_log, root, allow_empty=True)
        rejected = False
        try:
            capture.artifact_record(empty_log, root)
        except capture.CaptureError:
            rejected = True
        report = root / "report.json"
        report.write_text("{}\n", encoding="utf-8")
        nonempty = capture.artifact_record(report, root)
    if not rejected or accepted["bytes"] != 0 or nonempty["bytes"] != 3:
        raise SystemExit("capture artifact empty-log boundary check failed")
    payload = json.dumps({
            "schema_version": 1,
            "change": 419,
            "status": "pass",
            "checks": {
                "empty_stdout_or_stderr_allowed": True,
                "empty_report_rejected": True,
                "nonempty_report_accepted": True,
            },
        }, indent=2, sort_keys=True) + "\n"
    if CHECK.exists():
        if CHECK.read_text(encoding="utf-8") != payload:
            raise SystemExit("retained capture self-test differs from replay")
    else:
        CHECK.write_text(payload, encoding="utf-8")
    print(json.dumps({"change": 419, "status": "pass"}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
