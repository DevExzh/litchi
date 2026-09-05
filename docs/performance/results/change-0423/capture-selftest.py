#!/usr/bin/env python3
"""Small no-workload checks for the 0423 capture custody helpers."""

from __future__ import annotations

import argparse
import importlib.util
import json
from pathlib import Path
import tempfile


HERE = Path(__file__).resolve().parent


def capture_module():
    spec = importlib.util.spec_from_file_location("change0423_capture", HERE / "capture.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load capture.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def expect_failure(function, message: str) -> None:
    try:
        function()
    except Exception:
        return
    raise AssertionError(f"expected failure: {message}")


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, help="optional retained JSON check record")
    args = parser.parse_args()
    capture = capture_module()
    with tempfile.TemporaryDirectory(prefix="change-0423-capture-selftest-") as raw:
        root = Path(raw)
        empty_stderr = root / "stderr.txt"
        empty_stdout = root / "stdout.txt"
        empty_report = root / "empty-report.json"
        report = root / "report.json"
        empty_stderr.write_bytes(b"")
        empty_stdout.write_bytes(b"")
        empty_report.write_bytes(b"")
        report.write_text("{}\n", encoding="utf-8")
        accepted_stderr = capture.artifact_record(empty_stderr, root, allow_empty=True)
        accepted_stdout = capture.artifact_record(empty_stdout, root, allow_empty=True)
        expect_failure(
            lambda: capture.artifact_record(empty_report, root, allow_empty=False),
            "empty report should be rejected",
        )
        accepted_report = capture.artifact_record(report, root, allow_empty=False)
        verifier = root / "verifier.json"
        valid_verifier = {
            "status": "pass",
            "claim_authorized": False,
            "performance_claim": None,
            "selector": "pptx_cross_copy_plain_lifecycle",
            "lane": "normal",
            "samples": 100,
            "warmups": 10,
            "report_count": 1,
            "reports": [{}],
        }
        write_json(verifier, valid_verifier)
        capture.verify_output(
            verifier, "valid verifier", selector=valid_verifier["selector"],
            lane="normal", samples=100, warmups=10,
        )
        empty_verifier = root / "empty-verifier.json"
        write_json(empty_verifier, {})
        expect_failure(
            lambda: capture.verify_output(
                empty_verifier, "empty verifier", selector=valid_verifier["selector"],
                lane="normal", samples=100, warmups=10,
            ),
            "empty verifier object should be rejected",
        )
        wrong_lane = root / "wrong-lane.json"
        wrong_lane_value = dict(valid_verifier, lane="allocator")
        write_json(wrong_lane, wrong_lane_value)
        expect_failure(
            lambda: capture.verify_output(
                wrong_lane, "wrong lane", selector=valid_verifier["selector"],
                lane="normal", samples=100, warmups=10,
            ),
            "wrong verifier lane should be rejected",
        )
        claimed = root / "claimed.json"
        write_json(claimed, dict(valid_verifier, claim_authorized=True))
        expect_failure(
            lambda: capture.verify_output(
                claimed, "claimed verifier", selector=valid_verifier["selector"],
                lane="normal", samples=100, warmups=10,
            ),
            "authorized claim should be rejected",
        )
        missing_identity = root / "missing-identity.json"
        write_json(missing_identity, {"environment": {"git_revision": "a" * 40, "git_worktree_dirty": False}})
        expect_failure(
            lambda: capture.validate_report_identity(
                capture.load_json(missing_identity, "missing identity"),
                {"revision": "a" * 40}, {"sha256": "b" * 64, "bytes": 1},
                "missing identity",
            ),
            "missing report binary identity should be rejected",
        )
        result = {
            "change": 423,
            "status": "pass",
            "empty_stdout_allowed": accepted_stdout["bytes"] == 0,
            "empty_stderr_allowed": accepted_stderr["bytes"] == 0,
            "nonempty_report_required": accepted_report["bytes"] > 0,
            "verifier_envelope_required": True,
            "wrong_lane_rejected": True,
            "claim_rejected": True,
            "report_identity_required": True,
            "temporary_directory_cleaned_by_context": True,
        }
    if args.output is not None:
        output = args.output.expanduser().resolve()
        if output.exists():
            raise SystemExit(f"refusing to overwrite selftest output: {output}")
        output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
