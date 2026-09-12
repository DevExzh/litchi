#!/usr/bin/env python3
"""Small tamper checks for the 0517 evidence verifier."""

from __future__ import annotations

import copy
import importlib.util
import tempfile
from pathlib import Path


HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("change0517_verify_tested", HERE / "verify.py")
assert spec is not None and spec.loader is not None
verify = importlib.util.module_from_spec(spec)
spec.loader.exec_module(verify)


def rejects(function):
    try:
        function()
    except (verify.VerificationError, AssertionError, KeyError, ValueError):
        return
    raise AssertionError("tampered evidence was accepted")


def main() -> None:
    result = verify.verify_bundle()
    assert result["status"] in {"pending-candidate", "pending-quality", "pass"}
    assert result["candidate_complete"] is (result["status"] != "pending-candidate")

    name = "p128-k1-owned-batch"
    rows = verify.read_rows(HERE / "preflight" / f"{name}.csv")
    damaged_rows = copy.deepcopy(rows)
    damaged_rows[0]["output_sha256"] = "0" * 64
    rejects(lambda: verify.capture.validate(name, damaged_rows, 1, 0, 1))

    receipt = verify.load_json(HERE / "preflight" / f"{name}.json")
    damaged_receipt = copy.deepcopy(receipt)
    damaged_receipt["artifacts"][f"{name}.csv"] = "0" * 64
    rejects(lambda: verify.verify_artifacts(HERE / "preflight", name,
                                            damaged_receipt, "native"))

    raw = HERE / "profile-preflight" / f"{name}.callgrind"
    with tempfile.TemporaryDirectory(prefix="litchi-verify-0517-") as directory:
        damaged_raw = Path(directory) / raw.name
        content = raw.read_text(encoding="utf-8")
        damaged_raw.write_text(content.replace("summary: 6662867", "summary: 1", 1),
                               encoding="utf-8")
        rejects(lambda: verify.verify_profile_scope(damaged_raw))

    rejects(lambda: verify.verify_timeline([(0.0, 2.0, "build"),
                                             (1.0, 3.0, "capture")]))
    print("0517 verifier tamper checks passed")


if __name__ == "__main__":
    main()
