#!/usr/bin/env python3
"""Fail closed unless the selected PPT source uses the approved Cargo lock."""

from __future__ import annotations

import hashlib
from pathlib import Path
import subprocess
import sys


APPROVED_CANDIDATE_LOCK = (
    "4100b1fbb272dd24842f33a2d6f353a3f3a97dfa7e435c4649a0e1a029bfaea1"
)


def main(argv: list[str]) -> int:
    try:
        separator = argv.index("--")
    except ValueError:
        separator = len(argv)
    if len(argv) < 3 or separator < 3 or (separator == len(argv) and len(argv) != 3):
        raise SystemExit(
            "usage: validate_ppt_candidate_lock.py CARGO_LOCK EXPECTED_SHA256"
            " [-- COMMAND ...]"
        )
    lock = Path(argv[1])
    expected = argv[2]
    if expected != APPROVED_CANDIDATE_LOCK:
        raise SystemExit("requested lock is not the approved PPT candidate lock")
    if not lock.is_file():
        raise SystemExit(f"candidate Cargo.lock is missing: {lock}")
    observed = hashlib.sha256(lock.read_bytes()).hexdigest()
    if observed != expected:
        raise SystemExit(
            f"candidate Cargo.lock hash mismatch: expected {expected}, observed {observed}"
        )
    if separator == len(argv):
        print(observed)
        return 0
    command = argv[separator + 1 :]
    if not command or any(not item for item in command):
        raise SystemExit("the command after -- is empty")
    return subprocess.run(command, check=False).returncode


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
