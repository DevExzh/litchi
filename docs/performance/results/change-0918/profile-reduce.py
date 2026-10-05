#!/usr/bin/env python3
"""Reduce retained callgrind captures after the offline replay gates pass."""

from __future__ import annotations

import importlib.util
import json
from pathlib import Path


PACKET = Path(__file__).resolve().parent
REPLAY_PATH = PACKET / "profile-replay.py"


def replay_module():
    specification = importlib.util.spec_from_file_location("profile_replay_0918", REPLAY_PATH)
    if specification is None or specification.loader is None:
        raise RuntimeError(f"cannot load {REPLAY_PATH}")
    module = importlib.util.module_from_spec(specification)
    specification.loader.exec_module(module)
    return module


def main() -> None:
    replay = replay_module()
    _, _, summaries, _ = replay.validate_receipt()
    captures = []
    for summary in summaries:
        captures.append({
            "id": summary["id"],
            "raw_path": summary["raw_path"],
            "raw_bytes": summary["raw_bytes"],
            "raw_sha256": summary["raw_sha256"],
            "stderr_path": summary["stderr_path"],
            "stderr_bytes": summary["stderr_bytes"],
            "stderr_sha256": summary["stderr_sha256"],
            "run_case_calls": summary["run_case_calls"],
            "output_sha256": summary["output_sha256"],
            "output_bytes": summary["output_bytes"],
            "events": summary["events"],
        })
    print(json.dumps({
        "schema": "litchi.performance.0918.callgrind-reduction.v2",
        "captures": captures,
    }, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
