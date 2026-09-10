#!/usr/bin/env python3
"""Freeze one 0501 executable, source manifest, and protocol-driver identity."""

from __future__ import annotations

import argparse
import json
import os
import subprocess
from pathlib import Path

from custody import HERE, REPO, canonical_json, now, sha_bytes, sha_file, source_identity, source_snapshot


BOUND_FILES = (
    "protocol.json",
    "custody.py",
    "capture.py",
    "verify-report.py",
    "verify.py",
    "compare.py",
    "profile.py",
    "freeze.py",
    "crates/litchi-pptx/src/presentation/source_cross_copy.rs",
    "tools/perf-baseline/src/pptx_provider_lifecycle.rs",
    "tools/perf-baseline/src/lib.rs",
    "Cargo.lock",
)


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--phase", choices=("before", "after"), required=True)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--revision", default=None)
    parser.add_argument("--output", type=Path, default=None)
    args = parser.parse_args()

    output = (HERE / f"{args.phase}-freeze.json") if args.output is None else args.output
    output = output.resolve()
    if output.exists():
        raise SystemExit(f"refusing to replace existing freeze: {output}")

    binary = args.binary.resolve()
    if not binary.is_file():
        raise SystemExit(f"missing executable: {binary}")
    protocol_raw = (HERE / "protocol.json").read_bytes()
    protocol = json.loads(protocol_raw)
    if protocol["change"] != 501:
        raise SystemExit("protocol change identity mismatch")
    if protocol["status"] not in {"planned", "frozen"}:
        raise SystemExit("protocol must be planned or frozen")

    source = source_identity(source_snapshot())
    bound = {}
    for name in BOUND_FILES:
        path = REPO / name if not name.endswith(".py") and name != "protocol.json" and name != "custody.py" else HERE / name
        if not path.is_file():
            raise SystemExit(f"missing bound file: {name}")
        bound[name] = sha_file(path)

    revision = args.revision
    if revision is None:
        revision = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=REPO, text=True).strip()
    if len(revision) != 40 or revision.lower() != revision:
        raise SystemExit("revision must be a lowercase forty-character Git identity")

    record = {
        "change": 501,
        "phase": args.phase,
        "frozen_utc": now(),
        "revision": revision,
        "rust_toolchain": "1.98.1",
        "target_dir": os.environ.get("CARGO_TARGET_DIR", "/tmp/litchi-goal-0501-target"),
        "binary": {
            "path": str(binary),
            "bytes": binary.stat().st_size,
            "sha256": sha_file(binary),
        },
        "protocol_sha256": sha_bytes(protocol_raw),
        "source_manifest": source,
        "bound_files": bound,
        "source_candidate": "crates/litchi-pptx/src/presentation/source_cross_copy.rs",
        "harness": "tools/perf-baseline/src/pptx_provider_lifecycle.rs",
        "measurement_tmp_root": "/tmp/litchi-goal-0501",
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(canonical_json(record))
    print(json.dumps({"phase": args.phase, "binary_sha256": record["binary"]["sha256"], "source_files": source["files_count"]}, sort_keys=True))


if __name__ == "__main__":
    main()
