#!/usr/bin/env python3
"""Inventory one completed 0960 epoch without copying its raw capture files."""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
from datetime import datetime, timezone
from pathlib import Path


def descriptor(path: Path, root: Path) -> dict[str, object]:
    data = path.read_bytes()
    mode = stat.S_IMODE(path.stat().st_mode)
    return {
        "path": str(path),
        "relative_path": str(path.relative_to(root)),
        "bytes": len(data),
        "mode": mode,
        "sha256": hashlib.sha256(data).hexdigest(),
    }


def inventory(root: Path) -> dict[str, object]:
    if not root.is_dir():
        raise SystemExit(f"root is not a directory: {root}")
    files: list[dict[str, object]] = []
    symlinks: list[dict[str, object]] = []
    for path in sorted(root.rglob("*")):
        if path.is_symlink():
            symlinks.append({
                "path": str(path),
                "relative_path": str(path.relative_to(root)),
                "target": os.readlink(path),
            })
        elif path.is_file():
            files.append(descriptor(path, root))
    return {
        "path": str(root),
        "regular_file_count": len(files),
        "symlink_count": len(symlinks),
        "total_bytes": sum(int(item["bytes"]) for item in files),
        "max_file_bytes": max((int(item["bytes"]) for item in files), default=0),
        "files": files,
        "symlinks": symlinks,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--root", required=True, type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    output = args.output
    if output.exists():
        raise SystemExit(f"refusing to overwrite existing output: {output}")
    result = {
        "schema": "litchi.performance.0960.epoch2.external-manifest.v1",
        "generated_utc": datetime.now(timezone.utc).isoformat(),
        "scope": "All regular files and symlink entries under one completed epoch; raw reports remain external.",
        "root": inventory(args.root.resolve()),
        "regeneration": {
            "command": f"python3 -B {Path(__file__).resolve()} --root {args.root.resolve()} --output {output.resolve()}",
            "output_must_not_exist": True,
            "capture_native_historical": "LITCHI_PERF_0960_CAPTURE_APPROVED=1 /usr/bin/taskset --cpu-list 12-15 python3 -B /home/zhuhe/litchi-performance-0960/tools/screen_0960_v7.py --capture --lane native --go --plan /tmp/litchi-performance-0960-epoch2/plan-final-v1.json --packet /tmp/litchi-performance-0960-epoch2/capture-native-v1",
            "capture_allocator_historical": "LITCHI_PERF_0960_CAPTURE_APPROVED=1 /usr/bin/taskset --cpu-list 12-15 python3 -B /home/zhuhe/litchi-performance-0960/tools/screen_0960_v7.py --capture --lane allocator --go --plan /tmp/litchi-performance-0960-epoch2/plan-final-v1.json --packet /tmp/litchi-performance-0960-epoch2/capture-allocator-v1",
            "historical_capture_scope": "The two capture commands above describe the completed epoch and its retained packet roots. They are provenance only; do not rerun them against those roots or treat them as a safe recapture recipe.",
            "future_capture_template": "LITCHI_PERF_0960_CAPTURE_APPROVED=1 /usr/bin/taskset --cpu-list 12-15 python3 -B /home/zhuhe/litchi-performance-0960/tools/screen_0960_v7.py --capture --lane <native|allocator> --go --plan <FRESH_REVIEWED_PLAN> --packet <FRESH_EMPTY_PACKET>",
            "future_capture_requirements": "A future epoch requires a newly reviewed plan, source/build/binary bindings, root GO receipt, and fresh non-existing packet/output paths; the frozen plan's capture_root points at retained evidence and must not be reused for overwrite.",
            "reduce_historical_packets": "python3 -B /home/zhuhe/litchi-performance-0960/tools/reduce_0960_v3.py --plan /tmp/litchi-performance-0960-epoch2/plan-final-v1.json --native-packet /tmp/litchi-performance-0960-epoch2/capture-native-v1 --allocator-packet /tmp/litchi-performance-0960-epoch2/capture-allocator-v1 --output /tmp/litchi-performance-0960-epoch2-reduction-replay.json",
            "select_reduced_output": "python3 -B /home/zhuhe/litchi-performance-0960/tools/select_screen_0960_v1.py --report /tmp/litchi-performance-0960-epoch2-reduction-replay.json --rule /home/zhuhe/litchi-performance-0960/review/root-screening-rule-v2.json --output /tmp/litchi-performance-0960-epoch2-selection-replay.json",
            "raw_audit_v2_replay": "python3 -B /home/zhuhe/code/litchi/docs/performance/results/change-0960/epoch2/review/audit_epoch2_raw_v2.py /tmp/litchi-performance-0960-epoch2 --output /tmp/litchi-performance-0960-epoch2/review/independent-raw-audit-v2-replay.json",
        },
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x") as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write("\n")
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
