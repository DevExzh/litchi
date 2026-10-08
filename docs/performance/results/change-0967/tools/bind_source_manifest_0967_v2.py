#!/usr/bin/env python3
"""Refresh 0967 source manifests with the sealed matched control overlay.

This is a custody-only operation.  It never builds or launches a workload.
The v1 manifests remain historical; v2 records the identical
``tools/perf-revision-control-0967`` trees that the formal control binaries
are built from.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
from typing import Any

REVISION = "893072657ad8b1d726bfce351b5ec2e98f084e94"
OVERLAY = (
    "tools/perf-revision-control-0967/Cargo.toml",
    "tools/perf-revision-control-0967/Cargo.lock",
    "tools/perf-revision-control-0967/src/lib.rs",
    "tools/perf-revision-control-0967/src/main.rs",
    "tools/perf-revision-control-0967/src/bin/alloc.rs",
    "tools/perf-revision-control-0967/src/alloc_metrics.rs",
)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def load(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise SystemExit(f"{path}: expected JSON object")
    return value


def rows(root: Path) -> list[dict[str, Any]]:
    result = []
    for relative in OVERLAY:
        path = root / relative
        if not path.is_file():
            raise SystemExit(f"missing control overlay file: {path}")
        result.append({"path": relative, "bytes": path.stat().st_size, "sha256": sha(path)})
    cargo = (root / OVERLAY[0]).read_text(encoding="utf-8")
    for dependency in ("litchi-core = { path = \"../../crates/litchi-core\" }", "litchi-doc = { path = \"../../crates/litchi-doc\" }"):
        if dependency not in cargo:
            raise SystemExit(f"control overlay dependency is not the sealed relative path: {dependency}")
    return result


def main() -> int:
    base = Path("/home/zhuhe/litchi-performance-0967")
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidate-root", type=Path, default=base / "source-candidate")
    ap.add_argument("--baseline-root", type=Path, default=base / "harness-control/source-baseline")
    ap.add_argument("--candidate-v1", type=Path, default=base / "receipts/source-candidate-manifest-v1.json")
    ap.add_argument("--baseline-v1", type=Path, default=base / "receipts/source-baseline-manifest-v1.json")
    ap.add_argument("--candidate-output", type=Path, default=base / "receipts/source-candidate-manifest-v2.json")
    ap.add_argument("--baseline-output", type=Path, default=base / "receipts/source-baseline-manifest-v2.json")
    args = ap.parse_args()

    candidate = args.candidate_root.resolve()
    baseline = args.baseline_root.resolve()
    candidate_v1 = args.candidate_v1.resolve()
    baseline_v1 = args.baseline_v1.resolve()
    candidate_doc = load(candidate_v1)
    baseline_doc = load(baseline_v1)
    if candidate_doc.get("source_revision") != REVISION or baseline_doc.get("source_revision") != REVISION:
        raise SystemExit("source revision mismatch")
    candidate_rows = rows(candidate)
    baseline_rows = rows(baseline)
    if candidate_rows != baseline_rows:
        raise SystemExit("baseline/candidate control overlay differs")

    def common(variant: str, counterpart_root: Path) -> dict[str, Any]:
        return {
            "timed_control_overlay": {
                "relative_root": "tools/perf-revision-control-0967",
                "dependency_root": "../../crates",
                "files_equal": True,
                "variant": variant,
                "counterpart_root": str(counterpart_root),
                "files": candidate_rows,
            },
            "parent_manifest": {
                "candidate": {"path": str(candidate_v1), "sha256": sha(candidate_v1)},
                "baseline": {"path": str(baseline_v1), "sha256": sha(baseline_v1)},
            },
            "build_started": False,
            "workload_started": False,
        }
    candidate_out = {**candidate_doc, "schema": "litchi.performance.0967.source-candidate.v2", **common("candidate", baseline)}
    baseline_out = {**baseline_doc, "schema": "litchi.performance.0967.source-baseline.v2", **common("baseline", candidate)}
    for path, document in ((args.candidate_output.resolve(), candidate_out), (args.baseline_output.resolve(), baseline_out)):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(document, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({
        "status": "sealed_for_matched_build",
        "candidate": {"path": str(args.candidate_output.resolve()), "sha256": sha(args.candidate_output.resolve())},
        "baseline": {"path": str(args.baseline_output.resolve()), "sha256": sha(args.baseline_output.resolve())},
        "overlay_sha256": candidate_rows,
    }))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
