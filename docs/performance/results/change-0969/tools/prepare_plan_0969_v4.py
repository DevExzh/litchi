#!/usr/bin/env python3
"""Emit a bounded prospective or formal 0969 plan outside the repository.

The ``aa`` mode binds only the retained baseline controls and leaves candidate
executables unbound.  The ``formal`` mode additionally binds the candidate
executables and a completed A/A receipt/manifest.  This file emits custody
metadata only; it never builds or launches a workload.
"""
from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from protocol_common_v4 import (
    ALLOCATOR_PAIRS,
    ALLOCATOR_SAMPLES,
    ALLOCATOR_WARMUPS,
    BOOTSTRAP_RESAMPLES,
    BOOTSTRAP_SEED,
    CASES,
    CI_LIMIT,
    EDITS,
    NATIVE_PAIRS,
    NATIVE_SAMPLES,
    NATIVE_WARMUPS,
    PLAN_SCHEMA,
    RETENTIONS,
    expected_cells,
    sha256,
    write_json_exclusive,
)

ROOT = Path("/home/zhuhe/litchi-performance-0969")
SOURCE_MANIFEST = ROOT / "manifests/source-manifest-v1.json"
DESIGN_REVIEW = ROOT / "review/design-review-v1.json"
BASELINE_REVIEW = ROOT / "review/baseline-reuse-v1.json"
PROTOCOL = ROOT / "protocol-design-v4.md"
TREE_SHA = "df8a842b08fec1792c67921fb1ce3513d3c6cb3dd0149cf13672bad14090bfb0"
BASELINE_NATIVE = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
BASELINE_ALLOCATOR = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile_alloc")


def bound_file(path: Path) -> dict[str, Any]:
    return {"path": str(path), "sha256": sha256(path)}


def maybe_file(path_text: str | None) -> dict[str, Any]:
    if not path_text:
        return {"path": "", "sha256": None}
    path = Path(path_text)
    return bound_file(path)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--mode", choices=("aa", "formal"), required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--candidate-native", type=Path)
    parser.add_argument("--candidate-allocator", type=Path)
    parser.add_argument("--candidate-manifest", type=Path)
    parser.add_argument("--candidate-root", type=Path, default=ROOT / "source-candidate")
    parser.add_argument("--candidate-tree-sha256", default=TREE_SHA)
    parser.add_argument("--baseline-manifest", type=Path)
    parser.add_argument("--baseline-root", type=Path, default=ROOT / "source-baseline")
    parser.add_argument("--baseline-tree-sha256", default=TREE_SHA)
    parser.add_argument("--aa-receipt", type=Path)
    parser.add_argument("--aa-manifest", type=Path)
    args = parser.parse_args()

    fixtures = {
        "docnohf": bound_file(ROOT / "qa-core/test-data/ole/doc/NoHeadFoot.doc") | {"bytes": (ROOT / "qa-core/test-data/ole/doc/NoHeadFoot.doc").stat().st_size},
        "docfloat": bound_file(ROOT / "qa-core/test-data/ole/doc/FloatingPictures.doc") | {"bytes": (ROOT / "qa-core/test-data/ole/doc/FloatingPictures.doc").stat().st_size},
    }
    cells = [{"cell_id": identity, "case": case, "edits": edits, "retention": retention}
             for case in CASES for edits in EDITS for retention in RETENTIONS
             for identity in [f"{case}/e{edits}/{retention}"]]
    baseline_manifest = args.baseline_manifest or SOURCE_MANIFEST
    source_row = {
        "manifest": str(baseline_manifest),
        "manifest_sha256": sha256(baseline_manifest),
        "root": str(args.baseline_root),
        "tree_sha256": args.baseline_tree_sha256,
    }
    candidate_manifest = args.candidate_manifest or SOURCE_MANIFEST
    candidate_source = {
        "manifest": str(candidate_manifest),
        "manifest_sha256": sha256(candidate_manifest),
        "root": str(args.candidate_root),
        "tree_sha256": args.candidate_tree_sha256,
    }
    aa_mode = args.mode == "aa"
    binaries = {
        "native_baseline": bound_file(BASELINE_NATIVE),
        "allocator_baseline": bound_file(BASELINE_ALLOCATOR),
        "native_candidate": maybe_file(str(args.candidate_native) if args.candidate_native else None),
        "allocator_candidate": maybe_file(str(args.candidate_allocator) if args.candidate_allocator else None),
    }
    helper_paths = {
        "protocol_common": ROOT / "tools/protocol_common_v4.py",
        "runner": ROOT / "tools/run_0969_v4.py",
        "reducer": ROOT / "tools/reduce_0969_v4.py",
        "plan_generator": Path(__file__).resolve(),
    }
    helpers = {name: bound_file(path) for name, path in helper_paths.items()}
    plan: dict[str, Any] = {
        "schema": PLAN_SCHEMA,
        "status": "aa_authorized" if aa_mode else "capture_authorized",
        "created_utc": "2026-10-08",
        "no_build_or_workload": aa_mode,
        "protocol_document": bound_file(PROTOCOL),
        "design_review": {"status": "pass", **bound_file(DESIGN_REVIEW)},
        "baseline_reuse_review": {"status": "pass", **bound_file(BASELINE_REVIEW)},
        "helpers": helpers,
        "source": {"baseline": source_row, "candidate": candidate_source},
        "fixtures": fixtures,
        "binaries": binaries,
        "execution": {"serial": True, "retry": False, "cargo_processes": 0, "child_timeout_seconds": 900, "quiescence_required": True},
        "real_matrix": {
            "cells": cells,
            "native": {
                "pairs_per_cell": NATIVE_PAIRS, "warmups": NATIVE_WARMUPS, "samples": NATIVE_SAMPLES,
                "relative_ratio": "candidate_over_baseline", "bootstrap_seed": BOOTSTRAP_SEED,
                "bootstrap_resamples": BOOTSTRAP_RESAMPLES, "noninferiority_upper_ci": CI_LIMIT,
                "p99": "descriptive_only",
            },
            "allocator": {
                "pairs_per_cell": ALLOCATOR_PAIRS, "warmups": ALLOCATOR_WARMUPS, "samples": ALLOCATOR_SAMPLES,
                "relative_ratio": "candidate_over_baseline", "bootstrap_seed": BOOTSTRAP_SEED,
                "bootstrap_resamples": BOOTSTRAP_RESAMPLES,
                "requested_metric": "samples[i].allocation.allocated_bytes",
                "rss_metric": "packet.rss.kb", "per_repeat_guards": {"allocated_bytes": 1.01, "peak_live_bytes": 1.05, "rss_kb": 1.05},
                "allocation_calls_median_guard": 1.01,
            },
        },
        "aa_noise": {
            "cells": expected_cells(), "coverage": "all_real_cells", "same_binary": "baseline",
            "order": ["AB", "BA"], "native_pairs_per_cell": NATIVE_PAIRS, "warmups": NATIVE_WARMUPS,
            "samples": NATIVE_SAMPLES, "bootstrap_seed": BOOTSTRAP_SEED,
            "bootstrap_resamples": BOOTSTRAP_RESAMPLES, "upper_ci_limit": CI_LIMIT,
            "before_candidate_formal_launch": True,
            "complete_noise_failure": "readiness_fail_uninterpretable",
            "missing_or_invalid": "incomplete_no_decision",
        },
        "benefit": {
            "requested_metric": "samples[i].allocation.allocated_bytes",
            "requested_endpoints": ["docnohf/e32/default", "docfloat/e32/default"],
            "requested_endpoint_median_upper": 0.95,
            "native_p50_median_upper": 0.95, "native_p50_ci_upper": 1.0,
        },
        "outputs": {
            "aa_root": str(ROOT / "capture/aa-v4"),
            "capture_root": str(ROOT / ("capture/aa-v4" if aa_mode else "capture/formal-v4")),
            "analysis_root": str(ROOT / "analysis"),
        },
    }
    if not aa_mode:
        plan["aa_readiness_receipt"] = maybe_file(str(args.aa_receipt) if args.aa_receipt else None)
        plan["aa_capture_manifest"] = maybe_file(str(args.aa_manifest) if args.aa_manifest else None)
    write_json_exclusive(args.output.resolve(), plan)
    print(json.dumps({"schema": PLAN_SCHEMA, "status": "emitted", "mode": args.mode, "plan": str(args.output.resolve())}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
