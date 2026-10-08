#!/usr/bin/env python3
"""Emit the immutable, no-GO 0967 matched-build/capture plan.

The generator only hashes existing inputs and writes JSON. It never invokes
Cargo, the control probe, or a performance workload. A later binding step must
fill the six executable SHA-256 values and the sealed control receipt before a
runner can launch anything.
"""
from __future__ import annotations

import argparse
import hashlib
import json
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

REVISION = "893072657ad8b1d726bfce351b5ec2e98f084e94"
PLAN_SCHEMA = "litchi.performance.0967.plan.v1"
QUALITY_ACCEPTED = {"passed", "passed_with_retained_baseline_failure"}
REAL_FIXTURES = {
    "docnohf": {
        "relative_path": "test-data/ole/doc/NoHeadFoot.doc",
        "bytes": 26112,
        "sha256": "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48",
    },
    "docfloat": {
        "relative_path": "test-data/ole/doc/FloatingPictures.doc",
        "bytes": 335360,
        "sha256": "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d",
    },
}
QUERY_FIXTURE = {
    "id": "real-bookmark-delete-redline-query",
    "relative_path": "test-data/libreoffice-core/sw/qa/core/doc/data/bookmark-delete-redline.doc",
    "bytes": 66048,
    "sha256": "f8c19b5712e12244af5e243ab6b996af6bc32193dad208948a155c4a4d5bb3dd",
    "revision_count": 20,
    "paragraph_zero_unmarked": False,
}
HARNESS_FILES = {
    "tools/perf-cfb-profile/Cargo.toml": "41dee12e4eb7d1e3853a33a7743afdf6b5d6519d5e7456e61ce4978aa9177945",
    "tools/perf-cfb-profile/Cargo.lock": "343c5464ae1144cbb72284cbb3e9b52b5dff2b798d09462d3c0d6a74b317bac7",
    "tools/perf-cfb-profile/src/lib.rs": "b4893321ac9a9321a9055823a7f05ecb35e59319531913affc49342b8ccd73dc",
    "tools/perf-cfb-profile/src/main.rs": "13206a810887dfcb4897a265167edc8fa3706ae736a2785a27354c7617eb5ca9",
    "tools/perf-cfb-profile/src/bin/alloc.rs": "2542257a9ab68347d3e8133b5f65a2254cc1e439fef665e327e093470f365f5f",
    "tools/perf-cfb-profile/src/oracle_0730.rs": "2b7a4d8e41c53690c8abf8d656e4b90942a7b7a0d84f205f93a1ac811f5f61e6",
}
CONTROL_RESULT = "harness-control/control-result-v1.md"
CONTROL_RESULT_SHA256 = "c566adbd1dbb6da18913dea89692088777260ce2ba29e9f7d9de437153e374a3"
CONTROL_PROFILE_ROOT = "/home/zhuhe/litchi-performance-0967/harness-perf"
CONTROL_PROFILE_DESIGN = {"design-v1.md": "f75386a27bedb50de295d6cf37099b1cf6be89c01167baac3c513869ec41b37a"}
CONTROL_PROFILE_SOURCE_RELATIVE = (
    "tools/perf-revision-control-0967/Cargo.toml",
    "tools/perf-revision-control-0967/Cargo.lock",
    "tools/perf-revision-control-0967/src/lib.rs",
    "tools/perf-revision-control-0967/src/main.rs",
    "tools/perf-revision-control-0967/src/bin/alloc.rs",
    "tools/perf-revision-control-0967/src/alloc_metrics.rs",
)
CONTROL_PROFILE_SOURCE_SHA256 = {
    "tools/perf-revision-control-0967/Cargo.toml": "e1672b09fa896d0d561b994243466a99ace27d6f93bcf2bd4719933dc036efde",
    "tools/perf-revision-control-0967/Cargo.lock": "0a274917e3f7767a7cac2bd62c10325afa186a57da7e05dd8b6e7a0f4de84223",
    "tools/perf-revision-control-0967/src/lib.rs": "47e77b319d74ea91fc6952ae9e99c2fd43c735710b7f7e10783b90d192265712",
    "tools/perf-revision-control-0967/src/main.rs": "63ac2dc52077466b30ac6ee4abcfbd59a88789d35731fcf5c7f5a96d2dc95978",
    "tools/perf-revision-control-0967/src/bin/alloc.rs": "9aa8e131e2810e83de5bff1964b8bdbb0cdd894b6865e26986119c6238c9df08",
    "tools/perf-revision-control-0967/src/alloc_metrics.rs": "048680324eff200dead3b37cabb97582b5f76b6e7ca572b7e2bc2c5bb3a7475d",
}
CONTROL_PROFILE_SOURCE_ROOTS = {
    "baseline": "/home/zhuhe/litchi-performance-0967/harness-control/source-baseline",
    "candidate": "/home/zhuhe/litchi-performance-0967/source-candidate",
}
CONTROL_PROFILE_TARGET_ROOTS = {
    "baseline": "/home/zhuhe/litchi-performance-0967/builds/control-target-baseline-revision-v1",
    "candidate": "/home/zhuhe/litchi-performance-0967/builds/control-target-candidate-revision-v1",
}
CONTROL_PROFILE_RECEIPT_DEFAULT = "/home/zhuhe/litchi-performance-0967/receipts/control-build-0967-v2.json"
CONTROL_PROFILE_RECEIPT_STATUSES = {"passed", "sealed", "matched_ordinary_native_and_allocator_control_builds_passed"}
CONTROL_SCOPE_REVIEW = {
    "path": "/home/zhuhe/litchi-performance-0967/review/control-harness-scope-0967-v1.json",
    "sha256": "141e58bbe04e11606452ec51dc516a096260bc102fa8014f5c919d6ea4b6b9bf",
    "status": "scope_pass_with_formal_binding_pending",
}
CONTROL_PROFILE_INPUTS = {
    "generated-positive": {"path": "/home/zhuhe/litchi-performance-0967/harness-perf/inputs/generated-positive.doc", "bytes": 15360, "sha256": "47e6f76319a8086ebc36719bc9f1d894ad907a02856cfb842ac198d7879d8737"},
    "dense1024": {"path": "/home/zhuhe/litchi-performance-0967/harness-perf/inputs/dense1024.doc", "bytes": 93696, "sha256": "f29dc1e3e1c05dcf37ce0843551a21ca9423cb283d54599813b21d52fcd76def"},
    "false-marker": {"path": "/home/zhuhe/litchi-performance-0967/harness-perf/inputs/false-marker.doc", "bytes": 15360, "sha256": "67cc08d9c17b3e8336327b51ecd78e6f44b14929037e877559b451918159f4ee"},
    "real-positive": {"path": "/home/zhuhe/litchi-performance-0967/harness-perf/inputs/real-positive.doc", "bytes": 66048, "sha256": "f8c19b5712e12244af5e243ab6b996af6bc32193dad208948a155c4a4d5bb3dd"},
}
CONTROL_BINARIES = {
    "control_baseline": {
        "path": "/home/zhuhe/litchi-performance-0967/harness-control/target-baseline/release/doc_revision_positive_control_0967",
        "sha256": "2227105e81dc43f8e128904818c846955e5d8d1e477611679ebb1246ca3cbe0d",
    },
    "control_candidate": {
        "path": "/home/zhuhe/litchi-performance-0967/harness-control/target-candidate/release/doc_revision_positive_control_0967",
        "sha256": "fed4f00380289030558e56ccd08bdd5a7d524c9a2f648e375cec12b9baa93709",
    },
}
FUNCTIONAL_REPORTS = {
    "generated-positive": {
        "baseline": "harness-control/receipts/generated-baseline-final.json",
        "candidate": "harness-control/receipts/generated-candidate-controls.json",
    },
    "generated-dense-1024": {
        "baseline": "harness-control/receipts/dense-1024-baseline-release.json",
        "candidate": "harness-control/receipts/dense-1024-candidate-release.json",
    },
    "generated-false": {
        "baseline": "harness-control/receipts/false-marker-baseline-release.json",
        "candidate": "harness-control/receipts/false-marker-candidate-release.json",
    },
    "real-bookmark-delete-redline-query": {
        "baseline": "harness-control/receipts/query-baseline-real-final.json",
        "candidate": "harness-control/receipts/query-candidate-real-controls.json",
    },
}
PRODUCTION = [
    "crates/litchi-doc/src/sprm.rs",
    "crates/litchi-doc/src/tracked_revision/codec.rs",
    "crates/litchi-doc/src/tracked_revision/package.rs",
]
TEST_ONLY = [
    "crates/litchi-doc/src/sprm.rs::frozen_original_reference",
    "crates/litchi-doc/src/tracked_revision/package.rs::revision_scan_tests",
]
RETENTIONS = ["zero", "default", "release"]
EDITS = [1, 2, 8, 32]


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def load(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise SystemExit(f"{path}: expected JSON object")
    return value


def fixture_rows(root: Path) -> dict[str, dict[str, Any]]:
    rows: dict[str, dict[str, Any]] = {}
    for case, expected in REAL_FIXTURES.items():
        path = root / expected["relative_path"]
        if not path.is_file():
            raise SystemExit(f"missing fixture: {path}")
        actual = sha(path)
        if path.stat().st_size != expected["bytes"] or actual != expected["sha256"]:
            raise SystemExit(f"fixture changed: {path}")
        rows[case] = {"id": case, "path": str(path), **expected}
    query_path = root / QUERY_FIXTURE["relative_path"]
    if not query_path.is_file() or query_path.stat().st_size != QUERY_FIXTURE["bytes"] or sha(query_path) != QUERY_FIXTURE["sha256"]:
        raise SystemExit(f"query fixture changed: {query_path}")
    rows[QUERY_FIXTURE["id"]] = {"path": str(query_path), **QUERY_FIXTURE}
    return rows


def harness_rows(root: Path) -> list[dict[str, Any]]:
    rows = []
    for rel, expected in HARNESS_FILES.items():
        path = root / rel
        if not path.is_file() or sha(path) != expected:
            raise SystemExit(f"harness binding mismatch: {path}")
        rows.append({"path": rel, "bytes": path.stat().st_size, "sha256": expected})
    return rows


def timed_source_rows(root: Path) -> list[dict[str, Any]]:
    rows = []
    for rel in CONTROL_PROFILE_SOURCE_RELATIVE:
        path = root / rel
        expected = CONTROL_PROFILE_SOURCE_SHA256[rel]
        if not path.is_file() or sha(path) != expected:
            raise SystemExit(f"timed control source binding mismatch: {path}")
        rows.append({"path": rel, "bytes": path.stat().st_size, "sha256": expected})
    return rows


def build_command(source_root: Path, target_dir: Path, package: str, bins: list[str]) -> dict[str, Any]:
    return {
        "package": package,
        "source_root": str(source_root),
        "target_dir": str(target_dir),
        "argv": [
            "cargo", "+1.95.0", "build", "--release", "--locked", "--offline", "-j2",
            "--manifest-path", str(source_root / "tools/perf-cfb-profile/Cargo.toml"),
            "--target-dir", str(target_dir),
            *sum((["--bin", name] for name in bins), []),
        ],
        "expected_binaries": [str(target_dir / "release" / name) for name in bins],
        "env": {"CARGO_INCREMENTAL": "0", "CARGO_TERM_COLOR": "never", "RUSTFLAGS": ""},
        "launched": False,
    }


def control_profile_command(source_root: Path, target_dir: Path, variant: str, bins: list[str]) -> dict[str, Any]:
    overlay_root = source_root / "tools/perf-revision-control-0967"
    return {
        "package": "doc-revision-control-profile-0967",
        "source_root": str(source_root),
        "control_source": str(overlay_root),
        "source_active": str(overlay_root),
        "source_active_expected_target": str(overlay_root),
        "target_dir": str(target_dir),
        "variant": variant,
        "argv": [
            "cargo", "+1.95.0", "build", "--release", "--locked", "--offline", "-j2",
            "--manifest-path", str(overlay_root / "Cargo.toml"), "--target-dir", str(target_dir),
            *sum((["--bin", name] for name in bins), []),
        ],
        "expected_binaries": [str(target_dir / "release" / name) for name in bins],
        "env": {"CARGO_INCREMENTAL": "0", "CARGO_TERM_COLOR": "never", "RUSTFLAGS": ""},
        "source_binding_required": True,
        "launched": False,
    }


def control_command(source_root: Path, target_dir: Path, variant: str) -> dict[str, Any]:
    return {
        "package": "doc_revision_positive_control_0967",
        "source_root": str(source_root),
        "control_source": "/home/zhuhe/litchi-performance-0967/harness-control/probe",
        "source_active": str(source_root),
        "target_dir": str(target_dir),
        "variant": variant,
        "argv": [
            "cargo", "+1.95.0", "build", "--release", "--locked", "--offline", "-j2",
            "--manifest-path", "/home/zhuhe/litchi-performance-0967/harness-control/probe/Cargo.toml",
            "--target-dir", str(target_dir),
            "--bin", "doc_revision_positive_control_0967",
        ],
        "expected_binary": str(target_dir / "release" / "doc_revision_positive_control_0967"),
        "source_binding_required": True,
        "launched": False,
    }


def child_schedule(pairs: int) -> list[dict[str, Any]]:
    return [
        {"pair": pair, "relative_order": "baseline_then_candidate" if pair <= pairs // 2 else "candidate_then_baseline"}
        for pair in range(1, pairs + 1)
    ]


def real_cells() -> list[dict[str, Any]]:
    return [
        {"cell_id": f"{case}/e{edits}/{retention}", "case": case, "edits": edits, "retention": retention}
        for case in REAL_FIXTURES
        for edits in EDITS
        for retention in RETENTIONS
    ]


def control_matrix() -> list[dict[str, Any]]:
    # This is fixed before capture. The control probe is correctness-only; the
    # timed scope is separately bound to a timing-capable control harness.
    return [
        {
            "id": "generated-positive",
            "probe_mode": "run-generated",
            "fixture_pattern": "generated-positive",
            "edits": 1,
            "retention": "default",
            "timed_scope": False,
            "require_revision_survival": True,
        },
        {
            "id": "generated-dense-1024",
            "probe_mode": "run-dense",
            "runs": 1024,
            "fixture_pattern": "generated-dense-1024",
            "edits": 32,
            "retention": "default",
            "timed_scope": False,
            "require_revision_survival": True,
        },
        {
            "id": "generated-false",
            "probe_mode": "run-false-marker",
            "fixture_pattern": "generated-false-marker",
            "edits": 1,
            "retention": "default",
            "timed_scope": False,
            "require_revision_survival": False,
            "require_false_marker": True,
        },
        {
            "id": QUERY_FIXTURE["id"],
            "probe_mode": "query-input",
            "input_fixture": QUERY_FIXTURE["relative_path"],
            "timed_scope": False,
            "query_reopen": True,
        },
    ]


def main() -> int:
    ap = argparse.ArgumentParser()
    base = Path("/home/zhuhe/litchi-performance-0967")
    ap.add_argument("--baseline-root", type=Path, default=base / "harness-control/source-baseline")
    ap.add_argument("--candidate-root", type=Path, default=base / "source-candidate")
    ap.add_argument("--source-manifest", type=Path, default=base / "receipts/source-candidate-manifest-v2.json")
    ap.add_argument("--baseline-manifest", type=Path, default=base / "receipts/source-baseline-manifest-v2.json")
    ap.add_argument("--quality", type=Path, default=base / "receipts/quality-0967-coder-v1.json")
    ap.add_argument("--semantic", type=Path, default=base / "review/semantic-review-0967-v1.json")
    ap.add_argument("--control-receipt", type=Path)
    ap.add_argument("--control-profile-receipt", type=Path, default=Path(CONTROL_PROFILE_RECEIPT_DEFAULT))
    ap.add_argument("--output", type=Path, default=base / "plans/plan-0967-v1.json")
    args = ap.parse_args()
    baseline = args.baseline_root.resolve()
    candidate = args.candidate_root.resolve()
    source_manifest = args.source_manifest.resolve()
    baseline_manifest = args.baseline_manifest.resolve()
    quality_path = args.quality.resolve()
    semantic_path = args.semantic.resolve()
    manifest = load(source_manifest)
    quality = load(quality_path)
    semantic = load(semantic_path)
    baseline_manifest_doc = load(baseline_manifest)
    if manifest.get("source_revision") != REVISION or quality.get("source_commit") != REVISION or semantic.get("source", {}).get("baseline_commit") != REVISION:
        raise SystemExit("source revision mismatch")
    if manifest.get("status") != "sealed_for_matched_build":
        raise SystemExit("candidate source manifest is not sealed")
    if quality.get("status") not in QUALITY_ACCEPTED or semantic.get("status") not in QUALITY_ACCEPTED:
        raise SystemExit("quality or semantic receipt is not accepted")
    if not baseline_manifest_doc.get("source_path"):
        raise SystemExit("baseline manifest lacks source_path")
    if not baseline.is_dir() or not candidate.is_dir():
        raise SystemExit("source roots unavailable")
    fixtures = fixture_rows(candidate)
    # Verify baseline copies too, so a later build cannot silently use a drifted tree.
    fixture_rows(baseline)
    harness_candidate = harness_rows(candidate)
    harness_baseline = harness_rows(baseline)
    timed_source_candidate = timed_source_rows(candidate)
    timed_source_baseline = timed_source_rows(baseline)
    if timed_source_candidate != timed_source_baseline:
        raise SystemExit("timed control source overlay differs between baseline and candidate")
    functional_reports: dict[str, dict[str, dict[str, Any]]] = {}
    for control_id, variants in FUNCTIONAL_REPORTS.items():
        functional_reports[control_id] = {}
        for variant, rel in variants.items():
            path = base / rel
            if not path.is_file():
                raise SystemExit(f"missing sealed functional control report: {path}")
            functional_reports[control_id][variant] = {"path": str(path), "sha256": sha(path)}
    control_profile_files = []
    for rel, expected in CONTROL_PROFILE_DESIGN.items():
        path = Path(CONTROL_PROFILE_ROOT) / rel
        if not path.is_file() or sha(path) != expected:
            raise SystemExit(f"control profile binding mismatch: {path}")
        control_profile_files.append({"path": rel, "bytes": path.stat().st_size, "sha256": expected})
    control_design_text = (Path(CONTROL_PROFILE_ROOT) / "design-v1.md").read_text(encoding="utf-8")
    control_design_aligned = all(token in control_design_text for token in ("generated-positive", "dense1024", "false-marker", "real-positive"))
    control_inputs = {}
    for control_id, row in CONTROL_PROFILE_INPUTS.items():
        path = Path(row["path"])
        if not path.is_file() or path.stat().st_size != row["bytes"] or sha(path) != row["sha256"]:
            if control_id != "generated-positive":
                raise SystemExit(f"control profile input binding mismatch: {path}")
            control_inputs[control_id] = {**row, "available": False}
        else:
            control_inputs[control_id] = {**row, "available": True}
    now = datetime.now(timezone.utc).isoformat()
    build_root = base / "builds"
    control_profile_receipt_path = args.control_profile_receipt.resolve()
    control_profile_receipt = None
    if control_profile_receipt_path.is_file():
        receipt = load(control_profile_receipt_path)
        if receipt.get("status") not in CONTROL_PROFILE_RECEIPT_STATUSES:
            raise SystemExit("control profile build receipt is not sealed/passed")
        control_profile_receipt = {"path": str(control_profile_receipt_path), "sha256": sha(control_profile_receipt_path), "status": receipt.get("status")}
    commands = [
        build_command(baseline, build_root / "target-baseline-cfb", "doc-retention-probe-0942-cfb-profile", ["doc_retention_probe_0942_cfb_profile", "doc_retention_probe_0942_cfb_profile_alloc"]),
        build_command(candidate, build_root / "target-candidate-cfb", "doc-retention-probe-0942-cfb-profile", ["doc_retention_probe_0942_cfb_profile", "doc_retention_probe_0942_cfb_profile_alloc"]),
    ]
    control_receipt = None
    if args.control_receipt:
        receipt_path = args.control_receipt.resolve()
        control_receipt = {"path": str(receipt_path), "sha256": sha(receipt_path)}
        if receipt_path.suffix == ".json":
            document = load(receipt_path)
            if document.get("status") not in {"passed", "sealed"}:
                raise SystemExit("control receipt is not sealed/passed")
            control_receipt["status"] = document.get("status")
        else:
            text = receipt_path.read_text(encoding="utf-8")
            if "Status: functional control qualification complete" not in text:
                raise SystemExit("control markdown receipt is not sealed")
            control_receipt["status"] = "passed"
    else:
        default_receipt = base / CONTROL_RESULT
        if default_receipt.is_file() and sha(default_receipt) == CONTROL_RESULT_SHA256:
            control_receipt = {"path": str(default_receipt), "sha256": CONTROL_RESULT_SHA256, "status": "passed"}
    controls_status = "bound" if control_receipt else "pending_sealed_control_receipt"
    plan = {
        "schema": PLAN_SCHEMA,
        "version": 1,
        "status": "pending_build_and_capture_authorization",
        "created_utc": now,
        "timing_claim": False,
        "protocol_document": {"path": str(base / "protocol-design-v1.md"), "sha256": sha(base / "protocol-design-v1.md")},
        "build_started": False,
        "workload_started": False,
        "source": {
            "revision": REVISION,
            "baseline": {"root": str(baseline), "manifest": str(baseline_manifest), "manifest_sha256": sha(baseline_manifest), "status": baseline_manifest_doc.get("status")},
            "candidate": {"root": str(candidate), "manifest": str(source_manifest), "manifest_sha256": sha(source_manifest), "status": manifest.get("status")},
            "quality_receipt": {"path": str(quality_path), "sha256": sha(quality_path), "status": quality.get("status")},
            "semantic_review": {"path": str(semantic_path), "sha256": sha(semantic_path), "status": semantic.get("status")},
            "reproduction_receipt": manifest.get("reproduction_receipt"),
            "production_allowlist": PRODUCTION,
            "test_only_allowlist": TEST_ONLY,
            "unrelated_change_policy": "reject",
        },
        "harness": {
            "relative_root": "tools/perf-cfb-profile",
            "candidate_files": harness_candidate,
            "baseline_files": harness_baseline,
            "timed_overlay_relative_root": "tools/perf-revision-control-0967",
            "candidate_timed_files": timed_source_candidate,
            "baseline_timed_files": timed_source_baseline,
            "source_equivalence_required": True,
            "timer_wrapper": "original 0942 perf-cfb-profile; no added frame pointers or candidate-specific instrumentation",
        },
        "build_policy": {
            "toolchain": "+1.95.0",
            "locked": True,
            "offline": True,
            "cargo_jobs": 2,
            "concurrency": 1,
            "no_cargo_parallelism": "one Cargo process at a time; -j2 is the inherited matched-build setting",
            "rustflags": "",
            "commands": commands,
            "control_commands": [],
            "control_build_note": "functional control binaries are retained and bound by control-result-v1.md; timed controls require a separate reviewed harness",
            "control_profile_commands": [
                control_profile_command(baseline, Path(CONTROL_PROFILE_TARGET_ROOTS["baseline"]), "baseline", ["doc_revision_control_profile_0967", "doc_revision_control_profile_0967_alloc"]),
                control_profile_command(candidate, Path(CONTROL_PROFILE_TARGET_ROOTS["candidate"]), "candidate", ["doc_revision_control_profile_0967", "doc_revision_control_profile_0967_alloc"]),
            ],
            "control_profile_receipt": control_profile_receipt,
        },
        "host_state": {
            "status": "pending_root_before_capture",
            "before": None,
            "after": None,
            "required_before_GO": True,
            "note": "0967 does not impose CPU affinity; root binds host/lscpu/load/memory/ASLR state immediately before and after capture.",
        },
        "binaries": {
            "native_baseline": {"path": str(build_root / "target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile"), "sha256": None, "kind": "ordinary_uninstrumented"},
            "native_candidate": {"path": str(build_root / "target-candidate-cfb/release/doc_retention_probe_0942_cfb_profile"), "sha256": None, "kind": "ordinary_uninstrumented"},
            "allocator_baseline": {"path": str(build_root / "target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile_alloc"), "sha256": None, "kind": "counting_allocator"},
            "allocator_candidate": {"path": str(build_root / "target-candidate-cfb/release/doc_retention_probe_0942_cfb_profile_alloc"), "sha256": None, "kind": "counting_allocator"},
            "control_baseline": {**CONTROL_BINARIES["control_baseline"], "kind": "correctness_control"},
            "control_candidate": {**CONTROL_BINARIES["control_candidate"], "kind": "correctness_control"},
            "timed_control_native_baseline": {"path": str(Path(CONTROL_PROFILE_TARGET_ROOTS["baseline"]) / "release/doc_revision_control_profile_0967"), "sha256": None, "kind": "control_profile_native"},
            "timed_control_native_candidate": {"path": str(Path(CONTROL_PROFILE_TARGET_ROOTS["candidate"]) / "release/doc_revision_control_profile_0967"), "sha256": None, "kind": "control_profile_native"},
            "timed_control_allocator_baseline": {"path": str(Path(CONTROL_PROFILE_TARGET_ROOTS["baseline"]) / "release/doc_revision_control_profile_0967_alloc"), "sha256": None, "kind": "control_profile_allocator"},
            "timed_control_allocator_candidate": {"path": str(Path(CONTROL_PROFILE_TARGET_ROOTS["candidate"]) / "release/doc_revision_control_profile_0967_alloc"), "sha256": None, "kind": "control_profile_allocator"},
        },
        "fixtures": fixtures,
        "real_matrix": {
            "cells": real_cells(),
            "native": {"pairs_per_cell": 12, "warmups": 5, "samples": 80, "relative_ratio": "candidate_over_baseline", "schedule": child_schedule(12), "bootstrap_seed": 967001, "bootstrap_resamples": 10000, "confidence": 0.95, "p99_descriptive": {"ratio_median_field": "p99_ratio_median", "flag_field": "p99_descriptive_flag", "flag_if_median_gt": 1.05, "report_only": True}},
            "allocator": {"pairs_per_cell": 3, "warmups": 0, "samples": 1, "relative_ratio": "candidate_over_baseline", "schedule": child_schedule(3), "bootstrap_seed": 967002, "bootstrap_resamples": 10000},
        },
        "controls": {
            "status": controls_status,
            "correctness_probe_schema": "litchi.performance.0967.doc_revision_positive_control.v1",
            "probe_receipt": control_receipt,
            "functional_reports": functional_reports,
            "matrix": control_matrix(),
            "required_variants": ["baseline", "candidate"],
            "scenario_audit": {"path": str(base / "tools/scenario_audit_0967_v1.py"), "sha256": sha(base / "tools/scenario_audit_0967_v1.py"), "output": str(base / "analysis/scenario-audit-v1.json"), "status_required": "pass"},
            "timed_scope": {
                "status": "bound" if control_design_aligned and control_profile_receipt else "pending_design_alignment_or_build_receipt",
                "required": True,
                "schema": "litchi.performance.0967.control_profile.v1",
                "harness_root": CONTROL_PROFILE_ROOT,
                "harness_files": control_profile_files,
                "source_overlay_relative_root": "tools/perf-revision-control-0967",
                "native_timing_claim": True,
                "allocator_timing_claim": False,
                "allocator_metric_claim": True,
                "scope_review": CONTROL_SCOPE_REVIEW,
                "inputs": control_inputs,
                "cells": [
                    {"id": "generated-positive-e1-default", "control": "generated-positive", "scope": "edit", "edits": 1, "retention": "default"},
                    {"id": "dense1024-e32-default", "control": "dense1024", "scope": "edit", "edits": 32, "retention": "default"},
                    {"id": "false-marker-e1-default", "control": "false-marker", "scope": "edit", "edits": 1, "retention": "default"},
                    {"id": "false-marker-e32-default", "control": "false-marker", "scope": "edit", "edits": 32, "retention": "default"},
                ],
                "native_pairs_per_control": 12,
                "warmups": 5,
                "samples": 80,
                "allocator_pairs_per_control": 3,
                "allocator_warmups": 0,
                "allocator_samples": 1,
                "bootstrap_seed": 967003,
                "bootstrap_resamples": 10000,
                "p99_descriptive": {"ratio_median_field": "p99_ratio_median", "flag_field": "p99_descriptive_flag", "flag_if_median_gt": 1.05, "report_only": True},
                "gate": "same p50/p95 noninferiority upper CI <= 1.05; allocator memory safeguards apply",
                "note": "The correctness probe's timing_claim is false. This separate profile harness owns the timed control schema: native timing_claim=true, allocator timing_claim=false. Executable hashes remain unbound until binding.",
                "design_alignment": control_design_aligned,
            },
        },
        "tools": {
            "runner": {"path": str(base / "tools/run_0967_v1.py"), "sha256": sha(base / "tools/run_0967_v1.py")},
            "reducer": {"path": str(base / "tools/reduce_0967_v1.py"), "sha256": sha(base / "tools/reduce_0967_v1.py")},
            "scenario_audit": {"path": str(base / "tools/scenario_audit_0967_v1.py"), "sha256": sha(base / "tools/scenario_audit_0967_v1.py")},
            "source_binder": {"path": str(base / "tools/bind_source_manifest_0967_v2.py"), "sha256": sha(base / "tools/bind_source_manifest_0967_v2.py")},
        },
        "aa_noise": {
            "coverage": "endpoint_cells_only",
            "cells": ["docnohf/e32/default", "docfloat/e32/default"],
            "same_binary": "baseline",
            "order": ["AB", "BA"],
            "native_pairs_per_cell": 12,
            "warmups": 5,
            "samples": 80,
            "bootstrap_seed": 967004,
            "bootstrap_resamples": 10000,
            "symmetric_ratio": "max(r,1/r)",
            "upper_ci_limit": 1.05,
            "interpretation": "readiness only; never pooled into candidate ratios or used to qualify/reject",
        },
        "gates": {
            "native_noninferiority": "every real cell p50 and p95 paired-bootstrap upper CI <= 1.05",
            "native_review_flag": "any real-cell p50 median > 1.05 is reported for review even when CI passes",
            "native_p99_descriptive": "report p99_ratio_median and p99_descriptive_flag when median > 1.05; descriptive only, never a gate or waiver",
            "allocation_branch": {"docfloat/e32/default": 0.90, "docnohf/e32/default": 0.97, "metric": "paired median allocation_calls candidate_over_baseline"},
            "native_benefit_branch": "at least one endpoint p50 median <= 0.95 and 95% bootstrap upper CI < 1.00",
            "memory": {"requested_or_allocated_bytes": 1.01, "peak_live_bytes": 1.05, "process_rss": 1.05},
            "other_allocation_calls_median": 1.01,
            "controls": "all baseline/candidate positive, dense, false reports pass scenario audit; timed control scope is bound and passes noninferiority",
            "rejection": "any custody/correctness/missing/duplicate/terminal/metric mismatch, failed real-cell CI, guard violation, absent controls, unbounded positive double walk, or no benefit branch",
            "aa_failure": "comparison uninterpretable; retain packets and make no adoption/rejection decision",
            "no_retry": True,
        },
        "outputs": {
            "capture_root": str(base / "capture"),
            "reduction": str(base / "analysis/reduction-v1.json"),
            "selection": str(base / "analysis/selection-v1.json"),
            "raw_packets_retained": True,
        },
        "historical_note": "Quality and semantic receipts retain an unrelated baseline rustdoc failure; no candidate-file diagnostic was reported. This plan makes no adoption claim until builds, controls, and measurements pass.",
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(plan, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": plan["status"], "control_status": controls_status, "output": str(args.output), "sha256": sha(args.output)}))
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
