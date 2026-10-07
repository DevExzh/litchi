#!/usr/bin/env python3
"""Build a fresh epoch-2 0964 prospective/final screen plan without running work.

The local contract fragment contains only the frozen three-case oracles and
serial schedule.  No historical plan object, schema, helper path, or candidate
custody descriptor is copied into the emitted plan.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
import re
import stat
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0964")
REFERENCE = ROOT / "protocol/inherited-contract-0964-v1.json"
RETAINED_ROOT = Path("/home/zhuhe/litchi-performance-0957")
PRODUCTION_ALLOWLIST = [
    "crates/litchi-xlsx/src/cell_values/validation.rs",
    "crates/litchi-xlsx/src/raw/worksheet/codec.rs",
]
PERF_LOCK_SHA256 = "db62bdd9174484613acae1f49f8187cf105d0b4976e2fea4de91df7372a893d5"
MANAGED_SHA256 = "5886b0f86a697bae757994d5a757cffae4a7e8d61a2769ff1c967029cd1912c3"
ALLOCATOR_SHA256 = "9ba37a6b9796ad0e316e7571ebe117ba7a5f97bcf5d29153c7cb3d74c485e9a5"
CPU_MASK = [12, 13, 14, 15]
FIXED_ENV = {"PATH": "/usr/bin:/bin", "LANG": "C", "LC_ALL": "C", "RUST_BACKTRACE": "0", "TZ": "UTC"}
NATIVE_BLOCKS = [["A", "B", "B", "A"], ["B", "A", "A", "B"], ["A", "B", "B", "A"]]
ALLOCATOR_BLOCKS = [["A", "B"], ["B", "A"], ["A", "B"]]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, executable: bool | None = None) -> dict[str, Any]:
    path = path.expanduser().resolve()
    if not path.is_file() or path.is_symlink():
        raise RuntimeError(f"regular file required: {path}")
    mode = stat.S_IMODE(path.stat().st_mode)
    actual = bool(mode & 0o111)
    if executable is not None and actual is not executable:
        raise RuntimeError(f"executable mismatch: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "mode": mode,
            "executable": actual, "sha256": sha256(path)}


def maybe_descriptor(path: Path, executable: bool | None = None) -> dict[str, Any]:
    if not path.exists() or path.is_symlink():
        return {"status": "pending_root_binding", "path": str(path.resolve())}
    return descriptor(path, executable)


def read_test_paths(path: Path | None) -> list[str]:
    if path is None:
        return []
    if not path.is_file() or path.is_symlink():
        raise RuntimeError(f"test allowlist file missing: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        value = [line.strip() for line in path.read_text(encoding="utf-8").splitlines() if line.strip()]
    if isinstance(value, dict):
        value = value.get("paths")
    if not isinstance(value, list) or not all(isinstance(item, str) and item for item in value):
        raise RuntimeError("test allowlist must be a JSON list or one relative path per line")
    paths = sorted(set(value))
    if any(path.startswith("/") or ".." in Path(path).parts for path in paths):
        raise RuntimeError("test allowlist paths must be relative")
    if any(path in PRODUCTION_ALLOWLIST for path in paths):
        raise RuntimeError("production path repeated in test allowlist")
    return paths


def pending_or(path: Path | None, executable: bool | None = None) -> dict[str, Any]:
    return {"status": "pending_root_binding", "path": str(path.resolve())} if path is None else maybe_descriptor(path, executable)


def required_final(value: dict[str, Any], label: str) -> None:
    if value.get("status") == "pending_root_binding":
        raise RuntimeError(f"--final requires {label}: {value.get('path')}")


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path)
    parser.add_argument("--final", action="store_true", help="emit plan-final-v1.json and require candidate custody inputs")
    parser.add_argument("--candidate-manifest", type=Path)
    parser.add_argument("--baseline-manifest", type=Path)
    parser.add_argument("--baseline-receipt", type=Path)
    parser.add_argument("--baseline-runtime-manifest", type=Path)
    parser.add_argument("--source-start", type=Path)
    parser.add_argument("--mode-correction", type=Path)
    parser.add_argument("--source-export", type=Path)
    parser.add_argument("--candidate-freeze", type=Path)
    parser.add_argument("--candidate-build-receipt", type=Path)
    parser.add_argument("--candidate-build-tool", type=Path)
    parser.add_argument("--quality-receipt", type=Path)
    parser.add_argument("--candidate-benchmark", type=Path)
    parser.add_argument("--candidate-allocator", type=Path)
    parser.add_argument("--test-allowlist-file", type=Path)
    parser.add_argument("--host", type=Path)
    parser.add_argument("--root-go", type=Path)
    args = parser.parse_args()

    if args.output is None:
        args.output = ROOT / "protocol" / ("plan-epoch2-0964-final-v1.json" if args.final else "plan-epoch2-0964-v1.json")
    if args.output.exists() or args.output.is_symlink():
        raise RuntimeError(f"refusing to overwrite plan: {args.output}")
    reference = json.loads(REFERENCE.read_text(encoding="utf-8"))
    default_build_tool = ROOT / "tools/build_candidate_managed_v2.py"
    if not default_build_tool.is_file():
        default_build_tool = ROOT / "tools/build_candidate_managed_v1.py"
    build_tool_path = args.candidate_build_tool or default_build_tool
    revision_match = re.search(r"_v(\d+)\.py$", build_tool_path.name)
    build_revision = f"v{revision_match.group(1)}" if revision_match else "v1"
    test_paths = read_test_paths(args.test_allowlist_file)
    allowlist = PRODUCTION_ALLOWLIST + test_paths

    # These schedule and oracle fragments are retained because they are the
    # deterministic three-case XLSX contract.  Every path-bearing value is
    # rebound below; no historical descriptor is retained.
    cases = copy.deepcopy(reference["cases"])
    comparison = copy.deepcopy(reference["comparison"])
    comparison.update({
        "role_A": "retained 0957 baseline executable",
        "role_B": "0964 raw-reader and NamespaceResolver transport candidate executable",
        "id": "baseline-vs-0964-candidate",
        "screening_only": True,
        "full_followup_required": True,
    })
    report_contract = copy.deepcopy(reference["report_contract"])
    lanes = copy.deepcopy(reference["lanes"])
    for lane in ("native", "allocator"):
        lanes[lane]["baseline_binary"] = descriptor(
            RETAINED_ROOT / "executables" / (
                "xlsx_managed_compression_benchmark.v2" if lane == "native" else "xlsx_managed_compression_alloc.v2"),
            True,
        )
        candidate_default = ROOT / "executables" / (
            f"xlsx_managed_compression_benchmark.{build_revision}" if lane == "native" else
            f"xlsx_managed_compression_alloc.{build_revision}")
        candidate_arg = args.candidate_benchmark if lane == "native" else args.candidate_allocator
        lanes[lane]["candidate_binary"] = pending_or(candidate_arg or candidate_default, True)
        lanes[lane]["candidate_receipt"] = pending_or(args.candidate_build_receipt or ROOT / f"receipts/candidate-managed-{build_revision}-build-result.json")
        lanes[lane]["execution"]["workload_cwd"] = str(ROOT / "source-baseline/tools/perf-baseline")
        lanes[lane]["execution"]["environment"]["approval_variable"] = "LITCHI_PERF_0964_CAPTURE_APPROVED"
    comparison["blocks"] = [{"id": index, "order": order} for index, order in enumerate(NATIVE_BLOCKS)]
    comparison["allocator_blocks"] = [{"id": index, "order": order} for index, order in enumerate(ALLOCATOR_BLOCKS)]

    baseline_manifest_path = args.baseline_manifest or ROOT / "source-manifest-baseline-v1.json"
    candidate_manifest_path = args.candidate_manifest or ROOT / "source-manifest-candidate-v1.json"
    baseline_receipt_path = args.baseline_receipt or ROOT / "receipts/baseline-binding-v1.json"
    candidate_freeze_path = args.candidate_freeze or ROOT / f"receipts/candidate-managed-{build_revision}-build-start.json"
    source_start_path = args.source_start or ROOT / "source-start.json"
    test_allowlist_path = args.test_allowlist_file or ROOT / "protocol/candidate-test-allowlist-v1.json"
    source = {
        "base_commit": json.loads(source_start_path.read_text(encoding="utf-8")).get("head") if source_start_path.is_file() else "pending_root_binding",
        "baseline_source_root": str(ROOT / "source-baseline"),
        "baseline_runtime_source_root": str(RETAINED_ROOT / "source-v1"),
        "candidate_source_root": str(ROOT / "source-candidate"),
        "candidate_runtime_source_root": str(ROOT / "source-candidate"),
        "retained_0957_source_root": str(RETAINED_ROOT / "source-v1"),
        "baseline_manifest": maybe_descriptor(baseline_manifest_path),
        "candidate_manifest": maybe_descriptor(candidate_manifest_path),
        "baseline_receipt": maybe_descriptor(baseline_receipt_path),
        "candidate_freeze": maybe_descriptor(candidate_freeze_path),
        "baseline_runtime_manifest": maybe_descriptor(args.baseline_runtime_manifest or RETAINED_ROOT / "source-manifest-v2.json"),
        "source_start": maybe_descriptor(args.source_start or ROOT / "source-start.json"),
        "mode_correction": maybe_descriptor(args.mode_correction or ROOT / "source-mode-correction.json"),
        "source_export": maybe_descriptor(args.source_export or ROOT / "receipts/source-export-v1.json"),
        "candidate_change_allowlist": allowlist,
        "candidate_allowlist_file": maybe_descriptor(test_allowlist_path),
        "required_perf_baseline_lock_sha256": PERF_LOCK_SHA256,
        "root_lock_overlay_is_not_runtime_binding": True,
        "source_cache_state": "fresh process and in-memory OwnedSource; no cold filesystem claim",
    }

    candidate_change = {
        "description": "Replace NsReader transport in validation.rs and raw worksheet codec with Reader events plus borrowed NamespaceResolver resolution; pending-scope pops and all existing validation remain authoritative.",
        "expected_allowlist": allowlist,
        "production_allowlist": PRODUCTION_ALLOWLIST,
        "test_only_allowlist": test_paths,
        "test_only_allowlist_is_explicit": True,
        "validation_preserved": True,
        "authoritative_validation_and_fallback_unchanged": True,
        "namespace_scope_and_declaration_order_preserved": True,
        "borrowed_event_resolution_preserved": True,
        "mce_attribute_text_eof_and_error_order_preserved": True,
        "cancellation_limits_and_resource_cleanup_preserved": True,
        "public_api_or_dependency_change": False,
        "unsafe_or_executor_change": False,
        "production_adoption": False,
    }
    gates = {
        "adoption_allowed": False,
        "allocation_claim": False,
        "screening_only": True,
        "full_followup_required": True,
        "signal_threshold": 0.95,
        "safety_review_threshold": 1.05,
        "useful_signal": "native edit_commit_ns.p50/p95, total_ns.p50/p95, native max_rss_kib, or allocator region_peak_live_bytes median <=0.95",
        "requested_bytes_or_allocation_calls_alone_signal": False,
        "hard_safety": "native primary, native phase p99, native max_rss_kib, allocator region_peak_live_bytes, or allocator max_rss_kib median >1.05",
        "invalid_metric_or_oracle": "reject packet and stop without retry",
        "tail_resource_guards": [
            "all required p99 values finite and positive",
            "native resources_after_drop all zero",
            "allocator phase live boundaries and combined conservation exact",
            "allocator resources_after_core all zero and failed allocations zero",
        ],
        "confirmation_adoption_gate": "fresh confirmation plus named primary <=0.95, bootstrap CI upper <1.0, no safety stop, and all ordinary/fallback/differential/oracle controls pass",
    }
    confirmation = {
        "when": "only after qualifying screen signal and every safety, source, terminal, control, and oracle gate passes",
        "native": "fresh six ABBA/BAAB blocks, 20 warmups, 200 samples per child",
        "allocator": "fresh three AB/BA repeats, 2 warmups, 10 samples per child",
        "fresh_bindings_required": True,
        "screening_result_not_adoption_evidence": True,
        "additional_controls_before_adoption": [
            "ordinary full XML validation path",
            "fallback path and malformed-input/error-order differential",
            "namespace/MCE/BOM/cap-limit reader differential coverage",
        ],
    }
    controls = {
        "timed": False,
        "must_pass": True,
        "groups": [
            "ordinary full validation and source admission",
            "fallback and malformed-input/error-order behavior",
            "namespace rebinding/default/prefix and MCE behavior",
            "BOM, EOF/closure, depth/input limits, cancellation and cleanup",
            "measured raw worksheet codec path and deterministic output/inverse",
        ],
        "candidate_specific_tests": "test-only paths are explicitly bound in candidate_allowlist_file and candidate_change.test_only_allowlist",
        "oracles": ["output", "semantic", "source", "inverse", "resource counters", "sink publication", "error class/order"],
        "scope": "Existing deterministic reader differential and full XLSX quality gates; no separate unbuilt control harness.",
    }

    helper_paths = {
        "capture": ROOT / "tools/screen_0964_v2.py",
        "reducer": ROOT / "tools/reduce_0964_v1.py",
        "controls": ROOT / "tools/check_selector_0964_v1.py",
        "selector": ROOT / "tools/select_screen_0964_v1.py",
        "selector_controls": ROOT / "tools/check_selector_0964_v1.py",
        "rule": ROOT / "protocol/root-screening-rule-0964-v1.json",
    }
    helpers = {key: descriptor(path, True if key == "capture" else None) for key, path in helper_paths.items()}
    affinity_controls_path = ROOT / "protocol/affinity-controls-epoch2-v3.json"
    helpers["affinity_controls"] = maybe_descriptor(affinity_controls_path)

    plan = {
        "schema": "litchi.performance.0964.inline-number-screen-plan.v1",
        "protocol_schema": "litchi.performance.0964.xlsx-reader-namespace-transport-screen-plan.v1",
        "plan_id": "0964-xlsx-reader-namespace-transport-screen-epoch2",
        "plan_revision": "epoch2-prospective-v1" if not args.final else "epoch2-root-final-v1",
        "status": "authorized" if args.root_go else "prospective",
        "measurements_started": False,
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "purpose": "Prospective descriptive screen of a bounded XLSX NsReader-to-Reader/NamespaceResolver transport change; screen cannot authorize adoption.",
        "status_note": "No build or workload is started by this generator; candidate custody and root GO remain pending." if not args.final else "Fresh candidate custody and root bindings supplied; measurements remain false until capture authorization.",
        "analysis_root": str(ROOT / "analysis-screen-epoch2-v1"),
        "capture_root": {"native": str(ROOT / "capture-native-epoch2-v1"), "allocator": str(ROOT / "capture-allocator-epoch2-v1")},
        "capture_authorization": "root_GO_required",
        "cpu_mask": CPU_MASK,
        "deadline_seconds": 900,
        "workers": 1,
        "environment": FIXED_ENV,
        "execution": {
            "child_owns_output_directory": True,
            "fixed_affinity": True,
            "fresh_process_per_child": True,
            "no_retries": True,
            "output_absent_before_launch": True,
            "process_group": "start_new_session; bounded SIGTERM then SIGKILL; reap and verify clear",
            "serial": True,
            "stop_on_first_failure": True,
            "terminal_before_parse": True,
            "terminal_before_report_parse": True,
        },
        "scope": {
            "format": "XLSX managed-compression benchmark",
            "candidate_mechanism": "raw Reader plus borrowed NamespaceResolver transport in two measured worksheet loops",
            "cases": [case["case_id"] for case in cases],
            "native_schedule": "3 blocks ABBA/BAAB/ABBA, 10 warmups, 60 measured samples, 36 serial fresh-process children",
            "allocator_schedule": "3 paired AB/BA/AB repeats, 2 warmups, 10 measured rows, 18 serial fresh-process children",
            "capture_epoch": "epoch2-affinity-hardened",
            "fresh_epoch2_schedule_required": True,
            "prior_epoch_capture_reusable": False,
            "screening_only": True,
            "iwork_excluded": True,
        },
        "cases": cases,
        "comparison": comparison,
        "lanes": lanes,
        "report_contract": report_contract,
        "candidate_change": candidate_change,
        "controls": controls,
        "confirmation": confirmation,
        "gates": gates,
        "source": source,
        "helpers": helpers,
        "epoch2_affinity_hardening": {
            "runner": helpers["capture"],
            "controls": helpers["affinity_controls"],
            "parent_preexec_affinity": CPU_MASK,
            "exact_process_affinity_getter": "os.sched_getaffinity(pid)",
            "explicit_taskset_argv_preserved": True,
            "launch_mismatch": "fail before wait/report parsing; no retry or replacement",
            "prior_epoch_capture_status": "invalid_due_to_launch_affinity",
            "fresh_54_child_schedule_required": True,
        },
        "root_screen_selector": helpers["selector"],
        "root_screen_selector_controls": {**helpers["selector_controls"], "description": "synthetic equality, invalid, no-signal, safety-stop, and counter-only controls"},
        "root_screening_rule": helpers["rule"],
        "runner": {"schema": "litchi.performance.0964.capture.v1", "deadline_seconds": 900, "cpu_mask": CPU_MASK},
        "reused_0959": {
            "capture_validator": descriptor(Path("/home/zhuhe/litchi-performance-0959/tools/capture_0959_v7.py"), True),
            "reducer": descriptor(Path("/home/zhuhe/litchi-performance-0959/tools/reduce_0959_v3_finalplan.py"), True),
            "reuse_policy": "import only nearest-rank/geometric-mean/ratio/bootstrap after exact SHA validation; all schedule and sample validation is 0964-specific",
        },
        "build_policy": {
            "builds_started_by_this_plan": False,
            "no_workload_before_root_GO": True,
            "source_freeze_required_before_build": True,
            "build_authorization": "root separately authorizes one isolated release build after quality",
            "cargo_processes": 1,
            "cargo_build_jobs": 2,
            "release_only": True,
            "locked_offline": True,
            "perf_baseline_lock_sha256": PERF_LOCK_SHA256,
        "candidate_build_tool": maybe_descriptor(build_tool_path),
        "candidate_build_receipt": pending_or(args.candidate_build_receipt or ROOT / f"receipts/candidate-managed-{build_revision}-build-result.json"),
        },
        "quality": {"receipt": maybe_descriptor(args.quality_receipt or ROOT / "receipts/quality-0964-final-v1.json"), "scope": "full XLSX quality plus frozen reader differential; source verified after quality before candidate freeze"},
        "host": ({**descriptor(args.host), "status": "bound", "receipt": descriptor(args.host)} if args.host else {"status": "pending_root_binding", "path": str(ROOT / "protocol/host-v1.json")}),
        "root_go": ({"status": "GO", "receipt": descriptor(args.root_go)} if args.root_go else {"status": "pending_root_binding", "path": str(ROOT / "protocol/root-go-epoch2-v1.json")}),
    }
    if args.final:
        for label, value in (("source.baseline_manifest", source["baseline_manifest"]),
                             ("source.candidate_manifest", source["candidate_manifest"]),
                             ("source.baseline_receipt", source["baseline_receipt"]),
                             ("source.candidate_freeze", source["candidate_freeze"]),
                             ("source.candidate_allowlist_file", source["candidate_allowlist_file"]),
                             ("host", plan["host"]), ("root GO", plan["root_go"]),
                             ("candidate build tool", plan["build_policy"]["candidate_build_tool"]),
                             ("candidate build receipt", plan["build_policy"]["candidate_build_receipt"]),
                             ("quality receipt", plan["quality"]["receipt"]),
                             ("epoch2 affinity controls", plan["helpers"].get("affinity_controls"))):
            required_final(value, label)
        for lane in ("native", "allocator"):
            required_final(lanes[lane]["candidate_binary"], f"{lane} candidate executable")
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(plan, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": plan["status"], "plan": str(args.output.resolve()), "candidate_allowlist": allowlist}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
