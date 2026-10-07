#!/usr/bin/env python3
"""Materialize the capture-ready 0961 plan from the frozen 0960 schedule.

The inherited runner/reducer own the 3-case, ABBA/BAAB, report, and terminal
contracts.  This generator only rebinds those contracts to the fresh 0961
source, candidate build receipts, and executables.  It refuses to emit a
plan until every candidate custody input is a real descriptor.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0961")
PREVIOUS = Path("/home/zhuhe/litchi-performance-0960/plan-final-v1.json")
OLD_ROOT = "/home/zhuhe/litchi-performance-0960"
NEW_ROOT = str(ROOT)
ALLOWLIST = [
    "crates/litchi-xlsx/src/cell_values/admission_tests.rs",
    "crates/litchi-xlsx/src/cell_values/source.rs",
]


def digest(path: Path, executable: bool | None = None) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file() or path.is_symlink():
        raise RuntimeError(f"regular file required: {path}")
    st = path.stat()
    actual_exec = bool(st.st_mode & 0o111)
    if executable is not None and actual_exec is not executable:
        raise RuntimeError(f"executable mismatch: {path}")
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
    return {"path": str(path), "bytes": st.st_size, "mode": stat.S_IMODE(st.st_mode),
            "executable": actual_exec, "sha256": h.hexdigest()}


def replace_old(value: Any) -> Any:
    if isinstance(value, str):
        return value.replace(OLD_ROOT, NEW_ROOT)
    if isinstance(value, list):
        return [replace_old(x) for x in value]
    if isinstance(value, dict):
        return {key: replace_old(item) for key, item in value.items()}
    return value


def desc(path: str | Path, executable: bool | None = None) -> dict[str, Any]:
    return digest(Path(path), executable)


def manifest_entries(path: Path) -> dict[str, tuple[int, int, str]]:
    value = json.loads(path.read_text(encoding="utf-8"))
    rows = value.get("files")
    if not isinstance(rows, list) or value.get("file_count") != len(rows):
        raise RuntimeError(f"invalid source manifest: {path}")
    result: dict[str, tuple[int, int, str]] = {}
    for row in rows:
        name = row.get("path")
        if not isinstance(name, str) or name in result:
            raise RuntimeError(f"invalid or duplicate manifest path: {path}: {name!r}")
        result[name] = (row.get("bytes"), row.get("mode"), row.get("sha256"))
    return result


def verify_candidate_manifest(candidate_manifest: Path) -> None:
    before = manifest_entries(ROOT / "source-manifest-baseline-v1.json")
    after = manifest_entries(candidate_manifest)
    changed = sorted(name for name in set(before) | set(after) if before.get(name) != after.get(name))
    if changed != ALLOWLIST:
        raise RuntimeError(f"candidate manifest changed paths {changed!r}; expected {ALLOWLIST!r}")
    if len(after) != 10_933:
        raise RuntimeError(f"candidate manifest file count {len(after)} != 10933")
    for name, (size, mode, sha) in after.items():
        path = ROOT / "source-candidate" / name
        if not path.is_file() or path.is_symlink():
            raise RuntimeError(f"candidate manifest path missing: {name}")
        data = path.read_bytes()
        actual = (len(data), stat.S_IMODE(path.stat().st_mode), hashlib.sha256(data).hexdigest())
        if actual != (size, mode, sha):
            raise RuntimeError(f"candidate manifest/source mismatch: {name}")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--candidate-manifest", type=Path, required=True)
    ap.add_argument("--candidate-freeze", type=Path, required=True,
                    help="immutable source/build-start receipt")
    ap.add_argument("--candidate-benchmark", type=Path, required=True)
    ap.add_argument("--candidate-allocator", type=Path, required=True)
    ap.add_argument("--host", type=Path)
    ap.add_argument("--root-go", type=Path,
                    help="fresh root GO receipt; omit for prospective check-only plan")
    ap.add_argument("--output", type=Path, default=ROOT / "protocol/plan-final-v1.json")
    args = ap.parse_args()

    if args.output.exists() or args.output.is_symlink():
        raise RuntimeError(f"refusing to overwrite plan: {args.output}")
    baseline_manifest = ROOT / "source-manifest-baseline-v1.json"
    if not baseline_manifest.is_file():
        raise RuntimeError(f"missing baseline manifest: {baseline_manifest}")
    verify_candidate_manifest(args.candidate_manifest)
    candidate_manifest = desc(args.candidate_manifest)
    candidate_freeze = desc(args.candidate_freeze)
    candidate_benchmark = desc(args.candidate_benchmark, True)
    candidate_allocator = desc(args.candidate_allocator, True)
    baseline_benchmark = desc("/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_benchmark.v2", True)
    baseline_allocator = desc("/home/zhuhe/litchi-performance-0957/executables/xlsx_managed_compression_alloc.v2", True)

    plan = replace_old(json.loads(PREVIOUS.read_text(encoding="utf-8")))
    plan["schema"] = "litchi.performance.0960.inline-number-screen-plan.v1"
    plan["protocol_schema"] = "litchi.performance.0961.xlsx-apply-batch-capacity-screen-plan.v1"
    plan["created_utc"] = datetime.now(timezone.utc).isoformat()
    plan["plan_id"] = "0961-apply-batch-capacity-screen"
    plan["plan_revision"] = "root-final-v1"
    plan["status"] = "authorized" if args.root_go else "prospective"
    plan["measurements_started"] = False
    plan["purpose"] = "Prospective descriptive screening of a bounded MultiSourceEdit::apply_batch capacity hint; screening cannot authorize adoption."
    plan["status_note"] = "0961 fresh source/binary bindings; inherited 0960 runner and report contract"
    plan["capture_root"] = {"native": str(ROOT / "capture-native-v1"), "allocator": str(ROOT / "capture-allocator-v1")}
    plan["analysis_root"] = str(ROOT / "analysis-screen-v1")
    plan["candidate_change"] = {
        "description": "bounded advisory iterator capacity hint in MultiSourceEdit::apply_batch",
        "expected_allowlist": ALLOWLIST,
        "production_adoption": False,
        "status": "quality_and_build_validated",
        "size_hint_is_advisory": True,
        "fallible_growth_and_max_batch_limit_unchanged": True,
    }
    plan["scope"] = {
        "format": "XLSX managed-compression benchmark",
        "candidate_mechanism": "bounded advisory MultiSourceEdit::apply_batch capacity hint",
        "cases": ["medium-one-percent", "dense-sparse-one-percent", "medium-one-edit"],
        "native_schedule": "3 blocks ABBA/BAAB/ABBA, 10 warmups, 60 measured samples, 36 serial fresh-process children",
        "allocator_schedule": "3 paired AB/BA/AB repeats, 2 warmups, 10 measured rows, 18 serial fresh-process children",
        "screening_only": True,
        "iwork_excluded": True,
    }
    plan["confirmation"] = {
        "native": "fresh plan with six ABBA/BAAB blocks, 20 warmups, 200 samples per child",
        "allocator": "fresh plan with three paired AB/BA repeats, 2 warmups, 10 samples per child",
        "fresh_bindings_required": True,
        "when": "only after a qualifying screen signal and all safety/oracle/control checks pass",
        "screening_result_not_adoption_evidence": True,
        "additional_controls_before_adoption": ["MultiSourceEdit exact 8-item and 256-item success", "ordinary non-managed and fallback performance controls"],
    }
    plan["source"] = {
        "base_commit": json.loads((ROOT / "normative-inputs.json").read_text(encoding="utf-8"))["head"],
        "baseline_manifest": desc(baseline_manifest),
        "candidate_manifest": candidate_manifest,
        "baseline_receipt": desc(ROOT / "receipts/baseline-binding-v1.json"),
        "candidate_freeze": candidate_freeze,
        "baseline_runtime_manifest": desc("/home/zhuhe/litchi-performance-0957/source-manifest-v2.json"),
        "baseline_source_root": str(ROOT / "source-baseline"),
        "baseline_runtime_source_root": "/home/zhuhe/litchi-performance-0957/source-v1",
        "candidate_source_root": str(ROOT / "source-candidate"),
        "candidate_runtime_source_root": str(ROOT / "source-candidate"),
        "candidate_change_allowlist": ALLOWLIST,
        "source_start": str(ROOT / "source-start.json"),
        "mode_correction": str(ROOT / "source-mode-correction.json"),
        "source_cache_state": "fresh process and in-memory OwnedSource; no cold filesystem claim",
    }
    plan["lanes"]["native"]["baseline_binary"] = baseline_benchmark
    plan["lanes"]["native"]["candidate_binary"] = candidate_benchmark
    plan["lanes"]["allocator"]["baseline_binary"] = baseline_allocator
    plan["lanes"]["allocator"]["candidate_binary"] = candidate_allocator
    plan["comparison"]["role_B"] = "bounded edit-staging capacity candidate executable"
    for lane in ("native", "allocator"):
        plan["lanes"][lane]["candidate_receipt"] = desc(ROOT / "receipts/candidate-managed-v1-build-result.json")
        plan["lanes"][lane]["execution"]["workload_cwd"] = str(ROOT / "source-baseline/tools/perf-baseline")
    plan["helpers"]["capture"] = desc("/home/zhuhe/litchi-performance-0960/tools/screen_0960_v7.py", True)
    plan["helpers"]["reducer"] = desc("/home/zhuhe/litchi-performance-0960/tools/reduce_0960_v3.py")
    plan["helpers"]["controls"] = desc("/home/zhuhe/litchi-performance-0960/tools/check_screening_contract_v9.py", True)
    selector = ROOT / "tools/select_screen_0961_v1.py"
    rule = ROOT / "protocol/root-screening-rule-0961-v1.json"
    selector_controls = ROOT / "tools/check_selector_0961_v1.py"
    plan["root_screen_selector"] = desc(selector)
    plan["root_screen_selector_controls"] = {
        **desc(selector_controls),
        "description": "six synthetic boundary, invalid, no-signal, and safety-stop controls",
    }
    plan["root_screening_rule"] = desc(rule)
    plan["root_go"] = {"status": "pending_root_binding"}
    if args.root_go:
        plan["root_go"] = {"status": "GO", "receipt": desc(args.root_go)}
    plan["host"] = {"status": "pending_root_binding"}
    if args.host:
        host_desc = desc(args.host)
        plan["host"] = {**host_desc, "status": "bound", "receipt": host_desc}
    plan["root_screening_rule"] = desc(rule)
    plan["gates"]["adoption_allowed"] = False
    plan["gates"]["full_followup_required"] = True
    plan["gates"]["screening_safety"] = "Any custody, terminal, report, semantic, inverse, resource, allocator, or control failure rejects without retry; native primary/native RSS/allocator region peak/allocator RSS median >1.05 stops."
    plan["gates"]["useful_signal"] = "native primary or native max_rss_kib or allocator region_peak_live_bytes median <=0.95; requested bytes/calls alone never signal"
    plan["build_policy"] = {
        "builds_started_by_this_plan": False,
        "no_workload_before_root_GO": True,
        "build_authorization": "root separately authorized isolated release build after quality",
        "source_freeze_required_before_build": True,
        "candidate_build_flags": "Rust1.95.0 release opt3 fatLTO cg1 debug1 incremental0 jobs2 RUSTFLAGS unset; build is root-owned and already terminal before plan generation.",
        "candidate_build_tool": desc(ROOT / "build-tools/build_candidate_managed_v1.py"),
        "candidate_build_start": candidate_freeze,
        "candidate_build_terminal": desc(ROOT / "receipts/candidate-managed-v1-build-terminal.json"),
        "candidate_build_result": desc(ROOT / "receipts/candidate-managed-v1-build-result.json"),
    }
    plan["supersedes"] = {**desc(PREVIOUS), "role": "historical 0960 schedule source; runner contract reused unchanged"}
    pending = []
    if not args.host:
        pending.append("host")
    if not args.root_go:
        pending.append("root GO")
    plan["status_note"] = (
        "0961 source, quality, build, helper, host, and selector bindings are final; no measurements have started."
        if not pending else
        "0961 source, quality, build, helper, and selector bindings are final; pending root bindings: " + ", ".join(pending) + "."
    )
    plan["quality"] = {"receipt": desc(ROOT / "quality/0961-quality-receipt.txt"), "scope": "full 1403 on production-final source before final test-only additions; final focused34, Clippy and scoped format"}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(plan, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({"status": plan["status"], "plan": str(args.output), "candidate_manifest": candidate_manifest,
                      "candidate_benchmark": candidate_benchmark, "candidate_allocator": candidate_allocator}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
