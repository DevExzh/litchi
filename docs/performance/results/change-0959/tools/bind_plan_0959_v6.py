#!/usr/bin/env python3
"""Bind the reviewed 0959 Arc-shared-Number comparison plan with v6 report-schema controls.

This is a create-only plan binder.  It performs custody and source-manifest
checks, records actual candidate build/binary/host/helper descriptors, and
leaves the plan prospective unless a separately supplied root GO receipt is
passed.  It never builds, launches, or edits the checkout.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import stat
import sys
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0959").resolve()
TEMPLATE = ROOT / "plan-template-0959-v4.json"
SCHEMA = "litchi.performance.0959.inline-number-comparison-plan.v4"
ALLOWLIST = {
    "crates/litchi-xlsx/src/cell.rs",
    "crates/litchi-xlsx/src/raw/worksheet/semantic.rs",
    "crates/litchi-xlsx/src/raw/worksheet/selected.rs",
}
LOCK_PATH = "Cargo.lock"
VALID_PASS = {"pass", "passed", "complete", "bound", "authorized", "success"}


class BindError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise BindError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_descriptor(path: Path, *, executable: bool = False) -> dict[str, Any]:
    path = path.expanduser().resolve()
    require(path.exists() and path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    mode = stat.S_IMODE(path.stat().st_mode)
    if executable:
        require(bool(path.stat().st_mode & stat.S_IXUSR), f"file is not executable: {path}")
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "mode": mode,
        "executable": bool(path.stat().st_mode & (stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH)),
        "sha256": sha256(path),
    }


def read_json(path: Path, *, label: str, limit: int = 64 * 1024 * 1024) -> Any:
    path = path.expanduser().resolve()
    require(path.exists() and path.is_file() and not path.is_symlink(), f"{label} is missing: {path}")
    require(path.stat().st_size <= limit, f"{label} exceeds size limit")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise BindError(f"cannot read {label}: {path}: {exc}") from exc


def descriptor_from_arg(path: Path, label: str, *, executable: bool = False) -> dict[str, Any]:
    return file_descriptor(path, executable=executable)


def manifest_entries(path: Path, label: str) -> dict[str, tuple[int, int, str]]:
    value = read_json(path, label=label, limit=128 * 1024 * 1024)
    raw = value.get("files", value.get("entries")) if isinstance(value, dict) else None
    require(isinstance(raw, list), f"{label} has no files list")
    entries: dict[str, tuple[int, int, str]] = {}
    for index, item in enumerate(raw):
        require(isinstance(item, dict), f"{label}.files[{index}] is not an object")
        rel = item.get("path", item.get("name"))
        require(isinstance(rel, str) and rel and not rel.startswith("/"), f"{label}.files[{index}] path invalid")
        require(rel not in entries, f"{label} has duplicate path {rel}")
        size = item.get("bytes", item.get("size"))
        mode = item.get("mode", 0)
        digest = item.get("sha256")
        require(isinstance(size, int) and size >= 0, f"{label}.{rel} bytes invalid")
        require(isinstance(mode, int) and mode >= 0, f"{label}.{rel} mode invalid")
        require(isinstance(digest, str) and len(digest) == 64, f"{label}.{rel} SHA invalid")
        entries[rel] = (size, mode, digest)
    declared = value.get("file_count")
    if declared is not None:
        require(declared == len(entries), f"{label} file_count mismatch")
    return entries


def validate_overlay(
    probe_path: Path,
    comparison_path: Path,
    comparison_receipt_path: Path,
    candidate_path: Path,
) -> dict[str, Any]:
    probe = manifest_entries(probe_path, "original test-only probe manifest")
    comparison = manifest_entries(comparison_path, "baseline comparison manifest")
    candidate = manifest_entries(candidate_path, "candidate manifest")

    added_to_comparison = sorted(set(comparison) - set(probe))
    removed_from_comparison = sorted(set(probe) - set(comparison))
    require(added_to_comparison == [LOCK_PATH], f"baseline comparison added paths: {added_to_comparison}")
    require(not removed_from_comparison, f"baseline comparison removed paths: {removed_from_comparison}")
    common_probe = set(probe) & set(comparison)
    require(all(probe[p] == comparison[p] for p in common_probe), "baseline comparison changed a probe entry")
    require(comparison[LOCK_PATH][2] == "d80aa535dcf880a85fa31e92fe2e5ce78b91d93a55953ae734e513aed83e0061", "comparison Cargo.lock hash changed")

    require(set(candidate) == set(comparison), "candidate manifest paths differ from comparison baseline")
    changed = sorted(path for path in candidate if candidate[path] != comparison[path])
    require(set(changed) == ALLOWLIST, f"candidate manifest changed paths: {changed}")
    require(candidate[LOCK_PATH] == comparison[LOCK_PATH], "candidate Cargo.lock is not byte-identical to comparison baseline")

    receipt = read_json(comparison_receipt_path, label="baseline comparison freeze receipt")
    require(isinstance(receipt, dict) and receipt.get("status") in VALID_PASS, "baseline comparison freeze did not pass")
    manifest_ref = receipt.get("manifest")
    require(isinstance(manifest_ref, dict), "baseline comparison receipt lacks manifest binding")
    require(manifest_ref.get("sha256") == sha256(comparison_path), "baseline comparison receipt manifest differs")
    prior_ref = receipt.get("prior_probe_manifest")
    require(isinstance(prior_ref, dict) and prior_ref.get("sha256") == sha256(probe_path), "comparison receipt prior probe differs")
    lock = receipt.get("lock")
    require(isinstance(lock, dict) and lock.get("path") == str((ROOT / "source-baseline/Cargo.lock").resolve()), "comparison lock path differs")
    require(lock.get("sha256") == comparison[LOCK_PATH][2] and lock.get("bytes") == comparison[LOCK_PATH][0], "comparison receipt lock differs")
    receipt_ref = file_descriptor(comparison_receipt_path)
    return {
        "status": "reviewed_test_build_overlay",
        "runtime": False,
        "paths": [LOCK_PATH],
        "manifest": file_descriptor(comparison_path),
        "receipt": receipt_ref,
        "prior_probe_manifest": file_descriptor(probe_path),
        "candidate_manifest": file_descriptor(candidate_path),
        "candidate_delta_paths": sorted(ALLOWLIST),
        "interpretation": "identical generated root Cargo.lock is a comparison/build overlay; retained 0957 runtime baseline is unchanged",
    }


def resolve_nested(base: Path, value: Any, label: str) -> Path:
    if isinstance(value, str):
        raw = value
    elif isinstance(value, dict):
        raw = value.get("path")
    else:
        raw = None
    require(isinstance(raw, str) and raw, f"{label} path missing")
    path = Path(raw).expanduser()
    if path.is_absolute():
        actual = path
        require(actual.exists(), f"{label} path missing: {actual}")
        return actual.resolve()
    candidates: list[Path] = []
    for root in [base.parent, ROOT, *base.parents]:
        candidate = (root / path).resolve()
        if candidate.exists() and candidate.is_file() and candidate not in candidates:
            candidates.append(candidate)
    require(len(candidates) == 1, f"{label} relative path is missing or ambiguous")
    return candidates[0]


def ref_digest(value: Any, base: Path, label: str) -> tuple[Path, dict[str, Any]]:
    path = resolve_nested(base, value, label)
    expected = value if isinstance(value, dict) else {}
    actual = file_descriptor(path)
    if isinstance(expected, dict):
        for key in ("bytes", "sha256"):
            if key in expected:
                require(actual[key] == expected[key], f"{label} {key} mismatch")
    return path, actual


def status_passed(value: Any, label: str) -> None:
    status = value.get("status") if isinstance(value, dict) else None
    if status is not None:
        require(status in VALID_PASS, f"{label} status is not passed: {status}")


def build_receipt_contract(
    result_path: Path,
    binary_path: Path,
    lane: str,
    candidate_manifest_sha: str,
) -> dict[str, Any]:
    result = read_json(result_path, label=f"{lane} build result", limit=32 * 1024 * 1024)
    require(isinstance(result, dict), f"{lane} build result is not an object")
    status_passed(result, f"{lane} build result")
    if result.get("status") is None:
        require(result.get("returncode", result.get("return_code")) == 0, f"{lane} build return code is not zero")
        require(result.get("timed_out") is False, f"{lane} build timed out")
        require(result.get("source_unchanged") is True, f"{lane} build source changed")
        require(result.get("all_binaries_present") is True, f"{lane} build did not produce all binaries")
    result_dir = result_path.resolve().parent

    start_value = result.get("start_receipt", result.get("start"))
    if start_value is not None:
        start_path, _ = ref_digest(start_value, result_path, f"{lane} build start receipt")
        start = read_json(start_path, label=f"{lane} build start receipt")
        source_manifest = start.get("source_manifest") or start.get("manifest")
        if isinstance(source_manifest, dict):
            require(source_manifest.get("sha256") == candidate_manifest_sha, f"{lane} build source manifest differs")
        source_root = start.get("source_root")
        if source_root is not None:
            require(Path(str(source_root)).resolve().is_dir(), f"{lane} build source root unavailable")

    terminal_value = result.get("terminal_receipt", result.get("terminal"))
    terminal = None
    terminal_path = None
    if terminal_value is not None:
        terminal_path, _ = ref_digest(terminal_value, result_path, f"{lane} build terminal receipt")
        terminal = read_json(terminal_path, label=f"{lane} build terminal receipt")
        require(isinstance(terminal, dict), f"{lane} terminal is not an object")
        require(terminal.get("returncode", terminal.get("return_code")) == 0, f"{lane} terminal return code is not zero")
        require(terminal.get("timed_out") is False, f"{lane} terminal timed out")
        if "process_group_clear" in terminal:
            require(terminal["process_group_clear"] is True, f"{lane} build process group remains")
        source = terminal.get("source")
        if isinstance(source, dict):
            manifest = source.get("manifest")
            if isinstance(manifest, dict):
                require(manifest.get("sha256") == candidate_manifest_sha, f"{lane} terminal source manifest differs")

    binary_value = result.get("binary")
    if isinstance(binary_value, dict):
        _, reported = ref_digest(binary_value, result_path, f"{lane} result binary")
        require(reported["sha256"] == sha256(binary_path), f"{lane} result binary differs")
        require(reported["bytes"] == binary_path.stat().st_size, f"{lane} result binary size differs")

    binary_receipt_value = result.get("binary_receipt") or result.get("binaries_receipt")
    if binary_receipt_value is None and isinstance(result.get("build"), dict):
        binary_receipt_value = result["build"].get("binary_receipt")
    binary_receipt = None
    if binary_receipt_value is not None:
        binary_receipt_path, _ = ref_digest(binary_receipt_value, result_path, f"{lane} binary receipt")
        binary_receipt = read_json(binary_receipt_path, label=f"{lane} binary receipt")
        require(isinstance(binary_receipt, dict), f"{lane} binary receipt is not an object")
        binaries = binary_receipt.get("binaries", {})
        role_name = "xlsx_managed_compression_benchmark" if lane == "native" else "xlsx_managed_compression_alloc"
        reported = binaries.get(role_name) if isinstance(binaries, dict) else None
        require(isinstance(reported, dict) and reported.get("exists") is True, f"{lane} binary receipt lacks {role_name}")
        require(reported.get("sha256") == sha256(binary_path), f"{lane} binary receipt SHA differs")
        require(reported.get("bytes") == binary_path.stat().st_size, f"{lane} binary receipt size differs")
    require(terminal is not None or result.get("terminal") is not None, f"{lane} terminal receipt missing")
    require(binary_receipt is not None or isinstance(binary_value, dict), f"{lane} binary receipt missing")
    return {
        "status": "validated",
        "result": file_descriptor(result_path),
        "terminal": file_descriptor(terminal_path) if terminal_path else None,
        "binary_receipt": file_descriptor(resolve_nested(result_path, binary_receipt_value, f"{lane} binary receipt")) if binary_receipt_value is not None else None,
        "candidate_manifest_sha256": candidate_manifest_sha,
        "lane": lane,
    }


def validate_freeze(path: Path, candidate_manifest_sha: str) -> dict[str, Any]:
    value = read_json(path, label="candidate source freeze receipt")
    require(isinstance(value, dict), "candidate source freeze receipt is not an object")
    if value.get("status") is not None:
        require(value["status"] in VALID_PASS, f"candidate source freeze did not pass: {value['status']}")
    for key in ("manifest", "candidate_manifest", "source_manifest"):
        ref = value.get(key)
        if isinstance(ref, dict) and ref.get("sha256") is not None:
            require(ref["sha256"] == candidate_manifest_sha, f"candidate freeze {key} differs")
    return {"status": "validated", "receipt": file_descriptor(path)}


def validate_controls_receipt(path: Path, capture_path: Path, template_path: Path) -> dict[str, Any]:
    value = read_json(path, label="negative-control receipt")
    require(isinstance(value, dict) and value.get("status") == "passed", "negative-control receipt did not pass")
    summary_ref = value.get("summary")
    require(isinstance(summary_ref, dict), "negative-control receipt summary binding missing")
    summary_path, _ = ref_digest(summary_ref, path, "negative-control summary")
    summary = read_json(summary_path, label="negative-control summary")
    require(isinstance(summary, dict) and summary.get("status") == "passed", "negative-control summary did not pass")
    require(summary.get("runner", {}).get("sha256") == sha256(capture_path), "negative-control runner differs from capture helper")
    require(summary.get("plan", {}).get("sha256") == sha256(template_path), "negative-control plan differs from template")
    controls = summary.get("controls")
    require(isinstance(controls, list) and len(controls) == 4, "negative-control count is not four")
    require(all(isinstance(item, dict) and item.get("rejected") is True for item in controls), "negative-control admission was not fail-closed")
    positive = summary.get("positive_allocator_reports")
    require(isinstance(positive, list) and len(positive) == 3, "allocator positive-control count is not three")
    expected = {
        "medium-one-percent": "medium-one-percent-w1-r0",
        "dense-sparse-one-percent": "dense-sparse-one-percent-w1-r0",
        "medium-one-edit": "medium-one-edit-w1-r0",
    }
    seen: set[str] = set()
    for index, item in enumerate(positive):
        require(isinstance(item, dict) and item.get("status") == "passed", f"allocator positive control {index} did not pass")
        case_id = item.get("case_id")
        require(case_id in expected and case_id not in seen, f"allocator positive control {index} case is invalid")
        seen.add(case_id)
        require(item.get("child_id") == f"allocator__{case_id}__b00__p0", f"allocator positive control {index} child binding changed")
        report = item.get("report")
        require(isinstance(report, dict), f"allocator positive control {index} report binding missing")
        report_path = Path(report.get("path", "")).resolve()
        expected_path = Path("/home/zhuhe/litchi-performance-0957/capture-allocation-v1/reports") / expected[case_id] / "report.json"
        require(report_path == expected_path.resolve(), f"allocator positive control {index} report path changed")
        require(report.get("sha256") == sha256(report_path), f"allocator positive control {index} report digest changed")
        require(report.get("bytes") == report_path.stat().st_size, f"allocator positive control {index} report size changed")
    require(seen == set(expected), "allocator positive controls do not cover all cases")
    return {"status": "validated", "receipt": file_descriptor(path), "summary": file_descriptor(summary_path), "positive_allocator_reports": positive}


def bind(args: argparse.Namespace) -> dict[str, Any]:
    template_path = args.template.resolve()
    plan = read_json(template_path, label="plan template")
    require(isinstance(plan, dict) and plan.get("schema") == SCHEMA, "template is not the reviewed v4 plan")
    require(plan.get("status") == "prospective_candidate_pending", "template status was already bound or changed")

    probe_path = args.probe_manifest.resolve()
    comparison_path = args.comparison_manifest.resolve()
    comparison_receipt_path = args.comparison_receipt.resolve()
    candidate_path = args.candidate_manifest.resolve()
    overlay = validate_overlay(probe_path, comparison_path, comparison_receipt_path, candidate_path)
    candidate_manifest_desc = file_descriptor(candidate_path)
    comparison_manifest_desc = file_descriptor(comparison_path)
    comparison_receipt_desc = file_descriptor(comparison_receipt_path)
    probe_desc = file_descriptor(probe_path)
    prior_probe_receipt_path = args.probe_receipt.resolve()
    prior_probe_receipt_desc = file_descriptor(prior_probe_receipt_path)

    freeze_info = validate_freeze(args.candidate_freeze.resolve(), candidate_manifest_desc["sha256"])
    native_binary = file_descriptor(args.native_binary.resolve(), executable=True)
    allocator_binary = file_descriptor(args.allocator_binary.resolve(), executable=True)
    native_contract = build_receipt_contract(args.native_result.resolve(), args.native_binary.resolve(), "native", candidate_manifest_desc["sha256"])
    allocator_contract = build_receipt_contract(args.allocator_result.resolve(), args.allocator_binary.resolve(), "allocator", candidate_manifest_desc["sha256"])
    host_desc = file_descriptor(args.host.resolve())
    host_value = read_json(args.host.resolve(), label="host receipt")
    require(isinstance(host_value, dict), "host receipt is not an object")

    helper_desc = {
        "capture": file_descriptor(args.capture.resolve()),
        "reducer": file_descriptor(args.reducer.resolve()),
        "controls": file_descriptor(args.controls.resolve()),
    }
    controls_receipt = validate_controls_receipt(args.controls_receipt.resolve(), args.capture.resolve(), template_path)
    helper_desc["controls_receipt"] = file_descriptor(args.controls_receipt.resolve())
    root_go_desc = None
    root_go_value = None
    if args.root_go is not None:
        root_go_path = args.root_go.resolve()
        root_go_value = read_json(root_go_path, label="root GO receipt")
        require(isinstance(root_go_value, dict) and root_go_value.get("status") in {"GO", "passed", "authorized"}, "root GO receipt is not passed")
        root_go_desc = file_descriptor(root_go_path)

    plan["binding_revision"] = "v6"
    plan["status"] = "authorized" if root_go_desc else "prospective"
    plan["schema_note"] = "The v4 schema name is retained for the reviewed schedule and gate contract; this v6 binding records the Arc-shared Number candidate explicitly and binds positive allocator report-schema controls."
    plan["purpose"] = "Fresh-process paired native and allocator comparison for an isolated Arc-shared Number candidate; no adoption claim."
    plan["candidate_change"] = {
        **plan["candidate_change"],
        "description": "Arc-shared Number representation candidate under review; candidate source is bound to the reviewed three-path delta.",
        "representation": "Arc<str> shared lexical Number storage with borrowed lexical construction and exact integer construction",
        "candidate_source_manifest": candidate_manifest_desc,
        "candidate_source_freeze": file_descriptor(args.candidate_freeze.resolve()),
    }
    original_test_only_probe = plan["source"].get("test_only_probe")
    plan["source"] = {
        **plan["source"],
        "baseline_manifest": comparison_manifest_desc,
        "baseline_manifest_sha256": comparison_manifest_desc["sha256"],
        "baseline_receipt": comparison_receipt_desc,
        "probe_manifest": probe_desc,
        "probe_receipt": prior_probe_receipt_desc,
        "original_test_only_probe": original_test_only_probe,
        "test_only_probe": {
            "status": "explicit_test_only_overlay",
            "runtime": False,
            "paths": ["crates/litchi-xlsx/src/cell.rs"],
            "manifest": comparison_manifest_desc,
            "receipt": comparison_receipt_desc,
            "base_probe_manifest": probe_desc,
        },
        "comparison_manifest": comparison_manifest_desc,
        "comparison_receipt": comparison_receipt_desc,
        "comparison_overlay": overlay,
        "candidate_manifest": candidate_manifest_desc,
        "candidate_freeze": file_descriptor(args.candidate_freeze.resolve()),
        "baseline_runtime_source_is_0957_retained": True,
        "candidate_source_root": str(Path(str(read_json(candidate_path, label="candidate manifest").get("source_root", ""))).resolve()),
    }
    plan["comparison"] = {
        **plan["comparison"],
        "role_B": "candidate isolated Arc-shared Number executable",
    }
    plan["candidate_build"] = {
        **plan["candidate_build"],
        "status": "bound",
        "native_receipt": file_descriptor(args.native_result.resolve()),
        "allocator_receipt": file_descriptor(args.allocator_result.resolve()),
        "native_contract": native_contract,
        "allocator_contract": allocator_contract,
        "candidate_source_manifest": candidate_manifest_desc,
    }
    plan["lanes"]["native"]["candidate_binary"] = native_binary
    plan["lanes"]["allocator"]["candidate_binary"] = allocator_binary
    plan["helpers"] = helper_desc
    plan["host"] = host_desc
    plan["root_go"] = {
        "role": "root qualification GO receipt",
        "status": "GO" if root_go_desc else "pending_root_binding",
        **({"receipt": root_go_desc} if root_go_desc else {}),
    }
    plan["binding"] = {
        "schema": "litchi.performance.0959.final-plan-binding.v6",
        "status": "authorized" if root_go_desc else "prospective_pending_root_go",
        "script": file_descriptor(Path(__file__)),
        "template": file_descriptor(template_path),
        "host": host_desc,
        "source_baseline_comparison": comparison_manifest_desc,
        "source_baseline_comparison_receipt": comparison_receipt_desc,
        "controls_receipt": controls_receipt,
        "runtime_baseline_manifest_unchanged": True,
        "root_go_bound": bool(root_go_desc),
    }
    plan["pending_bindings"] = [] if root_go_desc else ["root qualification GO receipt"]
    return plan


def write_exclusive(path: Path, value: dict[str, Any]) -> None:
    path = path.expanduser().resolve()
    require(not path.exists(), f"refusing to overwrite existing plan: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def parser() -> argparse.ArgumentParser:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--template", type=Path, default=TEMPLATE)
    p.add_argument("--output", type=Path, required=True, help="new plan path; must not already exist")
    p.add_argument("--probe-manifest", type=Path, default=ROOT / "source-manifest-probe-v2.json")
    p.add_argument("--probe-receipt", type=Path, default=ROOT / "receipts/source-baseline-v2.json")
    p.add_argument("--comparison-manifest", type=Path, default=ROOT / "source-manifest-baseline-comparison-v1.json")
    p.add_argument("--comparison-receipt", type=Path, default=ROOT / "receipts/baseline-comparison-freeze-v1.json")
    p.add_argument("--candidate-manifest", type=Path, required=True)
    p.add_argument("--candidate-freeze", type=Path, required=True)
    p.add_argument("--native-result", type=Path, required=True)
    p.add_argument("--native-binary", type=Path, required=True)
    p.add_argument("--allocator-result", type=Path, required=True)
    p.add_argument("--allocator-binary", type=Path, required=True)
    p.add_argument("--host", type=Path, default=ROOT / "receipts/host-v1.json")
    p.add_argument("--capture", type=Path, required=True)
    p.add_argument("--reducer", type=Path, required=True)
    p.add_argument("--controls", type=Path, required=True)
    p.add_argument("--controls-receipt", type=Path, default=ROOT / "receipts/controls-0959-v6.json")
    p.add_argument("--root-go", type=Path)
    return p


def main(argv: list[str] | None = None) -> int:
    try:
        args = parser().parse_args(argv)
        plan = bind(args)
        write_exclusive(args.output, plan)
        print(args.output.resolve())
        return 0
    except BindError as exc:
        print(f"bind-plan error: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
