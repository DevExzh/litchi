#!/usr/bin/env python3
"""Validate one immutable 0978 decoder packet integration with reducer v2.

The input packets retain a 0977-origin native capture.  This validator
therefore skips only the original native command custody, original 0977 plan
contract, and packet output-boundary checks.  It still applies the runner
packet gate, every artifact binding, report/layout/loss/final-log oracle,
strict corrected parser and summary bridge, and the independent raw,
corrected, lifecycle, attribution, ledger, frame, map, and packet-claim
checks.  It never invokes perf, a workload, a build, or the raw dump scanner.
"""
from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import platform
import sys
from pathlib import Path
from typing import Any, Callable, Mapping, Sequence


ROOT = Path("/home/zhuhe/litchi-performance-0978")
DEFAULT_INTEGRATION = ROOT / "fresh-offline-child02-v1"
TOOLS = ROOT / "tools"
REDUCER = TOOLS / "reduce_0978_v2.py"
RETAINED_REDUCER = Path("/home/zhuhe/litchi-performance-0977/tools/reduce_0977.py")
RUNNER = TOOLS / "run_0978_v1.py"
LIFECYCLE = TOOLS / "lifecycle_ledger_0978.py"
FOCUSED_DECODER = TOOLS / "focuseddecode_0978.py"
FOCUSED_DECODER_V2 = TOOLS / "focuseddecode_0978_v2.py"
FOCUSED_DECODER_FROZEN_COPY = ROOT / "analysis/focuseddecode_0978-pre-full-lifetime-guard-866723.py"
MARKER = "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
PERIOD = 1_000_000
CPU = 31


class ValidationFailure(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    path = Path(path).absolute()
    if not path.is_file():
        raise ValidationFailure(f"missing bound file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def load(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(Path(path).read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValidationFailure(f"cannot load {path}: {exc}") from exc
    if not isinstance(value, dict):
        raise ValidationFailure(f"JSON root is not an object: {path}")
    return value


def load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ValidationFailure(f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as exc:  # noqa: BLE001
        raise ValidationFailure(f"cannot import {path}: {exc}") from exc
    return module


def path_from(row: Mapping[str, Any], label: str) -> Path:
    path = Path(row.get("path", "")).absolute()
    actual = binding(path)
    if row.get("bytes") != actual["bytes"] or row.get("sha256") != actual["sha256"]:
        raise ValidationFailure(f"{label} binding differs from file")
    return path


def result_call(label: str, function: Callable[[], Any]) -> tuple[dict[str, Any], Any | None]:
    try:
        value = function()
    except Exception as exc:  # noqa: BLE001
        return {"status": "fail", "error": f"{type(exc).__name__}: {exc}"}, None
    return {"status": "pass"}, value


def canonicalize_perf_ips(rows: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    """Normalize perf text's bare hexadecimal IP tokens before strict compare.

    ``_parse_corrected_text`` retains perf's hexadecimal token as a string,
    while JSON attribution rows retain the same IP as an integer.  This
    representation normalization changes no source claim or frame order; it
    lets the frozen frame comparator compare the same address exactly.
    """
    normalized: list[dict[str, Any]] = []
    for row in rows:
        copy = dict(row)
        frames: list[dict[str, Any]] = []
        for frame in row.get("frames", []):
            frame_copy = dict(frame)
            ip = frame_copy.get("ip")
            if isinstance(ip, str):
                try:
                    frame_copy["ip"] = int(ip, 16)
                except ValueError:
                    pass
            frames.append(frame_copy)
        copy["frames"] = frames
        normalized.append(copy)
    return normalized


def child_validation(
    *, old: Any, new: Any, runner: Any, lifecycle_module: Any,
    plan: Mapping[str, Any], plan_path: Path, packet_path: Path,
) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    packet = load(packet_path)
    packet_id = str(packet.get("packet_id"))
    child_dir = packet_path.parent
    child_results: dict[str, Any] = {"packet_id": packet_id, "packet": binding(packet_path), "checks": {}}

    artifacts: dict[str, dict[str, Any]] = {}
    retained: dict[str, dict[str, Any]] = {}
    canonical_result, canonical_value = result_call("canonical artifacts", lambda: new._artifact_map(packet))
    child_results["checks"]["canonical_artifacts"] = canonical_result
    if canonical_value is not None:
        artifacts = canonical_value
    retained_result, retained_value = result_call("retained artifacts", lambda: old._packet_artifacts(packet, packet_path))
    child_results["checks"]["retained_artifacts"] = retained_result
    if retained_value is not None:
        retained = retained_value
    if not artifacts or not retained:
        child_results["status"] = "rejected"
        return child_results, {}

    artifact_bindings: dict[str, dict[str, Any]] = {}
    for name, row in artifacts.items():
        path = path_from(row, f"{packet_id}.artifacts.{name}")
        artifact_bindings[f"{packet_id}.artifact.{name}"] = binding(path)
    child_results["artifact_count"] = len(artifact_bindings)
    child_results["checks"]["raw_task_exact_mmap_binding"] = {
        "status": "pass" if artifacts.get("raw_task") == artifacts.get("perf_mmap_dump") else "fail",
        "required": True,
    }

    packet_gate, gate_reasons = runner.packet_gate(packet)
    child_results["checks"]["runner_packet_gate"] = {
        "status": "pass" if packet_gate else "fail",
        "reasons": gate_reasons,
    }

    report_result, report = result_call("report", lambda: old._check_report(packet, retained, dict(plan)))
    child_results["checks"]["report_oracle"] = report_result
    layout_result, layout = result_call("layout", lambda: old._check_layout(packet, retained, dict(plan)))
    child_results["checks"]["layout_oracle"] = layout_result
    loss_result, loss = result_call("loss", lambda: old._check_loss(packet, retained))
    child_results["checks"]["loss_oracle"] = loss_result

    # The shim JSONL is the final decoder log.  corrected_log is the decoder's
    # stderr artifact and is still binding-checked above, but is not fed into
    # the final-log oracle.
    try:
        layout_for_log = old._load(old._artifact_path(retained, "layout"))
        shim_path = old._artifact_path(retained, "shim_log")
        final_log = old._check_final_log(shim_path, layout_for_log, plan["binary"]["path"])
        child_results["checks"]["final_log_oracle"] = {"status": "pass", "source": "shim_log", "rows": final_log["rows"], "exact_target_reports": final_log["exact_target_reports"], "exact_target_infos": final_log["exact_target_infos"]}
        child_results["final_log_binding"] = binding(shim_path)
    except Exception as exc:  # noqa: BLE001
        child_results["checks"]["final_log_oracle"] = {"status": "fail", "error": f"{type(exc).__name__}: {exc}"}

    marker_dso = str(plan["binary"]["path"])
    marker = str(plan["marker"]["symbol"])
    build_id = str(plan["marker"]["build_id"]).lower()
    launch = packet.get("launch") if isinstance(packet.get("launch"), Mapping) else {}
    script_path = old._artifact_path(retained, "corrected_script")
    parsed_old = None
    old_parse_result, parsed_old = result_call(
        "strict corrected parser",
        lambda: old.p.parse_perf_script_strict(
            script_path.read_text(encoding="utf-8", errors="strict"), marker=marker,
            marker_dso=marker_dso, expected_pid=int(launch["exec_pid"]), expected_cpu=CPU,
            expected_period=PERIOD,
        ),
    )
    child_results["checks"]["strict_corrected_parser"] = old_parse_result

    if parsed_old is not None:
        if loss is not None:
            bridge_result, bridge_value = result_call(
                "legacy summary bridge",
                lambda: new._compare_legacy_decoder_summary(old, packet, parsed_old, loss),
            )
            child_results["checks"]["legacy_summary_bridge"] = bridge_result
            if bridge_value is not None:
                child_results["legacy_summary_fields"] = {
                    "sample_count": bridge_value.get("sample_count"),
                    "marker_sample_count": bridge_value.get("marker_sample_count"),
                    "total_period": bridge_value.get("total_period"),
                    "marker_period": bridge_value.get("marker_period"),
                }
        else:
            child_results["checks"]["legacy_summary_bridge"] = {"status": "skipped", "reason": "strict loss oracle failed; no None/packet-loss coercion"}

    # Load the retained structured sidecars.  This parses their rows but does
    # not rerun the gigabyte raw dump scanner.
    try:
        raw_path = path_from(artifacts["raw_stack"], f"{packet_id}.raw_stack")
        lifecycle_path = path_from(artifacts["lifecycle_ledger"], f"{packet_id}.lifecycle_ledger")
        attribution_path = path_from(artifacts["attribution_samples"], f"{packet_id}.attribution_samples")
        corrected = new._parse_corrected_text(script_path, marker=marker, marker_dso=marker_dso, expected_period=PERIOD)
        raw_obj = load(raw_path)
        raw = new._parse_raw(raw_path, raw_obj, marker=marker, marker_dso=marker_dso)
        attribution_obj = load(attribution_path)
        attribution_corrected = new._parse_structured_rows(
            {"sample_ledger": attribution_obj["corrected_samples"]}, "attribution.corrected_samples",
            marker=marker, marker_dso=marker_dso, allow_summary_claims=True,
        )
        attribution_raw = new._parse_structured_rows(
            {"sample_ledger": attribution_obj["raw_samples"]}, "attribution.raw_samples",
            marker=marker, marker_dso=marker_dso, allow_summary_claims=True,
        )
        lifecycle_obj = load(lifecycle_path)
        lifecycle_rows, lifecycle_ledger = new._parse_lifecycle(lifecycle_path, lifecycle_obj, marker=marker, marker_dso=marker_dso)
        declared_lifecycle_rows = new._parse_structured_rows(
            {"sample_ledger": attribution_obj["lifecycle_samples"]}, "attribution.lifecycle_samples",
            marker=marker, marker_dso=marker_dso, allow_summary_claims=True, allow_unframed_claims=True,
        )
        child_results["sidecar_counts"] = {"corrected": len(corrected), "raw": len(raw), "attribution_corrected": len(attribution_corrected), "attribution_raw": len(attribution_raw), "lifecycle": len(lifecycle_rows)}

        # Canonicalize only the representation of bare perf hexadecimal IP
        # tokens, then run the frozen full frame/order comparator.
        frame_result, _ = result_call(
            "full corrected-frame linkage",
            lambda: new._compare_attribution_rows(
                canonicalize_perf_ips(corrected), canonicalize_perf_ips(attribution_corrected),
                "corrected(perf-hex-canonical)", "attribution.corrected_samples",
            ),
        )
        child_results["checks"]["full_corrected_frame_linkage"] = frame_result
        exact_result, _ = result_call(
            "exact attribution linkage",
            lambda: new._validate_exact_attribution_linkage(
                attribution_obj, corrected, marker=marker, marker_dso=marker_dso,
                target_build_id=build_id, allow_synthetic=False,
            ),
        )
        child_results["checks"]["exact_attribution_claim_linkage"] = exact_result
        raw_id_result, _ = result_call(
            "raw attribution identity",
            lambda: new._compare_identities(raw, attribution_raw, "raw_stack", "attribution.raw_samples"),
        )
        child_results["checks"]["raw_attribution_identity"] = raw_id_result
        corrected_lifecycle_result, _ = result_call(
            "corrected lifecycle identity",
            lambda: new._compare_identities(corrected, lifecycle_rows, "corrected", "lifecycle"),
        )
        child_results["checks"]["corrected_lifecycle_identity"] = corrected_lifecycle_result
        declared_lifecycle_result, _ = result_call(
            "declared lifecycle rows",
            lambda: new._compare_lifecycle_rows(lifecycle_rows, declared_lifecycle_rows, "lifecycle", "attribution.lifecycle_samples"),
        )
        child_results["checks"]["declared_lifecycle_rows"] = declared_lifecycle_result
        contract_result, _ = result_call(
            "lifecycle contract",
            lambda: new._require_live_lifecycle_contract(lifecycle_obj, lifecycle_rows, lifecycle_ledger),
        )
        child_results["checks"]["lifecycle_contract"] = contract_result

        # Derive disjoint D/M/R/T/U buckets from the retained normalized rows
        # and check every packet claim.  The lifecycle build uses the raw
        # scanner's original rows because it needs physical record offsets;
        # this is sidecar replay, not a fresh raw dump parse.
        actual_result, actual = result_call(
            "derived attribution",
            lambda: new._derive_child(corrected, raw, lifecycle_rows, lifecycle_ledger, marker_dso=marker_dso),
        )
        child_results["checks"]["derived_attribution"] = actual_result
        claims_result, _ = result_call(
            "packet claims",
            lambda: new._compare_packet_claims(packet, new._claim_object(packet), actual),
        ) if actual is not None else ({"status": "skipped", "reason": "derived attribution failed"}, None)
        child_results["checks"]["packet_claims"] = claims_result
        attribution_map = attribution_obj.get("attributions")
        if not isinstance(attribution_map, Mapping):
            raise ValidationFailure(f"{packet_id}: attribution map missing")
        attributions = {int(key): value for key, value in attribution_map.items()}
        stream = lifecycle_obj.get("stream")
        if not isinstance(stream, Mapping):
            raise ValidationFailure(f"{packet_id}: lifecycle stream missing")
        built_result, built = result_call(
            "lifecycle build",
            lambda: lifecycle_module.build_ledger(
                stream, raw_obj["sample_ledger"], expected_pid=int(launch["exec_pid"]), expected_cpu=CPU,
                expected_period=PERIOD, target_path=marker_dso, target_build_id=build_id,
                target_comm=Path(marker_dso).name[:15], terminal_lost_samples=0,
                marker=marker, marker_dso=marker_dso, attributions=attributions,
            ),
        )
        child_results["checks"]["lifecycle_build"] = built_result
        if built is not None:
            built_rows_result, _ = result_call("built lifecycle rows", lambda: new._compare_built_lifecycle_rows(built["samples"], lifecycle_rows))
            child_results["checks"]["built_lifecycle_rows"] = built_rows_result
            built_ledger_result, _ = result_call("built lifecycle ledger", lambda: new._compare_built_lifecycle_ledger(built["ledger"], lifecycle_obj, packet))
            child_results["checks"]["built_lifecycle_ledger"] = built_ledger_result
            if actual is not None:
                built_actual_result, _ = result_call("built ledger to derived", lambda: new._compare_built_ledger_to_actual(built["ledger"], actual))
                child_results["checks"]["built_ledger_to_derived"] = built_actual_result
            map_result, _ = result_call(
                "marker map linkage",
                lambda: new._validate_marker_map_linkage(
                    lifecycle_module, stream, attributions, expected_pid=int(launch["exec_pid"]),
                    target_comm=Path(marker_dso).name[:15], marker_dso=marker_dso, target_build_id=build_id,
                ),
            )
            child_results["checks"]["marker_map_linkage"] = map_result
    except Exception as exc:  # noqa: BLE001
        child_results["checks"]["sidecar_semantic_setup"] = {"status": "fail", "error": f"{type(exc).__name__}: {exc}"}

    required_failures = [name for name, result in child_results["checks"].items() if result.get("status") == "fail"]
    child_results["required_failures"] = required_failures
    child_results["status"] = "pass" if not required_failures else "rejected"
    return child_results, artifact_bindings


def write_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    path = Path(path).absolute()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
    except FileExistsError as exc:
        raise ValidationFailure(f"refusing to overwrite existing receipt: {path}") from exc


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--integration-root", type=Path, default=DEFAULT_INTEGRATION)
    parser.add_argument("--child", default="child-02", help="one child directory, for example child-02")
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    integration = args.integration_root.absolute()
    output = args.output.absolute()
    if output.exists():
        raise ValidationFailure(f"output already exists: {output}")
    plan_path = integration / "plan.json"
    plan = load(plan_path)
    if not isinstance(args.child, str) or len(args.child) != 8 or not args.child.startswith("child-") or not args.child[-2:].isdigit():
        raise ValidationFailure(f"--child must be child-NN, got {args.child!r}")
    child_number = int(args.child[-2:])
    if child_number <= 0:
        raise ValidationFailure(f"--child number must be positive, got {args.child!r}")
    packet_paths = [integration / args.child / f"{args.child}.packet.json"]
    dependency = plan.get("dependencies", {}).get("focused_decode") if isinstance(plan.get("dependencies"), Mapping) else None
    if not isinstance(dependency, Mapping):
        raise ValidationFailure("fresh plan focused_decode dependency is missing")
    if Path(str(dependency.get("path"))).absolute() != FOCUSED_DECODER.absolute():
        raise ValidationFailure("fresh plan focused_decode path is not the frozen original decoder path")
    if dependency.get("sha256") != "866723324881a57ce7c4282cdfd2dabe497397b1612e8d7c4f975cb7e9fe2134":
        raise ValidationFailure("fresh plan focused_decode dependency is not the frozen original SHA")
    if not FOCUSED_DECODER_V2.is_file() or not FOCUSED_DECODER_FROZEN_COPY.is_file():
        raise ValidationFailure("focused decoder lineage files are incomplete")
    source_paths: dict[str, Path] = {
        "validator": Path(__file__), "reducer_0978": REDUCER,
        "retained_reducer_0977": RETAINED_REDUCER, "runner": RUNNER,
        "lifecycle_helper": LIFECYCLE, "focused_decoder": FOCUSED_DECODER,
        "focused_decoder_frozen_copy": FOCUSED_DECODER_FROZEN_COPY,
        "focused_decoder_v2_candidate": FOCUSED_DECODER_V2, "plan": plan_path,
    }
    for packet_path in packet_paths:
        source_paths[f"packet_{args.child}"] = packet_path
        packet = load(packet_path)
        if packet.get("packet_id") != args.child:
            raise ValidationFailure(f"packet identity differs from --child: {packet.get('packet_id')!r}")
        for name, row in packet.get("artifacts", {}).items():
            source_paths[f"{args.child}.artifact.{name}"] = Path(row["path"])
    before = {name: binding(path) for name, path in source_paths.items()}
    if before["focused_decoder"]["sha256"] != dependency["sha256"]:
        raise ValidationFailure("current focused decoder does not match fresh plan's frozen original SHA")
    if before["focused_decoder_frozen_copy"]["sha256"] != dependency["sha256"]:
        raise ValidationFailure("frozen focused decoder copy does not match fresh plan SHA")
    if before["focused_decoder_v2_candidate"]["sha256"] == dependency["sha256"]:
        raise ValidationFailure("focused decoder v2 candidate unexpectedly equals frozen original")
    old = load_module("reduce_0977_packet_validator", RETAINED_REDUCER)
    new = load_module("reduce_0978_v2_packet_validator", REDUCER)
    regression_result, regression = result_call("reducer hexadecimal IP regression", new._perf_ip_hex_regression)
    if regression is None:
        raise ValidationFailure(f"reducer hexadecimal IP regression failed: {regression_result}")
    sys.path.insert(0, str(TOOLS))
    runner = load_module("run_0978_packet_validator", RUNNER)
    lifecycle_module = load_module("lifecycle_ledger_packet_validator", LIFECYCLE)
    children: list[dict[str, Any]] = []
    for packet_path in packet_paths:
        result, _ = child_validation(
            old=old, new=new, runner=runner, lifecycle_module=lifecycle_module,
            plan=plan, plan_path=plan_path, packet_path=packet_path,
        )
        children.append(result)
    after = {name: binding(path) for name, path in source_paths.items()}
    stable = before == after
    if not stable:
        raise ValidationFailure("bound source/artifact changed during validation")
    failures = [f"{child['packet_id']}:{name}" for child in children for name in child.get("required_failures", [])]
    receipt = {
        "schema": "litchi.performance.0978.packet_integration_validation.v4",
        "status": "pass" if not failures else "rejected",
        "scope": {
            "integration_root": str(integration),
            "packets": [child["packet_id"] for child in children],
            "requested_child": args.child,
            "native_actions": {"build": False, "recording": False, "workload": False},
            "raw_dump_scanner_rerun": False,
            "cohort_admitted": False,
        },
        "command": {
            "argv": [sys.executable, "-B", str(Path(__file__).absolute()), "--integration-root", str(integration), "--child", args.child, "--output", str(output)],
            "cwd": os.getcwd(), "offline_only": True, "perf_subprocesses": 0,
            "python": platform.python_version(),
        },
        "explicit_exclusions": {
            "original_0977_native_command_custody": "skipped: packets point to retained 0977-origin native command and this validator must not relabel it as 0978 custody",
            "original_0977_plan_contract": "skipped: retained capture predates the 0978 packet integration plan; plan marker/binary/event fields are still used as identity inputs",
            "packet_output_boundary": "skipped: retained packet directories are copied 0977 outputs, not a 0978 native capture output root",
        },
        "focused_decoder_lineage": {
            "plan_expected_original_sha256": dependency["sha256"],
            "original_path": str(FOCUSED_DECODER.absolute()),
            "original_frozen_copy_path": str(FOCUSED_DECODER_FROZEN_COPY.absolute()),
            "original_frozen_copy_matches_plan": before["focused_decoder"]["sha256"] == dependency["sha256"] and before["focused_decoder_frozen_copy"]["sha256"] == dependency["sha256"],
            "v2_candidate_path": str(FOCUSED_DECODER_V2.absolute()),
            "v2_candidate_sha256": before["focused_decoder_v2_candidate"]["sha256"],
            "v2_candidate_distinct_from_original": before["focused_decoder_v2_candidate"]["sha256"] != dependency["sha256"],
            "v2_candidate_used_for_this_packet": False,
            "v2_candidate_scope": "separate lineage guard only; no packet was generated or interpreted with this source",
        },
        "reducer_lineage": {
            "path": str(REDUCER.absolute()),
            "sha256": before["reducer_0978"]["sha256"],
            "perf_ip_hex_regression": regression,
            "scope": "final reducer v2 lineage; perf frame IP tokens are parsed base-16, including digits-only tokens, and attribution-contract labels distinguish all-DSO known physical ownership from target-source eligibility",
        },
        "source_bindings_before": before,
        "source_bindings_after": after,
        "source_inputs_stable_before_after": stable,
        "children": children,
        "required_failure_count": len(failures),
        "required_failures": failures,
        "limitations": [
            "A rejected receipt is diagnostic and does not admit a cohort or support speedup/causal claims.",
            "The final decoder oracle uses the explicitly bound shim_log DWFL JSONL; corrected_log remains the bound stderr artifact.",
            "The new reducer parses perf text IP tokens as base-16 integers, including digits-only tokens; the validator's compatibility canonicalizer is representation-only and drops no frame, owner, or order row.",
        ],
    }
    write_exclusive(output, receipt)
    print(json.dumps({"status": receipt["status"], "receipt": str(output), "required_failures": failures}, sort_keys=True))
    return 0 if not failures else 2


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except ValidationFailure as exc:
        print(json.dumps({"status": "failed", "error": str(exc)}, sort_keys=True), file=sys.stderr)
        raise SystemExit(1)
