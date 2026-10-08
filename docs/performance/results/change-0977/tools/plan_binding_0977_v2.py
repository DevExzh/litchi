#!/usr/bin/env python3
"""Fail-closed custody checks for the fresh 0977 qualification packet.

This module is deliberately independent of the retained 0974/0975 runners.
It validates the immutable external inputs and the small 0977 helper set, but
does not launch perf or an Office workload.
"""
from __future__ import annotations

import hashlib
import json
import os
import stat
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0977")
REPO = Path("/home/zhuhe/code/litchi")
PLAN_SCHEMA = "litchi.performance.0977.plan.v1"
HELPER_SCHEMA = "litchi.performance.0977.helper_seal.v1"
GO_SCHEMA = "litchi.performance.0977.root_go.v1"
RUNNER_SCHEMA = "litchi.performance.0977.runner.v1"

CPU = 31
RECORDER_CPU = 30
PERIOD = 1_000_000
STACK_CAP = 32_768
WARMUPS = 5
SAMPLES = 80
CHILDREN = 12
TIMEOUT_SECONDS = 900

PERF = Path("/usr/bin/perf")
TASKSET = Path("/usr/bin/taskset")
PYTHON = Path("/usr/bin/python3.14")
LAUNCHER = Path("/home/zhuhe/litchi-performance-0970/tools/launch_baseline_v2.py")
BRIDGE_LINEAGE = Path("/home/zhuhe/litchi-performance-0976/tools/affinity_bridge_0976_v2.py")
DECODER_SO = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.so")
DECODER_SOURCE = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.c")
DECODER_MAP = Path("/home/zhuhe/litchi-performance-0976/tools/dwfl_dynamic_shim_0976_cache_final_v2.map")
DECODER_VALIDATOR = Path("/home/zhuhe/litchi-performance-0976/tools/validate_and_noop_final_v2_0973_v3.py")
LIBDW = Path("/usr/lib/x86_64-linux-gnu/libdw-0.194.so")
LIBELF = Path("/usr/lib/x86_64-linux-gnu/libelf-0.194.so")
HEADER = Path("/home/zhuhe/litchi-performance-0972/header-v2/root/usr/include/elfutils/libdwfl.h")
CAPABILITY = Path("/home/zhuhe/litchi-performance-0974/capability-perthread-mmap16m-v1/capability-receipt-v1.json")
CONTROL_TARGET = Path("/home/zhuhe/litchi-performance-0976/tools/stdlib_cpu_burn_0976_v2.py")
SOURCE_MANIFEST = Path("/home/zhuhe/litchi-performance-0969/manifests/source-manifest-v1.json")
SOURCE_FILE_MANIFEST = Path("/home/zhuhe/litchi-performance-0969/manifests/source-file-manifest-v1.json")
DECODER_REVIEW = Path("/home/zhuhe/litchi-performance-0976/analysis/independent-decoder-review-0976-v1.json")
NOOP_REVIEW = Path("/home/zhuhe/litchi-performance-0976/receipts/final-v2-regression-0973-v2-posthoc-v3.json")
FOCUSED_CONTROLS = ROOT / "receipts/focuseddecode-controls-0977-v4.json"
PERF_REQUALIFICATION = ROOT / "receipts/current-perf-requalification-0977-v1-posthoc-v2.json"
PERF_REQUALIFICATION_SOURCE = ROOT / "receipts/current-perf-requalification-0977-v1.json"
PERF_REQUALIFICATION_POSTHOC_SOURCE = ROOT / "receipts/current-perf-requalification-0977-v1-posthoc.json"
PERF_TOOL_DRIFT = ROOT / "receipts/perf-tool-drift-0977-v1.json"
PERF_DRIFT_REVIEW = ROOT / "analysis/perf-drift-review-0977-v1.json"

# ``perf report --header-only`` on the current installed perf can emit an
# Invalid HEADER_EVENT_DESC warning and omit the event descriptor.  The
# decoder therefore treats the single ``perf evlist -v`` line as the
# authoritative event contract and retains the report-header output/stderr as
# descriptive evidence.  Keep this contract in the plan so a future decoder
# cannot silently switch authorities.
EVLIST_REQUIRED_ATTRIBUTES = (
    "type: 4 (cpu)",
    "size: 144",
    "config: 0xc0 (instructions)",
    "{ sample_period, sample_freq }: 1000000",
    "sample_type: IP|TID|TIME|CALLCHAIN|CPU|PERIOD|REGS_USER|STACK_USER",
    "read_format: ID|LOST",
    "disabled: 1",
    "exclude_kernel: 1",
    "exclude_hv: 1",
    "mmap: 1",
    "comm: 1",
    "enable_on_exec: 1",
    "task: 1",
    "mmap_data: 1",
    "sample_id_all: 1",
    "exclude_callchain_kernel: 1",
    "exclude_callchain_user: 1",
    "mmap2: 1",
    "comm_exec: 1",
    "ksymbol: 1",
    "bpf_event: 1",
    "build_id: 1",
    "sample_regs_user: 0xff0fff",
    "sample_stack_user: 32768",
)
PERF_CONTRACT = {
    "authority": "perf evlist -v; report header descriptor is retained separately and is not admission authority",
    "event": "instructions:u",
    "period": PERIOD,
    "stack_cap": STACK_CAP,
    "inherit": 0,
    "required_attributes": list(EVLIST_REQUIRED_ATTRIBUTES),
    "header": {
        "command": "perf report --header-only --stdio",
        "authority": "descriptive_only",
        "retain_stdout_and_stderr": True,
        "known_warning": "Invalid HEADER_EVENT_DESC: nre=1 sz=144 (min 64)",
        "event_attributes_source": "evlist-v",
    },
}

# Retained 0967 ordinary binary/fixture are the known 0975 protocol inputs.
# A later plan may replace these with a fully bound fresh binary, but the
# default generator refuses to invent an unbound candidate.
BINARY = Path("/home/zhuhe/litchi-performance-0967/builds/target-baseline-cfb/release/doc_retention_probe_0942_cfb_profile")
FIXTURE = Path("/home/zhuhe/litchi-performance-0969/qa-core/test-data/ole/doc/FloatingPictures.doc")


class BindingError(ValueError):
    pass


def utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def strict_load(path: Path) -> Any:
    def pairs(items: list[tuple[str, Any]]) -> dict[str, Any]:
        out: dict[str, Any] = {}
        for key, value in items:
            if key in out:
                raise BindingError(f"{path}: duplicate JSON key {key!r}")
            out[key] = value
        return out

    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise BindingError(f"cannot load JSON {path}: {exc}") from exc


def sha256(path: str | Path) -> str:
    path = Path(path)
    if not path.is_file() or path.is_symlink():
        raise BindingError(f"not a regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def binding(path: str | Path) -> dict[str, Any]:
    path = Path(path).resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def check_binding(row: Any, label: str, expected: Path | None = None) -> Path:
    if not isinstance(row, dict):
        raise BindingError(f"{label}: missing binding")
    path_value, digest = row.get("path"), row.get("sha256")
    if not isinstance(path_value, str) or not isinstance(digest, str) or len(digest) != 64:
        raise BindingError(f"{label}: malformed path/SHA")
    path = Path(path_value).resolve()
    if expected is not None and path != expected.resolve():
        raise BindingError(f"{label}: path drift ({path} != {expected.resolve()})")
    actual = binding(path)
    if actual["sha256"] != digest or (row.get("bytes") is not None and actual["bytes"] != row["bytes"]):
        raise BindingError(f"{label}: bytes or SHA drift")
    return path


def write_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    encoded = (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    except OSError as exc:
        raise BindingError(f"refusing to overwrite {path}: {exc}") from exc
    try:
        with os.fdopen(fd, "wb") as handle:
            handle.write(encoded)
            handle.flush()
            os.fsync(handle.fileno())
    except Exception:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def environment() -> dict[str, str]:
    return {"LANG": "C", "LC_ALL": "C", "TZ": "UTC", "LD_PRELOAD": "absent", "PYTHONDONTWRITEBYTECODE": "1"}


def helper_names() -> tuple[str, ...]:
    return ("bridge", "runner", "plan_binding", "plan_generator", "launcher", "decoder", "reducer")


def validate_seal(plan: dict[str, Any], plan_path: Path | None = None) -> dict[str, Any]:
    if plan.get("schema") != PLAN_SCHEMA or plan.get("status") != "pre_capture":
        raise BindingError("plan schema/status mismatch")
    if plan.get("no_build_or_workload") is not True:
        raise BindingError("plan must remain pre-capture")
    expected_plan_path = str(plan_path.resolve()) if plan_path is not None else str(Path(plan.get("plan_path", "")).resolve())
    if plan.get("plan_path") != expected_plan_path:
        raise BindingError("plan self path mismatch")
    batch = plan.get("batch") or {}
    expected_batch = {
        "children": CHILDREN, "serial": True, "cpu": CPU, "recorder_cpu": RECORDER_CPU,
        "period": PERIOD, "stack_cap": STACK_CAP, "warmups": WARMUPS, "samples": SAMPLES,
        "minimum_marker_samples": 100, "marker_period_positive": True,
        "raw_lost_sum": 0, "terminal_lost_samples": 0, "empty_callchain_count": 0,
        "stop_on_first_failure": True, "retry": False,
    }
    if batch != expected_batch:
        raise BindingError(f"batch contract drift: {batch}")
    execution = plan.get("execution") or {}
    if execution.get("serial") is not True or execution.get("retry") is not False or execution.get("timeout_seconds") != TIMEOUT_SECONDS:
        raise BindingError("execution contract drift")
    check_binding(plan.get("design"), "design")
    adr_path = check_binding(plan.get("adr_goal_revalidation"), "adr_goal_revalidation")
    adr = strict_load(adr_path)
    if adr.get("schema") != "litchi.performance.0977.adr_goal_revalidation.v1" or adr.get("status") != "pass":
        raise BindingError("ADR/GOAL revalidation is not a passing 0977 receipt")
    if adr.get("accepted_adr_count") != 32 or adr.get("build_started") is not False or adr.get("workload_started") is not False:
        raise BindingError("ADR/GOAL revalidation receipt has unsafe status")
    check_binding(plan.get("capability_probe"), "capability_probe", CAPABILITY)
    # Keep the original v1 receipt as the source experiment record.  The
    # posthoc v2 receipt below is the canonical current-perf admission record;
    # binding only the source receipt would omit its four-way canonical
    # sequence and evlist-v authority evidence.
    requal_path = check_binding(plan.get("perf_requalification"), "perf_requalification", PERF_REQUALIFICATION_SOURCE)
    requal = strict_load(requal_path)
    if requal.get("schema") != "litchi.performance.0977.current_perf_requalification.v1" or requal.get("status") != "complete":
        raise BindingError("current perf requalification is not complete")
    if requal.get("native_workload_started") is not False or requal.get("build_started") is not False:
        raise BindingError("current perf requalification has unsafe provenance")
    current_perf = requal.get("tool", {}).get("current", {})
    if current_perf.get("path") != str(PERF.resolve()) or current_perf.get("sha256") != sha256(PERF) or current_perf.get("sha256") == requal.get("tool", {}).get("prior_sealed_sha256"):
        raise BindingError("current perf requalification tool binding drift")
    commands = {row.get("name"): row for row in requal.get("commands", []) if isinstance(row, dict)}
    if set(commands) != {"original-corrected", "translated-corrected", "original-stock", "original-noop"}:
        raise BindingError("current perf requalification command set drift")
    for name in ("original-corrected", "translated-corrected"):
        row = commands[name]
        if row.get("returncode") != 0 or row.get("quality") != {
            "empty_zero": True, "marker_minimum_100": True, "period_6940000000": True,
            "sample_count_6940": True, "script_lost_zero": True,
        }:
            raise BindingError(f"current perf corrected requalification failed: {name}")
    if commands["original-stock"].get("returncode") != 0 or commands["original-noop"].get("returncode") != 0:
        raise BindingError("current perf stock/noop requalification did not complete")
    identity = requal.get("identity_controls", {})
    if identity.get("stock_noop_script_byte_identical") is not True or identity.get("stock_noop_stderr_byte_identical") is not True:
        raise BindingError("current perf stock/noop identity control failed")
    drift_path = check_binding(plan.get("perf_tool_drift"), "perf_tool_drift", PERF_TOOL_DRIFT)
    drift = strict_load(drift_path)
    if drift.get("schema") != "litchi.performance.0977.perf_tool_drift.v1" or drift.get("status") != "recorded":
        raise BindingError("perf tool drift receipt is not the recorded host transition")
    if drift.get("native_workload_started") is not False or drift.get("build_started") is not False:
        raise BindingError("perf tool drift receipt has unsafe provenance")
    drift_tool = drift.get("tool", {})
    if drift_tool.get("path") != str(PERF.resolve()) or drift_tool.get("current_sha256") != sha256(PERF):
        raise BindingError("perf tool drift current binding mismatch")
    if drift_tool.get("current_sha256") == drift_tool.get("frozen_prior_sha256"):
        raise BindingError("perf tool drift receipt does not record a changed perf identity")

    posthoc_path = check_binding(plan.get("perf_requalification_posthoc"), "perf_requalification_posthoc", PERF_REQUALIFICATION)
    posthoc = strict_load(posthoc_path)
    if posthoc.get("schema") != "litchi.performance.0977.current_perf_requalification_posthoc.v2" or posthoc.get("status") != "pass":
        raise BindingError("canonical current perf requalification is not passing")
    if posthoc.get("native_workload_started") is not False:
        raise BindingError("canonical current perf requalification has unsafe provenance")
    source_row = posthoc.get("source_receipt")
    source_path = check_binding(source_row, "perf_requalification_posthoc.source_receipt", PERF_REQUALIFICATION_POSTHOC_SOURCE)
    source = strict_load(source_path)
    if source.get("schema") != "litchi.performance.0977.current_perf_requalification_posthoc.v1" or source.get("status") != "pass":
        raise BindingError("canonical requalification source receipt is not passing")
    current_perf_posthoc = source.get("current_perf", {})
    if current_perf_posthoc.get("path") != str(PERF.resolve()) or current_perf_posthoc.get("sha256") != sha256(PERF):
        raise BindingError("canonical requalification current perf binding drift")
    canonical = posthoc.get("canonical_sequence", {})
    canonical_sha = "1f7335d982514a9f586e0a3a001343ce7332deb650a9bf6ebae41f60f62dca3c"
    if canonical.get("all_four_equal") is not True or canonical.get("count") != 6940 or canonical.get("sha256") != canonical_sha:
        raise BindingError("canonical four-way current perf sequence failed")
    members = canonical.get("members", {})
    if set(members) != {"current_original", "current_translated", "retained_0976_original", "retained_0976_translated"}:
        raise BindingError("canonical sequence member set drift")
    expected_summary = {
        "empty_callchain_count": 0, "explicit_stack_truncated_count": 0,
        "lost_events": 0, "marker_period": 951000000,
        "marker_sample_count": 951, "sample_count": 6940,
        "total_period": 6940000000, "unresolved_frames": 3289,
    }
    for name, member in members.items():
        if member.get("canonical_sha256") != canonical_sha or member.get("summary") != expected_summary:
            raise BindingError(f"canonical sequence member failed: {name}")
    evlist_rows = posthoc.get("evlist_v_authority", {})
    if set(evlist_rows) != {"original", "translated"}:
        raise BindingError("evlist-v authority lane set drift")
    for name, row in evlist_rows.items():
        contract = row.get("contract", {})
        if (contract.get("status") != "pass" or contract.get("schema") != "litchi.performance.0977.evlist_contract.v1"
                or contract.get("authority") != PERF_CONTRACT["authority"]
                or contract.get("event") != PERF_CONTRACT["event"]
                or contract.get("period") != PERF_CONTRACT["period"]
                or contract.get("stack_cap") != PERF_CONTRACT["stack_cap"]
                or contract.get("inherit") != PERF_CONTRACT["inherit"]
                or contract.get("required_attributes") != PERF_CONTRACT["required_attributes"]):
            raise BindingError(f"evlist-v authority contract failed: {name}")
        if row.get("returncode") != 0 or row.get("stderr", {}).get("bytes") != 0:
            raise BindingError(f"evlist-v command failed: {name}")
    header_warning = posthoc.get("header_warning", {})
    if (header_warning.get("schema") != "litchi.performance.0977.perf_header_metadata.v1"
            or header_warning.get("authority") != "evlist-v; header warning is descriptive and not silently ignored"
            or header_warning.get("known_invalid_header_event_desc_warning") is not True
            or header_warning.get("warnings") != [PERF_CONTRACT["header"]["known_warning"]]):
        raise BindingError("current perf report-header warning disposition drift")
    check_binding(plan.get("source_manifest"), "source_manifest", SOURCE_MANIFEST)
    check_binding(plan.get("source_file_manifest"), "source_file_manifest", SOURCE_FILE_MANIFEST)
    check_binding(plan.get("binary"), "binary", BINARY)
    check_binding(plan.get("fixture"), "fixture", FIXTURE)
    if plan["binary"].get("sha256") != "0db34357a1e725dc397a4fd895532433db1214c5862b7bc96bf23f21c0396ecb":
        raise BindingError("retained ordinary binary SHA is not the sealed 0967 baseline")
    if plan["fixture"].get("sha256") != "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d":
        raise BindingError("FloatingPictures fixture SHA is not the sealed input")
    for name, expected in (("perf", PERF), ("taskset", TASKSET), ("python", PYTHON), ("launcher", LAUNCHER),
                           ("bridge_lineage", BRIDGE_LINEAGE), ("decoder_so", DECODER_SO),
                           ("decoder_source", DECODER_SOURCE), ("decoder_map", DECODER_MAP),
                           ("decoder_validator", DECODER_VALIDATOR), ("libdw", LIBDW), ("libelf", LIBELF),
                           ("decoder_header", HEADER), ("control_target", CONTROL_TARGET)):
        check_binding(plan.get("dependencies", {}).get(name), f"dependencies.{name}", expected)
    for name, expected in (("perf", PERF), ("taskset", TASKSET), ("python", PYTHON), ("launcher", LAUNCHER)):
        top = plan.get(name)
        dep = plan.get("dependencies", {}).get(name)
        if not isinstance(top, dict) or top.get("path") != dep.get("path") or top.get("sha256") != dep.get("sha256"):
            raise BindingError(f"top-level {name} binding does not equal sealed dependency")
    if plan["dependencies"]["bridge_lineage"].get("scope") != "handoff lineage; 0977 bridge is a fresh target-generic adaptation":
        raise BindingError("bridge lineage scope missing")
    if plan.get("environment") != environment():
        raise BindingError("native environment contract drift")
    if plan.get("terminal_contract") != {
        "field": "PERF_RECORD_LOST_SAMPLES", "required_value": 0,
        "source": "decoder-owned offline perf report --stdio",
        "runner_precomputes": False, "raw_loss_not_added": True,
    }:
        raise BindingError("terminal report ownership contract drift")
    command_contract = plan.get("command_receipt_contract") or {}
    if command_contract.get("cwd") != str(REPO.resolve()) or command_contract.get("no_secret_environment") is not True or command_contract.get("LD_PRELOAD") != "absent":
        raise BindingError("command receipt/cwd contract drift")
    control = plan.get("control") or {}
    if control.get("target_script", {}).get("path") != str(CONTROL_TARGET.resolve()) or control.get("target_script", {}).get("sha256") != sha256(CONTROL_TARGET):
        raise BindingError("harmless control target binding drift")
    if control.get("seconds") != 0.25 or control.get("check_interval") != 0.02 or control.get("output_root") != str((ROOT / "capture/affinity-control-v1").resolve()):
        raise BindingError("harmless control contract drift")
    perf = plan.get("perf") or {}
    if perf.get("record_mode") != ["--per-thread"] or perf.get("record_memory") != ["-m", "16M"]:
        raise BindingError("perf record mode drift")
    if perf.get("callgraph") != "dwarf,32768" or perf.get("event") != "instructions:u" or perf.get("period") != PERIOD:
        raise BindingError("perf sampling contract drift")
    if plan.get("perf_contract") != PERF_CONTRACT:
        raise BindingError("perf event/header authority contract drift")
    workload = plan.get("workload") or {}
    for key in ("case", "edits", "retention", "target_argv_template"):
        if key not in workload:
            raise BindingError(f"workload.{key} missing")
    if workload.get("case") != "docfloat" or workload.get("edits") != 1 or workload.get("retention") != "default":
        raise BindingError("0977 workload cell must remain docfloat/e1/default")
    if workload.get("target_argv_template") != [
        "<binary>", "--case", "docfloat", "--input", "<fixture>", "--edits", "1",
        "--retention", "default", "--warmup", "5", "--samples", "80",
    ]:
        raise BindingError("workload argv template drift")
    child_gate = plan.get("child_gate")
    expected_child_gate = {
        "all_children": True, "raw_lost_sum": 0, "terminal_lost_samples": 0,
        "min_marker_samples": 100, "marker_period_share_positive": True,
        "empty_callchain_count": 0, "on_failure": "stop_without_retry",
    }
    if child_gate != expected_child_gate:
        raise BindingError("child gate contract drift")
    reviews = plan.get("decoder_reviews")
    if not isinstance(reviews, dict) or set(reviews) != {"final_v2_review", "noop_posthoc", "focused_controls"}:
        raise BindingError("decoder review binding set drift")
    review_path = check_binding(reviews["final_v2_review"], "decoder_reviews.final_v2_review", DECODER_REVIEW)
    review = strict_load(review_path)
    if review.get("status") != "pass" or review.get("decision") != "pass_for_offline_decoder_evidence_only":
        raise BindingError("final-v2 decoder review is not passing offline evidence")
    noop_path = check_binding(reviews["noop_posthoc"], "decoder_reviews.noop_posthoc", NOOP_REVIEW)
    noop = strict_load(noop_path)
    if noop.get("status") != "pass" or noop.get("native_capture_started") is not False or noop.get("native_workload_started") is not False:
        raise BindingError("no-op decoder review is not a passing posthoc receipt")
    focused_path = check_binding(reviews["focused_controls"], "decoder_reviews.focused_controls", FOCUSED_CONTROLS)
    focused = strict_load(focused_path)
    if focused.get("schema") != "litchi.performance.0977.focuseddecode_controls_receipt.v4" or focused.get("status") != "pass":
        raise BindingError("focused decoder control receipt is not passing")
    if focused.get("native_workload_started") is not False:
        raise BindingError("focused decoder controls have unsafe provenance")
    if focused.get("current_perf_sha256") != sha256(PERF) or focused.get("current_perf", {}).get("path") != str(PERF.resolve()):
        raise BindingError("focused decoder controls are bound to a different perf tool")
    if focused.get("decoder") != (plan.get("helpers") or {}).get("decoder"):
        raise BindingError("focused decoder controls are not bound to the sealed decoder helper")
    helpers = plan.get("helpers") or {}
    if set(helpers) != set(helper_names()):
        raise BindingError("helper binding set drift")
    seal_row = plan.get("helper_seal")
    seal_path = check_binding(seal_row, "helper_seal")
    seal = strict_load(seal_path)
    if seal.get("schema") != HELPER_SCHEMA or seal.get("status") != "sealed" or seal.get("helpers") != helpers:
        raise BindingError("helper seal mismatch")
    if plan.get("helper_seal_sha256") != seal_row.get("sha256"):
        raise BindingError("helper seal digest is not bound")
    for name, row in helpers.items():
        check_binding(row, f"helper.{name}")
    outputs = plan.get("outputs") or {}
    capture_root = Path(outputs.get("capture_root", "")).resolve()
    analysis_root = Path(outputs.get("analysis_root", "")).resolve()
    if capture_root != (ROOT / "capture/instruction-profile-v2").resolve() or analysis_root != (ROOT / "analysis").resolve():
        raise BindingError("fresh output roots drifted")
    return {"status": "pass", "schema": PLAN_SCHEMA, "plan": str(plan_path or plan["plan_path"]), "helper_count": len(helpers)}


def validate_plan(path: Path) -> dict[str, Any]:
    return validate_seal(strict_load(path), path)
