#!/usr/bin/env python3
"""Fail-closed custody checks for the fresh 0978 qualification packet.

This module is deliberately independent of the retained 0977 runner and
decoder.  It validates the immutable external inputs and the fresh 0978
variable-layout/lifecycle helper set, but does not launch perf or an Office
workload.
"""
from __future__ import annotations

import hashlib
import json
import os
import stat
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

ROOT = Path("/home/zhuhe/litchi-performance-0978")
REPO = Path("/home/zhuhe/code/litchi")
PLAN_SCHEMA = "litchi.performance.0978.plan.v1"
HELPER_SCHEMA = "litchi.performance.0978.helper_seal.v1"
GO_SCHEMA = "litchi.performance.0978.root_go.v1"
RUNNER_SCHEMA = "litchi.performance.0978.runner.v1"

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
# The 0977 bridge is retained byte-for-byte.  The fresh runner binds this
# exact helper instead of copying or silently modifying the native handoff.
BRIDGE = Path("/home/zhuhe/litchi-performance-0977/tools/affinity_bridge_0977.py")
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
# These are retained 0977 host/tool custody receipts.  The new policy and
# parser receipts live under ROOT and are bound independently below.
FOCUSED_CONTROLS = Path("/home/zhuhe/litchi-performance-0977/receipts/focuseddecode-controls-0977-v4.json")
PERF_REQUALIFICATION = Path("/home/zhuhe/litchi-performance-0977/receipts/current-perf-requalification-0977-v1-posthoc-v2.json")
PERF_REQUALIFICATION_SOURCE = Path("/home/zhuhe/litchi-performance-0977/receipts/current-perf-requalification-0977-v1.json")
PERF_REQUALIFICATION_POSTHOC_SOURCE = Path("/home/zhuhe/litchi-performance-0977/receipts/current-perf-requalification-0977-v1-posthoc.json")
PERF_TOOL_DRIFT = Path("/home/zhuhe/litchi-performance-0977/receipts/perf-tool-drift-0977-v1.json")
PERF_DRIFT_REVIEW = Path("/home/zhuhe/litchi-performance-0977/analysis/perf-drift-review-0977-v1.json")

PROTOCOL_DESIGN = ROOT / "protocol-design-v1.md"
ADR_RECEIPT = ROOT / "receipts/adr-goal-revalidation-0978-v1.json"
POLICY_REVIEW = ROOT / "review/independent-design-review-v3.json"
SCANNER = ROOT / "tools/raw_stack_0978.py"
LIFECYCLE = ROOT / "tools/lifecycle_ledger_0978.py"
SCANNER_TESTS = ROOT / "tools/test_raw_stack_0978.py"
LIFECYCLE_TESTS = ROOT / "tools/test_lifecycle_ledger_0978.py"
REDUCER_TESTS = ROOT / "tools/test_reduce_0978.py"
REPLAY_TOOL = ROOT / "tools/replay_raw_stack_0978.py"
REDUCER = ROOT / "tools/reduce_0978_v4.py"
RETAINED_REDUCER_0978_V1 = ROOT / "tools/reduce_0978.py"
# The fresh reducer delegates retained custody/oracle checks to this exact
# 0977 implementation.  Keep the complete imported closure explicit so a
# later plan cannot silently substitute a different arithmetic helper.
RETAINED_REDUCER_0977 = Path("/home/zhuhe/litchi-performance-0977/tools/reduce_0977.py")
RETAINED_REDUCER_0973 = Path("/home/zhuhe/litchi-performance-0973/tools-v2/reduce_0973_v2.py")
RETAINED_PROTOCOL_0973 = Path("/home/zhuhe/litchi-performance-0973/tools-v2/protocol_common_0973_v2.py")
RETAINED_PROTOCOL_0971 = Path("/home/zhuhe/litchi-performance-0971/tools/protocol_common_0971_v2.py")
RETAINED_PROTOCOL_0969 = Path("/home/zhuhe/litchi-performance-0969/tools/protocol_common_v4.py")
RETAINED_REDUCER_CLOSURE = (
    RETAINED_REDUCER_0977,
    RETAINED_REDUCER_0973,
    RETAINED_PROTOCOL_0973,
    RETAINED_PROTOCOL_0971,
    RETAINED_PROTOCOL_0969,
)
RETAINED_REDUCER_CLOSURE_SHA256 = {
    RETAINED_REDUCER_0977: "65746c0f22f3434567830a479775e568528657bcc9ecb7467e8f8ea116f817d1",
    RETAINED_REDUCER_0973: "f33f14724f6264d87b2c3a0082c16bca711787140d06f3c7cfdd6e78b6f4c63b",
    RETAINED_PROTOCOL_0973: "56b509fde01bbd66702ac8cd835da0d09d74d332935926078c6545093ed52bf4",
    RETAINED_PROTOCOL_0971: "4ff9af4c64b50896f39a6a30286d445ebe0c9364435888593ade45a11c5cab43",
    RETAINED_PROTOCOL_0969: "5ce0f2ac08fc7909df054a4cc2d5137b3279af4cd726acf5bfb4558d97b05c5c",
}
RETAINED_REDUCER_LINEAGE_SHA256 = {
    **RETAINED_REDUCER_CLOSURE_SHA256,
    RETAINED_REDUCER_0978_V1: "e904caa12e395cc7b6917cc03cf82509168338c94291575b3ecef282639ae626",
}
REPLAY_RECEIPT = ROOT / "receipts/replay-0977-child01-02-v1.json"
SCANNER_CONTROLS = ROOT / "receipts/raw-stack-scanner-controls-0978-v2.json"
RAW_TEST_REVIEW = ROOT / "review/independent-raw-stack-tests-v1-final.json"
CONTROL_COMMAND_RECEIPT = Path("/home/zhuhe/litchi-performance-0977/receipts/root-control-command-0977-v1.json")
CONTROL_TEST = Path("/home/zhuhe/litchi-performance-0977/tools/test_control_replay_0977_v2.py")
CONTROL_AUDIT = Path("/home/zhuhe/litchi-performance-0977/analysis/cpu-review-control-audit-0977-v2.json")
CONTROL_EQUIVALENCE = ROOT / "receipts/control-boundary-equivalence-0978-v5.json"
FOCUSED_DECODER = ROOT / "tools/focuseddecode_0978_v3.py"
FOCUSED_DECODER_SHA256 = "9596e246c7c627fb3810e7c6c6d3dda12af58b62477f1f1233654c4256e2dac9"

# Decoder-owned analysis artifacts.  The compact ABI summary remains an
# optional diagnostic; these full streams and their offline command custody
# rows are mandatory for a fresh child packet.
DECODER_CANONICAL_ARTIFACTS = (
    "corrected_script", "corrected_log", "shim_log", "corrected_command",
    "perf_header", "perf_header_stderr", "perf_header_command",
    "terminal_report", "terminal_report_stderr", "terminal_report_command",
    "perf_evlist", "perf_evlist_stderr", "perf_evlist_command",
    "perf_mmap_dump", "perf_mmap_dump_stderr", "perf_mmap_dump_command",
    "raw_stack", "lifecycle_ledger", "attribution_samples",
)
DECODER_CANONICAL_FILENAMES = (
    "<child>.perf-header.txt", "<child>.perf-header.stderr", "<child>.perf-header.command.json",
    "<child>.terminal-report.txt", "<child>.terminal-report.stderr", "<child>.terminal-report.command.json",
    "<child>.perf-evlist-v.txt", "<child>.perf-evlist-v.stderr", "<child>.perf-evlist-v.command.json",
    "<child>.perf-mmap-dump.txt", "<child>.perf-mmap-dump.stderr", "<child>.perf-mmap-dump.command.json",
    "<child>.corrected.perf-script-no-inline.txt", "<child>.corrected.perf.stderr",
    "<child>.corrected.dwfl.jsonl", "<child>.corrected.command.json",
    "<child>.raw-stack.json", "<child>.lifecycle-ledger.json", "<child>.attribution-samples.json",
)
DECODER_OPTIONAL_FILENAMES = ("<child>.abi-scan.json",)

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
    return (
        "bridge", "runner", "plan_binding", "plan_generator", "launcher", "decoder",
        "scanner", "lifecycle", "reducer", "tests", "policy",
    )


def check_helper_binding(value: Any, label: str) -> None:
    """Validate either one helper row or the immutable test-row collection."""
    if isinstance(value, list):
        if not value:
            raise BindingError(f"{label}: helper list is empty")
        for index, row in enumerate(value):
            check_binding(row, f"{label}[{index}]")
        return
    check_binding(value, label)


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
        "raw_lost_sum": 0, "terminal_lost_samples": 0,
        "structural_conservation_required": True, "valid_empty_is_unknown": True,
        "unknown_period_cutoff": None,
        "stop_on_first_failure": True, "retry": False,
    }
    if batch != expected_batch:
        raise BindingError(f"batch contract drift: {batch}")
    execution = plan.get("execution") or {}
    if execution.get("serial") is not True or execution.get("retry") is not False or execution.get("timeout_seconds") != TIMEOUT_SECONDS:
        raise BindingError("execution contract drift")
    check_binding(plan.get("design"), "design", PROTOCOL_DESIGN)
    policy_path = check_binding(plan.get("policy_review"), "policy_review", POLICY_REVIEW)
    policy = strict_load(policy_path)
    if (policy.get("schema") != "litchi.performance.0978.independent_design_review.v3"
            or policy.get("status") != "offline_policy_pass_with_required_variable_layout_protocol"):
        raise BindingError("0978 independent design policy is not the passing v3 receipt")
    if policy.get("decision", {}).get("recommendation") != "phase_scoped_marker_inclusive_attribution_with_explicit_unknown_remainder":
        raise BindingError("0978 policy recommendation drift")
    if "unknown_period" not in policy.get("recommended_future_protocol", {}).get("threshold_policy", {}):
        raise BindingError("0978 policy lacks explicit unknown-period rule")
    adr_path = check_binding(plan.get("adr_goal_revalidation"), "adr_goal_revalidation", ADR_RECEIPT)
    adr = strict_load(adr_path)
    if adr.get("schema") != "litchi.performance.0978.adr_goal_revalidation.v1" or adr.get("status") != "pass":
        raise BindingError("ADR/GOAL revalidation is not a passing 0978 receipt")
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
                           ("bridge", BRIDGE),
                           ("bridge_lineage", BRIDGE_LINEAGE), ("decoder_so", DECODER_SO),
                           ("decoder_source", DECODER_SOURCE), ("decoder_map", DECODER_MAP),
                           ("decoder_validator", DECODER_VALIDATOR), ("libdw", LIBDW), ("libelf", LIBELF),
                           ("decoder_header", HEADER), ("control_target", CONTROL_TARGET),
                           ("protocol_design", PROTOCOL_DESIGN), ("policy_review", POLICY_REVIEW),
                           ("scanner", SCANNER), ("lifecycle", LIFECYCLE),
                           ("reducer_tests", REDUCER_TESTS),
                           ("reducer", REDUCER),
                           ("scanner_tests", SCANNER_TESTS), ("lifecycle_tests", LIFECYCLE_TESTS),
                           ("replay_tool", REPLAY_TOOL), ("replay_receipt", REPLAY_RECEIPT),
                           ("scanner_controls", SCANNER_CONTROLS), ("raw_test_review", RAW_TEST_REVIEW),
                           ("control_command_receipt", CONTROL_COMMAND_RECEIPT),
                           ("control_test", CONTROL_TEST), ("control_audit", CONTROL_AUDIT),
                           ("control_equivalence", CONTROL_EQUIVALENCE),
                           ("focused_decode", FOCUSED_DECODER),
                           ("retained_reducer_0977", RETAINED_REDUCER_0977),
                           ("retained_reducer_0973", RETAINED_REDUCER_0973),
                           ("retained_protocol_0973", RETAINED_PROTOCOL_0973),
                           ("retained_protocol_0971", RETAINED_PROTOCOL_0971),
                           ("retained_protocol_0969", RETAINED_PROTOCOL_0969),
                           ("retained_reducer_0978_v1", RETAINED_REDUCER_0978_V1)):
        check_binding(plan.get("dependencies", {}).get(name), f"dependencies.{name}", expected)
    retained_dependency_paths = (
        ("retained_reducer_0977", RETAINED_REDUCER_0977),
        ("retained_reducer_0973", RETAINED_REDUCER_0973),
        ("retained_protocol_0973", RETAINED_PROTOCOL_0973),
        ("retained_protocol_0971", RETAINED_PROTOCOL_0971),
        ("retained_protocol_0969", RETAINED_PROTOCOL_0969),
        ("retained_reducer_0978_v1", RETAINED_REDUCER_0978_V1),
    )
    for path, expected_sha in RETAINED_REDUCER_LINEAGE_SHA256.items():
        name = next(name for name, candidate in retained_dependency_paths if candidate == path)
        if plan["dependencies"][name].get("sha256") != expected_sha:
            raise BindingError(f"{name} SHA is not the sealed immutable dependency")
    scanner_controls = strict_load(SCANNER_CONTROLS)
    if scanner_controls.get("schema") != "litchi.performance.0978.raw_stack_controls.v2" or scanner_controls.get("status") != "pass":
        raise BindingError("raw-stack scanner control receipt is not the passing v2 receipt")
    if scanner_controls.get("scanner") != plan["dependencies"]["scanner"]:
        raise BindingError("raw-stack scanner control receipt is bound to a different scanner")
    if scanner_controls.get("replay_tool") != plan["dependencies"]["replay_tool"]:
        raise BindingError("raw-stack scanner control receipt is bound to a different replay tool")
    if scanner_controls.get("retained_replay", {}).get("commands") is None:
        raise BindingError("raw-stack scanner control receipt lacks replay command custody")
    raw_test_review = strict_load(RAW_TEST_REVIEW)
    if raw_test_review.get("schema") != "litchi.performance.0978.independent_raw_stack_tests.v1" or raw_test_review.get("status") != "pass_independent_offline_mutation_controls":
        raise BindingError("independent raw-stack test review is not passing")
    if raw_test_review.get("source_binding", {}).get("sha256") != plan["dependencies"]["scanner"].get("sha256"):
        raise BindingError("independent raw-stack review is bound to a different scanner")
    control_command = strict_load(CONTROL_COMMAND_RECEIPT)
    if not isinstance(control_command.get("argv"), list) or "--control" not in control_command["argv"] or "--run" not in control_command["argv"]:
        raise BindingError("retained control command receipt does not prove control boundary")
    control_audit = strict_load(CONTROL_AUDIT)
    if control_audit.get("status") != "pass" or control_audit.get("mode") != "control":
        raise BindingError("retained control audit is not passing")
    control_equivalence = strict_load(CONTROL_EQUIVALENCE)
    if control_equivalence.get("schema") != "litchi.performance.0978.control_boundary_equivalence.v5" or control_equivalence.get("status") != "pass":
        raise BindingError("0978 retained-control equivalence receipt is not passing v5")
    if control_equivalence.get("native_workload_started") is not False or control_equivalence.get("build_started") is not False:
        raise BindingError("0978 control equivalence receipt has unsafe provenance")
    if control_equivalence.get("checks", {}).get("all_checks_passed") is not True:
        raise BindingError("0978 control equivalence checks did not all pass")
    candidate_runner = control_equivalence.get("candidate_runner", {})
    helper_runner = (plan.get("helpers") or {}).get("runner", {})
    if candidate_runner.get("path") != helper_runner.get("path") or candidate_runner.get("sha256") != helper_runner.get("sha256"):
        raise BindingError("0978 control equivalence is bound to a different runner")
    for name, expected in (("perf", PERF), ("taskset", TASKSET), ("python", PYTHON), ("launcher", LAUNCHER)):
        top = plan.get(name)
        dep = plan.get("dependencies", {}).get(name)
        if not isinstance(top, dict) or top.get("path") != dep.get("path") or top.get("sha256") != dep.get("sha256"):
            raise BindingError(f"top-level {name} binding does not equal sealed dependency")
    if plan["dependencies"]["bridge"].get("scope") != "retained 0977 bridge; exact native handoff and command boundary":
        raise BindingError("retained bridge scope missing")
    if plan["dependencies"]["bridge_lineage"].get("scope") != "handoff lineage; 0977 bridge retained byte-for-byte":
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
    if control.get("reuse_boundary") != "exact 0977 command/affinity boundary; independently command-self-tested before GO":
        raise BindingError("harmless control reuse boundary is not explicitly proven")
    if control.get("equivalence_receipt") != plan["dependencies"]["control_equivalence"]:
        raise BindingError("harmless control equivalence receipt is not bound")
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
        raise BindingError("0978 inherited workload cell must remain docfloat/e1/default")
    if workload.get("target_argv_template") != [
        "<binary>", "--case", "docfloat", "--input", "<fixture>", "--edits", "1",
        "--retention", "default", "--warmup", "5", "--samples", "80",
    ]:
        raise BindingError("workload argv template drift")
    child_gate = plan.get("child_gate")
    expected_child_gate = {
        "all_children": True, "raw_lost_sum": 0, "terminal_lost_samples": 0,
        "min_marker_samples": 100, "marker_period_share_positive": True,
        "structural_conservation": True, "valid_empty_is_unknown": True,
        "unknown_period_cutoff": None, "on_failure": "stop_without_retry",
    }
    if child_gate != expected_child_gate:
        raise BindingError("child gate contract drift")
    decoder_interface = plan.get("decoder_interface")
    if not isinstance(decoder_interface, dict):
        raise BindingError("decoder interface contract is missing")
    if decoder_interface.get("function") != "decode_child(plan_path, child_dir)":
        raise BindingError("decoder entry-point contract drift")
    if decoder_interface.get("artifact_bindings") != list(DECODER_CANONICAL_ARTIFACTS):
        raise BindingError("decoder canonical artifact binding set drift")
    expected_decoder_files = list(DECODER_CANONICAL_FILENAMES + ("<child>.target-layout.json", "<child>.decoder.json") + DECODER_OPTIONAL_FILENAMES)
    if decoder_interface.get("filenames") != expected_decoder_files:
        raise BindingError("decoder canonical filename contract drift")
    if decoder_interface.get("decoder_owned_receipts") != [
        "<child>.perf-header.command.json", "<child>.terminal-report.command.json",
        "<child>.perf-evlist-v.command.json", "<child>.perf-mmap-dump.command.json",
        "<child>.corrected.command.json", "<child>.terminal-summary.json", "<child>.loss.json",
    ]:
        raise BindingError("decoder offline command receipt contract drift")
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
    # This receipt is retained 0977 evidence.  The new decoder/scanner/helper
    # bytes are independently bound below; requiring this old receipt to name
    # the fresh decoder would silently rewrite historical evidence.
    helpers = plan.get("helpers") or {}
    if set(helpers) != set(helper_names()):
        raise BindingError("helper binding set drift")
    if (helpers.get("decoder", {}).get("path") != str(FOCUSED_DECODER.resolve())
            or helpers.get("decoder", {}).get("sha256") != FOCUSED_DECODER_SHA256):
        raise BindingError("focused decoder helper is not the sealed v3 source")
    focused_decode = plan.get("focused_decode")
    if (not isinstance(focused_decode, dict)
            or focused_decode.get("path") != helpers.get("decoder", {}).get("path")
            or focused_decode.get("sha256") != helpers.get("decoder", {}).get("sha256")):
        raise BindingError("focused decoder plan binding does not equal sealed decoder helper")
    seal_row = plan.get("helper_seal")
    seal_path = check_binding(seal_row, "helper_seal")
    seal = strict_load(seal_path)
    if seal.get("schema") != HELPER_SCHEMA or seal.get("status") != "sealed" or seal.get("helpers") != helpers:
        raise BindingError("helper seal mismatch")
    retained_lineage = seal.get("retained_lineage")
    if not isinstance(retained_lineage, dict):
        raise BindingError("helper seal retained reducer lineage is missing")
    retained_lineage_paths = {
        "base_0977_reducer": RETAINED_REDUCER_0977,
        "base_0973_reducer": RETAINED_REDUCER_0973,
        "base_0973_protocol_common": RETAINED_PROTOCOL_0973,
        "base_0971_protocol_common": RETAINED_PROTOCOL_0971,
        "base_0969_protocol_common": RETAINED_PROTOCOL_0969,
        "base_0978_reducer_v1": RETAINED_REDUCER_0978_V1,
    }
    for name, expected_path in retained_lineage_paths.items():
        row = retained_lineage.get(name)
        check_binding(row, f"helper_seal.retained_lineage.{name}", expected_path)
        if row.get("sha256") != RETAINED_REDUCER_LINEAGE_SHA256[expected_path]:
            raise BindingError(f"helper seal retained lineage SHA drift: {name}")
    if plan.get("helper_seal_sha256") != seal_row.get("sha256"):
        raise BindingError("helper seal digest is not bound")
    for name, row in helpers.items():
        check_helper_binding(row, f"helper.{name}")
    for name, dependency_name in (("bridge", "bridge"), ("scanner", "scanner"), ("lifecycle", "lifecycle"), ("reducer", "reducer")):
        row = helpers.get(name)
        dependency_row = plan["dependencies"][dependency_name]
        if not isinstance(row, dict) or row.get("path") != dependency_row.get("path") or row.get("sha256") != dependency_row.get("sha256"):
            raise BindingError(f"helper.{name} does not equal sealed dependency")
    if helpers.get("policy", {}).get("path") != plan.get("policy_review", {}).get("path") or helpers.get("policy", {}).get("sha256") != plan.get("policy_review", {}).get("sha256"):
        raise BindingError("helper.policy is not the reviewed policy receipt")
    outputs = plan.get("outputs") or {}
    capture_root = Path(outputs.get("capture_root", "")).resolve()
    analysis_root = Path(outputs.get("analysis_root", "")).resolve()
    if capture_root != (ROOT / "capture/instruction-profile-v1").resolve() or analysis_root != (ROOT / "analysis").resolve():
        raise BindingError("fresh output roots drifted")
    return {"status": "pass", "schema": PLAN_SCHEMA, "plan": str(plan_path or plan["plan_path"]), "helper_count": len(helpers)}


def validate_plan(path: Path) -> dict[str, Any]:
    return validate_seal(strict_load(path), path)
