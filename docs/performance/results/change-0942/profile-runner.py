#!/usr/bin/env python3
"""Prepare and capture the 0942 native CFB attribution packet.

The packet is deliberately a profile, rather than a benchmark qualification.
It binds the current source snapshot, the harness overlay, one already-built
frame-pointer/debug binary, the local ``perf`` contract, and the 0917 report
validator.  The command has no build mode.  ``write-plan`` only writes an
external prospective plan; ``real-preflight`` and ``capture`` refuse to run
unless root has supplied an authorized manifest and an explicit approval
environment variable.

Every child writes its terminal receipt before any report, RSS, or perf-script
data is parsed.  A failed child is retained and the schedule advances without
retry.  Raw ``perf.data`` and perf-script output are always outside Git.  The
reducer reports marker/frame coverage and correctness identities only; it does
not calculate a latency metric or an adoption decision.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
import time
from typing import Any, Iterable


REPO_ROOT = Path("/home/zhuhe/code/litchi").resolve()
HERE = Path(__file__).resolve()
ARTIFACT_ROOT = Path(
    os.environ.get("LITCHI_PERF_0942_ROOT", "/home/zhuhe/litchi-performance-0942")
).resolve()
SOURCE_ROOT = (ARTIFACT_ROOT / "source-v2").resolve()
PLAN_PATH = ARTIFACT_ROOT / "plans/cfb-profile-plan-v5.json"
MANIFEST_PATH = ARTIFACT_ROOT / "manifests/cfb-profile-manifest-v5.json"
DRAFT_MANIFEST_PATH = ARTIFACT_ROOT / "manifests/cfb-profile-manifest-v5-draft.json"
CHILD_ROOT = ARTIFACT_ROOT / "capture/formal-v5/children"
PREFLIGHT_ROOT = ARTIFACT_ROOT / "preflight/real-v5"
REPORT_ROOT = ARTIFACT_ROOT / "reports"
VALIDATOR_PATH = (REPO_ROOT / "docs/performance/results/change-0917/capture.py").resolve()
VALIDATOR_PROTOCOL = (
    REPO_ROOT / "docs/performance/results/change-0917/native-protocol.json"
).resolve()
HARNESS_ROOT = (REPO_ROOT / "tools/perf-cfb-profile").resolve()

SOURCE_REVISION = "73559c87c37898b675100b02985181a7764c66d2"
NATIVE_BINARY_NAME = "doc_retention_probe_0942_cfb_profile"
NATIVE_BINARY_PATH = (
    ARTIFACT_ROOT / "target-profile-v2/release" / NATIVE_BINARY_NAME
).resolve()
BUILD_MANIFEST = SOURCE_ROOT / "tools/perf-cfb-profile/Cargo.toml"
BUILD_LOCK = SOURCE_ROOT / "tools/perf-cfb-profile/Cargo.lock"
BUILD_TARGET = (ARTIFACT_ROOT / "target-profile-v2").resolve()
BUILD_RECEIPT = ARTIFACT_ROOT / "builds/profile-v3/terminal.json"
PLAN_V1_HISTORY = ARTIFACT_ROOT / "plans/cfb-profile-plan-v1.json"
PLAN_V2_HISTORY = ARTIFACT_ROOT / "plans/cfb-profile-plan-v2.json"
PLAN_V3_HISTORY = ARTIFACT_ROOT / "plans/cfb-profile-plan-v3.json"
PLAN_V4_HISTORY = ARTIFACT_ROOT / "plans/cfb-profile-plan-v4.json"
PERF_RECORD_HELP = ARTIFACT_ROOT / "contracts/perf-record-help-v1.txt"
PERF_SCRIPT_HELP = ARTIFACT_ROOT / "contracts/perf-script-help-v2.txt"
NM_RECEIPT = ARTIFACT_ROOT / "contracts/profile-binary-nm-v2.txt"

PERF = Path("/usr/bin/perf").resolve()
TASKSET = Path("/usr/bin/taskset").resolve()
TIME = Path("/usr/bin/time").resolve()
ADDR2LINE = Path("/usr/bin/addr2line").resolve()
NM = Path("/usr/bin/x86_64-linux-gnu-nm").resolve()
PERF_FIELDS = ("comm", "pid", "tid", "time", "event", "ip", "sym", "dso", "period")
PERF_FIELDS_TEXT = ",".join(PERF_FIELDS)
PERF_EVENT = "cycles:u"
PERF_FREQUENCY = 997
PERF_CALL_GRAPH = "fp"
CPU = 14
SAMPLES = 200
WARMUPS = 0
DEADLINE_SECONDS = 900
MARKER_SYMBOL = "profiled_measured_case"
# The bound profiling binary is a PIE.  perf's ``ip`` may be a runtime address,
# while addr2line consumes the ELF-relative address.  This value is pinned to
# the bound binary SHA and is used only as a deterministic fallback when the
# runtime address does not resolve.
STATIC_MARKER_ADDRESS = "0xdf130"
STATIC_MARKER_END = "0xdf1c0"
STATIC_MARKER_SYMBOL = (
    "doc_retention_probe_0942_cfb_profile::profiled_measured_case"
)
PROJECT_OWNERS = ("cfb", "ole-common", "doc", "core", "other")
SELECTED_FUNCTIONS = (
    "profiled_measured_case",
    "run_case",
    "Snapshot::open_bounded",
    "Edit::replace_paragraph",
    "Edit::commit",
)
TOP_N = 25
REPORT_SCHEMA = "litchi.performance.0917.doc_retention.v1"
FIXTURES: dict[str, dict[str, object]] = {
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
EDIT_COUNTS = (1, 2, 8, 32)
RETENTION = "default"
CAPTURE_ENV = {
    "PATH": "/usr/bin:/bin",
    "LC_ALL": "C",
    "LANG": "C",
    "TZ": "UTC",
    "RUST_BACKTRACE": "0",
    "PERF_PAGER": "cat",
    "PAGER": "cat",
}
APPROVAL_ENV = "LITCHI_PERF_0942_CAPTURE_APPROVED"
APPROVAL_VALUE = "1"
HEX = set("0123456789abcdef")
MAX_JSON_BYTES = 1 << 20


class RunnerError(RuntimeError):
    """A frozen custody, schedule, or semantic contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RunnerError(message)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256_path(path: Path) -> str:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path, *, packet_root: Path | None = None) -> dict[str, object]:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    rendered: str
    if packet_root is None:
        rendered = str(path)
    else:
        try:
            rendered = path.relative_to(packet_root.resolve()).as_posix()
        except ValueError as error:
            raise RunnerError(f"artifact is outside packet root: {path}") from error
    return {"path": rendered, "bytes": path.stat().st_size, "sha256": sha256_path(path)}


def packet_path(item: object, packet_root: Path) -> Path:
    require(isinstance(item, dict), "packet descriptor is not an object")
    value = item.get("path")
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            "packet descriptor path is not relative")
    path = (packet_root / value).resolve()
    require(path.is_relative_to(packet_root.resolve()), f"packet path escapes root: {value}")
    return path


def verify_descriptor(item: object, label: str, *, packet_root: Path | None = None,
                      allow_missing: bool = False) -> Path:
    require(isinstance(item, dict), f"{label}: descriptor missing")
    path_value = item.get("path")
    require(isinstance(path_value, str) and path_value, f"{label}: descriptor path missing")
    path = packet_path(item, packet_root) if packet_root is not None else Path(path_value).resolve()
    if allow_missing and not path.exists():
        return path
    require(path.is_file() and not path.is_symlink(), f"{label}: missing regular file: {path}")
    require(type(item.get("bytes")) is int and item["bytes"] >= 0,
            f"{label}: byte count missing")
    require(item["bytes"] == path.stat().st_size, f"{label}: byte count changed")
    expected = item.get("sha256")
    require(isinstance(expected, str) and len(expected) == 64 and set(expected) <= HEX,
            f"{label}: SHA-256 missing")
    require(sha256_path(path) == expected, f"{label}: SHA-256 changed")
    return path


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def read_json(path: Path, *, max_bytes: int | None = MAX_JSON_BYTES) -> Any:
    path = path.resolve()
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    if max_bytes is not None:
        require(path.stat().st_size < max_bytes, f"JSON exceeds limit: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"invalid JSON {path}: {error}") from error
    return value


def write_exclusive(path: Path, value: object) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    data = json_bytes(value)
    require(len(data) < MAX_JSON_BYTES, f"JSON output exceeds packet limit: {path}")
    try:
        with path.open("xb") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise RunnerError(f"refusing to overwrite retained file: {path}") from error


def write_text_exclusive(path: Path, data: str) -> None:
    path = path.resolve()
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8", newline="") as stream:
            stream.write(data)
            stream.flush()
            os.fsync(stream.fileno())
    except FileExistsError as error:
        raise RunnerError(f"refusing to overwrite retained file: {path}") from error


def harness_files() -> list[dict[str, object]]:
    paths = [
        HARNESS_ROOT / "Cargo.toml",
        HARNESS_ROOT / "Cargo.lock",
        HARNESS_ROOT / "README.md",
        HARNESS_ROOT / "src/alloc_metrics.rs",
        HARNESS_ROOT / "src/bin/alloc.rs",
        HARNESS_ROOT / "src/lib.rs",
        HARNESS_ROOT / "src/main.rs",
        HARNESS_ROOT / "src/oracle_0730.rs",
    ]
    return [descriptor(path) for path in paths]


def verify_marker_source(root: Path) -> None:
    source = root / "tools/perf-cfb-profile/src/lib.rs"
    text = source.read_text(encoding="utf-8")
    require(re.search(r"#\[inline\(never\)\]\s*fn\s+profiled_measured_case\b", text)
            is not None, "harness marker is not an inline-never function")
    require(text.count("profiled_measured_case(") == 3,
            "harness marker call count changed; measured-only scope is ambiguous")
    require(re.search(r"let\s+result\s*=\s*run_case\(source,\s*args,\s*retention\)\?;", text)
            is not None,
            "harness marker does not wrap run_case")


def schedule_rows() -> list[dict[str, object]]:
    cells = [
        {"case": case, "edits": edits, "retention": RETENTION}
        for case in ("docnohf", "docfloat")
        for edits in EDIT_COUNTS
    ]
    rows: list[dict[str, object]] = []
    for repetition, ordered in enumerate((cells, list(reversed(cells))), start=1):
        for position, cell in enumerate(ordered, start=1):
            case = str(cell["case"])
            edits = int(cell["edits"])
            rows.append({
                "id": f"profile-r{repetition:02d}-{case}-e{edits:02d}",
                "repetition": repetition,
                "position": position,
                "lane": "native_profile",
                "case": case,
                "edits": edits,
                "retention": RETENTION,
                "warmups": WARMUPS,
                "samples": SAMPLES,
                "cpu": CPU,
                "deadline_seconds": DEADLINE_SECONDS,
                "retry": False,
            })
    require(len(rows) == 16, "0942 formal schedule is not 16 children")
    return rows


def preflight_spec() -> dict[str, object]:
    return {
        "id": "real-preflight-docfloat-e01",
        "lane": "native_profile",
        "case": "docfloat",
        "edits": 1,
        "retention": RETENTION,
        "warmups": WARMUPS,
        "samples": SAMPLES,
        "cpu": CPU,
        "deadline_seconds": DEADLINE_SECONDS,
        "retry": False,
        "counts_toward_formal_schedule": False,
    }


def build_command() -> list[str]:
    return [
        "/home/zhuhe/.cargo/bin/cargo", "+1.95.0", "build", "--offline", "--locked",
        "-j2", "--release", "--message-format=json-render-diagnostics",
        "--bin", NATIVE_BINARY_NAME, "--manifest-path", str(BUILD_MANIFEST),
        "--target-dir", str(BUILD_TARGET),
    ]


def build_environment() -> dict[str, object]:
    return {
        "allowlist": ["HOME", "CARGO_HOME", "RUSTUP_HOME", "PATH", "LC_ALL", "LANG", "TZ"],
        "explicit": {
            "CARGO_TERM_COLOR": "never",
            "RUSTFLAGS": "-Cforce-frame-pointers=yes",
            "CARGO_PROFILE_RELEASE_DEBUG": "2",
            "CARGO_INCREMENTAL": "0",
        },
        "encoded_rustflags_forbidden": True,
    }


def record_argv(binary: Path, fixture: Path, spec: dict[str, object], child: Path) -> list[str]:
    perf_data = child / "perf.data"
    return [
        str(TASKSET), "--cpu-list", str(CPU), str(PERF), "record", "-e", PERF_EVENT,
        f"-F{PERF_FREQUENCY}", "--call-graph", PERF_CALL_GRAPH,
        "--output", str(perf_data), "--", str(binary),
        "--case", str(spec["case"]), "--input", str(fixture),
        "--edits", str(spec["edits"]), "--retention", str(spec["retention"]),
        "--warmup", str(WARMUPS), "--samples", str(SAMPLES),
    ]


def script_argv(perf_data: Path) -> list[str]:
    return [
        str(PERF), "script", "--no-inline", "--show-lost-events",
        "-i", str(perf_data), "-F", PERF_FIELDS_TEXT,
    ]


def profile_contract() -> dict[str, object]:
    return {
        "event": PERF_EVENT,
        "frequency_hz": PERF_FREQUENCY,
        "fields": list(PERF_FIELDS),
        "header_format": "comm pid/tid time: period event:",
        "call_graph": PERF_CALL_GRAPH,
        "dwarf_stack_bytes": None,
        "call_graph_rationale": "force-frame-pointers build; 0928 DWARF decoder limitations; 0929 FP guidance",
        "max_stack_option": None,
        "max_stack_help_checked": True,
        "marker_symbol": MARKER_SYMBOL,
        "marker_static_address": STATIC_MARKER_ADDRESS,
        "marker_static_end": STATIC_MARKER_END,
        "marker_static_symbol": STATIC_MARKER_SYMBOL,
        "marker_static_dso": NATIVE_BINARY_NAME,
        "marker_symbol_match": "exact static demangled symbol, optionally followed by +0x offset",
        "marker_dso_match": "bound absolute ELF path or exact ELF basename",
        "marker_scope": "only timed and allocator measured lifecycle calls profiled_measured_case",
        "native_frame_required": True,
        "project_owner_buckets": list(PROJECT_OWNERS),
        "owner_assignment": "first project-owned frame in perf's leaf-first chain before the exact marker; other if no project owner",
        "selected_functions": list(SELECTED_FUNCTIONS),
        "top_n": TOP_N,
        "addr2line": {
            "tool": str(ADDR2LINE),
            "argv": [str(ADDR2LINE), "-e", "<binary>", "-f", "-C", "<marker-ip>"],
            "required": True,
        },
        "script_options": ["--no-inline", "--show-lost-events"],
        "summary_schema": "litchi.performance.0942.perf-script-summary.v5",
        "marker_coverage_is_diagnostic": True,
        "incomplete_frames_are_unknown_not_nonmarker": True,
        "no_marker_membership": "unknown",
        "malformed_nonframe_lines_rejected": True,
        "raw_retention": ["perf.data", "perf-script.txt"],
    }


def verify_perf_help(record_help: Path, script_help: Path) -> None:
    record_text = record_help.read_text(encoding="utf-8", errors="strict")
    script_text = script_help.read_text(encoding="utf-8", errors="strict")
    require("--call-graph" in record_text and "fp" in record_text,
            "perf record help lacks frame-pointer call-graph contract")
    require("--max-stack" not in record_text,
            "local perf unexpectedly exposes --max-stack; freeze an explicit review first")
    require("-F, --fields" in script_text and "period" in script_text,
            "perf script help lacks requested fields/period contract")
    require("--no-inline" in script_text and "--show-lost-events" in script_text,
            "perf script help lacks no-inline/lost-event contract")


def verify_static_marker_binding(binary: Path, nm_output: Path) -> None:
    """Bind the PIE fallback to the exact symbol range in the exact ELF."""
    text = nm_output.read_text(encoding="utf-8", errors="strict")
    marker_line = re.compile(
        rf"^(?P<address>[0-9a-fA-F]+)\s+\S\s+{re.escape(STATIC_MARKER_SYMBOL)}$"
    )
    marker_address: int | None = None
    text_addresses: list[int] = []
    for line in text.splitlines():
        parts = line.split(maxsplit=2)
        if len(parts) == 3 and re.fullmatch(r"[0-9a-fA-F]+", parts[0]):
            address = int(parts[0], 16)
            if parts[1].lower() not in {"u", "?"}:
                text_addresses.append(address)
        match = marker_line.fullmatch(line)
        if match is not None:
            marker_address = int(match.group("address"), 16)
    require(marker_address == int(STATIC_MARKER_ADDRESS, 16),
            "nm marker address changed for bound binary")
    following = sorted(address for address in text_addresses if address > marker_address)
    require(following and following[0] == int(STATIC_MARKER_END, 16),
            "nm marker range changed for bound binary")
    require(binary.name == NATIVE_BINARY_NAME, "nm DSO binding is not the profile binary")


def build_binding() -> dict[str, object]:
    require(BUILD_RECEIPT.is_file(), f"profile build terminal missing: {BUILD_RECEIPT}")
    terminal = read_json(BUILD_RECEIPT)
    require(isinstance(terminal, dict) and terminal.get("exit_code") == 0,
            "profile build terminal is not successful")
    require(terminal.get("source_root") == str(SOURCE_ROOT)
            and terminal.get("target_dir") == str(BUILD_TARGET)
            and terminal.get("binary_name") == NATIVE_BINARY_NAME,
            "profile build terminal binding changed")
    binary = descriptor(NATIVE_BINARY_PATH)
    return {
        "receipt": descriptor(BUILD_RECEIPT),
        "binary": binary,
        "binary_name": NATIVE_BINARY_NAME,
        "manifest": str(BUILD_MANIFEST),
        "lock": str(BUILD_LOCK),
        "target_dir": str(BUILD_TARGET),
        "command": build_command(),
        "environment": build_environment(),
    }


def plan_value() -> dict[str, object]:
    verify_perf_help(PERF_RECORD_HELP, PERF_SCRIPT_HELP)
    verify_marker_source(SOURCE_ROOT)
    verify_static_marker_binding(NATIVE_BINARY_PATH, NM_RECEIPT)
    return {
        "schema": "litchi.performance.0942.cfb-profile-plan.v5",
        "status": "prospective_only_no_build_no_capture",
        "status_note": "root-owned source/build receipts are bound; this plan starts no build, profiler, or workload",
        "created_utc": utc_now(),
        "repository": str(REPO_ROOT),
        "source_revision": SOURCE_REVISION,
        "source_root": str(SOURCE_ROOT),
        "source_seed_receipt": descriptor(ARTIFACT_ROOT / "receipts/source-v2.json"),
        "supersedes": [
            descriptor(PLAN_V1_HISTORY), descriptor(PLAN_V2_HISTORY),
            descriptor(PLAN_V3_HISTORY), descriptor(PLAN_V4_HISTORY),
        ],
        "harness_overlay": {
            "source_root_relative": "tools/perf-cfb-profile",
            "files": harness_files(),
            "marker": MARKER_SYMBOL,
            "marker_source_contract": {
                "inline_never": True,
                "source_call_count": 3,
                "wraps": "run_case",
            },
        },
        "runner": descriptor(HERE),
        "correctness_validator": {
            "capture_script": descriptor(VALIDATOR_PATH),
            "protocol": descriptor(VALIDATOR_PROTOCOL),
            "function": "validate_report",
            "lane": "native",
            "expected_warmups": WARMUPS,
            "expected_samples": SAMPLES,
            "required_schema": REPORT_SCHEMA,
        },
        "fixtures": [
            {
                "case": case,
                "path": str(SOURCE_ROOT / str(row["relative_path"])),
                "relative_path": row["relative_path"],
                "bytes": row["bytes"],
                "sha256": row["sha256"],
            }
            for case, row in FIXTURES.items()
        ],
        "build": {
            "started_by_this_plan": False,
            "manifest": str(BUILD_MANIFEST),
            "lock": str(BUILD_LOCK),
            "target_dir": str(BUILD_TARGET),
            "binary_name": NATIVE_BINARY_NAME,
            "binary_path": str(NATIVE_BINARY_PATH),
            "command": build_command(),
            "environment": build_environment(),
            "profiling_requirements": ["force-frame-pointers", "release-debug-2"],
            "binding_required_before_capture": True,
            "terminal": descriptor(BUILD_RECEIPT),
            "binary": descriptor(NATIVE_BINARY_PATH),
            "marker_static_address": STATIC_MARKER_ADDRESS,
            "marker_static_end": STATIC_MARKER_END,
            "marker_static_symbol": STATIC_MARKER_SYMBOL,
        },
        "tools": {
            "perf": str(PERF),
            "taskset": str(TASKSET),
            "time": str(TIME),
            "addr2line": str(ADDR2LINE),
            "nm": str(NM),
            "help_checks": [
                [str(PERF), "record", "--help"],
                [str(PERF), "script", "--help"],
            ],
            "help_receipts": {
                "record": descriptor(PERF_RECORD_HELP),
                "script": descriptor(PERF_SCRIPT_HELP),
            },
            "nm_receipt": {
                "binary": descriptor(NATIVE_BINARY_PATH),
                "command": [str(NM), "-an", "--demangle", str(NATIVE_BINARY_PATH)],
                "output": descriptor(NM_RECEIPT),
                "marker_symbol": STATIC_MARKER_SYMBOL,
                "marker_start": STATIC_MARKER_ADDRESS,
                "marker_end": STATIC_MARKER_END,
            },
            "record_argv_template": [
                str(TASKSET), "--cpu-list", str(CPU), str(PERF), "record", "-e", PERF_EVENT,
                f"-F{PERF_FREQUENCY}", "--call-graph", PERF_CALL_GRAPH,
                "--output", "<child>/perf.data", "--", "<binary>", "<harness-args>",
            ],
            "script_argv_template": [
                str(PERF), "script", "--no-inline", "--show-lost-events",
                "-i", "<child>/perf.data", "-F", PERF_FIELDS_TEXT,
            ],
            "profile_contract": profile_contract(),
        },
        "environment": CAPTURE_ENV,
        "schedule": {
            "serial": True,
            "cpu": CPU,
            "warmups": WARMUPS,
            "samples": SAMPLES,
            "deadline_seconds": DEADLINE_SECONDS,
            "retry": False,
            "formal_children": schedule_rows(),
            "real_preflight": preflight_spec(),
            "cell_order": "forward cells docnohf/docfloat × edits 1/2/8/32; repetition two reverses all eight cells",
        },
        "custody": {
            "terminal_before_parse": True,
            "failed_children_retained": True,
            "no_retry_or_replacement": True,
            "raw_outside_git": True,
            "max_rss_units": "KiB from /usr/bin/time -f %M",
            "no_summed_absolute_gauges": True,
        },
        "claim_scope": {
            "purpose": "native CFB operation attribution around the measured lifecycle",
            "latency_comparison": False,
            "adoption_gate": False,
            "timing_claim": False,
            "harness_overhead_caveat": "instrumented marker/barrier and profiling build can change execution; report timing_claim is retained for validator compatibility and is excluded from claims",
            "profile_metadata_required": ["perf event", "field order", "marker coverage", "native DSO coverage"],
        },
    }


def write_plan() -> None:
    require(not PLAN_PATH.exists(), f"refusing to replace existing plan: {PLAN_PATH}")
    value = plan_value()
    write_exclusive(PLAN_PATH, value)
    print(json.dumps({"status": value["status"], "plan": str(PLAN_PATH),
                      "sha256": sha256_path(PLAN_PATH)}, sort_keys=True))


def manifest_draft_value(plan: dict[str, object]) -> dict[str, object]:
    """Build an unapproved manifest for root's final binding review.

    The draft deliberately uses a distinct path and status.  The capture
    actions only accept ``MANIFEST_PATH`` with status ``authorized``; writing
    this object can therefore never authorize a child by accident.
    """
    expected_harness = plan["harness_overlay"]["files"]
    source_harness: list[dict[str, object]] = []
    for expected in expected_harness:
        require(isinstance(expected, dict), "plan harness descriptor malformed")
        relative = Path(str(expected["path"])).resolve().relative_to(HARNESS_ROOT)
        source_path = SOURCE_ROOT / "tools/perf-cfb-profile" / relative
        source_item = descriptor(source_path)
        require(source_item["bytes"] == expected["bytes"]
                and source_item["sha256"] == expected["sha256"],
                f"frozen source harness differs from plan: {relative}")
        source_harness.append(source_item)
    tools = plan["tools"]
    build = plan["build"]
    return {
        "schema": "litchi.performance.0942.cfb-profile-manifest.v5",
        "status": "prospective_only_unapproved",
        "authorization": {
            "approved": False,
            "required_status_for_capture": "authorized",
            "approval_environment": APPROVAL_ENV,
            "approval_value": APPROVAL_VALUE,
            "note": "root must review and create the separately named authorized manifest",
        },
        "created_utc": utc_now(),
        "plan": descriptor(PLAN_PATH),
        "source_revision": SOURCE_REVISION,
        "runner": descriptor(HERE),
        "correctness_validator": plan["correctness_validator"],
        "source": {
            "root": str(SOURCE_ROOT),
            "seed_receipt": plan["source_seed_receipt"],
            "harness_files": source_harness,
        },
        "fixtures": plan["fixtures"],
        "build": {
            "receipt": build["terminal"],
            "binary": build["binary"],
            "binary_name": build["binary_name"],
            "manifest": build["manifest"],
            "lock": build["lock"],
            "target_dir": build["target_dir"],
            "command": build["command"],
            "environment": build["environment"],
            "marker_static_address": build["marker_static_address"],
            "marker_static_end": build["marker_static_end"],
            "marker_static_symbol": build["marker_static_symbol"],
        },
        "tools": {
            "perf": descriptor(PERF),
            "taskset": descriptor(TASKSET),
            "time": descriptor(TIME),
            "addr2line": descriptor(ADDR2LINE),
            "nm": descriptor(NM),
            "help_receipts": tools["help_receipts"],
            "nm_receipt": tools["nm_receipt"],
            "profile_contract": tools["profile_contract"],
        },
        "environment": plan["environment"],
        "schedule": plan["schedule"],
    }


def write_manifest_draft() -> None:
    require(PLAN_PATH.is_file(), f"write the v5 plan first: {PLAN_PATH}")
    require(not DRAFT_MANIFEST_PATH.exists(),
            f"refusing to replace existing draft: {DRAFT_MANIFEST_PATH}")
    plan = load_plan()
    value = manifest_draft_value(plan)
    write_exclusive(DRAFT_MANIFEST_PATH, value)
    print(json.dumps({"status": value["status"], "manifest": str(DRAFT_MANIFEST_PATH),
                      "sha256": sha256_path(DRAFT_MANIFEST_PATH)}, sort_keys=True))


def load_plan() -> dict[str, object]:
    value = read_json(PLAN_PATH)
    require(isinstance(value, dict), "plan is not an object")
    require(value.get("schema") == "litchi.performance.0942.cfb-profile-plan.v5",
            "plan schema mismatch")
    require(value.get("source_revision") == SOURCE_REVISION, "plan source revision changed")
    require(value.get("status") == "prospective_only_no_build_no_capture",
            "plan status changed")
    schedule = value.get("schedule")
    require(isinstance(schedule, dict), "plan schedule missing")
    require(schedule.get("formal_children") == schedule_rows(),
            "formal schedule changed")
    require(schedule.get("real_preflight") == preflight_spec(),
            "real preflight changed")
    require(schedule.get("serial") is True, "schedule is not serial")
    require(schedule.get("retry") is False, "retry policy changed")
    return value


def load_manifest(plan: dict[str, object]) -> dict[str, object]:
    value = read_json(MANIFEST_PATH)
    require(isinstance(value, dict), "manifest is not an object")
    require(value.get("schema") == "litchi.performance.0942.cfb-profile-manifest.v5",
            "manifest schema mismatch")
    require(value.get("status") == "authorized", "manifest is not authorized")
    require(value.get("plan") == descriptor(PLAN_PATH), "manifest plan binding changed")
    require(value.get("source_revision") == SOURCE_REVISION, "manifest source revision changed")
    require(value.get("runner") == descriptor(HERE), "runner binding changed")
    validator = value.get("correctness_validator")
    require(isinstance(validator, dict), "validator binding missing")
    verify_descriptor(validator.get("capture_script"), "validator capture script")
    verify_descriptor(validator.get("protocol"), "validator protocol")
    require(validator.get("capture_script") == descriptor(VALIDATOR_PATH),
            "validator capture script changed")
    require(validator.get("protocol") == descriptor(VALIDATOR_PROTOCOL),
            "validator protocol changed")
    source = value.get("source")
    require(isinstance(source, dict), "source binding missing")
    require(source.get("root") == str(SOURCE_ROOT), "source root changed")
    require(source.get("seed_receipt") == plan["source_seed_receipt"],
            "source seed receipt binding changed")
    require(source.get("seed_receipt") is not None, "source seed receipt missing")
    verify_descriptor(source.get("seed_receipt"), "source seed receipt")
    actual_harness = source.get("harness_files")
    expected_harness = plan["harness_overlay"]["files"]
    require(isinstance(actual_harness, list) and len(actual_harness) == len(expected_harness),
            "harness overlay binding changed")
    for expected, item in zip(expected_harness, actual_harness, strict=True):
        require(isinstance(expected, dict) and isinstance(item, dict),
                "harness descriptor malformed")
        expected_relative = Path(str(expected["path"])).resolve().relative_to(HARNESS_ROOT)
        actual_relative = Path(str(item["path"])).resolve().relative_to(SOURCE_ROOT / "tools/perf-cfb-profile")
        require(expected_relative == actual_relative
                and expected.get("bytes") == item.get("bytes")
                and expected.get("sha256") == item.get("sha256"),
                "harness overlay file changed")
        verify_descriptor(item, "harness overlay file")
    verify_marker_source(SOURCE_ROOT)
    fixtures = value.get("fixtures")
    require(isinstance(fixtures, list) and len(fixtures) == 2, "fixture bindings missing")
    expected_fixtures = []
    for case, row in FIXTURES.items():
        path = SOURCE_ROOT / str(row["relative_path"])
        expected_fixtures.append({
            "case": case,
            "path": str(path),
            "relative_path": row["relative_path"],
            "bytes": row["bytes"],
            "sha256": row["sha256"],
        })
        require(path.is_file() and not path.is_symlink(), f"fixture missing: {path}")
        require(path.stat().st_size == row["bytes"] and sha256_path(path) == row["sha256"],
                f"fixture identity changed: {case}")
    require(fixtures == expected_fixtures, "fixture bindings changed")
    build = value.get("build")
    require(isinstance(build, dict), "build binding missing")
    require(build.get("binary_name") == NATIVE_BINARY_NAME
            and build.get("manifest") == str(BUILD_MANIFEST)
            and build.get("target_dir") == str(BUILD_TARGET)
            and build.get("marker_static_address") == STATIC_MARKER_ADDRESS,
            "build identity changed")
    require(build.get("receipt") == plan["build"]["terminal"],
            "build terminal binding changed")
    require(build.get("binary") == plan["build"]["binary"],
            "profile binary binding changed")
    verify_descriptor(build.get("receipt"), "build terminal receipt")
    receipt = read_json(Path(str(build["receipt"]["path"])))
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            "build terminal is not successful")
    binary_path = verify_descriptor(build.get("binary"), "profile binary")
    require(binary_path == NATIVE_BINARY_PATH and os.access(binary_path, os.X_OK),
            "profile binary path or executable bit changed")
    require(build.get("command") == build_command(), "build command changed")
    require(build.get("environment") == build_environment(), "build environment changed")
    tools = value.get("tools")
    require(isinstance(tools, dict), "tool binding missing")
    for key, expected in (("perf", PERF), ("taskset", TASKSET), ("time", TIME),
                          ("addr2line", ADDR2LINE), ("nm", NM)):
        tool_path = verify_descriptor(tools.get(key), f"tool {key}")
        require(tool_path == expected, f"tool path changed: {key}")
    help_receipts = tools.get("help_receipts")
    require(isinstance(help_receipts, dict), "perf help receipts missing")
    verify_perf_help(
        verify_descriptor(help_receipts.get("record"), "perf record help"),
        verify_descriptor(help_receipts.get("script"), "perf script help"),
    )
    require(help_receipts == plan["tools"]["help_receipts"],
            "perf help receipts changed")
    nm_receipt = tools.get("nm_receipt")
    require(nm_receipt == plan["tools"]["nm_receipt"], "nm receipt binding changed")
    require(nm_receipt["binary"] == plan["build"]["binary"],
            "nm binary does not match profile binary")
    verify_descriptor(nm_receipt["binary"], "nm binary binding")
    nm_output = verify_descriptor(nm_receipt["output"], "nm output")
    require(nm_receipt["command"] == [str(NM), "-an", "--demangle", str(NATIVE_BINARY_PATH)],
            "nm command changed")
    verify_static_marker_binding(NATIVE_BINARY_PATH, nm_output)
    require(tools.get("profile_contract") == profile_contract(), "profile contract changed")
    require(value.get("environment") == CAPTURE_ENV, "capture environment changed")
    require(value.get("schedule") == plan["schedule"], "manifest schedule changed")
    return value


def require_approval() -> None:
    require(os.environ.get(APPROVAL_ENV) == APPROVAL_VALUE,
            f"{APPROVAL_ENV}=1 is required for real capture")


def fixture_path(case: str) -> Path:
    require(case in FIXTURES, f"unknown fixture: {case}")
    path = SOURCE_ROOT / str(FIXTURES[case]["relative_path"])
    require(path.is_file() and not path.is_symlink(), f"fixture missing: {path}")
    return path


def child_root_for(spec: dict[str, object], *, preflight: bool) -> Path:
    root = PREFLIGHT_ROOT if preflight else CHILD_ROOT
    return root / str(spec["id"])


def run_process(argv: list[str], *, cwd: Path, stdout_path: Path, stderr_path: Path,
                timeout: int, environment: dict[str, str]) -> tuple[int, bool]:
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        try:
            process = subprocess.Popen(
                argv,
                cwd=str(cwd),
                env=environment,
                stdin=subprocess.DEVNULL,
                stdout=stdout,
                stderr=stderr,
                start_new_session=True,
            )
        except OSError as error:
            stderr.write(f"process spawn failed: {error}\n")
            stderr.flush()
            return 127, False
        timed_out = False
        try:
            exit_code = process.wait(timeout=timeout)
        except subprocess.TimeoutExpired:
            timed_out = True
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except ProcessLookupError:
                pass
            exit_code = process.wait()
    return exit_code, timed_out


def parse_rss(path: Path) -> int:
    text = path.read_text(encoding="ascii")
    require(re.fullmatch(r"[0-9]+\n?", text) is not None,
            f"RSS sidecar is not a single KiB integer: {path}")
    value = int(text.strip())
    require(value > 0, f"RSS sidecar is non-positive: {path}")
    return value


def report_from_stdout(path: Path) -> dict[str, object]:
    text = path.read_text(encoding="utf-8")
    candidates = [line.strip() for line in text.splitlines() if line.strip()]
    require(candidates, f"probe stdout is empty: {path}")
    decoded: list[dict[str, object]] = []
    for line in candidates:
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        require(isinstance(value, dict), f"probe stdout JSON is not an object: {path}")
        decoded.append(value)
    require(len(decoded) == 1, f"probe stdout did not contain exactly one JSON report: {path}")
    return decoded[0]


def load_validator() -> Any:
    require(sha256_path(VALIDATOR_PATH) == descriptor(VALIDATOR_PATH)["sha256"],
            "validator changed while loading")
    name = "litchi_0942_0917_capture_validator"
    spec = importlib.util.spec_from_file_location(name, VALIDATOR_PATH)
    require(spec is not None and spec.loader is not None, "cannot load 0917 validator")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def validate_report(report: dict[str, object], row: dict[str, object]) -> dict[str, object]:
    validator = load_validator()
    case = str(row["case"])
    fixture = {
        "case": case,
        "path": str(FIXTURES[case]["relative_path"]),
        "path_abs": str(fixture_path(case)),
        "bytes": FIXTURES[case]["bytes"],
        "sha256": FIXTURES[case]["sha256"],
    }
    spec = {
        "label": str(row["id"]),
        "lane": "native",
        "case": case,
        "edits": int(row["edits"]),
        "retention": RETENTION,
    }
    result = validator.validate_report(
        report, spec, fixture, "native", expected_warmups=WARMUPS,
        expected_samples=SAMPLES,
    )
    require(report.get("schema") == REPORT_SCHEMA, f"{row['id']}: report schema changed")
    require(report.get("timing_claim") is True
            and report.get("allocator_instrumented") is False
            and report.get("allocator_timing_claim") is False,
            f"{row['id']}: native profile lane identity failed")
    require(report.get("input") in {str(fixture_path(case)), fixture["path"]},
            f"{row['id']}: report input identity failed")
    return {
        "case": case,
        "edits": int(row["edits"]),
        "retention": RETENTION,
        "report_sha256": hashlib.sha256(json_bytes(report)).hexdigest(),
        "output_sha256": report["output_sha256"],
        "output_bytes": report["output_bytes"],
        "output_capacity": report["output_capacity"],
        "samples": len(report["samples"]),
        "correctness": report["correctness"],
    }


# With ``-F comm,pid,tid,time,event,ip,sym,dso,period`` this perf version emits
# ``comm pid/tid time: period event:``.  In particular, period precedes event
# and pid/tid is one slash-separated field; treating the line like the older
# ``pid tid ... event period`` format silently turns event text into a period.
HEADER_RE = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>[0-9]+)/(?P<tid>[0-9]+)"
    r"\s+(?P<time>[0-9]+(?:\.[0-9]+)?)\s*:\s*"
    r"(?P<period>[0-9]+)\s+(?P<event>\S+)"
)
ADDRESS_RE = re.compile(r"^(?:0x)?[0-9a-fA-F]+$")
FRAME_WITH_DSO_RE = re.compile(
    r"^(?P<ip>(?:0x)?[0-9a-fA-F]+)\s+(?P<symbol>.*?)\s+"
    r"\((?P<dso>[^)]*)\)(?:\s+(?P<period>[0-9]+))?\s*$"
)
MARKER_RE = re.compile(r"(?<![A-Za-z0-9_])profiled_measured_case(?![A-Za-z0-9_])")
MARKER_SYMBOL_EXACT_RE = re.compile(
    rf"^{re.escape(STATIC_MARKER_SYMBOL)}(?:\+0x[0-9a-fA-F]+)?$"
)
UNKNOWN_RE = re.compile(r"^(?:\?\?|\[unknown\]|unknown)$", re.IGNORECASE)


def _period_from_tokens(tokens: list[str]) -> int | None:
    if not tokens:
        return None
    candidate = tokens[-1].rstrip(":,)")
    if re.fullmatch(r"[0-9]+", candidate) is None:
        return None
    return int(candidate)


def _frame_from_line(line: str) -> dict[str, object] | None:
    stripped = line.strip()
    if not stripped or stripped.startswith("#"):
        return None
    tokens = stripped.split()
    if len(tokens) < 2 or ADDRESS_RE.fullmatch(tokens[0]) is None:
        return None
    structured = FRAME_WITH_DSO_RE.match(stripped)
    if structured is not None:
        ip = structured.group("ip")
        symbol = structured.group("symbol").strip()
        dso = structured.group("dso") or "[unknown]"
        frame_period = (int(structured.group("period"))
                         if structured.group("period") is not None else None)
    else:
        ip = tokens[0]
        # Keep the complete demangled symbol.  ``sym`` can contain spaces, so
        # taking tokens[1] would silently truncate the physical frame name.
        frame_period = _period_from_tokens(tokens[2:])
        dso = "[unknown]"
        dso_start = None
        for index, token in enumerate(tokens[2:], start=2):
            candidate = token.strip("()")
            if "/" in candidate or candidate.startswith("["):
                dso_start = index
                dso = candidate
                break
        symbol_tokens = tokens[1:dso_start] if dso_start is not None else tokens[1:]
        if frame_period is not None and symbol_tokens and symbol_tokens[-1].rstrip(":,)") == str(frame_period):
            symbol_tokens = symbol_tokens[:-1]
        symbol = " ".join(symbol_tokens).strip()
    return {
        "ip": ip,
        "symbol": symbol,
        "dso": dso,
        "period": frame_period,
        "inlined": dso.strip().lower() == "inlined",
        "unresolved": bool(UNKNOWN_RE.fullmatch(symbol)
                            or UNKNOWN_RE.fullmatch(dso)),
    }


def _owner_for_frame(frame: dict[str, object]) -> str | None:
    if bool(frame.get("inlined")):
        return None
    value = f"{frame.get('dso', '')} {frame.get('symbol', '')}".lower()
    if "litchi_cfb" in value or "litchi-cfb" in value:
        return "cfb"
    if "litchi_ole_common" in value or "litchi-ole-common" in value:
        return "ole-common"
    if "litchi_doc" in value or "litchi-doc" in value:
        return "doc"
    if "litchi_core" in value or "litchi-core" in value:
        return "core"
    if bool(frame.get("unresolved")):
        return None
    return "other"


def _marker_frame_is_bound(frame: dict[str, object], binary: Path) -> bool:
    symbol = str(frame.get("symbol", ""))
    dso = str(frame.get("dso", ""))
    # perf can print either the complete DSO path or just its basename.  A
    # different path with the same basename is not accepted because it could
    # be another build of the PIE.
    return (MARKER_SYMBOL_EXACT_RE.fullmatch(symbol) is not None
            and dso in {str(binary), binary.name})


def _add_metric(table: dict[str, dict[str, int]], key: str, period: int) -> None:
    row = table.setdefault(key, {"samples": 0, "period": 0})
    row["samples"] += 1
    row["period"] += period


def _rank_metrics(table: dict[str, dict[str, int]], denominator: dict[str, int],
                  *, label: str = "symbol") -> dict[str, object]:
    ranked = sorted(table.items(), key=lambda item: (-item[1]["period"],
                                                      -item[1]["samples"], item[0]))
    top = [{label: key, **values} for key, values in ranked[:TOP_N]]
    remainder = ranked[TOP_N:]
    return {
        "top": top,
        "top_n": TOP_N,
        "denominator": denominator,
        "remainder": {
            "symbols": len(remainder),
            "samples": sum(values["samples"] for _, values in remainder),
            "period": sum(values["period"] for _, values in remainder),
        },
    }


def _selected_match(selected: str, symbol: str) -> bool:
    if selected == MARKER_SYMBOL:
        return MARKER_RE.search(symbol) is not None
    return symbol == selected or symbol.endswith(f"::{selected}") or selected in symbol


def _process_profile_record(
    record: dict[str, object],
    state: dict[str, object],
    binary: Path,
) -> None:
    period = record["period"]
    require(type(period) is int and period > 0, "perf sample period is missing or non-positive")
    frames = record["frames"]
    require(isinstance(frames, list), "perf sample frames are malformed")
    state["all_samples"] += 1
    state["all_period"] += period
    unresolved_frames = [frame for frame in frames if frame.get("unresolved")]
    inline_frames = [frame for frame in frames if frame.get("inlined")]
    state["inline_frames"] += len(inline_frames)
    # ``--no-inline`` is part of the frozen export contract.  Keep an
    # unexpected inline row visible in the summary, but do not let its pseudo
    # DSO stand in for a physical owner or sampled leaf.
    physical_frames = [frame for frame in frames if not frame.get("inlined")]
    truncated = record.get("truncated") is True
    if truncated:
        state["truncated_samples"] += 1
        state["truncated_period"] += period
    marker_indices = [index for index, frame in enumerate(frames)
                      if MARKER_RE.search(str(frame.get("symbol", ""))) is not None]
    marker = bool(marker_indices)
    for marker_index in marker_indices:
        marker_frame = frames[marker_index]
        require(_marker_frame_is_bound(marker_frame, binary),
                "marker symbol or DSO is not bound to the profile binary")
        state["marker_native_dso_lines"] += 1
    fully_resolved = marker and not unresolved_frames and not truncated
    if marker:
        state["marker_samples"] += 1
        state["marker_period"] += period
        if fully_resolved:
            state["fully_resolved_marker_samples"] += 1
        if unresolved_frames:
            state["marker_samples_with_unresolved_frames"] += 1
        marker_index = marker_indices[0]
        # perf prints the sampled leaf first.  The nearest physical owner is
        # therefore the first named project DSO before the marker frame.
        for frame in frames[:marker_index]:
            owner = _owner_for_frame(frame)
            if owner in {"cfb", "ole-common", "doc", "core"}:
                state["owner_buckets"][owner]["samples"] += 1
                state["owner_buckets"][owner]["period"] += period
                break
        else:
            state["owner_buckets"]["other"]["samples"] += 1
            state["owner_buckets"]["other"]["period"] += period
        leaf = physical_frames[0] if physical_frames else None
        if isinstance(leaf, dict):
            _add_metric(state["leaf_symbols"], str(leaf.get("symbol", "[unknown]")), period)
        for selected in SELECTED_FUNCTIONS:
            if any(_selected_match(selected, str(frame.get("symbol", "")))
                   for frame in frames):
                _add_metric(state["inclusive_functions"], selected, period)
        marker_address = next((str(frame["ip"]) for frame in frames
                               if MARKER_RE.search(str(frame.get("symbol", "")))
                               and frame.get("ip")), None)
        if marker_address is not None and marker_address not in state["marker_addresses"]:
            state["marker_addresses"].append(marker_address)
    else:
        leaf = physical_frames[0] if physical_frames else None
        if isinstance(leaf, dict):
            _add_metric(state["nonmarker_leaf_symbols"],
                        str(leaf.get("symbol", "[unknown]")), period)
        if unresolved_frames or truncated or not physical_frames:
            state["unknown_samples"] += 1
            state["unknown_period"] += period
        # A printed stack without the marker does not prove that the measured
        # path was absent: perf may have sampled outside it, or the chain may
        # be incomplete.  Retain the observation separately and label its
        # membership unknown even when every printed frame has a name.
        state["no_marker_samples"] += 1
        state["no_marker_period"] += period
    if unresolved_frames:
        state["unresolved_samples"] += 1
        state["unresolved_period"] += period
        state["unresolved_frames"] += len(unresolved_frames)


def parse_perf_script(path: Path, binary: Path) -> dict[str, object]:
    """Stream one ``perf script -F`` export into disjoint sample accounting.

    A sample without the marker is retained as ``no_marker_observed`` with
    unknown membership; incomplete/unresolved stacks are counted separately.
    A missing marker is never treated as proof that the measured path did not
    execute.
    """
    state: dict[str, object] = {
        "all_samples": 0,
        "all_period": 0,
        "marker_samples": 0,
        "marker_period": 0,
        "fully_resolved_marker_samples": 0,
        "marker_samples_with_unresolved_frames": 0,
        "no_marker_samples": 0,
        "no_marker_period": 0,
        "unknown_samples": 0,
        "unknown_period": 0,
        "unresolved_samples": 0,
        "unresolved_period": 0,
        "unresolved_frames": 0,
        "inline_frames": 0,
        "marker_native_dso_lines": 0,
        "truncated_samples": 0,
        "truncated_period": 0,
        "lost_records": 0,
        "owner_buckets": {owner: {"samples": 0, "period": 0}
                          for owner in PROJECT_OWNERS},
        "leaf_symbols": {},
        "nonmarker_leaf_symbols": {},
        "inclusive_functions": {},
        "marker_addresses": [],
        "headers": 0,
        "event_counts": {},
    }
    current: dict[str, object] | None = None
    with path.open("r", encoding="utf-8") as stream:
        for line in stream:
            stripped = line.strip()
            if not stripped or stripped.startswith("#"):
                continue
            if "PERF_RECORD_LOST" in line:
                state["lost_records"] += 1
                continue
            if stripped.startswith("..."):
                require(current is not None,
                        f"truncation marker appeared before a sample: {path}")
                current["truncated"] = True
                continue
            match = HEADER_RE.match(line)
            if match is not None:
                if current is not None:
                    _process_profile_record(current, state, binary)
                event = match.group("event").rstrip(":")
                period = int(match.group("period"))
                require(period > 0,
                        "perf script header exposes a non-positive period")
                state["headers"] += 1
                state["event_counts"][event] = state["event_counts"].get(event, 0) + 1
                current = {"event": event, "period": period, "frames": []}
                continue
            if current is not None:
                frame = _frame_from_line(line)
                if frame is not None:
                    current["frames"].append(frame)
                else:
                    raise RunnerError(
                        f"unparsed non-frame perf-script line: {path}: {stripped}"
                    )
            else:
                raise RunnerError(
                    f"unparsed perf-script line outside a frame/header: {path}: {stripped}"
                )
        if current is not None:
            _process_profile_record(current, state, binary)
    headers = int(state["headers"])
    require(headers > 0, f"perf script has no sample headers: {path}")
    require(state["event_counts"].get(PERF_EVENT, 0) == headers,
            f"perf script event field mismatch: {path}")
    require(state["marker_samples"] > 0,
            f"perf script has no {MARKER_SYMBOL} frames: {path}")
    marker_dso_lines = int(state["marker_native_dso_lines"])
    # A marker-qualified sample is native only when every marker frame names
    # the exact bound executable (or its exact basename).  The raw line is
    # retained; this count is a custody check, not a substitute for the raw
    # stack.
    require(marker_dso_lines > 0, f"native marker/DSO coverage missing: {path}")
    require(state["all_samples"] == state["marker_samples"]
            + state["no_marker_samples"],
            f"perf sample partition is not disjoint: {path}")
    all_period = int(state["all_period"])
    require(all_period == int(state["marker_period"])
            + int(state["no_marker_period"]),
            f"perf period partition is not disjoint: {path}")
    marker_denominator = {"samples": int(state["marker_samples"]),
                          "period": int(state["marker_period"])}
    no_marker_denominator = {"samples": int(state["no_marker_samples"]),
                             "period": int(state["no_marker_period"])}
    return {
        "schema": "litchi.performance.0942.perf-script-summary.v5",
        "fields": list(PERF_FIELDS),
        "event": PERF_EVENT,
        "sample_accounting": {
            "all": {"samples": headers, "period": all_period},
            "marker_qualified": marker_denominator,
            "no_marker_observed": {
                **no_marker_denominator,
                "membership": "unknown",
                "reason": "marker_not_observed_does_not_prove_path_absence",
            },
            "nonmarker_observed": no_marker_denominator,
            "unknown_incomplete_or_unresolved": {
                "samples": int(state["unknown_samples"]),
                "period": int(state["unknown_period"]),
            },
            "unresolved": {
                "samples": int(state["unresolved_samples"]),
                "period": int(state["unresolved_period"]),
                "frames": int(state["unresolved_frames"]),
            },
            "inline_frames_skipped": int(state["inline_frames"]),
            "truncated_or_incomplete": {
                "samples": int(state["truncated_samples"]),
                "period": int(state["truncated_period"]),
            },
            "lost": {"records": int(state["lost_records"]), "period": None},
            "partition_ok": True,
        },
        "marker": {
            "symbol": MARKER_SYMBOL,
            "samples": int(state["marker_samples"]),
            "period": int(state["marker_period"]),
            "fully_resolved_samples": int(state["fully_resolved_marker_samples"]),
            "samples_with_unresolved_frames": int(state["marker_samples_with_unresolved_frames"]),
            "native_dso_frame_lines": marker_dso_lines,
            "addresses": list(state["marker_addresses"]),
        },
        "nearest_physical_owner_buckets": state["owner_buckets"],
        "ranked_nearest_physical_owners": _rank_metrics(
            state["owner_buckets"], marker_denominator, label="owner"
        ),
        "exclusive_leaf_symbols": _rank_metrics(state["leaf_symbols"], marker_denominator),
        "no_marker_leaf_symbols": _rank_metrics(state["nonmarker_leaf_symbols"],
                                                  no_marker_denominator),
        "inclusive_selected_functions": _rank_metrics(
            state["inclusive_functions"], marker_denominator, label="function"
        ),
        "marker_ip": state["marker_addresses"][0] if state["marker_addresses"] else None,
        "script_bytes": path.stat().st_size,
        "native_frame_coverage": True,
        "marker_coverage_is_diagnostic": True,
        "incomplete_frames_are_unknown_not_nonmarker": True,
    }


def resolve_marker_addr2line(binary: Path, profile: dict[str, object]) -> None:
    marker_ip = profile.get("marker_ip")
    require(isinstance(marker_ip, str) and marker_ip, "marker IP is missing")
    attempts: list[dict[str, object]] = []
    for address in (marker_ip, STATIC_MARKER_ADDRESS):
        command = [str(ADDR2LINE), "-e", str(binary), "-f", "-C", address]
        try:
            result = subprocess.run(
                command, cwd=str(SOURCE_ROOT), env=dict(CAPTURE_ENV),
                stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                text=True, check=False, timeout=30,
            )
        except (OSError, subprocess.SubprocessError) as error:
            raise RunnerError(f"addr2line failed: {error}") from error
        output = result.stdout.strip()
        attempt = {
            "argv": command,
            "exit_code": result.returncode,
            "output": output,
            "marker_resolved": result.returncode == 0 and MARKER_SYMBOL in output,
        }
        attempts.append(attempt)
        if attempt["marker_resolved"]:
            profile["addr2line"] = {
                "attempts": attempts,
                "selected_address": address,
                "marker_resolved": True,
            }
            return
    raise RunnerError(f"addr2line did not resolve {MARKER_SYMBOL}: {attempts!r}")


def require_profile_coverage(profile: dict[str, object], *, preflight: bool) -> None:
    marker = profile["marker"]
    minimum = 100 if preflight else 1
    require(marker["samples"] >= minimum,
            f"marker sample coverage below {'preflight' if preflight else 'formal'} minimum")
    require(marker["fully_resolved_samples"] > 0,
            "no fully resolved marker stack coverage")
    require(profile.get("addr2line", {}).get("marker_resolved") is True,
            "marker addr2line resolution missing")


def raw_descriptor(path: Path) -> dict[str, object]:
    return descriptor(path, packet_root=ARTIFACT_ROOT)


def common_child_start(spec: dict[str, object], child: Path, argv: list[str],
                       export_argv: list[str]) -> dict[str, object]:
    return {
        "schema": "litchi.performance.0942.child-start.v5",
        "id": spec["id"],
        "spec": spec,
        "source_revision": SOURCE_REVISION,
        "source_root": str(SOURCE_ROOT),
        "binary": str(NATIVE_BINARY_PATH),
        "fixture": str(fixture_path(str(spec["case"]))),
        "cwd": str(SOURCE_ROOT),
        "environment": CAPTURE_ENV,
        "record_argv": argv,
        "script_argv": export_argv,
        "profile_contract": profile_contract(),
        "child_root": str(child),
        "started_utc": utc_now(),
    }


def run_child(spec: dict[str, object], manifest: dict[str, object], *, preflight: bool) -> dict[str, object]:
    child = child_root_for(spec, preflight=preflight)
    require(not child.exists(), f"refusing to replace retained child: {child}")
    child.mkdir(parents=True)
    binary = Path(str(manifest["build"]["binary"]["path"])).resolve()
    fixture = fixture_path(str(spec["case"]))
    record = child / "perf.data"
    rss = child / "max-rss-kib.txt"
    stdout = child / "report.stdout"
    stderr = child / "record.stderr"
    script_output = child / "perf-script.txt"
    script_stderr = child / "script.stderr"
    record_terminal = child / "record-terminal.json"
    script_terminal = child / "script-terminal.json"
    receipt_path = child / "receipt.json"
    record_command = record_argv(binary, fixture, spec, child)
    exported_command = script_argv(record)
    write_exclusive(child / "start.json", common_child_start(
        spec, child, record_command, exported_command
    ))
    timed_command = [str(TIME), "-f", "%M", "-o", str(rss), "--", *record_command]
    started = utc_now()
    started_monotonic = time.monotonic()
    exit_code, timed_out = run_process(
        timed_command, cwd=SOURCE_ROOT, stdout_path=stdout, stderr_path=stderr,
        timeout=DEADLINE_SECONDS, environment=dict(CAPTURE_ENV),
    )
    ended = utc_now()
    elapsed = time.monotonic() - started_monotonic
    # This terminal is deliberately written before opening report.stdout,
    # max-rss-kib.txt, or perf.data for interpretation.
    write_exclusive(record_terminal, {
        "schema": "litchi.performance.0942.record-terminal.v5",
        "id": spec["id"],
        "record_argv": timed_command,
        "started_utc": started,
        "ended_utc": ended,
        "elapsed_seconds": elapsed,
        "exit_code": exit_code,
        "timed_out": timed_out,
        "terminal_written_before_parse": True,
    })
    status = "failed"
    failure: str | None = None
    report_summary: dict[str, object] | None = None
    profile_summary: dict[str, object] | None = None
    rss_value: int | None = None
    if exit_code != 0 or timed_out:
        failure = f"perf record failed with exit {exit_code} (timed_out={timed_out})"
    elif not record.is_file():
        failure = "perf record exited zero without perf.data"
    else:
        script_started = utc_now()
        script_exit, script_timeout = run_process(
            exported_command, cwd=SOURCE_ROOT, stdout_path=script_output,
            stderr_path=script_stderr, timeout=DEADLINE_SECONDS,
            environment=dict(CAPTURE_ENV),
        )
        write_exclusive(script_terminal, {
        "schema": "litchi.performance.0942.script-terminal.v5",
            "id": spec["id"],
            "script_argv": exported_command,
            "started_utc": script_started,
            "ended_utc": utc_now(),
            "exit_code": script_exit,
            "timed_out": script_timeout,
            "terminal_written_before_parse": True,
        })
        if script_exit != 0 or script_timeout:
            failure = f"perf script failed with exit {script_exit} (timed_out={script_timeout})"
        else:
            try:
                rss_value = parse_rss(rss)
                report = report_from_stdout(stdout)
                report_summary = validate_report(report, spec)
                profile_summary = parse_perf_script(script_output, binary)
                resolve_marker_addr2line(binary, profile_summary)
                require_profile_coverage(profile_summary, preflight=preflight)
                status = "passed"
            except (OSError, UnicodeError, RunnerError, ValueError) as error:
                failure = str(error)
    receipt: dict[str, object] = {
        "schema": "litchi.performance.0942.child-receipt.v5",
        "id": spec["id"],
        "spec": spec,
        "status": status,
        "failure": failure,
        "source_revision": SOURCE_REVISION,
        "source_root": str(SOURCE_ROOT),
        "binary": descriptor(binary),
        "fixture": {
            "case": spec["case"],
            "path": str(fixture),
            "bytes": FIXTURES[str(spec["case"])]["bytes"],
            "sha256": FIXTURES[str(spec["case"])]["sha256"],
        },
        "environment": CAPTURE_ENV,
        "record_command": timed_command,
        "script_command": exported_command,
        "record_terminal": raw_descriptor(record_terminal),
        "start": raw_descriptor(child / "start.json"),
        "report_stdout": raw_descriptor(stdout),
        "record_stderr": raw_descriptor(stderr),
        "rss_sidecar": raw_descriptor(rss) if rss.is_file() else None,
        "perf_data": raw_descriptor(record) if record.is_file() else None,
        "script_terminal": raw_descriptor(script_terminal) if script_terminal.is_file() else None,
        "perf_script": raw_descriptor(script_output) if script_output.is_file() else None,
        "script_stderr": raw_descriptor(script_stderr) if script_stderr.is_file() else None,
        "max_rss_kib": rss_value,
        "report": report_summary,
        "profile": profile_summary,
        "terminal_before_parse": True,
        "retry": False,
        "started_utc": started,
        "ended_utc": ended,
    }
    write_exclusive(receipt_path, receipt)
    return receipt


def run_capture(plan: dict[str, object], manifest: dict[str, object], *, preflight: bool) -> None:
    specs = [preflight_spec()] if preflight else schedule_rows()
    root = PREFLIGHT_ROOT if preflight else CHILD_ROOT
    root.mkdir(parents=True, exist_ok=True)
    results = []
    for spec in specs:
        results.append(run_child(spec, manifest, preflight=preflight))
    summary = {
        "schema": "litchi.performance.0942.capture-run.v5",
        "mode": "real-preflight" if preflight else "formal",
        "plan": descriptor(PLAN_PATH),
        "manifest": descriptor(MANIFEST_PATH),
        "children": len(results),
        "passed": sum(item.get("status") == "passed" for item in results),
        "failed": sum(item.get("status") != "passed" for item in results),
        "retry": False,
        "results": [{"id": item["id"], "status": item["status"],
                     "failure": item.get("failure")} for item in results],
        "created_utc": utc_now(),
    }
    path = REPORT_ROOT / ("real-preflight-run-v5.json" if preflight else "formal-run-v5.json")
    write_exclusive(path, summary)
    print(json.dumps(summary, sort_keys=True))


def load_packet_receipt(path: Path, packet_root: Path) -> dict[str, object]:
    value = read_json(path)
    require(isinstance(value, dict), f"receipt is not an object: {path}")
    require(value.get("schema") == "litchi.performance.0942.child-receipt.v5",
            f"receipt schema mismatch: {path}")
    # The record terminal is the first custody boundary.  Rehash it and inspect
    # its exit code before opening report, RSS, or perf-script artifacts.
    terminal_path = packet_path(value.get("record_terminal"), packet_root)
    terminal = read_json(terminal_path)
    require(terminal.get("terminal_written_before_parse") is True,
            f"terminal custody marker missing: {path}")
    require(terminal.get("exit_code") == 0 and terminal.get("timed_out") is False,
            f"failed record terminal: {path}")
    for key in ("start", "record_terminal", "report_stdout", "record_stderr", "rss_sidecar",
                "perf_data", "script_terminal", "perf_script", "script_stderr"):
        item = value.get(key)
        if item is None:
            continue
        verify_descriptor(item, f"{path.name}:{key}", packet_root=packet_root)
    script_terminal_item = value.get("script_terminal")
    require(script_terminal_item is not None, f"script terminal missing: {path}")
    script_terminal = read_json(packet_path(script_terminal_item, packet_root))
    require(script_terminal.get("terminal_written_before_parse") is True
            and script_terminal.get("exit_code") == 0
            and script_terminal.get("timed_out") is False,
            f"failed script terminal: {path}")
    return value


def replay_child(receipt: dict[str, object], packet_root: Path, row: dict[str, object],
                 binary: Path, *, preflight: bool) -> dict[str, object]:
    require(receipt.get("status") == "passed", f"child is not passed: {row['id']}")
    require(receipt.get("spec") == row, f"child spec changed: {row['id']}")
    report_path = packet_path(receipt["report_stdout"], packet_root)
    rss_path = packet_path(receipt["rss_sidecar"], packet_root)
    script_path = packet_path(receipt["perf_script"], packet_root)
    require(parse_rss(rss_path) == receipt.get("max_rss_kib"),
            f"RSS receipt metric changed: {row['id']}")
    report = report_from_stdout(report_path)
    report_summary = validate_report(report, row)
    profile_summary = parse_perf_script(script_path, binary)
    resolve_marker_addr2line(binary, profile_summary)
    require_profile_coverage(profile_summary, preflight=preflight)
    require(report_summary == receipt.get("report"),
            f"report summary changed: {row['id']}")
    require(profile_summary == receipt.get("profile"),
            f"profile summary changed: {row['id']}")
    return {"id": row["id"], "report": report_summary, "profile": profile_summary,
            "max_rss_kib": receipt["max_rss_kib"]}


def reduce_packet(plan: dict[str, object], manifest: dict[str, object]) -> None:
    preflight_receipt_path = PREFLIGHT_ROOT / str(preflight_spec()["id"]) / "receipt.json"
    formal_paths = [CHILD_ROOT / str(row["id"]) / "receipt.json" for row in schedule_rows()]
    require(preflight_receipt_path.is_file(), "real preflight receipt missing")
    preflight_receipt = load_packet_receipt(preflight_receipt_path, ARTIFACT_ROOT)
    binary = Path(str(manifest["build"]["binary"]["path"])).resolve()
    preflight = replay_child(preflight_receipt, ARTIFACT_ROOT, preflight_spec(), binary,
                             preflight=True)
    require(all(path.is_file() for path in formal_paths), "formal schedule is incomplete")
    loaded: list[dict[str, object]] = []
    for row, path in zip(schedule_rows(), formal_paths, strict=True):
        receipt = load_packet_receipt(path, ARTIFACT_ROOT)
        loaded.append(replay_child(receipt, ARTIFACT_ROOT, row, binary, preflight=False))
    output = {
        "schema": "litchi.performance.0942.cfb-profile-reduction.v5",
        "claim": "native CFB attribution only; no latency comparison and no adoption gate",
        "timing_claim": False,
        "harness_overhead_caveat": "profiled marker/barrier and diagnostic build are attribution instrumentation; report timing arrays are retained only for correctness validation",
        "plan": descriptor(PLAN_PATH),
        "manifest": descriptor(MANIFEST_PATH),
        "preflight": preflight,
        "formal_children": loaded,
        "formal_child_count": len(loaded),
        "all_correctness_validated": True,
        "all_marker_native_frame_coverage_validated": True,
        "created_utc": utc_now(),
    }
    path = REPORT_ROOT / "formal-reduction-v5.json"
    write_exclusive(path, output)
    print(json.dumps({"status": "reduced", "children": len(loaded),
                      "report": str(path), "sha256": sha256_path(path)}, sort_keys=True))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    sub = parser.add_subparsers(dest="action", required=True)
    sub.add_parser("write-plan", help="write the inert external prospective plan")
    sub.add_parser("write-manifest-draft", help="write an unapproved external binding draft")
    sub.add_parser("check-plan", help="validate the frozen prospective plan")
    sub.add_parser("real-preflight", help="run one authorized real profile child")
    sub.add_parser("capture", help="run the authorized serial 16-child profile schedule")
    sub.add_parser("reduce", help="rehash, validate, and reduce retained children")
    args = parser.parse_args()
    try:
        if args.action == "write-plan":
            write_plan()
            return 0
        plan = load_plan()
        if args.action == "write-manifest-draft":
            write_manifest_draft()
            return 0
        if args.action == "check-plan":
            print(json.dumps({"status": plan["status"], "formal_children": 16,
                              "preflight": preflight_spec()["id"],
                              "runner_sha256": sha256_path(HERE)}, sort_keys=True))
            return 0
        require_approval()
        manifest = load_manifest(plan)
        if args.action == "real-preflight":
            run_capture(plan, manifest, preflight=True)
        elif args.action == "capture":
            run_capture(plan, manifest, preflight=False)
        elif args.action == "reduce":
            reduce_packet(plan, manifest)
        else:
            raise RunnerError(f"unknown action: {args.action}")
        return 0
    except RunnerError as error:
        print(f"profile-runner: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
