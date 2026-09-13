#!/usr/bin/env python3
"""Validate and compare the matched 0555 OLE2/OOXML captures.

The 0555 driver emits two non-instrumented native children and two allocator
children for each stage.  This analyzer owns only the numerical evidence in
those children.  It does not build, run, profile, or modify the capture.  The
allocator executable's elapsed samples are validated as evidence but are
never used as native latency.  Whole-child GNU-time RSS is retained once per
native job and is compared independently of the per-case elapsed samples.

The output keeps each native sample vector and all percentiles in the compact
summary rows, retains every allocation counter vector, and points at the raw
report/catalog/receipt/RSS files with hashes.  Large raw operation-metric
objects therefore remain in their original files rather than being copied
into the comparison repeatedly.  Missing, malformed, stale, or mismatched
evidence fails closed.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import math
import re
import sys
from pathlib import Path
from typing import Any


# Importing this module must never create __pycache__ in the sealed evidence
# directory.  The analyzer is read-only until its final exclusive report
# write.
sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))
if str(REPO) not in sys.path:
    sys.path.insert(0, str(REPO))

# The driver is imported only for its immutable path constants.  Load it by
# absolute bundle path so an older campaign's ``run`` module cannot leak into
# this consumer when it is embedded by another Python process.  Matrix
# reconstruction below is local so the driver's mutable stage globals do not
# affect analysis.
_RUN_SPEC = importlib.util.spec_from_file_location(
    "ole2_physical_marker_0555_run", HERE / "run.py"
)
if _RUN_SPEC is None or _RUN_SPEC.loader is None:
    raise ImportError(f"cannot load 0555 driver: {HERE / 'run.py'}")
RUN = importlib.util.module_from_spec(_RUN_SPEC)
_RUN_SPEC.loader.exec_module(RUN)

from tools.summarize_crud_baseline import (  # noqa: E402
    _rust_statistics,
    _validate_elapsed,
)
from tools import validate_perf_corpus_binding as _BINDING  # noqa: E402


SCHEMA = "ole2_physical_marker_metrics_0555_v1"
STAGES = ("baseline", "candidate")
LANES = ("native", "alloc")
NORMAL = "normal"
ALLOC = "alloc"
BINARY_NAMES = {
    NORMAL: "litchi-perf-baseline",
    ALLOC: "litchi-perf-baseline-alloc",
}
XLS_CASES = (
    "xls_semantic_open",
    "xls_eager_open_list_worksheets",
    "xls_eager_open_one_cell",
    "xls_source_backed_open",
    "xls_source_backed_open_list_worksheets",
    "xls_source_backed_open_one_cell",
    "xls_owned_source_open",
    "xls_owned_source_open_list_worksheets",
    "xls_owned_source_open_one_cell",
)
CFB_CASES = ("cfb_open",)
CFB_SHAPES = ("tiny", "many-small", "few-large")
ALLOCATION_FIELDS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
    "incremental_region_peak_live_bytes",
)
ALLOCATION_GATE_FIELDS = (
    "allocation_calls",
    "allocated_bytes",
    "incremental_region_peak_live_bytes",
)
SUMMARY_FIELDS = (
    "min",
    "p50",
    "p95",
    "p99",
    "max",
    "mean",
    "standard_deviation",
)
CI_FIELDS = ("lower", "upper")
RSS_FIELDS = ("max_rss_kib", "elapsed_seconds", "user_seconds", "system_seconds")
DRIFT_THRESHOLD_PERCENT = 5.0
PRIMARY_IMPROVEMENT_PERCENT = 3.0
HOST_SCOPE = "Accessible compiler processes; no host quiescence guarantee"
RECEIPT_KEYS = {
    "schema",
    "command",
    "start_utc",
    "end_utc",
    "seconds",
    "exit_code",
    "stage",
    "execution_stage",
    "execution_manifest_sha256",
    "binary_sha256",
    "source_manifest_sha256",
    "workspace_lock_sha256",
    "workspace_lock_binding_sha256",
    "script_sha256",
    "plan_sha256",
    "environment",
    "artifacts",
}
RECEIPT_ENVIRONMENT_KEYS = {
    "TMPDIR",
    "CARGO_TARGET_DIR",
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "RUSTFLAGS",
    "CARGO_ENCODED_RUSTFLAGS",
    "LD_PRELOAD",
    "MALLOC_CONF",
    "GLIBC_TUNABLES",
}
VECTOR_STATUSES = {"measured", "not_applicable", "unavailable", "overflow"}
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory capture artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def _reject_constant(value: str) -> None:
    raise EvidenceError(f"non-finite JSON number {value!r}")


def _no_duplicate_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise EvidenceError(f"duplicate JSON object key {key!r}")
        result[key] = value
    return result


def _finite_tree(value: Any, label: str) -> None:
    if value is None or isinstance(value, (bool, str)):
        return
    if isinstance(value, (int, float)):
        require(math.isfinite(float(value)), f"{label} contains a non-finite number")
        return
    if isinstance(value, list):
        for index, item in enumerate(value):
            _finite_tree(item, f"{label}[{index}]")
        return
    if isinstance(value, dict):
        for key, item in value.items():
            require(isinstance(key, str), f"{label} has a non-string key")
            _finite_tree(item, f"{label}.{key}")
        return
    raise EvidenceError(f"{label} has unsupported JSON type {type(value).__name__}")


def read_json(path: Path) -> Any:
    try:
        value = json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=_no_duplicate_pairs,
            parse_constant=_reject_constant,
        )
    except EvidenceError:
        raise
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error
    _finite_tree(value, str(path))
    return value


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def is_hash(value: Any) -> bool:
    return isinstance(value, str) and SHA256_RE.fullmatch(value) is not None


def check_hash(value: Any, label: str) -> str:
    require(is_hash(value), f"{label} is not a lowercase SHA-256")
    return value


def finite_number(value: Any, label: str, *, minimum: float | None = None) -> None:
    require(
        isinstance(value, (int, float)) and not isinstance(value, bool),
        f"{label} is not numeric",
    )
    number = float(value)
    require(math.isfinite(number), f"{label} is not finite")
    if minimum is not None:
        require(number >= minimum, f"{label} is below {minimum}")


def nonnegative_integer(value: Any, label: str) -> None:
    require(
        isinstance(value, int) and not isinstance(value, bool) and value >= 0,
        f"{label} is not a non-negative integer",
    )


def timestamp(value: Any, label: str) -> _datetime.datetime:
    require(isinstance(value, str), f"{label} timestamp is missing")
    try:
        parsed = _datetime.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} timestamp is malformed") from error
    require(parsed.tzinfo is not None, f"{label} timestamp has no timezone")
    return parsed


def canonical(value: Any) -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        raise EvidenceError(f"value is not canonical JSON: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"evidence path escapes campaign bundle: {path}") from error


def plan_data() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(REVISION_RE.fullmatch(plan.get("revision", "")) is not None,
            "plan revision is not a lowercase commit hash")
    require(plan.get("schema") == "ole2_physical_marker_0555_plan_v1",
            "plan schema differs")
    require(plan.get("scope") == (
        "Matched OLE2 physical-sector marker accounting experiment across the "
        "FAT, DIFAT, Directory, MiniFAT, MiniStream, and RegularStream roles; "
        "no public API or semantic format change"
    ),
            "plan scope differs")
    require(plan.get("priority") == (
        "OLE2/OOXML first; ODF deferred until the OLE2/OOXML optimization goal "
        "completes; iWork excluded"
    ),
            "plan priority differs")
    require(plan.get("status") == (
        "frozen before build, capture, candidate application, and result observation"
    ),
            "plan is not frozen before build and capture")
    require(plan.get("cpu") == 2, "plan CPU differs from pinned CPU")
    require(plan.get("owned_paths") == ["/home/zhuhe/litchi-goal-0555-target"],
            "plan owned target differs")
    require(plan.get("candidate_files") == ["crates/litchi-cfb/src/file.rs"],
            "plan candidate file scope differs")
    groups = plan.get("groups")
    require(isinstance(groups, dict), "plan groups are missing")
    require(groups.get("xls", {}).get("cases") == list(XLS_CASES),
            "plan XLS case matrix differs")
    require(groups.get("xls", {}).get("primary_cases") == [
        "xls_source_backed_open",
        "xls_source_backed_open_one_cell",
        "xls_owned_source_open",
        "xls_owned_source_open_one_cell",
    ], "plan primary XLS case matrix differs")
    require(groups.get("cfb", {}).get("cases") == list(CFB_CASES),
            "plan CFB case matrix differs")
    require(groups.get("cfb", {}).get("shapes") == list(CFB_SHAPES),
            "plan CFB shape matrix differs")
    require(groups.get("cfb", {}).get("payload") == "incompressible",
            "plan CFB payload differs")
    native = plan.get("native")
    require(native == {
        "repeats": 2,
        "warmup": 20,
        "samples": 1000,
        "order": [
            "baseline r1 xls/cfb",
            "candidate r1 xls/cfb",
            "candidate r2 cfb/xls",
            "baseline r2 cfb/xls under candidate execution-stage binding",
        ],
        "source_identity": (
            "Every receipt binds both output stage and live execution stage; retained "
            "baseline-r2 uses the baseline binary/output folder with candidate source "
            "execution identity"
        ),
    }, "plan native lane differs")
    allocation = plan.get("allocation")
    require(isinstance(allocation, dict), "plan allocation lane is missing")
    require(allocation.get("repeats") == 2 and allocation.get("warmup") == 3
            and allocation.get("samples") == 30,
            "plan allocation counts differ")
    require(allocation.get("order") == [
        "baseline r1 xls/cfb",
        "candidate r1 xls/cfb",
        "candidate r2 cfb/xls",
        "baseline r2 cfb/xls under candidate execution-stage binding",
    ], "plan allocation order differs")
    require(allocation.get("scope") == (
        "Existing constructor/operation clock with the canonical operation-global System allocator region; "
        "excludes fixture construction, correctness oracles, report construction, and object drop"
    ), "plan allocation scope differs")
    require(allocation.get("vector") == (
        "Every nine XLS case and three CFB shape in both repeats; calls, allocated bytes, "
        "and incremental region peak are retained per operation"
    ), "plan allocation vector scope differs")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "plan profile lane is missing")
    require(profile.get("repeats") == 2 and profile.get("warmup") == 0
            and profile.get("samples") == 5,
            "plan profile counts differ")
    require(profile.get("jobs") == ["xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large"],
            "plan profile jobs differ")
    require(profile.get("xls_owner") == (
        "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
    ), "plan XLS profile owner differs")
    require(profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open",
            "plan CFB profile owner differs")
    require(profile.get("positive_timed_dumps") == (
        "Exactly five positive timed constructor dumps per job and repeat; setup and "
        "termination dumps are retained and classified separately"
    ), "plan profile dump contract differs")
    require(profile.get("runtime_flags") == [
        "--vgdb=no",
        "--collect-atstart=no",
        "--toggle-collect=<owner>",
        "--zero-before=<owner>",
        "--dump-after=<owner>",
    ], "plan profile runtime flags differ")
    admission = plan.get("admission")
    require(isinstance(admission, dict), "plan admission is missing")
    for key in (
        "primary_xls_p50", "primary_xls_mean", "native_xls_controls",
        "native_cfb_controls", "native_rss", "allocation", "correctness",
        "mechanism", "review", "disposition",
    ):
        require(isinstance(admission.get(key), str) and admission[key].strip(),
                f"plan admission.{key} is missing")
    require(isinstance(plan.get("prior_scope_note"), str)
            and "0554" in plan["prior_scope_note"]
            and "many-small" in plan["prior_scope_note"]
            and "before" in plan["prior_scope_note"],
            "plan prior scope note is missing")
    return plan


def frozen_inputs(plan: dict[str, Any]) -> dict[str, str]:
    value = read_json(HERE / "frozen-inputs.json")
    require(isinstance(value, dict), "frozen-inputs.json is not an object")
    require(set(value) == {"schema", "frozen_utc", "files"},
            "frozen input envelope differs")
    require(value.get("schema") == "ole2_0555_frozen_inputs_v1",
            "frozen input schema differs")
    timestamp(value.get("frozen_utc"), "frozen input frozen_utc")
    files = value.get("files")
    require(isinstance(files, dict), "frozen input file map is missing")
    expected = {"plan.json", "run.py", "workspace-lock.json", "adr-manifest.json"}
    require(set(files) == expected, "frozen input inventory differs")
    for name, digest in files.items():
        check_hash(digest, f"frozen input {name}")
        require(digest == sha256(HERE / name), f"frozen input {name} changed")
    require(files["plan.json"] == sha256(HERE / "plan.json"),
            "frozen plan hash differs")
    require(files["run.py"] == sha256(HERE / "run.py"),
            "frozen driver hash differs")
    lock = read_json(HERE / "workspace-lock.json")
    require(isinstance(lock, dict), "workspace-lock.json is not an object")
    require(lock.get("path") == "Cargo.lock", "workspace lock path differs")
    check_hash(lock.get("sha256"), "workspace-lock.json.sha256")
    require(lock["sha256"] == sha256(REPO / "Cargo.lock"),
            "Cargo.lock differs from the frozen workspace lock")
    require(lock["sha256"] == sha256(HERE / "workspace-Cargo.lock"),
            "retained workspace Cargo.lock differs from the frozen workspace lock")
    require(files["workspace-lock.json"] == sha256(HERE / "workspace-lock.json"),
            "workspace-lock input hash differs")
    require(files["adr-manifest.json"] == sha256(HERE / "adr-manifest.json"),
            "ADR manifest input hash differs")
    require(files["plan.json"] == sha256(HERE / "plan.json"),
            "frozen input plan hash differs from the active plan")
    return {key: files[key] for key in sorted(files)}


def source_manifest(stage: str) -> dict[str, str]:
    path = HERE / stage / "source-manifest.json"
    value = read_json(path)
    require(isinstance(value, dict) and value, f"{stage} source manifest is empty")
    require(list(value) == sorted(value), f"{stage} source manifest paths are not sorted")
    for name, digest in value.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"{stage} source manifest path is invalid")
        require(
            name in {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
            or name.startswith((".cargo/", "crates/", "tools/perf-baseline/")),
            f"{stage} source manifest path is out of scope: {name}",
        )
        check_hash(digest, f"{stage} source manifest {name}")
    return {name: value[name] for name in sorted(value)}


def host_sidecar(path: Path) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict), f"{path.name} host sidecar is not an object")
    require(value.get("scope") == HOST_SCOPE, f"{path.name} host scope differs")
    processes = value.get("compiler_processes")
    require(isinstance(processes, list), f"{path.name} compiler process list is missing")
    for index, process in enumerate(processes):
        require(isinstance(process, dict), f"{path.name} process {index} is not an object")
        nonnegative_integer(process.get("pid"), f"{path.name} process {index}.pid")
        require(process["pid"] > 0, f"{path.name} process {index}.pid is zero")
        require(process.get("comm") in {"cargo", "rustc"},
                f"{path.name} process {index}.comm is unexpected")
        require(isinstance(process.get("cwd"), str) and process["cwd"],
                f"{path.name} process {index}.cwd is missing")
    return value


def _expected_build_command(kind: str, plan: dict[str, Any]) -> list[str]:
    target = Path(plan["owned_paths"][0])
    command = [
        "env",
        "TMPDIR=" + str(target / "tmp"),
        "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0",
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        "tools/perf-baseline/Cargo.toml",
        "--bin",
        BINARY_NAMES[kind],
        "--target-dir",
        str(target),
    ]
    if kind == ALLOC:
        command += ["--features", "allocator-metrics"]
    return command


def _expected_execution_stage(stage: str, repeat: int) -> str:
    # The final ABBA baseline control uses the candidate manifest while its
    # retained baseline executable is measured.  This mirrors the explicit
    # execution-stage option in run.py and keeps source custody visible.
    return "candidate" if stage == "baseline" and repeat == 2 else stage


def validate_receipt_common(value: dict[str, Any], label: str) -> None:
    require(set(value) == RECEIPT_KEYS, f"{label} receipt schema differs")
    start = timestamp(value.get("start_utc"), f"{label}.start_utc")
    end = timestamp(value.get("end_utc"), f"{label}.end_utc")
    require(start < end, f"{label} timestamps are not increasing")
    finite_number(value.get("seconds"), f"{label}.seconds", minimum=0.0)
    require(value["seconds"] > 0.0, f"{label}.seconds is not positive")
    wall = (end - start).total_seconds()
    require(abs(float(value["seconds"]) - wall) <= max(0.25, wall * 0.02 + 0.05),
            f"{label}.seconds does not match its UTC interval")
    require(isinstance(value.get("command"), list)
            and all(isinstance(item, str) and item for item in value["command"]),
            f"{label}.command is malformed")
    require(isinstance(value.get("exit_code"), int)
            and not isinstance(value["exit_code"], bool),
            f"{label}.exit_code is malformed")
    require(value.get("schema") == "ole2_0555_run_receipt_v1",
            f"{label}.schema differs")
    require(value.get("stage") in set(STAGES) | {"final"},
            f"{label}.stage is invalid")
    require(value.get("execution_stage") in set(STAGES) | {"final"},
            f"{label}.execution_stage is invalid")
    binary_sha256 = value.get("binary_sha256")
    # Build receipts are created before a child binary is copied and therefore
    # deliberately carry null. Capture receipts are bound to their copied
    # executable by stage_analysis below.
    require(binary_sha256 is None or is_hash(binary_sha256),
            f"{label}.binary_sha256 is neither null nor a lowercase SHA-256")
    check_hash(value.get("execution_manifest_sha256"), f"{label}.execution_manifest_sha256")
    check_hash(value.get("source_manifest_sha256"), f"{label}.source_manifest_sha256")
    check_hash(value.get("workspace_lock_sha256"), f"{label}.workspace_lock_sha256")
    check_hash(value.get("workspace_lock_binding_sha256"),
               f"{label}.workspace_lock_binding_sha256")
    check_hash(value.get("script_sha256"), f"{label}.script_sha256")
    check_hash(value.get("plan_sha256"), f"{label}.plan_sha256")
    environment = value.get("environment")
    require(isinstance(environment, dict)
            and set(environment) == RECEIPT_ENVIRONMENT_KEYS,
            f"{label}.environment fields differ")
    require(isinstance(environment.get("TMPDIR"), str)
            and environment["TMPDIR"].endswith("/tmp"),
            f"{label}.environment.TMPDIR is missing")
    require(environment.get("CARGO_BUILD_JOBS") == "2"
            and environment.get("CARGO_INCREMENTAL") == "0",
            f"{label}.environment build controls differ")
    for key in RECEIPT_ENVIRONMENT_KEYS - {"TMPDIR", "CARGO_TARGET_DIR",
                                            "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL"}:
        require(environment.get(key) is None,
                f"{label}.environment.{key} is uncontrolled")
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}.artifacts is missing")


def validate_artifacts(folder: Path, artifacts: dict[str, Any], label: str) -> None:
    for name, digest in artifacts.items():
        require(isinstance(name, str) and Path(name).name == name,
                f"{label} has a non-local artifact name")
        check_hash(digest, f"{label}.artifacts.{name}")
        path = folder / name
        require(path.is_file() and not path.is_symlink(),
                f"{label} artifact is missing: {name}")
        require(sha256(path) == digest, f"{label} artifact hash differs: {name}")
        if name.endswith(".host.json"):
            host_sidecar(path)


def binary_metadata(stage: str, kind: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    path = folder / f"binary-{kind}.json"
    value = read_json(path)
    require(isinstance(value, dict), f"{path.name} is not an object")
    require(set(value) == {
        "path", "sha256", "bytes", "build_receipt_sha256",
        "source_manifest_sha256", "workspace_lock_sha256",
    }, f"{path.name} field inventory differs")
    digest = check_hash(value.get("sha256"), f"{path.name}.sha256")
    nonnegative_integer(value.get("bytes"), f"{path.name}.bytes")
    require(value["bytes"] > 0, f"{path.name}.bytes is zero")
    expected_path = Path(RUN.SCRATCH_ROOT) / stage / kind
    require(value.get("path") == str(expected_path), f"{path.name} path differs")
    executable = Path(value["path"])
    require(not executable.is_symlink(), f"{path.name} points to a symlink")
    if executable.exists():
        require(executable.is_file(), f"{path.name} executable is not a regular file")
        require(sha256(executable) == digest, f"{path.name} executable hash differs")
        require(executable.stat().st_size == value["bytes"],
                f"{path.name} executable size differs")
    # Before cleanup the binary must be present.  After cleanup a strict
    # campaign verifier owns binary custody; this analyzer still validates
    # every descriptor and does not invent a hash for a missing executable.
    build_path = folder / f"build-{kind}.receipt.json"
    require(value.get("build_receipt_sha256") == sha256(build_path),
            f"{path.name} build receipt hash differs")
    require(value.get("source_manifest_sha256") == sha256(folder / "source-manifest.json"),
            f"{path.name} source manifest hash differs")
    lock = read_json(HERE / "workspace-lock.json")
    require(value.get("workspace_lock_sha256") == lock.get("sha256")
            and value["workspace_lock_sha256"] == sha256(HERE / "workspace-Cargo.lock"),
            f"{path.name} workspace lock hash differs")
    build = read_json(build_path)
    require(isinstance(build, dict), f"{build_path.name} is not an object")
    validate_receipt_common(build, build_path.name)
    require(build.get("exit_code") == 0 and build.get("binary_sha256") is None,
            f"{build_path.name} is not a successful build receipt")
    require(build.get("stage") == stage,
            f"{build_path.name} stage differs")
    require(build.get("execution_stage") == stage,
            f"{build_path.name} execution stage differs")
    manifest = folder / "source-manifest.json"
    require(build.get("execution_manifest_sha256") == sha256(manifest),
            f"{build_path.name} execution manifest differs")
    require(build.get("source_manifest_sha256") == sha256(manifest),
            f"{build_path.name} source manifest differs")
    require(build.get("script_sha256") == sha256(HERE / "run.py"),
            f"{build_path.name} driver hash differs")
    require(build.get("plan_sha256") == sha256(HERE / "plan.json"),
            f"{build_path.name} plan hash differs")
    require(build.get("workspace_lock_sha256") == sha256(HERE / "workspace-Cargo.lock"),
            f"{build_path.name} workspace lock hash differs")
    require(build.get("workspace_lock_binding_sha256") == sha256(HERE / "workspace-lock.json"),
            f"{build_path.name} workspace lock binding differs")
    require(build.get("command") == _expected_build_command(kind, plan),
            f"{build_path.name} command differs")
    expected_artifacts = {
        f"build-{kind}.host.json",
        f"build-{kind}.stdout",
        f"build-{kind}.stderr",
    }
    require(set(build["artifacts"]) == expected_artifacts,
            f"{build_path.name} artifact inventory differs")
    validate_artifacts(folder, build["artifacts"], build_path.name)
    return {
        "path": str(executable),
        "sha256": digest,
        "bytes": value["bytes"],
        "build_receipt_sha256": value["build_receipt_sha256"],
        "source_manifest_sha256": value["source_manifest_sha256"],
        "workspace_lock_sha256": value["workspace_lock_sha256"],
        "binary": BINARY_NAMES[kind],
        "stage": stage,
        "present": executable.exists(),
    }


def jobs_for(plan: dict[str, Any], lane: str) -> list[dict[str, Any]]:
    require(lane in LANES, f"unknown lane {lane}")
    output: list[dict[str, Any]] = []
    config = plan["allocation"] if lane == "alloc" else plan["native"]
    for repeat in (1, 2):
        groups = ["xls", "cfb"] if lane == "alloc" or repeat == 1 else ["cfb", "xls"]
        for group in groups:
            selection = plan["groups"][group]
            output.append({
                "name": f"{lane}-r{repeat}-{group}",
                "lane": lane,
                "repeat": repeat,
                "group": group,
                "selection": selection,
                "samples": config["samples"],
                "warmup": config["warmup"],
            })
    return output


def _command_option(command: list[str], option: str) -> str:
    positions = [index for index, item in enumerate(command) if item == option]
    require(len(positions) == 1 and positions[0] + 1 < len(command),
            f"command option {option} is missing or repeated")
    return command[positions[0] + 1]


def validate_capture_command(
    command: list[str], job: dict[str, Any], kind: str, folder: Path,
    binary: dict[str, Any], plan: dict[str, Any],
) -> None:
    selection = job["selection"]
    expected = ["taskset", "-c", str(plan["cpu"])]
    if kind == NORMAL:
        expected += [
            "/usr/bin/time", "-f",
            '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
            "-o", str(folder / f"{job['name']}.rss.json"),
        ]
    expected += [
        binary["path"],
        "--case", ",".join(selection["cases"]),
        "--warmup", str(job["warmup"]),
        "--samples", str(job["samples"]),
        "--json", str(folder / f"{job['name']}.json"),
        "--corpus-manifest", str(folder / f"{job['name']}.catalog.json"),
    ]
    if "shapes" in selection:
        expected += ["--shape", ",".join(selection["shapes"]),
                     "--payload", selection["payload"]]
    else:
        require("--shape" not in command and "--payload" not in command,
                f"{job['name']} XLS command has CFB selectors")
    require(command == expected, f"{job['name']} capture command differs")


def validate_rss(path: Path) -> dict[str, Any]:
    value = read_json(path)
    require(isinstance(value, dict), f"{path.name} RSS is not an object")
    require(set(value) == set(RSS_FIELDS), f"{path.name} RSS fields differ")
    nonnegative_integer(value.get("max_rss_kib"), f"{path.name}.max_rss_kib")
    for field in RSS_FIELDS[1:]:
        finite_number(value.get(field), f"{path.name}.{field}", minimum=0.0)
    return value


def _validate_metric_vector(value: dict[str, Any], count: int, label: str) -> None:
    require(set(value) <= {"status", "scope", "values"},
            f"{label} metric vector has unexpected fields")
    status = value.get("status")
    require(status in VECTOR_STATUSES, f"{label}.status is invalid")
    require(isinstance(value.get("scope"), str) and value["scope"],
            f"{label}.scope is missing")
    values = value.get("values")
    if status == "measured":
        require(isinstance(values, list) and len(values) == count,
                f"{label}.values must contain {count} values")
        for index, item in enumerate(values):
            if isinstance(item, bool):
                raise EvidenceError(f"{label}.values[{index}] is boolean")
            if isinstance(item, int):
                require(item >= 0, f"{label}.values[{index}] is negative")
            elif isinstance(item, str):
                require(item in {"sequential", "random", "unknown"},
                        f"{label}.values[{index}] is not a known category")
            else:
                raise EvidenceError(f"{label}.values[{index}] is not scalar")
    else:
        require(values is None or "values" not in value,
                f"{label} has values despite status {status}")


def validate_metric_tree(value: Any, count: int, label: str) -> None:
    if isinstance(value, dict):
        if "status" in value and "scope" in value and (
            "values" in value or set(value) <= {"status", "scope"}
        ):
            _validate_metric_vector(value, count, label)
            return
        for key, child in value.items():
            validate_metric_tree(child, count, f"{label}.{key}")
    elif isinstance(value, list):
        # Lists outside vectors are metadata containers in the producer JSON;
        # recursively validate any nested metric envelopes without assuming a
        # count for arbitrary semantic lists.
        for index, child in enumerate(value):
            validate_metric_tree(child, count, f"{label}[{index}]")
    elif value is not None:
        require(isinstance(value, (str, int, float, bool)),
                f"{label} has an unsupported scalar")


def operation_metrics(row: dict[str, Any], count: int, allocator: bool, label: str) -> dict[str, Any]:
    value = row.get("operation_metrics")
    require(isinstance(value, dict), f"{label}.operation_metrics is missing")
    require(value.get("sample_count") == count,
            f"{label}.operation_metrics.sample_count differs")
    indices = value.get("sample_indices")
    require(isinstance(indices, list) and sorted(indices) == list(range(count)),
            f"{label}.operation_metrics.sample_indices is not a permutation")
    elapsed = row["elapsed_ns"]
    require(value.get("sample_indices") == elapsed["sample_order"],
            f"{label}.operation_metrics is not aligned to elapsed_ns")
    require(value.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label}.operation_metrics alignment differs")
    validate_metric_tree(value, count, f"{label}.operation_metrics")
    allocation = value.get("allocation")
    require(isinstance(allocation, dict), f"{label}.operation_metrics.allocation is missing")
    require(allocation.get("scope") == "operation_global_system_allocator",
            f"{label}.operation_metrics.allocation scope differs")
    require(allocation.get("status") == ("measured" if allocator else "unavailable"),
            f"{label}.operation_metrics.allocation status differs")
    for field in ALLOCATION_FIELDS[:-1]:
        vector = allocation.get(field)
        require(isinstance(vector, dict), f"{label}.allocation.{field} is missing")
        if allocator:
            require(vector.get("status") == "measured",
                    f"{label}.allocation.{field} is not measured")
        else:
            require(vector.get("status") == "unavailable",
                    f"{label}.allocation.{field} is not unavailable")
    return value


def validate_report(
    path: Path, catalog_path: Path, job: dict[str, Any], kind: str,
    binary: dict[str, Any], plan: dict[str, Any],
) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    report = read_json(path)
    catalog = read_json(catalog_path)
    require(isinstance(report, dict), f"{path.name} report is not an object")
    require(isinstance(catalog, dict), f"{catalog_path.name} catalog is not an object")
    try:
        _BINDING.validate_binding(report, catalog)
    except Exception as error:  # the boundary validator has its own exception type
        raise EvidenceError(f"{path.name} corpus/catalog binding failed: {error}") from error
    require(report.get("schema_version") == 1, f"{path.name} schema version differs")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{path.name}.tool is missing")
    require(tool.get("binary") == BINARY_NAMES[kind], f"{path.name} tool binary differs")
    require(tool.get("profile") == "release", f"{path.name} is not release profile")
    expected_instrumentation = (
        "system_allocator_operation_scoped" if kind == ALLOC else "none"
    )
    require(tool.get("instrumentation") == expected_instrumentation,
            f"{path.name} instrumentation differs")
    if kind == ALLOC:
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{path.name} allocator counter revision differs")
    else:
        require(tool.get("allocator_counter_revision") is None,
                f"{path.name} native report carries allocator revision")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict), f"{path.name}.binary_identity is missing")
    require(identity.get("binary_sha256") == binary["sha256"],
            f"{path.name} binary hash differs")
    require(identity.get("binary_bytes") == binary["bytes"],
            f"{path.name} binary size differs")
    require(identity.get("path") == binary["path"], f"{path.name} binary path differs")
    require(identity.get("profile") == "release" and identity.get("executable") is True,
            f"{path.name} binary identity is incomplete")
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{path.name}.environment is missing")
    require(environment.get("git_revision") == plan["revision"],
            f"{path.name} git revision differs from plan")
    require(environment.get("cpu_affinity") == str(plan["cpu"]),
            f"{path.name} CPU affinity differs from plan")
    expected_allocator = (
        "CountingSystemAllocator(std::alloc::System)" if kind == ALLOC
        else "Rust system allocator"
    )
    require(environment.get("allocator") == expected_allocator,
            f"{path.name} allocator environment differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{path.name}.configuration is missing")
    selection = job["selection"]
    require(configuration.get("cases") == selection["cases"],
            f"{path.name} case selection differs")
    require(configuration.get("samples_per_case") == job["samples"],
            f"{path.name} sample count differs")
    require(configuration.get("warmup_iterations_per_case") == job["warmup"],
            f"{path.name} warmup count differs")
    require(configuration.get("filesystem_process_isolated") is True
            and configuration.get("filesystem_fresh_child_per_sample") is True,
            f"{path.name} process isolation contract differs")
    if "shapes" in selection:
        require(configuration.get("corpus_shapes") == selection["shapes"],
                f"{path.name} corpus shapes differ")
        require(configuration.get("payload_kinds") == [selection["payload"]],
                f"{path.name} payload differs")
    reference = report.get("corpus_catalog")
    require(isinstance(reference, dict), f"{path.name}.corpus_catalog is missing")
    for key in ("manifest_version", "catalog_id", "content_set_sha256"):
        require(reference.get(key) == catalog.get(key),
                f"{path.name} corpus catalog {key} differs")
    require(reference.get("catalog_sha256") == catalog.get("catalog_sha256"),
            f"{path.name} corpus catalog catalog_sha256 differs")
    rows = report.get("results")
    require(isinstance(rows, list), f"{path.name}.results is not a list")
    if job["group"] == "xls":
        require(len(rows) == len(XLS_CASES), f"{path.name} XLS row count differs")
        expected_keys = {(case, None) for case in XLS_CASES}
        actual_keys = {(row.get("case"), None) for row in rows if isinstance(row, dict)}
    else:
        require(len(rows) == len(CFB_CASES) * len(CFB_SHAPES),
                f"{path.name} CFB row count differs")
        expected_keys = {(case, shape) for case in CFB_CASES for shape in selection["shapes"]}
        actual_keys = {
            (row.get("case"), row.get("corpus", {}).get("shape"))
            for row in rows if isinstance(row, dict)
        }
    require(len(actual_keys) == len(rows) and actual_keys == expected_keys,
            f"{path.name} result matrix differs")
    for index, row in enumerate(rows):
        label = f"{path.name}.results[{index}]"
        require(isinstance(row, dict), f"{label} is not an object")
        elapsed = row.get("elapsed_ns")
        require(isinstance(elapsed, dict), f"{label}.elapsed_ns is missing")
        try:
            checked = _validate_elapsed(row, job["samples"], label)
        except Exception as error:
            raise EvidenceError(f"{label}.elapsed_ns validation failed: {error}") from error
        operation_metrics(row, job["samples"], kind == ALLOC, label)
        case = row.get("case")
        if job["group"] == "xls":
            require(case in XLS_CASES, f"{label} has an unexpected XLS case")
            require(row.get("sink") is None, f"{label} has sink evidence")
            check_hash(row.get("output_sha256"), f"{label}.output_sha256")
        else:
            require(case == "cfb_open", f"{label} is not cfb_open")
            require(row.get("source") is None and row.get("sink") is None
                    and row.get("output_sha256") is None,
                    f"{label} CFB row has fabricated publication evidence")
        # _validate_elapsed returns the producer-equivalent statistics.  The
        # producer's own object was checked by that helper; retain its raw
        # vector and source order from the original report in the row summary.
        require(checked["samples"] == elapsed["samples"]
                and checked["sample_order"] == elapsed["sample_order"],
                f"{label} elapsed evidence changed while validating")
    return report, rows, catalog


def normalize_identity(value: Any) -> Any:
    if isinstance(value, list):
        require(value, "identity vector is empty")
        values = [normalize_identity(item) for item in value]
        if all(canonical(item) == canonical(values[0]) for item in values[1:]):
            return values[0]
        return values
    if isinstance(value, dict):
        return {key: normalize_identity(value[key]) for key in sorted(value)}
    return value


def row_identity(row: dict[str, Any]) -> dict[str, Any]:
    # These are the complete semantic/source/corpus and publication identity
    # fields.  Timing and operation counters are intentionally excluded: they
    # are the measured values being compared, not workload identity.
    return {
        "case": row.get("case"),
        "corpus": normalize_identity(row.get("corpus")),
        "output_sha256": row.get("output_sha256"),
        "sink": normalize_identity(row.get("sink"))
        if row.get("sink") is not None else None,
        "source": normalize_identity(row.get("source"))
        if row.get("source") is not None else None,
    }


def identity_digest(identity: dict[str, Any]) -> str:
    return hashlib.sha256(canonical(identity).encode("utf-8")).hexdigest()


def summary(values: list[int], unit: str) -> dict[str, Any]:
    require(values, "summary values are empty")
    require(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0
                for item in values), "summary values are not non-negative integers")
    ordered = sorted(values)
    # _rust_statistics is a pure arithmetic helper.  Its input must be the
    # producer's sorted elapsed representation; using a fresh order here is
    # sufficient for allocation vectors, whose sample order is not semantic.
    value = _rust_statistics(ordered, list(range(len(ordered))))
    value["unit"] = unit
    value.pop("sample_order", None)
    return {key: value[key] for key in ("unit", *SUMMARY_FIELDS, "confidence_interval_95")}


def allocation_vectors(row: dict[str, Any], count: int, label: str) -> dict[str, list[int]]:
    allocation = row["operation_metrics"]["allocation"]
    output: dict[str, list[int]] = {}
    for field in ALLOCATION_FIELDS[:-1]:
        vector = allocation[field]
        require(vector.get("status") == "measured", f"{label}.{field} is not measured")
        values = vector.get("values")
        require(isinstance(values, list) and len(values) == count,
                f"{label}.{field} has the wrong sample count")
        checked: list[int] = []
        for index, item in enumerate(values):
            nonnegative_integer(item, f"{label}.{field}[{index}]")
            checked.append(item)
        output[field] = checked
    for index in range(count):
        require(output["failed_allocation_calls"][index] == 0,
                f"{label} recorded a failed allocation")
        require(
            output["live_bytes_after"][index]
            == output["live_bytes_before"][index]
            + output["allocated_bytes"][index]
            - output["deallocated_bytes"][index],
            f"{label} live-byte balance does not reconcile at sample {index}",
        )
        require(output["peak_live_bytes_before"][index] >= output["live_bytes_before"][index],
                f"{label} peak-before is below live-before at sample {index}")
        require(output["peak_live_bytes_after"][index] >= output["live_bytes_after"][index],
                f"{label} peak-after is below live-after at sample {index}")
        require(output["region_peak_live_bytes"][index] >= max(
            output["live_bytes_before"][index], output["live_bytes_after"][index]
        ), f"{label} region peak is below live bytes at sample {index}")
    output["incremental_region_peak_live_bytes"] = [
        peak - before
        for peak, before in zip(
            output["region_peak_live_bytes"], output["live_bytes_before"]
        )
    ]
    return output


def evidence_ref(path: Path) -> dict[str, str]:
    return {"path": relative(path), "sha256": sha256(path)}


def stage_analysis(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    require(folder.is_dir(), f"missing {stage} evidence directory")
    manifest = source_manifest(stage)
    binaries = {kind: binary_metadata(stage, kind, plan) for kind in (NORMAL, ALLOC)}
    rows: list[dict[str, Any]] = []
    identities: dict[str, dict[str, Any]] = {}
    rss_by_job: dict[str, dict[str, Any]] = {}
    receipts: list[dict[str, Any]] = []
    intervals: list[tuple[_datetime.datetime, _datetime.datetime, str]] = []
    for lane in LANES:
        kind = NORMAL if lane == "native" else ALLOC
        for job in jobs_for(plan, lane):
            name = job["name"]
            receipt_path = folder / f"{name}.receipt.json"
            receipt = read_json(receipt_path)
            require(isinstance(receipt, dict), f"{receipt_path.name} is not an object")
            validate_receipt_common(receipt, receipt_path.name)
            require(receipt.get("exit_code") == 0, f"{receipt_path.name} did not exit successfully")
            require(receipt.get("stage") == stage,
                    f"{receipt_path.name} output stage differs")
            require(receipt.get("execution_stage") == _expected_execution_stage(stage, job["repeat"]),
                    f"{receipt_path.name} execution stage differs")
            execution_manifest = HERE / receipt["execution_stage"] / "source-manifest.json"
            require(receipt.get("execution_manifest_sha256") == sha256(execution_manifest),
                    f"{receipt_path.name} execution manifest differs")
            require(receipt.get("source_manifest_sha256") == sha256(folder / "source-manifest.json"),
                    f"{receipt_path.name} source manifest differs")
            require(receipt.get("script_sha256") == sha256(HERE / "run.py"),
                    f"{receipt_path.name} driver hash differs")
            require(receipt.get("plan_sha256") == sha256(HERE / "plan.json"),
                    f"{receipt_path.name} plan hash differs")
            require(receipt.get("workspace_lock_sha256") == sha256(HERE / "workspace-Cargo.lock"),
                    f"{receipt_path.name} workspace lock hash differs")
            require(receipt.get("workspace_lock_binding_sha256") == sha256(HERE / "workspace-lock.json"),
                    f"{receipt_path.name} workspace lock binding differs")
            require(receipt.get("binary_sha256") == binaries[kind]["sha256"],
                    f"{receipt_path.name} binary hash differs")
            validate_artifacts(folder, receipt["artifacts"], receipt_path.name)
            expected_artifacts = {
                f"{name}.host.json",
                f"{name}.stdout",
                f"{name}.stderr",
                f"{name}.json",
                f"{name}.catalog.json",
            }
            if lane == "native":
                expected_artifacts.add(f"{name}.rss.json")
            require(set(receipt["artifacts"]) == expected_artifacts,
                    f"{receipt_path.name} artifact inventory differs")
            validate_capture_command(
                receipt["command"], job, kind, folder, binaries[kind], plan
            )
            report_path = folder / f"{name}.json"
            catalog_path = folder / f"{name}.catalog.json"
            report, raw_rows, catalog = validate_report(
                report_path, catalog_path, job, kind, binaries[kind], plan
            )
            rss = None
            if lane == "native":
                rss_path = folder / f"{name}.rss.json"
                rss_value = validate_rss(rss_path)
                rss = {
                    **rss_value,
                    "job": name,
                    "repeat": job["repeat"],
                    "group": job["group"],
                    "evidence": evidence_ref(rss_path),
                }
                rss_by_job[name] = rss
            start = timestamp(receipt["start_utc"], f"{name}.start_utc")
            end = timestamp(receipt["end_utc"], f"{name}.end_utc")
            intervals.append((start, end, name))
            receipts.append({
                "name": name,
                "lane": lane,
                "repeat": job["repeat"],
                "group": job["group"],
                "execution_stage": receipt["execution_stage"],
                "binary_sha256": receipt["binary_sha256"],
                "source_manifest_sha256": receipt["source_manifest_sha256"],
                "execution_manifest_sha256": receipt["execution_manifest_sha256"],
                "evidence": evidence_ref(receipt_path),
            })
            require(len(raw_rows) > 0, f"{name} has no result rows")
            for index, raw in enumerate(raw_rows):
                label = f"{name}.results[{index}]"
                identity = row_identity(raw)
                digest = identity_digest(identity)
                if digest in identities:
                    require(identities[digest] == identity,
                            f"{label} identity digest collision")
                else:
                    identities[digest] = identity
                corpus_shape = raw["corpus"]["shape"]
                key_shape = corpus_shape if job["group"] == "cfb" else "xls"
                item: dict[str, Any] = {
                    "stage": stage,
                    "lane": lane,
                    "repeat": job["repeat"],
                    "group": job["group"],
                    "job": name,
                    "case": raw["case"],
                    "shape": key_shape,
                    "corpus_shape": corpus_shape,
                    "identity_sha256": digest,
                    "identity_equal": True,
                    "samples": job["samples"],
                    "warmup": job["warmup"],
                    "binary_sha256": binaries[kind]["sha256"],
                    "execution_stage": receipt["execution_stage"],
                    "receipt_execution_manifest_sha256": receipt["execution_manifest_sha256"],
                    "receipt_source_manifest_sha256": receipt["source_manifest_sha256"],
                    "evidence": {
                        "report": evidence_ref(report_path),
                        "catalog": evidence_ref(catalog_path),
                        "receipt": evidence_ref(receipt_path),
                        "stdout": evidence_ref(folder / f"{name}.stdout"),
                        "stderr": evidence_ref(folder / f"{name}.stderr"),
                    },
                    "catalog_identity": {
                        "manifest_version": catalog["manifest_version"],
                        "catalog_id": catalog["catalog_id"],
                        "catalog_sha256": catalog["catalog_sha256"],
                        "content_set_sha256": catalog["content_set_sha256"],
                    },
                }
                if lane == "native":
                    item["elapsed_ns"] = {
                        key: raw["elapsed_ns"][key]
                        for key in (
                            "unit", "samples", "sample_order", *SUMMARY_FIELDS,
                            "confidence_interval_95",
                        )
                    }
                    item["rss_job"] = name
                else:
                    vectors = allocation_vectors(raw, job["samples"], label)
                    item["allocation_status"] = raw["operation_metrics"]["allocation"]["status"]
                    item["allocation_scope"] = raw["operation_metrics"]["allocation"]["scope"]
                    item["allocation_instrumented_elapsed_excluded"] = True
                    item["allocation"] = {
                        "samples": vectors,
                        "summary": {
                            field: summary(
                                vectors[field],
                                "count" if field.endswith("calls") else "bytes",
                            )
                            for field in ALLOCATION_FIELDS
                        },
                    }
                rows.append(item)
    ordered = sorted(intervals, key=lambda item: item[0])
    require(all(left[1] <= right[0] for left, right in zip(ordered, ordered[1:])),
            f"{stage} capture receipts overlap")
    expected_jobs = {lane: len(jobs_for(plan, lane)) for lane in LANES}
    expected_rows = {
        "native": len(XLS_CASES) + len(CFB_CASES) * len(CFB_SHAPES),
        "alloc": len(XLS_CASES) + len(CFB_CASES) * len(CFB_SHAPES),
    }
    for lane in LANES:
        current = [row for row in rows if row["lane"] == lane]
        require(len(current) == expected_rows[lane] * 2,
                f"{stage} {lane} row matrix is incomplete")
        expected_samples = plan["native" if lane == "native" else "allocation"]["samples"]
        require(sum(item["samples"] for item in current)
                == expected_samples * len(current),
                f"{stage} {lane} sample count differs")
    by_key = {
        (item["lane"], item["repeat"], item["group"], item["case"], item["shape"]): item
        for item in rows
    }
    require(len(by_key) == len(rows), f"{stage} row identity keys are duplicated")
    for item in [row for row in rows if row["lane"] == "alloc"]:
        counterpart = by_key.get(("native", item["repeat"], item["group"], item["case"], item["shape"]))
        require(counterpart is not None, f"{stage} allocation row lacks native counterpart")
        require(counterpart["identity_sha256"] == item["identity_sha256"],
                f"{stage} allocation/native identity differs for {item['case']}/{item['shape']}")
    identity_values = [
        {"sha256": digest, "value": identities[digest]}
        for digest in sorted(identities)
    ]
    return {
        "stage": stage,
        "source_manifest_sha256": sha256(folder / "source-manifest.json"),
        "source_manifest": manifest,
        "binaries": binaries,
        "job_counts": expected_jobs,
        "row_counts": {lane: len([row for row in rows if row["lane"] == lane]) for lane in LANES},
        "sample_counts": {
            lane: sum(row["samples"] for row in rows if row["lane"] == lane)
            for lane in LANES
        },
        "receipts": receipts,
        "native_rss": [rss_by_job[key] for key in sorted(rss_by_job)],
        "identities": identity_values,
        "rows": rows,
    }


def percent_change(first: float, second: float) -> float | None:
    if first == 0.0:
        return 0.0 if second == 0.0 else None
    return (second / first - 1.0) * 100.0


def comparison_record(
    first: Any, second: Any, *, lane: str, group: str, case: str | None,
    shape: str | None, repeat: int, metric: str,
) -> dict[str, Any]:
    finite_number(first, f"comparison {metric}.baseline")
    finite_number(second, f"comparison {metric}.candidate")
    change = percent_change(float(first), float(second))
    return {
        "lane": lane,
        "group": group,
        "case": case,
        "shape": shape,
        "repeat": repeat,
        "metric": metric,
        "baseline": first,
        "candidate": second,
        "delta": float(second) - float(first),
        "change_percent": change,
        "improvement_percent": -change if change is not None else None,
        "adverse_over_five_percent": change is None or change > DRIFT_THRESHOLD_PERCENT,
    }


def summary_comparisons(
    first: dict[str, Any], second: dict[str, Any], *, lane: str,
    group: str, case: str | None, shape: str | None, repeat: int,
    prefix: str,
) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    require(first.get("unit") == second.get("unit"),
            f"{prefix} summary units differ")
    for field in SUMMARY_FIELDS:
        output.append(comparison_record(
            first[field], second[field], lane=lane, group=group,
            case=case, shape=shape, repeat=repeat, metric=f"{prefix}.{field}",
        ))
    first_ci = first.get("confidence_interval_95")
    second_ci = second.get("confidence_interval_95")
    require(isinstance(first_ci, dict) and isinstance(second_ci, dict),
            f"{prefix} confidence interval is missing")
    require(first_ci.get("method") == second_ci.get("method"),
            f"{prefix} confidence interval method differs")
    for field in CI_FIELDS:
        output.append(comparison_record(
            first_ci[field], second_ci[field], lane=lane, group=group,
            case=case, shape=shape, repeat=repeat,
            metric=f"{prefix}.confidence_interval_95.{field}",
        ))
    return output


def compare_stages(
    baseline: dict[str, Any], candidate: dict[str, Any], plan: dict[str, Any],
) -> dict[str, Any]:
    base_rows = {
        (row["lane"], row["repeat"], row["group"], row["case"], row["shape"]): row
        for row in baseline["rows"]
    }
    cand_rows = {
        (row["lane"], row["repeat"], row["group"], row["case"], row["shape"]): row
        for row in candidate["rows"]
    }
    require(set(base_rows) == set(cand_rows), "baseline/candidate row keys differ")
    base_identities = {entry["sha256"]: entry["value"] for entry in baseline["identities"]}
    cand_identities = {entry["sha256"]: entry["value"] for entry in candidate["identities"]}
    native: list[dict[str, Any]] = []
    allocation: list[dict[str, Any]] = []
    adverse: list[dict[str, Any]] = []
    identity_rows: list[dict[str, Any]] = []
    for key in sorted(base_rows):
        left, right = base_rows[key], cand_rows[key]
        lane, repeat, group, case, shape = key
        require(left["identity_sha256"] in base_identities
                and right["identity_sha256"] in cand_identities,
                f"matched identity record missing for {key}")
        identity_equal = (
            left["identity_sha256"] == right["identity_sha256"]
            and base_identities[left["identity_sha256"]]
            == cand_identities[right["identity_sha256"]]
        )
        require(identity_equal, f"matched semantic/source/corpus identity differs for {key}")
        identity_rows.append({
            "lane": lane,
            "group": group,
            "case": case,
            "shape": shape,
            "repeat": repeat,
            "identity_equal": True,
            "baseline_identity_sha256": left["identity_sha256"],
            "candidate_identity_sha256": right["identity_sha256"],
        })
        if lane == "native":
            records = summary_comparisons(
                left["elapsed_ns"], right["elapsed_ns"], lane=lane,
                group=group, case=case, shape=shape, repeat=repeat,
                prefix="elapsed_ns",
            )
            adverse.extend(item for item in records if item["adverse_over_five_percent"])
            native.append({
                "lane": lane,
                "group": group,
                "case": case,
                "shape": shape,
                "repeat": repeat,
                "identity_equal": True,
                "comparisons": records,
                "baseline_evidence": left["evidence"],
                "candidate_evidence": right["evidence"],
            })
        else:
            left_summary = left["allocation"]["summary"]
            right_summary = right["allocation"]["summary"]
            require(set(left_summary) == set(right_summary),
                    f"allocation fields differ for {key}")
            records: list[dict[str, Any]] = []
            for field in ALLOCATION_FIELDS:
                records.extend(summary_comparisons(
                    left_summary[field], right_summary[field], lane=lane,
                    group=group, case=case, shape=shape, repeat=repeat,
                    prefix=f"allocation.{field}",
                ))
            adverse.extend(item for item in records if item["adverse_over_five_percent"])
            allocation.append({
                "lane": lane,
                "group": group,
                "case": case,
                "shape": shape,
                "repeat": repeat,
                "identity_equal": True,
                "status": left["allocation_status"],
                "scope": left["allocation_scope"],
                "comparisons": records,
                "baseline_samples": left["allocation"]["samples"],
                "candidate_samples": right["allocation"]["samples"],
                "baseline_evidence": left["evidence"],
                "candidate_evidence": right["evidence"],
            })
    base_rss = {item["job"]: item for item in baseline["native_rss"]}
    cand_rss = {item["job"]: item for item in candidate["native_rss"]}
    require(set(base_rss) == set(cand_rss), "native RSS job matrix differs")
    rss_comparisons: list[dict[str, Any]] = []
    for job in sorted(base_rss):
        left, right = base_rss[job], cand_rss[job]
        require(left["repeat"] == right["repeat"] and left["group"] == right["group"],
                f"native RSS job identity differs for {job}")
        records = [comparison_record(
            left[field], right[field], lane="native", group=left["group"],
            case=None, shape=None, repeat=left["repeat"], metric=f"rss.{field}",
        ) for field in RSS_FIELDS]
        adverse.extend(item for item in records if item["adverse_over_five_percent"])
        rss_comparisons.append({
            "job": job,
            "group": left["group"],
            "repeat": left["repeat"],
            "comparisons": records,
            "baseline_evidence": left["evidence"],
            "candidate_evidence": right["evidence"],
        })
    return {
        "identity_rows": identity_rows,
        "native": native,
        "allocation": allocation,
        "native_rss": rss_comparisons,
        "adverse_over_five_percent": adverse,
        "pair_counts": {
            "identity": len(identity_rows),
            "native": len(native),
            "allocation": len(allocation),
            "native_rss": len(rss_comparisons),
        },
    }


def _comparison_index(comparisons: dict[str, Any]) -> dict[tuple[Any, ...], dict[str, Any]]:
    output: dict[tuple[Any, ...], dict[str, Any]] = {}
    for item in comparisons["native"]:
        for record in item["comparisons"]:
            output[("native", item["group"], item["case"], item["shape"],
                   item["repeat"], record["metric"])] = record
    for item in comparisons["allocation"]:
        for record in item["comparisons"]:
            output[("alloc", item["group"], item["case"], item["shape"],
                   item["repeat"], record["metric"])] = record
    for item in comparisons["native_rss"]:
        for record in item["comparisons"]:
            output[("rss", item["group"], None, item["job"],
                   item["repeat"], record["metric"])] = record
    return output


def gate_check(
    index: dict[tuple[Any, ...], dict[str, Any]], key: tuple[Any, ...],
    *, criterion: str, maximum_change: float | None = None,
    minimum_improvement: float | None = None,
) -> dict[str, Any]:
    record = index.get(key)
    require(record is not None, f"missing gate comparison {key}")
    change = record["change_percent"]
    if maximum_change is not None:
        passed = change is not None and change <= maximum_change
    elif minimum_improvement is not None:
        passed = change is not None and change <= -minimum_improvement
    else:
        raise EvidenceError("gate has no threshold")
    return {
        "lane": record["lane"],
        "group": record["group"],
        "case": record["case"],
        "shape": record["shape"],
        "repeat": record["repeat"],
        "metric": record["metric"],
        "criterion": criterion,
        "pass": passed,
        "baseline": record["baseline"],
        "candidate": record["candidate"],
        "change_percent": change,
    }


def gate_group(name: str, description: str, checks: list[dict[str, Any]]) -> dict[str, Any]:
    return {
        "name": name,
        "description": description,
        "pass": bool(checks) and all(item["pass"] for item in checks),
        "check_count": len(checks),
        "checks": checks,
    }


def repeat_drift(
    stage_docs: dict[str, dict[str, Any]],
) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    all_records: list[dict[str, Any]] = []
    for stage in STAGES:
        document = stage_docs[stage]
        for lane in LANES:
            rows = [row for row in document["rows"] if row["lane"] == lane]
            by_key = {
                (row["group"], row["case"], row["shape"], row["repeat"]): row
                for row in rows
            }
            pairs = sorted({(row["group"], row["case"], row["shape"]) for row in rows})
            for group, case, shape in pairs:
                first = by_key[(group, case, shape, 1)]
                second = by_key[(group, case, shape, 2)]
                if lane == "native":
                    first_summary = first["elapsed_ns"]
                    second_summary = second["elapsed_ns"]
                    for field in SUMMARY_FIELDS:
                        change = percent_change(
                            float(first_summary[field]), float(second_summary[field])
                        )
                        all_records.append({
                            "stage": stage, "lane": lane, "group": group,
                            "case": case, "shape": shape,
                            "repeat_first": 1, "repeat_second": 2,
                            "metric": f"elapsed_ns.{field}",
                            "first": first_summary[field], "second": second_summary[field],
                            "delta": float(second_summary[field]) - float(first_summary[field]),
                            "change_percent": change,
                            "threshold_percent": DRIFT_THRESHOLD_PERCENT,
                            "over_five_percent": change is None or abs(change) > DRIFT_THRESHOLD_PERCENT,
                        })
                else:
                    for field in ALLOCATION_FIELDS:
                        first_summary = first["allocation"]["summary"][field]
                        second_summary = second["allocation"]["summary"][field]
                        for stat in SUMMARY_FIELDS:
                            change = percent_change(
                                float(first_summary[stat]), float(second_summary[stat])
                            )
                            all_records.append({
                                "stage": stage, "lane": lane, "group": group,
                                "case": case, "shape": shape,
                                "repeat_first": 1, "repeat_second": 2,
                                "metric": f"allocation.{field}.{stat}",
                                "first": first_summary[stat], "second": second_summary[stat],
                                "delta": float(second_summary[stat]) - float(first_summary[stat]),
                                "change_percent": change,
                                "threshold_percent": DRIFT_THRESHOLD_PERCENT,
                                "over_five_percent": change is None or abs(change) > DRIFT_THRESHOLD_PERCENT,
                            })
        rss = {item["job"]: item for item in document["native_rss"]}
        # Native RSS is one observation per job, rather than one repeated row.
        # It is emitted in the stage document and compared here without
        # duplicating it nine times for the XLS group.
        for job in sorted(rss):
            first = rss[job]
            # Reconstruct repeat pair from stage rows' RSS job references by
            # looking in the raw stage document's native_rss list.
            # The list contains both repeats under distinct names.
            candidates = [item for item in document["native_rss"]
                          if item["group"] == first["group"]]
            by_repeat = {item["repeat"]: item for item in candidates}
            if 1 not in by_repeat or 2 not in by_repeat:
                continue
            if first["repeat"] != 1:
                continue
            second = by_repeat[2]
            for field in RSS_FIELDS:
                change = percent_change(float(first[field]), float(second[field]))
                all_records.append({
                    "stage": stage, "lane": "native", "group": first["group"],
                    "case": None, "shape": None, "job": first["job"],
                    "repeat_first": 1, "repeat_second": 2,
                    "metric": f"rss.{field}", "first": first[field], "second": second[field],
                    "delta": float(second[field]) - float(first[field]),
                    "change_percent": change,
                    "threshold_percent": DRIFT_THRESHOLD_PERCENT,
                    "over_five_percent": change is None or abs(change) > DRIFT_THRESHOLD_PERCENT,
                })
    return all_records, [item for item in all_records if item["over_five_percent"]]


def main_gates(
    plan: dict[str, Any], baseline: dict[str, Any], candidate: dict[str, Any],
    comparisons: dict[str, Any],
) -> dict[str, Any]:
    index = _comparison_index(comparisons)
    primary_cases = plan["groups"]["xls"]["primary_cases"]
    primary_xls_p50: list[dict[str, Any]] = []
    for repeat in (1, 2):
        for case in primary_cases:
            primary_xls_p50.append(gate_check(
                index,
                ("native", "xls", case, "xls", repeat, "elapsed_ns.p50"),
                criterion="candidate change <= -3.00%",
                minimum_improvement=PRIMARY_IMPROVEMENT_PERCENT,
            ))
    primary_xls_mean: list[dict[str, Any]] = []
    for repeat in (1, 2):
        for case in primary_cases:
            primary_xls_mean.append(gate_check(
                index,
                ("native", "xls", case, "xls", repeat, "elapsed_ns.mean"),
                criterion="candidate change <= +5.00%",
                maximum_change=DRIFT_THRESHOLD_PERCENT,
            ))
    native_xls_controls: list[dict[str, Any]] = []
    for item in comparisons["native"]:
        if item["group"] != "xls" or item["case"] in primary_cases:
            continue
        for metric in ("elapsed_ns.p50", "elapsed_ns.mean"):
            native_xls_controls.append(gate_check(
                index,
                ("native", item["group"], item["case"], item["shape"],
                 item["repeat"], metric),
                criterion="candidate change <= +5.00%",
                maximum_change=DRIFT_THRESHOLD_PERCENT,
            ))
    native_cfb_controls: list[dict[str, Any]] = []
    for item in comparisons["native"]:
        if item["group"] != "cfb":
            continue
        for metric in ("elapsed_ns.p50", "elapsed_ns.mean"):
            native_cfb_controls.append(gate_check(
                index,
                ("native", item["group"], item["case"], item["shape"],
                 item["repeat"], metric),
                criterion="candidate change <= +5.00%",
                maximum_change=DRIFT_THRESHOLD_PERCENT,
            ))
    native_rss: list[dict[str, Any]] = []
    for item in comparisons["native_rss"]:
        native_rss.append(gate_check(
            index,
            ("rss", item["group"], None, item["job"], item["repeat"], "rss.max_rss_kib"),
            criterion="candidate change <= +5.00%",
            maximum_change=DRIFT_THRESHOLD_PERCENT,
        ))
    allocation_checks: list[dict[str, Any]] = []
    for item in comparisons["allocation"]:
        for field in ALLOCATION_GATE_FIELDS:
            basis = "max" if field == "incremental_region_peak_live_bytes" else "mean"
            allocation_checks.append(gate_check(
                index,
                ("alloc", item["group"], item["case"], item["shape"],
                 item["repeat"], f"allocation.{field}.{basis}"),
                criterion="candidate change <= +5.00%",
                maximum_change=DRIFT_THRESHOLD_PERCENT,
            ))
    groups = {
        "primary_xls_p50": gate_group(
            "primary_xls_p50",
            plan["admission"]["primary_xls_p50"],
            primary_xls_p50,
        ),
        "primary_xls_mean": gate_group(
            "primary_xls_mean",
            plan["admission"]["primary_xls_mean"],
            primary_xls_mean,
        ),
        "native_xls_controls": gate_group(
            "native_xls_controls",
            plan["admission"]["native_xls_controls"],
            native_xls_controls,
        ),
        "native_cfb_controls": gate_group(
            "native_cfb_controls",
            plan["admission"]["native_cfb_controls"],
            native_cfb_controls,
        ),
        "native_rss": gate_group(
            "native_rss",
            plan["admission"]["native_rss"],
            native_rss,
        ),
        "allocation": gate_group(
            "allocation",
            plan["admission"]["allocation"],
            allocation_checks,
        ),
    }
    groups["all_frozen_main_gates_pass"] = all(item["pass"] for item in groups.values())
    groups["external_controls_required"] = {
        "status": "pending",
        "validated_here": False,
        "required": ["profile", "correctness", "quality"],
        "reason": "profiles, source review, correctness, and quality are external to the metrics analyzer",
    }
    return groups


def analyze() -> dict[str, Any]:
    plan = plan_data()
    frozen = frozen_inputs(plan)
    stages = {stage: stage_analysis(stage, plan) for stage in STAGES}
    comparisons = compare_stages(stages["baseline"], stages["candidate"], plan)
    drift, drift_over_five = repeat_drift(stages)
    gates = main_gates(
        plan, stages["baseline"], stages["candidate"], comparisons
    )
    main_pass = gates["all_frozen_main_gates_pass"]
    disposition = (
        "pending external profile/correctness/quality evidence"
        if main_pass else "reject: one or more frozen main metrics gates failed"
    )
    return {
        "schema": SCHEMA,
        "status": "pass",
        "stage": "matched baseline/candidate",
        "scope": plan["scope"],
        "priority": plan["priority"],
        "performance_claim": "descriptive matched comparison; no production adoption claim",
        "disposition": disposition,
        "plan_sha256": sha256(HERE / "plan.json"),
        "run_sha256": sha256(HERE / "run.py"),
        "frozen_inputs": frozen,
        "prior_scope_note": plan["prior_scope_note"],
        "helpers": {
            "tools/summarize_crud_baseline.py": sha256(REPO / "tools/summarize_crud_baseline.py"),
            "tools/validate_perf_corpus_binding.py": sha256(REPO / "tools/validate_perf_corpus_binding.py"),
        },
        "source_manifests": {
            stage: stages[stage]["source_manifest"] for stage in STAGES
        },
        "host_scope": HOST_SCOPE,
        "binaries": {stage: stages[stage]["binaries"] for stage in STAGES},
        "job_counts": {stage: stages[stage]["job_counts"] for stage in STAGES},
        "row_counts": {stage: stages[stage]["row_counts"] for stage in STAGES},
        "sample_counts": {stage: stages[stage]["sample_counts"] for stage in STAGES},
        "identities": {
            stage: stages[stage]["identities"] for stage in STAGES
        },
        "native": {
            "timing_scope": "elapsed_ns from non-instrumented litchi-perf-baseline only",
            "reported_statistics": [*SUMMARY_FIELDS, "confidence_interval_95"],
            "baseline_rows": [row for row in stages["baseline"]["rows"] if row["lane"] == "native"],
            "candidate_rows": [row for row in stages["candidate"]["rows"] if row["lane"] == "native"],
            "baseline_rss": stages["baseline"]["native_rss"],
            "candidate_rss": stages["candidate"]["native_rss"],
            "comparison": comparisons["native"],
            "rss_comparison": comparisons["native_rss"],
            "instrumented_elapsed_excluded": True,
        },
        "allocation": {
            "scope": "operation_global_system_allocator",
            "reported_metrics": list(ALLOCATION_FIELDS),
            "instrumented_elapsed_excluded": True,
            "baseline_rows": [row for row in stages["baseline"]["rows"] if row["lane"] == "alloc"],
            "candidate_rows": [row for row in stages["candidate"]["rows"] if row["lane"] == "alloc"],
            "comparison": comparisons["allocation"],
        },
        "matched_identity": {
            "keys": ["case", "corpus", "source", "sink", "output_sha256", "catalog_identity"],
            "all_lanes_repeats_equal": all(item["identity_equal"] for item in comparisons["identity_rows"]),
            "comparison": comparisons["identity_rows"],
        },
        "comparisons": {
            "native": comparisons["native"],
            "native_rss": comparisons["native_rss"],
            "allocation": comparisons["allocation"],
            "adverse_over_five_percent": comparisons["adverse_over_five_percent"],
        },
        "repeat_drift": drift,
        "repeat_drift_over_five_percent": drift_over_five,
        "main_gates": gates,
        "limits": [
            "Allocator-instrumented elapsed samples are excluded from native latency interpretation.",
            "GNU time max RSS is a whole-child high-water observation, not an operation-local peak.",
            "Matched evidence is limited to the frozen CFB/XLS corpus and pinned CPU matrix.",
            "Profile, correctness, and quality evidence are external to this analyzer.",
            "The 0554 many-small positive target is not reused; many-small is a required CFB control in this frozen plan.",
            "No hardware, cold-cache, provider, scaling, ODF, or iWork claim follows from this matrix.",
        ],
    }


def write_identical(path: Path, value: dict[str, Any]) -> None:
    encoded = (json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n").encode("utf-8")
    try:
        with path.open("xb") as stream:
            stream.write(encoded)
    except FileExistsError:
        require(path.is_file() and not path.is_symlink(),
                f"refusing to read non-regular output {path}")
        require(path.read_bytes() == encoded,
                f"existing output differs from deterministic replay: {path}")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output_positional", nargs="?", type=Path)
    parser.add_argument("--output", dest="output_option", type=Path)
    parser.add_argument("--stage", choices=("matched",), default="matched")
    args = parser.parse_args(argv)
    if args.output_positional is not None and args.output_option is not None:
        parser.error("provide output either positionally or with --output")
    output = args.output_option or args.output_positional or HERE / "metrics-analysis.json"
    try:
        document = analyze()
        write_identical(output, document)
    except (EvidenceError, AssertionError, KeyError, OSError, TypeError, ValueError) as error:
        print(f"0555 metrics evidence check failed: {error}", file=sys.stderr)
        return 1
    print(f"0555 matched metrics {document['status']}: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
