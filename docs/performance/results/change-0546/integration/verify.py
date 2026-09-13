"""Read-only custody verifier for the integrated 0546 XLSX traversal campaign.

The pilot has several independently useful evidence lanes.  This verifier
keeps them separate while checking the bindings that make a comparison
meaningful: frozen inputs, source manifests, retained binaries, serialized
timing reports, and serial ABBA execution.  A failed build or test command is
retained as an auditable attempt; it is not silently treated as a successful
stage.  Missing later evidence is reported as ``incomplete`` until the final
decision, cleanup record, and seal exist.

The verifier never starts a build, benchmark, profiler, analyzer, or cleanup
operation.  Use ``python3 -B verify.py --precleanup`` while a campaign is in
progress and ``python3 -B verify.py --strict`` after sealing.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import subprocess
import sys
import tempfile
from typing import Any, Callable


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[4]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
GUARD_RUN = HERE / "guard_run.py"
QUALITY_PLAN = HERE / "quality-plan.json"
ALLOCATION_GATES = HERE / "allocation-gates.json"
PROTOCOL = HERE / "protocol.md"
ADR = HERE / "adr-manifest.json"
FROZEN = HERE / "frozen-inputs.json"
CANDIDATE_FROZEN = HERE / "candidate-frozen-inputs.json"
DECISION = HERE / "decision.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
RUNTIME_RESTORATION = HERE / "runtime-restoration.json"
TARGET = Path("/home/zhuhe/litchi-goal-0546-target")
CAP_DIR = HERE / "cap-boundary"
CAP_PLAN = CAP_DIR / "plan.json"
CAP_RUN = CAP_DIR / "cap_run.py"
CAP_ANALYZER = CAP_DIR / "analyze.py"
CAP_ANALYSIS = CAP_DIR / "cap-analysis.json"
CAP_EXAMPLE = "perf_cap_boundary"
CAP_BINARY_DIR = TARGET / "cap-boundary-binaries"
CAP_ENABLER = "crates/litchi-xlsx/examples/perf_cap_boundary.rs"
CAP_SIZES = (1, 2, 160, 164, 256)
CAP_REPEATS = (1, 2)
CAP_NATIVE_ORDER = (
    ("baseline", 1), ("candidate", 1),
    ("candidate", 2), ("baseline", 2),
)
CAP_TIME_FORMAT = (
    '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,'
    '"system_seconds":%S}'
)
FROZEN_INITIAL = HERE / "frozen-inputs.initial.json"
CONDITIONAL_FROZEN = HERE / "conditional-frozen-inputs.json"
SYMBOL_OBSERVATION = HERE / "symbol-observation.json"
EAGER_PLAN = HERE / "eager-plan.json"
EAGER_RUN = HERE / "eager_run.py"
PROFILE_ANALYSIS = HERE / "planning-profile-analysis.json"
HARDWARE_ANALYSIS = HERE / "hardware-analysis.json"
EAGER_ANALYSIS = HERE / "eager-analysis.json"
CONDITIONAL_FROZEN_FILES = {
    "plan.json", "run.py", "eager-plan.json", "eager_run.py",
    "analyze_eager.py", "analyze_planning.py", "analyze_hardware.py",
    "conditional_phase.py", "comparison.json", "allocation-analysis.json",
    "guard-analysis.json", "symbol-observation.json",
    "cap-boundary/cap-analysis.json",
}
EAGER_TIME_FORMAT = (
    '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,'
    '"system_seconds":%S}'
)

SOURCE_EXACT = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml"}
SOURCE_PREFIXES = (".cargo/", "crates/", "tools/perf-baseline/")
ADR_ROOT = "docs/adr/"
PRIMARY_CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
MAIN_QUALITY_REQUIRED = {
    "quality-tests",
    "quality-check",
    "quality-clippy",
    "quality-rustdoc",
    "quality-fmt",
    "quality-boundaries",
    "quality-guard-clippy",
    "quality-guard-fmt",
    "quality-cap-clippy",
}
EXPECTED_HARNESS_PATHS = {
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/bin/xlsx_planning_guard.rs",
    "tools/perf-baseline/src/xlsx_planning_guard.rs",
}
TIMING_STATS = ("p50", "p95", "p99", "mean")
TEST_RESULT = re.compile(
    r"test result:\s+(ok|FAILED)\.\s+(\d+) passed;\s+(\d+) failed;"
    r"\s+(\d+) ignored;\s+(\d+) measured;\s+(\d+) filtered out"
)
HEX64 = re.compile(r"[0-9a-f]{64}\Z")
HEX40 = re.compile(r"[0-9a-f]{40}\Z")


class EvidenceError(ValueError):
    """Malformed or contradictory retained evidence."""


class Pending(EvidenceError):
    """Evidence needed by a later campaign phase has not arrived."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence bundle: {path}") from error


def display_path(path: Path) -> str:
    try:
        return relative(path)
    except EvidenceError:
        return path.as_posix()


def need(path: Path, label: str | None = None, *, directory: bool = False) -> Path:
    label = label or display_path(path)
    if not path.exists():
        raise Pending(f"{label} is missing")
    require(not path.is_symlink(), f"{label} is a symlink")
    if directory:
        require(path.is_dir(), f"{label} is not a directory")
    else:
        require(path.is_file(), f"{label} is not a regular file")
    return path


def read_bytes(path: Path, label: str | None = None) -> bytes:
    need(path, label)
    try:
        return path.read_bytes()
    except OSError as error:
        raise EvidenceError(f"cannot read {label or display_path(path)}: {error}") from error


def read_text(path: Path, label: str | None = None) -> str:
    try:
        return read_bytes(path, label).decode("utf-8")
    except UnicodeDecodeError as error:
        raise EvidenceError(f"{label or display_path(path)} is not UTF-8") from error


def read_json(path: Path, label: str | None = None) -> Any:
    label = label or display_path(path)
    try:
        return json.loads(read_text(path, label))
    except json.JSONDecodeError as error:
        raise EvidenceError(f"cannot parse {label}: {error}") from error


def sha(path: Path) -> str:
    need(path, f"artifact for hashing: {display_path(path)}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        raise EvidenceError(f"cannot hash {display_path(path)}: {error}") from error
    return digest.hexdigest()


def valid_digest(value: Any, *, length: int = 64) -> bool:
    return isinstance(value, str) and bool(
        (HEX64 if length == 64 else HEX40).fullmatch(value)
    )


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute(),
            f"{label} is not a safe relative path")
    path = Path(value)
    require(path.as_posix() == value and "." not in path.parts and ".." not in path.parts,
            f"{label} escapes its root")
    return value


def parse_time(value: Any, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} is not a timestamp")
    try:
        result = dt.datetime.fromisoformat(value)
    except ValueError as error:
        raise EvidenceError(f"{label} is invalid") from error
    require(result.tzinfo is not None, f"{label} has no timezone")
    return result


def interval(value: dict[str, Any], label: str) -> tuple[dt.datetime, dt.datetime]:
    require(isinstance(value, dict), f"{label} is not an object")
    start = parse_time(value.get("start_utc"), f"{label}.start_utc")
    end = parse_time(value.get("end_utc"), f"{label}.end_utc")
    seconds = value.get("seconds")
    require(isinstance(seconds, (int, float)) and not isinstance(seconds, bool)
            and math.isfinite(float(seconds)) and seconds >= 0.0 and end > start,
            f"{label} interval is invalid")
    # ``run.py`` records the timestamps with wall-clock UTC and measures the
    # duration with a monotonic clock.  Bind the two serialized views while
    # allowing a small scheduling/clock-read difference.  A large mismatch
    # would make receipt ordering auditable but its claimed duration opaque.
    wall_seconds = (end - start).total_seconds()
    tolerance = max(0.05, wall_seconds * 0.01)
    require(abs(float(seconds) - wall_seconds) <= tolerance,
            f"{label} seconds does not match its UTC interval")
    return start, end


def git_output(args: list[str], *, input_data: bytes | None = None) -> bytes:
    try:
        return subprocess.check_output(args, cwd=REPO, input=input_data, stderr=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        detail = getattr(error, "stderr", b"")
        raise EvidenceError(
            f"Git command failed ({' '.join(args)}): {detail.decode(errors='replace')[-2000:]}"
        ) from error


def source_name(name: str) -> bool:
    return name in SOURCE_EXACT or any(name.startswith(prefix) for prefix in SOURCE_PREFIXES)


def source_inventory() -> set[str]:
    tracked = git_output([
        "git", "ls-files", "-z", "--", "crates", "tools/perf-baseline", "Cargo.toml",
        "Cargo.lock", ".cargo", "rust-toolchain.toml",
    ]).split(b"\0")
    untracked = git_output([
        "git", "ls-files", "--others", "--exclude-standard", "-z", "--",
        "crates", "tools/perf-baseline",
    ]).split(b"\0")
    names = {item.decode("utf-8") for item in tracked if item}
    names.update(item.decode("utf-8") for item in untracked if item and item.endswith(b".rs"))
    return {
        name for name in names
        if source_name(name) and (REPO / name).is_file() and not (REPO / name).is_symlink()
    }


def revision_manifest(revision: str) -> dict[str, str]:
    raw = git_output([
        "git", "ls-tree", "-r", "-z", revision, "--", "crates", "tools/perf-baseline",
        "Cargo.toml", "Cargo.lock", ".cargo", "rust-toolchain.toml",
    ])
    entries: list[tuple[str, bytes]] = []
    for item in raw.split(b"\0"):
        if not item:
            continue
        try:
            metadata, encoded_name = item.split(b"\t", 1)
            fields = metadata.split()
            require(len(fields) == 3 and fields[1] == b"blob",
                    "Git source tree entry is malformed")
            name = encoded_name.decode("utf-8")
        except (UnicodeDecodeError, ValueError) as error:
            raise EvidenceError("Git source tree entry is malformed") from error
        require(source_name(name), f"Git source tree contains out-of-scope source: {name}")
        entries.append((name, fields[2]))
    require(entries, "Git source tree is empty")
    request = b"".join(oid + b"\n" for _, oid in entries)
    raw = git_output(["git", "cat-file", "--batch"], input_data=request)
    result: dict[str, str] = {}
    position = 0
    for name, expected_oid in entries:
        end = raw.find(b"\n", position)
        require(end >= 0, "Git source object response is truncated")
        fields = raw[position:end].split()
        require(len(fields) == 3 and fields[0] == expected_oid and fields[1] == b"blob",
                "Git source object response is malformed")
        position = end + 1
        try:
            length = int(fields[2])
        except ValueError as error:
            raise EvidenceError("Git source object length is malformed") from error
        data = raw[position:position + length]
        require(len(data) == length, "Git source object is truncated")
        result[name] = hashlib.sha256(data).hexdigest()
        position += length
        require(raw[position:position + 1] == b"\n", "Git source object separator is missing")
        position += 1
    require(position == len(raw), "Git source object response has trailing data")
    require(len(result) == len(entries), "Git source tree repeats a path")
    return result


def source_manifest(path: Path) -> dict[str, str]:
    value = read_json(path, f"{display_path(path)} source manifest")
    require(isinstance(value, dict) and value,
            f"{display_path(path)} is not a nonempty source manifest")
    result: dict[str, str] = {}
    for name, digest in value.items():
        safe_relative(name, f"{display_path(path)} source path")
        require(source_name(name) and valid_digest(digest),
                f"{display_path(path)} has an invalid source entry: {name}")
        require(name not in result, f"{display_path(path)} repeats {name}")
        result[name] = digest
    return result


def changed_paths(left: dict[str, str], right: dict[str, str]) -> set[str]:
    return {
        name for name in set(left) | set(right)
        if left.get(name) != right.get(name)
    }


def path_in_roots(name: str, roots: list[str]) -> bool:
    return any(name.startswith(root.rstrip("/") + "/") for root in roots)


def shared_test_prefixes(plan: dict[str, Any]) -> list[str]:
    value = plan.get("shared_test_source_prefixes")
    require(isinstance(value, list) and value,
            "plan.shared_test_source_prefixes is not a nonempty list")
    result: list[str] = []
    for prefix in value:
        safe_relative(prefix, "shared test source prefix")
        require(prefix.startswith("crates/litchi-xlsx/tests/"),
                f"shared test source prefix is outside XLSX tests: {prefix}")
        require(prefix == prefix.rstrip("/") and prefix,
                f"shared test source prefix is malformed: {prefix}")
        result.append(prefix)
    require(len(set(result)) == len(result),
            "plan.shared_test_source_prefixes repeats a prefix")
    return result


def is_shared_test_path(name: str, plan: dict[str, Any]) -> bool:
    return any(name == prefix or name == prefix + ".rs"
               or name.startswith(prefix + "/")
               for prefix in shared_test_prefixes(plan))


def shared_source_paths(snapshot: dict[str, str], revision: dict[str, str],
                        plan: dict[str, Any]) -> set[str]:
    return {name for name in set(snapshot) | set(revision)
            if is_shared_test_path(name, plan)}


def plan_data() -> dict[str, Any]:
    value = read_json(PLAN, "plan.json")
    require(isinstance(value, dict), "plan is not an object")
    revision = value.get("revision")
    require(valid_digest(revision, length=40), "plan revision is malformed")
    try:
        subprocess.run(["git", "cat-file", "-e", revision + "^{commit}"], cwd=REPO,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise EvidenceError("plan revision is not a Git commit") from error
    created = parse_time(value.get("created_utc"), "plan.created_utc")
    status = value.get("status")
    require(isinstance(status, str) and status.startswith("frozen-before"),
            "plan is not frozen before builds and captures")
    priority = value.get("priority")
    if priority is not None:
        require(isinstance(priority, str) and "OLE2/OOXML" in priority and "ODF" in priority,
                "plan priority does not preserve OLE2/OOXML first and ODF deferral")
    roots = value.get("candidate_source_roots")
    require(isinstance(roots, list) and roots, "candidate_source_roots is empty")
    clean_roots: list[str] = []
    for root in roots:
        safe_relative(root.rstrip("/"), "candidate source root")
        require(root.endswith("/") and root.startswith("crates/"),
                f"candidate source root is invalid: {root}")
        require((REPO / root.rstrip("/")).is_dir(),
                f"candidate source root is not a directory: {root}")
        clean_roots.append(root)
    require(len(set(clean_roots)) == len(clean_roots), "candidate source roots repeat")
    require(value.get("owned_paths") == [str(TARGET)], "plan owned paths differ")
    cpu = value.get("cpu")
    require(isinstance(cpu, int) and not isinstance(cpu, bool) and cpu >= 0,
            "plan.cpu is not a valid nonnegative integer")
    primary = value.get("primary")
    require(isinstance(primary, dict), "plan.primary is not an object")
    require(primary.get("case") == PRIMARY_CASE, "plan primary case differs")
    _positive_integer(primary.get("repeats"), "plan.primary.repeats")
    _positive_integer(primary.get("warmup"), "plan.primary.warmup", allow_zero=True)
    _positive_integer(primary.get("samples"), "plan.primary.samples")
    _shape_list(primary.get("shapes"), "plan.primary.shapes")
    guards = value.get("guards")
    require(isinstance(guards, list) and guards, "plan.guards is empty")
    for index, guard in enumerate(guards):
        require(isinstance(guard, dict), f"plan.guards[{index}] is not an object")
        require(isinstance(guard.get("case"), str) and guard["case"],
                f"plan.guards[{index}].case is invalid")
        _shape_list(guard.get("shapes"), f"plan.guards[{index}].shapes")
    conditional = value.get("conditional_lanes")
    require(isinstance(conditional, dict)
            and set(conditional) == {"eager", "hardware", "profile"},
            "plan.conditional_lanes inventory differs")
    for name, description in conditional.items():
        require(isinstance(description, str) and description.strip(),
                f"plan.conditional_lanes.{name} is invalid")
    for key in ("guard_repeats", "guard_warmup", "guard_samples"):
        _positive_integer(value.get(key), f"plan.{key}", allow_zero=key.endswith("warmup"))
    allocation = value.get("allocation")
    require(isinstance(allocation, dict), "plan.allocation is not an object")
    _positive_integer(allocation.get("repeats"), "plan.allocation.repeats")
    _positive_integer(allocation.get("warmup"), "plan.allocation.warmup", allow_zero=True)
    _positive_integer(allocation.get("samples"), "plan.allocation.samples")
    _shape_list(allocation.get("shapes"), "plan.allocation.shapes")
    require(isinstance(allocation.get("scope"), str) and allocation["scope"],
            "plan.allocation.scope is invalid")
    refusal = value.get("refusal_guard")
    require(isinstance(refusal, dict), "plan.refusal_guard is not an object")
    require(refusal.get("bin") == "xlsx_planning_guard",
            "plan refusal guard binary differs")
    _shape_list(refusal.get("shapes"), "plan.refusal_guard.shapes")
    cases = refusal.get("cases")
    require(isinstance(cases, list) and cases == ["valid", "late-validator", "late-raw"],
            "plan refusal guard cases differ")
    _positive_integer(refusal.get("repeats"), "plan.refusal_guard.repeats")
    for key in ("native_warmup", "allocation_warmup"):
        _positive_integer(refusal.get(key), f"plan.refusal_guard.{key}", allow_zero=True)
    for key in ("native_samples", "allocation_samples"):
        _positive_integer(refusal.get(key), f"plan.refusal_guard.{key}")
    for key in ("native_invalid_max_baseline_valid_ratio",
                "allocation_invalid_peak_max_baseline_valid_ratio",
                "ordinary_valid_max_regression_percent"):
        _finite_nonnegative(refusal.get(key), f"plan.refusal_guard.{key}")
    require(isinstance(refusal.get("scope"), str) and refusal["scope"],
            "plan.refusal_guard.scope is invalid")
    gates = value.get("gates")
    require(isinstance(gates, dict), "plan.gates is not an object")
    expected_gates = {
        "total_p50_reduction_percent", "total_mean_reduction_percent",
        "planning_p50_reduction_percent", "planning_ir_reduction_percent",
        "require_every_shape_repeat",
    }
    require(set(gates) == expected_gates, "plan gate inventory differs")
    for key in expected_gates - {"require_every_shape_repeat"}:
        _finite_nonnegative(gates.get(key), f"plan.gates.{key}")
    require(gates.get("require_every_shape_repeat") is True,
            "plan.gates.require_every_shape_repeat must be true")
    profile = value.get("profile")
    require(isinstance(profile, dict), "plan.profile is not an object")
    _shape_list(profile.get("shapes"), "plan.profile.shapes")
    _positive_integer(profile.get("repeats"), "plan.profile.repeats")
    _positive_integer(profile.get("warmup"), "plan.profile.warmup", allow_zero=True)
    _positive_integer(profile.get("samples"), "plan.profile.samples")
    require(isinstance(profile.get("owner"), str) and profile["owner"],
            "plan.profile.owner is invalid")
    require(isinstance(profile.get("scope"), str) and profile["scope"],
            "plan.profile.scope is invalid")
    hardware = value.get("hardware")
    require(isinstance(hardware, dict), "plan.hardware is not an object")
    _shape_list(hardware.get("shapes"), "plan.hardware.shapes")
    _positive_integer(hardware.get("repeats"), "plan.hardware.repeats")
    _positive_integer(hardware.get("warmup"), "plan.hardware.warmup", allow_zero=True)
    _positive_integer(hardware.get("samples"), "plan.hardware.samples")
    require(isinstance(hardware.get("events"), str) and hardware["events"],
            "plan.hardware.events is invalid")
    require(isinstance(hardware.get("scope"), str) and hardware["scope"],
            "plan.hardware.scope is invalid")
    planned = value.get("candidate_files")
    if planned is not None:
        require(isinstance(planned, list) and planned,
                "plan.candidate_files is not a nonempty list")
        for name in planned:
            safe_relative(name, "plan candidate file")
            require(path_in_roots(name, clean_roots),
                    f"plan candidate file is outside candidate roots: {name}")
        require(len(set(planned)) == len(planned), "plan.candidate_files repeats a path")
    if "candidate_patch_sha256" in value:
        require(valid_digest(value["candidate_patch_sha256"]),
                "plan candidate patch digest is malformed")
    shared_test_prefixes(value)
    cap = value.get("cap_boundary")
    require(isinstance(cap, dict), "plan.cap_boundary is not an object")
    require(cap.get("sizes") == list(CAP_SIZES)
            and cap.get("repeats") == len(CAP_REPEATS)
            and cap.get("warmup") == 10 and cap.get("samples") == 100
            and cap.get("required_before_conditional_lanes") is True,
            "plan.cap_boundary matrix differs")
    _finite_nonnegative(cap.get("candidate_planning_p50_mean_max_ratio"),
                        "plan.cap_boundary.candidate_planning_p50_mean_max_ratio")
    require(float(cap["candidate_planning_p50_mean_max_ratio"]) == 1.05,
            "plan.cap_boundary ratio differs")
    enabler = value.get("harness_enabler")
    require(isinstance(enabler, str) and "perf_cap_boundary" in enabler
            and ("sparse" in enabler.lower() or "1/2" in enabler
                 or "comment" in enabler.lower()),
            "plan does not describe the retained cap benchmark enabler")
    return {**value, "_created": created, "_roots": clean_roots}


def _positive_integer(value: Any, label: str, *, allow_zero: bool = False) -> None:
    require(isinstance(value, int) and not isinstance(value, bool)
            and (value >= 0 if allow_zero else value > 0),
            f"{label} is not a valid nonnegative/positive integer")


def _finite_nonnegative(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)) and float(value) >= 0.0,
            f"{label} is not a finite nonnegative number")


def _shape_list(value: Any, label: str) -> None:
    require(isinstance(value, list) and value and all(isinstance(item, str) and item for item in value),
            f"{label} is not a nonempty string list")
    require(len(set(value)) == len(value), f"{label} repeats a shape")


def quality_commands() -> dict[str, list[str]]:
    value = read_json(QUALITY_PLAN, "quality-plan.json")
    require(isinstance(value, dict) and isinstance(value.get("commands"), dict),
            "quality plan is not an object")
    commands = value["commands"]
    require(MAIN_QUALITY_REQUIRED <= set(commands),
            "quality plan omits required production quality checks")
    result: dict[str, list[str]] = {}
    for name, command in commands.items():
        require(isinstance(name, str) and name and isinstance(command, list) and command
                and all(isinstance(item, str) and item for item in command),
                f"quality command {name} is malformed")
        result[name] = command
    require("cargo" in result["quality-tests"] and "test" in result["quality-tests"],
            "quality-tests is not a cargo test command")
    return result


def allocation_gate_data() -> dict[str, Any]:
    value = read_json(ALLOCATION_GATES, "allocation-gates.json")
    require(isinstance(value, dict), "allocation gates are not an object")
    required = {
        "status", "planning_allocated_bytes_max_increase_percent",
        "planning_incremental_peak_live_bytes_max_increase_percent",
        "comparison", "scope", "created_utc",
    }
    require(set(value) == required, "allocation gate inventory differs")
    require(value.get("status") in {
        "frozen-before-first-allocation-capture",
        "frozen-before-build-and-capture",
    },
            "allocation gates were not frozen before capture")
    require("planning_allocation_calls_reduction_percent" not in value,
            "allocation calls must remain diagnostic in 0546")
    for key in required - {"status", "comparison", "scope", "created_utc"}:
        _finite_nonnegative(value[key], f"allocation-gates.{key}")
        require(float(value[key]) == 1.0, f"allocation-gates.{key} must be 1.0")
    require(isinstance(value["scope"], str) and value["scope"],
            "allocation-gates.scope is invalid")
    require(value["comparison"] == "maximum candidate <= 1.01 * minimum baseline in every shape/repeat",
            "allocation comparison rule differs")
    return {**value, "_created": parse_time(value["created_utc"], "allocation-gates.created_utc")}


def frozen_inputs(plan: dict[str, Any]) -> tuple[dict[str, Any], dt.datetime]:
    """Validate either the integrated or legacy frozen-input envelope.

    The integrated 0546 campaign records a compact ``sha256`` map and binds
    the diagnostic sibling freeze.  Earlier campaign bundles used a
    ``files`` map plus a shared-enabler object.  Both forms are checked when
    present; no historical failed-attempt envelope is required for 0546.
    """
    value = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(value, dict), "frozen input envelope is malformed")
    created = parse_time(value.get("created_utc", value.get("utc")),
                         "frozen-inputs.created_utc")
    require(created >= plan["_created"], "frozen inputs predate the plan")
    files = value.get("files", value.get("sha256"))
    require(isinstance(files, dict), "frozen input digest map is malformed")
    # These are the custody authorities used before the first baseline
    # capture.  Additional entries are allowed and are checked below.
    required = {
        "plan.json", "run.py", "guard_run.py", "baseline_phase.py",
        "candidate_phase.py", "quality-plan.json", "environment.json",
        "analyze.py", "analyze_allocation.py", "analyze_guard.py",
        "cap-boundary/plan.json", "cap-boundary/cap_run.py",
        "cap-boundary/analyze.py",
    }
    require(required <= set(files),
            f"frozen inputs omit required files: {sorted(required - set(files))}")
    for name, digest in files.items():
        safe_relative(name, "frozen input path")
        require(valid_digest(digest), f"frozen input digest is malformed: {name}")
        target = HERE / name
        require(target.is_file() and not target.is_symlink(),
                f"frozen input file is missing: {name}")
        actual_digest = sha(target)
        if actual_digest != digest:
            # The two analysis modules were rebound after the initial freeze
            # solely to resolve the nested 0521 helper.  The amendment keeps
            # both hashes, so validate that explicit custody record instead
            # of treating the old analyzer hash as the live input.
            amendment = HERE / "analysis-amendment.json"
            changes = {}
            if amendment.is_file() and not amendment.is_symlink():
                amended = read_json(amendment, "analysis-amendment.json")
                if isinstance(amended, dict):
                    changes = amended.get("changes", {})
            change = changes.get(name) if isinstance(changes, dict) else None
            require(isinstance(change, dict)
                    and digest in (change.get("initial_sha256"), change.get("final_sha256"))
                    and change.get("final_sha256") == actual_digest,
                    f"frozen input hash differs: {name}")
    amendment_path = HERE / "analysis-amendment.json"
    if amendment_path.is_file() and not amendment_path.is_symlink():
        amendment = read_json(amendment_path, "analysis-amendment.json")
        require(isinstance(amendment, dict)
                and isinstance(amendment.get("changes"), dict),
                "analysis amendment is malformed")
        for name, change in amendment["changes"].items():
            safe_relative(name, "analysis amendment path")
            require(isinstance(change, dict)
                    and valid_digest(change.get("initial_sha256"))
                    and valid_digest(change.get("final_sha256"))
                    and files.get(name) in (change["initial_sha256"], change["final_sha256"]),
                    f"analysis amendment hashes are malformed: {name}")
            current = HERE / name
            require(current.is_file() and sha(current) == change["final_sha256"],
                    f"analysis amendment final hash differs: {name}")
            original = HERE / str(change.get("original", ""))
            require(original.is_file() and sha(original) == change["initial_sha256"],
                    f"analysis amendment original hash differs: {name}")
        if "allocation_gates_sha256" in amendment:
            require(valid_digest(amendment["allocation_gates_sha256"])
                    and ALLOCATION_GATES.is_file()
                    and amendment["allocation_gates_sha256"] == sha(ALLOCATION_GATES),
                    "analysis amendment allocation gate hash differs")
    enabler = value.get("shared_enabler")
    enabler_digest = value.get("enabler_sha256")
    if isinstance(enabler, dict):
        require(enabler.get("path") == CAP_ENABLER
                and valid_digest(enabler.get("sha256")),
                "frozen input shared enabler binding is malformed")
        enabler_digest = enabler["sha256"]
    require(valid_digest(enabler_digest),
            "frozen input shared enabler hash is malformed")
    require((REPO / CAP_ENABLER).is_file()
            and sha(REPO / CAP_ENABLER) == enabler_digest,
            "frozen input shared enabler hash differs")
    if "cap-boundary/guard.rs" in files:
        require(files["cap-boundary/guard.rs"] == enabler_digest,
                "frozen input cap guard does not bind shared enabler")
    # The integrated envelope may bind the sibling diagnostic freeze without
    # copying that independent evidence into this nested bundle.
    diagnostic_digest = value.get("diagnostic_freeze_sha256")
    if diagnostic_digest is not None:
        require(valid_digest(diagnostic_digest),
                "frozen diagnostic freeze hash is malformed")
        diagnostic = next((candidate for candidate in (
            HERE.parent / "freeze.json", HERE / "diagnostic-freeze.json",
        ) if candidate.is_file() and not candidate.is_symlink()), None)
        require(diagnostic is not None and sha(diagnostic) == diagnostic_digest,
                "frozen diagnostic freeze hash differs")
    # Retain compatibility with a legacy corrected envelope only when it is
    # explicitly in that legacy ``files`` form.  A fresh integrated campaign
    # never needs frozen-inputs.initial.json or failed-build custody.
    if isinstance(value.get("files"), dict) and FROZEN_INITIAL.exists():
        initial = read_json(FROZEN_INITIAL, "frozen-inputs.initial.json")
        require(isinstance(initial, dict) and isinstance(initial.get("files"), dict),
                "initial frozen input envelope is malformed")
        initial_created = parse_time(initial.get("created_utc"),
                                    "frozen-inputs.initial.created_utc")
        require(initial_created <= created,
                "initial frozen inputs follow corrected freeze")
        require(value.get("prior_frozen_inputs_sha256") == sha(FROZEN_INITIAL),
                "corrected frozen inputs do not bind initial freeze")
        initial_files = initial["files"]
        for name, digest in initial_files.items():
            safe_relative(name, "initial frozen input path")
            require(valid_digest(digest),
                    f"initial frozen input digest is malformed: {name}")
            target = (CAP_DIR / "guard.initial.rs"
                      if name == "cap-boundary/guard.rs" else HERE / name)
            require(target.is_file() and not target.is_symlink()
                    and sha(target) == digest,
                    f"initial frozen input hash differs: {name}")
        require(initial_files.get("cap-boundary/guard.rs") ==
                sha(CAP_DIR / "guard.initial.rs"),
                "initial frozen inputs do not bind retained failed guard")
        require(initial.get("shared_enabler", {}).get("path") == CAP_ENABLER
                and initial.get("shared_enabler", {}).get("sha256") ==
                initial_files.get("cap-boundary/guard.rs"),
                "initial shared enabler binding differs")
        require(files["cap-boundary/guard.rs"] !=
                initial_files["cap-boundary/guard.rs"],
                "corrected freeze did not change the retained guard binding")
        correction = value.get("correction")
        require(isinstance(correction, str) and "failed" in correction.lower()
                and "clippy" in correction.lower(),
                "corrected freeze lacks the failed cap-clippy explanation")
    return value, created


def candidate_frozen_inputs(plan: dict[str, Any], baseline: dict[str, str],
                            frozen_created: dt.datetime) -> tuple[dict[str, Any], dt.datetime]:
    """Validate the second freeze that binds the applied candidate patch.

    The first freeze predates the candidate application and therefore cannot
    bind the reviewed candidate patch.  The candidate envelope records the
    exact patch/review inputs and source hashes used for the candidate freeze;
    every candidate receipt must occur after this point.
    """
    if not CANDIDATE_FROZEN.exists():
        # The integrated runner binds candidate source manifests and the
        # reviewed patch directly.  A separate envelope is optional; use the
        # baseline freeze as the conservative lower bound until that stage is
        # present.
        return {"status": "implicit-source-manifest-freeze"}, frozen_created
    value = read_json(CANDIDATE_FROZEN, "candidate-frozen-inputs.json")
    require(isinstance(value, dict), "candidate frozen input envelope is malformed")
    created = parse_time(value.get("created_utc", value.get("utc")),
                         "candidate-frozen-inputs.created_utc")
    require(created >= frozen_created, "candidate freeze predates the baseline freeze")
    files = value.get("files", value.get("sha256", {}))
    require(isinstance(files, dict), "candidate frozen input digest map is malformed")
    for name, digest in files.items():
        safe_relative(name, "candidate frozen input path")
        require(valid_digest(digest), f"candidate frozen input digest is malformed: {name}")
        candidates = [HERE / name]
        if name.startswith("candidate-sources/"):
            candidates.append(HERE / name)
        else:
            candidates.append(HERE / "candidate-sources" / Path(name).name)
        existing = [candidate for candidate in candidates
                    if candidate.is_file() and not candidate.is_symlink()]
        if len(existing) > 1:
            digests = {sha(candidate) for candidate in existing}
            if len(digests) != 1:
                adaptation_path = HERE / "format-adaptation.json"
                require(adaptation_path.is_file() and not adaptation_path.is_symlink(),
                        f"candidate frozen duplicate differs: {name}")
                adaptation = read_json(adaptation_path, "format-adaptation.json")
                prepared = adaptation.get("prepared_patch_sha256") if isinstance(adaptation, dict) else None
                applied = adaptation.get("applied_patch_sha256") if isinstance(adaptation, dict) else None
                require(valid_digest(prepared) and valid_digest(applied)
                        and digests == {prepared, applied},
                        f"candidate frozen duplicate differs: {name}")
        target = existing[0] if existing else None
        require(target is not None and not target.is_symlink(),
                f"candidate frozen input file is missing: {name}")
        require(sha(target) == digest, f"candidate frozen input hash differs: {name}")
    baseline_path = HERE / "baseline/source-manifest.json"
    if "baseline_manifest_sha256" in value:
        require(value["baseline_manifest_sha256"] == sha(baseline_path),
                "candidate freeze baseline manifest binding differs")
    sources = value.get("candidate_sources")
    candidate_manifest_path = HERE / "candidate/source-manifest.json"
    if sources is not None:
        require(isinstance(sources, dict) and sources,
                "candidate frozen source inventory is malformed")
        roots = plan["_roots"]
        for name, digest in sources.items():
            safe_relative(name, "candidate frozen source path")
            require(path_in_roots(name, roots) and valid_digest(digest),
                    f"candidate frozen source entry is malformed: {name}")
        if candidate_manifest_path.is_file():
            candidate_manifest = source_manifest(candidate_manifest_path)
            changed = changed_paths(candidate_manifest, baseline)
            require(set(sources) == changed,
                    "candidate frozen source inventory differs from candidate manifest diff")
            planned = plan.get("candidate_files")
            if planned is not None:
                require(set(sources) == set(planned),
                        "candidate frozen source inventory differs from plan")
            for name, digest in sources.items():
                require(candidate_manifest.get(name) == digest,
                        f"candidate frozen source hash differs: {name}")
    return value, created


def validate_candidate_source_hashes(plan: dict[str, Any],
                                     baseline: dict[str, str]) -> dict[str, Any]:
    """Check the review-time five-file source inventory when retained.

    This compact inventory is independent of the later candidate checkout.
    It lets the verifier establish patch/source custody while the baseline
    lane is still running, then cross-checks the same hashes against the
    candidate stage once that stage exists.
    """
    path = HERE / "candidate-source-hashes.json"
    if not path.exists():
        return {"status": "not-present"}
    value = read_json(path, "candidate-source-hashes.json")
    require(isinstance(value, dict), "candidate source hashes are not an object")
    require(value.get("schema") == "litchi.xlsx.change-0546-candidate-source-hashes.v1",
            "candidate source hash schema differs")
    require(value.get("baseline_revision") == plan["revision"],
            "candidate source hash baseline revision differs")
    candidate_patch = value.get("candidate_patch")
    require(isinstance(candidate_patch, dict)
            and candidate_patch.get("path") == "candidate.patch"
            and valid_digest(candidate_patch.get("sha256")),
            "candidate source hash patch binding is malformed")
    reviewed, _ = measured_candidate_patches()
    require(reviewed.is_file() and sha(reviewed) == candidate_patch["sha256"],
            "candidate source hash patch digest differs")
    patch_names = patch_paths(read_bytes(reviewed, display_path(reviewed)),
                             display_path(reviewed))
    planned = plan.get("candidate_files")
    require(isinstance(planned, list) and patch_names == set(planned),
            "candidate source hash patch paths differ from plan")
    files = value.get("files")
    require(isinstance(files, dict) and set(files) == set(planned),
            "candidate source hash file inventory differs")
    candidate_sources = HERE / "candidate-sources"
    adaptation_path = HERE / "format-adaptation.json"
    adaptation = read_json(adaptation_path, "format-adaptation.json") \
        if adaptation_path.is_file() and not adaptation_path.is_symlink() else {}
    adaptation_changes = adaptation.get("changes", {}) \
        if isinstance(adaptation, dict) else {}
    for name, item in files.items():
        require(isinstance(item, dict)
                and item.get("baseline_sha256") == baseline.get(name)
                and valid_digest(item.get("candidate_sha256")),
                f"candidate source hash entry is malformed: {name}")
        candidate_copy = candidate_sources / name
        require(candidate_copy.is_file() and not candidate_copy.is_symlink(),
                f"candidate source review copy is missing: {name}")
        candidate_copy_sha = sha(candidate_copy)
        if candidate_copy_sha != item["candidate_sha256"]:
            change = adaptation_changes.get(name) if isinstance(adaptation_changes, dict) else None
            require(isinstance(change, dict)
                    and change.get("prepared_sha256") == candidate_copy_sha
                    and change.get("formatted_sha256") == item["candidate_sha256"],
                    f"candidate source review copy hash differs: {name}")
        if name in baseline:
            baseline_copy = HERE / "baseline/sources" / name
            require(baseline_copy.is_file()
                    and sha(baseline_copy) == item["baseline_sha256"],
                    f"candidate source baseline hash differs: {name}")
        else:
            require(item["baseline_sha256"] is None,
                    f"candidate source new-file baseline hash differs: {name}")
    changed = {name for name, item in files.items()
               if item.get("baseline_sha256") != item.get("candidate_sha256")}
    require(changed == set(planned),
            "candidate source hash inventory contains an unchanged planned file")
    if (HERE / "candidate/source-manifest.json").is_file():
        candidate_manifest = source_manifest(HERE / "candidate/source-manifest.json")
        require(all(candidate_manifest.get(name) == item["candidate_sha256"]
                    for name, item in files.items()),
                "candidate source hash inventory differs from candidate manifest")
    return {"status": "pass", "schema": value["schema"],
            "source_hashes": relative(path), "source_hashes_sha256": sha(path),
            "candidate_patch_sha256": candidate_patch["sha256"],
            "files": len(files)}


def adr_manifest() -> dict[str, str]:
    path = next((candidate for candidate in (
        ADR, HERE.parent / "adr-manifest.json",
    ) if candidate.is_file() and not candidate.is_symlink()), None)
    require(path is not None, "adr-manifest.json is missing")
    value = read_json(path, display_path(path))
    require(isinstance(value, dict) and isinstance(value.get("files"), dict),
            "ADR manifest is malformed")
    expected = {
        item.decode("utf-8") for item in git_output(["git", "ls-files", "-z", "--", "docs/adr"])
        .split(b"\0") if item
    }
    files = value["files"]
    require(set(files) == expected, "ADR manifest inventory differs")
    for name, digest in files.items():
        safe_relative(name, "ADR path")
        require(name.startswith(ADR_ROOT) and valid_digest(digest),
                f"ADR entry is malformed: {name}")
        require((REPO / name).is_file() and not (REPO / name).is_symlink()
                and sha(REPO / name) == digest, f"ADR hash differs: {name}")
    return files


def validate_protocol() -> None:
    text = read_text(PROTOCOL, "protocol.md")
    lowered = text.lower()
    require("ole2" in lowered and "ooxml" in lowered and "odf" in lowered,
            "protocol does not preserve the OLE2/OOXML priority and ODF deferral")
    has_abba_label = "abba" in lowered
    has_abba_sequence = all(token in lowered for token in ("a1", "b1", "b2", "a2"))
    require((has_abba_label or has_abba_sequence)
            and "allocation" in lowered
            and ("cleanup" in lowered or "custody" in lowered
                 or "retained" in lowered),
            "protocol omits required custody controls")


def patch_paths(data: bytes, label: str) -> set[str]:
    try:
        text = data.decode("utf-8")
    except UnicodeDecodeError as error:
        raise EvidenceError(f"{label} is not UTF-8") from error

    def header_path(field: str) -> str | None:
        # ``git diff`` uses a//dev/null or b//dev/null for one side of a
        # new/deleted file, while plain unified diffs use /dev/null.  The
        # timestamp portion of a unified header is separated by a tab.
        field = field.split("\t", 1)[0]
        if field in ("/dev/null", "a//dev/null", "b//dev/null"):
            return None
        if field.startswith(("a/", "b/")):
            field = field[2:]
        if field == "/dev/null":
            return None
        safe_relative(field, f"{label} path")
        return field

    result: set[str] = set()
    unified_left: str | None = None
    unified_pending = False
    for line in text.splitlines():
        if line.startswith("diff --git "):
            fields = line.split()
            require(len(fields) >= 4, f"{label} has a malformed diff header")
            left = header_path(fields[2])
            right = header_path(fields[3])
            require(left == right or left is None or right is None,
                    f"{label} diff header paths differ")
            if left is not None:
                result.add(left)
        elif line.startswith("--- ") or line.startswith("+++ "):
            # Candidate review patches are sometimes retained as a plain
            # unified diff without git's ``diff --git`` headers.
            if line.startswith("--- "):
                require(not unified_pending, f"{label} has an unmatched unified diff header")
                unified_left = header_path(line[4:])
                unified_pending = True
                continue
            require(unified_pending, f"{label} has an unmatched unified diff header")
            right = header_path(line[4:])
            require(unified_left == right or unified_left is None or right is None,
                    f"{label} unified diff paths differ")
            if right is not None:
                result.add(right)
            elif unified_left is not None:
                result.add(unified_left)
            unified_left = None
            unified_pending = False
        else:
            continue
    require(not unified_pending, f"{label} has an unmatched unified diff header")
    return result


def git_show(revision: str, name: str) -> bytes:
    return git_output(["git", "show", f"{revision}:{name}"])


def replay_candidate_patch(stage: Path, baseline: dict[str, str], candidate: dict[str, str],
                           changed: set[str], label: str) -> None:
    """Replay the definitive baseline-to-candidate patch in an isolated tree."""
    patch_path = Path(label)
    if not patch_path.is_absolute():
        patch_path = HERE / patch_path
    data = read_bytes(patch_path, label)
    paths = patch_paths(data, label)
    require(paths == changed, f"{label} paths differ from candidate source diff")
    roots = plan_data()["_roots"]
    require(all(path_in_roots(name, roots) for name in paths),
            f"{label} contains a path outside candidate roots")
    with tempfile.TemporaryDirectory(prefix="litchi-0546-verify-candidate-", dir="/dev/shm") as temporary:
        root = Path(temporary)
        for name in changed & set(baseline):
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes((HERE / "baseline/sources" / name).read_bytes())
        checked = subprocess.run(
            ["git", "apply", "--check", "--whitespace=nowarn", "-"], cwd=root,
            input=data, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(checked.returncode == 0,
                f"{label} does not apply: {checked.stderr.decode(errors='replace')[-1200:]}")
        applied = subprocess.run(
            ["git", "apply", "--whitespace=nowarn", "-"], cwd=root,
            input=data, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(applied.returncode == 0,
                f"{label} could not be replayed: {applied.stderr.decode(errors='replace')[-1200:]}")
        for name in changed:
            produced = root / name
            retained = stage / "sources" / name
            require(produced.is_file() and retained.is_file()
                    and produced.read_bytes() == retained.read_bytes(),
                    f"{label} output differs for {name}")


def measured_candidate_patches() -> tuple[Path, Path]:
    """Locate the reviewed and applied patch custody inputs for 0546.

    The integrated runner keeps its authoritative source patch in the stage
    directory.  The reviewed patch is copied into ``candidate-sources`` by
    the review freeze; older bundles placed the same two files at the bundle
    root.  Accept these equivalent layouts while binding whichever bytes are
    actually retained.
    """
    def choose(candidates: list[Path]) -> Path | None:
        existing = [path for path in candidates
                    if path.is_file() and not path.is_symlink()]
        if not existing:
            return None
        digests = {sha(path) for path in existing}
        if len(digests) != 1:
            adaptation_path = HERE / "format-adaptation.json"
            require(adaptation_path.is_file() and not adaptation_path.is_symlink(),
                    f"duplicate candidate patch custody differs: "
                    f"{[display_path(path) for path in existing]}")
            adaptation = read_json(adaptation_path, "format-adaptation.json")
            prepared = adaptation.get("prepared_patch_sha256") if isinstance(adaptation, dict) else None
            applied = adaptation.get("applied_patch_sha256") if isinstance(adaptation, dict) else None
            require(valid_digest(prepared) and valid_digest(applied)
                    and prepared in digests and applied in digests,
                    "format adaptation does not bind differing candidate patches")
        return existing[0]

    reviewed = choose([
        HERE / "candidate.patch",
        HERE / "candidate-sources/candidate.patch",
    ])
    applied = choose([
        HERE / "applied-candidate.patch",
        HERE / "candidate-sources/applied-candidate.patch",
    ])
    if reviewed is None:
        reviewed = HERE / "candidate/source.patch"
    if applied is None:
        applied = reviewed
    return reviewed, applied


def _canonical_json(value: Any, label: str) -> str:
    """Return a stable key for multiset custody comparisons."""
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True,
                          separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        raise EvidenceError(f"{label} is not canonical JSON") from error


def _multiset(values: list[Any], label: str) -> dict[str, int]:
    result: dict[str, int] = {}
    for index, value in enumerate(values):
        key = _canonical_json(value, f"{label}[{index}]")
        result[key] = result.get(key, 0) + 1
    return result


def validate_next_candidate(candidate: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    """Custody-check the unmeasured follow-up proposal, when retained.

    The follow-up is deliberately optional: an early verifier run can finish
    the measured pilot before the next proposal is written.  A priority note
    may stand alone as design-only evidence; a supplied patch additionally
    receives isolated provenance replay.  No timing or admission result is
    inferred from either form.
    """
    patch_path = HERE / "next-candidate.patch"
    note_path = HERE / "next-priority.md"
    if not patch_path.exists() and not note_path.exists():
        return {"status": "not-present", "admission": "unmeasured-only"}
    require(note_path.is_file() and not note_path.is_symlink(),
            "next-priority.md is missing or not a regular file")
    note = read_text(note_path, "next-priority.md").lower()
    require("proposal" in note,
            "next-priority.md lacks an explicit proposal description")
    require("0546" in note and "ooxml" in note and "odf" in note,
            "next-priority.md omits pilot or format priority scope")
    require("no implementation or performance admission" in note
            or "no performance admission" in note,
            "next-priority.md does not disclaim implementation or performance admission")

    # A priority note is valid design-only evidence by itself.  A patch is a
    # separate, optional unmeasured proposal and is replayed only when the
    # author actually supplies one; the note must never manufacture a patch
    # requirement for this rejected campaign.
    if not patch_path.exists():
        design_path = HERE / "next-design-review.md"
        design_result: dict[str, Any] = {}
        if design_path.exists():
            require(design_path.is_file() and not design_path.is_symlink(),
                    "next-design-review.md is not a regular file")
            design = read_text(design_path, "next-design-review.md").lower()
            require("design-only" in design and "no patch" in design
                    and "no code should be applied" in design,
                    "next-design-review.md does not remain an unimplemented proposal")
            design_result = {
                "design": relative(design_path),
                "design_sha256": sha(design_path),
            }
        return {
            "status": "pass", "admission": "unmeasured-design-only",
            "label": "unmeasured design-only proposal",
            "note": relative(note_path), "note_sha256": sha(note_path),
            **design_result,
        }

    require(patch_path.is_file() and not patch_path.is_symlink(),
            "next-candidate.patch is not a regular file")
    patch = read_bytes(patch_path, "next-candidate.patch")
    paths = patch_paths(patch, "next-candidate.patch")
    require(paths, "next-candidate.patch is empty")
    roots = plan["_roots"]
    require(all(path_in_roots(name, roots) for name in paths),
            "next-candidate.patch contains a path outside candidate roots")
    candidate_paths = set(candidate["manifest"])
    require(paths <= candidate_paths,
            "next-candidate.patch is not based on frozen candidate source files")
    applied_path = HERE / "applied-candidate.patch"
    require(sha(patch_path) != sha(applied_path),
            "next-candidate.patch repeats the measured applied candidate patch")
    # Replay only against the retained candidate bytes.  The produced bytes
    # are intentionally not published as a stage: this checks context and
    # path provenance for an unmeasured proposal without making it evidence.
    with tempfile.TemporaryDirectory(prefix="litchi-0546-verify-next-", dir="/dev/shm") as temporary:
        root = Path(temporary)
        for name in paths:
            source = HERE / "candidate/sources" / name
            require(source.is_file() and not source.is_symlink(),
                    f"candidate source for next proposal is missing: {name}")
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(source.read_bytes())
        checked = subprocess.run(
            ["git", "apply", "--check", "--whitespace=nowarn", "-"], cwd=root,
            input=patch, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(checked.returncode == 0,
                "next-candidate.patch does not apply to frozen candidate sources: "
                f"{checked.stderr.decode(errors='replace')[-1200:]}")
        applied = subprocess.run(
            ["git", "apply", "--whitespace=nowarn", "-"], cwd=root,
            input=patch, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(applied.returncode == 0,
                "next-candidate.patch could not be replayed: "
                f"{applied.stderr.decode(errors='replace')[-1200:]}")
    # The patch branch has already validated the common note above.
    return {
        "status": "pass", "admission": "unmeasured-only",
        "patch": relative(patch_path), "patch_sha256": sha(patch_path),
        "patch_paths": sorted(paths), "label": "unmeasured proposal",
        "provenance": "replayed against frozen candidate sources",
        "note": relative(note_path), "note_sha256": sha(note_path),
    }


def validate_patch(stage: Path, snapshot: dict[str, str], revision: dict[str, str],
                   allowed: set[str], revision_name: str) -> set[str]:
    patch_path = need(stage / "source.patch", f"{relative(stage)}/source.patch")
    data = read_bytes(patch_path)
    changed = changed_paths(snapshot, revision)
    require(changed <= allowed,
            f"{relative(stage)} changes outside its permitted source scope: "
            f"{sorted(changed - allowed)}")
    paths = patch_paths(data, relative(patch_path))
    tracked_changed = changed & set(revision)
    require(paths == tracked_changed,
            f"{relative(stage)} source patch paths differ: "
            f"{sorted(paths ^ tracked_changed)}")
    if not changed:
        require(not data, f"{relative(stage)} has a patch without source changes")
        return changed
    if changed:
        # Replay in an isolated tmpfs directory.  This proves the retained
        # source snapshot corresponds to the patch without mutating the repo.
        with tempfile.TemporaryDirectory(prefix="litchi-0546-verify-", dir="/dev/shm") as temporary:
            root = Path(temporary)
            for name in tracked_changed:
                target = root / name
                target.parent.mkdir(parents=True, exist_ok=True)
                target.write_bytes(git_show(revision_name, name))
            if data:
                checked = subprocess.run(
                    ["git", "apply", "--check", "--whitespace=nowarn", "-"],
                    cwd=root, input=data, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                    check=False,
                )
                require(checked.returncode == 0,
                        f"{relative(patch_path)} does not apply: "
                        f"{checked.stderr.decode(errors='replace')[-1200:]}")
                applied = subprocess.run(
                    ["git", "apply", "--whitespace=nowarn", "-"], cwd=root,
                    input=data, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
                )
                require(applied.returncode == 0,
                        f"{relative(patch_path)} could not be replayed: "
                        f"{applied.stderr.decode(errors='replace')[-1200:]}")
            retained = stage / "sources"
            # Untracked harness/candidate additions are deliberately included
            # in the manifest and source copies, but a plain ``git diff``
            # cannot encode them.  Their byte custody is checked below; only
            # paths represented by the replayed patch are compared here.
            for name in tracked_changed:
                produced = root / name
                expected = retained / name
                if name in snapshot:
                    require(produced.is_file() and expected.is_file()
                            and produced.read_bytes() == expected.read_bytes(),
                            f"{relative(stage)} patch output differs for {name}")
                else:
                    require(not produced.exists() and not expected.exists(),
                            f"{relative(stage)} deletion custody differs for {name}")
    return changed


def harness_paths(plan: dict[str, Any]) -> set[str]:
    configured = plan.get("harness_source_paths")
    if configured is None:
        return set(EXPECTED_HARNESS_PATHS)
    require(isinstance(configured, list) and configured,
            "plan.harness_source_paths is not a nonempty list")
    paths = set()
    for name in configured:
        safe_relative(name, "harness source path")
        require(name.startswith("tools/perf-baseline/"),
                f"harness source path is outside perf harness: {name}")
        paths.add(name)
    require(paths == EXPECTED_HARNESS_PATHS,
            "harness source paths differ from the standalone guard enabler")
    return paths


def stage_names() -> list[str]:
    result: list[str] = []
    for path in sorted(HERE.iterdir(), key=lambda item: item.name):
        if not path.is_dir() or path.is_symlink() or not (path / "source-manifest.json").exists():
            continue
        safe_relative(path.name, "stage name")
        require("/" not in path.name, "stage name is nested")
        if path.name in {"baseline", "candidate"}:
            result.append(path.name)
        elif not re.fullmatch(r"failed-(baseline|candidate)-[1-9][0-9]*", path.name):
            raise EvidenceError(f"unrecognized source-attempt directory: {path.name}")
    if "baseline" not in result:
        raise Pending("baseline source attempt has not been frozen")
    if "candidate" not in result:
        raise Pending("candidate source attempt has not been frozen")
    return ["baseline", "candidate"]


def failed_stage_names() -> list[str]:
    result: list[str] = []
    for path in sorted(HERE.iterdir(), key=lambda item: item.name):
        if not path.is_dir() or path.is_symlink() or not (path / "source-manifest.json").exists():
            continue
        if path.name.startswith("failed-"):
            require(re.fullmatch(r"failed-(baseline|candidate)-[1-9][0-9]*", path.name),
                    f"invalid failed stage name: {path.name}")
            result.append(path.name)
    return result


def validate_stage_sources(stage_name: str, plan: dict[str, Any], revision: dict[str, str],
                            roots: list[str], harness: set[str], baseline: dict[str, str] | None,
                            *, final_candidate: bool = False) -> dict[str, Any]:
    stage = HERE / stage_name
    require(stage.is_dir() and not stage.is_symlink(),
            f"{stage_name} is not a regular stage directory")
    snapshot = source_manifest(stage / "source-manifest.json")
    if stage_name == "baseline":
        allowed = set(harness) | {CAP_ENABLER} | shared_source_paths(snapshot, revision, plan)
        changed = validate_patch(stage, snapshot, revision, allowed, plan["revision"])
    else:
        require(baseline is not None, "candidate validation has no baseline manifest")
        changed_from_baseline = changed_paths(snapshot, baseline)
        require(changed_from_baseline,
                f"{stage_name} candidate source diff is empty")
        planned = plan.get("candidate_files")
        require(isinstance(planned, list) and planned
                and changed_from_baseline == set(planned),
                "candidate source diff does not equal the frozen five-file inventory")
        require(all(path_in_roots(name, roots) for name in changed_from_baseline),
                f"{stage_name} changes outside candidate source roots: "
                f"{sorted(name for name in changed_from_baseline if not path_in_roots(name, roots))}")
        # A candidate snapshot still includes the common guard-enabler edits;
        # only its production diff is permitted against the baseline.
        allowed = set(harness) | {
            name for name in set(snapshot) | set(revision) if path_in_roots(name, roots)
        }
        changed = validate_patch(stage, snapshot, revision, allowed, plan["revision"])
        diff_path = need(stage / "source-diff.json", f"{stage_name}/source-diff.json")
        diff = read_json(diff_path)
        require(isinstance(diff, dict), f"{stage_name}/source-diff.json is not an object")
        require(diff.get("baseline_manifest_sha256") == sha(HERE / "baseline/source-manifest.json")
                and diff.get("candidate_manifest_sha256") == sha(stage / "source-manifest.json"),
                f"{stage_name} source-diff manifest binding differs")
        require(diff.get("candidate_source_roots") == roots,
                f"{stage_name} source-diff roots differ")
        changes = diff.get("changed_files")
        require(isinstance(changes, dict) and set(changes) == changed_from_baseline,
                f"{stage_name} source-diff changed file set differs")
        for name in changed_from_baseline:
            item = changes[name]
            require(isinstance(item, dict)
                    and item.get("baseline_sha256") == baseline.get(name)
                    and item.get("candidate_sha256") == snapshot.get(name),
                    f"{stage_name} source-diff hash differs for {name}")
        # The reviewed candidate patch is the definitive five-file patch for
        # this campaign.  The integrated runner may retain it under the
        # candidate-sources review freeze, while the stage source.patch is
        # the applied checkout spelling.
        draft, applied_patch = measured_candidate_patches()
        need(draft, display_path(draft))
        draft_paths = patch_paths(read_bytes(draft, display_path(draft)),
                                  display_path(draft))
        require(draft_paths == changed_from_baseline,
                "reviewed candidate patch paths differ from the final candidate source diff")
        require(all(path_in_roots(name, roots) for name in draft_paths),
                "reviewed candidate patch contains a path outside candidate roots")
        if "candidate_patch_sha256" in plan:
            require(sha(draft) == plan["candidate_patch_sha256"],
                    "reviewed candidate patch hash differs from plan")
        need(applied_patch, display_path(applied_patch))
        replay_candidate_patch(stage, baseline, snapshot, changed_from_baseline,
                               display_path(draft))
        replay_candidate_patch(stage, baseline, snapshot, changed_from_baseline,
                               display_path(applied_patch))
    sources = need(stage / "sources", f"{stage_name}/sources", directory=True)
    actual: set[str] = set()
    for path in sources.rglob("*"):
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            name = path.relative_to(sources).as_posix()
            safe_relative(name, f"{stage_name}/sources path")
            actual.add(name)
    expected = {name for name in snapshot
                if path_in_roots(name, roots) or name in harness
                or is_shared_test_path(name, plan)}
    require(actual == expected, f"{stage_name}/sources inventory differs")
    for name in actual:
        require(sha(sources / name) == snapshot[name],
                f"{stage_name}/sources hash differs: {name}")
    # The harness addition must be present in both source attempts, with the
    # same bytes.  This catches a source-boundary change hidden outside the
    # candidate roots.
    for name in harness:
        require(snapshot.get(name) is not None,
                f"{stage_name} omits bound harness source: {name}")
    return {
        "manifest": snapshot,
        "manifest_sha256": sha(stage / "source-manifest.json"),
        "changed_from_revision": sorted(changed_paths(snapshot, revision)),
        "changed_from_baseline": sorted(changed_paths(snapshot, baseline or snapshot)),
    }


def validate_failed_stage_sources(stage_name: str, plan: dict[str, Any],
                                  revision: dict[str, str], baseline: dict[str, str],
                                  harness: set[str], frozen_created: dt.datetime,
                                  corrected_created: dt.datetime) -> dict[str, Any]:
    """Audit a retained failed attempt without admitting it to measurement."""
    require(re.fullmatch(r"failed-baseline-[1-9][0-9]*", stage_name) is not None,
            f"unsupported failed stage: {stage_name}")
    stage = HERE / stage_name
    snapshot = source_manifest(stage / "source-manifest.json")
    expected = dict(baseline)
    require(CAP_ENABLER in expected,
            "baseline manifest omits the permanent cap enabler")
    expected[CAP_ENABLER] = sha(CAP_DIR / "guard.initial.rs")
    require(snapshot == expected,
            f"{stage_name} source snapshot differs from retained initial freeze")
    require(not read_bytes(stage / "source.patch", f"{stage_name}/source.patch"),
            f"{stage_name} unexpectedly contains a source patch")
    changed = changed_paths(snapshot, revision)
    shared_changed = shared_source_paths(snapshot, revision, plan) & changed
    require(changed == {CAP_ENABLER} | shared_changed,
            f"{stage_name} source difference is outside initial shared inputs")
    sources = need(stage / "sources", f"{stage_name}/sources", directory=True)
    actual: set[str] = set()
    for path in sources.rglob("*"):
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            name = path.relative_to(sources).as_posix()
            safe_relative(name, f"{stage_name}/sources path")
            actual.add(name)
    expected_sources = {
        name for name in snapshot
        if path_in_roots(name, plan["_roots"]) or name in harness
        or is_shared_test_path(name, plan)
    }
    require(actual == expected_sources,
            f"{stage_name}/sources inventory differs")
    for name in actual:
        require(sha(sources / name) == snapshot[name],
                f"{stage_name}/sources hash differs: {name}")
    receipt_paths = sorted(stage.glob("*.receipt.json"), key=lambda p: p.name)
    require(receipt_paths, f"{stage_name} has no failed receipt")
    require(len(receipt_paths) == 1,
            f"{stage_name} has unexpected failed receipt count")
    action = receipt_paths[0].name.removesuffix(".receipt.json")
    require(action == "quality-cap-clippy",
            f"{stage_name} failed action differs: {action}")
    manifest_data = {
        stage_name: {"manifest": snapshot,
                     "manifest_sha256": sha(stage / "source-manifest.json")}
    }
    record = validate_receipt(
        stage_name, action, read_json(receipt_paths[0], relative(receipt_paths[0])),
        plan, quality_commands(), {}, manifest_data,
    )
    require(record["exit_code"] != 0,
            f"{stage_name} retained receipt is not a failed attempt")
    require(frozen_created < record["start"] < record["end"] <= corrected_created,
            f"{stage_name} failed receipt is outside correction freeze window")
    bound = {receipt_paths[0].name} | record["artifacts"]
    allowed = {"source-manifest.json", "source.patch"}
    for path in stage.iterdir():
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            require(path.name in bound or path.name in allowed,
                    f"{relative(path)} is an unbound failed-attempt artifact")
    return {
        "status": "pass", "stage": stage_name,
        "manifest_sha256": sha(stage / "source-manifest.json"),
        "receipt_sha256": sha(receipt_paths[0]), "action": action,
        "exit_code": record["exit_code"],
        "start_utc": record["start"].isoformat(),
        "end_utc": record["end"].isoformat(),
    }


def validate_interrupted_baseline_attempt(baseline_manifest_sha256: str) -> dict[str, Any]:
    """Audit the preserved partial build without turning it into a receipt.

    The session that produced these logs no longer has a terminal handle.  Its
    two output files are useful custody evidence, but they carry no exit code,
    duration, or binary claim and therefore stay outside the measured receipt
    timeline.
    """
    folder = HERE / "interrupted-baseline-build-1"
    if not folder.exists():
        return {"status": "not-present"}
    require(folder.is_dir() and not folder.is_symlink(),
            "interrupted baseline attempt is not a regular directory")
    expected = {"build-normal.stdout", "build-normal.stderr", "interruption.json"}
    actual = {path.name for path in folder.iterdir()
              if path.is_file() and not path.is_symlink()}
    require(actual == expected,
            f"interrupted baseline attempt artifact inventory differs: {sorted(actual)}")
    value = read_json(folder / "interruption.json", "interrupted-baseline-build-1/interruption.json")
    require(isinstance(value, dict), "interrupted baseline interruption record is not an object")
    require(value.get("status") == "interrupted-no-terminal-receipt"
            and value.get("session_missing") is True
            and value.get("matching_build_processes") == [],
            "interrupted baseline status does not establish a missing terminal receipt")
    _positive_integer(value.get("session_id"),
                      "interrupted baseline session_id")
    parse_time(value.get("observed_utc"),
               "interrupted baseline observed_utc")
    require(value.get("action") and isinstance(value["action"], str)
            and "no exit status" in value["action"].lower()
            and "successful binary" in value["action"].lower(),
            "interrupted baseline action does not disclaim exit and binary claims")
    require(value.get("source_manifest_sha256") == baseline_manifest_sha256,
            "interrupted baseline source manifest differs")
    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == {
        "build-normal.stdout", "build-normal.stderr"},
            "interrupted baseline artifact hashes differ")
    for name, digest in artifacts.items():
        require(valid_digest(digest) and sha(folder / name) == digest,
                f"interrupted baseline artifact hash differs: {name}")
    require(not list(folder.glob("*.receipt.json")),
            "interrupted baseline attempt contains an invented receipt")
    resumed_name = value.get("resumed_build_receipt")
    resumed_digest = value.get("resumed_build_receipt_sha256")
    resumed_exit = value.get("resumed_build_exit_code")
    if resumed_name is not None or resumed_digest is not None or resumed_exit is not None:
        require(resumed_name == "baseline/build-normal.receipt.json"
                and valid_digest(resumed_digest) and resumed_exit == 0,
                "interrupted baseline resumed receipt binding is malformed")
        resumed_path = HERE / resumed_name
        require(resumed_path.is_file() and sha(resumed_path) == resumed_digest,
                "interrupted baseline resumed receipt hash differs")
    return {
        "status": "pass", "directory": relative(folder),
        "source_manifest_sha256": baseline_manifest_sha256,
        "artifacts": {name: artifacts[name] for name in sorted(artifacts)},
        "receipt": "none; terminal status intentionally unavailable",
        "resumed_receipt": resumed_name,
        "resumed_receipt_sha256": resumed_digest,
    }


def validate_harness_identity(stages: dict[str, dict[str, Any]], harness: set[str]) -> None:
    baseline = stages["baseline"]["manifest"]
    for name, data in stages.items():
        if name == "baseline":
            continue
        snapshot = data["manifest"]
        for path in set(baseline) | set(snapshot):
            if path in harness or not path.startswith("crates/"):
                require(baseline.get(path) == snapshot.get(path),
                        f"harness/common source differs between baseline and {name}: {path}")


def test_counts(stdout: str, label: str) -> dict[str, int]:
    rows = list(TEST_RESULT.finditer(stdout))
    if "test result:" not in stdout:
        return {"passed": 0, "failed": 0, "ignored": 0, "executed": 0, "result_lines": 0}
    require(rows, f"{label} has malformed Cargo test result output")
    passed = failed = ignored = 0
    for row in rows:
        passed += int(row.group(2))
        failed += int(row.group(3))
        ignored += int(row.group(4))
    return {
        "passed": passed, "failed": failed, "ignored": ignored,
        "executed": passed + failed, "result_lines": len(rows),
    }


def expected_main_jobs(plan: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    primary = plan["primary"]
    native: list[dict[str, Any]] = []
    for repeat in range(1, int(primary["repeats"]) + 1):
        for shape in primary["shapes"]:
            native.append({
                "name": f"native-r{repeat}-primary-{shape}", "kind": "primary",
                "guard": None, "repeat": repeat, "case": primary["case"],
                "shape": shape, "warmup": int(primary["warmup"]),
                "samples": int(primary["samples"]),
            })
    for repeat in range(1, int(plan["guard_repeats"]) + 1):
        for guard, item in enumerate(plan["guards"]):
            for shape in item["shapes"]:
                native.append({
                    "name": f"native-r{repeat}-guard{guard}-{shape}", "kind": "guard",
                    "guard": guard, "repeat": repeat, "case": item["case"],
                    "shape": shape, "warmup": int(plan["guard_warmup"]),
                    "samples": int(plan["guard_samples"]),
                })
    alloc = [
        {
            "name": f"alloc-r{repeat}-{shape}", "kind": "allocation", "guard": None,
            "repeat": repeat, "case": primary["case"], "shape": shape,
            "warmup": int(plan["allocation"]["warmup"]),
            "samples": int(plan["allocation"]["samples"]),
        }
        for repeat in range(1, int(plan["allocation"]["repeats"]) + 1)
        for shape in plan["allocation"]["shapes"]
    ]
    return native, alloc


def expected_guard_jobs(plan: dict[str, Any]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    config = plan["refusal_guard"]
    normal: list[dict[str, Any]] = []
    alloc: list[dict[str, Any]] = []
    for repeat in range(1, int(config["repeats"]) + 1):
        for shape in config["shapes"]:
            for case in config["cases"]:
                base = {
                    "repeat": repeat, "shape": shape, "case": case,
                    "warmup": int(config["native_warmup"]),
                    "samples": int(config["native_samples"]),
                }
                normal.append({"name": f"guard-native-r{repeat}-{shape}-{case}", **base})
                alloc_base = {**base, "warmup": int(config["allocation_warmup"]),
                              "samples": int(config["allocation_samples"])}
                alloc.append({"name": f"guard-alloc-r{repeat}-{shape}-{case}", **alloc_base})
    return normal, alloc


def expected_actions(plan: dict[str, Any], commands: dict[str, list[str]]) -> tuple[set[str], dict[str, dict[str, Any]]]:
    native, alloc = expected_main_jobs(plan)
    guard_native, guard_alloc = expected_guard_jobs(plan)
    actions = set(commands) | {
        "build-normal", "build-alloc", "build-guard-normal", "build-guard-alloc",
    } | {item["name"] for item in native + alloc + guard_native + guard_alloc}
    metadata: dict[str, dict[str, Any]] = {}
    for item in native + alloc:
        metadata[item["name"]] = {**item, "guard_binary": False,
                                   "allocator": item["kind"] == "allocation"}
    for item in guard_native + guard_alloc:
        metadata[item["name"]] = {**item, "guard_binary": True,
                                   "allocator": item["name"].startswith("guard-alloc-")}
    # The final quality lane is intentionally an action prefix: all frozen
    # quality commands must be present on the selected final source stage.
    actions |= {f"final-{name}" for name in commands}
    return actions, metadata


def artifact_map(stage: Path, action: str, receipt: dict[str, Any]) -> set[str]:
    entries = receipt.get("artifacts")
    require(isinstance(entries, dict), f"{relative(stage)}/{action} artifacts are missing")
    # Profile annotation files and eager binding records are written after
    # ``run.run`` has serialized its receipt.  They have their own analyzer
    # custody checks and therefore must not be mistaken for floating child
    # artifacts of the original command.
    def is_conditional_sidecar(name: str) -> bool:
        return (action.startswith("profile-")
                and name.endswith((".inclusive.txt", ".self.txt"))) \
            or (action.startswith("eager-") and name.endswith(".binding.json"))

    actual = {
        path.name for path in stage.glob(action + ".*")
        if path.is_file() and not path.is_symlink()
        and path.name != f"{action}.receipt.json"
        and not is_conditional_sidecar(path.name)
    }
    require(set(entries) == actual,
            f"{relative(stage)}/{action} artifact inventory differs")
    require({f"{action}.stdout", f"{action}.stderr"} <= actual,
            f"{relative(stage)}/{action} omits stdout or stderr")
    for name, digest in entries.items():
        safe_relative(name, f"{relative(stage)}/{action} artifact")
        require(Path(name).name == name and valid_digest(digest),
                f"{relative(stage)}/{action} artifact entry is malformed")
        target = stage / name
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"{relative(stage)}/{action} artifact hash differs: {name}")
    return actual


def build_command_ok(command: Any, binary: str, allocator: bool, *, guard: bool) -> bool:
    if not isinstance(command, list):
        return False
    required = ["cargo", "build", "--release", "--locked", "--manifest-path", "--bin"]
    if any(item not in command for item in required):
        return False
    try:
        if command[command.index("--bin") + 1] != binary:
            return False
    except (IndexError, ValueError):
        return False
    if "--target-dir" not in command:
        return False
    try:
        if command[command.index("--target-dir") + 1] != str(TARGET):
            return False
    except (IndexError, ValueError):
        return False
    if allocator:
        return "--features" in command and command[command.index("--features") + 1] == "allocator-metrics"
    return "--features" not in command


def main_capture_command_ok(command: Any, plan: dict[str, Any], job: dict[str, Any],
                            binary: dict[str, Any], stage: Path) -> bool:
    if not isinstance(command, list) or command[:3] != ["taskset", "-c", str(plan["cpu"])]:
        return False
    binary_path = binary["path"]
    if binary_path not in command:
        return False
    expected = {
        "--warmup": str(job["warmup"]), "--samples": str(job["samples"]),
        "--case": str(job["case"]), "--xlsx-cell-crud-shape": str(job["shape"]),
        "--json": str(stage / (job["name"] + ".json")),
    }
    try:
        for option, value in expected.items():
            if command.count(option) != 1 or command[command.index(option) + 1] != value:
                return False
    except IndexError:
        return False
    if job.get("allocator"):
        return "/usr/bin/time" not in command and "valgrind" not in command
    rss = str(stage / (job["name"] + ".rss.json"))
    return "/usr/bin/time" in command and "-o" in command \
        and command[command.index("-o") + 1] == rss


def guard_capture_command_ok(command: Any, plan: dict[str, Any], job: dict[str, Any],
                             binary: dict[str, Any], stage: Path) -> bool:
    expected = [
        "taskset", "-c", str(plan["cpu"]), binary["path"], "--shape", job["shape"],
        "--case", job["case"], "--warmup", str(job["warmup"]),
        "--samples", str(job["samples"]), "--json", str(stage / (job["name"] + ".json")),
    ]
    return command == expected


def binary_identity(stage: Path, label: str, plan: dict[str, Any], *, guard: bool,
                    allocator: bool) -> dict[str, Any] | None:
    identity_path = stage / (f"binary-guard-{label}.json" if guard else f"binary-{label}.json")
    if not identity_path.exists():
        return None
    identity = read_json(identity_path)
    require(isinstance(identity, dict), f"{relative(identity_path)} is not an object")
    expected_name = f"{stage.name}-guard-{label}" if guard else f"{stage.name}-{label}"
    expected_path = TARGET / "retained-binaries" / expected_name
    require(identity.get("path") == str(expected_path),
            f"{relative(identity_path)} binary path differs")
    require(valid_digest(identity.get("sha256")), f"{relative(identity_path)} digest is malformed")
    _positive_integer(identity.get("bytes"), f"{relative(identity_path)}.bytes")
    require(identity["bytes"] > 0, f"{relative(identity_path)} binary is empty")
    require(identity.get("source_manifest_sha256") == sha(stage / "source-manifest.json"),
            f"{relative(identity_path)} source manifest differs")
    build_name = "build-guard-" + label if guard else "build-" + label
    receipt_path = stage / f"{build_name}.receipt.json"
    receipt = read_json(receipt_path, f"{relative(receipt_path)}")
    require(identity.get("build_receipt_sha256") == sha(receipt_path),
            f"{relative(identity_path)} build receipt differs")
    require(receipt.get("exit_code") == 0 and receipt.get("binary_sha256") is None,
            f"{relative(receipt_path)} is not a successful build receipt")
    expected_binary = "xlsx_planning_guard" if guard else (
        "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline")
    require(build_command_ok(receipt.get("command"), expected_binary, allocator, guard=guard),
            f"{relative(receipt_path)} build command differs")
    binary_path = Path(identity["path"])
    if binary_path.exists():
        require(binary_path.is_file() and not binary_path.is_symlink(),
                f"{relative(identity_path)} retained binary is not regular")
        require(sha(binary_path) == identity["sha256"]
                and binary_path.stat().st_size == identity["bytes"],
                f"{relative(identity_path)} retained binary hash or size differs")
    else:
        # After cleanup, the retained hash must be present in the custody
        # record.  The record itself is checked by validate_cleanup below.
        require(CLEANUP.exists(), f"{relative(identity_path)} binary vanished before cleanup")
    return {
        "path": str(binary_path), "sha256": identity["sha256"], "bytes": identity["bytes"],
        "identity_path": identity_path, "build_name": build_name,
    }


def validate_receipt(stage_name: str, action: str, receipt: dict[str, Any],
                     plan: dict[str, Any], commands: dict[str, list[str]],
                     metadata: dict[str, dict[str, Any]],
                     manifests: dict[str, dict[str, Any]]) -> dict[str, Any]:
    stage = HERE / stage_name
    path = stage / f"{action}.receipt.json"
    require(isinstance(receipt, dict), f"{relative(path)} is not an object")
    start, end = interval(receipt, relative(path))
    exit_code = receipt.get("exit_code")
    require(isinstance(exit_code, int) and not isinstance(exit_code, bool),
            f"{relative(path)} exit code is malformed")
    stage_manifest_sha = manifests[stage_name]["manifest_sha256"]
    working_sha = receipt.get("working_source_manifest_sha256")
    require(receipt.get("source_manifest_sha256") == stage_manifest_sha
            and valid_digest(working_sha),
            f"{relative(path)} source manifest binding differs")
    retained_baseline_action = action.startswith((
        "profile-", "hardware-", "eager-", "build-eager-"
    ))
    if (stage_name == "baseline"
            and ("-r2-" in action or retained_baseline_action)
            and (HERE / "candidate/source-manifest.json").exists()):
        require(working_sha == manifests["candidate"]["manifest_sha256"],
                f"{relative(path)} retained baseline does not bind candidate source")
    else:
        require(working_sha == stage_manifest_sha,
                f"{relative(path)} working source differs from stage source")
    require(receipt.get("script_sha256") == sha(RUN)
            and receipt.get("plan_sha256") == sha(PLAN),
            f"{relative(path)} driver or plan binding differs")
    environment = receipt.get("environment")
    require(isinstance(environment, dict) and environment.get("TMPDIR") == str(TARGET / "test-tmp"),
            f"{relative(path)} temporary directory binding differs")
    if receipt.get("source_unchanged") is not None:
        require(receipt.get("source_unchanged") is True,
                f"{relative(path)} reports source mutation")
    artifacts = artifact_map(stage, action, receipt)

    command = receipt.get("command")
    expected_command: bool | None = None
    job = metadata.get(action)
    if action in commands:
        expected_command = command == commands[action]
    elif action.startswith("final-") and action[6:] in commands:
        expected_command = command == commands[action[6:]]
    elif action in {"build-normal", "build-alloc", "build-guard-normal", "build-guard-alloc"}:
        guard = action.startswith("build-guard-")
        allocator = action.endswith("alloc")
        binary = "xlsx_planning_guard" if guard else (
            "litchi-perf-baseline-alloc" if allocator else "litchi-perf-baseline")
        expected_command = build_command_ok(command, binary, allocator, guard=guard)
    elif job is not None:
        identity = binary_identity(
            stage, "alloc" if job.get("allocator") else "normal", plan,
            guard=bool(job.get("guard_binary")), allocator=bool(job.get("allocator")),
        )
        require(identity is not None,
                f"{relative(path)} has no bound binary identity")
        require(receipt.get("binary_sha256") == identity["sha256"],
                f"{relative(path)} binary digest differs")
        expected_command = (
            guard_capture_command_ok(command, plan, job, identity, stage)
            if job.get("guard_binary") else
            main_capture_command_ok(command, plan, job, identity, stage)
        )
    elif action.startswith("profile-") or action.startswith("hardware-") \
            or action.startswith("eager-") or action == "symbols" \
            or action.startswith("build-eager-"):
        expected_command = isinstance(command, list) and bool(command)
    else:
        raise EvidenceError(f"{relative(path)} action is not in the frozen action inventory")
    require(expected_command is True, f"{relative(path)} command differs from the frozen plan")

    if exit_code == 0:
        if action.startswith("native-"):
            required = {f"{action}.json", f"{action}.rss.json"}
            require(required <= artifacts, f"{relative(path)} omits native report or RSS")
        elif action.startswith("alloc-") or action.startswith("guard-"):
            require(f"{action}.json" in artifacts, f"{relative(path)} omits timed report")
        elif action.startswith("build-"):
            require(artifacts == {f"{action}.stdout", f"{action}.stderr"},
                    f"{relative(path)} build artifact set differs")
        elif action in commands or action.startswith("final-"):
            require(artifacts == {f"{action}.stdout", f"{action}.stderr"},
                    f"{relative(path)} quality artifact set differs")
    counts = None
    if action in {"quality-tests", "final-quality-tests"}:
        counts = test_counts(read_text(stage / f"{action}.stdout"), relative(stage / f"{action}.stdout"))
        if exit_code == 0:
            require(counts["failed"] == 0, f"{relative(path)} reports failed tests")
    return {
        "stage": stage_name, "name": action, "value": receipt,
        "start": start, "end": end, "exit_code": exit_code,
        "artifacts": artifacts, "counts": counts,
    }


def validate_stage_receipts(stage_name: str, plan: dict[str, Any], commands: dict[str, list[str]],
                            metadata: dict[str, dict[str, Any]],
                            manifests: dict[str, dict[str, Any]],
                            *, allow_inflight: bool = False) -> list[dict[str, Any]]:
    stage = HERE / stage_name
    receipts = sorted(stage.glob("*.receipt.json"), key=lambda item: item.name)
    if not receipts:
        raise Pending(f"{stage_name} has no retained command receipts")
    allowed, _ = expected_actions(plan, commands)
    records: list[dict[str, Any]] = []
    receipt_names: set[str] = set()
    for path in receipts:
        action = path.name.removesuffix(".receipt.json")
        receipt_names.add(action)
        if action not in allowed and not (
            action.startswith(("profile-", "hardware-", "eager-", "build-eager-"))
            or action == "symbols"
        ):
            raise EvidenceError(f"{relative(path)} action is not frozen")
        records.append(validate_receipt(stage_name, action, read_json(path), plan, commands,
                                        metadata, manifests))
    # No generated artifact may float beside a receipt without being listed in
    # that receipt.  Metadata and retained source copies have their own
    # explicit custody checks.
    bound = {f"{action}.receipt.json" for action in receipt_names}
    bound |= {name for record in records for name in record["artifacts"]}
    allowed_metadata = {"source-manifest.json", "source.patch", "source-diff.json"}
    allowed_metadata |= {
        path.name for path in stage.glob("binary-*.json")
        if path.is_file() and not path.is_symlink()
    }
    # Conditional analyzers retain deterministic sidecars beside their stage
    # receipts.  They are not child artifacts and therefore cannot appear in
    # the receipt's artifact map; the conditional lane validator binds their
    # contents to the corresponding report replay below.
    allowed_metadata |= {
        path.name for pattern in ("profile-*.inclusive.txt", "profile-*.self.txt",
                                  "eager-*.binding.json")
        for path in stage.glob(pattern)
        if path.is_file() and not path.is_symlink()
    }
    for path in stage.iterdir():
        if path.is_symlink():
            raise EvidenceError(f"{relative(path)} is a symlink")
        if path.is_file() and path.name not in bound and path.name not in allowed_metadata:
            conditional_prefix = path.name.startswith((
                "profile-", "hardware-", "eager-", "build-eager-", "symbols."
            ))
            if allow_inflight and (conditional_prefix or
                                   any(path.name.startswith(action + ".") for action in allowed)):
                raise Pending(f"{relative(path)} has no receipt yet; command is still being recorded")
            raise EvidenceError(f"{relative(path)} is an unbound stage artifact")
    return records


def validate_binary_set(stage_name: str, plan: dict[str, Any], records: list[dict[str, Any]]) -> dict[str, Any]:
    stage = HERE / stage_name
    result: dict[str, Any] = {}
    required = {
        "normal": "build-normal" in {record["name"] for record in records}
                   and any(record["name"].startswith("native-") for record in records),
        "alloc": "build-alloc" in {record["name"] for record in records}
                 and any(record["name"].startswith("alloc-") for record in records),
        "guard-normal": "build-guard-normal" in {record["name"] for record in records}
                        and any(record["name"].startswith("guard-native-") for record in records),
        "guard-alloc": "build-guard-alloc" in {record["name"] for record in records}
                       and any(record["name"].startswith("guard-alloc-") for record in records),
    }
    for label, needed in required.items():
        guard = label.startswith("guard-")
        kind = "alloc" if label.endswith("alloc") else "normal"
        identity = binary_identity(stage, kind, plan, guard=guard, allocator=kind == "alloc")
        if needed:
            require(identity is not None, f"{stage_name} is missing bound {label} binary identity")
        elif identity is not None:
            # A retained identity without its corresponding command attempt is
            # stale evidence and must not be used by an analyzer.
            require(f"build-{label}" in {record["name"] for record in records},
                    f"{stage_name} has an unbound {label} binary identity")
        if identity is not None:
            result[label] = identity
    return result


def validate_intervals(records: list[dict[str, Any]]) -> None:
    ordered = sorted(records, key=lambda record: record["start"])
    require(all(left["end"] <= right["start"]
                for left, right in zip(ordered, ordered[1:])),
            "command receipt intervals overlap")


def validate_abba(records_by_stage: dict[str, list[dict[str, Any]]], plan: dict[str, Any]) -> dict[str, Any]:
    """Check the planned baseline/candidate/baseline child order per lane."""
    native_jobs, allocation_jobs = expected_main_jobs(plan)
    guard_native_jobs, guard_allocation_jobs = expected_guard_jobs(plan)
    expected_by_lane_repeat: dict[str, dict[int, set[str]]] = {}
    for lane, jobs in (
        ("native", native_jobs), ("alloc", allocation_jobs),
        ("guard-native", guard_native_jobs), ("guard-alloc", guard_allocation_jobs),
    ):
        expected_by_lane_repeat[lane] = {}
        for job in jobs:
            expected_by_lane_repeat[lane].setdefault(job["repeat"], set()).add(job["name"])
    groups: dict[tuple[str, str, int], list[dict[str, Any]]] = {}
    for lane, prefix in (("native", "native-"), ("alloc", "alloc-"),
                         ("guard-native", "guard-native-"), ("guard-alloc", "guard-alloc-")):
        for stage_name in ("baseline", "candidate"):
            for record in records_by_stage.get(stage_name, []):
                action = record["name"]
                if not action.startswith(prefix):
                    continue
                match = re.search(r"-r([0-9]+)-", action)
                if match:
                    groups.setdefault((lane, stage_name, int(match.group(1))), []).append(record)
    order: list[tuple[str, str, int, dt.datetime]] = []
    for lane in ("native", "alloc", "guard-native", "guard-alloc"):
        for stage_name, repeat in (("baseline", 1), ("candidate", 1),
                                   ("candidate", 2), ("baseline", 2)):
            group = groups.get((lane, stage_name, repeat), [])
            expected_names = expected_by_lane_repeat[lane].get(repeat, set())
            actual_names = {item["name"] for item in group}
            if actual_names != expected_names:
                missing = sorted(expected_names - actual_names)
                extra = sorted(actual_names - expected_names)
                raise Pending(
                    f"ABBA {lane} {stage_name} repeat {repeat} is incomplete "
                    f"(missing={missing}, extra={extra})"
                )
            if any(item["exit_code"] != 0 for item in group):
                raise Pending(f"ABBA {lane} {stage_name} repeat {repeat} contains a failed capture")
            order.append((lane, stage_name, repeat, min(item["start"] for item in group)))
    for lane in ("native", "alloc", "guard-native", "guard-alloc"):
        values = [item for item in order if item[0] == lane]
        values.sort(key=lambda item: item[3])
        expected = [(lane, "baseline", 1), (lane, "candidate", 1),
                    (lane, "candidate", 2), (lane, "baseline", 2)]
        require([(item[0], item[1], item[2]) for item in values] == expected,
                f"{lane} capture order is not baseline/candidate/candidate/baseline")
    return {"status": "pass", "lanes": 4, "order": [
        {"lane": lane, "stage": stage, "repeat": repeat, "start_utc": start.isoformat()}
        for lane, stage, repeat, start in order
    ]}


def validate_guard_report(path: Path, job: dict[str, Any], binary: dict[str, Any],
                          allocator: bool) -> dict[str, Any]:
    raw = read_json(path, relative(path))
    label = relative(path)
    require(isinstance(raw, dict), f"{label} is not an object")
    require(raw.get("schema") == "litchi.xlsx.planning-refusal-guard.v1",
            f"{label} schema differs")
    require(raw.get("tool") == "xlsx_planning_guard"
            and raw.get("case") == job["case"] and raw.get("shape") == job["shape"],
            f"{label} tool/case/shape differs")
    require(raw.get("warmup_iterations") == job["warmup"]
            and raw.get("samples") == job["samples"], f"{label} iteration counts differ")
    source = raw.get("source")
    require(isinstance(source, dict), f"{label}.source is not an object")
    require(source.get("shape") == job["shape"] and source.get("worksheet_member") == "xl/worksheets/sheet1.xml"
            and source.get("compression") == "stored", f"{label} source identity differs")
    for key in ("rows", "columns", "archive_bytes", "worksheet_bytes"):
        _positive_integer(source.get(key), f"{label}.source.{key}")
    for key in ("source_sha256", "worksheet_sha256"):
        require(valid_digest(source.get(key)), f"{label}.source.{key} is malformed")
    require(raw.get("source_sha256") == source["source_sha256"],
            f"{label} source digest is not promoted consistently")
    logical = raw.get("logical_error")
    require(isinstance(logical, dict), f"{label}.logical_error is not an object")
    if job["case"] == "valid":
        require(logical.get("status") == "accepted"
                and logical.get("variant") is None and logical.get("message") is None,
                f"{label} valid outcome differs")
    else:
        require(logical.get("status") == "expected_failure"
                and logical.get("variant") == "Invalid"
                and isinstance(logical.get("message"), str) and logical["message"],
                f"{label} invalid outcome differs")
    phase = raw.get("phase")
    require(isinstance(phase, dict) and phase.get("name") == "edit_sheets",
            f"{label}.phase differs")
    durations = phase.get("duration_ns")
    order = phase.get("sample_order")
    metrics = phase.get("allocation_metrics")
    samples = phase.get("samples")
    require(isinstance(durations, list) and len(durations) == job["samples"],
            f"{label}.phase.duration_ns cardinality differs")
    require(isinstance(order, list) and sorted(order) == list(range(job["samples"])),
            f"{label}.phase.sample_order is not a permutation")
    require(isinstance(metrics, list) and len(metrics) == job["samples"]
            and isinstance(samples, list) and len(samples) == job["samples"],
            f"{label}.phase sample vectors differ")
    for index, duration in enumerate(durations):
        _positive_integer(duration, f"{label}.phase.duration_ns[{index}]", allow_zero=True)
        item = samples[index]
        require(isinstance(item, dict) and item.get("order") == index
                and item.get("duration_ns") == duration
                and item.get("allocation_metrics") == metrics[index],
                f"{label}.phase sample {index} is not serialized consistently")
        validate_guard_allocation(metrics[index], f"{label}.phase.allocation_metrics[{index}]", allocator)
    correctness = raw.get("correctness")
    require(isinstance(correctness, dict), f"{label}.correctness is not an object")
    require(correctness.get("source_unchanged") is True,
            f"{label} source/error correctness is incomplete")
    if job["case"] == "valid":
        expected = {
            "retry_preserved_error": False, "valid_snapshot_values": True,
            "empty_commit_is_noop": True, "expected_error_exact": False,
            "commit_outside_timing": True,
        }
    else:
        expected = {
            "retry_preserved_error": True, "valid_snapshot_values": False,
            "empty_commit_is_noop": False, "expected_error_exact": True,
            "commit_outside_timing": False,
        }
    for key, value in expected.items():
        require(correctness.get(key) is value, f"{label}.correctness.{key} differs")
    binary_identity = raw.get("binary")
    require(isinstance(binary_identity, dict)
            and binary_identity.get("sha256") == binary["sha256"]
            and binary_identity.get("profile", "release") == "release",
            f"{label}.binary identity differs")
    runner = raw.get("runner")
    require(isinstance(runner, dict) and runner.get("profile") == "release",
            f"{label}.runner identity differs")
    require(raw.get("allocation_scope") == "operation_global_system_allocator",
            f"{label} allocation scope differs")
    require(isinstance(raw.get("timing_scope"), str) and raw["timing_scope"]
            and isinstance(raw.get("performance_claim"), str)
            and raw["performance_claim"].startswith("none"),
            f"{label} timing/performance scope differs")
    return raw


def validate_guard_allocation(value: Any, label: str, allocator: bool) -> None:
    require(isinstance(value, dict), f"{label} is not an object")
    require(value.get("scope") == "operation_global_system_allocator",
            f"{label}.scope differs")
    if allocator:
        require(value.get("status") in {"measured", "overflow", "unavailable"},
                f"{label}.status is invalid")
        if value.get("status") == "measured":
            for key in ("allocation_calls", "deallocation_calls", "reallocation_calls",
                        "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
                        "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
                        "peak_live_bytes_after", "region_peak_live_bytes"):
                _positive_integer(value.get(key), f"{label}.{key}", allow_zero=True)
    else:
        require(value.get("status") == "unavailable"
                and set(value) == {"status", "scope"}, f"{label} normal allocation differs")


def validate_guard_reports(records_by_stage: dict[str, list[dict[str, Any]]],
                           plan: dict[str, Any], binaries: dict[str, dict[str, Any]]) -> dict[str, Any]:
    normal, alloc = expected_guard_jobs(plan)
    expected_by_name = {item["name"]: item for item in normal + alloc}
    identities: dict[str, dict[str, Any]] = {}
    count = 0
    for stage_name, records in records_by_stage.items():
        for record in records:
            action = record["name"]
            if not action.startswith("guard-") or record["exit_code"] != 0:
                continue
            job = expected_by_name[action]
            guard_key = "guard-alloc" if action.startswith("guard-alloc-") else "guard-normal"
            binary = binaries[stage_name][guard_key]
            raw = validate_guard_report(HERE / stage_name / f"{action}.json", job, binary,
                                        action.startswith("guard-alloc-"))
            identity = {
                "case": raw["case"], "shape": raw["shape"], "source": raw["source"],
                "logical_error": raw["logical_error"], "correctness": raw["correctness"],
            }
            key = action
            if key in identities:
                require(identities[key] == identity,
                        f"guard identity differs across stages for {key}")
            else:
                identities[key] = identity
            count += 1
    expected_count = (len(normal) + len(alloc)) * 2
    if count != expected_count:
        expected_names = {item["name"] for item in normal + alloc}
        successful_names = {
            record["name"] for rows in records_by_stage.values() for record in rows
            if record["name"].startswith("guard-") and record["exit_code"] == 0
        }
        return {
            "status": "pending", "reports": count, "expected": expected_count,
            "missing": sorted(expected_names - successful_names),
        }
    return {"status": "pass", "reports": count, "expected": expected_count}


def load_module(path: Path, name: str, *, register: bool = False) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise EvidenceError(f"cannot load analyzer: {display_path(path)}")
    module = importlib.util.module_from_spec(spec)
    if register:
        # analyze_allocation.py imports the native analyzer by this stable
        # name.  Register the exact retained module before executing the
        # allocation analyzer; no source files are imported from the checkout.
        sys.modules[name] = module
    try:
        spec.loader.exec_module(module)
    except Exception as error:  # noqa: BLE001 - preserve analyzer evidence context
        if register and sys.modules.get(name) is module:
            del sys.modules[name]
        raise EvidenceError(f"analyzer import failed for {display_path(path)}: {error}") from error
    return module


def _guard_analyzer_path() -> Path | None:
    for name in ("analyze_guard.py", "analyze_refusal_guard.py", "guard_analysis.py"):
        path = HERE / name
        if path.is_file() and not path.is_symlink():
            return path
    return None


def replay_analyzer(path: Path, report_path: Path, *, function: str = "analyze") -> dict[str, Any]:
    need(path, display_path(path))
    expected = read_json(report_path, display_path(report_path))
    module = load_module(
        path,
        "analyze" if path.name == "analyze.py" else
        "litchi_0546_" + path.stem.replace("-", "_"),
        register=path.name == "analyze.py",
    )
    method = getattr(module, function, None)
    require(callable(method), f"{display_path(path)} has no callable {function}()")
    try:
        result = method()
    except Exception as error:  # noqa: BLE001 - analyzer defines evidence errors
        raise EvidenceError(f"{display_path(path)} replay failed: {error}") from error
    require(_canonical_json(result, "analyzer replay") ==
            _canonical_json(expected, display_path(report_path)),
            f"{display_path(report_path)} differs from deterministic analyzer replay")
    summary: dict[str, Any] = {
        "status": "pass", "report": relative(report_path), "report_sha256": sha(report_path),
        "analyzer": relative(path), "analyzer_sha256": sha(path),
        "report_status": result.get("status") if isinstance(result, dict) else None,
    }
    # Preserve only the admission fields needed to bind an eventual decision;
    # the complete report remains hash-bound and is never rewritten here.
    if path.name == "analyze.py":
        summary["admission_status"] = result.get("admission_status")
        summary["native_admission"] = result.get("native_admission")
    elif path.name == "analyze_allocation.py":
        summary["planning_gate_passed"] = result.get("planning_gate_passed")
    elif path.name in {"analyze_guard.py", "analyze_refusal_guard.py", "guard_analysis.py"}:
        summary["admission_status"] = result.get("admission_status")
    return summary


def first_existing(paths: list[Path]) -> Path | None:
    existing = [path for path in paths if path.is_file() and not path.is_symlink()]
    require(len(existing) <= 1, f"duplicate analyzer reports: {[relative(path) for path in existing]}")
    return existing[0] if existing else None


def replay_analyses(plan: dict[str, Any], stages: dict[str, dict[str, Any]],
                    records_by_stage: dict[str, list[dict[str, Any]]],
                    binaries: dict[str, dict[str, Any]]) -> dict[str, Any]:
    analyzer = HERE / "analyze.py"
    allocation_analyzer = HERE / "analyze_allocation.py"
    guard_analyzer = _guard_analyzer_path()
    native_report = HERE / "comparison.json"
    alloc_report = HERE / "allocation-analysis.json"
    guard_report = first_existing([
        HERE / "guard-analysis.json", HERE / "guard-comparison.json",
        HERE / "refusal-guard-analysis.json",
    ])
    native_jobs, allocation_jobs = expected_main_jobs(plan)
    guard_native_jobs, guard_allocation_jobs = expected_guard_jobs(plan)

    def lane_ready(prefix: str, jobs: list[dict[str, Any]]) -> bool:
        expected = {job["name"] for job in jobs}
        for stage in ("baseline", "candidate"):
            rows = {record["name"]: record for record in records_by_stage.get(stage, [])
                    if record["name"].startswith(prefix)}
            if set(rows) != expected or any(record["exit_code"] != 0 for record in rows.values()):
                return False
        return True

    native_ready = lane_ready("native-", native_jobs)
    # Replaying a present report is always strict, even if a failed attempt
    # means the full comparison is not yet admissible.
    result: dict[str, Any] = {}
    if native_report.exists():
        result["native"] = replay_analyzer(analyzer, native_report)
    elif native_ready:
        raise Pending("comparison.json is missing after native captures")
    else:
        result["native"] = {"status": "pending", "reason": "native captures are incomplete"}

    allocation_ready = lane_ready("alloc-", allocation_jobs)
    if alloc_report.exists():
        # The allocation analyzer imports ``analyze`` for the shared schema
        # and numerical helper.  It must replay against this bundle's exact
        # native analyzer even when comparison.json is still pending.
        if "analyze" not in sys.modules:
            need(analyzer, display_path(analyzer))
            load_module(analyzer, "analyze", register=True)
        result["allocation"] = replay_analyzer(allocation_analyzer, alloc_report)
    elif allocation_ready:
        raise Pending("allocation-analysis.json is missing after allocation captures")
    else:
        result["allocation"] = {"status": "pending", "reason": "allocation captures are incomplete"}

    guard_ready = lane_ready("guard-native-", guard_native_jobs) \
        and lane_ready("guard-alloc-", guard_allocation_jobs)
    if guard_report is not None:
        require(guard_analyzer is not None, "guard analysis report has no bound analyzer")
        result["guard"] = replay_analyzer(guard_analyzer, guard_report)
    elif guard_ready:
        raise Pending("guard analysis report is missing after guard captures")
    else:
        result["guard"] = {"status": "pending", "reason": "guard captures are incomplete"}
    return result


def _conditional_lane_records(
    records_by_stage: dict[str, list[dict[str, Any]]], prefix: str,
) -> dict[str, dict[str, dict[str, Any]]]:
    """Index conditional receipts by stage and action without accepting extras."""
    result: dict[str, dict[str, dict[str, Any]]] = {}
    for stage_name in ("baseline", "candidate"):
        result[stage_name] = {
            record["name"]: record
            for record in records_by_stage.get(stage_name, [])
            if record["name"].startswith(prefix)
        }
    return result


def _conditional_profile_or_hardware_jobs(
    plan: dict[str, Any], lane: str,
) -> dict[str, dict[str, Any]]:
    require(lane in {"profile", "hardware"},
            f"unsupported conditional lane: {lane}")
    section = plan[lane]
    return {
        f"{lane}-r{repeat}-{shape}": {
            "name": f"{lane}-r{repeat}-{shape}",
            "repeat": repeat,
            "shape": shape,
            "case": plan["primary"]["case"],
            "warmup": int(section["warmup"]),
            "samples": int(section["samples"]),
        }
        for repeat in range(1, int(section["repeats"]) + 1)
        for shape in section["shapes"]
    }


def _conditional_expected_command(
    stage_name: str, lane: str, job: dict[str, Any],
    plan: dict[str, Any], binary: dict[str, Any],
    eager: dict[str, Any] | None = None,
) -> list[str]:
    stage = HERE / stage_name
    name = job["name"]
    if lane == "profile":
        owner = plan["profile"]["owner"]
        return [
            "taskset", "-c", str(plan["cpu"]), "valgrind", "--tool=callgrind",
            "--collect-atstart=no", "--toggle-collect=" + owner,
            "--zero-before=" + owner, "--dump-after=" + owner,
            "--callgrind-out-file=" + str(stage / (name + ".callgrind")),
            binary["path"], "--warmup", str(job["warmup"]),
            "--samples", str(job["samples"]), "--case", job["case"],
            "--xlsx-cell-crud-shape", job["shape"], "--json",
            str(stage / (name + ".json")),
        ]
    if lane == "hardware":
        return [
            "taskset", "-c", str(plan["cpu"]), "perf", "stat", "-x", ",",
            "-o", str(stage / (name + ".csv")), "-e", plan["hardware"]["events"],
            "--", binary["path"], "--warmup", str(job["warmup"]),
            "--samples", str(job["samples"]), "--case", job["case"],
            "--xlsx-cell-crud-shape", job["shape"], "--json",
            str(stage / (name + ".json")),
        ]
    require(lane == "eager" and eager is not None,
            f"unsupported conditional command lane: {lane}")
    return [
        "taskset", "-c", str(eager["cpu"]), "/usr/bin/time", "-f", EAGER_TIME_FORMAT,
        "-o", str(stage / (name + ".rss.json")), binary["path"],
        "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
        "--case", job["case"], "--xlsx-shape", job["shape"], "--json",
        str(stage / (name + ".json")),
    ]


def _validate_conditional_receipt_lane(
    lane: str, plan: dict[str, Any], records_by_stage: dict[str, list[dict[str, Any]]],
    binaries: dict[str, dict[str, Any]], eager: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Check exact conditional jobs, binary custody, commands, and ABBA order."""
    prefix = lane + "-"
    indexed = _conditional_lane_records(records_by_stage, prefix)
    expected_by_stage: dict[str, list[dict[str, Any]]] = {}
    if lane in {"profile", "hardware"}:
        jobs = _conditional_profile_or_hardware_jobs(plan, lane)
        expected_order = [jobs[f"{lane}-r{repeat}-{shape}"]
                          for repeat in (1, 2)
                          for shape in plan[lane]["shapes"]]
        for stage_name in ("baseline", "candidate"):
            expected_by_stage[stage_name] = expected_order
    else:
        require(isinstance(eager, dict), "eager plan is missing for eager receipts")
        all_jobs = eager.get("jobs")
        require(isinstance(all_jobs, list), "eager plan jobs are not a list")
        expected_by_stage = {
            stage_name: [job for job in all_jobs if job.get("stage") == stage_name]
            for stage_name in ("baseline", "candidate")
        }

    for stage_name in ("baseline", "candidate"):
        expected = expected_by_stage[stage_name]
        expected_names = {job["name"] for job in expected}
        actual_names = set(indexed[stage_name])
        require(actual_names == expected_names,
                f"{stage_name} {lane} receipt set differs: "
                f"missing={sorted(expected_names - actual_names)}, "
                f"extra={sorted(actual_names - expected_names)}")
        for job in expected:
            record = indexed[stage_name][job["name"]]
            require(record["exit_code"] == 0,
                    f"{stage_name}/{job['name']} conditional capture failed")
            receipt = record["value"]
            binary = binaries[stage_name]["normal"]
            require(receipt.get("binary_sha256") == binary["sha256"],
                    f"{stage_name}/{job['name']} binary binding differs")
            require(receipt.get("command") == _conditional_expected_command(
                stage_name, lane, job, plan, binary, eager),
                f"{stage_name}/{job['name']} command differs")

    # A conditional lane is four serial ABBA groups.  Within each group the
    # driver order is also frozen (profile/hardware shape order, eager case
    # then shape order), so a report cannot hide a reordered child sequence.
    groups: list[tuple[str, int, list[dict[str, Any]]]] = []
    for stage_name, repeat in (("baseline", 1), ("candidate", 1),
                               ("candidate", 2), ("baseline", 2)):
        group = [indexed[stage_name][job["name"]]
                 for job in expected_by_stage[stage_name]
                 if job["repeat"] == repeat]
        require(group, f"{stage_name} {lane} repeat {repeat} is empty")
        ordered = sorted(group, key=lambda record: record["start"])
        expected_names = [job["name"] for job in expected_by_stage[stage_name]
                          if job["repeat"] == repeat]
        require([record["name"] for record in ordered] == expected_names,
                f"{stage_name} {lane} repeat {repeat} child order differs")
        groups.append((stage_name, repeat, ordered))
    previous_end: dt.datetime | None = None
    for stage_name, repeat, group in groups:
        start = min(record["start"] for record in group)
        end = max(record["end"] for record in group)
        if previous_end is not None:
            require(previous_end <= start,
                    f"{lane} ABBA groups overlap at {stage_name} repeat {repeat}")
        previous_end = end
    return {
        "status": "pass",
        "lane": lane,
        "receipts": sum(len(value) for value in indexed.values()),
        "groups": [
            {"stage": stage_name, "repeat": repeat,
             "start_utc": min(record["start"] for record in group).isoformat(),
             "end_utc": max(record["end"] for record in group).isoformat()}
            for stage_name, repeat, group in groups
        ],
    }


def validate_conditional_frozen_inputs(
    candidate_frozen_created: dt.datetime,
) -> tuple[dict[str, Any], dt.datetime]:
    """Validate the exact envelope written immediately before conditional capture."""
    value = read_json(CONDITIONAL_FROZEN, "conditional-frozen-inputs.json")
    require(isinstance(value, dict)
            and set(value) == {"created_utc", "stage", "files"},
            "conditional frozen input envelope is malformed")
    created = parse_time(value.get("created_utc"),
                         "conditional-frozen-inputs.created_utc")
    require(created > candidate_frozen_created,
            "conditional frozen inputs do not follow candidate freeze")
    require(value.get("stage") == "before first conditional capture",
            "conditional frozen input stage differs")
    files = value.get("files")
    require(isinstance(files, dict)
            and set(files) == CONDITIONAL_FROZEN_FILES,
            "conditional frozen input file inventory differs")
    for name, digest in files.items():
        safe_relative(name, "conditional frozen input path")
        require(valid_digest(digest),
                f"conditional frozen input digest is malformed: {name}")
        target = HERE / name
        require(target.is_file() and not target.is_symlink(),
                f"conditional frozen input file is missing: {name}")
        require(sha(target) == digest,
                f"conditional frozen input hash differs: {name}")
    return value, created


def validate_symbol_observation(
    plan: dict[str, Any], binaries: dict[str, dict[str, Any]],
    candidate_frozen_created: dt.datetime, conditional_created: dt.datetime,
) -> dict[str, Any]:
    value = read_json(SYMBOL_OBSERVATION, "symbol-observation.json")
    require(isinstance(value, dict)
            and set(value) == {"created_utc", "owner", "plan_sha256", "stages"},
            "symbol observation envelope is malformed")
    created = parse_time(value.get("created_utc"), "symbol-observation.created_utc")
    require(candidate_frozen_created < created <= conditional_created,
            "symbol observation is outside conditional freeze window")
    owner = plan["profile"]["owner"]
    require(value.get("owner") == owner and value.get("plan_sha256") == sha(PLAN),
            "symbol observation plan or owner binding differs")
    stages = value.get("stages")
    require(isinstance(stages, dict) and set(stages) == {"baseline", "candidate"},
            "symbol observation stage inventory differs")
    expected_fields = {"command", "binary_sha256", "source_manifest_sha256", "matching_lines"}
    for stage_name in ("baseline", "candidate"):
        observation = stages[stage_name]
        require(isinstance(observation, dict) and set(observation) == expected_fields,
                f"symbol observation {stage_name} fields differ")
        binary = binaries[stage_name]["normal"]
        require(observation["command"] == [
            "nm", "-C", "--defined-only", binary["path"]
        ], f"symbol observation {stage_name} command differs")
        require(observation["binary_sha256"] == binary["sha256"],
                f"symbol observation {stage_name} binary differs")
        require(observation["source_manifest_sha256"] ==
                sha(HERE / stage_name / "source-manifest.json"),
                f"symbol observation {stage_name} source differs")
        lines = observation["matching_lines"]
        require(isinstance(lines, list) and lines
                and all(isinstance(line, str) and owner in line for line in lines),
                f"symbol observation {stage_name} has no exact owner match")
    return {"status": "pass", "sha256": sha(SYMBOL_OBSERVATION),
            "owner": owner, "created_utc": created.isoformat()}


def validate_eager_plan() -> dict[str, Any]:
    """Replay eager-run's frozen plan validator without capturing anything."""
    value = read_json(EAGER_PLAN, "eager-plan.json")
    module = load_module(EAGER_RUN, "litchi_0546_eager_run_plan")
    method = getattr(module, "validate_plan", None)
    require(callable(method), "eager_run.py has no callable validate_plan()")
    try:
        result = method(value, plan_data())
    except Exception as error:  # noqa: BLE001 - retain analyzer context
        raise EvidenceError(f"eager plan replay failed: {error}") from error
    require(result == value, "eager-plan.json differs from its frozen validator replay")
    return value


def replay_conditional_analyzer(
    analyzer_path: Path, report_path: Path, *, planning: bool = False,
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Replay one conditional analyzer and compare canonical JSON values."""
    need(analyzer_path, display_path(analyzer_path))
    expected = read_json(report_path, display_path(report_path))
    module = load_module(
        analyzer_path, "litchi_0546_" + analyzer_path.stem.replace("-", "_"),
    )
    method = getattr(module, "analyze", None)
    require(callable(method), f"{display_path(analyzer_path)} has no analyze()")
    try:
        result = method(False) if planning else method()
    except Exception as error:  # noqa: BLE001 - preserve analyzer evidence context
        raise EvidenceError(f"{display_path(analyzer_path)} replay failed: {error}") from error
    require(_canonical_json(result, "conditional analyzer replay") ==
            _canonical_json(expected, display_path(report_path)),
            f"{display_path(report_path)} differs from canonical analyzer replay")
    require(isinstance(result, dict),
            f"{display_path(report_path)} replay result is not an object")
    return result, {
        "status": "pass", "report": relative(report_path),
        "report_sha256": sha(report_path),
        "analyzer": relative(analyzer_path),
        "analyzer_sha256": sha(analyzer_path),
        "report_status": result.get("status"),
    }


def validate_profile_analysis(plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    result, summary = replay_conditional_analyzer(
        HERE / "analyze_planning.py", PROFILE_ANALYSIS, planning=True,
    )
    require(result.get("schema") == "xlsx_0546_paired_planning_profile_analysis_v1"
            and result.get("status") == "pass"
            and result.get("plan_sha256") == sha(PLAN),
            "planning profile analysis envelope differs")
    require(result.get("symbol_observation_sha256") == sha(SYMBOL_OBSERVATION),
            "planning profile analysis does not bind symbol observation")
    required = float(plan["gates"]["planning_ir_reduction_percent"])
    require(required >= 5.0, "planning Ir gate is below the required 5% floor")
    expected_keys = {
        (repeat, shape)
        for repeat in range(1, int(plan["profile"]["repeats"]) + 1)
        for shape in plan["profile"]["shapes"]
    }
    rows = result.get("rows")
    require(isinstance(rows, list) and all(isinstance(row, dict) for row in rows),
            "planning profile analysis rows are malformed")
    require({(row.get("repeat"), row.get("shape")) for row in rows} == expected_keys,
            "planning profile analysis row matrix differs")
    for row in rows:
        require(isinstance(row, dict) and set(row) == {
            "repeat", "shape", "baseline", "candidate", "planning_ir"
        }, "planning profile analysis row fields differ")
        ir = row["planning_ir"]
        required_value = ir.get("required_reduction_percent") if isinstance(ir, dict) else None
        reduction_value = ir.get("reduction_percent") if isinstance(ir, dict) else None
        require(isinstance(ir, dict)
                and isinstance(ir.get("passed"), bool)
                and ir["passed"] is True
                and isinstance(required_value, (int, float))
                and not isinstance(required_value, bool)
                and math.isfinite(float(required_value))
                and float(required_value) == required
                and isinstance(reduction_value, (int, float))
                and not isinstance(reduction_value, bool)
                and math.isfinite(float(reduction_value))
                and float(reduction_value) >= 5.0,
                f"planning Ir gate failed for {row.get('repeat')}/{row.get('shape')}")
        for field in ("baseline", "candidate"):
            value = row[field]
            require(isinstance(value, dict)
                    and isinstance(value.get("planning_ir"), int)
                    and value["planning_ir"] > 0,
                    f"planning profile {field} Ir is invalid")
    stage_rows = result.get("stages")
    require(isinstance(stage_rows, dict)
            and set(stage_rows) == {"baseline", "candidate"},
            "planning profile stage inventory differs")
    expected_count = len(expected_keys)
    for stage_name, values in stage_rows.items():
        require(isinstance(values, list) and len(values) == expected_count,
                f"planning profile {stage_name} row count differs")
    summary["ir_gate"] = {"required_reduction_percent": required,
                           "minimum_observed_reduction_percent": min(
                               float(row["planning_ir"]["reduction_percent"])
                               for row in rows),
                           "rows": len(rows)}
    return result, summary


def validate_eager_analysis() -> tuple[dict[str, Any], dict[str, Any]]:
    result, summary = replay_conditional_analyzer(
        HERE / "analyze_eager.py", EAGER_ANALYSIS,
    )
    require(result.get("schema") == "litchi-0546-eager-analysis-v1"
            and result.get("status") == "pass"
            and result.get("primary_plan_sha256") == sha(PLAN)
            and result.get("eager_plan_sha256") == sha(EAGER_PLAN)
            and result.get("no_speedup_requirement") is True,
            "eager analysis envelope differs")
    require(isinstance(result.get("capture_order"), list)
            and [item.get("group") for item in result["capture_order"]] == [
                "baseline-r1", "candidate-r1", "candidate-r2", "baseline-r2"
            ], "eager analysis ABBA order differs")
    stages = result.get("baseline"), result.get("candidate")
    require(all(isinstance(stage, dict)
                and isinstance(stage.get("custody"), dict)
                and stage["custody"].get("receipt_count") == 8
                for stage in stages),
            "eager analysis does not retain the complete 8-row stage controls")
    comparison = result.get("comparison")
    require(isinstance(comparison, dict)
            and isinstance(comparison.get("adverse_flags_over_five_percent"), list)
            and isinstance(comparison.get("same_build_drift_over_five_percent"), list),
            "eager analysis comparison flags are missing")
    summary["adverse_rows"] = len(comparison["adverse_flags_over_five_percent"])
    summary["drift_rows"] = len(comparison["same_build_drift_over_five_percent"])
    return result, summary


def validate_hardware_analysis() -> tuple[dict[str, Any], dict[str, Any]]:
    result, summary = replay_conditional_analyzer(
        HERE / "analyze_hardware.py", HARDWARE_ANALYSIS,
    )
    require(result.get("schema") ==
            "xlsx_shared_traversal_whole_child_hardware_analysis_v1"
            and result.get("status") == "pass"
            and result.get("plan_sha256") == sha(PLAN),
            "hardware analysis envelope differs")
    stages = result.get("stages")
    require(isinstance(stages, dict) and set(stages) == {"baseline", "candidate"}
            and all(isinstance(value, dict) and value.get("status") == "pass"
                    for value in stages.values()),
            "hardware analysis stage diagnostics are incomplete")
    comparison = result.get("comparison")
    require(isinstance(comparison, dict)
            and comparison.get("latency_gate") is False
            and comparison.get("operation_local_claim") is False
            and comparison.get("isolated_planning_counter_claim") is False,
            "hardware analysis makes an inadmissible performance claim")
    summary["diagnostic_only"] = True
    summary["captures"] = sum(stage.get("capture_count", 0)
                               for stage in stages.values())
    return result, summary


def validate_hardware_review(hardware_result: dict[str, Any] | None) -> dict[str, Any]:
    """Keep hardware review separate from the timing adverse-row groups.

    Hardware counters are diagnostic and have no admission gate.  A retained
    prose note is checked when supplied, while the analyzer report itself is
    sufficient custody evidence when no separate note was written.
    """
    if hardware_result is None:
        return {"status": "not-present"}
    note_path = HERE / "hardware-review.md"
    if not note_path.exists():
        return {"status": "diagnostic-only", "report_sha256": sha(HARDWARE_ANALYSIS)}
    note = read_text(note_path, "hardware-review.md").lower()
    require("0546" in note and "hardware" in note and "diagnostic" in note,
            "hardware-review.md does not identify diagnostic 0546 hardware review")
    require("latency" in note and ("no" in note or "not" in note),
            "hardware-review.md does not disclaim a latency gate")
    return {"status": "pass", "report_sha256": sha(HARDWARE_ANALYSIS),
            "note_sha256": sha(note_path)}


def validate_conditional_lanes(
    plan: dict[str, Any], records_by_stage: dict[str, list[dict[str, Any]]],
    binaries: dict[str, dict[str, Any]], candidate_frozen_created: dt.datetime,
    analyses: dict[str, Any],
) -> dict[str, Any]:
    """Validate optional conditionals only after the pilot gates authorize them."""
    prefixes = ("profile-", "hardware-", "eager-")
    stage_has_receipts = any(
        record["name"].startswith(prefix)
        for records in records_by_stage.values()
        for record in records
        for prefix in prefixes
    )
    # A rejection may retain planning/setup notes that deliberately stop
    # before any conditional capture.  Those notes are not conditional
    # evidence and must not turn an otherwise complete rejection into a
    # fabricated pending lane.  Enter the strict conditional path only when a
    # capture receipt or a completed analyzer report is present.
    evidence_present = stage_has_receipts or any(path.is_file() for path in (
        PROFILE_ANALYSIS, HARDWARE_ANALYSIS, EAGER_ANALYSIS,
    ))
    if not evidence_present:
        return {"status": "not-present", "analyses": {},
                "hardware_review": {"status": "not-present"}, "receipts": 0}

    require(all(analyses.get(name, {}).get("status") == "pass"
                for name in ("native", "allocation", "guard", "cap")),
            "conditional evidence exists before complete pilot analyzer replay")
    require(isinstance(analyses["native"].get("native_admission"), dict)
            and analyses["native"]["native_admission"].get("passed") is True
            and analyses["allocation"].get("planning_gate_passed") is True
            and analyses["guard"].get("admission_status") == "pass"
            and analyses["cap"].get("admission_passed") is True,
            "conditional evidence exists before pilot gates pass")
    if not CONDITIONAL_FROZEN.exists():
        raise Pending("conditional-frozen-inputs.json is missing")
    _, conditional_created = validate_conditional_frozen_inputs(candidate_frozen_created)
    symbol = validate_symbol_observation(
        plan, binaries, candidate_frozen_created, conditional_created,
    )
    eager_plan = validate_eager_plan()

    lane_paths = {
        "profile": (PROFILE_ANALYSIS, "profile-"),
        "eager": (EAGER_ANALYSIS, "eager-"),
        "hardware": (HARDWARE_ANALYSIS, "hardware-"),
    }
    lane_results: dict[str, Any] = {}
    lane_summaries: dict[str, Any] = {}
    for lane, (report_path, prefix) in lane_paths.items():
        report_exists = report_path.exists()
        records = [record for rows in records_by_stage.values()
                    for record in rows if record["name"].startswith(prefix)]
        if not report_exists and not records:
            if lane in {"profile", "eager"}:
                raise Pending(f"{report_path.name} is missing before retention")
            continue
        if not report_exists:
            raise Pending(f"{report_path.name} is missing after {lane} captures")
        if lane == "profile":
            result, summary = validate_profile_analysis(plan)
        elif lane == "eager":
            result, summary = validate_eager_analysis()
        else:
            result, summary = validate_hardware_analysis()
        _validate_conditional_receipt_lane(
            lane, plan, records_by_stage, binaries,
            eager_plan if lane == "eager" else None,
        )
        lane_results[lane] = result
        lane_summaries[lane] = summary

    require({"profile", "eager"} <= set(lane_results),
            "conditional profile/eager analyses are incomplete")
    hardware_review = validate_hardware_review(lane_results.get("hardware"))
    receipt_count = sum(
        1 for rows in records_by_stage.values() for record in rows
        if record["name"].startswith(prefixes)
    )
    return {
        "status": "pass", "created_utc": conditional_created.isoformat(),
        "symbol_observation": symbol, "eager_plan_sha256": sha(EAGER_PLAN),
        "analyses": lane_summaries, "hardware_review": hardware_review,
        "receipts": receipt_count,
        "_results": lane_results,
    }


def cap_plan_data(plan: dict[str, Any], candidate_changed: set[str]) -> dict[str, Any]:
    """Validate the supplemental cap plan before admitting its evidence."""
    for path, label in ((CAP_PLAN, "cap-boundary/plan.json"),
                        (CAP_RUN, "cap-boundary/cap_run.py"),
                        (CAP_ANALYZER, "cap-boundary/analyze.py")):
        need(path, label)
    value = read_json(CAP_PLAN, "cap-boundary/plan.json")
    require(isinstance(value, dict), "cap-boundary plan is not an object")
    created = parse_time(value.get("created_utc"), "cap-boundary.plan.created_utc")
    require(created >= plan["_created"], "cap-boundary plan predates main plan")
    require(isinstance(value.get("status"), str)
            and value["status"].startswith("frozen-before"),
            "cap-boundary plan is not frozen before build and capture")
    require(value.get("revision") == plan["revision"],
            "cap-boundary plan revision differs")
    priority = value.get("priority")
    require(isinstance(priority, str) and "ole2/ooxml" in priority.lower()
            and "odf" in priority.lower(),
            "cap-boundary plan priority omits OLE2/OOXML first and ODF deferral")
    require(value.get("owned_paths") == plan["owned_paths"],
            "cap-boundary owned target differs")
    require(value.get("cpu") == plan["cpu"], "cap-boundary CPU differs")
    require(value.get("native_order") == [
        "baseline repeat 1", "candidate repeat 1", "candidate repeat 2",
        "retained baseline repeat 2 under candidate source checkout",
    ], "cap-boundary native order differs")
    boundary = value.get("cap_boundary")
    require(isinstance(boundary, dict), "cap-boundary section is missing")
    require(boundary.get("sizes") == list(CAP_SIZES)
            and boundary.get("repeats") == len(CAP_REPEATS)
            and boundary.get("warmup") == 10
            and boundary.get("samples") == 100
            and boundary.get("case") == "valid"
            and boundary.get("binary") == CAP_EXAMPLE,
            "cap-boundary size, repeat, or child configuration differs")
    require(boundary.get("fixture_output") ==
            "stage/cap-native-r{repeat}-{size}.fixture.bin",
            "cap-boundary fixture output template differs")
    require(isinstance(boundary.get("timing_scope"), str)
            and boundary["timing_scope"].strip()
            and isinstance(boundary.get("performance_claim"), str)
            and "planning" in boundary["performance_claim"].lower(),
            "cap-boundary timing scope or claim is missing")
    expected_roots = [
        "crates/litchi-xlsx/src/cell_values/",
        "crates/litchi-xlsx/src/raw/worksheet/",
        "crates/litchi-xlsx/examples/",
    ]
    require(value.get("candidate_source_roots") == expected_roots,
            "cap-boundary candidate roots differ")
    require(value.get("candidate_files") == plan.get("candidate_files")
            and set(candidate_changed) == set(plan.get("candidate_files", [])),
            "cap-boundary candidate file inventory differs")
    require(value.get("shared_test_source_prefixes") ==
            plan.get("shared_test_source_prefixes"),
            "cap-boundary shared test prefix differs")
    enabler = value.get("temporary_enabler")
    require(isinstance(enabler, dict)
            and enabler.get("source") == CAP_ENABLER
            and enabler.get("source_root_included_for_snapshot") is True
            and enabler.get("retained_runtime_api") is False,
            "cap-boundary enabler binding differs")
    build = value.get("build")
    require(isinstance(build, dict)
            and build.get("package") == "litchi-xlsx"
            and build.get("example") == CAP_EXAMPLE
            and build.get("profile") == "release"
            and build.get("features") == "all"
            and build.get("jobs") == 2
            and build.get("incremental") is False
            and build.get("locked") is True
            and build.get("target_dir") == str(TARGET),
            "cap-boundary build envelope differs")
    guard = value.get("guard")
    require(isinstance(guard, dict)
            and guard.get("schema") == "litchi.xlsx.cap-boundary-guard.v1"
            and guard.get("tool") == CAP_EXAMPLE
            and guard.get("case") == "valid"
            and guard.get("allocator") == "not measured"
            and guard.get("profile") == "not measured",
            "cap-boundary guard envelope differs")
    admission = value.get("admission")
    require(isinstance(admission, str) and "1.05" in admission
            and "every size" in admission.lower()
            and "individual" in admission.lower(),
            "cap-boundary admission rule is missing")
    return {**value, "_created": created}


def cap_source_stage(stage_name: str, main_stage: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    """Bind each cap source snapshot to the already frozen main snapshot."""
    stage = CAP_DIR / stage_name
    require(stage.is_dir() and not stage.is_symlink(),
            f"cap-boundary/{stage_name} is not a regular stage directory")
    manifest_path = stage / "source-manifest.json"
    manifest = source_manifest(manifest_path)
    require(manifest == main_stage["manifest"],
            f"cap-boundary/{stage_name} manifest differs from main {stage_name}")
    source_patch = read_bytes(stage / "source.patch",
                              f"cap-boundary/{stage_name}/source.patch")
    main_patch = read_bytes(HERE / stage_name / "source.patch",
                            f"{stage_name}/source.patch")
    require(source_patch == main_patch,
            f"cap-boundary/{stage_name} source patch differs from main stage")
    roots = plan["_roots"]
    harness = harness_paths(plan)
    expected = {name for name in manifest
                if path_in_roots(name, roots) or name in harness
                or is_shared_test_path(name, plan)}
    sources = need(stage / "sources", f"cap-boundary/{stage_name}/sources",
                   directory=True)
    actual: set[str] = set()
    for path in sources.rglob("*"):
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            name = path.relative_to(sources).as_posix()
            safe_relative(name, f"cap-boundary/{stage_name}/sources path")
            actual.add(name)
    require(actual == expected,
            f"cap-boundary/{stage_name}/sources inventory differs")
    for name in actual:
        require(sha(sources / name) == manifest[name],
                f"cap-boundary/{stage_name}/sources hash differs: {name}")
    if stage_name == "candidate":
        main_diff = read_json(HERE / "candidate/source-diff.json",
                              "candidate/source-diff.json")
        cap_diff = read_json(stage / "source-diff.json",
                             "cap-boundary/candidate/source-diff.json")
        require(_canonical_json(cap_diff, "cap candidate source diff") ==
                _canonical_json(main_diff, "main candidate source diff"),
                "cap-boundary candidate source-diff differs from main")
    return {"manifest": manifest, "manifest_sha256": sha(manifest_path),
            "stage": stage_name}


def cap_build_command() -> list[str]:
    return [
        "env", "CARGO_BUILD_JOBS=2", "CARGO_INCREMENTAL=0", "cargo", "build",
        "--release", "--locked", "-p", "litchi-xlsx", "--all-features",
        "--example", CAP_EXAMPLE, "--target-dir", str(TARGET),
    ]


def cap_artifacts(stage: Path, action: str, receipt: dict[str, Any],
                  expected: set[str]) -> set[str]:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{relative(stage)}/{action} artifact inventory differs")
    for name, digest in artifacts.items():
        safe_relative(name, f"{relative(stage)}/{action} artifact")
        require(Path(name).name == name and valid_digest(digest),
                f"{relative(stage)}/{action} artifact entry is malformed")
        target = stage / name
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"{relative(stage)}/{action} artifact hash differs: {name}")
    return set(artifacts)


def cap_receipt(stage_name: str, action: str, cap_plan: dict[str, Any],
                manifest_sha256: str, working_sha256: str,
                *, binary_sha256: str | None = None) -> dict[str, Any]:
    stage = CAP_DIR / stage_name
    path = stage / f"{action}.receipt.json"
    receipt = read_json(path, relative(path))
    require(isinstance(receipt, dict), f"{relative(path)} is not an object")
    start, end = interval(receipt, relative(path))
    require(receipt.get("exit_code") == 0,
            f"{relative(path)} is not a successful cap receipt")
    require(receipt.get("source_manifest_sha256") == manifest_sha256
            and receipt.get("working_source_manifest_sha256") == working_sha256,
            f"{relative(path)} source binding differs")
    require(receipt.get("script_sha256") == sha(RUN)
            and receipt.get("plan_sha256") == sha(CAP_PLAN)
            and receipt.get("cap_driver_sha256") == sha(CAP_RUN),
            f"{relative(path)} driver or plan binding differs")
    environment = receipt.get("environment")
    require(isinstance(environment, dict)
            and environment.get("TMPDIR") == str(TARGET / "test-tmp"),
            f"{relative(path)} temporary directory binding differs")
    if binary_sha256 is not None:
        require(receipt.get("binary_sha256") == binary_sha256,
                f"{relative(path)} binary binding differs")
    return {"name": action, "stage": stage_name, "start": start, "end": end,
            "exit_code": receipt["exit_code"], "value": receipt,
            "receipt_sha256": sha(path)}


def cap_binary(stage_name: str, source: dict[str, Any], cap_plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    stage = CAP_DIR / stage_name
    identity_path = stage / "binary-cap-boundary.json"
    identity = read_json(identity_path, relative(identity_path))
    require(isinstance(identity, dict), f"{relative(identity_path)} is not an object")
    expected_path = CAP_BINARY_DIR / f"{stage_name}-cap-boundary"
    require(identity.get("path") == str(expected_path),
            f"{relative(identity_path)} path differs")
    binary_sha256 = identity.get("sha256")
    require(valid_digest(binary_sha256), f"{relative(identity_path)} digest is malformed")
    _positive_integer(identity.get("bytes"), f"{relative(identity_path)}.bytes")
    require(identity["bytes"] > 0, f"{relative(identity_path)} binary is empty")
    require(identity.get("source_manifest_sha256") == source["manifest_sha256"],
            f"{relative(identity_path)} source manifest differs")
    binary_path = Path(identity["path"])
    if binary_path.exists():
        require(binary_path.is_file() and not binary_path.is_symlink()
                and sha(binary_path) == binary_sha256
                and binary_path.stat().st_size == identity["bytes"],
                f"{relative(identity_path)} binary hash or size differs")
    else:
        require(CLEANUP.exists(),
                f"{relative(identity_path)} binary vanished before cleanup")
    build = cap_receipt(stage_name, "build-cap-boundary", cap_plan,
                        source["manifest_sha256"], source["manifest_sha256"])
    require(build["value"].get("binary_sha256") is None,
            f"{stage_name} cap build unexpectedly binds a binary")
    require(build["value"].get("command") == cap_build_command(),
            f"{stage_name} cap build command differs")
    cap_artifacts(stage, "build-cap-boundary", build["value"],
                  {"build-cap-boundary.stdout", "build-cap-boundary.stderr"})
    require(identity.get("build_receipt_sha256") == build["receipt_sha256"],
            f"{relative(identity_path)} build receipt differs")
    return ({"path": str(binary_path), "sha256": binary_sha256,
             "bytes": identity["bytes"],
             "source_manifest_sha256": source["manifest_sha256"],
             "build_receipt_sha256": build["receipt_sha256"]}, build)


def cap_column_name(column: int) -> str:
    value = column + 1
    result = ""
    while value:
        value, remainder = divmod(value - 1, 26)
        result = chr(ord("A") + remainder) + result
    return result


def cap_report(stage_name: str, repeat: int, size: int, binary: dict[str, Any],
               source: dict[str, Any], cap_plan: dict[str, Any]) -> dict[str, Any]:
    action = f"cap-native-r{repeat}-{size}"
    stage = CAP_DIR / stage_name
    report_path = stage / f"{action}.json"
    fixture_path = stage / f"{action}.fixture.bin"
    rss_path = stage / f"{action}.rss.json"
    stdout_path = stage / f"{action}.stdout"
    stderr_path = stage / f"{action}.stderr"
    for path in (report_path, fixture_path, rss_path, stdout_path, stderr_path):
        need(path, relative(path))
    receipt_record = cap_receipt(
        stage_name, action, cap_plan, source["manifest_sha256"],
        source["manifest_sha256"] if not (stage_name == "baseline" and repeat == 2)
        else sha(CAP_DIR / "candidate/source-manifest.json"),
        binary_sha256=binary["sha256"],
    )
    receipt = receipt_record["value"]
    expected_command = [
        "taskset", "-c", str(cap_plan["cpu"]), "/usr/bin/time", "-f", CAP_TIME_FORMAT,
        "-o", str(rss_path), binary["path"], "--size", str(size),
        "--warmup", "10", "--samples", "100", "--json", str(report_path),
        "--fixture-out", str(fixture_path),
    ]
    require(receipt.get("command") == expected_command,
            f"{relative(stage / (action + '.receipt.json'))} command differs")
    cap_artifacts(stage, action, receipt,
                  {report_path.name, fixture_path.name, rss_path.name,
                   stdout_path.name, stderr_path.name})
    report = read_json(report_path, relative(report_path))
    require(isinstance(report, dict), f"{relative(report_path)} is not an object")
    require(stdout_path.read_bytes().rstrip(b"\n") == report_path.read_bytes(),
            f"{relative(stdout_path)} does not reproduce report")
    require(report.get("schema") == "litchi.xlsx.cap-boundary-guard.v1"
            and report.get("tool") == CAP_EXAMPLE
            and report.get("case") == "valid"
            and report.get("size") == size
            and report.get("warmup_iterations") == 10
            and report.get("samples") == 100,
            f"{relative(report_path)} report envelope differs")
    cells = size * size
    event_count = 5 * cells + 2 * size + 6 + int(size <= 2)
    require(report.get("rows") == size and report.get("columns") == size
            and report.get("cells") == cells
            and report.get("last_cell") == f"{cap_column_name(size - 1)}{size}"
            and report.get("event_count") == event_count
            and report.get("expected_event_count") == event_count
            and report.get("event_count_formula") == "5*N*N+2*N+6+I(N<=2) (including EOF)"
            and report.get("comment_bytes") == (1024 * 1024 if size <= 2 else 0)
            and report.get("shared_provisional_event_cap") == 131072
            and report.get("ordinary_parser_event_cap") == 1000000
            and report.get("event_cap_relation") ==
            ("below" if event_count <= 131072 else "above")
            and report.get("source_stream_byte_limit") == 8 * 1024 * 1024
            and report.get("source_stream_eligible") is True,
            f"{relative(report_path)} event-bound identity differs")
    fixture_sha256 = sha(fixture_path)
    fixture_bytes = fixture_path.stat().st_size
    require(receipt["artifacts"].get(fixture_path.name) == fixture_sha256,
            f"{relative(report_path)} fixture receipt hash differs")
    require(report.get("archive_bytes") == fixture_bytes,
            f"{relative(report_path)} archive size differs")
    require(report.get("fixture_out") == str(fixture_path.resolve()),
            f"{relative(report_path)} fixture path differs")
    source_report = report.get("source")
    require(isinstance(source_report, dict), f"{relative(report_path)} source is missing")
    expected_source = {
        "shape": f"{size}x{size}", "rows": size, "columns": size,
        "format": "OOXML/XLSX",
        "fixture_kind": (
            "single_worksheet_sparse_numeric_with_comment"
            if size <= 2 else "single_worksheet_dense_numeric_grid"
        ),
        "generator": "litchi-xlsx-cap-boundary-stored-grid-v1",
        "worksheet_member": "xl/worksheets/sheet1.xml", "compression": "stored",
        "encoding": "UTF-8", "marker_free": True,
        "identity_method": "parent binds SHA-256 of fixture_dump bytes",
    }
    for key, expected in expected_source.items():
        require(source_report.get(key) == expected,
                f"{relative(report_path)} source.{key} differs")
    for key in ("source_xml_bytes", "worksheet_bytes", "worksheet_xml_bytes", "archive_bytes"):
        _positive_integer(source_report.get(key), f"{relative(report_path)} source.{key}")
    require(source_report["archive_bytes"] == fixture_bytes
            and source_report.get("fixture_dump") == str(fixture_path.resolve()),
            f"{relative(report_path)} source fixture binding differs")
    binary_report = report.get("binary")
    require(isinstance(binary_report, dict)
            and binary_report.get("path") == binary["path"]
            and binary_report.get("bytes") == binary["bytes"]
            and binary_report.get("profile") == "release"
            and binary_report.get("identity_method") ==
            "parent binds SHA-256 of the captured binary",
            f"{relative(report_path)} binary identity differs")
    correctness = report.get("correctness")
    require(isinstance(correctness, dict), f"{relative(report_path)} correctness is missing")
    for key in ("source_unchanged", "source_bytes_unchanged", "source_xml_unchanged",
                "valid_snapshot_values", "empty_commit_is_noop", "commit_outside_timing",
                "no_op_publication_exact"):
        require(correctness.get(key) is True,
                f"{relative(report_path)} correctness.{key} is not true")
    phase = report.get("phase")
    require(isinstance(phase, dict) and phase.get("name") == "edit_sheets"
            and phase.get("warmup_iterations") == 10
            and phase.get("samples") == 100
            and phase.get("sample_order") == list(range(100)),
            f"{relative(report_path)} phase envelope differs")
    durations = phase.get("duration_ns")
    require(isinstance(durations, list) and len(durations) == 100,
            f"{relative(report_path)} duration vector differs")
    for index, value in enumerate(durations):
        _positive_integer(value, f"{relative(report_path)} duration_ns[{index}]")
    ordered = sorted(durations)
    stats = {"p50": (ordered[49] + ordered[50]) // 2,
             "mean": sum(durations) / len(durations),
             "min": ordered[0], "max": ordered[-1]}
    return {
        "name": action, "stage": stage_name, "repeat": repeat, "size": size,
        "exit_code": 0,
        "start": receipt_record["start"], "end": receipt_record["end"],
        "receipt_sha256": receipt_record["receipt_sha256"],
        "report_sha256": sha(report_path), "fixture_sha256": fixture_sha256,
        "fixture_bytes": fixture_bytes, "rss_sha256": sha(rss_path),
        "binary_sha256": binary["sha256"], "manifest_sha256": source["manifest_sha256"],
        "statistics": stats, "rss": read_json(rss_path),
        "report": report,
    }


def cap_stage_records(stage_name: str, main_stage: dict[str, Any],
                      cap_plan: dict[str, Any]) -> dict[str, Any]:
    source = cap_source_stage(stage_name, main_stage, plan_data())
    binary, build = cap_binary(stage_name, source, cap_plan)
    rows = []
    for repeat in CAP_REPEATS:
        for size in (CAP_SIZES if repeat == 1 else tuple(reversed(CAP_SIZES))):
            rows.append(cap_report(stage_name, repeat, size, binary, source, cap_plan))
    expected = {"build-cap-boundary", *[row["name"] for row in rows]}
    actual = {path.name.removesuffix(".receipt.json")
              for path in (CAP_DIR / stage_name).glob("*.receipt.json")
              if path.is_file() and not path.is_symlink()}
    require(actual == expected,
            f"cap-boundary/{stage_name} receipt inventory differs")
    bound = {f"{name}.receipt.json" for name in expected}
    bound |= {name for row in rows for name in (
        f"{row['name']}.json", f"{row['name']}.fixture.bin",
        f"{row['name']}.rss.json", f"{row['name']}.stdout", f"{row['name']}.stderr")}
    bound |= {"source-manifest.json", "source.patch", "sources",
              "binary-cap-boundary.json", "build-cap-boundary.stdout",
              "build-cap-boundary.stderr"}
    if stage_name == "candidate":
        bound.add("source-diff.json")
    for path in (CAP_DIR / stage_name).iterdir():
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        require(path.name in bound,
                f"{relative(path)} is an unbound cap stage artifact")
    return {"stage": stage_name, "source": source, "manifest": source["manifest"],
            "manifest_sha256": source["manifest_sha256"], "binary": binary,
            "build": build, "rows": rows,
            "records": [build, *rows]}


def cap_timeline(cap_stages: dict[str, dict[str, Any]],
                 cap_plan: dict[str, Any], candidate_frozen_created: dt.datetime,
                 all_main_records: list[dict[str, Any]]) -> dict[str, Any]:
    records = [record for stage in cap_stages.values() for record in stage["records"]]
    validate_intervals(records)
    validate_intervals(all_main_records + records)
    require(all(record["start"] > cap_plan["_created"] for record in records),
            "a cap receipt predates the cap plan freeze")
    require(all(record["start"] > candidate_frozen_created
                for stage_name, stage in cap_stages.items()
                for record in stage["records"]
                if stage_name == "candidate"
                or (stage_name == "baseline" and record["name"] != "build-cap-boundary"
                    and "r2" in record["name"])),
            "candidate cap evidence predates candidate source freeze")
    groups: dict[tuple[str, int], list[dict[str, Any]]] = {}
    for stage_name, stage in cap_stages.items():
        for record in stage["rows"]:
            groups.setdefault((stage_name, record["repeat"]), []).append(record)
    expected_order = list(CAP_NATIVE_ORDER)
    starts = []
    for stage_name, repeat in expected_order:
        rows = groups.get((stage_name, repeat), [])
        require(len(rows) == len(CAP_SIZES),
                f"cap ABBA {stage_name} repeat {repeat} is incomplete")
        expected_sizes = list(CAP_SIZES if repeat == 1 else reversed(CAP_SIZES))
        ordered = sorted(rows, key=lambda row: row["start"])
        require([row["size"] for row in ordered] == expected_sizes,
                f"cap {stage_name} repeat {repeat} size order differs")
        starts.append((stage_name, repeat, ordered[0]["start"]))
    require([(stage, repeat) for stage, repeat, _ in sorted(starts, key=lambda item: item[2])]
            == expected_order,
            "cap capture order is not baseline/candidate/candidate/baseline")
    b = cap_stages["baseline"]["build"]
    c = cap_stages["candidate"]["build"]
    require(b["end"] <= min(row["start"] for row in groups[("baseline", 1)]),
            "baseline cap build does not precede repeat 1")
    require(max(row["end"] for row in groups[("baseline", 1)]) <= c["start"],
            "candidate cap build does not follow baseline repeat 1")
    require(c["end"] <= min(row["start"] for row in groups[("candidate", 1)]),
            "candidate cap build does not precede candidate repeat 1")
    require(max(row["end"] for row in groups[("candidate", 1)]) <=
            min(row["start"] for row in groups[("candidate", 2)])
            and max(row["end"] for row in groups[("candidate", 2)]) <=
            min(row["start"] for row in groups[("baseline", 2)]),
            "cap repeat groups overlap or violate ABBA")
    return {"status": "pass", "receipts": len(records),
            "order": [{"stage": stage, "repeat": repeat,
                       "start_utc": start.isoformat()}
                      for stage, repeat, start in sorted(starts, key=lambda item: item[2])],
            "_records": records}


def replay_cap_analysis(cap_plan: dict[str, Any]) -> dict[str, Any]:
    """Replay cap analysis and compare its JSON-serializable value exactly."""
    need(CAP_ANALYSIS, "cap-boundary/cap-analysis.json")
    expected = read_json(CAP_ANALYSIS, "cap-boundary/cap-analysis.json")
    # cap_run.py intentionally rebinds the shared ``run`` module's custody
    # paths to the cap subbundle while its analyzer is running.  Restore the
    # main-runner module (and import path) before replaying eager evidence;
    # otherwise eager_run.py would compare retained normal binaries against
    # cap-boundary-binaries and fail for a verifier-internal reason.
    runner = sys.modules.get("run")
    runner_state = {
        name: getattr(runner, name)
        for name in ("HERE", "REPO", "TARGET", "SCRATCH")
        if runner is not None and hasattr(runner, name)
    }
    path_state = list(sys.path)
    try:
        module = load_module(CAP_ANALYZER, "litchi_0546_cap_analysis")
        method = getattr(module, "analyze", None)
        require(callable(method), "cap-boundary/analyze.py has no analyze()")
        result = method()
    except Exception as error:  # noqa: BLE001 - retain analyzer context
        raise EvidenceError(f"cap-boundary analyzer replay failed: {error}") from error
    finally:
        sys.path[:] = path_state
        if runner is None:
            sys.modules.pop("run", None)
        else:
            for name, value in runner_state.items():
                setattr(runner, name, value)
    try:
        normalized = json.loads(json.dumps(result, ensure_ascii=False,
                                           sort_keys=True, allow_nan=False))
    except (TypeError, ValueError) as error:
        raise EvidenceError("cap-boundary analyzer result is not canonical JSON") from error
    require(normalized == expected,
            "cap-boundary/cap-analysis.json differs from deterministic replay")
    require(isinstance(result, dict)
            and result.get("schema") == "litchi.xlsx.cap-boundary-analysis.v1"
            and result.get("status") in {"pass", "reject"}
            and result.get("plan_sha256") == sha(CAP_PLAN)
            and result.get("driver_sha256") == sha(CAP_RUN),
            "cap-boundary analysis envelope differs")
    stages = result.get("stages")
    comparison = result.get("comparison")
    require(isinstance(stages, dict) and set(stages) == {"baseline", "candidate"}
            and all(isinstance(value, dict) and value.get("status") == "pass"
                    for value in stages.values()),
            "cap-boundary analysis stages are incomplete")
    require(isinstance(comparison, dict)
            and comparison.get("status") == result["status"]
            and isinstance(comparison.get("admission_passed"), bool)
            and ((result["status"] == "pass") is comparison["admission_passed"]),
            "cap-boundary admission result differs")
    rows = comparison.get("rows")
    require(isinstance(rows, list)
            and {(row.get("repeat"), row.get("size")) for row in rows}
            == {(repeat, size) for repeat in CAP_REPEATS for size in CAP_SIZES},
            "cap-boundary comparison matrix differs")
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("metrics"), dict)
                and set(row["metrics"]) == {"p50", "mean"},
                "cap-boundary comparison metric inventory differs")
        for metric in row["metrics"].values():
            require(isinstance(metric, dict) and isinstance(metric.get("passed"), bool),
                    "cap-boundary comparison metric is malformed")
    require(isinstance(comparison.get("adverse"), list)
            and isinstance(comparison.get("drift"), list),
            "cap-boundary adverse/drift vectors are missing")
    return {"status": "pass", "report_status": result["status"],
            "report_sha256": sha(CAP_ANALYSIS), "plan_sha256": sha(CAP_PLAN),
            "driver_sha256": sha(CAP_RUN),
            "admission_passed": comparison["admission_passed"],
            "rows": rows, "adverse": comparison["adverse"],
            "drift": comparison["drift"]}


def validate_cap_boundary(plan: dict[str, Any], stage_data: dict[str, dict[str, Any]],
                          candidate_changed: set[str], candidate_frozen_created: dt.datetime,
                          all_main_records: list[dict[str, Any]]) -> dict[str, Any]:
    cap_plan = cap_plan_data(plan, candidate_changed)
    for name in ("baseline", "candidate"):
        require((CAP_DIR / name).is_dir(),
                f"cap-boundary/{name} source stage is missing")
    cap_stages = {
        name: cap_stage_records(name, stage_data[name], cap_plan)
        for name in ("baseline", "candidate")
    }
    timeline = cap_timeline(cap_stages, cap_plan, candidate_frozen_created,
                            all_main_records)
    analysis = replay_cap_analysis(cap_plan)
    return {
        "status": "pass", "plan_sha256": sha(CAP_PLAN),
        "plan_created_utc": cap_plan["_created"].isoformat(),
        "stages": {name: {"manifest_sha256": cap_stages[name]["manifest_sha256"],
                           "binary": cap_stages[name]["binary"],
                           "receipts": len(cap_stages[name]["records"])}
                    for name in ("baseline", "candidate")},
        "timeline": timeline, "analysis": analysis,
        "_records": timeline["_records"],
        "_binaries": {name: {"cap-boundary": cap_stages[name]["binary"]}
                      for name in ("baseline", "candidate")},
    }


def validate_adverse_review(analyses: dict[str, Any], decision: dict[str, Any]) -> dict[str, Any]:
    """Bind every retained adverse/drift row to replayed analyzer output.

    0546 accepts the compact generic ``groups`` form (one exact source pointer
    and one list of reviewed rows) as well as the earlier named-list form.
    Both forms compare the original analyzer rows as canonical JSON multisets,
    so a count cannot conceal an omitted or substituted flag.
    """
    review_path = HERE / "adverse-review.json"
    review = read_json(review_path, "adverse-review.json")
    require(isinstance(review, dict), "adverse review is not an object")
    require(isinstance(review.get("schema_version"), int)
            and review["schema_version"] >= 2
            and review.get("status") == "complete",
            "adverse review is not complete evidence")
    disposition = review.get("disposition")
    require(isinstance(disposition, str), "adverse review disposition is missing")
    word = disposition.strip().lower()
    if decision["decision"] == "rejected":
        require(word in {"reject", "rejected", "rejected_and_reverted", "baseline_retained"},
                "adverse review disposition does not record rejection")
    else:
        require(word in {"accept", "accepted", "retain", "retained"},
                "adverse review disposition does not record acceptance")
    scope = review.get("scope")
    require(isinstance(scope, str) and "0546" in scope.lower()
            and "ooxml" in scope.lower() and "odf" in scope.lower(),
            "adverse review scope omits pilot or format priority")

    guard_report = first_existing([
        HERE / "guard-analysis.json", HERE / "guard-comparison.json",
        HERE / "refusal-guard-analysis.json",
    ])
    require(guard_report is not None, "guard analysis report is missing")
    report_bindings = {
        "comparison_sha256": HERE / "comparison.json",
        "allocation_analysis_sha256": HERE / "allocation-analysis.json",
        "guard_analysis_sha256": guard_report,
        "cap_analysis_sha256": CAP_ANALYSIS,
        "plan_sha256": PLAN,
    }
    # Conditional reports are optional for a rejected pilot.  If a later
    # review retains one, its digest may be bound here without changing the
    # required eight source groups for a pilot with no conditional captures.
    for field, path in (
        ("profile_analysis_sha256", PROFILE_ANALYSIS),
        ("hardware_analysis_sha256", HARDWARE_ANALYSIS),
        ("eager_analysis_sha256", EAGER_ANALYSIS),
    ):
        if field in review:
            require(path.is_file(), f"adverse review binds missing {path.name}")
            report_bindings[field] = path
    for field, path in report_bindings.items():
        require(valid_digest(review.get(field)) and review[field] == sha(path),
                f"adverse review {field} does not bind {path.name}")
    require(all(analyses.get(name, {}).get("status") == "pass"
                for name in ("native", "allocation", "guard", "cap")),
            "adverse review is not backed by complete analyzer replay")
    native = analyses["native"]
    native_review = review.get("native_admission")
    if native_review is not None:
        require(isinstance(native_review, dict), "adverse review native admission is malformed")
        if "decision" in native_review:
            require(native_review["decision"] == native.get("admission_status"),
                    "adverse review native admission differs")
        if "passed" in native_review:
            require(native_review["passed"] is
                    (isinstance(native.get("native_admission"), dict)
                     and native["native_admission"].get("passed") is True),
                    "adverse review native pass flag differs")
    allocation_review = review.get("allocation_diagnostics")
    if allocation_review is not None:
        require(isinstance(allocation_review, dict),
                "adverse review allocation diagnostics are malformed")
        if "planning_gate_passed" in allocation_review:
            require(allocation_review["planning_gate_passed"] is
                    (analyses["allocation"].get("planning_gate_passed") is True),
                    "adverse review allocation pass flag differs")
    guard_review = review.get("guard_admission")
    if guard_review is not None:
        require(isinstance(guard_review, dict), "adverse review guard admission is malformed")
        if "admission_status" in guard_review:
            require(guard_review["admission_status"] == analyses["guard"].get("admission_status"),
                    "adverse review guard admission differs")

    comparison = read_json(HERE / "comparison.json", "comparison.json")
    allocation = read_json(HERE / "allocation-analysis.json", "allocation-analysis.json")
    guard = read_json(guard_report, display_path(guard_report))
    require(isinstance(comparison, dict) and isinstance(comparison.get("comparison"), dict),
            "comparison report lacks comparison rows")
    require(isinstance(allocation, dict), "allocation report is not an object")
    require(isinstance(guard, dict) and isinstance(guard.get("comparison"), dict),
            "guard report lacks comparison rows")
    normal_guard = guard["comparison"].get("normal")
    allocation_guard = guard["comparison"].get("allocation")
    require(isinstance(normal_guard, dict) and isinstance(allocation_guard, dict),
            "guard report lacks normal/allocation comparison rows")
    expected_groups: list[tuple[str, str, list[Any]]] = [
        ("adverse_flags", "comparison.json:comparison.adverse_flags_over_five_percent",
         comparison["comparison"].get("adverse_flags_over_five_percent")),
        ("same_build_drift_flags", "comparison.json:comparison.same_build_drift_over_five_percent",
         comparison["comparison"].get("same_build_drift_over_five_percent")),
        ("allocation_adverse_flags", "allocation-analysis.json:adverse_flags",
         allocation.get("adverse_flags")),
        ("guard_adverse_flags", "guard-analysis.json:comparison.normal.adverse_flags_over_five_percent",
         normal_guard.get("adverse_flags_over_five_percent")),
        ("guard_adverse_flags", "guard-analysis.json:comparison.allocation.adverse_flags_over_five_percent",
         allocation_guard.get("adverse_flags_over_five_percent")),
        ("guard_same_build_drift_flags", "guard-analysis.json:comparison.same_build_drift_over_five_percent",
         guard["comparison"].get("same_build_drift_over_five_percent")),
        ("cap_adverse_flags", "cap-analysis.json:comparison.adverse", analyses["cap"].get("adverse")),
        ("cap_same_build_drift_flags", "cap-analysis.json:comparison.drift", analyses["cap"].get("drift")),
    ]
    if analyses.get("eager", {}).get("status") == "pass":
        eager_report = read_json(EAGER_ANALYSIS, "eager-analysis.json")
        eager_comparison = eager_report.get("comparison")
        require(isinstance(eager_comparison, dict),
                "eager analysis lacks comparison rows")
        expected_groups.extend([
            ("eager_adverse_flags",
             "eager-analysis.json:comparison.adverse_flags_over_five_percent",
             eager_comparison.get("adverse_flags_over_five_percent")),
            ("eager_same_build_drift_flags",
             "eager-analysis.json:comparison.same_build_drift_over_five_percent",
             eager_comparison.get("same_build_drift_over_five_percent")),
        ])
    for _, source, rows in expected_groups:
        require(isinstance(rows, list), f"{source} is not a list")
    expected_sources = {source for _, source, _ in expected_groups}

    def normalize_source(source: Any) -> str:
        require(isinstance(source, str) and source.strip(),
                "adverse review source pointer is missing")
        value = source.strip().replace("cap-boundary/", "")
        require(value in expected_sources, f"adverse review source pointer is unknown: {source}")
        return value

    reviewed: list[tuple[str, dict[str, Any]]] = []
    identifiers: set[str] = set()

    def add_row(source: str, item: Any, label: str, generic: bool) -> None:
        require(isinstance(item, dict), f"{label} is not an object")
        required = {"id", "original", "classification", "interpretation", "disposition"}
        if generic:
            require(set(item) == required, f"{label} fields differ")
        else:
            require(set(item) == required | {"source"}, f"{label} fields differ")
            require(normalize_source(item.get("source")) == source,
                    f"{label} source pointer differs")
        identifier = item.get("id")
        require(isinstance(identifier, str) and identifier and identifier not in identifiers,
                f"{label} ID is missing or duplicated")
        identifiers.add(identifier)
        require(isinstance(item.get("classification"), str)
                and item["classification"].strip(), f"{label} classification is missing")
        require(isinstance(item.get("interpretation"), str)
                and item["interpretation"].strip(), f"{label} interpretation is missing")
        require(isinstance(item.get("disposition"), str)
                and item["disposition"].strip(), f"{label} disposition is missing")
        reviewed.append((source, item["original"]))

    groups = review.get("groups")
    if groups is not None:
        require(isinstance(groups, list), "adverse review groups are not a list")
        by_source: dict[str, list[Any]] = {}
        for index, group in enumerate(groups):
            require(isinstance(group, dict) and set(group) == {"source", "rows"},
                    f"adverse-review.groups[{index}] fields differ")
            source = normalize_source(group["source"])
            require(source not in by_source, f"adverse review repeats source: {source}")
            require(isinstance(group["rows"], list),
                    f"adverse-review.groups[{index}].rows is not a list")
            by_source[source] = group["rows"]
        require(set(by_source) == expected_sources,
                "adverse review generic source inventory differs")
        for source, rows in by_source.items():
            for index, item in enumerate(rows):
                add_row(source, item, f"adverse-review.groups[{source}][{index}]", True)
    else:
        # Compatibility for the 0543-style named lists.  A duplicated review
        # key (the two guard adverse lists) is checked by its row source field.
        review_keys = list(dict.fromkeys(key for key, _, _ in expected_groups))
        for key in review_keys:
            require(isinstance(review.get(key), list), f"adverse review {key} is not a list")
        for key, source, rows in expected_groups:
            # Pull rows from the shared named list by the exact source pointer.
            expected_count = len(rows)
            matching = [item for item in review[key]
                        if isinstance(item, dict)
                        and isinstance(item.get("source"), str)
                        and normalize_source(item["source"]) == source]
            require(len(matching) == expected_count,
                    f"adverse review {key} does not cover {source}")
            for index, item in enumerate(matching):
                add_row(source, item, f"adverse-review.{key}[{index}]", False)

    expected_rows = [(source, row) for _, source, rows in expected_groups for row in rows]
    require(_multiset(reviewed, "adverse-review originals") ==
            _multiset(expected_rows, "analyzer adverse rows"),
            "adverse review originals do not exactly match analyzer flags")
    coverage = review.get("source_coverage")
    if coverage is not None:
        require(isinstance(coverage, list), "adverse review source coverage is not a list")
        expected_coverage = [{"source": source, "review_key": key, "count": len(rows)}
                             for key, source, rows in expected_groups]
        require(_multiset(coverage, "source coverage") ==
                _multiset(expected_coverage, "expected source coverage"),
                "adverse review source coverage differs")
    counts = review.get("counts")
    if counts is not None:
        require(isinstance(counts, dict)
                and counts.get("reviewed_flags") == len(expected_rows),
                "adverse review count does not cover every analyzer flag")
    summary = review.get("classification_summary")
    if summary is not None:
        classes: dict[str, int] = {}
        for source, item in reviewed:
            # The generic list has no source in each row; locate its
            # classification through the retained row object in the review.
            for group in groups or []:
                if normalize_source(group.get("source")) == source:
                    for candidate in group.get("rows", []):
                        if candidate.get("original") == item:
                            name = candidate["classification"]
                            classes[name] = classes.get(name, 0) + 1
                            break
        if not classes:
            for key in dict.fromkeys(key for key, _, _ in expected_groups):
                for item in review[key]:
                    classes[item["classification"]] = classes.get(item["classification"], 0) + 1
        require(isinstance(summary, dict) and set(summary) == set(classes),
                "adverse review classification summary differs")
        for name, count in classes.items():
            require(isinstance(summary[name], dict) and summary[name].get("count") == count,
                    f"adverse review classification count differs: {name}")
    note = read_text(HERE / "results-review.md", "results-review.md").lower()
    require("0546" in note and "individually" in note and "decision" in note,
            "results-review.md lacks explicit individual decision review")
    return {"status": "pass", "records": len(expected_rows),
            "unique_ids": len(identifiers),
            "comparison_sha256": review["comparison_sha256"],
            "allocation_analysis_sha256": review["allocation_analysis_sha256"],
            "guard_analysis_sha256": review["guard_analysis_sha256"],
            "cap_analysis_sha256": review["cap_analysis_sha256"],
            "plan_sha256": review["plan_sha256"]}


def final_quality(stage_name: str, records: list[dict[str, Any]], commands: dict[str, list[str]]) -> dict[str, Any]:
    expected = {f"final-{name}" for name in commands}
    selected = {record["name"] for record in records if record["name"].startswith("final-")}
    require(selected == expected,
            f"{stage_name} final quality receipt set differs: {sorted(selected ^ expected)}")
    chosen = [record for record in records if record["name"] in expected]
    require(all(record["exit_code"] == 0 for record in chosen),
            f"{stage_name} final quality contains a failed command")
    test = next(record for record in chosen if record["name"] == "final-quality-tests")
    require(test["counts"] is not None and test["counts"]["failed"] == 0,
            f"{stage_name} final quality tests have no successful Cargo summary")
    return {
        "stage": stage_name, "receipts": len(chosen),
        "test_counts": test["counts"],
    }


def conditional_receipts(records_by_stage: dict[str, list[dict[str, Any]]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for prefix in ("profile-", "hardware-", "eager-"):
        records = [record for rows in records_by_stage.values() for record in rows
                   if record["name"].startswith(prefix)]
        result[prefix[:-1]] = {
            "status": "pass" if records and all(record["exit_code"] == 0 for record in records)
            else ("failed" if records else "not-captured"),
            "receipts": len(records),
        }
    return result


def validate_decision(stages: dict[str, dict[str, Any]], records_by_stage: dict[str, list[dict[str, Any]]],
                      analyses: dict[str, Any], commands: dict[str, list[str]], plan: dict[str, Any]) -> dict[str, Any]:
    decision = read_json(DECISION, "decision.json")
    require(isinstance(decision, dict), "decision is not an object")
    final = decision.get("final_stage", decision.get("final_source"))
    safe_relative(final, "decision.final_stage")
    # Earlier campaign decisions used ``status``/``disposition`` for the
    # outcome while the 0546 driver uses ``decision``.  Accept the aliases
    # only for the same explicit accepted/rejected vocabulary.
    require(final in stages, "decision final stage is not a retained source attempt")
    word = decision.get("decision", decision.get("status", decision.get("disposition")))
    require(isinstance(word, str), "decision.decision is missing")
    normalized = word.strip().lower()
    accepted = normalized in {"accept", "accepted", "retain", "retained"}
    rejected = normalized in {"reject", "rejected", "rejected_and_reverted", "baseline_retained"}
    require(accepted ^ rejected, f"decision outcome is not accepted or rejected: {word}")
    quality = final_quality(final, records_by_stage[final], commands)
    # A supplied count is an assertion about the final stdout, never a value
    # invented by this verifier.
    for key in ("full_test_count", "focused_test_count"):
        if key in decision:
            _positive_integer(decision[key], f"decision.{key}", allow_zero=True)
    if "full_test_count" in decision:
        require(decision["full_test_count"] == quality["test_counts"]["passed"],
                "decision full test count differs from final quality output")
    digest_bindings = {
        "comparison_sha256": "comparison.json",
        "allocation_analysis_sha256": "allocation-analysis.json",
    }
    for field, name in digest_bindings.items():
        if field in decision:
            require(valid_digest(decision[field]) and (HERE / name).is_file()
                    and decision[field] == sha(HERE / name),
                    f"decision.{field} does not bind {name}")
    if "guard_analysis_sha256" in decision:
        guard_report = first_existing([
            HERE / "guard-analysis.json", HERE / "guard-comparison.json",
            HERE / "refusal-guard-analysis.json",
        ])
        require(guard_report is not None and decision["guard_analysis_sha256"] == sha(guard_report),
                "decision.guard_analysis_sha256 does not bind guard evidence")
    if "plan_sha256" in decision:
        require(decision["plan_sha256"] == sha(PLAN), "decision.plan_sha256 differs")
    if "cap_analysis_sha256" in decision:
        require(valid_digest(decision["cap_analysis_sha256"])
                and decision["cap_analysis_sha256"] == sha(CAP_ANALYSIS),
                "decision.cap_analysis_sha256 does not bind cap evidence")
    conditional_report_bindings = {
        "profile_analysis_sha256": PROFILE_ANALYSIS,
        "hardware_analysis_sha256": HARDWARE_ANALYSIS,
        "eager_analysis_sha256": EAGER_ANALYSIS,
    }
    for field, path in conditional_report_bindings.items():
        if field in decision:
            require(valid_digest(decision[field]) and path.is_file()
                    and decision[field] == sha(path),
                    f"decision.{field} does not bind conditional evidence")
    for field, stage_name in (("baseline_manifest_sha256", "baseline"),
                              ("candidate_manifest_sha256", "candidate"),
                              ("final_source_manifest_sha256", final)):
        if field in decision:
            require(valid_digest(decision[field])
                    and decision[field] == stages[stage_name]["manifest_sha256"],
                    f"decision.{field} does not bind {stage_name} manifest")
    native = analyses.get("native", {})
    allocation = analyses.get("allocation", {})
    guard = analyses.get("guard", {})
    cap = analyses.get("cap", {})
    if "native_passed" in decision:
        require(isinstance(decision["native_passed"], bool),
                "decision.native_passed is malformed")
        native_passed = (isinstance(native.get("native_admission"), dict)
                         and native["native_admission"].get("passed") is True)
        require(decision["native_passed"] == native_passed,
                "decision.native_passed disagrees with native analyzer")
    if "allocation_passed" in decision:
        require(isinstance(decision["allocation_passed"], bool),
                "decision.allocation_passed is malformed")
        require(decision["allocation_passed"] ==
                (allocation.get("planning_gate_passed") is True),
                "decision.allocation_passed disagrees with allocation analyzer")
    if "guard_passed" in decision:
        require(isinstance(decision["guard_passed"], bool),
                "decision.guard_passed is malformed")
        require(decision["guard_passed"] is (guard.get("admission_status") == "pass"),
                "decision.guard_passed disagrees with guard analyzer")
    for cap_field in ("cap_passed", "cap_boundary_passed"):
        if cap_field in decision:
            require(isinstance(decision[cap_field], bool),
                    f"decision.{cap_field} is malformed")
            require(decision[cap_field] ==
                    (cap.get("report_status") == "pass"
                     and cap.get("admission_passed") is True),
                    f"decision.{cap_field} disagrees with cap analyzer")
    if accepted:
        require(native.get("status") == "pass" and allocation.get("status") == "pass"
                and guard.get("status") == "pass",
                "accepted decision lacks complete native/allocation/guard analyzer replay")
        require(native.get("admission_status") == "eligible-for-conditional-lanes"
                and isinstance(native.get("native_admission"), dict)
                and native["native_admission"].get("passed") is True,
                "accepted decision does not satisfy native performance gates")
        require(allocation.get("planning_gate_passed") is True,
                "accepted decision does not satisfy allocation gates")
        require(guard.get("admission_status") == "pass",
                "accepted decision does not satisfy refusal-guard gates")
        require(cap.get("status") == "pass"
                and cap.get("report_status") == "pass"
                and cap.get("admission_passed") is True,
                "accepted decision does not satisfy cap-boundary gates")
        conditional_analysis = analyses.get("conditional", {})
        require(conditional_analysis.get("status") == "pass"
                and analyses.get("profile", {}).get("status") == "pass"
                and analyses.get("eager", {}).get("status") == "pass",
                "accepted decision lacks complete conditional lane replay")
        conditional = conditional_receipts(records_by_stage)
        require(conditional["profile"]["status"] == "pass"
                and conditional["eager"]["status"] == "pass",
                "accepted decision lacks successful profile and eager guard receipts")
    runtime_restored = decision.get("runtime_restored")
    require(isinstance(runtime_restored, bool),
            "decision.runtime_restored is required for integrated admission custody")
    require(runtime_restored is rejected,
            "decision runtime_restored flag contradicts accepted/rejected outcome")
    return {
        "decision": "accepted" if accepted else "rejected", "final_stage": final,
        "quality": quality, "native_status": native.get("status"),
        "allocation_status": allocation.get("status"), "guard_status": guard.get("status"),
        "conditional": conditional_receipts(records_by_stage),
    }


def validate_runtime_restoration(
    decision: dict[str, Any], stages: dict[str, dict[str, Any]],
    plan: dict[str, Any], records: list[dict[str, Any]],
) -> dict[str, Any]:
    """Validate an optional source-restoration custody note.

    The decision and current source manifest are the authoritative admission
    controls.  A separate note is useful when the rejected candidate was
    explicitly reverted, but an integrated run may omit it.  If present, its
    timestamp, selected manifest, and restored candidate-file inventory must
    agree with those controls and with the recorded command timeline.
    """
    if not RUNTIME_RESTORATION.exists():
        return {"status": "not-present"}
    value = read_json(RUNTIME_RESTORATION, "runtime-restoration.json")
    require(isinstance(value, dict), "runtime-restoration.json is not an object")
    created = parse_time(value.get("created_utc", value.get("utc")),
                         "runtime-restoration.created_utc")
    # Rejection restoration is normally written immediately after the
    # measured lanes and before the nine final baseline quality commands.
    # Permit either ordering relative to that final-quality lane, while still
    # requiring restoration to follow every measured (non-final) command.
    measured = [record for record in records
                if not record["name"].startswith("final-")]
    require(all(created > record["end"] for record in measured),
            "runtime restoration predates the measured command timeline")
    final = decision["final_stage"]
    require(value.get("final_stage", value.get("final_source")) == final,
            "runtime restoration final stage differs from decision")
    manifest_sha = value.get("final_source_manifest_sha256")
    require(valid_digest(manifest_sha)
            and manifest_sha == stages[final]["manifest_sha256"],
            "runtime restoration final source manifest differs")
    restored = value.get("candidate_files_restored")
    if restored is not None:
        require(isinstance(restored, list)
                and all(isinstance(name, str) for name in restored),
                "runtime restoration candidate file list is malformed")
        expected = sorted(changed_paths(stages["candidate"]["manifest"],
                                         stages["baseline"]["manifest"]))
        require(sorted(restored) == expected,
                "runtime restoration candidate file inventory differs")
    retained = value.get("retained_enabler")
    if retained is not None:
        if isinstance(retained, str):
            retained = [retained]
        require(isinstance(retained, list) and retained == [CAP_ENABLER],
                "runtime restoration retained enabler differs")
    reason = value.get("reason")
    require(isinstance(reason, str) and reason.strip(),
            "runtime restoration reason is missing")
    return {
        "status": "pass", "path": relative(RUNTIME_RESTORATION),
        "sha256": sha(RUNTIME_RESTORATION), "created_utc": created.isoformat(),
        "final_stage": final, "final_source_manifest_sha256": manifest_sha,
    }


def current_source(final_stage: str, stage_data: dict[str, dict[str, Any]], plan: dict[str, Any],
                   decision: dict[str, Any]) -> dict[str, Any]:
    final_manifest = stage_data[final_stage]["manifest"]
    current = source_inventory()
    require(set(final_manifest) == current, "current source inventory differs from final manifest")
    for name, digest in final_manifest.items():
        require(sha(REPO / name) == digest, f"current source hash differs: {name}")
    revision = revision_manifest(plan["revision"])
    changed = changed_paths(final_manifest, revision)
    roots = plan["_roots"]
    harness = harness_paths(plan)
    # Candidate roots are prefixes, not literal source entries.  The
    # standalone cap example is retained in the final baseline checkout as a
    # measured harness enabler and is admitted only with its exact frozen
    # source hash.
    for path in changed:
        require(path in harness or path == CAP_ENABLER or path_in_roots(path, roots),
                f"current source changed outside plan: {path}")
    if CAP_ENABLER in changed:
        require(path_in_roots(CAP_ENABLER, roots),
                "retained cap enabler is outside candidate source roots")
        require(final_manifest.get(CAP_ENABLER) == sha(REPO / CAP_ENABLER),
                "current cap enabler hash differs from retained source")
    if decision["decision"] == "rejected":
        baseline = stage_data["baseline"]["manifest"]
        require(final_manifest == baseline,
                "rejected decision final source is not the restored baseline")
    else:
        require(final_stage == "candidate", "accepted decision does not select candidate source")
    return {"manifest_sha256": stage_data[final_stage]["manifest_sha256"],
            "entries": len(final_manifest), "changed_files": sorted(changed)}


def validate_cleanup(plan: dict[str, Any], binaries: dict[str, dict[str, Any]],
                     records: list[dict[str, Any]]) -> dict[str, Any]:
    value = read_json(CLEANUP, "cleanup.json")
    require(isinstance(value, dict)
            and value.get("removed") == plan["owned_paths"]
            and value.get("owned_paths_absent") is True
            and value.get("accessible_process_references") == []
            and value.get("patch_replay_temporary_directories_absent") is True,
            "cleanup record differs")
    completed = parse_time(value.get("completed_utc"), "cleanup.completed_utc")
    require(all(completed > record["end"] for record in records),
            "cleanup completed before the last recorded command")
    for path in plan["owned_paths"]:
        target = Path(path)
        require(not target.exists() and not target.is_symlink(), f"owned target remains: {path}")
    retained = value.get("retained_binary_sha256_before_removal")
    expected_retained = {
        identity["path"]: identity["sha256"]
        for stage_binaries in binaries.values()
        for identity in stage_binaries.values()
    }
    require(len(expected_retained) == 10 and retained == expected_retained,
            "cleanup must bind exactly all ten measured binaries")
    return {"status": "pass", "owned_paths_absent": True, "removed": plan["owned_paths"],
            "completed_utc": completed.isoformat()}


def snapshot() -> dict[str, str]:
    result: dict[str, str] = {}
    for path in HERE.rglob("*"):
        require(not path.is_symlink(), f"evidence bundle contains symlink: {relative(path)}")
        if path.is_file() and path != SEAL:
            result[relative(path)] = sha(path)
    return result


def validate_seal() -> dict[str, Any]:
    need(SEAL, "SHA256SUMS")
    expected: dict[str, str] = {}
    for line in read_text(SEAL, "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2 and valid_digest(fields[0]), "malformed SHA256SUMS line")
        name = safe_relative(fields[1], "SHA256SUMS path")
        require(name != "SHA256SUMS" and name not in expected,
                "SHA256SUMS repeats or seals itself")
        expected[name] = fields[0]
    actual = snapshot()
    require(actual == expected, "SHA256SUMS inventory or digest differs")
    return {"status": "pass", "entries": len(expected), "sha256": sha(SEAL)}


def verify(*, precleanup: bool = False) -> dict[str, Any]:
    plan = plan_data()
    commands = quality_commands()
    gates = allocation_gate_data()
    frozen, frozen_created = frozen_inputs(plan)
    adrs = adr_manifest()
    validate_protocol()
    revision = revision_manifest(plan["revision"])
    roots = plan["_roots"]
    harness = harness_paths(plan)
    names = stage_names()
    baseline_manifest = source_manifest(HERE / "baseline/source-manifest.json")
    candidate_source_hashes = validate_candidate_source_hashes(plan, baseline_manifest)
    interrupted = validate_interrupted_baseline_attempt(
        sha(HERE / "baseline/source-manifest.json"))
    _, candidate_frozen_created = candidate_frozen_inputs(plan, baseline_manifest, frozen_created)
    # 0546 has no historical failed-quality or interrupted-build prerequisite.
    # If a later run retains such an attempt, its own validator below audits
    # it; the normal campaign timeline starts at the integrated freeze.
    failed_lower_bound = frozen_created
    stage_data: dict[str, dict[str, Any]] = {}
    for name in names:
        stage_data[name] = validate_stage_sources(
            name, plan, revision, roots, harness,
            baseline_manifest if name != "baseline" else None,
            final_candidate=False,
        )
    validate_harness_identity(stage_data, harness)
    if stage_data["baseline"]["changed_from_revision"]:
        require(set(stage_data["baseline"]["changed_from_revision"])
                <= harness | {CAP_ENABLER}
                | shared_source_paths(stage_data["baseline"]["manifest"], revision, plan),
                "baseline source contains a production candidate change")
    for name in names:
        if name != "baseline":
            require(all(path_in_roots(path, roots)
                        for path in stage_data[name]["changed_from_baseline"]),
                    f"{name} source changes outside candidate roots")
    failed_data = []
    for failed_name in failed_stage_names():
        failed_data.append(validate_failed_stage_sources(
            failed_name, plan, revision, baseline_manifest, harness,
            failed_lower_bound, candidate_frozen_created,
        ))
    next_candidate = validate_next_candidate(stage_data["candidate"], plan)

    allowed, metadata = expected_actions(plan, commands)
    records_by_stage: dict[str, list[dict[str, Any]]] = {}
    all_records: list[dict[str, Any]] = []
    for name in names:
        records = validate_stage_receipts(name, plan, commands, metadata, stage_data,
                                          allow_inflight=precleanup)
        records_by_stage[name] = records
        all_records.extend(records)
    validate_intervals(all_records)
    require(all(record["start"] > frozen_created for record in all_records),
            "a command receipt predates frozen inputs")
    require(all(record["start"] > candidate_frozen_created
                for record in all_records if record["stage"] == "candidate"
                or (record["stage"] == "baseline" and ("-r2-" in record["name"]
                    or record["name"].startswith(("profile-", "hardware-", "eager-",
                                                   "build-eager-", "symbols"))))),
            "a candidate or retained baseline receipt predates candidate frozen inputs")
    allocation_records = [record for record in all_records if record["name"].startswith(("alloc-", "guard-alloc-"))]
    require(all(record["start"] > gates["_created"] for record in allocation_records),
            "an allocation capture predates allocation gates")
    binaries: dict[str, dict[str, Any]] = {}
    for name, records in records_by_stage.items():
        binaries[name] = validate_binary_set(name, plan, records)
    # The strict ABBA matrix only applies once both named source stages have
    # all four capture lanes.  A partial early campaign remains incomplete.
    abba = validate_abba(records_by_stage, plan)
    guard_reports = validate_guard_reports(records_by_stage, plan, binaries)
    analyses = replay_analyses(plan, stage_data, records_by_stage, binaries)
    cap = validate_cap_boundary(
        plan, stage_data,
        set(stage_data["candidate"]["changed_from_baseline"]),
        candidate_frozen_created, all_records,
    )
    analyses["cap"] = cap["analysis"]
    conditional = validate_conditional_lanes(
        plan, records_by_stage, binaries, candidate_frozen_created, analyses,
    )
    analyses["conditional"] = {
        key: value for key, value in conditional.items() if key not in {"_results"}
    }
    analyses.update(conditional.get("analyses", {}))
    cap_records = cap["_records"]
    failed_intervals = [{
        "start": parse_time(row["start_utc"], "failed.start"),
        "end": parse_time(row["end_utc"], "failed.end"),
        "exit_code": row["exit_code"],
    } for row in failed_data]
    all_evidence_records = all_records + cap_records + failed_intervals
    validate_intervals(all_evidence_records)
    binaries_with_cap = {
        stage: {**binaries[stage], **cap["_binaries"][stage]}
        for stage in binaries
    }
    if not DECISION.exists():
        raise Pending("decision.json is missing")
    decision = validate_decision(stage_data, records_by_stage, analyses, commands, plan)
    adverse_review = validate_adverse_review(analyses, decision)
    # The selected final source is checked after decision semantics are known.
    final_name = decision["final_stage"]
    if final_name != "baseline":
        validate_stage_sources(final_name, plan, revision, roots, harness, baseline_manifest,
                               final_candidate=True)
    source = current_source(final_name, stage_data, plan, decision)
    runtime_restoration = validate_runtime_restoration(
        decision, stage_data, plan, all_evidence_records,
    )
    if CLEANUP.exists():
        cleanup: dict[str, Any] = validate_cleanup(plan, binaries_with_cap,
                                                   all_evidence_records)
    elif SEAL.exists():
        raise EvidenceError("SHA256SUMS exists before cleanup.json")
    elif precleanup:
        cleanup = {"status": "pending", "reason": "cleanup.json is pending"}
    else:
        raise Pending("cleanup.json is missing")
    if SEAL.exists():
        seal = validate_seal()
    elif precleanup:
        seal = {"status": "pending", "reason": "SHA256SUMS is pending"}
    else:
        raise Pending("SHA256SUMS is missing")
    status = "pass" if cleanup["status"] == "pass" and seal["status"] == "pass" else "incomplete"
    cap_public = {
        key: value for key, value in cap.items() if not key.startswith("_")
    }
    # ``cap_timeline`` keeps private datetime-bearing receipt rows for
    # cleanup and interval checks.  Do not expose those implementation rows
    # through the compact JSON result; its public order already contains UTC
    # strings and is the retained receipt summary.
    if isinstance(cap_public.get("timeline"), dict):
        cap_public["timeline"] = {
            key: value for key, value in cap_public["timeline"].items()
            if key != "_records"
        }
    return {
        "status": status,
        "scope": "0546 OOXML/XLSX shared worksheet traversal campaign; ODF deferred",
        "decision": decision,
        "stages": names,
        "failed_attempts": failed_data,
        "interrupted_attempt": interrupted,
        "candidate_source_hashes": candidate_source_hashes,
        "receipts": len(all_evidence_records),
        "failed_attempt_receipts": sum(record["exit_code"] != 0
                                       for record in all_evidence_records),
        "abba": abba,
        "cap_boundary": cap_public,
        "guard_reports": guard_reports,
        "analyses": analyses,
        "adverse_review": adverse_review,
        "next_candidate": next_candidate,
        "runtime_restoration": runtime_restoration,
        "source": source,
        "frozen_inputs_created_utc": frozen_created.isoformat(),
        "allocation_gates_created_utc": gates["_created"].isoformat(),
        "adr_entries": len(adrs),
        "cleanup": cleanup,
        "seal": seal,
        "performance_claim": "candidate decision is bound to replayed analyzer evidence",
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--precleanup", action="store_true",
                        help="allow cleanup.json and SHA256SUMS to remain pending")
    parser.add_argument("--strict", action="store_true",
                        help="require cleanup.json and SHA256SUMS")
    parser.add_argument("--output", type=Path,
                        help="write compact JSON outside this evidence bundle")
    args = parser.parse_args()
    precleanup = args.precleanup and not args.strict
    try:
        result = verify(precleanup=precleanup)
        output = result
    except Pending as error:
        output = {
            "status": "incomplete", "reason": str(error),
            "scope": "0546 OOXML/XLSX shared worksheet traversal campaign; ODF deferred",
        }
    except EvidenceError as error:
        output = {
            "status": "fail", "reason": str(error),
            "scope": "0546 OOXML/XLSX shared worksheet traversal campaign; ODF deferred",
        }
    text = json.dumps(output, indent=2, sort_keys=True) + "\n"
    if args.output:
        target = args.output.resolve()
        require(target != HERE.resolve() and HERE.resolve() not in target.parents,
                "refusing output inside evidence bundle")
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text(text, encoding="utf-8")
    print(text, end="")
    return 1 if output["status"] == "fail" or (
        output["status"] == "incomplete" and not precleanup
    ) else 0


if __name__ == "__main__":
    raise SystemExit(main())
