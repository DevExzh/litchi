"""Read-only custody verifier for the 0543 XLSX traversal follow-up.

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
REPO = HERE.parents[3]
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
GUARD_RUN = HERE / "guard_run.py"
QUALITY_PLAN = HERE / "quality-plan.json"
ALLOCATION_GATES = HERE / "allocation-gates.json"
PROTOCOL = HERE / "protocol.md"
ADR = HERE / "adr-manifest.json"
FROZEN = HERE / "frozen-inputs.json"
BASELINE_FROZEN = HERE / "baseline-frozen-inputs.json"
CANDIDATE_FROZEN = HERE / "candidate-frozen-inputs.json"
CONDITIONAL_FROZEN = HERE / "conditional-frozen-inputs.json"
DECISION = HERE / "decision.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
TARGET = Path("/home/zhuhe/litchi-goal-0543-target")
CAP_DIR = HERE / "cap-boundary"
CAP_PLAN = CAP_DIR / "plan.json"
CAP_RUN = CAP_DIR / "cap_run.py"
CAP_ANALYZER = CAP_DIR / "analyze.py"
CAP_ANALYSIS = CAP_DIR / "cap-analysis.json"
CAP_GUARD = CAP_DIR / "guard.rs"
CAP_SOURCE_PREPARATION = CAP_DIR / "source-preparation.json"
CAP_TARGET = TARGET / "cap-boundary-binaries"
HARDWARE_ANALYSIS = HERE / "hardware-analysis.json"
HARDWARE_REVIEW = HERE / "hardware-review.json"
HARDWARE_REVIEW_SCRIPT = HERE / "review_hardware.py"
HARDWARE_REVIEW_MARKDOWN = HERE / "hardware-review.md"
CAP_EXAMPLE = "perf_cap_boundary"
CAP_SIZES = (160, 164, 256)
CAP_REPEATS = (1, 2)
CAP_NATIVE_ORDER = (
    ("baseline", 1), ("candidate", 1),
    ("candidate", 2), ("baseline", 2),
)
CAP_TIME_FORMAT = (
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


def sha_bytes(value: bytes) -> str:
    """Hash retained bytes with the same SHA-256 convention as stage files."""

    return hashlib.sha256(value).hexdigest()


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
    """Return the frozen integration-test prefixes shared by both stages."""

    value = plan.get("shared_test_source_prefixes")
    require(isinstance(value, list) and value,
            "plan.shared_test_source_prefixes is not a nonempty list")
    result: list[str] = []
    for prefix in value:
        safe_relative(prefix, "shared test source prefix")
        require(prefix.startswith("crates/litchi-xlsx/tests/"),
                f"shared test source prefix is outside XLSX integration tests: {prefix}")
        require(prefix == prefix.rstrip("/") and prefix,
                f"shared test source prefix is malformed: {prefix}")
        result.append(prefix)
    require(len(set(result)) == len(result),
            "plan.shared_test_source_prefixes repeats a prefix")
    return result


def is_shared_test_path(name: str, plan: dict[str, Any]) -> bool:
    """Match an owner file and descendants named by a frozen test prefix."""

    return any(
        name == prefix or name == prefix + ".rs" or name.startswith(prefix + "/")
        for prefix in shared_test_prefixes(plan)
    )


def test_only_path(name: str, plan: dict[str, Any]) -> bool:
    """Allow only the explicitly frozen integration-test source inventory."""

    return is_shared_test_path(name, plan)


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
    shared_test_prefixes(value)
    if "candidate_patch_sha256" in value:
        require(valid_digest(value["candidate_patch_sha256"]),
                "plan candidate patch digest is malformed")
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
    }, "allocation gates were not frozen before capture")
    require("planning_allocation_calls_reduction_percent" not in value,
            "allocation calls must remain diagnostic in 0543")
    for key in required - {"status", "comparison", "scope", "created_utc"}:
        _finite_nonnegative(value[key], f"allocation-gates.{key}")
        require(float(value[key]) == 1.0, f"allocation-gates.{key} must be 1.0")
    require(isinstance(value["scope"], str) and value["scope"],
            "allocation-gates.scope is invalid")
    require(value["comparison"] == "maximum candidate <= 1.01 * minimum baseline in every shape/repeat",
            "allocation comparison rule differs")
    return {**value, "_created": parse_time(value["created_utc"], "allocation-gates.created_utc")}


def validate_allocation_status_correction() -> dict[str, Any]:
    """Bind a retained analyzer compatibility fix to its before/after bytes.

    The 0543 gate envelope was frozen with the broader campaign status
    spelling.  The copied 0542 analyzer initially accepted only its older
    spelling; root retained that failed replay and the exact one-line fix
    before replaying the allocation report.  Keep the diagnostic chain
    auditable without treating it as a measurement receipt.
    """

    before_path = HERE / "analyze_allocation.before-status-fix.py"
    after_path = HERE / "analyze_allocation.py"
    correction_path = HERE / "analysis-status-correction.json"
    present = [path.exists() for path in (before_path, correction_path)]
    if not any(present):
        return {"status": "not-present"}
    require(all(present),
            "allocation analyzer status correction is only partially retained")
    before = read_text(before_path, relative(before_path))
    after = read_text(after_path, relative(after_path))
    value = read_json(correction_path, relative(correction_path))
    require(isinstance(value, dict),
            "allocation analyzer status correction is not an object")
    expected = {
        "created_utc", "failed_command", "exit_code", "diagnostic", "cause",
        "before_sha256", "after_sha256", "plan_sha256",
        "measurement_receipts_changed",
    }
    require(set(value) == expected,
            "allocation analyzer status correction fields differ")
    created = parse_time(value.get("created_utc"),
                         "analysis-status-correction.created_utc")
    require(isinstance(value.get("failed_command"), list)
            and all(isinstance(item, str) and item for item in value["failed_command"])
            and value["failed_command"] == [
                "python3", "-B", after_path.relative_to(REPO).as_posix(),
            ], "allocation analyzer correction command differs")
    require(value.get("exit_code") == 1,
            "allocation analyzer correction did not retain the failed replay")
    require(isinstance(value.get("diagnostic"), str)
            and value["diagnostic"].strip()
            and isinstance(value.get("cause"), str)
            and value["cause"].strip(),
            "allocation analyzer correction explanation is missing")
    require(value.get("before_sha256") == sha(before_path)
            and value.get("after_sha256") == sha(after_path),
            "allocation analyzer correction before/after hash differs")
    # The correction record calls this binding ``plan_sha256`` for historical
    # compatibility; its value is the separately frozen allocation-gates
    # envelope, whose hash is what the corrected analyzer consumes here.
    require(value.get("plan_sha256") == sha(ALLOCATION_GATES),
            "allocation analyzer correction gate-envelope hash differs")
    require(value.get("measurement_receipts_changed") is False,
            "allocation analyzer correction claims measurement changes")
    require(
        'gates.get("status") == "frozen-before-first-allocation-capture"' in before,
        "retained pre-fix allocation analyzer lacks the failed status check",
    )
    require(
        'gates.get("status") in {"frozen-before-first-allocation-capture",' in after
        and '"frozen-before-build-and-capture"' in after,
        "retained allocation analyzer lacks the bounded status compatibility fix",
    )
    return {
        "status": "pass", "created_utc": created.isoformat(),
        "before_sha256": value["before_sha256"],
        "after_sha256": value["after_sha256"],
        "correction_sha256": sha(correction_path),
        "measurement_receipts_changed": False,
    }


def frozen_inputs(plan: dict[str, Any]) -> tuple[dict[str, Any], dt.datetime]:
    value = read_json(FROZEN, "frozen-inputs.json")
    require(isinstance(value, dict) and isinstance(value.get("files"), dict),
            "frozen input envelope is malformed")
    created = parse_time(value.get("created_utc"), "frozen-inputs.created_utc")
    require(created >= plan["_created"], "frozen inputs predate the plan")
    files = value["files"]
    # The capture drivers and their phase wrappers are frozen before the
    # first baseline build.  Analysis scripts, protocol prose, and the ADR
    # manifest have their own hash/replay checks and may be added while the
    # pilot is being assembled (the initial baseline is frozen before those
    # analyzers were copied into this bundle).
    required = {
        "plan.json", "run.py", "guard_run.py", "baseline_phase.py",
        "candidate_phase.py", "final_quality.py", "quality-plan.json",
        "allocation-gates.json", "differential-tests.patch", "test-design.md",
        "protocol.md", "adr-manifest.json",
    }
    require(required <= set(files),
            f"frozen inputs omit required files: {sorted(required - set(files))}")
    for name, digest in files.items():
        safe_relative(name, "frozen input path")
        require(valid_digest(digest), f"frozen input digest is malformed: {name}")
        target = HERE / name
        require(target.is_file() and not target.is_symlink(),
                f"frozen input file is missing: {name}")
        require(sha(target) == digest, f"frozen input hash differs: {name}")
    return value, created


def candidate_frozen_inputs(plan: dict[str, Any], baseline: dict[str, str],
                            frozen_created: dt.datetime) -> tuple[dict[str, Any], dt.datetime]:
    """Validate the second freeze that binds the applied candidate patch.

    The first freeze predates the baseline campaign and therefore cannot bind
    the later differential-test and candidate patches.  The candidate
    envelope records those exact inputs and the source hashes used for the
    candidate freeze; every candidate receipt must occur after this point.
    """
    value = read_json(CANDIDATE_FROZEN, "candidate-frozen-inputs.json")
    require(isinstance(value, dict) and isinstance(value.get("files"), dict)
            and isinstance(value.get("candidate_sources"), dict),
            "candidate frozen input envelope is malformed")
    created = parse_time(value.get("created_utc"), "candidate-frozen-inputs.created_utc")
    require(created >= frozen_created, "candidate freeze predates the initial freeze")
    required = {
        "candidate.patch", "applied-candidate.patch", "candidate-design.md",
        "source-review.md", "test-review.md", "apply_candidate.py",
        "differential-tests-final.patch", "baseline-frozen-inputs.json",
    }
    files = value["files"]
    require(required <= set(files),
            f"candidate frozen inputs omit required files: {sorted(required - set(files))}")
    for name, digest in files.items():
        safe_relative(name, "candidate frozen input path")
        require(valid_digest(digest), f"candidate frozen input digest is malformed: {name}")
        target = HERE / name
        require(target.is_file() and not target.is_symlink(),
                f"candidate frozen input file is missing: {name}")
        require(sha(target) == digest, f"candidate frozen input hash differs: {name}")
    require(value.get("baseline_manifest_sha256") == sha(HERE / "baseline/source-manifest.json"),
            "candidate freeze baseline manifest binding differs")
    sources = value["candidate_sources"]
    require(sources, "candidate frozen source inventory is empty")
    roots = plan["_roots"]
    for name, digest in sources.items():
        safe_relative(name, "candidate frozen source path")
        require(path_in_roots(name, roots) and valid_digest(digest),
                f"candidate frozen source entry is malformed: {name}")
    candidate_manifest_path = HERE / "candidate/source-manifest.json"
    candidate_manifest = source_manifest(candidate_manifest_path)
    changed = changed_paths(candidate_manifest, baseline)
    require(set(sources) == changed,
            "candidate frozen source inventory differs from candidate manifest diff")
    for name, digest in sources.items():
        require(candidate_manifest.get(name) == digest,
                f"candidate frozen source hash differs: {name}")
    return value, created


def conditional_frozen_inputs(plan: dict[str, Any],
                              candidate_frozen_created: dt.datetime) -> dict[str, Any]:
    """Validate the freeze separating pilot admission from conditional lanes.

    The profile, hardware, and eager controls are optional while the pilot is
    incomplete.  Once their plan or receipts exist, this second freeze must
    bind the exact runners, analyzers, and primary plan before any conditional
    child starts.  An accepted outcome is required to carry this freeze.
    """

    conditional_paths = {
        "plan.json": PLAN,
        "run.py": RUN,
        "eager-plan.json": HERE / "eager-plan.json",
        "eager_run.py": HERE / "eager_run.py",
        "analyze_eager.py": HERE / "analyze_eager.py",
        "analyze_planning.py": HERE / "analyze_planning.py",
        "analyze_hardware.py": HERE / "analyze_hardware.py",
    }
    value_exists = CONDITIONAL_FROZEN.exists()
    evidence_exists = (HERE / "eager-plan.json").exists() or any(
        path.is_file() and not path.is_symlink()
        for stage in ("baseline", "candidate")
        for path in (HERE / stage).glob(
            "*"
        )
        if path.name.startswith(("profile-", "hardware-", "eager-", "symbols"))
    ) or any(
        (HERE / name).is_file() and not (HERE / name).is_symlink()
        for name in (
            "symbol-observation.json", "planning-profile-analysis.json",
            "profile-analysis.json", "profile-comparison.json",
            "eager-analysis.json", "eager-guard-comparison.json",
            "hardware-analysis.json", "hardware-comparison.json",
        )
    )
    if not value_exists:
        require(not evidence_exists,
                "conditional evidence exists without conditional-frozen-inputs.json")
        return {"status": "not-present"}
    value = read_json(CONDITIONAL_FROZEN, "conditional-frozen-inputs.json")
    require(isinstance(value, dict), "conditional frozen-inputs envelope is malformed")
    created = parse_time(value.get("created_utc"),
                         "conditional-frozen-inputs.created_utc")
    require(created > candidate_frozen_created,
            "conditional freeze does not follow candidate frozen inputs")
    files = value.get("files", value.get("hashes"))
    require(isinstance(files, dict),
            "conditional frozen-inputs file hash map is missing")
    required = set(conditional_paths)
    require(required <= set(files),
            f"conditional frozen-inputs omit required files: {sorted(required - set(files))}")
    for name, digest in files.items():
        safe_relative(name, "conditional frozen-input path")
        require(valid_digest(digest),
                f"conditional frozen-input digest is malformed: {name}")
        target = HERE / name
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"conditional frozen-input hash differs: {name}")
    for name, target in conditional_paths.items():
        require(files[name] == sha(target),
                f"conditional frozen-input binding differs: {name}")
    eager_plan = read_json(HERE / "eager-plan.json", "eager-plan.json")
    require(isinstance(eager_plan, dict)
            and eager_plan.get("schema") == "litchi-0543-eager-guard-plan-v1"
            and eager_plan.get("status") == "frozen-before-capture",
            "conditional eager plan is not frozen before capture")
    require(eager_plan.get("primary_plan_sha256") == sha(PLAN)
            and eager_plan.get("run_script_sha256") == sha(RUN)
            and eager_plan.get("eager_run_script_sha256") == sha(HERE / "eager_run.py")
            and eager_plan.get("analyzer_script_sha256") == sha(HERE / "analyze_eager.py"),
            "conditional eager plan script bindings differ")
    return {
        "status": "pass", "created_utc": created.isoformat(),
        "files": {name: files[name] for name in sorted(files)},
        "eager_plan_sha256": sha(HERE / "eager-plan.json"),
    }


def cap_plan_data(plan: dict[str, Any], candidate_changed: set[str],
                  candidate_frozen_created: dt.datetime) -> tuple[dict[str, Any], dt.datetime]:
    """Validate the supplemental cap-boundary plan before admitting its lane.

    The cap guard has its own evidence directory and source root list because
    its temporary example is intentionally outside the main 0543 snapshot.
    The example is a measurement enabler only: the candidate diff remains the
    five files frozen by the main candidate envelope.
    """

    if not CAP_DIR.exists():
        raise Pending("cap-boundary evidence directory is missing")
    require(CAP_DIR.is_dir() and not CAP_DIR.is_symlink(),
            "cap-boundary evidence directory is not regular")
    for path, label in (
        (CAP_PLAN, "cap-boundary/plan.json"),
        (CAP_RUN, "cap-boundary/cap_run.py"),
        (CAP_ANALYZER, "cap-boundary/analyze.py"),
        (CAP_GUARD, "cap-boundary/guard.rs"),
        (CAP_SOURCE_PREPARATION, "cap-boundary/source-preparation.json"),
    ):
        need(path, label)
    value = read_json(CAP_PLAN, "cap-boundary/plan.json")
    require(isinstance(value, dict), "cap-boundary plan is not an object")
    created = parse_time(value.get("created_utc"), "cap-boundary.plan.created_utc")
    require(created > candidate_frozen_created,
            "cap-boundary plan predates candidate frozen inputs")
    status = value.get("status")
    if status == "draft":
        raise Pending("cap-boundary plan is still draft")
    require(status in {"frozen-before-build-and-capture", "frozen-before-capture"},
            "cap-boundary plan is not frozen before build and capture")
    require(value.get("revision") == plan["revision"],
            "cap-boundary plan revision differs from main plan")
    priority = value.get("priority")
    require(isinstance(priority, str) and "ole2/ooxml" in priority.lower()
            and "odf" in priority.lower(),
            "cap-boundary plan priority omits OLE2/OOXML first and ODF deferral")
    require(value.get("owned_paths") == plan["owned_paths"],
            "cap-boundary owned target differs from main plan")
    require(value.get("cpu") == plan["cpu"], "cap-boundary CPU differs from main plan")
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
            "cap-boundary size, repetition, or child configuration differs")
    require(isinstance(boundary.get("fixture_output"), str)
            and boundary["fixture_output"] == "stage/cap-native-r{repeat}-{size}.fixture.bin",
            "cap-boundary fixture output template differs")
    require(isinstance(boundary.get("timing_scope"), str)
            and boundary["timing_scope"].strip()
            and isinstance(boundary.get("performance_claim"), str)
            and "planning" in boundary["performance_claim"].lower(),
            "cap-boundary timing scope or claim is missing")
    roots = value.get("candidate_source_roots")
    expected_roots = [
        "crates/litchi-xlsx/src/cell_values/",
        "crates/litchi-xlsx/src/raw/worksheet/",
        "crates/litchi-xlsx/examples/",
    ]
    require(roots == expected_roots, "cap-boundary candidate roots differ")
    for root in roots:
        safe_relative(root.rstrip("/"), "cap-boundary candidate root")
        require(root.endswith("/") and (REPO / root.rstrip("/")).is_dir(),
                f"cap-boundary candidate root is invalid: {root}")
    candidate_files = value.get("candidate_files")
    require(isinstance(candidate_files, list) and len(candidate_files) == len(candidate_changed)
            and len(set(candidate_files)) == len(candidate_files)
            and set(candidate_files) == candidate_changed
            and all(path_in_roots(name, plan["_roots"]) for name in candidate_files),
            "cap-boundary candidate file inventory differs from main candidate diff")
    require(value.get("shared_test_source_prefixes") == shared_test_prefixes(plan),
            "cap-boundary shared test prefix differs")
    temporary = value.get("temporary_enabler")
    require(isinstance(temporary, dict)
            and temporary.get("source") == "crates/litchi-xlsx/examples/perf_cap_boundary.rs"
            and temporary.get("source_root_included_for_snapshot") is True
            and temporary.get("retained_runtime_api") is False,
            "cap-boundary temporary enabler binding differs")
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
    require(isinstance(admission, str)
            and "1.05" in admission
            and "every size" in admission.lower()
            and "individual" in admission.lower(),
            "cap-boundary admission rule is missing")
    return {**value, "_created": created, "_roots": roots}, created


def cap_source_preparation(plan: dict[str, Any], main_baseline: dict[str, str],
                           main_candidate: dict[str, str], candidate_changed: set[str],
                           candidate_frozen_created: dt.datetime,
                           cap_plan_created: dt.datetime) -> tuple[dict[str, Any], dt.datetime]:
    """Bind the temporary example preparation to the already frozen main stages."""

    value = read_json(CAP_SOURCE_PREPARATION, "cap-boundary/source-preparation.json")
    require(isinstance(value, dict), "cap-boundary source preparation is not an object")
    created = parse_time(value.get("created_utc"),
                         "cap-boundary.source-preparation.created_utc")
    require(created > candidate_frozen_created and created <= cap_plan_created,
            "cap-boundary source preparation does not precede the cap plan freeze")
    action = value.get("action")
    require(isinstance(action, str)
            and "restored main baseline" in action.lower()
            and "temporary cap guard example" in action.lower(),
            "cap-boundary source preparation action is not explicit")
    require(value.get("baseline_source_manifest_sha256") ==
            sha(HERE / "baseline/source-manifest.json")
            and value.get("candidate_source_manifest_sha256") ==
            sha(HERE / "candidate/source-manifest.json"),
            "cap-boundary source preparation main manifest binding differs")
    files = value.get("from_candidate_files")
    require(isinstance(files, dict) and set(files) == candidate_changed,
            "cap-boundary source preparation candidate file inventory differs")
    for name in candidate_changed:
        require(files[name] == main_candidate.get(name)
                and valid_digest(files[name]),
                f"cap-boundary source preparation candidate hash differs: {name}")
    last = value.get("last_main_receipt_sha256")
    require(valid_digest(last), "cap-boundary last main receipt digest is malformed")
    matching: list[tuple[dt.datetime, dt.datetime]] = []
    for stage in ["baseline", "candidate", *failed_stage_names()]:
        folder = HERE / stage
        if not folder.is_dir() or folder.is_symlink():
            continue
        for path in folder.glob("*.receipt.json"):
            if path.is_file() and not path.is_symlink() and sha(path) == last:
                start, end = interval(read_json(path), relative(path))
                matching.append((start, end))
    require(matching and all(end <= created for _, end in matching),
            "cap-boundary last main receipt is not retained before preparation")
    require(main_baseline and main_candidate,
            "main source manifests are empty before cap preparation")
    return {
        "status": "pass", "created_utc": created.isoformat(),
        "last_main_receipt_sha256": last,
    }, created


def cap_analyzer_correction(initial_files: dict[str, Any],
                            current_files: dict[str, Any],
                            frozen_created: dt.datetime) -> dict[str, Any]:
    """Validate a post-capture analyzer correction without reopening inputs.

    The cap freeze intentionally records the analyzer bytes used when the
    lane was prepared.  The final analyzer was corrected after the captures to
    make its fixture-path comparison stage-independent and to retain RSS
    diagnostics.  Keep that correction as a separately bound replay artifact;
    it must not be mistaken for a source, plan, driver, guard, or capture
    change.
    """

    path = CAP_DIR / "analyzer-correction.json"
    need(path, "cap-boundary/analyzer-correction.json")
    value = read_json(path, "cap-boundary/analyzer-correction.json")
    require(isinstance(value, dict),
            "cap-boundary analyzer correction is not an object")
    require(value.get("schema") ==
            "litchi.xlsx.cap-boundary-analyzer-correction.v1"
            and value.get("status") == "corrected",
            "cap-boundary analyzer correction envelope differs")
    created = parse_time(value.get("created_utc"),
                         "cap-boundary.analyzer-correction.created_utc")
    require(created > frozen_created,
            "cap-boundary analyzer correction predates corrected freeze")
    initial = value.get("initial")
    final = value.get("final")
    require(isinstance(initial, dict) and isinstance(final, dict)
            and initial.get("path") == "analyze.initial.py"
            and final.get("path") == "analyze.py"
            and initial.get("sha256") == initial.get("frozen_analyzer_sha256")
            and initial.get("sha256") == initial_files.get("cap-boundary/analyze.py")
            and valid_digest(initial.get("sha256"))
            and valid_digest(final.get("sha256"))
            and final.get("sha256") == sha(CAP_ANALYZER),
            "cap-boundary analyzer correction source hashes differ")
    require(current_files.get("cap-boundary/analyze.py") ==
            initial.get("sha256"),
            "cap-boundary correction freeze analyzer hash was rewritten")
    corrections = value.get("corrections")
    require(isinstance(corrections, list) and corrections,
            "cap-boundary analyzer correction list is missing")
    correction_ids = {item.get("id") for item in corrections
                      if isinstance(item, dict)}
    require(correction_ids == {
        "fixture-matrix-path-independent",
        "rss-adverse-and-drift-diagnostics",
        "terminal-replay-bindings",
    } and all(isinstance(item.get("rationale"), str)
              and item["rationale"].strip() for item in corrections
              if isinstance(item, dict)),
            "cap-boundary analyzer correction details differ")
    admission = value.get("admission")
    require(isinstance(admission, dict)
            and admission.get("unchanged") is True
            and admission.get("gate_metrics") == ["planning p50", "planning mean"]
            and admission.get("candidate_to_baseline_max_ratio") == 1.05
            and admission.get("rss_used_for_gate") is False
            and admission.get("captures_changed") is False
            and admission.get("plan_driver_guard_changed") is False,
            "cap-boundary analyzer correction changes the admission envelope")
    unchanged = value.get("frozen_input_hashes_unchanged")
    require(isinstance(unchanged, dict)
            and set(unchanged) == {"plan.json", "cap_run.py", "guard.rs"}
            and unchanged["plan.json"] == sha(CAP_PLAN)
            and unchanged["cap_run.py"] == sha(CAP_RUN)
            and unchanged["guard.rs"] == sha(CAP_GUARD),
            "cap-boundary analyzer correction input bindings differ")
    replayed = value.get("replayed_analysis")
    require(isinstance(replayed, dict)
            and replayed.get("path") == "cap-analysis.json"
            and replayed.get("sha256") == sha(CAP_ANALYSIS)
            and replayed.get("status") in {"pass", "reject"}
            and isinstance(replayed.get("reason"), str)
            and replayed["reason"].strip(),
            "cap-boundary analyzer correction replay binding differs")
    return {"status": "pass", "created_utc": created.isoformat(),
            "initial_sha256": initial["sha256"],
            "final_sha256": final["sha256"],
            "correction_sha256": sha(path),
            "replayed_analysis_sha256": replayed["sha256"]}


def cap_frozen_inputs(cap_plan_created: dt.datetime,
                      candidate_frozen_created: dt.datetime,
                      preparation_created: dt.datetime) -> tuple[dict[str, Any], dt.datetime]:
    """Validate the cap correction freeze and its retained first attempt."""

    initial_path = CAP_DIR / "frozen-inputs.initial.json"
    for path, label in ((initial_path, "cap-boundary/frozen-inputs.initial.json"),
                        (CAP_DIR / "frozen-inputs.json", "cap-boundary/frozen-inputs.json")):
        need(path, label)
    initial = read_json(initial_path, "cap-boundary/frozen-inputs.initial.json")
    value = read_json(CAP_DIR / "frozen-inputs.json", "cap-boundary/frozen-inputs.json")
    require(isinstance(initial, dict) and isinstance(value, dict),
            "cap-boundary frozen-input envelopes are malformed")
    created = parse_time(value.get("created_utc"),
                         "cap-boundary.frozen-inputs.created_utc")
    initial_created = parse_time(initial.get("created_utc"),
                                "cap-boundary.frozen-inputs.initial.created_utc")
    require(initial_created >= cap_plan_created
            and created > initial_created
            and created > preparation_created
            and created > candidate_frozen_created,
            "cap-boundary correction freeze ordering differs")
    require(value.get("status") == "frozen-before-build-and-capture",
            "cap-boundary frozen-inputs status differs")
    initial_files = initial.get("sha256")
    files = value.get("sha256")
    require(isinstance(initial_files, dict) and isinstance(files, dict)
            and initial_files and files,
            "cap-boundary frozen-input hash maps are missing")
    initial_aliases = {
        "cap-boundary/guard.rs": CAP_DIR / "guard.initial.rs",
        "cap-boundary/analyze.py": CAP_DIR / "analyze.initial.py",
    }
    for envelope, mapping, label in (
        (initial, initial_files, "cap-boundary initial frozen input"),
        (value, files, "cap-boundary frozen input"),
    ):
        for name, digest in mapping.items():
            safe_relative(name, f"{label} path")
            require(valid_digest(digest), f"{label} digest is malformed: {name}")
            target = HERE / name
            comparison_target = initial_aliases.get(name) if envelope is initial else None
            if comparison_target is not None:
                require(comparison_target.is_file() and not comparison_target.is_symlink()
                        and sha(comparison_target) == digest,
                        f"{label} retained initial snapshot differs: {name}")
            elif name == "cap-boundary/analyze.py" and digest == initial_files.get(name):
                # The corrected analyzer is written after the corrected cap
                # freeze.  Its immutable initial bytes remain bound above;
                # cap_analyzer_correction() binds the final replay bytes.
                require(target.is_file() and not target.is_symlink(),
                        f"{label} artifact is missing: {name}")
            else:
                require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                        f"{label} hash differs: {name}")
    required = {
        "cap-boundary/plan.json", "cap-boundary/cap_run.py",
        "cap-boundary/guard.rs", "cap-boundary/analyze.py",
        "cap-boundary/source-preparation.json", "run.py", "plan.json",
        "applied-candidate.patch", "baseline/source-manifest.json",
        "candidate/source-manifest.json",
    }
    require(required <= set(files),
            f"cap-boundary frozen inputs omit required files: {sorted(required - set(files))}")
    require(value.get("prior_frozen_inputs_sha256") == sha(initial_path),
            "cap-boundary correction does not bind initial frozen inputs")
    correction = value.get("correction")
    require(isinstance(correction, str)
            and "failed" in correction.lower()
            and "build" in correction.lower()
            and "no gates changed" in correction.lower(),
            "cap-boundary correction explanation is missing")
    require(initial_files.get("cap-boundary/guard.rs") == sha(CAP_DIR / "guard.initial.rs")
            and initial_files.get("cap-boundary/analyze.py") == sha(CAP_DIR / "analyze.initial.py"),
            "cap-boundary initial guard/analyzer snapshots are not bound")
    require(files["cap-boundary/guard.rs"] == sha(CAP_GUARD),
            "cap-boundary corrected guard hash differs")
    require(files["cap-boundary/guard.rs"] != initial_files["cap-boundary/guard.rs"],
            "cap-boundary correction did not change the retained guard source")
    analyzer_correction = cap_analyzer_correction(initial_files, files, created)
    return {"status": "pass", "created_utc": created.isoformat(),
            "initial_created_utc": initial_created.isoformat(),
            "sha256": sha(CAP_DIR / "frozen-inputs.json"),
            "prior_sha256": value["prior_frozen_inputs_sha256"],
            "analyzer_correction": analyzer_correction}, created


def cap_failed_stage(stage_name: str, main_manifest: dict[str, str],
                     main_stage: dict[str, Any], revision: dict[str, str],
                     cap_plan_sha: str, preparation_created: dt.datetime,
                     frozen_created: dt.datetime) -> dict[str, Any]:
    """Audit a failed cap source/build attempt without admitting it to ABBA."""

    stage = CAP_DIR / stage_name
    require(stage_name == "failed-baseline-1",
            f"unsupported failed cap source attempt: {stage_name}")
    require(stage.is_dir() and not stage.is_symlink(),
            f"cap-boundary/{stage_name} is not a regular directory")
    manifest = source_manifest(stage / "source-manifest.json")
    initial_guard = sha(CAP_DIR / "guard.initial.rs")
    expected = dict(main_manifest)
    expected["crates/litchi-xlsx/examples/perf_cap_boundary.rs"] = initial_guard
    require(manifest == expected,
            f"cap-boundary/{stage_name} source manifest differs from initial snapshot")
    patch = stage / "source.patch"
    main_patch = HERE / "baseline/source.patch"
    require(read_bytes(patch, relative(patch)) == read_bytes(main_patch, relative(main_patch)),
            f"cap-boundary/{stage_name} source patch differs")
    roots = [
        "crates/litchi-xlsx/src/cell_values/",
        "crates/litchi-xlsx/src/raw/worksheet/",
        "crates/litchi-xlsx/examples/",
    ]
    harness = harness_paths(plan_data())
    shared = {name for name in set(manifest) | set(revision)
              if test_only_path(name, plan_data())}
    expected_sources = {name for name in manifest
                        if path_in_roots(name, roots) or name in harness or name in shared}
    sources = need(stage / "sources", f"cap-boundary/{stage_name}/sources", directory=True)
    actual = {path.relative_to(sources).as_posix() for path in sources.rglob("*")
              if path.is_file() and not path.is_symlink()}
    require(actual == expected_sources,
            f"cap-boundary/{stage_name}/sources inventory differs")
    for path in sources.rglob("*"):
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
    for name in actual:
        require(sha(sources / name) == manifest[name],
                f"cap-boundary/{stage_name}/sources hash differs: {name}")
    require((sources / "crates/litchi-xlsx/examples/perf_cap_boundary.rs").read_bytes()
            == read_bytes(CAP_DIR / "guard.initial.rs", "cap-boundary/guard.initial.rs"),
            f"cap-boundary/{stage_name} initial guard source differs")
    receipt_path = stage / "build-cap-boundary.receipt.json"
    receipt = read_json(receipt_path, relative(receipt_path))
    start, end = interval(receipt, relative(receipt_path))
    require(start > preparation_created and end <= frozen_created,
            f"cap-boundary/{stage_name} failed receipt is outside correction freeze")
    require(receipt.get("exit_code") != 0 and receipt.get("binary_sha256") is None
            and receipt.get("source_manifest_sha256") == sha(stage / "source-manifest.json")
            and receipt.get("working_source_manifest_sha256") == sha(stage / "source-manifest.json")
            and receipt.get("script_sha256") == sha(RUN)
            and receipt.get("plan_sha256") == cap_plan_sha
            and "cap_driver_sha256" not in receipt
            and receipt.get("environment", {}).get("TMPDIR") == str(TARGET / "test-tmp")
            and receipt.get("command") == cap_build_command(),
            f"cap-boundary/{stage_name} failed build receipt binding differs")
    cap_artifacts(stage, "build-cap-boundary", receipt,
                  {"build-cap-boundary.stdout", "build-cap-boundary.stderr"})
    allowed = {"source-manifest.json", "source.patch", "build-cap-boundary.receipt.json",
               "build-cap-boundary.stdout", "build-cap-boundary.stderr"}
    for path in stage.iterdir():
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            require(path.name in allowed,
                    f"{relative(path)} is an unbound failed cap artifact")
    return {"status": "pass", "stage": stage_name,
            "manifest_sha256": sha(stage / "source-manifest.json"),
            "receipt_sha256": sha(receipt_path), "start": start, "end": end}


def cap_source_stage(stage_name: str, cap_plan: dict[str, Any],
                     main_manifest: dict[str, str], revision: dict[str, str],
                     main_stage: dict[str, Any], harness: set[str],
                     main_candidate_changed: set[str]) -> dict[str, Any]:
    """Validate one cap source snapshot, including the temporary example bytes."""

    stage = CAP_DIR / stage_name
    require(stage.is_dir() and not stage.is_symlink(),
            f"cap-boundary/{stage_name} is not a regular stage directory")
    manifest = source_manifest(stage / "source-manifest.json")
    enabler = cap_plan["temporary_enabler"]["source"]
    expected = dict(main_manifest)
    expected[enabler] = sha(CAP_GUARD)
    require(manifest == expected,
            f"cap-boundary/{stage_name} source manifest differs from main snapshot")
    patch = stage / "source.patch"
    main_patch = HERE / stage_name / "source.patch"
    require(read_bytes(patch, relative(patch)) == read_bytes(main_patch, relative(main_patch)),
            f"cap-boundary/{stage_name} source patch differs from main snapshot")
    roots = cap_plan["_roots"]
    shared = {
        name for name in set(manifest) | set(revision)
        if test_only_path(name, plan_data())
    }
    expected_sources = {
        name for name in manifest
        if path_in_roots(name, roots) or name in harness or name in shared
    }
    sources = need(stage / "sources", f"cap-boundary/{stage_name}/sources", directory=True)
    actual: set[str] = set()
    for path in sources.rglob("*"):
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            name = path.relative_to(sources).as_posix()
            safe_relative(name, f"cap-boundary/{stage_name}/sources path")
            actual.add(name)
    require(actual == expected_sources,
            f"cap-boundary/{stage_name}/sources inventory differs")
    for name in actual:
        require(sha(sources / name) == manifest[name],
                f"cap-boundary/{stage_name}/sources hash differs: {name}")
    require((sources / enabler).read_bytes() == read_bytes(CAP_GUARD, relative(CAP_GUARD)),
            f"cap-boundary/{stage_name} temporary enabler differs from retained guard source")
    if stage_name == "candidate":
        diff_path = need(stage / "source-diff.json", "cap-boundary/candidate/source-diff.json")
        diff = read_json(diff_path, relative(diff_path))
        require(isinstance(diff, dict)
                and diff.get("baseline_manifest_sha256") ==
                sha(CAP_DIR / "baseline/source-manifest.json")
                and diff.get("candidate_manifest_sha256") ==
                sha(stage / "source-manifest.json")
                and diff.get("candidate_source_roots") == roots,
                "cap-boundary candidate source-diff envelope differs")
        changes = diff.get("changed_files")
        require(isinstance(changes, dict) and set(changes) == main_candidate_changed,
                "cap-boundary candidate source-diff file set differs")
        baseline_cap = source_manifest(CAP_DIR / "baseline/source-manifest.json")
        for name in main_candidate_changed:
            item = changes[name]
            require(isinstance(item, dict)
                    and item.get("baseline_sha256") == baseline_cap.get(name)
                    and item.get("candidate_sha256") == manifest.get(name),
                    f"cap-boundary candidate source-diff hash differs: {name}")
    return {
        "manifest": manifest,
        "manifest_sha256": sha(stage / "source-manifest.json"),
        "enabler": enabler,
        "changed_from_main": sorted(changed_paths(manifest, main_manifest)),
    }


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


def cap_receipt_binding(stage_name: str, action: str, receipt: dict[str, Any],
                        manifest_sha: str, working_sha: str,
                        cap_plan: dict[str, Any], cap_plan_sha: str) -> tuple[dt.datetime, dt.datetime]:
    stage = CAP_DIR / stage_name
    path = stage / f"{action}.receipt.json"
    require(isinstance(receipt, dict), f"{relative(path)} is not an object")
    start, end = interval(receipt, relative(path))
    require(isinstance(receipt.get("exit_code"), int)
            and not isinstance(receipt.get("exit_code"), bool),
            f"{relative(path)} exit code is malformed")
    require(receipt.get("source_manifest_sha256") == manifest_sha
            and receipt.get("working_source_manifest_sha256") == working_sha,
            f"{relative(path)} source manifest binding differs")
    require(receipt.get("script_sha256") == sha(RUN)
            and receipt.get("plan_sha256") == cap_plan_sha
            and receipt.get("cap_driver_sha256") == sha(CAP_RUN),
            f"{relative(path)} driver or plan binding differs")
    environment = receipt.get("environment")
    require(isinstance(environment, dict)
            and environment.get("TMPDIR") == str(TARGET / "test-tmp"),
            f"{relative(path)} temporary directory binding differs")
    if "source_unchanged" in receipt:
        require(receipt["source_unchanged"] is True,
                f"{relative(path)} reports source mutation")
    return start, end


def cap_binary_identity(stage_name: str, manifest_sha: str, cap_plan: dict[str, Any],
                        cap_plan_sha: str) -> tuple[dict[str, Any], dict[str, Any]]:
    stage = CAP_DIR / stage_name
    identity_path = stage / "binary-cap-boundary.json"
    identity = read_json(identity_path, relative(identity_path))
    require(isinstance(identity, dict), f"{relative(identity_path)} is not an object")
    expected_path = CAP_TARGET / f"{stage_name}-cap-boundary"
    require(identity.get("path") == str(expected_path),
            f"{relative(identity_path)} binary path differs")
    require(valid_digest(identity.get("sha256")),
            f"{relative(identity_path)} digest is malformed")
    _positive_integer(identity.get("bytes"), f"{relative(identity_path)}.bytes")
    require(identity["bytes"] > 0
            and identity.get("source_manifest_sha256") == manifest_sha,
            f"{relative(identity_path)} identity differs")
    build_path = stage / "build-cap-boundary.receipt.json"
    build = read_json(build_path, relative(build_path))
    require(identity.get("build_receipt_sha256") == sha(build_path),
            f"{relative(identity_path)} build receipt differs")
    build_start, build_end = cap_receipt_binding(
        stage_name, "build-cap-boundary", build, manifest_sha, manifest_sha,
        cap_plan, cap_plan_sha,
    )
    require(build.get("exit_code") == 0 and build.get("binary_sha256") is None,
            f"{relative(build_path)} is not a successful cap build receipt")
    require(build.get("command") == cap_build_command(),
            f"{relative(build_path)} command differs")
    cap_artifacts(stage, "build-cap-boundary", build,
                  {"build-cap-boundary.stdout", "build-cap-boundary.stderr"})
    binary_path = Path(identity["path"])
    if binary_path.exists():
        require(binary_path.is_file() and not binary_path.is_symlink()
                and sha(binary_path) == identity["sha256"]
                and binary_path.stat().st_size == identity["bytes"],
                f"{relative(identity_path)} retained binary hash or size differs")
    else:
        require(CLEANUP.exists(),
                f"{relative(identity_path)} binary vanished before cleanup")
    return {
        "path": str(binary_path), "sha256": identity["sha256"],
        "bytes": identity["bytes"],
        "source_manifest_sha256": manifest_sha,
        "build_receipt_sha256": sha(build_path),
    }, {"name": "build-cap-boundary", "stage": stage_name,
        "start": build_start, "end": build_end, "exit_code": build["exit_code"]}


def cap_capture_record(stage_name: str, repeat: int, size: int, binary: dict[str, Any],
                       manifest_sha: str, working_sha: str, cap_plan: dict[str, Any],
                       cap_plan_sha: str) -> dict[str, Any]:
    stage = CAP_DIR / stage_name
    action = f"cap-native-r{repeat}-{size}"
    report = stage / f"{action}.json"
    fixture = stage / f"{action}.fixture.bin"
    rss = stage / f"{action}.rss.json"
    receipt_path = stage / f"{action}.receipt.json"
    for path in (report, fixture, rss, receipt_path,
                 stage / f"{action}.stdout", stage / f"{action}.stderr"):
        need(path, relative(path))
    receipt = read_json(receipt_path, relative(receipt_path))
    start, end = cap_receipt_binding(
        stage_name, action, receipt, manifest_sha, working_sha,
        cap_plan, cap_plan_sha,
    )
    require(receipt.get("exit_code") == 0
            and receipt.get("binary_sha256") == binary["sha256"],
            f"{relative(receipt_path)} is not a successful cap capture")
    command = [
        "taskset", "-c", str(cap_plan["cpu"]), "/usr/bin/time", "-f", CAP_TIME_FORMAT,
        "-o", str(rss), binary["path"], "--size", str(size), "--warmup", "10",
        "--samples", "100", "--json", str(report), "--fixture-out", str(fixture),
    ]
    require(receipt.get("command") == command,
            f"{relative(receipt_path)} command differs")
    cap_artifacts(stage, action, receipt,
                  {report.name, fixture.name, rss.name,
                   f"{action}.stdout", f"{action}.stderr"})
    return {
        "stage": stage_name, "name": action, "repeat": repeat, "size": size,
        "start": start, "end": end, "exit_code": receipt["exit_code"],
        "report_sha256": sha(report), "receipt_sha256": sha(receipt_path),
        "fixture_sha256": sha(fixture), "fixture_bytes": fixture.stat().st_size,
        "binary_sha256": binary["sha256"], "manifest_sha256": manifest_sha,
    }


def cap_stage_records(stage_name: str, cap_plan: dict[str, Any],
                      source_data: dict[str, Any], cap_plan_sha: str,
                      candidate_cap_manifest_sha: str | None) -> dict[str, Any]:
    manifest_sha = source_data["manifest_sha256"]
    manifest = source_data["manifest"]
    binary, build_record = cap_binary_identity(
        stage_name, manifest_sha, cap_plan, cap_plan_sha,
    )
    captures: list[dict[str, Any]] = []
    for repeat in CAP_REPEATS:
        sizes = CAP_SIZES if repeat == 1 else tuple(reversed(CAP_SIZES))
        working_sha = (candidate_cap_manifest_sha
                       if stage_name == "baseline" and repeat == 2
                       else manifest_sha)
        for size in sizes:
            captures.append(cap_capture_record(
                stage_name, repeat, size, binary, manifest_sha, working_sha,
                cap_plan, cap_plan_sha,
            ))
    expected_receipts = {"build-cap-boundary", *[record["name"] for record in captures]}
    actual_receipts = {
        path.name.removesuffix(".receipt.json")
        for path in (CAP_DIR / stage_name).glob("*.receipt.json")
        if path.is_file() and not path.is_symlink()
    }
    require(actual_receipts == expected_receipts,
            f"cap-boundary/{stage_name} receipt inventory differs")
    bound = {f"{name}.receipt.json" for name in expected_receipts}
    bound |= {name for record in captures for name in (
        f"{record['name']}.json", f"{record['name']}.fixture.bin",
        f"{record['name']}.rss.json", f"{record['name']}.stdout",
        f"{record['name']}.stderr",
    )}
    bound |= {"build-cap-boundary.stdout", "build-cap-boundary.stderr",
              "binary-cap-boundary.json", "source-manifest.json", "source.patch"}
    if stage_name == "candidate":
        bound.add("source-diff.json")
    for path in (CAP_DIR / stage_name).iterdir():
        if path.is_symlink():
            raise EvidenceError(f"{relative(path)} is a symlink")
        if path.is_file():
            require(path.name in bound,
                    f"{relative(path)} is an unbound cap stage artifact")
    return {"manifest": manifest, "manifest_sha256": manifest_sha,
            "binary": binary, "build": build_record, "captures": captures}


def cap_timeline(cap_stages: dict[str, dict[str, Any]],
                 preparation_created: dt.datetime,
                 cap_plan_created: dt.datetime,
                 cap_frozen_created: dt.datetime) -> dict[str, Any]:
    """Require the cap captures to remain serial and in the frozen ABBA order."""

    records: list[dict[str, Any]] = []
    for stage_data in cap_stages.values():
        records.append(stage_data["build"])
        records.extend(stage_data["captures"])
    validate_intervals(records)
    require(all(record["start"] > preparation_created
                and record["start"] > cap_plan_created
                and record["start"] > cap_frozen_created for record in records),
            "a cap receipt predates the cap source preparation or corrected freeze")
    groups: dict[tuple[str, int], list[dict[str, Any]]] = {}
    for stage_name, stage_data in cap_stages.items():
        for record in stage_data["captures"]:
            groups.setdefault((stage_name, record["repeat"]), []).append(record)
    expected = list(CAP_NATIVE_ORDER)
    starts: list[tuple[str, int, dt.datetime]] = []
    for stage_name, repeat in expected:
        rows = groups.get((stage_name, repeat), [])
        require(len(rows) == len(CAP_SIZES),
                f"cap ABBA {stage_name} repeat {repeat} is incomplete")
        ordered = sorted(rows, key=lambda record: record["start"])
        expected_sizes = list(CAP_SIZES if repeat == 1 else reversed(CAP_SIZES))
        require([record["size"] for record in ordered] == expected_sizes,
                f"cap {stage_name} repeat {repeat} size order differs")
        starts.append((stage_name, repeat, ordered[0]["start"]))
    require([(stage, repeat) for stage, repeat, _ in sorted(starts, key=lambda row: row[2])]
            == expected,
            "cap capture order is not baseline/candidate/candidate/baseline")
    by_stage_build = {stage: data["build"] for stage, data in cap_stages.items()}
    b1 = groups[("baseline", 1)]
    c1 = groups[("candidate", 1)]
    c2 = groups[("candidate", 2)]
    b2 = groups[("baseline", 2)]
    require(by_stage_build["baseline"]["end"] <= min(row["start"] for row in b1),
            "baseline cap build does not precede baseline repeat 1")
    require(max(row["end"] for row in b1) <= by_stage_build["candidate"]["start"],
            "candidate cap build overlaps or precedes baseline repeat 1")
    require(by_stage_build["candidate"]["end"] <= min(row["start"] for row in c1),
            "candidate cap build does not precede candidate repeat 1")
    require(max(row["end"] for row in c1) <= min(row["start"] for row in c2)
            and max(row["end"] for row in c2) <= min(row["start"] for row in b2),
            "cap repeat groups overlap or violate serial ABBA")
    return {
        "status": "pass", "receipts": len(records),
        "order": [{"stage": stage, "repeat": repeat,
                   "start_utc": start.isoformat()}
                  for stage, repeat, start in sorted(starts, key=lambda row: row[2])],
    }


def replay_cap_analysis(cap_plan_sha: str) -> dict[str, Any]:
    """Replay the independent cap analyzer and retain its admission result."""

    need(CAP_ANALYSIS, "cap-boundary/cap-analysis.json")
    expected = read_json(CAP_ANALYSIS, "cap-boundary/cap-analysis.json")
    module = load_module(CAP_ANALYZER, "litchi_0543_cap_analysis", register=False)
    method = getattr(module, "analyze", None)
    require(callable(method), "cap-boundary/analyze.py has no callable analyze()")
    try:
        result = method()
    except Exception as error:  # noqa: BLE001 - preserve analyzer evidence context
        raise EvidenceError(f"cap-boundary analyzer replay failed: {error}") from error
    # The analyzer keeps ``fixture_identities_by_size`` keyed by integer
    # sizes in memory, while its retained JSON necessarily has string object
    # keys after serialization.  Compare the JSON round-trip representation
    # so this custody check tests the bytes that can actually be retained,
    # without permitting any semantic difference in the report.
    try:
        replayed_json = json.loads(json.dumps(
            result, ensure_ascii=False, sort_keys=True, allow_nan=False,
        ))
    except (TypeError, ValueError) as error:
        raise EvidenceError(
            "cap-boundary analyzer replay is not JSON serializable"
        ) from error
    require(replayed_json == expected,
            "cap-boundary/cap-analysis.json differs from deterministic analyzer replay")
    require(isinstance(result, dict)
            and result.get("schema") == "litchi.xlsx.cap-boundary-analysis.v1"
            and result.get("status") in {"pass", "reject"}
            and result.get("plan_sha256") == cap_plan_sha
            and result.get("driver_sha256") == sha(CAP_RUN),
            "cap-boundary analysis envelope differs")
    stages = result.get("stages")
    comparison = result.get("comparison")
    require(isinstance(stages, dict) and set(stages) == {"baseline", "candidate"}
            and all(isinstance(item, dict) and item.get("status") == "pass"
                    for item in stages.values()),
            "cap-boundary analysis stages are incomplete")
    require(isinstance(comparison, dict)
            and comparison.get("status") == result["status"]
            and isinstance(comparison.get("admission_passed"), bool)
            and ((result["status"] == "pass") is comparison["admission_passed"]),
            "cap-boundary comparison admission result differs")
    rows = comparison.get("rows")
    require(isinstance(rows, list) and len(rows) == len(CAP_SIZES) * len(CAP_REPEATS),
            "cap-boundary comparison row cardinality differs")
    keys = {(row.get("repeat"), row.get("size")) for row in rows
            if isinstance(row, dict)}
    require(keys == {(repeat, size) for repeat in CAP_REPEATS for size in CAP_SIZES},
            "cap-boundary comparison matrix differs")
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("metrics"), dict)
                and set(row["metrics"]) == {"p50", "mean"},
                "cap-boundary comparison metrics differ")
        for metric in row["metrics"].values():
            require(isinstance(metric, dict) and isinstance(metric.get("passed"), bool),
                    "cap-boundary comparison metric is malformed")
    adverse = comparison.get("adverse")
    drift = comparison.get("drift")
    require(isinstance(adverse, list) and isinstance(drift, list),
            "cap-boundary adverse/drift vectors are missing")
    return {
        "status": "pass", "report_status": result["status"],
        "report_path": relative(CAP_ANALYSIS),
        "report_sha256": sha(CAP_ANALYSIS), "plan_sha256": cap_plan_sha,
        "admission_passed": comparison["admission_passed"],
        "rows": rows, "adverse": adverse, "drift": drift,
    }


def validate_cap_boundary(plan: dict[str, Any], baseline: dict[str, str],
                          candidate: dict[str, str], candidate_changed: set[str],
                          revision: dict[str, str], stages: dict[str, dict[str, Any]],
                          harness: set[str], candidate_frozen_created: dt.datetime,
                          decision: dict[str, Any]) -> dict[str, Any]:
    """Validate the supplemental valid-path cap lane before final retention."""

    cap_plan, cap_plan_created = cap_plan_data(
        plan, candidate_changed, candidate_frozen_created,
    )
    preparation, preparation_created = cap_source_preparation(
        plan, baseline, candidate, candidate_changed,
        candidate_frozen_created, cap_plan_created,
    )
    cap_freeze, cap_frozen_created = cap_frozen_inputs(
        cap_plan_created, candidate_frozen_created, preparation_created,
    )
    failed_cap = cap_failed_stage(
        "failed-baseline-1", baseline, stages["baseline"], revision,
        sha(CAP_PLAN), preparation_created, cap_frozen_created,
    )
    cap_source: dict[str, dict[str, Any]] = {}
    for stage_name, main_manifest in (("baseline", baseline), ("candidate", candidate)):
        cap_source[stage_name] = cap_source_stage(
            stage_name, cap_plan, main_manifest, revision,
            stages[stage_name], harness, candidate_changed,
        )
    cap_plan_sha = sha(CAP_PLAN)
    cap_stage_data = {
        stage_name: cap_stage_records(
            stage_name, cap_plan, cap_source[stage_name], cap_plan_sha,
            cap_source["candidate"]["manifest_sha256"],
        )
        for stage_name in ("baseline", "candidate")
    }
    timeline = cap_timeline(
        cap_stage_data, preparation_created, cap_plan_created, cap_frozen_created,
    )
    analysis = replay_cap_analysis(cap_plan_sha)
    if decision["decision"] == "accepted":
        require(analysis["report_status"] == "pass" and analysis["admission_passed"] is True,
                "accepted decision does not satisfy supplemental cap-boundary gate")
    return {
        "status": "pass", "plan_sha256": cap_plan_sha,
        "plan_created_utc": cap_plan_created.isoformat(),
        "source_preparation": preparation,
        "frozen_inputs": cap_freeze,
        "failed_attempts": [{
            "status": failed_cap["status"],
            "stage": failed_cap["stage"],
            "manifest_sha256": failed_cap["manifest_sha256"],
            "receipt_sha256": failed_cap["receipt_sha256"],
            "start_utc": failed_cap["start"].isoformat(),
            "end_utc": failed_cap["end"].isoformat(),
        }],
        "stages": {
            stage_name: {
                "manifest_sha256": cap_stage_data[stage_name]["manifest_sha256"],
                "binary": cap_stage_data[stage_name]["binary"],
                "receipts": len(cap_stage_data[stage_name]["captures"]) + 1,
            }
            for stage_name in ("baseline", "candidate")
        },
        "timeline": timeline,
        "analysis": analysis,
        "_records": [failed_cap]
        + [cap_stage_data[stage_name]["build"]
                     for stage_name in ("baseline", "candidate")]
        + [record for stage_name in ("baseline", "candidate")
           for record in cap_stage_data[stage_name]["captures"]],
        "binaries": {
            stage_name: {"cap-boundary": cap_stage_data[stage_name]["binary"]}
            for stage_name in ("baseline", "candidate")
        },
    }


def adr_manifest() -> dict[str, str]:
    value = read_json(ADR, "adr-manifest.json")
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
    require("abba" in lowered and "allocation" in lowered
            and ("cleanup" in lowered
                 or ("removed" in lowered and "sealing" in lowered)),
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


def replay_differential_patch(plan: dict[str, Any], revision: dict[str, str],
                              target_manifest: dict[str, str], stage_name: str,
                              patch_path: Path, label: str) -> dict[str, Any]:
    """Replay one frozen public differential-test patch into a stage.

    ``run.py`` deliberately keeps the new integration test outside the
    candidate implementation roots.  Its source bytes therefore need an
    independent binding: a normal ``git diff`` omits an untracked new test,
    while the retained differential patch records both the owner edit and the
    new module.  This check proves that patch and stage snapshot describe the
    same test-only change without changing the checkout.  The initial patch
    is replayed against the retained failed attempt; the corrected final
    patch is replayed against the fresh baseline.
    """

    data = read_bytes(patch_path, label)
    paths = patch_paths(data, label)
    require(paths, f"{label} is empty")
    require(all(test_only_path(name, plan) for name in paths),
            f"{label} contains a non-test path")
    changed = changed_paths(target_manifest, revision)
    require(changed and all(test_only_path(name, plan) for name in changed),
            f"{stage_name} source difference is not exactly test-only")
    tracked_changed = changed & set(revision)
    require(paths == changed,
            f"{label} paths differ from the frozen test inventory")
    with tempfile.TemporaryDirectory(prefix="litchi-0543-verify-differential-",
                                     dir="/dev/shm") as temporary:
        root = Path(temporary)
        for name in tracked_changed:
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes(git_show(plan["revision"], name))
        checked = subprocess.run(
            ["git", "apply", "--check", "--whitespace=nowarn", "-"], cwd=root,
            input=data, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(checked.returncode == 0,
                f"{label} does not apply: "
                f"{checked.stderr.decode(errors='replace')[-1200:]}")
        applied = subprocess.run(
            ["git", "apply", "--whitespace=nowarn", "-"], cwd=root,
            input=data, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(applied.returncode == 0,
                f"{label} could not be replayed: "
                f"{applied.stderr.decode(errors='replace')[-1200:]}")
        for name in changed:
            produced = root / name
            retained = HERE / stage_name / "sources" / name
            require(produced.is_file() and retained.is_file()
                    and produced.read_bytes() == retained.read_bytes(),
                    f"{label} replay differs for {name}")
    return {"status": "pass", "paths": sorted(paths),
            "patch_sha256": sha(patch_path), "stage": stage_name,
            "test_paths": sorted(changed)}


def validate_differential_tests(plan: dict[str, Any], revision: dict[str, str],
                                baseline: dict[str, str]) -> dict[str, Any]:
    """Replay the corrected public differential-test patch into the baseline."""

    patch_path = HERE / "differential-tests-final.patch"
    return replay_differential_patch(
        plan, revision, baseline, "baseline", patch_path,
        "differential-tests-final.patch",
    )


def replay_candidate_patch(stage: Path, baseline: dict[str, str], candidate: dict[str, str],
                           changed: set[str], label: str) -> None:
    """Replay the definitive baseline-to-candidate patch in an isolated tree."""
    patch_path = HERE / label
    data = read_bytes(patch_path, label)
    paths = patch_paths(data, label)
    require(paths == changed, f"{label} paths differ from candidate source diff")
    roots = plan_data()["_roots"]
    require(all(path_in_roots(name, roots) for name in paths),
            f"{label} contains a path outside candidate roots")
    with tempfile.TemporaryDirectory(prefix="litchi-0543-verify-candidate-", dir="/dev/shm") as temporary:
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
    the measured pilot before the next proposal is written.  If either part
    is present, however, both the patch provenance and its explicit prose
    label are required.  No timing or admission result is inferred from it.
    """
    patch_path = HERE / "next-candidate.patch"
    note_path = HERE / "next-priority.md"
    if not patch_path.exists():
        # 0543 deliberately does not require an unmeasured follow-up patch.
        # A priority note may still be retained for the next OLE2/OOXML
        # campaign, but its presence must not turn the optional patch into a
        # required input for this measured decision.
        if note_path.exists():
            note = read_text(note_path, "next-priority.md").lower()
            require("unmeasured" in note and "0543" in note
                    and "ooxml" in note and "odf" in note,
                    "next-priority.md omits the unmeasured 0543 format priority")
            return {
                "status": "not-present", "admission": "unmeasured-only",
                "note": relative(note_path), "note_sha256": sha(note_path),
            }
        return {"status": "not-present", "admission": "unmeasured-only"}
    if not note_path.exists():
        raise EvidenceError("next-candidate.patch is retained without next-priority.md")
    require(patch_path.is_file() and not patch_path.is_symlink(),
            "next-candidate.patch is missing or not a regular file")
    require(note_path.is_file() and not note_path.is_symlink(),
            "next-priority.md is missing or not a regular file")
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
    with tempfile.TemporaryDirectory(prefix="litchi-0543-verify-next-", dir="/dev/shm") as temporary:
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
    note = read_text(note_path, "next-priority.md").lower()
    require("unmeasured proposal" in note,
            "next-priority.md lacks the explicit unmeasured proposal label")
    require("0543" in note and "ooxml" in note and "odf" in note,
            "next-priority.md omits pilot or format priority scope")
    return {
        "status": "pass", "admission": "unmeasured-only",
        "patch": relative(patch_path), "patch_sha256": sha(patch_path),
        "patch_paths": sorted(paths), "label": "unmeasured proposal",
        "provenance": "replayed against frozen candidate sources",
        "note": relative(note_path), "note_sha256": sha(note_path),
    }


def validate_cap_next_provenance(plan: dict[str, Any], candidate: dict[str, str],
                                 candidate_frozen_created: dt.datetime,
                                 cap_boundary: dict[str, Any]) -> dict[str, Any]:
    """Replay the optional cap-boundary follow-up against frozen candidate bytes.

    The cap follow-up is a proposal for a later campaign.  Its patch and
    provenance are useful custody evidence, but no generated source is
    admitted to the measured 0543 candidate and no performance claim is made
    here.  If any provenance artifact is retained, validate the complete set.
    """

    provenance_path = CAP_DIR / "next-provenance.json"
    patch_path = CAP_DIR / "next-candidate.patch"
    design_path = CAP_DIR / "next-design.md"
    review_path = CAP_DIR / "next-review.md"
    present = [path.exists() for path in
               (provenance_path, patch_path, design_path, review_path)]
    if not any(present):
        return {"status": "not-present", "admission": "unmeasured-only"}
    require(all(present),
            "cap-boundary next proposal provenance is only partially retained")
    value = read_json(provenance_path, "cap-boundary/next-provenance.json")
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.xlsx.cap-boundary-provenance.v1"
            and value.get("status") == "patch-ready-unmeasured",
            "cap-boundary next provenance envelope differs")
    created = parse_time(value.get("created_utc"),
                         "cap-boundary.next-provenance.created_utc")
    cap_freeze_created = parse_time(
        cap_boundary["frozen_inputs"]["created_utc"],
        "cap-boundary.frozen-inputs.created_utc",
    )
    require(created > cap_freeze_created and created > candidate_frozen_created,
            "cap-boundary next proposal predates its frozen inputs")
    review_disposition = value.get("review_disposition")
    require(isinstance(review_disposition, str)
            and "static" in review_disposition.lower()
            and "conditional" in review_disposition.lower()
            and "fresh performance" in review_disposition.lower(),
            "cap-boundary next proposal review disposition is missing")

    cap_plan = read_json(CAP_PLAN, "cap-boundary/plan.json")
    candidate_files = cap_plan.get("candidate_files") if isinstance(cap_plan, dict) else None
    require(isinstance(candidate_files, list) and candidate_files
            and len(set(candidate_files)) == len(candidate_files),
            "cap-boundary next proposal candidate file inventory is missing")
    expected_files = set(candidate_files)
    roots = cap_plan.get("candidate_source_roots")
    require(isinstance(roots, list) and roots,
            "cap-boundary next proposal source roots are missing")
    base = value.get("base")
    after = value.get("after")
    require(isinstance(base, dict) and isinstance(after, dict),
            "cap-boundary next proposal source envelopes are missing")
    require(base.get("label") == "0543-frozen-five-file-candidate"
            and after.get("label") == "0543-cap-boundary-next-candidate"
            and base.get("source_root") == str(TARGET / "candidate-src")
            and after.get("source_root") == str(TARGET / "next-candidate-src"),
            "cap-boundary next proposal source labels differ")

    def source_hashes(envelope: dict[str, Any], label: str) -> dict[str, str]:
        files = envelope.get("files_sha256")
        require(isinstance(files, dict) and set(files) == expected_files,
                f"{label} source inventory differs")
        for name, digest in files.items():
            safe_relative(name, f"{label} source path")
            require(path_in_roots(name, roots) and valid_digest(digest),
                    f"{label} source entry is malformed: {name}")
        return files

    base_files = source_hashes(base, "cap-boundary next base")
    after_files = source_hashes(after, "cap-boundary next result")
    candidate_sources = HERE / "candidate/sources"
    for name, digest in base_files.items():
        path = candidate_sources / name
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"cap-boundary next base does not match frozen candidate: {name}")

    # Check external source snapshots while they remain available.  After
    # campaign cleanup the owned target may be absent; the frozen candidate
    # snapshot and this isolated patch replay remain sufficient custody.
    for envelope, files, label in ((base, base_files, "cap-boundary next base"),
                                   (after, after_files, "cap-boundary next result")):
        root = Path(envelope["source_root"])
        if root.exists():
            require(root.is_dir() and not root.is_symlink(),
                    f"{label} source root is not a regular directory")
            actual: set[str] = set()
            for path in root.rglob("*"):
                require(not path.is_symlink(), f"{label} source snapshot has a symlink")
                if path.is_file():
                    name = path.relative_to(root).as_posix()
                    safe_relative(name, f"{label} source path")
                    actual.add(name)
            require(actual == expected_files,
                    f"{label} source snapshot inventory differs")
            for name, digest in files.items():
                require(sha(root / name) == digest,
                        f"{label} source snapshot hash differs: {name}")
        else:
            require(CLEANUP.exists(),
                    f"{label} source root vanished before cleanup")

    artifacts = value.get("artifacts")
    require(isinstance(artifacts, dict)
            and set(artifacts) == {"patch", "design", "independent_review",
                                   "frozen_0543_patch"},
            "cap-boundary next proposal artifact inventory differs")
    expected_artifacts = {
        "patch": ("docs/performance/results/change-0543/cap-boundary/next-candidate.patch",
                   patch_path),
        "design": ("docs/performance/results/change-0543/cap-boundary/next-design.md",
                    design_path),
        "independent_review": (
            "docs/performance/results/change-0543/cap-boundary/next-review.md", review_path,
        ),
        "frozen_0543_patch": ("docs/performance/results/change-0543/candidate.patch",
                               HERE / "candidate.patch"),
    }
    for name, (expected_path, target) in expected_artifacts.items():
        item = artifacts[name]
        require(isinstance(item, dict) and item.get("path") == expected_path
                and valid_digest(item.get("sha256"))
                and item.get("sha256") == sha(target),
                f"cap-boundary next proposal {name} artifact binding differs")
        if name != "frozen_0543_patch":
            require(item.get("bytes") == target.stat().st_size,
                    f"cap-boundary next proposal {name} byte count differs")
    require(sha(HERE / "candidate.patch") == sha(HERE / "applied-candidate.patch"),
            "cap-boundary next proposal does not bind identical definitive candidate patches")

    patch = read_bytes(patch_path, "cap-boundary/next-candidate.patch")
    paths = patch_paths(patch, "cap-boundary/next-candidate.patch")
    changed = value.get("changed_files")
    unchanged = value.get("unchanged_files")
    require(isinstance(changed, list) and isinstance(unchanged, list)
            and len(set(changed)) == len(changed)
            and len(set(unchanged)) == len(unchanged)
            and set(changed) | set(unchanged) == expected_files
            and set(changed) & set(unchanged) == set()
            and paths == set(changed)
            and set(changed) == expected_files - set(unchanged)
            and all(path_in_roots(name, roots) for name in paths),
            "cap-boundary next proposal changed-file inventory differs")
    with tempfile.TemporaryDirectory(prefix="litchi-0543-verify-cap-next-",
                                     dir="/dev/shm") as temporary:
        root = Path(temporary)
        for name in expected_files:
            target = root / name
            target.parent.mkdir(parents=True, exist_ok=True)
            target.write_bytes((candidate_sources / name).read_bytes())
        checked = subprocess.run(
            ["git", "apply", "--check", "--whitespace=nowarn", "-"], cwd=root,
            input=patch, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(checked.returncode == 0,
                "cap-boundary next proposal patch does not apply to frozen candidate: "
                f"{checked.stderr.decode(errors='replace')[-1200:]}")
        applied = subprocess.run(
            ["git", "apply", "--whitespace=nowarn", "-"], cwd=root,
            input=patch, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False,
        )
        require(applied.returncode == 0,
                "cap-boundary next proposal patch replay failed: "
                f"{applied.stderr.decode(errors='replace')[-1200:]}")
        for name, digest in after_files.items():
            produced = root / name
            require(produced.is_file() and sha(produced) == digest,
                    f"cap-boundary next proposal replay hash differs: {name}")

    scope = value.get("scope")
    require(isinstance(scope, dict)
            and scope.get("candidate_file_count") == len(expected_files)
            and scope.get("candidate_patch_changed_file_count") == len(changed)
            and scope.get("private_cap_tests_retained") is True
            and scope.get("public_integration_tests_included") is False
            and scope.get("ole2_ooxml_priority") is True
            and scope.get("odf_deferred") is True,
            "cap-boundary next proposal scope differs")
    safety = value.get("safety_invariants")
    expected_safety = {
        "preflight_before_reader_and_raw_parser": True,
        "runtime_event_cap_retained": True,
        "authoritative_validation_fallback_retained": True,
        "source_payload_ownership_retained": True,
        "mce_x14ac_utf8_and_8mib_fences_retained": True,
        "post_eof_result_and_x14ac_retry_boundary_retained": True,
        "higher_ranked_observer_lifetime_retained": True,
        "public_api_dependency_and_cargo_changes": False,
        "unsafe_code_added": False,
        "performance_claim_made": False,
    }
    require(safety == expected_safety,
            "cap-boundary next proposal safety invariants differ")
    execution = value.get("execution")
    require(execution == {
        "builds_run_by_this_task": False,
        "tests_run_by_this_task": False,
        "commits_created_by_this_task": False,
        "live_source_edits_by_this_task": False,
        "fresh_measurement_required": True,
    }, "cap-boundary next proposal execution claims differ")
    design = read_text(design_path, "cap-boundary/next-design.md").lower()
    review = read_text(review_path, "cap-boundary/next-review.md").lower()
    require("unmeasured" in design and "cap" in design and "preflight" in design
            and "fresh" in design and "ole2" in design and "odf" in design,
            "cap-boundary next design omits unmeasured follow-up controls")
    require("static" in review and "sound" in review and "conditional" in review
            and "edge" in review and "fresh" in review and "ole2" in review
            and "odf" in review,
            "cap-boundary next review omits conditional follow-up controls")
    return {
        "status": "pass", "admission": "unmeasured-only",
        "created_utc": created.isoformat(),
        "patch": relative(patch_path), "patch_sha256": sha(patch_path),
        "patch_paths": sorted(paths), "base_sha256":
        {name: digest for name, digest in sorted(base_files.items())},
        "after_sha256":
        {name: digest for name, digest in sorted(after_files.items())},
        "provenance": relative(provenance_path),
        "provenance_sha256": sha(provenance_path),
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
        with tempfile.TemporaryDirectory(prefix="litchi-0543-verify-", dir="/dev/shm") as temporary:
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


def _stage_dirs() -> list[Path]:
    """Return top-level source-attempt directories, rejecting symlinks."""

    result: list[Path] = []
    for path in sorted(HERE.iterdir(), key=lambda item: item.name):
        if not path.is_dir():
            continue
        if path.is_symlink():
            raise EvidenceError(f"{relative(path)} is a symlink")
        if (path / "source-manifest.json").exists():
            safe_relative(path.name, "stage name")
            require("/" not in path.name, "stage name is nested")
            result.append(path)
    return result


def stage_names() -> list[str]:
    """Return the two measured source stages in their fixed semantic order.

    Historical failed attempts live beside these stages.  They are audited
    separately and must never be admitted to the ABBA matrix or final-source
    decision merely because they happen to contain a source manifest.
    """

    available = {path.name for path in _stage_dirs()}
    if "baseline" not in available:
        raise Pending("baseline source attempt has not been frozen")
    if "candidate" not in available:
        raise Pending("candidate source attempt has not been frozen")
    unexpected = available - {"baseline", "candidate"} - {
        path.name for path in _stage_dirs() if path.name.startswith("failed-")
    }
    require(not unexpected, f"unrecognized source-attempt directories: {sorted(unexpected)}")
    return ["baseline", "candidate"]


FAILED_STAGE_RE = re.compile(r"failed-(baseline|candidate)-([1-9][0-9]*)\Z")


def failed_stage_names() -> list[str]:
    """Return retained failed-attempt directories in deterministic order."""

    result: list[str] = []
    for path in _stage_dirs():
        match = FAILED_STAGE_RE.fullmatch(path.name)
        if path.name.startswith("failed-"):
            require(match is not None,
                    f"failed source-attempt directory has an invalid name: {path.name}")
            result.append(path.name)
    result.sort(key=lambda name: (
        0 if name.startswith("failed-baseline-") else 1,
        int(name.rsplit("-", 1)[1]),
    ))
    return result


def validate_stage_sources(stage_name: str, plan: dict[str, Any], revision: dict[str, str],
                            roots: list[str], harness: set[str], baseline: dict[str, str] | None,
                            *, final_candidate: bool = False) -> dict[str, Any]:
    stage = HERE / stage_name
    require(stage.is_dir() and not stage.is_symlink(),
            f"{stage_name} is not a regular stage directory")
    snapshot = source_manifest(stage / "source-manifest.json")
    baseline_test_changes = (
        {name for name in changed_paths(baseline, revision)
         if test_only_path(name, plan)}
        if baseline is not None else set()
    )
    if stage_name == "baseline":
        changed_from_revision = changed_paths(snapshot, revision)
        require(changed_from_revision,
                "baseline source difference is empty; differential tests were not frozen")
        require(all(test_only_path(name, plan) for name in changed_from_revision),
                "baseline source patch contains a non-test change")
        allowed = changed_from_revision
        changed = validate_patch(stage, snapshot, revision, allowed, plan["revision"])
    else:
        require(baseline is not None, "candidate validation has no baseline manifest")
        changed_from_baseline = changed_paths(snapshot, baseline)
        require(changed_from_baseline,
                f"{stage_name} candidate source diff is empty")
        require(all(path_in_roots(name, roots) for name in changed_from_baseline),
                f"{stage_name} changes outside candidate source roots: "
                f"{sorted(name for name in changed_from_baseline if not path_in_roots(name, roots))}")
        # A candidate snapshot still includes the common guard-enabler edits;
        # only its production diff is permitted against the baseline.  The
        # baseline's shared integration-test patch is also present when the
        # candidate source.patch is replayed directly from the Git revision.
        allowed = baseline_test_changes | {
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
        # 0543 freezes one definitive baseline-to-candidate patch.  It may be
        # byte-identical to applied-candidate.patch, but both names remain
        # bound so a future draft cannot silently replace the replay target.
        draft = need(HERE / "candidate.patch", "candidate.patch")
        draft_paths = patch_paths(read_bytes(draft, "candidate.patch"), "candidate.patch")
        require(draft_paths == changed_from_baseline,
                "candidate.patch paths differ from the final candidate source diff")
        require(all(path_in_roots(name, roots) for name in draft_paths),
                "candidate.patch contains a path outside candidate roots")
        if "candidate_patch_sha256" in plan:
            require(sha(draft) == plan["candidate_patch_sha256"],
                    "candidate.patch hash differs from plan")
        need(HERE / "applied-candidate.patch", "applied-candidate.patch")
        # Both retained names are definitive replay inputs for 0543.  The
        # applied patch may be byte-identical to candidate.patch, but either
        # spelling must independently produce the frozen candidate bytes.
        replay_candidate_patch(stage, baseline, snapshot, changed_from_baseline,
                               "candidate.patch")
        replay_candidate_patch(stage, baseline, snapshot, changed_from_baseline,
                               "applied-candidate.patch")
        if final_candidate:
            planned = plan.get("candidate_files", [])
            require(isinstance(planned, list) and planned,
                    "final candidate has no planned candidate_files")
            require(set(planned) == changed_from_baseline,
                    "final candidate file inventory differs from plan")
    sources = need(stage / "sources", f"{stage_name}/sources", directory=True)
    actual: set[str] = set()
    for path in sources.rglob("*"):
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            name = path.relative_to(sources).as_posix()
            safe_relative(name, f"{stage_name}/sources path")
            actual.add(name)
    shared_tests = {
        name for name in set(snapshot) | set(revision)
        if test_only_path(name, plan)
    }
    expected = {
        name for name in snapshot
        if path_in_roots(name, roots) or name in harness or name in shared_tests
    }
    require(actual == expected, f"{stage_name}/sources inventory differs")
    for name in actual:
        require(sha(sources / name) == snapshot[name],
                f"{stage_name}/sources hash differs: {name}")
    # The committed guard harness must be present in both source attempts,
    # with the same bytes.  This catches a source-boundary change hidden
    # outside the candidate roots.
    for name in harness:
        require(snapshot.get(name) is not None,
                f"{stage_name} omits bound harness source: {name}")
    return {
        "manifest": snapshot,
        "manifest_sha256": sha(stage / "source-manifest.json"),
        "changed_from_revision": sorted(changed_paths(snapshot, revision)),
        "changed_from_baseline": sorted(changed_paths(snapshot, baseline or snapshot)),
        "shared_test_paths": sorted(shared_tests),
    }


def validate_failed_stage_sources(stage_name: str, plan: dict[str, Any],
                                  revision: dict[str, str], roots: list[str],
                                  harness: set[str],
                                  baseline: dict[str, str] | None) -> dict[str, Any]:
    """Validate a retained source snapshot from a failed campaign attempt.

    Failed attempts are historical custody records.  A failed baseline is
    expected to contain only the frozen public test patch; a failed candidate
    may additionally contain a production candidate-root patch.  Neither is
    eligible for ABBA or final-source selection, but both source manifests,
    patches, and receipt bindings remain auditable.
    """

    match = FAILED_STAGE_RE.fullmatch(stage_name)
    require(match is not None, f"invalid failed stage name: {stage_name}")
    stage = HERE / stage_name
    snapshot = source_manifest(stage / "source-manifest.json")
    kind = match.group(1)
    changed_from_revision = changed_paths(snapshot, revision)
    require(changed_from_revision,
            f"{stage_name} source difference is empty")
    if kind == "baseline":
        require(all(test_only_path(name, plan) for name in changed_from_revision),
                f"{stage_name} source difference contains a non-test path")
        allowed = changed_from_revision
    else:
        require(baseline is not None,
                f"{stage_name} candidate source validation has no baseline manifest")
        changed_from_baseline = changed_paths(snapshot, baseline)
        require(changed_from_baseline,
                f"{stage_name} candidate source difference is empty")
        require(all(path_in_roots(name, roots) for name in changed_from_baseline),
                f"{stage_name} changes outside candidate source roots: "
                f"{sorted(name for name in changed_from_baseline if not path_in_roots(name, roots))}")
        allowed = {
            name for name in set(snapshot) | set(revision)
            if path_in_roots(name, roots)
        }
        allowed |= {
            name for name in set(snapshot) | set(revision)
            if test_only_path(name, plan)
        }
    changed = validate_patch(stage, snapshot, revision, allowed, plan["revision"])

    sources = need(stage / "sources", f"{stage_name}/sources", directory=True)
    actual: set[str] = set()
    for path in sources.rglob("*"):
        require(not path.is_symlink(), f"{relative(path)} is a symlink")
        if path.is_file():
            name = path.relative_to(sources).as_posix()
            safe_relative(name, f"{stage_name}/sources path")
            actual.add(name)
    shared_tests = {
        name for name in set(snapshot) | set(revision)
        if test_only_path(name, plan)
    }
    expected = {
        name for name in snapshot
        if path_in_roots(name, roots) or name in harness or name in shared_tests
    }
    require(actual == expected, f"{stage_name}/sources inventory differs")
    for name in actual:
        require(sha(sources / name) == snapshot[name],
                f"{stage_name}/sources hash differs: {name}")
    for name in harness:
        require(snapshot.get(name) is not None,
                f"{stage_name} omits bound harness source: {name}")
        require(snapshot[name] == sha_bytes(git_show(plan["revision"], name)),
                f"{stage_name} guard harness differs from committed revision: {name}")

    differential: dict[str, Any] | None = None
    if kind == "baseline":
        # The initial patch is deliberately retained as the failed-attempt
        # receipt.  It is replayed against this exact source snapshot; the
        # corrected final patch is checked independently against fresh
        # baseline below.
        differential = replay_differential_patch(
            plan, revision, snapshot, stage_name,
            HERE / "differential-tests.patch", "differential-tests.patch",
        )
    return {
        "manifest": snapshot,
        "manifest_sha256": sha(stage / "source-manifest.json"),
        "changed_from_revision": sorted(changed_from_revision),
        "changed_from_baseline": sorted(
            changed_paths(snapshot, baseline or snapshot)
        ),
        "shared_test_paths": sorted(shared_tests),
        "kind": kind,
        "differential_tests": differential,
    }


def baseline_freeze_data(plan: dict[str, Any], frozen_created: dt.datetime,
                         baseline: dict[str, str], failed_name: str,
                         failed_records: list[dict[str, Any]]) -> dict[str, Any]:
    """Validate the correction freeze that precedes the fresh baseline."""

    value = read_json(BASELINE_FROZEN, "baseline-frozen-inputs.json")
    require(isinstance(value, dict), "baseline frozen-inputs envelope is malformed")
    created = parse_time(value.get("created_utc"),
                         "baseline-frozen-inputs.created_utc")
    require(created > frozen_created,
            "corrected baseline freeze does not follow initial frozen inputs")
    require(failed_records and created > max(record["end"] for record in failed_records),
            "corrected baseline freeze does not follow the failed attempt")
    require(value.get("stage") == "before corrected baseline freeze/build/capture",
            "baseline frozen-inputs stage differs")
    require(value.get("failed_attempt") == failed_name,
            "baseline frozen-inputs failed attempt differs")
    files = value.get("files")
    require(isinstance(files, dict) and "differential-tests-final.patch" in files,
            "baseline frozen-inputs omit the corrected differential patch")
    for name, digest in files.items():
        safe_relative(name, "baseline frozen-input path")
        require(valid_digest(digest),
                f"baseline frozen-input digest is malformed: {name}")
        target = HERE / name
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"baseline frozen-input hash differs: {name}")
    require(files["differential-tests-final.patch"] ==
            sha(HERE / "differential-tests-final.patch"),
            "baseline frozen-input final patch binding differs")
    require(value.get("initial_frozen_inputs_sha256") == sha(FROZEN),
            "baseline correction does not bind the initial frozen inputs")
    failed_stage = HERE / failed_name
    require(value.get("failed_source_manifest_sha256") ==
            sha(failed_stage / "source-manifest.json"),
            "baseline correction does not bind failed source snapshot")
    failed_receipt = failed_stage / "quality-tests.receipt.json"
    require(value.get("failed_receipt_sha256") == sha(failed_receipt),
            "baseline correction does not bind failed test receipt")
    test_sources = value.get("test_sources")
    require(isinstance(test_sources, dict) and test_sources,
            "baseline correction test source inventory is empty")
    changed = changed_paths(baseline, revision_manifest(plan["revision"]))
    require(set(test_sources) == changed
            and all(test_only_path(name, plan) for name in changed),
            "baseline correction test source inventory differs")
    for name, digest in test_sources.items():
        require(valid_digest(digest) and baseline.get(name) == digest,
                f"baseline correction test source hash differs: {name}")
    reason = value.get("reason")
    require(isinstance(reason, str) and reason.strip(),
            "baseline correction reason is missing")
    return {
        "status": "pass", "created_utc": created.isoformat(),
        "failed_attempt": failed_name,
        "failed_source_manifest_sha256": value["failed_source_manifest_sha256"],
        "failed_receipt_sha256": value["failed_receipt_sha256"],
        "final_patch_sha256": files["differential-tests-final.patch"],
        "test_sources": sorted(test_sources),
    }


def validate_harness_identity(stages: dict[str, dict[str, Any]], harness: set[str],
                              revision: dict[str, str], plan: dict[str, Any]) -> None:
    baseline = stages["baseline"]["manifest"]
    for name in harness:
        require(name in revision and name in baseline,
                f"guard harness path is not present in the frozen revision: {name}")
        # revision_manifest stores SHA-256 values of the Git blob bytes.  The
        # explicit byte check below protects against a changed harness even
        # when a stage manifest happens to be copied consistently.
        require(sha_bytes(git_show(plan["revision"], name)) == baseline[name],
                f"baseline guard harness differs from the committed revision: {name}")
    for name, data in stages.items():
        if name == "baseline":
            continue
        snapshot = data["manifest"]
        for path in set(baseline) | set(snapshot):
            if path in harness or not path.startswith("crates/") \
                    or is_shared_test_path(path, plan):
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
    actual = {
        path.name for path in stage.glob(action + ".*")
        if path.is_file() and not path.is_symlink()
        and path.name != f"{action}.receipt.json"
        # eager_run writes a post-run binding beside the run.py receipt.  It
        # is checked as custody metadata below, rather than as a child output
        # that could have been emitted by the command itself.
        and path.name != f"{action}.binding.json"
        # analyze_planning --write creates deterministic inclusive/self
        # annotations after the profile receipt.  They are analyzer inputs,
        # not outputs emitted by the profiled child.
        and not (
            action.startswith("profile-")
            and path.name in {f"{action}.inclusive.txt", f"{action}.self.txt"}
        )
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


def conditional_capture_command_ok(command: Any, plan: dict[str, Any],
                                   action: str, binary: dict[str, Any],
                                   stage: Path) -> bool:
    """Match the exact profile or whole-child hardware command from run.py."""

    match = re.fullmatch(r"(profile|hardware)-r([1-9][0-9]*)-(.+)", action)
    if match is None or not isinstance(command, list):
        return False
    lane, repeat_text, shape = match.groups()
    config = plan.get(lane)
    if not isinstance(config, dict) or shape not in config.get("shapes", []):
        return False
    try:
        repeat = int(repeat_text)
    except ValueError:
        return False
    if repeat < 1 or repeat > int(config.get("repeats", 0)):
        return False
    primary = plan["primary"]
    expected = ["taskset", "-c", str(plan["cpu"])]
    if lane == "profile":
        owner = plan["profile"]["owner"]
        expected += [
            "valgrind", "--tool=callgrind", "--collect-atstart=no",
            "--toggle-collect=" + owner, "--zero-before=" + owner,
            "--dump-after=" + owner,
            "--callgrind-out-file=" + str(stage / (action + ".callgrind")),
        ]
    else:
        expected += [
            "perf", "stat", "-x", ",", "-o", str(stage / (action + ".csv")),
            "-e", plan["hardware"]["events"], "--",
        ]
    expected += [
        binary["path"], "--warmup", str(config["warmup"]),
        "--samples", str(config["samples"]), "--case", primary["case"],
        "--xlsx-cell-crud-shape", shape,
        "--json", str(stage / (action + ".json")),
    ]
    return command == expected


def eager_job_for_action(action: str, stage_name: str | None = None) -> dict[str, Any] | None:
    """Look up one eager child in the retained eager plan."""

    eager_path = HERE / "eager-plan.json"
    if not eager_path.is_file() or eager_path.is_symlink():
        return None
    eager = read_json(eager_path, "eager-plan.json")
    jobs = eager.get("jobs") if isinstance(eager, dict) else None
    if not isinstance(jobs, list):
        return None
    matches = [item for item in jobs
               if isinstance(item, dict) and item.get("name") == action
               and (stage_name is None or item.get("stage") == stage_name)]
    require(len(matches) <= 1, f"eager plan repeats action name: {action}")
    return matches[0] if matches else None


def eager_capture_command_ok(command: Any, plan: dict[str, Any],
                             action: str, binary: dict[str, Any],
                             stage: Path) -> bool:
    """Match the exact ordinary eager command emitted by eager_run.py."""

    job = eager_job_for_action(action, stage.name)
    if job is None or not isinstance(command, list):
        return False
    time_format = (
        '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,'
        '"system_seconds":%S}'
    )
    expected = [
        "taskset", "-c", str(plan["cpu"]), "/usr/bin/time", "-f", time_format,
        "-o", str(stage / (action + ".rss.json")), binary["path"],
        "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
        "--case", job["case"], "--xlsx-shape", job["shape"],
        "--json", str(stage / (action + ".json")),
    ]
    return command == expected


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
    retained_candidate = (
        stage_name == "baseline"
        and (HERE / "candidate/source-manifest.json").exists()
        and (
            "-r2-" in action
            or action.startswith(("profile-", "hardware-", "eager-",
                                  "build-eager-", "symbols"))
        )
    )
    if retained_candidate:
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

    if action.startswith("eager-"):
        # eager_run emits a binding after run.run writes the generic receipt;
        # validate its complete digest chain here so a pending analyzer cannot
        # hide a substituted child or binary behind the unbound sidecar.
        binding_path = stage / f"{action}.binding.json"
        binding = read_json(binding_path, relative(binding_path))
        require(isinstance(binding, dict),
                f"{relative(binding_path)} is not an object")
        expected_fields = {
            "schema", "stage", "repeat", "group", "name", "case", "shape",
            "source_checkout", "retained_baseline", "primary_plan_sha256",
            "eager_plan_sha256", "run_script_sha256", "eager_run_script_sha256",
            "analyzer_script_sha256", "source_manifest_sha256",
            "working_source_manifest_sha256", "binary_sha256", "receipt_sha256",
            "command_sha256", "output", "rss",
        }
        require(set(binding) == expected_fields,
                f"{relative(binding_path)} fields differ")
        eager_plan_path = HERE / "eager-plan.json"
        require(eager_plan_path.is_file() and not eager_plan_path.is_symlink(),
                "eager receipt has no eager-plan.json")
        eager_plan = read_json(eager_plan_path, "eager-plan.json")
        jobs = eager_plan.get("jobs") if isinstance(eager_plan, dict) else None
        job = next((item for item in jobs or []
                    if isinstance(item, dict) and item.get("name") == action
                    and item.get("stage") == stage_name), None)
        require(job is not None, f"{relative(binding_path)} action is not in eager plan")
        identity = binary_identity(stage, "normal", plan, guard=False, allocator=False)
        require(identity is not None, f"{relative(binding_path)} has no normal binary identity")
        require(binding.get("schema") == "litchi-0543-eager-guard-binding-v1"
                and binding.get("stage") == stage_name
                and binding.get("repeat") == job.get("repeat")
                and binding.get("group") == job.get("group")
                and binding.get("name") == action
                and binding.get("case") == job.get("case")
                and binding.get("shape") == job.get("shape")
                and binding.get("source_checkout") == job.get("source_checkout")
                and binding.get("retained_baseline") == job.get("retained_baseline"),
                f"{relative(binding_path)} job identity differs")
        require(binding.get("primary_plan_sha256") == sha(PLAN)
                and binding.get("eager_plan_sha256") == sha(eager_plan_path)
                and binding.get("run_script_sha256") == sha(RUN)
                and binding.get("eager_run_script_sha256") == sha(HERE / "eager_run.py")
                and binding.get("analyzer_script_sha256") == sha(HERE / "analyze_eager.py")
                and binding.get("source_manifest_sha256") == stage_manifest_sha
                and binding.get("working_source_manifest_sha256") == working_sha
                and binding.get("binary_sha256") == identity["sha256"]
                and binding.get("receipt_sha256") == sha(path)
                and binding.get("command_sha256") == hashlib.sha256(
                    json.dumps(command, separators=(",", ":")).encode()
                ).hexdigest()
                and binding.get("output") == relative(stage / f"{action}.json")
                and binding.get("rss") == relative(stage / f"{action}.rss.json"),
                f"{relative(binding_path)} digest or output binding differs")

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
    elif action.startswith("profile-") or action.startswith("hardware-"):
        identity = binary_identity(stage, "normal", plan, guard=False, allocator=False)
        require(identity is not None,
                f"{relative(path)} has no bound normal binary identity")
        require(receipt.get("binary_sha256") == identity["sha256"],
                f"{relative(path)} binary digest differs")
        expected_command = conditional_capture_command_ok(
            command, plan, action, identity, stage,
        )
    elif action.startswith("eager-"):
        identity = binary_identity(stage, "normal", plan, guard=False, allocator=False)
        require(identity is not None,
                f"{relative(path)} has no bound normal binary identity")
        require(receipt.get("binary_sha256") == identity["sha256"],
                f"{relative(path)} binary digest differs")
        expected_command = eager_capture_command_ok(
            command, plan, action, identity, stage,
        )
    elif action == "symbols":
        identity = binary_identity(stage, "normal", plan, guard=False, allocator=False)
        require(identity is not None,
                f"{relative(path)} has no bound normal binary identity")
        require(receipt.get("binary_sha256") == identity["sha256"],
                f"{relative(path)} binary digest differs")
        expected_command = command == ["nm", "-C", identity["path"]]
    elif action.startswith("build-eager-"):
        # The eager lane is intentionally capture-only.  Retain compatibility
        # with an explicitly named build attempt, but bind it to the same
        # normal release command rather than accepting an arbitrary command.
        expected_command = build_command_ok(
            command, "litchi-perf-baseline", False, guard=False,
        )
    else:
        raise EvidenceError(f"{relative(path)} action is not in the frozen action inventory")
    require(expected_command is True, f"{relative(path)} command differs from the frozen plan")

    if exit_code == 0:
        if action.startswith("native-"):
            required = {f"{action}.json", f"{action}.rss.json"}
            require(required <= artifacts, f"{relative(path)} omits native report or RSS")
        elif action.startswith("alloc-") or action.startswith("guard-"):
            require(f"{action}.json" in artifacts, f"{relative(path)} omits timed report")
        elif action.startswith("profile-"):
            require({f"{action}.json", f"{action}.callgrind"} <= artifacts,
                    f"{relative(path)} omits profile report or Callgrind output")
        elif action.startswith("hardware-"):
            require({f"{action}.json", f"{action}.csv"} <= artifacts,
                    f"{relative(path)} omits hardware report or perf output")
        elif action.startswith("eager-"):
            require({f"{action}.json", f"{action}.rss.json"} <= artifacts,
                    f"{relative(path)} omits eager report or RSS")
        elif action == "symbols":
            require(artifacts == {"symbols.stdout", "symbols.stderr"},
                    f"{relative(path)} symbol artifact set differs")
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
    bound |= {
        f"{action}.binding.json" for action in receipt_names
        if action.startswith("eager-")
    }
    bound |= {name for record in records for name in record["artifacts"]}
    allowed_metadata = {"source-manifest.json", "source.patch", "source-diff.json"}
    allowed_metadata |= {
        path.name for path in stage.glob("binary-*.json")
        if path.is_file() and not path.is_symlink()
    }
    # The planning analyzer creates deterministic inclusive/self annotation
    # sidecars after the profile receipt.  They are analyzer inputs rather
    # than child outputs, so they are allowed only for an already-retained
    # profile action; replay_analyses verifies their exact bytes and paths.
    profile = plan.get("profile")
    if isinstance(profile, dict) and isinstance(profile.get("shapes"), list) \
            and isinstance(profile.get("repeats"), int):
        for repeat in range(1, profile["repeats"] + 1):
            for shape in profile["shapes"]:
                action = f"profile-r{repeat}-{shape}"
                if action in receipt_names:
                    allowed_metadata.update({
                        f"{action}.inclusive.txt", f"{action}.self.txt",
                    })
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


def validate_symbol_observation(
    plan: dict[str, Any],
    records_by_stage: dict[str, list[dict[str, Any]]],
    binaries: dict[str, dict[str, Any]],
) -> dict[str, Any]:
    """Bind the optional ``nm -C`` owner observation to one retained binary.

    The profile analyzer can replay its exact owner edge without this helper,
    but a retained symbol receipt/observation must never be left as loose
    metadata.  The observation is deliberately singular: selecting between
    multiple binaries or owner spellings would make the profile attribution
    non-reproducible.
    """

    observation_path = HERE / "symbol-observation.json"
    rows = [
        (stage, record)
        for stage, records in records_by_stage.items()
        for record in records
        if record["name"] == "symbols"
    ]
    if not rows and not observation_path.exists():
        return {"status": "not-present"}
    require(observation_path.is_file() and not observation_path.is_symlink(),
            "symbol observation is not a regular file")
    observation = read_json(observation_path, relative(observation_path))
    require(isinstance(observation, dict),
            "symbol observation is not an object")
    owner = plan["profile"]["owner"]
    require(observation.get("owner") == owner,
            "symbol observation owner differs from plan.profile.owner")
    require(observation.get("plan_sha256") == sha(PLAN),
            "symbol observation plan binding differs")

    # The fresh 0543 conditional coordinator records both direct ``nm``
    # invocations in one root-owned observation, without creating runner
    # receipts or persisting raw nm output.  Validate that form first.
    stages = observation.get("stages")
    if stages is not None:
        require(not rows, "symbol observation mixes stage records with symbols receipts")
        require(isinstance(stages, dict) and set(stages) == {"baseline", "candidate"},
                "symbol observation stage inventory differs")
        result_stages: dict[str, Any] = {}
        for stage_name in ("baseline", "candidate"):
            item = stages[stage_name]
            require(isinstance(item, dict)
                    and set(item) == {
                        "command", "binary_sha256", "source_manifest_sha256",
                        "matching_lines",
                    },
                    f"symbol observation {stage_name} fields differ")
            identity = binaries[stage_name].get("normal")
            require(identity is not None,
                    f"symbol observation {stage_name} has no normal binary identity")
            expected_command = ["nm", "-C", "--defined-only", identity["path"]]
            require(item.get("command") == expected_command
                    and item.get("binary_sha256") == identity["sha256"]
                    and item.get("source_manifest_sha256") == sha(
                        HERE / stage_name / "source-manifest.json"
                    ),
                    f"symbol observation {stage_name} binary or command binding differs")
            lines = item.get("matching_lines")
            require(isinstance(lines, list) and lines
                    and all(isinstance(line, str) and line and owner in line
                            for line in lines),
                    f"symbol observation {stage_name} owner matches are malformed")
            result_stages[stage_name] = {
                "binary_sha256": identity["sha256"], "matches": len(lines),
            }
        created = parse_time(observation.get("created_utc"),
                             "symbol-observation.created_utc")
        return {
            "status": "pass", "owner": owner, "stages": result_stages,
            "created_utc": created.isoformat(),
            "observation_sha256": sha(observation_path),
        }

    # Retain compatibility with the earlier profile coordinator's single
    # baseline ``run.run`` receipt.  This branch is strict when that older
    # schema is used and cannot silently accept an arbitrary symbol command.
    require(len(rows) == 1, "symbol observation has zero or multiple symbol receipts")
    stage_name, record = rows[0]
    require(record["exit_code"] == 0,
            f"{stage_name}/symbols receipt is not successful")
    required = {
        "owner", "candidates", "binary_sha256", "plan_sha256", "command",
        "receipt_sha256", "stdout_sha256", "selection", "selected_utc",
    }
    require(required <= set(observation),
            "symbol observation omits required custody fields")
    identity = binaries[stage_name].get("normal")
    require(identity is not None,
            f"{stage_name}/symbols has no normal binary identity")
    receipt_path = HERE / stage_name / "symbols.receipt.json"
    stdout_path = HERE / stage_name / "symbols.stdout"
    require(observation.get("binary_sha256") == identity["sha256"]
            and observation.get("command") == ["nm", "-C", identity["path"]]
            and observation.get("receipt_sha256") == sha(receipt_path)
            and observation.get("stdout_sha256") == sha(stdout_path),
            "symbol observation binary, command, or digest binding differs")
    candidates = observation.get("candidates")
    require(isinstance(candidates, dict) and set(candidates) == {owner},
            "symbol observation candidate inventory differs from the direct owner")
    lines = candidates.get(owner)
    require(isinstance(lines, list) and lines and
            all(isinstance(line, str) and line for line in lines),
            "symbol observation owner matches are malformed")
    raw_lines = read_text(stdout_path, relative(stdout_path)).splitlines()
    require(all(line in raw_lines for line in lines),
            "symbol observation owner matches are absent from nm output")
    selection = observation.get("selection")
    require(isinstance(selection, str) and selection.strip(),
            "symbol observation selection explanation is missing")
    selected = parse_time(observation.get("selected_utc"),
                          "symbol-observation.selected_utc")
    require(selected >= record["end"],
            "symbol observation was selected before nm receipt completed")
    return {
        "status": "pass", "stage": stage_name, "owner": owner,
        "matches": len(lines), "binary_sha256": identity["sha256"],
        "receipt_sha256": sha(receipt_path),
        "stdout_sha256": sha(stdout_path),
        "selected_utc": selected.isoformat(),
        "observation_sha256": sha(observation_path),
    }


def conditional_jobs(plan: dict[str, Any], lane: str) -> dict[tuple[str, int], set[str]]:
    """Return the frozen per-stage/repeat job names for one conditional lane."""

    if lane == "profile":
        config = plan["profile"]
    elif lane == "hardware":
        config = plan["hardware"]
    else:
        eager_path = HERE / "eager-plan.json"
        require(eager_path.is_file() and not eager_path.is_symlink(),
                "eager receipts have no eager plan")
        eager = read_json(eager_path, "eager-plan.json")
        jobs = eager.get("jobs") if isinstance(eager, dict) else None
        require(isinstance(jobs, list) and jobs,
                "eager plan has no job inventory")
        result: dict[tuple[str, int], set[str]] = {}
        for item in jobs:
            require(isinstance(item, dict)
                    and item.get("stage") in {"baseline", "candidate"}
                    and isinstance(item.get("repeat"), int)
                    and isinstance(item.get("name"), str),
                    "eager plan job is malformed")
            result.setdefault((item["stage"], item["repeat"]), set()).add(item["name"])
        return result
    require(isinstance(config, dict), f"plan.{lane} is not an object")
    shapes = config.get("shapes")
    repeats = config.get("repeats")
    require(isinstance(shapes, list) and shapes and isinstance(repeats, int)
            and repeats > 0, f"plan.{lane} job matrix is malformed")
    return {
        (stage, repeat): {
            f"{lane}-r{repeat}-{shape}" for shape in shapes
        }
        for stage in ("baseline", "candidate")
        for repeat in range(1, repeats + 1)
    }


def validate_conditional_abba(records_by_stage: dict[str, list[dict[str, Any]]],
                              plan: dict[str, Any]) -> dict[str, Any]:
    """Check ABBA ordering for profile, hardware, and eager child groups."""

    result: dict[str, Any] = {}
    for lane in ("profile", "hardware", "eager"):
        records = {
            stage: [record for record in rows
                    if record["name"].startswith(lane + "-")]
            for stage, rows in records_by_stage.items()
        }
        all_lane = [record for rows in records.values() for record in rows]
        if not all_lane:
            result[lane] = {"status": "unmeasured", "receipts": 0}
            continue
        jobs = conditional_jobs(plan, lane)
        expected_names = {
            (stage, repeat): names
            for (stage, repeat), names in jobs.items()
        }
        actual_names: dict[tuple[str, int], set[str]] = {}
        for stage, rows in records.items():
            for record in rows:
                match = re.search(r"-r([0-9]+)-", record["name"])
                require(match is not None,
                        f"{stage}/{record['name']} has no conditional repeat")
                key = (stage, int(match.group(1)))
                actual_names.setdefault(key, set()).add(record["name"])
        extras = set(actual_names) - set(expected_names)
        require(not extras,
                f"{lane} conditional receipts contain unexpected groups: {sorted(extras)}")
        missing_groups = set(expected_names) - set(actual_names)
        missing_jobs = {
            key: sorted(expected_names[key] - actual_names.get(key, set()))
            for key in expected_names
            if expected_names[key] != actual_names.get(key, set())
        }
        if missing_jobs:
            # A failed conditional command can leave a partial retained
            # matrix; this is auditable and forces an incomplete/rejected
            # outcome.  An accepted outcome is checked again below and must
            # have every group.
            result[lane] = {
                "status": "pending", "receipts": len(all_lane),
                "missing_groups": sorted(missing_groups),
                "missing_jobs": missing_jobs,
            }
            continue
        if any(record["exit_code"] != 0 for record in all_lane):
            result[lane] = {
                "status": "failed", "receipts": len(all_lane),
                "failed": sorted(record["name"] for record in all_lane
                                  if record["exit_code"] != 0),
            }
            continue
        groups: list[tuple[str, str, int, dt.datetime]] = []
        for (stage, repeat), names in expected_names.items():
            group = [record for record in records.get(stage, [])
                     if record["name"] in names]
            require(group, f"{lane} {stage} repeat {repeat} is empty")
            groups.append((lane, stage, repeat,
                           min(record["start"] for record in group)))
        groups.sort(key=lambda item: item[3])
        expected_order = [
            (lane, "baseline", 1), (lane, "candidate", 1),
            (lane, "candidate", 2), (lane, "baseline", 2),
        ]
        actual_order = [(lane, stage, repeat) for _, stage, repeat, _ in groups]
        require(actual_order == expected_order,
                f"{lane} conditional capture order is not baseline/candidate/candidate/baseline")
        result[lane] = {
            "status": "pass", "receipts": len(all_lane),
            "order": [
                {"stage": stage, "repeat": repeat, "start_utc": start.isoformat()}
                for _, stage, repeat, start in groups
            ],
        }
    return result


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
        "litchi_0543_" + path.stem.replace("-", "_"),
        register=path.name == "analyze.py",
    )
    method = getattr(module, function, None)
    require(callable(method), f"{display_path(path)} has no callable {function}()")
    try:
        result = method()
    except Exception as error:  # noqa: BLE001 - analyzer defines evidence errors
        raise EvidenceError(f"{display_path(path)} replay failed: {error}") from error
    require(result == expected,
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

    # Conditional lanes are deliberately replayed when present.  Their
    # reports are not folded into the native/allocation admission, but an
    # accepted candidate may not claim retention without the profile Ir gate
    # and the ordinary eager-read guard having been run and replayed.
    conditional_specs = {
        "profile": (HERE / "analyze_planning.py",
                     [HERE / "planning-profile-analysis.json",
                      HERE / "profile-analysis.json", HERE / "profile-comparison.json"]),
        "eager": (HERE / "analyze_eager.py",
                  [HERE / "eager-analysis.json", HERE / "eager-guard-comparison.json"]),
        "hardware": (HERE / "analyze_hardware.py",
                     [HERE / "hardware-analysis.json", HERE / "hardware-comparison.json"]),
    }
    for lane, (lane_analyzer, report_candidates) in conditional_specs.items():
        report = first_existing(report_candidates)
        receipt_prefix = lane + "-"
        receipt_present = any(
            record["name"].startswith(receipt_prefix)
            for rows in records_by_stage.values() for record in rows
        )
        if report is None:
            result[lane] = {
                "status": "pending" if receipt_present else "unmeasured",
                "reason": ("conditional receipts exist without an analysis report"
                           if receipt_present else "conditional lane was not captured"),
            }
            continue
        require(lane_analyzer.is_file() and not lane_analyzer.is_symlink(),
                f"{lane} analysis report has no analyzer")
        result[lane] = replay_analyzer(lane_analyzer, report)
        result[lane]["report_path"] = relative(report)
        value = read_json(report, relative(report))
        require(isinstance(value, dict) and value.get("status") == "pass",
                f"{relative(report)} is not a passing conditional report")
        if lane == "profile":
            rows = value.get("rows")
            require(isinstance(rows, list) and rows,
                    f"{relative(report)} has no paired Ir rows")
            required_ir = float(plan["gates"]["planning_ir_reduction_percent"])
            for index, row in enumerate(rows):
                require(isinstance(row, dict),
                        f"{relative(report)} rows[{index}] is not an object")
                metric = row.get("planning_ir")
                require(isinstance(metric, dict)
                        and metric.get("passed") is True
                        and isinstance(metric.get("reduction_percent"), (int, float))
                        and float(metric["reduction_percent"]) >= required_ir,
                        f"{relative(report)} profile Ir gate is not met for row {index}")
            result[lane]["ir_rows"] = len(rows)
            result[lane]["ir_gate_passed"] = True
        elif lane == "eager":
            comparison = value.get("comparison")
            require(isinstance(comparison, dict)
                    and comparison.get("runtime_comparison_performed") is True,
                    f"{relative(report)} eager comparison is incomplete")
            result[lane]["eager_comparison_passed"] = True
            result[lane]["individual_review_count"] = comparison.get(
                "individual_review_count", 0)
        else:
            # Hardware is diagnostic and has no admission gate, but a
            # retained report must still be a complete deterministic replay.
            result[lane]["diagnostic"] = True
    return result


def replay_hardware_review(analyses: dict[str, Any]) -> dict[str, Any]:
    """Replay the separately derived individual hardware flag review."""

    present = [path.exists() for path in
               (HARDWARE_REVIEW, HARDWARE_REVIEW_SCRIPT, HARDWARE_REVIEW_MARKDOWN)]
    if not any(present):
        require(analyses.get("hardware", {}).get("status") != "pass",
                "hardware review artifacts are missing after hardware analysis")
        return {"status": "not-present"}
    require(all(present), "hardware review artifacts are only partially retained")
    require(analyses.get("hardware", {}).get("status") == "pass",
            "hardware review exists without a passing hardware analysis")
    expected = read_json(HARDWARE_REVIEW, "hardware-review.json")
    module = load_module(HARDWARE_REVIEW_SCRIPT, "litchi_0543_hardware_review",
                         register=False)
    method = getattr(module, "review", None)
    require(callable(method), "review_hardware.py has no callable review()")
    try:
        result = method()
    except Exception as error:  # noqa: BLE001 - preserve review context
        raise EvidenceError(f"hardware review replay failed: {error}") from error
    try:
        replayed_json = json.loads(json.dumps(
            result, ensure_ascii=False, sort_keys=True, allow_nan=False,
        ))
    except (TypeError, ValueError) as error:
        raise EvidenceError("hardware review replay is not JSON serializable") from error
    require(replayed_json == expected,
            "hardware-review.json differs from deterministic review replay")
    required = {
        "schema", "status", "hardware_analysis_sha256", "derivation_sha256",
        "threshold_percent", "ipc_adverse_direction", "other_metrics_adverse_direction",
        "same_build_drift_direction", "scope", "adverse_count", "drift_count", "flags",
    }
    require(isinstance(result, dict) and set(result) == required
            and result.get("schema") == "litchi.0543.hardware-review.v1"
            and result.get("status") == "complete"
            and result.get("hardware_analysis_sha256") == sha(HARDWARE_ANALYSIS)
            and result.get("derivation_sha256") == sha(HARDWARE_REVIEW_SCRIPT)
            and result.get("threshold_percent") == 5
            and result.get("ipc_adverse_direction") == "decrease"
            and result.get("other_metrics_adverse_direction") == "increase"
            and result.get("same_build_drift_direction") == "absolute"
            and isinstance(result.get("scope"), str)
            and "whole-child" in result["scope"].lower()
            and "diagnostic" in result["scope"].lower(),
            "hardware review envelope differs")
    flags = result.get("flags")
    require(isinstance(flags, list), "hardware review flags are missing")
    adverse = [flag for flag in flags
               if isinstance(flag, dict) and flag.get("kind") == "candidate_adverse"]
    drift = [flag for flag in flags
             if isinstance(flag, dict) and flag.get("kind") == "same_build_drift"]
    require(len(adverse) == result["adverse_count"]
            and len(drift) == result["drift_count"]
            and result["adverse_count"] == 1
            and result["drift_count"] == 4,
            "hardware review flag counts differ")
    ids: set[str] = set()
    for index, flag in enumerate(flags):
        require(isinstance(flag, dict), f"hardware-review.flags[{index}] is not an object")
        expected_keys = {
            "kind", "shape", "repeat", "metric", "baseline", "candidate",
            "change_percent", "id", "interpretation", "disposition",
        }
        if flag.get("kind") == "same_build_drift":
            expected_keys = {
                "kind", "stage", "shape", "metric", "first", "second",
                "change_percent", "id", "interpretation", "disposition",
            }
        require(set(flag) == expected_keys,
                f"hardware-review.flags[{index}] fields differ")
        identifier = flag.get("id")
        require(identifier == f"0543-hardware-{index + 1:03d}"
                and identifier not in ids, f"hardware-review.flags[{index}] ID differs")
        ids.add(identifier)
        require(isinstance(flag.get("metric"), str) and flag["metric"]
                and isinstance(flag.get("shape"), str) and flag["shape"]
                and isinstance(flag.get("interpretation"), str)
                and flag["interpretation"].strip()
                and isinstance(flag.get("disposition"), str)
                and flag["disposition"].strip(),
                f"hardware-review.flags[{index}] review fields are incomplete")
        change = flag.get("change_percent")
        require(isinstance(change, (int, float)) and not isinstance(change, bool)
                and math.isfinite(float(change)) and abs(float(change)) > 5.0,
                f"hardware-review.flags[{index}] change is not an adverse finite value")
        if flag["kind"] == "candidate_adverse":
            require(isinstance(flag.get("repeat"), int)
                    and not isinstance(flag["repeat"], bool)
                    and ((flag["metric"] == "ipc" and change < -5.0)
                         or (flag["metric"] != "ipc" and change > 5.0)),
                    f"hardware-review.flags[{index}] adverse direction differs")
        else:
            require(isinstance(flag.get("stage"), str) and flag["stage"] in
                    {"baseline", "candidate"},
                    f"hardware-review.flags[{index}] drift stage differs")
    markdown = read_text(HARDWARE_REVIEW_MARKDOWN, "hardware-review.md").lower()
    require("hardware" in markdown and "diagnostic" in markdown
            and "individual" in markdown and "context-switch" in markdown
            and all(identifier.lower() in markdown for identifier in ids),
            "hardware-review.md does not cover every individual flag")
    return {
        "status": "pass", "report_path": relative(HARDWARE_REVIEW),
        "report_sha256": sha(HARDWARE_REVIEW),
        "derivation_path": relative(HARDWARE_REVIEW_SCRIPT),
        "derivation_sha256": sha(HARDWARE_REVIEW_SCRIPT),
        "hardware_analysis_sha256": result["hardware_analysis_sha256"],
        "adverse_count": result["adverse_count"],
        "drift_count": result["drift_count"],
    }


def validate_adverse_review(analyses: dict[str, Any], decision: dict[str, Any]) -> dict[str, Any]:
    """Bind every reviewed adverse/drift row to the replayed reports.

    A summary count alone can hide an omitted or substituted flag.  The
    review therefore carries the original analyzer row for each flag; this
    check compares those rows as a multiset, while also requiring an ID,
    interpretation, and disposition for every retained record.
    """
    review_path = HERE / "adverse-review.json"
    review = read_json(review_path, "adverse-review.json")
    require(isinstance(review, dict), "adverse review is not an object")
    schema_version = review.get("schema_version")
    require(schema_version in {2, 3} and review.get("status") == "complete",
            "adverse review is not a complete supported schema version")
    review_word = review.get("disposition")
    require(isinstance(review_word, str), "adverse review disposition is missing")
    review_normalized = review_word.strip().lower()
    expected_decision = decision["decision"]
    if expected_decision == "rejected":
        require(review_normalized in {
            "reject", "rejected", "rejected_and_reverted", "baseline_retained",
            "rejected-and-restored", "rejected_and_restored", "restore_baseline",
        },
                "adverse review disposition does not record rejection")
    else:
        require(review_normalized in {"accept", "accepted", "retain", "retained"},
                "adverse review disposition does not record acceptance")
    scope = review.get("scope")
    require(isinstance(scope, str) and "0543" in scope.lower()
            and "ooxml" in scope.lower() and "odf" in scope.lower(),
            "adverse review scope omits pilot or format priority")

    report_bindings = {
        "comparison_sha256": HERE / "comparison.json",
        "allocation_analysis_sha256": HERE / "allocation-analysis.json",
        "guard_analysis_sha256": HERE / "guard-analysis.json",
        "plan_sha256": PLAN,
    }
    for field, path in report_bindings.items():
        require(valid_digest(review.get(field)) and review[field] == sha(path),
                f"adverse review {field} does not bind {path.name}")
    if "analysis_status_correction_sha256" in review:
        correction_path = HERE / "analysis-status-correction.json"
        require(correction_path.is_file()
                and review["analysis_status_correction_sha256"] == sha(correction_path),
                "adverse review allocation status correction binding differs")
    if "decision_status" in review:
        decision_status = review["decision_status"]
        require(isinstance(decision_status, str) and decision_status.strip(),
                "adverse review decision_status is malformed")
        normalized_status = decision_status.strip().lower()
        accepted_statuses = {"accept", "accepted", "retain", "retained"}
        rejected_statuses = {
            "reject", "rejected", "rejected_and_reverted", "baseline_retained",
            "rejected-and-restored", "rejected_and_restored", "restore_baseline",
        }
        require(
            normalized_status in (accepted_statuses if expected_decision == "accepted"
                                  else rejected_statuses),
            "adverse review decision_status differs from decision.json",
        )
    # Conditional reports are optional after a rejected pilot, but once one
    # is retained its exact JSON digest must be carried by the review.  The
    # aliases keep the custody schema readable for the two report spellings
    # used by the profile/eager drivers.
    conditional_binding_fields = {
        "profile": ("planning_profile_analysis_sha256", "profile_analysis_sha256"),
        "eager": ("eager_analysis_sha256", "eager_guard_analysis_sha256"),
        "hardware": ("hardware_analysis_sha256", "hardware_comparison_sha256"),
        "hardware_review": ("hardware_review_sha256", "hardware_review_analysis_sha256"),
        "cap": ("cap_analysis_sha256", "cap_boundary_analysis_sha256"),
    }
    for lane, fields in conditional_binding_fields.items():
        summary = analyses.get(lane, {})
        if summary.get("status") != "pass":
            continue
        report_name = summary.get("report_path")
        require(isinstance(report_name, str),
                f"replayed {lane} report has no path")
        report_path = HERE / report_name
        bound = [field for field in fields if field in review]
        require(len(bound) == 1 and review[bound[0]] == sha(report_path),
                f"adverse review does not bind the replayed {lane} report")
    require(analyses.get("native", {}).get("status") == "pass"
            and analyses.get("allocation", {}).get("status") == "pass"
            and analyses.get("guard", {}).get("status") == "pass",
            "adverse review is not backed by complete analyzer replay")
    native_review = review.get("native_admission")
    require(isinstance(native_review, dict), "adverse review native admission is missing")
    native = analyses["native"]
    if "decision" in native_review:
        require(native_review["decision"] == native.get("admission_status"),
                "adverse review native admission differs")
    if "passed" in native_review:
        require(native_review["passed"] is
                (isinstance(native.get("native_admission"), dict)
                 and native["native_admission"].get("passed") is True),
                "adverse review native pass flag differs")
    allocation_review = review.get("allocation_diagnostics")
    require(isinstance(allocation_review, dict), "adverse review allocation diagnostics are missing")
    if "planning_gate_passed" in allocation_review:
        require(allocation_review["planning_gate_passed"] is
                (analyses["allocation"].get("planning_gate_passed") is True),
                "adverse review allocation pass flag differs")
    guard_review = review.get("guard_admission")
    require(isinstance(guard_review, dict), "adverse review guard admission is missing")
    if "admission_status" in guard_review:
        require(guard_review["admission_status"] == analyses["guard"].get("admission_status"),
                "adverse review guard admission differs")

    comparison = read_json(HERE / "comparison.json", "comparison.json")
    allocation = read_json(HERE / "allocation-analysis.json", "allocation-analysis.json")
    guard = read_json(HERE / "guard-analysis.json", "guard-analysis.json")
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
    ]
    conditional_review_sources: list[tuple[str, str, list[Any]]] = []
    for lane, summary in analyses.items():
        if lane not in {"profile", "eager", "hardware", "cap"} or summary.get("status") != "pass":
            continue
        report_name = summary.get("report_path")
        require(isinstance(report_name, str), f"{lane} analysis report path is missing")
        report = read_json(HERE / report_name, report_name)
        if lane == "eager":
            comparison_value = report.get("comparison")
            require(isinstance(comparison_value, dict),
                    f"{report_name} lacks eager comparison rows")
            conditional_review_sources.extend([
                ("eager_adverse_flags", f"{report_name}:comparison.adverse_flags_over_five_percent",
                 comparison_value.get("adverse_flags_over_five_percent")),
                ("eager_same_build_drift_flags",
                 f"{report_name}:comparison.same_build_drift_over_five_percent",
                 comparison_value.get("same_build_drift_over_five_percent")),
            ])
        elif lane == "hardware":
            comparison_value = report.get("comparison")
            if isinstance(comparison_value, dict):
                # The 0543 hardware analyzer is diagnostic and emits no
                # review vectors today.  Accept future explicit vectors under
                # the same naming convention and bind them if present.
                for key, field in (("hardware_adverse_flags",
                                    "adverse_flags_over_five_percent"),
                                   ("hardware_same_build_drift_flags",
                                    "same_build_drift_over_five_percent")):
                    if field in comparison_value:
                        conditional_review_sources.append(
                            (key, f"{report_name}:comparison.{field}",
                             comparison_value[field]))
        elif lane == "profile":
            # Profile analyzer failures are gate failures, not an adverse
            # timing lane.  If a future report emits review vectors, include
            # them rather than silently accepting an unreviewed list.
            for key, field in (("profile_adverse_flags",
                                "adverse_flags_over_five_percent"),
                               ("profile_same_build_drift_flags",
                                "same_build_drift_over_five_percent")):
                if field in report:
                    conditional_review_sources.append(
                        (key, f"{report_name}:{field}", report[field]))
        elif lane == "cap":
            comparison_value = report.get("comparison")
            require(isinstance(comparison_value, dict),
                    f"{report_name} lacks cap-boundary comparison rows")
            conditional_review_sources.extend([
                ("cap_adverse_flags", f"{report_name}:comparison.adverse",
                 comparison_value.get("adverse")),
                ("cap_same_build_drift_flags", f"{report_name}:comparison.drift",
                 comparison_value.get("drift")),
            ])
    # Schema v3 names the conditional lists explicitly.  Accept the compact
    # ``conditional_*`` aliases only when a producer used those names
    # consistently; either spelling still carries the exact source rows.
    review_key_aliases = {
        "eager_adverse_flags": ("eager_adverse_flags", "conditional_adverse_flags"),
        "eager_same_build_drift_flags": (
            "eager_same_build_drift_flags", "conditional_same_build_drift_flags",
        ),
        "hardware_adverse_flags": ("hardware_adverse_flags",),
        "hardware_same_build_drift_flags": ("hardware_same_build_drift_flags",),
        "profile_adverse_flags": ("profile_adverse_flags",),
        "profile_same_build_drift_flags": ("profile_same_build_drift_flags",),
        "cap_adverse_flags": (
            "cap_adverse_flags", "cap_boundary_adverse_flags",
            "conditional_cap_adverse_flags",
        ),
        "cap_same_build_drift_flags": (
            "cap_same_build_drift_flags", "cap_boundary_same_build_drift_flags",
            "conditional_cap_same_build_drift_flags",
        ),
    }
    if schema_version == 3:
        remapped: list[tuple[str, str, list[Any]]] = []
        for key, source, rows in conditional_review_sources:
            aliases = review_key_aliases.get(key, (key,))
            present = [name for name in aliases if name in review]
            require(len(present) <= 1,
                    f"adverse review has duplicate conditional list aliases for {key}")
            remapped.append((present[0] if present else key, source, rows))
        conditional_review_sources = remapped
    for group in conditional_review_sources:
        require(isinstance(group[2], list), f"{group[1]} is not a list")
    expected_groups.extend(conditional_review_sources)
    for key, source, rows in expected_groups:
        require(isinstance(rows, list), f"{source} is not a list")
        require(isinstance(review.get(key), list), f"adverse review {key} is not a list")

    expected_rows = [(source, row) for _, source, rows in expected_groups for row in rows]
    review_rows = []
    ids: set[str] = set()
    review_keys = list(dict.fromkeys(key for key, _, _ in expected_groups))
    for key in review_keys:
        for index, item in enumerate(review[key]):
            label = f"adverse-review.{key}[{index}]"
            require(isinstance(item, dict), f"{label} is not an object")
            require(set(item) == {"id", "source", "original", "classification",
                                  "interpretation", "disposition"},
                    f"{label} fields differ")
            identifier = item.get("id")
            require(isinstance(identifier, str) and identifier and identifier not in ids,
                    f"{label} ID is missing or duplicated")
            ids.add(identifier)
            require(isinstance(item.get("source"), str) and item["source"],
                    f"{label} source is missing")
            require(isinstance(item.get("classification"), str) and item["classification"],
                    f"{label} classification is missing")
            require(isinstance(item.get("interpretation"), str) and item["interpretation"].strip(),
                    f"{label} interpretation is missing")
            require(isinstance(item.get("disposition"), str) and item["disposition"].strip(),
                    f"{label} disposition is missing")
            review_rows.append((item["source"], item["original"]))
    require(_multiset(review_rows, "adverse-review originals") ==
            _multiset(expected_rows, "analyzer adverse rows"),
            "adverse review originals do not exactly match analyzer flags")

    coverage = review.get("source_coverage")
    expected_coverage = [{"source": source, "review_key": key, "count": len(rows)}
                         for key, source, rows in expected_groups]
    # The two guard adverse sources intentionally share one review list; both
    # source rows remain separate in the coverage inventory.
    require(_multiset(coverage if isinstance(coverage, list) else [], "source coverage") ==
            _multiset(expected_coverage, "expected source coverage"),
            "adverse review source coverage differs")
    counts = review.get("counts")
    expected_counts = {
        "native_adverse_flags_over_five_percent": len(expected_groups[0][2]),
        "native_same_build_drift_over_five_percent": len(expected_groups[1][2]),
        "allocation_adverse_flags": len(expected_groups[2][2]),
        "guard_normal_adverse_flags_over_five_percent": len(expected_groups[3][2]),
        "guard_allocation_adverse_flags_over_five_percent": len(expected_groups[4][2]),
        "guard_same_build_drift_over_five_percent": len(expected_groups[5][2]),
        "adverse_flags_over_five_percent": len(expected_groups[0][2])
        + len(expected_groups[2][2]) + len(expected_groups[3][2]) + len(expected_groups[4][2]),
        "same_build_drift_over_five_percent": len(expected_groups[1][2])
        + len(expected_groups[5][2]),
        "reviewed_flags": len(expected_rows),
    }
    for key, _, rows in conditional_review_sources:
        expected_counts[key] = len(rows)
    if conditional_review_sources:
        expected_counts["conditional_adverse_flags_over_five_percent"] = sum(
            len(rows) for key, _, rows in conditional_review_sources
            if "adverse" in key
        )
        expected_counts["conditional_same_build_drift_over_five_percent"] = sum(
            len(rows) for key, _, rows in conditional_review_sources
            if "drift" in key
        )
    if schema_version == 2:
        require(counts == expected_counts, "adverse review counts differ from analyzer flags")
    else:
        require(isinstance(counts, dict), "adverse review counts are missing")
        base_counts = {
            key: value for key, value in expected_counts.items()
            if not key.startswith("conditional_")
        }
        for key, value in base_counts.items():
            require(counts.get(key) == value,
                    f"adverse review count differs: {key}")
        if conditional_review_sources:
            conditional_adverse = expected_counts.get(
                "conditional_adverse_flags_over_five_percent", 0)
            conditional_drift = expected_counts.get(
                "conditional_same_build_drift_over_five_percent", 0)
            aliases = {
                "adverse": (
                    "conditional_adverse_flags_over_five_percent",
                    "conditional_adverse_flags",
                ),
                "drift": (
                    "conditional_same_build_drift_over_five_percent",
                    "conditional_same_build_drift_flags",
                ),
            }
            for kind, expected in (("adverse", conditional_adverse),
                                   ("drift", conditional_drift)):
                present = [name for name in aliases[kind] if name in counts]
                require(len(present) == 1 and counts[present[0]] == expected,
                        f"adverse review conditional {kind} count differs")
            known = set(base_counts) | set(aliases["adverse"]) | set(aliases["drift"])
            pending = {
                "conditional_adverse_flags_pending",
                "conditional_same_build_drift_flags_pending",
            }
            require(set(counts) <= known | pending,
                    "adverse review contains an unknown count field")
            for key in pending:
                if key in counts:
                    require(counts[key] == 0,
                            f"adverse review retains pending conditional count: {key}")
        else:
            require(not any(key.startswith("conditional_") for key in counts),
                    "adverse review claims conditional rows without conditional reports")

        conditional_lane = review.get("conditional_lanes")
        require(isinstance(conditional_lane, dict),
                "schema v3 adverse review conditional_lanes is missing")
        expected_conditional_adverse = sum(
            len(rows) for key, _, rows in conditional_review_sources if "adverse" in key
        )
        expected_conditional_drift = sum(
            len(rows) for key, _, rows in conditional_review_sources if "drift" in key
        )
        if conditional_review_sources:
            require(conditional_lane.get("status") in {"complete", "pass", "reviewed"},
                    "schema v3 conditional review is not complete")
        for key, expected in (
            ("adverse_flags_reviewed", expected_conditional_adverse),
            ("same_build_drift_flags_reviewed", expected_conditional_drift),
        ):
            require(conditional_lane.get(key) == expected,
                    f"schema v3 conditional review count differs: {key}")
    classes: dict[str, int] = {}
    for key in review_keys:
        for item in review[key]:
            classes[item["classification"]] = classes.get(item["classification"], 0) + 1
    summary = review.get("classification_summary")
    require(isinstance(summary, dict) and set(summary) == set(classes),
            "adverse review classification summary differs")
    for name, count in classes.items():
        require(isinstance(summary[name], dict) and summary[name].get("count") == count,
                f"adverse review classification count differs: {name}")
    markdown = HERE / "adverse-review.md"
    note = read_text(markdown, "adverse-review.md").lower()
    require("0543" in note and "individually" in note and "decision" in note,
            "adverse-review.md lacks explicit individual decision review")
    return {
        "status": "pass", "records": len(expected_rows), "unique_ids": len(ids),
        "comparison_sha256": review["comparison_sha256"],
        "allocation_analysis_sha256": review["allocation_analysis_sha256"],
        "guard_analysis_sha256": review["guard_analysis_sha256"],
        "plan_sha256": review["plan_sha256"],
    }


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
    # outcome while the 0543 driver uses ``decision``.  Accept the aliases
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
        conditional = conditional_receipts(records_by_stage)
        require(conditional["profile"]["status"] == "pass"
                and conditional["eager"]["status"] == "pass"
                and analyses.get("profile", {}).get("ir_gate_passed") is True
                and analyses.get("eager", {}).get("eager_comparison_passed") is True,
                "accepted decision lacks replayed profile Ir and eager guard evidence")
    runtime_restored = decision.get("runtime_restored")
    if runtime_restored is not None:
        require(isinstance(runtime_restored, bool), "decision.runtime_restored is malformed")
        require(runtime_restored is rejected,
                "decision runtime_restored flag contradicts accepted/rejected outcome")
    return {
        "decision": "accepted" if accepted else "rejected", "final_stage": final,
        "quality": quality, "native_status": native.get("status"),
        "allocation_status": allocation.get("status"), "guard_status": guard.get("status"),
        "conditional": conditional_receipts(records_by_stage),
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
    require(changed <= set(harness) | set(roots)
            | {path for path in changed if test_only_path(path, plan)},
            "current source changed outside plan")
    if decision["decision"] == "rejected":
        baseline = stage_data["baseline"]["manifest"]
        require(final_manifest == baseline,
                "rejected decision final source is not the restored baseline")
    else:
        require(final_stage == "candidate", "accepted decision does not select candidate source")
    return {"manifest_sha256": stage_data[final_stage]["manifest_sha256"],
            "entries": len(final_manifest), "changed_files": sorted(changed)}


def validate_cleanup(plan: dict[str, Any], binaries: dict[str, dict[str, Any]],
                     records: list[dict[str, Any]],
                     *, extra_binaries: dict[str, dict[str, Any]] | None = None,
                     extra_records: list[dict[str, Any]] | None = None) -> dict[str, Any]:
    value = read_json(CLEANUP, "cleanup.json")
    require(isinstance(value, dict)
            and value.get("removed") == plan["owned_paths"]
            and value.get("owned_paths_absent") is True
            and value.get("accessible_process_references") == []
            and value.get("patch_replay_temporary_directories_absent") is True,
            "cleanup record differs")
    completed = parse_time(value.get("completed_utc"), "cleanup.completed_utc")
    all_records = records + (extra_records or [])
    require(all(completed > record["end"] for record in all_records),
            "cleanup completed before the last recorded command")
    for path in plan["owned_paths"]:
        target = Path(path)
        require(not target.exists() and not target.is_symlink(), f"owned target remains: {path}")
    retained = value.get("retained_binary_sha256_before_removal")
    require(isinstance(retained, dict), "cleanup retained binary map is missing")
    all_binaries = {stage: dict(stage_binaries)
                    for stage, stage_binaries in binaries.items()}
    for stage, stage_binaries in (extra_binaries or {}).items():
        all_binaries.setdefault(stage, {}).update(stage_binaries)
    expected_retained = {
        identity["path"]: identity["sha256"]
        for stage_binaries in all_binaries.values()
        for identity in stage_binaries.values()
    }
    require(retained == expected_retained,
            "cleanup retained binary map differs from all measured binaries")
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
    allocation_correction = validate_allocation_status_correction()
    frozen, frozen_created = frozen_inputs(plan)
    adrs = adr_manifest()
    validate_protocol()
    revision = revision_manifest(plan["revision"])
    roots = plan["_roots"]
    harness = harness_paths(plan)
    names = stage_names()
    failed_names = failed_stage_names()
    baseline_manifest = source_manifest(HERE / "baseline/source-manifest.json")
    stage_data: dict[str, dict[str, Any]] = {}
    for name in names:
        stage_data[name] = validate_stage_sources(
            name, plan, revision, roots, harness,
            baseline_manifest if name != "baseline" else None,
            final_candidate=False,
        )
    failed_data: dict[str, dict[str, Any]] = {}
    for name in failed_names:
        failed_data[name] = validate_failed_stage_sources(
            name, plan, revision, roots, harness,
            baseline_manifest if name.startswith("failed-candidate-") else None,
        )
    # Receipts bind to every retained source manifest, including failed
    # historical attempts.  They remain outside the measured stage map used
    # by ABBA and analyzer admission.
    manifest_data = {**stage_data, **failed_data}
    commands_data = quality_commands()
    metadata_data = expected_actions(plan, commands_data)[1]
    failed_records_by_stage: dict[str, list[dict[str, Any]]] = {}
    failed_records: list[dict[str, Any]] = []
    for name in failed_names:
        records = validate_stage_receipts(
            name, plan, commands_data, metadata_data, manifest_data,
            allow_inflight=precleanup,
        )
        failed_records_by_stage[name] = records
        failed_records.extend(records)
    require(failed_names, "baseline correction has no retained failed attempt")
    failed_baselines = [name for name in failed_names
                        if name.startswith("failed-baseline-")]
    require(failed_baselines, "baseline correction has no failed baseline attempt")
    correction_envelope = read_json(BASELINE_FROZEN, "baseline-frozen-inputs.json")
    correction_name = correction_envelope.get("failed_attempt") \
        if isinstance(correction_envelope, dict) else None
    require(correction_name in failed_baselines,
            "baseline correction names an unretained failed baseline attempt")
    correction = baseline_freeze_data(
        plan, frozen_created, baseline_manifest,
        correction_name, failed_records_by_stage[correction_name],
    )
    _, candidate_frozen_created = candidate_frozen_inputs(plan, baseline_manifest, frozen_created)
    require(candidate_frozen_created > parse_time(
        correction["created_utc"], "baseline-frozen-inputs.created_utc"
    ), "candidate freeze does not follow corrected baseline freeze")
    conditional_freeze = conditional_frozen_inputs(plan, candidate_frozen_created)
    differential = validate_differential_tests(plan, revision, baseline_manifest)
    validate_harness_identity(stage_data, harness, revision, plan)
    if stage_data["baseline"]["changed_from_revision"]:
        require(all(test_only_path(path, plan)
                    for path in stage_data["baseline"]["changed_from_revision"]),
                "baseline source contains a production candidate change")
    for name in names:
        if name != "baseline":
            require(all(path_in_roots(path, roots)
                        for path in stage_data[name]["changed_from_baseline"]),
                    f"{name} source changes outside candidate roots")
    next_candidate = validate_next_candidate(stage_data["candidate"], plan)

    allowed, metadata = expected_actions(plan, commands)
    records_by_stage: dict[str, list[dict[str, Any]]] = {}
    all_records: list[dict[str, Any]] = list(failed_records)
    for name in names:
        records = validate_stage_receipts(name, plan, commands, metadata, manifest_data,
                                          allow_inflight=precleanup)
        records_by_stage[name] = records
        all_records.extend(records)
    validate_intervals(all_records)
    conditional_abba = validate_conditional_abba(records_by_stage, plan)
    require(all(record["start"] > frozen_created for record in all_records),
            "a command receipt predates frozen inputs")
    require(all(record["start"] > candidate_frozen_created
                for record in all_records if record["stage"] == "candidate"
                or (record["stage"] == "baseline" and ("-r2-" in record["name"]
                    or record["name"].startswith(("profile-", "hardware-", "eager-",
                                                   "build-eager-", "symbols"))))),
            "a candidate or retained baseline receipt predates candidate frozen inputs")
    conditional_records = [
        record for record in all_records
        if record["name"].startswith(("profile-", "hardware-", "eager-",
                                      "build-eager-"))
        or record["name"] == "symbols"
    ]
    if conditional_records:
        require(conditional_freeze.get("status") == "pass",
                "conditional receipts exist without a valid conditional freeze")
        conditional_created = parse_time(
            conditional_freeze["created_utc"],
            "conditional-frozen-inputs.created_utc",
        )
        require(all(record["start"] > conditional_created
                    for record in conditional_records),
                "a conditional receipt predates conditional frozen inputs")
    allocation_records = [record for record in all_records if record["name"].startswith(("alloc-", "guard-alloc-"))]
    require(all(record["start"] > gates["_created"] for record in allocation_records),
            "an allocation capture predates allocation gates")
    binaries: dict[str, dict[str, Any]] = {}
    for name, records in records_by_stage.items():
        binaries[name] = validate_binary_set(name, plan, records)
    symbol_observation = validate_symbol_observation(plan, records_by_stage, binaries)
    # The strict ABBA matrix only applies once both named source stages have
    # all four capture lanes.  A partial early campaign remains incomplete.
    abba = validate_abba(records_by_stage, plan)
    guard_reports = validate_guard_reports(records_by_stage, plan, binaries)
    analyses = replay_analyses(plan, stage_data, records_by_stage, binaries)
    analyses["hardware_review"] = replay_hardware_review(analyses)
    if not DECISION.exists():
        raise Pending("decision.json is missing")
    decision = validate_decision(stage_data, records_by_stage, analyses, commands, plan)
    if decision["decision"] == "accepted":
        require(conditional_abba["profile"].get("status") == "pass"
                and conditional_abba["eager"].get("status") == "pass",
                "accepted decision lacks complete conditional ABBA captures")
    candidate_changed = set(stage_data["candidate"]["changed_from_baseline"])
    cap_boundary = validate_cap_boundary(
        plan, baseline_manifest, stage_data["candidate"]["manifest"], candidate_changed,
        revision, stage_data, harness, candidate_frozen_created, decision,
    )
    analyses["cap"] = cap_boundary["analysis"]
    cap_next_candidate = validate_cap_next_provenance(
        plan, stage_data["candidate"]["manifest"], candidate_frozen_created,
        cap_boundary,
    )
    adverse_review = validate_adverse_review(analyses, decision)
    # The selected final source is checked after decision semantics are known.
    final_name = decision["final_stage"]
    if final_name != "baseline":
        validate_stage_sources(final_name, plan, revision, roots, harness, baseline_manifest,
                               final_candidate=True)
    source = current_source(final_name, stage_data, plan, decision)
    cap_records = cap_boundary.pop("_records")
    validate_intervals(all_records + cap_records)
    if CLEANUP.exists():
        cleanup: dict[str, Any] = validate_cleanup(
            plan, binaries, all_records,
            extra_binaries=cap_boundary["binaries"], extra_records=cap_records,
        )
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
    return {
        "status": status,
        "scope": "0543 OOXML/XLSX shared worksheet traversal follow-up; ODF deferred",
        "decision": decision,
        "stages": names,
        "failed_attempts": [
            {
                "stage": name,
                "source_manifest_sha256": failed_data[name]["manifest_sha256"],
                "receipts": len(failed_records_by_stage[name]),
                "failed_receipts": sum(
                    record["exit_code"] != 0
                    for record in failed_records_by_stage[name]
                ),
                "differential_tests": failed_data[name].get("differential_tests"),
            }
            for name in failed_names
        ],
        "receipts": len(all_records),
        "failed_attempt_receipts": sum(record["exit_code"] != 0 for record in all_records),
        "abba": abba,
        "conditional_abba": conditional_abba,
        "guard_reports": guard_reports,
        "symbol_observation": symbol_observation,
        "analyses": analyses,
        "adverse_review": adverse_review,
        "next_candidate": next_candidate,
        "cap_next_candidate": cap_next_candidate,
        "differential_tests": differential,
        "baseline_correction": correction,
        "allocation_status_correction": allocation_correction,
        "conditional_freeze": conditional_freeze,
        "cap_boundary": cap_boundary,
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
            "scope": "0543 OOXML/XLSX shared worksheet traversal follow-up; ODF deferred",
        }
    except EvidenceError as error:
        output = {
            "status": "fail", "reason": str(error),
            "scope": "0543 OOXML/XLSX shared worksheet traversal follow-up; ODF deferred",
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
