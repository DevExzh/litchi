"""Read-only custody verifier for the 0542 XLSX traversal pilot.

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
CANDIDATE_FROZEN = HERE / "candidate-frozen-inputs.json"
DECISION = HERE / "decision.json"
CLEANUP = HERE / "cleanup.json"
SEAL = HERE / "SHA256SUMS"
TARGET = Path("/home/zhuhe/litchi-goal-0542-target")

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
    require(value.get("status") == "frozen-before-first-allocation-capture",
            "allocation gates were not frozen before capture")
    require("planning_allocation_calls_reduction_percent" not in value,
            "allocation calls must remain diagnostic in 0542")
    for key in required - {"status", "comparison", "scope", "created_utc"}:
        _finite_nonnegative(value[key], f"allocation-gates.{key}")
        require(float(value[key]) == 1.0, f"allocation-gates.{key} must be 1.0")
    require(isinstance(value["scope"], str) and value["scope"],
            "allocation-gates.scope is invalid")
    require(value["comparison"] == "maximum candidate <= 1.01 * minimum baseline in every shape/repeat",
            "allocation comparison rule differs")
    return {**value, "_created": parse_time(value["created_utc"], "allocation-gates.created_utc")}


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
    # pilot is being assembled (the 0542 baseline was frozen before those
    # analyzers were copied into this bundle).
    required = {
        "plan.json", "run.py", "guard_run.py", "baseline_phase.py",
        "candidate_phase.py", "final_quality.py", "quality-plan.json",
        "allocation-gates.json",
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
    a later proposal, cap tests, or formatting supplement.  The candidate
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
        "candidate.patch", "cap-tests.patch", "early-fallback.patch",
        "applied-candidate.patch", "candidate-design.md",
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
    require("abba" in lowered and "allocation" in lowered and "cleanup" in lowered,
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
    patch_path = HERE / label
    data = read_bytes(patch_path, label)
    paths = patch_paths(data, label)
    require(paths == changed, f"{label} paths differ from candidate source diff")
    roots = plan_data()["_roots"]
    require(all(path_in_roots(name, roots) for name in paths),
            f"{label} contains a path outside candidate roots")
    with tempfile.TemporaryDirectory(prefix="litchi-0542-verify-candidate-", dir="/dev/shm") as temporary:
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
    if not patch_path.exists() and not note_path.exists():
        return {"status": "not-present", "admission": "unmeasured-only"}
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
    with tempfile.TemporaryDirectory(prefix="litchi-0542-verify-next-", dir="/dev/shm") as temporary:
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
    require("0542" in note and "ooxml" in note and "odf" in note,
            "next-priority.md omits pilot or format priority scope")
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
        with tempfile.TemporaryDirectory(prefix="litchi-0542-verify-", dir="/dev/shm") as temporary:
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
        result.append(path.name)
    if "baseline" not in result:
        raise Pending("baseline source attempt has not been frozen")
    if "candidate" not in result:
        raise Pending("candidate source attempt has not been frozen")
    return result


def validate_stage_sources(stage_name: str, plan: dict[str, Any], revision: dict[str, str],
                            roots: list[str], harness: set[str], baseline: dict[str, str] | None,
                            *, final_candidate: bool = False) -> dict[str, Any]:
    stage = HERE / stage_name
    require(stage.is_dir() and not stage.is_symlink(),
            f"{stage_name} is not a regular stage directory")
    snapshot = source_manifest(stage / "source-manifest.json")
    if stage_name == "baseline":
        allowed = set(harness)
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
        # The original proposal is retained for review, but it may be a
        # strict subset of the applied candidate when a later cap-test
        # patch is folded in.  The applied patch is authoritative whenever
        # a candidate source attempt exists, including a later rejection.
        draft = need(HERE / "candidate.patch", "candidate.patch")
        draft_paths = patch_paths(read_bytes(draft, "candidate.patch"), "candidate.patch")
        require(draft_paths and draft_paths <= changed_from_baseline,
                "candidate.patch is not a subset of the final candidate source diff")
        require(all(path_in_roots(name, roots) for name in draft_paths),
                "candidate.patch contains a path outside candidate roots")
        if "candidate_patch_sha256" in plan:
            require(sha(draft) == plan["candidate_patch_sha256"],
                    "candidate.patch hash differs from plan")
        need(HERE / "applied-candidate.patch", "applied-candidate.patch")
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
    expected = {name for name in snapshot if path_in_roots(name, roots) or name in harness}
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
    actual = {
        path.name for path in stage.glob(action + ".*")
        if path.is_file() and not path.is_symlink() and path.name != f"{action}.receipt.json"
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
    if stage_name == "baseline" and "-r2-" in action and (HERE / "candidate/source-manifest.json").exists():
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
        "litchi_0542_" + path.stem.replace("-", "_"),
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
    return result


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
    require(review.get("schema_version") == 2 and review.get("status") == "complete",
            "adverse review is not complete schema version 2 evidence")
    review_word = review.get("disposition")
    require(isinstance(review_word, str), "adverse review disposition is missing")
    review_normalized = review_word.strip().lower()
    expected_decision = decision["decision"]
    if expected_decision == "rejected":
        require(review_normalized in {"reject", "rejected", "rejected_and_reverted",
                                      "baseline_retained"},
                "adverse review disposition does not record rejection")
    else:
        require(review_normalized in {"accept", "accepted", "retain", "retained"},
                "adverse review disposition does not record acceptance")
    scope = review.get("scope")
    require(isinstance(scope, str) and "0542" in scope.lower()
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
    require(counts == expected_counts, "adverse review counts differ from analyzer flags")
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
    require("0542" in note and "individually" in note and "decision" in note,
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
    # outcome while the 0542 driver uses ``decision``.  Accept the aliases
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
                and conditional["eager"]["status"] == "pass",
                "accepted decision lacks successful profile and eager guard receipts")
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
    require(changed <= set(harness) | set(roots), "current source changed outside plan")
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
    if retained is not None:
        require(isinstance(retained, dict), "cleanup retained binary map is malformed")
        for stage_binaries in binaries.values():
            for identity in stage_binaries.values():
                key = identity["path"]
                require(retained.get(key) == identity["sha256"],
                        f"cleanup does not retain binary hash for {key}")
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
    _, candidate_frozen_created = candidate_frozen_inputs(plan, baseline_manifest, frozen_created)
    stage_data: dict[str, dict[str, Any]] = {}
    for name in names:
        stage_data[name] = validate_stage_sources(
            name, plan, revision, roots, harness,
            baseline_manifest if name != "baseline" else None,
            final_candidate=False,
        )
    validate_harness_identity(stage_data, harness)
    if stage_data["baseline"]["changed_from_revision"]:
        require(set(stage_data["baseline"]["changed_from_revision"]) <= harness,
                "baseline source contains a production candidate change")
    for name in names:
        if name != "baseline":
            require(all(path_in_roots(path, roots)
                        for path in stage_data[name]["changed_from_baseline"]),
                    f"{name} source changes outside candidate roots")
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
    if CLEANUP.exists():
        cleanup: dict[str, Any] = validate_cleanup(plan, binaries, all_records)
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
        "scope": "0542 OOXML/XLSX shared worksheet traversal pilot; ODF deferred",
        "decision": decision,
        "stages": names,
        "receipts": len(all_records),
        "failed_attempt_receipts": sum(record["exit_code"] != 0 for record in all_records),
        "abba": abba,
        "guard_reports": guard_reports,
        "analyses": analyses,
        "adverse_review": adverse_review,
        "next_candidate": next_candidate,
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
            "scope": "0542 OOXML/XLSX shared worksheet traversal pilot; ODF deferred",
        }
    except EvidenceError as error:
        output = {
            "status": "fail", "reason": str(error),
            "scope": "0542 OOXML/XLSX shared worksheet traversal pilot; ODF deferred",
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
