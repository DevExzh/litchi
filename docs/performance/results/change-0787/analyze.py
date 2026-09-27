"""Offline paired replay and analysis for the 0787 cached-Part packet.

This module is intentionally independent of the benchmark executable.  It
only consumes retained receipts and JSON reports, verifies their custody, and
derives deterministic summaries.  Absolute paths in receipts are treated as
historical paths; once the owned worktree and target have been removed they
are resolved relative to this packet.
"""

from __future__ import annotations

import csv
import hashlib
import io
import json
import math
import random
import re
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable, Iterator


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
ORIGIN_PATH = PACKET / "origin.json"
PLAN_PATH = PACKET / "plan.json"
REPORT_SCHEMA = "litchi.execution-baseline.v1"
ANALYSIS_SCHEMA = "litchi.cached-part-scheduling-analysis.v1"
BOOTSTRAP_SEED = 787078
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95
ROUTES = ("parts",)
SHAPES = ("small", "large", "mixed")
STATES = ("fresh", "primed")
WIDTHS = (1, 2, 4, 8, 32)
FLOORS = (0, 65536)
LEGS = ("before", "after")
CANDIDATE_FILES = (
    "crates/litchi-opc/src/source_backed.rs",
    "crates/litchi-opc/src/source_backed/batch.rs",
    "crates/litchi-opc/tests/source_backed_batch.rs",
)
NATIVE_BLOCKS = 6
OBSERVER_BLOCKS = 2
QUALIFICATION_BLOCKS = 1
NATIVE_SAMPLES = 30
OBSERVER_SAMPLES = 2
QUALIFICATION_SAMPLES = 1
EXPECTED_NATIVE_REPORTS = NATIVE_BLOCKS * 60 * 2
EXPECTED_NATIVE_SAMPLES = EXPECTED_NATIVE_REPORTS * NATIVE_SAMPLES
EXPECTED_OBSERVER_REPORTS = OBSERVER_BLOCKS * 60 * 2
EXPECTED_OBSERVER_SAMPLES = EXPECTED_OBSERVER_REPORTS * OBSERVER_SAMPLES
EXPECTED_QUALIFICATION_REPORTS_PER_LEG = 60
EXPECTED_QUALIFICATION_SAMPLES_PER_LEG = 60
EXPECTED_QUALIFICATION_REPORTS = EXPECTED_QUALIFICATION_REPORTS_PER_LEG * 2
EXPECTED_QUALIFICATION_SAMPLES = EXPECTED_QUALIFICATION_SAMPLES_PER_LEG * 2
PRODUCTION_PATHSPEC = (
    "crates", "Cargo.toml", "clippy.toml", ".cargo/config.toml",
    "rust-toolchain.toml",
)
EXPECTED_PRODUCTION_FILES = 9196
EXPECTED_ARCHITECTURE_FILES = 35
FROZEN_INPUTS = (
    "plan.json", "adoption-policy.json", "build.py", "capture.py", "quality.py",
    "custody.py", "architecture-inputs.json", "host.json", "origin.json",
    "profile-plan.json",
)
QUALITY_COMMAND_COUNT = 6
HEX = frozenset("0123456789abcdefABCDEF")


class ReplayError(RuntimeError):
    """Raised for missing, stale, malformed, or contradictory evidence."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def is_revision(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 40 and all(c in HEX for c in value)


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def finite_number(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def origin() -> dict[str, Any]:
    value = read_json(ORIGIN_PATH)
    require(isinstance(value, dict), "origin.json is malformed")
    require(is_revision(value.get("base")), "origin base revision is invalid")
    require(isinstance(value.get("owned_worktree"), str)
            and value["owned_worktree"], "origin owned worktree is missing")
    return value


def _owned_root() -> Path:
    return Path(origin()["owned_worktree"]).resolve()


def _path_candidates(raw: Path) -> list[Path]:
    candidates: list[Path] = []
    if raw.is_absolute():
        try:
            candidates.append(PACKET / raw.resolve().relative_to(_owned_root()))
        except ValueError:
            pass
        text = raw.as_posix()
        marker = "/docs/performance/results/change-0787/"
        if marker in text:
            candidates.append(PACKET / text.split(marker, 1)[1])
        candidates.append(raw)
    else:
        text = raw.as_posix()
        prefix = "docs/performance/results/change-0787/"
        if text.startswith(prefix):
            candidates.append(PACKET / text[len(prefix):])
        candidates.extend((PACKET / raw, ROOT / raw))
    unique: list[Path] = []
    for candidate in candidates:
        candidate = candidate.resolve(strict=False)
        if candidate not in unique:
            unique.append(candidate)
    return unique


def resolve_path(value: Any, *, packet_bound: bool = True) -> Path:
    require(isinstance(value, str) and value, f"invalid artifact path: {value!r}")
    candidates = _path_candidates(Path(value))
    for candidate in candidates:
        if candidate.is_file() and not candidate.is_symlink():
            if packet_bound:
                try:
                    candidate.relative_to(PACKET.resolve())
                except ValueError:
                    continue
            return candidate
    fallback = candidates[0] if candidates else Path(value).resolve(strict=False)
    if packet_bound:
        try:
            fallback.relative_to(PACKET.resolve())
        except ValueError:
            fail(f"artifact path escaped packet: {value}")
    return fallback


def artifact(value: Any, label: str, *, packet_bound: bool = True,
             allow_missing: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    size = value.get("bytes", value.get("size"))
    nonnegative_int(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256 is invalid")
    path = resolve_path(raw, packet_bound=packet_bound)
    if not path.is_file():
        if allow_missing:
            return None
        fail(f"missing {label}: {raw}")
    require(not path.is_symlink(), f"{label} is a symlink: {raw}")
    require(path.stat().st_size == size, f"{label}.bytes changed")
    require(sha256(path) == digest, f"{label}.sha256 changed")
    return path


def artifact_path(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    path = artifact(value, label, packet_bound=packet_bound)
    assert path is not None
    return path


def file_identity(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def expected_cases() -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for route in ROUTES:
        for shape in SHAPES:
            for floor in FLOORS:
                for workers in WIDTHS:
                    result.append({
                        "route": route, "shape": shape, "state": "fresh",
                        "task_floor": floor, "workers": workers,
                    })
            for floor in FLOORS:
                for workers in WIDTHS:
                    result.append({
                        "route": route, "shape": shape, "state": "primed",
                        "task_floor": floor, "workers": workers,
                    })
    return result


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(set(plan) == {"schema", "scope", "affinity", "cases", "bootstrap",
                          "native", "observer", "qualification"},
            "plan fields changed")
    require(plan.get("schema") == "litchi.cached-part-scheduling.0787.v1", "plan schema changed")
    require(plan.get("cases") == expected_cases(), "frozen case order or cardinality changed")
    require(plan.get("affinity") == list(range(32)), "capture affinity changed")
    require(plan.get("bootstrap") == {
        "confidence": BOOTSTRAP_CONFIDENCE,
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
    }, "bootstrap contract changed")
    expected_lanes = {
        "native": {"blocks": NATIVE_BLOCKS, "samples": NATIVE_SAMPLES, "warmup": 3,
                    "orders": [["before", "after"], ["after", "before"],
                               ["before", "after"], ["after", "before"],
                               ["after", "before"], ["before", "after"]]},
        "observer": {"blocks": OBSERVER_BLOCKS, "samples": OBSERVER_SAMPLES, "warmup": 0,
                      "orders": [["before", "after"], ["after", "before"]]},
        "qualification": {"blocks": QUALIFICATION_BLOCKS,
                           "samples": QUALIFICATION_SAMPLES, "warmup": 0},
    }
    for lane, expected in expected_lanes.items():
        value = plan.get(lane)
        require(isinstance(value, dict), f"{lane} lane is missing")
        for key, wanted in expected.items():
            require(value.get(key) == wanted, f"{lane} {key} configuration changed")
    require(isinstance(plan.get("scope"), str) and plan["scope"], "plan scope is missing")
    changed_path = PACKET / "candidate/changed-files.json"
    require(changed_path.is_file() and not changed_path.is_symlink(),
            "candidate changed-file manifest is missing")
    changed = read_json(changed_path)
    require(changed == list(CANDIDATE_FILES), "candidate changed-file manifest changed")
    return plan


def _git_output(arguments: list[str], label: str, *, binary: bool = False) -> bytes | str:
    try:
        completed = subprocess.run(["git", *arguments], cwd=ROOT, check=True,
                                   stdout=subprocess.PIPE)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {label}: {error}")
    return completed.stdout if binary else completed.stdout.decode()


def _git_names(revision: str) -> list[str]:
    raw = _git_output(["ls-tree", "-r", "-z", "--name-only", revision,
                       "--", *PRODUCTION_PATHSPEC], "base production census", binary=True)
    assert isinstance(raw, bytes)
    names = [item.decode() for item in raw.split(b"\0") if item]
    require(len(names) == EXPECTED_PRODUCTION_FILES,
            f"base production census changed: {len(names)}")
    require(len(set(names)) == len(names), "base production census contains duplicates")
    return names


def _git_blob_bytes(revision: str, names: Iterable[str], label: str) -> dict[str, bytes]:
    ordered = list(names)
    request = b"".join(f"{revision}:{name}\n".encode() for name in ordered)
    try:
        completed = subprocess.run(["git", "cat-file", "--batch"], cwd=ROOT,
                                   input=request, stdout=subprocess.PIPE, check=True)
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read {label} Git blobs: {error}")
    data = completed.stdout
    offset = 0
    result: dict[str, bytes] = {}
    for name in ordered:
        end = data.find(b"\n", offset)
        require(end >= 0, f"{label} blob header is truncated: {name}")
        header = data[offset:end].split()
        require(len(header) == 3 and header[1] == b"blob", f"{label} blob is invalid: {name}")
        try:
            size = int(header[2])
        except ValueError:
            fail(f"{label} blob size is invalid: {name}")
        offset = end + 1
        require(size >= 0 and offset + size <= len(data), f"{label} blob is truncated: {name}")
        result[name] = data[offset:offset + size]
        offset += size
        require(data[offset:offset + 1] == b"\n", f"{label} blob terminator is missing: {name}")
        offset += 1
    require(offset == len(data), f"{label} blob batch has trailing data")
    return result


def load_architecture_inputs() -> dict[str, Any]:
    value = read_json(PACKET / "architecture-inputs.json")
    require(isinstance(value, dict) and len(value) == EXPECTED_ARCHITECTURE_FILES,
            "architecture input cardinality changed")
    base = origin()["base"]
    recorded: dict[str, str] = {}
    for name, digest in value.items():
        require(isinstance(name, str) and name and not name.startswith("/"),
                "architecture input path is invalid")
        require(is_sha(digest), f"architecture input digest is invalid: {name}")
        recorded[name] = digest
    base_bytes = _git_blob_bytes(base, sorted(recorded), "architecture inputs")
    base_hashes = {name: hashlib.sha256(data).hexdigest() for name, data in base_bytes.items()}
    require(base_hashes == recorded, "architecture input origin Git blobs changed")
    return {
        "count": len(recorded), "revision": base, "files": recorded,
        "base_git_blob_hashes_match": True,
        "receipt": file_identity(PACKET / "architecture-inputs.json"),
    }


def load_source_manifest(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is malformed")
    revision = value.get("revision")
    require(is_revision(revision), f"{label}.revision is invalid")
    files = value.get("files")
    require(isinstance(files, dict) and files, f"{label}.files is missing")
    for name, digest in files.items():
        require(isinstance(name, str) and name and is_sha(digest),
                f"{label} contains an invalid file digest")
    return {"revision": revision, "files": dict(files)}


def normalise_tool_manifest(value: Any, revision: str, label: str) -> dict[str, Any]:
    if isinstance(value, dict) and "files" in value:
        return load_source_manifest(value, label)
    require(isinstance(value, dict) and value, f"{label} is missing")
    require(all(isinstance(name, str) and name and is_sha(digest)
                for name, digest in value.items()), f"{label} contains an invalid digest")
    return {"revision": revision, "files": dict(value)}


def _cleanup_has(cleanup: Any, receipt: dict[str, Any]) -> bool:
    expected = (receipt.get("path"), receipt.get("bytes", receipt.get("size")),
                receipt.get("sha256", receipt.get("digest")))
    if not (isinstance(expected[0], str) and isinstance(expected[1], int)
            and expected[1] >= 0 and is_sha(expected[2])):
        return False
    if isinstance(cleanup, dict):
        actual = (cleanup.get("path"), cleanup.get("bytes", cleanup.get("size")),
                  cleanup.get("sha256", cleanup.get("digest")))
        if actual == expected:
            return True
        return any(_cleanup_has(value, receipt) for value in cleanup.values())
    if isinstance(cleanup, list):
        return any(_cleanup_has(value, receipt) for value in cleanup)
    return False


def load_cleanup() -> tuple[Any, bool]:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None, False
    value = read_json(path)
    require(isinstance(value, dict), "cleanup.json is malformed")
    return value, value.get("verified") is True or value.get(
        "executables_verified_before_removal") is True


def validate_binary(receipt: Any, label: str, cleanup: Any, verified: bool) -> dict[str, Any]:
    require(isinstance(receipt, dict), f"{label} receipt is missing")
    path = artifact(receipt, label, packet_bound=False, allow_missing=True)
    if path is None:
        require(verified and _cleanup_has(cleanup, receipt),
                f"{label} is missing without an exact cleanup witness")
    # Keep the derived analysis stable across the final target cleanup.  The
    # live-vs-cleanup distinction is checked above but is deliberately not
    # serialized: otherwise replay after removing the target would change
    # analysis.json even though the retained binary receipt is identical.
    return {"path": receipt["path"], "bytes": receipt["bytes"],
            "sha256": receipt["sha256"], "custody_verified": True}


def _normalise_token(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    for prefix in (str(_owned_root()), str(ROOT.resolve())):
        value = value.replace(prefix, str(ROOT.resolve()))
    return value


def normalise_command(command: Any) -> list[Any]:
    require(isinstance(command, list), "command receipt is malformed")
    return [_normalise_token(item) for item in command]


def load_frozen_inputs(directory: Path, label: str) -> dict[str, str]:
    path = directory / "frozen-inputs.json"
    value = read_json(path)
    require(isinstance(value, dict) and set(value) == set(FROZEN_INPUTS),
            f"{label} frozen input set changed")
    for name, digest in value.items():
        require(is_sha(digest), f"{label} frozen input digest is invalid: {name}")
        current = PACKET / name
        require(current.is_file() and sha256(current) == digest,
                f"{label} frozen input changed: {name}")
    return dict(value)


def current_tool_source() -> dict[str, str]:
    tool = ROOT / "tools/perf-execution"
    require(tool.is_dir(), "standalone tool source is missing")
    result: dict[str, str] = {}
    for path in sorted(tool.rglob("*")):
        if path.is_file() and "target" not in path.parts:
            result[str(path.relative_to(ROOT))] = sha256(path)
    require(result, "standalone tool source is empty")
    return result


def _base_production_files(base: str) -> dict[str, str]:
    names = _git_names(base)
    base_bytes = _git_blob_bytes(base, names, "base production census")
    return {name: hashlib.sha256(data).hexdigest() for name, data in base_bytes.items()}


def load_build(leg: str, plan: dict[str, Any], base_files: dict[str, str],
               cleanup: Any, cleanup_verified: bool) -> dict[str, Any]:
    require(leg in LEGS, f"invalid build leg: {leg}")
    build_path = PACKET / f"build-{leg}" / "build.json"
    build = read_json(build_path)
    require(isinstance(build, dict), f"{leg} build.json is malformed")
    source_path = artifact_path(build.get("source"), f"{leg} build source")
    frozen_path = artifact_path(build.get("frozen_inputs"), f"{leg} frozen inputs")
    require(frozen_path == (source_path.parent / "frozen-inputs.json").resolve(),
            f"{leg} frozen input artifact path changed")
    source_value = read_json(source_path)
    require(isinstance(source_value, dict), f"{leg} build source manifest is malformed")
    production = load_source_manifest(source_value.get("production"),
                                      f"{leg} build production source")
    tool = normalise_tool_manifest(source_value.get("tool"), production["revision"],
                                   f"{leg} build tool source")
    base = origin()["base"]
    require(production["revision"] == base, f"{leg} production revision changed")
    if leg == "before":
        require(production["files"] == base_files,
                "before production census does not match base Git blobs")
    require(len(production["files"]) == EXPECTED_PRODUCTION_FILES,
            f"{leg} production census cardinality changed")
    frozen = load_frozen_inputs(source_path.parent, f"{leg} build")
    require(frozen["plan.json"] == sha256(PLAN_PATH), "plan changed after freeze")
    require(frozen["adoption-policy.json"] == sha256(PACKET / "adoption-policy.json"),
            f"{leg} adoption policy changed after freeze")
    rows = build.get("rows")
    require(isinstance(rows, list) and len(rows) == 2,
            f"{leg} build row cardinality changed")
    expected_manifest = str((ROOT / "tools/perf-execution/Cargo.toml").resolve())
    expected_commands = {
        "native": ["cargo", "build", "--offline", "--locked", "--release",
                    "--manifest-path", expected_manifest],
        "observer": ["cargo", "build", "--offline", "--locked", "--release",
                      "--manifest-path", expected_manifest, "--features", "source-metrics"],
    }
    logs: dict[str, str] = {}
    for row in rows:
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"{leg} build command failed")
        command = normalise_command(row.get("command"))
        kind = "observer" if "source-metrics" in command else "native"
        require(command == expected_commands[kind], f"{leg} {kind} build command changed")
        log = artifact_path(row.get("log"), f"{leg} {kind} build log")
        logs[kind] = rel(log)
    require(set(logs) == {"native", "observer"}, f"{leg} build command set changed")
    env = build.get("environment")
    require(isinstance(env, dict) and env.get("CARGO_BUILD_JOBS") == "2"
            and env.get("CARGO_INCREMENTAL") == "0", f"{leg} build environment changed")
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "observer"},
            f"{leg} build binary map changed")
    checked_binaries = {
        name: validate_binary(value, f"{leg} {name} binary", cleanup, cleanup_verified)
        for name, value in binaries.items()
    }
    canonical_commands = {
        "native": ["cargo", "build", "--offline", "--locked", "--release",
                    "--manifest-path", "tools/perf-execution/Cargo.toml"],
        "observer": ["cargo", "build", "--offline", "--locked", "--release",
                      "--manifest-path", "tools/perf-execution/Cargo.toml",
                      "--features", "source-metrics"],
    }
    return {
        "leg": leg, "receipt": file_identity(build_path),
        "source": {"path": rel(source_path), "production": production, "tool": tool},
        "frozen_inputs": frozen, "commands": canonical_commands, "logs": logs,
        "environment": {key: env.get(key) for key in
                         ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL")},
        "binaries": checked_binaries,
    }


QUALITY_COMMANDS = (
    ("cargo", "fmt", "-p", "litchi-opc", "--", "--check"),
    ("cargo", "check", "--offline", "--locked", "-p", "litchi-opc",
     "--all-features", "--all-targets"),
    ("cargo", "test", "--offline", "--locked", "-p", "litchi-opc",
     "--all-features", "--", "--test-threads=2"),
    ("cargo", "clippy", "--offline", "--locked", "-p", "litchi-opc",
     "--all-features", "--lib", "--", "-D", "warnings"),
    ("cargo", "doc", "--offline", "--locked", "-p", "litchi-opc",
     "--all-features", "--no-deps"),
    ("python3", "-B", "tools/check_crate_boundaries.py"),
)


def load_quality(after: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    quality = read_json(path)
    require(isinstance(quality, dict), "quality.json is malformed")
    source_path = artifact_path(quality.get("source"), "quality source")
    source = read_json(source_path)
    require(isinstance(source, dict), "quality source is malformed")
    require(set(source) == {"revision", "files"}, "quality source manifest changed")
    quality_production = load_source_manifest(source, "quality source")
    require(quality_production == after["source"]["production"],
            "quality source differs from after build source")
    rows = quality.get("rows")
    require(isinstance(rows, list) and len(rows) == QUALITY_COMMAND_COUNT,
            "quality gate cardinality changed")
    templates = [list(item) for item in QUALITY_COMMANDS]
    live_commands = templates
    commands = templates
    logs: list[str] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"quality gate {index} failed")
        require(normalise_command(row.get("command")) == list(live_commands[index]),
                f"quality gate {index} command changed")
        logs.append(rel(artifact_path(row.get("log"), f"quality gate {index} log")))
    env = quality.get("environment")
    require(isinstance(env, dict) and env.get("CARGO_BUILD_JOBS") == "2"
            and env.get("CARGO_INCREMENTAL") == "0"
            and env.get("CARGO_PROFILE_DEV_DEBUG") == "0"
            and env.get("RUSTDOCFLAGS") == "-D warnings",
            "quality environment changed")
    return {"receipt": file_identity(path), "gates": len(rows), "commands": [list(v) for v in commands],
            "logs": logs, "source": rel(source_path),
            "environment": {key: env[key] for key in
                             ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                              "CARGO_PROFILE_DEV_DEBUG", "RUSTDOCFLAGS")}}


def _candidate_archive_file(path: Path, expected: str, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    digest = sha256(path)
    require(digest == expected, f"{label} differs from source receipt")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": digest}


def load_candidate_archive(before: dict[str, Any], after: dict[str, Any]) -> dict[str, Any]:
    """Bind the planned candidate patch to archived before/after file bytes.

    The candidate worker retains ``candidate/before`` and ``candidate/files``
    plus an archive receipt.  The receipt is deliberately checked as a
    manifest of the planned diff, rather than inferred from the live checkout;
    the checkout is restored to production after the candidate run.
    """

    root = PACKET / "candidate"
    require(root.is_dir() and not root.is_symlink(), "candidate archive is missing")
    changed_manifest_path = root / "changed-files.json"
    require(changed_manifest_path.is_file() and not changed_manifest_path.is_symlink(),
            "candidate changed-file manifest is missing")
    require(read_json(changed_manifest_path) == list(CANDIDATE_FILES),
            "candidate changed-file manifest changed")
    receipt_path = root / "archive-receipt.json"
    require(receipt_path.is_file() and not receipt_path.is_symlink(),
            "candidate archive receipt is missing")
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), "candidate archive receipt is malformed")
    require(receipt.get("change") == "0787-cached-part-scheduling",
            "candidate archive change identity changed")
    require(receipt.get("base_commit") == origin()["base"],
            "candidate archive base commit changed")
    for key, wanted in (("archive_only", True), ("live_production_edited", False),
                        ("committed", False), ("cargo_or_compiler_run", False),
                        ("native_or_profiler_run", False)):
        require(receipt.get(key) is wanted, f"candidate archive {key} changed")
    require(receipt.get("changed_files") == list(CANDIDATE_FILES),
            "candidate archive changed-file set changed")
    checks = receipt.get("checks")
    live_status = checks.get("live_source_status") if isinstance(checks, dict) else None
    require(isinstance(checks, dict)
            and checks.get("git_apply_check", {}).get("status") == "pass"
            and checks.get("rustfmt", {}).get("status") == "pass"
            and isinstance(live_status, str)
            and "root-owned initial candidate applied" in live_status
            and "archive author" in live_status
            and "no live source edits" in live_status,
            "candidate archive checks changed")
    action_scope = receipt.get("action_scope")
    require(isinstance(action_scope, str)
            and "archive author" in action_scope
            and "root-owned" in action_scope
            and "recorded separately" in action_scope,
            "candidate archive action scope changed")
    before_files = before["source"]["production"]["files"]
    after_files = after["source"]["production"]["files"]
    changed = list(CANDIDATE_FILES)
    require(set(before_files) == set(after_files),
            "before/after production census file set changed")
    actual_diff = {name for name in before_files if before_files[name] != after_files[name]}
    require(actual_diff == set(changed),
            "after production source contains an unplanned change")
    require(set(changed).issubset(set(before_files) | set(after_files)),
            "candidate changed-file set escapes production census")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), "candidate archive artifact manifest is missing")
    artifact_specs = {
        "changed_files_json": root / "changed-files.json",
        "model_patch": root / "model.patch",
        "implementation_notes": root / "implementation-notes.md",
    }
    checked_artifacts: dict[str, dict[str, Any]] = {}
    for key, path in artifact_specs.items():
        spec = artifacts.get(key)
        require(isinstance(spec, dict), f"candidate archive artifact is missing: {key}")
        require(path.is_file() and not path.is_symlink(), f"candidate archive artifact is missing: {path}")
        require(spec.get("bytes") == path.stat().st_size
                and spec.get("sha256") == sha256(path),
                f"candidate archive artifact changed: {key}")
        checked_artifacts[key] = file_identity(path)
    entries = receipt.get("files")
    require(isinstance(entries, list) and len(entries) == len(CANDIDATE_FILES),
            "candidate archive file entries are missing")
    entry_by_name: dict[str, dict[str, Any]] = {}
    for entry in entries:
        require(isinstance(entry, dict) and isinstance(entry.get("path"), str),
                "candidate archive file entry is malformed")
        name = entry["path"]
        require(name not in entry_by_name, f"duplicate candidate archive file: {name}")
        entry_by_name[name] = entry
    require(list(entry_by_name) == list(CANDIDATE_FILES),
            "candidate archive file order or set changed")
    rows: list[dict[str, Any]] = []
    for name in sorted(changed):
        entry = entry_by_name[name]
        old = entry.get("before_sha256")
        new = entry.get("candidate_sha256")
        require(is_sha(old) and is_sha(new), f"candidate file hashes are invalid: {name}")
        require(before_files.get(name) == old and after_files.get(name) == new,
                f"candidate source diff does not match build manifests: {name}")
        require(old != new, f"candidate file is not changed: {name}")
        before_path = root / "before" / name
        after_path = root / "files" / name
        before_receipt = _candidate_archive_file(
            before_path, old, f"candidate before/{name}")
        after_receipt = _candidate_archive_file(
            after_path, new, f"candidate files/{name}")
        require(entry.get("before_bytes") == before_receipt["bytes"]
                and entry.get("candidate_bytes") == after_receipt["bytes"],
                f"candidate archive byte counts changed: {name}")
        rows.append({
            "name": name,
            "before": before_receipt,
            "after": after_receipt,
        })
    require(set(changed) == {
        row["name"] for row in rows
    }, "candidate archive file set is incomplete")
    for directory in (root / "before", root / "files"):
        for path in directory.rglob("*"):
            require(not path.is_symlink(), f"symlink in candidate archive: {path}")
        archived = sorted(str(path.relative_to(directory)) for path in directory.rglob("*")
                          if path.is_file())
        require(archived == sorted(CANDIDATE_FILES),
                f"candidate archive contains unplanned files: {directory}")
    return {
        "receipt": file_identity(receipt_path),
        "base_commit": receipt["base_commit"],
        "live_source_status": live_status,
        "action_scope": action_scope,
        "artifacts": checked_artifacts,
        "changed_files": rows,
        "changed_file_count": len(rows),
    }


def historical_tool_identity(before: dict[str, Any]) -> dict[str, Any]:
    """Check 0786 tool identity without importing any 0786 timings."""

    historical = PACKET.parent / "change-0786" / "analysis.json"
    require(historical.is_file() and not historical.is_symlink(),
            "historical 0786 analysis is missing")
    value = read_json(historical)
    require(isinstance(value, dict), "historical 0786 analysis is malformed")
    tool = value.get("source", {}).get("build", {}).get("tool", {})
    if isinstance(tool, dict) and isinstance(tool.get("files"), dict):
        historical_files = tool["files"]
    elif isinstance(tool, dict):
        historical_files = tool
    else:
        fail("historical 0786 tool source is missing")
    require(historical_files == before["source"]["tool"]["files"],
            "0787 tool source differs from historical 0786 tool source")
    return {
        "packet": "../change-0786/analysis.json",
        "analysis_sha256": sha256(historical),
        "tool_source_identical": True,
        "timings_pooled": False,
    }


def _case_key(case: dict[str, Any]) -> tuple[Any, ...]:
    return (case["route"], case["shape"], case["state"], case["task_floor"], case["workers"])


def _walk(value: Any, path: tuple[str, ...] = ()) -> Iterator[tuple[tuple[str, ...], Any]]:
    yield path, value
    if isinstance(value, dict):
        for key, child in value.items():
            yield from _walk(child, path + (str(key),))
    elif isinstance(value, list):
        for index, child in enumerate(value):
            yield from _walk(child, path + (str(index),))


def _norm_key(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", str(value).lower())


def _lookup(value: Any, names: Iterable[str]) -> Any:
    wanted = {_norm_key(name) for name in names}
    if isinstance(value, dict):
        for key, child in value.items():
            if _norm_key(key) in wanted:
                return child
        for key in ("config", "configuration", "case", "request", "timing", "metrics",
                    "corpus", "execution", "result", "resources", "resource_snapshots",
                    "availability", "source_metrics", "verification"):
            if key in value:
                result = _lookup(value[key], names)
                if result is not None:
                    return result
    return None


def _has_false_verification(value: Any) -> bool:
    for path, child in _walk(value):
        if not path or not isinstance(child, bool):
            continue
        key = _norm_key(path[-1])
        if child is False and any(marker in key for marker in
                                  ("verify", "valid", "parity", "release", "within",
                                   "under", "exact", "correct", "bounded", "match")):
            return True
    return False


def _verification_ok(report: dict[str, Any]) -> bool:
    candidate = report.get("verification", report.get("verifications"))
    require(candidate is not None, "report verification section is missing")
    require(not _has_false_verification(candidate), "report contains a failed verification")
    if isinstance(candidate, dict):
        require(candidate.get("ordered") is True,
                "report member order verification is missing or false")
        require(candidate.get("all_member_sha256_match") is True,
                "report member digest verification is missing or false")
        members = candidate.get("members")
        logical = candidate.get("logical_bytes")
        sequence = candidate.get("sequence_sha256")
        positive_int(members, "verified member count")
        nonnegative_int(logical, "verified logical bytes")
        require(is_sha(sequence), "verified sequence digest is invalid")
    # At least one positive verification marker must be present.  This avoids
    # accepting an empty object while allowing the probe to add named checks.
    positives = []
    for path, child in _walk(candidate):
        if isinstance(child, bool) and child is True:
            positives.append(path)
    require(positives, "report verification section has no positive checks")
    return True


def _json_fingerprint(value: Any) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def _output_fingerprint(report: dict[str, Any]) -> str:
    # Preserve only deterministic output/corpus identities.  Resource and
    # timing values are deliberately excluded from this parity witness.
    selected: list[tuple[tuple[str, ...], Any]] = []
    for path, value in _walk(report):
        if not path:
            continue
        # Sample count differs by lane (native 30, observer 2, qualification
        # 1).  One verified sample is enough for deterministic output parity;
        # including every sample would make a correct native/observer pair
        # appear different solely because of the lane sample budget.
        if path[0] == "samples" and (len(path) < 2 or path[1] != "0"):
            continue
        key = _norm_key(path[-1])
        if isinstance(value, (str, int, list, dict)) and (
                "output" in key or "corpus" in key or "member" in key
                or key.endswith("sha256") or key.endswith("digest")):
            selected.append((path, value))
    require(selected, "report has no deterministic output/corpus identity")
    return _json_fingerprint(selected)


def _number(value: Any, label: str) -> float:
    finite_number(value, label)
    require(float(value) >= 0, f"{label} is negative")
    return float(value)


def _sample_values(report: dict[str, Any], samples: list[dict[str, Any]]) -> tuple[list[float], list[float] | None, list[float]]:
    walls: list[float] = []
    cpus: list[float] = []
    logical: list[float] = []
    cpu_available = True
    for index, sample in enumerate(samples):
        wall = _lookup(sample, ("wall_ns", "wall_time_ns", "elapsed_ns", "duration_ns"))
        if wall is None:
            timing = sample.get("timing") if isinstance(sample, dict) else None
            wall = _lookup(timing, ("wall_ns", "wall_time_ns", "elapsed_ns", "duration_ns"))
        require(wall is not None, f"sample {index} wall_ns is missing")
        walls.append(_number(wall, f"sample {index} wall_ns"))
        cpu = _lookup(sample, ("cpu_ns", "cpu_time_ns", "process_cpu_ns"))
        if cpu is None:
            timing = sample.get("timing") if isinstance(sample, dict) else None
            cpu = _lookup(timing, ("cpu_ns", "cpu_time_ns", "process_cpu_ns"))
        if cpu is None:
            cpu_available = False
        else:
            cpus.append(_number(cpu, f"sample {index} cpu_ns"))
        amount = _lookup(sample, ("logical_bytes", "logical_read_bytes", "read_bytes",
                                  "work_bytes", "output_bytes", "bytes"))
        if amount is None:
            amount = _lookup(report, ("logical_bytes", "logical_read_bytes", "read_bytes",
                                      "work_bytes", "output_bytes"))
        require(amount is not None, f"sample {index} logical bytes is missing")
        logical.append(_number(amount, f"sample {index} logical bytes"))
    return walls, (cpus if cpu_available else None), logical


def _resource_ok(report: dict[str, Any], samples: list[dict[str, Any]]) -> tuple[bool, dict[str, Any]]:
    all_resources = [sample.get("resources", sample.get("resource_snapshots"))
                     for sample in samples]
    if not all_resources and isinstance(report.get("resources"), (dict, list)):
        all_resources = [report["resources"]]
    require(all_resources and all(value is not None for value in all_resources),
            "report resource snapshots are missing")
    for value in all_resources:
        require(isinstance(value, (dict, list)), "report resource snapshot is malformed")
    # Check each sample independently.  The returned detail is a compact
    # witness; raw snapshots remain bound to their report artifact.
    details: list[dict[str, Any]] = []
    for resources in all_resources:
        details.append(_resource_one(resources))
    require(all(item["ok"] for item in details), "resource snapshot budget check failed")
    return True, {"samples": len(details), "marker_hits": details[0]["marker_hits"],
                  "usage_limit_checks": sum(item["usage_limit_checks"] for item in details),
                  "final_release_observed": all(item["final_release_observed"] for item in details)}


def _resource_contract(samples: list[dict[str, Any]], label: str) -> list[dict[str, Any]]:
    """Return the stable permit/cpu-task witness used for before/after parity."""

    result: list[dict[str, Any]] = []
    for index, sample in enumerate(samples):
        resources = sample.get("resources", sample.get("resource_snapshots"))
        require(isinstance(resources, dict), f"{label} sample {index} resources are missing")
        snapshots: dict[str, dict[str, int]] = {}
        for marker in ("before_operation", "after_operation", "after_drop"):
            value = resources.get(marker)
            require(isinstance(value, dict), f"{label} sample {index} {marker} is missing")
            compact: dict[str, int] = {}
            for key in ("workers", "io_concurrency", "cpu_tasks"):
                counter = value.get(key)
                nonnegative_int(counter, f"{label} sample {index} {marker}.{key}")
                compact[key] = counter
            snapshots[marker] = compact
        released = resources.get("worker_and_io_released")
        cpu_bounded = resources.get("cpu_tasks_within_limit")
        require(released is True and cpu_bounded is True,
                f"{label} sample {index} permit/cpu-task witness failed")
        result.append({"snapshots": snapshots, "worker_and_io_released": released,
                       "cpu_tasks_within_limit": cpu_bounded})
    return result


def _resource_one(resources: Any) -> dict[str, Any]:
    keys = {_norm_key(path[-1]) for path, _ in _walk(resources) if path}
    required_markers = ("before", "after", "final", "drop", "output")
    marker_hits = {marker: any(marker in key for key in keys) for marker in required_markers}
    require(marker_hits["before"] and marker_hits["after"],
            "report resource snapshots lack before/after points")
    violations: list[str] = []
    if _has_false_verification(resources):
        violations.append("resource verification flag is false")
    limits: dict[str, float] = {}
    usages: list[tuple[str, float, tuple[str, ...]]] = []
    for path, value in _walk(resources):
        if not path or not isinstance(value, (int, float)) or isinstance(value, bool):
            continue
        key = _norm_key(path[-1])
        if (any(marker in key for marker in ("limit", "maximum", "capacity"))
                or (len(path) > 1 and _norm_key(path[-2]) in {"limits", "limit"})):
            if float(value) < 0:
                violations.append("negative limit at " + ".".join(path))
            else:
                limits[_norm_key(path[-1])] = float(value)
        is_snapshot_counter = len(path) > 1 and any(
            marker in _norm_key(part) for part in path[:-1]
            for marker in ("before", "after", "final", "drop", "snapshot"))
        if (any(marker in key for marker in ("used", "inuse", "consumed", "active", "outstanding"))
                or (is_snapshot_counter and key in {"workers", "ioconcurrency", "cputasks"})):
            if float(value) < 0:
                violations.append("negative usage at " + ".".join(path))
            else:
                usages.append((_norm_key(path[-1]), float(value), path))
    usage_limit_aliases = {
        "workers": "workers", "ioconcurrency": "ioconcurrency", "cputasks": "cputasks",
    }
    for usage_key, usage, usage_path in usages:
        limit_key = usage_limit_aliases.get(usage_key)
        if limit_key in limits and usage > limits[limit_key] + 1e-9:
            violations.append(f"usage exceeds limit at {'.'.join(usage_path)}")
    # Final snapshots must explicitly show released worker/I/O permits when
    # the probe publishes those counters.  A missing final zero is invalid.
    final_objects = []
    for path, value in _walk(resources):
        if path and any(marker in _norm_key(part) for part in path
                        for marker in ("final", "drop", "afteroutputs", "afterdrop")):
            if isinstance(value, dict):
                final_objects.append(value)
    released_seen = False
    for obj in final_objects:
        for key, value in obj.items():
            norm = _norm_key(key)
            if norm in {"workers", "ioconcurrency", "workerpermits", "iopermits"} \
                    and isinstance(value, (int, float)):
                released_seen = True
                if float(value) != 0:
                    violations.append(f"final permit usage is nonzero: {key}")
    return {"ok": not violations, "marker_hits": marker_hits,
            "usage_limit_checks": len(usages), "final_release_observed": released_seen,
            "violations": violations}


def _source_metrics(report: dict[str, Any]) -> dict[str, Any] | None:
    source = report.get("source_metrics", report.get("source"))
    if source is None:
        availability = report.get("availability")
        source = _lookup(availability, ("source_metrics", "source_metrics_available"))
        if isinstance(source, bool):
            return {"available": source}
    if source is None or isinstance(source, bool):
        return {"available": bool(source)} if source is not None else None
    if not isinstance(source, dict):
        return None
    availability = str(source.get("availability", ""))
    result: dict[str, Any] = {
        "available": bool(availability and "feature" in availability
                           and "unavailable" not in availability
                           and "not-applicable" not in availability),
        "availability": availability,
    }
    aliases = {
        "calls": ("calls", "read_calls", "logical_calls"),
        "requested_bytes": ("requested_bytes", "bytes_requested"),
        "returned_bytes": ("returned_bytes", "bytes_returned"),
        "short_reads": ("short_reads", "short_read_count"),
        "max_active": ("max_active", "max_concurrent", "max_reads",
                       "max_simultaneous_reads"),
        "active": ("active", "active_reads", "concurrent_reads",
                    "active_reads_after_operation"),
        "request_histogram": ("request_histogram", "request_size_histogram",
                               "request_sizes", "histogram"),
    }
    for canonical, names in aliases.items():
        value = _lookup(source, names)
        if value is not None:
            result[canonical] = value
    return result


def _contract_value(report: dict[str, Any], name: str) -> Any:
    value = _lookup(report, (name,))
    return value


def _check_corpus(report: dict[str, Any], expected: dict[str, Any], label: str) -> None:
    corpus = report.get("corpus")
    require(isinstance(corpus, dict), f"{label} corpus manifest is missing")
    require(corpus.get("shape") == expected["shape"], f"{label} corpus shape changed")
    metadata = corpus.get("metadata_members")
    require(metadata == ["[Content_Types].xml", "_rels/.rels"],
            f"{label} metadata member manifest changed")
    members = corpus.get("members")
    require(isinstance(members, list) and len(members) == 32,
            f"{label} payload member cardinality changed")
    require(corpus.get("selected_payload_member_count") == 32,
            f"{label} selected member count changed")
    expected_sizes = {
        "small": [4096] * 32,
        "large": [256 * 1024] * 32,
        "mixed": [256 * 1024] * 31 + [4096],
    }[expected["shape"]]
    for index, (member, size) in enumerate(zip(members, expected_sizes)):
        require(isinstance(member, dict), f"{label} member {index} is malformed")
        require(member.get("index") == index and member.get("bytes") == size,
                f"{label} member {index} size/order changed")
        require(is_sha(member.get("sha256")), f"{label} member {index} digest is invalid")
        require(member.get("opc_name") == f"custom/member{index:02}.bin"
                and member.get("opc_uri") == f"/custom/member{index:02}.bin"
                and member.get("cfb_name") == f"Member{index:02}",
                f"{label} member {index} name/order changed")
    require(is_sha(corpus.get("opc_sha256")) and is_sha(corpus.get("cfb_sha256")),
            f"{label} archive digest is invalid")


def parse_report(path: Path, receipt: dict[str, Any], expected: dict[str, Any],
                 lane: str, rss_kib: int) -> dict[str, Any]:
    report = read_json(path)
    require(isinstance(report, dict), f"{lane} report is malformed: {path}")
    schema = report.get("schema", report.get("schema_version"))
    require(schema == REPORT_SCHEMA, f"report schema changed: {path}")
    # The command configuration is part of the report contract.  Probe
    # versions have used a top-level config and a flattened shape; accept both
    # spellings but require every frozen case field to agree.
    for field in ("route", "shape", "state", "workers", "task_floor"):
        actual = _contract_value(report, field)
        require(actual == expected[field], f"{path} {field} differs from receipt")
    _check_corpus(report, expected, str(path))
    samples = report.get("samples")
    require(isinstance(samples, list), f"{path} samples are missing")
    expected_count = {"native": NATIVE_SAMPLES, "observer": OBSERVER_SAMPLES,
                      "qualification": QUALIFICATION_SAMPLES}[lane]
    require(len(samples) == expected_count, f"{path} sample count changed")
    config = report.get("config")
    require(isinstance(config, dict) and config.get("samples") == expected_count
            and config.get("warmup") == {"native": 3, "observer": 0, "qualification": 0}[lane]
            and config.get("aggregate_parallel_bytes") == 65536
            and config.get("cpu_task_limit") == 1_000_000,
            f"{path} benchmark configuration changed")
    metrics = report.get("metrics")
    require(isinstance(metrics, dict)
            and metrics.get("source_metrics_feature") is (lane != "native")
            and isinstance(metrics.get("cpu_clock"), str)
            and "ProcessCPUTime" in metrics["cpu_clock"],
            f"{path} metric availability changed")
    if expected["route"] == "opc":
        require(metrics.get("opc_external_read_at") == "not-applicable",
                f"{path} OPC source availability changed")
    else:
        require(metrics.get("opc_external_read_at") == "not-used",
                f"{path} source availability changed")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"{path} sample {index} is malformed")
        expected_logical = {
            "small": 4096 * 32,
            "large": 262144 * 32,
            "mixed": 262144 * 31 + 4096,
        }[expected["shape"]]
        require(_lookup(sample, ("logical_bytes",)) == expected_logical,
                f"{path} logical output bytes changed")
    for index, sample in enumerate(samples):
        require(_verification_ok(sample), f"{path} sample {index} verification failed")
    walls, cpus, logical = _sample_values(report, samples)
    resources_ok, resource_detail = _resource_ok(report, samples)
    resource_contract = _resource_contract(samples, str(path))
    output_fingerprint = _output_fingerprint(report)
    source_rows = [_source_metrics(sample) for sample in samples]
    require(all(value is not None for value in source_rows),
            f"{path} sample source metrics are missing")
    source_rows = [value for value in source_rows if value is not None]
    source_metrics: dict[str, Any] = {
        "available": all(value.get("available") is True for value in source_rows),
        "availability": source_rows[0].get("availability", ""),
    }
    for key in ("calls", "requested_bytes", "returned_bytes", "short_reads",
                "active", "max_active"):
        values = [value.get(key) for value in source_rows
                  if isinstance(value.get(key), (int, float)) and not isinstance(value.get(key), bool)]
        if values:
            source_metrics[key] = statistics.median(values)
    histograms = [value.get("request_histogram") for value in source_rows
                  if value.get("request_histogram") is not None]
    if histograms:
        source_metrics["request_histogram"] = histograms[0]
    if lane == "native":
        # The normal binary must not silently carry source-counter diagnostics.
        require(source_metrics.get("available") is False,
                f"native report carries source metrics: {path}")
    if lane in ("observer", "qualification"):
        require(source_metrics is not None, f"observer report lacks source metrics: {path}")
        require(source_metrics.get("available") is True,
                f"observer source metrics unavailable: {path}")
        for field in ("calls", "requested_bytes", "returned_bytes", "short_reads",
                      "max_active", "request_histogram"):
            require(field in source_metrics, f"observer metric is missing: {field}")
        require(source_metrics["calls"] >= 0
                and source_metrics["requested_bytes"] >= 0
                and source_metrics["returned_bytes"] >= 0
                and source_metrics["short_reads"] >= 0
                and source_metrics["returned_bytes"] <= source_metrics["requested_bytes"]
                and source_metrics["short_reads"] <= source_metrics["calls"]
                and source_metrics["max_active"] <= expected["workers"],
                f"observer source counters violate route bounds: {path}")
        histogram = source_metrics["request_histogram"]
        require(isinstance(histogram, list)
                and all(isinstance(value, int) and value >= 0 for value in histogram)
                and sum(histogram) == source_metrics["calls"]
                and source_metrics.get("active", 0) == 0
                and source_metrics["returned_bytes"] == source_metrics["requested_bytes"],
                f"observer request accounting changed: {path}")
        expected_calls = 64 if expected["state"] == "fresh" else 0
        require(source_metrics["calls"] == expected_calls,
                f"observer source call count changed: {path}")
        if expected["state"] == "primed":
            require(source_metrics["requested_bytes"] == 0
                    and source_metrics["returned_bytes"] == 0
                    and source_metrics["short_reads"] == 0
                    and source_metrics["max_active"] == 0,
                    f"primed observer source counters are not zero: {path}")
    for contract in resource_contract:
        snapshots = contract["snapshots"]
        expected_cpu = 0 if expected["state"] == "fresh" else 32
        require(snapshots["before_operation"]["cpu_tasks"] == expected_cpu
                and snapshots["after_operation"]["cpu_tasks"] == expected_cpu + 32
                and snapshots["after_drop"]["cpu_tasks"] == expected_cpu + 32,
                f"cpu-task snapshot changed: {path}")
    cpu_wall = None if cpus is None else [cpu / wall if wall else 0.0
                                          for cpu, wall in zip(cpus, walls)]
    throughputs = [amount / wall * 1_000_000_000 if wall else 0.0
                   for amount, wall in zip(logical, walls)]
    return {
        "lane": lane, "block": receipt["block"], "leg": receipt["leg"], **expected,
        "report": file_identity(path), "rss_kib": rss_kib,
        "samples": len(samples), "wall_ns": walls,
        "cpu_ns": cpus, "logical_bytes": logical,
        "p50_ns": nearest_rank(walls, 0.50),
        "p95_ns": nearest_rank(walls, 0.95),
        "p99_ns": nearest_rank(walls, 0.99),
        "mean_ns": statistics.fmean(walls),
        "cpu_wall_ratio": None if cpu_wall is None else statistics.fmean(cpu_wall),
        "throughput_bytes_s": statistics.fmean(throughputs),
        "output_fingerprint": output_fingerprint,
        "corpus_fingerprint": _json_fingerprint(report["corpus"]),
        "verification_ok": True, "resource_ok": resources_ok,
        "resource_detail": resource_detail,
        "resource_contract": resource_contract,
        "source_metrics": source_metrics,
    }


def parse_rss(path: Path) -> int:
    text = path.read_text().strip()
    require(re.fullmatch(r"[0-9]+", text) is not None, f"RSS receipt is not an integer: {path}")
    value = int(text)
    nonnegative_int(value, "RSS")
    return value


def capture_command_expected(row: dict[str, Any], binary: dict[str, Any],
                             report: dict[str, Any], rss: dict[str, Any],
                             plan: dict[str, Any]) -> list[str]:
    binary_path = _normalise_token(binary["path"])
    report_path = _normalise_token(report["path"])
    rss_path = _normalise_token(rss["path"])
    affinity = ",".join(str(item) for item in plan["affinity"])
    command = ["/usr/bin/time", "-f", "%M", "-o", rss_path, "taskset", "-c", affinity,
               binary_path]
    for key in ("route", "shape", "state", "task_floor", "workers"):
        command.extend(["--" + key.replace("_", "-"), str(row[key])])
    lane = row["lane"]
    spec_name = "qualification" if lane.startswith("qualification-") else lane
    spec = plan[spec_name]
    command.extend(["--samples", str(spec["samples"]), "--warmup", str(spec["warmup"]),
                    "--output", report_path])
    return command


def _binary_identity(value: Any) -> tuple[Any, Any, Any]:
    require(isinstance(value, dict), "binary identity is malformed")
    return value.get("path"), value.get("bytes"), value.get("sha256")


def load_lane(lane: str, plan: dict[str, Any], builds: dict[str, Any]) -> list[dict[str, Any]]:
    require(lane in ("native", "observer", "qualification-before",
                     "qualification-after"), f"invalid capture lane: {lane}")
    directory = PACKET / lane
    receipts_path = directory / "receipts.json"
    receipts = read_json(receipts_path)
    require(isinstance(receipts, list), f"{lane} receipts are malformed")
    expected_count = {
        "native": EXPECTED_NATIVE_REPORTS, "observer": EXPECTED_OBSERVER_REPORTS,
        "qualification-before": EXPECTED_QUALIFICATION_REPORTS_PER_LEG,
        "qualification-after": EXPECTED_QUALIFICATION_REPORTS_PER_LEG,
    }[lane]
    require(len(receipts) == expected_count, f"{lane} child cardinality changed")
    binary_kind = "native" if lane == "native" else "observer"
    qualification = lane.startswith("qualification-")
    logical_lane = "qualification" if qualification else lane
    fixed_leg = lane.split("-", 1)[1] if qualification else None
    expected: list[dict[str, Any]] = []
    spec = plan[logical_lane]
    for block in range(spec["blocks"]):
        order = [fixed_leg] if fixed_leg is not None else spec["orders"][block]
        cases = plan["cases"]
        for case in cases:
            for leg in order:
                expected.append({"lane": lane, "block": block, "leg": leg, **case})
    result: list[dict[str, Any]] = []
    seen_outputs: dict[tuple[Any, ...], tuple[str, str]] = {}
    for index, (row, wanted) in enumerate(zip(receipts, expected)):
        require(isinstance(row, dict), f"{lane} receipt {index} is malformed")
        require(row.get("exit_code") == 0, f"{lane} receipt {index} command failed")
        for field, value in wanted.items():
            require(row.get(field) == value, f"{lane} receipt {index} {field} changed")
        build = builds[wanted["leg"]]
        binary = build["binaries"][binary_kind]
        require(_binary_identity(row.get("binary")) == _binary_identity(binary),
                f"{lane} receipt {index} binary identity changed")
        log = artifact_path(row.get("log"), f"{lane} receipt {index} log")
        rss_path = artifact_path(row.get("rss"), f"{lane} receipt {index} RSS")
        report_path = artifact_path(row.get("report"), f"{lane} receipt {index} report")
        require(normalise_command(row.get("command")) == capture_command_expected(
            row, binary, row["report"], row["rss"], plan),
            f"{lane} receipt {index} command changed")
        rss_kib = parse_rss(rss_path)
        parsed = parse_report(report_path, row, wanted, logical_lane, rss_kib)
        group_key = (wanted["shape"], wanted["state"], wanted["task_floor"], wanted["workers"])
        previous = seen_outputs.get(group_key)
        if previous is None:
            seen_outputs[group_key] = (parsed["output_fingerprint"], parsed["corpus_fingerprint"])
        else:
            require(previous == (parsed["output_fingerprint"], parsed["corpus_fingerprint"]),
                    f"{lane} output parity changed for {group_key}")
        parsed["log"] = rel(log)
        result.append(parsed)
    complete_path = directory / "complete.json"
    complete = read_json(complete_path)
    require(isinstance(complete, dict) and complete.get("children") == expected_count,
            f"{lane} completion receipt changed")
    complete_receipts = artifact_path(complete.get("receipts"), f"{lane} completion receipts")
    require(complete_receipts == receipts_path.resolve(), f"{lane} completion receipt path changed")
    source_path = directory / "source.json"
    complete_source = artifact_path(complete.get("source"), f"{lane} completion source")
    require(complete_source == source_path.resolve(), f"{lane} completion source path changed")
    source = read_json(source_path)
    require(isinstance(source, dict), f"{lane} source census is malformed")
    source_normalized = {
        "production": load_source_manifest(source.get("production"),
                                            f"{lane} production source"),
        "tool": normalise_tool_manifest(source.get("tool"),
                                         builds["after"]["source"]["production"]["revision"],
                                         f"{lane} tool source"),
    }
    source_leg = fixed_leg or "after"
    require(source_normalized == {
        "production": builds[source_leg]["source"]["production"],
        "tool": builds[source_leg]["source"]["tool"],
    }, f"{lane} source census differs from expected build")
    return result


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    ordered = sorted(float(value) for value in values)
    require(ordered, "nearest-rank received no values")
    index = max(0, min(len(ordered) - 1, math.ceil(quantile * len(ordered)) - 1))
    return ordered[index]


def bootstrap_median(values: list[float]) -> dict[str, Any]:
    require(values, "bootstrap received no paired values")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        estimates.append(statistics.median(rng.choice(values) for _ in values))
    estimates.sort()
    # The packet freezes the exact zero-based nearest-rank endpoint indexes;
    # spelling them out avoids a platform-dependent ceil(0.025 * 10000).
    return {
        "estimate": statistics.median(values),
        "lower": estimates[249],
        "upper": estimates[9749],
        "endpoint_indexes": [249, 9749],
        "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
        "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median",
    }


PAIRED_METRICS = (
    ("p50_ns", "p50"),
    ("p95_ns", "p95"),
    ("p99_ns", "p99"),
    ("rss_kib", "rss"),
    ("cpu_wall_ratio", "cpu_wall_ratio"),
    ("throughput_bytes_s", "throughput_bytes_s"),
)


def _pair_metric(before: list[dict[str, Any]], after: list[dict[str, Any]],
                 field: str, label: str) -> dict[str, Any]:
    ratios: list[float] = []
    before_values: list[float] = []
    after_values: list[float] = []
    for index, (left, right) in enumerate(zip(before, after)):
        old = left.get(field)
        new = right.get(field)
        require(old is not None and new is not None,
                f"{label} metric is unavailable at block {index}")
        finite_number(old, f"{label} before block {index}")
        finite_number(new, f"{label} after block {index}")
        require(float(old) > 0, f"{label} before block {index} is not positive")
        require(float(new) >= 0, f"{label} after block {index} is negative")
        before_values.append(float(old))
        after_values.append(float(new))
        ratios.append(float(new) / float(old))
    confidence = bootstrap_median(ratios)
    spread = max(ratios) / min(ratios) - 1.0 if min(ratios) > 0 else None
    flags = [
        {"block": index, "ratio": ratio, "delta": ratio - 1.0,
         "absolute_over_5_percent": abs(ratio - 1.0) > 0.05,
         "regression_over_5_percent": ratio - 1.0 > 0.05}
        for index, ratio in enumerate(ratios)
    ]
    return {
        "field": field, "estimate": confidence["estimate"],
        "ci95_low": confidence["lower"], "ci95_high": confidence["upper"],
        "bootstrap": confidence, "before_block_values": before_values,
        "after_block_values": after_values, "block_ratios": ratios,
        "block_spread_ratio": spread, "block_flags": flags,
        "over_5_percent_blocks": [row["block"] for row in flags
                                   if row["absolute_over_5_percent"]],
        "regression_over_5_percent_blocks": [row["block"] for row in flags
                                              if row["regression_over_5_percent"]],
        "descriptive_only": field in {"cpu_wall_ratio", "throughput_bytes_s"},
    }


def _check_pair_invariants(old: dict[str, Any], new: dict[str, Any], label: str) -> None:
    require(old["corpus_fingerprint"] == new["corpus_fingerprint"],
            f"{label} corpus identity changed")
    require(old["output_fingerprint"] == new["output_fingerprint"],
            f"{label} output identity changed")
    require(old["logical_bytes"] == new["logical_bytes"],
            f"{label} logical output bytes changed")
    require(old["resource_contract"] == new["resource_contract"],
            f"{label} permit/cpu-task snapshots changed")
    if old["lane"] in ("observer", "qualification"):
        # max_active is a scheduler observation, so equal source work can
        # expose different maxima between paired process runs.  Each report
        # is still checked against its requested worker bound in parse_report;
        # retain both values in the observer rows instead of treating them as
        # a deterministic parity identity.
        for field in ("calls", "requested_bytes", "returned_bytes", "short_reads",
                      "active", "request_histogram"):
            require(old["source_metrics"].get(field) == new["source_metrics"].get(field),
                    f"{label} source counter changed: {field}")


def make_paired(native: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[Any, ...], dict[str, dict[str, Any]]] = {}
    for row in native:
        key = (row["route"], row["shape"], row["state"], row["task_floor"],
               row["workers"], row["block"])
        require(row["leg"] in LEGS, f"native leg changed for {key}")
        require(row["leg"] not in grouped.setdefault(key, {}),
                f"duplicate native paired leg: {key} {row['leg']}")
        grouped[key][row["leg"]] = row
    cases: dict[tuple[Any, ...], list[dict[str, Any]]] = {}
    for key, legs in grouped.items():
        require(set(legs) == set(LEGS), f"native before/after pair is incomplete: {key}")
        old, new = legs["before"], legs["after"]
        _check_pair_invariants(old, new, f"native {key}")
        case_key = key[:-1]
        cases.setdefault(case_key, []).append({"block": key[-1], "before": old, "after": new})
    require(len(cases) == 60, f"paired case cardinality changed: {len(cases)}")
    output: list[dict[str, Any]] = []
    expected_order = [_case_key(case) for case in expected_cases()]
    require(set(cases) == set(expected_order), "paired case matrix changed")
    for case_key in expected_order:
        blocks = sorted(cases[case_key], key=lambda row: row["block"])
        require([row["block"] for row in blocks] == list(range(NATIVE_BLOCKS)),
                f"paired block coverage changed: {case_key}")
        metrics: dict[str, Any] = {}
        for field, name in PAIRED_METRICS:
            metrics[name] = _pair_metric([row["before"] for row in blocks],
                                         [row["after"] for row in blocks],
                                         field, f"{case_key} {name}")
        output.append({
            "route": case_key[0], "shape": case_key[1], "state": case_key[2],
            "task_floor": case_key[3], "workers": case_key[4],
            "blocks": NATIVE_BLOCKS, "samples_per_report": NATIVE_SAMPLES,
            "metrics": metrics,
            "p50": metrics["p50"], "p95": metrics["p95"],
            "p99": metrics["p99"], "rss": metrics["rss"],
            "cpu_wall_ratio": metrics["cpu_wall_ratio"],
            "throughput_bytes_s": metrics["throughput_bytes_s"],
            "material_benefit_eligible": case_key[2] == "primed" and case_key[4] > 1,
        })
    require(len(output) == 60, "paired output cardinality changed")
    return output


def adoption_summary(paired: list[dict[str, Any]], policy: dict[str, Any]) -> dict[str, Any]:
    require(isinstance(policy, dict) and policy.get("schema") ==
            "litchi.cached-part-adoption.0787.v1", "adoption policy changed")
    latency = policy.get("latency")
    rss = policy.get("rss")
    benefit = policy.get("benefit")
    require(isinstance(latency, dict) and isinstance(rss, dict) and isinstance(benefit, dict),
            "adoption policy sections are missing")
    require(latency.get("maximum_median_paired_ratio") == 1.05
            and latency.get("reject_when_ci95_low_exceeds") == 1
            and rss.get("maximum_median_paired_ratio") == 1.05
            and rss.get("reject_when_ci95_low_exceeds") == 1
            and benefit.get("eligible_state") == "primed"
            and benefit.get("eligible_workers") == [2, 4, 8, 32]
            and benefit.get("minimum_improvement_percent") == 3
            and benefit.get("ci95_high_below") == 1,
            "adoption thresholds changed")
    require(benefit.get("at_least_one_case_required") is True,
            "benefit cardinality policy changed")
    reject: list[dict[str, Any]] = []
    benefit_rows: list[dict[str, Any]] = []
    for row in paired:
        key = {k: row[k] for k in ("route", "shape", "state", "task_floor", "workers")}
        for metric_name, section in (("p50", latency), ("rss", rss)):
            metric = row["metrics"][metric_name]
            if (metric["estimate"] > float(section["maximum_median_paired_ratio"])
                    and metric["ci95_low"] > float(section["reject_when_ci95_low_exceeds"])):
                reject.append({**key, "kind": metric_name, "reason": "median_ratio_limit",
                               "value": metric["estimate"], "ci95_low": metric["ci95_low"]})
        if (row["state"] == benefit["eligible_state"]
                and row["workers"] in benefit["eligible_workers"]):
            metric = row["metrics"]["p50"]
            threshold = 1.0 - float(benefit["minimum_improvement_percent"]) / 100.0
            if (metric["estimate"] <= threshold
                    and metric["ci95_high"] < float(benefit["ci95_high_below"])):
                benefit_rows.append({**key, "ci95_high": metric["ci95_high"],
                                     "estimate": metric["estimate"],
                                     "threshold": threshold})
    return {
        "policy": policy,
        "latency_or_rss_rejections": reject,
        "material_benefit_rows": benefit_rows,
        "material_benefit_found": bool(benefit_rows),
        "candidate_eligible_for_retention": not reject and bool(benefit_rows),
        "cpu_is_descriptive_only": True,
    }


def load_policy_interpretation() -> dict[str, Any]:
    path = PACKET / "policy-interpretation.json"
    value = read_json(path)
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.cached-part-policy-interpretation.0787.v1",
            "policy interpretation receipt changed")
    require(value.get("before_after_build_and_paired_capture") is True,
            "policy interpretation was not frozen before measurement")
    require(value.get("frozen_policy_sha256") == sha256(PACKET / "adoption-policy.json"),
            "policy interpretation is bound to a different policy")
    rejection = str(value.get("rejection", ""))
    benefit = str(value.get("benefit", ""))
    require("both" in rejection and "> 1.05" in rejection and "> 1.0" in rejection,
            "latency/RSS conjunction interpretation changed")
    require("<= 0.97" in benefit and "< 1.0" in benefit,
            "benefit interpretation changed")
    return {"receipt": file_identity(path),
            "rejection_is_conjunctive": True, "benefit_is_conjunctive": True}


def _observer_rows_0787(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for row in records:
        metrics = row.get("source_metrics") or {"available": False}
        result.append({
            "lane": row["lane"], "leg": row["leg"], "block": row["block"],
            "route": row["route"], "shape": row["shape"], "state": row["state"],
            "task_floor": row["task_floor"], "workers": row["workers"],
            "samples": row["samples"], "available": metrics.get("available", False),
            "calls": metrics.get("calls"), "requested_bytes": metrics.get("requested_bytes"),
            "returned_bytes": metrics.get("returned_bytes"),
            "short_reads": metrics.get("short_reads"),
            "max_active": metrics.get("max_active"),
            "request_histogram": json.dumps(metrics.get("request_histogram"), sort_keys=True)
            if metrics.get("request_histogram") is not None else "",
        })
    return result


def _artifact_receipts_0787(records: list[dict[str, Any]]) -> dict[str, Any]:
    return {"reports": len(records), "samples": sum(row["samples"] for row in records),
            "report_files": [row["report"] for row in records]}


def _csv_text_0787(rows: list[dict[str, Any]], fields: list[str]) -> str:
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=fields, lineterminator="\n",
                            extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue()


def _paired_csv_rows_0787(paired: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for row in paired:
        out = {k: row[k] for k in ("route", "shape", "state", "task_floor", "workers",
                                   "blocks", "samples_per_report", "material_benefit_eligible")}
        for metric_name in ("p50", "p95", "p99", "rss", "cpu_wall_ratio", "throughput_bytes_s"):
            metric = row["metrics"][metric_name]
            out[f"{metric_name}_ratio"] = metric["estimate"]
            out[f"{metric_name}_ci95_low"] = metric["ci95_low"]
            out[f"{metric_name}_ci95_high"] = metric["ci95_high"]
            out[f"{metric_name}_block_spread_ratio"] = metric["block_spread_ratio"]
            out[f"{metric_name}_block_ratios"] = json.dumps(metric["block_ratios"],
                                                              separators=(",", ":"))
            out[f"{metric_name}_over_5_percent_blocks"] = json.dumps(
                metric["over_5_percent_blocks"], separators=(",", ":"))
            out[f"{metric_name}_regression_over_5_percent_blocks"] = json.dumps(
                metric["regression_over_5_percent_blocks"], separators=(",", ":"))
        rows.append(out)
    return rows


PAIRED_FIELDS_0787 = [
    "route", "shape", "state", "task_floor", "workers", "blocks", "samples_per_report",
    "material_benefit_eligible",
]
for _metric_name in ("p50", "p95", "p99", "rss", "cpu_wall_ratio", "throughput_bytes_s"):
    PAIRED_FIELDS_0787.extend([
        f"{_metric_name}_ratio", f"{_metric_name}_ci95_low", f"{_metric_name}_ci95_high",
        f"{_metric_name}_block_spread_ratio", f"{_metric_name}_block_ratios",
        f"{_metric_name}_over_5_percent_blocks",
        f"{_metric_name}_regression_over_5_percent_blocks",
    ])


def render_markdown_0787(analysis: dict[str, Any], paired: list[dict[str, Any]]) -> str:
    benefit_policy = analysis["adoption"]["policy"]["benefit"]
    benefit_median_limit = 1.0 - float(benefit_policy["minimum_improvement_percent"]) / 100.0
    benefit_ci_limit = float(benefit_policy["ci95_high_below"])
    lines = [
        "# 0787 cached-Part scheduling paired replay",
        "",
        "This report replays retained before/after native process measurements.",
        "The route is the public ordered source-backed Part read under finite budgets.",
        "Bytes are in-memory and the primed state is a cache-hit control.",
        "CPU time and throughput ratios are descriptive diagnostics; CPU time does not infer worker count.",
        "Observer source counters are kept separate from native timing.",
        "Observer max simultaneous reads are scheduler-dependent; both leg values are retained and each is bounded by workers.",
        "",
        f"- Native reports/samples: {analysis['counts']['native_reports']} / {analysis['counts']['native_samples']}",
        f"- Observer reports/samples: {analysis['counts']['observer_reports']} / {analysis['counts']['observer_samples']}",
        f"- Qualification reports/samples: {analysis['counts']['qualification_reports']} / {analysis['counts']['qualification_samples']}",
        f"- Paired cases: {len(paired)}; six blocks per case; bootstrap seed {BOOTSTRAP_SEED}",
        f"- Candidate retention eligibility: {analysis['adoption']['candidate_eligible_for_retention']}",
        "",
        "| Shape | State | Floor | Workers | p50 ratio | p50 95% CI | p95 ratio | p99 ratio | RSS ratio | CPU ratio | Throughput ratio | Flags |",
        "|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---|",
    ]
    for row in paired:
        p50 = row["metrics"]["p50"]
        p95 = row["metrics"]["p95"]
        p99 = row["metrics"]["p99"]
        rss = row["metrics"]["rss"]
        cpu = row["metrics"]["cpu_wall_ratio"]
        throughput = row["metrics"]["throughput_bytes_s"]
        flags = []
        if p50["over_5_percent_blocks"]:
            flags.append("p50>5%")
        if rss["over_5_percent_blocks"]:
            flags.append("rss>5%")
        if p50["regression_over_5_percent_blocks"]:
            flags.append("p50-regression")
        if rss["regression_over_5_percent_blocks"]:
            flags.append("rss-regression")
        if (row["material_benefit_eligible"]
                and p50["estimate"] <= benefit_median_limit
                and p50["ci95_high"] < benefit_ci_limit):
            flags.append("benefit")
        lines.append("| " + " | ".join([
            row["shape"], row["state"], str(row["task_floor"]), str(row["workers"]),
            f"{p50['estimate']:.6g}", f"[{p50['ci95_low']:.6g}, {p50['ci95_high']:.6g}]",
            f"{p95['estimate']:.6g}", f"{p99['estimate']:.6g}",
            f"{rss['estimate']:.6g}", f"{cpu['estimate']:.6g}",
            f"{throughput['estimate']:.6g}", ",".join(flags) or "-",
        ]) + " |")
    lines.extend(["", "Block spreads for p50/p95/p99/RSS and every individual >5% ratio are retained in paired.csv.", ""])
    return "\n".join(lines)


def _pair_lane_records_0787(left: list[dict[str, Any]], right: list[dict[str, Any]],
                             label: str) -> None:
    def index(records: list[dict[str, Any]]) -> dict[tuple[Any, ...], dict[str, Any]]:
        result: dict[tuple[Any, ...], dict[str, Any]] = {}
        for row in records:
            key = (row["block"], row["shape"], row["state"], row["task_floor"],
                   row["workers"])
            require(key not in result, f"{label} duplicate key: {key}")
            result[key] = row
        return result
    before, after = index(left), index(right)
    require(set(before) == set(after), f"{label} before/after key set changed")
    for key in sorted(before):
        _check_pair_invariants(before[key], after[key], f"{label} {key}")


def build_analysis_0787() -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    plan = load_plan()
    base_files = _base_production_files(origin()["base"])
    cleanup, cleanup_verified = load_cleanup()
    builds = {leg: load_build(leg, plan, base_files, cleanup, cleanup_verified)
              for leg in LEGS}
    require(builds["before"]["source"]["tool"] == builds["after"]["source"]["tool"],
            "before/after tool source differs")
    require(builds["before"]["source"]["tool"]["files"] == current_tool_source(),
            "standalone tool source changed after build")
    historical = historical_tool_identity(builds["before"])
    candidate = load_candidate_archive(builds["before"], builds["after"])
    policy_interpretation = load_policy_interpretation()
    architecture = load_architecture_inputs()
    quality = load_quality(builds["after"])
    native = load_lane("native", plan, builds)
    observer = load_lane("observer", plan, builds)
    qualification_before = load_lane("qualification-before", plan, builds)
    qualification_after = load_lane("qualification-after", plan, builds)
    observer_before = [row for row in observer if row["leg"] == "before"]
    observer_after = [row for row in observer if row["leg"] == "after"]
    require(len(observer_before) == len(observer_after) == 120,
            "observer before/after leg cardinality changed")
    _pair_lane_records_0787(observer_before, observer_after, "observer")
    qualification = qualification_before + qualification_after
    _pair_lane_records_0787(qualification_before, qualification_after, "qualification")
    paired = make_paired(native)
    all_records = native + observer + qualification
    identities: dict[tuple[Any, ...], tuple[str, str]] = {}
    for row in all_records:
        key = (row["shape"], row["state"], row["task_floor"], row["workers"])
        identity = (row["corpus_fingerprint"], row["output_fingerprint"])
        previous = identities.setdefault(key, identity)
        require(previous == identity, f"cross-lane payload/output parity changed: {key}")
    require(len(native) == EXPECTED_NATIVE_REPORTS
            and sum(row["samples"] for row in native) == EXPECTED_NATIVE_SAMPLES,
            "native aggregate cardinality changed")
    require(len(observer) == EXPECTED_OBSERVER_REPORTS
            and sum(row["samples"] for row in observer) == EXPECTED_OBSERVER_SAMPLES,
            "observer aggregate cardinality changed")
    require(len(qualification) == EXPECTED_QUALIFICATION_REPORTS
            and sum(row["samples"] for row in qualification) == EXPECTED_QUALIFICATION_SAMPLES,
            "qualification aggregate cardinality changed")
    observer_records = observer + qualification
    require(all(row["source_metrics"]["calls"] == (64 if row["state"] == "fresh" else 0)
                for row in observer_records), "observer source call-count contract changed")
    policy = read_json(PACKET / "adoption-policy.json")
    adoption = adoption_summary(paired, policy)
    observer_rows = _observer_rows_0787(observer_records)
    require(len(observer_rows) == EXPECTED_OBSERVER_REPORTS + EXPECTED_QUALIFICATION_REPORTS,
            "observer row cardinality changed")
    output = {
        "schema": ANALYSIS_SCHEMA,
        "plan_schema": plan["schema"],
        "report_schema": REPORT_SCHEMA,
        "scope": plan["scope"],
        "host": {"host": file_identity(PACKET / "host.json")},
        "source": {
            "base_revision": origin()["base"],
            "production_file_count": EXPECTED_PRODUCTION_FILES,
            "before_production_byte_identical_to_base":
                builds["before"]["source"]["production"]["files"] == base_files,
            "after_candidate_diff_bound_to_archive": True,
            "tool_source_frozen": True,
            "before": builds["before"]["source"],
            "after": builds["after"]["source"],
            "candidate_archive": candidate,
        },
        "architecture_inputs": architecture,
        "build": builds,
        "quality": quality,
        "historical_0786": historical,
        "policy_interpretation": policy_interpretation,
        "counts": {
            "reports": len(all_records),
            "samples": sum(row["samples"] for row in all_records),
            "native_reports": len(native),
            "native_samples": sum(row["samples"] for row in native),
            "observer_reports": len(observer),
            "observer_samples": sum(row["samples"] for row in observer),
            "qualification_reports": len(qualification),
            "qualification_samples": sum(row["samples"] for row in qualification),
        },
        "lanes": {
            "native": _artifact_receipts_0787(native),
            "observer": _artifact_receipts_0787(observer),
            "qualification_before": _artifact_receipts_0787(qualification_before),
            "qualification_after": _artifact_receipts_0787(qualification_after),
        },
        "bootstrap": {
            "seed": BOOTSTRAP_SEED,
            "resamples": BOOTSTRAP_RESAMPLES,
            "confidence": BOOTSTRAP_CONFIDENCE,
            "statistic": "median",
            "endpoint_indexes": [249, 9749],
        },
        "paired": paired,
        "adoption": adoption,
        "observer": {
            "timings_pooled_with_native": False,
            "rows": observer_rows,
            "reports": len(observer_rows),
            "fresh_logical_calls": 64,
            "primed_logical_calls": 0,
            "max_active_parity": "scheduler-dependent; retained per leg and bounded by workers",
        },
        "verification": {
            "plan_and_order_checked": True,
            "receipts_checked": True,
            "report_schema_checked": True,
            "source_payload_output_parity_checked": True,
            "resource_limits_checked": True,
            "cpu_tasks_unchanged_checked": True,
            "final_permits_released_checked": True,
            "production_base_git_blobs_checked": True,
            "candidate_archive_checked": True,
            "architecture_inputs_checked": True,
            "tool_source_checked": True,
            "historical_0786_identity_checked": True,
            "quality_commands_checked_exactly": True,
            "observer_timing_separated": True,
            "observer_source_counts_checked": True,
            "observer_max_active_scheduler_dependent": True,
            "paired_bootstrap_endpoints_checked": True,
            "cpu_descriptive_only": True,
            "profile_timing_separated": True,
        },
        "reports": all_records,
    }
    return output, paired, observer_rows


def write_outputs_0787(analysis: dict[str, Any], paired: list[dict[str, Any]]) -> None:
    for name in ("analysis.json", "paired.csv", "paired.md"):
        require(not (PACKET / name).exists(), f"refusing to overwrite retained {name}")
    (PACKET / "analysis.json").write_text(json.dumps(analysis, indent=2, sort_keys=True) + "\n")
    (PACKET / "paired.csv").write_text(
        _csv_text_0787(_paired_csv_rows_0787(paired), PAIRED_FIELDS_0787)
    )
    (PACKET / "paired.md").write_text(render_markdown_0787(analysis, paired))


def check_outputs_0787(analysis: dict[str, Any], paired: list[dict[str, Any]]) -> None:
    require(read_json(PACKET / "analysis.json") == analysis,
            "analysis.json does not replay byte-for-byte")
    require((PACKET / "paired.csv").is_file()
            and (PACKET / "paired.csv").read_text() ==
            _csv_text_0787(_paired_csv_rows_0787(paired), PAIRED_FIELDS_0787),
            "paired.csv does not replay byte-for-byte")
    require((PACKET / "paired.md").is_file()
            and (PACKET / "paired.md").read_text() ==
            render_markdown_0787(analysis, paired),
            "paired.md does not replay byte-for-byte")


def analyze_0787(*, write: bool = False, check: bool = False) -> dict[str, Any]:
    result, paired, _ = build_analysis_0787()
    if write:
        write_outputs_0787(result, paired)
    if check:
        check_outputs_0787(result, paired)
    return result


# Keep the packet module's conventional import surface while retaining the
# explicit numbered entry point used by the validator.
analyze = analyze_0787


def main_0787(argv: list[str] | None = None) -> int:
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="write replay outputs")
    mode.add_argument("--check", action="store_true", help="check retained replay outputs")
    args = parser.parse_args(argv)
    try:
        analyze_0787(write=args.write or not args.check, check=args.check)
    except ReplayError as error:
        print(f"0787 replay failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main_0787())
