"""Offline replay and analysis for the 0786 execution-budget packet.

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
ANALYSIS_SCHEMA = "litchi.execution-scaling-analysis.v1"
BOOTSTRAP_SEED = 786078
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_CONFIDENCE = 0.95
ROUTES = ("opc", "cfb", "parts")
SHAPES = ("small", "large", "mixed")
STATES = ("fresh", "primed")
WIDTHS = (1, 2, 4, 8, 32)
FLOORS = (0, 65536)
PRODUCTION_PATHSPEC = (
    "crates", "Cargo.toml", "clippy.toml", ".cargo/config.toml",
    "rust-toolchain.toml",
)
EXPECTED_PRODUCTION_FILES = 9196
EXPECTED_ARCHITECTURE_FILES = 35
FROZEN_INPUTS = (
    "plan.json", "build.py", "capture.py", "quality.py", "custody.py",
    "architecture-inputs.json", "host.json", "cgroup-limits.json", "origin.json",
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
        marker = "/docs/performance/results/change-0786/"
        if marker in text:
            candidates.append(PACKET / text.split(marker, 1)[1])
        candidates.append(raw)
    else:
        text = raw.as_posix()
        prefix = "docs/performance/results/change-0786/"
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
                    "route": route, "shape": "large", "state": "primed",
                    "task_floor": floor, "workers": workers,
                })
    return result


def load_plan() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema") == "litchi.execution-scaling.0786.v1", "plan schema changed")
    require(plan.get("cases") == expected_cases(), "frozen case order or cardinality changed")
    require(plan.get("widths") == list(WIDTHS), "frozen worker widths changed")
    require(plan.get("affinity") == list(range(32)), "capture affinity changed")
    require(plan.get("bootstrap") == {
        "confidence": BOOTSTRAP_CONFIDENCE,
        "resamples": BOOTSTRAP_RESAMPLES,
        "seed": BOOTSTRAP_SEED,
    }, "bootstrap contract changed")
    for lane, blocks, samples, warmups, orders in (
        ("native", 6, 30, 3, ["forward", "reverse", "forward", "reverse", "reverse", "forward"]),
        ("observer", 2, 2, 0, ["forward", "reverse"]),
        ("qualification", 1, 1, 0, ["forward"]),
    ):
        value = plan.get(lane)
        require(isinstance(value, dict), f"{lane} lane is missing")
        require(value.get("blocks") == blocks and value.get("samples") == samples
                and value.get("warmup") == warmups and value.get("orders") == orders,
                f"{lane} lane configuration changed")
    require(isinstance(plan.get("scope"), str) and plan["scope"], "plan scope is missing")
    require(isinstance(plan.get("source_policy"), str) and plan["source_policy"],
            "source policy is missing")
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
    live: dict[str, str] = {}
    for name, digest in value.items():
        require(isinstance(name, str) and name and not name.startswith("/"),
                "architecture input path is invalid")
        require(is_sha(digest), f"architecture input digest is invalid: {name}")
        path = ROOT / name
        require(path.is_file() and not path.is_symlink(), f"architecture input is missing: {name}")
        actual = sha256(path)
        require(actual == digest, f"architecture input changed: {name}")
        live[name] = actual
    base_bytes = _git_blob_bytes(base, sorted(live), "architecture inputs")
    base_hashes = {name: hashlib.sha256(data).hexdigest() for name, data in base_bytes.items()}
    require(base_hashes == live, "architecture input origin Git blobs changed")
    return {
        "count": len(live), "revision": base, "files": live,
        "live_files_match": True, "origin_blob_hashes_match": True,
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


def _production_source(base: str) -> dict[str, str]:
    names = _git_names(base)
    base_bytes = _git_blob_bytes(base, names, "base production census")
    expected = {name: hashlib.sha256(data).hexdigest() for name, data in base_bytes.items()}
    for name, digest in expected.items():
        path = ROOT / name
        require(path.is_file() and not path.is_symlink(), f"production source is missing: {name}")
        require(sha256(path) == digest, f"production source differs from base Git blob: {name}")
    return expected


def load_build(plan: dict[str, Any]) -> dict[str, Any]:
    build_path = PACKET / "build.json"
    build = read_json(build_path)
    require(isinstance(build, dict), "build.json is malformed")
    source_path = artifact_path(build.get("source"), "build source")
    source_value = read_json(source_path)
    require(isinstance(source_value, dict), "build source manifest is malformed")
    production = load_source_manifest(source_value.get("production"), "build production source")
    # custody.tool_source() is intentionally a flat path -> SHA-256 map; it
    # has no Git revision because the standalone package is measured from the
    # worktree.  Accept the wrapped form too for replay compatibility with a
    # future custody helper, but normalize both to the same semantic object.
    tool = normalise_tool_manifest(source_value.get("tool"), production["revision"],
                                   "build tool source")
    base = origin()["base"]
    require(production["revision"] == base, "build production revision changed")
    expected_production = _production_source(base)
    require(production["files"] == expected_production,
            "build production census does not match base Git blobs")
    require(len(production["files"]) == EXPECTED_PRODUCTION_FILES,
            "build production census cardinality changed")
    require(tool["files"] == current_tool_source(), "tool source changed after freeze")
    frozen = load_frozen_inputs(source_path.parent, "build")
    require(frozen["plan.json"] == sha256(PLAN_PATH), "plan changed after freeze")
    rows = build.get("rows")
    require(isinstance(rows, list) and len(rows) == 2, "build row cardinality changed")
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
                "build command failed")
        command = normalise_command(row.get("command"))
        kind = "observer" if "source-metrics" in command else "native"
        require(command == expected_commands[kind], f"{kind} build command changed")
        log = artifact_path(row.get("log"), f"{kind} build log")
        logs[kind] = rel(log)
    env = build.get("environment")
    require(isinstance(env, dict) and env.get("CARGO_BUILD_JOBS") == "2"
            and env.get("CARGO_INCREMENTAL") == "0", "build environment changed")
    cleanup, cleanup_verified = load_cleanup()
    binaries = build.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"native", "observer"},
            "build binary map changed")
    checked_binaries = {
        name: validate_binary(value, f"{name} binary", cleanup, cleanup_verified)
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
        "receipt": file_identity(build_path),
        "source": {"path": rel(source_path), "production": production, "tool": tool},
        "frozen_inputs": frozen, "commands": canonical_commands, "logs": logs,
        "environment": {key: env.get(key) for key in
                         ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL")},
        "binaries": checked_binaries,
    }


QUALITY_COMMANDS = (
    ("cargo", "fmt", "--manifest-path", "{manifest}", "--", "--check"),
    ("cargo", "check", "--offline", "--locked", "--manifest-path", "{manifest}",
     "--all-features", "--all-targets"),
    ("cargo", "test", "--offline", "--locked", "--manifest-path", "{manifest}",
     "--all-features", "--", "--test-threads=2"),
    ("cargo", "clippy", "--offline", "--locked", "--manifest-path", "{manifest}",
     "--all-features", "--all-targets", "--", "-D", "warnings"),
    ("cargo", "doc", "--offline", "--locked", "--manifest-path", "{manifest}",
     "--all-features", "--no-deps"),
    ("python3", "-B", "tools/check_crate_boundaries.py"),
)


def load_quality(build: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    quality = read_json(path)
    require(isinstance(quality, dict), "quality.json is malformed")
    source_path = artifact_path(quality.get("source"), "quality source")
    source = read_json(source_path)
    require(isinstance(source, dict), "quality source is malformed")
    quality_tool = normalise_tool_manifest(source.get("tool"),
                                           build["source"]["production"]["revision"],
                                           "quality tool source")
    require(source.get("production") == build["source"]["production"]
            and quality_tool == build["source"]["tool"],
            "quality source differs from build source")
    rows = quality.get("rows")
    require(isinstance(rows, list) and len(rows) == QUALITY_COMMAND_COUNT,
            "quality gate cardinality changed")
    manifest = str((ROOT / "tools/perf-execution/Cargo.toml").resolve())
    templates = [list(item) for item in QUALITY_COMMANDS]
    live_commands = [[manifest if token == "{manifest}" else token for token in row]
                     for row in templates]
    commands = [["tools/perf-execution/Cargo.toml" if token == "{manifest}" else token
                 for token in row] for row in templates]
    logs: list[str] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"quality gate {index} failed")
        require(normalise_command(row.get("command")) == live_commands[index],
                f"quality gate {index} command changed")
        logs.append(rel(artifact_path(row.get("log"), f"quality gate {index} log")))
    env = quality.get("environment")
    require(isinstance(env, dict) and env.get("CARGO_BUILD_JOBS") == "2"
            and env.get("CARGO_INCREMENTAL") == "0"
            and env.get("CARGO_PROFILE_DEV_DEBUG") == "0"
            and env.get("RUSTDOCFLAGS") == "-D warnings",
            "quality environment changed")
    return {"receipt": file_identity(path), "gates": len(rows), "commands": commands,
            "logs": logs, "source": rel(source_path),
            "environment": {key: env[key] for key in
                             ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                              "CARGO_PROFILE_DEV_DEBUG", "RUSTDOCFLAGS")}}


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
    require(isinstance(metadata, list) and len(metadata) == 2,
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
        for field in ("opc_name", "opc_uri", "cfb_name"):
            require(isinstance(member.get(field), str) and member[field],
                    f"{label} member {index} {field} is missing")
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
    expected_count = {"native": 30, "observer": 2, "qualification": 1}[lane]
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
    for index, sample in enumerate(samples):
        require(_verification_ok(sample), f"{path} sample {index} verification failed")
    walls, cpus, logical = _sample_values(report, samples)
    resources_ok, resource_detail = _resource_ok(report, samples)
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
        if expected["route"] != "opc":
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
    cpu_wall = None if cpus is None else [cpu / wall if wall else 0.0
                                          for cpu, wall in zip(cpus, walls)]
    throughputs = [amount / wall * 1_000_000_000 if wall else 0.0
                   for amount, wall in zip(logical, walls)]
    return {
        "lane": lane, "block": receipt["block"], **expected,
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
        "verification_ok": True, "resource_ok": resources_ok,
        "resource_detail": resource_detail,
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
    spec = plan[lane]
    command.extend(["--samples", str(spec["samples"]), "--warmup", str(spec["warmup"]),
                    "--output", report_path])
    return command


def load_lane(lane: str, plan: dict[str, Any], build: dict[str, Any]) -> list[dict[str, Any]]:
    directory = PACKET / lane
    receipts_path = directory / "receipts.json"
    receipts = read_json(receipts_path)
    require(isinstance(receipts, list), f"{lane} receipts are malformed")
    expected_count = {"native": 720, "observer": 240, "qualification": 120}[lane]
    require(len(receipts) == expected_count, f"{lane} child cardinality changed")
    binary_kind = "native" if lane == "native" else "observer"
    binary = build["binaries"][binary_kind]
    expected: list[dict[str, Any]] = []
    for block, order in enumerate(plan[lane]["orders"]):
        cases = plan["cases"] if order == "forward" else list(reversed(plan["cases"]))
        for case in cases:
            expected.append({"lane": lane, "block": block, **case})
    result: list[dict[str, Any]] = []
    seen_outputs: dict[tuple[str, str], str] = {}
    for index, (row, wanted) in enumerate(zip(receipts, expected)):
        require(isinstance(row, dict), f"{lane} receipt {index} is malformed")
        for field, value in wanted.items():
            require(row.get(field) == value, f"{lane} receipt {index} {field} changed")
        require(row.get("binary") == build["binaries"][binary_kind]
                or row.get("binary") == {
                    "path": binary["path"], "bytes": binary["bytes"], "sha256": binary["sha256"]
                }, f"{lane} receipt {index} binary identity changed")
        log = artifact_path(row.get("log"), f"{lane} receipt {index} log")
        rss_path = artifact_path(row.get("rss"), f"{lane} receipt {index} RSS")
        report_path = artifact_path(row.get("report"), f"{lane} receipt {index} report")
        require(normalise_command(row.get("command")) == capture_command_expected(
            row, binary, row["report"], row["rss"], plan),
            f"{lane} receipt {index} command changed")
        rss_kib = parse_rss(rss_path)
        parsed = parse_report(report_path, row, wanted, lane, rss_kib)
        group_key = (wanted["route"], wanted["shape"])
        previous = seen_outputs.get(group_key)
        if previous is None:
            seen_outputs[group_key] = parsed["output_fingerprint"]
        else:
            require(previous == parsed["output_fingerprint"],
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
    source = read_json(source_path)
    require(isinstance(source, dict), f"{lane} source census is malformed")
    source_normalized = {
        "production": load_source_manifest(source.get("production"),
                                            f"{lane} production source"),
        "tool": normalise_tool_manifest(source.get("tool"),
                                         build["source"]["production"]["revision"],
                                         f"{lane} tool source"),
    }
    require(source_normalized == {
        "production": build["source"]["production"],
        "tool": build["source"]["tool"],
    }, f"{lane} source census differs from build")
    return result


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    ordered = sorted(float(value) for value in values)
    require(ordered, "nearest-rank received no values")
    index = max(0, min(len(ordered) - 1, math.ceil(quantile * len(ordered)) - 1))
    return ordered[index]


def percentile(values: list[float], quantile: float) -> float:
    return nearest_rank(values, quantile)


def bootstrap_median(values: list[float]) -> dict[str, Any]:
    require(values, "bootstrap received no paired values")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates: list[float] = []
    for _ in range(BOOTSTRAP_RESAMPLES):
        estimates.append(statistics.median(rng.choice(values) for _ in values))
    return {
        "estimate": statistics.median(values),
        # Keep the frozen nearest-rank endpoints as decimal constants.  The
        # expression (.05 / 2) can evaluate just above .025 on this Python
        # build, moving ceil(250.000...) to rank 251 instead of rank 250.
        "lower": percentile(estimates, 0.025),
        "upper": percentile(estimates, 0.975),
        "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
        "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median",
    }


def _amdahl_fit(rows: list[dict[str, Any]]) -> dict[str, float]:
    observations = [(float(row["workers"]), float(row["time_ratio"]))
                    for row in rows if row["workers"] > 1 and row["time_ratio"] is not None]
    if not observations:
        return {"unconstrained_serial_fraction": 0.0,
                "constrained_serial_fraction": 0.0, "rmse_unconstrained": 0.0,
                "rmse_constrained": 0.0}
    terms = [(1.0 - 1.0 / width, ratio - 1.0 / width) for width, ratio in observations]
    denominator = sum(x * x for x, _ in terms)
    unconstrained = sum(x * z for x, z in terms) / denominator if denominator else 0.0
    constrained = min(1.0, max(0.0, unconstrained))
    residual_unconstrained = [z - unconstrained * x for x, z in terms]
    residual_constrained = [z - constrained * x for x, z in terms]
    return {
        "unconstrained_serial_fraction": unconstrained,
        "constrained_serial_fraction": constrained,
        "rmse_unconstrained": math.sqrt(statistics.fmean(x * x for x in residual_unconstrained)),
        "rmse_constrained": math.sqrt(statistics.fmean(x * x for x in residual_constrained)),
    }


def make_scaling(records: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    native = [row for row in records if row["lane"] == "native"]
    grouped: dict[tuple[str, str, str, int], dict[int, dict[int, dict[str, Any]]]] = {}
    for row in native:
        key = (row["route"], row["shape"], row["state"], row["task_floor"])
        grouped.setdefault(key, {}).setdefault(row["block"], {})[row["workers"]] = row
    output: list[dict[str, Any]] = []
    fit_summary: dict[str, Any] = {}
    for key in sorted(grouped):
        blocks = grouped[key]
        require(set(blocks) == set(range(6)), f"native block coverage changed for {key}")
        per_width: dict[int, list[dict[str, Any]]] = {width: [] for width in WIDTHS}
        for block in range(6):
            require(set(blocks[block]) == set(WIDTHS), f"native width coverage changed for {key} block {block}")
            base = blocks[block][1]
            for width in WIDTHS:
                row = blocks[block][width]
                ratio = base["p50_ns"] / row["p50_ns"] if row["p50_ns"] else None
                per_width[width].append({"block": block, "row": row, "speedup": ratio,
                                         "time_ratio": row["p50_ns"] / base["p50_ns"] if base["p50_ns"] else None})
        aggregate_fit_rows = []
        for width in WIDTHS:
            entries = per_width[width]
            ratios = [entry["speedup"] for entry in entries if entry["speedup"] is not None]
            speedup = bootstrap_median(ratios)
            p50s = [entry["row"]["p50_ns"] for entry in entries]
            p95s = [entry["row"]["p95_ns"] for entry in entries]
            p99s = [entry["row"]["p99_ns"] for entry in entries]
            means = [entry["row"]["mean_ns"] for entry in entries]
            cpus = [entry["row"]["cpu_wall_ratio"] for entry in entries
                    if entry["row"]["cpu_wall_ratio"] is not None]
            throughput = [entry["row"]["throughput_bytes_s"] for entry in entries]
            rss = [entry["row"]["rss_kib"] for entry in entries]
            p50_spread = (max(p50s) / min(p50s) - 1.0) if min(p50s) else None
            p95_spread = (max(p95s) / min(p95s) - 1.0) if min(p95s) else None
            p99_spread = (max(p99s) / min(p99s) - 1.0) if min(p99s) else None
            rss_spread = (max(rss) / min(rss) - 1.0) if min(rss) else None
            over_5pct: list[str] = []
            regressions: list[str] = []
            for entry in entries:
                block = entry["block"]
                base = blocks[block][1]
                for metric in ("p50_ns", "p95_ns", "p99_ns", "rss_kib"):
                    current = entry["row"][metric]
                    reference = base[metric]
                    if not reference:
                        continue
                    delta = current / reference - 1.0
                    if abs(delta) > 0.05:
                        over_5pct.append(f"{metric}:block{block}:{delta:+.6f}")
                    if delta > 0.05:
                        regressions.append(f"{metric}:block{block}:{delta:+.6f}")
            time_ratio = statistics.median([entry["time_ratio"] for entry in entries
                                            if entry["time_ratio"] is not None])
            aggregate_fit_rows.append({"workers": width, "time_ratio": time_ratio})
            tail = statistics.median([p99 / p50 if p50 else 0.0 for p99, p50 in zip(p99s, p50s)])
            apparent = None if width == 1 else (
                (1.0 / speedup["estimate"] - 1.0 / width) / (1.0 - 1.0 / width)
                if speedup["estimate"] else None
            )
            spread_flags = [metric for metric, value in (
                ("p50", p50_spread), ("p95", p95_spread),
                ("p99", p99_spread), ("rss", rss_spread),
            ) if value is not None and value > 0.05]
            output.append({
                "route": key[0], "shape": key[1], "state": key[2], "task_floor": key[3],
                "workers": width, "blocks": len(entries), "samples_per_report": 30,
                "p50_ns": statistics.median(p50s), "p95_ns": statistics.median(p95s),
                "p99_ns": statistics.median(p99s), "mean_ns": statistics.median(means),
                "speedup": speedup["estimate"], "speedup_ci_low": speedup["lower"],
                "speedup_ci_high": speedup["upper"],
                "efficiency": speedup["estimate"] / width,
                "cpu_wall_ratio": None if not cpus else statistics.median(cpus),
                "logical_throughput_bytes_s": statistics.median(throughput),
                "rss_kib": statistics.median(rss),
                "p50_block_spread_ratio": p50_spread,
                "p95_block_spread_ratio": p95_spread,
                "p99_block_spread_ratio": p99_spread,
                "rss_block_spread_ratio": rss_spread,
                "p99_to_p50_ratio": tail,
                "spread_flag": bool(spread_flags),
                "spread_flags": spread_flags,
                "rss_spread_flag": "rss" in spread_flags,
                "tail_flag": bool(tail > 1.05),
                "over_5pct_vs_width1": over_5pct,
                "regression_vs_width1": regressions,
                "negative_scaling": bool(width > 1 and speedup["estimate"] < 1.0),
                "apparent_serial_fraction": apparent,
                "block_speedups": [entry["speedup"] for entry in entries],
                "block_p50_ns": p50s,
                "block_p95_ns": p95s,
                "block_p99_ns": p99s,
                "block_rss_kib": rss,
                "bootstrap": speedup,
            })
        fit = _amdahl_fit(aggregate_fit_rows)
        family = "/".join(map(str, key))
        fit_summary[family] = {**fit, "descriptive_only": True,
                               "interpretation": "not a causal decomposition"}
        for row in output:
            if (row["route"], row["shape"], row["state"], row["task_floor"]) == key:
                row.update({"amdahl_unconstrained_serial_fraction": fit["unconstrained_serial_fraction"],
                            "amdahl_constrained_serial_fraction": fit["constrained_serial_fraction"],
                            "amdahl_rmse_unconstrained": fit["rmse_unconstrained"],
                            "amdahl_rmse_constrained": fit["rmse_constrained"]})
    require(len(output) == 120, f"scaling row cardinality changed: {len(output)}")
    return output, fit_summary


def load_raw_audit(scaling: list[dict[str, Any]]) -> dict[str, Any]:
    """Bind the independent raw-report/corpus audit to the derived curves."""

    path = PACKET / "raw-audit.json"
    value = read_json(path)
    require(isinstance(value, dict), "raw-audit.json is malformed")
    require(value.get("reports") == 1080 and value.get("samples") == 22200
            and value.get("independently_reconstructed_payloads") == 96,
            "raw audit aggregate counts changed")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 120, "raw audit row cardinality changed")
    by_case = {(row["route"], row["shape"], row["state"], row["task_floor"], row["workers"]): row
               for row in scaling}
    require(len(by_case) == 120, "scaling case key cardinality changed")
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and isinstance(row.get("case"), list)
                and len(row["case"]) == 5, f"raw audit row {index} is malformed")
        key = tuple(row["case"])
        require(key in by_case, f"raw audit case is absent from scaling: {key}")
        derived = by_case[key]
        paired = row.get("paired_speedups")
        require(isinstance(paired, list) and len(paired) == 6,
                f"raw audit paired speeds changed: {key}")
        for actual, expected in zip(derived["block_speedups"], paired):
            require(actual == expected, f"raw audit paired speed mismatch: {key}")
        require(derived["speedup"] == row.get("median"),
                f"raw audit median mismatch: {key}")
        ci = row.get("ci95")
        require(isinstance(ci, list) and len(ci) == 2
                and derived["speedup_ci_low"] == ci[0]
                and derived["speedup_ci_high"] == ci[1],
                f"raw audit CI mismatch: {key}")
    corpora = value.get("corpora")
    identities = value.get("container_identities")
    require(isinstance(corpora, dict) and set(corpora) == set(SHAPES)
            and isinstance(identities, dict) and set(identities) == set(SHAPES),
            "raw audit corpus identity set changed")
    return {
        "receipt": file_identity(path),
        "script": file_identity(PACKET / "raw_audit.py"),
        "reports": value["reports"], "samples": value["samples"],
        "independently_reconstructed_payloads": value["independently_reconstructed_payloads"],
        "rows": len(rows), "paired_curves_match": True,
        "corpus_shapes": sorted(corpora), "container_identity_shapes": sorted(identities),
    }


def _observer_rows(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for row in records:
        metrics = row.get("source_metrics") or {"available": False}
        output = {"lane": row["lane"], "block": row["block"],
                  "route": row["route"], "shape": row["shape"], "state": row["state"],
                  "task_floor": row["task_floor"], "workers": row["workers"],
                  "samples": row["samples"], "available": metrics.get("available", False),
                  "calls": metrics.get("calls"),
                  "requested_bytes": metrics.get("requested_bytes"),
                  "returned_bytes": metrics.get("returned_bytes"),
                  "short_reads": metrics.get("short_reads"),
                  "active": metrics.get("active"), "max_active": metrics.get("max_active"),
                  "request_histogram": json.dumps(metrics.get("request_histogram"), sort_keys=True)
                  if metrics.get("request_histogram") is not None else ""}
        result.append(output)
    return result


def _artifact_receipts(records: list[dict[str, Any]]) -> dict[str, Any]:
    return {"reports": len(records), "samples": sum(row["samples"] for row in records),
            "report_files": [row["report"] for row in records]}


def _csv_text(rows: list[dict[str, Any]], fields: list[str]) -> str:
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=fields, lineterminator="\n", extrasaction="ignore")
    writer.writeheader()
    writer.writerows(rows)
    return output.getvalue()


SCALING_FIELDS = [
    "route", "shape", "state", "task_floor", "workers", "blocks", "samples_per_report",
    "p50_ns", "p95_ns", "p99_ns", "mean_ns", "speedup", "speedup_ci_low",
    "speedup_ci_high", "efficiency", "cpu_wall_ratio", "logical_throughput_bytes_s",
    "rss_kib", "p50_block_spread_ratio", "p95_block_spread_ratio",
    "p99_block_spread_ratio", "rss_block_spread_ratio", "p99_to_p50_ratio",
    "spread_flag", "rss_spread_flag", "tail_flag", "negative_scaling",
    "apparent_serial_fraction", "over_5pct_vs_width1", "regression_vs_width1",
    "amdahl_unconstrained_serial_fraction", "amdahl_constrained_serial_fraction",
    "amdahl_rmse_unconstrained", "amdahl_rmse_constrained",
]
OBSERVER_FIELDS = [
    "lane", "block", "route", "shape", "state", "task_floor", "workers", "samples",
    "available", "calls", "requested_bytes", "returned_bytes", "short_reads", "active",
    "max_active", "request_histogram",
]


def render_markdown(analysis: dict[str, Any], scaling: list[dict[str, Any]]) -> str:
    lines = [
        "# 0786 finite execution-budget scaling",
        "",
        "This report is an offline replay of the retained native, observer, and qualification receipts.",
        "The measurements use warm in-memory bytes and finite hierarchical execution budgets.",
        "Requested-width efficiency and Amdahl fits are descriptive; they do not infer active worker counts",
        "or establish a causal decomposition.",
        "",
        f"- Native reports/samples: {analysis['counts']['native_reports']} / {analysis['counts']['native_samples']}",
        f"- Observer reports/samples: {analysis['counts']['observer_reports']} / {analysis['counts']['observer_samples']}",
        f"- Qualification reports/samples: {analysis['counts']['qualification_reports']} / {analysis['counts']['qualification_samples']}",
        f"- Scaling rows: {len(scaling)}; bootstrap: {BOOTSTRAP_RESAMPLES} resamples, seed {BOOTSTRAP_SEED}",
        "",
        "| Route | Shape | State | Floor | Width | p50 ns | Speedup | Efficiency | CPU/wall | RSS KiB | Flags |",
        "|---|---|---:|---:|---:|---:|---:|---:|---:|---:|---|",
    ]
    for row in scaling:
        flags = ",".join(flag for flag, enabled in (
            ("spread", row["spread_flag"]), ("rss-spread", row["rss_spread_flag"]),
            ("tail", row["tail_flag"]), ("negative", row["negative_scaling"]),
            ("vs-width1", bool(row["regression_vs_width1"])),
        ) if enabled) or "-"
        def fmt(value: Any) -> str:
            return "n/a" if value is None else f"{value:.6g}" if isinstance(value, float) else str(value)
        lines.append("| " + " | ".join((row["route"], row["shape"], row["state"],
                                         str(row["task_floor"]), str(row["workers"]),
                                         fmt(row["p50_ns"]), fmt(row["speedup"]),
                                         fmt(row["efficiency"]), fmt(row["cpu_wall_ratio"]),
                                         fmt(row["rss_kib"]), flags)) + " |")
    lines.extend(["", "Observer source counters are retained in `observer.csv` and are never pooled into native timing.", ""])
    return "\n".join(lines)


def build_analysis() -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    plan = load_plan()
    build = load_build(plan)
    quality = load_quality(build)
    architecture = load_architecture_inputs()
    native = load_lane("native", plan, build)
    observer = load_lane("observer", plan, build)
    qualification = load_lane("qualification", plan, build)
    all_records = native + observer + qualification
    parity: dict[tuple[str, str], str] = {}
    for row in all_records:
        key = (row["route"], row["shape"])
        previous = parity.setdefault(key, row["output_fingerprint"])
        require(previous == row["output_fingerprint"],
                f"cross-lane output parity changed for {key}")
    require(len(native) == 720 and sum(row["samples"] for row in native) == 21600,
            "native aggregate cardinality changed")
    require(len(observer) == 240 and sum(row["samples"] for row in observer) == 480,
            "observer aggregate cardinality changed")
    require(len(qualification) == 120 and sum(row["samples"] for row in qualification) == 120,
            "qualification aggregate cardinality changed")
    scaling, amdahl = make_scaling(all_records)
    raw_audit = load_raw_audit(scaling)
    observer_rows = _observer_rows(observer + qualification)
    require(len(observer_rows) == 360, "observer CSV cardinality changed")
    output = {
        "schema": ANALYSIS_SCHEMA,
        "plan_schema": plan["schema"],
        "report_schema": REPORT_SCHEMA,
        "scope": plan["scope"],
        "host": {"host": file_identity(PACKET / "host.json"),
                 "cgroup_limits": file_identity(PACKET / "cgroup-limits.json")},
        "source": {
            "base_revision": origin()["base"],
            "production_file_count": EXPECTED_PRODUCTION_FILES,
            "production_byte_identical_to_base": True,
            "tool_source_frozen": True,
            "build": build["source"],
        },
        "architecture_inputs": architecture,
        "build": build,
        "quality": quality,
        "counts": {
            "reports": len(all_records), "samples": sum(row["samples"] for row in all_records),
            "native_reports": len(native), "native_samples": sum(row["samples"] for row in native),
            "observer_reports": len(observer), "observer_samples": sum(row["samples"] for row in observer),
            "qualification_reports": len(qualification),
            "qualification_samples": sum(row["samples"] for row in qualification),
        },
        "lanes": {
            "native": _artifact_receipts(native), "observer": _artifact_receipts(observer),
            "qualification": _artifact_receipts(qualification),
        },
        "bootstrap": {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
                      "confidence": BOOTSTRAP_CONFIDENCE, "statistic": "median"},
        "scaling": scaling,
        "amdahl": amdahl,
        "raw_audit": raw_audit,
        "observer": {"timings_pooled_with_native": False,
                      "rows": observer_rows, "reports": len(observer_rows)},
        "verification": {
            "plan_and_order_checked": True, "receipts_checked": True,
            "report_schema_checked": True, "output_parity_checked": True,
            "resource_limits_checked": True, "final_permits_released_checked": True,
            "production_base_git_blobs_checked": True,
            "architecture_inputs_checked": True, "tool_source_checked": True,
            "quality_commands_checked_exactly": True,
            "observer_timing_separated": True,
            "amdahl_descriptive_only": True,
            "raw_audit_checked": True,
        },
        "reports": all_records,
    }
    return output, scaling, observer_rows


def write_outputs(analysis: dict[str, Any], scaling: list[dict[str, Any]],
                  observer_rows: list[dict[str, Any]]) -> None:
    for name in ("analysis.json", "scaling.csv", "observer.csv", "scaling.md"):
        candidate = PACKET / name
        require(not candidate.exists(), f"refusing to overwrite retained {name}")
    (PACKET / "analysis.json").write_text(json.dumps(analysis, indent=2, sort_keys=True) + "\n")
    (PACKET / "scaling.csv").write_text(_csv_text(scaling, SCALING_FIELDS))
    (PACKET / "observer.csv").write_text(_csv_text(observer_rows, OBSERVER_FIELDS))
    # Re-render from the just-written semantic object so the retained prose is
    # deterministic and replayable.
    (PACKET / "scaling.md").write_text(render_markdown(analysis, scaling))


def check_outputs(analysis: dict[str, Any], scaling: list[dict[str, Any]],
                  observer_rows: list[dict[str, Any]]) -> None:
    path = PACKET / "analysis.json"
    retained = read_json(path)
    require(retained == analysis, "analysis.json does not replay byte-for-byte")
    expected = {
        "scaling.csv": _csv_text(scaling, SCALING_FIELDS),
        "observer.csv": _csv_text(observer_rows, OBSERVER_FIELDS),
        "scaling.md": render_markdown(analysis, scaling),
    }
    for name, text in expected.items():
        candidate = PACKET / name
        require(candidate.is_file() and not candidate.is_symlink(), f"{name} is missing")
        require(candidate.read_text() == text, f"{name} does not replay byte-for-byte")


def analyze(*, write: bool = False, check: bool = False) -> dict[str, Any]:
    result, scaling, observer_rows = build_analysis()
    if write:
        write_outputs(result, scaling, observer_rows)
    if check:
        check_outputs(result, scaling, observer_rows)
    return result


def main(argv: list[str] | None = None) -> int:
    import argparse
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="write replay outputs")
    mode.add_argument("--check", action="store_true", help="replay and check retained outputs")
    args = parser.parse_args(argv)
    try:
        analyze(write=args.write or not args.check, check=args.check)
    except ReplayError as error:
        print(f"0786 replay failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
