#!/usr/bin/env python3
"""Shared, fail-closed helpers for the 0418 measurement scripts.

The three public scripts in this directory are deliberately thin wrappers
around this module.  This file contains no workload invocation at import
time.  It only describes how the build, capture, and profile records bind to
the protocol, source worktrees, and copied binaries.
"""

from __future__ import annotations

import datetime as _datetime
import hashlib
import json
import math
import os
import platform
import re
import shutil
import stat
import subprocess
from pathlib import Path
from typing import Any, Mapping, Sequence


REPO_ROOT = Path(__file__).resolve().parents[5]
CHANGE_ROOT = REPO_ROOT / "docs" / "performance" / "results" / "change-0418"
ABBA_LEGS = ("A1", "B1", "B2", "A2")
ABBA_ROLES = ("control", "candidate", "candidate", "control")
ROLE_BY_LEG = dict(zip(ABBA_LEGS, ABBA_ROLES))
RUNTIME_TOOLCHAIN = "1.98.1"
BUILD_RUSTFLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
NORMAL_BINARY_NAME = "litchi-perf-baseline"
ALLOCATOR_BINARY_NAME = "litchi-perf-baseline-alloc"
DEFAULT_TARGET_DIR = REPO_ROOT / "tools" / "perf-baseline" / "target"
DEFAULT_BINARY_PREFIX = Path("/tmp/litchi-goal-0418")
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")
SAFE_SELECTOR_RE = re.compile(r"^[A-Za-z0-9_.-]+$")
SOURCE_FILE_SPECS = (
    "Cargo.toml",
    "Cargo.lock",
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/Cargo.lock",
    "tools/perf-baseline/src",
    "crates/litchi-opc/Cargo.toml",
    "crates/litchi-opc/src",
    "crates/litchi-pptx/Cargo.toml",
    "crates/litchi-pptx/src",
)


class BatchError(RuntimeError):
    """An input, identity, or subprocess failure that must stop the batch."""


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON number is not permitted: {value}")


def _reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON object key: {key}")
        result[key] = value
    return result


def load_json(path: Path) -> Any:
    try:
        return json.loads(
            path.read_text(),
            object_pairs_hook=_reject_duplicate_keys,
            parse_constant=_reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as exc:
        raise BatchError(f"cannot load strict JSON {path}: {exc}") from exc


def write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = path.with_name(f".{path.name}.tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=False) + "\n")
    os.replace(temporary, path)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def sha256_json(path: Path) -> str:
    return sha256_file(path)


def relative_path(path: Path, root: Path) -> str:
    """Return a portable artifact path while retaining absolute argv fields."""

    try:
        return path.resolve().relative_to(root.resolve()).as_posix()
    except ValueError:
        return str(path)


def require_string(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise BatchError(f"{label} must be a non-empty string")
    return value


def require_int(value: Any, label: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        raise BatchError(f"{label} must be an integer >= {minimum}")
    return value


def run_text(argv: Sequence[str], *, cwd: Path | None = None,
             env: Mapping[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        list(argv), cwd=cwd, env=None if env is None else dict(env),
        capture_output=True, text=True, check=False,
    )


def git_value(worktree: Path, *arguments: str) -> str:
    process = run_text(["git", *arguments], cwd=worktree)
    if process.returncode:
        detail = (process.stderr or process.stdout).strip()
        raise BatchError(
            f"git {' '.join(arguments)} failed in {worktree} "
            f"(exit {process.returncode}): {detail}"
        )
    return process.stdout


def source_identity(worktree: Path, *, require_clean: bool = True) -> dict[str, Any]:
    worktree = worktree.resolve()
    if not worktree.is_dir():
        raise BatchError(f"worktree does not exist: {worktree}")
    revision = git_value(worktree, "rev-parse", "HEAD").strip()
    status = git_value(
        worktree, "status", "--porcelain=v1", "--untracked-files=all"
    )
    if require_clean and status:
        raise BatchError(f"worktree is dirty: {worktree}: {status!r}")
    tracked = run_text(["git", "ls-files", "-z", "--", *SOURCE_FILE_SPECS], cwd=worktree)
    if tracked.returncode:
        raise BatchError(f"git ls-files failed in {worktree}: {tracked.stderr.strip()}")
    source_files: list[dict[str, Any]] = []
    for relative in sorted(path for path in tracked.stdout.split("\0") if path):
        source_path = worktree / relative
        if not source_path.is_file():
            raise BatchError(f"tracked source binding is missing: {source_path}")
        source_files.append({
            "path": relative,
            "bytes": source_path.stat().st_size,
            "sha256": sha256_file(source_path),
        })
    if not source_files:
        raise BatchError(f"no source files matched binding in {worktree}")
    return {
        "worktree": str(worktree),
        "revision": revision,
        "git_status_porcelain": status,
        "clean": not bool(status),
        "source_files": source_files,
    }


def add_role_arguments(parser: Any) -> None:
    parser.add_argument(
        "--roles-json",
        type=Path,
        help="JSON containing control/candidate revision and worktree fields",
    )
    for role in ("control", "candidate"):
        parser.add_argument(f"--{role}-revision")
        parser.add_argument(f"--{role}-worktree", type=Path)


def load_role_specs(args: Any) -> dict[str, dict[str, Any]]:
    values: dict[str, dict[str, Any]] = {}
    if args.roles_json is not None:
        raw = load_json(args.roles_json)
        if isinstance(raw, dict) and isinstance(raw.get("roles"), dict):
            raw = raw["roles"]
        if not isinstance(raw, dict):
            raise BatchError("roles JSON must be an object or an object with roles")
        for role in ("control", "candidate"):
            item = raw.get(role)
            if not isinstance(item, dict):
                raise BatchError(f"roles JSON is missing object {role!r}")
            revision = item.get("revision", item.get("git_revision", item.get("source_revision")))
            worktree = item.get("worktree")
            values[role] = {
                "revision": require_string(revision, f"{role}.revision"),
                "worktree": Path(require_string(worktree, f"{role}.worktree")),
            }
        return values

    for role in ("control", "candidate"):
        revision = getattr(args, f"{role}_revision")
        worktree = getattr(args, f"{role}_worktree")
        if revision is None or worktree is None:
            raise BatchError(
                "provide --roles-json or all four --control/--candidate "
                "revision and worktree arguments"
            )
        values[role] = {"revision": revision, "worktree": worktree}
    return values


def load_protocol(root: Path, protocol_path: Path | None = None) -> dict[str, Any]:
    path = (protocol_path or (root / "protocol.json")).resolve()
    protocol = load_json(path)
    if not isinstance(protocol, dict):
        raise BatchError("0418 protocol must be a JSON object")
    if protocol.get("change") != 418:
        raise BatchError(f"expected protocol change 418, got {protocol.get('change')!r}")

    order = protocol.get("order")
    if order != list(ABBA_ROLES):
        raise BatchError(
            f"protocol order must be {list(ABBA_ROLES)!r}, got {order!r}"
        )

    jobs = protocol.get("jobs")
    if not isinstance(jobs, list) or not jobs:
        raise BatchError("protocol jobs must be a non-empty list")
    normalized_jobs: list[dict[str, Any]] = []
    seen: set[str] = set()
    for index, job in enumerate(jobs):
        if not isinstance(job, dict):
            raise BatchError(f"protocol jobs[{index}] must be an object")
        selector = require_string(job.get("selector"), f"jobs[{index}].selector")
        if not SAFE_SELECTOR_RE.fullmatch(selector):
            raise BatchError(f"unsafe selector name {selector!r}")
        if selector in seen:
            raise BatchError(f"duplicate protocol selector {selector!r}")
        seen.add(selector)
        role = require_string(job.get("role"), f"jobs[{index}].role")
        samples = require_int(job.get("samples"), f"jobs[{index}].samples", minimum=1)
        warmups = require_int(job.get("warmups"), f"jobs[{index}].warmups", minimum=0)
        normalized_jobs.append({
            "selector": selector,
            "role": role,
            "samples": samples,
            "warmups": warmups,
        })

    allocator = protocol.get("allocator")
    if not isinstance(allocator, dict):
        raise BatchError("protocol allocator must be an object")
    allocator_samples = require_int(
        allocator.get("samples"), "allocator.samples", minimum=1
    )
    allocator_warmups = require_int(
        allocator.get("warmups"), "allocator.warmups", minimum=0
    )
    if allocator.get("order") not in ("same ABBA", "same_abba", list(ABBA_ROLES)):
        raise BatchError(f"allocator order is not same ABBA: {allocator.get('order')!r}")
    if allocator.get("latency_comparison") != "excluded":
        raise BatchError("allocator latency comparison must be explicitly excluded")

    explicit_allocator = allocator.get("selectors")
    if explicit_allocator is None:
        # Keep the allocator batch aligned with every declared selector.  The
        # protocol's scope field determines which rows have operation-local
        # allocation evidence; it does not silently remove the phase guards
        # from the reproducible ABBA capture.
        allocator_selectors = [job["selector"] for job in normalized_jobs]
    else:
        if not isinstance(explicit_allocator, list):
            raise BatchError("allocator.selectors must be a list when present")
        allocator_selectors = [
            require_string(value, "allocator.selectors[]") for value in explicit_allocator
        ]
        if len(set(allocator_selectors)) != len(allocator_selectors):
            raise BatchError("allocator selectors contain duplicates")
        if any(value not in seen for value in allocator_selectors):
            raise BatchError("allocator selector is absent from protocol jobs")
        # The protocol's job list is authoritative for selector order.
        allocator_selectors = [
            job["selector"] for job in normalized_jobs if job["selector"] in allocator_selectors
        ]
    if not allocator_selectors:
        raise BatchError("allocator lane has no lifecycle selectors")
    allocator_jobs = [
        job for job in normalized_jobs if job["selector"] in allocator_selectors
    ]

    flags = protocol.get("common_flags")
    flags_source = "protocol.common_flags"
    if flags is None:
        # 0418 intentionally keeps the matched PPTX command flags from 0417;
        # bind the fallback explicitly so a later matrix change is visible.
        fallback_path = root.parent / "change-0417" / "matrix.json"
        fallback = load_json(fallback_path)
        flags = fallback.get("common_flags") if isinstance(fallback, dict) else None
        flags_source = relative_path(fallback_path, root)
    if not isinstance(flags, list) or not all(isinstance(value, str) for value in flags):
        raise BatchError("common_flags must be a list of strings")
    forbidden = {"--case", "--samples", "--warmup", "--json", "--corpus-manifest"}
    if forbidden.intersection(flags):
        raise BatchError("common_flags must not contain harness output or case flags")
    try:
        worker_index = flags.index("--workers")
    except ValueError as exc:
        raise BatchError("common_flags must pin --workers 1") from exc
    if worker_index + 1 >= len(flags) or flags[worker_index + 1] != "1":
        raise BatchError("common_flags must pin --workers 1")

    cpu = require_int(protocol.get("cpu"), "protocol.cpu", minimum=0)
    workers = require_int(protocol.get("workers"), "protocol.workers", minimum=1)
    if workers != 1:
        raise BatchError("0418 capture requires protocol workers=1")
    profile = protocol.get("profile")
    if not isinstance(profile, dict):
        raise BatchError("protocol profile must be an object")
    profile_config = {
        "samples": require_int(profile.get("samples"), "profile.samples", minimum=1),
        "warmups": require_int(profile.get("warmups"), "profile.warmups", minimum=0),
        "event": require_string(profile.get("event"), "profile.event"),
        "frequency": require_int(profile.get("frequency"), "profile.frequency", minimum=1),
        "call_graph": require_string(profile.get("call_graph"), "profile.call_graph"),
        "scope": require_string(profile.get("scope"), "profile.scope"),
    }
    profile_names = profile.get("selectors")
    if profile_names is None:
        profile_names = [
            job["selector"] for job in normalized_jobs
            if job["role"].lower().startswith("primary")
        ] or [normalized_jobs[0]["selector"]]
    if isinstance(profile_names, str):
        profile_names = [profile_names]
    if not isinstance(profile_names, list) or not profile_names:
        raise BatchError("profile.selectors must be a non-empty list when present")
    profile_names = [require_string(value, "profile.selectors[]") for value in profile_names]
    if any(value not in seen for value in profile_names):
        raise BatchError("profile selector is absent from protocol jobs")

    return {
        "path": path,
        "sha256": sha256_json(path),
        "raw": protocol,
        "jobs": normalized_jobs,
        "job_by_selector": {job["selector"]: job for job in normalized_jobs},
        "allocator_jobs": allocator_jobs,
        "allocator": {
            "samples": allocator_samples,
            "warmups": allocator_warmups,
            "selectors": allocator_selectors,
        },
        "common_flags": list(flags),
        "common_flags_source": flags_source,
        "cpu": cpu,
        "workers": workers,
        "profile": profile_config,
        "profile_selectors": profile_names,
    }


def host_identity(*, include_perf: bool = False) -> dict[str, Any]:
    affinity = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else None
    commands: dict[str, list[str]] = {
        "git": ["git", "--version"],
        "rustc": ["rustc", f"+{RUNTIME_TOOLCHAIN}", "--version"],
        "cargo": ["cargo", f"+{RUNTIME_TOOLCHAIN}", "--version"],
        "taskset": ["taskset", "--version"],
        "time": ["/usr/bin/time", "--version"],
    }
    if include_perf:
        commands["perf"] = ["perf", "--version"]
    versions: dict[str, Any] = {}
    for name, argv in commands.items():
        process = run_text(argv, cwd=REPO_ROOT)
        versions[name] = {
            "argv": argv,
            "path": shutil.which(argv[0]) if not Path(argv[0]).is_absolute() else argv[0],
            "exit_code": process.returncode,
            "stdout": process.stdout,
            "stderr": process.stderr,
        }
    return {
        "utc": utc_now(),
        "platform": platform.platform(),
        "python": platform.python_version(),
        "uname": dict(zip(("system", "node", "release", "version", "machine"), os.uname())),
        "logical_cpu_count": os.cpu_count(),
        "process_affinity": affinity,
        "tools": versions,
    }


def build_environment(target_dir: Path) -> dict[str, str]:
    return {
        "RUSTUP_TOOLCHAIN": RUNTIME_TOOLCHAIN,
        "CARGO_TARGET_DIR": str(target_dir.resolve()),
        "CARGO_BUILD_JOBS": "4",
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "RUSTFLAGS": BUILD_RUSTFLAGS,
    }


def binary_identity(path: Path, *, label: str) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file():
        raise BatchError(f"binary does not exist: {path}")
    mode_bits = stat.S_IMODE(path.stat().st_mode)
    executable = bool(mode_bits & (stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH))
    if not executable:
        raise BatchError(f"binary is not executable: {path}")
    digest = sha256_file(path)
    return {
        "label": label,
        "path": str(path),
        "sha256": digest,
        "binary_sha256": digest,
        "bytes": path.stat().st_size,
        "binary_bytes": path.stat().st_size,
        "mode_bits": mode_bits,
        "executable": executable,
        "profile": "release-debug1-frame-pointer",
    }


def descriptor_path(descriptor: Mapping[str, Any]) -> Path:
    return Path(require_string(descriptor.get("path"), "binary.path"))


def validate_binary_descriptor(descriptor: Mapping[str, Any], *, label: str) -> dict[str, Any]:
    path = descriptor_path(descriptor)
    actual = binary_identity(path, label=label)
    expected_hash = descriptor.get("sha256", descriptor.get("binary_sha256"))
    if expected_hash != actual["sha256"]:
        raise BatchError(f"{label} binary hash changed: expected {expected_hash}, got {actual['sha256']}")
    expected_bytes = descriptor.get("bytes", descriptor.get("binary_bytes"))
    if expected_bytes != actual["bytes"]:
        raise BatchError(f"{label} binary size changed: expected {expected_bytes}, got {actual['bytes']}")
    return actual


def append_failure(root: Path, record: Mapping[str, Any]) -> None:
    journal = root / "checks" / "failure-journal.jsonl"
    journal.parent.mkdir(parents=True, exist_ok=True)
    with journal.open("a", encoding="utf-8") as stream:
        stream.write(json.dumps(dict(record), sort_keys=True) + "\n")
        stream.flush()


def load_build_identity(path: Path) -> dict[str, Any]:
    identity = load_json(path)
    if not isinstance(identity, dict) or identity.get("status") != "complete":
        raise BatchError(f"build identity is absent or incomplete: {path}")
    roles = identity.get("roles")
    if not isinstance(roles, dict):
        raise BatchError("build identity must contain roles")
    for role in ("control", "candidate"):
        item = roles.get(role)
        if not isinstance(item, dict):
            raise BatchError(f"build identity is missing role {role!r}")
        if not item.get("source", {}).get("clean", False):
            raise BatchError(f"build identity marks {role} worktree dirty")
        binaries = item.get("binaries")
        if not isinstance(binaries, dict):
            raise BatchError(f"build identity is missing {role} binaries")
        for phase in ("normal", "allocator"):
            descriptor = binaries.get(phase)
            if not isinstance(descriptor, dict):
                raise BatchError(f"build identity is missing {role}/{phase} binary")
            validate_binary_descriptor(descriptor, label=f"{role}/{phase}")
    return identity


def role_binary(identity: Mapping[str, Any], role: str, phase: str) -> dict[str, Any]:
    try:
        descriptor = identity["roles"][role]["binaries"][phase]
    except (KeyError, TypeError) as exc:
        raise BatchError(f"missing build identity binary {role}/{phase}") from exc
    if not isinstance(descriptor, dict):
        raise BatchError(f"invalid build identity binary {role}/{phase}")
    return descriptor


def assert_source_matches(identity: Mapping[str, Any], role: str) -> dict[str, Any]:
    item = identity["roles"][role]
    source = item["source"]
    current = source_identity(Path(source["worktree"]), require_clean=True)
    if current["revision"] != source["revision"]:
        raise BatchError(
            f"{role} worktree revision changed: expected {source['revision']}, "
            f"got {current['revision']}"
        )
    return current


def run_with_files(
    argv: Sequence[str], *, cwd: Path, env: Mapping[str, str],
    stdout_path: Path, stderr_path: Path,
) -> subprocess.CompletedProcess[bytes]:
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    stderr_path.parent.mkdir(parents=True, exist_ok=True)
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        return subprocess.run(
            list(argv), cwd=cwd, env=dict(env), stdout=stdout, stderr=stderr, check=False
        )


def require_new(paths: Sequence[Path]) -> None:
    existing = [str(path) for path in paths if path.exists()]
    if existing:
        raise BatchError("refusing to overwrite existing artifacts: " + ", ".join(existing))


def require_output_files(paths: Sequence[Path]) -> None:
    missing = [str(path) for path in paths if not path.is_file() or path.stat().st_size == 0]
    if missing:
        raise BatchError("successful command did not produce non-empty artifacts: " + ", ".join(missing))


def selector_stem(selector: str) -> str:
    if not SAFE_SELECTOR_RE.fullmatch(selector):
        raise BatchError(f"unsafe selector name {selector!r}")
    return selector
