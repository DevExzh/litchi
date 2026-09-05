#!/usr/bin/env python3
"""Build the single clean 0423 source revision and copy both binaries.

The 0423 workload compares two public API roles in one revision, so there is
one candidate build identity.  The normal and allocator binaries are copied
to ``/tmp`` before capture and are never rebuilt or overwritten by this
driver.  The source manifest is intentionally retained in the build record;
capture journals refer to its digest rather than repeating it.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import platform
import re
import shutil
import stat
import subprocess
import sys
from typing import Any


CHANGE = 423
TOOLCHAIN = "1.98.1"
BUILD_RUSTFLAGS = "-C force-frame-pointers=yes -C force-unwind-tables=yes"
NORMAL_BINARY = "litchi-perf-baseline"
ALLOCATOR_BINARY = "litchi-perf-baseline-alloc"
SOURCE_SPECS = (
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
SHA256 = re.compile(r"^[0-9a-f]{64}$")


class BuildError(RuntimeError):
    """A fail-closed build input or identity error."""


def fail(message: str) -> None:
    raise BuildError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON value {value!r}")


def reject_duplicates(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def load_json(path: Path, label: str) -> Any:
    try:
        return json.loads(
            path.read_text(encoding="utf-8"),
            object_pairs_hook=reject_duplicates,
            parse_constant=reject_constant,
        )
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        fail(f"cannot load {label} {path}: {error}")


def write_json(path: Path, value: Any) -> None:
    temporary = path.with_name(f".{path.name}.tmp")
    try:
        temporary.write_text(
            json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n",
            encoding="utf-8",
        )
        temporary.replace(path)
    except (OSError, TypeError, ValueError, OverflowError) as error:
        try:
            temporary.unlink()
        except OSError:
            pass
        fail(f"cannot write {path}: {error}")


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def run_git(worktree: Path, *arguments: str) -> str:
    try:
        process = subprocess.run(
            ["git", *arguments], cwd=worktree, capture_output=True,
            text=True, check=False,
        )
    except OSError as error:
        fail(f"cannot run git in {worktree}: {error}")
    if process.returncode != 0:
        detail = (process.stderr or process.stdout).strip()
        fail(f"git {' '.join(arguments)} failed in {worktree}: {detail}")
    return process.stdout


def source_identity(worktree: Path) -> dict[str, Any]:
    worktree = worktree.expanduser().resolve()
    if not worktree.is_dir():
        fail(f"source worktree is missing: {worktree}")
    revision = run_git(worktree, "rev-parse", "HEAD").strip()
    status = run_git(worktree, "status", "--porcelain=v1", "--untracked-files=all")
    tracked = run_git(worktree, "ls-files", "-z", "--", *SOURCE_SPECS)
    files: list[dict[str, Any]] = []
    for relative in sorted(item for item in tracked.split("\0") if item):
        path = worktree / relative
        if not path.is_file():
            fail(f"tracked source file is missing: {path}")
        files.append({
            "path": relative,
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        })
    if not files:
        fail(f"source binding matched no files in {worktree}")
    return {
        "worktree": str(worktree),
        "revision": revision,
        "git_status_porcelain": status,
        "clean": status == "",
        "source_files": files,
    }


def host_identity() -> dict[str, Any]:
    commands = {
        "rustc": ["rustc", f"+{TOOLCHAIN}", "--version"],
        "cargo": ["cargo", f"+{TOOLCHAIN}", "--version"],
        "taskset": ["taskset", "--version"],
        "time": ["/usr/bin/time", "--version"],
    }
    tools: dict[str, Any] = {}
    for name, argv in commands.items():
        try:
            process = subprocess.run(argv, capture_output=True, text=True, check=False)
        except OSError as error:
            tools[name] = {"argv": argv, "error": str(error)}
            continue
        tools[name] = {
            "argv": argv,
            "path": shutil.which(argv[0]) if not Path(argv[0]).is_absolute() else argv[0],
            "exit_code": process.returncode,
            "stdout": process.stdout,
            "stderr": process.stderr,
        }
    affinity = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else None
    cpu_model = None
    memory_total_kib = None
    try:
        for line in Path("/proc/cpuinfo").read_text(encoding="utf-8").splitlines():
            if line.startswith("model name") and ":" in line:
                cpu_model = line.split(":", 1)[1].strip()
                break
    except OSError:
        pass
    try:
        for line in Path("/proc/meminfo").read_text(encoding="utf-8").splitlines():
            if line.startswith("MemTotal:"):
                memory_total_kib = int(line.split()[1])
                break
    except (OSError, ValueError, IndexError):
        pass
    return {
        "utc": utc_now(),
        "platform": platform.platform(),
        "python": platform.python_version(),
        "uname": dict(zip(("system", "node", "release", "version", "machine"), os.uname())),
        "logical_cpu_count": os.cpu_count(),
        "process_affinity": affinity,
        "cpu_model": cpu_model,
        "memory_total_kib": memory_total_kib,
        "tools": tools,
    }


def binary_identity(path: Path, label: str) -> dict[str, Any]:
    path = path.expanduser().resolve()
    if not path.is_file():
        fail(f"{label} binary does not exist: {path}")
    mode_bits = stat.S_IMODE(path.stat().st_mode)
    executable = bool(mode_bits & (stat.S_IXUSR | stat.S_IXGRP | stat.S_IXOTH))
    if not executable:
        fail(f"{label} binary is not executable: {path}")
    digest = sha256_file(path)
    if SHA256.fullmatch(digest) is None:
        fail(f"{label} binary hash is malformed")
    size = path.stat().st_size
    return {
        "label": label,
        "path": str(path),
        "sha256": digest,
        "binary_sha256": digest,
        "bytes": size,
        "binary_bytes": size,
        "mode_bits": mode_bits,
        "executable": executable,
        "profile": "release-debug1-frame-pointer",
    }


def protocol_binding(root: Path) -> tuple[dict[str, Any], str]:
    path = root / "protocol.json"
    if not path.is_file():
        fail(f"frozen protocol is missing: {path}")
    protocol = load_json(path, "protocol")
    if not isinstance(protocol, dict) or protocol.get("change") != CHANGE:
        fail("protocol.json must be the 0423 protocol")
    revision = protocol.get("baseline_revision")
    if not isinstance(revision, str) or not revision or any(
        character not in "0123456789abcdef" for character in revision.lower()
    ) or revision != revision.lower():
        fail("protocol.baseline_revision must be a lowercase git prefix")
    return protocol, sha256_file(path)


def require_baseline_ancestor(worktree: Path, baseline: str, revision: str) -> None:
    try:
        process = subprocess.run(
            ["git", "merge-base", "--is-ancestor", baseline, revision],
            cwd=worktree, capture_output=True, text=True, check=False,
        )
    except OSError as error:
        fail(f"cannot check protocol baseline ancestry: {error}")
    if process.returncode != 0:
        fail(
            f"protocol baseline {baseline} is not an ancestor of source revision "
            f"{revision}: {(process.stderr or process.stdout).strip()}"
        )


def build_environment(target: Path) -> dict[str, str]:
    return {
        "RUSTUP_TOOLCHAIN": TOOLCHAIN,
        "CARGO_TARGET_DIR": str(target.resolve()),
        "CARGO_BUILD_JOBS": "4",
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "RUSTFLAGS": BUILD_RUSTFLAGS,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("role", choices=("candidate",), help="the one 0423 source build")
    parser.add_argument("worktree", type=Path)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--binary-prefix", type=Path, default=Path("/tmp/litchi-goal-0423"))
    parser.add_argument("--target-dir", type=Path)
    parser.add_argument("--protocol", type=Path, help="seed protocol.json into --root")
    args = parser.parse_args()

    root = args.root.expanduser().resolve()
    root.mkdir(parents=True, exist_ok=True)
    protocol_path = root / "protocol.json"
    if args.protocol is not None:
        seed = args.protocol.expanduser().resolve()
        if not seed.is_file():
            fail(f"protocol seed is missing: {seed}")
        if protocol_path.exists():
            fail(f"refusing to overwrite protocol: {protocol_path}")
        shutil.copy2(seed, protocol_path)
    protocol, protocol_sha256 = protocol_binding(root)
    output = root / "build-candidate.json"
    if output.exists():
        fail(f"refusing to overwrite build record: {output}")
    source_before = source_identity(args.worktree)
    if not source_before["clean"]:
        fail(f"source worktree is dirty: {source_before['worktree']}")
    baseline = protocol["baseline_revision"]
    require_baseline_ancestor(args.worktree, baseline, source_before["revision"])
    repo_root = Path(__file__).resolve().parents[4]
    target = (args.target_dir or (repo_root / "tools" / "perf-baseline" / "target")).resolve()
    environment = build_environment(target)
    argv = [
        "cargo", f"+{TOOLCHAIN}", "build", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--release", "--features",
        "allocator-metrics", "--bin", NORMAL_BINARY, "--bin", ALLOCATOR_BINARY,
    ]
    record: dict[str, Any] = {
        "schema_version": 1,
        "change": CHANGE,
        "role": "candidate",
        "status": "running",
        "started_utc": utc_now(),
        "protocol": {"path": "protocol.json", "sha256": protocol_sha256},
        "protocol_sha256": protocol_sha256,
        "baseline_revision": baseline,
        "source_before": source_before,
        "environment": environment,
        "host": host_identity(),
        "argv": argv,
        "target_dir": str(target),
    }
    log = root / "build-candidate.log"
    if log.exists():
        fail(f"refusing to overwrite build log: {log}")
    write_json(output, record)
    try:
        with log.open("wb") as stream:
            result = subprocess.run(
                argv, cwd=args.worktree, env=os.environ | environment,
                stdout=stream, stderr=subprocess.STDOUT, check=False,
            )
    except (OSError, subprocess.SubprocessError) as error:
        record.update({"status": "failed", "error": str(error), "exit_code": None})
        try:
            record["source_after"] = source_identity(args.worktree)
        except BuildError as source_error:
            record["source_after_error"] = str(source_error)
        record["finished_utc"] = utc_now()
        write_json(output, record)
        print(f"0423 build failed: {error}", file=sys.stderr)
        return 1

    record["exit_code"] = result.returncode
    try:
        source_after = source_identity(args.worktree)
        record["source_after"] = source_after
    except BuildError as error:
        source_after = None
        record["source_after_error"] = str(error)
    record["binaries"] = {}
    try:
        if result.returncode == 0 and source_after == source_before:
            for mode, name in (("normal", NORMAL_BINARY), ("allocator", ALLOCATOR_BINARY)):
                destination = Path(f"{args.binary_prefix}-{args.role}-{mode}").expanduser()
                if destination.exists():
                    fail(f"refusing to overwrite copied binary: {destination}")
                built = target / "release" / name
                if not built.is_file():
                    fail(f"cargo did not produce {built}")
                destination.parent.mkdir(parents=True, exist_ok=True)
                shutil.copy2(built, destination)
                record["binaries"][mode] = binary_identity(destination, f"candidate/{mode}")
            record["status"] = "pass"
        else:
            record["status"] = "failed"
            if result.returncode == 0 and source_after != source_before:
                record["error"] = "source changed during build"
    except BuildError as error:
        record["status"] = "failed"
        record["error"] = str(error)
    record["finished_utc"] = utc_now()
    write_json(output, record)
    print(json.dumps({"change": CHANGE, "role": "candidate", "status": record["status"]}))
    return 0 if record["status"] == "pass" else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except BuildError as error:
        print(f"0423 build failed: {error}", file=sys.stderr)
        raise SystemExit(1)
