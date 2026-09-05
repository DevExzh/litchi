#!/usr/bin/env python3
"""Build and bind the two 0418 release binaries serially.

The control and candidate worktrees are supplied by the caller.  Both are
built with identical release flags and one shared target directory; each
binary is copied immediately after its role's build so the second build
cannot replace the first role's executable.  A dirty worktree, revision
mismatch, failed build, or missing binary stops the batch and is retained in
the failure journal.
"""

from __future__ import annotations

import argparse
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

from _common import (
    ALLOCATOR_BINARY_NAME,
    CHANGE_ROOT,
    DEFAULT_BINARY_PREFIX,
    DEFAULT_TARGET_DIR,
    NORMAL_BINARY_NAME,
    BatchError,
    add_role_arguments,
    append_failure,
    binary_identity,
    build_environment,
    host_identity,
    load_protocol,
    load_role_specs,
    relative_path,
    require_new,
    source_identity,
    utc_now,
    write_json,
)


BUILD_ARGV = [
    "cargo",
    "+1.98.1",
    "build",
    "--locked",
    "--manifest-path",
    "tools/perf-baseline/Cargo.toml",
    "--release",
    "--features",
    "allocator-metrics",
    "--bin",
    NORMAL_BINARY_NAME,
    "--bin",
    ALLOCATOR_BINARY_NAME,
]


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=CHANGE_ROOT)
    parser.add_argument("--protocol", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--target-dir", type=Path, default=DEFAULT_TARGET_DIR)
    parser.add_argument(
        "--binary-prefix",
        type=Path,
        default=DEFAULT_BINARY_PREFIX,
        help="prefix for copied binaries (default: /tmp/litchi-goal-0418)",
    )
    add_role_arguments(parser)
    return parser.parse_args()


def _copy_binary(
    source: Path, destination: Path, *, role: str, phase: str
) -> dict[str, Any]:
    require_new([destination])
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(source, destination)
    # copy2 preserves the executable mode; make the contract explicit in case
    # the destination directory applies an unusual umask or ACL.
    destination.chmod(destination.stat().st_mode | 0o111)
    return binary_identity(destination, label=f"{role}/{phase}")


def _write_failure(
    root: Path, identity: dict[str, Any], message: str, *, output_preexisted: bool
) -> None:
    identity["status"] = "failed"
    identity["failure"] = message
    # A failed retry must never replace a prior build identity.  The journal
    # still records the attempted command and reason below.
    if not output_preexisted:
        write_json(Path(identity["output"]), identity)
    append_failure(
        root,
        {
            "stage": "build",
            "utc": utc_now(),
            "message": message,
            "identity": relative_path(Path(identity["output"]), root),
        },
    )


def main() -> int:
    args = parse_args()
    root = args.root.resolve()
    output = (args.output or (root / "build-identity.json")).resolve()
    target_dir = args.target_dir.resolve()
    binary_prefix = args.binary_prefix.resolve()
    protocol_path = (args.protocol or (root / "protocol.json")).resolve()
    output_preexisted = output.exists()

    identity: dict[str, Any] = {
        "schema_version": 1,
        "change": 418,
        "status": "starting",
        "output": str(output),
        "protocol": {"path": str(protocol_path)},
        "target_dir": str(target_dir),
        "build_argv": BUILD_ARGV,
        "environment": build_environment(target_dir),
        "roles": {},
    }
    try:
        if output.exists():
            raise BatchError(f"refusing to overwrite existing build identity: {output}")
        protocol = load_protocol(root, protocol_path)
        role_specs = load_role_specs(args)
        identity["protocol"]["sha256"] = protocol["sha256"]
        identity["common_flags"] = protocol["common_flags"]
        identity["common_flags_source"] = protocol["common_flags_source"]
        identity["host"] = host_identity(include_perf=False)
        for name in ("rustc", "cargo"):
            if identity["host"]["tools"][name]["exit_code"] != 0:
                raise BatchError(f"required build tool {name} is unavailable")
        identity["status"] = "running"
        write_json(output, identity)

        env = os.environ.copy()
        env.update(identity["environment"])
        target_dir.mkdir(parents=True, exist_ok=True)

        for role in ("control", "candidate"):
            spec = role_specs[role]
            worktree = Path(spec["worktree"]).resolve()
            source_before = source_identity(worktree, require_clean=True)
            if source_before["revision"] != spec["revision"]:
                raise BatchError(
                    f"{role} revision mismatch: requested {spec['revision']}, "
                    f"worktree has {source_before['revision']}"
                )
            log = root / "checks" / f"build-{role}.log"
            require_new([log])
            started = utc_now()
            monotonic_started = time.monotonic()
            log.parent.mkdir(parents=True, exist_ok=True)
            with log.open("wb") as stream:
                process = subprocess.run(
                    BUILD_ARGV,
                    cwd=worktree,
                    env=env,
                    stdout=stream,
                    stderr=subprocess.STDOUT,
                    check=False,
                )
            finished = utc_now()
            source_after = source_identity(worktree, require_clean=True)
            if source_after["revision"] != source_before["revision"]:
                raise BatchError(f"{role} revision changed during build")
            build_record: dict[str, Any] = {
                "role": role,
                "source": source_before,
                "source_after": source_after,
                "argv": BUILD_ARGV,
                "cwd": str(worktree),
                "environment": identity["environment"],
                "log": relative_path(log, root),
                "started_utc": started,
                "finished_utc": finished,
                "elapsed_seconds": time.monotonic() - monotonic_started,
                "exit_code": process.returncode,
            }
            identity["roles"][role] = build_record
            write_json(output, identity)
            if process.returncode != 0:
                raise BatchError(f"{role} build failed with exit {process.returncode}")

            release_dir = target_dir / "release"
            source_normal = release_dir / NORMAL_BINARY_NAME
            source_allocator = release_dir / ALLOCATOR_BINARY_NAME
            if not source_normal.is_file() or not source_allocator.is_file():
                raise BatchError(
                    f"{role} build did not produce both release binaries in {release_dir}"
                )
            binaries: dict[str, Any] = {}
            for phase, source, suffix in (
                ("normal", source_normal, "normal"),
                ("allocator", source_allocator, "allocator"),
            ):
                destination = binary_prefix.parent / f"{binary_prefix.name}-{role}-{suffix}"
                binaries[phase] = _copy_binary(
                    source, destination, role=role, phase=phase
                )
            identity["roles"][role]["binaries"] = binaries
            write_json(output, identity)

        identity["status"] = "complete"
        identity["completed_utc"] = utc_now()
        write_json(output, identity)
        print(f"built 0418 control and candidate binaries: {output}")
        return 0
    except (BatchError, OSError, subprocess.SubprocessError) as exc:
        message = str(exc)
        _write_failure(root, identity, message, output_preexisted=output_preexisted)
        print(f"0418 build failed: {message}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
