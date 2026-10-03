#!/usr/bin/env python3
"""Build one immutable 0894 observable probe role.

The driver materializes the standalone probe manifest against an explicit
source tree.  A baseline invocation may generate the offline Cargo.lock; the
candidate invocation must receive that exact lock through ``--lock-source``.
Each role owns ``target/litchi-0894-probe-{role}`` and refuses to reuse it.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
import subprocess
from typing import Any


PROBE = Path(__file__).resolve().parent
PACKET = PROBE.parent
ROOT = PROBE.parents[4]
TEMPLATE = PROBE / "Cargo.toml.template"
SOURCE = PROBE / "probe.rs"
HELPERS = (TEMPLATE, SOURCE, PROBE / "input_census.py", PROBE / "compare.py", Path(__file__).resolve())
TOOLCHAIN = "1.95.0"
JOBS = "2"
PROFILE = "dev"


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ValueError(message)


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"expected regular non-symlink file: {path}")
    path = path.resolve()
    mode = stat.S_IMODE(path.stat().st_mode)
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
        "mode": mode,
        "executable": bool(mode & stat.S_IXUSR),
    }


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as stream:
        json.dump(value, stream, indent=2, sort_keys=True)
        stream.write("\n")


def run_logged(argv: list[str], cwd: Path, env: dict[str, str], stdout: Path, stderr: Path) -> subprocess.CompletedProcess[str]:
    with stdout.open("x", encoding="utf-8") as out, stderr.open("x", encoding="utf-8") as err:
        return subprocess.run(
            argv,
            cwd=cwd,
            env=env,
            stdout=out,
            stderr=err,
            text=True,
            check=False,
        )


def git_revision(tree: Path) -> str | None:
    top = subprocess.run(
        ["git", "-C", str(tree), "rev-parse", "--show-toplevel"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if top.returncode != 0 or Path(top.stdout.strip()).resolve() != tree.resolve():
        return None
    result = subprocess.run(
        ["git", "-C", str(tree), "rev-parse", "HEAD"],
        cwd=ROOT,
        capture_output=True,
        text=True,
        check=False,
    )
    if result.returncode != 0:
        return None
    return result.stdout.strip() or None


def environment(target: Path) -> dict[str, str]:
    tmp = target / "tmp"
    tmp.mkdir(parents=True, exist_ok=True)
    inherited = {
        key: os.environ[key]
        for key in ("PATH", "HOME", "RUSTUP_HOME", "CARGO_HOME", "USER", "LOGNAME", "TERM")
        if key in os.environ
    }
    values = inherited
    values.update(
        {
            "RUSTUP_TOOLCHAIN": TOOLCHAIN,
            "CARGO_NET_OFFLINE": "true",
            "CARGO_TARGET_DIR": str((target / "cargo-target").resolve()),
            "TMPDIR": str(tmp.resolve()),
            "CARGO_TERM_COLOR": "never",
        }
    )
    for key in (
        "RUSTFLAGS",
        "RUSTC_WRAPPER",
        "CARGO_BUILD_RUSTC_WRAPPER",
        "CARGO_ENCODED_RUSTFLAGS",
        "CARGO_PROFILE_RELEASE_LTO",
        "CARGO_PROFILE_DEV_LTO",
    ):
        values.pop(key, None)
    return values


def materialize_manifest(target: Path, tree: Path) -> Path:
    text = TEMPLATE.read_text(encoding="utf-8")
    manifest = text.replace("@TREE@", str(tree))
    path = target / "Cargo.toml"
    path.write_text(manifest, encoding="utf-8", newline="\n")
    shutil.copyfile(SOURCE, target / "probe.rs")
    return path


def args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--tree", type=Path, required=True, help="immutable baseline/candidate source tree")
    parser.add_argument("--role", choices=("baseline", "candidate"), required=True)
    parser.add_argument("--lock-source", type=Path, help="exact Cargo.lock from the baseline probe role")
    parser.add_argument(
        "--source-binding",
        type=Path,
        action="append",
        default=[],
        help="source archive/manifest or before/after receipt to bind by descriptor",
    )
    parser.add_argument(
        "--target-root",
        type=Path,
        help="owned build root; defaults to repository target/litchi-0894-probe-{role}",
    )
    return parser.parse_args()


def main() -> int:
    options = args()
    require(options.tree.is_dir() and not options.tree.is_symlink(), f"source tree is not a regular directory: {options.tree}")
    tree = options.tree.resolve()
    target = (options.target_root or ROOT / "target" / f"litchi-0894-probe-{options.role}").resolve()
    require(not target.exists(), f"owned probe target already exists: {target}")
    require(TEMPLATE.is_file() and SOURCE.is_file(), "probe template/source is missing")
    if options.role == "candidate":
        require(options.lock_source is not None, "candidate requires --lock-source from baseline")
    if options.lock_source is not None:
        require(options.lock_source.is_file() and not options.lock_source.is_symlink(), f"lock source is missing: {options.lock_source}")
        lock_source = options.lock_source.resolve()
    else:
        lock_source = None
    source_bindings = []
    for binding in options.source_binding:
        require(binding.is_file() and not binding.is_symlink(), f"source binding is missing: {binding}")
        source_bindings.append(descriptor(binding))

    target.mkdir(parents=True)
    marker = target / f".owner-litchi-0894-probe-{options.role}"
    marker.write_text(f"litchi 0894 observable probe {options.role}\n", encoding="utf-8")
    manifest = materialize_manifest(target, tree)
    env = environment(target)
    lock = target / "Cargo.lock"
    if lock_source is not None:
        shutil.copyfile(lock_source, lock)
    start = {
        "schema": "litchi.xlsx-observable-probe-build-start-0894.v1",
        "status": "building",
        "role": options.role,
        "started_utc": utc_now(),
        "tree": str(tree),
        "tree_revision": git_revision(tree),
        "target_root": str(target),
        "marker": descriptor(marker),
        "template": descriptor(TEMPLATE),
        "probe_source": descriptor(SOURCE),
        "helpers": {str(path.name): descriptor(path) for path in HELPERS},
        "manifest": descriptor(manifest),
        "lock_source": descriptor(lock_source) if lock_source is not None else None,
        "environment": {
            "RUSTUP_TOOLCHAIN": TOOLCHAIN,
            "CARGO_NET_OFFLINE": "true",
            "CARGO_TARGET_DIR": env["CARGO_TARGET_DIR"],
            "TMPDIR": env["TMPDIR"],
            "CARGO_TERM_COLOR": "never",
        },
        "source_bindings": source_bindings,
        "profile": PROFILE,
        "debug": 0,
        "incremental": False,
        "jobs": int(JOBS),
    }
    write_once(target / "build-start.json", start)

    rustc_result = run_logged(
        ["rustc", "--version"],
        target,
        env,
        target / "rustc.stdout.txt",
        target / "rustc.stderr.txt",
    )
    cargo_result = run_logged(
        ["cargo", "--version"],
        target,
        env,
        target / "cargo.stdout.txt",
        target / "cargo.stderr.txt",
    )
    require(rustc_result.returncode == 0 and cargo_result.returncode == 0, "toolchain version probe failed")
    rustc_version = (target / "rustc.stdout.txt").read_text(encoding="utf-8").strip()
    cargo_version = (target / "cargo.stdout.txt").read_text(encoding="utf-8").strip()
    require(rustc_version.startswith(f"rustc {TOOLCHAIN}"), f"unexpected rustc version: {rustc_version}")
    require(cargo_version.startswith("cargo "), f"unexpected cargo version: {cargo_version}")

    generate = ["cargo", "generate-lockfile", "--offline", "--manifest-path", str(manifest)]
    if lock_source is None:
        generate_result = run_logged(generate, target, env, target / "lock.stdout.txt", target / "lock.stderr.txt")
        require(generate_result.returncode == 0, f"offline Cargo.lock generation failed; see {target / 'lock.stderr.txt'}")
    require(lock.is_file(), f"Cargo.lock was not materialized: {lock}")
    lock_descriptor = descriptor(lock)
    command = [
        "cargo",
        "build",
        "--locked",
        "--offline",
        "--manifest-path",
        str(manifest),
        "--bin",
        "xlsx0894-probe",
        "-j",
        JOBS,
    ]
    result = run_logged(command, target, env, target / "build.stdout.txt", target / "build.stderr.txt")
    if result.returncode != 0:
        write_once(
            target / "build-failure.json",
            {
                "schema": "litchi.xlsx-observable-probe-build-failure-0894.v1",
                "status": "fail",
                "role": options.role,
                "returncode": result.returncode,
                "command": command,
                "lock": lock_descriptor,
            },
        )
        raise RuntimeError(f"probe build failed ({result.returncode}); see {target / 'build.stderr.txt'}")

    built = target / "cargo-target" / "debug" / "xlsx0894-probe"
    require(built.is_file(), f"Cargo did not produce the probe binary: {built}")
    staged = target / "xlsx0894-probe"
    shutil.copy2(built, staged)
    staged.chmod(staged.stat().st_mode | stat.S_IXUSR)
    write_once(
        target / "build.json",
        {
            "schema": "litchi.xlsx-observable-probe-build-0894.v1",
            "status": "pass",
            "role": options.role,
            "started_utc": start["started_utc"],
            "finished_utc": utc_now(),
            "tree": str(tree),
            "tree_revision": git_revision(tree),
            "source_bindings": source_bindings,
            "target_root": str(target),
            "marker": descriptor(marker),
            "manifest": descriptor(manifest),
            "lock": lock_descriptor,
            "lock_source": descriptor(lock_source) if lock_source is not None else None,
            "binary": descriptor(staged),
            "cargo_binary": descriptor(built),
            "command": command,
            "lock_command": generate if lock_source is None else None,
            "environment": start["environment"],
            "profile": PROFILE,
            "debug": 0,
            "incremental": False,
            "jobs": int(JOBS),
            "toolchain": TOOLCHAIN,
            "rustc_version": rustc_version,
            "cargo_version": cargo_version,
            "helpers": start["helpers"],
            "logs": {
                "rustc_stdout": descriptor(target / "rustc.stdout.txt"),
                "rustc_stderr": descriptor(target / "rustc.stderr.txt"),
                "cargo_stdout": descriptor(target / "cargo.stdout.txt"),
                "cargo_stderr": descriptor(target / "cargo.stderr.txt"),
                "lock_stdout": descriptor(target / "lock.stdout.txt") if lock_source is None else None,
                "lock_stderr": descriptor(target / "lock.stderr.txt") if lock_source is None else None,
                "build_stdout": descriptor(target / "build.stdout.txt"),
                "build_stderr": descriptor(target / "build.stderr.txt"),
            },
        },
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
