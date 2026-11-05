#!/usr/bin/env python3
"""Prepare and compile the minimal DrawingML qualification driver.

This helper deliberately lives outside ``tools/perf-baseline`` at execution
time.  It materializes only the workspace crates needed by
``litchi-drawingml`` and builds a tiny package containing the reviewed probe
and allocator binary.  Each invocation owns its generated ``Cargo.lock``;
the lock is generated offline, hashed, and then consumed with ``--locked``.
No benchmark, source tree, or target directory is removed by this tool.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
from typing import Any, Sequence


SCHEMA = "litchi.drawingml.compile-preparation.v1"
FEATURES = (
    "drawingml-probe",
    "drawingml-raw-current",
    "drawingml-owner-current",
    "drawingml-theme-family-current",
    "drawingml-theme-family-owner-limits-current",
    "drawingml-ink-limits-current",
    "drawingml-transform-limits-current",
)
REQUIRED_SOURCE_PATHS = (
    "Cargo.toml",
    "crates/litchi-core",
    "crates/litchi-drawingml",
    "crates/litchi-ooxml-common",
    "crates/litchi-opc",
    "crates/soapberry-zip",
    "crates/xml-minifier",
    "crates/xml-minifier-macros",
)
PROBE_RELATIVE_PATH = Path("tools/perf-baseline/src/drawingml_raw_probe.rs")
ALLOCATOR_RELATIVE_PATH = Path(
    "tools/perf-baseline/src/bin/support/drawingml_counting_allocator.rs"
)
HEX_SHA = re.compile(r"^[0-9a-f]{40}$")


class DriverError(RuntimeError):
    """A source, preparation, or compile identity failed closed."""


def _run(
    argv: Sequence[str],
    *,
    cwd: Path | None = None,
    env: dict[str, str] | None = None,
    capture: bool = True,
) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        list(argv),
        cwd=cwd,
        env=env,
        check=False,
        text=True,
        stdout=subprocess.PIPE if capture else None,
        stderr=subprocess.PIPE if capture else None,
    )


def _git(root: Path, *args: str) -> str:
    result = _run(["git", "-C", str(root), *args])
    if result.returncode != 0:
        raise DriverError(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout


def _sha256(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise DriverError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def _driver_identity() -> dict[str, Any]:
    path = Path(__file__).resolve()
    digest, size = _sha256(path)
    return {"path": str(path), "sha256": digest, "bytes": size}


def _require_full_sha(value: str, label: str) -> str:
    if not HEX_SHA.fullmatch(value):
        raise DriverError(f"{label} must be a full 40-hex commit SHA")
    return value


def _ensure_absent(path: Path, label: str) -> None:
    if path.exists() or path.is_symlink():
        raise DriverError(f"{label} already exists; refusing to overwrite it: {path}")


def _ensure_empty_directory(path: Path, label: str) -> None:
    if path.is_symlink() or path.exists() and not path.is_dir():
        raise DriverError(f"{label} is not a directory: {path}")
    if path.exists() and any(path.iterdir()):
        raise DriverError(f"{label} must be absent or empty: {path}")
    path.mkdir(parents=True, exist_ok=True)


def _source_identity(root: Path, expected_revision: str | None = None) -> dict[str, Any]:
    if not root.is_dir() or root.is_symlink():
        raise DriverError(f"source root is not a regular directory: {root}")
    root = root.resolve()
    head = _git(root, "rev-parse", "HEAD").strip()
    if expected_revision is not None and head != _require_full_sha(expected_revision, "source revision"):
        raise DriverError(f"source HEAD {head} differs from requested {expected_revision}")
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    if status:
        raise DriverError(f"source worktree is dirty: {root}\n{status}")
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    paths: dict[str, dict[str, Any]] = {}
    for relative in REQUIRED_SOURCE_PATHS:
        path = root / relative
        if relative != "Cargo.toml" and (path.is_symlink() or not path.is_dir()):
            raise DriverError(f"required sparse source path is absent: {path}")
        if relative == "Cargo.toml" and (path.is_symlink() or not path.is_file()):
            raise DriverError(f"required source manifest is absent: {path}")
        if path.is_file():
            digest, size = _sha256(path)
            paths[relative] = {"sha256": digest, "bytes": size}
        else:
            paths[relative] = {"directory": True}
    return {"root": str(root), "revision": head, "tree": tree, "status": status, "paths": paths}


def _prepare_source(args: argparse.Namespace) -> int:
    git_root = Path(args.git_root).resolve()
    source_root = Path(args.source_root).resolve()
    revision = _require_full_sha(args.revision, "revision")
    _ensure_absent(source_root, "sparse source root")
    _git(git_root, "rev-parse", "--show-toplevel")
    check = _run(["git", "-C", str(git_root), "cat-file", "-e", f"{revision}^{{commit}}"])
    if check.returncode != 0:
        raise DriverError(f"revision is not available in {git_root}: {revision}")
    source_root.parent.mkdir(parents=True, exist_ok=True)
    result = _run(
        ["git", "-C", str(git_root), "worktree", "add", "--detach", "--no-checkout", str(source_root), revision]
    )
    if result.returncode != 0:
        raise DriverError(f"cannot create sparse source worktree: {result.stderr.strip()}")
    patterns = ["/Cargo.toml", "/Cargo.lock", "/.cargo/", *[f"/{path}/" for path in REQUIRED_SOURCE_PATHS[1:]]]
    sparse = _run(["git", "-C", str(source_root), "sparse-checkout", "init", "--no-cone"])
    if sparse.returncode == 0:
        sparse = _run(["git", "-C", str(source_root), "sparse-checkout", "set", "--no-cone", *patterns])
    if sparse.returncode != 0:
        raise DriverError(f"cannot materialize sparse source closure: {sparse.stderr.strip()}")
    # `worktree add --no-checkout` starts with an empty index.  Some Git
    # versions update the sparse patterns but defer the initial index
    # population, so make that step explicit before recording the receipt.
    populated = _run(["git", "-C", str(source_root), "read-tree", "-mu", "HEAD"])
    if populated.returncode != 0:
        raise DriverError(f"cannot populate sparse source index: {populated.stderr.strip()}")
    identity = _source_identity(source_root, revision)
    payload = {
        "schema": SCHEMA,
        "operation": "prepare",
        "driver": _driver_identity(),
        "required_paths": list(REQUIRED_SOURCE_PATHS),
        "sparse_patterns": patterns,
        "source": identity,
    }
    output = Path(args.output).resolve()
    _ensure_absent(output, "preparation receipt")
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return 0


def _write_harness(work: Path, source_root: Path, probe_root: Path) -> dict[str, Any]:
    source = source_root / "crates/litchi-drawingml"
    common = source_root / "crates/litchi-ooxml-common"
    probe = probe_root / PROBE_RELATIVE_PATH
    allocator = probe_root / ALLOCATOR_RELATIVE_PATH
    if probe.is_symlink() or not probe.is_file():
        raise DriverError(f"probe source is absent or symlinked: {probe}")
    if allocator.is_symlink() or not allocator.is_file():
        raise DriverError(f"allocator source is absent or symlinked: {allocator}")
    probe_hash, probe_bytes = _sha256(probe)
    allocator_hash, allocator_bytes = _sha256(allocator)
    src = work / "src"
    support = src / "support"
    src.mkdir(parents=True, exist_ok=True)
    support.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(probe, src / "drawingml_raw_probe.rs")
    shutil.copyfile(allocator, support / "drawingml_counting_allocator.rs")
    (src / "lib.rs").write_text(
        "#![forbid(unsafe_code)]\n\npub mod drawingml_raw_probe;\n",
        encoding="utf-8",
    )
    (src / "bin").mkdir()
    (src / "bin/drawingml_raw_native.rs").write_text(
        "fn main() -> std::io::Result<()> {\n"
        "    litchi_perf_baseline::drawingml_raw_probe::run(None)\n"
        "}\n",
        encoding="utf-8",
    )
    (src / "bin/drawingml_raw_allocator.rs").write_text(
        "#[path = \"../support/drawingml_counting_allocator.rs\"]\n"
        "mod allocator;\n\n"
        "fn main() -> std::io::Result<()> {\n"
        "    litchi_perf_baseline::drawingml_raw_probe::run(Some(&allocator::OBSERVER))\n"
        "}\n",
        encoding="utf-8",
    )
    manifest = work / "Cargo.toml"
    source_literal = str(source).replace("\\", "\\\\").replace('"', '\\"')
    common_literal = str(common).replace("\\", "\\\\").replace('"', '\\"')
    manifest.write_text(
        f"""[package]
name = "litchi-perf-baseline"
version = "0.1.0"
edition = "2024"
publish = false

[features]
drawingml-probe = []
drawingml-raw-current = []
drawingml-owner-current = []
drawingml-theme-family-current = []
drawingml-theme-family-owner-limits-current = []
drawingml-ink-limits-current = []
drawingml-transform-limits-current = []

[lib]
path = "src/lib.rs"

[[bin]]
name = "drawingml-raw-native"
path = "src/bin/drawingml_raw_native.rs"

[[bin]]
name = "drawingml-raw-allocator"
path = "src/bin/drawingml_raw_allocator.rs"

[dependencies]
litchi-drawingml = {{ path = "{source_literal}" }}
litchi-ooxml-common = {{ path = "{common_literal}" }}
quick-xml = "0.41"
serde = {{ version = "1", features = ["derive"] }}
serde_json = "1"
sha2 = "=0.11.0"
""",
        encoding="utf-8",
    )
    return {
        "probe": {"path": str(probe), "sha256": probe_hash, "bytes": probe_bytes},
        "allocator": {"path": str(allocator), "sha256": allocator_hash, "bytes": allocator_bytes},
        "manifest": {"path": str(manifest), "sha256": _sha256(manifest)[0]},
    }


def _compile(args: argparse.Namespace) -> int:
    source_root = Path(args.source_root).resolve()
    probe_root = Path(args.probe_root).resolve()
    output = Path(args.output).resolve()
    work = Path(args.work_dir).resolve()
    target = Path(args.target_dir).resolve()
    features = tuple(dict.fromkeys(filter(None, args.features.split(","))))
    unknown = sorted(set(features) - set(FEATURES))
    if unknown:
        raise DriverError(f"unknown DrawingML feature(s): {unknown}")
    source = _source_identity(source_root, args.source_revision)
    if not probe_root.is_dir():
        raise DriverError(f"probe root is not a directory: {probe_root}")
    _ensure_absent(output, "compile receipt")
    _ensure_empty_directory(work, "driver work directory")
    _ensure_empty_directory(target, "Cargo target directory")
    work.mkdir(parents=True, exist_ok=True)
    adapter = _write_harness(work, source_root, probe_root)
    manifest = work / "Cargo.toml"
    env = os.environ.copy()
    env.update({"CARGO_BUILD_JOBS": "1", "CARGO_INCREMENTAL": "0", "CARGO_TARGET_DIR": str(target)})
    cargo = args.cargo
    generate_argv = [cargo, "generate-lockfile", "--manifest-path", str(manifest), "--offline"]
    generate = _run(generate_argv, cwd=work, env=env)
    (work / "generate.stdout.log").write_text(generate.stdout or "", encoding="utf-8")
    (work / "generate.stderr.log").write_text(generate.stderr or "", encoding="utf-8")
    lock = work / "Cargo.lock"
    lock_identity = None
    check = None
    if generate.returncode == 0 and lock.is_file():
        lock_hash, lock_bytes = _sha256(lock)
        lock_identity = {"path": str(lock), "sha256": lock_hash, "bytes": lock_bytes}
        check_argv = [cargo, "check", "--manifest-path", str(manifest), "--locked", "--offline", "--bins"]
        if features:
            check_argv.extend(["--features", ",".join(features)])
        check = _run(check_argv, cwd=work, env=env)
        (work / "check.stdout.log").write_text(check.stdout or "", encoding="utf-8")
        (work / "check.stderr.log").write_text(check.stderr or "", encoding="utf-8")
    payload = {
        "schema": SCHEMA,
        "operation": "compile",
        "driver": _driver_identity(),
        "role": args.role,
        "source": source,
        "features": list(features),
        "adapter": adapter,
        "lock": lock_identity,
        "commands": {
            "generate_lockfile": generate_argv,
            "check": check_argv if check is not None else None,
        },
        "results": {
            "generate_lockfile": generate.returncode,
            "check": check.returncode if check is not None else None,
        },
        "cargo": {"version": _run([cargo, "--version"]).stdout.strip()},
        "work_dir": str(work),
        "target_dir": str(target),
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    if generate.returncode != 0:
        raise DriverError(f"Cargo lock generation failed; see {work / 'generate.stderr.log'}")
    if check is None or check.returncode != 0:
        raise DriverError(f"Cargo check failed; see {work / 'check.stderr.log'}")
    return 0


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    prepare = subparsers.add_parser("prepare", help="create a sparse historical source worktree")
    prepare.add_argument("--git-root", required=True)
    prepare.add_argument("--revision", required=True)
    prepare.add_argument("--source-root", required=True)
    prepare.add_argument("--output", required=True)
    prepare.set_defaults(function=_prepare_source)
    compile_parser = subparsers.add_parser("compile", help="compile the minimal probe package")
    compile_parser.add_argument("--source-root", required=True)
    compile_parser.add_argument("--source-revision", required=True)
    compile_parser.add_argument("--probe-root", required=True)
    compile_parser.add_argument("--features", default="")
    compile_parser.add_argument("--role", required=True)
    compile_parser.add_argument("--work-dir", required=True)
    compile_parser.add_argument("--target-dir", required=True)
    compile_parser.add_argument("--output", required=True)
    compile_parser.add_argument("--cargo", default="cargo")
    compile_parser.set_defaults(function=_compile)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = _parser().parse_args(argv)
    try:
        return int(args.function(args))
    except DriverError as error:
        print(f"drawingml compile driver: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
