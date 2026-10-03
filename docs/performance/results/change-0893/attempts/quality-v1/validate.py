#!/usr/bin/env python3
"""Run the non-performance 0893 producer-shape quality gates.

The runner is deliberately create-only per attempt.  It records the exact
source, command, environment, tool versions, output logs, and test counts;
the first failed gate stops the attempt while retaining its evidence.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
from pathlib import Path
import re
import stat
import subprocess
import time
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
MANIFEST = Path("tools/perf-baseline/Cargo.toml")
TARGET = Path("/tmp/litchi-0893-quality")
MARKER = TARGET / ".owner-0893-quality"
MARKER_TEXT = "litchi 0893 producer-shape quality target\n"
TOOLCHAIN = "1.95.0-x86_64-unknown-linux-gnu"
TOOLCHAIN_BIN = Path("/home/zhuhe/.rustup/toolchains") / TOOLCHAIN / "bin"
CARGO = TOOLCHAIN_BIN / "cargo"
RUSTFMT = TOOLCHAIN_BIN / "rustfmt"
RUSTC = TOOLCHAIN_BIN / "rustc"
PYTHON = "python3"

PRIMARY_SOURCE = Path("tools/perf-baseline/src/producer_shape.rs")
NESTED_SOURCE_DIR = Path("tools/perf-baseline/src/producer_shape")
EXPECTED_NESTED_SOURCES = {"full_cell_guard_tests.rs", "selected_route_tests.rs"}
TEST_GATES = {
    "producer_full_cell",
    "producer_shape_selected_route_tests",
    "producer_shape_full_cell_guard_tests",
    "producer_shape_tests",
}
TEST_RESULT = re.compile(r"test result: (?:ok|FAILED)\.\s+(\d+) passed;")
ENV_KEYS = (
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "CARGO_NET_OFFLINE",
    "CARGO_PROFILE_DEV_DEBUG",
    "CARGO_PROFILE_RELEASE_DEBUG",
    "CARGO_PROFILE_TEST_DEBUG",
    "CARGO_TARGET_DIR",
    "CARGO_HOME",
    "HOME",
    "LANG",
    "LC_ALL",
    "PATH",
    "RUSTDOCFLAGS",
    "RUSTFLAGS",
    "RUSTUP_HOME",
    "RUSTUP_TOOLCHAIN",
    "TMPDIR",
    "TZ",
)


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat().replace("+00:00", "Z")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def identity(path: Path) -> dict[str, Any]:
    path = path.resolve()
    if path.is_symlink() or not path.is_file():
        raise RuntimeError(f"expected a regular non-symlink file: {path}")
    metadata = path.stat()
    return {
        "path": str(path),
        "bytes": metadata.st_size,
        "mode": stat.S_IMODE(metadata.st_mode),
        "executable": bool(metadata.st_mode & stat.S_IXUSR),
        "sha256": sha256(path),
    }


def write_exclusive(path: Path, value: Any) -> None:
    if path.exists() or path.is_symlink():
        raise RuntimeError(f"create-only output already exists: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as destination:
        payload = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
        destination.write(payload)
        destination.flush()
        os.fsync(destination.fileno())
    directory = os.open(path.parent, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def replace_json(path: Path, value: Any) -> None:
    """Update the one mutable progress record atomically inside this attempt."""
    temporary = path.with_name(path.name + ".tmp")
    if temporary.exists() or temporary.is_symlink():
        raise RuntimeError(f"stale progress temporary exists: {temporary}")
    with temporary.open("xb") as destination:
        payload = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
        destination.write(payload)
        destination.flush()
        os.fsync(destination.fileno())
    os.replace(temporary, path)
    directory = os.open(path.parent, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
    try:
        os.fsync(directory)
    finally:
        os.close(directory)


def run_git(*args: str) -> bytes:
    result = subprocess.run(
        ["git", *args], cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False
    )
    if result.returncode != 0:
        detail = result.stderr.decode("utf-8", errors="replace").strip()
        raise RuntimeError(f"git {' '.join(args)} failed ({result.returncode}): {detail}")
    return result.stdout


def source_head() -> str:
    return run_git("rev-parse", "HEAD").decode("ascii").strip()


def source_git_diff(attempt: Path, name: str, paths: tuple[Path, ...]) -> dict[str, Any]:
    diff_paths = (*paths, Path("tools/perf-baseline/Cargo.toml"), Path("Cargo.lock"))
    output = run_git("diff", "--no-ext-diff", "--binary", "--", *(str(path) for path in diff_paths))
    path = attempt / name
    with path.open("xb") as destination:
        destination.write(output)
        destination.flush()
        os.fsync(destination.fileno())
    return identity(path)


def same_file_content(left: dict[str, Any] | None, right: dict[str, Any] | None) -> bool:
    if left is None or right is None:
        return left is right
    return all(left.get(field) == right.get(field) for field in ("bytes", "mode", "executable", "sha256"))


def affected_sources(nested_source: str | None) -> tuple[Path, ...]:
    primary = ROOT / PRIMARY_SOURCE
    if primary.is_symlink() or not primary.is_file():
        raise RuntimeError(f"reviewed producer source is missing: {primary}")
    candidates = sorted(
        path
        for path in (ROOT / NESTED_SOURCE_DIR).glob("*.rs")
        if path.is_file() and not path.is_symlink()
    )
    if {path.name for path in candidates} != EXPECTED_NESTED_SOURCES:
        raise RuntimeError(
            "expected the two reviewed nested producer-shape .rs sources; "
            f"found {[str(path.relative_to(ROOT)) for path in candidates]}"
        )
    if nested_source is not None:
        candidate = Path(nested_source)
        nested = (ROOT / candidate if not candidate.is_absolute() else candidate).resolve()
        expected_parent = (ROOT / NESTED_SOURCE_DIR).resolve()
        if nested.parent != expected_parent or nested.suffix != ".rs":
            raise RuntimeError(f"nested source must be one .rs file under {expected_parent}")
        if nested not in candidates:
            raise RuntimeError(f"nested source is missing or symlinked: {nested}")
    return (PRIMARY_SOURCE, *(path.relative_to(ROOT) for path in candidates))


def source_snapshot(attempt: Path, paths: tuple[Path, ...], label: str) -> list[dict[str, Any]]:
    destination_root = attempt / ("source" if label == "before" else "source-after")
    records: list[dict[str, Any]] = []
    for path in paths:
        source = ROOT / path
        source_identity = identity(source)
        destination = destination_root / path
        destination.parent.mkdir(parents=True, exist_ok=True)
        with destination.open("xb") as output:
            output.write(source.read_bytes())
            output.flush()
            os.fsync(output.fileno())
        destination.chmod(stat.S_IMODE(source.stat().st_mode))
        copy_identity = identity(destination)
        if not same_file_content(source_identity, copy_identity):
            raise RuntimeError(f"source snapshot differs from live source: {path}")
        records.append({"relative_path": str(path), "source": source_identity, "snapshot": copy_identity})
    return records


def ensure_target() -> dict[str, Any]:
    if TARGET.is_symlink() or (TARGET.exists() and not TARGET.is_dir()):
        raise RuntimeError(f"dedicated target is not a directory: {TARGET}")
    if not TARGET.exists():
        TARGET.mkdir(mode=0o755)
        with MARKER.open("xb") as destination:
            destination.write(MARKER_TEXT.encode())
            destination.flush()
            os.fsync(destination.fileno())
    if MARKER.is_symlink() or not MARKER.is_file() or MARKER.read_text() != MARKER_TEXT:
        raise RuntimeError(f"dedicated target marker is missing or mismatched: {MARKER}")
    return identity(MARKER)


def environment() -> dict[str, str]:
    inherited = os.environ.copy()
    cargo_home = inherited.get("CARGO_HOME", "/home/zhuhe/.cargo")
    rustup_home = inherited.get("RUSTUP_HOME", "/home/zhuhe/.rustup")
    path = f"{TOOLCHAIN_BIN}:/usr/bin:/bin"
    fixed = {
        "CARGO_BUILD_JOBS": "4",
        "CARGO_INCREMENTAL": "0",
        "CARGO_NET_OFFLINE": "true",
        "CARGO_PROFILE_DEV_DEBUG": "0",
        "CARGO_PROFILE_RELEASE_DEBUG": "0",
        "CARGO_PROFILE_TEST_DEBUG": "0",
        "CARGO_TARGET_DIR": str(TARGET),
        "CARGO_HOME": cargo_home,
        "HOME": inherited.get("HOME", "/home/zhuhe"),
        "LANG": "C",
        "LC_ALL": "C",
        "PATH": path,
        "RUSTDOCFLAGS": "-D warnings",
        "RUSTFLAGS": "",
        "RUSTUP_HOME": rustup_home,
        "RUSTUP_TOOLCHAIN": TOOLCHAIN,
        "TMPDIR": "/tmp",
        "TZ": "UTC",
    }
    inherited.update(fixed)
    return inherited


def recorded_environment(env: dict[str, str]) -> dict[str, str]:
    return {key: env.get(key, "") for key in ENV_KEYS}


def commands(sources: tuple[Path, ...]) -> list[tuple[str, list[str], tuple[str, ...]]]:
    cargo = str(CARGO)
    manifest = str(MANIFEST)
    fmt_sources = [str(path) for path in sources]
    return [
        (
            "rustfmt_changed_sources",
            [str(RUSTFMT), "--check", "--edition", "2024", *fmt_sources],
            ("rustfmt",),
        ),
        (
            "producer_full_cell",
            [cargo, "test", "--manifest-path", manifest, "--locked", "--offline", "--lib", "producer_full_cell", "--", "--nocapture"],
            ("cargo", "rustc"),
        ),
        (
            "producer_shape_selected_route_tests",
            [cargo, "test", "--manifest-path", manifest, "--locked", "--offline", "--lib", "producer_shape::selected_route_tests", "--", "--nocapture"],
            ("cargo", "rustc"),
        ),
        (
            "producer_shape_full_cell_guard_tests",
            [cargo, "test", "--manifest-path", manifest, "--locked", "--offline", "--lib", "producer_shape::full_cell_guard_tests", "--", "--nocapture"],
            ("cargo", "rustc"),
        ),
        (
            "producer_shape_tests",
            [cargo, "test", "--manifest-path", manifest, "--locked", "--offline", "--lib", "producer_shape::tests", "--", "--nocapture"],
            ("cargo", "rustc"),
        ),
        (
            "cargo_clippy",
            [cargo, "clippy", "--manifest-path", manifest, "--locked", "--offline", "--lib", "--", "-D", "warnings"],
            ("cargo", "rustc"),
        ),
        (
            "cargo_doc",
            [cargo, "doc", "--manifest-path", manifest, "--locked", "--offline", "--no-deps", "--lib"],
            ("cargo", "rustc"),
        ),
        (
            "crate_boundaries",
            [PYTHON, "-B", "tools/check_crate_boundaries.py"],
            ("python3",),
        ),
    ]


def tool_versions(env: dict[str, str]) -> dict[str, str]:
    probes = {
        "cargo": [str(CARGO), "--version", "--verbose"],
        "rustc": [str(RUSTC), "--version", "--verbose"],
        "rustfmt": [str(RUSTFMT), "--version", "--verbose"],
        "python3": [PYTHON, "--version"],
    }
    versions: dict[str, str] = {}
    for name, argv in probes.items():
        result = subprocess.run(
            argv,
            cwd=ROOT,
            env=env,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            check=False,
            text=True,
        )
        output = result.stdout
        if result.returncode != 0:
            raise RuntimeError(f"{name} version probe failed ({result.returncode}): {output.strip()}")
        versions[name] = output
    return versions


def runner_copy(attempt: Path) -> dict[str, Any]:
    destination = attempt / "validate.py"
    if destination.exists() or destination.is_symlink():
        raise RuntimeError(f"attempt runner copy already exists: {destination}")
    with destination.open("xb") as output:
        output.write(Path(__file__).read_bytes())
        output.flush()
        os.fsync(output.fileno())
    destination.chmod(0o755)
    return identity(destination)


def run_gate(
    attempt: Path,
    name: str,
    argv: list[str],
    tools: tuple[str, ...],
    env: dict[str, str],
    versions: dict[str, str],
) -> dict[str, Any]:
    log = attempt / "logs" / f"{name}.log"
    started = utc_now()
    monotonic = time.monotonic_ns()
    launch_error: str | None = None
    returncode: int | None = None
    with log.open("xb") as output:
        header = {
            "cwd": str(ROOT),
            "argv": argv,
            "environment": recorded_environment(env),
            "tool_versions": {tool: versions[tool] for tool in tools},
        }
        output.write((json.dumps(header, sort_keys=True) + "\n\n").encode())
        output.flush()
        try:
            result = subprocess.run(argv, cwd=ROOT, env=env, stdout=output, stderr=subprocess.STDOUT, check=False)
            returncode = result.returncode
        except Exception as error:  # retain launch failures in the command record
            launch_error = repr(error)
        output.flush()
        os.fsync(output.fileno())
    text = log.read_text(encoding="utf-8", errors="replace")
    matches = [int(value) for value in TEST_RESULT.findall(text)] if name in TEST_GATES else []
    passed_count = sum(matches)
    failures: list[str] = []
    if launch_error is not None:
        failures.append("launch_error")
    if returncode != 0:
        failures.append("nonzero_exit")
    if name in TEST_GATES and passed_count <= 0:
        failures.append("no_positive_test_count")
    return {
        "name": name,
        "argv": argv,
        "cwd": str(ROOT),
        "environment": recorded_environment(env),
        "tool_versions": {tool: versions[tool] for tool in tools},
        "started_utc": started,
        "finished_utc": utc_now(),
        "elapsed_ns": time.monotonic_ns() - monotonic,
        "returncode": returncode,
        "terminal_exit": returncode,
        "pass": not failures,
        "log": str(log.relative_to(PACKET)),
        "log_identity": identity(log),
        "test_result_matches": matches,
        "passed_count": passed_count if name in TEST_GATES else None,
        "failures": failures,
        **({"launch_error": launch_error} if launch_error is not None else {}),
    }


def valid_attempt_name(value: str) -> bool:
    path = Path(value)
    return bool(value) and not path.is_absolute() and path.name == value and value not in {".", ".."}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run", action="store_true", help="run the eight quality gates")
    parser.add_argument("--attempt", default="quality-v1", help="new attempt directory name")
    parser.add_argument("--nested-source", help="optionally name one of the two reviewed nested .rs paths")
    args = parser.parse_args()
    if not args.run:
        print(json.dumps({"status": "prepared", "performance_claim": False, "message": "no build or gate launched"}, sort_keys=True))
        return 0
    if not valid_attempt_name(args.attempt):
        raise SystemExit("--attempt must be one path component")

    paths = affected_sources(args.nested_source)
    marker = ensure_target()
    attempt = PACKET / "attempts" / args.attempt
    if attempt.exists() or attempt.is_symlink():
        raise SystemExit(f"refusing to reuse existing quality attempt: {attempt}")
    attempt.mkdir(parents=True)
    (attempt / "logs").mkdir()

    env = environment()
    versions = tool_versions(env)
    head_before = source_head()
    sources_before = source_snapshot(attempt, paths, "before")
    diff_before = source_git_diff(attempt, "git-diff-before.txt", paths)
    runner = runner_copy(attempt)
    planned = [{"name": name, "argv": argv} for name, argv, _ in commands(paths)]
    started = utc_now()
    start = {
        "schema": "litchi.xlsx-producer-shape-quality-start-0893.v1",
        "status": "running",
        "performance_claim": False,
        "started_utc": started,
        "cwd": str(ROOT),
        "manifest": str(ROOT / MANIFEST),
        "target": str(TARGET),
        "target_marker": marker,
        "toolchain": TOOLCHAIN,
        "environment": recorded_environment(env),
        "tool_versions": versions,
        "source_head_before": head_before,
        "git_diff_before": diff_before,
        "affected_sources_before": sources_before,
        "affected_source_paths": [str(path) for path in paths],
        "planned_commands": planned,
        "runner": runner,
    }
    write_exclusive(attempt / "quality-start.json", start)

    progress_path = attempt / "quality-progress.json"
    command_records: list[dict[str, Any]] = []
    failures: list[str] = []
    skipped: list[str] = []
    progress = {
        **start,
        "schema": "litchi.xlsx-producer-shape-quality-progress-0893.v1",
        "commands": command_records,
        "failures": failures,
        "skipped_after_first_failure": skipped,
        "next_gate": planned[0]["name"] if planned else None,
    }
    replace_json(progress_path, progress)

    for index, (name, argv, tools) in enumerate(commands(paths)):
        try:
            command = run_gate(attempt, name, argv, tools, env, versions)
        except Exception as error:
            command = {
                "name": name,
                "argv": argv,
                "pass": False,
                "returncode": None,
                "terminal_exit": None,
                "failures": ["runner_error"],
                "runner_error": repr(error),
            }
        command_records.append(command)
        if not command.get("pass"):
            failures.append(name)
            skipped.extend(item[0] for item in commands(paths)[index + 1 :])
        progress["commands"] = command_records
        progress["failures"] = failures
        progress["skipped_after_first_failure"] = skipped
        progress["next_gate"] = None if failures else (planned[index + 1]["name"] if index + 1 < len(planned) else None)
        progress["updated_utc"] = utc_now()
        replace_json(progress_path, progress)
        print(
            f"{name}: {'PASS' if command.get('pass') else 'FAIL'} "
            f"exit={command.get('returncode')} passed={command.get('passed_count')}",
            flush=True,
        )
        if failures:
            break

    post_errors: list[str] = []
    try:
        head_after = source_head()
    except Exception as error:
        head_after = None
        post_errors.append(f"source_head_after:{error!r}")
    try:
        sources_after = source_snapshot(attempt, paths, "after")
    except Exception as error:
        sources_after = []
        post_errors.append(f"affected_sources_after:{error!r}")
    try:
        diff_after = source_git_diff(attempt, "git-diff-after.txt", paths)
    except Exception as error:
        diff_after = None
        post_errors.append(f"git_diff_after:{error!r}")
    if head_after != head_before:
        post_errors.append("source_head_changed")
    if [item.get("source") for item in sources_after] != [item.get("source") for item in sources_before]:
        post_errors.append("affected_sources_changed")
    if not same_file_content(diff_after, diff_before):
        post_errors.append("git_diff_changed")
    failures.extend(post_errors)

    status = "pass" if not failures else "fail"
    terminal = {
        **start,
        "schema": "litchi.xlsx-producer-shape-quality-0893.v1",
        "status": status,
        "finished_utc": utc_now(),
        "commands": command_records,
        "failures": failures,
        "skipped_after_first_failure": skipped,
        "stop_after_first_failure": True,
        "source_head_after": head_after,
        "git_diff_after": diff_after,
        "affected_sources_after": sources_after,
        "progress": identity(progress_path),
    }
    progress.update({
        "status": status,
        "finished_utc": terminal["finished_utc"],
        "source_head_after": head_after,
        "git_diff_after": diff_after,
        "affected_sources_after": sources_after,
    })
    replace_json(progress_path, progress)
    terminal["progress"] = identity(progress_path)
    terminal_path = attempt / "quality-receipt.json"
    write_exclusive(terminal_path, terminal)
    print(json.dumps({"status": status, "attempt": str(attempt), "receipt": str(terminal_path), "failures": failures}, sort_keys=True))
    return 0 if status == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
