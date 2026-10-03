#!/usr/bin/env python3
"""Run the bounded 0894 selected-metadata quality gates.

The runner is deliberately create-only per attempt.  It records the exact
candidate source inventory, source copies, commands, environment, tool
versions, logs, and terminal result.  It never launches a workload or claims
performance; ``--run`` is the only mode that invokes Cargo or creates the
workspace-owned quality target.  The default and ``--print-plan`` modes are
read-only.

The candidate is allowed to add one implementation helper under the selected
worksheet module.  The required paths and the optional path are explicit so a
new source file cannot enter a quality attempt accidentally, while the
``--candidate-source`` option lets the coordinator freeze the exact inventory
when the implementation shape is chosen.
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
from typing import Any, Iterable


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
MANIFEST = Path("crates/litchi-xlsx/Cargo.toml")
PERF_MANIFEST = Path("tools/perf-baseline/Cargo.toml")
LOCKFILE = Path("Cargo.lock")

TARGET = ROOT / "target" / "litchi-0894-quality"
MARKER = TARGET / ".owner-0894-quality"
MARKER_TEXT = "litchi 0894 selected-metadata quality target\n"
PRIVATE_TMP = TARGET / "private-tmp"

TOOLCHAIN = "1.95.0-x86_64-unknown-linux-gnu"
TOOLCHAIN_BIN = Path("/home/zhuhe/.rustup/toolchains") / TOOLCHAIN / "bin"
CARGO = TOOLCHAIN_BIN / "cargo"
RUSTFMT = TOOLCHAIN_BIN / "rustfmt"
RUSTC = TOOLCHAIN_BIN / "rustc"
PYTHON = "python3"

SELECTED_SOURCE = Path("crates/litchi-xlsx/src/raw/worksheet/selected.rs")
METADATA_SOURCE = Path("crates/litchi-xlsx/src/raw/worksheet/selected/producer_metadata.rs")
METADATA_TESTS = Path("crates/litchi-xlsx/src/raw/worksheet/selected/producer_metadata_tests.rs")
WORKSHEET_TESTS = Path("crates/litchi-xlsx/src/raw/worksheet/tests.rs")
HARNESS_SOURCE = Path("tools/perf-baseline/src/producer_shape/selected_route_tests.rs")
REQUIRED_CANDIDATE_SOURCES = frozenset(
    {SELECTED_SOURCE, METADATA_TESTS, WORKSHEET_TESTS, HARNESS_SOURCE}
)
OPTIONAL_CANDIDATE_SOURCES = frozenset({METADATA_SOURCE})
ALLOWED_CANDIDATE_SOURCES = REQUIRED_CANDIDATE_SOURCES | OPTIONAL_CANDIDATE_SOURCES

TEST_RESULT = re.compile(r"test result: (?:ok|FAILED)\.\s+(\d+) passed;")
FOCUSED_GATES = frozenset(
    {
        "focused_producer_metadata_tests",
        "xlsx_lib_tests",
        "producer_shape_harness_tests",
        "producer_full_cell_selector_test",
    }
)

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
    original = Path(path)
    if original.is_symlink():
        raise RuntimeError(f"expected a regular non-symlink file: {original}")
    path = original.resolve()
    if not path.is_file():
        raise RuntimeError(f"expected a regular non-symlink file: {path}")
    metadata = path.stat()
    return {
        "path": str(path),
        "bytes": metadata.st_size,
        "mode": stat.S_IMODE(metadata.st_mode),
        "executable": bool(metadata.st_mode & stat.S_IXUSR),
        "sha256": sha256(path),
    }


def same_identity(left: dict[str, Any] | None, right: dict[str, Any] | None) -> bool:
    if left is None or right is None:
        return left is right
    return all(left.get(field) == right.get(field) for field in ("bytes", "mode", "executable", "sha256"))


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
    """Atomically update the one mutable progress record in an attempt."""

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


def candidate_path(raw: str | Path) -> Path:
    value = Path(raw)
    if value.is_absolute():
        resolved = value.resolve()
    else:
        resolved = (ROOT / value).resolve()
    try:
        relative = resolved.relative_to(ROOT)
    except ValueError as error:
        raise RuntimeError(f"candidate source escapes repository: {raw}") from error
    if relative not in ALLOWED_CANDIDATE_SOURCES:
        allowed = ", ".join(str(path) for path in sorted(ALLOWED_CANDIDATE_SOURCES, key=str))
        raise RuntimeError(f"candidate source is outside the reviewed inventory: {relative}; allowed: {allowed}")
    return relative


def candidate_sources(explicit: Iterable[str] | None = None) -> tuple[Path, ...]:
    """Resolve the exact candidate inventory without assuming one module shape."""

    optional = ROOT / METADATA_SOURCE
    if optional.is_symlink() or (optional.exists() and not optional.is_file()):
        raise RuntimeError(f"optional candidate source is not a regular file: {optional}")
    if explicit is None:
        selected = set(REQUIRED_CANDIDATE_SOURCES)
        if optional.is_file() and not optional.is_symlink():
            selected.add(METADATA_SOURCE)
    else:
        raw_paths = list(explicit)
        selected = {candidate_path(raw) for raw in raw_paths}
        if len(selected) != len(raw_paths):
            raise RuntimeError("candidate inventory contains duplicate paths")
    missing = REQUIRED_CANDIDATE_SOURCES - selected
    if missing:
        raise RuntimeError("candidate inventory omits required paths: " + ", ".join(sorted(map(str, missing))))
    unexpected = selected - ALLOWED_CANDIDATE_SOURCES
    if unexpected:
        raise RuntimeError("candidate inventory contains unreviewed paths: " + ", ".join(sorted(map(str, unexpected))))
    ordered = [SELECTED_SOURCE]
    if METADATA_SOURCE in selected:
        ordered.append(METADATA_SOURCE)
    ordered.extend((METADATA_TESTS, WORKSHEET_TESTS, HARNESS_SOURCE))
    for path in ordered:
        identity(ROOT / path)
    return tuple(ordered)


def source_git_diff(attempt: Path, name: str, paths: tuple[Path, ...]) -> dict[str, Any]:
    """Record only the candidate and lock/manifest diff; unrelated docs stay out."""

    diff_paths = (*paths, MANIFEST, LOCKFILE)
    output = run_git("diff", "--no-ext-diff", "--binary", "--", *(str(path) for path in diff_paths))
    destination = attempt / name
    with destination.open("xb") as output_file:
        output_file.write(output)
        output_file.flush()
        os.fsync(output_file.fileno())
    return identity(destination)


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
        if not same_identity(source_identity, copy_identity):
            raise RuntimeError(f"source snapshot differs from live source: {path}")
        records.append({"relative_path": str(path), "source": source_identity, "snapshot": copy_identity})
    return records


def build_input_identities() -> dict[str, dict[str, Any]]:
    return {str(path): identity(ROOT / path) for path in (MANIFEST, LOCKFILE)}


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
    if PRIVATE_TMP.is_symlink() or (PRIVATE_TMP.exists() and not PRIVATE_TMP.is_dir()):
        raise RuntimeError(f"private TMPDIR is not a directory: {PRIVATE_TMP}")
    if not PRIVATE_TMP.exists():
        PRIVATE_TMP.mkdir(mode=0o700)
    return {"target": identity(MARKER), "private_tmp": identity(PRIVATE_TMP) if PRIVATE_TMP.is_file() else {"path": str(PRIVATE_TMP.resolve()), "mode": stat.S_IMODE(PRIVATE_TMP.stat().st_mode)}}


def environment() -> dict[str, str]:
    inherited = os.environ.copy()
    # Do not let a caller inject compiler flags or a wrapper into a quality
    # attempt.  The full inherited mapping is used only as a starting point;
    # receipts serialize the allowlisted ``ENV_KEYS`` below.
    for key in (
        "CARGO_BUILD_RUSTC",
        "CARGO_BUILD_RUSTC_WRAPPER",
        "CARGO_BUILD_RUSTC_WORKSPACE_WRAPPER",
        "CARGO_ENCODED_RUSTFLAGS",
        "RUSTC",
        "RUSTC_WRAPPER",
        "RUSTC_WORKSPACE_WRAPPER",
        "RUSTFLAGS",
    ):
        inherited.pop(key, None)
    cargo_home = inherited.get("CARGO_HOME", "/home/zhuhe/.cargo")
    rustup_home = inherited.get("RUSTUP_HOME", "/home/zhuhe/.rustup")
    inherited.update(
        {
            "CARGO_BUILD_JOBS": "2",
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
            "PATH": f"{TOOLCHAIN_BIN}:/usr/bin:/bin",
            "RUSTDOCFLAGS": "-D warnings",
            "RUSTFLAGS": "",
            "RUSTUP_HOME": rustup_home,
            "RUSTUP_TOOLCHAIN": TOOLCHAIN,
            "TMPDIR": str(PRIVATE_TMP),
            "TZ": "UTC",
        }
    )
    return inherited


def recorded_environment(env: dict[str, str]) -> dict[str, str]:
    return {key: env.get(key, "") for key in ENV_KEYS}


def commands() -> list[tuple[str, list[str], tuple[str, ...]]]:
    cargo = str(CARGO)
    manifest = str(MANIFEST)
    perf_manifest = str(PERF_MANIFEST)
    return [
        (
            "rustfmt_candidate_sources",
            [str(RUSTFMT), "--check", "--edition", "2024"],
            ("rustfmt",),
        ),
        (
            "focused_producer_metadata_tests",
            [cargo, "test", "--manifest-path", manifest, "--locked", "--offline", "--lib", "producer_metadata", "--", "--nocapture"],
            ("cargo", "rustc"),
        ),
        (
            "xlsx_lib_tests",
            [cargo, "test", "--manifest-path", manifest, "--locked", "--offline", "--lib", "--", "--nocapture"],
            ("cargo", "rustc"),
        ),
        (
            "xlsx_all_features_check",
            [cargo, "check", "--manifest-path", manifest, "--locked", "--offline", "--all-features", "--lib"],
            ("cargo", "rustc"),
        ),
        (
            "xlsx_clippy_all_features",
            [cargo, "clippy", "--manifest-path", manifest, "--locked", "--offline", "--all-features", "--lib", "--", "-D", "warnings"],
            ("cargo", "rustc"),
        ),
        (
            "xlsx_rustdoc",
            [cargo, "doc", "--manifest-path", manifest, "--locked", "--offline", "--all-features", "--no-deps", "--lib"],
            ("cargo", "rustc"),
        ),
        (
            "producer_shape_harness_tests",
            [cargo, "test", "--manifest-path", perf_manifest, "--locked", "--offline", "--lib", "producer_shape::", "--", "--nocapture"],
            ("cargo", "rustc"),
        ),
        (
            "producer_full_cell_selector_test",
            [cargo, "test", "--manifest-path", perf_manifest, "--locked", "--offline", "--lib", "producer_full_cell_scan_selector", "--", "--nocapture"],
            ("cargo", "rustc"),
        ),
        (
            "crate_boundaries",
            [PYTHON, "-B", "tools/check_crate_boundaries.py"],
            ("python3",),
        ),
    ]


def rendered_commands(candidate: tuple[Path, ...]) -> list[dict[str, Any]]:
    return [
        {
            "name": name,
            "argv": [*argv, *(str(path) for path in candidate)]
            if name == "rustfmt_candidate_sources"
            else argv,
            "tools": tools,
        }
        for name, argv, tools in commands()
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
        if result.returncode != 0:
            raise RuntimeError(f"{name} version probe failed ({result.returncode}): {result.stdout.strip()}")
        versions[name] = result.stdout
    return versions


def runner_copy(attempt: Path) -> dict[str, Any]:
    destination = attempt / "quality.py"
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
    candidate: tuple[Path, ...],
) -> dict[str, Any]:
    if name == "rustfmt_candidate_sources":
        argv = [*argv, *(str(path) for path in candidate)]
    log = attempt / "logs" / f"{name}.log"
    started = utc_now()
    monotonic = time.monotonic_ns()
    launch_error: str | None = None
    returncode: int | None = None
    with log.open("xb") as output:
        header = {
            "cwd": str(ROOT),
            "argv": argv,
            "candidate_sources": [str(path) for path in candidate],
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
    matches = [int(value) for value in TEST_RESULT.findall(text)] if name in FOCUSED_GATES else []
    passed_count = sum(matches)
    failures: list[str] = []
    if launch_error is not None:
        failures.append("launch_error")
    if returncode != 0:
        failures.append("nonzero_exit")
    if name in FOCUSED_GATES and passed_count <= 0:
        failures.append("no_positive_test_count")
    return {
        "name": name,
        "argv": argv,
        "cwd": str(ROOT),
        "candidate_sources": [str(path) for path in candidate],
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
        "passed_count": passed_count if name in FOCUSED_GATES else None,
        "failures": failures,
        **({"launch_error": launch_error} if launch_error is not None else {}),
    }


def valid_attempt_name(value: str) -> bool:
    path = Path(value)
    return bool(value) and not path.is_absolute() and path.name == value and value not in {".", ".."}


def print_plan(candidate: tuple[Path, ...] | None = None) -> int:
    planned = commands()
    if candidate is None:
        candidate = tuple(sorted(REQUIRED_CANDIDATE_SOURCES, key=str))
    rendered = []
    for name, argv, _tools in planned:
        rendered.append({"name": name, "argv": [*argv, *(str(path) for path in candidate)] if name == "rustfmt_candidate_sources" else argv})
    print(
        json.dumps(
            {
                "status": "prepared",
                "schema": "litchi.xlsx-selected-metadata-quality-plan-0894.v1",
                "performance_claim": False,
                "run_required": True,
                "candidate_sources": [str(path) for path in candidate],
                "target": str(TARGET),
                "private_tmp": str(PRIVATE_TMP),
                "toolchain": TOOLCHAIN,
                "environment": {
                    "CARGO_BUILD_JOBS": "2",
                    "CARGO_INCREMENTAL": "0",
                    "CARGO_NET_OFFLINE": "true",
                    "CARGO_PROFILE_DEV_DEBUG": "0",
                    "CARGO_PROFILE_TEST_DEBUG": "0",
                    "TMPDIR": str(PRIVATE_TMP),
                },
                "commands": rendered,
            },
            sort_keys=True,
        )
    )
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--run", action="store_true", help="run all quality gates sequentially")
    parser.add_argument("--print-plan", action="store_true", help="print the frozen gate plan without creating files")
    parser.add_argument("--attempt", default="quality-v1", help="new immutable attempt directory name")
    parser.add_argument(
        "--candidate-source",
        action="append",
        dest="candidate_sources",
        help="candidate Rust source path; repeat to freeze the exact reviewed inventory",
    )
    args = parser.parse_args()
    if args.run and args.print_plan:
        raise SystemExit("--run and --print-plan are mutually exclusive")
    if args.run and not valid_attempt_name(args.attempt):
        raise SystemExit("--attempt must be one path component")

    candidate = candidate_sources(args.candidate_sources) if args.candidate_sources is not None or args.run or args.print_plan else None
    if not args.run:
        return print_plan(candidate)

    marker = ensure_target()
    assert candidate is not None
    attempt = PACKET / "attempts" / args.attempt
    if attempt.exists() or attempt.is_symlink():
        raise SystemExit(f"refusing to reuse existing quality attempt: {attempt}")
    attempt.mkdir(parents=True)
    (attempt / "logs").mkdir()

    env = environment()
    versions = tool_versions(env)
    head_before = source_head()
    sources_before = source_snapshot(attempt, candidate, "before")
    build_inputs_before = build_input_identities()
    diff_before = source_git_diff(attempt, "git-diff-before.txt", candidate)
    runner = runner_copy(attempt)
    planned = rendered_commands(candidate)
    start = {
        "schema": "litchi.xlsx-selected-metadata-quality-start-0894.v1",
        "status": "running",
        "performance_claim": False,
        "started_utc": utc_now(),
        "cwd": str(ROOT),
        "manifest": str(ROOT / MANIFEST),
        "perf_manifest": str(ROOT / PERF_MANIFEST),
        "target": str(TARGET),
        "target_marker": marker,
        "private_tmp": str(PRIVATE_TMP),
        "toolchain": TOOLCHAIN,
        "environment": recorded_environment(env),
        "tool_versions": versions,
        "source_head_before": head_before,
        "git_diff_before": diff_before,
        "candidate_sources_before": sources_before,
        "build_inputs_before": build_inputs_before,
        "candidate_source_paths": [str(path) for path in candidate],
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
        "schema": "litchi.xlsx-selected-metadata-quality-progress-0894.v1",
        "commands": command_records,
        "failures": failures,
        "skipped_after_first_failure": skipped,
        "next_gate": planned[0]["name"] if planned else None,
    }
    replace_json(progress_path, progress)

    planned_commands = commands()
    for index, (name, argv, tools) in enumerate(planned_commands):
        try:
            command = run_gate(attempt, name, argv, tools, env, versions, candidate)
        except Exception as error:
            command = {
                "name": name,
                "argv": argv,
                "candidate_sources": [str(path) for path in candidate],
                "pass": False,
                "returncode": None,
                "terminal_exit": None,
                "failures": ["runner_error"],
                "runner_error": repr(error),
            }
        command_records.append(command)
        if not command.get("pass"):
            failures.append(name)
            skipped.extend(item[0] for item in planned_commands[index + 1 :])
        progress["commands"] = command_records
        progress["failures"] = failures
        progress["skipped_after_first_failure"] = skipped
        progress["next_gate"] = None if failures else (planned[index + 1]["name"] if index + 1 < len(planned) else None)
        progress["updated_utc"] = utc_now()
        replace_json(progress_path, progress)
        print(f"{name}: {'PASS' if command.get('pass') else 'FAIL'} exit={command.get('returncode')} passed={command.get('passed_count')}", flush=True)
        if failures:
            break

    post_errors: list[str] = []
    try:
        head_after = source_head()
    except Exception as error:
        head_after = None
        post_errors.append(f"source_head_after:{error!r}")
    try:
        sources_after = source_snapshot(attempt, candidate, "after")
    except Exception as error:
        sources_after = []
        post_errors.append(f"candidate_sources_after:{error!r}")
    try:
        build_inputs_after = build_input_identities()
    except Exception as error:
        build_inputs_after = {}
        post_errors.append(f"build_inputs_after:{error!r}")
    try:
        diff_after = source_git_diff(attempt, "git-diff-after.txt", candidate)
    except Exception as error:
        diff_after = None
        post_errors.append(f"git_diff_after:{error!r}")
    if head_after != head_before:
        post_errors.append("source_head_changed")
    if [item.get("source") for item in sources_after] != [item.get("source") for item in sources_before]:
        post_errors.append("candidate_sources_changed")
    if build_inputs_after != build_inputs_before:
        post_errors.append("build_inputs_changed")
    if not same_identity(diff_after, diff_before):
        post_errors.append("git_diff_changed")
    failures.extend(post_errors)

    status = "pass" if not failures else "fail"
    terminal = {
        **start,
        "schema": "litchi.xlsx-selected-metadata-quality-receipt-0894.v1",
        "status": status,
        "finished_utc": utc_now(),
        "commands": command_records,
        "failures": failures,
        "skipped_after_first_failure": skipped,
        "stop_after_first_failure": True,
        "source_head_after": head_after,
        "git_diff_after": diff_after,
        "candidate_sources_after": sources_after,
        "build_inputs_after": build_inputs_after,
        "progress": identity(progress_path),
    }
    progress.update(
        {
            "status": status,
            "finished_utc": terminal["finished_utc"],
            "source_head_after": head_after,
            "git_diff_after": diff_after,
            "candidate_sources_after": sources_after,
        }
    )
    replace_json(progress_path, progress)
    terminal["progress"] = identity(progress_path)
    terminal_path = attempt / "quality-receipt.json"
    write_exclusive(terminal_path, terminal)
    print(json.dumps({"status": status, "attempt": str(attempt), "receipt": str(terminal_path), "failures": failures}, sort_keys=True))
    return 0 if status == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
