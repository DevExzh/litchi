#!/usr/bin/env python3
"""Build and replay the bounded 0773 durability syscall probe.

This runner is deliberately serial and makes no timing claim. It builds one
debug-profile probe against the exact baseline source and one against the exact
integration source, runs the baseline first and the candidate second in the
same owned ``.../tracedir`` directory, archives both raw ``strace`` files,
and invokes :mod:`verify_trace` fail-closed.

The runner does not alter either source worktree.  It creates standalone probe
workspaces and logs below one fresh output directory.  The repository's
workspace locks are recorded separately; a small probe lock is generated once
from the baseline probe and copied byte-for-byte to the candidate probe so the
two binaries have one declared dependency resolution.

Usage::

    python3 trace.py [--output trace-0]

The default roots are the current baseline and this integration worktree.  A
caller may override them when reproducing the packet, but the expected commit
IDs remain enforced unless explicitly changed in the source.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import time

import verify_trace


PACKET = Path(__file__).resolve().parent
# PACKET is .../<repo>/docs/performance/results/change-0773.
INTEGRATION_ROOT = PACKET.parents[3]
DEFAULT_BASELINE = Path("/home/zhuhe/code/litchi")
DEFAULT_CANDIDATE = INTEGRATION_ROOT
DEFAULT_FIXTURE = Path("test-data/ooxml/xlsb/sample.xlsb")
EXPECTED_BASELINE = "6074e10e57"
EXPECTED_CANDIDATE = "339572acbf"
TOOLCHAIN = "1.95.0"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def canonical_sha(value: object) -> str:
    encoded = json.dumps(value, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def git(repo: Path, *arguments: str) -> str:
    result = subprocess.run(
        ["git", *arguments], cwd=repo, check=True, capture_output=True, text=True
    )
    return result.stdout.strip()


def source_clean(repo: Path, reference: str) -> tuple[bool, int]:
    # The baseline worktree may carry unrelated documentation/index work.  The
    # guard is intentionally scoped to source crates, the perf harness, and
    # the manifests that determine this probe's dependency graph, and compares
    # those paths to the requested production reference rather than requiring
    # the worktree HEAD itself to remain at that reference.
    pathspec = [
        "Cargo.toml",
        "Cargo.lock",
        "rust-toolchain.toml",
        ".cargo/config.toml",
        "crates",
        "tools/perf-baseline",
    ]
    changed = subprocess.run(
        ["git", "diff", "--quiet", reference, "--", *pathspec], cwd=repo
    ).returncode
    status = git(repo, "status", "--porcelain", "--untracked-files=all")
    untracked = sum(line.startswith("?? ") for line in status.splitlines())
    source_untracked = any(
        line.startswith("?? ")
        and (
            line[3:].startswith("crates/")
            or line[3:].startswith("tools/perf-baseline/")
            or line[3:] in {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml", ".cargo/config.toml"}
        )
        for line in status.splitlines()
    )
    # Cargo.lock is intentionally ignored by this repository.  It still
    # participates in the build and is recorded below, so fail closed if a
    # worktree has no usable workspace lock at all.
    lock_present = (repo / "Cargo.lock").is_file()
    return changed == 0 and not source_untracked and lock_present, untracked


def source_files(repo: Path) -> list[str]:
    names = git(repo, "ls-files", "-z").split("\0")
    roots = {"Cargo.toml", "Cargo.lock", "rust-toolchain.toml", ".cargo/config.toml"}
    return sorted(
        name
        for name in names
        if name and (name in roots or name.startswith("crates/") or name.startswith("tools/perf-baseline/"))
    )


def source_manifest(repo: Path, reference: str) -> dict[str, object]:
    head = git(repo, "rev-parse", "HEAD")
    clean, untracked = source_clean(repo, reference)
    if not clean:
        raise RuntimeError(
            f"{repo}: probe source crates/harness or manifests differ from {reference}; refusing ambiguous provenance"
        )
    files = {name: sha256(repo / name) for name in source_files(repo)}
    # Cargo.lock is ignored in the production repository, but it is part of
    # the exact dependency input and therefore belongs in this manifest.
    files["Cargo.lock"] = sha256(repo / "Cargo.lock")
    return {
        "repo": str(repo),
        "observed_head": head,
        "reference_revision": reference,
        "source_matches_reference": clean,
        "untracked_file_count": untracked,
        "files": files,
        "source_sha256": canonical_sha(files),
        "workspace_lock_sha256": sha256(repo / "Cargo.lock"),
    }


def verify_source_custody(
    repo: Path, reference: str, initial: dict[str, object], phase: str
) -> dict[str, object]:
    current = source_manifest(repo, reference)
    if (
        current["source_sha256"] != initial["source_sha256"]
        or current["workspace_lock_sha256"] != initial["workspace_lock_sha256"]
    ):
        raise RuntimeError(f"{phase}: source or workspace lock changed in {repo}")
    return {
        "phase": phase,
        "repo": str(repo),
        "observed_head": current["observed_head"],
        "source_sha256": current["source_sha256"],
        "workspace_lock_sha256": current["workspace_lock_sha256"],
    }


def probe_source_manifest() -> dict[str, str]:
    files = {}
    for path in sorted((PACKET / "probe-src").rglob("*")):
        if path.is_file():
            files[str(path.relative_to(PACKET / "probe-src"))] = sha256(path)
    return {"files": files, "source_sha256": canonical_sha(files)}


def write_json(path: Path, value: object) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def command_env(target: Path) -> dict[str, str]:
    env = os.environ.copy()
    env.update(
        {
            "RUSTUP_TOOLCHAIN": TOOLCHAIN,
            "CARGO_TARGET_DIR": str(target),
            "CARGO_BUILD_JOBS": os.environ.get("CARGO_BUILD_JOBS", "2"),
            "CARGO_INCREMENTAL": "0",
            "CARGO_PROFILE_DEV_DEBUG": "0",
        }
    )
    return env


def run_logged(command: list[str], cwd: Path, env: dict[str, str], log: Path) -> dict[str, object]:
    started = time.time()
    log.parent.mkdir(parents=True, exist_ok=True)
    with log.open("w", encoding="utf-8") as output:
        result = subprocess.run(command, cwd=cwd, env=env, stdout=output, stderr=subprocess.STDOUT)
    row = {
        "command": command,
        "cwd": str(cwd),
        "returncode": result.returncode,
        "started": started,
        "ended": time.time(),
        "log": str(log),
        "log_sha256": sha256(log),
    }
    if result.returncode != 0:
        raise RuntimeError(f"command failed ({result.returncode}): {' '.join(command)}; see {log}")
    return row


def prepare_probe(
    leg: str, repo: Path, output: Path, target: Path, levels: bool
) -> tuple[Path, Path, dict[str, object]]:
    build = output / "build" / leg
    probe = build / "probe"
    probe.mkdir(parents=True)
    source_dir = PACKET / "probe-src"
    shutil.copytree(source_dir / "src", probe / "src")
    template = (source_dir / "Cargo.toml.template").read_text(encoding="utf-8")
    cargo_toml = probe / "Cargo.toml"
    cargo_toml.write_text(template.replace("@SRC@", str(repo)), encoding="utf-8")
    workspace_lock = repo / "Cargo.lock"
    shutil.copy2(workspace_lock, build / "workspace.Cargo.lock")
    log = output / "build" / f"{leg}-lock.log"
    env = command_env(target)
    command = [
        "cargo",
        "generate-lockfile",
        "--offline",
        "--manifest-path",
        str(cargo_toml),
    ]
    # Only the baseline generates the probe lock.  The candidate is assigned
    # the exact resulting bytes by main() after both standalone manifests are
    # prepared.
    lock_row = None
    if leg == "before":
        lock_row = run_logged(command, repo, env, log)
    else:
        log.write_text("candidate probe lock deferred to the baseline lock\n", encoding="utf-8")
        lock_row = {
            "command": command,
            "cwd": str(repo),
            "returncode": 0,
            "deferred": True,
            "log": str(log),
            "log_sha256": sha256(log),
        }
    if not (probe / "Cargo.lock").is_file() and leg == "before":
        raise RuntimeError("baseline probe did not produce Cargo.lock")
    metadata = {
        "leg": leg,
        "repo": str(repo),
        "target": str(target),
        "features": ["levels"] if levels else [],
        "cargo_toml_sha256": sha256(cargo_toml),
        "probe_source_sha256": probe_source_manifest()["source_sha256"],
        "workspace_lock_sha256": sha256(workspace_lock),
        "lock_generation": lock_row,
    }
    return probe, target / "debug" / "durability-probe", metadata


def run_probe(
    leg: str,
    binary: Path,
    fixture: Path,
    output: Path,
    expected_windows: int,
    levels: bool,
) -> dict[str, object]:
    tracedir = output / "tracedir"
    if tracedir.exists():
        shutil.rmtree(tracedir)
    tracedir.mkdir(parents=True)
    trace = output / f"{leg}.strace"
    stdout = output / f"{leg}.stdout"
    stderr = output / f"{leg}.stderr"
    command = [
        "strace",
        "-f",
        "-s",
        "256",
        "-yy",
        "-o",
        str(trace),
        str(binary),
        str(tracedir),
        str(fixture),
    ]
    started = time.time()
    with stdout.open("w", encoding="utf-8") as out, stderr.open("w", encoding="utf-8") as err:
        result = subprocess.run(command, cwd=INTEGRATION_ROOT, stdout=out, stderr=err)
    published = output / "published" / leg
    published.mkdir(parents=True, exist_ok=True)
    output_hashes: dict[str, dict[str, object]] = {}
    for path in sorted(tracedir.iterdir()):
        if not path.is_file():
            continue
        shutil.copy2(path, published / path.name)
        output_hashes[path.name] = {
            "bytes": path.stat().st_size,
            "sha256": sha256(path),
        }
    row = {
        "leg": leg,
        "command": command,
        "cwd": str(INTEGRATION_ROOT),
        "features": ["levels"] if levels else [],
        "expected_windows": expected_windows,
        "returncode": result.returncode,
        "started": started,
        "ended": time.time(),
        "trace": trace.name,
        "stdout": stdout.name,
        "stderr": stderr.name,
        "published": str(published.relative_to(output)),
        "output_sha256": output_hashes,
        "trace_sha256": sha256(trace) if trace.exists() else None,
        "stdout_sha256": sha256(stdout),
        "stderr_sha256": sha256(stderr),
    }
    return row


def archive_hashes(output: Path) -> dict[str, str]:
    entries = {}
    for path in sorted(output.rglob("*")):
        if path.is_file() and path.name != "hashes.json":
            entries[str(path.relative_to(output))] = sha256(path)
    manifest = {
        "schema": 1,
        "files": entries,
        "entries_sha256": canonical_sha(entries),
    }
    write_json(output / "hashes.json", manifest)
    return entries


def choose_output(requested: Path | None) -> Path:
    if requested is not None:
        return requested if requested.is_absolute() else PACKET / requested
    index = 0
    while (PACKET / f"trace-{index}").exists():
        index += 1
    return PACKET / f"trace-{index}"


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", type=Path, default=DEFAULT_BASELINE)
    parser.add_argument("--candidate", type=Path, default=DEFAULT_CANDIDATE)
    parser.add_argument("--fixture", type=Path, default=None)
    parser.add_argument("--output", type=Path, default=None)
    parser.add_argument(
        "--before-target",
        type=Path,
        default=Path("/home/zhuhe/code/litchi-target-0773-trace-before"),
    )
    parser.add_argument(
        "--after-target",
        type=Path,
        default=Path("/home/zhuhe/code/litchi-target-0773-quality"),
    )
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)
    baseline = args.baseline.resolve()
    candidate = args.candidate.resolve()
    output = choose_output(args.output).resolve()
    fixture_relative = args.fixture or DEFAULT_FIXTURE
    fixture_before = (fixture_relative if fixture_relative.is_absolute() else baseline / fixture_relative).resolve()
    fixture_after = (fixture_relative if fixture_relative.is_absolute() else candidate / fixture_relative).resolve()
    if output.exists():
        raise SystemExit(f"refusing to overwrite existing replay directory: {output}")
    if not fixture_before.is_file() or not fixture_after.is_file():
        raise SystemExit(f"fixture missing in one leg: {fixture_before}, {fixture_after}")
    output.mkdir(parents=True)

    before_source = source_manifest(baseline, EXPECTED_BASELINE)
    after_source = source_manifest(candidate, EXPECTED_CANDIDATE)
    if sha256(fixture_before) != sha256(fixture_after):
        raise SystemExit("baseline and candidate fixture bytes differ")
    if before_source["workspace_lock_sha256"] != after_source["workspace_lock_sha256"]:
        raise SystemExit("baseline and candidate workspace lock bytes differ")
    probe_source = probe_source_manifest()
    write_json(output / "source-before.json", before_source)
    write_json(output / "source-after.json", after_source)
    write_json(output / "probe-source.json", probe_source)

    before_probe, before_binary, before_meta = prepare_probe(
        "before", baseline, output, args.before_target.resolve(), levels=False
    )
    after_probe, after_binary, after_meta = prepare_probe(
        "after", candidate, output, args.after_target.resolve(), levels=True
    )
    shutil.copy2(before_probe / "Cargo.lock", after_probe / "Cargo.lock")
    probe_lock_sha = sha256(before_probe / "Cargo.lock")
    before_meta["probe_lock_sha256"] = probe_lock_sha
    after_meta["probe_lock_sha256"] = sha256(after_probe / "Cargo.lock")
    if after_meta["probe_lock_sha256"] != probe_lock_sha:
        raise RuntimeError("candidate probe lock differs after exact lock copy")

    build_rows = []
    before_build_command = [
        "cargo",
        "build",
        "--offline",
        "--locked",
        "--manifest-path",
        str(before_probe / "Cargo.toml"),
    ]
    after_build_command = [*before_build_command[:-1], str(after_probe / "Cargo.toml"), "--features", "levels"]
    build_rows.append(
        run_logged(
            before_build_command,
            baseline,
            command_env(args.before_target.resolve()),
            output / "build" / "before-build.log",
        )
    )
    build_rows.append(
        run_logged(
            after_build_command,
            candidate,
            command_env(args.after_target.resolve()),
            output / "build" / "after-build.log",
        )
    )
    post_build_custody = {
        "before": verify_source_custody(baseline, EXPECTED_BASELINE, before_source, "after-build"),
        "after": verify_source_custody(candidate, EXPECTED_CANDIDATE, after_source, "after-build"),
    }
    write_json(output / "source-post-build.json", post_build_custody)
    if not before_binary.is_file() or not after_binary.is_file():
        raise RuntimeError("one probe binary was not produced")
    binary_dir = output / "binary"
    binary_dir.mkdir()
    before_binary_copy = binary_dir / "before-durability-probe"
    after_binary_copy = binary_dir / "after-durability-probe"
    shutil.copy2(before_binary, before_binary_copy)
    shutil.copy2(after_binary, after_binary_copy)
    before_meta["binary_sha256"] = sha256(before_binary_copy)
    after_meta["binary_sha256"] = sha256(after_binary_copy)
    before_meta["binary_source"] = str(before_binary)
    after_meta["binary_source"] = str(after_binary)
    write_json(output / "probe-before.json", before_meta)
    write_json(output / "probe-after.json", after_meta)

    run_manifest = {
        "schema": 1,
        "toolchain": TOOLCHAIN,
        "source_heads": {
            "before": before_source["observed_head"],
            "after": after_source["observed_head"],
        },
        "fixture": {
            "before": str(fixture_before),
            "after": str(fixture_after),
            "sha256": sha256(fixture_before),
        },
        "build_rows": build_rows,
        "legs": {},
    }
    run_manifest["legs"]["before"] = run_probe(
        "before", before_binary_copy, fixture_before, output, expected_windows=24, levels=False
    )
    post_before_run_custody = {
        "before": verify_source_custody(baseline, EXPECTED_BASELINE, before_source, "after-before-run"),
        "after": verify_source_custody(candidate, EXPECTED_CANDIDATE, after_source, "after-before-run"),
    }
    write_json(output / "source-post-before-run.json", post_before_run_custody)
    if run_manifest["legs"]["before"]["returncode"] != 0:
        write_json(output / "run.json", run_manifest)
        archive_hashes(output)
        raise RuntimeError("baseline probe failed; candidate replay was not started")
    run_manifest["legs"]["after"] = run_probe(
        "after", after_binary_copy, fixture_after, output, expected_windows=96, levels=True
    )
    post_run_custody = {
        "before": verify_source_custody(baseline, EXPECTED_BASELINE, before_source, "after-run"),
        "after": verify_source_custody(candidate, EXPECTED_CANDIDATE, after_source, "after-run"),
    }
    write_json(output / "source-post-run.json", post_run_custody)
    write_json(output / "run.json", run_manifest)

    windows = output / "windows.json"
    verification = verify_trace.verify(
        output / "before.strace", output / "after.strace", output / "run.json", windows
    )
    if not verification["ok"]:
        archive_hashes(output)
        raise RuntimeError("syscall replay verification failed; see windows.json")
    archive_hashes(output)
    provenance = {
        "schema": 1,
        "baseline": before_source,
        "candidate": after_source,
        "probe_source": probe_source,
        "fixture_sha256": sha256(fixture_before),
        "probe_lock_sha256": probe_lock_sha,
        "binary_sha256": {
            "before": before_meta["binary_sha256"],
            "after": after_meta["binary_sha256"],
        },
        "run_manifest": "run.json",
        "verification": "windows.json",
        "source_custody": [
            "source-post-build.json",
            "source-post-before-run.json",
            "source-post-run.json",
        ],
        "raw_traces": ["before.strace", "after.strace"],
        "build_profile": {
            "profile": "dev",
            "debug": "0",
            "incremental": "0",
            "timing_claim": False,
        },
    }
    write_json(output / "provenance.json", provenance)
    archive_hashes(output)
    print(json.dumps({"output": str(output), "ok": True, "entries": len(provenance["raw_traces"])}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main(sys.argv[1:]))
    except (RuntimeError, subprocess.CalledProcessError) as error:
        print(f"trace.py: {error}", file=sys.stderr)
        raise SystemExit(1) from error
