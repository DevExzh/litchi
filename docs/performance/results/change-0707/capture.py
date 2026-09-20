#!/usr/bin/env python3
"""Capture the serialized 0707 XLSX planning matrix.

This driver never builds a binary.  ``build.py`` freezes the native and
allocator binaries first; this script binds every child to those frozen
artifacts, the exact source census, the plan, and the constraints.  Native,
allocator, and Callgrind phases are invoked separately so the coordinator can
control source transitions and avoid overlapping Cargo or benchmark work.

Canonical phases are:

    native-baseline allocator-baseline profile-baseline
    native-candidate allocator-candidate profile-candidate

The candidate phases are intentionally available even though the baseline
phase is the first operation in this batch.  Every child is refusal-safe: an
existing artifact is never replaced.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
BIN_ROOT = REPO.parent / "litchi-0707-bin"

NATIVE_PHASES = {
    "native-baseline": ("baseline", "native"),
    "native-candidate": ("candidate", "native"),
    "allocator-baseline": ("baseline", "alloc"),
    "allocator-candidate": ("candidate", "alloc"),
    "profile-baseline": ("baseline", "profile"),
    "profile-candidate": ("candidate", "profile"),
}

ALIASES = {
    "baseline": "native-baseline",
    "candidate": "native-candidate",
    "native-A": "native-baseline",
    "native-B": "native-candidate",
    "allocator-A": "allocator-baseline",
    "allocator-B": "allocator-candidate",
    "alloc-A": "allocator-baseline",
    "alloc-B": "allocator-candidate",
    "profile-A": "profile-baseline",
    "profile-B": "profile-candidate",
}


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest_json(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def load_json(path: Path) -> Any:
    require(path.is_file(), f"missing JSON input: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise RuntimeError(f"invalid JSON in {path}: {error}") from error


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def plan() -> dict[str, Any]:
    value = load_json(HERE / "plan.json")
    require(isinstance(value, dict), "plan is not an object")
    return value


def profile_plan() -> dict[str, Any]:
    value = load_json(HERE / "profile-plan.json")
    require(isinstance(value, dict), "profile plan is not an object")
    return value


def source_census() -> dict[str, str]:
    """Match the exact Rust/source census used by build.py."""

    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path
            for path in (REPO / folder).rglob("*")
            if path.is_file()
            and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {
        str(path.relative_to(REPO)): sha(path)
        for path in sorted(set(paths))
    }


def constraints_check() -> None:
    constraints = load_json(HERE / "constraints.json")
    require(isinstance(constraints, dict), "constraints is not an object")
    for name, digest in constraints.items():
        path = REPO / name
        require(path.is_file(), f"constraint input is absent: {name}")
        require(sha(path) == digest, f"constraint changed during capture: {name}")


def normalized_phase(raw: str) -> str:
    phase = ALIASES.get(raw, raw)
    require(phase in NATIVE_PHASES, f"unknown phase: {raw}")
    return phase


def build_record(role: str, lane: str) -> tuple[dict[str, Any], dict[str, str], Path]:
    records_path = HERE / f"build-{role}.json"
    records = load_json(records_path)
    require(isinstance(records, list), f"{records_path} is not a build-record list")
    binary_name = f"{role}-{'alloc' if lane == 'alloc' else 'native'}"
    matches = [
        item
        for item in records
        if isinstance(item, dict) and Path(str(item.get("binary", ""))).name == binary_name
    ]
    require(len(matches) == 1, f"{records_path}: expected one {binary_name} record")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{records_path}: {binary_name} build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    require(binary.is_file() and not binary.is_symlink(), f"missing frozen binary: {binary}")
    binary_digest = record.get("binary_sha256")
    require(isinstance(binary_digest, str) and len(binary_digest) == 64,
            f"{records_path}: binary digest is missing")
    require(sha(binary) == binary_digest, f"frozen binary digest changed: {binary}")
    source_path = HERE / f"source-{role}.json"
    expected = load_json(source_path)
    require(isinstance(expected, dict), f"{source_path} is not a source map")
    require(record.get("source_manifest_sha256") == sha(source_path),
            f"{records_path}: source manifest binding differs")
    return record, expected, binary


def source_relation(expected: dict[str, str], current: dict[str, str], name: str) -> dict[str, Any]:
    changed = sorted(
        path for path in set(expected) | set(current)
        if expected.get(path) != current.get(path)
    )
    require(not changed, f"{name}: current source differs from frozen source: {changed}")
    return {
        "mode": "exact",
        "changed_paths": changed,
        "expected_entry_count": len(expected),
        "current_entry_count": len(current),
    }


def ordered_jobs(p: dict[str, Any], kind: str, role: str, lane: str) -> Iterable[dict[str, Any]]:
    if lane == "native":
        repeats = int(p["native_repeats"])
        samples = int(p["samples"])
        warmup = int(p["warmup"])
    elif lane == "alloc":
        repeats = int(p["allocation_repeats"])
        samples = int(p["allocation_samples"])
        warmup = int(p["allocation_warmup"])
    else:
        repeats = int(p["profile_repeats"])
        samples = int(p["profile_samples"])
        warmup = int(p["profile_warmup"])
    for repeat in range(1, repeats + 1):
        shapes = list(p["shapes"])
        if repeat == 2:
            shapes.reverse()
        for shape in shapes:
            prefix = {"native": "native", "alloc": "alloc", "profile": "profile"}[lane]
            yield {
                "kind": kind,
                "role": role,
                "lane": lane,
                "repeat": repeat,
                "shape": shape,
                "case": p["case"],
                "samples": samples,
                "warmup": warmup,
                "name": f"{prefix}-{role}-r{repeat}-{shape}",
            }


def profile_command(
    job: dict[str, Any], p: dict[str, Any], profile: dict[str, Any], binary: Path,
    output_path: Path, callgrind_path: Path,
) -> list[str]:
    command = ["taskset", "-c", str(p["cpu"]), "valgrind"]
    command.extend(profile["options"])
    command.append(f"--callgrind-out-file={callgrind_path}")
    command.extend(
        [
            str(binary),
            "--warmup", str(job["warmup"]),
            "--samples", str(job["samples"]),
            "--case", job["case"],
            "--xlsx-cell-crud-shape", job["shape"],
            "--json", str(output_path),
        ]
    )
    return command


def child_artifacts(name: str, profile: bool) -> list[Path]:
    paths = [HERE / f"{name}.json", HERE / f"{name}.stdout", HERE / f"{name}.stderr"]
    if profile:
        paths.extend(sorted(HERE.glob(f"{name}.callgrind*")))
    return paths


def validate_profile_matrix(p: dict[str, Any], profile: dict[str, Any]) -> None:
    require(profile.get("shapes") == p.get("shapes"), "profile shapes differ from plan")
    require(profile.get("repeats") == p.get("profile_repeats"),
            "profile repeats differ from plan")
    require(profile.get("warmup") == p.get("profile_warmup"),
            "profile warmup differs from plan")
    require(profile.get("samples") == p.get("profile_samples"),
            "profile samples differ from plan")
    required = {
        "--tool=callgrind",
        "--collect-atstart=no",
        f"--toggle-collect={profile.get('owner')}",
        f"--zero-before={profile.get('owner')}",
        f"--dump-after={profile.get('owner')}",
    }
    options = profile.get("options")
    require(isinstance(options, list) and required.issubset(options),
            "profile Callgrind options are incomplete")
    require(profile.get("expected_lifecycle_calls") == 3,
            "profile lifecycle call policy differs")
    require(profile.get("expected_measured_calls") == 1,
            "profile measured call policy differs")


def run_child(
    job: dict[str, Any], p: dict[str, Any], profile: dict[str, Any],
    build: dict[str, Any], expected: dict[str, str], binary: Path,
) -> None:
    name = job["name"]
    is_profile = job["lane"] == "profile"
    receipt_path = HERE / f"{name}.receipt.json"
    require(not receipt_path.exists(), f"refusing to replace receipt: {receipt_path}")
    require(not any(path.exists() or path.is_symlink() for path in child_artifacts(name, is_profile)),
            f"refusing to replace child artifacts: {name}")
    constraints_check()
    before = source_census()
    relation_before = source_relation(expected, before, name)
    before_path = HERE / f"{name}.source-before.json"
    after_path = HERE / f"{name}.source-after.json"
    output_path = HERE / f"{name}.json"
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    callgrind_path = HERE / f"{name}.callgrind"
    require(not before_path.exists() and not after_path.exists(),
            f"refusing to replace source custody artifacts: {name}")
    write_json(before_path, before)
    if is_profile:
        command = profile_command(job, p, profile, binary, output_path, callgrind_path)
    else:
        command = [
            "taskset", "-c", str(p["cpu"]), str(binary),
            "--warmup", str(job["warmup"]),
            "--samples", str(job["samples"]),
            "--case", job["case"],
            "--xlsx-cell-crud-shape", job["shape"],
            "--json", str(output_path),
        ]
    started = utc_now()
    tick = time.monotonic()
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        result = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr)
    elapsed = time.monotonic() - tick

    after = source_census()
    relation_after = source_relation(expected, after, name)
    write_json(after_path, after)
    require(before == after, f"{name}: source census changed during child")
    require(sha(binary) == build["binary_sha256"], f"{name}: frozen binary changed")

    artifacts: dict[str, str] = {}
    for path in sorted(child_artifacts(name, is_profile) + [before_path, after_path]):
        require(path.is_file() and not path.is_symlink(), f"{name}: missing artifact {path.name}")
        artifacts[path.name] = sha(path)
    receipt = {
        "schema_version": 1,
        "name": name,
        "kind": job["kind"],
        "role": job["role"],
        "lane": job["lane"],
        "repeat": job["repeat"],
        "shape": job["shape"],
        "case": job["case"],
        "command": command,
        "start_utc": started,
        "end_utc": utc_now(),
        "seconds": elapsed,
        "exit_code": result.returncode,
        "cpu": p["cpu"],
        "binary_path": str(binary),
        "binary_sha256": sha(binary),
        "binary_bytes": binary.stat().st_size,
        "build_record_sha256": sha(HERE / f"build-{job['role']}.json"),
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "retained_source_manifest": f"source-{job['role']}.json",
        "retained_source_manifest_sha256": sha(HERE / f"source-{job['role']}.json"),
        "retained_source_census_sha256": digest_json(expected),
        "retained_source_entry_count": len(expected),
        "current_checkout_source": {
            "before_artifact": before_path.name,
            "after_artifact": after_path.name,
            "before_sha256": digest_json(before),
            "after_sha256": digest_json(after),
            "before_file_sha256": sha(before_path),
            "after_file_sha256": sha(after_path),
            "before_entry_count": len(before),
            "after_entry_count": len(after),
            "relation_before": relation_before,
            "relation_after": relation_after,
            "unchanged_during_child": before == after,
        },
        "plan_sha256": sha(HERE / "plan.json"),
        "script_sha256": sha(Path(__file__)),
        "constraints_sha256": sha(HERE / "constraints.json"),
        "profile_plan_sha256": sha(HERE / "profile-plan.json") if is_profile else None,
        "environment": {
            key: os.environ.get(key)
            for key in ("RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF", "GLIBC_TUNABLES")
        },
        "artifacts": artifacts,
    }
    write_json(receipt_path, receipt)
    require(result.returncode == 0, f"{name}: child failed with exit code {result.returncode}")
    if is_profile:
        dumps = sorted(HERE.glob(f"{name}.callgrind.*"))
        require(dumps, f"{name}: no numbered Callgrind dumps were retained")
    print(f"{name} passed", flush=True)


def run_phase(raw_phase: str) -> None:
    phase = normalized_phase(raw_phase)
    p = plan()
    profile = profile_plan()
    validate_profile_matrix(p, profile)
    role, lane = NATIVE_PHASES[phase]
    build_lane = "alloc" if lane == "alloc" else "native"
    build, expected, binary = build_record(role, build_lane)
    jobs = list(ordered_jobs(p, "profile" if lane == "profile" else lane, role, lane))
    require(jobs, f"{phase}: no jobs selected")
    for job in jobs:
        run_child(job, p, profile, build, expected, binary)
    print(f"phase {phase} complete ({len(jobs)} children)", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("phase", help="serialized native, allocator, or profile phase")
    args = parser.parse_args()
    try:
        run_phase(args.phase)
    except (OSError, RuntimeError, ValueError) as error:
        print(f"capture failed: {error}", file=os.sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
