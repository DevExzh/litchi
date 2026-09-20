#!/usr/bin/env python3
"""Capture the frozen 0706 XLSX matrix one serialized child at a time.

This module is deliberately a capture-only tool.  It never builds Cargo
artifacts and never runs a profiler.  The coordinator invokes one phase at a
time so the four ABBA legs can be separated by whatever source checkout
transition is required by the experiment.

The native phase names are:

    baseline-noise1 baseline-noise2 baseline-A1 candidate-B1 candidate-B2
    baseline-A2

Noise phases run the primary matrix only.  The A/B phases run the primary,
guard, and producer matrices.  ``allocator-baseline`` and
``allocator-candidate`` are separate operation-scoped allocator lanes; they
never contribute native timing evidence.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
BIN_ROOT = REPO.parent / "litchi-0706-bin"

NATIVE_PHASES = {
    "baseline-noise1": ("baseline", "noise-1", "primary"),
    "baseline-noise2": ("baseline", "noise-2", "primary"),
    "baseline-A1": ("baseline", "A1", "full"),
    "candidate-B1": ("candidate", "B1", "full"),
    "candidate-B2": ("candidate", "B2", "full"),
    "baseline-A2": ("baseline", "A2", "full"),
}

# A few shell-friendly aliases are accepted, but the canonical names above
# are recorded in every receipt and are the only names consumed by analysis.
PHASE_ALIASES = {
    "baseline-noise-1": "baseline-noise1",
    "baseline-noise-2": "baseline-noise2",
    "baseline-a1": "baseline-A1",
    "candidate-b1": "candidate-B1",
    "candidate-b2": "candidate-B2",
    "baseline-a2": "baseline-A2",
    "alloc-A": "allocator-baseline",
    "alloc-B": "allocator-candidate",
    "allocator-A": "allocator-baseline",
    "allocator-B": "allocator-candidate",
    "allocation-A": "allocator-baseline",
    "allocation-B": "allocator-candidate",
}


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest_json(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def fail(message: str) -> "NoReturn":
    raise RuntimeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def load_json(path: Path) -> Any:
    require(path.is_file(), f"missing JSON input: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def plan() -> dict[str, Any]:
    value = load_json(HERE / "plan.json")
    require(isinstance(value, dict), "plan is not an object")
    return value


def source_census() -> dict[str, str]:
    """Match build.py's source census exactly.

    The census intentionally excludes documentation and this measurement
    packet.  It binds binaries to Rust/Cargo inputs while permitting the
    packet itself to grow as children are captured.
    """

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


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n")


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def source_relation(
    role: str, expected: dict[str, str], current: dict[str, str],
    *, phase: str,
) -> dict[str, Any]:
    changed = sorted(
        name for name in set(expected) | set(current)
        if expected.get(name) != current.get(name)
    )
    allowed_roots = tuple(plan().get("candidate_roots", ()))
    if role == "candidate":
        require(not changed, f"{phase}: candidate source differs from retained source: {changed}")
        mode = "exact"
    else:
        require(
            all(any(name.startswith(root) for root in allowed_roots) for name in changed),
            f"{phase}: baseline source changed outside candidate roots: {changed}",
        )
        mode = "exact" if not changed else "baseline-retained-under-allowed-candidate-delta"
    return {
        "mode": mode,
        "changed_paths": changed,
        "allowed_roots": list(allowed_roots),
        "expected_entry_count": len(expected),
        "current_entry_count": len(current),
    }


def build_record(role: str, lane: str) -> tuple[dict[str, Any], dict[str, str], Path]:
    build_path = HERE / f"build-{role}.json"
    records = load_json(build_path)
    require(isinstance(records, list), f"{build_path} is not a build-record list")
    matches = [item for item in records if isinstance(item, dict) and item.get("label") == lane]
    # build.py's current record does not explicitly name the label in the
    # record; retain compatibility with that exact frozen helper by matching
    # the binary name as a fallback.
    if not matches:
        expected_name = f"{role}-{lane}"
        matches = [
            item for item in records
            if isinstance(item, dict)
            and Path(str(item.get("binary", ""))).name == expected_name
        ]
    require(len(matches) == 1, f"{build_path}: expected one {lane} build record")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{build_path}: {lane} build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    require(binary.is_file() and not binary.is_symlink(), f"missing frozen binary: {binary}")
    binary_digest = record.get("binary_sha256")
    require(isinstance(binary_digest, str) and len(binary_digest) == 64,
            f"{build_path}: {lane} binary digest is missing")
    require(sha(binary) == binary_digest, f"frozen {lane} binary digest changed")
    source_manifest_name = f"source-{role}.json"
    source_manifest_path = HERE / source_manifest_name
    expected = load_json(source_manifest_path)
    require(isinstance(expected, dict), f"{source_manifest_name} is not a source map")
    retained_manifest_sha = sha(source_manifest_path)
    require(
        record.get("source_manifest_sha256") == retained_manifest_sha,
        f"{build_path}: {lane} build is not bound to {source_manifest_name}",
    )
    return record, expected, binary


def constraints_check() -> None:
    constraints = load_json(HERE / "constraints.json")
    require(isinstance(constraints, dict), "constraints is not an object")
    for name, digest in constraints.items():
        path = REPO / name
        require(path.is_file(), f"constraint input is absent: {name}")
        require(sha(path) == digest, f"constraint changed during capture: {name}")


def normalized_phase(raw: str) -> str:
    phase = PHASE_ALIASES.get(raw, raw)
    if phase in NATIVE_PHASES or phase in {"allocator-baseline", "allocator-candidate"}:
        return phase
    fail(
        "unknown phase; use baseline-noise1, baseline-noise2, baseline-A1, "
        "candidate-B1, candidate-B2, baseline-A2, allocator-baseline, or "
        "allocator-candidate"
    )


def slug(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_" else "_" for char in value)


def primary_jobs(p: dict[str, Any], phase: str) -> Iterable[dict[str, Any]]:
    primary = p["primary"]
    shapes = list(primary["shapes"])
    repeats = range(1, int(primary["repeats"]) + 1)
    for repeat in repeats:
        ordered = shapes if repeat == 1 else list(reversed(shapes))
        for shape in ordered:
            yield {
                "kind": "primary",
                "repeat": repeat,
                "shape": shape,
                "case": primary["case"],
                "samples": primary["samples"],
                "warmup": primary["warmup"],
                "name": f"native-{slug(phase)}-r{repeat}-{slug(shape)}",
            }


def expanded_guards(p: dict[str, Any]) -> list[tuple[str, str]]:
    values: list[tuple[str, str]] = []
    for guard in p["guards"]:
        case = guard["case"]
        for shape in guard["shapes"]:
            values.append((case, shape))
    require(len(values) == 7, f"plan guard expansion has {len(values)} entries; expected 7")
    require(len(set(values)) == len(values), "plan guard expansion contains duplicate entries")
    return values


def full_jobs(p: dict[str, Any], phase: str) -> Iterable[dict[str, Any]]:
    yield from primary_jobs(p, phase)
    for case, shape in expanded_guards(p):
        yield {
            "kind": "guard",
            "repeat": None,
            "shape": shape,
            "case": case,
            "samples": p["guard_samples"],
            "warmup": p["guard_warmup"],
            "name": f"guard-{slug(phase)}-{slug(case)}-{slug(shape)}",
        }
    producer = p.get("producer")
    if producer:
        for repeat in range(1, int(p["primary"]["repeats"]) + 1):
            yield {
                "kind": "producer",
                "repeat": repeat,
                "shape": "medium",
                "case": producer,
                "samples": p["primary"]["samples"],
                "warmup": p["primary"]["warmup"],
                "name": f"producer-{slug(phase)}-r{repeat}",
            }


def allocation_jobs(p: dict[str, Any], phase: str) -> Iterable[dict[str, Any]]:
    config = p["allocation"]
    shapes = list(config["shapes"])
    for repeat in range(1, int(config["repeats"]) + 1):
        ordered = shapes if repeat == 1 else list(reversed(shapes))
        for shape in ordered:
            yield {
                "kind": "allocation",
                "repeat": repeat,
                "shape": shape,
                "case": p["primary"]["case"],
                "samples": config["samples"],
                "warmup": config["warmup"],
                "name": f"alloc-{slug(phase)}-r{repeat}-{slug(shape)}",
            }


def run_child(
    *, phase: str, role: str, lane: str, job: dict[str, Any],
    p: dict[str, Any], expected: dict[str, str], build: dict[str, Any],
    binary: Path, source_manifest_name: str,
) -> None:
    name = job["name"]
    receipt_path = HERE / f"{name}.receipt.json"
    require(not receipt_path.exists(), f"refusing to replace existing receipt: {receipt_path}")
    output_path = HERE / f"{name}.json"
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    before_path = HERE / f"{name}.source-before.json"
    after_path = HERE / f"{name}.source-after.json"
    for path in (output_path, stdout_path, stderr_path, before_path, after_path):
        require(not path.exists(), f"refusing to replace existing artifact: {path}")

    constraints_check()
    current_before = source_census()
    relation_before = source_relation(role, expected, current_before, phase=phase)
    write_json(before_path, current_before)

    command = [
        "taskset", "-c", str(p["cpu"]), str(binary),
        "--warmup", str(job["warmup"]),
        "--samples", str(job["samples"]),
        "--case", job["case"],
        "--xlsx-cell-crud-shape", job["shape"],
        "--json", str(output_path),
    ]
    if job["kind"] == "producer":
        command += ["--producer-evidence", str(HERE / f"{name}.producer.json")]
    started = utc_now()
    tick = time.monotonic()
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        result = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr)
    elapsed = time.monotonic() - tick

    current_after = source_census()
    relation_after = source_relation(role, expected, current_after, phase=phase)
    write_json(after_path, current_after)
    require(current_before == current_after, f"{name}: source census changed during child")
    require(sha(binary) == build["binary_sha256"], f"{name}: retained binary changed")

    retained_source_sha = sha(HERE / source_manifest_name)
    artifacts: dict[str, str] = {}
    for path in sorted(HERE.glob(name + "*")):
        if path.is_file() and path != receipt_path:
            artifacts[path.name] = sha(path)
    producer_path = HERE / f"{name}.producer.json"
    if producer_path.is_file():
        artifacts[producer_path.name] = sha(producer_path)

    receipt = {
        "schema_version": 2,
        "name": name,
        "phase": phase,
        "role": role,
        "lane": lane,
        "kind": job["kind"],
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
        "retained_binary_source": {
            "manifest": source_manifest_name,
            "manifest_sha256": retained_source_sha,
            "source_census_sha256": digest_json(expected),
            "source_entry_count": len(expected),
        },
        "current_checkout_source": {
            "before_artifact": before_path.name,
            "after_artifact": after_path.name,
            "before_sha256": digest_json(current_before),
            "after_sha256": digest_json(current_after),
            "before_file_sha256": sha(before_path),
            "after_file_sha256": sha(after_path),
            "before_entry_count": len(current_before),
            "after_entry_count": len(current_after),
            "relation_before": relation_before,
            "relation_after": relation_after,
            "unchanged_during_child": current_before == current_after,
        },
        "build_record_sha256": sha(HERE / f"build-{role}.json"),
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "plan_sha256": sha(HERE / "plan.json"),
        "script_sha256": sha(Path(__file__)),
        "constraints_sha256": sha(HERE / "constraints.json"),
        "environment": {
            key: os.environ.get(key)
            for key in ("RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF", "GLIBC_TUNABLES")
        },
        "artifacts": artifacts,
    }
    write_json(receipt_path, receipt)
    require(result.returncode == 0, f"{name}: child failed with exit code {result.returncode}")
    print(f"{name} passed", flush=True)


def run_phase(raw_phase: str) -> None:
    phase = normalized_phase(raw_phase)
    p = plan()
    require(p.get("revision"), "plan revision is missing")
    if phase in NATIVE_PHASES:
        role, label, matrix = NATIVE_PHASES[phase]
        lane = "native"
        build, expected, binary = build_record(role, lane)
        source_manifest_name = f"source-{role}.json"
        jobs = primary_jobs(p, phase) if matrix == "primary" else full_jobs(p, phase)
    else:
        role = "baseline" if phase == "allocator-baseline" else "candidate"
        label = "allocator-A" if role == "baseline" else "allocator-B"
        lane = "alloc"
        build, expected, binary = build_record(role, lane)
        source_manifest_name = f"source-{role}.json"
        jobs = allocation_jobs(p, phase)

    # Materialize the generator before running so a malformed plan cannot
    # leave a partially captured matrix with an apparently valid prefix.
    jobs = list(jobs)
    require(jobs, f"{phase}: no jobs selected")
    for job in jobs:
        run_child(
            phase=phase, role=role, lane=lane, job=job, p=p, expected=expected,
            build=build, binary=binary, source_manifest_name=source_manifest_name,
        )
    print(f"phase {phase} complete ({len(jobs)} children)", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("phase", help="one serialized native or allocator phase")
    args = parser.parse_args()
    try:
        run_phase(args.phase)
    except (AssertionError, OSError, RuntimeError, ValueError) as error:
        print(f"capture failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
