#!/usr/bin/env python3
"""Capture the 0709 DOCX ordinary-save baseline packet.

The packet is deliberately split into one benchmark process per
``(corpus, phase, repeat, lane)``.  A process therefore receives at most one
``--ooxml-file`` input, which keeps the DOCX classifier and the fixture
identity independent of every other corpus.  This script never builds Cargo
artifacts and never invokes a profiler; the coordinator supplies the two
already-frozen baseline binaries.

The native lane has three same-build repeats of the four ordinary-save
phases, with the second repeat in reverse phase order.  The allocator lane
has two repeats and uses the same order rule.  Allocator elapsed values are
retained for instrumentation evidence only and are never a speedup claim.
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
from typing import Any, NoReturn


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

DEFAULT_PHASE_ORDER = ["lifecycle", "edit", "atomic_publish", "counting_publish"]
PHASE_CASE_SUFFIX = {
    "lifecycle": "lifecycle",
    "edit": "edit",
    "atomic_publish": "atomic_publish",
    "counting_publish": "counting_publish",
}
PHASE_ALIASES = {
    "atomic": "atomic_publish",
    "counting": "counting_publish",
    "counting_sink": "counting_publish",
}


def fail(message: str) -> NoReturn:
    raise RuntimeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest_json(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON input: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n")


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def slug(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_" else "_" for char in value)


def source_census() -> dict[str, str]:
    """Match the 0709 build.py Rust/Cargo source census exactly."""

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


def load_plan() -> dict[str, Any]:
    value = read_json(HERE / "plan.json")
    require(isinstance(value, dict), "plan is not an object")
    require(isinstance(value.get("revision"), str) and value["revision"],
            "plan revision is missing")
    require(isinstance(value.get("cpu"), int) and value["cpu"] >= 0,
            "plan cpu must be a non-negative integer")
    for lane in ("native", "allocator"):
        config = value.get(lane)
        require(isinstance(config, dict), f"plan {lane} configuration is missing")
        for key in ("repeats", "samples", "warmup"):
            require(isinstance(config.get(key), int) and config[key] >= 0,
                    f"plan {lane}.{key} must be a non-negative integer")
        require(config["repeats"] > 0 and config["samples"] > 0,
                f"plan {lane} repeats and samples must be positive")
    require(value["native"] == {"repeats": 3, "samples": 100, "warmup": 10},
            "native sample plan changed")
    require(value["allocator"] == {"repeats": 2, "samples": 3, "warmup": 0},
            "allocator sample plan changed")
    order_description = value.get("order", "")
    require(isinstance(order_description, str)
            and "repeat 1" in order_description
            and "repeat 2" in order_description
            and "reverse" in order_description,
            "repeat order description is missing")
    require(value.get("review_threshold_percent") == 5,
            "repeat review threshold changed")
    phase_order = value.get("phase_order", DEFAULT_PHASE_ORDER)
    require(isinstance(phase_order, list) and phase_order == DEFAULT_PHASE_ORDER,
            "ordinary-save phase order changed")
    corpora = value.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 3,
            "plan must contain exactly three DOCX corpora")
    ids: set[str] = set()
    for corpus in corpora:
        require(isinstance(corpus, dict), "corpus entry is not an object")
        for key in ("id", "label", "origin", "expected_edit_admitted"):
            require(key in corpus, f"corpus entry lacks {key}")
        identity = corpus["id"]
        require(isinstance(identity, str) and identity and identity not in ids,
                "corpus ids must be unique non-empty strings")
        ids.add(identity)
        require(corpus["origin"] in {"generated-harness-corpus", "caller-named-real-file"},
                f"unsupported corpus origin: {corpus['origin']}")
        require(isinstance(corpus["expected_edit_admitted"], bool),
                f"{identity}: expected_edit_admitted must be boolean")
        path = corpus.get("path")
        digest = corpus.get("sha256")
        if corpus["origin"] == "generated-harness-corpus":
            require(path is None and digest is None,
                    f"{identity}: generated corpus must not bind a fixture")
        else:
            require(isinstance(path, str) and path and isinstance(digest, str)
                    and len(digest) == 64 and set(digest) <= set("0123456789abcdef"),
                    f"{identity}: real corpus fixture binding is malformed")
    return value


def constraints_check() -> None:
    constraints = read_json(HERE / "constraints.json")
    require(isinstance(constraints, dict), "constraints is not an object")
    for name, digest in constraints.items():
        path = REPO / name
        require(path.is_file(), f"constraint input is absent: {name}")
        require(sha(path) == digest, f"constraint changed during capture: {name}")


def source_relation(expected: dict[str, str], current: dict[str, str], label: str) -> dict[str, Any]:
    changed = sorted(
        name for name in set(expected) | set(current)
        if expected.get(name) != current.get(name)
    )
    require(not changed, f"{label}: current source differs from baseline source: {changed}")
    return {
        "mode": "exact",
        "changed_paths": [],
        "expected_entry_count": len(expected),
        "current_entry_count": len(current),
    }


def build_record(lane: str) -> tuple[dict[str, Any], dict[str, str], Path]:
    records_path = HERE / "build-baseline.json"
    records = read_json(records_path)
    require(isinstance(records, list), "build-baseline.json is not a record list")
    binary_name = f"baseline-{lane}"
    matches = [
        item for item in records
        if isinstance(item, dict) and Path(str(item.get("binary", ""))).name == binary_name
    ]
    require(len(matches) == 1, f"build-baseline.json has no unique {binary_name} record")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{binary_name} build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    require(binary.is_file() and not binary.is_symlink(), f"missing frozen binary: {binary}")
    digest = record.get("binary_sha256")
    require(isinstance(digest, str) and len(digest) == 64, f"{binary_name} binary digest missing")
    require(sha(binary) == digest, f"{binary_name} binary digest changed")
    byte_count = record.get("binary_bytes")
    require(isinstance(byte_count, int) and byte_count > 0 and binary.stat().st_size == byte_count,
            f"{binary_name} binary byte count changed")
    manifest_path = HERE / "source-baseline.json"
    expected = read_json(manifest_path)
    require(isinstance(expected, dict) and expected, "source-baseline.json is not a source map")
    require(record.get("source_manifest_sha256") == sha(manifest_path),
            f"{binary_name} is not bound to source-baseline.json")
    return record, expected, binary


def corpus_entries(p: dict[str, Any]) -> list[dict[str, Any]]:
    # Preserve plan order in receipts and analysis.  The analyzer independently
    # checks the three required semantic roles, so labels remain human-facing.
    return list(p["corpora"])


def phase_order(p: dict[str, Any], repeat: int) -> list[str]:
    order = list(p.get("phase_order", DEFAULT_PHASE_ORDER))
    return order if repeat % 2 == 1 else list(reversed(order))


def phase_case(corpus: dict[str, Any], phase: str) -> str:
    prefix = "docx_ordinary_save_" if corpus["origin"] == "generated-harness-corpus" else "docx_real_file_ordinary_save_"
    return prefix + PHASE_CASE_SUFFIX[phase]


def corpus_fixture(corpus: dict[str, Any]) -> tuple[Path | None, dict[str, Any] | None]:
    if corpus["origin"] == "generated-harness-corpus":
        return None, None
    raw = str(corpus["path"])
    path = (REPO / raw).resolve() if not Path(raw).is_absolute() else Path(raw).resolve()
    require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
    actual_sha = sha(path)
    require(actual_sha == corpus["sha256"],
            f"{corpus['id']}: fixture digest changed: {actual_sha}")
    if corpus.get("bytes") is not None:
        require(path.stat().st_size == corpus["bytes"],
                f"{corpus['id']}: fixture byte count changed")
    return path, {"path": raw, "resolved_path": str(path), "bytes": path.stat().st_size,
                  "sha256": actual_sha}


def child_name(lane: str, repeat: int, corpus: dict[str, Any], phase: str) -> str:
    return f"{lane}-r{repeat}-{slug(corpus['id'])}-{phase}"


def expected_artifacts(name: str) -> list[Path]:
    return [
        HERE / f"{name}.json",
        HERE / f"{name}.stdout",
        HERE / f"{name}.stderr",
        HERE / f"{name}.source-before.json",
        HERE / f"{name}.source-after.json",
        HERE / f"{name}.receipt.json",
    ]


def run_child(
    *, p: dict[str, Any], lane: str, repeat: int, corpus: dict[str, Any], phase: str,
    expected_source: dict[str, str], build: dict[str, Any], binary: Path,
    order_index: int,
) -> None:
    name = child_name(lane, repeat, corpus, phase)
    paths = expected_artifacts(name)
    for path in paths:
        require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")

    constraints_check()
    current_before = source_census()
    relation_before = source_relation(expected_source, current_before, f"{name} before")
    before_path = HERE / f"{name}.source-before.json"
    write_json(before_path, current_before)

    fixture_path, fixture_before = corpus_fixture(corpus)
    command = [
        "taskset", "-c", str(p["cpu"]), str(binary),
        "--warmup", str(p[lane]["warmup"]),
        "--samples", str(p[lane]["samples"]),
        "--case", phase_case(corpus, phase),
        "--json", str(HERE / f"{name}.json"),
    ]
    filesystem_root = p.get("filesystem_root")
    if filesystem_root:
        command += ["--filesystem-root", str(filesystem_root)]
    if fixture_path is not None:
        # Use the plan spelling in the command.  This preserves the exact
        # source evidence path inside the benchmark's real_file provenance.
        command += ["--ooxml-file", str(corpus["path"])]

    started = utc_now()
    tick = time.monotonic()
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        result = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr)
    seconds = time.monotonic() - tick

    current_after = source_census()
    relation_after = source_relation(expected_source, current_after, f"{name} after")
    after_path = HERE / f"{name}.source-after.json"
    write_json(after_path, current_after)
    require(current_before == current_after, f"{name}: source changed during child")
    require(sha(binary) == build["binary_sha256"], f"{name}: binary changed during child")

    fixture_after = None
    if fixture_path is not None:
        fixture_after = corpus_fixture(corpus)[1]
        require(fixture_before == fixture_after, f"{name}: fixture changed during child")

    output_path = HERE / f"{name}.json"
    artifacts: dict[str, str] = {}
    for path in (output_path, stdout_path, stderr_path, before_path, after_path):
        require(path.is_file() and not path.is_symlink(), f"{name}: child did not create {path.name}")
        artifacts[path.name] = sha(path)

    receipt = {
        "schema_version": 1,
        "name": name,
        "lane": lane,
        "lane_order_index": order_index,
        "repeat": repeat,
        "corpus_id": corpus["id"],
        "corpus_label": corpus["label"],
        "corpus_origin": corpus["origin"],
        "phase": phase,
        "case": phase_case(corpus, phase),
        "samples": p[lane]["samples"],
        "warmup": p[lane]["warmup"],
        "command": command,
        "start_utc": started,
        "end_utc": utc_now(),
        "seconds": seconds,
        "exit_code": result.returncode,
        "cpu": p["cpu"],
        "binary_path": str(binary),
        "binary_sha256": sha(binary),
        "binary_bytes": binary.stat().st_size,
        "build_record_sha256": sha(HERE / "build-baseline.json"),
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "retained_binary_source": {
            "manifest": "source-baseline.json",
            "manifest_sha256": sha(HERE / "source-baseline.json"),
            "source_census_sha256": digest_json(expected_source),
            "source_entry_count": len(expected_source),
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
        "fixture": {
            "plan_path": corpus.get("path"),
            "plan_sha256": corpus.get("sha256"),
            "before": fixture_before,
            "after": fixture_after,
        },
        "plan_sha256": sha(HERE / "plan.json"),
        "script_sha256": sha(Path(__file__)),
        "constraints_sha256": sha(HERE / "constraints.json"),
        "environment": {
            key: os.environ.get(key)
            for key in ("RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF", "GLIBC_TUNABLES")
        },
        "artifacts": artifacts,
    }
    receipt_path = HERE / f"{name}.receipt.json"
    write_json(receipt_path, receipt)
    require(result.returncode == 0, f"{name}: benchmark child failed with exit code {result.returncode}")
    print(f"{name} passed", flush=True)


def run_lane(raw_lane: str) -> None:
    lane = {"alloc": "allocator", "allocation": "allocator"}.get(raw_lane, raw_lane)
    require(lane in {"native", "allocator"}, "lane must be native or allocator")
    p = load_plan()
    build, expected_source, binary = build_record("native" if lane == "native" else "alloc")
    jobs: list[tuple[int, dict[str, Any], str]] = []
    for repeat in range(1, p[lane]["repeats"] + 1):
        for phase in phase_order(p, repeat):
            ordered_corpora = corpus_entries(p)
            if repeat % 2 == 0:
                ordered_corpora = list(reversed(ordered_corpora))
            for corpus in ordered_corpora:
                jobs.append((repeat, corpus, phase))
    require(len(jobs) == (36 if lane == "native" else 24),
            f"{lane}: expected fixed child count, got {len(jobs)}")
    for order_index, (repeat, corpus, phase) in enumerate(jobs):
        run_child(
            p=p, lane=lane, repeat=repeat, corpus=corpus, phase=phase,
            expected_source=expected_source, build=build, binary=binary,
            order_index=order_index,
        )
    print(f"lane {lane} complete ({len(jobs)} children)", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("lane", help="native or allocator")
    args = parser.parse_args()
    try:
        run_lane(args.lane)
    except (AssertionError, OSError, RuntimeError, ValueError) as error:
        print(f"capture failed: {error}", file=os.sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
