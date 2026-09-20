#!/usr/bin/env python3
"""Capture one frozen stage of the 0711 paired DOCX ordinary-save pilot.

The coordinator changes the checkout and builds the named baseline or
candidate stage outside this script.  A stage is then captured as four
isolated children (two corpora by two phases) for one lane.  The four stage
labels are deliberately explicit: baseline-A1, candidate-B1, candidate-B2,
and baseline-A2.  B2 and A2 reverse both corpus and phase order to reduce
ordering bias.  This script never builds Cargo artifacts, edits production,
restores a checkout, or replaces an existing evidence file.
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
from typing import Any, NoReturn

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]

PHASES = ("edit", "lifecycle")
LANE_ALIASES = {"alloc": "allocator", "allocation": "allocator"}
STAGE_BY_LABEL = {
    "baseline-A1": {"source": "baseline", "build": "baseline", "pair": "pair-1", "order": "forward"},
    "candidate-B1": {"source": "candidate", "build": "candidate", "pair": "pair-1", "order": "forward"},
    "candidate-B2": {"source": "candidate", "build": "candidate", "pair": "pair-2", "order": "reverse"},
    "baseline-A2": {"source": "baseline", "build": "baseline", "pair": "pair-2", "order": "reverse"},
}
PHASE_CASE_SUFFIX = {"edit": "edit", "lifecycle": "lifecycle"}


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
    """Use the exact census used by custody.py and build.py."""

    # Importing the packet's single custody implementation prevents a later
    # change in source selection from making captures and build manifests
    # silently incomparable.
    if str(HERE) not in sys.path:
        sys.path.insert(0, str(HERE))
    from custody import census

    return census()


def load_plan() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema_version") == 1, "plan schema changed")
    require(plan.get("revision") == "9862aeb599ca623d75e5a10e4f297b472442e7d3",
            "plan revision changed")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("filesystem_root") == "/home/zhuhe/code/litchi-0711-fs",
            "filesystem root changed")
    require(plan.get("phase_order") == list(PHASES), "phase order changed")
    require(plan.get("native") == {"samples": 100, "warmup": 10},
            "native sample plan changed")
    require(plan.get("allocator") == {"samples": 3, "warmup": 0},
            "allocator sample plan changed")
    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 2,
            "pilot must contain exactly two DOCX corpora")
    ids: set[str] = set()
    for corpus in corpora:
        require(isinstance(corpus, dict), "corpus entry is malformed")
        identity = corpus.get("id")
        require(isinstance(identity, str) and identity and identity not in ids,
                "corpus ids must be unique")
        ids.add(identity)
        origin = corpus.get("origin")
        require(origin in {"generated-harness-corpus", "caller-named-real-file"},
                f"{identity}: unsupported corpus origin")
        require(isinstance(corpus.get("expected_edit_admitted"), bool),
                f"{identity}: edit admission missing")
        if origin == "generated-harness-corpus":
            require(corpus.get("path") is None and corpus.get("sha256") is None,
                    f"{identity}: generated corpus has fixture binding")
        else:
            digest = corpus.get("sha256")
            require(isinstance(corpus.get("path"), str) and corpus["path"],
                    f"{identity}: real fixture path missing")
            require(isinstance(digest, str) and len(digest) == 64
                    and set(digest) <= set("0123456789abcdef"),
                    f"{identity}: real fixture digest malformed")
    stages = plan.get("stages")
    require(isinstance(stages, list) and len(stages) == 4, "stage count changed")
    require([item.get("label") for item in stages] == list(STAGE_BY_LABEL),
            "stage order changed")
    for item in stages:
        require(item == {"label": item["label"], **STAGE_BY_LABEL[item["label"]]},
                f"stage metadata changed: {item.get('label')}")
    thresholds = plan.get("thresholds")
    require(thresholds == {
        "edit_improvement_percent": 3,
        "lifecycle_regression_percent": 3,
        "allocation_regression_percent": 3,
        "tail_flag_percent": 5,
        "repeat_drift_flag_percent": 5,
    }, "thresholds changed")
    return plan


def constraints_check() -> None:
    constraints = read_json(HERE / "constraints.json")
    require(isinstance(constraints, dict), "constraints is not an object")
    for name, digest in constraints.items():
        path = REPO / name
        require(path.is_file() and sha(path) == digest,
                f"constraint changed during capture: {name}")


def fixture_info(corpus: dict[str, Any]) -> tuple[Path | None, dict[str, Any] | None]:
    if corpus["origin"] == "generated-harness-corpus":
        return None, None
    raw = str(corpus["path"])
    path = (REPO / raw).resolve() if not Path(raw).is_absolute() else Path(raw).resolve()
    require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
    actual = sha(path)
    require(actual == corpus["sha256"], f"fixture digest changed: {corpus['id']}")
    require(path.stat().st_size == corpus["bytes"], f"fixture size changed: {corpus['id']}")
    return path, {"path": raw, "resolved_path": str(path), "bytes": path.stat().st_size,
                  "sha256": actual}


def stage_record(stage: str, lane: str) -> tuple[dict[str, Any], dict[str, str], Path]:
    metadata = STAGE_BY_LABEL[stage]
    build_path = HERE / f"build-{metadata['build']}.json"
    records = read_json(build_path)
    require(isinstance(records, list), f"{build_path.name} is not a record list")
    binary_label = "native" if lane == "native" else "alloc"
    binary_name = f"{metadata['build']}-{binary_label}"
    matches = [item for item in records if isinstance(item, dict)
               and Path(str(item.get("binary", ""))).name == binary_name]
    require(len(matches) == 1, f"no unique build record for {binary_name}")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{binary_name} build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    require(binary.is_file() and not binary.is_symlink(), f"missing binary: {binary}")
    digest = record.get("binary_sha256")
    require(isinstance(digest, str) and len(digest) == 64 and sha(binary) == digest,
            f"{binary_name} digest changed")
    size = record.get("binary_bytes")
    require(isinstance(size, int) and size > 0 and binary.stat().st_size == size,
            f"{binary_name} size changed")
    source_path = HERE / f"source-{metadata['source']}.json"
    expected_source = read_json(source_path)
    require(isinstance(expected_source, dict) and expected_source,
            f"{source_path.name} is not a source map")
    require(record.get("source_manifest_sha256") == sha(source_path),
            f"{binary_name} source manifest binding changed")
    return record, expected_source, binary


def source_relation(expected: dict[str, str], current: dict[str, str], label: str) -> dict[str, Any]:
    changed = sorted(name for name in set(expected) | set(current)
                     if expected.get(name) != current.get(name))
    require(not changed, f"{label}: source census differs: {changed}")
    return {"mode": "exact", "changed_paths": [],
            "expected_entry_count": len(expected), "current_entry_count": len(current)}


def phase_case(corpus: dict[str, Any], phase: str) -> str:
    prefix = ("docx_ordinary_save_" if corpus["origin"] == "generated-harness-corpus"
              else "docx_real_file_ordinary_save_")
    return prefix + PHASE_CASE_SUFFIX[phase]


def child_name(stage: str, lane: str, corpus: dict[str, Any], phase: str) -> str:
    return f"{slug(stage)}-{lane}-{slug(corpus['id'])}-{phase}"


def artifacts_for(name: str) -> list[Path]:
    return [HERE / f"{name}{suffix}" for suffix in (
        ".json", ".stdout", ".stderr", ".source-before.json",
        ".source-after.json", ".receipt.json")]


def ordered_jobs(plan: dict[str, Any], stage: str) -> list[tuple[dict[str, Any], str]]:
    metadata = STAGE_BY_LABEL[stage]
    phases = list(plan["phase_order"])
    corpora = list(plan["corpora"])
    if metadata["order"] == "reverse":
        phases.reverse()
        corpora.reverse()
    return [(corpus, phase) for phase in phases for corpus in corpora]


def run_child(*, plan: dict[str, Any], stage: str, lane: str,
              corpus: dict[str, Any], phase: str, expected_source: dict[str, str],
              build: dict[str, Any], binary: Path, order_index: int) -> None:
    name = child_name(stage, lane, corpus, phase)
    paths = artifacts_for(name)
    for path in paths:
        require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    constraints_check()
    before = source_census()
    relation_before = source_relation(expected_source, before, f"{name} before")
    before_path = HERE / f"{name}.source-before.json"
    write_json(before_path, before)
    fixture_path, fixture_before = fixture_info(corpus)
    command = [
        "taskset", "-c", str(plan["cpu"]), str(binary),
        "--warmup", str(plan[lane]["warmup"]),
        "--samples", str(plan[lane]["samples"]),
        "--case", phase_case(corpus, phase),
        "--json", str(HERE / f"{name}.json"),
        "--filesystem-root", str(plan["filesystem_root"]),
    ]
    if fixture_path is not None:
        command += ["--ooxml-file", str(corpus["path"])]
    started = utc_now()
    tick = time.monotonic()
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
        result = subprocess.run(command, cwd=REPO, stdout=stdout, stderr=stderr)
    seconds = time.monotonic() - tick
    after = source_census()
    relation_after = source_relation(expected_source, after, f"{name} after")
    after_path = HERE / f"{name}.source-after.json"
    write_json(after_path, after)
    require(before == after, f"{name}: source changed during child")
    require(sha(binary) == build["binary_sha256"], f"{name}: binary changed during child")
    fixture_after = None if fixture_path is None else fixture_info(corpus)[1]
    require(fixture_before == fixture_after, f"{name}: fixture changed during child")
    output_path = HERE / f"{name}.json"
    artifacts: dict[str, str] = {}
    for path in (output_path, stdout_path, stderr_path, before_path, after_path):
        require(path.is_file() and not path.is_symlink(), f"{name}: missing {path.name}")
        artifacts[path.name] = sha(path)
    metadata = STAGE_BY_LABEL[stage]
    receipt = {
        "schema_version": 1,
        "packet": "change-0711-docx-ordinary-save-paired-pilot",
        "name": name,
        "stage": stage,
        "source_label": metadata["source"],
        "build_label": metadata["build"],
        "pair": metadata["pair"],
        "lane": lane,
        "stage_order": metadata["order"],
        "stage_order_index": order_index,
        "corpus_id": corpus["id"],
        "corpus_label": corpus["label"],
        "corpus_origin": corpus["origin"],
        "phase": phase,
        "case": phase_case(corpus, phase),
        "samples": plan[lane]["samples"],
        "warmup": plan[lane]["warmup"],
        "command": command,
        "start_utc": started,
        "end_utc": utc_now(),
        "seconds": seconds,
        "exit_code": result.returncode,
        "cpu": plan["cpu"],
        "binary_path": str(binary),
        "binary_sha256": sha(binary),
        "binary_bytes": binary.stat().st_size,
        "build_record": build_path_name(metadata["build"]),
        "build_record_sha256": sha(HERE / build_path_name(metadata["build"])),
        "build_source_manifest": f"source-{metadata['source']}.json",
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "retained_binary_source": {
            "manifest": f"source-{metadata['source']}.json",
            "manifest_sha256": sha(HERE / f"source-{metadata['source']}.json"),
            "source_census_sha256": digest_json(expected_source),
            "source_entry_count": len(expected_source),
        },
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
        "fixture": {"plan_path": corpus.get("path"), "plan_sha256": corpus.get("sha256"),
                    "before": fixture_before, "after": fixture_after},
        "plan_sha256": sha(HERE / "plan.json"),
        "script_sha256": sha(Path(__file__).resolve()),
        "constraints_sha256": sha(HERE / "constraints.json"),
        "environment": {key: os.environ.get(key)
                         for key in ("RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF", "GLIBC_TUNABLES")},
        "artifacts": artifacts,
    }
    write_json(HERE / f"{name}.receipt.json", receipt)
    require(result.returncode == 0, f"{name}: benchmark failed with exit code {result.returncode}")
    print(f"{name} passed", flush=True)


def build_path_name(label: str) -> str:
    return f"build-{label}.json"


def run_lane(stage: str, raw_lane: str) -> None:
    plan = load_plan()
    require(stage in STAGE_BY_LABEL, f"unknown stage {stage!r}")
    lane = LANE_ALIASES.get(raw_lane, raw_lane)
    require(lane in {"native", "allocator"}, "lane must be native or alloc")
    build, expected_source, binary = stage_record(stage, lane)
    jobs = ordered_jobs(plan, stage)
    require(len(jobs) == 4, f"fixed stage child count changed: {len(jobs)}")
    for index, (corpus, phase) in enumerate(jobs):
        run_child(plan=plan, stage=stage, lane=lane, corpus=corpus, phase=phase,
                  expected_source=expected_source, build=build, binary=binary,
                  order_index=index)
    print(f"stage {stage} lane {lane} complete ({len(jobs)} children)", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", help="baseline-A1, candidate-B1, candidate-B2, or baseline-A2")
    parser.add_argument("lane", help="native or alloc")
    args = parser.parse_args()
    # Accept the accidental lane-first spelling while retaining one frozen
    # canonical argv in every receipt.
    stage, lane = args.stage, args.lane
    if stage in {"native", "alloc", "allocator", "allocation"} and lane in STAGE_BY_LABEL:
        stage, lane = lane, stage
    try:
        run_lane(stage, lane)
    except (AssertionError, OSError, RuntimeError, ValueError, KeyError) as error:
        print(f"capture failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
