#!/usr/bin/env python3
"""Capture one source-bound stage of the 0722 DOCX fusion pilot.

The coordinator builds the baseline and candidate binaries before invoking this
driver.  The driver never edits or restores the checkout and never builds
Cargo artifacts.  Every benchmark invocation is a separate child process.  A
child records the frozen binary source and the live candidate checkout source
by digest; the full source maps are retained once as ``source-baseline.json``
and ``source-candidate.json``.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from typing import Any, NoReturn

sys.dont_write_bytecode = True


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PHASES = ("edit", "lifecycle")
LANES = ("native", "allocator")
LANE_ALIASES = {"alloc": "allocator", "allocation": "allocator"}
STAGES = (
    "baseline-A1", "candidate-B1", "candidate-B2", "baseline-A2",
    "baseline-A3", "candidate-B3", "candidate-B4", "baseline-A4",
)
ALLOCATOR_STAGES = STAGES[:4]
PHASE_CASE_SUFFIX = {"edit": "edit", "lifecycle": "lifecycle"}
ENVIRONMENT_KEYS = (
    "LC_ALL", "LANG", "TZ", "RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
    "GLIBC_TUNABLES", "PERL_HASH_SEED", "PERL_PERTURB_KEYS",
)


_CUSTODY_SPEC = importlib.util.spec_from_file_location(
    "change0722_custody", HERE / "custody.py"
)
if _CUSTODY_SPEC is None or _CUSTODY_SPEC.loader is None:
    raise RuntimeError(f"cannot load packet custody module: {HERE / 'custody.py'}")
_CUSTODY = importlib.util.module_from_spec(_CUSTODY_SPEC)
_CUSTODY_SPEC.loader.exec_module(_CUSTODY)


def fail(message: str) -> NoReturn:
    raise RuntimeError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest_json(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON input: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON in {path}: {error}")


def write_json(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def slug(value: str) -> str:
    return "".join(char if char.isalnum() or char in "-_" else "_" for char in value)


def source_census() -> dict[str, str]:
    value = _CUSTODY.census()
    require(isinstance(value, dict) and all(
        isinstance(key, str) and isinstance(item, str)
        for key, item in value.items()
    ), "custody census is not a raw path-to-SHA map")
    return value


def load_plan() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema_version") == 1, "plan schema changed")
    require(isinstance(plan.get("packet"), str) and plan["packet"],
            "plan packet is missing")
    require(isinstance(plan.get("revision"), str) and plan["revision"],
            "plan revision is missing")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("filesystem_root") == "/home/zhuhe/code/litchi-0722-fs",
            "filesystem root changed")
    require(plan.get("phase_order") == list(PHASES), "phase order changed")
    require(plan.get("native") == {"samples": 200, "warmup": 100},
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
                f"{identity}: edit admission is missing")
        if origin == "generated-harness-corpus":
            require(corpus.get("path") is None and corpus.get("sha256") is None,
                    f"{identity}: generated corpus has fixture binding")
        else:
            digest = corpus.get("sha256")
            require(isinstance(corpus.get("path"), str) and corpus["path"],
                    f"{identity}: real fixture path is missing")
            require(isinstance(digest, str) and len(digest) == 64
                    and set(digest) <= set("0123456789abcdef"),
                    f"{identity}: real fixture digest is malformed")
    stages = plan.get("stages")
    require(isinstance(stages, list) and len(stages) == len(STAGES),
            "stage count changed")
    require([item.get("label") for item in stages] == list(STAGES),
            "stage order changed")
    for item in stages:
        require(isinstance(item, dict), "stage entry is malformed")
        require(item.get("label") in STAGES, "unknown stage label")
        require(item.get("source") in {"baseline", "candidate"},
                f"{item.get('label')}: source changed")
        require(item.get("build") == item.get("source"),
                f"{item.get('label')}: build/source pairing changed")
        require(item.get("pair") in {"pair-1", "pair-2", "pair-3", "pair-4"},
                f"{item.get('label')}: pair changed")
        require(item.get("cycle") in {1, 2}, f"{item.get('label')}: cycle changed")
        require(item.get("order") in {"forward", "reverse"},
                f"{item.get('label')}: order changed")
    require(plan.get("allocator_stages") == list(ALLOCATOR_STAGES),
            "allocator stage scope changed")
    require(plan.get("thresholds") == {
        "edit_improvement_percent": 3,
        "lifecycle_regression_percent": 3,
        "allocation_regression_percent": 3,
        "allocation_net_live_regression_percent": 0,
        "tail_flag_percent": 5,
        "repeat_drift_flag_percent": 5,
    }, "thresholds changed")
    allowlist = plan.get("source_delta_allowlist")
    require(isinstance(allowlist, list) and allowlist and
            all(isinstance(item, str) and item for item in allowlist),
            "source delta allowlist is malformed")
    require(plan.get("environment") == {
        "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "RUSTFLAGS": None, "LD_PRELOAD": None, "MALLOC_CONF": None,
        "GLIBC_TUNABLES": None, "PERL_HASH_SEED": "0",
        "PERL_PERTURB_KEYS": "0",
    }, "child environment changed")
    return plan


def constraints_check() -> None:
    constraints = read_json(HERE / "constraints.json")
    require(isinstance(constraints, dict), "constraints is not an object")
    for name, digest in constraints.items():
        require(isinstance(name, str) and isinstance(digest, str),
                "constraint entry is malformed")
        path = REPO / name
        require(path.is_file() and sha(path) == digest,
                f"constraint changed: {name}")


def source_pair() -> tuple[dict[str, str], dict[str, str], list[str]]:
    baseline = read_json(HERE / "source-baseline.json")
    candidate = read_json(HERE / "source-candidate.json")
    require(isinstance(baseline, dict) and baseline,
            "source-baseline.json is not a raw source map")
    require(isinstance(candidate, dict) and candidate,
            "source-candidate.json is not a raw source map")
    plan = load_plan()
    changed = sorted(name for name in set(baseline) | set(candidate)
                     if baseline.get(name) != candidate.get(name))
    require(changed, "candidate source is identical to baseline source")
    require(set(changed) <= set(plan["source_delta_allowlist"]),
            f"candidate source delta exceeds allowlist: {changed}")
    return baseline, candidate, changed


def source_relation(expected: dict[str, str], current: dict[str, str], label: str) -> dict[str, Any]:
    changed = sorted(name for name in set(expected) | set(current)
                     if expected.get(name) != current.get(name))
    require(not changed, f"{label}: live source differs: {changed}")
    return {
        "mode": "exact",
        "changed_paths": [],
        "expected_entry_count": len(expected),
        "current_entry_count": len(current),
        "expected_sha256": digest_json(expected),
        "current_sha256": digest_json(current),
    }


def cleanup_witnesses() -> list[dict[str, Any]]:
    witnesses: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = HERE / filename
        if not path.is_file() or path.is_symlink():
            continue
        value = read_json(path)

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw_path, str) and isinstance(digest, str):
                    witnesses.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return witnesses


def _resolve_path(raw: str) -> Path:
    path = Path(raw)
    return (REPO / path).resolve() if not path.is_absolute() else path.resolve()


def binary_custody(binary: Path, digest: str, size: int) -> str:
    if binary.is_file() and not binary.is_symlink():
        require(sha(binary) == digest and binary.stat().st_size == size,
                f"live binary identity changed: {binary}")
        return "live-binary"
    for witness in cleanup_witnesses():
        candidate = _resolve_path(str(witness["path"]))
        if (candidate == binary.resolve() and witness.get("sha256") == digest
                and witness.get("bytes") == size):
            return "exact-cleanup-witness"
    fail(f"binary is absent without an exact cleanup witness: {binary}")


def build_info(source_label: str, lane: str, *, executable: bool = False) -> dict[str, Any]:
    path = HERE / f"build-{source_label}.json"
    records = read_json(path)
    require(isinstance(records, list), f"{path.name} is not a build record list")
    binary_label = "native" if lane == "native" else "alloc"
    binary_name = f"{source_label}-{binary_label}"
    matches = [item for item in records if isinstance(item, dict)
               and Path(str(item.get("binary", ""))).name == binary_name]
    require(len(matches) == 1, f"no unique build record for {binary_name}")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{binary_name} build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    digest = record.get("binary_sha256")
    size = record.get("binary_bytes")
    require(isinstance(digest, str) and len(digest) == 64
            and set(digest) <= set("0123456789abcdef"),
            f"{binary_name} binary digest is malformed")
    require(isinstance(size, int) and size > 0, f"{binary_name} binary size is malformed")
    source_path = HERE / f"source-{source_label}.json"
    expected_source = read_json(source_path)
    require(isinstance(expected_source, dict) and expected_source,
            f"{source_path.name} is not a source map")
    require(record.get("source_manifest_sha256") == sha(source_path),
            f"{binary_name} source manifest binding changed")
    custody = binary_custody(binary, digest, size)
    if executable:
        require(binary.is_file() and not binary.is_symlink(),
                f"capture binary is unavailable: {binary}")
    return {
        "source_label": source_label,
        "lane": lane,
        "name": binary_name,
        "binary": str(binary),
        "binary_sha256": digest,
        "binary_bytes": size,
        "binary_custody": custody,
        "record": record,
        "record_sha256": sha(path),
        "source_manifest": source_path.name,
        "source_manifest_sha256": sha(source_path),
        "source": expected_source,
    }


def fixture_info(corpus: dict[str, Any]) -> tuple[Path | None, dict[str, Any] | None]:
    if corpus["origin"] == "generated-harness-corpus":
        return None, None
    path = _resolve_path(str(corpus["path"]))
    require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
    actual = sha(path)
    require(actual == corpus["sha256"], f"fixture digest changed: {corpus['id']}")
    require(path.stat().st_size == corpus["bytes"], f"fixture size changed: {corpus['id']}")
    return path, {
        "path": corpus["path"],
        "resolved_path": str(path),
        "bytes": path.stat().st_size,
        "sha256": actual,
    }


def stage_metadata(plan: dict[str, Any], stage: str) -> dict[str, Any]:
    require(stage in STAGES, f"unknown stage {stage!r}")
    item = next(row for row in plan["stages"] if row["label"] == stage)
    return dict(item)


def phase_case(corpus: dict[str, Any], phase: str) -> str:
    prefix = ("docx_ordinary_save_" if corpus["origin"] == "generated-harness-corpus"
              else "docx_real_file_ordinary_save_")
    return prefix + PHASE_CASE_SUFFIX[phase]


def child_name(stage: str, lane: str, corpus: dict[str, Any], phase: str) -> str:
    return f"{slug(stage)}-{lane}-{slug(corpus['id'])}-{phase}"


def ordered_jobs(plan: dict[str, Any], stage: str) -> list[tuple[dict[str, Any], str]]:
    metadata = stage_metadata(plan, stage)
    phases = list(plan["phase_order"])
    corpora = list(plan["corpora"])
    if metadata["order"] == "reverse":
        phases.reverse()
        corpora.reverse()
    return [(corpus, phase) for phase in phases for corpus in corpora]


def expected_command(plan: dict[str, Any], build: dict[str, Any],
                     corpus: dict[str, Any], phase: str, name: str) -> list[str]:
    command = [
        "taskset", "-c", str(plan["cpu"]), build["binary"],
        "--warmup", str(plan[build["lane"]]["warmup"]),
        "--samples", str(plan[build["lane"]]["samples"]),
        "--case", phase_case(corpus, phase),
        "--json", str(HERE / f"{name}.json"),
        "--filesystem-root", str(plan["filesystem_root"]),
    ]
    if corpus["origin"] == "caller-named-real-file":
        command += ["--ooxml-file", str(corpus["path"])]
    return command


def child_environment(plan: dict[str, Any]) -> tuple[dict[str, str], dict[str, str | None]]:
    expected = dict(plan["environment"])
    environment = os.environ.copy()
    for key, value in expected.items():
        if value is None:
            environment.pop(key, None)
        else:
            environment[key] = value
    observed = {key: environment.get(key) for key in ENVIRONMENT_KEYS}
    require(observed == expected, "child environment could not be frozen")
    return environment, observed


def run_child(*, plan: dict[str, Any], stage: str, lane: str,
              corpus: dict[str, Any], phase: str, build: dict[str, Any],
              baseline_source: dict[str, str], candidate_source: dict[str, str],
              order_index: int) -> None:
    metadata = stage_metadata(plan, stage)
    name = child_name(stage, lane, corpus, phase)
    output_path = HERE / f"{name}.json"
    stdout_path = HERE / f"{name}.stdout"
    stderr_path = HERE / f"{name}.stderr"
    receipt_path = HERE / f"{name}.receipt.json"
    for path in (output_path, stdout_path, stderr_path, receipt_path):
        require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    constraints_check()
    require(build["source"] == (baseline_source if metadata["source"] == "baseline"
                                 else candidate_source),
            f"{name}: binary source manifest does not match stage source")
    before = source_census()
    relation_before = source_relation(candidate_source, before, f"{name} before")
    fixture_path, fixture_before = fixture_info(corpus)
    command = expected_command(plan, build, corpus, phase, name)
    environment, environment_observed = child_environment(plan)
    started = utc_now()
    tick = time.monotonic()
    with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
        result = subprocess.run(command, cwd=REPO, env=environment,
                                stdout=stdout, stderr=stderr)
    seconds = time.monotonic() - tick
    after = source_census()
    relation_after = source_relation(candidate_source, after, f"{name} after")
    require(before == after, f"{name}: live source changed during child")
    require(sha(Path(build["binary"])) == build["binary_sha256"],
            f"{name}: binary changed during child")
    fixture_after = None if fixture_path is None else fixture_info(corpus)[1]
    require(fixture_before == fixture_after, f"{name}: fixture changed during child")
    require(output_path.is_file() and not output_path.is_symlink(),
            f"{name}: benchmark report is missing")
    artifacts = {
        path.name: sha(path)
        for path in (output_path, stdout_path, stderr_path)
    }
    live_source = {
        "manifest": "source-candidate.json",
        "manifest_sha256": sha(HERE / "source-candidate.json"),
        "source_census_sha256": digest_json(candidate_source),
        "source_entry_count": len(candidate_source),
        "before_sha256": digest_json(before),
        "after_sha256": digest_json(after),
        "recensus_before": True,
        "recensus_after": True,
        "unchanged_during_child": before == after,
        "relation_before": relation_before,
        "relation_after": relation_after,
    }
    binary_source = {
        "manifest": build["source_manifest"],
        "manifest_sha256": build["source_manifest_sha256"],
        "source_census_sha256": digest_json(build["source"]),
        "source_entry_count": len(build["source"]),
    }
    receipt = {
        "schema_version": 1,
        "packet": plan["packet"],
        "name": name,
        "stage": stage,
        "source_label": metadata["source"],
        "build_label": metadata["build"],
        "pair": metadata["pair"],
        "cycle": metadata["cycle"],
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
        "fresh_child_process": True,
        "command": command,
        "start_utc": started,
        "end_utc": utc_now(),
        "seconds": seconds,
        "exit_code": result.returncode,
        "cpu": plan["cpu"],
        "binary_path": build["binary"],
        "binary_sha256": build["binary_sha256"],
        "binary_bytes": build["binary_bytes"],
        "build_record": f"build-{metadata['build']}.json",
        "build_record_sha256": build["record_sha256"],
        "build_source_manifest": build["source_manifest"],
        "build_source_manifest_sha256": build["source_manifest_sha256"],
        "binary_source": binary_source,
        "retained_binary_source": binary_source,
        "live_checkout_source": live_source,
        "current_checkout_source": live_source,
        "fixture": {
            "plan_path": corpus.get("path"),
            "plan_sha256": corpus.get("sha256"),
            "before": fixture_before,
            "after": fixture_after,
        },
        "plan_sha256": sha(HERE / "plan.json"),
        "script_sha256": sha(Path(__file__).resolve()),
        "constraints_sha256": sha(HERE / "constraints.json"),
        "environment": environment_observed,
        "artifacts": artifacts,
    }
    write_json(receipt_path, receipt)
    require(result.returncode == 0,
            f"{name}: benchmark failed with exit code {result.returncode}")
    print(f"{name} passed", flush=True)


def run_lane(stage: str, raw_lane: str) -> None:
    plan = load_plan()
    lane = LANE_ALIASES.get(raw_lane, raw_lane)
    require(stage in STAGES, f"unknown stage {stage!r}")
    require(lane in LANES, "lane must be native or allocator")
    if lane == "allocator":
        require(stage in plan["allocator_stages"],
                "allocator capture is frozen to the first ABBA cycle")
    baseline_source, candidate_source, _ = source_pair()
    build = build_info(stage_metadata(plan, stage)["build"], lane, executable=True)
    jobs = ordered_jobs(plan, stage)
    require(len(jobs) == 4, "fixed stage child count changed")
    for order_index, (corpus, phase) in enumerate(jobs):
        run_child(
            plan=plan, stage=stage, lane=lane, corpus=corpus, phase=phase,
            build=build, baseline_source=baseline_source,
            candidate_source=candidate_source, order_index=order_index,
        )
    print(f"stage {stage} lane {lane} complete ({len(jobs)} children)", flush=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", help="one of the eight frozen ABBA stage labels")
    parser.add_argument("lane", help="native or allocator (alloc is accepted as an alias)")
    args = parser.parse_args()
    stage, lane = args.stage, args.lane
    if stage in {"native", "alloc", "allocator", "allocation"} and lane in STAGES:
        stage, lane = lane, stage
    try:
        run_lane(stage, lane)
    except (AssertionError, KeyError, OSError, RuntimeError, ValueError) as error:
        print(f"capture failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
