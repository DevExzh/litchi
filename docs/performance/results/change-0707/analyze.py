#!/usr/bin/env python3
"""Verify the 0707 current-source XLSX diagnostic packet.

This analyzer is deliberately diagnostic.  It does not build or execute the
benchmark and it never makes a before/after performance claim.  The detailed
XLSX source, corpus, sink, budget, and native timing validators are imported
from the retained 0705 helper.  The allocator lane is checked independently:
its operation-local counters are useful diagnostics, while its instrumented
elapsed time and absolute process gauges are never treated as native timing
evidence.

The command has one bounded operation::

    python3 -B analyze.py [--output PATH]

    The default output is ``analysis.json`` beside this file.  An explicit
    ``--output`` path may be used for an independent replay destination.
"""

from __future__ import annotations

import argparse
import copy
import datetime as _datetime
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import re
import statistics
import sys
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
RESULTS = HERE.parent

CASE = "xlsx_source_backed_cell_values_one_percent_edit_save"
SHAPES = ("medium", "dense-sparse")
NATIVE_REPEATS = 2
NATIVE_SAMPLES = 100
NATIVE_WARMUP = 20
ALLOC_REPEATS = 2
ALLOC_SAMPLES = 3
ALLOC_WARMUP = 2
CPU = 12

TIMING_PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
ALL_PHASES = TIMING_PHASES + ("reopen_ns",)
ALLOC_FIELDS = (
    "plan_allocation_metrics",
    "staging_allocation_metrics",
    "commit_allocation_metrics",
    "commit_core_allocation_metrics",
    "publication_allocation_metrics",
)
ALLOC_PHASES = ("plan", "staging", "commit_core", "commit", "publication")
ALLOC_PHASE_FIELDS = {
    "plan": "plan_allocation_metrics",
    "staging": "staging_allocation_metrics",
    "commit_core": "commit_core_allocation_metrics",
    "commit": "commit_allocation_metrics",
    "publication": "publication_allocation_metrics",
}
ALLOC_METRICS = (
    "allocation_calls",
    "deallocation_calls",
    "reallocation_calls",
    "failed_allocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "live_bytes_before",
    "live_bytes_after",
    "peak_live_bytes_before",
    "peak_live_bytes_after",
    "region_peak_live_bytes",
)
RELATIVE_METRICS = (
    "allocation_calls",
    "reallocation_calls",
    "allocated_bytes",
    "deallocated_bytes",
    "net_live",
    "peak_above_start",
)
HEX = set("0123456789abcdef")
RECEIPT_KEYS = {
    "schema_version", "name", "kind", "role", "lane", "repeat", "shape", "case",
    "command", "start_utc", "end_utc", "seconds", "exit_code", "cpu", "binary_path",
    "binary_sha256", "binary_bytes", "build_record_sha256",
    "build_source_manifest_sha256", "retained_source_manifest",
    "retained_source_manifest_sha256", "retained_source_census_sha256",
    "retained_source_entry_count", "current_checkout_source", "plan_sha256",
    "script_sha256", "constraints_sha256", "profile_plan_sha256", "environment",
    "artifacts",
}
CURRENT_SOURCE_KEYS = {
    "before_artifact", "after_artifact", "before_sha256", "after_sha256",
    "before_file_sha256", "after_file_sha256", "before_entry_count",
    "after_entry_count", "relation_before", "relation_after", "unchanged_during_child",
}
ENVIRONMENT_KEYS = {"RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF", "GLIBC_TUNABLES"}
RETAINED_HELPERS = {
    "native_validator": RESULTS / "change-0705" / "analyze.py",
    "allocator_validator": RESULTS / "change-0705" / "analyze_allocations.py",
}


def load_helper(name: str, path: Path) -> Any:
    """Load a retained packet helper without copying its validator code."""

    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load retained helper: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


NATIVE_HELPER = load_helper("litchi_change_0705_native", RESULTS / "change-0705" / "analyze.py")
ALLOC_HELPER = load_helper(
    "litchi_change_0705_allocator", RESULTS / "change-0705" / "analyze_allocations.py"
)


def sha(path: Path) -> str:
    """Use the retained helper's exact file-digest convention."""

    return NATIVE_HELPER.sha(path)


def retained_helper_identities() -> dict[str, dict[str, str]]:
    """Bind imported validator code to immutable packet identities."""

    identities: dict[str, dict[str, str]] = {}
    for name, path in RETAINED_HELPERS.items():
        require(path.is_file() and not path.is_symlink(), f"retained helper is missing: {path}")
        identities[name] = {
            "path": str(path.relative_to(REPO)),
            "sha256": sha(path),
        }
    return identities


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as error:
        raise AssertionError(f"invalid JSON in {path}: {error}") from error


def read_object(path: Path) -> dict[str, Any]:
    value = read(path)
    require(isinstance(value, dict), f"{path.name} is not a JSON object")
    return value


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def nonnegative_int(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")
    return value


def finite_number(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")
    return float(value)


def list_of(value: Any, count: int, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label} is not a list")
    require(len(value) == count, f"{label} has {len(value)} values; expected {count}")
    return value


def canonical_json(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def json_sha(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def source_census() -> dict[str, str]:
    """Reconstruct build.py's source census for custody verification."""

    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path for path in (REPO / folder).rglob("*")
            if path.is_file() and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {
        str(path.relative_to(REPO)): sha(path)
        for path in sorted(set(paths))
    }


def validate_source_manifest() -> tuple[dict[str, str], str]:
    path = HERE / "source-baseline.json"
    value = read_object(path)
    require(value, "source-baseline.json is empty")
    for name, digest in value.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute(),
                f"source-baseline path is invalid: {name!r}")
        check_hex(digest, f"source-baseline[{name}]")
    current = source_census()
    require(current == value, "current source census differs from source-baseline.json")
    return value, sha(path)


def cleanup_witnesses() -> list[dict[str, Any]]:
    """Read exact post-cleanup binary witnesses, if targets were removed."""

    values: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = HERE / filename
        if not path.is_file() or path.is_symlink():
            continue
        value = read(path)

        def visit(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes", item.get("size")))
                if isinstance(raw_path, str) and isinstance(digest, str):
                    check_hex(digest, f"{filename}:{raw_path}")
                    if size is not None:
                        nonnegative_int(size, f"{filename}:{raw_path}.bytes")
                    values.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    visit(child)
            elif isinstance(item, list):
                for child in item:
                    visit(child)

        visit(value)
    return values


def resolved(path: str | Path) -> str:
    value = Path(path)
    return str(value.resolve() if value.is_absolute() else (REPO / value).resolve())


def binary_identity(path: Path, digest: str, size: int, label: str,
                    witnesses: list[dict[str, Any]]) -> dict[str, Any]:
    require(not path.is_symlink(), f"{label} binary is a symlink")
    if path.is_file():
        require(sha(path) == digest, f"{label} binary digest changed")
        require(path.stat().st_size == size, f"{label} binary byte count changed")
        # The evidence is valid either while the frozen file is retained or
        # after its exact cleanup witness has been checked.  Keep custody out
        # of the serialized identity so pre/post-cleanup analyses replay
        # byte-for-byte.
        return {"path": str(path), "sha256": digest, "bytes": size, "custody": "validated"}
    target = str(path.resolve())
    for witness in witnesses:
        if resolved(witness["path"]) == target and witness["sha256"] == digest:
            if witness["bytes"] is None or witness["bytes"] == size:
                return {
                    "path": str(path), "sha256": digest, "bytes": size,
                    "custody": "validated",
                }
    raise AssertionError(f"{label} binary is absent without an exact cleanup witness")


def build_record(lane: str, source_digest: str,
                 witnesses: list[dict[str, Any]]) -> dict[str, Any]:
    path = HERE / "build-baseline.json"
    records = read(path)
    require(isinstance(records, list), "build-baseline.json is not a record list")
    expected_name = f"baseline-{lane}"
    matches = [
        item for item in records
        if isinstance(item, dict)
        and (item.get("label") == lane
             or Path(str(item.get("binary", ""))).name == expected_name)
    ]
    require(len(matches) == 1, f"build-baseline.json has no unique {expected_name} record")
    record = matches[0]
    require(set(record) == {
        "command", "exit_code", "seconds", "source_manifest_sha256", "environment",
        "binary", "binary_sha256", "binary_bytes",
    }, f"{lane} build record keys changed")
    require(record.get("exit_code") == 0, f"{lane} build failed")
    require(isinstance(record.get("command"), list) and record["command"],
            f"{lane} build command is missing")
    environment = record.get("environment")
    require(isinstance(environment, dict) and set(environment) == ENVIRONMENT_KEYS,
            f"{lane} build environment keys changed")
    digest = record.get("binary_sha256")
    check_hex(digest, f"{lane} build binary_sha256")
    require(record.get("source_manifest_sha256") == source_digest,
            f"{lane} build is not bound to source-baseline.json")
    expected_path = (REPO.parent / "litchi-0707-bin" / expected_name).resolve()
    actual_path = Path(str(record.get("binary", ""))).resolve()
    require(actual_path == expected_path,
            f"{lane} binary path is not the frozen 0707 path: {actual_path}")
    size = nonnegative_int(record.get("binary_bytes"), f"{lane} binary_bytes")
    require(size > 0, f"{lane} binary_bytes is zero")
    identity = binary_identity(actual_path, digest, size, lane, witnesses)
    return {
        "lane": lane,
        "path": str(path),
        "sha256": sha(path),
        "binary": identity,
        "binary_sha256": digest,
        "source_manifest_sha256": source_digest,
        "command": record.get("command"),
        "seconds": finite_number(record.get("seconds"), f"{lane} build seconds"),
    }


def validate_plan() -> dict[str, Any]:
    plan = read_object(HERE / "plan.json")
    require(plan.get("case") == CASE, "plan case changed")
    require(plan.get("shapes") == list(SHAPES), "plan shape order changed")
    for key, expected in (
        ("native_repeats", NATIVE_REPEATS), ("samples", NATIVE_SAMPLES),
        ("warmup", NATIVE_WARMUP), ("allocation_repeats", ALLOC_REPEATS),
        ("allocation_samples", ALLOC_SAMPLES), ("allocation_warmup", ALLOC_WARMUP),
        ("cpu", CPU),
    ):
        require(plan.get(key) == expected, f"plan.{key} changed")
    claim = str(plan.get("claim", "")).lower()
    require("no" in claim and "speed" in claim, "plan permits a speedup claim")
    return plan


def expected_job_from_receipt(name: str, receipt: dict[str, Any], *, allocator: bool,
                              plan: dict[str, Any]) -> dict[str, Any]:
    shape = receipt.get("shape")
    repeat = receipt.get("repeat")
    if shape not in SHAPES or not isinstance(repeat, int):
        match = re.search(r"r([12])-(medium|dense-sparse)$", name)
        require(match is not None, f"{name} does not identify a shape/repeat")
        repeat = int(match.group(1))
        shape = match.group(2)
    require(shape in SHAPES and repeat in (1, 2), f"{name} shape/repeat is invalid")
    return {
        "name": name,
        "shape": shape,
        "repeat": repeat,
        "case": receipt.get("case", plan["case"]),
        "samples": ALLOC_SAMPLES if allocator else NATIVE_SAMPLES,
        "warmup": ALLOC_WARMUP if allocator else NATIVE_WARMUP,
        "kind": receipt.get("kind", "allocation" if allocator else "primary"),
        "phase": receipt.get("phase", "baseline"),
    }


def validate_receipt(root: Path, receipt_path: Path, job: dict[str, Any],
                     build: dict[str, Any], source_digest: str, plan: dict[str, Any],
                     allocator: bool) -> dict[str, Any]:
    receipt = read_object(receipt_path)
    name = job["name"]
    require(set(receipt) == RECEIPT_KEYS, f"{name} receipt keys changed")
    require(receipt["schema_version"] == 1, f"{name} receipt schema changed")
    require(receipt["name"] == name, f"{name} receipt name differs")
    require(receipt["kind"] == ("alloc" if allocator else "native"),
            f"{name} receipt kind differs")
    require(receipt["role"] == "baseline", f"{name} receipt role differs")
    require(receipt["lane"] == ("alloc" if allocator else "native"),
            f"{name} receipt lane differs")
    for key, value in (("shape", job["shape"]), ("repeat", job["repeat"]),
                       ("case", plan["case"]), ("cpu", CPU)):
        require(receipt[key] == value, f"{name} receipt {key} differs")
    require(receipt["exit_code"] == 0, f"{name} child failed")
    require(receipt["binary_sha256"] == build["binary_sha256"],
            f"{name} receipt binary digest differs")
    require(receipt["binary_bytes"] == build["binary"]["bytes"],
            f"{name} receipt binary size differs")
    require(resolved(receipt["binary_path"]) == resolved(build["binary"]["path"]),
            f"{name} receipt binary path differs")
    require(receipt["build_record_sha256"] == build["sha256"],
            f"{name} receipt build record binding differs")
    require(receipt["build_source_manifest_sha256"] == source_digest,
            f"{name} receipt source binding differs")
    retained_source = read(HERE / "source-baseline.json")
    require(receipt["retained_source_manifest"] == "source-baseline.json",
            f"{name} retained source manifest differs")
    require(receipt["retained_source_manifest_sha256"] == source_digest,
            f"{name} retained source manifest digest differs")
    require(receipt["retained_source_census_sha256"] == json_sha(retained_source),
            f"{name} retained source census digest differs")
    require(receipt["retained_source_entry_count"] == len(retained_source),
            f"{name} retained source entry count differs")
    require(receipt["plan_sha256"] == sha(root / "plan.json"),
            f"{name} receipt plan binding differs")
    require(receipt["script_sha256"] == sha(root / "capture.py"),
            f"{name} receipt capture script binding differs")
    require(receipt["constraints_sha256"] == sha(root / "constraints.json"),
            f"{name} receipt constraints binding differs")
    require(receipt["profile_plan_sha256"] is None,
            f"{name} baseline receipt unexpectedly binds a profile plan")
    environment = receipt["environment"]
    require(isinstance(environment, dict) and set(environment) == ENVIRONMENT_KEYS,
            f"{name} receipt environment keys changed")

    current = receipt["current_checkout_source"]
    require(isinstance(current, dict) and set(current) == CURRENT_SOURCE_KEYS,
            f"{name} current source receipt keys changed")
    before_name = current["before_artifact"]
    after_name = current["after_artifact"]
    require(isinstance(before_name, str) and isinstance(after_name, str),
            f"{name} source census artifact names are invalid")
    before = read_object(root / before_name)
    after = read_object(root / after_name)
    require(before == after, f"{name} source census changed during child")
    require(before == retained_source,
            f"{name} child source census differs from retained baseline")
    require(current["before_sha256"] == json_sha(before)
            and current["after_sha256"] == json_sha(after),
            f"{name} source census digest binding differs")
    require(current["before_file_sha256"] == sha(root / before_name)
            and current["after_file_sha256"] == sha(root / after_name),
            f"{name} source census file digest binding differs")
    require(current["before_entry_count"] == len(before)
            and current["after_entry_count"] == len(after),
            f"{name} source census entry count differs")
    require(current["unchanged_during_child"] is True,
            f"{name} source census changed during child")
    for relation_key in ("relation_before", "relation_after"):
        relation = current[relation_key]
        require(isinstance(relation, dict), f"{name} {relation_key} is not an object")
        require(set(relation) == {
            "mode", "changed_paths", "expected_entry_count", "current_entry_count",
        }, f"{name} {relation_key} keys changed")
        require(relation["mode"] == "exact" and relation["changed_paths"] == []
                and relation["expected_entry_count"] == len(retained_source)
                and relation["current_entry_count"] == len(retained_source),
                f"{name} {relation_key} source relation changed")
    artifacts = receipt["artifacts"]
    require(isinstance(artifacts, dict), f"{name} artifact map is invalid")
    expected_artifacts = {
        f"{name}.json", f"{name}.stdout", f"{name}.stderr",
        before_name, after_name,
    }
    require(set(artifacts) == expected_artifacts, f"{name} artifact inventory changed")
    for artifact_name, digest in artifacts.items():
        check_hex(digest, f"{name} artifact {artifact_name}")
        artifact_path = root / artifact_name
        require(artifact_path.is_file() and not artifact_path.is_symlink(),
                f"{name} artifact is missing: {artifact_name}")
        require(sha(artifact_path) == digest, f"{name} artifact digest differs: {artifact_name}")
    expected_command = [
        "taskset", "-c", str(CPU), str(build["binary"]["path"]),
        "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
        "--case", plan["case"], "--xlsx-cell-crud-shape", job["shape"],
        "--json", str(root / f"{name}.json"),
    ]
    require(receipt["command"] == expected_command,
            f"{name} command differs from frozen baseline command")
    try:
        start = _datetime.datetime.fromisoformat(receipt["start_utc"])
        end = _datetime.datetime.fromisoformat(receipt["end_utc"])
        require(start <= end, f"{name} receipt interval is inverted")
    except (KeyError, TypeError, ValueError) as error:
        raise AssertionError(f"{name} receipt timestamps are invalid") from error
    require(finite_number(receipt["seconds"], f"{name} receipt seconds") > 0,
            f"{name} receipt duration is not positive")
    return {"path": str(receipt_path), "sha256": sha(receipt_path), "receipt": receipt}


def validate_basic_report(raw: dict[str, Any], job: dict[str, Any], build: dict[str, Any],
                          plan: dict[str, Any], *, allocator: bool
                          ) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    name = job["name"]
    require(raw.get("schema_version") == 1, f"{name} report schema changed")
    tool = raw.get("tool")
    require(isinstance(tool, dict) and tool.get("profile") == "release",
            f"{name} tool profile changed")
    require(tool.get("name") == "litchi-perf-baseline",
            f"{name} tool name changed")
    require(tool.get("binary") == ("litchi-perf-baseline-alloc" if allocator
                                    else "litchi-perf-baseline"),
            f"{name} tool binary identity changed")
    require(tool.get("target_os") == "linux" and tool.get("target_arch") == "x86_64",
            f"{name} target identity changed")
    expected_instrumentation = (
        "system_allocator_operation_scoped" if allocator else "none"
    )
    require(tool.get("instrumentation") == expected_instrumentation,
            f"{name} instrumentation changed")
    if allocator:
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{name} allocator counter revision changed")
    binary = raw.get("binary_identity")
    require(isinstance(binary, dict), f"{name} binary identity is missing")
    require(binary.get("binary_sha256") == build["binary_sha256"],
            f"{name} report binary digest differs")
    require(resolved(binary.get("path", "")) == resolved(build["binary"]["path"]),
            f"{name} report binary path differs")
    require(binary.get("profile") == "release", f"{name} report binary profile changed")
    require(binary.get("binary_bytes") == build["binary"]["bytes"],
            f"{name} report binary size differs")
    environment = raw.get("environment")
    require(isinstance(environment, dict), f"{name} report environment is missing")
    if environment.get("allocator") is not None:
        require(environment["allocator"] == (
            "CountingSystemAllocator(std::alloc::System)" if allocator
            else "Rust system allocator"
        ), f"{name} allocator identity changed")
    if environment.get("git_revision") is not None:
        require(environment["git_revision"] == plan["revision"],
                f"{name} report revision differs")
    if environment.get("cpu_affinity") is not None:
        require(str(environment["cpu_affinity"]) == str(CPU),
                f"{name} report CPU affinity differs")
    configuration = raw.get("configuration")
    require(isinstance(configuration, dict), f"{name} report configuration is missing")
    require(configuration.get("cases") == [CASE], f"{name} report case configuration differs")
    require(configuration.get("xlsx_cell_crud_shapes") == [job["shape"]],
            f"{name} report shape configuration differs")
    require(configuration.get("samples_per_case") == job["samples"],
            f"{name} report sample count differs")
    require(configuration.get("warmup_iterations_per_case") == job["warmup"],
            f"{name} report warmup count differs")
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1, f"{name} report result count changed")
    result = results[0]
    require(isinstance(result, dict) and result.get("case") == CASE,
            f"{name} report case differs")
    corpus = NATIVE_HELPER.validate_corpus(result.get("corpus", {}), job["shape"])
    NATIVE_HELPER.validate_sink(result.get("sink"), name)
    check_hex(result.get("output_sha256"), f"{name}.output_sha256")
    return result, corpus, result["sink"]


def elapsed_shape(result: dict[str, Any], samples: int, label: str) -> tuple[list[int], list[int]]:
    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns",
            f"{label}.elapsed_ns is malformed")
    values = [nonnegative_int(item, f"{label}.elapsed_ns.samples[{i}]")
              for i, item in enumerate(list_of(elapsed.get("samples"), samples,
                                                f"{label}.elapsed_ns.samples"))]
    require(all(value > 0 for value in values), f"{label} elapsed sample is zero")
    order = [nonnegative_int(item, f"{label}.elapsed_ns.sample_order[{i}]")
             for i, item in enumerate(list_of(elapsed.get("sample_order"), samples,
                                               f"{label}.elapsed_ns.sample_order"))]
    require(values == sorted(values), f"{label} elapsed samples are not sorted")
    require(sorted(order) == list(range(samples)), f"{label} sample order is not a permutation")
    # The allocator lane is explicitly not a timing lane.  Check that its
    # runner did emit a coherent vector, but do not accept its reported
    # confidence interval as native evidence.
    return values, order


def source_identity(result: dict[str, Any]) -> dict[str, Any]:
    value = copy.deepcopy(result.get("source"))
    require(isinstance(value, dict), "source identity is missing")
    summary = value.get("xlsx_cell_values")
    require(isinstance(summary, dict), "XLSX source identity is missing")
    for key in ALL_PHASES:
        summary.pop(key, None)
    # Native reports carry unavailable allocator markers while allocator
    # reports carry measured counters.  Those lane-specific fields are
    # checked separately and cannot be part of cross-lane semantic identity.
    for key in ALLOC_FIELDS:
        summary.pop(key, None)

    def collapse_lists(item: Any, label: str) -> Any:
        if isinstance(item, dict):
            return {key: collapse_lists(child, f"{label}.{key}")
                    for key, child in item.items()}
        if isinstance(item, list):
            require(item, f"{label} is an empty identity vector")
            require(all(child == item[0] for child in item),
                    f"{label} varies across samples")
            return collapse_lists(item[0], label + "[0]")
        return item

    # The native and allocator lanes intentionally use different sample
    # counts.  All retained source counters are constant within a report, so
    # collapse those acquisition vectors to their validated scalar identity.
    return collapse_lists(value, "source")


def native_row(root: Path, receipt_path: Path, job: dict[str, Any], build: dict[str, Any],
               source_digest: str, plan: dict[str, Any]) -> dict[str, Any]:
    receipt_info = validate_receipt(root, receipt_path, job, build, source_digest, plan, False)
    raw = read_object(root / f"{job['name']}.json")
    result, corpus, sink = validate_basic_report(raw, job, build, plan, allocator=False)
    elapsed = NATIVE_HELPER.check_reported_statistics(result.get("elapsed_ns"), job["samples"])
    phase_vectors, constants = NATIVE_HELPER.validate_source(
        result, corpus, job["shape"], job["samples"], compatibility=False
    )
    phase_stats = {phase: NATIVE_HELPER.stats(values) for phase, values in phase_vectors.items()}
    phase_sums = [sum(phase_vectors[phase][index] for phase in TIMING_PHASES)
                  for index in range(job["samples"])]
    phase_sum_stats = NATIVE_HELPER.stats(phase_sums)
    identity = {
        "corpus": corpus,
        "sink": sink,
        "output_sha256": result["output_sha256"],
        "semantic_sha256": constants["semantic_sha256"],
        "untouched_member_count": constants["untouched_member_count"],
        "untouched_member_sha256": constants["untouched_member_sha256"],
        "source": source_identity(result),
    }
    return {
        "name": job["name"], "shape": job["shape"], "repeat": job["repeat"],
        "samples": job["samples"], "phase": job["phase"],
        "elapsed_ns": elapsed, "phase_stats": phase_stats,
        "phase_sum_stats": phase_sum_stats,
        "phase_sum_alignment": True,
        "identity": identity,
        "receipt": receipt_info,
    }


def normalize_allocator_fields(result: dict[str, Any], shape: str, samples: int, label: str) -> tuple[dict[str, Any], dict[str, list[dict[str, int]]]]:
    source = result.get("source")
    require(isinstance(source, dict), f"{label} source is missing")
    summary = source.get("xlsx_cell_values")
    require(isinstance(summary, dict), f"{label} XLSX source summary is missing")
    # Reuse the complete 0705 source/budget/resource validator after changing
    # only the lane-specific measured fields to its accepted unavailable form.
    compatibility = copy.deepcopy(result)
    compatibility_summary = compatibility["source"]["xlsx_cell_values"]
    for key in ALLOC_FIELDS:
        value = summary.get(key)
        if value is not None:
            list_of(value, samples, f"{label}.{key}")
            compatibility_summary[key] = [{"status": "unavailable"} for _ in range(samples)]
    corpus = result["corpus"]
    NATIVE_HELPER.validate_source(compatibility, corpus, shape, samples,
                                  compatibility=False)
    phases: dict[str, list[dict[str, int]]] = {}
    for phase in ALLOC_PHASES:
        key = ALLOC_PHASE_FIELDS[phase]
        raw_values = list_of(summary.get(key), samples, f"{label}.{key}")
        parsed: list[dict[str, int]] = []
        for index, item in enumerate(raw_values):
            parsed.append(ALLOC_HELPER.validate_allocator_sample(item, f"{label}.{key}[{index}]"))
        phases[phase] = parsed
    return source_identity(result), phases


def metric_summary(values: Iterable[int]) -> dict[str, Any]:
    return ALLOC_HELPER.metric_summary(list(values))


def allocator_phase_record(samples: list[dict[str, int]]) -> dict[str, Any]:
    vectors = {field: [sample[field] for sample in samples] for field in ALLOC_METRICS}
    relative = {
        "allocation_calls": vectors["allocation_calls"],
        "reallocation_calls": vectors["reallocation_calls"],
        "allocated_bytes": vectors["allocated_bytes"],
        "deallocated_bytes": vectors["deallocated_bytes"],
        "net_live": [after - before for before, after in
                      zip(vectors["live_bytes_before"], vectors["live_bytes_after"])],
        "peak_above_start": [peak - before for before, peak in
                              zip(vectors["live_bytes_before"], vectors["region_peak_live_bytes"])],
    }
    require(all(value >= 0 for key, values in relative.items() if key != "net_live"
                for value in values), "allocator relative gauge became negative")
    return {
        "sample_count": len(samples),
        "vectors": vectors,
        "metrics": {field: metric_summary(values) for field, values in vectors.items()},
        "relative_vectors": relative,
        "relative_metrics": {field: metric_summary(values) for field, values in relative.items()},
        "absolute_process_gauges": {
            field: metric_summary(vectors[field])
            for field in ("live_bytes_before", "live_bytes_after",
                          "peak_live_bytes_before", "peak_live_bytes_after",
                          "region_peak_live_bytes")
        },
    }


def allocator_row(root: Path, receipt_path: Path, job: dict[str, Any], build: dict[str, Any],
                  source_digest: str, plan: dict[str, Any]) -> dict[str, Any]:
    receipt_info = validate_receipt(root, receipt_path, job, build, source_digest, plan, True)
    raw = read_object(root / f"{job['name']}.json")
    result, corpus, sink = validate_basic_report(raw, job, build, plan, allocator=True)
    elapsed_values, elapsed_order = elapsed_shape(result, job["samples"], job["name"])
    source_info, phase_samples = normalize_allocator_fields(
        result, job["shape"], job["samples"], job["name"]
    )
    reconciliation = ALLOC_HELPER.validate_phase_reconciliation(phase_samples, job["name"])
    phases = {phase: allocator_phase_record(values) for phase, values in phase_samples.items()}
    constants = source_info["xlsx_cell_values"]
    identity = {
        "corpus": corpus,
        "sink": sink,
        "output_sha256": result["output_sha256"],
        "semantic_sha256": constants["semantic_sha256"],
        "untouched_member_count": constants["untouched_member_count"],
        "untouched_member_sha256": constants["untouched_member_sha256"],
        "source": source_info,
    }
    return {
        "name": job["name"], "shape": job["shape"], "repeat": job["repeat"],
        "samples": job["samples"], "phase": job["phase"],
        "timing_diagnostic": {
            "samples": elapsed_values, "sample_order": elapsed_order,
            "compared_with_native": False,
        },
        "phase_metrics": phases,
        "phase_reconciliation": reconciliation,
        "identity": identity,
        "receipt": receipt_info,
    }


def repeat_flags(rows: list[dict[str, Any]], shapes: Iterable[str]) -> list[dict[str, Any]]:
    flags: list[dict[str, Any]] = []
    for shape in shapes:
        pair = sorted((row for row in rows if row["shape"] == shape),
                      key=lambda row: row["repeat"])
        require(len(pair) == 2, f"{shape} does not have two native repeats")
        first, second = pair
        for label, left, right in (
            ("elapsed_ns", first["elapsed_ns"], second["elapsed_ns"]),
            *[(phase, first["phase_stats"][phase], second["phase_stats"][phase])
              for phase in ALL_PHASES],
            ("phase_sum_ns", first["phase_sum_stats"], second["phase_sum_stats"]),
        ):
            for metric in ("p50", "p95", "p99", "mean"):
                if left[metric] == 0:
                    raise AssertionError(f"{shape} {label}.{metric} has zero repeat-1 value")
                change = (float(right[metric]) / float(left[metric]) - 1.0) * 100.0
                if abs(change) > 5.0:
                    flags.append({
                        "shape": shape, "metric_group": label, "metric": metric,
                        "repeat1": left[metric], "repeat2": right[metric],
                        "change_percent": change,
                    })
    return flags


def relative_repeat_records(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for shape in SHAPES:
        pair = sorted((row for row in rows if row["shape"] == shape),
                      key=lambda row: row["repeat"])
        require(len(pair) == 2, f"{shape} does not have two allocator repeats")
        for phase in ALLOC_PHASES:
            vectors = {
                field: pair[0]["phase_metrics"][phase]["relative_vectors"][field]
                + pair[1]["phase_metrics"][phase]["relative_vectors"][field]
                for field in RELATIVE_METRICS
            }
            stable = all(len(set(values)) == 1 for values in vectors.values())
            records.append({
                "shape": shape, "phase": phase, "samples": 2 * ALLOC_SAMPLES,
                "metrics": {field: values[0] for field, values in vectors.items()}
                if stable else None,
                "vectors": vectors,
                "all_six_relative_samples_identical": stable,
                "absolute_process_gauges_not_compared": True,
                "whole_operation_peak_not_added": True,
            })
    return records


def identity_parity(native: list[dict[str, Any]], allocator: list[dict[str, Any]]) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for shape in SHAPES:
        for repeat in (1, 2):
            n = next(row for row in native if row["shape"] == shape and row["repeat"] == repeat)
            a = next(row for row in allocator if row["shape"] == shape and row["repeat"] == repeat)
            require(n["identity"] == a["identity"],
                    f"allocator/native identity differs for {shape} repeat {repeat}")
            records.append({
                "shape": shape, "repeat": repeat,
                "corpus_equal": True, "output_sha256_equal": True,
                "source_identity_equal": True, "sink_equal": True,
                "resource_identity_equal": True,
                "allocator_timing_compared": False,
                "absolute_process_gauges_compared": False,
            })
    return records


def discover_receipts(prefix: str) -> list[Path]:
    return sorted(path for path in HERE.glob(f"{prefix}-*.receipt.json")
                   if path.is_file() and not path.is_symlink())


def analyze(root: Path = HERE) -> dict[str, Any]:
    root = Path(root).resolve()
    require(root == HERE, "analyzer root must be its own packet directory")
    plan = validate_plan()
    source, source_digest = validate_source_manifest()
    helpers = retained_helper_identities()
    witnesses = cleanup_witnesses()
    builds = {
        "native": build_record("native", source_digest, witnesses),
        "alloc": build_record("alloc", source_digest, witnesses),
    }
    native_receipts = discover_receipts("native")
    alloc_receipts = discover_receipts("alloc")
    expected_count = len(SHAPES) * NATIVE_REPEATS
    require(len(native_receipts) == expected_count,
            f"native receipt count is {len(native_receipts)}; expected {expected_count}")
    require(len(alloc_receipts) == expected_count,
            f"allocator receipt count is {len(alloc_receipts)}; expected {expected_count}")
    native_rows: list[dict[str, Any]] = []
    alloc_rows: list[dict[str, Any]] = []
    for path in native_receipts:
        receipt = read_object(path)
        job = expected_job_from_receipt(path.name[:-len(".receipt.json")], receipt,
                                         allocator=False, plan=plan)
        native_rows.append(native_row(root, path, job, builds["native"], source_digest, plan))
    for path in alloc_receipts:
        receipt = read_object(path)
        job = expected_job_from_receipt(path.name[:-len(".receipt.json")], receipt,
                                         allocator=True, plan=plan)
        alloc_rows.append(allocator_row(root, path, job, builds["alloc"], source_digest, plan))
    native_rows.sort(key=lambda row: (row["repeat"], SHAPES.index(row["shape"])))
    alloc_rows.sort(key=lambda row: (row["repeat"], SHAPES.index(row["shape"])))
    require({(row["shape"], row["repeat"]) for row in native_rows}
            == {(shape, repeat) for repeat in (1, 2) for shape in SHAPES},
            "native receipt shape/repeat set is incomplete or duplicated")
    require({(row["shape"], row["repeat"]) for row in alloc_rows}
            == {(shape, repeat) for repeat in (1, 2) for shape in SHAPES},
            "allocator receipt shape/repeat set is incomplete or duplicated")
    flags = repeat_flags(native_rows, SHAPES)
    native_identities: dict[str, Any] = {}
    for shape in SHAPES:
        pair = [row for row in native_rows if row["shape"] == shape]
        require(pair[0]["identity"] == pair[1]["identity"],
                f"native repeat identity differs for {shape}")
        native_identities[shape] = pair[0]["identity"]
    allocator_identity: dict[str, Any] = {}
    for shape in SHAPES:
        pair = [row for row in alloc_rows if row["shape"] == shape]
        require(pair[0]["identity"] == pair[1]["identity"],
                f"allocator repeat identity differs for {shape}")
        allocator_identity[shape] = pair[0]["identity"]
    return {
        "schema_version": 1,
        "status": "pass",
        "scope": "0707 current-source XLSX source-backed edit/save diagnostic; no before/after speedup claim",
        "claim": "current-source phase attribution and allocator diagnostics only",
        "revision": plan["revision"],
        "case": CASE,
        "plan": plan,
        "source": {
            "manifest": "source-baseline.json",
            "manifest_sha256": source_digest,
            "entry_count": len(source),
            "current_census_equal": True,
        },
        "retained_helpers": helpers,
        "builds": builds,
        "native": {
            "runs": native_rows,
            "repeat_drift_over_five_percent": flags,
            "aa_repeat_flags": flags,
            "identities": native_identities,
            "timing_comparison": None,
        },
        "allocator": {
            "runs": alloc_rows,
            "relative_repeat_consistency": relative_repeat_records(alloc_rows),
            "identities": allocator_identity,
            "native_identity_parity": identity_parity(native_rows, alloc_rows),
            "whole_operation_peak": None,
            "whole_operation_peak_policy": (
                "never sum phase peaks or phase live deltas; commit retains its own observer-ordered region peak"
            ),
            "timing_comparison": None,
            "absolute_process_gauges_comparison": None,
        },
        "checks": {
            "retained_0705_native_validators": True,
            "retained_helper_sha256_identities": True,
            "build_record_and_binary_custody": True,
            "source_manifest_and_current_census": True,
            "native_phase_sums_recomputed": True,
            "native_corpus_output_source_sink_resource_identity": True,
            "native_aa_repeat_flags": True,
            "allocator_all_phase_vectors": True,
            "allocator_six_relative_metrics": list(RELATIVE_METRICS),
            "allocator_peak_not_added": True,
            "allocator_absolute_process_gauges_not_parity": True,
        },
        "unavailable": [
            "historical before/after speedup",
            "native Office producer",
            "cold-cache physical I/O",
            "parallel scaling",
            "allocator-instrumented elapsed timing as native evidence",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=HERE / "analysis.json")
    args = parser.parse_args()
    output = args.output.resolve()
    try:
        report = analyze()
    except (AssertionError, OSError, TypeError, ValueError, KeyError, RuntimeError,
            NATIVE_HELPER.__dict__.get("EvidenceError", RuntimeError),
            ALLOC_HELPER.__dict__.get("EvidenceError", RuntimeError)) as error:
        report = {
            "schema_version": 1,
            "status": "error",
            "scope": "0707 current-source diagnostic",
            "error": str(error),
            "whole_operation_peak": None,
            "whole_operation_peak_policy": "not derived by summing phase peaks",
        }
        output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(f"0707 analysis error: {error}", file=sys.stderr)
        return 1
    output.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(f"0707 analysis {report['status']}: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
