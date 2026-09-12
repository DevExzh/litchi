#!/usr/bin/env python3
"""Validate the conditional 0531 XLSX planning Callgrind lane.

The native 0531 admission is produced by ``analyze.py``.  This companion
analyzer is deliberately independent of that admission calculation: it
binds each profile receipt to its stage source manifest, build receipt and
binary, checks the four raw dumps, and then compares the selected planning
edge between the frozen baseline and final stages.

The retained 0530 planning analyzer supplies the raw Callgrind parser and
annotation helpers.  It is loaded by path, and its evidence directory is
rebound to this campaign.  No profiling or annotation is started unless the
caller explicitly passes ``--create-annotations``.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import re
import sys
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
PLANNING_PATH = HERE.parent / "change-0530" / "analyze_planning.py"
_planning_spec = importlib.util.spec_from_file_location(
    "litchi_xlsx_planning_0530_for_0531", PLANNING_PATH
)
if _planning_spec is None or _planning_spec.loader is None:
    raise ImportError(f"cannot load retained planning analyzer: {PLANNING_PATH}")
PLANNING = importlib.util.module_from_spec(_planning_spec)
_planning_spec.loader.exec_module(PLANNING)

# The 0530 module loaded its immutable 0521/0519 helpers with the old
# evidence directory.  Rebinding only path presentation is enough: all raw
# parser and annotation code is otherwise retained unchanged.
RAW = PLANNING.h
PLANNING.HERE = HERE
RAW.HERE = HERE


OWNER = "litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets"
LIFECYCLE_PARENT = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
MEASURED_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
WORKSHEET_PARSE = "litchi_xlsx::raw::worksheet::parse"
MCE_PROCESS = "litchi_ooxml_common::mce::codec::process_ooxml"
MCE_NAMESPACE = "litchi_ooxml_common::mce::codec::process_markup_compatibility"
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns")
IDENTITY_TIMING_FIELDS = frozenset(PHASES)
IDENTITY_ALLOCATION_FIELDS = frozenset(
    {"commit_allocation_metrics", "publication_allocation_metrics"}
)
NUMBERED_PARTS = (1, 2, 3, 4)
PROFILE_PREFIX = "profile-"


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def read_text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside the evidence directory: {path}") from error


def regular(path: Path, label: str) -> None:
    require(path.is_file() and not path.is_symlink(), f"{label} is not a regular file")


def plan_data() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "plan.profile is not an object")
    require(profile.get("owner") == OWNER, "profile owner differs from the frozen plan")
    require(profile.get("shapes") == ["medium", "dense-sparse"],
            "profile shape matrix differs from the frozen plan")
    require(profile.get("repeats") == 2, "profile repeat count differs from the frozen plan")
    require(profile.get("warmup") == 0, "profile warmup differs from the frozen plan")
    require(profile.get("samples") == 1, "profile sample count differs from the frozen plan")
    primary = plan.get("primary")
    require(isinstance(primary, dict), "plan.primary is not an object")
    require(primary.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save",
            "profile case differs from the frozen primary case")
    gates = plan.get("gates")
    require(isinstance(gates, dict), "plan.gates is not an object")
    required_gate = gates.get("planning_ir_reduction_percent")
    require(isinstance(required_gate, (int, float)) and not isinstance(required_gate, bool)
            and required_gate >= 0,
            "planning Ir gate is malformed")
    return plan


def expected_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    return [
        {
            "name": f"profile-r{repeat}-{shape}",
            "repeat": repeat,
            "shape": shape,
            "case": plan["primary"]["case"],
            "warmup": int(profile["warmup"]),
            "samples": int(profile["samples"]),
        }
        for repeat in range(1, int(profile["repeats"]) + 1)
        for shape in profile["shapes"]
    ]


def native_name(job: dict[str, Any]) -> str:
    return f"final-native-r{job['repeat']}-primary-{job['shape']}"


def _nonnegative_integer(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")


def _canonical(value: Any, label: str) -> Any:
    """Collapse constant sample vectors for profile/native identity parity."""

    if isinstance(value, list):
        if not value:
            return []
        values = [_canonical(item, f"{label}[{index}]")
                  for index, item in enumerate(value)]
        require(all(item == values[0] for item in values),
                f"{label} is not constant across retained samples")
        return values[0]
    if isinstance(value, dict):
        return {key: _canonical(item, f"{label}.{key}")
                for key, item in sorted(value.items())}
    return value


def source_identity(source: Any, label: str) -> dict[str, Any]:
    require(isinstance(source, dict), f"{label} is not an object")
    result: dict[str, Any] = {}
    for key, value in sorted(source.items()):
        if key != "xlsx_cell_values":
            result[key] = _canonical(value, f"{label}.{key}")
            continue
        require(isinstance(value, dict), f"{label}.xlsx_cell_values is not an object")
        inner: dict[str, Any] = {}
        for inner_key, inner_value in sorted(value.items()):
            if inner_key in IDENTITY_TIMING_FIELDS or inner_key in IDENTITY_ALLOCATION_FIELDS:
                require(isinstance(inner_value, list),
                        f"{label}.xlsx_cell_values.{inner_key} is not a vector")
                continue
            inner[inner_key] = _canonical(
                inner_value, f"{label}.xlsx_cell_values.{inner_key}"
            )
        result[key] = inner
    return result


def normalized_configuration(configuration: Any, label: str) -> dict[str, Any]:
    require(isinstance(configuration, dict), f"{label}.configuration is not an object")
    result = dict(configuration)
    for field in ("samples_per_case", "warmup_iterations_per_case"):
        require(field in result, f"{label}.configuration omits {field}")
        result[field] = "validated-planned-count"
    return result


def result_identity(result: dict[str, Any], source: dict[str, Any],
                    configuration: Any, label: str) -> dict[str, Any]:
    corpus = result.get("corpus")
    sink = result.get("sink")
    output = result.get("output_sha256")
    require(isinstance(corpus, dict), f"{label}.corpus is not an object")
    require(isinstance(sink, dict), f"{label}.sink is not an object")
    require(isinstance(output, str) and len(output) == 64,
            f"{label}.output_sha256 is malformed")
    return {
        "configuration": normalized_configuration(configuration, label),
        "corpus": corpus,
        "sink": sink,
        "source": source_identity(source, f"{label}.source"),
        "output_sha256": output,
    }


def validate_sink(sink: Any, label: str) -> dict[str, Any]:
    require(isinstance(sink, dict), f"{label}.sink is not an object")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        _nonnegative_integer(sink.get(key), f"{label}.sink.{key}")
    buckets = sink.get("write_size_buckets")
    require(isinstance(buckets, dict), f"{label}.sink.write_size_buckets is missing")
    for key, value in buckets.items():
        _nonnegative_integer(value, f"{label}.sink.write_size_buckets.{key}")
    require(sum(buckets.values()) == sink["write_calls"],
            f"{label}.sink.write_size_buckets does not reconcile")
    require(sink["largest_write"] <= 65_536,
            f"{label}.sink.largest_write exceeds the bounded write size")
    return sink


def validate_result(path: Path, plan: dict[str, Any], job: dict[str, Any],
                    binary_sha: str, samples: int, warmup: int) -> dict[str, Any]:
    label = relative(path)
    raw = read_json(path)
    require(isinstance(raw, dict), f"{label} is not an object")
    require(raw.get("schema_version") == 1, f"{label}.schema_version is not 1")
    tool = raw.get("tool")
    require(isinstance(tool, dict), f"{label}.tool is not an object")
    require(tool.get("binary") == "litchi-perf-baseline",
            f"{label}.tool.binary is unexpected")
    require(tool.get("profile") == "release", f"{label}.tool.profile is not release")
    require(tool.get("instrumentation") == "none",
            f"{label}.tool.instrumentation is not none")
    identity = raw.get("binary_identity")
    require(isinstance(identity, dict) and identity.get("binary_sha256") == binary_sha,
            f"{label}.binary_identity is bound to a different binary")
    environment = raw.get("environment")
    require(isinstance(environment, dict), f"{label}.environment is not an object")
    require(environment.get("git_revision") == plan["revision"],
            f"{label}.environment.git_revision differs from plan")
    require(environment.get("cpu_affinity") == str(plan["cpu"]),
            f"{label}.environment.cpu_affinity differs from plan")
    configuration = raw.get("configuration")
    require(isinstance(configuration, dict), f"{label}.configuration is not an object")
    require(configuration.get("cases") == [job["case"]],
            f"{label}.configuration.cases differs from command")
    require(configuration.get("xlsx_cell_crud_shapes") == [job["shape"]],
            f"{label}.configuration.xlsx_cell_crud_shapes differs from command")
    require(configuration.get("samples_per_case") == samples,
            f"{label}.configuration.samples_per_case differs from command")
    require(configuration.get("warmup_iterations_per_case") == warmup,
            f"{label}.configuration.warmup_iterations_per_case differs from command")
    results = raw.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}.results must contain one result")
    result = results[0]
    require(isinstance(result, dict), f"{label}.results[0] is not an object")
    require(result.get("case") == job["case"], f"{label}.results[0].case differs")
    corpus = result.get("corpus")
    require(isinstance(corpus, dict) and corpus.get("shape") == job["shape"],
            f"{label}.corpus.shape differs from command")

    elapsed = result.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{label}.elapsed_ns is not an object")
    elapsed_values = elapsed.get("samples")
    require(isinstance(elapsed_values, list) and len(elapsed_values) == samples,
            f"{label}.elapsed_ns.samples has the wrong length")
    for index, value in enumerate(elapsed_values):
        _nonnegative_integer(value, f"{label}.elapsed_ns.samples[{index}]")
    sample_order = elapsed.get("sample_order")
    require(isinstance(sample_order, list) and sorted(sample_order) == list(range(samples)),
            f"{label}.elapsed_ns.sample_order is not a permutation")

    source = result.get("source")
    require(isinstance(source, dict), f"{label}.source is not an object")
    xlsx = source.get("xlsx_cell_values")
    require(isinstance(xlsx, dict), f"{label}.source.xlsx_cell_values is missing")
    require(xlsx.get("implementation") == "source-backed",
            f"{label}.source.xlsx_cell_values implementation is unexpected")
    require(xlsx.get("cache_mode") == "unmanaged-control",
            f"{label}.source.xlsx_cell_values cache mode is unexpected")
    phase_values: dict[str, list[int]] = {}
    for phase in PHASES:
        values = xlsx.get(phase)
        require(isinstance(values, list) and len(values) == samples,
                f"{label}.source.xlsx_cell_values.{phase} has the wrong length")
        for index, value in enumerate(values):
            _nonnegative_integer(value, f"{label}.{phase}[{index}]")
        phase_values[phase] = values
    output = result.get("output_sha256")
    require(isinstance(output, str) and len(output) == 64,
            f"{label}.output_sha256 is malformed")
    source_output = xlsx.get("output_sha256")
    require(isinstance(source_output, list) and len(source_output) == samples
            and all(value == output for value in source_output),
            f"{label}.source output digest differs from result output")
    for sorted_index, acquisition_index in enumerate(sample_order):
        phase_sum = sum(phase_values[phase][acquisition_index] for phase in PHASES[:-1])
        require(phase_sum == elapsed_values[sorted_index],
                f"{label} phase sum does not reproduce elapsed sample {sorted_index}")
    validate_sink(result.get("sink"), label)
    logical = result_identity(result, source, configuration, label)
    return {
        "file": label,
        "sha256": sha(path),
        "raw": raw,
        "logical_identity": logical,
        "logical_identity_sha256": hashlib.sha256(
            json.dumps(logical, sort_keys=True, separators=(",", ":")).encode()
        ).hexdigest(),
        "elapsed_samples": elapsed_values,
    }


def validate_build(stage: str, plan: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    manifest = folder / "source-manifest.json"
    binary_path = folder / "binary-normal.json"
    build_path = folder / "build-normal.receipt.json"
    for path, label in ((manifest, f"{stage} source manifest"),
                        (binary_path, f"{stage} binary identity"),
                        (build_path, f"{stage} build receipt")):
        regular(path, label)
    manifest_value = read_json(manifest)
    require(isinstance(manifest_value, dict) and manifest_value,
            f"{stage} source manifest is empty")
    manifest_sha = sha(manifest)
    build = read_json(build_path)
    require(isinstance(build, dict), f"{stage} build receipt is not an object")
    require(build.get("exit_code") == 0, f"{stage} build did not exit successfully")
    require(build.get("binary_sha256") is None, f"{stage} build receipt has a child binary hash")
    require(build.get("source_manifest_sha256") == manifest_sha,
            f"{stage} build/source manifest binding differs")
    require(build.get("working_source_manifest_sha256") == manifest_sha,
            f"{stage} build working source binding differs")
    require(build.get("plan_sha256") == sha(HERE / "plan.json"),
            f"{stage} build/plan binding differs")
    require(build.get("script_sha256") == sha(HERE / "run.py"),
            f"{stage} build/script binding differs")
    command = build.get("command")
    require(isinstance(command, list), f"{stage} build command is not a list")
    for token in ("cargo", "build", "--release", "--locked", "--manifest-path", "--bin"):
        require(token in command, f"{stage} build command omits {token}")
    require(command[command.index("--bin") + 1] == "litchi-perf-baseline",
            f"{stage} build binary differs")
    build_artifacts = {"build-normal.stdout", "build-normal.stderr"}
    require(set(build.get("artifacts", {})) == build_artifacts,
            f"{stage} build artifact inventory differs")
    for filename, digest in build["artifacts"].items():
        artifact = folder / filename
        regular(artifact, f"{stage}/{filename}")
        require(sha(artifact) == digest, f"{stage}/{filename} hash differs from build receipt")

    binary = read_json(binary_path)
    require(isinstance(binary, dict), f"{stage} binary identity is not an object")
    expected_path = Path(plan["owned_paths"][0]) / f"{stage}-normal"
    require(binary.get("path") == str(expected_path), f"{stage} binary path differs")
    digest = binary.get("sha256")
    require(isinstance(digest, str) and re.fullmatch(r"[0-9a-f]{64}", digest),
            f"{stage} binary hash is malformed")
    _nonnegative_integer(binary.get("bytes"), f"{stage} binary byte count")
    require(binary["bytes"] > 0, f"{stage} binary is empty")
    require(binary.get("source_manifest_sha256") == manifest_sha,
            f"{stage} binary/source manifest binding differs")
    require(binary.get("build_receipt_sha256") == sha(build_path),
            f"{stage} binary/build receipt binding differs")
    binary_file = Path(binary["path"])
    if binary_file.exists():
        regular(binary_file, f"{stage} retained binary")
        require(binary_file.stat().st_size == binary["bytes"], f"{stage} binary size differs")
        require(sha(binary_file) == digest, f"{stage} binary hash differs from identity")
    else:
        cleanup = read_json(HERE / "cleanup.json")
        require(cleanup.get("plan_sha256") == sha(HERE / "plan.json")
                and cleanup.get("owned_paths_absent") is True
                and cleanup.get("removed") == plan["owned_paths"]
                and all(not Path(path).exists() for path in plan["owned_paths"]),
                "missing binary has no bound cleanup receipt")
    return {
        "stage": stage,
        "source_manifest": relative(manifest),
        "source_manifest_sha256": manifest_sha,
        "build_receipt": relative(build_path),
        "build_receipt_sha256": sha(build_path),
        "binary_identity": relative(binary_path),
        "binary_sha256": digest,
        "binary_bytes": binary["bytes"],
        "binary_path": str(binary_file),
        "source_manifest_entries": len(manifest_value),
    }


def profile_command(stage: str, job: dict[str, Any], plan: dict[str, Any],
                    binary: str) -> list[str]:
    folder = HERE / stage
    name = job["name"]
    owner = plan["profile"]["owner"]
    return [
        "taskset", "-c", str(plan["cpu"]),
        "valgrind", "--tool=callgrind", "--collect-atstart=no",
        f"--toggle-collect={owner}", f"--zero-before={owner}",
        f"--dump-after={owner}",
        f"--callgrind-out-file={folder / (name + '.callgrind')}",
        binary,
        "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
        "--case", job["case"], "--xlsx-cell-crud-shape", job["shape"],
        "--json", str(folder / (name + ".json")),
    ]


def validate_profile_receipt(stage: str, job: dict[str, Any], plan: dict[str, Any],
                             meta: dict[str, Any]) -> dict[str, Any]:
    folder = HERE / stage
    name = job["name"]
    path = folder / f"{name}.receipt.json"
    regular(path, relative(path))
    receipt = read_json(path)
    require(isinstance(receipt, dict), f"{relative(path)} is not an object")
    require(receipt.get("exit_code") == 0, f"{relative(path)} child failed")
    require(receipt.get("binary_sha256") == meta["binary_sha256"],
            f"{relative(path)} binary binding differs")
    require(receipt.get("source_manifest_sha256") == meta["source_manifest_sha256"],
            f"{relative(path)} source binding differs")
    candidate_manifest = HERE / ("candidate" if stage == "baseline" else "final") / "source-manifest.json"
    expected_working = sha(candidate_manifest) if candidate_manifest.is_file() \
        else meta["source_manifest_sha256"]
    require(receipt.get("working_source_manifest_sha256") == expected_working,
            f"{relative(path)} working source binding differs")
    require(receipt.get("plan_sha256") == sha(HERE / "plan.json"),
            f"{relative(path)} plan binding differs")
    require(receipt.get("script_sha256") == sha(HERE / "run.py"),
            f"{relative(path)} script binding differs")
    command = receipt.get("command")
    require(command == profile_command(stage, job, plan, meta["binary_path"]),
            f"{relative(path)} command differs from the frozen profile command")
    expected_artifacts = {
        f"{name}.json", f"{name}.stdout", f"{name}.stderr",
        f"{name}.callgrind",
        *(f"{name}.callgrind.{part}" for part in NUMBERED_PARTS),
    }
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts,
            f"{relative(path)} artifact inventory differs")
    for filename, digest in artifacts.items():
        artifact = folder / filename
        regular(artifact, f"{stage}/{filename}")
        require(sha(artifact) == digest,
                f"{stage}/{filename} hash differs from profile receipt")
    return {
        "file": relative(path),
        "sha256": sha(path),
        "start_utc": receipt.get("start_utc"),
        "end_utc": receipt.get("end_utc"),
        "artifacts": {key: artifacts[key] for key in sorted(artifacts)},
    }


def validate_raw_dump(path: Path, part: int, expected_parent: str,
                      plan: dict[str, Any]) -> dict[str, Any]:
    label = relative(path)
    text = read_text(path)
    lines = [line.strip() for line in text.splitlines()]
    require("events: Ir" in lines and all(
        not line.startswith("events:") or line == "events: Ir" for line in lines
    ), f"{label} event set is not exactly Ir")
    require(RAW.part_number(text, label) == part,
            f"{label} part number differs from suffix")
    require(RAW.trigger(text, label) == f"--dump-after={OWNER}",
            f"{label} trigger does not identify the selected owner")
    total = RAW.summary_ir(text, label)
    require(total > 0, f"{label} has no positive selected Ir")
    incoming = RAW.target_edge_summary(path, OWNER, None)
    require(incoming["positive_edge_count"] == 1 and incoming["calls"] == 1,
            f"{label} selected owner does not have one positive incoming call")
    edge = incoming["edges"][0]
    require(edge["caller"] == expected_parent,
            f"{label} selected owner caller differs: {edge['caller']!r}")
    require(edge["inclusive_ir"] == total,
            f"{label} selected owner edge does not equal raw summary")
    return {
        "part": part,
        "file": label,
        "sha256": sha(path),
        "summary_ir": total,
        "parent_edge": edge,
        "parent_edge_unique": True,
    }


def validate_termination(path: Path) -> dict[str, Any]:
    label = relative(path)
    text = read_text(path)
    lines = [line.strip() for line in text.splitlines()]
    require("events: Ir" in lines and all(
        not line.startswith("events:") or line == "events: Ir" for line in lines
    ), f"{label} event set is not exactly Ir")
    require(RAW.part_number(text, label) == 5, f"{label} termination part is not 5")
    require(RAW.trigger(text, label) == "Program termination",
            f"{label} is not a termination dump")
    total = RAW.summary_ir(text, label)
    require(total == 0, f"{label} termination Ir is {total}, expected zero")
    return {"part": 5, "file": label, "sha256": sha(path), "summary_ir": total}


def _edge(path: Path, target: str, caller: str) -> dict[str, Any]:
    value = RAW.target_edge_summary(path, target, caller)
    require(value["positive_edge_count"] == 1 and value["calls"] > 0,
            f"{relative(path)} missing unique MCE edge {caller} -> {target}")
    return value


def validate_annotations(stage: str, job: dict[str, Any], selected: Path,
                          total: int, create_annotations: bool) -> dict[str, Any]:
    folder = HERE / stage
    name = job["name"]
    outputs: dict[bool, tuple[Path, str]] = {}
    for inclusive in (True, False):
        suffix = ".inclusive.txt" if inclusive else ".self.txt"
        path = folder / (name + suffix)
        output, command = RAW.run_annotation(selected, inclusive)
        if create_annotations:
            if path.exists():
                require(read_text(path) == output,
                        f"{relative(path)} differs from deterministic annotation replay")
            else:
                path.write_text(output, encoding="utf-8")
        else:
            regular(path, relative(path))
            require(read_text(path) == output,
                    f"{relative(path)} differs from deterministic annotation replay")
        outputs[inclusive] = (path, output)

    inclusive_path, inclusive_text = outputs[True]
    self_path, self_text = outputs[False]
    inclusive = RAW.parse_annotation(inclusive_text, OWNER, relative(inclusive_path))
    exclusive = RAW.parse_annotation(self_text, OWNER, relative(self_path))
    inclusive_direct = RAW.direct_map(inclusive["direct"])
    self_direct = RAW.direct_map(exclusive["direct"])
    require(inclusive["selected_ir"] == total,
            f"{relative(inclusive_path)} owner Ir differs from raw edge")
    require(inclusive_direct == self_direct,
            f"{name} inclusive/self direct edges differ")
    require(inclusive["selected_ir"] == exclusive["selected_ir"] + sum(inclusive_direct.values()),
            f"{name} self plus direct Ir equation does not hold")
    raw_direct_edges = {}
    for child, cost in sorted(inclusive_direct.items()):
        edge = RAW.target_edge_summary(selected, child, OWNER)
        require(edge["positive_edge_count"] > 0 and edge["inclusive_ir"] == cost,
                f"{name} annotation edge does not match raw edge: {child}")
        raw_direct_edges[child] = edge

    # Keep the shared OOXML mechanism tied to the worksheet parse parent in
    # both stages.  The exact search helper may be inlined or selected by a
    # different architecture-specific memchr symbol, so only stable MCE
    # parent edges are required here.
    mce_edges = {
        "worksheet_parse_to_process_ooxml": _edge(
            selected, MCE_PROCESS, WORKSHEET_PARSE
        ),
        "process_ooxml_to_process_markup_compatibility": _edge(
            selected, MCE_NAMESPACE, MCE_PROCESS
        ),
    }
    mce_inclusive = RAW.parse_annotation(
        inclusive_text, MCE_NAMESPACE, relative(inclusive_path)
    )
    mce_self = RAW.parse_annotation(self_text, MCE_NAMESPACE, relative(self_path))
    mce_direct = RAW.direct_map(mce_inclusive["direct"])
    require(mce_direct == RAW.direct_map(mce_self["direct"]),
            f"{name} MCE inclusive/self direct edges differ")
    require(mce_inclusive["selected_ir"] == mce_self["selected_ir"] + sum(mce_direct.values()),
            f"{name} MCE self plus direct Ir equation does not hold")
    return {
        "selected_dump": relative(selected),
        "inclusive": {
            "file": relative(inclusive_path),
            "sha256": sha(inclusive_path),
            "owner_ir": inclusive["selected_ir"],
            "owner_self_ir": exclusive["selected_ir"],
            "direct_callee_ir": dict(sorted(inclusive_direct.items())),
        },
        "self": {"file": relative(self_path), "sha256": sha(self_path)},
        "mce": {
            "parent_edges": mce_edges,
            "process_markup_compatibility_ir": mce_inclusive["selected_ir"],
            "process_markup_compatibility_self_ir": mce_self["selected_ir"],
            "direct_callee_ir": dict(sorted(mce_direct.items())),
        },
        "raw_direct_edges": raw_direct_edges,
        "validation": {
            "owner_raw_inclusive_matches": True,
            "owner_self_plus_direct_equation": True,
            "owner_annotation_edges_match_raw": True,
            "mce_parent_edges_present": True,
            "mce_self_plus_direct_equation": True,
        },
    }


def validate_profile(stage: str, job: dict[str, Any], plan: dict[str, Any],
                     meta: dict[str, Any], create_annotations: bool) -> dict[str, Any]:
    folder = HERE / stage
    receipt = validate_profile_receipt(stage, job, plan, meta)
    profile_path = folder / f"{job['name']}.json"
    native_path = folder / f"{native_name(job)}.json"
    regular(profile_path, relative(profile_path))
    regular(native_path, relative(native_path))
    profile = validate_result(profile_path, plan, job, meta["binary_sha256"],
                              job["samples"], job["warmup"])
    native = validate_result(
        native_path, plan, job, meta["binary_sha256"],
        int(plan["primary"]["samples"]), int(plan["primary"]["warmup"])
    )
    require(profile["logical_identity"] == native["logical_identity"],
            f"{job['name']} profile/native logical identity differs")

    numbered = sorted(
        folder.glob(f"{job['name']}.callgrind.[0-9]*"),
        key=lambda path: int(path.name.rsplit(".", 1)[1])
    )
    require([int(path.name.rsplit(".", 1)[1]) for path in numbered] == list(NUMBERED_PARTS),
            f"{job['name']} must retain exactly four numbered Callgrind dumps")
    raw_parts = []
    for part, path in zip(NUMBERED_PARTS, numbered):
        expected_parent = MEASURED_PARENT if part == 4 else LIFECYCLE_PARENT
        raw_parts.append(validate_raw_dump(path, part, expected_parent, plan))
        if part < 4:
            lifecycle = RAW.target_edge_summary(path, OWNER, MEASURED_PARENT)
            require(lifecycle["positive_edge_count"] == 0,
                    f"{relative(path)} has an unexpected measured-parent edge")
        else:
            lifecycle = RAW.target_edge_summary(path, OWNER, LIFECYCLE_PARENT)
            require(lifecycle["positive_edge_count"] == 0,
                    f"{relative(path)} has an unexpected lifecycle-parent edge")
    termination = validate_termination(folder / f"{job['name']}.callgrind")
    annotation = validate_annotations(
        stage, job, numbered[-1], raw_parts[-1]["summary_ir"], create_annotations
    )
    return {
        "name": job["name"],
        "repeat": job["repeat"],
        "shape": job["shape"],
        "receipt": receipt,
        "profile_result": {key: profile[key] for key in
                           ("file", "sha256", "logical_identity_sha256")},
        "native_result": {key: native[key] for key in
                           ("file", "sha256", "logical_identity_sha256")},
        "raw_dumps": raw_parts,
        "termination": termination,
        "planning_ir": raw_parts[-1]["summary_ir"],
        "annotations": annotation,
        "validation": {
            "profile_native_identity_equal": True,
            "three_lifecycle_dumps": True,
            "unique_final_measured_parent": True,
            "termination_zero_ir": True,
            **annotation["validation"],
        },
    }


def _missing(stage: str, plan: dict[str, Any]) -> list[str]:
    folder = HERE / stage
    if not folder.is_dir():
        return [stage]
    missing: list[str] = []
    for name in ("source-manifest.json", "binary-normal.json", "build-normal.receipt.json"):
        if not (folder / name).is_file():
            missing.append(f"{stage}/{name}")
    for job in expected_jobs(plan):
        for suffix in (".receipt.json", ".json", ".stdout", ".stderr", ".callgrind"):
            if not (folder / (job["name"] + suffix)).is_file():
                missing.append(f"{stage}/{job['name']}{suffix}")
        for part in NUMBERED_PARTS:
            if not (folder / f"{job['name']}.callgrind.{part}").is_file():
                missing.append(f"{stage}/{job['name']}.callgrind.{part}")
    return missing


def analyze_stage(stage: str, plan: dict[str, Any], create_annotations: bool) -> dict[str, Any]:
    missing = _missing(stage, plan)
    if missing:
        return {"stage": stage, "status": "pending", "missing_artifacts": missing}
    meta = validate_build(stage, plan)
    jobs = expected_jobs(plan)
    profiles = [validate_profile(stage, job, plan, meta, create_annotations)
                for job in jobs]
    return {
        "stage": stage,
        "status": "pass",
        "metadata": meta,
        "profiles": profiles,
        "profile_count": len(profiles),
        "validation": {
            "exact_source_binary_receipt_bindings": True,
            "profile_matrix_complete": True,
            "raw_dumps_complete": True,
        },
    }


def percent_reduction(baseline: int, final: int) -> float:
    require(baseline > 0, "baseline planning Ir must be positive")
    return (baseline - final) / baseline * 100.0


def native_admission() -> dict[str, Any]:
    path = HERE / "final-native-comparison.json"
    regular(path, "0531 native comparison")
    comparison = read_json(path)
    require(isinstance(comparison, dict), "0531 native comparison is not an object")
    require(comparison.get("plan_sha256") == sha(HERE / "plan.json"),
            "0531 native comparison/plan binding differs")
    require(comparison.get("status") == "pass",
            "0531 native comparison did not pass")
    require(comparison.get("admission_status") in ("eligible-for-conditional-lanes", "reject"),
            "0531 native admission did not authorize conditional lanes")
    admission = comparison.get("native_admission")
    require(isinstance(admission, dict) and isinstance(admission.get("passed"), bool),
            "0531 native admission is not passed")
    return {
        "file": relative(path),
        "sha256": sha(path),
        "admission_status": comparison["admission_status"],
        "native_rows": len(admission.get("rows", [])),
    }


def compare_stages(baseline: dict[str, Any], final: dict[str, Any],
                   plan: dict[str, Any]) -> dict[str, Any]:
    left = {(row["repeat"], row["shape"]): row for row in baseline["profiles"]}
    right = {(row["repeat"], row["shape"]): row for row in final["profiles"]}
    require(set(left) == set(right), "baseline/final profile matrices differ")
    required = float(plan["gates"]["planning_ir_reduction_percent"])
    rows: list[dict[str, Any]] = []
    for key in sorted(left):
        base, cand = left[key], right[key]
        reduction = percent_reduction(base["planning_ir"], cand["planning_ir"])
        base_mce = base["annotations"]["mce"]
        cand_mce = cand["annotations"]["mce"]
        rows.append({
            "repeat": key[0],
            "shape": key[1],
            "baseline_profile": base["name"],
            "candidate_profile": cand["name"],
            "planning_ir": {
                "baseline": base["planning_ir"],
                "final": cand["planning_ir"],
                "delta_ir": cand["planning_ir"] - base["planning_ir"],
                "reduction_percent": reduction,
                "required_reduction_percent": required,
                "passed": reduction >= required,
            },
            "mce_process_markup_compatibility_ir": {
                "baseline": base_mce["process_markup_compatibility_ir"],
                "final": cand_mce["process_markup_compatibility_ir"],
                "delta_ir": cand_mce["process_markup_compatibility_ir"]
                - base_mce["process_markup_compatibility_ir"],
            },
            "profile_identity_equal": (
                base["profile_result"]["logical_identity_sha256"]
                == cand["profile_result"]["logical_identity_sha256"]
            ),
        })
    return {
        "rows": rows,
        "required_reduction_percent": required,
        "passed": all(row["planning_ir"]["passed"] for row in rows),
        "decision": "instruction-gate-passed-diagnostic-only" if all(
            row["planning_ir"]["passed"] for row in rows
        ) else "reject",
        "scope": "Planning Ir is conditional Callgrind mechanism evidence; it is not latency or a whole-workflow speedup.",
    }


def render_markdown(document: dict[str, Any]) -> str:
    lines = [
        "# 0531 conditional planning profile",
        "",
        "This report validates the source-bound XLSX `edit_sheets` Callgrind lane.",
        "Planning Ir is conditional mechanism evidence and is not converted to latency.",
        "",
        f"Status: **{document['status']}**.",
        "",
    ]
    for stage in ("baseline", "final"):
        value = document.get("stages", {}).get(stage)
        if not isinstance(value, dict):
            continue
        lines.append(f"## {stage}")
        lines.append("")
        lines.append(f"Status: **{value.get('status')}**.")
        if value.get("status") == "pending":
            lines.append(f"Missing artifacts: `{len(value.get('missing_artifacts', []))}`.")
        else:
            lines.extend(("", "| Repeat | Shape | Planning Ir |", "| ---: | --- | ---: |"))
            for row in value.get("profiles", []):
                lines.append(f"| {row['repeat']} | {row['shape']} | {row['planning_ir']:,} |")
        lines.append("")
    comparison = document.get("comparison")
    if isinstance(comparison, dict):
        lines.extend(("## Comparison", "", "| Repeat | Shape | Baseline Ir | Candidate Ir | Reduction | Gate |", "| ---: | --- | ---: | ---: | ---: | ---: |"))
        for row in comparison.get("rows", []):
            metric = row["planning_ir"]
            lines.append(
                f"| {row['repeat']} | {row['shape']} | {metric['baseline']:,} | "
                f"{metric['final']:,} | {metric['reduction_percent']:.3f}% | "
                f"{'pass' if metric['passed'] else 'fail'} |"
            )
        lines.extend(("", f"Diagnostic profile decision: **{comparison.get('decision')}**."))
    lines.append("")
    return "\n".join(lines)


def analyze(stage_selection: str, create_annotations: bool) -> dict[str, Any]:
    plan = plan_data()
    selected = ("baseline",) if stage_selection == "baseline" else \
        ("final",) if stage_selection == "final" else ("baseline", "final")
    stages = {stage: analyze_stage(stage, plan, create_annotations) for stage in selected}
    document: dict[str, Any] = {
        "schema": "xlsx_mce_conditional_planning_profile_analysis_v1",
        "status": "pass" if all(value["status"] == "pass" for value in stages.values()) else "pending",
        "stage_selection": list(selected),
        "selected_function": OWNER,
        "plan": relative(HERE / "plan.json"),
        "plan_sha256": sha(HERE / "plan.json"),
        "planning_analyzer": {
            "path": str(PLANNING_PATH.relative_to(HERE.parent)),
            "sha256": sha(PLANNING_PATH),
        },
        "raw_helpers": {
            "path": str(PLANNING.HELPER.relative_to(HERE.parent)),
            "sha256": sha(PLANNING.HELPER),
        },
        "stages": stages,
        "scope": "Conditional XLSX planning profile for the shared OOXML MCE namespace-search pilot.",
        "limitations": [
            "The rebuilt candidate failed native admission. These completed profiles are diagnostic only and cannot authorize retention.",
            "Callgrind Ir is a mechanism diagnostic, not native latency, hardware cycles, or allocation counts.",
            "The four generated profiles cover two shapes and two repeats; no cold-cache, range, scaling, or native-producer claim follows.",
        ],
    }
    if set(stages) == {"baseline", "final"} \
            and all(value["status"] == "pass" for value in stages.values()):
        document["native_admission"] = native_admission()
        document["comparison"] = compare_stages(
            stages["baseline"], stages["final"], plan
        )
    return document


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("baseline", "final", "both"), default="both")
    parser.add_argument("--create-annotations", action="store_true",
                        help="write deterministic annotations for each final dump")
    parser.add_argument("--output", type=Path,
                        help="JSON output path; defaults to profile-comparison.json")
    parser.add_argument("--markdown-output", type=Path,
                        help="Markdown output path; defaults beside the JSON output")
    args = parser.parse_args()
    output = args.output or HERE / "final-profile-comparison.json"
    markdown = args.markdown_output or output.with_suffix(".md")
    try:
        document = analyze(args.stage, args.create_annotations)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(document, indent=2, sort_keys=True) + "\n",
                          encoding="utf-8")
        markdown.parent.mkdir(parents=True, exist_ok=True)
        markdown.write_text(render_markdown(document), encoding="utf-8")
    except (EvidenceError, OSError, json.JSONDecodeError, ValueError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0531 conditional profile evidence {document['status']}: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
