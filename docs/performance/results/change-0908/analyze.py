#!/usr/bin/env python3
"""Offline reducer for the 0908 PPTX multi-slide Callgrind capture.

The retained 0845 reader owns the compressed Callgrind grammar.  This file is
only the packet adapter: it binds the frozen binary/runner/owner contract,
admits the first child, dynamically records whether that child produced one
or two scoped dumps, and then applies the same qualification to all four
children.  It never starts a workload, a build, or a profiler.

The output is an instruction-count report.  ``Ir`` is kept as an attribution
counter; it is deliberately never converted into native elapsed time,
hardware cycles, allocations, or a speedup.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import shlex
import sys
from typing import Any, Mapping, Sequence


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
READER_PATH = HERE.parent / "change-0845" / "reader_v2.py"

OWNER = "litchi_perf_baseline::publish_pptx_multi_slide_batch_edit"
MEASURED_PARENT = "litchi_perf_baseline::run_pptx_multi_slide_batch_edit_save"
MEASURED_CALLSITE = "51174"
EXPECTED_CALLSITE = "51143"
EXPECTED_BINARY_SHA256 = "e0a7707dcd849c92e01c8dc9115d61414fba10584bb8415e442f4c225828af84"
EXPECTED_CPU = 12
LABELS = ("eager-a", "source-a", "source-b", "eager-b")
KINDS = {
    "eager-a": "eager",
    "source-a": "source_backed",
    "source-b": "source_backed",
    "eager-b": "eager",
}
EXPECTED_OPTIONS = (
    "--tool=callgrind",
    "--collect-atstart=no",
    "--toggle-collect=" + OWNER,
)
ACCEPTED_ZERO_OPTIONS = ("--zero-before=" + OWNER,)
ACCEPTED_DUMP_OPTIONS = ("--dump-after=" + OWNER,)
ACCEPTED_TRIGGERS = frozenset(ACCEPTED_DUMP_OPTIONS)

HARNESS_NAMESPACE = "litchi_perf_baseline::"
PRODUCTION_NAMESPACES = (
    "litchi_pptx::",
    "litchi_opc::",
    "litchi_iwa::",
    "litchi_xlsx::",
    "litchi_docx::",
    "litchi_presentation::",
    "soapberry_zip::",
    "xml_minifier::",
    "zlib_rs::",
)
RUNTIME_NAMESPACES = (
    "std::",
    "core::",
    "alloc::",
    "__rust_",
    "rustc_",
    "malloc",
    "calloc",
    "realloc",
    "free",
    "memcpy",
    "memmove",
    "memset",
)


class EvidenceError(ValueError):
    """A retained artifact is missing, malformed, or contradictory."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def packet_relative(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(HERE.resolve()))
    except ValueError as error:
        raise EvidenceError(f"path escapes 0908 packet: {path}") from error


def repo_relative(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(REPO.resolve()))
    except ValueError as error:
        raise EvidenceError(f"path escapes repository: {path}") from error


def load_reader() -> Any:
    require(READER_PATH.is_file(), f"missing retained Callgrind reader: {READER_PATH}")
    spec = importlib.util.spec_from_file_location("retained_callgrind_reader_0845_0908", READER_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load retained Callgrind reader: {READER_PATH}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


READER = load_reader()


def descriptor(path: Path) -> dict[str, Any]:
    return {"file": packet_relative(path), "bytes": path.stat().st_size,
            "sha256": sha256(path)}


def checked_descriptor(value: Any, label: str, root: Path | None = None) -> dict[str, Any]:
    """Verify a retained absolute-path descriptor and return its actual value."""

    require(isinstance(value, Mapping) and isinstance(value.get("path"), str),
            f"{label} is not a path descriptor")
    raw_path = Path(value["path"])
    require(not raw_path.is_symlink(), f"{label} path is a symlink: {raw_path}")
    path = raw_path.resolve()
    require(path.is_file() and not path.is_symlink(), f"{label} path is not a regular file: {path}")
    if root is not None:
        try:
            path.relative_to(root.resolve())
        except ValueError as error:
            raise EvidenceError(f"{label} escapes {root}: {path}") from error
    actual = {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}
    require(value.get("bytes") == actual["bytes"] and value.get("sha256") == actual["sha256"],
            f"{label} descriptor does not match {path}")
    return actual


def require_descriptor_for(value: Any, path: Path, label: str) -> dict[str, Any]:
    actual = checked_descriptor(value, label)
    require(Path(actual["path"]).resolve() == path.resolve(),
            f"{label} path differs: {actual['path']} != {path}")
    return actual


def validate_plan_custody(plan: Mapping[str, Any], path: Path) -> None:
    """Bind every frozen plan/helper/input/tool descriptor to current bytes."""

    if isinstance(plan.get("binary"), Mapping):
        checked_descriptor(plan["binary"], f"{packet_relative(path)}.binary")
    for key in ("helper", "initial_helper", "initial_plan"):
        value = plan.get(key)
        if value is not None:
            checked_descriptor(value, f"{packet_relative(path)}.{key}")
    for key in ("inputs", "tools"):
        value = plan.get(key)
        if value is None:
            continue
        require(isinstance(value, list), f"{packet_relative(path)}.{key} is not a list")
        for index, item in enumerate(value):
            checked_descriptor(item, f"{packet_relative(path)}.{key}[{index}]")


def _plan_path() -> Path:
    # The v2 semantic wrapper is the current plan for later children.  The
    # original plan remains retained because eager-a was captured under it;
    # both descriptors are deliberately accepted by the admission binding.
    candidates = (HERE / "plan-v2.json", HERE / "plan.json", HERE / "capture-plan.json")
    existing = [path for path in candidates if path.is_file() and not path.is_symlink()]
    require(existing, "0908 requires plan-v2.json, plan.json, or capture-plan.json")
    return existing[0]


def _binary_record(plan: Mapping[str, Any]) -> Mapping[str, Any]:
    value = plan.get("binary")
    if isinstance(value, str):
        return {"path": value}
    require(isinstance(value, Mapping), "plan.binary is missing or malformed")
    return value


def _plan_rows(plan: Mapping[str, Any]) -> list[dict[str, Any]]:
    raw = plan.get("legs")
    if raw is None:
        raw = plan.get("children")
    if raw is None:
        raw = plan.get("schedule")
    require(isinstance(raw, list) and raw, "plan has no capture legs")
    rows: list[dict[str, Any]] = []
    for index, item in enumerate(raw):
        if isinstance(item, (list, tuple)):
            require(len(item) == 2 and all(isinstance(part, str) for part in item),
                    f"plan leg {index} is malformed")
            label, kind = item
            row = {"index": index, "label": label, "kind": kind}
        else:
            require(isinstance(item, Mapping), f"plan leg {index} is malformed")
            label = item.get("label", item.get("name", item.get("leg")))
            kind = item.get("kind", item.get("mode"))
            require(isinstance(label, str) and isinstance(kind, str),
                    f"plan leg {index} lacks label/kind")
            row = dict(item)
            row.update(index=index, label=label, kind=kind)
        rows.append(row)
    require([(row["label"], row["kind"]) for row in rows]
            == [(label, KINDS[label]) for label in LABELS],
            "plan leg order or eager/source classification differs")
    return rows


def _option_list(plan: Mapping[str, Any]) -> list[str]:
    value = plan.get("options", plan.get("callgrind_options", plan.get("valgrind_options")))
    if value is None:
        return list(EXPECTED_OPTIONS + ACCEPTED_ZERO_OPTIONS + ACCEPTED_DUMP_OPTIONS)
    require(isinstance(value, list) and all(isinstance(item, str) for item in value),
            "plan Callgrind options are malformed")
    return list(value)


def validate_plan() -> dict[str, Any]:
    path = _plan_path()
    raw = read_json(path)
    require(isinstance(raw, Mapping), "0908 plan is not an object")
    schema = raw.get("schema")
    require((isinstance(schema, str) and "0908" in schema)
            or raw.get("plan_id") == "0908",
            "plan is not the 0908 PPTX Callgrind plan")
    validate_plan_custody(raw, path)
    # Both descriptors are retained in the actual packet: eager-a was started
    # with the original wrapper and later children with the v2 wrapper.  Check
    # the original plan independently so a v2 receipt cannot hide drift in it.
    for related in (HERE / "plan.json", HERE / "plan-v2.json"):
        if related.is_file() and not related.is_symlink() and related != path:
            related_plan = read_json(related)
            require(isinstance(related_plan, Mapping), f"{packet_relative(related)} is not an object")
            validate_plan_custody(related_plan, related)

    rows = _plan_rows(raw)
    binary = _binary_record(raw)
    binary_sha = binary.get("sha256", binary.get("binary_sha256"))
    require(binary_sha == EXPECTED_BINARY_SHA256,
            "plan binary SHA-256 differs from frozen 0907 native binary")
    if "bytes" in binary:
        require(type(binary["bytes"]) is int and binary["bytes"] > 0,
                "plan binary byte count is invalid")
    if isinstance(binary.get("path"), str):
        binary_path = Path(binary["path"])
        if binary_path.is_file() and not binary_path.is_symlink():
            require(sha256(binary_path) == EXPECTED_BINARY_SHA256,
                    "available plan binary differs from frozen native binary")

    cpu = raw.get("cpu", raw.get("cpu_affinity"))
    require(cpu == EXPECTED_CPU, "plan CPU differs from fixed CPU 12")
    warmup = raw.get("warmup", raw.get("warmup_iterations"))
    samples = raw.get("samples_per_leg", raw.get("samples"))
    require(warmup == 0 and samples == 1,
            "0908 Callgrind plan must use warmup=0 and samples=1")

    owner = raw.get("owner", OWNER)
    parent = raw.get("measured_parent", raw.get("parent", MEASURED_PARENT))
    require(owner == OWNER and parent == MEASURED_PARENT,
            "plan owner or measured parent differs")
    options = _option_list(raw)
    require(all(option in options for option in EXPECTED_OPTIONS),
            "plan omits a required Callgrind collection option")
    require(any(option in options for option in ACCEPTED_ZERO_OPTIONS),
            "plan omits --zero-before=<exact owner>")
    require(any(option in options for option in ACCEPTED_DUMP_OPTIONS),
            "plan omits --dump-after=<exact owner>")

    normalized = dict(raw)
    normalized.update(
        _plan_path=path,
        _rows=rows,
        _binary=binary,
        _options=options,
        _plan_sha256=sha256(path),
        _reader_sha256=sha256(READER_PATH),
    )
    return normalized


def _row_capture_path(row: Mapping[str, Any]) -> Path | None:
    value = row.get("capture", row.get("capture_dir", row.get("path")))
    if not isinstance(value, str):
        return None
    path = Path(value)
    if not path.is_absolute():
        path = HERE / path
    path = path.resolve()
    try:
        path.relative_to(HERE.resolve())
    except ValueError as error:
        raise EvidenceError(f"capture path escapes packet: {path}") from error
    return path


def locate_capture(row: Mapping[str, Any]) -> Path:
    label = str(row["label"])
    explicit = _row_capture_path(row)
    if explicit is not None:
        require(explicit.is_dir() and not explicit.is_symlink(),
                f"missing capture directory for {label}: {explicit}")
        return explicit

    candidates: list[Path] = []
    for root in (HERE / "capture", HERE):
        if not root.is_dir():
            continue
        for path in root.iterdir():
            if not path.is_dir() or path.is_symlink():
                continue
            if path.name == label or path.name.endswith("-" + label):
                if any(item.is_file() and item.name.startswith("callgrind.out.")
                       and item.name[len("callgrind.out."):].isdigit()
                       for item in path.iterdir()):
                    candidates.append(path.resolve())
    unique = sorted(set(candidates))
    require(len(unique) == 1,
            f"expected one capture directory for {label}, found {unique}")
    return unique[0]


def numbered_paths(folder: Path) -> list[Path]:
    paths = [path for path in folder.iterdir()
             if path.is_file() and not path.is_symlink()
             and path.name.startswith("callgrind.out.")
             and path.name[len("callgrind.out."):].isdigit()]
    paths.sort(key=lambda path: int(path.name.rsplit(".", 1)[1]))
    require(paths, f"{packet_relative(folder)} has no numbered Callgrind dump")
    numbers = [int(path.name.rsplit(".", 1)[1]) for path in paths]
    require(numbers == list(range(1, len(numbers) + 1)),
            f"{packet_relative(folder)} numbered dumps are not contiguous: {numbers}")
    require(len(paths) in (1, 2),
            f"{packet_relative(folder)} has {len(paths)} scoped dumps; expected dynamically qualified 1 or 2")
    return paths


def _header_command(parsed: Mapping[str, Any], label: str) -> str:
    value = parsed["header"].get("command")
    require(isinstance(value, str) and value, f"{label}: Callgrind command header is missing")
    return value


def _command_binary_token(command: str, planned_path: str | None) -> str | None:
    try:
        tokens = shlex.split(command)
    except ValueError as error:
        raise EvidenceError(f"invalid Callgrind command header: {error}") from error
    planned_name = Path(planned_path).name if planned_path else None
    for token in tokens:
        if planned_name and Path(token).name == planned_name:
            return token
        if Path(token).name in {"litchi-perf-baseline", "native"}:
            return token
    return None


def _expected_wrapper(label: str) -> tuple[Path, Path]:
    if label == "eager-a":
        return HERE / "capture.py", HERE / "plan.json"
    return HERE / "capture_v2.py", HERE / "plan-v2.json"


def _expected_case(row: Mapping[str, Any]) -> str:
    expected = f"pptx_{row['kind']}_multi_slide_batch_edit_save"
    require(row.get("case", expected) == expected,
            f"{row['label']}: plan case differs from its kind")
    return expected


def _expected_argv(plan: Mapping[str, Any], row: Mapping[str, Any], folder: Path) -> list[str]:
    binary = _binary_record(plan)
    require(isinstance(binary.get("path"), str), "plan binary path is missing")
    return [
        "/usr/bin/taskset", "-c", str(EXPECTED_CPU), "/usr/bin/valgrind",
        *plan["_options"],
        "--callgrind-out-file=" + str((folder / "callgrind.out").resolve()),
        str(Path(binary["path"]).resolve()),
        "--workers", "1", "--warmup", "0", "--samples", "1",
        "--case", _expected_case(row),
        "--json", str((folder / "report.json").resolve()),
    ]


def _validate_receipt_artifacts(command: Mapping[str, Any], folder: Path, label: str) -> dict[str, Any]:
    raw = command.get("artifacts")
    require(isinstance(raw, list) and raw, f"{label}: command artifact list is missing")
    artifacts: dict[str, dict[str, Any]] = {}
    for index, item in enumerate(raw):
        actual = checked_descriptor(item, f"{label}.command.artifacts[{index}]", folder)
        name = Path(actual["path"]).name
        require(name not in artifacts, f"{label}: duplicate command artifact {name}")
        artifacts[name] = actual
    required = {"callgrind.out", "report.json", "start.json", "stdout.txt", "stderr.txt"}
    require(required <= artifacts.keys(),
            f"{label}: command artifacts omit {sorted(required - artifacts.keys())}")
    scoped = {name for name in artifacts if name.startswith("callgrind.out.")
              and name[len("callgrind.out."):].isdigit()}
    require(scoped, f"{label}: command artifacts omit numbered Callgrind output")
    return artifacts


def _one_native_vector(value: Any, label: str) -> list[Any]:
    require(isinstance(value, list) and len(value) == 200,
            f"{label}: native reference does not contain 200 values")
    require(all(item == value[0] for item in value),
            f"{label}: native reference vector is not constant and cannot be reduced to one sample")
    return [value[0]]


def _collapse_native_metrics(value: Any, label: str = "operation_metrics") -> Any:
    require(isinstance(value, Mapping), f"{label}: native operation metrics are not an object")
    result: dict[str, Any] = {}
    for key, child in value.items():
        child_label = f"{label}.{key}"
        if key == "sample_count":
            require(child == 200, f"{child_label}: native count is not 200")
            result[key] = 1
        elif key == "sample_indices":
            require(isinstance(child, list) and sorted(child) == list(range(200)),
                    f"{child_label}: native indices are not a permutation of 0..199")
            result[key] = [0]
        elif key == "values":
            result[key] = _one_native_vector(child, child_label)
        elif isinstance(child, Mapping):
            result[key] = _collapse_native_metrics(child, child_label)
        else:
            result[key] = child
    return result


def _validate_semantics(folder: Path, plan: Mapping[str, Any], row: Mapping[str, Any],
                        command: Mapping[str, Any]) -> dict[str, Any]:
    """Recheck v2's semantic receipt and its 1-vs-200 native reference binding."""

    label = str(row["label"])
    semantic_path = folder / "semantics.json"
    semantic = read_json(semantic_path)
    require(isinstance(semantic, Mapping) and semantic.get("status") == "pass",
            f"{label}: semantic receipt is not pass")
    case = _expected_case(row)
    require(semantic.get("case") == case, f"{label}: semantic case differs")

    reference_descriptor = semantic.get("reference")
    reference_actual = checked_descriptor(reference_descriptor, f"{label}.semantics.reference")
    reference_path = Path(reference_actual["path"])
    require(reference_path.name == label + ".json"
            and reference_path.parent.resolve() == (HERE.parent / "change-0907").resolve(),
            f"{label}: semantic reference is not the matching retained 0907 native report")
    reference = read_json(reference_path)
    require(isinstance(reference, Mapping), f"{label}: native reference is not an object")
    reference_results = reference.get("results")
    require(isinstance(reference_results, list) and len(reference_results) == 1,
            f"{label}: native reference result cardinality differs")
    expected = reference_results[0]
    require(isinstance(expected, Mapping) and expected.get("case") == case,
            f"{label}: native reference case differs")

    report_path = folder / "report.json"
    report = read_json(report_path)
    require(isinstance(report, Mapping) and report.get("schema_version") == 1,
            f"{label}: workload report schema differs")
    identity = report.get("binary_identity")
    require(isinstance(identity, Mapping)
            and identity.get("binary_sha256") == EXPECTED_BINARY_SHA256,
            f"{label}: report binary identity differs")
    binary = _binary_record(plan)
    require(identity.get("binary_bytes") == binary.get("bytes"),
            f"{label}: report binary byte count differs")
    tool = report.get("tool")
    require(isinstance(tool, Mapping)
            and tool.get("profile") == "release"
            and tool.get("instrumentation") == "none",
            f"{label}: report tool identity differs")
    environment = report.get("environment")
    require(isinstance(environment, Mapping)
            and environment.get("cpu_affinity") == str(EXPECTED_CPU)
            and all(environment.get(key) is None
                    for key in ("git_revision", "git_worktree_dirty", "rustc_version")),
            f"{label}: report environment binding differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, Mapping)
            and configuration.get("samples_per_case") == 1
            and configuration.get("warmup_iterations_per_case") == 0
            and configuration.get("cases") == [case],
            f"{label}: report sample or case configuration differs")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}: workload result cardinality differs")
    result = results[0]
    require(isinstance(result, Mapping) and result.get("case") == case,
            f"{label}: workload result case differs")

    for key in ("corpus", "sink", "output_sha256"):
        require(result.get(key) == expected.get(key),
                f"{label}: report {key} differs from native reference")
        require(semantic.get(key) == expected.get(key),
                f"{label}: semantic {key} differs from native reference")

    expected_source = expected.get("source")
    report_source = result.get("source")
    semantic_source = semantic.get("source")
    require(isinstance(expected_source, Mapping)
            and isinstance(report_source, Mapping)
            and isinstance(semantic_source, Mapping)
            and set(report_source) == set(expected_source)
            and set(semantic_source) == set(expected_source),
            f"{label}: source counter fields differ")
    for key, native_values in expected_source.items():
        reduced = _one_native_vector(native_values, f"{label}.native.source.{key}")
        require(report_source[key] == reduced and semantic_source[key] == reduced,
                f"{label}: source counter {key} is not the validated one-sample reduction")

    native_metrics = expected.get("operation_metrics")
    measured_metrics = result.get("operation_metrics")
    require(measured_metrics == _collapse_native_metrics(native_metrics,
                                                         f"{label}.native.operation_metrics"),
            f"{label}: operation_metrics is not the validated 1-sample reduction of native 200-sample values")
    require(measured_metrics.get("sample_count") == 1
            and measured_metrics.get("sample_indices") == [0],
            f"{label}: operation_metrics sample binding differs")

    semantic_reference = semantic.get("reference")
    require(semantic_reference == reference_actual,
            f"{label}: semantic reference descriptor was not retained exactly")
    if label == "eager-a":
        require_descriptor_for(semantic.get("original_successful_command"),
                               folder / "command.json",
                               f"{label}.semantics.original_successful_command")
        require_descriptor_for(semantic.get("original_failure"),
                               folder / "semantic-failure.json",
                               f"{label}.semantics.original_failure")
    return {
        "file": packet_relative(semantic_path),
        "sha256": sha256(semantic_path),
        "status": "pass",
        "reference": reference_actual,
        "native_sample_count": 200,
        "captured_sample_count": 1,
        "native_vectors_constant": True,
        "operation_metrics_reduced": True,
    }


def validate_capture_metadata(folder: Path, plan: Mapping[str, Any], row: Mapping[str, Any]) -> dict[str, Any]:
    """Validate the actual capture.py/capture_v2.py receipt contract."""

    label = str(row["label"])
    start_path = folder / "start.json"
    command_path = folder / "command.json"
    start = read_json(start_path)
    command = read_json(command_path)
    require(isinstance(start, Mapping) and isinstance(command, Mapping),
            f"{label}: start/command receipts are malformed")
    require(command.get("exit_code") == 0, f"{label}: capture child did not exit zero")
    expected_argv = _expected_argv(plan, row, folder)
    require(start.get("argv") == expected_argv and command.get("argv") == expected_argv,
            f"{label}: start/command argv differs from frozen capture command")
    require(start.get("argv") == command.get("argv"), f"{label}: start and command argv differ")
    require(start.get("cwd") == plan.get("cwd") and command.get("cwd") == plan.get("cwd"),
            f"{label}: start/command cwd differs from plan")
    environment = plan.get("environment")
    require(start.get("environment") == environment and command.get("environment") == environment,
            f"{label}: start/command environment differs from plan")
    require(start.get("row") == dict(row) and command.get("row") == dict(row),
            f"{label}: start/command row differs from plan")

    binary = _binary_record(plan)
    binary_path = Path(str(binary["path"])).resolve()
    for receipt_name, receipt in (("start", start), ("command", command)):
        require_descriptor_for(receipt.get("binary"), binary_path, f"{label}.{receipt_name}.binary")
    helper_path, receipt_plan_path = _expected_wrapper(label)
    require_descriptor_for(start.get("helper"), helper_path, f"{label}.start.helper")
    require_descriptor_for(command.get("helper"), helper_path, f"{label}.command.helper")
    require_descriptor_for(start.get("plan"), receipt_plan_path, f"{label}.start.plan")
    require_descriptor_for(command.get("plan"), receipt_plan_path, f"{label}.command.plan")

    artifacts = _validate_receipt_artifacts(command, folder, label)
    semantic = _validate_semantics(folder, plan, row, command)
    original_failure: dict[str, Any] | None = None
    failure_path = folder / "semantic-failure.json"
    if failure_path.is_file() and not failure_path.is_symlink():
        require(label == "eager-a", f"{label}: unexpected semantic-failure.json")
        failure = read_json(failure_path)
        require(isinstance(failure, Mapping) and failure.get("status") == "fail",
                f"{label}: retained original semantic failure is malformed")
        original_failure = descriptor(failure_path)
    return {
        "start": descriptor(start_path),
        "command": descriptor(command_path),
        "artifacts": artifacts,
        "semantic": semantic,
        "original_wrapper_failure": original_failure,
        "validation": {
            "start_command_argv_identical": True,
            "fixed_cpu_and_workload_arguments": True,
            "binary_helper_plan_bindings_exact": True,
            "all_command_artifacts_hashed": True,
            "semantic_status_pass": True,
            "native_200_to_capture_1_reduction_verified": True,
        },
    }


def _owner_row(parsed: Mapping[str, Any], label: str) -> dict[str, Any]:
    names = parsed["names"]
    owner_ids = sorted(function_id for function_id, name in names.items() if name == OWNER)
    require(len(owner_ids) == 1,
            f"{label}: owner resolved to {owner_ids}; one target-owner symbol is required")
    owner_id = owner_ids[0]
    incoming = READER.incoming_edges(parsed, owner_id)
    require(len(incoming) == 1,
            f"{label}: owner has ambiguous positive incoming edges: {incoming}")
    edge = incoming[0]
    require(edge["caller"] == MEASURED_PARENT,
            f"{label}: owner caller is {edge['caller']!r}, expected {MEASURED_PARENT!r}")
    # The expected-output helper call at line 51143 is outside the retained
    # scoped owner edge in this capture.  The raw edge must be the measured
    # call at 51174; accepting the caller name alone would make that boundary
    # too weak if the harness is rearranged later.
    require(str(edge["position"]) == MEASURED_CALLSITE,
            f"{label}: owner edge position is {edge['position']!r}; expected measured callsite {MEASURED_CALLSITE} (expected-output callsite {EXPECTED_CALLSITE} is excluded)")
    require(edge["calls"] == 1,
            f"{label}: owner incoming call count is {edge['calls']}, expected one")
    summary = parsed["header"]["summary_ir"]
    require(summary > 0 and parsed["header"]["totals_ir"] == summary,
            f"{label}: scoped Ir totals do not reconstruct the summary")
    require(edge["inclusive_ir"] == summary,
            f"{label}: owner incoming Ir differs from scoped summary")
    require(owner_id in parsed["functions"], f"{label}: owner function record is missing")
    view = READER.function_view(parsed["functions"][owner_id], names)
    require(view["self_ir"] + view["direct_children_ir"] == summary,
            f"{label}: owner self plus direct callees does not conserve Ir")
    all_self = sum(function["self_ir"] for function in parsed["functions"].values())
    require(all_self == summary,
            f"{label}: all function self costs do not conserve scoped Ir")
    return {
        "owner_id": owner_id,
        "incoming": incoming,
        "measured_callsite": MEASURED_CALLSITE,
        "excluded_expected_callsite": EXPECTED_CALLSITE,
        "function": view,
        "summary_ir": summary,
        "all_function_self_ir": all_self,
        "validation": {
            "one_owner_symbol": True,
            "one_positive_owner_edge": True,
            "owner_parent_matches": True,
            "one_owner_call": True,
            "owner_inclusive_matches_summary": True,
            "owner_self_plus_direct_conserves": True,
            "all_function_self_conserves": True,
        },
    }


def _source_class(name: str, source: str) -> dict[str, str]:
    normalized = source.replace("\\", "/")
    if name.startswith(HARNESS_NAMESPACE) or "/tools/perf-baseline/" in normalized:
        return {
            "class": "harness_guard",
            "confidence": "high",
            "basis": "perf-baseline namespace or source path",
            "source_caveat": "This row belongs to the measurement harness/guard path, not a production crate.",
        }
    if "/crates/" in normalized and "/tools/" not in normalized:
        return {
            "class": "production",
            "confidence": "high",
            "basis": "Callgrind source attribution under crates/",
            "source_caveat": "The source path identifies a production crate, but Ir remains inclusive at this call edge.",
        }
    if name.startswith(PRODUCTION_NAMESPACES):
        return {
            "class": "production_candidate",
            "confidence": "medium",
            "basis": "known production namespace without a crates/ source path",
            "source_caveat": "Namespace suggests production code; source attribution was not under crates/, so this is not a high-confidence source proof.",
        }
    if name.startswith(RUNTIME_NAMESPACES) or "/rustc/" in normalized or normalized.startswith("<"):
        return {
            "class": "runtime_or_dependency",
            "confidence": "medium",
            "basis": "runtime/toolchain namespace or source",
            "source_caveat": "This row is retained for conservation and is excluded from production totals.",
        }
    return {
        "class": "unclassified",
        "confidence": "low",
        "basis": "no production or harness evidence matched",
        "source_caveat": "Do not treat this row as production work without an independent source review.",
    }


def direct_callee_rows(parsed: Mapping[str, Any], owner: Mapping[str, Any]) -> dict[str, Any]:
    function_rows = parsed["functions"]
    names = parsed["names"]
    rows: list[dict[str, Any]] = []
    category_ir: dict[str, int] = {}
    for edge in owner["function"]["direct_children"]:
        callee_id = edge["callee_id"]
        name = str(edge["callee"])
        callee = function_rows.get(callee_id)
        source = str(callee.get("file", "")) if callee is not None else ""
        classification = _source_class(name, source)
        row = {
            "callee_id": callee_id,
            "name": name,
            "source": source,
            "calls": edge["calls"],
            "edge_count": edge["edge_count"],
            "inclusive_ir": edge["inclusive_ir"],
            "callee_self_ir": callee["self_ir"] if callee is not None else None,
            "callee_records": callee["records"] if callee is not None else None,
            "classification": classification,
        }
        rows.append(row)
        category = classification["class"]
        category_ir[category] = category_ir.get(category, 0) + edge["inclusive_ir"]
    require(sum(row["inclusive_ir"] for row in rows) == owner["function"]["direct_children_ir"],
            "direct callee rows do not conserve owner direct Ir")
    rows.sort(key=lambda row: (-row["inclusive_ir"], row["name"], row["callee_id"]))
    return {
        "rows": rows,
        "category_inclusive_ir": dict(sorted(category_ir.items())),
        "direct_children_ir": owner["function"]["direct_children_ir"],
        "classification_policy": {
            "production": "high confidence only when the reader source path is under crates/",
            "production_candidate": "known production namespace with weaker source evidence",
            "harness_guard": "perf-baseline namespace or tools/perf-baseline source path",
            "runtime_or_dependency": "toolchain/runtime attribution retained but excluded from production totals",
            "unclassified": "never promoted to production without independent source review",
        },
    }


def parse_part(path: Path, plan: Mapping[str, Any], label: str) -> dict[str, Any]:
    parsed = READER.parse_raw(path)
    header = parsed["header"]
    expected_part = int(path.name.rsplit(".", 1)[1])
    require(header["part"] == expected_part,
            f"{packet_relative(path)} part differs from filename")
    require(header["trigger"] in ACCEPTED_TRIGGERS,
            f"{packet_relative(path)} trigger {header['trigger']!r} is not the scoped dump trigger")
    command = _header_command(parsed, packet_relative(path))
    owner = _owner_row(parsed, packet_relative(path))
    owner["path"] = packet_relative(path)
    owner["sha256"] = sha256(path)
    owner["part"] = expected_part
    owner["trigger"] = header["trigger"]
    owner["command"] = command
    return {
        "file": packet_relative(path),
        "sha256": owner["sha256"],
        "bytes": path.stat().st_size,
        "header": header,
        "owner": owner,
        "direct_callees": direct_callee_rows(parsed, owner),
        "_parsed": parsed,
    }


def parse_terminal(path: Path, last_part: int, command: str, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(),
            f"{label}: missing terminal Callgrind dump {path}")
    parsed = READER.parse_raw(path, allow_empty=True)
    header = parsed["header"]
    require(header["part"] == last_part + 1,
            f"{packet_relative(path)} terminal part does not follow scoped dumps")
    require(header["trigger"] == "Program termination",
            f"{packet_relative(path)} terminal trigger differs")
    require(header["summary_ir"] == 0 and header["totals_ir"] == 0 and not parsed["functions"],
            f"{packet_relative(path)} terminal dump is not empty")
    require(_header_command(parsed, packet_relative(path)) == command,
            f"{packet_relative(path)} terminal command differs from scoped dump")
    return {
        "file": packet_relative(path),
        "sha256": sha256(path),
        "bytes": path.stat().st_size,
        "header": header,
    }


def analyze_capture(plan: Mapping[str, Any], row: Mapping[str, Any], expected_count: int | None) -> dict[str, Any]:
    label = str(row["label"])
    folder = locate_capture(row)
    metadata = validate_capture_metadata(folder, plan, row)
    numbered = numbered_paths(folder)
    if expected_count is not None:
        require(len(numbered) == expected_count,
                f"{label}: scoped dump count {len(numbered)} differs from first-child qualification {expected_count}")
    parts = [parse_part(path, plan, label) for path in numbered]
    command = parts[0]["owner"]["command"]
    require(all(part["owner"]["command"] == command for part in parts),
            f"{label}: scoped dump commands differ")
    binary = _binary_record(plan)
    planned_path = binary.get("path")
    binary_token = _command_binary_token(command, planned_path if isinstance(planned_path, str) else None)
    require(binary_token is not None,
            f"{label}: Callgrind command does not identify litchi-perf-baseline")
    if isinstance(planned_path, str):
        require(Path(planned_path).name == Path(binary_token).name,
                f"{label}: Callgrind command binary differs from plan")
    terminal = parse_terminal(folder / "callgrind.out", len(numbered), command, label)
    return {
        "label": label,
        "kind": row["kind"],
        "capture_directory": packet_relative(folder),
        "metadata": metadata,
        "binary_command_token": binary_token,
        "numbered_dump_count": len(numbered),
        "parts": [{key: value for key, value in part.items() if key != "_parsed"}
                  for part in parts],
        "terminal": terminal,
        "validation": {
            "scoped_dump_count_is_one_or_two": True,
            "all_scoped_triggers_match": True,
            "one_owner_per_scoped_dump": True,
            "one_measured_parent_per_owner": True,
            "owner_and_direct_callee_ir_conserved": True,
            "terminal_dump_is_empty_and_retained": True,
        },
    }


def _admission_path() -> Path:
    return HERE / "first-child-admission.json"


def _admission_plan_binding(value: Any) -> str:
    if isinstance(value, Mapping):
        actual = checked_descriptor(value, "first-child-admission.plan")
        path = Path(actual["path"]).resolve()
    elif isinstance(value, str):
        path = Path(value).resolve()
        require(path.is_file() and not path.is_symlink(),
                f"first-child-admission plan is missing: {path}")
        actual = {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}
    else:
        raise EvidenceError("first-child-admission.plan is missing")
    accepted = {path.resolve() for path in
                (HERE / "plan.json", HERE / "plan-v2.json", HERE / "capture-plan.json")
                if path.is_file() and not path.is_symlink()}
    require(path in accepted, "first-child admission is bound to no retained 0908 plan")
    return actual["sha256"]


def _validate_first_admission(prior: Mapping[str, Any], plan: Mapping[str, Any]) -> int:
    """Validate the root admission receipt before inferring its one dump."""

    require(prior.get("status") == "pass", "first-child admission is not pass")
    _admission_plan_binding(prior.get("plan"))
    raw_descriptor = prior.get("raw_dump")
    raw_actual = checked_descriptor(raw_descriptor, "first-child-admission.raw_dump", HERE)
    raw_path = Path(raw_actual["path"])
    folder = HERE / "capture" / "eager-a"
    require(raw_path.parent.resolve() == folder.resolve()
            and raw_path.name == "callgrind.out.1",
            "first-child raw dump is not capture/eager-a/callgrind.out.1")
    numbered = numbered_paths(folder)
    require(len(numbered) == 1 and numbered[0].resolve() == raw_path,
            "first-child admission raw dump is not the sole retained scoped dump")

    parsed = READER.parse_raw(raw_path)
    header = parsed["header"]
    require(header["part"] == 1 and header["summary_ir"] == header["totals_ir"] > 0,
            "first-child admission raw header is not a positive conserved part 1")
    require(header["trigger"] in ACCEPTED_TRIGGERS,
            "first-child admission raw trigger differs")
    owner = _owner_row(parsed, "first-child-admission.raw_dump")
    edge = owner["incoming"][0]
    receipt_edge = prior.get("measured_edge")
    require(isinstance(receipt_edge, Mapping), "first-child admission measured_edge is missing")
    for key in ("caller", "callee", "calls", "inclusive_ir", "position"):
        require(receipt_edge.get(key) == edge.get(key),
                f"first-child admission measured_edge.{key} differs from raw")
    receipt_header = prior.get("header")
    require(isinstance(receipt_header, Mapping), "first-child admission header is missing")
    for key in ("part", "events", "summary_ir", "totals_ir", "trigger", "command"):
        require(receipt_header.get(key) == header.get(key),
                f"first-child admission header.{key} differs from raw")

    terminal_descriptor = checked_descriptor(prior.get("terminal"),
                                             "first-child-admission.terminal", HERE)
    terminal_path = Path(terminal_descriptor["path"])
    require(terminal_path.resolve() == (folder / "callgrind.out").resolve(),
            "first-child admission terminal path differs")
    parse_terminal(terminal_path, 1, header["command"], "first-child-admission")
    require_descriptor_for(prior.get("original_command"), folder / "command.json",
                           "first-child-admission.original_command")
    require_descriptor_for(prior.get("corrected_semantics"), folder / "semantics.json",
                           "first-child-admission.corrected_semantics")
    require_descriptor_for(prior.get("reader"), READER_PATH,
                           "first-child-admission.reader")
    helper = prior.get("helper")
    if helper is not None:
        checked_descriptor(helper, "first-child-admission.helper")
    semantic = read_json(folder / "semantics.json")
    require(isinstance(semantic, Mapping) and semantic.get("status") == "pass",
            "first-child corrected semantics is not pass")
    return 1


def analyze(output: Path, first_only: bool) -> dict[str, Any]:
    plan = validate_plan()
    rows = plan["_rows"]
    expected_count: int | None = None
    prior_admission: dict[str, Any] | None = None
    if not first_only:
        admission_path = _admission_path()
        prior_admission = read_json(admission_path)
        require(isinstance(prior_admission, Mapping),
                "full reduction requires a first-child-admission object")
        expected_count = _validate_first_admission(prior_admission, plan)

    selected_rows = rows[:1] if first_only else rows
    captures = [analyze_capture(plan, row, expected_count) for row in selected_rows]
    qualified_count = captures[0]["numbered_dump_count"]
    require(qualified_count in (1, 2), "first capture did not dynamically qualify one or two dumps")
    if not first_only:
        require(all(item["numbered_dump_count"] == qualified_count for item in captures),
                "capture children do not share the first-child qualified dump count")

    result: dict[str, Any] = {
        "schema": "litchi.pptx-multi-slide-callgrind-analysis-0908.v1",
        "status": "pass",
        "mode": "first-child-admission" if first_only else "full-reduction",
        "plan": {"file": packet_relative(plan["_plan_path"]),
                 "sha256": plan["_plan_sha256"]},
        "retained_reader": {"file": repo_relative(READER_PATH),
                            "sha256": plan["_reader_sha256"]},
        "binary": {"sha256": EXPECTED_BINARY_SHA256,
                   "bytes": plan["_binary"].get("bytes"),
                   "path": plan["_binary"].get("path")},
        "cpu": EXPECTED_CPU,
        "owner": OWNER,
        "measured_parent": MEASURED_PARENT,
        "measured_callsite": MEASURED_CALLSITE,
        "excluded_expected_callsite": EXPECTED_CALLSITE,
        "options": plan["_options"],
        "qualification": {
            "numbered_dump_count": qualified_count,
            "accepted_counts": [1, 2],
            "same_count_across_reduction": not first_only,
            "reason": "Callgrind may emit one or two owner-scoped parts depending on emitted/inlined call structure; the first retained child determines the count.",
        },
        "captures": captures,
        "scope": "Ir counts inside the target owner only; direct-callee rows are the attribution boundary.",
        "limitations": [
            "Callgrind Ir is a guest-instruction attribution counter, not native wall time, hardware cycles, allocations, RSS, or a speedup.",
            "Direct-callee inclusive Ir rows are a disjoint immediate-owner partition; nested production rows must not be added to those totals.",
            "Harness guards, runtime/dependency rows, and unclassified rows are retained for conservation and are excluded from high-confidence production totals.",
            "Source-path classification is evidence from Callgrind records; it does not prove semantic ownership of an inclusive callee body.",
        ],
        "validation": {
            "frozen_binary_bound": True,
            "fixed_cpu_bound": True,
            "fixed_owner_and_parent": True,
            "first_child_admitted_before_full_reduction": not first_only,
            "all_selected_children_reduced": not first_only or len(captures) == 1,
        },
    }
    if prior_admission is not None:
        result["prior_admission"] = {
            "file": packet_relative(_admission_path()),
            "sha256": sha256(_admission_path()),
        }
    output = output if output.is_absolute() else HERE / output
    output.parent.mkdir(parents=True, exist_ok=True)
    require(not output.exists() and not output.is_symlink(),
            f"refusing to overwrite existing analysis output: {output}")
    with output.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, indent=2, sort_keys=True)
        stream.write("\n")
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--first-child", action="store_true",
                        help="admit only eager-a and write first-child-admission.json")
    parser.add_argument("--output", type=Path,
                        default=None,
                        help="result JSON path; defaults to first-child-admission.json or analysis.json")
    args = parser.parse_args()
    output = args.output or (_admission_path() if args.first_child else HERE / "analysis.json")
    try:
        report = analyze(output, first_only=args.first_child)
    except (EvidenceError, OSError, ImportError, TypeError, ValueError) as error:
        print(f"analyze.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0908 Callgrind {report['mode']} passed: {output} ({len(report['captures'])} capture(s))")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
