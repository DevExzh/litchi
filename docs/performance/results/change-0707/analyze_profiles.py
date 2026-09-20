#!/usr/bin/env python3
"""Validate and attribute the 0707 XLSX planning Callgrind profiles.

The selected owner is the outer SourceBackedEditor::edit_sheets planning
operation.  Each profile child executes three lifecycle calls and one timed
call.  The analyzer selects those dumps by their positive raw incoming edges,
retains every raw dump and annotation, and keeps nested inclusive rows outside
the disjoint immediate-child partition.

Callgrind Ir is guest-instruction attribution.  This packet makes no native
latency, allocation, hardware-counter, or production-change claim.
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
RETAINED_0705 = HERE.parent / "change-0705" / "analyze_profiles.py"
RETAINED_0521 = HERE.parent / "change-0521" / "analyze_profiles.py"

OWNER = "litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets"
LIFECYCLE_PARENT = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
MEASURED_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
NUMBERED_PARTS = (1, 2, 3, 4)
TIMING_FIELDS = frozenset(
    {"open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns"}
)
PERL_ENV = {"PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"}
VALIDATOR_OBSERVE = "litchi_xlsx::cell_values::validation::Validator::observe"
VALIDATE_ELEMENT = "litchi_xlsx::cell_values::validation::validate_element"


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def digest_json(value: Any) -> str:
    encoded = (
        json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n"
    ).encode("utf-8")
    return hashlib.sha256(encoded).hexdigest()


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
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
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def load_retained_helpers() -> tuple[Any, Any]:
    """Load the retained 0705 wrapper and its immutable 0521 helpers."""

    require(RETAINED_0705.is_file(), f"missing retained helper {RETAINED_0705}")
    require(RETAINED_0521.is_file(), f"missing retained helper {RETAINED_0521}")
    spec = importlib.util.spec_from_file_location(
        "retained_0705_profile_analyzer_0707", RETAINED_0705
    )
    require(spec is not None and spec.loader is not None,
            f"cannot load retained helper {RETAINED_0705}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    helper = module.H
    helper.HERE = HERE
    return module, helper


RETAINED_0705_MODULE, H = load_retained_helpers()


def option(command: list[str], name: str) -> str:
    for token in command:
        if token.startswith(name + "="):
            return token.split("=", 1)[1]
    try:
        index = command.index(name)
        return command[index + 1]
    except (ValueError, IndexError) as error:
        raise EvidenceError(f"command omits {name}") from error


def plan_data() -> dict[str, Any]:
    plan = read_json(HERE / "plan.json")
    require(isinstance(plan, dict), "plan.json is not an object")
    require(isinstance(plan.get("revision"), str) and plan["revision"],
            "plan revision is missing")
    require(plan.get("case") == "xlsx_source_backed_cell_values_one_percent_edit_save",
            "planning profile case differs")
    require(plan.get("shapes") == ["medium", "dense-sparse"],
            "planning profile shape matrix differs")
    require(plan.get("profile_repeats") == 2 and plan.get("profile_samples") == 1,
            "planning profile repeat/sample policy differs")
    require(plan.get("profile_warmup") == 0, "planning profile warmup differs")
    require(isinstance(plan.get("cpu"), int), "plan CPU is missing")
    return plan


def profile_plan() -> dict[str, Any]:
    plan = read_json(HERE / "profile-plan.json")
    require(isinstance(plan, dict), "profile-plan.json is not an object")
    require(plan.get("owner") == OWNER, "profile owner differs")
    require(plan.get("shapes") == ["medium", "dense-sparse"],
            "profile shape matrix differs")
    require(plan.get("repeats") == 2, "profile repeat count differs")
    require(plan.get("warmup") == 0 and plan.get("samples") == 1,
            "profile warmup/sample policy differs")
    require(plan.get("expected_numbered_dumps") == len(NUMBERED_PARTS),
            "profile numbered-dump policy differs")
    require(plan.get("expected_lifecycle_calls") == 3,
            "profile lifecycle call policy differs")
    require(plan.get("expected_measured_calls") == 1,
            "profile measured call policy differs")
    require(plan.get("lifecycle_parent") == LIFECYCLE_PARENT,
            "lifecycle parent differs")
    require(plan.get("measured_parent") == MEASURED_PARENT,
            "measured parent differs")
    options = plan.get("options")
    require(isinstance(options, list) and all(isinstance(item, str) for item in options),
            "profile Callgrind options are invalid")
    for token in (
        "--tool=callgrind",
        "--collect-atstart=no",
        "--toggle-collect=" + OWNER,
        "--zero-before=" + OWNER,
        "--dump-after=" + OWNER,
    ):
        require(token in options, f"profile options omit {token}")
    return plan


def validate_binary_or_cleanup(
    binary: Path, expected_sha: str, label: str
) -> dict[str, Any]:
    """Bind a binary either to bytes or to an exact cleanup witness."""

    if binary.is_file() and not binary.is_symlink():
        require(sha256(binary) == expected_sha,
                f"{label}: frozen binary hash differs")
        return {
            "path": str(binary),
            "sha256": expected_sha,
            "bytes": binary.stat().st_size,
            "verified": True,
        }
    cleanup_path = HERE / "cleanup.json"
    require(cleanup_path.is_file(),
            f"{label}: frozen binary is missing without cleanup witness: {binary}")
    cleanup = read_json(cleanup_path)
    require(cleanup.get("owned_paths_absent") is True,
            f"{label}: cleanup witness does not establish owned paths are absent")
    witnesses = cleanup.get("binaries")
    require(isinstance(witnesses, list),
            f"{label}: cleanup witness binaries are not a list")
    matching = [
        item for item in witnesses
        if isinstance(item, dict) and item.get("path") == str(binary)
    ]
    require(len(matching) == 1, f"{label}: cleanup witness does not bind {binary}")
    witness = matching[0]
    require(witness.get("sha256") == expected_sha,
            f"{label}: cleanup witness binary hash differs")
    require(isinstance(witness.get("bytes"), int) and witness["bytes"] > 0,
            f"{label}: cleanup witness binary size is invalid")
    return {
        "path": str(binary),
        "sha256": expected_sha,
        "bytes": witness["bytes"],
        # Keep custody mode out of the report.  The same normalized identity
        # must replay byte-for-byte before and after owned-binary cleanup.
        "verified": True,
    }


def validate_planning_symbols(build_sha: str, stage: str) -> dict[str, Any]:
    """Validate the frozen two-specialization symbol discovery artifact."""

    text_path = HERE / "planning-symbols.txt"
    json_path = HERE / "planning-symbols.json"
    require(text_path.is_file() and json_path.is_file(),
            "planning-symbol discovery artifacts are missing")
    symbols = read_json(json_path)
    require(isinstance(symbols, dict), "planning-symbols.json is not an object")
    filters = symbols.get("filter_substrings", [])
    require(
        OWNER in filters or OWNER.removeprefix("litchi_xlsx::") in filters,
        "planning symbols omit the exact edit_sheets filter",
    )
    require(symbols.get("artifact_sha256") == sha256(text_path),
            "planning-symbol text hash differs")
    discovered_sha = symbols.get("binary_sha256")
    require(isinstance(discovered_sha, str) and len(discovered_sha) == 64,
            "planning-symbol binary hash is missing")
    lines = [line.strip() for line in read_text(text_path).splitlines()]
    owner_lines = [line for line in lines if line.endswith(" " + OWNER)]
    require(len(owner_lines) == 2,
            f"expected two edit_sheets specializations, got {len(owner_lines)}")
    if stage == "baseline":
        require(discovered_sha == build_sha,
                "planning-symbol binary differs from baseline native binary")
    return {
        "text": relative(text_path),
        "text_sha256": sha256(text_path),
        "json": relative(json_path),
        "json_sha256": sha256(json_path),
        "binary_sha256": discovered_sha,
        "owner_specialization_count": len(owner_lines),
        "binary_matches_stage": discovered_sha == build_sha,
        "stage": stage,
    }


def validate_build(stage: str) -> dict[str, Any]:
    records_path = HERE / f"build-{stage}.json"
    source_path = HERE / f"source-{stage}.json"
    records = read_json(records_path)
    expected_source = read_json(source_path)
    require(isinstance(records, list), f"{relative(records_path)} is not a record list")
    require(isinstance(expected_source, dict),
            f"{relative(source_path)} is not a source map")
    binary_name = f"{stage}-native"
    matches = [
        record for record in records
        if isinstance(record, dict)
        and Path(str(record.get("binary", ""))).name == binary_name
    ]
    require(len(matches) == 1,
            f"{relative(records_path)}: expected one {binary_name} record")
    record = matches[0]
    require(record.get("exit_code") == 0,
            f"{relative(records_path)}: {binary_name} build failed")
    binary = Path(str(record.get("binary", ""))).resolve()
    binary_sha = record.get("binary_sha256")
    require(isinstance(binary_sha, str) and len(binary_sha) == 64,
            f"{relative(records_path)}: native binary hash is missing")
    source_sha = sha256(source_path)
    require(record.get("source_manifest_sha256") == source_sha,
            f"{relative(records_path)}: source manifest binding differs")
    custody = validate_binary_or_cleanup(binary, binary_sha, stage)
    return {
        "stage": stage,
        "record": relative(records_path),
        "record_sha256": sha256(records_path),
        "source_manifest": relative(source_path),
        "source_manifest_sha256": source_sha,
        "source_entry_count": len(expected_source),
        "binary": str(binary),
        "binary_sha256": binary_sha,
        "binary_bytes": record.get("binary_bytes"),
        "binary_custody": custody,
        "expected_source": expected_source,
    }


def retained_helper_identities() -> dict[str, str]:
    """Validate and return the packet's transitive retained-helper custody."""

    root = HERE.parents[3]
    identity_path = HERE / "helper-identities.json"
    if identity_path.is_file():
        identities = read_json(identity_path)
        require(isinstance(identities, dict),
                "helper-identities.json is not an object")
        for name, digest in identities.items():
            require(isinstance(name, str) and isinstance(digest, str),
                    "helper identity entry is malformed")
            path = root / name
            require(path.is_file() and not path.is_symlink(),
                    f"retained helper is missing: {name}")
            require(sha256(path) == digest,
                    f"retained helper hash differs: {name}")
        return {str(name): str(digest) for name, digest in sorted(identities.items())}
    return {
        str(RETAINED_0705.relative_to(root)): sha256(RETAINED_0705),
        str(RETAINED_0521.relative_to(root)): sha256(RETAINED_0521),
    }


def validate_source_receipt(
    receipt: dict[str, Any], expected_source: dict[str, str], label: str
) -> dict[str, Any]:
    current = receipt.get("current_checkout_source")
    require(isinstance(current, dict), f"{label}: current source custody is missing")
    before_name = current.get("before_artifact")
    after_name = current.get("after_artifact")
    require(isinstance(before_name, str) and isinstance(after_name, str),
            f"{label}: source custody artifact names are missing")
    for filename in (before_name, after_name):
        require(Path(filename).name == filename,
                f"{label}: source custody path escapes packet: {filename}")
    before_path = HERE / before_name
    after_path = HERE / after_name
    for path in (before_path, after_path):
        require(path.is_file() and not path.is_symlink(),
                f"{label}: source custody artifact is missing: {path}")
    before = read_json(before_path)
    after = read_json(after_path)
    require(before == expected_source and after == expected_source,
            f"{label}: current source census differs from retained source")
    require(before == after and current.get("unchanged_during_child") is True,
            f"{label}: source changed during child")
    require(current.get("before_sha256") == digest_json(before),
            f"{label}: before source digest differs")
    require(current.get("after_sha256") == digest_json(after),
            f"{label}: after source digest differs")
    require(current.get("before_file_sha256") == sha256(before_path),
            f"{label}: before source file hash differs")
    require(current.get("after_file_sha256") == sha256(after_path),
            f"{label}: after source file hash differs")
    return {
        "before": relative(before_path),
        "after": relative(after_path),
        "before_sha256": sha256(before_path),
        "after_sha256": sha256(after_path),
        "source_equal": True,
    }


def validate_receipt_artifacts(
    receipt: dict[str, Any], required: set[str], label: str
) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifacts are not an object")
    require(required.issubset(artifacts),
            f"{label}: required artifacts omitted: {sorted(required - set(artifacts))}")
    for filename, digest in artifacts.items():
        require(isinstance(filename, str) and Path(filename).name == filename,
                f"{label}: artifact path escapes packet: {filename!r}")
        path = HERE / filename
        require(path.is_file() and not path.is_symlink(),
                f"{label}: artifact is missing: {filename}")
        require(sha256(path) == digest, f"{label}: artifact hash differs: {filename}")


def validate_common_receipt(
    receipt: dict[str, Any], build: dict[str, Any], plan: dict[str, Any],
    profile: dict[str, Any], label: str,
) -> dict[str, Any]:
    require(receipt.get("exit_code") == 0, f"{label}: child failed")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{label}: binary binding differs")
    require(receipt.get("build_record_sha256") == build["record_sha256"],
            f"{label}: build record binding differs")
    require(receipt.get("build_source_manifest_sha256") ==
            build["source_manifest_sha256"],
            f"{label}: build source binding differs")
    require(receipt.get("retained_source_manifest") == build["source_manifest"],
            f"{label}: retained source manifest path differs")
    require(receipt.get("retained_source_manifest_sha256") ==
            build["source_manifest_sha256"],
            f"{label}: retained source manifest hash differs")
    require(receipt.get("retained_source_census_sha256") ==
            digest_json(build["expected_source"]),
            f"{label}: retained source census digest differs")
    require(receipt.get("retained_source_entry_count") ==
            len(build["expected_source"]),
            f"{label}: retained source entry count differs")
    require(receipt.get("plan_sha256") == sha256(HERE / "plan.json"),
            f"{label}: plan binding differs")
    require(receipt.get("script_sha256") == sha256(HERE / "capture.py"),
            f"{label}: capture script binding differs")
    require(receipt.get("constraints_sha256") == sha256(HERE / "constraints.json"),
            f"{label}: constraints binding differs")
    if receipt.get("lane") == "profile":
        require(receipt.get("profile_plan_sha256") ==
                sha256(HERE / "profile-plan.json"),
                f"{label}: profile plan binding differs")
    return validate_source_receipt(receipt, build["expected_source"], label)


def validate_native_receipt(
    name: str, receipt: dict[str, Any], build: dict[str, Any],
    plan: dict[str, Any], profile: dict[str, Any], shape: str,
) -> dict[str, Any]:
    label = f"{name}.receipt.json"
    source_custody = validate_common_receipt(receipt, build, plan, profile, label)
    command = receipt.get("command")
    require(isinstance(command, list), f"{label}: command is not a list")
    require(command[:3] == ["taskset", "-c", str(plan["cpu"])],
            f"{label}: command prefix differs")
    require(option(command, "--warmup") == str(plan["warmup"]),
            f"{label}: native warmup differs")
    require(option(command, "--samples") == str(plan["samples"]),
            f"{label}: native samples differs")
    require(option(command, "--case") == plan["case"],
            f"{label}: native case differs")
    require(option(command, "--xlsx-cell-crud-shape") == shape,
            f"{label}: native shape differs")
    require(Path(option(command, "--json")).name == name + ".json",
            f"{label}: native JSON output path differs")
    required = {name + suffix for suffix in (".json", ".stdout", ".stderr")}
    required.update({
        receipt["current_checkout_source"][key]
        for key in ("before_artifact", "after_artifact")
    })
    validate_receipt_artifacts(receipt, required, label)
    return source_custody


def validate_profile_receipt(
    name: str, receipt: dict[str, Any], build: dict[str, Any],
    plan: dict[str, Any], profile: dict[str, Any], shape: str,
) -> dict[str, Any]:
    label = f"{name}.receipt.json"
    source_custody = validate_common_receipt(receipt, build, plan, profile, label)
    command = receipt.get("command")
    require(isinstance(command, list), f"{label}: command is not a list")
    require(command[:4] == ["taskset", "-c", str(plan["cpu"]), "valgrind"],
            f"{label}: profile command prefix differs")
    for token in profile["options"]:
        require(token in command, f"{label}: profile command omits {token}")
    require(option(command, "--warmup") == "0", f"{label}: profile warmup differs")
    require(option(command, "--samples") == "1", f"{label}: profile samples differs")
    require(option(command, "--case") == plan["case"],
            f"{label}: profile case differs")
    require(option(command, "--xlsx-cell-crud-shape") == shape,
            f"{label}: profile shape differs")
    require(Path(option(command, "--json")).name == name + ".json",
            f"{label}: profile JSON output path differs")
    require(Path(option(command, "--callgrind-out-file")).name ==
            name + ".callgrind",
            f"{label}: Callgrind output path differs")
    required = {
        name + suffix for suffix in (".json", ".stdout", ".stderr", ".callgrind")
    }
    required.update(f"{name}.callgrind.{part}" for part in NUMBERED_PARTS)
    required.update({
        receipt["current_checkout_source"][key]
        for key in ("before_artifact", "after_artifact")
    })
    validate_receipt_artifacts(receipt, required, label)
    return source_custody


def canonical(value: Any, label: str) -> Any:
    if isinstance(value, list):
        require(bool(value), f"{label}: empty vector")
        values = [canonical(item, f"{label}[{index}]")
                  for index, item in enumerate(value)]
        require(all(item == values[0] for item in values),
                f"{label}: vector is not constant")
        return values[0]
    if isinstance(value, dict):
        return {
            key: canonical(item, f"{label}.{key}")
            for key, item in sorted(value.items())
        }
    return value


def source_identity(source: Any, label: str) -> Any:
    require(isinstance(source, dict), f"{label}: source is not an object")
    result: dict[str, Any] = {}
    for key, value in sorted(source.items()):
        if key == "xlsx_cell_values":
            require(isinstance(value, dict), f"{label}.{key}: not an object")
            nested = {}
            for nested_key, nested_value in sorted(value.items()):
                if nested_key in TIMING_FIELDS or nested_key.endswith("allocation_metrics"):
                    continue
                nested[nested_key] = canonical(
                    nested_value, f"{label}.{key}.{nested_key}"
                )
            result[key] = nested
        else:
            result[key] = canonical(value, f"{label}.{key}")
    return result


def result_identity(result: Any, label: str) -> dict[str, Any]:
    require(isinstance(result, dict), f"{label}: result is not an object")
    require(all(key in result for key in ("case", "corpus", "sink", "source")),
            f"{label}: result lacks case/corpus/sink/source")
    return {
        "case": result["case"],
        "corpus": result["corpus"],
        "sink": result["sink"],
        "source": source_identity(result["source"], label + ".source"),
        "output_sha256": result.get("output_sha256"),
    }


def validate_result_report(
    path: Path, build: dict[str, Any], plan: dict[str, Any],
    shape: str, profile_report: bool,
) -> tuple[dict[str, Any], dict[str, Any]]:
    report = read_json(path)
    label = relative(path)
    require(report.get("schema_version") == 1, f"{label}: schema version differs")
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{label}: tool identity is missing")
    require(tool.get("binary") == "litchi-perf-baseline",
            f"{label}: unexpected benchmark binary")
    require(tool.get("profile") == "release", f"{label}: report is not release")
    require(tool.get("instrumentation") == "none", f"{label}: report is instrumented")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict)
            and identity.get("binary_sha256") == build["binary_sha256"],
            f"{label}: binary identity differs")
    environment = report.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == plan["revision"],
            f"{label}: report revision differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{label}: configuration is missing")
    require(configuration.get("cases") == [plan["case"]],
            f"{label}: case configuration differs")
    require(configuration.get("xlsx_cell_crud_shapes") == [shape],
            f"{label}: shape configuration differs")
    expected_samples = 1 if profile_report else plan["samples"]
    expected_warmup = 0 if profile_report else plan["warmup"]
    require(configuration.get("samples_per_case") == expected_samples,
            f"{label}: sample configuration differs")
    require(configuration.get("warmup_iterations_per_case") == expected_warmup,
            f"{label}: warmup configuration differs")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}: expected one result")
    identity_value = result_identity(results[0], label + ".results[0]")
    return report, identity_value


def profile_name(stage: str, repeat: int, shape: str) -> str:
    return f"profile-{stage}-r{repeat}-{shape}"


def native_name(stage: str, repeat: int, shape: str) -> str:
    return f"native-{stage}-r{repeat}-{shape}"


def numbered_dumps(stem: Path, expected: tuple[int, ...]) -> list[Path]:
    paths = [
        path for path in HERE.glob(stem.name + ".[0-9]*")
        if path.is_file() and path.suffix[1:].isdigit()
    ]
    paths.sort(key=lambda path: int(path.suffix[1:]))
    parts = [int(path.suffix[1:]) for path in paths]
    require(parts == list(expected),
            f"{stem.name}: numbered dumps differ from {list(expected)}: {parts}")
    return paths


def edge(path: Path, target: str, caller: str | None = None) -> dict[str, Any]:
    return H.target_edge_summary(path, target, caller)


def classify_dumps(
    paths: list[Path], profile: dict[str, Any]
) -> tuple[list[dict[str, Any]], Path]:
    rows: list[dict[str, Any]] = []
    measured: list[Path] = []
    lifecycle: list[Path] = []
    owner = profile["owner"]
    for path in paths:
        label = relative(path)
        text = read_text(path)
        part = int(path.suffix[1:])
        require(H.part_number(text, label) == part,
                f"{label}: Callgrind part does not match suffix")
        require(H.trigger(text, label) == f"--dump-after={owner}",
                f"{label}: dump trigger differs")
        require("events: Ir" in text.splitlines(), f"{label}: Ir event is absent")
        summary = H.summary_ir(text, label)
        all_incoming = edge(path, owner)
        measured_edge = edge(path, owner, profile["measured_parent"])
        lifecycle_edge = edge(path, owner, profile["lifecycle_parent"])
        require(all_incoming["calls"] == 1,
                f"{label}: owner incoming call count is {all_incoming['calls']}")
        if measured_edge["calls"]:
            require(measured_edge["calls"] == 1,
                    f"{label}: measured owner call count differs")
            require(lifecycle_edge["calls"] == 0,
                    f"{label}: dump has both measured and lifecycle owner calls")
            require(measured_edge["inclusive_ir"] == summary,
                    f"{label}: measured owner Ir differs from dump summary")
            measured.append(path)
            role = "measured"
        else:
            require(lifecycle_edge["calls"] == 1,
                    f"{label}: dump has neither expected owner parent")
            require(measured_edge["calls"] == 0,
                    f"{label}: lifecycle dump has a measured owner call")
            require(lifecycle_edge["inclusive_ir"] == summary,
                    f"{label}: lifecycle owner Ir differs from dump summary")
            lifecycle.append(path)
            role = "lifecycle"
        rows.append({
            "file": label,
            "part": part,
            "sha256": sha256(path),
            "summary_ir": summary,
            "role": role,
            "all_owner_incoming": all_incoming,
            "measured_parent": measured_edge,
            "lifecycle_parent": lifecycle_edge,
        })
    require(len(lifecycle) == profile["expected_lifecycle_calls"],
            f"expected {profile['expected_lifecycle_calls']} lifecycle dumps, got {len(lifecycle)}")
    require(len(measured) == profile["expected_measured_calls"],
            f"expected {profile['expected_measured_calls']} measured dumps, got {len(measured)}")
    return rows, measured[0]


def terminal_dump(path: Path, numbered: list[Path]) -> dict[str, Any]:
    label = relative(path)
    require(path.is_file(), f"{label}: termination dump is missing")
    text = read_text(path)
    part = H.part_number(text, label)
    require(part == int(numbered[-1].suffix[1:]) + 1,
            f"{label}: termination part does not follow numbered dumps")
    require(H.trigger(text, label) == "Program termination",
            f"{label}: termination trigger differs")
    require(H.summary_ir(text, label) == 0,
            f"{label}: termination dump is not zero-Ir")
    require("events: Ir" in text.splitlines(), f"{label}: termination Ir is absent")
    return {
        "file": label,
        "part": part,
        "sha256": sha256(path),
        "summary_ir": 0,
        "trigger": "Program termination",
    }


def write_or_check(path: Path, text: str) -> None:
    if path.exists():
        require(read_text(path) == text,
                f"annotation replay differs: {relative(path)}")
    else:
        path.write_text(text, encoding="utf-8")


def annotate(
    selected: Path, name: str, owner: str, summary: int
) -> tuple[dict[str, Any], str, str]:
    inclusive_text, inclusive_command = H.run_annotation(selected, True)
    self_text, self_command = H.run_annotation(selected, False)
    inclusive_path = HERE / (name + ".inclusive.txt")
    self_path = HERE / (name + ".self.txt")
    write_or_check(inclusive_path, inclusive_text)
    write_or_check(self_path, self_text)
    inclusive = H.parse_annotation(inclusive_text, owner, relative(inclusive_path))
    exclusive = H.parse_annotation(self_text, owner, relative(self_path))
    require(inclusive["selected_ir"] == summary,
            f"{name}: inclusive owner Ir differs from raw summary")
    require(exclusive["direct"] == inclusive["direct"],
            f"{name}: inclusive/self direct children differ")
    direct_ir = sum(item["inclusive_ir"] for item in inclusive["direct"])
    require(exclusive["selected_ir"] + direct_ir == inclusive["selected_ir"],
            f"{name}: self plus immediate children does not reconstruct owner Ir")
    annotation = {
        "environment": dict(PERL_ENV),
        "command": {"inclusive": inclusive_command, "self": self_command},
        "files": {
            "inclusive": relative(inclusive_path),
            "inclusive_sha256": sha256(inclusive_path),
            "self": relative(self_path),
            "self_sha256": sha256(self_path),
        },
        "owner": {
            "inclusive_ir": inclusive["selected_ir"],
            "self_ir": exclusive["selected_ir"],
            "direct_callees": inclusive["direct"],
            "direct_callee_ir": H.direct_map(inclusive["direct"]),
        },
        "immediate_child_partition": {
            "self_ir": exclusive["selected_ir"],
            "direct_children_ir": direct_ir,
            "owner_ir": inclusive["selected_ir"],
            "disjoint": True,
            "equation": "self_ir + sum(immediate direct-child inclusive Ir) = owner inclusive Ir",
            "nested_inclusive_costs_excluded": True,
        },
        "validation": {
            "inclusive_owner_matches_raw": True,
            "self_plus_immediate_children_equals_inclusive": True,
            "inclusive_and_self_direct_children_match": True,
        },
    }
    return annotation, inclusive_text, self_text


def annotation_rows(text: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    star_re = H.EDGES.STAR_RE
    for line_number, line in enumerate(text.splitlines(), 1):
        match = star_re.match(line)
        if match:
            rows.append({
                "line": line_number,
                "name": H.display_name(match.group(2)),
                "inclusive_ir": int(match.group(1).replace(",", "")),
            })
    return rows


def raw_function_names(path: Path) -> set[str]:
    """Collect demangled names from raw records for optional nested probes."""

    function_re = getattr(getattr(H, "EDGES", None), "raw", None)
    function_re = getattr(function_re, "FUNCTION_RE", None)
    fallback = re.compile(r"^(?:fn|cfn)=\(\d+\)(?:\s+(.*))?$")
    names: set[str] = set()
    for line in read_text(path).splitlines():
        stripped = line.strip()
        match = function_re.match(stripped) if function_re else fallback.match(stripped)
        if not match:
            continue
        raw_name = match.group(3) if function_re else match.group(1)
        if raw_name:
            names.add(H.display_name(raw_name))
    return names


def nested_category(name: str) -> str | None:
    if name == VALIDATOR_OBSERVE:
        return "validator_observe"
    if name == VALIDATE_ELEMENT:
        return "validator_validate_element"
    if "FactsBuilder::" in name:
        return "facts_builder"
    if name.startswith("litchi_xlsx::cell::Store::"):
        return "cell_store"
    if name.startswith("litchi_xlsx::raw::worksheet::") and (
        "Parser::" in name or "::parse" in name or name.endswith("::parse")
    ):
        return "raw_worksheet_parser"
    return None


def nested_diagnostics(
    selected: Path, inclusive_text: str, self_text: str
) -> dict[str, Any]:
    """Report nested inclusive rows without adding them to the partition."""

    rows = annotation_rows(inclusive_text)
    targets: dict[str, set[str]] = {}
    for row in rows:
        category = nested_category(row["name"])
        if category is not None:
            targets.setdefault(category, set()).add(row["name"])
    for name in raw_function_names(selected):
        category = nested_category(name)
        if category is not None:
            targets.setdefault(category, set()).add(name)

    result: dict[str, Any] = {}
    categories = (
        "validator_observe",
        "validator_validate_element",
        "raw_worksheet_parser",
        "facts_builder",
        "cell_store",
    )
    for category in categories:
        entries: list[dict[str, Any]] = []
        for target in sorted(targets.get(category, ())):
            annotation_matches = [row for row in rows if row["name"] == target]
            raw_edge = edge(selected, target)
            nested_direct: dict[str, Any] | None = None
            required_target = category in {
                "validator_observe",
                "validator_validate_element",
            }
            if len(annotation_matches) == 1:
                try:
                    inclusive = H.parse_annotation(
                        inclusive_text, target, f"{relative(selected)}:{target}"
                    )
                    exclusive = H.parse_annotation(
                        self_text, target, f"{relative(selected)}:{target}:self"
                    )
                    require(exclusive["direct"] == inclusive["direct"],
                            f"{relative(selected)}: nested inclusive/self children differ for {target}")
                    direct_children = []
                    for child in inclusive["direct"]:
                        raw_child = edge(selected, child["name"], target)
                        if required_target:
                            require(raw_child["positive_edge_count"] > 0,
                                    f"{relative(selected)}: raw edge missing for {target} -> {child['name']}")
                            require(raw_child["calls"] == child["calls"],
                                    f"{relative(selected)}: raw/annotation calls differ for {target} -> {child['name']}")
                            require(raw_child["inclusive_ir"] == child["inclusive_ir"],
                                    f"{relative(selected)}: raw/annotation Ir differs for {target} -> {child['name']}")
                        direct_children.append({**child, "raw_edge": raw_child})
                    allocator_edges = [
                        child for child in direct_children
                        if "__rust_alloc" in child["name"]
                        or "__rust_dealloc" in child["name"]
                    ]
                    nested_direct = {
                        "inclusive_ir": inclusive["selected_ir"],
                        "self_ir": exclusive["selected_ir"],
                        "direct_children": direct_children,
                        "direct_children_ir": sum(
                            child["inclusive_ir"] for child in direct_children
                        ),
                        "equation_valid": (
                            exclusive["selected_ir"]
                            + sum(child["inclusive_ir"] for child in direct_children)
                            == inclusive["selected_ir"]
                        ),
                        "allocator_edges": allocator_edges,
                        "allocator_edge_interpretation": (
                            "Raw Callgrind calls fields on positive-Ir allocator "
                            "edges, retained without interpreting their collection "
                            "scope as measured-region allocation counts. Dense "
                            "counts exceed the independent planning allocation "
                            "counter; this discrepancy is unresolved."
                        ),
                        "partition_boundary": (
                            "Immediate children of this nested owner are retained "
                            "as diagnostics only and are not added to the outer "
                            "edit_sheets partition."
                        ),
                    }
                except (AssertionError, EvidenceError) as error:
                    if required_target:
                        raise EvidenceError(
                            f"{relative(selected)}: required nested target {target} "
                            f"could not be validated: {error}"
                        ) from error
                    # The retained annotation helper can omit a nested row's
                    # tree when callgrind_annotate cannot represent a unique
                    # parent path.  Raw incoming edges remain useful evidence.
                    nested_direct = None
            entries.append({
                "target": target,
                "annotation_rows": annotation_matches,
                "raw_incoming": raw_edge,
                "nested_direct": nested_direct,
                "available": bool(annotation_matches or raw_edge["positive_edge_count"]),
                "overlap": "inclusive nested diagnostic; excluded from immediate-child sums",
            })
        result[category] = {
            "available": bool(entries),
            "entries": entries,
            "aggregation": (
                "Rows and raw incoming edges are retained individually. Their "
                "inclusive Ir is overlapping diagnostic evidence and is never "
                "added to owner or immediate-child totals."
            ),
        }
    return result


def analyze_profile(
    stage: str, plan: dict[str, Any], profile: dict[str, Any],
    build: dict[str, Any], repeat: int, shape: str,
) -> dict[str, Any]:
    name = profile_name(stage, repeat, shape)
    receipt_path = HERE / f"{name}.receipt.json"
    profile_path = HERE / f"{name}.json"
    require(receipt_path.is_file(), f"{relative(receipt_path)}: receipt is missing")
    require(profile_path.is_file(), f"{relative(profile_path)}: result is missing")
    receipt = read_json(receipt_path)
    source_custody = validate_profile_receipt(
        name, receipt, build, plan, profile, shape
    )
    _, profile_identity = validate_result_report(
        profile_path, build, plan, shape, True
    )

    native = native_name(stage, repeat, shape)
    native_path = HERE / f"{native}.json"
    native_receipt_path = HERE / f"{native}.receipt.json"
    require(native_path.is_file(), f"{relative(native_path)}: native result is missing")
    require(native_receipt_path.is_file(),
            f"{relative(native_receipt_path)}: native receipt is missing")
    native_receipt = read_json(native_receipt_path)
    native_source_custody = validate_native_receipt(
        native, native_receipt, build, plan, profile, shape
    )
    _, native_identity = validate_result_report(
        native_path, build, plan, shape, False
    )
    require(profile_identity == native_identity,
            f"{name}: profile/native corpus, source, sink, or output identity differs")

    stem = HERE / f"{name}.callgrind"
    dumps = numbered_dumps(stem, NUMBERED_PARTS)
    parts, selected = classify_dumps(dumps, profile)
    termination = terminal_dump(stem, dumps)
    selected_part = next(item for item in parts if item["file"] == relative(selected))
    annotations, inclusive_text, self_text = annotate(
        selected, name, profile["owner"], selected_part["summary_ir"]
    )
    nested = nested_diagnostics(selected, inclusive_text, self_text)
    owner = annotations["owner"]
    return {
        "name": name,
        "stage": stage,
        "repeat": repeat,
        "shape": shape,
        "profile_receipt": relative(receipt_path),
        "profile_receipt_sha256": sha256(receipt_path),
        "profile_result": relative(profile_path),
        "profile_result_sha256": sha256(profile_path),
        "source_custody": source_custody,
        "native": {
            "name": native,
            "receipt": relative(native_receipt_path),
            "receipt_sha256": sha256(native_receipt_path),
            "result": relative(native_path),
            "result_sha256": sha256(native_path),
            "source_custody": native_source_custody,
            "identity": native_identity,
        },
        "result_identity": profile_identity,
        "parts": parts,
        "termination": termination,
        "selected": relative(selected),
        "selected_by": {
            "caller": profile["measured_parent"],
            "positive_owner_call_count": 1,
            "selection_is_raw_incoming_edge_based": True,
        },
        "annotations": annotations,
        "nested_diagnostics": nested,
        "total_ir": owner["inclusive_ir"],
        "self_ir": owner["self_ir"],
        "direct_ir": annotations["immediate_child_partition"]["direct_children_ir"],
        "direct_callee_ir": owner["direct_callee_ir"],
    }


def expected_profiles(profile: dict[str, Any]) -> list[tuple[int, str]]:
    return [
        (repeat, shape)
        for repeat in range(1, profile["repeats"] + 1)
        for shape in profile["shapes"]
    ]


def stage_has_profiles(stage: str, profile: dict[str, Any]) -> bool:
    return all(
        (HERE / f"{profile_name(stage, repeat, shape)}.callgrind.1").is_file()
        for repeat, shape in expected_profiles(profile)
    )


def analyze_stage(
    stage: str, plan: dict[str, Any], profile: dict[str, Any]
) -> dict[str, Any]:
    build = validate_build(stage)
    planning_symbols = validate_planning_symbols(build["binary_sha256"], stage)
    rows = [
        analyze_profile(stage, plan, profile, build, repeat, shape)
        for repeat, shape in expected_profiles(profile)
    ]
    require(len(rows) == profile["repeats"] * len(profile["shapes"]),
            f"{stage}: profile matrix is incomplete")

    child_totals: dict[str, int] = {}
    child_calls: dict[str, int] = {}
    for row in rows:
        for name, value in row["direct_callee_ir"].items():
            child_totals[name] = child_totals.get(name, 0) + value
        for part in row["annotations"]["owner"]["direct_callees"]:
            child_calls[part["name"]] = child_calls.get(part["name"], 0) + (
                part["calls"] or 0
            )
    require(child_totals, f"{stage}: immediate direct-child partition is empty")
    total_ir = sum(row["total_ir"] for row in rows)
    self_ir = sum(row["self_ir"] for row in rows)
    direct_ir = sum(row["direct_ir"] for row in rows)
    return {
        "stage": stage,
        "metadata": {
            key: value for key, value in build.items() if key != "expected_source"
        },
        "planning_symbols": planning_symbols,
        "profile_count": len(rows),
        "profiles": rows,
        "aggregate": {
            "measured_owner_total_ir": total_ir,
            "measured_owner_self_ir": self_ir,
            "immediate_child_partition_ir": direct_ir,
            "immediate_direct_children": [
                {
                    "name": name,
                    "aggregate_inclusive_ir": child_totals[name],
                    "aggregate_calls": child_calls.get(name, 0),
                    "share_of_owner_total_ir": child_totals[name] / total_ir,
                }
                for name in sorted(child_totals,
                                   key=lambda item: (-child_totals[item], item))
            ],
            "nested_costs_added": False,
            "partition_equation": (
                "aggregate self Ir + aggregate immediate-child Ir = aggregate owner Ir"
            ),
        },
        "validation": {
            "expected_profile_matrix": True,
            "all_profile_receipts_and_artifacts_valid": True,
            "all_native_output_corpus_source_identity_valid": True,
            "all_numbered_raw_dumps_retained_and_valid": True,
            "all_termination_dumps_zero_ir": True,
            "immediate_children_are_disjoint_partition": True,
            "nested_inclusive_costs_excluded_from_partition": True,
        },
    }


def metric(baseline: int, candidate: int) -> dict[str, Any]:
    return {
        "baseline": baseline,
        "candidate": candidate,
        "delta_ir": candidate - baseline,
        "candidate_over_baseline": candidate / baseline if baseline else None,
        "delta_percent": ((candidate / baseline - 1.0) * 100.0
                          if baseline else None),
    }


def compare_stages(
    baseline: dict[str, Any], candidate: dict[str, Any]
) -> dict[str, Any]:
    left = {(row["repeat"], row["shape"]): row for row in baseline["profiles"]}
    right = {(row["repeat"], row["shape"]): row for row in candidate["profiles"]}
    require(set(left) == set(right), "baseline/candidate profile matrices differ")
    rows = []
    for key in sorted(left):
        base = left[key]
        cand = right[key]
        identity_equal = base["result_identity"] == cand["result_identity"]
        require(identity_equal,
                f"{base['name']}: baseline/candidate native identity differs")
        names = set(base["direct_callee_ir"]) | set(cand["direct_callee_ir"])
        rows.append({
            "repeat": key[0],
            "shape": key[1],
            "baseline_profile": base["name"],
            "candidate_profile": cand["name"],
            "native_identity_equal": identity_equal,
            "metrics": {
                "owner_total_ir": metric(base["total_ir"], cand["total_ir"]),
                "owner_self_ir": metric(base["self_ir"], cand["self_ir"]),
                "immediate_child_partition_ir": metric(
                    base["direct_ir"], cand["direct_ir"]
                ),
            },
            "immediate_direct_children": [
                {
                    "name": name,
                    "metric": metric(
                        base["direct_callee_ir"].get(name, 0),
                        cand["direct_callee_ir"].get(name, 0),
                    ),
                }
                for name in sorted(names)
            ],
            "nested_comparison": (
                "Nested Validator/parser/FactsBuilder/Store rows remain separate "
                "overlapping diagnostics and are not numerically aggregated here."
            ),
        })
    return {
        "available": True,
        "profile_count": len(rows),
        "profiles": rows,
        "validation": {
            "profile_matrix_matches": True,
            "native_identity_parity": True,
            "immediate_partition_compared": True,
            "nested_costs_not_added": True,
        },
    }


def analyze(stage_selection: str = "both") -> dict[str, Any]:
    plan = plan_data()
    profile = profile_plan()
    if stage_selection == "baseline":
        selected = ["baseline"]
    elif stage_selection == "candidate":
        selected = ["candidate"]
    else:
        selected = [
            stage for stage in ("baseline", "candidate")
            if stage_has_profiles(stage, profile)
        ]
        require(selected, "no completed profile stage is available")
    stages = {stage: analyze_stage(stage, plan, profile) for stage in selected}
    comparison = None
    if "baseline" in stages and "candidate" in stages:
        comparison = compare_stages(stages["baseline"], stages["candidate"])
    return {
        "schema": "xlsx_callgrind_edit_sheets_planning_analysis_v1",
        "status": "pass",
        "selected_function": OWNER,
        "plan": relative(HERE / "plan.json"),
        "plan_sha256": sha256(HERE / "plan.json"),
        "profile_plan": relative(HERE / "profile-plan.json"),
        "profile_plan_sha256": sha256(HERE / "profile-plan.json"),
        "stage_selection": selected,
        "stages": stages,
        "comparison": comparison,
        "helpers": retained_helper_identities(),
        "limitations": [
            "Callgrind Ir is guest-instruction attribution, not native latency, hardware instructions, cycles, allocation counts, RSS, or cache counters.",
            "The selected owner covers edit_sheets planning, including source-backed selector/catalog/closure work; staging, commit, publication, and returned-edit teardown are outside this owner.",
            "Immediate direct children form the only disjoint attribution partition. Validator, parser, FactsBuilder, and Store rows are nested inclusive diagnostics and are never added to owner or direct-child totals.",
            "Profile/native parity checks captured output, corpus, source counters, and output digest. It is not a correctness proof for untested corpora or a native Office claim.",
            "The two edit_sheets specializations are proven only for the frozen planning-symbol binary and captured route; the artifact does not establish broad monomorphization coverage.",
            "Valgrind allocator replacement can expose allocator edge counts and guest Ir, but those edges do not predict native malloc latency or production allocator cost.",
            "Raw __rust_alloc and __rust_dealloc calls fields are not admitted as measured-region allocation counts: dense reports 233339 per edge versus 129411 total planning allocation requests in the independent allocator lane. Their collection-scope discrepancy is unresolved; no causal explanation or count-based allocation attribution is claimed.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", nargs="?", type=Path,
                        help="JSON destination; defaults to profile-analysis.json")
    parser.add_argument("--output", dest="output_option", type=Path,
                        help="JSON destination (alternative to positional path)")
    parser.add_argument("--stage", choices=("baseline", "candidate", "both"),
                        default="both",
                        help="validate one stage or all available stages")
    args = parser.parse_args()
    if args.output is not None and args.output_option is not None:
        parser.error("provide the output path either positionally or with --output")
    output = args.output_option or args.output or (HERE / "profile-analysis.json")
    try:
        document = analyze(args.stage)
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(document, indent=2, sort_keys=True) + "\n",
                          encoding="utf-8")
    except (EvidenceError, OSError, json.JSONDecodeError) as error:
        print(f"analyze_profiles.py: error: {error}", file=sys.stderr)
        return 2
    print(f"Planning profile edges, native identity, and annotations verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
