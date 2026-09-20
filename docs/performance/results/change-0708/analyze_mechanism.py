#!/usr/bin/env python3
"""Validate the conditional 0708 Callgrind and RSS mechanism evidence.

The analyzer binds every child to the frozen source/build custody and to the
0708 ABBA native result with the same shape and repeat.  It selects the
Callgrind measured dump from its raw incoming edge into the direct runner and
retains the three lifecycle dumps, the termination dump, and all annotations.
Only the outer ``edit_sheets`` Ir total is compared for the primary gate.
Callgrind allocation edges are retained as raw evidence by the files but are
not parsed, summed, or interpreted because the 0707 discrepancy is unresolved.
RSS is parsed from GNU ``time -v`` as a whole-child process signal; setup and
oracle work are included and no operation-level peak claim is made.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
RETAINED_0707 = HERE.parent / "change-0707" / "analyze_profiles.py"

OWNER = "litchi_xlsx::cell_values::source::SourceBackedEditor::edit_sheets"
LIFECYCLE_PARENT = "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates"
MEASURED_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
TIMING_FIELDS = frozenset(
    {"open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns"}
)
PERL_ENV = {"PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"}
RSS_RE = re.compile(r"^Maximum resident set size \(kbytes\):\s*([0-9][0-9,]*)\s*$")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory mechanism artifact."""


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


def digest_json(value: Any) -> str:
    encoded = (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()
    return hashlib.sha256(encoded).hexdigest()


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


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path escapes mechanism packet: {path}") from error


def load_retained_helpers() -> Any:
    require(RETAINED_0707.is_file(), f"missing retained helper: {RETAINED_0707}")
    spec = importlib.util.spec_from_file_location("retained_0707_mechanism_helpers", RETAINED_0707)
    require(spec is not None and spec.loader is not None, "cannot load retained 0707 analyzer")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    helper = module.H
    helper.HERE = HERE
    return helper


H = load_retained_helpers()


def plan_data() -> tuple[dict[str, Any], dict[str, Any]]:
    main = read_json(HERE / "plan.json")
    mechanism = read_json(HERE / "mechanism-plan.json")
    require(isinstance(main, dict) and isinstance(mechanism, dict), "plans are not objects")
    require(main.get("revision") == mechanism.get("revision"), "plan revisions differ")
    require(isinstance(main.get("primary"), dict), "primary native plan is missing")
    require(main["primary"].get("case") == mechanism.get("case"), "plan cases differ")
    require(main["primary"].get("shapes") == mechanism.get("shapes"), "plan shapes differ")
    primary = main.get("primary")
    require(isinstance(primary, dict), "primary native plan is missing")
    require(primary.get("repeats") == 2 and primary.get("samples") == 200
            and primary.get("warmup") == 20, "primary native plan differs")
    profile = mechanism.get("profile")
    rss = mechanism.get("rss")
    require(isinstance(profile, dict) and profile.get("repeats") == 2
            and profile.get("samples") == 1 and profile.get("warmup") == 0,
            "mechanism profile plan differs")
    require(profile.get("owner") == OWNER,
            "mechanism profile owner differs")
    require(profile.get("lifecycle_parent") == LIFECYCLE_PARENT
            and profile.get("measured_parent") == MEASURED_PARENT,
            "mechanism profile parent differs")
    require(profile.get("numbered_parts") == [1, 2, 3, 4],
            "mechanism numbered dump policy differs")
    require(profile.get("expected_lifecycle_calls") == 3
            and profile.get("expected_measured_calls") == 1,
            "mechanism call-count policy differs")
    require(isinstance(rss, dict) and rss.get("repeats") == 2
            and rss.get("samples") == 3 and rss.get("warmup") == 2,
            "mechanism RSS plan differs")
    require(mechanism.get("cpu") == 12, "mechanism CPU differs")
    options = profile.get("options")
    require(isinstance(options, list) and all(isinstance(item, str) for item in options),
            "mechanism Callgrind options are invalid")
    for token in (
        "--tool=callgrind",
        "--collect-atstart=no",
        "--toggle-collect=" + OWNER,
        "--zero-before=" + OWNER,
        "--dump-after=" + OWNER,
    ):
        require(token in options, f"mechanism profile omits {token}")
    return main, mechanism


def source_census() -> dict[str, str]:
    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path for path in (REPO / folder).rglob("*")
            if path.is_file()
            and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {str(path.relative_to(REPO)): sha(path) for path in sorted(set(paths))}


def validate_constraints() -> str:
    path = HERE / "constraints.json"
    constraints = read_json(path)
    require(isinstance(constraints, dict), "constraints are not an object")
    for name, expected in constraints.items():
        source = REPO / str(name)
        require(source.is_file() and sha(source) == expected,
                f"constraint changed: {name}")
    return sha(path)


def binary_custody(path: Path, expected_sha: str, expected_bytes: int | None,
                   stage: str) -> dict[str, Any]:
    """Verify bytes now or the exact cleanup witness, with stable identity."""

    if path.is_file() and not path.is_symlink():
        require(sha(path) == expected_sha, f"{stage}: binary hash differs")
        actual_bytes = path.stat().st_size
        require(expected_bytes is None or actual_bytes == expected_bytes,
                f"{stage}: binary byte count differs")
        return {"path": str(path), "sha256": expected_sha, "bytes": actual_bytes,
                "verified": True}
    cleanup_path = HERE / "cleanup.json"
    require(cleanup_path.is_file(), f"{stage}: binary missing without cleanup witness")
    cleanup = read_json(cleanup_path)
    require(cleanup.get("owned_paths_absent") is True,
            f"{stage}: cleanup witness does not establish owned paths absent")
    witnesses = cleanup.get("binaries")
    require(isinstance(witnesses, list), f"{stage}: cleanup binary witness is not a list")
    matches = [item for item in witnesses if isinstance(item, dict)
               and item.get("path") == str(path)]
    require(len(matches) == 1, f"{stage}: cleanup witness does not bind {path}")
    witness = matches[0]
    require(witness.get("sha256") == expected_sha
            and isinstance(witness.get("bytes"), int)
            and witness["bytes"] > 0, f"{stage}: cleanup binary identity differs")
    return {"path": str(path), "sha256": expected_sha, "bytes": witness["bytes"],
            "verified": True}


def build_data(stage: str) -> dict[str, Any]:
    records_path = HERE / f"build-{stage}.json"
    source_path = HERE / f"source-{stage}.json"
    records = read_json(records_path)
    expected = read_json(source_path)
    require(isinstance(records, list) and isinstance(expected, dict),
            f"{stage}: build/source records are malformed")
    matches = [item for item in records if isinstance(item, dict)
               and Path(str(item.get("binary", ""))).name == f"{stage}-native"]
    require(len(matches) == 1, f"{stage}: expected one native build record")
    record = matches[0]
    require(record.get("exit_code") == 0, f"{stage}: native build failed")
    source_sha = sha(source_path)
    require(record.get("source_manifest_sha256") == source_sha,
            f"{stage}: source manifest binding differs")
    binary = Path(str(record.get("binary", ""))).resolve()
    binary_sha = record.get("binary_sha256")
    require(isinstance(binary_sha, str) and len(binary_sha) == 64,
            f"{stage}: binary hash is missing")
    return {
        "stage": stage,
        "record": rel(records_path),
        "record_sha256": sha(records_path),
        "source": rel(source_path),
        "source_sha256": source_sha,
        "expected_source": expected,
        "binary": binary,
        "binary_sha256": binary_sha,
        "binary_bytes": record.get("binary_bytes"),
        "binary_identity": binary_custody(binary, binary_sha,
                                           record.get("binary_bytes"), stage),
    }


def validate_source_receipt(receipt: dict[str, Any], build: dict[str, Any],
                            main: dict[str, Any], stage: str, label: str) -> dict[str, Any]:
    custody = receipt.get("current_checkout_source")
    require(isinstance(custody, dict), f"{label}: source custody is missing")
    before_name = custody.get("before_artifact")
    after_name = custody.get("after_artifact")
    require(isinstance(before_name, str) and isinstance(after_name, str),
            f"{label}: source custody artifact names are missing")
    before_path = HERE / before_name
    after_path = HERE / after_name
    for name, path in ((before_name, before_path), (after_name, after_path)):
        require(Path(name).name == name and not Path(name).is_absolute()
                and path.is_file() and not path.is_symlink(),
                f"{label}: source custody artifact is invalid: {path}")
    before = read_json(before_path)
    after = read_json(after_path)
    expected = build["expected_source"]
    allowed_roots = tuple(main.get("candidate_roots", ()))
    if stage == "candidate":
        require(before == expected and after == expected,
                f"{label}: candidate source census differs")
        mode = "exact"
    else:
        candidate_path = HERE / "source-candidate.json"
        candidate = read_json(candidate_path) if candidate_path.is_file() else None
        require(before == expected or (isinstance(candidate, dict) and before == candidate),
                f"{label}: baseline binary ran under an unbound source state")
        require(after == expected or (isinstance(candidate, dict) and after == candidate),
                f"{label}: baseline source after-census is unbound")
        changed = [name for name in set(expected) | set(before)
                   if expected.get(name) != before.get(name)]
        require(all(any(name.startswith(root) for root in allowed_roots) for name in changed),
                f"{label}: baseline source delta escapes candidate roots")
        mode = "exact" if not changed else "baseline-retained-under-allowed-candidate-delta"
    require(before == after, f"{label}: source changed during child")
    require(custody.get("unchanged_during_child") is True,
            f"{label}: source changed during child")
    require(custody.get("before_sha256") == digest_json(before)
            and custody.get("after_sha256") == digest_json(after)
            and custody.get("before_file_sha256") == sha(before_path)
            and custody.get("after_file_sha256") == sha(after_path),
            f"{label}: source custody digest differs")
    relation = custody.get("relation_before")
    relation_after = custody.get("relation_after")
    require(isinstance(relation, dict) and relation.get("mode") == mode
            and isinstance(relation_after, dict) and relation_after.get("mode") == mode
            and relation_after.get("changed_paths") == relation.get("changed_paths"),
            f"{label}: source relation mode differs")
    return {"before": rel(before_path), "after": rel(after_path),
            "before_sha256": sha(before_path), "after_sha256": sha(after_path),
            "source_equal": True, "mode": mode,
            "changed_paths": relation.get("changed_paths", [])}


def validate_artifacts(receipt: dict[str, Any], required: set[str], label: str) -> None:
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifact map is missing")
    require(required.issubset(artifacts),
            f"{label}: required artifacts omitted: {sorted(required - set(artifacts))}")
    for name, expected in artifacts.items():
        path = HERE / str(name)
        require(path.name == str(name) and path.is_file() and not path.is_symlink(),
                f"{label}: artifact path is invalid: {name}")
        require(sha(path) == expected, f"{label}: artifact hash differs: {name}")


def validate_common_receipt(receipt: dict[str, Any], build: dict[str, Any],
                            main: dict[str, Any], mechanism: dict[str, Any],
                            stage: str, kind: str, repeat: int, shape: str,
                            constraints_sha: str) -> dict[str, Any]:
    label = f"{receipt.get('name', '<unknown>')}.receipt.json"
    require(receipt.get("schema_version") == 1 and receipt.get("exit_code") == 0,
            f"{label}: receipt failed or schema differs")
    require(receipt.get("role") == stage and receipt.get("kind") == kind
            and receipt.get("repeat") == repeat and receipt.get("shape") == shape,
            f"{label}: lane identity differs")
    require(receipt.get("cpu") == mechanism["cpu"], f"{label}: CPU differs")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{label}: binary binding differs")
    require(receipt.get("build_record") == build["record"]
            and receipt.get("build_record_sha256") == build["record_sha256"],
            f"{label}: build record binding differs")
    require(receipt.get("build_source_manifest") == build["source"]
            and receipt.get("build_source_manifest_sha256") == build["source_sha256"],
            f"{label}: source manifest binding differs")
    require(receipt.get("retained_source_census_sha256") == digest_json(build["expected_source"])
            and receipt.get("retained_source_entry_count") == len(build["expected_source"]),
            f"{label}: retained source identity differs")
    retained = receipt.get("retained_binary_source")
    require(isinstance(retained, dict)
            and retained.get("manifest") == build["source"]
            and retained.get("manifest_sha256") == build["source_sha256"]
            and retained.get("source_census_sha256") == digest_json(build["expected_source"]),
            f"{label}: retained binary source binding differs")
    require(receipt.get("plan_sha256") == sha(HERE / "plan.json")
            and receipt.get("mechanism_plan_sha256") == sha(HERE / "mechanism-plan.json"),
            f"{label}: plan binding differs")
    require(receipt.get("script_sha256") == sha(HERE / "mechanism.py"),
            f"{label}: capture script binding differs")
    require(receipt.get("constraints_sha256") == constraints_sha,
            f"{label}: constraints binding differs")
    overflow = receipt.get("brk_segment_overflow")
    require(isinstance(overflow, dict)
            and isinstance(overflow.get("present"), bool)
            and isinstance(overflow.get("evidence"), list),
            f"{label}: brk-segment overflow observation is missing")
    custody = validate_source_receipt(receipt, build, main, stage, label)
    return {"label": label, "source_custody": custody}


def option(command: list[str], name: str) -> str:
    for token in command:
        if token.startswith(name + "="):
            return token.split("=", 1)[1]
    try:
        index = command.index(name)
        return command[index + 1]
    except (ValueError, IndexError) as error:
        raise EvidenceError(f"command omits {name}") from error


def canonical(value: Any, label: str) -> Any:
    if isinstance(value, list):
        require(bool(value), f"{label}: empty identity vector")
        values = [canonical(item, f"{label}[{index}]") for index, item in enumerate(value)]
        require(all(item == values[0] for item in values),
                f"{label}: per-sample identity values differ")
        return values[0]
    if isinstance(value, dict):
        return {key: canonical(item, f"{label}.{key}") for key, item in sorted(value.items())}
    return value


def source_identity(source: Any, label: str) -> Any:
    require(isinstance(source, dict), f"{label}: source is not an object")
    result: dict[str, Any] = {}
    for key, value in sorted(source.items()):
        if key == "xlsx_cell_values":
            require(isinstance(value, dict), f"{label}.{key}: not an object")
            nested: dict[str, Any] = {}
            for nested_key, nested_value in sorted(value.items()):
                if nested_key in TIMING_FIELDS or nested_key.endswith("allocation_metrics"):
                    continue
                nested[nested_key] = canonical(nested_value, f"{label}.{key}.{nested_key}")
            result[key] = nested
        else:
            result[key] = canonical(value, f"{label}.{key}")
    return result


def result_identity(report: dict[str, Any], label: str) -> dict[str, Any]:
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{label}: expected exactly one result")
    result = results[0]
    require(isinstance(result, dict)
            and all(key in result for key in ("case", "corpus", "sink", "source")),
            f"{label}: result identity is incomplete")
    return {
        "case": result["case"],
        "corpus": result["corpus"],
        "sink": result["sink"],
        "source": source_identity(result["source"], label + ".source"),
        "output_sha256": result.get("output_sha256"),
    }


def validate_result(path: Path, build: dict[str, Any], main: dict[str, Any],
                    shape: str, samples: int, warmup: int,
                    label: str) -> tuple[dict[str, Any], dict[str, Any]]:
    report = read_json(path)
    require(isinstance(report, dict) and report.get("schema_version") == 1,
            f"{label}: report schema differs")
    tool = report.get("tool")
    require(isinstance(tool, dict) and tool.get("binary") == "litchi-perf-baseline"
            and tool.get("profile") == "release"
            and tool.get("instrumentation") == "none",
            f"{label}: tool identity differs")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict)
            and identity.get("binary_sha256") == build["binary_sha256"],
            f"{label}: report binary identity differs")
    environment = report.get("environment")
    require(isinstance(environment, dict)
            and environment.get("git_revision") == main["revision"],
            f"{label}: report revision differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict)
            and configuration.get("cases") == [main["primary"]["case"]]
            and configuration.get("xlsx_cell_crud_shapes") == [shape]
            and configuration.get("samples_per_case") == samples
            and configuration.get("warmup_iterations_per_case") == warmup,
            f"{label}: report configuration differs")
    identity_value = result_identity(report, label)
    return report, identity_value


def native_reference(stage: str, repeat: int, shape: str, build: dict[str, Any],
                     main: dict[str, Any], mechanism: dict[str, Any]) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    phase = mechanism["native_reference"]["phase_by_stage"][stage]
    name = f"native-{phase}-r{repeat}-{shape}"
    result_path = HERE / f"{name}.json"
    receipt_path = HERE / f"{name}.receipt.json"
    require(result_path.is_file() and receipt_path.is_file(),
            f"{name}: 0708 ABBA native output is missing")
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict) and receipt.get("exit_code") == 0,
            f"{name}: native reference failed")
    require(receipt.get("binary_sha256") == build["binary_sha256"]
            and receipt.get("role") == stage and receipt.get("phase") == phase
            and receipt.get("kind") == "primary" and receipt.get("shape") == shape,
            f"{name}: native reference custody differs")
    _, identity = validate_result(
        result_path, build, main, shape,
        int(main["primary"]["samples"]), int(main["primary"]["warmup"]),
        name,
    )
    return result_path, receipt, identity


def numbered_dumps(stem: Path, expected_parts: list[int]) -> list[Path]:
    paths = [
        path for path in HERE.glob(stem.name + ".*")
        if path.is_file() and path.suffix[1:].isdigit()
    ]
    paths.sort(key=lambda path: int(path.suffix[1:]))
    require([int(path.suffix[1:]) for path in paths] == expected_parts,
            f"{stem.name}: numbered raw dumps differ")
    return paths


def classify_dumps(paths: list[Path], profile: dict[str, Any]) -> tuple[list[dict[str, Any]], Path]:
    rows: list[dict[str, Any]] = []
    measured: list[Path] = []
    lifecycle: list[Path] = []
    for path in paths:
        label = rel(path)
        text = read_text(path)
        part = int(path.suffix[1:])
        require(H.part_number(text, label) == part, f"{label}: part number differs")
        require(H.trigger(text, label) == "--dump-after=" + OWNER,
                f"{label}: dump trigger differs")
        require("events: Ir" in text.splitlines(), f"{label}: Ir event is missing")
        summary = H.summary_ir(text, label)
        all_incoming = H.target_edge_summary(path, OWNER)
        measured_edge = H.target_edge_summary(path, OWNER, MEASURED_PARENT)
        lifecycle_edge = H.target_edge_summary(path, OWNER, LIFECYCLE_PARENT)
        require(all_incoming["calls"] == 1,
                f"{label}: owner incoming call count differs")
        if measured_edge["calls"]:
            require(measured_edge["calls"] == 1 and lifecycle_edge["calls"] == 0,
                    f"{label}: measured/lifecycle owner edges overlap")
            require(measured_edge["inclusive_ir"] == summary,
                    f"{label}: measured owner Ir differs from summary")
            measured.append(path)
            role = "measured"
        else:
            require(lifecycle_edge["calls"] == 1 and measured_edge["calls"] == 0,
                    f"{label}: lifecycle owner edge is missing or duplicated")
            require(lifecycle_edge["inclusive_ir"] == summary,
                    f"{label}: lifecycle owner Ir differs from summary")
            lifecycle.append(path)
            role = "lifecycle"
        rows.append({
            "file": label,
            "part": part,
            "sha256": sha(path),
            "summary_ir": summary,
            "role": role,
            "owner_incoming": all_incoming,
            "selected_by_raw_incoming_edge": True,
        })
    require(len(lifecycle) == profile["expected_lifecycle_calls"],
            f"expected {profile['expected_lifecycle_calls']} lifecycle dumps")
    require(len(measured) == profile["expected_measured_calls"],
            f"expected {profile['expected_measured_calls']} measured dumps")
    return rows, measured[0]


def validate_termination(path: Path, numbered: list[Path]) -> dict[str, Any]:
    label = rel(path)
    require(path.is_file(), f"{label}: termination dump is missing")
    text = read_text(path)
    require(H.part_number(text, label) == int(numbered[-1].suffix[1:]) + 1,
            f"{label}: termination part differs")
    require(H.trigger(text, label) == "Program termination",
            f"{label}: termination trigger differs")
    require(H.summary_ir(text, label) == 0, f"{label}: termination Ir is nonzero")
    return {"file": label, "part": int(numbered[-1].suffix[1:]) + 1,
            "sha256": sha(path), "summary_ir": 0, "trigger": "Program termination"}


def write_or_check(path: Path, text: str) -> None:
    if path.exists():
        require(read_text(path) == text, f"annotation replay differs: {rel(path)}")
    else:
        path.write_text(text, encoding="utf-8")


def annotate(selected: Path, name: str, summary: int) -> dict[str, Any]:
    inclusive_text, inclusive_command = H.run_annotation(selected, True)
    self_text, self_command = H.run_annotation(selected, False)
    inclusive_path = HERE / f"{name}.inclusive.txt"
    self_path = HERE / f"{name}.self.txt"
    write_or_check(inclusive_path, inclusive_text)
    write_or_check(self_path, self_text)
    inclusive = H.parse_annotation(inclusive_text, OWNER, rel(inclusive_path))
    exclusive = H.parse_annotation(self_text, OWNER, rel(self_path))
    require(inclusive["selected_ir"] == summary,
            f"{name}: annotation owner Ir differs from raw summary")
    require(exclusive["direct"] == inclusive["direct"],
            f"{name}: annotation direct children differ")
    direct_ir = sum(item["inclusive_ir"] for item in inclusive["direct"])
    require(exclusive["selected_ir"] + direct_ir == inclusive["selected_ir"],
            f"{name}: owner partition does not reconstruct inclusive Ir")
    direct = [{"name": item["name"], "inclusive_ir": item["inclusive_ir"]}
              for item in inclusive["direct"]]
    nested: dict[str, Any] = {}
    for target in (
        "litchi_xlsx::cell_values::validation::Validator::observe",
        "litchi_xlsx::cell_values::validation::validate_element",
    ):
        nested[target] = H.target_edge_summary(selected, target)
    return {
        "environment": dict(PERL_ENV),
        "command": {"inclusive": inclusive_command, "self": self_command},
        "files": {
            "inclusive": rel(inclusive_path), "inclusive_sha256": sha(inclusive_path),
            "self": rel(self_path), "self_sha256": sha(self_path),
        },
        "owner": {"inclusive_ir": inclusive["selected_ir"],
                  "self_ir": exclusive["selected_ir"],
                  "direct_callees": direct},
        "immediate_child_partition": {
            "self_ir": exclusive["selected_ir"],
            "direct_children_ir": direct_ir,
            "owner_ir": inclusive["selected_ir"],
            "disjoint": True,
            "nested_inclusive_costs_excluded": True,
        },
        "nested_targets": nested,
        "allocation_call_interpretation": (
            "disabled; raw allocator edges are not parsed, summed, or interpreted"
        ),
    }


def profile_name(stage: str, repeat: int, shape: str) -> str:
    return f"mechanism-profile-{stage}-r{repeat}-{shape}"


def analyze_profile(stage: str, repeat: int, shape: str, build: dict[str, Any],
                    main: dict[str, Any], mechanism: dict[str, Any],
                    constraints_sha: str) -> dict[str, Any]:
    name = profile_name(stage, repeat, shape)
    receipt_path = HERE / f"{name}.receipt.json"
    result_path = HERE / f"{name}.json"
    require(receipt_path.is_file() and result_path.is_file(), f"{name}: profile artifacts missing")
    receipt = read_json(receipt_path)
    common = validate_common_receipt(receipt, build, main, mechanism, stage, "profile",
                                     repeat, shape, constraints_sha)
    command = receipt.get("command")
    require(isinstance(command, list)
            and command[:4] == ["taskset", "-c", str(mechanism["cpu"]), "valgrind"],
            f"{name}: profile command prefix differs")
    for token in mechanism["profile"]["options"]:
        require(token in command, f"{name}: profile option omitted: {token}")
    require(option(command, "--warmup") == "0" and option(command, "--samples") == "1"
            and option(command, "--case") == main["primary"]["case"]
            and option(command, "--xlsx-cell-crud-shape") == shape,
            f"{name}: profile command matrix differs")
    require(Path(option(command, "--json")).name == result_path.name
            and Path(option(command, "--callgrind-out-file")).name == f"{name}.callgrind",
            f"{name}: profile output paths differ")
    numbered_names = {f"{name}.callgrind.{part}" for part in [1, 2, 3, 4]}
    required = {f"{name}{suffix}" for suffix in (".json", ".stdout", ".stderr", ".callgrind")}
    required.update(numbered_names)
    required.update({common["source_custody"]["before"], common["source_custody"]["after"]})
    validate_artifacts(receipt, required, f"{name}.receipt.json")
    _, identity = validate_result(result_path, build, main, shape, 1, 0, name)
    _, native_receipt, native_identity = native_reference(
        stage, repeat, shape, build, main, mechanism
    )
    require(identity == native_identity, f"{name}: profile/native result identity differs")
    stem = HERE / f"{name}.callgrind"
    dumps = numbered_dumps(stem, mechanism["profile"]["numbered_parts"])
    rows, selected = classify_dumps(dumps, mechanism["profile"])
    termination = validate_termination(stem, dumps)
    selected_row = next(row for row in rows if row["file"] == rel(selected))
    annotations = annotate(selected, name, selected_row["summary_ir"])
    return {
        "name": name,
        "stage": stage,
        "repeat": repeat,
        "shape": shape,
        "receipt": rel(receipt_path),
        "receipt_sha256": sha(receipt_path),
        "result": rel(result_path),
        "result_sha256": sha(result_path),
        "source_custody": common["source_custody"],
        "brk_segment_overflow": receipt["brk_segment_overflow"],
        "native_reference": {
            "result": native_receipt["native_reference"]["result"]
            if isinstance(native_receipt.get("native_reference"), dict)
            else f"native-{mechanism['native_reference']['phase_by_stage'][stage]}-r{repeat}-{shape}.json",
            "native_result_identity": native_identity,
        },
        "result_identity": identity,
        "parts": rows,
        "termination": termination,
        "selected": rel(selected),
        "selected_by": {
            "caller": MEASURED_PARENT,
            "positive_owner_call_count": 1,
            "raw_incoming_edge": True,
        },
        "annotations": annotations,
        "total_ir": annotations["owner"]["inclusive_ir"],
        "self_ir": annotations["owner"]["self_ir"],
        "direct_ir": annotations["immediate_child_partition"]["direct_children_ir"],
        "allocation_call_interpretation": "disabled",
    }


def rss_name(stage: str, repeat: int, shape: str) -> str:
    return f"mechanism-rss-{stage}-r{repeat}-{shape}"


def parse_rss(path: Path, label: str) -> int:
    matches = [int(match.group(1).replace(",", ""))
               for match in (RSS_RE.match(line) for line in read_text(path).splitlines())
               if match]
    require(len(matches) == 1, f"{label}: expected one GNU time RSS record")
    return matches[0]


def analyze_rss(stage: str, repeat: int, shape: str, build: dict[str, Any],
                main: dict[str, Any], mechanism: dict[str, Any],
                constraints_sha: str) -> dict[str, Any]:
    name = rss_name(stage, repeat, shape)
    receipt_path = HERE / f"{name}.receipt.json"
    result_path = HERE / f"{name}.json"
    time_path = HERE / f"{name}.time.txt"
    require(receipt_path.is_file() and result_path.is_file() and time_path.is_file(),
            f"{name}: RSS artifacts missing")
    receipt = read_json(receipt_path)
    common = validate_common_receipt(receipt, build, main, mechanism, stage, "rss",
                                     repeat, shape, constraints_sha)
    command = receipt.get("command")
    require(isinstance(command, list)
            and command[:6] == ["taskset", "-c", str(mechanism["cpu"]), "/usr/bin/time", "-v", "-o"],
            f"{name}: RSS command prefix differs")
    require(Path(option(command, "-o")).name == time_path.name
            and option(command, "--warmup") == "2"
            and option(command, "--samples") == "3"
            and option(command, "--case") == main["primary"]["case"]
            and option(command, "--xlsx-cell-crud-shape") == shape
            and Path(option(command, "--json")).name == result_path.name,
            f"{name}: RSS command matrix differs")
    required = {f"{name}{suffix}" for suffix in (".json", ".stdout", ".stderr", ".time.txt")}
    required.update({common["source_custody"]["before"], common["source_custody"]["after"]})
    validate_artifacts(receipt, required, f"{name}.receipt.json")
    _, identity = validate_result(result_path, build, main, shape, 3, 2, name)
    _, native_receipt, native_identity = native_reference(
        stage, repeat, shape, build, main, mechanism
    )
    require(identity == native_identity, f"{name}: RSS/native result identity differs")
    rss_kib = parse_rss(time_path, rel(time_path))
    return {
        "name": name,
        "stage": stage,
        "repeat": repeat,
        "shape": shape,
        "receipt": rel(receipt_path),
        "receipt_sha256": sha(receipt_path),
        "result": rel(result_path),
        "result_sha256": sha(result_path),
        "time_output": rel(time_path),
        "time_output_sha256": sha(time_path),
        "source_custody": common["source_custody"],
        "brk_segment_overflow": receipt["brk_segment_overflow"],
        "native_reference": {
            "result": native_receipt["native_reference"]["result"]
            if isinstance(native_receipt.get("native_reference"), dict)
            else f"native-{mechanism['native_reference']['phase_by_stage'][stage]}-r{repeat}-{shape}.json",
            "native_result_identity": native_identity,
        },
        "result_identity": identity,
        "maximum_resident_set_size_kib": rss_kib,
        "scope": mechanism["rss"]["scope"],
        "operation_peak_claim": False,
    }


def reduction_percent(baseline: int, candidate: int) -> float:
    require(baseline > 0, "baseline planning Ir is zero")
    return (baseline - candidate) * 100.0 / baseline


def compare_profiles(rows: list[dict[str, Any]], mechanism: dict[str, Any]) -> dict[str, Any]:
    indexed = {(row["repeat"], row["shape"], row["stage"]): row for row in rows}
    comparisons: list[dict[str, Any]] = []
    gate = float(mechanism["gates"]["planning_ir_reduction_percent"])
    for repeat in range(1, int(mechanism["profile"]["repeats"]) + 1):
        for shape in mechanism["shapes"]:
            baseline = indexed[(repeat, shape, "baseline")]
            candidate = indexed[(repeat, shape, "candidate")]
            reduction = reduction_percent(baseline["total_ir"], candidate["total_ir"])
            comparisons.append({
                "repeat": repeat,
                "shape": shape,
                "baseline_total_ir": baseline["total_ir"],
                "candidate_total_ir": candidate["total_ir"],
                "planning_ir_reduction_percent": reduction,
                "threshold_percent": gate,
                "passes": reduction >= gate,
            })
    return {
        "threshold_percent": gate,
        "require_every_shape_repeat": True,
        "comparisons": comparisons,
        "passes": all(row["passes"] for row in comparisons),
        "claim": "No production or native benefit claim follows from this gate alone.",
    }


def compare_rss(rows: list[dict[str, Any]], mechanism: dict[str, Any]) -> dict[str, Any]:
    indexed = {(row["repeat"], row["shape"], row["stage"]): row for row in rows}
    threshold = float(mechanism["rss"]["review_threshold_percent"])
    comparisons: list[dict[str, Any]] = []
    for repeat in range(1, int(mechanism["rss"]["repeats"]) + 1):
        for shape in mechanism["shapes"]:
            baseline = indexed[(repeat, shape, "baseline")]["maximum_resident_set_size_kib"]
            candidate = indexed[(repeat, shape, "candidate")]["maximum_resident_set_size_kib"]
            delta = 0.0 if baseline == 0 else (candidate - baseline) * 100.0 / baseline
            comparisons.append({
                "repeat": repeat,
                "shape": shape,
                "baseline_kib": baseline,
                "candidate_kib": candidate,
                "candidate_minus_baseline_percent": delta,
                "review_threshold_percent": threshold,
                "review_flag": abs(delta) > threshold,
            })
    return {
        "threshold_percent": threshold,
        "comparisons": comparisons,
        "review_flags": sum(1 for row in comparisons if row["review_flag"]),
        "scope": mechanism["rss"]["scope"],
        "operation_peak_claim": False,
    }


def analyze(output: Path) -> dict[str, Any]:
    main, mechanism = plan_data()
    constraints_sha = validate_constraints()
    builds = {stage: build_data(stage) for stage in ("baseline", "candidate")}
    profile_rows = [
        analyze_profile(stage, repeat, shape, builds[stage], main, mechanism, constraints_sha)
        for stage in ("baseline", "candidate")
        for repeat in range(1, int(mechanism["profile"]["repeats"]) + 1)
        for shape in mechanism["shapes"]
    ]
    rss_rows = [
        analyze_rss(stage, repeat, shape, builds[stage], main, mechanism, constraints_sha)
        for stage in ("baseline", "candidate")
        for repeat in range(1, int(mechanism["rss"]["repeats"]) + 1)
        for shape in mechanism["shapes"]
    ]
    profile_comparison = compare_profiles(profile_rows, mechanism)
    rss_comparison = compare_rss(rss_rows, mechanism)
    report = {
        "schema_version": 1,
        "status": "pass" if profile_comparison["passes"] else "rejected",
        "revision": main["revision"],
        "case": main["primary"]["case"],
        "scope": "Conditional post-pilot mechanism evidence for XLSX source-backed planning",
        "plan": {
            "main": "plan.json",
            "main_sha256": sha(HERE / "plan.json"),
            "mechanism": "mechanism-plan.json",
            "mechanism_sha256": sha(HERE / "mechanism-plan.json"),
        },
        "custody": {
            "capture_script": "mechanism.py",
            "capture_script_sha256": sha(HERE / "mechanism.py"),
            "analyzer_script": "analyze_mechanism.py",
            "analyzer_script_sha256": sha(Path(__file__)),
            "constraints": "constraints.json",
            "constraints_sha256": constraints_sha,
            "builds": {
                stage: {
                    "record": data["record"],
                    "record_sha256": data["record_sha256"],
                    "source": data["source"],
                    "source_sha256": data["source_sha256"],
                    "binary_identity": data["binary_identity"],
                }
                for stage, data in builds.items()
            },
            "normalized_identity_replay": (
                "Build paths, hashes, and byte counts are retained; whether the owned binary "
                "is present or represented by cleanup.json is excluded from normalized identity."
            ),
        },
        "profiles": profile_rows,
        "profile_comparison": profile_comparison,
        "rss": rss_rows,
        "rss_comparison": rss_comparison,
        "allocation_call_interpretation": (
            "disabled; 0707 Callgrind allocation-call discrepancy unresolved"
        ),
        "native_parity": (
            "Every profile and RSS result was compared with its same-stage, same-repeat, "
            "same-shape 0708 ABBA native result after excluding timing and per-sample fields."
        ),
        "conditional_admission": (
            "The primary profile gate requires at least 3% planning Ir reduction for every "
            "shape and repeat. A rejected gate carries no benefit or shipping claim."
        ),
    }
    output.write_text(json.dumps(report, indent=2) + "\n", encoding="utf-8")
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path,
                        default=HERE / "mechanism-analysis.json")
    parser.add_argument("--require-pass", action="store_true",
                        help="return failure if every-shape/repeat Ir gate does not pass")
    args = parser.parse_args()
    try:
        report = analyze(args.output)
    except (OSError, EvidenceError, AssertionError, ValueError) as error:
        print(f"mechanism analysis failed: {error}", file=sys.stderr)
        return 1
    print(f"mechanism analysis {report['status']}: {args.output}")
    return 0 if report["status"] == "pass" or not args.require_pass else 2


if __name__ == "__main__":
    raise SystemExit(main())
