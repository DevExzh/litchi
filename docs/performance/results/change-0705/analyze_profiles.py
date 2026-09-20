#!/usr/bin/env python3
"""Validate and attribute the current-head XLSX publication profiles.

Each child contains the three lifecycle publications followed by the one
timed publication.  The analyzer keeps every Callgrind dump and selects the
timed dump from its positive raw edge into the direct edit/save runner.  The
selection therefore remains valid if Callgrind changes dump numbering.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import re
import subprocess
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
HELPER = HERE.parent / "change-0521" / "analyze_profiles.py"


def load_helper() -> Any:
    spec = importlib.util.spec_from_file_location("retained_0521_profile_helpers", HELPER)
    require(spec is not None and spec.loader is not None, f"cannot load {HELPER}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    module.HERE = HERE
    return module


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


H = load_helper()


def display_name(text: str) -> str:
    # Strip only the final executable suffix.  Rust symbols can contain
    # bracketed slice types before that suffix.
    text = text.rsplit(" [", 1)[0].strip()
    text = re.sub(r"\s+\([\d,]+x\)$", "", text)
    if text.startswith("???:"):
        return text[4:]
    return text.rsplit(":", 1)[-1] if text.startswith(("./", "/")) else text


H.display_name = display_name

TOPOLOGY_OWNER = (
    "litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream"
)
VALIDATE_SOURCE_PART_XML = "litchi_opc::source_backed::validate_source_part_xml"
WRITE_CHANGED_OVERLAYS = (
    "litchi_opc::source_backed::SourceBackedPackage::"
    "write_changed_overlays_with_appended_inner"
)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def relative(path: Path) -> str:
    return str(path.relative_to(HERE))


def option(command: list[str], name: str) -> str:
    for token in command:
        if token.startswith(name + "="):
            return token.split("=", 1)[1]
    try:
        index = command.index(name)
        return command[index + 1]
    except (ValueError, IndexError) as error:
        raise AssertionError(f"command omits {name}") from error


def profile_plan() -> dict[str, Any]:
    plan = read_json(HERE / "profile-plan.json")
    require(isinstance(plan, dict), "profile-plan.json is not an object")
    require(plan.get("owner") == (
        "litchi_xlsx::cell_values::source::SourceBackedEditor::"
        "publish_multi_commit_to_stream"
    ), "profile owner differs")
    require(plan.get("shapes") == ["medium", "dense-sparse"], "profile shapes differ")
    require(plan.get("repeats") == 2, "profile repeat count differs")
    require(plan.get("warmup") == 0 and plan.get("samples") == 1,
            "profile warmup/sample policy differs")
    options = plan.get("options")
    require(isinstance(options, list) and all(isinstance(item, str) for item in options),
            "profile Callgrind options are invalid")
    for option_name in ("--tool=callgrind", "--collect-atstart=no",
                        "--toggle-collect=" + plan["owner"],
                        "--zero-before=" + plan["owner"],
                        "--dump-after=" + plan["owner"]):
        require(option_name in options, f"profile options omit {option_name}")
    require(plan.get("lifecycle_parent") ==
            "litchi_perf_baseline::run_xlsx_cell_value_lifecycle_gates",
            "lifecycle parent differs")
    require(plan.get("measured_parent") ==
            "litchi_perf_baseline::run_xlsx_cell_values_edit_save",
            "measured parent differs")
    require(isinstance(plan.get("expected_lifecycle_calls"), int) and
            plan["expected_lifecycle_calls"] > 0 and
            plan.get("expected_measured_calls") == 1,
            "profile call-count policy differs")
    return plan


def canonical(value: Any, label: str) -> Any:
    if isinstance(value, list):
        require(bool(value), f"{label}: empty vector")
        values = [canonical(item, f"{label}[{index}]")
                  for index, item in enumerate(value)]
        require(all(item == values[0] for item in values),
                f"{label}: vector is not constant")
        return values[0]
    if isinstance(value, dict):
        return {key: canonical(item, f"{label}.{key}")
                for key, item in sorted(value.items())}
    return value


TIMING_FIELDS = frozenset({"open_ns", "plan_ns", "commit_ns", "publication_ns", "reopen_ns"})


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
                nested[nested_key] = canonical(nested_value, f"{label}.{key}.{nested_key}")
            result[key] = nested
        else:
            result[key] = canonical(value, f"{label}.{key}")
    return result


def result_identity(result: Any, label: str) -> dict[str, Any]:
    require(isinstance(result, dict), f"{label}: result is not an object")
    require(all(key in result for key in ("corpus", "sink", "source")),
            f"{label}: result lacks corpus/sink/source")
    return {
        "corpus": result["corpus"],
        "sink": result["sink"],
        "source": source_identity(result["source"], label + ".source"),
        "output_sha256": result.get("output_sha256"),
    }


def validate_build(native: dict[str, Any]) -> dict[str, Any]:
    build = read_json(HERE / "build.json")
    require(build.get("exit_code") == 0, "build receipt did not pass")
    binary_sha = build.get("binary_sha256")
    require(isinstance(binary_sha, str) and len(binary_sha) == 64,
            "build binary hash is missing")
    binary = Path(build.get("binary", ""))
    if binary.is_file() and not binary.is_symlink():
        require(sha(binary) == binary_sha, "frozen native binary hash differs")
    else:
        # Final cleanup may remove the owned target after capture.  Preserve
        # source/binary custody through the cleanup witness so replay remains
        # independently checkable and emits the same report before/after it.
        cleanup_path = HERE / "cleanup.json"
        require(cleanup_path.is_file(),
                f"frozen native binary is missing without cleanup witness: {binary}")
        cleanup = read_json(cleanup_path)
        require(cleanup.get("owned_paths_absent") is True,
                "cleanup witness does not establish owned paths are absent")
        witnesses = cleanup.get("binaries")
        require(isinstance(witnesses, list), "cleanup witness binaries are not a list")
        matching = [item for item in witnesses
                    if isinstance(item, dict) and item.get("path") == str(binary)]
        require(len(matching) == 1 and matching[0].get("sha256") == binary_sha,
                "cleanup witness does not bind the removed native binary")
        require(isinstance(matching[0].get("bytes"), int) and matching[0]["bytes"] > 0,
                "cleanup witness binary size is invalid")
    require(build.get("revision") == native.get("revision"),
            "build and native plan revisions differ")
    return {
        "binary_sha256": binary_sha,
        "build_sha256": sha(HERE / "build.json"),
        "source_manifest_sha256": build.get("source_manifest_sha256"),
    }


def validate_receipt(name: str, receipt: dict[str, Any], build: dict[str, Any],
                     profile: dict[str, Any], native: dict[str, Any], shape: str) -> None:
    label = f"{name}.receipt.json"
    require(receipt.get("exit_code") == 0, f"{label}: child failed")
    require(receipt.get("binary_sha256") == build["binary_sha256"],
            f"{label}: binary binding differs")
    require(receipt.get("source_manifest_sha256") == build["source_manifest_sha256"],
            f"{label}: source binding differs")
    require(receipt.get("plan_sha256") == sha(HERE / "plan.json"),
            f"{label}: native plan binding differs")
    require(receipt.get("script_sha256") == sha(HERE / "capture.py"),
            f"{label}: capture script binding differs")
    command = receipt.get("command")
    require(isinstance(command, list), f"{label}: command is not a list")
    require(command[:4] == ["taskset", "-c", str(native["cpu"]), "valgrind"],
            f"{label}: command prefix differs")
    for token in profile["options"]:
        require(token in command, f"{label}: command omits {token}")
    require(option(command, "--warmup") == "0", f"{label}: warmup differs")
    require(option(command, "--samples") == "1", f"{label}: samples differs")
    require(option(command, "--case") == native["case"], f"{label}: case differs")
    require(option(command, "--xlsx-cell-crud-shape") == shape,
            f"{label}: shape differs")
    require(Path(option(command, "--json")).name == name + ".json",
            f"{label}: JSON output path differs")
    require(Path(option(command, "--callgrind-out-file")).name == name + ".callgrind",
            f"{label}: Callgrind output path differs")

    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{label}: artifacts are not an object")
    expected = {name + suffix for suffix in (".json", ".stdout", ".stderr", ".callgrind")}
    require(expected <= set(artifacts), f"{label}: required artifact omitted")
    for filename, digest in artifacts.items():
        path = HERE / filename
        require(Path(filename).name == filename and path.is_file() and not path.is_symlink(),
                f"{label}: artifact missing or escapes packet: {filename}")
        require(sha(path) == digest, f"{label}: artifact hash differs: {filename}")


def validate_result(path: Path, native_path: Path, build: dict[str, Any],
                    native: dict[str, Any], shape: str) -> dict[str, Any]:
    report = read_json(path)
    require(report.get("schema_version") == 1, f"{relative(path)}: schema differs")
    identity = report.get("binary_identity")
    require(isinstance(identity, dict) and identity.get("binary_sha256") == build["binary_sha256"],
            f"{relative(path)}: binary identity differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{relative(path)}: configuration missing")
    require(configuration.get("cases") == [native["case"]],
            f"{relative(path)}: case configuration differs")
    require(configuration.get("xlsx_cell_crud_shapes") == [shape],
            f"{relative(path)}: shape configuration differs")
    require(configuration.get("samples_per_case") == 1 and
            configuration.get("warmup_iterations_per_case") == 0,
            f"{relative(path)}: profile sample policy differs")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1,
            f"{relative(path)}: expected one result")
    profile_identity = result_identity(results[0], relative(path) + ".results[0]")
    native_report = read_json(native_path)
    native_results = native_report.get("results")
    require(isinstance(native_results, list) and len(native_results) == 1,
            f"{relative(native_path)}: expected one result")
    native_identity = result_identity(native_results[0], relative(native_path) + ".results[0]")
    require(profile_identity == native_identity,
            f"{relative(path)}: profile/native correctness identity differs")
    return profile_identity


def numbered_dumps(stem: Path) -> list[Path]:
    paths = [path for path in HERE.glob(stem.name + ".[0-9]*")
             if path.is_file() and path.suffix[1:].isdigit()]
    paths.sort(key=lambda path: int(path.suffix[1:]))
    require(paths, f"{stem.name}: no numbered Callgrind dumps")
    parts = [int(path.suffix[1:]) for path in paths]
    require(parts == list(range(1, len(parts) + 1)),
            f"{stem.name}: numbered dumps are not contiguous: {parts}")
    return paths


def edge(path: Path, owner: str, parent: str) -> dict[str, Any]:
    return H.target_edge_summary(path, owner, parent)


def classify_dumps(paths: list[Path], profile: dict[str, Any]) -> tuple[list[dict[str, Any]], Path]:
    owner = profile["owner"]
    lifecycle_parent = profile["lifecycle_parent"]
    measured_parent = profile["measured_parent"]
    rows: list[dict[str, Any]] = []
    measured: list[Path] = []
    lifecycle: list[Path] = []
    for path in paths:
        text = path.read_text(encoding="utf-8", errors="replace")
        part = int(path.suffix[1:])
        require(H.part_number(text, relative(path)) == part,
                f"{relative(path)}: Callgrind part does not match suffix")
        require(H.trigger(text, relative(path)) == f"--dump-after={owner}",
                f"{relative(path)}: unexpected dump trigger")
        require("events: Ir" in text.splitlines(), f"{relative(path)}: Ir event is absent")
        summary = H.summary_ir(text, relative(path))
        direct = edge(path, owner, measured_parent)
        life = edge(path, owner, lifecycle_parent)
        role: str
        if direct["positive_edge_count"]:
            require(direct["positive_edge_count"] == 1 and direct["calls"] == 1,
                    f"{relative(path)}: measured owner edge is not one call")
            require(direct["inclusive_ir"] == summary,
                    f"{relative(path)}: measured owner Ir differs from summary")
            require(life["positive_edge_count"] == 0,
                    f"{relative(path)}: dump has both measured and lifecycle parents")
            measured.append(path)
            role = "measured"
        else:
            require(life["positive_edge_count"] == 1 and life["calls"] == 1,
                    f"{relative(path)}: dump is not a single lifecycle publication")
            require(life["inclusive_ir"] == summary,
                    f"{relative(path)}: lifecycle owner Ir differs from summary")
            require(direct["positive_edge_count"] == 0,
                    f"{relative(path)}: lifecycle dump has a measured parent")
            lifecycle.append(path)
            role = "lifecycle"
        rows.append({
            "file": relative(path),
            "part": part,
            "sha256": sha(path),
            "summary_ir": summary,
            "role": role,
            "measured_parent": direct,
            "lifecycle_parent": life,
        })
    require(len(lifecycle) == profile["expected_lifecycle_calls"],
            f"expected {profile['expected_lifecycle_calls']} lifecycle dumps, got {len(lifecycle)}")
    require(len(measured) == profile["expected_measured_calls"],
            f"expected one measured dump, got {len(measured)}")
    return rows, measured[0]


def terminal_dump(path: Path, numbered: list[Path]) -> dict[str, Any]:
    require(path.is_file(), f"{relative(path)}: final Callgrind dump is missing")
    text = path.read_text(encoding="utf-8", errors="replace")
    part = H.part_number(text, relative(path))
    require(part > int(numbered[-1].suffix[1:]),
            f"{relative(path)}: final part does not follow numbered dumps")
    require(H.trigger(text, relative(path)) == "Program termination",
            f"{relative(path)}: final dump trigger differs")
    require(H.summary_ir(text, relative(path)) == 0,
            f"{relative(path)}: final process dump is not zero-Ir")
    require("events: Ir" in text.splitlines(), f"{relative(path)}: final Ir event is absent")
    return {"file": relative(path), "part": part, "sha256": sha(path),
            "summary_ir": 0, "trigger": "Program termination"}


def write_or_check(path: Path, text: str) -> None:
    if path.exists():
        require(path.read_text(encoding="utf-8") == text,
                f"annotation replay differs: {relative(path)}")
    else:
        path.write_text(text, encoding="utf-8")


def annotate(selected: Path, name: str, owner: str, summary: int) -> dict[str, Any]:
    inclusive_text, inclusive_command = H.run_annotation(selected, True)
    self_text, self_command = H.run_annotation(selected, False)
    inclusive_path = HERE / (name + ".inclusive.txt")
    self_path = HERE / (name + ".self.txt")
    write_or_check(inclusive_path, inclusive_text)
    write_or_check(self_path, self_text)
    inclusive = H.parse_annotation(inclusive_text, owner, relative(inclusive_path))
    exclusive = H.parse_annotation(self_text, owner, relative(self_path))
    require(inclusive["selected_ir"] == summary,
            f"{name}: annotation owner Ir differs from raw summary")
    require(exclusive["direct"] == inclusive["direct"],
            f"{name}: inclusive/self direct edges differ")
    require(exclusive["selected_ir"] + sum(item["inclusive_ir"] for item in exclusive["direct"])
            == inclusive["selected_ir"],
            f"{name}: self plus direct Ir does not reconstruct inclusive owner")
    return {
        "environment": {"PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0"},
        "command": {"inclusive": inclusive_command, "self": self_command},
        "files": {"inclusive": relative(inclusive_path),
                  "inclusive_sha256": sha(inclusive_path),
                  "self": relative(self_path), "self_sha256": sha(self_path)},
        "owner": {
            "inclusive_ir": inclusive["selected_ir"],
            "self_ir": exclusive["selected_ir"],
            "direct_callees": inclusive["direct"],
            "direct_callee_ir": H.direct_map(inclusive["direct"]),
        },
        "validation": {
            "inclusive_owner_matches_raw": True,
            "self_plus_direct_equals_inclusive": True,
            "inclusive_and_self_direct_children_match": True,
        },
    }


def attribute_topology(selected: Path, name: str, publication_owner: str,
                       publication_annotations: dict[str, Any]) -> dict[str, Any]:
    """Validate the physical-writer partition nested under publication.

    The raw edge parser aggregates duplicate function records for a symbol;
    this matters for ``validate_source_part_xml``, whose ten calls are split
    over two positive raw edges.  Immediate children of this one topology
    owner are the partition boundary.  Their nested callees are retained only
    through annotations and are never added to this direct-child total.
    """
    inclusive_text = (HERE / (name + ".inclusive.txt")).read_text(
        encoding="utf-8", errors="replace"
    )
    self_text = (HERE / (name + ".self.txt")).read_text(
        encoding="utf-8", errors="replace"
    )
    incoming = edge(selected, TOPOLOGY_OWNER, publication_owner)
    require(incoming["positive_edge_count"] == 1 and incoming["calls"] == 1,
            f"{name}: topology owner incoming edge is not one publication call")
    topology_ir = incoming["inclusive_ir"]
    inclusive = H.parse_annotation(inclusive_text, TOPOLOGY_OWNER, name + ".inclusive")
    exclusive = H.parse_annotation(self_text, TOPOLOGY_OWNER, name + ".self")
    require(inclusive["selected_ir"] == topology_ir,
            f"{name}: topology annotation Ir differs from raw incoming edge")
    require(exclusive["direct"] == inclusive["direct"],
            f"{name}: topology inclusive/self direct edges differ")
    require(exclusive["selected_ir"] + sum(item["inclusive_ir"]
                                            for item in exclusive["direct"])
            == inclusive["selected_ir"],
            f"{name}: topology self plus direct Ir does not reconstruct inclusive Ir")

    direct_children = []
    for child in inclusive["direct"]:
        raw_child = edge(selected, child["name"], TOPOLOGY_OWNER)
        require(raw_child["positive_edge_count"] > 0,
                f"{name}: raw edge missing for topology child {child['name']}")
        require(raw_child["calls"] == child["calls"],
                f"{name}: raw/annotation call count differs for {child['name']}")
        require(raw_child["inclusive_ir"] == child["inclusive_ir"],
                f"{name}: raw/annotation Ir differs for {child['name']}")
        direct_children.append({**child, "raw_edge": raw_child})

    by_name = {item["name"]: item for item in direct_children}
    for required_name, required_calls in {
        VALIDATE_SOURCE_PART_XML: 10,
        WRITE_CHANGED_OVERLAYS: 1,
    }.items():
        require(required_name in by_name,
                f"{name}: required topology child is absent: {required_name}")
        require(by_name[required_name]["calls"] == required_calls,
                f"{name}: required topology child call count differs: {required_name}")

    return {
        "owner": TOPOLOGY_OWNER,
        "incoming_from_publication": incoming,
        "inclusive_ir": inclusive["selected_ir"],
        "self_ir": exclusive["selected_ir"],
        "direct_ir": sum(item["inclusive_ir"] for item in direct_children),
        "direct_children": direct_children,
        "required_children": {
            VALIDATE_SOURCE_PART_XML: {"calls": 10,
                                       "inclusive_ir": by_name[VALIDATE_SOURCE_PART_XML]["inclusive_ir"]},
            WRITE_CHANGED_OVERLAYS: {"calls": 1,
                                     "inclusive_ir": by_name[WRITE_CHANGED_OVERLAYS]["inclusive_ir"]},
        },
        "partition_boundary": (
            "Immediate direct children of write_topology_to_stream. Nested child "
            "owners are overlapping diagnostics and are not added to this total."
        ),
        "validation": {
            "incoming_owner_edge_matches_annotation": True,
            "self_plus_direct_equals_inclusive": True,
            "every_direct_child_has_matching_raw_edge": True,
            "required_child_call_counts": True,
        },
    }


def analyze(output: Path) -> dict[str, Any]:
    profile = profile_plan()
    native = read_json(HERE / "plan.json")
    build = validate_build(native)
    rows: list[dict[str, Any]] = []
    for repeat in range(1, profile["repeats"] + 1):
        for shape in profile["shapes"]:
            name = f"profile-r{repeat}-{shape}"
            receipt = read_json(HERE / (name + ".receipt.json"))
            validate_receipt(name, receipt, build, profile, native, shape)
            result_identity_value = validate_result(
                HERE / (name + ".json"),
                HERE / f"native-r{repeat}-{shape}.json",
                build, native, shape,
            )
            stem = HERE / (name + ".callgrind")
            dumps = numbered_dumps(stem)
            parts, selected = classify_dumps(dumps, profile)
            termination = terminal_dump(stem, dumps)
            selected_part = next(item for item in parts if item["file"] == relative(selected))
            annotations = annotate(selected, name, profile["owner"], selected_part["summary_ir"])
            topology = attribute_topology(selected, name, profile["owner"], annotations)
            owner_data = annotations["owner"]
            rows.append({
                "name": name,
                "repeat": repeat,
                "shape": shape,
                "result_identity": result_identity_value,
                "parts": parts,
                "termination": termination,
                "selected": relative(selected),
                "selected_by": {
                    "caller": profile["measured_parent"],
                    "positive_owner_call_count": 1,
                    "selection_is_edge_based": True,
                },
                "annotations": annotations,
                "nested_topology": topology,
                "total_ir": owner_data["inclusive_ir"],
                "self_ir": owner_data["self_ir"],
                "direct_ir": sum(owner_data["direct_callee_ir"].values()),
                "direct_callee_ir": owner_data["direct_callee_ir"],
            })
    require(len(rows) == profile["repeats"] * len(profile["shapes"]),
            "profile matrix is incomplete")
    topology_total = sum(row["nested_topology"]["inclusive_ir"] for row in rows)
    child_totals: dict[str, int] = {}
    child_calls: dict[str, int] = {}
    for row in rows:
        for child in row["nested_topology"]["direct_children"]:
            child_totals[child["name"]] = child_totals.get(child["name"], 0) + child["inclusive_ir"]
            child_calls[child["name"]] = child_calls.get(child["name"], 0) + (child["calls"] or 0)
    require(child_totals, "topology direct-child partition is empty")
    dominant_name = max(child_totals, key=child_totals.get)
    result = {
        "status": "pass",
        "plan_sha256": sha(HERE / "profile-plan.json"),
        "native_plan_sha256": sha(HERE / "plan.json"),
        "build_sha256": build["build_sha256"],
        "owner": profile["owner"],
        "nested_topology_summary": {
            "owner": TOPOLOGY_OWNER,
            "aggregate_inclusive_ir": topology_total,
            "direct_child_partition": [
                {
                    "name": child_name,
                    "aggregate_inclusive_ir": child_totals[child_name],
                    "aggregate_calls": child_calls[child_name],
                    "share_of_topology_inclusive_ir": child_totals[child_name] / topology_total,
                }
                for child_name in sorted(child_totals,
                                          key=lambda item: (-child_totals[item], item))
            ],
            "dominant_next_owner": dominant_name,
            "dominant_next_owner_selection": (
                "Largest immediate direct child of write_topology_to_stream by "
                "aggregate inclusive Ir across the four selected publication captures."
            ),
            "partition_boundary": (
                "Only immediate direct children are summed. Nested owners overlap "
                "and are not added to the direct-child partition."
            ),
        },
        "rows": rows,
        "scope": profile["scope"],
        "limitations": [
            "Callgrind Ir is a guest-instruction attribution mechanism, not a hardware instruction, cycle, or latency measurement.",
            "Nested annotation owners overlap; direct children of the selected owner are the disjoint attribution boundary.",
            "The native publication timer includes destruction of the returned snapshot; this profile owner ends at the method return.",
        ],
    }
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return result


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, default=HERE / "profile-analysis.json")
    args = parser.parse_args()
    report = analyze(args.output)
    print(f"Publication profile edges verified: {args.output} ({len(report['rows'])} profiles)")
