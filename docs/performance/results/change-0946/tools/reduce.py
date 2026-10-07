#!/usr/bin/env python3
"""Bounded, offline child/self attribution for retained 0944 Callgrind dumps.

This reducer consumes only frozen packet inputs. It never launches a workload or
an external profiler. The retained reader is used for the Callgrind grammar;
all parsed direct and incoming edges, including zero-cost/zero-call records,
are retained instead of being filtered by the older summary adapter.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
import argparse
from pathlib import Path
import sys
from typing import Any, Mapping

HERE = Path(__file__).resolve().parent.parent
PLAN_PATH = HERE / "plan-v2.json"
READER_PATH = HERE / "inputs" / "reader_v2.py"
DEFAULT_OUTPUT_PATH = HERE / "attribution-summary-v2.json"


class ReductionError(ValueError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReductionError(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON input: {path}")
    try:
        def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
            result: dict[str, Any] = {}
            for key, value in pairs:
                require(key not in result, f"duplicate JSON key {key!r} in {path}")
                result[key] = value
            return result

        return json.loads(path.read_text(encoding="utf-8"),
                          object_pairs_hook=reject_duplicate_keys)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ReductionError(f"cannot read JSON input {path}: {error}") from error


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def load_reader() -> Any:
    require(READER_PATH.is_file(), f"missing retained reader: {READER_PATH}")
    spec = importlib.util.spec_from_file_location("retained_reader_0946", READER_PATH)
    require(spec is not None and spec.loader is not None, "cannot load retained reader")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


READER: Any | None = None
OWNER = "litchi_xlsx::cell_values::source::SourceBackedEditor::publish_multi_commit_to_stream"
MEASURED_PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
TARGETS = [
    OWNER,
    "litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream",
    "litchi_opc::source_backed::validate_source_part_replacement_xml",
    "xml_minifier::audit::verify_source_replacement",
    "xml_minifier::audit::verify_with_policy",
    "xml_minifier::audit::check_start",
    "xml_minifier::audit::inspect_attributes",
    "xml_minifier::audit::check_end",
    "xml_minifier::audit::wellformed::scan_qname",
    "quick_xml::reader::Reader<R>::read_event_impl",
    "quick_xml::reader::state::ReaderState::emit_start",
    "quick_xml::reader::state::ReaderState::emit_end",
]


def source_label(path: str) -> str:
    """Make source attribution portable while retaining crate-relative identity."""
    marker = "/source-build/"
    if marker in path:
        return "source-build/" + path.split(marker, 1)[1]
    registry = "/.cargo/registry/src/"
    if registry in path:
        tail = path.split(registry, 1)[1]
        parts = tail.split("/", 2)
        if len(parts) == 3:
            return "registry/" + parts[1] + "/" + parts[2]
    return path


def edge_view(edge: Mapping[str, Any], names: Mapping[int, str], *, incoming: bool) -> dict[str, Any]:
    caller_id = int(edge["caller_id"])
    callee_id = int(edge["callee_id"])
    return {
        "caller_id": caller_id,
        "caller": names.get(caller_id, "<unnamed>"),
        "callee_id": callee_id,
        "callee": names.get(callee_id, "<unnamed>"),
        "calls": int(edge["calls"]),
        "inclusive_ir": int(edge["inclusive_ir"]),
        "line": int(edge["line"]),
        "position": str(edge["position"]),
        "file": source_label(str(edge["file"])),
        "direction": "incoming" if incoming else "direct",
    }


def incoming_edges(parsed: Mapping[str, Any], target_id: int) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for function in parsed["functions"].values():
        for edge in function["edges"]:
            if int(edge["callee_id"]) == target_id:
                result.append(edge_view(edge, parsed["names"], incoming=True))
    return result


def direct_edges(parsed: Mapping[str, Any], function: Mapping[str, Any]) -> list[dict[str, Any]]:
    return [edge_view(edge, parsed["names"], incoming=False)
            for edge in function["edges"]]


def target_ids(parsed: Mapping[str, Any], name: str) -> list[int]:
    return sorted(function_id for function_id, function_name in parsed["names"].items()
                  if function_name == name and function_id in parsed["functions"])


def packet_file(relative: Any) -> Path:
    require(isinstance(relative, str) and relative, "packet path is not a non-empty string")
    path = Path(relative)
    require(not path.is_absolute() and ".." not in path.parts and "." not in path.parts,
            f"packet path is not a strict relative path: {relative!r}")
    current = HERE
    for component in path.parts:
        current /= component
        require(not current.is_symlink(), f"packet path contains symlink: {relative!r}")
    require(current.is_file(), f"packet input is missing or not regular: {relative!r}")
    return current


def node_record(parsed: Mapping[str, Any], function_id: int) -> dict[str, Any]:
    function = parsed["functions"][function_id]
    direct = direct_edges(parsed, function)
    incoming = incoming_edges(parsed, function_id)
    self_ir = int(function["self_ir"])
    direct_ir = sum(item["inclusive_ir"] for item in direct)
    incoming_ir = sum(item["inclusive_ir"] for item in incoming)
    require(self_ir + direct_ir == incoming_ir,
            f"{parsed['path']}: {function['name']}#{function_id} conservation failed "
            f"self={self_ir} direct={direct_ir} incoming={incoming_ir}")
    return {
        "id": int(function_id),
        "name": str(function["name"]),
        "source": source_label(str(function["file"])),
        "records": int(function["records"]),
        "exclusive_self_ir": self_ir,
        "inclusive_ir": incoming_ir,
        "direct_children_ir": direct_ir,
        "direct_edges": direct,
        "incoming_edges": incoming,
        "edge_census": {
            "direct_count": len(direct),
            "direct_zero_cost_count": sum(item["inclusive_ir"] == 0 for item in direct),
            "direct_zero_call_count": sum(item["calls"] == 0 for item in direct),
            "incoming_count": len(incoming),
            "incoming_zero_cost_count": sum(item["inclusive_ir"] == 0 for item in incoming),
            "incoming_zero_call_count": sum(item["calls"] == 0 for item in incoming),
        },
        "validation": {
            "exclusive_self_plus_all_direct_edges_equals_all_incoming_edges": True,
            "zero_cost_edges_retained": True,
        },
    }


def target_group(parsed: Mapping[str, Any], name: str, publication_ir: int) -> dict[str, Any]:
    ids = target_ids(parsed, name)
    require(ids, f"{parsed['path']}: target has no function record: {name}")
    records = [node_record(parsed, function_id) for function_id in ids]
    exclusive = sum(item["exclusive_self_ir"] for item in records)
    inclusive = sum(item["inclusive_ir"] for item in records)
    direct = sum(item["direct_children_ir"] for item in records)
    require(exclusive + direct == inclusive,
            f"{parsed['path']}: aggregate conservation failed for {name}")
    return {
        "name": name,
        "record_count": len(records),
        "records": records,
        "exclusive_self_ir": exclusive,
        "inclusive_ir": inclusive,
        "direct_children_ir": direct,
        "inclusive_share_of_publication": inclusive / publication_ir if publication_ir else None,
        "exclusive_share_of_publication": exclusive / publication_ir if publication_ir else None,
        "overlap_policy": "diagnostic only; do not add to another target or topology partition",
        "validation": {"all_function_records_conserve": True, "aggregate_conserves": True},
    }


def owner_record(group: Mapping[str, Any]) -> dict[str, Any]:
    matches = []
    for record in group["records"]:
        matches.extend(edge for edge in record["incoming_edges"]
                       if edge["caller"] == MEASURED_PARENT and edge["calls"] > 0)
    require(len(matches) == 1, f"measured owner edge is ambiguous: {matches}")
    require(matches[0]["calls"] == 1, "measured owner edge must be one call")
    return matches[0]


def capture(plan: Mapping[str, Any], prior: Mapping[str, Any], row: Mapping[str, Any]) -> dict[str, Any]:
    path = packet_file(row["input_path"])
    require(path.stat().st_size == row["bytes"], f"capture byte count differs: {path}")
    require(sha256(path) == row["sha256"], f"capture hash differs: {path}")
    require(READER is not None, "retained reader was not loaded after input validation")
    parsed = READER.parse_raw(path)
    header = parsed["header"]
    require(header["part"] == 6, f"{path}: measured part is not 6")
    require(header["trigger"] == "--dump-after=" + OWNER, f"{path}: trigger differs")
    publication_ir = int(header["summary_ir"])
    require(header["totals_ir"] == publication_ir, f"{path}: summary/totals differ")
    all_self_ir = sum(int(function["self_ir"])
                      for function in parsed["functions"].values())
    require(all_self_ir == publication_ir,
            f"{path}: all parsed exclusive self rows do not conserve summary")
    require(publication_ir == prior["publication_ir"],
            f"{path}: prior publication Ir differs: {publication_ir} vs {prior['publication_ir']}")
    groups = {name: target_group(parsed, name, publication_ir) for name in TARGETS}
    owner = groups[OWNER]
    measured_edge = owner_record(owner)
    require(measured_edge["inclusive_ir"] == publication_ir,
            f"{path}: owner incoming Ir differs from Callgrind summary")
    topology_name = "litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream"
    topology = groups[topology_name]
    wrapper_name = "litchi_opc::source_backed::validate_source_part_replacement_xml"
    wrapper = groups[wrapper_name]
    require(wrapper["inclusive_ir"] == prior["nested_xml_ir"],
            f"{path}: wrapper Ir differs from prior nested XML Ir")
    prior_exclusive = int(prior["exclusive_self_ir"])
    require(prior_exclusive == publication_ir,
            f"{path}: prior exclusive-self denominator differs from publication Ir")
    denominators = {
        "publication_owner_inclusive_ir": publication_ir,
        "publication_whole_dump_exclusive_self_ir": prior_exclusive,
        "topology_owner_inclusive_ir": topology["inclusive_ir"],
        "xml_wrapper_inclusive_ir": wrapper["inclusive_ir"],
    }
    return {
        "index": int(row["index"]),
        "shape": str(row["shape"]),
        "repeat": int(row["repeat"]),
        "input": {"path": str(path.relative_to(HERE)), "bytes": path.stat().st_size,
                  "sha256": row["sha256"]},
        "header": {"part": header["part"], "trigger": header["trigger"],
                   "summary_ir": publication_ir, "totals_ir": header["totals_ir"]},
        "denominators": denominators,
        "measured_owner_edge": measured_edge,
        "targets": groups,
        "validation": {
            "owner_incoming_equals_summary": True,
            "owner_summary_equals_prior_publication_ir": True,
            "prior_whole_dump_exclusive_equals_publication": True,
            "xml_wrapper_equals_prior_nested_xml_ir": True,
            "all_selected_nodes_conserve_self_plus_direct_vs_incoming": True,
            "all_direct_and_incoming_edges_including_zero_retained": True,
            "nested_rows_are_diagnostic_only": True,
        },
    }


def validate_input_descriptors(plan: Mapping[str, Any]) -> None:
    descriptors = plan.get("inputs", {}).get("descriptors")
    require(isinstance(descriptors, dict), "plan input descriptors are missing")
    expected = {
        "inputs/reader_v2.py",
        "inputs/plan.json",
        "inputs/root-reduction-v1.json",
        "inputs/source-build-manifest.sha256",
        "inputs/source-manifest.sha256",
        "inputs/focused-source-hashes.json",
    }
    require(set(descriptors) == expected, "plan input descriptor coverage differs")
    for key, item in descriptors.items():
        require(isinstance(item, dict), "input descriptor is malformed")
        require(item.get("path") == key, f"input descriptor key/path differs: {key}")
        path = packet_file(item["path"])
        require(path.stat().st_size == item["bytes"], f"input byte count differs: {path}")
        require(sha256(path) == item["sha256"], f"input hash differs: {path}")


def main(output_path: Path) -> int:
    try:
        plan = read_json(PLAN_PATH)
        require(plan.get("schema") == "litchi.xlsx-publication-attribution-0946.plan.v2",
                "plan schema differs")
        require(plan.get("status") == "frozen", "plan is not frozen")
        require(plan.get("targets") == TARGETS, "target list differs from frozen plan")
        validate_input_descriptors(plan)
        global READER
        READER = load_reader()
        prior_root = read_json(HERE / "inputs/root-reduction-v1.json")
        prior_by_index = {int(row["index"]): row for row in prior_root["captures"]}
        rows = plan["captures"]
        require(len(rows) == 4, "capture matrix must contain four rows")
        require([row.get("index") for row in rows] == [0, 1, 2, 3],
                "capture indices are not the exact frozen schedule")
        require([(row.get("shape"), row.get("repeat")) for row in rows]
                == [("medium", 0), ("dense-sparse", 0),
                    ("dense-sparse", 1), ("medium", 1)],
                "capture shape/repeat schedule differs")
        require(len({row.get("input_path") for row in rows}) == 4,
                "capture input paths are not unique")
        captures = []
        for row in rows:
            prior = prior_by_index.get(int(row["index"]))
            require(prior is not None and prior["shape"] == row["shape"]
                    and prior["repeat"] == row["repeat"],
                    f"prior capture binding differs at {row['index']}")
            prior = dict(prior)
            prior["exclusive_self_ir"] = int(prior_root["captures"][int(row["index"])]
                                             ["exclusive_self"]["all_self_ir"])
            captures.append(capture(plan, prior, row))
        output = {
            "schema": "litchi.xlsx-publication-attribution-0946.v1",
            "status": "pass",
            "tool": {"file": "tools/reduce.py", "sha256": sha256(Path(__file__))},
            "plan": {"file": "plan-v2.json", "sha256": sha256(PLAN_PATH)},
            "retained_reader": {"file": "inputs/reader_v2.py", "sha256": sha256(READER_PATH)},
            "source_and_binary": {
                "binary": plan["binary"],
                "source_commit": "933ee45cbb5d2ef9dc29583f7ec890ad234e04fb",
                "0944_xml_audit_sha256": plan["source_bindings"]["audit_rs"]["sha256"],
                "current_xml_audit_sha256": plan["source_bindings"]["current_checkout_audit_rs"]["sha256"],
                "xml_audit_source_unchanged": True,
                "current_native_or_0945_sharing_claim": False,
            },
            "scope": "Retained 0944 Callgrind Ir graph only; nested rows overlap and no native/adoption claim is made.",
            "captures": captures,
            "limitations": [
                "Inclusive edge costs are Callgrind Ir and can overlap across nested targets.",
                "Exclusive self rows are exact only for the selected Callgrind dump and are not latency, allocation, or RSS measures.",
                "Call counts are provenance and are not used to normalize costs or infer per-call equality.",
                "The profile predates 0945 SourcePayload sharing; this packet does not qualify the 0945 candidate.",
            ],
        }
        output_path = output_path.resolve()
        require(not output_path.exists(), f"refusing to overwrite output: {output_path}")
        with output_path.open("x", encoding="utf-8") as stream:
            stream.write(json.dumps(output, indent=2, sort_keys=True) + "\n")
        print(f"0946 bounded attribution passed: {output_path} ({len(captures)} captures)")
        return 0
    except (ReductionError, OSError, ImportError) as error:
        print(f"reduce.py: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_PATH)
    args = parser.parse_args()
    raise SystemExit(main(args.output))
