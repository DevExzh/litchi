#!/usr/bin/env python3
"""Validate and attribute the retained 0709 DOCX edit Callgrind parts.

This is a read-only analysis of the four measured ``.callgrind.5`` parts from
the 0709 ordinary-save profile packet.  The measured owner is
``ordinary_save::Owner::edit`` and the requested nested boundary is
``Package::document_mut``.  Immediate children are the only disjoint
partition at each boundary.  Deeper rows are diagnostics inside one selected
child and are never added to an ancestor partition.

The output deliberately binds the result to the historical 0709 binary and
source manifest.  The 0710 preservation change guarded custom-property
publication in ``write_plain``; it did not change this edit capture.  Nothing
in this file infers current-source latency or a causal production speedup.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any, NoReturn


sys.dont_write_bytecode = True

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PACKET_0709 = REPO / "docs/performance/results/change-0709"
PACKET_0710 = REPO / "docs/performance/results/change-0710"
OUTPUT = HERE / "edit-attribution.json"

EDIT_OWNER = "litchi_perf_baseline::ordinary_save::Owner::edit"
MEASURED_PARENT = "litchi_perf_baseline::ordinary_save::run_case"
DOCUMENT_MUT = (
    "litchi_docx::package::package::document::<impl "
    "litchi_docx::package::model::Package>::document_mut"
)
MUTABLE_FROM_XML = "litchi_docx::writer::doc::model::MutableDocument::from_xml"
BODY_FROM_XML = "litchi_docx::writer::doc::package::DocumentBody::from_xml"
VALIDATE_SECTION = (
    "litchi_docx::writer::doc::package::DocumentBody::validate_section_placement"
)
ACTIVE_BLOCK_RANGES = "litchi_docx::parts::document_part::active_block_ranges"
ALT_SCAN = "litchi_docx::alt::codec::scan"

HISTORICAL_REVISION = "9b6a9bbfc7bcc2fbf103419eb6968e1025c529c7"
HISTORICAL_BINARY_SHA256 = (
    "042587f297602881a313881c77017bb53af286e7ee24d42906a24ed3b08455c0"
)
HISTORICAL_SOURCE_MANIFEST_SHA256 = (
    "7b322d8d66abde992b0e05fae73c7388c3c09fcd233257de0f8291e6c8eb87ea"
)
HISTORICAL_SOURCE_CENSUS_SHA256 = (
    "db839c00b1048798c2141047394e6286c0d558f56cc814a32ada1c6410fa8ae8"
)
HISTORICAL_PROFILE_PLAN_SHA256 = (
    "62bb99d3b05d6da34a5adb51e5e85cc327bba8f6f0271d52c6de21158db9cd77"
)
HISTORICAL_NATIVE_PLAN_SHA256 = (
    "d69de12689b7537fd00165b48c8f1ba5ca081975f4b68caea0c27879adb2fd17"
)

SOURCE_PATHS = (
    "crates/litchi-docx/src/alt/codec.rs",
    "crates/litchi-docx/src/namespace.rs",
    "crates/litchi-docx/src/package/codec.rs",
    "crates/litchi-docx/src/parts/document_part.rs",
    "crates/litchi-docx/src/writer/doc/model.rs",
    "crates/litchi-docx/src/writer/doc/package.rs",
    "crates/litchi-ooxml-common/src/mce/codec.rs",
)

PROFILE_NAMES = (
    "profile-r1-generated-edit",
    "profile-r2-generated-edit",
    "profile-r1-numbered-list-edit",
    "profile-r2-numbered-list-edit",
)

class EvidenceError(ValueError):
    """A missing, malformed, or contradictory retained artifact."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def read_text(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing text artifact: {path}")
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        fail(f"cannot read {path}: {error}")


def read_json(path: Path) -> Any:
    try:
        return json.loads(read_text(path))
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def relative_0709(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET_0709))
    except ValueError as error:
        fail(f"path is outside change-0709 packet: {path}")


def relative_repo(path: Path) -> str:
    try:
        return str(path.relative_to(REPO))
    except ValueError as error:
        fail(f"path is outside repository: {path}")


def load_retained_parser() -> Any:
    path = PACKET_0709 / "analyze_profiles.py"
    spec = importlib.util.spec_from_file_location("retained_profile_0709", path)
    require(spec is not None and spec.loader is not None,
            f"cannot load retained parser: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


H = load_retained_parser()


def raw_edges(path: Path) -> list[dict[str, Any]]:
    """Use the retained parser's Callgrind edge decoder, without annotation writes."""

    return H.raw_edges(path)


def edges_between(raw: list[dict[str, Any]], caller: str, callee: str) -> list[dict[str, Any]]:
    return [edge for edge in raw if edge["caller"] == caller and edge["callee"] == callee]


def aggregate_children(raw: list[dict[str, Any]], caller: str) -> list[dict[str, Any]]:
    grouped: dict[str, dict[str, int]] = {}
    for edge in raw:
        if edge["caller"] != caller:
            continue
        value = grouped.setdefault(edge["callee"], {
            "name": edge["callee"],
            "inclusive_ir": 0,
            "calls": 0,
            "raw_edge_count": 0,
        })
        value["inclusive_ir"] += edge["inclusive_ir"]
        value["calls"] += edge["calls"]
        value["raw_edge_count"] += 1
    return sorted(grouped.values(), key=lambda item: (-item["inclusive_ir"], item["name"]))


def partition(
    raw: list[dict[str, Any]],
    owner: str,
    incoming_caller: str,
    incoming_callee: str,
    label: str,
) -> dict[str, Any]:
    """Return one disjoint immediate-child partition at an explicit edge."""

    incoming = edges_between(raw, incoming_caller, incoming_callee)
    require(incoming, f"{label}: missing incoming edge for {owner}")
    require(incoming_callee == owner,
            f"{label}: partition owner differs from incoming callee")
    owner_ir = sum(edge["inclusive_ir"] for edge in incoming)
    children = aggregate_children(raw, owner)
    direct_ir = sum(item["inclusive_ir"] for item in children)
    self_ir = owner_ir - direct_ir
    require(self_ir >= 0, f"{label}: negative owner self Ir {self_ir}")
    require(self_ir + direct_ir == owner_ir,
            f"{label}: immediate-child partition does not reconstruct owner Ir")
    return {
        "owner": owner,
        "incoming": {
            "caller": incoming_caller,
            "callee": incoming_callee,
            "edge_count": len(incoming),
            "calls": sum(edge["calls"] for edge in incoming),
            "inclusive_ir": owner_ir,
            "raw_edges": incoming,
        },
        "self_ir": self_ir,
        "direct_children_ir": direct_ir,
        "children": children,
        "partition_equation": "self_ir + sum(immediate direct-child inclusive Ir) = owner inclusive Ir",
        "disjoint": True,
        "nested_inclusive_costs_excluded": True,
    }


def child(part: dict[str, Any], name: str, label: str) -> dict[str, Any]:
    matches = [item for item in part["children"] if item["name"] == name]
    require(len(matches) == 1, f"{label}: expected one direct child named {name!r}")
    return matches[0]


def share(value: int, denominator: int) -> float:
    require(denominator > 0, "share denominator must be positive")
    return value / denominator


def file_binding(path: Path) -> dict[str, Any]:
    return {"file": relative_repo(path), "sha256": sha256(path), "bytes": path.stat().st_size}


def validate_historical_packet() -> dict[str, Any]:
    revision = read_json(PACKET_0709 / "revision.json")
    require(revision.get("revision") == HISTORICAL_REVISION,
            "change-0709 revision does not match historical profile source")

    source_manifest_path = PACKET_0709 / "source-baseline.json"
    source_manifest = read_json(source_manifest_path)
    require(sha256(source_manifest_path) == HISTORICAL_SOURCE_MANIFEST_SHA256,
            "historical source manifest hash changed")
    source_files: dict[str, str] = {}
    for path in SOURCE_PATHS:
        digest = source_manifest.get(path)
        require(isinstance(digest, str) and len(digest) == 64,
                f"historical source manifest omits {path}")
        source_files[path] = digest

    transition_path = PACKET_0709 / "profile-revision-transition.json"
    transition = read_json(transition_path)
    require(transition.get("build_revision") == HISTORICAL_REVISION,
            "profile transition build revision differs")
    require(transition.get("source_manifest_sha256") == HISTORICAL_SOURCE_MANIFEST_SHA256,
            "profile transition source manifest differs")

    build_path = PACKET_0709 / "build-baseline.json"
    builds = read_json(build_path)
    require(isinstance(builds, list), "build-baseline.json is not a list")
    native_builds = [item for item in builds
                     if isinstance(item, dict)
                     and item.get("binary_sha256") == HISTORICAL_BINARY_SHA256]
    require(len(native_builds) == 1, "historical native build binding is not unique")
    native_build = native_builds[0]
    require(native_build.get("source_manifest_sha256") == HISTORICAL_SOURCE_MANIFEST_SHA256,
            "historical native build source binding differs")

    change_patch = PACKET_0710 / "change.patch"
    change_text = read_text(change_patch)
    require("if self.custom_props_dirty" in change_text,
            "0710 custom-properties guard is absent from retained patch")
    require(".write_for(&mut self.opc, CustomPropsHost::Word)" in change_text,
            "0710 custom-properties publication call is absent from retained patch")
    production_diff = "diff --git a/crates/litchi-docx/src/package/codec.rs b/crates/litchi-docx/src/package/codec.rs"
    test_diff = "diff --git a/crates/litchi-docx/tests/custom_properties.rs b/crates/litchi-docx/tests/custom_properties.rs"
    require(production_diff in change_text and test_diff in change_text,
            "0710 retained patch paths differ")

    return {
        "profile_packet": "change-0709-docx-ordinary-save",
        "profile_source_revision": HISTORICAL_REVISION,
        "profile_checkout_revision": transition["profile_checkout_revision"],
        "binary": {
            "path": native_build["binary"],
            "sha256": HISTORICAL_BINARY_SHA256,
            "bytes": native_build["binary_bytes"],
            "custody": "historical receipt binding; scratch binary is no longer required to replay raw attribution",
        },
        "source_manifest": file_binding(source_manifest_path)
        | {"sha256": HISTORICAL_SOURCE_MANIFEST_SHA256},
        "source_files": source_files,
        "profile_transition": file_binding(transition_path)
        | {"sha256": sha256(transition_path)},
        "0710_publication_change": {
            "patch": file_binding(change_patch),
            "production_path": "crates/litchi-docx/src/package/codec.rs",
            "test_path": "crates/litchi-docx/tests/custom_properties.rs",
            "scope": "guard untouched custom-property publication in write_plain; no edit-owner capture was rebuilt",
        },
    }


def validate_profile(name: str, analysis: dict[str, Any]) -> dict[str, Any]:
    raw_path = PACKET_0709 / f"{name}.callgrind.5"
    receipt_path = PACKET_0709 / f"{name}.receipt.json"
    profile_path = PACKET_0709 / f"{name}.json"
    receipt = read_json(receipt_path)
    profile_result = read_json(profile_path)
    binary_identity = profile_result.get("binary_identity") if isinstance(profile_result, dict) else None
    require(isinstance(binary_identity, dict), f"{name}: profile binary identity is malformed")
    require(binary_identity.get("binary_sha256") == HISTORICAL_BINARY_SHA256,
            f"{name}: profile result binary binding differs")
    require(binary_identity.get("binary_bytes") == 62889832,
            f"{name}: profile result binary size differs")
    raw_text = read_text(raw_path)
    label = relative_0709(raw_path)

    require(receipt.get("name") == name, f"{name}: receipt name differs")
    require(receipt.get("phase") == "edit", f"{name}: receipt phase is not edit")
    require(receipt.get("owner") == EDIT_OWNER, f"{name}: receipt owner differs")
    require(receipt.get("measured_parent") == MEASURED_PARENT,
            f"{name}: receipt measured parent differs")
    require(receipt.get("binary_sha256") == HISTORICAL_BINARY_SHA256,
            f"{name}: receipt binary binding differs")
    require(receipt.get("build_source_manifest_sha256") == HISTORICAL_SOURCE_MANIFEST_SHA256,
            f"{name}: receipt source manifest binding differs")
    require(receipt.get("retained_source_census_sha256") == HISTORICAL_SOURCE_CENSUS_SHA256,
            f"{name}: receipt source census binding differs")
    require(receipt.get("profile_plan_sha256") == HISTORICAL_PROFILE_PLAN_SHA256,
            f"{name}: receipt profile plan binding differs")
    require(receipt.get("plan_sha256") == HISTORICAL_NATIVE_PLAN_SHA256,
            f"{name}: receipt native plan binding differs")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{name}: receipt artifact map is malformed")
    require(artifacts.get(raw_path.name) == sha256(raw_path),
            f"{name}: raw .5 artifact hash differs from receipt")

    require(H.H.part_number(raw_text, label) == 5, f"{name}: retained part is not part 5")
    require(H.H.trigger(raw_text, label) == f"--dump-after={EDIT_OWNER}",
            f"{name}: retained trigger does not select exact edit owner")
    summary_ir = H.callgrind_summary(raw_text, label)
    raw = raw_edges(raw_path)
    owner_in = edges_between(raw, MEASURED_PARENT, EDIT_OWNER)
    require(len(owner_in) == 1 and owner_in[0]["calls"] == 1,
            f"{name}: run_case -> Owner::edit edge is not exactly one call")
    require(owner_in[0]["inclusive_ir"] == summary_ir,
            f"{name}: Owner::edit incoming Ir differs from raw summary")

    rows = analysis.get("profiles")
    require(isinstance(rows, list), "profile-analysis profiles are malformed")
    row_matches = [row for row in rows if isinstance(row, dict) and row.get("name") == name]
    require(len(row_matches) == 1, f"{name}: profile-analysis row is not unique")
    analysis_row = row_matches[0]
    owner_annotation = analysis_row["annotations"]["owner"]
    require(owner_annotation["inclusive_ir"] == summary_ir,
            f"{name}: retained annotation owner Ir differs from raw summary")
    require(analysis_row["parts"][4]["part"] == 5
            and analysis_row["parts"][4]["role"] == "measured",
            f"{name}: analysis part 5 is not the measured part")

    top_partition = partition(raw, EDIT_OWNER, MEASURED_PARENT, EDIT_OWNER, name)
    require(top_partition["self_ir"] == owner_annotation["self_ir"],
            f"{name}: raw edit self Ir differs from retained annotation")
    require({item["name"]: item["inclusive_ir"] for item in top_partition["children"]}
            == owner_annotation["direct_callee_ir"],
            f"{name}: raw edit direct children differ from retained annotation")

    document_partition = partition(raw, DOCUMENT_MUT, EDIT_OWNER, DOCUMENT_MUT, name)
    require(document_partition["incoming"]["inclusive_ir"]
            == owner_annotation["direct_callee_ir"][DOCUMENT_MUT],
            f"{name}: document_mut Ir differs from edit annotation")

    mutable_partition = partition(
        raw, MUTABLE_FROM_XML, DOCUMENT_MUT, MUTABLE_FROM_XML, name
    )
    body_partition = partition(
        raw, BODY_FROM_XML, MUTABLE_FROM_XML, BODY_FROM_XML, name
    )

    # This is a second call site in MutableDocument::from_xml.  Keep its
    # incoming edge separate; aggregating it with validate_section_placement
    # would make a misleading partition for SectionProperties::from_xml.
    validate_partition = partition(
        raw, VALIDATE_SECTION, MUTABLE_FROM_XML, VALIDATE_SECTION, name
    )
    alt = child(body_partition, ALT_SCAN, name)
    active = child(body_partition, ACTIVE_BLOCK_RANGES, name)

    return {
        "name": name,
        "corpus": receipt["corpus_id"],
        "repeat": receipt["repeat"],
        "historical_binary_sha256": HISTORICAL_BINARY_SHA256,
        "raw_artifact": file_binding(raw_path)
        | {"receipt": relative_0709(receipt_path), "receipt_sha256": sha256(receipt_path)},
        "profile_result": {
            "file": relative_0709(profile_path),
            "sha256": sha256(profile_path),
            "binary_identity": binary_identity,
        },
        "owner_summary_ir": summary_ir,
        "part": 5,
        "part_trigger": f"--dump-after={EDIT_OWNER}",
        "partitions": {
            "owner_edit": top_partition,
            "document_mut": document_partition,
            "mutable_document_from_xml": mutable_partition,
            "document_body_from_xml": body_partition,
            "validate_section_placement": validate_partition,
        },
        "deeper_diagnostics": {
            "alt_scan": {
                "inclusive_ir": alt["inclusive_ir"],
                "share_of_document_mut": share(
                    alt["inclusive_ir"], document_partition["incoming"]["inclusive_ir"]
                ),
                "share_of_document_body_from_xml": share(
                    alt["inclusive_ir"], body_partition["incoming"]["inclusive_ir"]
                ),
                "calls": alt["calls"],
            },
            "active_block_ranges": {
                "inclusive_ir": active["inclusive_ir"],
                "share_of_document_mut": share(
                    active["inclusive_ir"], document_partition["incoming"]["inclusive_ir"]
                ),
                "share_of_document_body_from_xml": share(
                    active["inclusive_ir"], body_partition["incoming"]["inclusive_ir"]
                ),
                "calls": active["calls"],
            },
            "dominant_body_child": body_partition["children"][0]["name"],
            "dominant_mutable_document_child": mutable_partition["children"][0]["name"],
        },
        "validation": {
            "raw_owner_edge_matches_summary": True,
            "raw_edit_partition_reconciles": True,
            "raw_document_mut_partition_reconciles": True,
            "raw_mutable_document_partition_reconciles": True,
            "raw_document_body_partition_reconciles": True,
            "nested_partitions_excluded_from_ancestor_sums": True,
        },
    }


def aggregate_partitions(rows: list[dict[str, Any]], partition_key: str) -> dict[str, Any]:
    by_corpus: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        by_corpus.setdefault(row["corpus"], []).append(row["partitions"][partition_key])
    result: dict[str, Any] = {}
    for corpus, parts in sorted(by_corpus.items()):
        owner_ir = sum(part["incoming"]["inclusive_ir"] for part in parts)
        self_ir = sum(part["self_ir"] for part in parts)
        children: dict[str, dict[str, int]] = {}
        for part in parts:
            for item in part["children"]:
                value = children.setdefault(item["name"], {
                    "name": item["name"],
                    "inclusive_ir": 0,
                    "calls": 0,
                    "raw_edge_count": 0,
                })
                value["inclusive_ir"] += item["inclusive_ir"]
                value["calls"] += item["calls"]
                value["raw_edge_count"] += item["raw_edge_count"]
        direct_ir = sum(item["inclusive_ir"] for item in children.values())
        require(self_ir + direct_ir == owner_ir,
                f"{corpus}/{partition_key}: aggregate partition does not reconcile")
        for item in children.values():
            item["share_of_owner"] = share(item["inclusive_ir"], owner_ir)
        result[corpus] = {
            "profile_count": len(parts),
            "owner_ir": owner_ir,
            "self_ir": self_ir,
            "direct_children_ir": direct_ir,
            "children": sorted(children.values(),
                                key=lambda item: (-item["inclusive_ir"], item["name"])),
            "partition_equation": "self_ir + sum(immediate direct-child inclusive Ir) = owner inclusive Ir",
            "disjoint": True,
            "nested_inclusive_costs_excluded": True,
        }
    return result


def aggregate_diagnostics(rows: list[dict[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for corpus in sorted({row["corpus"] for row in rows}):
        selected = [row for row in rows if row["corpus"] == corpus]
        dm_ir = sum(row["partitions"]["document_mut"]["incoming"]["inclusive_ir"]
                    for row in selected)
        body_ir = sum(row["partitions"]["document_body_from_xml"]["incoming"]["inclusive_ir"]
                      for row in selected)
        diagnostics = {}
        for key in ("alt_scan", "active_block_ranges"):
            ir = sum(row["deeper_diagnostics"][key]["inclusive_ir"] for row in selected)
            diagnostics[key] = {
                "inclusive_ir": ir,
                "share_of_document_mut": share(ir, dm_ir),
                "share_of_document_body_from_xml": share(ir, body_ir),
            }
        result[corpus] = {
            "profile_count": len(selected),
            "document_mut_ir": dm_ir,
            "document_body_from_xml_ir": body_ir,
            "children": diagnostics,
            "dominant_body_child": max(
                diagnostics, key=lambda key: diagnostics[key]["inclusive_ir"]
            ),
        }
    return result


def recommendation(aggregates: dict[str, Any]) -> dict[str, Any]:
    generated = aggregates["generated"]
    numbered = aggregates["numbered-list"]
    generated_alt = generated["children"]["alt_scan"]
    numbered_alt = numbered["children"]["alt_scan"]
    numbered_active = numbered["children"]["active_block_ranges"]
    return {
        "status": "diagnostic_only_no_production_change",
        "candidate_owner": ALT_SCAN,
        "boundary": f"{BODY_FROM_XML} -> {ALT_SCAN}",
        "selection": (
            "The generated corpus selects alt::codec::scan because it is the largest "
            "immediate child of DocumentBody::from_xml in both retained repeats. "
            "The admitted NumberedList corpus is a required contrast: its largest child "
            "is active_block_ranges, so this is not a universal hotspot claim."
        ),
        "instruction_share": {
            "generated": generated_alt,
            "numbered_list_contrast": numbered_alt,
            "numbered_list_dominant_alternative": numbered_active,
        },
        "bounded_next_experiment": (
            "Prototype one source-preserving seam at the DocumentBody::from_xml -> "
            "alt::codec::scan boundary: reuse a validated event/namespace context or "
            "prove a bounded no-altChunk fast path, then recapture both corpora. Keep "
            "the seam local to DOCX parsing and retain the existing ownership boundary."
        ),
        "constraints": [
            "Preserve exact source ranges, opaque/unknown markup, namespace choices, and lexical bytes.",
            "Preserve active mc:Choice and mc:Fallback selection and strict/transitional namespace resolution.",
            "Preserve altChunk relationship validation, source-order offsets, duplicate detection, and typed refusals.",
            "Retain XML byte/depth/node limits, MAX_CHUNKS, offset conversions, checked arithmetic, and allocation failures.",
            "Do not classify comments, attributes, or arbitrary prefixes as altChunk anchors through a byte-only shortcut.",
            "Run preservation, refusal, negative, native latency, allocator, and instruction gates before promotion; this packet has no causal speedup claim.",
        ],
        "deferred_contrast": (
            "active_block_ranges is the larger retained real-fixture child and should be "
            "measured separately before a shared scanner abstraction is proposed."
        ),
    }


def analyze() -> dict[str, Any]:
    historical = validate_historical_packet()
    analysis_path = PACKET_0709 / "profile-analysis.json"
    analysis = read_json(analysis_path)
    require(analysis.get("build", {}).get("binary_sha256") == HISTORICAL_BINARY_SHA256,
            "profile-analysis binary binding differs")
    require(analysis.get("build", {}).get("source_manifest_sha256")
            == HISTORICAL_SOURCE_MANIFEST_SHA256,
            "profile-analysis source manifest binding differs")
    require(analysis.get("analyzer_script_sha256") == sha256(PACKET_0709 / "analyze_profiles.py"),
            "profile-analysis parser hash differs")

    rows = [validate_profile(name, analysis) for name in PROFILE_NAMES]
    require({row["corpus"] for row in rows} == {"generated", "numbered-list"},
            "edit attribution corpus set differs")
    require({row["repeat"] for row in rows} == {1, 2},
            "edit attribution repeat set differs")

    aggregates = {
        "document_mut": aggregate_partitions(rows, "document_mut"),
        "mutable_document_from_xml": aggregate_partitions(rows, "mutable_document_from_xml"),
        "document_body_from_xml": aggregate_partitions(rows, "document_body_from_xml"),
        "validate_section_placement": aggregate_partitions(rows, "validate_section_placement"),
    }
    diagnostics = aggregate_diagnostics(rows)
    return {
        "schema": "docx_edit_nested_attribution_v1",
        "status": "pass",
        "scope": {
            "phase": "edit",
            "owner": EDIT_OWNER,
            "measured_parent": MEASURED_PARENT,
            "selected_part": 5,
            "profile_count": len(rows),
            "corpora": ["generated", "numbered-list"],
            "metric": "Callgrind Ir guest instructions",
            "selection_method": "retained raw positive incoming edge from run_case and exact caller/callee names",
        },
        "historical_binding": historical
        | {
            "analysis": file_binding(analysis_path),
            "analysis_sha256": sha256(analysis_path),
            "native_plan": file_binding(PACKET_0709 / "plan.json")
            | {"sha256": HISTORICAL_NATIVE_PLAN_SHA256},
            "profile_plan": file_binding(PACKET_0709 / "profile-plan.json")
            | {"sha256": HISTORICAL_PROFILE_PLAN_SHA256},
        },
        "rows": rows,
        "aggregate_partitions": aggregates,
        "aggregate_diagnostics": diagnostics,
        "recommendation": recommendation(diagnostics),
        "limitations": [
            "The four parts are historical 0709 profiles, not a current-source or 0710 rebuild.",
            "Callgrind Ir is guest-instruction attribution, not native latency, hardware cycles, allocation counts, RSS, or cache counters.",
            "Immediate children at one owner boundary are disjoint; every deeper partition is nested inside one child and is excluded from ancestor totals.",
            "Instruction shares describe this capture and do not establish causality or predict a speedup without a fresh candidate capture.",
            "The generated and NumberedList corpora are contrasting fixtures; neither represents all DOCX producer or document shapes.",
        ],
    }


def main() -> int:
    try:
        result = analyze()
        OUTPUT.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    except (EvidenceError, OSError, KeyError, TypeError) as error:
        print(f"edit-attribution.py: error: {error}", file=sys.stderr)
        return 1
    print(f"verified {len(result['rows'])} retained DOCX edit profiles; wrote {OUTPUT}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
