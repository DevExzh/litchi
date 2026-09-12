#!/usr/bin/env python3
"""Recompute the sealed 0525 scanner ceiling and 0527 phase context.

This read-only audit reuses the 0521 adapter and 0519 raw-edge parser.  The
reported resolver value is the complete scanner -> ``resolve_event`` edge,
so it is an upper bound for End-only resolver removal.  Start and Empty
lookups remain required.  Call metadata is deliberately not used for a
per-call estimate.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
from pathlib import Path
from typing import Any

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
P0525 = HERE.parent / "change-0525"
P0526 = HERE.parent / "change-0526"
P0527 = HERE.parent / "change-0527"
PLAN = HERE / "plan.json"
RUN = HERE / "run.py"
FROZEN = HERE / "frozen-inputs.json"
BINDING = HERE / "source-binding.json"
CANDIDATE_MANIFEST = P0525 / "candidate" / "source-manifest.json"
CANDIDATE_DIFF = P0525 / "candidate" / "source-diff.json"
HELPER_PATH = HERE.parent / "change-0521" / "analyze_profiles.py"
RAW_HELPER_PATH = HERE.parent / "change-0519" / "analyze_profiles.py"
EDGE_HELPER_PATH = HERE.parent / "change-0519" / "compare_profile_lanes.py"

OWNER = "litchi_xlsx::cell_values::source::MultiSourceEdit::commit"
SCAN = "litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::scan_with_limit"
START = "litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::Scanner::start_cell"
ADDRESS = "litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::Scanner::cell_address"
RESOLVE = "quick_xml::name::NamespaceResolver::resolve_event"
PARENT = "litchi_perf_baseline::run_xlsx_cell_values_edit_save"
ROWS = tuple((repeat, shape) for repeat in (1, 2) for shape in ("medium", "dense-sparse"))
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")


class EvidenceError(ValueError):
    pass


def require(ok: bool, message: str) -> None:
    if not ok:
        raise EvidenceError(message)


def digest(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def read(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def text(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def rel(path: Path) -> str:
    return path.relative_to(REPO).as_posix()


spec = importlib.util.spec_from_file_location("litchi_0521_profile_helpers", HELPER_PATH)
require(spec is not None and spec.loader is not None, "cannot load 0521 profile helper")
HELPER = importlib.util.module_from_spec(spec)
spec.loader.exec_module(HELPER)


def check_seal(root: Path) -> int:
    expected: dict[str, str] = {}
    for line in text(root / "SHA256SUMS").splitlines():
        fields = line.split("  ", 1)
        require(len(fields) == 2, f"malformed seal line in {root}")
        value, name = fields
        p = Path(name)
        require(not p.is_absolute() and ".." not in p.parts and name != "SHA256SUMS"
                and name not in expected, f"unsafe/duplicate sealed path: {name}")
        expected[name] = value
    actual = {
        path.relative_to(root).as_posix(): digest(path)
        for path in root.rglob("*")
        if path.is_file() and not path.is_symlink() and path.name != "SHA256SUMS"
    }
    require(expected == actual, f"sealed inventory mismatch: {root}")
    return len(expected)


def source_binding() -> dict[str, Any]:
    binding = read(BINDING)
    current: dict[str, str] = {}
    for name, expected in binding.get("source_files", {}).items():
        path = REPO / name
        require(path.is_file() and digest(path) == expected,
                f"bound source differs: {name}")
        current[name] = expected
    for name, item in binding.get("dependency_inputs", {}).items():
        source, retained = Path(name), HERE / item["retained"]
        require(source.is_file() and retained.is_file(), f"missing dependency: {name}")
        require(digest(source) == item["sha256"] == digest(retained),
                f"dependency binding differs: {name}")
    for directory, expected in binding.get("prior_seals", {}).items():
        seal = REPO / directory / "SHA256SUMS"
        require(seal.is_file() and digest(seal) == expected,
                f"prior seal binding differs: {directory}")

    candidate, diff = read(CANDIDATE_MANIFEST), read(CANDIDATE_DIFF)
    production: dict[str, dict[str, str]] = {}
    test_only: dict[str, dict[str, str | None]] = {}
    for name, item in diff.get("changed_files", {}).items():
        if not name.startswith("crates/litchi-xlsx/src/"):
            continue
        path = REPO / name
        current_hash = digest(path) if path.is_file() else None
        record = {"candidate_sha256": item["candidate_sha256"],
                  "current_sha256": current_hash}
        if name.endswith("/cell_values/snapshot.rs"):
            # 0527 retained additions in this file are cfg(test) only.
            test_only[name] = record
        else:
            require(current_hash == item["candidate_sha256"],
                    f"0525 production source differs: {name}")
            production[name] = record
    require(production, "no 0525 production source was bound")
    return {
        "binding_sha256": digest(BINDING), "revision": binding.get("revision"),
        "bound_source_files": current,
        "dependency_inputs": binding.get("dependency_inputs", {}),
        "prior_seals": binding.get("prior_seals", {}),
        "candidate_manifest_sha256": digest(CANDIDATE_MANIFEST),
        "candidate_diff_sha256": digest(CANDIDATE_DIFF),
        "0525_production_changed_files": production,
        "0527_test_only_source_observation": test_only,
        "production_unchanged_from_0525_candidate": True,
        "candidate_manifest_contains_bound_sources": all(
            name in candidate and candidate[name] == value["candidate_sha256"]
            for name, value in production.items()
        ),
    }


def input_hash(path: Path) -> dict[str, str]:
    require(path.is_file(), f"missing input: {path}")
    return {"path": rel(path), "sha256": digest(path)}


def inputs() -> dict[str, Any]:
    frozen = read(FROZEN)
    require(frozen.get("plan.json") == digest(PLAN), "frozen plan hash differs")
    require(frozen.get("run.py") == digest(RUN), "frozen driver hash differs")
    return {
        "plan": input_hash(PLAN), "run": input_hash(RUN),
        "frozen_inputs": input_hash(FROZEN),
        "prior_0525_seal": input_hash(P0525 / "SHA256SUMS"),
        "prior_0526_profile_analysis": input_hash(P0526 / "profile-analysis.json"),
        "prior_0526_source_binding": input_hash(P0526 / "source-binding.json"),
        "prior_0525_candidate_manifest": input_hash(CANDIDATE_MANIFEST),
        "prior_0525_candidate_diff": input_hash(CANDIDATE_DIFF),
        "helpers": {rel(path): digest(path) for path in
                     (HELPER_PATH, RAW_HELPER_PATH, EDGE_HELPER_PATH)},
    }


def raw_view(edge_summary: dict[str, Any]) -> dict[str, Any]:
    # The helper uses calls to identify positive edges, but call counts are
    # intentionally absent here because collection-off metadata is polluted.
    return {
        "target": edge_summary["target"], "caller": edge_summary.get("caller"),
        "positive_edge_count": edge_summary["positive_edge_count"],
        "inclusive_ir": edge_summary["inclusive_ir"],
        "edges": [{key: edge[key] for key in ("caller", "callee", "inclusive_ir")}
                  for edge in edge_summary["edges"]],
    }


def annotation(inc: str, self_text: str, name: str, label: str) -> dict[str, Any]:
    inclusive = HELPER.parse_annotation(inc, name, label + " inclusive")
    exclusive = HELPER.parse_annotation(self_text, name, label + " self")
    require(inclusive["direct"] == exclusive["direct"], f"annotation children differ: {name}")
    direct = HELPER.direct_map(inclusive["direct"])
    direct_sum = sum(direct.values())
    require(inclusive["selected_ir"] == exclusive["selected_ir"] + direct_sum,
            f"self plus direct equation differs: {name}")
    return {
        "inclusive_ir": inclusive["selected_ir"], "self_ir": exclusive["selected_ir"],
        "direct": dict(sorted(direct.items(), key=lambda item: (-item[1], item[0]))),
        "equation": {
            "inclusive_ir": inclusive["selected_ir"], "self_ir": exclusive["selected_ir"],
            "direct_sum_ir": direct_sum,
            "self_plus_direct_ir": exclusive["selected_ir"] + direct_sum,
            "verified": True,
        },
    }


def validate_children(raw: Path, owner: str, data: dict[str, Any]) -> None:
    for child, cost in data["direct"].items():
        edge = HELPER.target_edge_summary(raw, child, owner)
        require(edge["positive_edge_count"] > 0 and edge["inclusive_ir"] == cost,
                f"raw edge differs: {owner} -> {child}")


def prior_row(prior: dict[str, Any], repeat: int, shape: str) -> dict[str, Any]:
    found = [row for row in prior["rows"]
             if row["repeat"] == repeat and row["shape"] == shape]
    require(len(found) == 1, f"prior row missing/duplicated: {repeat}/{shape}")
    return found[0]


def ranked(data: dict[str, Any], denominator: int, share_name: str) -> list[dict[str, Any]]:
    return [
        {"rank": rank, "name": name, "inclusive_ir": cost,
         share_name: cost / denominator * 100, "immediate_disjoint_child": True}
        for rank, (name, cost) in enumerate(
            sorted(data["direct"].items(), key=lambda item: (-item[1], item[0])), 1
        )
    ]


def profile_row(repeat: int, shape: str, prior: dict[str, Any]) -> dict[str, Any]:
    stem = P0525 / "candidate" / f"profile-r{repeat}-{shape}"
    raw, inc_path, self_path = (Path(str(stem) + suffix)
                                for suffix in (".callgrind.4", ".inclusive.txt", ".self.txt"))
    require(all(path.is_file() for path in (raw, inc_path, self_path)),
            f"missing profile input: {stem}")
    inc, self_text = text(inc_path), text(self_path)
    data = {name: annotation(inc, self_text, name, rel(inc_path))
            for name in (SCAN, START, ADDRESS)}
    for name, owner_data in data.items():
        validate_children(raw, name, owner_data)

    old = prior_row(prior, repeat, shape)["owners"]
    for name in (SCAN, START, ADDRESS):
        require(data[name]["inclusive_ir"] == old[name]["inclusive_ir"]
                and data[name]["self_ir"] == old[name]["self_ir"]
                and data[name]["direct"] == old[name]["direct"],
                f"0526 retained owner changed: {name} {repeat}/{shape}")
    require(data[SCAN]["direct"].get(START) == data[START]["inclusive_ir"],
            "scanner -> start_cell chain differs")
    require(data[START]["direct"].get(ADDRESS) == data[ADDRESS]["inclusive_ir"],
            "start_cell -> cell_address chain differs")

    edge = HELPER.target_edge_summary(raw, RESOLVE, SCAN)
    annotation_ir = data[SCAN]["direct"].get(RESOLVE)
    require(annotation_ir is not None and edge["positive_edge_count"] == 1
            and edge["inclusive_ir"] == annotation_ir,
            f"scanner -> resolve_event edge differs: {rel(raw)}")
    commit_ir = old[OWNER]["inclusive_ir"]
    return {
        "repeat": repeat, "shape": shape,
        "inputs": {rel(path): digest(path) for path in (raw, inc_path, self_path)},
        "commit_inclusive_ir": commit_ir, "owners": data,
        "resolver_edge": {
            "annotation_direct_ir": annotation_ir, "raw_edge": raw_view(edge),
            "inclusive_ir": edge["inclusive_ir"],
            "scan_inclusive_ir": data[SCAN]["inclusive_ir"],
            "commit_inclusive_ir": commit_ir,
            "share_of_scan_percent": edge["inclusive_ir"] / data[SCAN]["inclusive_ir"] * 100,
            "share_of_commit_percent": edge["inclusive_ir"] / commit_ir * 100,
            "upper_bound_only": True,
            "upper_bound_scope": "complete scanner -> resolve_event edge; Start and Empty lookups remain required",
            "raw_edge_matches_annotation": True,
        },
        "start_cell_direct_children_ranked": ranked(
            data[START], data[START]["inclusive_ir"], "share_of_start_cell_percent"),
        "cell_address_direct_children_ranked": ranked(
            data[ADDRESS], data[ADDRESS]["inclusive_ir"], "share_of_cell_address_percent"),
        "validation": {
            "raw_direct_edges_match_annotations": True, "self_plus_direct_equations": True,
            "prior_0526_owner_values_match": True,
            "scanner_to_start_cell_to_cell_address": True,
            "resolver_raw_edge_matches_annotation": True,
            "call_counts_used_for_estimate": False,
        },
    }


def aggregate(rows: list[dict[str, Any]]) -> dict[str, Any]:
    owners: dict[str, dict[str, Any]] = {
        OWNER: {"inclusive_ir": sum(row["commit_inclusive_ir"] for row in rows)},
    }
    for name in (SCAN, START, ADDRESS):
        owners[name] = {"inclusive_ir": 0, "self_ir": 0, "direct": {}}
    for row in rows:
        for name, value in row["owners"].items():
            out = owners[name]
            out["inclusive_ir"] += value["inclusive_ir"]
            out["self_ir"] += value["self_ir"]
            for child, cost in value["direct"].items():
                out["direct"][child] = out["direct"].get(child, 0) + cost
    for name in (SCAN, START, ADDRESS):
        out = owners[name]
        out["direct"] = dict(sorted(out["direct"].items(), key=lambda item: (-item[1], item[0])))
        out["equation"] = {
            "inclusive_ir": out["inclusive_ir"], "self_ir": out["self_ir"],
            "direct_sum_ir": sum(out["direct"].values()),
            "self_plus_direct_ir": out["self_ir"] + sum(out["direct"].values()),
            "verified": out["inclusive_ir"] == out["self_ir"] + sum(out["direct"].values()),
        }
        out["share_of_commit_percent"] = out["inclusive_ir"] / owners[OWNER]["inclusive_ir"] * 100

    resolver_ir = sum(row["resolver_edge"]["inclusive_ir"] for row in rows)
    resolver = {
        "inclusive_ir": resolver_ir, "scan_inclusive_ir": owners[SCAN]["inclusive_ir"],
        "commit_inclusive_ir": owners[OWNER]["inclusive_ir"],
        "share_of_scan_percent": resolver_ir / owners[SCAN]["inclusive_ir"] * 100,
        "share_of_commit_percent": resolver_ir / owners[OWNER]["inclusive_ir"] * 100,
        "upper_bound_only": True,
        "upper_bound_scope": "complete scanner -> resolve_event edge; Start and Empty lookups remain required",
        "raw_edge_matches_annotation_per_row": True, "call_counts_used_for_estimate": False,
    }
    return {
        "owners": owners, "scanner_to_resolve_event": resolver,
        "scanner_to_start_cell_to_cell_address": {
            "start_cell_inclusive_ir": owners[START]["inclusive_ir"],
            "cell_address_inclusive_ir": owners[ADDRESS]["inclusive_ir"],
            "start_cell_direct_children_ranked": ranked(
                owners[START], owners[START]["inclusive_ir"], "share_of_start_cell_percent"),
            "cell_address_direct_children_ranked": ranked(
                owners[ADDRESS], owners[ADDRESS]["inclusive_ir"], "share_of_cell_address_percent"),
            "nested_costs_not_added_to_scanner_again": True,
        },
        "prior_0526_selected_owner_values_match": True,
    }


def phase_row(repeat: int, shape: str) -> dict[str, Any]:
    path = P0527 / "baseline" / f"native-r{repeat}-primary-{shape}.json"
    result = read(path)["results"]
    require(len(result) == 1, f"unexpected native result count: {rel(path)}")
    result = result[0]
    xlsx = result["source"]["xlsx_cell_values"]
    require(xlsx["timing_scope"] == "open, selector planning, commit, and stream publication; reopen/verification is separate and excluded",
            f"phase scope differs: {rel(path)}")
    vectors = {phase: xlsx[phase] for phase in PHASES}
    require(all(isinstance(values, list) and len(values) == 200
                and all(isinstance(value, int) and value >= 0 for value in values)
                for values in vectors.values()), f"invalid phase vector: {rel(path)}")
    sums = {phase: sum(values) for phase, values in vectors.items()}
    total = sum(sums.values())
    mean = total / 200
    require(math.isclose(mean, result["elapsed_ns"]["mean"], rel_tol=0, abs_tol=1e-6),
            f"phase sum does not reproduce elapsed mean: {rel(path)}")
    return {
        "repeat": repeat, "shape": shape, "input": input_hash(path), "sample_count": 200,
        "phase_sample_sums_ns": sums,
        "phase_means_ns": {phase: value / 200 for phase, value in sums.items()},
        "total_phase_sample_sum_ns": total, "total_phase_mean_ns": mean,
        "reported_elapsed_mean_ns": result["elapsed_ns"]["mean"],
        "phase_mean_share_percent": {phase: value / total * 100 for phase, value in sums.items()},
        "formula": "phase_mean_share = sum_i(phase_i) / sum_i(sum_p(phase_p)); medians are not summed",
        "reopen_excluded": True, "validated": True,
    }


def analyze() -> dict[str, Any]:
    seal_entries = check_seal(P0525)
    prior = read(P0526 / "profile-analysis.json")
    require(prior.get("status") == "pass", "0526 profile report is not passing")
    rows = [profile_row(repeat, shape, prior) for repeat, shape in ROWS]
    return {
        "schema": "litchi-0528-scanner-ceiling-audit-v1", "status": "pass",
        "scope": "Read-only sealed 0525 XLSX scanner attribution and 0527 phase context; no build, capture, latency, or production edit.",
        "selected_owner": OWNER, "inputs": inputs(), "source_binding": source_binding(),
        "prior_0525_seal_entries": seal_entries, "rows": rows,
        "aggregate": aggregate(rows),
        "phase_context": {
            "source": "0527 sealed baseline native primary phase vectors",
            "phases": list(PHASES), "rows": [phase_row(repeat, shape) for repeat, shape in ROWS],
            "scope": "per-sample vector means for open/plan/commit/publication; reopen is separate",
            "medians_summed": False,
        },
        "limitations": [
            "The complete scanner -> resolve_event edge is an upper bound for End-only resolver work; Start and Empty lookups remain required.",
            "Call metadata is collection-off polluted and no per-call split or estimate is reported.",
            "Nested inclusive owners overlap and must not be added together.",
            "Callgrind Ir is a mechanism diagnostic, not a native latency, RSS, allocation, I/O, or end-to-end speedup claim.",
        ],
        "next_roi": {
            "ranking_basis": "aggregate immediate-child inclusive Ir within each parent",
            "first": "Scanner::start_cell -> Scanner::cell_address",
            "second": "Scanner::start_cell -> wire::cell_tag",
            "resolver_ceiling_is_not_end_only": True, "publication_phase_context_only": True,
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    output = json.dumps(analyze(), indent=2, sort_keys=True) + "\n"
    if args.output:
        args.output.write_text(output, encoding="utf-8")
    else:
        print(output, end="")


if __name__ == "__main__":
    main()
