#!/usr/bin/env python3
"""Compare the six matched managed-DOCX Callgrind lanes."""

from __future__ import annotations

import argparse
import json
import re
import subprocess
from pathlib import Path

import analyze_profiles as raw


HERE = Path(__file__).resolve().parent
PUBLICATION = "litchi_docx::source_backed::Package::publish_document_commit_to_stream"
TOPOLOGY = "litchi_opc::source_backed::SourceBackedPackage::write_topology_to_stream"
XML_VALIDATOR = "litchi_opc::xml_splice::validate_source_xml"
SNAPSHOT_OWNER_PREFIX = "litchi_docx::source_backed::Package::main_document_snapshot"
SNAPSHOT_OWNER_NAMES = frozenset(
    {
        SNAPSHOT_OWNER_PREFIX,
        SNAPSHOT_OWNER_PREFIX + "_with_hint",
        SNAPSHOT_OWNER_PREFIX + "_with_before",
    }
)
FRESH_DOCX_SCAN_PREFIX = "litchi_docx::source_backed::ensure_source_document_xml"
FRESH_SNAPSHOT_BUILD_PREFIX = (
    "litchi_docx::document::transaction::Snapshot::from_source_xml"
)
# The baseline records its existing validation row.  The 0519 candidate is the
# source-proof reuse hit: no positive validator call or annotation row is
# expected in the publication-method scope.  Keep this as a route policy,
# rather than comparing against a copied prior call count.
XML_VALIDATOR_POLICY = {
    "baseline": "present",
    "candidate": "absent",
}
ARMS = (
    "p128-k1-owned-batch",
    "p128-k1-owned-repeated",
    "p512-k1-owned-batch",
    "p512-k1-owned-repeated",
    "p512-k32-owned-batch",
    "p512-k32-owned-repeated",
)
VARIANTS = {
    "baseline": ("profile-r1", "profile-r2"),
    "candidate": ("profile-after-r1", "profile-after-r2"),
}
STAR_RE = re.compile(r"^\s*([\d,]+)\s+\*\s+(.*)$")
EDGE_RE = re.compile(r"^\s*([\d,]+)\s+>\s+(.*)$")


def display_name(text: str) -> str:
    text = text.split(" [", 1)[0].strip()
    text = re.sub(r"\s+\([\d,]+x\)$", "", text)
    if text.startswith("???:"):
        return text[4:]
    # libc rows retain a source path, while Rust rows use `::` in names.
    return text.rsplit(":", 1)[-1] if text.startswith(("./", "/")) else text


def option(command: list[str], name: str) -> str:
    index = command.index(name)
    return command[index + 1]


def api_options(receipt: dict) -> dict[str, str]:
    command = receipt["command"]
    return {
        name: option(command, "--" + name)
        for name in ("paragraphs", "replacements", "source", "mode", "samples", "warmups", "repeats")
    }


def route_for(api: dict[str, str]) -> str:
    return (
        f"p{api['paragraphs']}-k{api['replacements']}-"
        f"{api['source']}-{api['mode']}"
    )


def is_snapshot_owner(name: str) -> bool:
    """Recognize the source-snapshot owner across baseline/candidate spellings."""
    return name in SNAPSHOT_OWNER_NAMES


def raw_positive_function_names(path: Path, prefixes: tuple[str, ...]) -> dict[str, list[str]]:
    """Return named functions with positive Ir in the scoped raw profile.

    A candidate hit may inline the source XML handoff, so the absence check is
    based on positive raw costs instead of requiring an annotation row for a
    particular helper.  The first pass resolves names for function records;
    the second pass marks a function only when its own section has positive Ir.
    """
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    events: list[str] = []
    names: dict[int, str] = {}
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("events:"):
            events = stripped.split(":", 1)[1].split()
        match = raw.FUNCTION_RE.match(stripped)
        if match and match.group(3):
            names[int(match.group(2))] = match.group(3)
    if "Ir" not in events:
        raise RuntimeError(f"{path}: raw profile has no Ir event")
    event_index = events.index("Ir")
    positive: dict[str, set[str]] = {prefix: set() for prefix in prefixes}
    current_id: int | None = None
    for line in lines:
        stripped = line.strip()
        match = raw.FUNCTION_RE.match(stripped)
        if match and match.group(1) == "fn":
            current_id = int(match.group(2))
            continue
        if current_id is None:
            continue
        cost = raw._cost(stripped, event_index)
        if cost is None or cost <= 0:
            continue
        name = names.get(current_id, "")
        for prefix in prefixes:
            if name.startswith(prefix):
                positive[prefix].add(name)
    return {prefix: sorted(values) for prefix, values in positive.items()}


def raw_incoming_call_summary(path: Path, target: str) -> dict:
    """Count positive raw call edges into one function in the scoped profile."""
    lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    events: list[str] = []
    names: dict[int, str] = {}
    for line in lines:
        stripped = line.strip()
        if stripped.startswith("events:"):
            events = stripped.split(":", 1)[1].split()
        match = raw.FUNCTION_RE.match(stripped)
        if match and match.group(3):
            names[int(match.group(2))] = match.group(3)
    if "Ir" not in events:
        raise RuntimeError(f"{path}: raw profile has no Ir event")
    event_index = events.index("Ir")
    current_id: int | None = None
    pending_callee_id: int | None = None
    pending_calls: int | None = None
    edges: list[dict] = []
    for line in lines:
        stripped = line.strip()
        match = raw.FUNCTION_RE.match(stripped)
        if match:
            kind = match.group(1)
            function_id = int(match.group(2))
            if kind == "fn":
                current_id = function_id
                pending_callee_id = pending_calls = None
            else:
                if current_id is None:
                    raise RuntimeError(f"{path}: cfn outside fn")
                pending_callee_id = function_id
                pending_calls = None
            continue
        calls_match = raw.CALLS_RE.match(stripped)
        if calls_match:
            if current_id is None or pending_callee_id is None:
                raise RuntimeError(f"{path}: calls outside cfn")
            pending_calls = int(calls_match.group(1).replace(",", ""))
            continue
        if current_id is None:
            continue
        cost = raw._cost(stripped, event_index)
        if cost is None:
            continue
        if pending_callee_id is not None and pending_calls is not None:
            if cost > 0 and pending_calls > 0:
                edges.append(
                    {
                        "caller": names.get(current_id, ""),
                        "callee": names.get(pending_callee_id, ""),
                        "calls": pending_calls,
                        "inclusive_ir": cost,
                    }
                )
            pending_callee_id = pending_calls = None
    matching = [edge for edge in edges if edge["callee"] == target]
    return {
        "target": target,
        "positive_edge_count": len(matching),
        "calls": sum(edge["calls"] for edge in matching),
        "edges": matching,
    }


def annotation(path: Path, inclusive: bool, xml_validator_policy: str) -> dict:
    output = subprocess.check_output(
        [
            "callgrind_annotate",
            "--auto=no",
            "--threshold=100",
            "--show-percs=no",
            "--inclusive=" + ("yes" if inclusive else "no"),
            "--tree=both",
            str(path),
        ],
        text=True,
    )
    lines = output.splitlines()
    stars: dict[str, list[int]] = {}
    for line in lines:
        match = STAR_RE.match(line)
        if match:
            stars.setdefault(display_name(match.group(2)), []).append(
                int(match.group(1).replace(",", ""))
            )

    target_hits = [
        index
        for index, line in enumerate(lines)
        if (match := STAR_RE.match(line))
        and display_name(match.group(2)) == PUBLICATION
    ]
    if len(target_hits) != 1:
        raise RuntimeError(f"{path}: expected one publication annotation row")
    direct: list[dict] = []
    for line in lines[target_hits[0] + 1 :]:
        match = EDGE_RE.match(line)
        if not match:
            break
        direct.append(
            {
                "name": display_name(match.group(2)),
                "inclusive_ir": int(match.group(1).replace(",", "")),
            }
        )
    return {
        "publication_ir": stars[PUBLICATION][0],
        "topology_ir": _single(stars, TOPOLOGY, path),
        "xml_validator_ir": _optional_single(stars, XML_VALIDATOR, path),
        "direct": direct,
    }


def _single(stars: dict[str, list[int]], name: str, path: Path) -> int:
    values = stars.get(name, [])
    if len(values) != 1:
        raise RuntimeError(f"{path}: expected one annotation row for {name}, got {values}")
    return values[0]


def _optional_single(stars: dict[str, list[int]], name: str, path: Path) -> int | None:
    values = stars.get(name, [])
    if len(values) > 1:
        raise RuntimeError(f"{path}: expected at most one annotation row for {name}, got {values}")
    return values[0] if values else None


def direct_map(edges: list[dict]) -> dict[str, int]:
    values: dict[str, int] = {}
    for edge in edges:
        name = edge.get("callee", edge.get("name"))
        values[name] = values.get(name, 0) + edge["inclusive_ir"]
    return values


def load_profile(path: Path, variant: str, repeat: str) -> dict:
    parsed = raw.parse_profile(path, PUBLICATION)
    receipt_path = path.with_suffix(".json")
    receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    api = api_options(receipt)
    route = route_for(api)
    if route not in ARMS:
        raise RuntimeError(f"{path}: unexpected route {route}")
    owner_edges = [
        edge for edge in parsed["selected_direct_callees"] if edge["callee"] == TOPOLOGY
    ]
    if len(owner_edges) != 1:
        raise RuntimeError(f"{path}: expected one direct topology owner edge")
    owner = owner_edges[0]
    snapshot_edges = [
        edge
        for edge in parsed["selected_direct_callees"]
        if is_snapshot_owner(edge["callee"])
    ]
    if len(snapshot_edges) > 1:
        raise RuntimeError(f"{path}: ambiguous semantic current-snapshot owner edges")
    if variant == "baseline" and len(snapshot_edges) != 1:
        raise RuntimeError(f"{path}: baseline requires one semantic current-snapshot owner edge")
    snapshot_owner = snapshot_edges[0] if snapshot_edges else None
    xml_validator_policy = XML_VALIDATOR_POLICY[variant]
    inclusive = annotation(path, True, xml_validator_policy)
    exclusive = annotation(path, False, xml_validator_policy)
    fresh_functions = raw_positive_function_names(
        path, (FRESH_DOCX_SCAN_PREFIX, FRESH_SNAPSHOT_BUILD_PREFIX)
    )
    xml_calls = raw_incoming_call_summary(path, XML_VALIDATOR)
    raw_direct = direct_map(parsed["selected_direct_callees"])
    annotated_direct = direct_map(inclusive["direct"])
    direct_edges_match = raw_direct == annotated_direct
    annotated_owner = [edge for edge in inclusive["direct"] if edge["name"] == TOPOLOGY]
    owner_matches = (
        len(annotated_owner) == 1
        and annotated_owner[0]["inclusive_ir"] == owner["inclusive_ir"]
    )
    checks = {
        "raw_exactly_one_positive_call": parsed["validation"]["exactly_one_positive_call"],
        "raw_publication_total_scoped": parsed["validation"][
            "publication_total_scoped_to_summary"
        ],
        "inclusive_publication_matches_raw": inclusive["publication_ir"]
        == parsed["selected_inclusive_ir"],
        "exclusive_self_matches_raw": exclusive["publication_ir"]
        == parsed["selected_self_ir"],
        "inclusive_direct_edges_match_raw": direct_edges_match,
        "exclusive_direct_edges_match_inclusive": direct_map(exclusive["direct"])
        == annotated_direct,
        "direct_owner_matches_raw": owner_matches,
        "direct_owner_matches_topology_inclusive": owner["inclusive_ir"]
        == inclusive["topology_ir"],
        "semantic_snapshot_owner_requirement_satisfied": (
            variant == "candidate" or snapshot_owner is not None
        ),
        "xml_validator_annotation_policy_satisfied": (
            inclusive["xml_validator_ir"] is not None
            if xml_validator_policy == "present"
            else inclusive["xml_validator_ir"] is None
        ),
        "xml_validator_call_policy_satisfied": (
            xml_calls["calls"] > 0
            if xml_validator_policy == "present"
            else xml_calls["calls"] == 0
        ),
        "candidate_fresh_docx_scan_absent": (
            variant != "candidate" or not fresh_functions[FRESH_DOCX_SCAN_PREFIX]
        ),
        "candidate_fresh_snapshot_build_absent": (
            variant != "candidate" or not fresh_functions[FRESH_SNAPSHOT_BUILD_PREFIX]
        ),
    }
    if not all(checks.values()):
        raise RuntimeError(f"{path}: annotation checks failed: {checks}")
    return {
        "variant": variant,
        "repeat": repeat,
        "route": route,
        "case": route.rsplit("-", 1)[0],
        "mode": api["mode"],
        "api": api,
        "profile": str(path.relative_to(HERE)),
        "profile_sha256": parsed["sha256"],
        "receipt": str(receipt_path.relative_to(HERE)),
        "binary_sha256": parsed["binary_sha256"],
        "source_manifest_sha256": parsed["source_manifest_sha256"],
        "plan_sha256": receipt.get("plan_sha256"),
        "candidate_plan_sha256": receipt.get("candidate_plan_sha256"),
        "publication_inclusive_ir": parsed["selected_inclusive_ir"],
        "publication_self_ir": parsed["selected_self_ir"],
        "publication_direct_ir": parsed["selected_direct_ir"],
        "publication_calls": parsed["selected_call_count"],
        "direct_owner": {
            "name": TOPOLOGY,
            "calls": owner["calls"],
            "inclusive_ir": owner["inclusive_ir"],
        },
        "snapshot_owner": (
            {
                "name": snapshot_owner["callee"],
                "calls": snapshot_owner["calls"],
                "inclusive_ir": snapshot_owner["inclusive_ir"],
            }
            if snapshot_owner is not None
            else None
        ),
        "semantic_snapshot_owner_present": snapshot_owner is not None,
        "fresh_positive_functions": fresh_functions,
        "xml_validator_calls": xml_calls,
        "xml_validator_policy": xml_validator_policy,
        "topology_inclusive_ir": inclusive["topology_ir"],
        "xml_validator_inclusive_ir": inclusive["xml_validator_ir"],
        "annotation_checks": checks,
    }


def load_variant(variant: str) -> dict[tuple[str, str], dict]:
    result: dict[tuple[str, str], dict] = {}
    for lane in VARIANTS[variant]:
        repeat = lane.rsplit("-", 1)[1]
        paths = sorted((HERE / lane).glob("*.callgrind"))
        if len(paths) != len(ARMS):
            raise RuntimeError(f"{lane}: expected {len(ARMS)} profiles, got {len(paths)}")
        for path in paths:
            profile = load_profile(path, variant, repeat)
            key = (repeat, profile["route"])
            if key in result:
                raise RuntimeError(f"{variant}: duplicate profile {key}")
            result[key] = profile
    expected = {(repeat, arm) for repeat in ("r1", "r2") for arm in ARMS}
    if set(result) != expected:
        raise RuntimeError(f"{variant}: profile matrix differs: {sorted(result)}")
    return result


def binding(result: dict[tuple[str, str], dict]) -> dict:
    fields = ("binary_sha256", "source_manifest_sha256", "plan_sha256", "candidate_plan_sha256")
    values = {field: sorted({profile.get(field) for profile in result.values()}) for field in fields}
    return values


def metric(left: int, right: int) -> dict:
    ratio = right / left
    return {
        "baseline": left,
        "candidate": right,
        "delta_ir": right - left,
        "ratio_candidate_over_baseline": ratio,
        "delta_pct": (ratio - 1.0) * 100.0,
    }


def optional_metric(left: int | None, right: int | None) -> dict:
    """Compare an optional semantic owner call without inventing a zero cost."""
    if left is None or right is None:
        return {
            "baseline": left,
            "candidate": right,
            "delta_ir": None,
            "ratio_candidate_over_baseline": None,
            "delta_pct": None,
            "candidate_call_absent": right is None,
        }
    return metric(left, right) | {"candidate_call_absent": False}


def compare(baseline: dict[tuple[str, str], dict], candidate: dict[tuple[str, str], dict]) -> list[dict]:
    comparisons = []
    for repeat in ("r1", "r2"):
        for arm in ARMS:
            left = baseline[repeat, arm]
            right = candidate[repeat, arm]
            same_api = left["api"] == right["api"]
            if not same_api:
                raise RuntimeError(f"{repeat}/{arm}: baseline/candidate CLI differs")
            comparisons.append(
                {
                    "repeat": repeat,
                    "route": arm,
                    "case": left["case"],
                    "mode": left["mode"],
                    "same_api": same_api,
                    "baseline": left,
                    "candidate": right,
                    "metrics": {
                        "publication_inclusive_ir": metric(
                            left["publication_inclusive_ir"], right["publication_inclusive_ir"]
                        ),
                        "direct_owner_inclusive_ir": metric(
                            left["direct_owner"]["inclusive_ir"], right["direct_owner"]["inclusive_ir"]
                        ),
                        "snapshot_owner_inclusive_ir": optional_metric(
                            left["snapshot_owner"]["inclusive_ir"]
                            if left["snapshot_owner"]
                            else None,
                            right["snapshot_owner"]["inclusive_ir"]
                            if right["snapshot_owner"]
                            else None,
                        ),
                        "topology_inclusive_ir": metric(
                            left["topology_inclusive_ir"], right["topology_inclusive_ir"]
                        ),
                        "xml_validator_inclusive_ir": optional_metric(
                            left["xml_validator_inclusive_ir"], right["xml_validator_inclusive_ir"]
                        ),
                    },
                }
            )
    return comparisons


def fmt(value: int) -> str:
    return f"{value:,}"


def metric_text(value: dict) -> str:
    if value["candidate"] is None:
        baseline = "—" if value["baseline"] is None else fmt(value["baseline"])
        return f"{baseline} → — (candidate direct call absent)"
    return f"{fmt(value['baseline'])} → {fmt(value['candidate'])} ({value['delta_pct']:+.2f}%)"


def write_markdown(document: dict, path: Path) -> None:
    lines = [
        "# 0519 managed DOCX publication profile comparison",
        "",
        "This is a matched six-arm × two-repeat Callgrind comparison of the same managed paragraph API routes on the frozen baseline and candidate binaries. All profiles use owned source, one measured sample, zero warmups, one repeat, and CPU 2.",
        "",
        "The publication value is inclusive Ir for `Package::publish_document_commit_to_stream`; its direct owner is the `SourceBackedPackage::write_topology_to_stream` edge. Topology and XML-validator values are inclusive `callgrind_annotate --tree=both` function rows. The publication method ends before the caller drops the returned Snapshot, so these are method-scope instruction diagnostics and do not include that caller drop or RSS. The candidate snapshot-owner edge is optional: an optimized hit may inline the source-XML handoff, and the table uses an em dash when no direct semantic snapshot call remains. The candidate XML-validator row is also intentionally absent when the retained source proof is accepted.",
        "",
        "Raw parsing and annotation checks pass for all 24 profiles: one positive publication edge with one call, publication incoming Ir equals the raw summary, self plus direct edges equals the publication total, annotation direct edges match raw edges, the direct owner equals topology inclusive Ir, and candidate hit profiles have no positive raw cost in either fresh DOCX scan/build helper. The baseline XML-validator policy is present; the candidate source-proof-reuse policy is absent in both positive raw calls and annotation rows.",
        "",
        "| Repeat | Route | Publication inclusive Ir (baseline → candidate) | Direct owner Ir | Snapshot owner Ir | Topology inclusive Ir | XML validator inclusive Ir |",
        "| --- | --- | ---: | ---: | ---: | ---: | ---: |",
    ]
    for record in document["comparisons"]:
        metrics = record["metrics"]
        lines.append(
            f"| {record['repeat']} | {record['route']} | "
            f"{metric_text(metrics['publication_inclusive_ir'])} | "
            f"{metric_text(metrics['direct_owner_inclusive_ir'])} | "
            f"{metric_text(metrics['snapshot_owner_inclusive_ir'])} | "
            f"{metric_text(metrics['topology_inclusive_ir'])} | "
            f"{metric_text(metrics['xml_validator_inclusive_ir'])} |"
        )
    lines += [
        "",
        "Use the per-route deltas to attribute candidate SourceXmlPart proof reuse and XML-validator elision. A Callgrind reduction is an instruction-count mechanism diagnostic; it does not replace native elapsed-time or RSS guards.",
        "",
        "Baseline profiles: `profile-r1/` and `profile-r2/`. Candidate profiles: `profile-after-r1/` and `profile-after-r2/`. Full raw/direct-callee details remain in `profile-analysis.json` and `profile-after-analysis.json`.",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> int:
    parser_cli = argparse.ArgumentParser(description=__doc__)
    parser_cli.add_argument("--json", type=Path, default=HERE / "profile-comparison.json")
    parser_cli.add_argument("--markdown", type=Path, default=HERE / "profile-comparison.md")
    args = parser_cli.parse_args()
    baseline = load_variant("baseline")
    candidate = load_variant("candidate")
    comparisons = compare(baseline, candidate)
    validation = {
        "baseline_profile_count": len(baseline),
        "candidate_profile_count": len(candidate),
        "same_api_for_all_pairs": all(record["same_api"] for record in comparisons),
        "baseline_exact_one_positive_call": all(
            profile["annotation_checks"]["raw_exactly_one_positive_call"]
            and profile["publication_calls"] == 1
            for profile in baseline.values()
        ),
        "candidate_exact_one_positive_call": all(
            profile["annotation_checks"]["raw_exactly_one_positive_call"]
            and profile["publication_calls"] == 1
            for profile in candidate.values()
        ),
        "all_annotation_checks": all(
            all(profile["annotation_checks"].values())
            for profile in (*baseline.values(), *candidate.values())
        ),
    }
    document = {
        "schema": "managed_docx_publication_profile_comparison_v1",
        "matrix": {
            "arms": list(ARMS),
            "repeats": ["r1", "r2"],
            "profiles_per_variant": len(baseline),
            "matched_pairs": len(comparisons),
        },
        "scope": "Callgrind Ir publication-method diagnostic; owned source; one measured sample; zero warmups; one harness repeat; publication excludes caller Snapshot drop; no RSS claim",
        "functions": {
            "publication": PUBLICATION,
            "direct_owner": TOPOLOGY,
            "snapshot_owner": sorted(SNAPSHOT_OWNER_NAMES),
            "fresh_docx_scan": FRESH_DOCX_SCAN_PREFIX,
            "fresh_snapshot_build": FRESH_SNAPSHOT_BUILD_PREFIX,
            "topology": TOPOLOGY,
            "xml_validator": XML_VALIDATOR,
            "xml_validator_policy": XML_VALIDATOR_POLICY,
        },
        "bindings": {"baseline": binding(baseline), "candidate": binding(candidate)},
        "validation": validation,
        "comparisons": comparisons,
    }
    args.json.parent.mkdir(parents=True, exist_ok=True)
    args.json.write_text(json.dumps(document, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    write_markdown(document, args.markdown)
    print(json.dumps({"matched_pairs": len(comparisons), "validation": validation}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
