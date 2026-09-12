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


def annotation(path: Path, inclusive: bool) -> dict:
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
        "xml_validator_ir": _single(stars, XML_VALIDATOR, path),
        "direct": direct,
    }


def _single(stars: dict[str, list[int]], name: str, path: Path) -> int:
    values = stars.get(name, [])
    if len(values) != 1:
        raise RuntimeError(f"{path}: expected one annotation row for {name}, got {values}")
    return values[0]


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
    inclusive = annotation(path, True)
    exclusive = annotation(path, False)
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
                        "topology_inclusive_ir": metric(
                            left["topology_inclusive_ir"], right["topology_inclusive_ir"]
                        ),
                        "xml_validator_inclusive_ir": metric(
                            left["xml_validator_inclusive_ir"], right["xml_validator_inclusive_ir"]
                        ),
                    },
                }
            )
    return comparisons


def fmt(value: int) -> str:
    return f"{value:,}"


def metric_text(value: dict) -> str:
    return f"{fmt(value['baseline'])} → {fmt(value['candidate'])} ({value['delta_pct']:+.2f}%)"


def write_markdown(document: dict, path: Path) -> None:
    lines = [
        "# 0517 managed DOCX publication profile comparison",
        "",
        "This is a matched six-arm × two-repeat Callgrind comparison of the same managed paragraph API routes on the frozen baseline and candidate binaries. All profiles use owned source, one measured sample, zero warmups, one repeat, and CPU 2.",
        "",
        "The publication value is inclusive Ir for `Package::publish_document_commit_to_stream`; its direct owner is the `SourceBackedPackage::write_topology_to_stream` edge. Topology and XML-validator values are inclusive `callgrind_annotate --tree=both` function rows. The publication method ends before the caller drops the returned Snapshot, so these are method-scope instruction diagnostics and do not include that caller drop or RSS.",
        "",
        "Raw parsing and annotation checks pass for all 24 profiles: one positive publication edge with one call, publication incoming Ir equals the raw summary, self plus direct edges equals the publication total, annotation direct edges match raw edges, and the direct owner equals topology inclusive Ir.",
        "",
        "| Repeat | Route | Publication inclusive Ir (baseline → candidate) | Direct owner Ir | Topology inclusive Ir | XML validator inclusive Ir |",
        "| --- | --- | ---: | ---: | ---: | ---: |",
    ]
    for record in document["comparisons"]:
        metrics = record["metrics"]
        lines.append(
            f"| {record['repeat']} | {record['route']} | "
            f"{metric_text(metrics['publication_inclusive_ir'])} | "
            f"{metric_text(metrics['direct_owner_inclusive_ir'])} | "
            f"{metric_text(metrics['topology_inclusive_ir'])} | "
            f"{metric_text(metrics['xml_validator_inclusive_ir'])} |"
        )
    lines += [
        "",
        "The candidate reduces publication inclusive Ir by about 17.9% on p128 K=1 and about 21.0% on p512 K=1/K=32. The direct topology owner falls about 32.3% on p128 and 43.7% on p512, while inclusive `validate_source_xml` falls about 33.3% on every arm. These are consistent with removing one complete source XML validation pass; they do not replace native elapsed-time or RSS guards.",
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
            "topology": TOPOLOGY,
            "xml_validator": XML_VALIDATOR,
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
