#!/usr/bin/env python3
"""Derive 0529 baseline primary phase shares from validated sample vectors.

This is a small context report, rather than a performance comparison.  It
loads the retained 0529 analyzer and asks it to validate the baseline stage,
then sums the validated 200-sample phase vectors.  Shares use the sum of the
four measured lifecycle phases as their denominator; post-publication
``reopen_ns`` is reported separately and is never folded into that share.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
from collections import defaultdict
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
RETAINED = REPO / "docs" / "performance" / "results" / "change-0529"
ANALYZER_PATH = RETAINED / "analyze.py"
PLAN_PATH = RETAINED / "plan.json"
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
ALL_PHASES = PHASES + ("reopen_ns",)
VECTOR_PHASES = ("elapsed_ns",) + ALL_PHASES


class ContextError(ValueError):
    """A missing or inconsistent retained evidence input."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ContextError(message)


def sha(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise ContextError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise ContextError(f"cannot read JSON {path}: {error}") from error


def load_retained_analyzer() -> Any:
    spec = importlib.util.spec_from_file_location("xlsx_0529_context_analyzer", ANALYZER_PATH)
    require(spec is not None and spec.loader is not None,
            f"cannot load retained analyzer: {ANALYZER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def nonnegative_integer(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a nonnegative integer")
    return value


def vector(row: dict[str, Any], phase: str) -> list[int]:
    timing = row.get("timing")
    require(isinstance(timing, dict), f"{row.get('name')}.timing is missing")
    stats = timing.get(phase)
    require(isinstance(stats, dict), f"{row.get('name')}.{phase} timing is missing")
    values = stats.get("samples")
    require(isinstance(values, list) and values,
            f"{row.get('name')}.{phase}.samples is missing")
    return [nonnegative_integer(value, f"{row.get('name')}.{phase}[{index}]")
            for index, value in enumerate(values)]


def phase_totals(rows: Iterable[dict[str, Any]], label: str) -> dict[str, Any]:
    rows = list(rows)
    require(rows, f"{label} has no rows")
    vectors = {phase: [value for row in rows for value in vector(row, phase)]
               for phase in VECTOR_PHASES}
    sample_count = len(vectors["elapsed_ns"])
    for phase in VECTOR_PHASES:
        require(len(vectors[phase]) == sample_count,
                f"{label}.{phase} sample count differs")
    sums = {phase: sum(values) for phase, values in vectors.items()}
    lifecycle_sum = sum(sums[phase] for phase in PHASES)
    elapsed_sum = sums["elapsed_ns"]
    # The retained analyzer checks phase sums in acquisition order.  This
    # second sum equation ensures the aggregate here uses the same measured
    # lifecycle phases before deriving shares.
    require(lifecycle_sum == elapsed_sum,
            f"{label} lifecycle phase sums do not equal elapsed sum")
    require(elapsed_sum > 0, f"{label} elapsed sum is not positive")

    def measure(phase: str, denominator: int | None = None) -> dict[str, Any]:
        total = sums[phase]
        count = len(vectors[phase])
        result: dict[str, Any] = {
            "sum_ns": total,
            "sample_count": count,
            "mean_ns": total / count,
        }
        if denominator is not None:
            result["share_of_elapsed_percent"] = total / denominator * 100.0
            result["denominator_elapsed_sum_ns"] = denominator
        return result

    return {
        "label": label,
        "sample_count": sample_count,
        "elapsed": measure("elapsed_ns"),
        "phases": {phase: measure(phase, elapsed_sum) for phase in PHASES},
        "reopen": {
            **measure("reopen_ns"),
            "excluded_from_lifecycle_share": True,
            "reason": "post-publication verification is outside elapsed_ns phase-sum denominator",
        },
    }


def row_record(row: dict[str, Any]) -> dict[str, Any]:
    identity = row.get("identity")
    require(isinstance(identity, dict), f"{row.get('name')}.identity is missing")
    corpus = identity.get("corpus")
    require(isinstance(corpus, dict), f"{row.get('name')}.identity.corpus is missing")
    record = phase_totals([row], f"primary {row.get('repeat')}/{row.get('shape')}")
    record.update({
        "repeat": row.get("repeat"),
        "shape": row.get("shape"),
        "case": row.get("case"),
        "name": row.get("name"),
        "identity_sha256": row.get("identity_sha256"),
        "corpus": {
            "name": corpus.get("name"),
            "generator": corpus.get("generator"),
            "shape": corpus.get("shape"),
            "archive_sha256": corpus.get("archive_sha256"),
            "target_entry": corpus.get("target_entry"),
        },
        "validation": (
            "Source identity, corpus identity, output/sink identity, phase "
            "vectors, and report statistics were validated by the retained "
            "0529 analyzer before these sums were computed."
        ),
    })
    return record


def aggregate_records(rows: list[dict[str, Any]], label: str) -> dict[str, Any]:
    return phase_totals(rows, label)


def analyze() -> dict[str, Any]:
    require(ANALYZER_PATH.is_file(), f"retained analyzer is missing: {ANALYZER_PATH}")
    require(PLAN_PATH.is_file(), f"retained plan is missing: {PLAN_PATH}")
    analyzer = load_retained_analyzer()
    validated = analyzer.analyze("baseline")
    require(validated.get("status") == "pass" and validated.get("stage") == "baseline",
            "retained 0529 baseline analyzer did not pass")
    evidence = validated.get("evidence")
    require(isinstance(evidence, dict), "retained baseline evidence is missing")
    native = evidence.get("native")
    require(isinstance(native, dict) and isinstance(native.get("rows"), list),
            "retained baseline native rows are missing")

    plan = read_json(PLAN_PATH)
    require(isinstance(plan, dict), "retained plan is not an object")
    primary_plan = plan.get("primary")
    require(isinstance(primary_plan, dict), "retained primary plan is missing")
    expected = {(repeat, shape)
                for repeat in range(1, int(primary_plan["repeats"]) + 1)
                for shape in primary_plan["shapes"]}
    primary = [row for row in native["rows"]
               if isinstance(row, dict) and row.get("kind") == "primary"
               and row.get("guard") is None]
    actual = [(row.get("repeat"), row.get("shape")) for row in primary]
    require(len(primary) == len(expected) and len(set(actual)) == len(actual)
            and set(actual) == expected,
            "retained baseline primary matrix differs from the frozen plan")
    expected_samples = int(primary_plan["samples"])
    require(expected_samples == 200, "0529 primary sample count is not 200")
    for row in primary:
        require(row.get("samples") == expected_samples,
                f"{row.get('name')} sample count differs from plan")

    primary.sort(key=lambda row: (row["repeat"], row["shape"]))
    records = [row_record(row) for row in primary]
    by_shape: dict[str, list[dict[str, Any]]] = defaultdict(list)
    by_repeat: dict[int, list[dict[str, Any]]] = defaultdict(list)
    for row in primary:
        by_shape[row["shape"]].append(row)
        by_repeat[row["repeat"]].append(row)
    shape_aggregates = [aggregate_records(by_shape[shape], f"shape {shape}")
                        for shape in sorted(by_shape)]
    repeat_aggregates = [aggregate_records(by_repeat[repeat], f"repeat {repeat}")
                         for repeat in sorted(by_repeat)]
    overall = aggregate_records(primary, "all primary rows")

    phase_ranking = sorted(
        ((phase, overall["phases"][phase]["share_of_elapsed_percent"])
         for phase in PHASES),
        key=lambda item: (-item[1], item[0]),
    )
    source_manifest_sha = evidence.get("manifest_sha256")
    require(isinstance(source_manifest_sha, str) and len(source_manifest_sha) == 64,
            "retained baseline source manifest digest is missing")
    return {
        "schema": "litchi-0530-xlsx-planning-context-v1",
        "status": "pass",
        "scope": (
            "0529 retained baseline primary XLSX source-backed cell-values "
            "planning phase sums, arithmetic means, and elapsed shares"
        ),
        "method": {
            "source": "change-0529 baseline primary timing vectors after current analyzer validation",
            "phases_in_share_denominator": list(PHASES),
            "excluded_phases": {
                "reopen_ns": "reported separately; excluded from elapsed lifecycle phase shares",
            },
            "equation": "share_percent = sum(phase_ns) / sum(elapsed_ns) * 100",
            "mean_equation": "mean_ns = sum(phase_ns) / sample_count",
            "sample_count": expected_samples,
            "repeats": int(primary_plan["repeats"]),
            "used_medians": False,
            "used_latency_instruction_conversion": False,
            "used_profile_or_Ir_data": False,
        },
        "validation": {
            "analyzer": "docs/performance/results/change-0529/analyze.py",
            "analyzer_sha256": sha(ANALYZER_PATH),
            "analyzer_status": validated["status"],
            "plan_sha256": validated.get("plan_sha256"),
            "source_manifest_sha256": source_manifest_sha,
            "primary_rows_validated": len(primary),
            "source_identity_and_corpus_validated": True,
        },
        "primary_rows": records,
        "shape_aggregates": shape_aggregates,
        "repeat_aggregates": repeat_aggregates,
        "overall": overall,
        "phase_ranking_by_elapsed_share": [
            {"phase": phase, "share_of_elapsed_percent": share}
            for phase, share in phase_ranking
        ],
    }


def markdown(value: dict[str, Any]) -> str:
    overall = value["overall"]
    lines = [
        "# 0530 XLSX planning phase context",
        "",
        "Status: validated context report from the retained 0529 baseline.",
        "",
        f"Retained analyzer SHA-256: `{value['validation']['analyzer_sha256']}`.",
        f"Retained plan SHA-256: `{value['validation']['plan_sha256']}`.",
        f"Retained baseline source manifest SHA-256: `{value['validation']['source_manifest_sha256']}`.",
        "",
        "The current 0529 analyzer validated source identity, corpus identity, "
        "sink/output identity, phase vectors, and report statistics before this "
        "report summed the four 200-sample primary rows. Shares use sums and "
        "arithmetic means only. `reopen_ns` is reported separately because it "
        "is post-publication verification and is outside the lifecycle elapsed "
        "phase-sum denominator.",
        "",
        "## Primary rows",
        "",
        "| repeat | shape | samples | open mean/share | plan mean/share | commit mean/share | publication mean/share | reopen mean |",
        "| ---: | --- | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for row in value["primary_rows"]:
        phases = row["phases"]
        lines.append(
            f"| {row['repeat']} | {row['shape']} | {row['sample_count']} | "
            f"{phases['open_ns']['mean_ns']:.3f} ns / {phases['open_ns']['share_of_elapsed_percent']:.6f}% | "
            f"{phases['plan_ns']['mean_ns']:.3f} ns / {phases['plan_ns']['share_of_elapsed_percent']:.6f}% | "
            f"{phases['commit_ns']['mean_ns']:.3f} ns / {phases['commit_ns']['share_of_elapsed_percent']:.6f}% | "
            f"{phases['publication_ns']['mean_ns']:.3f} ns / {phases['publication_ns']['share_of_elapsed_percent']:.6f}% | "
            f"{row['reopen']['mean_ns']:.3f} ns |"
        )
    lines += [
        "",
        "## Aggregate context",
        "",
        f"Across all four primary rows ({overall['sample_count']} samples), the "
        "lifecycle phase shares are:",
        "",
        "| phase | sum | arithmetic mean | share of lifecycle elapsed |",
        "| --- | ---: | ---: | ---: |",
    ]
    for phase in PHASES:
        item = overall["phases"][phase]
        lines.append(
            f"| `{phase}` | {item['sum_ns']} ns | {item['mean_ns']:.3f} ns | "
            f"{item['share_of_elapsed_percent']:.6f}% |"
        )
    lines += [
        f"| `elapsed_ns` denominator | {overall['elapsed']['sum_ns']} ns | "
        f"{overall['elapsed']['mean_ns']:.3f} ns | 100.000000% |",
        "",
        f"`reopen_ns` sums to {overall['reopen']['sum_ns']} ns (mean "
        f"{overall['reopen']['mean_ns']:.3f} ns) and is excluded from the "
        "lifecycle share denominator.",
        "",
        "The aggregate share ranking is "
        + ", ".join(
            f"`{row['phase']}` {row['share_of_elapsed_percent']:.6f}%"
            for row in value["phase_ranking_by_elapsed_share"]
        )
        + ". These shares describe where measured lifecycle time was spent; "
        "they do not convert latency to instruction counts or make a new "
        "performance claim.",
        "",
        "The 0529 pilot was rejected on its native timing gates, so this context "
        "does not admit a conditional profile or an OLE2/OOXML speedup claim. "
        "ODF remains deferred under the active OLE2/OOXML priority.",
    ]
    return "\n".join(lines) + "\n"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--json-output", type=Path,
                        default=HERE / "planning-context.json")
    parser.add_argument("--markdown-output", type=Path,
                        default=HERE / "planning-context.md")
    args = parser.parse_args()
    try:
        value = analyze()
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")
        args.markdown_output.parent.mkdir(parents=True, exist_ok=True)
        args.markdown_output.write_text(markdown(value), encoding="utf-8")
    except ContextError as error:
        print(f"planning context failed: {error}")
        return 1
    print(f"wrote {args.json_output}")
    print(f"wrote {args.markdown_output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
