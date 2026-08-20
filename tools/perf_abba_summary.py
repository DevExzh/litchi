#!/usr/bin/env python3
"""Build a deterministic, fail-closed summary from four ABBA reports.

The input reports are the JSON files emitted by ``tools/perf-baseline``.  The
four legs are ordered ``A1, B1, B2, A2``: A is the control implementation and
B is the candidate implementation.  This module deliberately has no
third-party dependencies so that it can be used from a clean checkout.

The summary is descriptive evidence.  A statistic is marked accepted only
when both candidate directions are lower and both same-implementation drift
values are within the configured ceilings.  No speedup claim is inferred from
an accepted statistic.
"""

from __future__ import annotations

import argparse
import json
import math
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


SCHEMA_VERSION = 1
HARNESS_SCHEMA_VERSION = 1
TOOL_NAME = "litchi-perf-abba-summary"
TOOL_VERSION = "0.1.0"
LEG_ORDER = ("a1", "b1", "b2", "a2")
STATISTICS = ("p50", "mean", "p95", "p99")
ENVIRONMENT_VARIANTS = frozenset(("git_revision", "git_worktree_dirty"))
DEFAULT_DRIFT_CEILINGS: dict[str, float] = {
    "p50": 5.0,
    "mean": 5.0,
    "p95": 10.0,
    "p99": 15.0,
}


class AbbaSummaryInputError(ValueError):
    """Raised when ABBA reports are not safely comparable."""


def _canonical_json(value: Any, location: str) -> str:
    """Return compact canonical JSON, rejecting values JSON cannot preserve."""

    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError, OverflowError) as error:
        raise AbbaSummaryInputError(f"{location} is not canonical JSON: {error}") from error


def _require_object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise AbbaSummaryInputError(f"{location} must be an object")
    return value


def _finite_number(value: Any, location: str, *, positive: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise AbbaSummaryInputError(f"{location} must be a number")
    try:
        number = float(value)
    except (OverflowError, ValueError) as error:
        raise AbbaSummaryInputError(f"{location} is outside the finite range") from error
    if not math.isfinite(number):
        raise AbbaSummaryInputError(f"{location} must be finite")
    if positive and number <= 0:
        raise AbbaSummaryInputError(f"{location} must be positive")
    return number


def _percentile(samples: Sequence[float], percentile: int) -> float:
    ordered = sorted(samples)
    if percentile == 50:
        return float(statistics.median(ordered))
    rank = max(1, math.ceil((percentile / 100.0) * len(ordered)))
    return float(ordered[rank - 1])


def recompute_statistics(elapsed: Any, location: str) -> dict[str, Any]:
    """Validate an ``elapsed_ns`` object and recompute its four statistics."""

    elapsed_object = _require_object(elapsed, location)
    if elapsed_object.get("unit") != "ns":
        raise AbbaSummaryInputError(f"{location}.unit must be 'ns'")
    samples_raw = elapsed_object.get("samples")
    if not isinstance(samples_raw, list) or not samples_raw:
        raise AbbaSummaryInputError(f"{location}.samples must be a non-empty list")
    samples = [
        _finite_number(value, f"{location}.samples[{index}]", positive=True)
        for index, value in enumerate(samples_raw)
    ]
    try:
        mean = float(statistics.fmean(samples))
    except (OverflowError, ValueError) as error:
        raise AbbaSummaryInputError(f"{location}.mean cannot be represented finitely") from error
    if not math.isfinite(mean):
        raise AbbaSummaryInputError(f"{location}.mean must be finite")
    computed = {
        "p50": _percentile(samples, 50),
        "mean": mean,
        "p95": _percentile(samples, 95),
        "p99": _percentile(samples, 99),
    }
    for name, expected in computed.items():
        if name not in elapsed_object:
            raise AbbaSummaryInputError(f"{location}.{name} is required")
        reported = _finite_number(elapsed_object[name], f"{location}.{name}", positive=True)
        # Integer nanosecond percentiles are commonly serialized as floats;
        # permit only insignificant serialization/rounding noise.  A changed
        # reported statistic still fails closed.
        tolerance = max(0.5, abs(expected) * 1e-12)
        if abs(reported - expected) > tolerance:
            raise AbbaSummaryInputError(
                f"{location}.{name}={reported} disagrees with samples ({expected})"
            )
    if not computed["p50"] <= computed["p95"] <= computed["p99"]:
        raise AbbaSummaryInputError(f"{location} percentiles are not non-decreasing")
    reported_values = {name: float(elapsed_object[name]) for name in STATISTICS}
    if not reported_values["p50"] <= reported_values["p95"] <= reported_values["p99"]:
        raise AbbaSummaryInputError(
            f"{location} reported percentiles are not non-decreasing"
        )
    return {
        "sample_count": len(samples),
        "p50": computed["p50"],
        "mean": computed["mean"],
        "p95": computed["p95"],
        "p99": computed["p99"],
    }


def _result_key(result: Any, location: str) -> tuple[str, str, dict[str, Any]]:
    row = _require_object(result, location)
    case = row.get("case")
    if not isinstance(case, str) or not case:
        raise AbbaSummaryInputError(f"{location}.case must be a non-empty string")
    corpus = _require_object(row.get("corpus"), f"{location}.corpus")
    if not corpus:
        raise AbbaSummaryInputError(f"{location}.corpus must not be empty")
    corpus_identity = _canonical_json(corpus, f"{location}.corpus")
    return case, corpus_identity, row


def _index_results(report: dict[str, Any], label: str) -> dict[tuple[str, str], dict[str, Any]]:
    results = report.get("results")
    if not isinstance(results, list) or not results:
        raise AbbaSummaryInputError(f"{label}.results must be a non-empty list")
    indexed: dict[tuple[str, str], dict[str, Any]] = {}
    for index, result in enumerate(results):
        case, corpus_identity, row = _result_key(result, f"{label}.results[{index}]")
        key = (case, corpus_identity)
        if key in indexed:
            raise AbbaSummaryInputError(
                f"{label}.results contains duplicate case/corpus identity for {case!r}"
            )
        indexed[key] = row
    return indexed


def _validate_report(
    report: Any, label: str
) -> tuple[
    int,
    dict[str, Any],
    dict[str, Any],
    dict[str, Any],
    dict[tuple[str, str], dict[str, Any]],
]:
    root = _require_object(report, label)
    schema_version = root.get("schema_version")
    if isinstance(schema_version, bool) or not isinstance(schema_version, int):
        raise AbbaSummaryInputError(f"{label}.schema_version must be an integer")
    if schema_version != HARNESS_SCHEMA_VERSION:
        raise AbbaSummaryInputError(
            f"{label}.schema_version must be {HARNESS_SCHEMA_VERSION!r}"
        )
    tool = _require_object(root.get("tool"), f"{label}.tool")
    if not tool:
        raise AbbaSummaryInputError(f"{label}.tool must not be empty")
    environment = _require_object(root.get("environment"), f"{label}.environment")
    if not environment:
        raise AbbaSummaryInputError(f"{label}.environment must not be empty")
    if "git_revision" in environment and environment["git_revision"] is not None and (
        not isinstance(environment["git_revision"], str) or not environment["git_revision"]
    ):
        raise AbbaSummaryInputError(
            f"{label}.environment.git_revision must be null or a non-empty string"
        )
    if (
        "git_worktree_dirty" in environment
        and environment["git_worktree_dirty"] is not None
        and not isinstance(environment["git_worktree_dirty"], bool)
    ):
        raise AbbaSummaryInputError(
            f"{label}.environment.git_worktree_dirty must be null or a boolean"
        )
    configuration = _require_object(root.get("configuration"), f"{label}.configuration")
    if not configuration:
        raise AbbaSummaryInputError(f"{label}.configuration must not be empty")
    _canonical_json(tool, f"{label}.tool")
    _canonical_json(environment, f"{label}.environment")
    _canonical_json(configuration, f"{label}.configuration")
    return schema_version, tool, environment, configuration, _index_results(root, label)


def _stable_environment(environment: dict[str, Any]) -> dict[str, Any]:
    """Return environment facts expected to remain fixed across ABBA legs."""

    return {
        key: value for key, value in environment.items() if key not in ENVIRONMENT_VARIANTS
    }


def _validate_configuration_rows(
    configuration: dict[str, Any],
    indexed: Mapping[tuple[str, str], dict[str, Any]],
    label: str,
) -> None:
    """Check optional harness cardinality/selector declarations when present."""

    cases = configuration.get("cases")
    if cases is not None:
        if (
            not isinstance(cases, list)
            or not cases
            or any(not isinstance(case, str) or not case for case in cases)
            or len(cases) != len(set(cases))
        ):
            raise AbbaSummaryInputError(f"{label}.configuration.cases must be unique strings")
        actual_cases = {case for case, _ in indexed}
        if actual_cases != set(cases):
            raise AbbaSummaryInputError(
                f"{label}.configuration.cases does not match result cases"
            )
    shapes = configuration.get("corpus_shapes")
    if shapes is not None:
        if (
            not isinstance(shapes, list)
            or any(not isinstance(shape, str) or not shape for shape in shapes)
            or len(shapes) != len(set(shapes))
        ):
            raise AbbaSummaryInputError(
                f"{label}.configuration.corpus_shapes must be unique strings"
            )
    samples_per_case = configuration.get("samples_per_case")
    if samples_per_case is not None:
        if (
            isinstance(samples_per_case, bool)
            or not isinstance(samples_per_case, int)
            or samples_per_case < 1
        ):
            raise AbbaSummaryInputError(
                f"{label}.configuration.samples_per_case must be a positive integer"
            )
        for key, row in indexed.items():
            case = key[0]
            elapsed = _require_object(row.get("elapsed_ns"), f"{label}.{case}.elapsed_ns")
            samples = elapsed.get("samples")
            if not isinstance(samples, list) or len(samples) != samples_per_case:
                raise AbbaSummaryInputError(
                    f"{label}.{case}.elapsed_ns.samples does not match samples_per_case"
                )


def _validate_drift_ceilings(value: Mapping[str, Any] | None) -> dict[str, float]:
    if value is None:
        return dict(DEFAULT_DRIFT_CEILINGS)
    if not isinstance(value, Mapping):
        raise AbbaSummaryInputError("drift ceilings must be an object")
    if set(value) != set(STATISTICS):
        raise AbbaSummaryInputError(
            f"drift ceilings must contain exactly {list(STATISTICS)!r}"
        )
    ceilings = {
        name: _finite_number(value[name], f"drift ceilings.{name}")
        for name in STATISTICS
    }
    if any(ceiling < 0 for ceiling in ceilings.values()):
        raise AbbaSummaryInputError("drift ceilings must be non-negative")
    return ceilings


def _identity_value(row: dict[str, Any], field: str, location: str) -> tuple[bool, str]:
    present = field in row
    value = row[field] if present else None
    return present, _canonical_json(value, f"{location}.{field}")


def _coerce_reports(
    a1: Mapping[str, Any] | Sequence[Mapping[str, Any]],
    b1: Mapping[str, Any] | None,
    b2: Mapping[str, Any] | None,
    a2: Mapping[str, Any] | None,
) -> tuple[Mapping[str, Any], Mapping[str, Any], Mapping[str, Any], Mapping[str, Any]]:
    if b1 is None and b2 is None and a2 is None:
        if isinstance(a1, Mapping):
            for labels in (LEG_ORDER, tuple(label.upper() for label in LEG_ORDER)):
                if all(label in a1 for label in labels):
                    return tuple(a1[label] for label in labels)  # type: ignore[return-value]
            raise AbbaSummaryInputError(
                f"reports mapping must contain {list(LEG_ORDER)!r}"
            )
        if isinstance(a1, Sequence) and not isinstance(a1, (str, bytes)) and len(a1) == 4:
            return tuple(a1)  # type: ignore[return-value]
        raise AbbaSummaryInputError("reports must be four reports in A1,B1,B2,A2 order")
    if b1 is None or b2 is None or a2 is None:
        raise AbbaSummaryInputError("all four ABBA reports are required")
    return a1, b1, b2, a2


def _parse_selectors(values: Iterable[str] | None) -> set[str] | None:
    if values is None:
        return None
    if isinstance(values, str):
        values = (values,)
    selectors: set[str] = set()
    for value in values:
        for selector in value.split(","):
            selector = selector.strip()
            if selector:
                selectors.add(selector)
    if not selectors:
        raise AbbaSummaryInputError("selectors must contain at least one non-empty value")
    return selectors


def _delta_percent(control: float, candidate: float) -> float:
    if control <= 0:
        raise AbbaSummaryInputError("elapsed statistics must have a positive control value")
    delta = (control - candidate) / control * 100.0
    if not math.isfinite(delta):
        raise AbbaSummaryInputError("candidate reduction is not finite")
    return delta


def _drift_percent(first: float, second: float) -> float:
    if first <= 0:
        raise AbbaSummaryInputError("elapsed statistics must have positive drift baselines")
    drift = (second - first) / first * 100.0
    if not math.isfinite(drift):
        raise AbbaSummaryInputError("same-implementation drift is not finite")
    return drift


def _result_summary(
    rows: Mapping[str, dict[str, Any]],
    *,
    case: str,
    corpus_identity: str,
    drift_ceilings: Mapping[str, float],
) -> dict[str, Any]:
    source_identities = {
        label: _identity_value(rows[label], "source", f"{label}.{case}") for label in LEG_ORDER
    }
    sink_identities = {
        label: _identity_value(rows[label], "sink", f"{label}.{case}") for label in LEG_ORDER
    }
    for field, identities in (("source", source_identities), ("sink", sink_identities)):
        expected = identities["a1"]
        if any(identity != expected for identity in identities.values()):
            raise AbbaSummaryInputError(
                f"{case}[{corpus_identity}] {field} identity differs between ABBA legs"
            )

    elapsed: dict[str, dict[str, Any]] = {}
    for label in LEG_ORDER:
        elapsed[label] = recompute_statistics(
            rows[label].get("elapsed_ns"), f"{label}.{case}.elapsed_ns"
        )
    sample_counts = {elapsed[label]["sample_count"] for label in LEG_ORDER}
    if len(sample_counts) != 1:
        raise AbbaSummaryInputError(
            f"{case}[{corpus_identity}] ABBA legs have different sample counts"
        )

    candidate_reduction = {
        "a1_to_b1": {
            name: _delta_percent(elapsed["a1"][name], elapsed["b1"][name])
            for name in STATISTICS
        },
        "a2_to_b2": {
            name: _delta_percent(elapsed["a2"][name], elapsed["b2"][name])
            for name in STATISTICS
        },
    }
    adverse_both: list[str] = []
    drift = {
        "control": {
            name: _drift_percent(elapsed["a1"][name], elapsed["a2"][name])
            for name in STATISTICS
        },
        "candidate": {
            name: _drift_percent(elapsed["b1"][name], elapsed["b2"][name])
            for name in STATISTICS
        },
    }
    drift_within_ceiling = {
        implementation: {
            name: abs(drift[implementation][name]) <= drift_ceilings[name]
            for name in STATISTICS
        }
        for implementation in ("control", "candidate")
    }
    accepted: list[str] = []
    rejected: dict[str, str] = {}
    for name in STATISTICS:
        reasons: list[str] = []
        first_reduction = candidate_reduction["a1_to_b1"][name]
        second_reduction = candidate_reduction["a2_to_b2"][name]
        if first_reduction < 0 and second_reduction < 0:
            adverse_both.append(name)
            reasons.append("candidate is not lower in both paired directions")
        elif first_reduction <= 0 or second_reduction <= 0:
            if (first_reduction < 0) != (second_reduction < 0):
                reasons.append("paired directions disagree")
            else:
                reasons.append("candidate is not lower in both paired directions")
        for implementation in ("control", "candidate"):
            if not drift_within_ceiling[implementation][name]:
                reasons.append(
                    f"{implementation} drift {drift[implementation][name]:+.6f}% "
                    f"exceeds {drift_ceilings[name]:g}% ceiling"
                )
        if reasons:
            rejected[name] = "; ".join(reasons)
        else:
            accepted.append(name)

    a1_row = rows["a1"]
    source_present, source_identity = source_identities["a1"]
    sink_present, sink_identity = sink_identities["a1"]
    corpus = json.loads(corpus_identity)
    return {
        "case": case,
        "shape": corpus.get("shape"),
        "corpus": corpus,
        "source": a1_row.get("source") if source_present else None,
        "sink": a1_row.get("sink") if sink_present else None,
        "identity": {
            "corpus": corpus_identity,
            "source_present": source_present,
            "source_canonical_json": source_identity,
            "sink_present": sink_present,
            "sink_canonical_json": sink_identity,
        },
        "elapsed_ns": {
            "sample_count": next(iter(sample_counts)),
            "legs_ns": elapsed,
            "candidate_reduction_percent": candidate_reduction,
            "same_implementation_drift_percent": drift,
            "drift_ceiling_percent": dict(drift_ceilings),
            "drift_within_ceiling": drift_within_ceiling,
            "adverse_both_statistics": adverse_both,
            "accepted_statistics": accepted,
            "rejected_statistics": rejected,
        },
    }


def summarize_reports(
    a1: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
    b1: Mapping[str, Any] | None = None,
    b2: Mapping[str, Any] | None = None,
    a2: Mapping[str, Any] | None = None,
    *,
    drift_ceilings: Mapping[str, Any] | None = None,
    cases: Iterable[str] | None = None,
    shapes: Iterable[str] | None = None,
    reports: Mapping[str, Any] | Sequence[Mapping[str, Any]] | None = None,
    ceilings: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Validate four reports and return a deterministic machine-readable summary.

    ``a1`` may be a four-item sequence, a mapping keyed by ``a1``, ``b1``,
    ``b2`` and ``a2``, or the first report when the remaining three reports are
    passed as positional arguments.
    """

    if reports is not None:
        if a1 is not None or b1 is not None or b2 is not None or a2 is not None:
            raise AbbaSummaryInputError("reports cannot be combined with positional ABBA reports")
        a1 = reports
    if a1 is None:
        raise AbbaSummaryInputError("four ABBA reports are required")
    if drift_ceilings is not None and ceilings is not None:
        raise AbbaSummaryInputError("use drift_ceilings or ceilings, not both")
    ceiling_values = _validate_drift_ceilings(
        drift_ceilings if drift_ceilings is not None else ceilings
    )
    report_values = _coerce_reports(a1, b1, b2, a2)
    validated = {
        label: _validate_report(report, label)
        for label, report in zip(LEG_ORDER, report_values)
    }
    schema_versions = {item[0] for item in validated.values()}
    if len(schema_versions) != 1:
        raise AbbaSummaryInputError("harness schema_version differs between ABBA legs")
    tool_identities = {
        _canonical_json(item[1], f"{label}.tool") for label, item in validated.items()
    }
    if len(tool_identities) != 1:
        raise AbbaSummaryInputError("harness tool identity differs between ABBA legs")
    environment_identities = {
        label: _canonical_json(item[2], f"{label}.environment")
        for label, item in validated.items()
    }
    stable_environment_identities = {
        label: _canonical_json(_stable_environment(item[2]), f"{label}.environment")
        for label, item in validated.items()
    }
    if len(set(stable_environment_identities.values())) != 1:
        raise AbbaSummaryInputError(
            "stable environment identity differs between ABBA legs"
        )
    configurations = {
        _canonical_json(item[3], f"{label}.configuration") for label, item in validated.items()
    }
    if len(configurations) != 1:
        raise AbbaSummaryInputError("harness configuration differs between ABBA legs")

    for label, (_, _, _, configuration, indexed) in validated.items():
        _validate_configuration_rows(configuration, indexed, label)

    result_sets = {frozenset(item[4]) for item in validated.values()}
    if len(result_sets) != 1:
        raise AbbaSummaryInputError("case/corpus result identities differ between ABBA legs")
    selected_cases = _parse_selectors(cases)
    selected_shapes = _parse_selectors(shapes)
    first_index = validated["a1"][4]
    selected_keys = []
    for case, corpus_identity in sorted(first_index):
        corpus = json.loads(corpus_identity)
        if selected_cases is not None and case not in selected_cases:
            continue
        shape = corpus.get("shape")
        if selected_shapes is not None and not isinstance(shape, str):
            raise AbbaSummaryInputError(
                f"{case}[{corpus_identity}] cannot be selected by shape without a string shape"
            )
        if selected_shapes is not None and shape not in selected_shapes:
            continue
        selected_keys.append((case, corpus_identity))
    if not selected_keys:
        raise AbbaSummaryInputError("selectors did not match any case/corpus result")

    # Validate every row before applying selectors.  A selector controls what
    # is emitted; it must not hide an unsafe mismatch in another report row.
    all_summaries: dict[tuple[str, str], dict[str, Any]] = {}
    for case, corpus_identity in sorted(first_index):
        rows = {label: validated[label][4][(case, corpus_identity)] for label in LEG_ORDER}
        all_summaries[(case, corpus_identity)] = _result_summary(
            rows,
            case=case,
            corpus_identity=corpus_identity,
            drift_ceilings=ceiling_values,
        )

    results = [all_summaries[key] for key in selected_keys]

    tool = json.loads(next(iter(tool_identities)))
    configuration = json.loads(next(iter(configurations)))
    environments = {
        label: json.loads(environment_identities[label]) for label in LEG_ORDER
    }
    stable_environment = json.loads(stable_environment_identities["a1"])
    return {
        "schema_version": SCHEMA_VERSION,
        "tool": {"name": TOOL_NAME, "version": TOOL_VERSION},
        "protocol": {
            "order": ["a1_control", "b1_candidate", "b2_candidate", "a2_control"],
            "statistics": list(STATISTICS),
            "drift_ceiling_percent": ceiling_values,
            "percentiles": "p50 = statistics.median; p95/p99 = nearest-rank",
        },
        "harness_identity": {
            "schema_version": next(iter(schema_versions)),
            "tool": tool,
            "configuration": configuration,
        },
        "environment": {
            "stable": stable_environment,
            "legs": environments,
        },
        "results": results,
        "verification": {
            "result_count": len(results),
            "tool_identity_verified": True,
            "configuration_identity_verified": True,
            "environment_stable_identity_verified": True,
            "environment_legs_recorded": True,
            "case_corpus_identity_verified": True,
            "source_identity_verified": True,
            "sink_identity_verified": True,
            "statistics_recomputed_from_samples": True,
        },
    }


build_summary = summarize_reports
summarize_abba = summarize_reports


def load_report(path: Path) -> dict[str, Any]:
    try:
        with path.open(encoding="utf-8") as handle:
            value = json.load(handle)
    except (OSError, json.JSONDecodeError) as error:
        raise AbbaSummaryInputError(f"cannot read {path}: {error}") from error
    return _require_object(value, str(path))


def _parse_ceiling_argument(value: str) -> dict[str, float]:
    try:
        parsed = json.loads(value)
    except json.JSONDecodeError:
        parsed = {}
        for item in value.split(","):
            name, separator, number = item.partition("=")
            if not separator:
                raise AbbaSummaryInputError(
                    "--drift-ceilings must be JSON or comma-separated name=value pairs"
                )
            parsed[name.strip()] = float(number)
    if not isinstance(parsed, dict):
        raise AbbaSummaryInputError("--drift-ceilings must describe an object")
    return _validate_drift_ceilings(parsed)


def _write_json(path: Path, value: dict[str, Any]) -> None:
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open("x", encoding="utf-8") as handle:
            json.dump(value, handle, indent=2, sort_keys=True, allow_nan=False)
            handle.write("\n")
    except FileExistsError as error:
        raise AbbaSummaryInputError(f"output already exists: {path}") from error
    except OSError as error:
        raise AbbaSummaryInputError(f"cannot write {path}: {error}") from error


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "reports",
        nargs="*",
        type=Path,
        metavar="REPORT",
        help="four reports in A1,B1,B2,A2 order",
    )
    for label in LEG_ORDER:
        aliases = {
            "a1": ("--a1", "--control-a", "--before-a"),
            "b1": ("--b1", "--candidate-a", "--after-a"),
            "b2": ("--b2", "--candidate-b", "--after-b"),
            "a2": ("--a2", "--control-b", "--before-b"),
        }[label]
        parser.add_argument(*aliases, type=Path, help=f"{label.upper()} harness report")
    parser.add_argument("--json-out", "--output", type=Path)
    parser.add_argument(
        "--drift-ceilings",
        help="JSON object or comma-separated p50=5,mean=5,p95=10,p99=15",
    )
    parser.add_argument("--case", dest="cases", action="append")
    parser.add_argument("--shape", dest="shapes", action="append")
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    try:
        named = [getattr(arguments, label) for label in LEG_ORDER]
        if arguments.reports and any(path is not None for path in named):
            raise AbbaSummaryInputError("use positional reports or --a1/--b1/--b2/--a2, not both")
        if arguments.reports:
            if len(arguments.reports) != 4:
                raise AbbaSummaryInputError("exactly four positional reports are required")
            paths = arguments.reports
        else:
            if any(path is None for path in named):
                raise AbbaSummaryInputError("all four --a1/--b1/--b2/--a2 reports are required")
            paths = named
        reports = [load_report(path) for path in paths]
        ceilings = (
            _parse_ceiling_argument(arguments.drift_ceilings)
            if arguments.drift_ceilings is not None
            else None
        )
        summary = summarize_reports(
            reports,
            drift_ceilings=ceilings,
            cases=arguments.cases,
            shapes=arguments.shapes,
        )
        if arguments.json_out is not None:
            _write_json(arguments.json_out, summary)
        json.dump(summary, sys.stdout, indent=2, sort_keys=True, allow_nan=False)
        sys.stdout.write("\n")
        return 0
    except (OSError, AbbaSummaryInputError, ValueError) as error:
        print(f"{TOOL_NAME}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
