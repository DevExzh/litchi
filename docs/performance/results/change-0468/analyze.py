#!/usr/bin/env python3
"""Analyze the retained 0468 XLSX frame-pointer sample export.

The parser and the generic whole-process attribution rules live in the 0466
helper.  This wrapper loads that helper from its explicit repository path and
adds the two views needed for the current optimization: an inclusive ranking
inside the exact ``Edit::commit`` context and small, source-motivated views of
the parser, worksheet rewriter, and XML-attribute symbols that remain there.

All values are descriptive ``cycles:u`` sample-period weights.  They are not
phase timers, operation counters, or proof that an absent symbol did no work.
The wrapper never invokes ``perf`` and accepts either an uncompressed or a
gzip-compressed script export.
"""

from __future__ import annotations

import argparse
from collections import Counter
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any, Callable, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
LEGACY_PATH = ROOT.parent / "change-0466" / "analyze.py"
LEGACY_MODULE_NAME = "litchi_change0466_cpu_analyze"
REPO_ROOT = next(
    (candidate for candidate in ROOT.parents if (candidate / "tools").is_dir()),
    ROOT,
)


def _load_legacy() -> Any:
    """Load the shared 0466 implementation without copying its parser."""

    if not LEGACY_PATH.is_file():
        raise RuntimeError(f"missing shared analyzer: {LEGACY_PATH}")
    spec = importlib.util.spec_from_file_location(LEGACY_MODULE_NAME, LEGACY_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load shared analyzer: {LEGACY_PATH}")
    module = importlib.util.module_from_spec(spec)
    # The dataclass declaration in the shared module consults sys.modules while
    # it is being executed.  Register the explicit module before exec_module.
    sys.modules[LEGACY_MODULE_NAME] = module
    spec.loader.exec_module(module)
    return module


legacy = _load_legacy()

# Deliberately re-export the shared parser primitives so tests and small
# replay scripts use exactly the same implementation as the 0466 evidence.
Sample = legacy.Sample
MARKERS = legacy.MARKERS
DEFAULT_TOP = legacy.DEFAULT_TOP
EXPECTED_EVENT = legacy.EXPECTED_EVENT
parse_samples = legacy.parse_samples
rank_symbols = legacy.rank_symbols
weighted_metric = legacy.weighted_metric
marker_matches = legacy.marker_matches
file_binding = legacy.file_binding


SCHEMA = "litchi-0468-xlsx-cpu-attribution-v1"
COMMIT_MARKER = MARKERS["edit_commit"][0]

# These are intentionally narrow and inspectable source-symbol families.  A
# symbol can belong to more than one family (for example quick-xml attribute
# iteration is part of worksheet parsing), so family rows are explicitly
# overlapping.  The patterns are substring anchors because perf's demangled
# output adds generic arguments and closure suffixes to many Rust methods.
MATERIAL_SYMBOL_PATTERNS: dict[str, tuple[str, ...]] = {
    "parser": (
        "litchi_xlsx::raw::worksheet::model::Parser",
        "litchi_xlsx::raw::worksheet::parse",
        "litchi_xlsx::raw::worksheet::validation",
        "quick_xml::reader::",
        "quick_xml::name::NamespaceResolver",
    ),
    "rewriter": (
        "litchi_xlsx::raw::worksheet::edit::",
        "litchi_xlsx::raw::compact::",
        "litchi_xlsx::raw::web::read",
    ),
    "attribute": (
        "litchi_ooxml_common::xml::unqualified_attribute_value",
        "quick_xml::events::attributes::Attribute",
        "quick_xml::events::attributes::Attributes",
        "quick_xml::events::attributes::IterState",
        "quick_xml::escape::",
    ),
}


def _matches_any(symbol: str, patterns: Sequence[str]) -> bool:
    return any(pattern in symbol for pattern in patterns)


def _has_commit_marker(sample: Sample) -> bool:
    return any(marker_matches(symbol, COMMIT_MARKER) for symbol in sample.symbols)


def _rank_material_symbols(
    samples: Iterable[Sample],
    patterns: Sequence[str],
    subset_denominator: int,
    whole_denominator: int,
    limit: int,
) -> list[dict[str, Any]]:
    """Rank only matched symbols, de-duplicating repeated frames per stack."""

    periods: Counter[str] = Counter()
    blocks: Counter[str] = Counter()
    for sample in samples:
        for symbol in {
            symbol for symbol in sample.symbols if _matches_any(symbol, patterns)
        }:
            periods[symbol] += sample.period
            blocks[symbol] += 1
    rows: list[dict[str, Any]] = []
    for symbol, period in sorted(
        periods.items(), key=lambda item: (-item[1], item[0])
    )[:limit]:
        rows.append(
            {
                "symbol": symbol,
                **weighted_metric(
                    period, blocks[symbol], subset_denominator, whole_denominator
                ),
            }
        )
    return rows


def _material_context(
    samples: Sequence[Sample],
    patterns: Sequence[str],
    denominator: int,
    whole: int,
    top: int,
) -> dict[str, Any]:
    matched = [
        sample
        for sample in samples
        if any(_matches_any(symbol, patterns) for symbol in sample.symbols)
    ]
    period = sum(sample.period for sample in matched)
    return {
        "matched_sample_blocks": len(matched),
        "matched_weighted_event_period": period,
        "share_of_subset_percent": period / denominator * 100.0
        if denominator
        else None,
        "share_of_whole_process_percent": period / whole * 100.0 if whole else None,
        "inclusive_period_weighted_ranking": _rank_material_symbols(
            matched, patterns, denominator, whole, top
        ),
    }


def material_symbol_views(
    samples: Sequence[Sample], *, top: int = DEFAULT_TOP
) -> dict[str, Any]:
    """Return overlapping material symbol families for whole and commit scopes."""

    whole = sum(sample.period for sample in samples)
    commits = [sample for sample in samples if _has_commit_marker(sample)]
    commit_period = sum(sample.period for sample in commits)
    categories: dict[str, Any] = {}
    for name, patterns in MATERIAL_SYMBOL_PATTERNS.items():
        categories[name] = {
            "patterns": list(patterns),
            "whole_process": _material_context(
                samples, patterns, whole, whole, top
            ),
            "exact_commit_context": _material_context(
                commits, patterns, commit_period, whole, top
            ),
        }
    return {
        "scope": (
            "inclusive whole-process and exact Edit::commit ancestor contexts; "
            "family rows overlap and are descriptive sample-period weights"
        ),
        "commit_marker": COMMIT_MARKER,
        "whole_process": {
            "weighted_event_period": whole,
            "raw_stack_blocks": len(samples),
        },
        "exact_commit_context": {
            "weighted_event_period": commit_period,
            "raw_stack_blocks": len(commits),
        },
        "families_overlap": True,
        "categories": categories,
    }


def _counter_rows(
    periods: Counter[str],
    blocks: Counter[str],
    subset_denominator: int,
    whole_denominator: int,
    limit: int,
) -> list[dict[str, Any]]:
    return [
        {
            "symbol": symbol,
            **weighted_metric(
                period, blocks[symbol], subset_denominator, whole_denominator
            ),
        }
        for symbol, period in sorted(
            periods.items(), key=lambda item: (-item[1], item[0])
        )[:limit]
    ]


def _sample_group(
    samples: Sequence[Sample], matcher: Callable[[str], bool]
) -> list[Sample]:
    return [sample for sample in samples if any(matcher(symbol) for symbol in sample.symbols)]


def _sample_period_bound(
    name: str,
    group: Sequence[Sample],
    commit_period: int,
    whole_period: int,
) -> dict[str, Any]:
    period = sum(sample.period for sample in group)
    commit_fraction = period / commit_period if commit_period else None
    whole_fraction = period / whole_period if whole_period else None

    def speedup(fraction: float | None) -> float | None:
        if fraction is None or fraction >= 1.0:
            return None
        return 1.0 / (1.0 - fraction)

    return {
        "name": name,
        "matched_sample_blocks": len(group),
        "weighted_event_period": period,
        "share_of_exact_commit_context_percent": commit_fraction * 100.0
        if commit_fraction is not None
        else None,
        "share_of_whole_process_percent": whole_fraction * 100.0
        if whole_fraction is not None
        else None,
        "sample_period_amdahl_proxy": {
            "formula": "1 / (1 - attributed_fraction)",
            "exact_commit_context_speedup_if_all_attributed_period_disappeared": speedup(
                commit_fraction
            ),
            "whole_process_speedup_if_all_attributed_period_disappeared": speedup(
                whole_fraction
            ),
            "status": "descriptive_upper_bound_not_measured_latency",
        },
    }


def _repository_binding(path: Path) -> dict[str, Any]:
    # Preserve source context across later production edits and portable replay.
    relative = path.relative_to(REPO_ROOT).as_posix()
    snapshot = ROOT / "source-context" / (path.name + ".txt")
    binding = file_binding(snapshot, ROOT)
    source_manifest = json.loads((ROOT / "sources.json").read_text())
    if binding["sha256"] != source_manifest[relative]:
        raise ValueError(f"retained source context differs from build inventory: {relative}")
    binding["repository_path"] = relative
    binding["scope"] = "retained-source-snapshot"
    return binding


def _web_read_context(
    samples: Sequence[Sample], *, top: int = DEFAULT_TOP
) -> dict[str, Any]:
    """Explain the retained callchain context around ``raw::web::read``."""

    target = "litchi_xlsx::raw::web::read"
    target_samples: list[Sample] = []
    target_periods = Counter()
    target_blocks = Counter()
    caller_periods = Counter()
    caller_blocks = Counter()
    child_periods = Counter()
    child_blocks = Counter()
    positions = Counter()
    occurrence_count = 0
    for sample in samples:
        indexes = [index for index, symbol in enumerate(sample.symbols) if target in symbol]
        if not indexes:
            continue
        target_samples.append(sample)
        target_periods[target] += sample.period
        target_blocks[target] += 1
        occurrence_count += len(indexes)
        for index in indexes:
            positions[index] += sample.period
            if index + 1 < len(sample.symbols):
                caller = sample.symbols[index + 1]
                caller_periods[caller] += sample.period
                caller_blocks[caller] += 1
            if index:
                child = sample.symbols[index - 1]
                child_periods[child] += sample.period
                child_blocks[child] += 1
            else:
                child_periods["<leaf/raw::web::read>"] += sample.period
                child_blocks["<leaf/raw::web::read>"] += 1

    whole_period = sum(sample.period for sample in samples)
    commit_samples = _sample_group(samples, lambda symbol: marker_matches(symbol, COMMIT_MARKER))
    commit_period = sum(sample.period for sample in commit_samples)
    target_commit_samples = _sample_group(
        commit_samples, lambda symbol: target in symbol
    )
    target_period = sum(target_periods.values())
    return {
        "symbol": target,
        "callchain_order": "leaf to caller/root",
        "matched_sample_blocks": len(target_samples),
        "target_occurrences": occurrence_count,
        "target_occurrences_per_sample": {
            "one": len(target_samples) if occurrence_count == len(target_samples) else None,
            "histogram": {
                str(count): sum(
                    1
                    for sample in target_samples
                    if sum(target in symbol for symbol in sample.symbols) == count
                )
                for count in sorted(
                    {
                        sum(target in symbol for symbol in sample.symbols)
                        for sample in target_samples
                    }
                )
            },
        },
        "weighted_event_period": target_period,
        "share_of_exact_commit_context_percent": target_period / commit_period * 100.0
        if commit_period
        else None,
        "share_of_whole_process_percent": target_period / whole_period * 100.0
        if whole_period
        else None,
        "target_samples_inside_exact_commit_context": len(target_commit_samples),
        "target_period_inside_exact_commit_context": sum(
            sample.period for sample in target_commit_samples
        ),
        "retained_frame_position_period": [
            {
                "leaf_distance": index,
                "weighted_event_period": period,
                "share_of_target_percent": period / target_period * 100.0
                if target_period
                else None,
            }
            for index, period in sorted(positions.items())
        ],
        "nearest_retained_caller": _counter_rows(
            caller_periods, caller_blocks, target_period, whole_period, top
        ),
        "nearest_retained_child": _counter_rows(
            child_periods, child_blocks, target_period, whole_period, top
        ),
        "source_context": [
            {
                "file": _repository_binding(
                    REPO_ROOT / "crates/litchi-xlsx/src/workbook/edit/semantic/transaction.rs"
                ),
                "line": 1677,
                "role": "commit validates the complete post-edit worksheet through raw::web::read",
            },
            {
                "file": _repository_binding(REPO_ROOT / "crates/litchi-xlsx/src/raw/web.rs"),
                "line": 33,
                "role": "raw::web::read constructs an NsReader over worksheet bytes",
            },
            {
                "file": _repository_binding(REPO_ROOT / "crates/litchi-xlsx/src/raw/web.rs"),
                "line": 47,
                "role": "the web-binding validator reads the worksheet event stream to EOF",
            },
        ],
        "interpretation": (
            "The nearest retained caller is the Edit::commit validation path. "
            "The nearest child frames are XML event/name/decoding work inside the "
            "whole-worksheet web-binding scan; raw::web::read is not a filesystem-I/O claim."
        ),
    }


def profile_context(
    script: Path,
    reports: Sequence[Path] = (),
    *,
    top: int = DEFAULT_TOP,
) -> dict[str, Any]:
    """Build a compact, reproducible explanation artifact from retained samples."""

    samples, stats = parse_samples(script)
    if not samples:
        raise ValueError("no positive cycles:u sample blocks parsed")
    whole_period = sum(sample.period for sample in samples)
    commit_samples = _sample_group(
        samples, lambda symbol: marker_matches(symbol, COMMIT_MARKER)
    )
    commit_period = sum(sample.period for sample in commit_samples)
    compact_samples = _sample_group(
        commit_samples,
        lambda symbol: "litchi_xlsx::raw::compact::changed" in symbol,
    )
    compact_period = sum(sample.period for sample in compact_samples)
    selectors: dict[str, Callable[[str], bool]] = {
        "worksheet_parser_union": lambda symbol: marker_matches(
            symbol, "<litchi_xlsx::raw::worksheet::model::Parser>::parse"
        ),
        "snapshot_scan": lambda symbol: "litchi_xlsx::raw::worksheet::edit::codec::snapshot::scan::scan_with_limit"
        in symbol,
        "compact_changed": lambda symbol: "litchi_xlsx::raw::compact::changed" in symbol,
        "raw_web_read": lambda symbol: "litchi_xlsx::raw::web::read" in symbol,
        "attribute_helper": lambda symbol: "litchi_ooxml_common::xml::unqualified_attribute_value"
        in symbol,
    }
    amdahl_rows = [
        _sample_period_bound(
            name,
            _sample_group(commit_samples, matcher),
            commit_period,
            whole_period,
        )
        for name, matcher in selectors.items()
    ]
    return {
        "schema": "litchi-0468-profile-context-v1",
        "purpose": (
            "descriptive callchain context and Amdahl-style sample-period bounds "
            "for the 0468 XLSX CPU profile"
        ),
        "inputs": {
            "perf_script": file_binding(script, ROOT),
            "normal_reports": [file_binding(path, ROOT) for path in reports],
            "analyzer": file_binding(Path(__file__).resolve(), ROOT),
            "shared_analyzer": _source_binding(LEGACY_PATH),
        },
        "sample_parser": {
            "accepted_sample_blocks": len(samples),
            "accepted_weighted_event_period": whole_period,
            "parser_diagnostics": stats,
        },
        "scope_denominators": {
            "whole_process": {
                "raw_stack_blocks": len(samples),
                "weighted_event_period": whole_period,
            },
            "exact_commit_context": {
                "raw_stack_blocks": len(commit_samples),
                "weighted_event_period": commit_period,
                "marker": COMMIT_MARKER,
            },
        },
        "raw_web_read_context": _web_read_context(samples, top=top),
        "compact_changed_subtree": {
            "marker": "litchi_xlsx::raw::compact::changed",
            "scope": "samples retaining the compact::changed ancestor inside exact Edit::commit",
            "raw_stack_blocks": len(compact_samples),
            "weighted_event_period": compact_period,
            "share_of_exact_commit_context_percent": compact_period / commit_period * 100.0
            if commit_period
            else None,
            "share_of_whole_process_percent": compact_period / whole_period * 100.0
            if whole_period
            else None,
            "inclusive_period_weighted_ranking": rank_symbols(
                compact_samples,
                lambda _sample: True,
                compact_period,
                whole_period,
                top,
                leaf=False,
            ),
            "leaf_period_weighted_ranking": rank_symbols(
                compact_samples,
                lambda _sample: True,
                compact_period,
                whole_period,
                top,
                leaf=True,
            ),
            "interpretive_buckets": {
                "scope": "overlapping symbol-family views inside the compact subtree; not additive",
                "xml_read_and_namespace": [
                    "quick_xml::reader::",
                    "quick_xml::name::NamespaceResolver",
                ],
                "event_ownership_and_copy": [
                    "quick_xml::events::Event>::into_owned",
                    "__memmove",
                    "malloc",
                    "__rustc::__rust_alloc",
                    "RawVec",
                    "Vec<",
                ],
                "attribute_iteration_and_duplicate_checks": [
                    "quick_xml::events::attributes::Attributes",
                    "quick_xml::events::attributes::IterState",
                ],
                "caution": (
                    "These are inclusive CPU symbol families. Their sampled periods "
                    "may overlap, and CPU stacks do not establish allocation counts."
                ),
            },
        },
        "sample_period_amdahl_proxy": {
            "rows_are_overlapping": True,
            "cannot_be_summed": True,
            "interpretation": (
                "Each row is a separate conditional upper-bound proxy if all of "
                "that row's inclusive sampled period vanished while the rest stayed "
                "fixed. It is not a measured elapsed-latency prediction."
            ),
            "rows": amdahl_rows,
        },
        "limitations": [
            "cycles:u periods are sampled event weights, not elapsed phase durations",
            "the whole-process denominator includes setup, warmups, verification, and teardown when sampled",
            "inclusive rows overlap across symbols and source-motivated families",
            "a missing or inlined frame does not establish zero work",
            "this artifact does not compare 0466 and 0468 numerically",
        ],
    }


def _source_binding(path: Path) -> dict[str, Any]:
    """Bind the shared analyzer with a stable repository-relative label."""

    binding = file_binding(path, ROOT)
    # The shared implementation is deliberately outside the 0468 evidence
    # directory.  Keep its repository-relative identity explicit rather than
    # presenting an absolute checkout path in the JSON report.
    if path == LEGACY_PATH:
        binding["path"] = "../change-0466/analyze.py"
        binding["scope"] = "adjacent-retained-repository-source"
    return binding


def analyze_script(
    script: Path,
    reports: Sequence[Path] = (),
    *,
    top: int = DEFAULT_TOP,
) -> dict[str, Any]:
    """Analyze one current 0468 sample export and contextual reports."""

    if top <= 0:
        raise ValueError("top must be positive")
    # Use the old helper for parsing, coverage, whole-process ranking, exact
    # contexts, and report binding.  It accepts both plain text and .gz paths.
    summary = legacy.analyze_script(
        script, reports, top=top, markers=MARKERS, bundle_root=ROOT
    )
    samples, _stats = parse_samples(script)
    summary["schema"] = SCHEMA
    summary["purpose"] = (
        "descriptive whole-process XLSX CPU attribution for 0468; no phase-time "
        "or operation-counter claim"
    )
    summary["inputs"]["parser"] = file_binding(Path(__file__).resolve(), ROOT)
    summary["inputs"]["shared_analyzer"] = _source_binding(LEGACY_PATH)
    summary["commit_context"] = {
        "marker": COMMIT_MARKER,
        "inclusive_period_weighted_ranking": rank_symbols(
            [sample for sample in samples if _has_commit_marker(sample)],
            lambda _sample: True,
            sum(sample.period for sample in samples if _has_commit_marker(sample)),
            sum(sample.period for sample in samples),
            top,
            leaf=False,
        ),
        "scope": (
            "exact Edit::commit ancestor context; inclusive rows de-duplicate "
            "repeated ancestors within each stack and overlap across rows"
        ),
    }
    summary["material_symbols"] = material_symbol_views(samples, top=top)
    summary["limitations"] = list(summary.get("limitations", ())) + [
        "Material parser, rewriter, and attribute families overlap by design; their rows are not additive.",
        "The exact commit context is an inclusive callchain selection and can include untimed setup, verification, or expected-output work when those frames are retained.",
        "The 0466 profile is not a numeric baseline here: source/build revisions and sampled inclusive denominators differ.",
    ]
    return summary


# A descriptive alias makes the intended current-profile entry point explicit
# to small replay scripts without creating a second implementation.
analyze_current = analyze_script


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--script", type=Path, required=True)
    parser.add_argument("--report", type=Path, action="append", default=[])
    parser.add_argument("--top", type=int, default=DEFAULT_TOP)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument(
        "--additional-output",
        type=Path,
        default=None,
        help="also write compact callchain-context/Amdahl proxy JSON",
    )
    args = parser.parse_args(argv)
    if args.top <= 0:
        raise SystemExit("--top must be positive")
    summary = analyze_script(args.script, args.report, top=args.top)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    if args.additional_output is not None:
        context = profile_context(args.script, args.report, top=args.top)
        args.additional_output.parent.mkdir(parents=True, exist_ok=True)
        args.additional_output.write_text(
            json.dumps(context, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
            encoding="utf-8",
        )
    print(
        json.dumps(
            {
                "output": str(args.output.resolve()),
                "whole_process_blocks": summary["whole_process"]["raw_stack_blocks"],
                "whole_process_period": summary["whole_process"]["weighted_event_period"],
                "commit_context_blocks": summary["material_symbols"][
                    "exact_commit_context"
                ]["raw_stack_blocks"],
                "commit_context_period": summary["material_symbols"][
                    "exact_commit_context"
                ]["weighted_event_period"],
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
