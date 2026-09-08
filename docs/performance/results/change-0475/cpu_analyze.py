#!/usr/bin/env python3
"""Analyze portable CPU profiles for the 0475 PPTX streaming experiment.

This command consumes only a ``perf script`` text export and optional normal
reports.  It reports cycles:u period weights and callchain attribution.  It
does not turn sampled periods into elapsed phase time, and independent owner
families are deliberately allowed to overlap.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any, Mapping, Sequence

import cpu_parser


SCHEMA = "litchi-0475-pptx-streaming-cpu-attribution-v1"
DEFAULT_CONTEXT_MARKERS: dict[str, tuple[str, ...]] = {
    "writer_under_run": (
        "litchi_perf_baseline::pptx_streaming_create::run",
    ),
    "materialized_preflight": (
        "litchi_perf_baseline::pptx_streaming_create::build_corpus",
        "litchi_perf_baseline::pptx_streaming_create::inspect_materialized_archive",
    ),
}

# These are attribution lenses rather than disjoint scopes.  A deflate write
# normally sits below a ZIP-entry writer and may therefore occur in both rows.
# Keeping the literals in the output makes the analysis auditable when a
# profile build changes its demangled generic spelling.
DEFAULT_OWNER_FAMILIES: dict[str, tuple[str, ...]] = {
    "deflate_initialization": (
        "flate2::deflate::write::DeflateEncoder",
        "soapberry_zip::writer::OwnedCompressor",
        "flate2::mem::Compress",
        "zlib_rs::stable::Deflate::new",
        "zlib_rs::deflate::init",
        "zlib_rs::allocate::zalloc_rust",
    ),
    "deflate_compression": (
        "flate2::zio::Writer",
        "flate2::mem::Compress",
        "zlib_rs::stable::Deflate::compress",
        "zlib_rs::stable::Deflate::compress_uninit",
        "zlib_rs::deflate::deflate",
        "zlib_rs::deflate::algorithm::",
    ),
    "opc_part_name_set_name_validation": (
        "litchi_opc::phys_pkg::PartNameSet",
        "litchi_opc::phys_pkg::PreparedPartName",
        "litchi_opc::phys_pkg::validate_part_name",
        "litchi_opc::package::OpcPackage>::validate_new_part_name",
        "litchi_opc::packuri::PackURI",
        "litchi_pptx::writer::streaming::part_uri",
        "litchi_opc::packuri::validate_percent_encoding",
    ),
    "zip_entry_directory": (
        "soapberry_zip::office::StreamingArchiveWriter",
        "soapberry_zip::writer::ZipArchiveWriter",
        "soapberry_zip::writer::ZipOwnedEntryWriter",
        "soapberry_zip::writer::PreparedSizedMember",
        "central_bytes",
        "finalize_central",
    ),
    "hashing_discard_sink": (
        "litchi_perf_baseline::HashingDiscardSink",
    ),
}


class AnalysisError(ValueError):
    """Raised when a profile cannot support the declared evidence."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def _marker_map(value: Any, label: str) -> dict[str, tuple[str, ...]]:
    if not isinstance(value, dict):
        fail(f"{label} must be an object")
    result: dict[str, tuple[str, ...]] = {}
    for name, literals in value.items():
        if not isinstance(name, str) or not name:
            fail(f"{label} contains an invalid family name")
        if isinstance(literals, str) and literals:
            result[name] = (literals,)
        elif isinstance(literals, list) and all(
            isinstance(item, str) and item for item in literals
        ):
            result[name] = tuple(literals)
        else:
            fail(f"{label}.{name} must contain non-empty string literals")
    if not result:
        fail(f"{label} contains no marker literals")
    return result


def load_anchors(path: Path | None) -> tuple[
    Mapping[str, Sequence[str]], Mapping[str, Sequence[str]]
]:
    """Load optional source-derived marker overrides without importing code."""

    if path is None:
        return DEFAULT_CONTEXT_MARKERS, DEFAULT_OWNER_FAMILIES
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AnalysisError(f"cannot read source anchors {path}: {error}") from error
    if not isinstance(value, dict):
        fail("source anchors must be an object")
    contexts_value = value.get("contexts", value.get("markers"))
    families_value = value.get("owner_families", value.get("families"))
    contexts = (
        _marker_map(contexts_value, "source anchors contexts")
        if contexts_value is not None
        else DEFAULT_CONTEXT_MARKERS
    )
    families = (
        _marker_map(families_value, "source anchors owner_families")
        if families_value is not None
        else DEFAULT_OWNER_FAMILIES
    )
    return contexts, families


def _load_report(path: Path, bundle_root: Path) -> dict[str, Any]:
    return cpu_parser._report_summary(path, bundle_root)


def analyze_script(
    script: Path,
    reports: Sequence[Path] = (),
    *,
    profile_label: str | None = None,
    top: int = cpu_parser.DEFAULT_TOP,
    context_markers: Mapping[str, Sequence[str]] = DEFAULT_CONTEXT_MARKERS,
    owner_families: Mapping[str, Sequence[str]] = DEFAULT_OWNER_FAMILIES,
    bundle_root: Path | None = None,
    anchors_path: Path | None = None,
) -> dict[str, Any]:
    if top <= 0:
        raise AnalysisError("top must be positive")
    if bundle_root is None:
        bundle_root = Path(__file__).resolve().parent
    samples, stats = cpu_parser.parse_samples(script)
    if not samples:
        raise AnalysisError("no positive cycles:u sample blocks parsed")
    whole = sum(sample.period for sample in samples)
    contexts = cpu_parser.context_summaries(samples, top, context_markers)
    # Lost samples are process-level perf metadata with no callchain and
    # therefore cannot be assigned to one context.  Repeat that limitation
    # beside each context row rather than hiding it in the global parser block.
    for context in contexts["categories"].values():
        context["lost_sample_count"] = stats.get("lost_sample_count", 0)
        context["lost_period_is_unavailable"] = True
    families = cpu_parser.owner_family_summaries(samples, top, owner_families)
    parser_path = Path(__file__).resolve().with_name("cpu_parser.py")
    analyzer_path = Path(__file__).resolve()
    return {
        "schema": SCHEMA,
        "purpose": (
            "descriptive whole-process cycles:u callchain attribution for public "
            "PPTX streaming slide creation"
        ),
        "profile": {
            "label": profile_label,
            "event": cpu_parser.EXPECTED_EVENT,
            "callchain_order": "perf frames are leaf to caller/root",
            "sampled_period_is_elapsed_time": False,
        },
        "timing_semantics": {
            "weight": "cycles:u perf sample period",
            "phase_latency": False,
            "operation_counter": False,
            "sample_period_to_elapsed_conversion": False,
            "writer_context": (
                "nearest disjoint context marker identifies samples retaining the "
                "public run frame; it does not measure a phase duration"
            ),
            "owner_families": (
                "independent inclusive lenses; one sample may match several families "
                "and family periods must not be summed"
            ),
            "normal_report_elapsed": (
                "optional harness elapsed vectors are workload context only; they "
                "are not combined with sample periods"
            ),
        },
        "inputs": {
            "bundle_root": ".",
            "perf_script": cpu_parser.file_binding(script, bundle_root),
            "normal_reports": [
                _load_report(path, bundle_root) for path in reports
            ],
            "parser": cpu_parser.file_binding(parser_path, bundle_root),
            "analyzer": cpu_parser.file_binding(analyzer_path, bundle_root),
            "source_anchors": (
                cpu_parser.file_binding(anchors_path, bundle_root)
                if anchors_path is not None
                else None
            ),
        },
        "sample_parser": cpu_parser.coverage(stats, samples),
        "whole_process": {
            "weighted_event_period": whole,
            "raw_stack_blocks": len(samples),
            "leaf_period_weighted_ranking": cpu_parser.rank_symbols(
                samples, lambda _sample: True, whole, whole, top, leaf=True
            ),
            "inclusive_period_weighted_ranking": cpu_parser.rank_symbols(
                samples, lambda _sample: True, whole, whole, top, leaf=False
            ),
        },
        "writer_context": {
            "name": "writer_under_run",
            "marker": list(context_markers.get("writer_under_run", ())),
            "scope": contexts["categories"].get(
                "writer_under_run", {"scope": cpu_parser.weighted_metric(0, 0, 0, whole)}
            ),
            "exact_context": (
                "accepted samples whose callchain contains the exact "
                "litchi_perf_baseline::pptx_streaming_create::run marker"
            ),
        },
        "materialized_preflight": {
            "name": "materialized_preflight",
            "marker": list(context_markers.get("materialized_preflight", ())),
            "scope": contexts["categories"].get(
                "materialized_preflight",
                {"scope": cpu_parser.weighted_metric(0, 0, 0, whole)},
            ),
            "exact_context": (
                "accepted samples retaining materialized corpus construction or "
                "its semantic/physical preflight markers"
            ),
        },
        "contexts": contexts,
        "owner_families": families,
        "limitations": [
            "Sample periods are weighted cycles:u event periods, not wall time or elapsed phase duration.",
            "The whole-process profile includes setup, materialized preflight, timed writer work, verification, and teardown whenever sampled.",
            "A missing marker means no matching frame was retained; it does not prove zero work.",
            "Unknown, malformed, empty, unparsed, truncated, and lost coverage remain explicit denominators.",
            "Writer and preflight context rows use a nearest-marker partition; owner-family rows are inclusive and may overlap.",
            "This profile does not claim total-RSS causality, physical I/O, or broad PPTX creation performance.",
        ],
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--script", type=Path, required=True)
    parser.add_argument("--report", type=Path, action="append", default=[])
    parser.add_argument("--source-anchors", type=Path, default=None)
    parser.add_argument("--profile-label", default=None)
    parser.add_argument("--top", type=int, default=cpu_parser.DEFAULT_TOP)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args(argv)
    contexts, families = load_anchors(args.source_anchors)
    bundle_root = Path(__file__).resolve().parent
    summary = analyze_script(
        args.script,
        args.report,
        profile_label=args.profile_label,
        top=args.top,
        context_markers=contexts,
        owner_families=families,
        bundle_root=bundle_root,
        anchors_path=args.source_anchors,
    )
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(summary, ensure_ascii=False, indent=2, sort_keys=True) + "\n",
        encoding="utf-8",
    )
    print(
        json.dumps(
            {
                "output": str(args.output.resolve()),
                "whole_process_blocks": summary["whole_process"]["raw_stack_blocks"],
                "whole_process_period": summary["whole_process"]["weighted_event_period"],
                "writer_blocks": summary["writer_context"]["scope"]["scope"]["raw_stack_blocks"],
                "writer_period": summary["writer_context"]["scope"]["scope"]["weighted_event_period"],
                "preflight_blocks": summary["materialized_preflight"]["scope"]["scope"]["raw_stack_blocks"],
                "preflight_period": summary["materialized_preflight"]["scope"]["scope"]["weighted_event_period"],
                "lost_samples": summary["sample_parser"]["lost_sample_count"],
                "truncated_period": summary["sample_parser"]["truncated_weighted_event_period"],
            },
            sort_keys=True,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
