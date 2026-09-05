#!/usr/bin/env python3
"""Run fail-closed mutation probes against one validated 0423 report.

This checker deliberately works on an in-memory copy of the report.  It never
rewrites the captured report or catalog; each probe must be rejected by the
same single-report verifier used for acceptance.
"""

from __future__ import annotations

import argparse
import copy
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any, Callable


HERE = Path(__file__).resolve().parent


def load_verifier() -> Any:
    path = HERE / "pinned" / "verify-report.py"
    spec = importlib.util.spec_from_file_location("change_0423_verify", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load verifier module {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def summary_for(report: dict[str, Any], verifier: Any, selector: str) -> dict[str, Any]:
    result = verifier.object_value(report["results"][0], "report.results[0]")
    source = verifier.object_value(result["source"], "report.results[0].source")
    key = "pptx_cross_copy" if selector in verifier.OWNED_SELECTORS else "pptx_source_backed_cross_copy_lifecycle"
    return verifier.object_value(source[key], f"report.results[0].source.{key}")


def validate_copy(
    report: dict[str, Any],
    catalog: dict[str, Any],
    *,
    verifier: Any,
    validators: tuple[Any, Any, Any],
    selector: str,
    lane: str,
    samples: int,
    warmups: int,
) -> None:
    perf_abba_summary, perf_compare, corpus_binding = validators
    verifier.validate_report(
        report,
        catalog,
        selector=selector,
        lane=lane,
        samples=samples,
        warmups=warmups,
        perf_abba_summary=perf_abba_summary,
        perf_compare=perf_compare,
        corpus_binding=corpus_binding,
    )


def mutation_probes(
    report: dict[str, Any],
    *,
    verifier: Any,
    catalog: dict[str, Any],
    validators: tuple[Any, Any, Any],
    selector: str,
    lane: str,
    samples: int,
    warmups: int,
) -> list[dict[str, Any]]:
    def trial(name: str, mutate: Callable[[dict[str, Any]], None]) -> dict[str, Any]:
        candidate = copy.deepcopy(report)
        mutate(candidate)
        try:
            validate_copy(
                candidate,
                catalog,
                verifier=verifier,
                validators=validators,
                selector=selector,
                lane=lane,
                samples=samples,
                warmups=warmups,
            )
        except verifier.VerificationError as error:
            return {"name": name, "rejected": True, "error": str(error)}
        except Exception as error:  # pragma: no cover - a verifier defect
            return {
                "name": name,
                "rejected": True,
                "error": f"unexpected verifier exception: {type(error).__name__}: {error}",
                "unexpected_exception": True,
            }
        return {"name": name, "rejected": False}

    probes: list[dict[str, Any]] = []

    def disable_gate(candidate: dict[str, Any]) -> None:
        summary = summary_for(candidate, verifier, selector)
        gate = next(iter(summary["gates"]))
        summary["gates"][gate] = False

    probes.append(trial("gate_false", disable_gate))

    def alter_output_digest(candidate: dict[str, Any]) -> None:
        summary = summary_for(candidate, verifier, selector)
        summary["output_sha256"][0] = "0" * 64

    probes.append(trial("output_digest_mutated", alter_output_digest))

    def alter_elapsed(candidate: dict[str, Any]) -> None:
        elapsed = candidate["results"][0]["elapsed_ns"]
        if samples > 1:
            elapsed["sample_order"][1] = elapsed["sample_order"][0]
        else:
            elapsed["samples"][0] = 0

    probes.append(trial("elapsed_or_sample_order_mutated", alter_elapsed))

    def exceed_sink_write(candidate: dict[str, Any]) -> None:
        candidate["results"][0]["sink"]["largest_write"] = verifier.SINK_MAX_WRITE + 1

    probes.append(trial("sink_write_ceiling_exceeded", exceed_sink_write))

    def alter_copy_plan(candidate: dict[str, Any]) -> None:
        summary = summary_for(candidate, verifier, selector)
        field = "added_opc_part_count" if selector in verifier.SOURCE_SELECTORS else "planned_part_count"
        summary[field] += 1

    probes.append(trial("copy_plan_count_mutated", alter_copy_plan))

    def add_unknown_configuration(candidate: dict[str, Any]) -> None:
        candidate["configuration"]["unregistered_option"] = True

    def change_shape_configuration(candidate: dict[str, Any]) -> None:
        candidate["configuration"]["corpus_shapes"] = ["plain"]

    def change_source_identity(candidate: dict[str, Any]) -> None:
        summary_for(candidate, verifier, selector)["source_archive_sha256"] = "0" * 64

    probes.append(trial("unknown_configuration_field", add_unknown_configuration))
    probes.append(trial("configuration_shape_mutated", change_shape_configuration))
    probes.append(trial("fixed_source_identity_mutated", change_source_identity))

    if selector in verifier.SOURCE_SELECTORS:
        def alter_source_output_length(candidate: dict[str, Any]) -> None:
            summary = summary_for(candidate, verifier, selector)
            summary["source_expected_output_bytes"] += 1

        probes.append(trial("source_output_length_mutated", alter_source_output_length))

        def exceed_lifecycle_with_open(candidate: dict[str, Any]) -> None:
            summary = summary_for(candidate, verifier, selector)
            summary["open_ns"][0] = summary["lifecycle_ns"][0] + 1

        def remove_source_reads(candidate: dict[str, Any]) -> None:
            summary_for(candidate, verifier, selector)["source_read_calls"][0] = 0

        def change_owned_sink_ceiling(candidate: dict[str, Any]) -> None:
            summary_for(candidate, verifier, selector)["matched_owned_output_ceiling"] += 1

        probes.append(trial("open_phase_exceeds_lifecycle", exceed_lifecycle_with_open))
        probes.append(trial("source_reads_omitted", remove_source_reads))
        probes.append(trial("common_owned_sink_ceiling_mutated", change_owned_sink_ceiling))

    if lane == "allocator":
        def region_below_endpoint(candidate: dict[str, Any]) -> None:
            allocation = candidate["results"][0]["operation_metrics"]["allocation"]
            before = allocation["live_bytes_before"]["values"][0]
            after = allocation["live_bytes_after"]["values"][0]
            endpoint = max(before, after)
            allocation["region_peak_live_bytes"]["values"][0] = endpoint - 1 if endpoint else -1

        def region_above_peak(candidate: dict[str, Any]) -> None:
            allocation = candidate["results"][0]["operation_metrics"]["allocation"]
            peak = allocation["peak_live_bytes_after"]["values"][0]
            allocation["region_peak_live_bytes"]["values"][0] = peak + 1

        probes.append(trial("allocator_region_peak_below_endpoint", region_below_endpoint))
        probes.append(trial("allocator_region_peak_above_lifetime_peak", region_above_peak))

    return probes


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo-root", type=Path)
    parser.add_argument("--report", type=Path, required=True)
    parser.add_argument("--catalog", type=Path, required=True)
    parser.add_argument("--selector", choices=sorted({
        "pptx_cross_copy_plain_lifecycle",
        "pptx_cross_copy_media_rich_lifecycle",
        "pptx_source_backed_cross_copy_plain_lifecycle",
        "pptx_source_backed_cross_copy_media_rich_lifecycle",
    }), required=True)
    parser.add_argument("--lane", choices=("normal", "allocator"), default="normal")
    parser.add_argument("--contract", choices=("formal", "functional"), default="formal")
    parser.add_argument("--samples", type=int)
    parser.add_argument("--warmups", type=int)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        verifier = load_verifier()
        repo_root = verifier.find_repo_root(args.repo_root)
        validators = verifier.import_validators(repo_root)
        expected_samples, expected_warmups = verifier.contract_samples(args.lane, args.contract)
        samples = expected_samples if args.samples is None else args.samples
        warmups = expected_warmups if args.warmups is None else args.warmups
        if (samples, warmups) != (expected_samples, expected_warmups):
            verifier.fail(
                "contract",
                f"{args.lane}/{args.contract} requires {expected_samples} samples and {expected_warmups} warmups",
            )
        report, report_sha256 = verifier.load_json(args.report.expanduser(), "report")
        catalog, catalog_sha256 = verifier.load_json(args.catalog.expanduser(), "catalog")
        validate_copy(
            report,
            catalog,
            verifier=verifier,
            validators=validators,
            selector=args.selector,
            lane=args.lane,
            samples=samples,
            warmups=warmups,
        )
        checks = mutation_probes(
            report,
            verifier=verifier,
            catalog=catalog,
            validators=validators,
            selector=args.selector,
            lane=args.lane,
            samples=samples,
            warmups=warmups,
        )
        if any(not check["rejected"] or check.get("unexpected_exception") for check in checks):
            raise verifier.VerificationError("one or more mutation probes were not cleanly rejected")
        output = {
            "change": verifier.EXPECTED_CHANGE,
            "selector": args.selector,
            "lane": args.lane,
            "samples": samples,
            "warmups": warmups,
            "report_sha256": report_sha256,
            "catalog_sha256": catalog_sha256,
            "validated_original": True,
            "checks": checks,
            "claim_authorized": False,
        }
        if args.output is not None:
            output_path = args.output.expanduser()
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(json.dumps(output, sort_keys=True, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(output, sort_keys=True, indent=2))
        return 0
    except Exception as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
