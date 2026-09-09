#!/usr/bin/env python3
"""Derive the 0483 validation and fuzz plans from retained receipts.

This helper only reads completed custody records.  It never builds, runs a
harness, runs a test, or invents a label.  The accepted attempt, final gate
labels, pilot specifications, developmental classifications, and fuzz receipt
labels are supplied by the coordinator; every supplied item must already be a
regular file with a complete receipt before either plan is written.

Pilot syntax is::

    LABEL:INSTRUMENTATION:ROUTE:COUNT:SAMPLES:WARMUPS:REPORT_PATH

The retained gate argv may use the direct form above or the exact formal
``/usr/bin/time -v -o PILOT.resource /usr/bin/taskset -c 2`` wrapper around
that harness command.

Fuzz receipt syntax is::

    LABEL=KIND:PATH

where KIND is one of ``prepared``, ``build`` or ``smoke``.  A developmental
classification uses ``LABEL=CLASSIFICATION:REASON``.  All paths in plans are
relative to this bundle, while the recorded command retains the actual
absolute executable/report paths from its gate receipt.
"""

from __future__ import annotations

import argparse
import json
from pathlib import Path
import sys
from typing import Any, NoReturn

import analyze
import verify
from common import ROOT, environment, write


class PlanError(ValueError):
    """A fail-closed plan input or retained receipt error."""


VALIDATION_FIELDS = {
    "schema", "label", "attempt", "argv", "cwd", "environment", "driver_sha256",
    "common_sha256", "started_utc", "source_before", "exit_code", "finished_utc",
    "source_after", "source_unchanged", "artifacts",
}
STARTED_VALIDATION_FIELDS = VALIDATION_FIELDS - {
    "exit_code", "finished_utc", "source_after", "source_unchanged", "artifacts",
}
FUZZ_KINDS = {"prepared", "build", "smoke"}


def fail(message: str) -> NoReturn:
    raise PlanError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path, label: str) -> Any:
    try:
        return verify.read_json(path, label)
    except verify.VerificationError as error:
        fail(str(error))


def relative_file(value: str, label: str) -> tuple[str, Path]:
    try:
        relative = verify.safe_relative(value, label)
        path = verify.bundle_file(relative, label)
    except verify.VerificationError as error:
        fail(str(error))
    return relative, path


def metadata_reference(value: str, label: str) -> dict[str, Any]:
    relative, path = relative_file(value, label)
    try:
        metadata = verify.metadata(path, label)
    except verify.VerificationError as error:
        fail(str(error))
    return {"path": relative, **metadata}


def output_relative(value: str, label: str) -> str:
    """Validate an output path without requiring the file to exist yet."""

    try:
        relative = verify.safe_relative(value, label)
    except verify.VerificationError as error:
        fail(str(error))
    current = ROOT
    for part in Path(relative).parts:
        current /= part
        require(not current.is_symlink(), f"{label}: symlink component is forbidden")
    return relative


def parse_positive(text: str, label: str) -> int:
    require(text.isdecimal(), f"{label} must be a positive decimal integer")
    value = int(text)
    require(value > 0, f"{label} must be positive")
    return value


def parse_attempt(value: str) -> str:
    require(bool(verify.LABEL.fullmatch(value)), "--attempt must be a path-safe token")
    return value


def parse_pilot(value: str, attempt: str) -> dict[str, Any]:
    parts = value.split(":", 6)
    require(len(parts) == 7, "--pilot must be LABEL:INSTRUMENTATION:ROUTE:COUNT:SAMPLES:WARMUPS:REPORT_PATH")
    label, instrumentation, route, count_text, samples_text, warmups_text, report_path = parts
    require(bool(verify.LABEL.fullmatch(label)), f"pilot label is not path-safe: {label!r}")
    require(instrumentation in analyze.INSTRUMENTATIONS, f"pilot {label}: instrumentation is invalid")
    require(route in analyze.ROUTES, f"pilot {label}: route is invalid")
    count = parse_positive(count_text, f"pilot {label}.count")
    require(count in analyze.COUNTS, f"pilot {label}: count is not in the frozen matrix")
    samples = parse_positive(samples_text, f"pilot {label}.samples")
    warmups = parse_positive(warmups_text, f"pilot {label}.warmups")
    relative_file(report_path, f"pilot {label}.report_path")
    return {
        "label": label,
        "path": report_path,
        "samples": samples,
        "warmups": warmups,
        "spec": {
            "label": label,
            "route": route,
            "route_name": analyze.ROUTE_NAMES[route],
            "instrumentation": instrumentation,
            "count": count,
            "attempt": attempt,
        },
    }


def parse_classification(value: str) -> tuple[str, dict[str, str]]:
    try:
        label, body = value.split("=", 1)
        classification, reason = body.split(":", 1)
    except ValueError:
        fail("--classify must be LABEL=developmental|historical:REASON")
    require(bool(verify.LABEL.fullmatch(label)), f"classification label is not path-safe: {label!r}")
    require(classification in {"developmental", "historical"}, f"classification {label}: kind is invalid")
    require(bool(reason.strip()), f"classification {label}: reason is empty")
    return label, {"classification": classification, "reason": reason}


def load_classifications(options: argparse.Namespace) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    if options.classification_file:
        path = Path(options.classification_file)
        require(path.is_file() and not path.is_symlink(), f"classification file is missing: {path}")
        value = read_json(path, "classification file")
        require(isinstance(value, dict), "classification file must contain an object")
        for label, item in value.items():
            require(isinstance(label, str), "classification file contains a malformed label")
            require(
                isinstance(item, dict) and set(item) == {"classification", "reason"},
                f"classification file entry {label} is malformed",
            )
            result[label] = {"classification": item["classification"], "reason": item["reason"]}
    for raw in options.classify:
        label, item = parse_classification(raw)
        require(label not in result, f"classification is duplicated: {label}")
        result[label] = item
    for label, item in result.items():
        require(bool(verify.LABEL.fullmatch(label)), f"classification label is not path-safe: {label!r}")
        require(item["classification"] in {"developmental", "historical"}, f"classification {label}: kind is invalid")
        require(isinstance(item["reason"], str) and item["reason"].strip(), f"classification {label}: reason is empty")
    return result


def parse_fuzz_receipt(value: str, attempt: str) -> tuple[str, str, str]:
    try:
        label, body = value.split("=", 1)
        kind, path = body.split(":", 1)
    except ValueError:
        fail("--fuzz-receipt must be LABEL=prepared|build|smoke:PATH")
    require(bool(verify.LABEL.fullmatch(label)), f"fuzz receipt label is not path-safe: {label!r}")
    require(kind in FUZZ_KINDS, f"fuzz receipt {label}: kind is invalid")
    relative, _ = relative_file(path, f"fuzz receipt {label}")
    prefix = f"fuzz/{attempt}/"
    require(relative.startswith(prefix), f"fuzz receipt {label}: path must be under {prefix}")
    return label, kind, relative


def discover_validation() -> tuple[dict[str, Path], dict[str, Path]]:
    directory = ROOT / "validation"
    require(directory.is_dir() and not directory.is_symlink(), "validation directory is missing")
    terminal = {
        path.stem: path
        for path in sorted(directory.glob("*.json"))
        if path.is_file() and not path.is_symlink() and not path.name.endswith(".started.json")
    }
    started = {
        path.name.removesuffix(".started.json"): path
        for path in sorted(directory.glob("*.started.json"))
        if path.is_file() and not path.is_symlink()
    }
    require(terminal, "validation contains no completed receipts")
    require(set(terminal) == set(started), "validation has orphaned started or terminal receipts")
    return terminal, started


def check_validation_receipt(label: str, path: Path, started_path: Path) -> dict[str, Any]:
    receipt = read_json(path, f"validation/{label}.json")
    started = read_json(started_path, f"validation/{label}.started.json")
    require(isinstance(receipt, dict) and set(receipt) == VALIDATION_FIELDS, f"validation/{label}: receipt fields differ")
    require(isinstance(started, dict) and set(started) == STARTED_VALIDATION_FIELDS, f"validation/{label}: started fields differ")
    require(receipt.get("schema") == "docx-tail-append-gate-v1", f"validation/{label}: schema differs")
    require(receipt.get("label") == label and started.get("label") == label, f"validation/{label}: label differs")
    require(isinstance(receipt.get("attempt"), (str, type(None))), f"validation/{label}: attempt is malformed")
    require(isinstance(receipt.get("argv"), list) and receipt["argv"] and all(isinstance(item, str) for item in receipt["argv"]), f"validation/{label}: argv is malformed")
    require(isinstance(receipt.get("cwd"), str) and Path(receipt["cwd"]).is_absolute(), f"validation/{label}: cwd is malformed")
    require(type(receipt.get("exit_code")) is int, f"validation/{label}: exit code is malformed")
    require(isinstance(receipt.get("source_unchanged"), bool), f"validation/{label}: source flag is malformed")
    require(receipt.get("environment") == environment(), f"validation/{label}: environment differs")
    for key in STARTED_VALIDATION_FIELDS:
        require(receipt.get(key) == started.get(key), f"validation/{label}: started field {key} changed")
    try:
        verify.interval(receipt.get("started_utc"), receipt.get("finished_utc"), f"validation/{label}")
        verify.check_source_reference(receipt.get("source_before"), f"validation/{label}.source_before")
        verify.check_source_reference(receipt.get("source_after"), f"validation/{label}.source_after")
        verify.check_digest(receipt.get("driver_sha256"), f"validation/{label}.driver_sha256")
        verify.check_digest(receipt.get("common_sha256"), f"validation/{label}.common_sha256")
    except verify.VerificationError as error:
        fail(str(error))
    require(verify._historical_helper("gate.py", receipt["driver_sha256"]) is not None, f"validation/{label}: gate helper custody is missing")
    require(verify._historical_helper("common.py", receipt["common_sha256"]) is not None, f"validation/{label}: common helper custody is missing")
    artifacts = receipt.get("artifacts")
    expected_artifacts = {f"{label}.stdout", f"{label}.stderr"}
    require(isinstance(artifacts, dict) and set(artifacts) == expected_artifacts, f"validation/{label}: artifacts differ")
    for name, expected in artifacts.items():
        artifact_path = ROOT / "validation" / name
        require(artifact_path.is_file() and not artifact_path.is_symlink(), f"validation/{label}: artifact is missing: {name}")
        try:
            verify.check_metadata(artifact_path, expected, f"validation/{label}.{name}")
        except verify.VerificationError as error:
            fail(str(error))
    return receipt


def load_accepted_builds(attempt: str) -> tuple[dict[str, Any], str, set[str]]:
    """Use the strict build custody checker before accepting any gate label."""

    try:
        binaries = verify.check_builds({"attempt": attempt, "environment": environment()})
    except verify.VerificationError as error:
        fail(str(error))
    require(binaries.get("attempt") == attempt, "accepted binaries attempt differs")
    source = binaries.get("source_manifest_sha256")
    try:
        source = verify.check_digest(source, "accepted source manifest")
    except verify.VerificationError as error:
        fail(str(error))
    gate_labels: set[str] = set()
    for instrumentation in ("normal", "allocator"):
        build_path = verify.bundle_file(binaries["binaries"][instrumentation]["build_path"], f"accepted {instrumentation} build")
        build = read_json(build_path, f"accepted {instrumentation} build")
        require(build.get("gate_label"), f"accepted {instrumentation} build gate label is missing")
        gate_labels.add(build["gate_label"])
    return binaries, source, gate_labels


def make_validation_plan(
    attempt: str,
    required_labels: list[str],
    pilots: list[dict[str, Any]],
    classifications: dict[str, dict[str, str]],
    terminal_paths: dict[str, Path],
    started_paths: dict[str, Path],
    accepted_source: str,
    binaries: dict[str, Any],
    build_gate_labels: set[str],
) -> dict[str, Any]:
    pilot_labels = [pilot["label"] for pilot in pilots]
    final_labels = required_labels + pilot_labels
    require(required_labels, "at least one --required-label is required")
    require(pilot_labels, "at least one --pilot is required")
    require(len(set(final_labels)) == len(final_labels), "final validation labels are duplicated")
    require(build_gate_labels <= set(required_labels), f"accepted build gate labels are not final required labels: {sorted(build_gate_labels - set(required_labels))}")
    require(set(final_labels) <= set(terminal_paths), f"final validation receipts are missing: {sorted(set(final_labels) - set(terminal_paths))}")
    require(set(classifications) <= set(terminal_paths) - set(final_labels), "classifications contain a final or unknown label")
    other_labels = set(terminal_paths) - set(final_labels)
    require(other_labels <= set(classifications), f"unclassified completed validation receipts: {sorted(other_labels - set(classifications))}")

    receipts = {
        label: check_validation_receipt(label, terminal_paths[label], started_paths[label])
        for label in terminal_paths
    }
    for label in final_labels:
        receipt = receipts[label]
        require(receipt["attempt"] == attempt, f"validation/{label}: accepted attempt differs")
        require(receipt["exit_code"] == 0 and receipt["source_unchanged"] is True, f"validation/{label}: accepted gate did not pass source-stably")
        require(receipt["source_before"] == receipt["source_after"], f"validation/{label}: accepted source changed")
        require(receipt["source_after"]["sha256"] == accepted_source, f"validation/{label}: accepted source binding differs")

    binary_by_instrumentation = binaries["binaries"]
    pilot_reports: dict[str, dict[str, Any]] = {}
    for pilot in pilots:
        label = pilot["label"]
        receipt = receipts[label]
        spec = pilot["spec"]
        report_path = verify.bundle_file(pilot["path"], f"validation/{label}.pilot_report")
        report_metadata = verify.metadata(report_path, f"validation/{label}.pilot_report")
        expected_binary = binary_by_instrumentation[spec["instrumentation"]]
        try:
            verify.check_pilot_argv(receipt["argv"], receipt, pilot, expected_binary, report_path, f"validation/{label}")
            report = verify.read_json(report_path, f"validation/{label}.pilot_report")
            analyze.validate_report(
                report,
                spec,
                f"validation/{label}.pilot_report",
                expected_samples=pilot["samples"],
                expected_warmups=pilot["warmups"],
            )
        except (verify.VerificationError, analyze.AnalysisError) as error:
            fail(str(error))
        pilot_reports[label] = {
            "path": pilot["path"],
            **report_metadata,
            "spec": spec,
            "samples": pilot["samples"],
            "warmups": pilot["warmups"],
        }

    developmental: dict[str, dict[str, Any]] = {}
    for label in sorted(other_labels):
        item = classifications[label]
        receipt = receipts[label]
        developmental[label] = {
            "classification": item["classification"],
            "reason": item["reason"],
            "source_before_sha256": receipt["source_before"]["sha256"],
            "source_after_sha256": receipt["source_after"]["sha256"],
            "current_source_differs": receipt["source_after"]["sha256"] != accepted_source,
        }

    return {
        "required_labels": required_labels,
        "pilot_labels": pilot_labels,
        "argv": {label: receipts[label]["argv"] for label in final_labels},
        "pilot_reports": pilot_reports,
        "developmental": developmental,
    }


def make_fuzz_plan(
    attempt: str,
    raw_receipts: list[str],
    seed_manifest_path: str,
    generator_path: str,
    seed_root: str,
    accepted_source: str,
) -> dict[str, Any]:
    require(raw_receipts, "at least one --fuzz-receipt is required")
    parsed = [parse_fuzz_receipt(value, attempt) for value in raw_receipts]
    labels = [item[0] for item in parsed]
    require(len(set(labels)) == len(labels), "fuzz receipt labels are duplicated")
    kinds = [item[1] for item in parsed]
    require(set(kinds) == FUZZ_KINDS and len(kinds) == 3, "fuzz receipts must contain exactly one prepared, build, and smoke receipt")
    paths = [item[2] for item in parsed]
    require(len(set(paths)) == len(paths), "fuzz receipt paths are duplicated")
    try:
        seed_root_relative = verify.safe_relative(seed_root, "fuzz seed root")
    except verify.VerificationError as error:
        fail(str(error))
    seed_root_path = ROOT / seed_root_relative
    require(seed_root_path.is_dir() and not seed_root_path.is_symlink(), "fuzz seed root is missing")
    seed_manifest = metadata_reference(seed_manifest_path, "fuzz seed manifest")
    generator = metadata_reference(generator_path, "fuzz generator")
    receipts = [
        {"label": label, "kind": kind, **metadata_reference(path, f"fuzz receipt {label}")}
        for label, kind, path in parsed
    ]
    plan: dict[str, Any] = {
        "attempt": attempt,
        "required_labels": labels,
        "seed_root": seed_root_relative,
        "seed_manifest": seed_manifest,
        "generator": generator,
        "receipts": receipts,
    }
    try:
        verify.check_fuzz_custody({"fuzz": plan})
    except verify.VerificationError as error:
        fail(str(error))
    build_ref = next(item for item in receipts if item["kind"] == "build")
    build_path = verify.bundle_file(build_ref["path"], "fuzz build receipt")
    build = read_json(build_path, "fuzz build receipt")
    require(isinstance(build.get("source_snapshot"), dict), "fuzz build source snapshot is missing")
    require(build["source_snapshot"].get("sha256") == accepted_source, "fuzz build source is not the accepted source")
    return plan


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--attempt", required=True)
    parser.add_argument("--required-label", action="append", default=[])
    parser.add_argument("--pilot", action="append", default=[])
    parser.add_argument("--classify", action="append", default=[])
    parser.add_argument("--classification-file")
    parser.add_argument("--fuzz-receipt", action="append", default=[])
    parser.add_argument("--seed-manifest", default="fuzz/seed-manifest.json")
    parser.add_argument("--generator", default="fuzz/generator.json")
    parser.add_argument("--seed-root", default="fuzz/seeds")
    parser.add_argument("--validation-output", default="validation-plan.json")
    parser.add_argument("--fuzz-output", default="fuzz-plan.json")
    return parser.parse_args()


def main() -> int:
    options = parse_args()
    try:
        attempt = parse_attempt(options.attempt)
        required_labels = list(options.required_label)
        require(all(isinstance(label, str) and verify.LABEL.fullmatch(label) for label in required_labels), "required labels are not path-safe")
        pilots = [parse_pilot(value, attempt) for value in options.pilot]
        classifications = load_classifications(options)
        binaries, accepted_source, build_gate_labels = load_accepted_builds(attempt)
        terminal_paths, started_paths = discover_validation()
        validation_plan = make_validation_plan(
            attempt,
            required_labels,
            pilots,
            classifications,
            terminal_paths,
            started_paths,
            accepted_source,
            binaries,
            build_gate_labels,
        )
        fuzz_plan = make_fuzz_plan(
            attempt,
            options.fuzz_receipt,
            options.seed_manifest,
            options.generator,
            options.seed_root,
            accepted_source,
        )
        validation_output = output_relative(options.validation_output, "validation plan output")
        fuzz_output = output_relative(options.fuzz_output, "fuzz plan output")
        validation_path = ROOT / validation_output
        fuzz_path = ROOT / fuzz_output
        require(not validation_path.exists() and not fuzz_path.exists(), "refusing to replace an existing plan")
        write(validation_path, validation_plan)
        write(fuzz_path, fuzz_plan)
        print(json.dumps({"validation": validation_output, "fuzz": fuzz_output}, sort_keys=True))
        return 0
    except (PlanError, verify.VerificationError, analyze.AnalysisError, OSError, KeyError, TypeError) as error:
        print(f"write-plans failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
