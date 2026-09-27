#!/usr/bin/env python3
"""Offline replay for the 0798 checked-attribute consumption census.

The root driver owns both builds and every child process.  This module only
reads retained receipts and JSON reports after capture has terminated.  It
does not invoke Cargo, the probe, a profiler, or a timing command.

The census is deliberately treated as lifecycle evidence for the checked
iterator instances observed in the calling thread.  It is not an allocator,
RSS, instruction, or native-latency measurement.  The baseline and census
reports are compared to the sealed 0794 qualification reports for semantic
and publication parity while elapsed values are retained only as opaque
evidence and are never summarized.
"""

from __future__ import annotations

import argparse
import collections
import hashlib
import json
import math
import re
import sys
import subprocess
from pathlib import Path
from typing import Any, Iterable, Mapping, NoReturn, Sequence

import custody as c


HERE = Path(__file__).resolve().parent
SCHEMA = "litchi.performance.0798.analysis.v1"
PLAN_SCHEMA = "litchi.performance.0798.v1"
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor")
MODES = ("capture", "commit", "lifecycle")
LEGS = ("before", "after")
SAFE_NAME = re.compile(r"^[A-Za-z0-9_.-]+$")
SHA256 = re.compile(r"^[0-9a-f]{64}$")
REVISION = re.compile(r"^[0-9a-f]{40}$")


class EvidenceError(ValueError):
    """A retained artifact is missing, stale, malformed, or contradictory."""


def fail(message: str) -> NoReturn:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path, label: str) -> Any:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"{label}: invalid JSON: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and SHA256.fullmatch(value) is not None


def integer(value: Any, label: str, *, positive: bool = False) -> int:
    require(isinstance(value, int) and not isinstance(value, bool),
            f"{label}: expected an integer")
    require(value > 0 if positive else value >= 0,
            f"{label}: expected {'a positive' if positive else 'a non-negative'} integer")
    return value


def finite(value: Any, label: str) -> float:
    require(isinstance(value, (int, float)) and not isinstance(value, bool),
            f"{label}: expected a number")
    result = float(value)
    require(math.isfinite(result), f"{label}: number is not finite")
    return result


def packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: path is missing")
    marker = f"/{HERE.name}/"
    if marker in raw:
        path = HERE / raw.split(marker, 1)[1]
    else:
        candidate = Path(raw)
        path = candidate if candidate.is_absolute() else HERE / candidate
    path = path.resolve()
    try:
        path.relative_to(HERE.resolve())
    except ValueError as error:
        raise EvidenceError(f"{label}: path escapes packet: {raw}") from error
    return path


def external_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: path is missing")
    path = Path(raw)
    require(path.is_absolute(), f"{label}: path is not absolute")
    return path.resolve()


def artifact(value: Any, label: str, *, allow_missing: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: artifact descriptor is missing")
    path = packet_path(value.get("path"), label)
    size = value.get("bytes", value.get("size"))
    integer(size, f"{label}.bytes")
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256: invalid digest")
    if not path.is_file() or path.is_symlink():
        require(allow_missing, f"{label}: artifact is missing: {path}")
        return {"path": str(path.relative_to(HERE)), "bytes": size, "sha256": digest}
    require(path.stat().st_size == size, f"{label}: byte count changed")
    require(sha256(path) == digest, f"{label}: SHA-256 changed")
    return {"path": str(path.relative_to(HERE)), "bytes": size, "sha256": digest}


def external_artifact(value: Any, label: str, *, allow_missing: bool = True) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: artifact descriptor is missing")
    path = external_path(value.get("path"), label)
    size = value.get("bytes", value.get("size"))
    integer(size, f"{label}.bytes", positive=True)
    digest = value.get("sha256", value.get("digest"))
    require(is_sha(digest), f"{label}.sha256: invalid digest")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == size, f"{label}: byte count changed")
        require(sha256(path) == digest, f"{label}: SHA-256 changed")
    else:
        require(allow_missing, f"{label}: binary is missing: {path}")
        cleanup = read_json(HERE / "cleanup.json", "cleanup witness")
        require(isinstance(cleanup, dict), "cleanup witness is malformed")
        require(cleanup.get("target_removed") is True and not c.TARGET.exists(),
                "cleanup witness does not prove target removal")
        expected = {"path": str(path), "bytes": size, "sha256": digest}
        require(expected in cleanup.get("removed_binaries", []),
                f"{label}: missing binary is not in cleanup witness")
    return {"path": str(path), "bytes": size, "sha256": digest}


def artifact_descriptor(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing: {path}")
    return {"path": str(path.relative_to(HERE)), "bytes": path.stat().st_size,
            "sha256": sha256(path)}


def frozen_plan() -> tuple[dict[str, Any], list[dict[str, str]]]:
    plan = read_json(HERE / "plan.json", "plan")
    require(isinstance(plan, dict) and plan.get("schema") == PLAN_SCHEMA,
            "plan schema changed")
    require(plan.get("cpu") == 12, "plan CPU changed")
    require(plan.get("counts") == {"reports": 45, "samples": 45},
            "plan report counts changed")
    cases = plan.get("cases")
    require(isinstance(cases, list) and len(cases) == 15, "plan case count changed")
    normalized: list[dict[str, str]] = []
    expected = [{"mode": mode, "shape": shape} for shape in SHAPES for mode in MODES]
    require(cases == expected, "plan case order changed")
    for index, case in enumerate(cases):
        require(isinstance(case, dict), f"plan case {index}: malformed")
        shape, mode = case.get("shape"), case.get("mode")
        require(shape in SHAPES and mode in MODES, f"plan case {index}: unknown case")
        normalized.append({"shape": shape, "mode": mode})
    control = plan.get("control")
    require(control == {"blocks": 1, "leg": "before", "samples": 1, "warmup": 0},
            "control schedule changed")
    census = plan.get("census")
    require(isinstance(census, dict), "census schedule is missing")
    require(census.get("blocks") == 2 and census.get("leg") == "after"
            and census.get("samples") == 1 and census.get("warmup") == 0
            and census.get("case_orders") == ["forward", "reverse"],
            "census schedule changed")
    qualification = plan.get("qualification")
    require(isinstance(qualification, dict), "qualification policy is missing")
    for key in ("exact_semantic_and_output_parity", "instance_conservation",
                "live_at_finish_zero", "overflow_absent", "repeat_histograms_equal"):
        require(qualification.get(key) is True, f"qualification policy {key} changed")
    return plan, normalized


def frozen_source() -> dict[str, Any]:
    value = read_json(HERE / "source.json", "source")
    require(isinstance(value, dict), "source manifest is malformed")
    revision = value.get("revision")
    require(isinstance(revision, str) and REVISION.fullmatch(revision),
            "source revision is malformed")
    files = value.get("files")
    require(isinstance(files, dict) and files, "source file census is missing")
    for name, digest in files.items():
        require(isinstance(name, str) and name and is_sha(digest),
                "source file census contains an invalid entry")
    current = c.source()
    require(current["files"] == files, "production source differs from frozen baseline")
    require(subprocess.run(["git", "merge-base", "--is-ancestor", revision, "HEAD"],
                           cwd=c.ROOT, check=False).returncode == 0,
            "frozen source revision is not an ancestor of HEAD")
    return value


def inherited_qualification() -> dict[tuple[str, str], dict[str, Any]]:
    inheritance = read_json(HERE / "inheritance.json", "inheritance")
    descriptor = inheritance.get("qualification_seal")
    require(isinstance(descriptor, dict), "qualification seal descriptor missing")
    seal_path = Path(descriptor["path"])
    require(seal_path.is_file() and sha256(seal_path) == descriptor["sha256"],
            "qualification seal identity changed")
    seal = read_json(seal_path, "qualification seal")
    files = seal.get("files")
    require(isinstance(files, dict), "qualification seal file map missing")
    root = seal_path.parent
    result: dict[tuple[str, str], dict[str, Any]] = {}
    for shape in SHAPES:
        for mode in MODES:
            relative = f"qualification/0-{shape}-{mode}-before.json"
            digest = files.get(relative)
            require(is_sha(digest), f"qualification seal lacks {relative}")
            path = root / relative
            require(path.is_file() and sha256(path) == digest,
                    f"qualification artifact changed: {relative}")
            report = read_json(path, relative)
            require(isinstance(report, dict), f"qualification report malformed: {relative}")
            result[(shape, mode)] = report
    return result


def validate_build(leg: str, source: Mapping[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    directory = HERE / f"build-{leg}"
    build = read_json(directory / "build.json", f"build-{leg}")
    require(isinstance(build, dict), f"build-{leg}: malformed")
    binary = external_artifact(build.get("binary"), f"build-{leg} binary")
    source_descriptor = artifact(build.get("source"), f"build-{leg} source")
    source_path = HERE / f"build-{leg}" / "source.json"
    require(source_descriptor["path"] == str(source_path.relative_to(HERE)),
            f"build-{leg}: source path changed")
    build_source = read_json(source_path, f"build-{leg} source manifest")
    require(isinstance(build_source, dict) and build_source.get("revision") == source["revision"],
            f"build-{leg}: source revision changed")
    rows = build.get("rows")
    require(isinstance(rows, list) and len(rows) == 4, f"build-{leg}: gate count changed")
    previous_end = float("-inf")
    for index, row in enumerate(rows):
        label = f"build-{leg}[{index}]"
        require(isinstance(row, dict) and row.get("exit_code") == 0, f"{label}: failed")
        started, ended = finite(row.get("started"), f"{label}.started"), finite(row.get("ended"), f"{label}.ended")
        require(previous_end <= started <= ended, f"{label}: serial timing order changed")
        previous_end = ended
        artifact(row.get("log"), f"{label}.log")
        command = row.get("command")
        require(isinstance(command, list) and all(isinstance(x, str) for x in command),
                f"{label}: command malformed")
        if command and command[0] == "cargo":
            require("--locked" in command, f"{label}: unlocked Cargo command")
    inputs = artifact(build.get("inputs"), f"build-{leg} inputs")
    require(inputs["path"] == str((directory / "inputs.json").relative_to(HERE)),
            f"build-{leg}: inputs path changed")
    return build, binary


def validate_receipts(lane: str, plan: Mapping[str, Any], source: Mapping[str, Any],
                      binary: Mapping[str, Any], expected_leg: str) -> list[dict[str, Any]]:
    directory = HERE / lane
    rows = read_json(directory / "receipts.json", f"{lane} receipts")
    require(isinstance(rows, list), f"{lane}: receipts malformed")
    count = 15 if lane == "control" else 30
    require(len(rows) == count, f"{lane}: receipt count changed")
    complete = read_json(directory / "complete.json", f"{lane} complete")
    require(isinstance(complete, dict) and complete.get("children") == count,
            f"{lane}: completion count changed")
    artifact(complete.get("source"), f"{lane} completion source")
    artifact(complete.get("receipts"), f"{lane} completion receipts")
    lane_source = read_json(directory / "source.json", f"{lane} source")
    require(lane_source == source, f"{lane}: source changed during capture")
    expected: list[tuple[int, dict[str, str]]] = []
    blocks = 1 if lane == "control" else 2
    for block in range(blocks):
        order = plan["cases"] if block == 0 else list(reversed(plan["cases"]))
        expected.extend((block, case) for case in order)
    previous_end = float("-inf")
    for index, (row, (block, case)) in enumerate(zip(rows, expected)):
        label = f"{lane}[{index}]"
        require(isinstance(row, dict), f"{label}: receipt malformed")
        for key, value in (("lane", lane), ("block", block), ("shape", case["shape"]),
                           ("mode", case["mode"]), ("leg", expected_leg),
                           ("exit_code", 0)):
            require(row.get(key) == value, f"{label}: {key} changed")
        started, ended = finite(row.get("started"), f"{label}.started"), finite(row.get("ended"), f"{label}.ended")
        require(previous_end <= started <= ended, f"{label}: receipt order changed")
        previous_end = ended
        row_binary = row.get("binary")
        require(isinstance(row_binary, dict)
                and row_binary.get("path") == binary["path"]
                and row_binary.get("bytes") == binary["bytes"]
                and row_binary.get("sha256") == binary["sha256"],
                f"{label}: binary identity changed")
        external_artifact(row_binary, f"{label} binary")
        report = artifact(row.get("report"), f"{label} report")
        log = artifact(row.get("log"), f"{label} log")
        require(Path(report["path"]).name == f"{block}-{case['shape']}-{case['mode']}.json",
                f"{label}: report name changed")
        require(Path(log["path"]).name == f"{block}-{case['shape']}-{case['mode']}.log",
                f"{label}: log name changed")
        command = row.get("command")
        require(isinstance(command, list) and all(isinstance(x, str) for x in command),
                f"{label}: command malformed")
        require("taskset" in command and "12" in command,
                f"{label}: command is not pinned to CPU 12")
        require(str(binary["path"]) in command, f"{label}: binary missing from command")
        for option, value in (("--shape", case["shape"]), ("--mode", case["mode"]),
                              ("--samples", "1"), ("--warmup", "0")):
            require(option in command and command[command.index(option) + 1] == value,
                    f"{label}: {option} changed")
        require("--output" in command, f"{label}: output option missing")
        require(packet_path(command[command.index("--output") + 1], f"{label} output")
                == HERE / report["path"], f"{label}: output path changed")
        row["report_descriptor"] = report
        row["log_descriptor"] = log
    return rows


def report_sample(report: Mapping[str, Any], label: str) -> dict[str, Any]:
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == 1, f"{label}: sample count changed")
    sample = samples[0]
    require(isinstance(sample, dict) and sample.get("index") == 0,
            f"{label}: sample malformed")
    integer(sample.get("elapsed_ns"), f"{label}.elapsed_ns", positive=True)
    return sample


def parity_projection(report: Mapping[str, Any], sample: Mapping[str, Any]) -> dict[str, Any]:
    """Select semantic/publication fields and deliberately omit elapsed data."""

    projection: dict[str, Any] = {}
    for key in ("mode", "shape", "slides", "shapes_per_slide",
                "timing_scope", "marker", "source", "fixture", "warmup",
                "samples_requested"):
        projection[key] = report.get(key)
    projection["source_sha256"] = sample.get("source_sha256")
    projection["output"] = sample.get("output")
    projection["verification"] = sample.get("verification")
    metrics = sample.get("metrics")
    require(isinstance(metrics, dict), "sample metrics are missing")
    projection["metrics"] = {key: value for key, value in metrics.items()
                              if key != "elapsed_ns"}
    return projection


def validate_report(report: Any, shape: str, mode: str, leg: str,
                    reference: Mapping[str, Any], label: str,
                    *, census_expected: bool) -> tuple[dict[str, Any], dict[str, Any] | None]:
    require(isinstance(report, dict), f"{label}: report malformed")
    require(report.get("schema") == "litchi.pptx.attribute-census-probe.v1",
            f"{label}: report schema changed")
    require(report.get("shape") == shape and report.get("mode") == mode,
            f"{label}: case identity changed")
    require(report.get("warmup") == 0 and report.get("samples_requested") == 1,
            f"{label}: sample policy changed")
    require(report.get("timing_scope") in {
        "Package::opened_presentation only",
        "Transaction::commit only; package capture and one set_shape_text staging are outside the clock",
        "Package::opened_presentation, edit, set_shape_text, commit, apply_opened_presentation_commit, and Package::to_bytes",
    }, f"{label}: timing scope changed")
    sample = report_sample(report, label)
    reference_sample = report_sample(reference, f"{label} reference")
    require(parity_projection(report, sample) == parity_projection(reference, reference_sample),
            f"{label}: semantic/publication parity differs from sealed qualification")
    raw_census = sample.get("census", report.get("census"))
    if census_expected:
        require(raw_census is not None, f"{label}: census evidence missing")
        return dict(report), normalize_census(raw_census, label)
    require(raw_census is None, f"{label}: baseline unexpectedly contains census evidence")
    return dict(report), None


def _bytes(value: Any, label: str) -> bytes:
    require(isinstance(value, list), f"{label}: expected a byte array")
    result = bytearray()
    for index, item in enumerate(value):
        integer(item, f"{label}[{index}]")
        require(item <= 255, f"{label}[{index}]: byte is out of range")
        result.append(item)
    return bytes(result)


def lexical_scan(source: bytes) -> tuple[bytes, int, int, int, bool]:
    """Independently count the unchecked lexical items in a BytesStart body.

    quick-xml's `BytesStart::as_ref()` contains the element name and its
    attributes, without the surrounding `<` and `>`.  The generated fixtures
    contain ordinary quoted XML attributes; this scanner also reports the
    first malformed value so a future fixture cannot silently make the census
    counts look plausible.
    """

    whitespace = b" \t\r\n"
    index = 0
    while index < len(source) and source[index] in whitespace:
        index += 1
    name_start = index
    while index < len(source) and source[index] not in whitespace:
        index += 1
    name = source[name_start:index]
    require(name, "census tag has no element name")
    attributes = items = errors = 0
    complete = True
    while True:
        while index < len(source) and source[index] in whitespace:
            index += 1
        if index >= len(source):
            break
        if source[index] == ord('/'):
            index += 1
            while index < len(source) and source[index] in whitespace:
                index += 1
            require(index == len(source), "census tag has trailing bytes after slash")
            break
        items += 1
        attr_start = index
        while index < len(source) and source[index] not in whitespace + b"=":
            index += 1
        if attr_start == index:
            errors += 1
            complete = False
            break
        while index < len(source) and source[index] in whitespace:
            index += 1
        if index >= len(source) or source[index] != ord('='):
            errors += 1
            complete = False
            break
        index += 1
        while index < len(source) and source[index] in whitespace:
            index += 1
        if index >= len(source):
            errors += 1
            complete = False
            break
        quote = source[index]
        if quote in (ord('"'), ord("'")):
            index += 1
            value_start = index
            while index < len(source) and source[index] != quote:
                index += 1
            if index >= len(source):
                errors += 1
                complete = False
                break
            index += 1
            del value_start
        else:
            value_start = index
            while index < len(source) and source[index] not in whitespace:
                index += 1
            if value_start == index:
                errors += 1
                complete = False
                break
        attributes += 1
    return name, attributes, items, errors, complete


def normalize_census(raw: Any, label: str) -> dict[str, Any]:
    """Validate and aggregate the exact 0798 per-instance observation rows."""

    require(isinstance(raw, dict), f"{label}: census is not an object")
    require(raw.get("owner") == "litchi-opc::xml_attributes::CheckedAttributes",
            f"{label}: census owner changed")
    starts = integer(raw.get("iterator_starts"), f"{label}.iterator_starts")
    clones = integer(raw.get("iterator_clones"), f"{label}.iterator_clones")
    drops = integer(raw.get("iterator_drops"), f"{label}.iterator_drops")
    live = integer(raw.get("live_instances_at_finish"), f"{label}.live_instances_at_finish")
    require(raw.get("counter_saturated") is False, f"{label}: census counter saturated")
    raw_rows = integer(raw.get("raw_instance_rows"), f"{label}.raw_instance_rows")
    require(raw.get("instance_identity_qualified") is True,
            f"{label}: probe lineage/instance qualification failed")
    require(raw.get("aggregation_conserved") is True,
            f"{label}: probe aggregation conservation failed")
    rows = raw.get("rows")
    require(isinstance(rows, list), f"{label}: census rows missing")

    scalar_totals = collections.Counter()
    tag_histogram: collections.Counter[str] = collections.Counter()
    element_histogram: collections.Counter[str] = collections.Counter()
    attribute_histogram: collections.Counter[str] = collections.Counter()
    tag_records: dict[str, dict[str, Any]] = {}
    drop_classes = collections.Counter()
    early_drop_flags = collections.Counter()
    partial_flags = collections.Counter()
    row_frequency_total = 0
    for index, row in enumerate(rows):
        row_label = f"{label}.rows[{index}]"
        require(isinstance(row, dict), f"{row_label}: row is malformed")
        frequency = integer(row.get("frequency"), f"{row_label}.frequency", positive=True)
        row_frequency_total += frequency
        clone = row.get("is_clone")
        require(isinstance(clone, bool), f"{row_label}.is_clone: expected boolean")
        start_prefix = integer(row.get("starting_successful_yields"),
                               f"{row_label}.starting_successful_yields")
        source_bytes = _bytes(row.get("tag_bytes"), f"{row_label}.tag_bytes")
        element_bytes = _bytes(row.get("element_name_bytes"), f"{row_label}.element_name_bytes")
        name, lexical_attributes, lexical_items, lexical_errors, lexical_complete = lexical_scan(source_bytes)
        require(element_bytes == name, f"{row_label}: element name differs from tag source")
        require(row.get("lexical_attribute_count") == lexical_attributes,
                f"{row_label}: lexical attribute count differs from independent scan")
        require(row.get("lexical_item_count") == lexical_items,
                f"{row_label}: lexical item count differs from independent scan")
        require(row.get("lexical_error_count") == lexical_errors,
                f"{row_label}: lexical error count differs from independent scan")
        require(row.get("lexical_scan_completed") is lexical_complete,
                f"{row_label}: lexical completion differs from independent scan")
        require(lexical_errors == 0 and lexical_items == lexical_attributes,
                f"{row_label}: generated tag has an unexpected lexical error")
        next_calls = integer(row.get("next_calls"), f"{row_label}.next_calls")
        successful = integer(row.get("successful_yields"), f"{row_label}.successful_yields")
        errors = integer(row.get("error_yields"), f"{row_label}.error_yields")
        ends = integer(row.get("end_yields"), f"{row_label}.end_yields")
        require(next_calls == successful + errors + ends,
                f"{row_label}: next transition counts do not conserve")
        termination = row.get("termination")
        require(termination in {"active", "error", "exhausted", "dropped", "live-at-finish"},
                f"{row_label}: unknown termination")
        for key in ("dropped", "early_drop", "partial_consumption", "live_at_finish"):
            require(isinstance(row.get(key), bool), f"{row_label}.{key}: expected boolean")
        require(row.get("counter_saturated") is False, f"{row_label}: row counter saturated")
        require(row["live_at_finish"] is False, f"{row_label}: row live at finish")
        require(row["dropped"] is True, f"{row_label}: row was not dropped")
        require(termination not in {"active", "live-at-finish"},
                f"{row_label}: non-terminal row survived finish")
        expected_termination = ("error" if errors else
                                ("exhausted" if ends else "dropped"))
        require(termination == expected_termination,
                f"{row_label}: termination does not match observed next outcomes")
        require(row["early_drop"] is (expected_termination == "dropped"),
                f"{row_label}: early-drop flag does not match termination")
        if termination == "exhausted":
            classification = "full"
        elif next_calls == 0:
            classification = "never"
        else:
            classification = "early"
        drop_classes[classification] += frequency
        early_drop_flags[str(row["early_drop"])] += frequency
        partial_flags[str(row["partial_consumption"])] += frequency
        for key, value in (("next_calls", next_calls), ("next_ok", successful),
                           ("next_err", errors), ("next_none", ends),
                           ("lexical_attribute_total", lexical_attributes),
                           ("lexical_item_total", lexical_items),
                           ("lexical_error_total", lexical_errors)):
            scalar_totals[key] += value * frequency
        scalar_totals["attribute_total"] += lexical_attributes * frequency
        tag_key = source_bytes.hex()
        tag_histogram[tag_key] += frequency
        element_histogram[element_bytes.hex()] += frequency
        attribute_histogram[str(lexical_attributes)] += frequency
        record = tag_records.setdefault(tag_key, {
            "tag_bytes": list(source_bytes), "element_name_bytes": list(element_bytes),
            "frequency": 0,
        })
        record["frequency"] += frequency

    require(row_frequency_total == raw_rows,
            f"{label}: aggregated row frequencies do not equal raw_instance_rows")
    require(starts + clones == drops + live,
            f"{label}: instance conservation failed")
    require(live == 0, f"{label}: live instances remain at finish")
    require(drops == row_frequency_total, f"{label}: drop count differs from rows")
    require(sum(drop_classes.values()) == drops, f"{label}: drop classification does not conserve")
    require(scalar_totals["next_calls"] == scalar_totals["next_ok"]
            + scalar_totals["next_err"] + scalar_totals["next_none"],
            f"{label}: aggregate next counts do not conserve")
    require(sum(tag_histogram.values()) == raw_rows,
            f"{label}: tag histogram does not sum to raw rows")
    return {
        "iterator_starts": starts,
        "iterator_clones": clones,
        "iterator_drops": drops,
        "live_instances_at_finish": live,
        "raw_instance_rows": raw_rows,
        "counter_saturated": False,
        "instance_identity_qualified": True,
        "aggregation_conserved": True,
        "next_transitions": {key: scalar_totals[key] for key in
                              ("next_calls", "next_ok", "next_err", "next_none")},
        "drop_classification": dict(sorted(drop_classes.items())),
        "early_drop_flags": dict(sorted(early_drop_flags.items())),
        "partial_consumption_flags": dict(sorted(partial_flags.items())),
        "lexical_totals": {key: scalar_totals[key] for key in
                           ("lexical_attribute_total", "lexical_item_total",
                            "lexical_error_total")},
        "attribute_total": scalar_totals["attribute_total"],
        "attribute_count_histogram": dict(sorted(attribute_histogram.items(),
                                                  key=lambda item: int(item[0]))),
        "tag_histogram": dict(sorted(tag_histogram.items())),
        "element_name_hex_histogram": dict(sorted(element_histogram.items())),
        "tag_records": [tag_records[key] for key in sorted(tag_records)],
        "row_count": len(rows),
    }


def census_totals(rows: Iterable[Mapping[str, Any]]) -> dict[str, Any]:
    scalar_keys = ("iterator_starts", "iterator_clones", "iterator_drops",
                   "live_instances_at_finish", "raw_instance_rows", "attribute_total")
    totals = {key: 0 for key in scalar_keys}
    totals["next_transitions"] = collections.Counter()
    totals["drop_classification"] = collections.Counter()
    totals["lexical_totals"] = collections.Counter()
    tags: collections.Counter[str] = collections.Counter()
    elements: collections.Counter[str] = collections.Counter()
    attributes: collections.Counter[str] = collections.Counter()
    tag_records: dict[str, dict[str, Any]] = {}
    for row in rows:
        for key in scalar_keys:
            totals[key] += row[key]
        totals["next_transitions"].update(row["next_transitions"])
        totals["drop_classification"].update(row["drop_classification"])
        totals["lexical_totals"].update(row["lexical_totals"])
        tags.update(row["tag_histogram"])
        elements.update(row["element_name_hex_histogram"])
        attributes.update(row["attribute_count_histogram"])
        for record in row["tag_records"]:
            key = bytes(record["tag_bytes"]).hex()
            current = tag_records.setdefault(key, {
                "tag_bytes": list(record["tag_bytes"]),
                "element_name_bytes": list(record["element_name_bytes"]),
                "frequency": 0,
            })
            require(current["element_name_bytes"] == record["element_name_bytes"],
                    f"tag record {key}: element name differs across rows")
            current["frequency"] += record["frequency"]
    totals["next_transitions"] = dict(sorted(totals["next_transitions"].items()))
    totals["drop_classification"] = dict(sorted(totals["drop_classification"].items()))
    totals["lexical_totals"] = dict(sorted(totals["lexical_totals"].items()))
    totals["tag_histogram"] = dict(sorted(tags.items()))
    totals["element_name_hex_histogram"] = dict(sorted(elements.items()))
    totals["tag_records"] = [tag_records[key] for key in sorted(tag_records)]
    totals["attribute_count_histogram"] = dict(sorted(attributes.items(),
                                                       key=lambda item: int(item[0])))
    return totals


def analyze() -> dict[str, Any]:
    plan, cases = frozen_plan()
    source = frozen_source()
    references = inherited_qualification()
    before_build, before_binary = validate_build("before", source)
    after_build, after_binary = validate_build("after", source)
    control_rows = validate_receipts("control", plan, source, before_binary, "before")
    census_rows = validate_receipts("census", plan, source, after_binary, "after")

    control_results: dict[tuple[str, str], dict[str, Any]] = {}
    for row in control_rows:
        key = (row["shape"], row["mode"])
        report = read_json(HERE / row["report_descriptor"]["path"], f"control {key}")
        checked, census = validate_report(report, *key, "before", references[key],
                                           f"control/{key[0]}/{key[1]}", census_expected=False)
        require(census is None, "baseline census unexpectedly present")
        control_results[key] = checked
    require(len(control_results) == 15, "control cases are incomplete")

    census_results: list[dict[str, Any]] = []
    by_identity: dict[tuple[int, str, str], dict[str, Any]] = {}
    for row in census_rows:
        key = (row["shape"], row["mode"])
        report = read_json(HERE / row["report_descriptor"]["path"], f"census {key}")
        checked, census = validate_report(report, *key, "after", references[key],
                                           f"census/{row['block']}/{key[0]}/{key[1]}",
                                           census_expected=True)
        require(census is not None, "census normalization unexpectedly empty")
        identity = (row["block"], key[0], key[1])
        require(identity not in by_identity, f"duplicate census identity: {identity}")
        by_identity[identity] = census
        census_results.append({
            "block": row["block"], "shape": key[0], "mode": key[1],
            "report": row["report_descriptor"], "report_sha256": row["report_descriptor"]["sha256"],
            "census": census,
        })
        del checked
    require(len(by_identity) == 30, "census cases are incomplete")

    repeat_rows: list[dict[str, Any]] = []
    repeat_equal = True
    for case in cases:
        key = (case["shape"], case["mode"])
        first, second = by_identity[(0, *key)], by_identity[(1, *key)]
        equal = first == second
        repeat_equal = repeat_equal and equal
        repeat_rows.append({"shape": key[0], "mode": key[1], "equal": equal,
                            "block0": first, "block1": second})
    require(repeat_equal, "reverse census repeat differs in counters or histograms")

    all_rows = [entry["census"] for entry in census_results]
    total = census_totals(all_rows)
    by_case_mode: dict[str, Any] = {}
    for case in cases:
        key = (case["shape"], case["mode"])
        rows = [by_identity[(block, *key)] for block in range(2)]
        by_case_mode[f"{key[0]}/{key[1]}"] = {
            "shape": key[0], "mode": key[1], "repeats": rows,
            "repeat_equal": rows[0] == rows[1], "totals": census_totals(rows),
        }

    return {
        "schema": SCHEMA,
        "reports": 45,
        "samples": 45,
        "baseline_reports": 15,
        "census_reports": 30,
        "baseline_semantic_output_parity": True,
        "census_semantic_output_parity": True,
        "source_files": len(source["files"]),
        "builds": {"before": 4, "after": 4},
        "census": {
            "scope": "OPC checked-iterator instances in the calling thread during generated PPTX operation regions",
            "by_case_mode": by_case_mode,
            "repeat_comparison": repeat_rows,
            "repeat_histograms_equal": repeat_equal,
            "totals": total,
            "next_transitions": total["next_transitions"],
            "drop_classification": total["drop_classification"],
            "tag_histogram": total["tag_histogram"],
            "element_name_hex_histogram": total["element_name_hex_histogram"],
            "tag_records": total["tag_records"],
            "attribute_count_histogram": total["attribute_count_histogram"],
            "attribute_total": total["attribute_total"],
            "qualification": {
                "instance_conservation": all(row["iterator_starts"] + row["iterator_clones"]
                                              == row["iterator_drops"]
                                              + row["live_instances_at_finish"] for row in all_rows),
                "live_at_finish_zero": all(row["live_instances_at_finish"] == 0 for row in all_rows),
                "overflow_absent": all(row["counter_saturated"] is False for row in all_rows),
                "repeat_histograms_equal": repeat_equal,
            },
            "diagnostic_only": True,
        },
        "claims": {
            "allocation_calls": False,
            "native_latency": False,
            "rss": False,
            "production_adoption": False,
            "public_workflow_speedup": False,
        },
        "binary_identities": {"before": before_binary, "after": after_binary},
    }


def main(argv: Sequence[str] | None = None) -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args(argv)
    try:
        result = analyze()
        if args.write:
            c.write(HERE / "analysis.json", result)
        if args.check:
            require(read_json(HERE / "analysis.json", "analysis") == result,
                    "analysis.json differs from deterministic replay")
    except EvidenceError as error:
        print(f"0798 replay FAIL: {error}", file=sys.stderr)
        return 1
    print("0798 replay PASS: 45 reports / 45 samples; census is diagnostic only")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
