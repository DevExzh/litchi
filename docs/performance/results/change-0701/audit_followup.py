#!/usr/bin/env python3
"""Independently verify the 0701 native/refusal follow-up packet.

This verifier reads only the follow-up raw process records and frozen build
receipts.  It recomputes sample statistics, pair deltas, and >5% triggers so
the follow-up driver is not its own evidence authority.  It never builds or
executes a probe.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BATCH = "0701"
LEGS = (("a0", "baseline"), ("b0", "candidate"), ("b1", "candidate"), ("a3", "baseline"))
PHASES = dict(LEGS)
SAMPLES = 300
WARMUPS = 10
STAT_FIELDS = ("p50_ns", "mean_ns", "p95_ns", "p99_ns", "min_ns", "max_ns")
NATIVE_SOURCES: dict[str, tuple[str, str]] = {
    "one-real": (str(ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"), "one"),
    "noop-real": (str(ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"), "noop"),
    "two-real": (str(ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx"), "two"),
    "one-control": (str(P / "marker-control.pptx"), "one"),
    "one-generated": ("generated:12x8", "one"),
    "one-notes-poi": (str(ROOT / "test-data/poi/test-data/slideshow/prProps.pptx"), "one"),
}
REFUSAL_CASES = {
    "small-valid",
    "generated-12x8-valid",
    "early-name-error",
    "late-root-error",
    "late-root-error-mce",
    "late-missing-relationship",
    "late-missing-relationship-mce",
    "notes-invalid-tail",
    "mixed-conformance",
    "slide-raw-overlimit-16m-to-64m",
}
EXPECTED_COLUMNS = ("capture_ns", "clone_ns", "settext_ns", "commit_ns", "apply_ns", "total_ns")
PAIRS = (("b0", "a0"), ("b1", "a3"), ("a3", "a0"))


def fail(message: str) -> None:
    raise AssertionError(message)


def need(path: Path) -> Path:
    if not path.exists():
        fail(f"missing follow-up evidence: {path.relative_to(P)}")
    return path


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(name: str) -> Any:
    return json.loads(need(P / name).read_text())


def quantile(values: list[int], fraction: float) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * fraction) - 1)]


def stats(values: list[int]) -> dict[str, float | int]:
    if len(values) != SAMPLES or any(value < 0 for value in values):
        fail("follow-up sample vector is malformed")
    return {
        "samples": len(values),
        "p50_ns": statistics.median(values),
        "mean_ns": statistics.mean(values),
        "p95_ns": quantile(values, 0.95),
        "p99_ns": quantile(values, 0.99),
        "min_ns": min(values),
        "max_ns": max(values),
    }


def semantic_identity(value: Any) -> str:
    canonical = json.dumps(value, ensure_ascii=True, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode()).hexdigest()


def parse_native(path: Path) -> tuple[list[str], list[dict[str, int]], dict[str, list[str]]]:
    lines = path.read_text().splitlines()
    header_line = next((line for line in lines if line.startswith("sample\t")), None)
    if header_line is None:
        fail(f"{path}: missing native sample header")
    header = header_line.split("\t")
    rows: list[dict[str, int]] = []
    metadata: dict[str, list[str]] = {}
    for line in lines:
        fields = line.split("\t")
        if fields[0].isdigit():
            if len(fields) != len(header):
                fail(f"{path}: native row/header mismatch")
            values = [int(value) for value in fields]
            if any(value < 0 for value in values):
                fail(f"{path}: negative native sample")
            rows.append(dict(zip(header, values)))
        elif fields[0] not in {"sample", "sample_ns"}:
            metadata[fields[0]] = fields[1:]
    if len(rows) != SAMPLES or [row["sample"] for row in rows] != list(range(SAMPLES)):
        fail(f"{path}: native sample index/count mismatch")
    if metadata.get("samples") != [str(SAMPLES)] or metadata.get("warmups") != [str(WARMUPS)]:
        fail(f"{path}: native sample metadata mismatch")
    if metadata.get("probe") != [BATCH]:
        fail(f"{path}: native probe identity mismatch")
    return header[1:], rows, metadata


def parse_refusal(path: Path) -> tuple[dict[str, str], dict[str, list[int]], dict[str, dict[str, list[str]]]]:
    lines = path.read_text().splitlines()
    header: dict[str, str] = {}
    cases: dict[str, list[int]] = {}
    metadata: dict[str, dict[str, list[str]]] = {}
    current: str | None = None
    for line in lines:
        fields = line.split("\t")
        key = fields[0]
        if key == "case":
            if len(fields) != 2 or fields[1] in cases:
                fail(f"{path}: duplicate/malformed refusal case")
            current = fields[1]
            cases[current] = []
            metadata[current] = {}
        elif key == "sample_ns":
            continue
        elif key.isdigit():
            if current is None or len(fields) != 2 or int(key) != len(cases[current]):
                fail(f"{path}: refusal sample index mismatch")
            value = int(fields[1])
            if value < 0:
                fail(f"{path}: negative refusal sample")
            cases[current].append(value)
        elif key == "all_iterations_passed":
            if fields != ["all_iterations_passed", "true"]:
                fail(f"{path}: refusal iteration failure")
            header[key] = fields[1]
        elif current is None:
            if len(fields) != 2:
                fail(f"{path}: malformed refusal header")
            header[key] = fields[1]
        else:
            metadata[current][key] = fields[1:]
    if header.get("probe") != "0693-refusal" or header.get("samples") != str(SAMPLES) or header.get("warmups") != str(WARMUPS):
        fail(f"{path}: refusal identity/sample metadata mismatch")
    if header.get("all_iterations_passed") != "true" or set(cases) != REFUSAL_CASES:
        fail(f"{path}: refusal case census mismatch")
    for case in REFUSAL_CASES:
        if len(cases[case]) != SAMPLES:
            fail(f"{path}/{case}: refusal sample count mismatch")
        if metadata[case].get("expected_error_debug") is None:
            fail(f"{path}/{case}: expected error identity missing")
        if metadata[case].get("observed_error_debug") not in (None, metadata[case]["expected_error_debug"]):
            fail(f"{path}/{case}: observed error identity differs")
    return header, cases, metadata


def build_receipts() -> dict[tuple[str, str], str]:
    result: dict[tuple[str, str], str] = {}
    for phase in ("baseline", "candidate"):
        native = next(row for row in read(f"builds-{phase}.json") if row.get("label") == "native")
        result[phase, "native"] = native["binary_sha256"]
        result[phase, "refusal"] = read(f"build-refusal-{phase}.json")["binary_sha256"]
    return result


def source_digest(command_source: str, control_sha: str) -> str | None:
    if command_source.startswith("generated:"):
        return None
    path = Path(command_source)
    if path.exists():
        return sha(path)
    if command_source == str(P / "marker-control.pptx"):
        return control_sha
    fail(f"follow-up source disappeared without a retained hash: {command_source}")


def validate_runs() -> tuple[list[dict[str, Any]], dict[str, Any]]:
    records = read("followup-runs.json")
    if not isinstance(records, list) or len(records) != len(NATIVE_SOURCES) * len(LEGS) + len(LEGS):
        fail("follow-up run record count mismatch")
    binaries = build_receipts()
    control = read("control-manifest.json")
    control_sha = control["control_sha256"]
    seen: set[tuple[str, str, str | None]] = set()
    native_identity: dict[str, str] = {}
    refusal_identity: dict[str, str] = {}
    native_rows: list[dict[str, Any]] = []
    refusal_rows: list[dict[str, Any]] = []
    for record in records:
        kind = record.get("kind")
        leg = record.get("leg")
        phase = PHASES.get(leg)
        if kind not in {"native", "refusal"} or phase is None or record.get("phase") != phase or record.get("exit_code") != 0:
            fail(f"invalid follow-up record: {record}")
        output = P / record.get("output", "")
        stderr = P / record.get("stderr", "")
        if not output.is_relative_to(P) or not stderr.is_relative_to(P):
            fail("follow-up output escapes packet")
        if sha(output) != record.get("output_sha256") or sha(stderr) != record.get("stderr_sha256"):
            fail(f"follow-up raw hash mismatch: {output}")
        binary_kind = "native" if kind == "native" else "refusal"
        if record.get("binary_sha256") != binaries[phase, binary_kind]:
            fail(f"follow-up binary binding mismatch: {kind}/{leg}")
        if kind == "native":
            case = record.get("case")
            if case not in NATIVE_SOURCES:
                fail(f"unknown follow-up native case: {case}")
            key = (kind, leg, case)
            expected_source, workflow = NATIVE_SOURCES[case]
            if key in seen:
                fail(f"duplicate follow-up native row: {key}")
            seen.add(key)
            source_arg = record.get("command", [None] * 6)[5]
            if source_arg != expected_source:
                fail(f"native source command mismatch: {case}/{leg}")
            command = ["taskset", "-c", "12", str(ROOT.parent / f"litchi-{BATCH}-bin" / f"{phase}-native"), "phases", str(expected_source), str(SAMPLES), str(WARMUPS), workflow]
            if record.get("command") != command or record.get("workflow") != workflow:
                fail(f"native command/workflow mismatch: {case}/{leg}")
            expected_digest = source_digest(str(expected_source), control_sha)
            if record.get("source_sha256") != expected_digest:
                fail(f"native source hash mismatch: {case}/{leg}")
            expected_output = P / "followup" / "native" / f"{case}-{leg}.tsv"
            if output != expected_output or stderr != expected_output.with_suffix(".stderr"):
                fail(f"native output path mismatch: {case}/{leg}")
            columns, samples, metadata = parse_native(output)
            if set(EXPECTED_COLUMNS) - set(columns):
                fail(f"native phase column census mismatch: {case}/{leg}")
            digest = semantic_identity(metadata)
            if record.get("semantic_identity_sha256") != digest:
                fail(f"native semantic identity receipt mismatch: {case}/{leg}")
            if case in native_identity and native_identity[case] != digest:
                fail(f"native metadata changed across legs: {case}")
            native_identity.setdefault(case, digest)
            native_rows.append({
                "kind": "native", "case": case, "leg": leg, "phase": phase,
                "workflow": workflow, "binary_sha256": record["binary_sha256"],
                "source_sha256": record["source_sha256"],
                "semantic_identity_sha256": digest, "metadata": metadata,
                "stats": {column: stats([row[column] for row in samples]) for column in columns},
            })
        else:
            key = (kind, leg, None)
            if key in seen:
                fail(f"duplicate follow-up refusal row: {leg}")
            seen.add(key)
            command = ["taskset", "-c", "12", str(ROOT.parent / f"litchi-{BATCH}-bin" / f"{phase}-refusal"), "matrix", str(SAMPLES), str(WARMUPS)]
            if record.get("command") != command:
                fail(f"refusal command mismatch: {leg}")
            expected_output = P / "followup" / "refusal" / f"{leg}.tsv"
            if output != expected_output or stderr != expected_output.with_suffix(".stderr"):
                fail(f"refusal output path mismatch: {leg}")
            _header, cases, metadata = parse_refusal(output)
            if record.get("case_count") != len(REFUSAL_CASES):
                fail(f"refusal case count receipt mismatch: {leg}")
            digest_map = {case: semantic_identity(fields) for case, fields in metadata.items()}
            if record.get("semantic_identity_sha256") != digest_map:
                fail(f"refusal semantic identity receipt mismatch: {leg}")
            for case, digest in digest_map.items():
                if case in refusal_identity and refusal_identity[case] != digest:
                    fail(f"refusal metadata changed across legs: {case}")
                refusal_identity.setdefault(case, digest)
                refusal_rows.append({
                    "kind": "refusal", "case": case, "leg": leg, "phase": phase,
                    "binary_sha256": record["binary_sha256"],
                    "semantic_identity_sha256": digest, "metadata": metadata[case],
                    "stats": {"total_ns": stats(cases[case])},
                })
    expected_native = {("native", leg, case) for leg, _phase in LEGS for case in NATIVE_SOURCES}
    expected_refusal = {("refusal", leg, None) for leg, _phase in LEGS}
    if seen != expected_native | expected_refusal:
        fail("follow-up native/refusal coverage is incomplete")
    return native_rows + refusal_rows, {"native": native_rows, "refusal": refusal_rows}


def expected_comparisons(summary: dict[str, list[dict[str, Any]]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for kind in ("native", "refusal"):
        rows = summary[kind]
        for case in sorted({row["case"] for row in rows}):
            by_leg = {row["leg"]: row for row in rows if row["case"] == case}
            for metric in sorted({metric for row in by_leg.values() for metric in row["stats"]}):
                for candidate, baseline in PAIRS:
                    candidate_stats = by_leg[candidate]["stats"][metric]
                    baseline_stats = by_leg[baseline]["stats"][metric]
                    result.append({
                        "kind": kind,
                        "case": case,
                        "phase": metric,
                        "pair": f"{candidate}/{baseline}",
                        "candidate_leg": candidate,
                        "baseline_leg": baseline,
                        "candidate": {name: candidate_stats[name] for name in STAT_FIELDS},
                        "baseline": {name: baseline_stats[name] for name in STAT_FIELDS},
                        "delta_pct": {name: (candidate_stats[name] / baseline_stats[name] - 1) * 100 for name in STAT_FIELDS},
                        "delta_ns": {name: candidate_stats[name] - baseline_stats[name] for name in STAT_FIELDS},
                    })
    return result


def compare_json(actual: Any, expected: Any, label: str) -> None:
    if isinstance(expected, float):
        if not isinstance(actual, (float, int)) or not math.isclose(float(actual), expected, rel_tol=0.0, abs_tol=1e-9):
            fail(f"{label}: numeric mismatch {actual!r} != {expected!r}")
    elif isinstance(expected, dict):
        if not isinstance(actual, dict) or set(actual) != set(expected):
            fail(f"{label}: object shape mismatch")
        for key in expected:
            compare_json(actual[key], expected[key], f"{label}/{key}")
    elif isinstance(expected, list):
        if not isinstance(actual, list) or len(actual) != len(expected):
            fail(f"{label}: list shape mismatch")
        for index, (left, right) in enumerate(zip(actual, expected)):
            compare_json(left, right, f"{label}/{index}")
    elif actual != expected:
        fail(f"{label}: value mismatch {actual!r} != {expected!r}")


def main() -> None:
    _rows, summary = validate_runs()
    for kind in summary:
        summary[kind].sort(key=lambda row: (row["case"], row["leg"]))
    recorded_summary = read("followup-summary.json")
    compare_json(recorded_summary, summary, "followup-summary")
    comparisons = expected_comparisons(summary)
    compare_json(read("followup-comparisons.json"), comparisons, "followup-comparisons")
    triggers = [
        {
            "kind": row["kind"], "case": row["case"], "phase": row["phase"],
            "pair": row["pair"], "metric": metric,
            "delta_pct": value, "delta_ns": row["delta_ns"][metric],
        }
        for row in comparisons
        if row["pair"] in {"b0/a0", "b1/a3"}
        for metric, value in row["delta_pct"].items()
        if value > 5
    ]
    compare_json(read("followup-triggers.json"), triggers, "followup-triggers")
    print(f"PASS: 0701 follow-up {len(summary['native'])} native and {len(summary['refusal'])} refusal rows; {len(comparisons)} comparisons")


if __name__ == "__main__":
    main()
