#!/usr/bin/env python3
"""Measure bounded MCE-marker search controls against the frozen oracle.

The source cases are deliberately small and deterministic.  This driver writes
their bytes before any oracle invocation, records the exact identity command
and its stdout/stderr, then runs the frozen oracle's timing mode in the fixed
AA/ABBA schedule.  It does not build Cargo targets.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
import subprocess
import sys
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BATCH = "0700"
PROBE = f"{BATCH}-mce-oracle"
PROFILE = "baseline"
CPU = "12"
WARMUPS = 10
SAMPLES = 300
URI = b"http://schemas.openxmlformats.org/markup-compatibility/2006"
URI_TEXT = URI.decode("ascii")
URI_LEN = len(URI)
NEAR_PREFIX_REPETITIONS = 4096
LATE_TEXT_BYTES = 131072

CASE_MANIFEST = "marker-control-cases.json"
RUN_MANIFEST = "marker-control-runs.json"
SUMMARY = "marker-control-summary.json"
COMPARISONS = "marker-control-comparisons.json"
TRIGGERS = "marker-control-triggers.json"
CASE_ROOT = P / "marker-controls" / "cases"
RUN_ROOT = P / "marker-controls" / "runs"

GENERATION_COMMAND = [
    "python3",
    "docs/performance/results/change-0700/measure-marker-controls.py",
    "--generate-only",
]
CENSUS_COMMAND = [
    "python3",
    "docs/performance/results/change-0700/measure-marker-controls.py",
    "--census-only",
]
LEGS = (
    ("a0", "baseline"),
    ("a1", "baseline"),
    ("a2", "baseline"),
    ("b0", "candidate"),
    ("b1", "candidate"),
    ("a3", "baseline"),
)
PAIRS = (("a1", "a0"), ("b0", "a2"), ("b1", "a3"), ("a3", "a2"))
METRICS = ("p50_ns", "mean_ns", "p95_ns", "p99_ns")


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def rel(path: Path) -> str:
    return str(path.relative_to(P))


def occurrences(value: bytes, needle: bytes) -> list[int]:
    positions: list[int] = []
    start = 0
    while True:
        position = value.find(needle, start)
        if position < 0:
            return positions
        positions.append(position)
        start = position + 1


def case_bytes() -> list[tuple[str, str, bytes, dict[str, Any]]]:
    """Return the complete deterministic case census and its predicates."""

    tiny_free = b"<r/>"
    tiny_marked = b'<r xmlns:mc="' + URI + b'"/>'
    near = URI[:-1] + b"X"
    near_prefix = b"<r>" + b" ".join([near] * NEAR_PREFIX_REPETITIONS) + b"</r>"
    late_comment = (
        b"<r>"
        + (b"x" * LATE_TEXT_BYTES)
        + b"<!--"
        + URI
        + b"-->"
        + b"</r>"
    )
    root_hit = b"<r><!--" + URI + b"--><n/></r>"
    return [
        (
            "tiny-marker-free",
            "tiny-free",
            tiny_free,
            {"description": "valid marker-free XML shorter than the 59-byte URI"},
        ),
        (
            "tiny-valid-marked",
            "tiny-marked",
            tiny_marked,
            {"description": "valid XML with the marker in a root namespace declaration"},
        ),
        (
            "large-near-prefix-marker-free",
            "near-prefix",
            near_prefix,
            {
                "description": "large valid text made from repeated 58-byte URI prefixes",
                "near_prefix_repetitions": NEAR_PREFIX_REPETITIONS,
            },
        ),
        (
            "long-late-comment-hit",
            "late-comment",
            late_comment,
            {
                "description": "long valid XML whose only exact URI is in the trailing comment",
                "late_text_bytes": LATE_TEXT_BYTES,
            },
        ),
        (
            "root-comment-hit",
            "root-hit",
            root_hit,
            {"description": "valid XML with the exact URI at the root comment"},
        ),
    ]


def census_case(name: str, kind: str, value: bytes, extra: dict[str, Any]) -> dict[str, Any]:
    positions = occurrences(value, URI)
    return {
        "case": name,
        "kind": kind,
        "file": rel(CASE_ROOT / f"{name}.xml"),
        "bytes": len(value),
        "sha256": sha_bytes(value),
        "marker_uri": URI_TEXT,
        "marker_bytes": URI_LEN,
        "marker_count": len(positions),
        "marker_positions": positions,
        **extra,
    }


def write_case_manifest(cases: list[dict[str, Any]]) -> None:
    manifest = {
        "schema": "0700-marker-control-cases-v1",
        "probe": PROBE,
        "profile": PROFILE,
        "marker_uri": URI_TEXT,
        "marker_bytes": URI_LEN,
        "case_order": [case["case"] for case in cases],
        "generation_command": GENERATION_COMMAND,
        "census_command": CENSUS_COMMAND,
        "case_root": rel(CASE_ROOT),
        "cases": cases,
    }
    (P / CASE_MANIFEST).write_text(json.dumps(manifest, indent=2) + "\n")


def generate_cases() -> list[dict[str, Any]]:
    CASE_ROOT.mkdir(parents=True, exist_ok=True)
    cases: list[dict[str, Any]] = []
    for name, kind, value, extra in case_bytes():
        path = CASE_ROOT / f"{name}.xml"
        path.write_bytes(value)
        cases.append(census_case(name, kind, value, extra))
    write_case_manifest(cases)
    return cases


def verify_generated_cases(cases: list[dict[str, Any]]) -> None:
    expected = case_bytes()
    if [case["case"] for case in cases] != [name for name, *_ in expected]:
        raise AssertionError("marker-control case order changed")
    for (name, kind, value, extra), case in zip(expected, cases):
        if case != census_case(name, kind, value, extra):
            raise AssertionError(f"marker-control census mismatch: {name}")
        path = CASE_ROOT / f"{name}.xml"
        if path.read_bytes() != value:
            raise AssertionError(f"marker-control source changed: {name}")


def parse_identity(raw: bytes, source: Path) -> str:
    text = raw.decode("utf-8")
    if not text.endswith("\n") or len(text.splitlines()) != 1:
        raise AssertionError(f"oracle identity is not one newline-terminated line: {source}")
    fields = text.rstrip("\n").split("\t")
    if fields[0] not in {"OK", "ERR"}:
        raise AssertionError(f"oracle identity record kind is invalid: {source}")
    parsed: dict[str, str] = {}
    for field in fields[1:]:
        key, separator, value = field.partition("=")
        if not separator or key in parsed:
            raise AssertionError(f"oracle identity field is malformed: {source}")
        parsed[key] = value
    if parsed.get("probe") != PROBE or parsed.get("profile") != PROFILE:
        raise AssertionError(f"oracle identity probe/profile mismatch: {source}")
    if parsed.get("input_len") != str(source.stat().st_size):
        raise AssertionError(f"oracle identity input length mismatch: {source}")
    return text


def parse_timing(raw: bytes, source: Path) -> list[int]:
    text = raw.decode("utf-8")
    lines = text.splitlines()
    if not text.endswith("\n") or len(lines) != SAMPLES + 1:
        raise AssertionError(f"oracle timing line census mismatch: {source}")
    header = lines[0].split("\t")
    if header != [
        "TIMING",
        f"probe={PROBE}",
        f"profile={PROFILE}",
        f"input_len={source.stat().st_size}",
        f"warmups={WARMUPS}",
        f"samples={SAMPLES}",
    ]:
        raise AssertionError(f"oracle timing header mismatch: {source}")
    values: list[int] = []
    for index, line in enumerate(lines[1:]):
        fields = line.split("\t")
        if len(fields) != 3 or fields[0] != "SAMPLE":
            raise AssertionError(f"oracle timing sample shape mismatch: {source}")
        if fields[1] != f"index={index}":
            raise AssertionError(f"oracle timing sample index mismatch: {source}/{index}")
        prefix = "elapsed_ns="
        if not fields[2].startswith(prefix) or not fields[2][len(prefix) :].isdigit():
            raise AssertionError(f"oracle timing sample duration mismatch: {source}/{index}")
        value = int(fields[2][len(prefix) :])
        if value <= 0:
            raise AssertionError(f"oracle timing sample duration is not positive: {source}/{index}")
        values.append(value)
    return values


def stats(values: list[int]) -> dict[str, float | int]:
    ordered = sorted(values)
    return {
        "samples": len(values),
        "p50_ns": statistics.median(values),
        "mean_ns": statistics.mean(values),
        "p95_ns": ordered[max(0, math.ceil(len(values) * 0.95) - 1)],
        "p99_ns": ordered[max(0, math.ceil(len(values) * 0.99) - 1)],
        "min_ns": min(values),
        "max_ns": max(values),
    }


def build_row(
    case: dict[str, Any],
    leg: str,
    phase: str,
    binary: Path,
    source: Path,
    identity: str,
    identity_path: Path,
    identity_error_path: Path,
    timing_path: Path,
    timing_error_path: Path,
    values: list[int],
    identity_command: list[str],
    timing_command: list[str],
    identity_exit_code: int,
    timing_exit_code: int,
) -> dict[str, Any]:
    return {
        "schema": "0700-marker-control-run-v1",
        "kind": "oracle-marker-control",
        "case": case["case"],
        "leg": leg,
        "phase": phase,
        "profile": PROFILE,
        "cpu": int(CPU),
        "warmups": WARMUPS,
        "samples": SAMPLES,
        "source": rel(source),
        "source_sha256": sha(source),
        "binary": str(binary),
        "binary_sha256": sha(binary),
        "identity_command": identity_command,
        "identity_exit_code": identity_exit_code,
        "identity_stdout": rel(identity_path),
        "identity_stdout_sha256": sha(identity_path),
        "identity_stderr": rel(identity_error_path),
        "identity_stderr_sha256": sha(identity_error_path),
        "identity": identity,
        "command": timing_command,
        "exit_code": timing_exit_code,
        "output": rel(timing_path),
        "output_sha256": sha(timing_path),
        "stderr": rel(timing_error_path),
        "stderr_sha256": sha(timing_error_path),
        "stats": stats(values),
    }


def compare_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_key = {(row["case"], row["leg"]): row for row in rows}
    result: list[dict[str, Any]] = []
    for case in [case["case"] for case in generate_case_descriptors()]:
        for candidate_leg, baseline_leg in PAIRS:
            candidate = by_key[case, candidate_leg]
            baseline = by_key[case, baseline_leg]
            delta_pct = {
                metric: (candidate["stats"][metric] / baseline["stats"][metric] - 1) * 100
                for metric in METRICS
            }
            delta_ns = {
                metric: candidate["stats"][metric] - baseline["stats"][metric]
                for metric in METRICS
            }
            result.append(
                {
                    "schema": "0700-marker-control-comparison-v1",
                    "case": case,
                    "profile": PROFILE,
                    "pair": f"{candidate_leg}/{baseline_leg}",
                    "candidate_leg": candidate_leg,
                    "baseline_leg": baseline_leg,
                    "candidate": {metric: candidate["stats"][metric] for metric in METRICS},
                    "baseline": {metric: baseline["stats"][metric] for metric in METRICS},
                    "delta_pct": delta_pct,
                    "delta_ns": delta_ns,
                }
            )
    return result


def generate_case_descriptors() -> list[dict[str, Any]]:
    """Read the manifest's case descriptors in the stable execution order."""
    manifest = json.loads((P / CASE_MANIFEST).read_text())
    return manifest["cases"]


def write_derived(rows: list[dict[str, Any]]) -> None:
    summary = {
        "schema": "0700-marker-control-summary-v1",
        "probe": PROBE,
        "profile": PROFILE,
        "cpu": int(CPU),
        "warmups": WARMUPS,
        "samples": SAMPLES,
        "rows": [
            {
                "case": row["case"],
                "leg": row["leg"],
                "phase": row["phase"],
                "binary_sha256": row["binary_sha256"],
                "source_sha256": row["source_sha256"],
                "stats": row["stats"],
            }
            for row in rows
        ],
    }
    comparisons = compare_rows(rows)
    triggers = [
        {
            "case": row["case"],
            "profile": row["profile"],
            "pair": row["pair"],
            "metric": metric,
            "delta_pct": row["delta_pct"][metric],
            "delta_ns": row["delta_ns"][metric],
        }
        for row in comparisons
        if row["pair"] in {"b0/a2", "b1/a3"}
        for metric in METRICS
        if row["delta_pct"][metric] > 5
    ]
    (P / SUMMARY).write_text(json.dumps(summary, indent=2) + "\n")
    (P / COMPARISONS).write_text(json.dumps(comparisons, indent=2) + "\n")
    (P / TRIGGERS).write_text(json.dumps(triggers, indent=2) + "\n")


def measure() -> None:
    cases = generate_cases()
    verify_generated_cases(cases)
    RUN_ROOT.mkdir(parents=True, exist_ok=True)
    rows: list[dict[str, Any]] = []
    (P / RUN_MANIFEST).write_text("[]\n")
    for leg, phase in LEGS:
        binary = ROOT.parent / f"litchi-{BATCH}-bin" / f"{phase}-oracle"
        if not binary.exists():
            raise FileNotFoundError(binary)
        for case in cases:
            source = P / case["file"]
            stem = f'{case["case"]}-{leg}'
            identity_path = RUN_ROOT / f"{stem}.identity"
            identity_error_path = RUN_ROOT / f"{stem}.identity.stderr"
            timing_path = RUN_ROOT / f"{stem}.tsv"
            timing_error_path = RUN_ROOT / f"{stem}.stderr"
            identity_command = [str(binary), "run", PROFILE, str(source)]
            identity_result = subprocess.run(identity_command, capture_output=True)
            identity_path.write_bytes(identity_result.stdout)
            identity_error_path.write_bytes(identity_result.stderr)
            if identity_result.returncode != 0 or identity_result.stderr:
                raise RuntimeError(f"oracle identity failed: {identity_command}")
            identity = parse_identity(identity_result.stdout, source)
            timing_command = [
                "taskset",
                "-c",
                CPU,
                str(binary),
                "time",
                PROFILE,
                str(source),
                str(WARMUPS),
                str(SAMPLES),
            ]
            timing_result = subprocess.run(timing_command, capture_output=True)
            timing_path.write_bytes(timing_result.stdout)
            timing_error_path.write_bytes(timing_result.stderr)
            if timing_result.returncode != 0 or timing_result.stderr:
                raise RuntimeError(f"oracle timing failed: {timing_command}")
            values = parse_timing(timing_result.stdout, source)
            row = build_row(
                case,
                leg,
                phase,
                binary,
                source,
                identity,
                identity_path,
                identity_error_path,
                timing_path,
                timing_error_path,
                values,
                identity_command,
                timing_command,
                identity_result.returncode,
                timing_result.returncode,
            )
            rows.append(row)
            (P / RUN_MANIFEST).write_text(json.dumps(rows, indent=2) + "\n")
            print(case["case"], leg, flush=True)
    identities: dict[str, set[str]] = {}
    for row in rows:
        identities.setdefault(row["case"], set()).add(row["identity"])
    if any(len(values) != 1 for values in identities.values()):
        raise AssertionError("baseline/candidate marker-control identity changed")
    write_derived(rows)


def main() -> None:
    if len(sys.argv) == 2 and sys.argv[1] == "--generate-only":
        cases = generate_cases()
        verify_generated_cases(cases)
        print(f"generated and verified {len(cases)} marker-control cases")
        return
    if len(sys.argv) == 2 and sys.argv[1] == "--census-only":
        manifest = json.loads((P / CASE_MANIFEST).read_text())
        cases = manifest["cases"]
        verify_generated_cases(cases)
        print(f"verified {len(cases)} marker-control cases")
        return
    if len(sys.argv) != 1:
        raise SystemExit("usage: measure-marker-controls.py [--generate-only|--census-only]")
    measure()


if __name__ == "__main__":
    main()
