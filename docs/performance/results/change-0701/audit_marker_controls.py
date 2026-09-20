#!/usr/bin/env python3
"""Independently audit the 0701 bounded marker-search control packet.

This module intentionally repeats the case construction and all raw-output
parsing instead of importing the measurement driver.  It can therefore catch
case-generation drift, receipt drift, stale binary/source bindings, malformed
sample indices, and summaries or >5% flags that were not recomputed from raw
oracle output.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
import sys
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BATCH = "0701"
PROBE = f"{BATCH}-mce-oracle"
PROFILE = "baseline"
CPU = 12
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
CASE_ROOT = "marker-controls/cases"
RUN_ROOT = "marker-controls/runs"
GENERATION_COMMAND = [
    "python3",
    "docs/performance/results/change-0701/measure-marker-controls.py",
    "--generate-only",
]
CENSUS_COMMAND = [
    "python3",
    "docs/performance/results/change-0701/measure-marker-controls.py",
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


def fail(message: str) -> None:
    raise AssertionError(message)


def need(packet: Path, relative: str) -> Path:
    path = packet / relative
    if not path.exists():
        fail(f"missing marker-control evidence: {relative}")
    return path


def read(packet: Path, relative: str) -> Any:
    return json.loads(need(packet, relative).read_text())


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def occurrences(value: bytes, needle: bytes) -> list[int]:
    positions: list[int] = []
    start = 0
    while True:
        position = value.find(needle, start)
        if position < 0:
            return positions
        positions.append(position)
        start = position + 1


def expected_cases() -> list[tuple[str, str, bytes, dict[str, Any]]]:
    """Independent copy of the bounded case construction."""

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


def descriptor(packet: Path, name: str, kind: str, value: bytes, extra: dict[str, Any]) -> dict[str, Any]:
    return {
        "case": name,
        "kind": kind,
        "file": f"{CASE_ROOT}/{name}.xml",
        "bytes": len(value),
        "sha256": sha_bytes(value),
        "marker_uri": URI_TEXT,
        "marker_bytes": URI_LEN,
        "marker_count": len(occurrences(value, URI)),
        "marker_positions": occurrences(value, URI),
        **extra,
    }


def validate_cases(packet: Path) -> list[dict[str, Any]]:
    manifest = read(packet, CASE_MANIFEST)
    required = {
        "schema",
        "probe",
        "profile",
        "marker_uri",
        "marker_bytes",
        "case_order",
        "generation_command",
        "census_command",
        "case_root",
        "cases",
    }
    if set(manifest) != required:
        fail("marker-control case manifest schema mismatch")
    if manifest["schema"] != "0701-marker-control-cases-v1" or manifest["probe"] != PROBE or manifest["profile"] != PROFILE:
        fail("marker-control case manifest identity mismatch")
    if manifest["marker_uri"] != URI_TEXT or manifest["marker_bytes"] != URI_LEN or manifest["case_root"] != CASE_ROOT:
        fail("marker-control case manifest marker metadata mismatch")
    if manifest["generation_command"] != GENERATION_COMMAND or manifest["census_command"] != CENSUS_COMMAND:
        fail("marker-control case generation/census command mismatch")
    expected = expected_cases()
    if manifest["case_order"] != [name for name, *_ in expected]:
        fail("marker-control case order mismatch")
    entries = manifest["cases"]
    if not isinstance(entries, list) or len(entries) != len(expected):
        fail("marker-control case entry census mismatch")
    for (name, kind, value, extra), entry in zip(expected, entries):
        if entry != descriptor(packet, name, kind, value, extra):
            fail(f"marker-control case descriptor mismatch: {name}")
        path = need(packet, entry["file"])
        if path.read_bytes() != value or sha(path) != entry["sha256"]:
            fail(f"marker-control case bytes/hash mismatch: {name}")
        if len(value) != entry["bytes"] or occurrences(value, URI) != entry["marker_positions"]:
            fail(f"marker-control case census mismatch: {name}")
        try:
            ET.fromstring(value)
        except ET.ParseError as error:
            fail(f"marker-control case is not valid XML: {name}: {error}")
        if name == "tiny-marker-free" and not (len(value) < URI_LEN and not occurrences(value, URI)):
            fail("tiny marker-free bound/predicate mismatch")
        if name == "tiny-valid-marked" and occurrences(value, URI) != [13]:
            fail("tiny marked URI position mismatch")
        if name == "large-near-prefix-marker-free":
            if occurrences(value, URI) or value.count(URI[:-1] + b"X") != NEAR_PREFIX_REPETITIONS:
                fail("near-prefix marker-free predicate mismatch")
        if name == "long-late-comment-hit":
            positions = occurrences(value, URI)
            if len(positions) != 1 or positions[0] <= len(value) // 2:
                fail("late-comment marker position mismatch")
            position = positions[0]
            if value[position - 4 : position] != b"<!--" or value[position + URI_LEN : position + URI_LEN + 3] != b"-->":
                fail("late-comment marker wrapper mismatch")
        if name == "root-comment-hit":
            positions = occurrences(value, URI)
            if len(positions) != 1 or positions[0] >= 32:
                fail("root marker position mismatch")
            position = positions[0]
            if value[position - 4 : position] != b"<!--" or value[position + URI_LEN : position + URI_LEN + 3] != b"-->":
                fail("root marker wrapper mismatch")
    return entries


def parse_fields(line: str, kind: str, label: str) -> dict[str, str]:
    fields = line.split("\t")
    if not fields or fields[0] != kind:
        fail(f"{label}: expected {kind} record")
    result: dict[str, str] = {}
    for field in fields[1:]:
        key, separator, value = field.partition("=")
        if not separator or not key or key in result:
            fail(f"{label}: malformed field")
        result[key] = value
    return result


def parse_identity(raw: bytes, source: Path, label: str) -> str:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"{label}: identity is not UTF-8: {error}")
    lines = text.splitlines()
    if len(lines) != 1 or not text.endswith("\n"):
        fail(f"{label}: identity is not one newline-terminated line")
    kind = lines[0].split("\t", 1)[0]
    if kind not in {"OK", "ERR"}:
        fail(f"{label}: identity kind mismatch")
    fields = parse_fields(lines[0], kind, label)
    if fields.get("probe") != PROBE or fields.get("profile") != PROFILE:
        fail(f"{label}: identity probe/profile mismatch")
    if fields.get("input_len") != str(source.stat().st_size):
        fail(f"{label}: identity input length mismatch")
    return text


def parse_timing(raw: bytes, source: Path, label: str) -> list[int]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"{label}: timing is not UTF-8: {error}")
    lines = text.splitlines()
    if len(lines) != SAMPLES + 1 or not text.endswith("\n"):
        fail(f"{label}: timing line census mismatch")
    header = parse_fields(lines[0], "TIMING", label)
    if header != {
        "probe": PROBE,
        "profile": PROFILE,
        "input_len": str(source.stat().st_size),
        "warmups": str(WARMUPS),
        "samples": str(SAMPLES),
    }:
        fail(f"{label}: timing header mismatch")
    values: list[int] = []
    for index, line in enumerate(lines[1:]):
        fields = line.split("\t")
        if len(fields) != 3 or fields[0] != "SAMPLE" or fields[1] != f"index={index}":
            fail(f"{label}: timing index/shape mismatch at {index}")
        prefix = "elapsed_ns="
        if not fields[2].startswith(prefix) or not fields[2][len(prefix) :].isdigit():
            fail(f"{label}: timing duration shape mismatch at {index}")
        elapsed = int(fields[2][len(prefix) :])
        if elapsed <= 0:
            fail(f"{label}: timing duration is not positive at {index}")
        values.append(elapsed)
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


def validate_runs(packet: Path, entries: list[dict[str, Any]]) -> list[dict[str, Any]]:
    rows = read(packet, RUN_MANIFEST)
    if not isinstance(rows, list) or len(rows) != len(entries) * len(LEGS):
        fail("marker-control run count mismatch")
    builds: dict[str, dict[str, Any]] = {}
    binaries: dict[str, Path] = {}
    for phase in ("baseline", "candidate"):
        build = read(packet, f"build-oracle-{phase}.json")
        if build.get("phase") != phase or build.get("exit_code") != 0:
            fail(f"marker-control oracle build receipt invalid: {phase}")
        binary = ROOT.parent / f"litchi-{BATCH}-bin" / f"{phase}-oracle"
        if binary.exists():
            if sha(binary) != build.get("binary_sha256"):
                fail(f"marker-control oracle binary binding invalid: {phase}")
        elif not (packet / "cleanup.json").is_file():
            fail(f"marker-control oracle binary is missing before cleanup: {phase}")
        else:
            cleanup = read(packet, "cleanup.json")
            removed = {row.get("path") for row in cleanup.get("removed", [])}
            if str(binary.parent) not in removed:
                fail(f"marker-control missing binary lacks cleanup binding: {phase}")
        builds[phase] = build
        binaries[phase] = binary
    expected_keys = {
        "schema",
        "kind",
        "case",
        "leg",
        "phase",
        "profile",
        "cpu",
        "warmups",
        "samples",
        "source",
        "source_sha256",
        "binary",
        "binary_sha256",
        "identity_command",
        "identity_exit_code",
        "identity_stdout",
        "identity_stdout_sha256",
        "identity_stderr",
        "identity_stderr_sha256",
        "identity",
        "command",
        "exit_code",
        "output",
        "output_sha256",
        "stderr",
        "stderr_sha256",
        "stats",
    }
    entry_by_name = {entry["case"]: entry for entry in entries}
    expected_order = [(leg, entry["case"]) for leg, _ in LEGS for entry in entries]
    if [(row.get("leg"), row.get("case")) for row in rows] != expected_order:
        fail("marker-control run order/coverage mismatch")
    identities: dict[str, bytes] = {}
    parsed_rows: list[dict[str, Any]] = []
    for row in rows:
        if set(row) != expected_keys:
            fail(f"marker-control run schema mismatch: {row.get('case')}/{row.get('leg')}")
        leg = row["leg"]
        phase = dict(LEGS).get(leg)
        case_name = row["case"]
        if row["schema"] != "0701-marker-control-run-v1" or row["kind"] != "oracle-marker-control" or phase is None:
            fail(f"marker-control run identity mismatch: {case_name}/{leg}")
        if row["phase"] != phase or row["profile"] != PROFILE or row["cpu"] != CPU or row["warmups"] != WARMUPS or row["samples"] != SAMPLES:
            fail(f"marker-control run metadata mismatch: {case_name}/{leg}")
        entry = entry_by_name[case_name]
        source = need(packet, entry["file"])
        binary = binaries[phase]
        expected_source = entry["file"]
        if row["source"] != expected_source or row["source_sha256"] != sha(source):
            fail(f"marker-control source binding mismatch: {case_name}/{leg}")
        if row["binary"] != str(binary) or row["binary_sha256"] != builds[phase]["binary_sha256"]:
            fail(f"marker-control binary binding mismatch: {case_name}/{leg}")
        stem = f"{case_name}-{leg}"
        identity_command = [str(binary), "run", PROFILE, str(source)]
        timing_command = [
            "taskset",
            "-c",
            str(CPU),
            str(binary),
            "time",
            PROFILE,
            str(source),
            str(WARMUPS),
            str(SAMPLES),
        ]
        if row["identity_command"] != identity_command or row["command"] != timing_command:
            fail(f"marker-control command binding mismatch: {case_name}/{leg}")
        if row["identity_exit_code"] != 0 or row["exit_code"] != 0:
            fail(f"marker-control exit status mismatch: {case_name}/{leg}")
        identity_path = need(packet, row["identity_stdout"])
        identity_error_path = need(packet, row["identity_stderr"])
        timing_path = need(packet, row["output"])
        timing_error_path = need(packet, row["stderr"])
        if row["identity_stdout"] != f"{RUN_ROOT}/{stem}.identity" or row["identity_stderr"] != f"{RUN_ROOT}/{stem}.identity.stderr":
            fail(f"marker-control identity path mismatch: {case_name}/{leg}")
        if row["output"] != f"{RUN_ROOT}/{stem}.tsv" or row["stderr"] != f"{RUN_ROOT}/{stem}.stderr":
            fail(f"marker-control timing path mismatch: {case_name}/{leg}")
        if row["identity_stdout_sha256"] != sha(identity_path) or row["identity_stderr_sha256"] != sha(identity_error_path):
            fail(f"marker-control identity raw hash mismatch: {case_name}/{leg}")
        if row["output_sha256"] != sha(timing_path) or row["stderr_sha256"] != sha(timing_error_path):
            fail(f"marker-control timing raw hash mismatch: {case_name}/{leg}")
        if identity_error_path.read_bytes() != b"" or timing_error_path.read_bytes() != b"":
            fail(f"marker-control emitted stderr: {case_name}/{leg}")
        identity_raw = identity_path.read_bytes()
        identity = parse_identity(identity_raw, source, f"{case_name}/{leg}")
        if row["identity"] != identity:
            fail(f"marker-control identity receipt mismatch: {case_name}/{leg}")
        if case_name in identities and identities[case_name] != identity_raw:
            fail(f"marker-control semantic identity changed: {case_name}")
        identities.setdefault(case_name, identity_raw)
        values = parse_timing(timing_path.read_bytes(), source, f"{case_name}/{leg}")
        expected_stats = stats(values)
        if row["stats"] != expected_stats:
            fail(f"marker-control raw stats mismatch: {case_name}/{leg}")
        parsed_rows.append({"row": row, "values": values, "stats": expected_stats})
    if set(identities) != set(entry_by_name):
        fail("marker-control identity coverage mismatch")
    return parsed_rows


def validate_derived(packet: Path, parsed_rows: list[dict[str, Any]], entries: list[dict[str, Any]]) -> None:
    summary = read(packet, SUMMARY)
    expected_summary = {
        "schema": "0701-marker-control-summary-v1",
        "probe": PROBE,
        "profile": PROFILE,
        "cpu": CPU,
        "warmups": WARMUPS,
        "samples": SAMPLES,
        "rows": [
            {
                "case": item["row"]["case"],
                "leg": item["row"]["leg"],
                "phase": item["row"]["phase"],
                "binary_sha256": item["row"]["binary_sha256"],
                "source_sha256": item["row"]["source_sha256"],
                "stats": item["stats"],
            }
            for item in parsed_rows
        ],
    }
    if summary != expected_summary:
        fail("marker-control summary is not a raw recomputation")
    by_key = {(item["row"]["case"], item["row"]["leg"]): item for item in parsed_rows}
    expected_comparisons: list[dict[str, Any]] = []
    for entry in entries:
        case = entry["case"]
        for candidate_leg, baseline_leg in PAIRS:
            candidate = by_key[case, candidate_leg]["stats"]
            baseline = by_key[case, baseline_leg]["stats"]
            delta_pct = {metric: (candidate[metric] / baseline[metric] - 1) * 100 for metric in METRICS}
            delta_ns = {metric: candidate[metric] - baseline[metric] for metric in METRICS}
            expected_comparisons.append(
                {
                    "schema": "0701-marker-control-comparison-v1",
                    "case": case,
                    "profile": PROFILE,
                    "pair": f"{candidate_leg}/{baseline_leg}",
                    "candidate_leg": candidate_leg,
                    "baseline_leg": baseline_leg,
                    "candidate": {metric: candidate[metric] for metric in METRICS},
                    "baseline": {metric: baseline[metric] for metric in METRICS},
                    "delta_pct": delta_pct,
                    "delta_ns": delta_ns,
                }
            )
    comparisons = read(packet, COMPARISONS)
    if comparisons != expected_comparisons:
        fail("marker-control comparisons are not a raw recomputation")
    expected_triggers = [
        {
            "case": comparison["case"],
            "profile": comparison["profile"],
            "pair": comparison["pair"],
            "metric": metric,
            "delta_pct": comparison["delta_pct"][metric],
            "delta_ns": comparison["delta_ns"][metric],
        }
        for comparison in expected_comparisons
        if comparison["pair"] in {"b0/a2", "b1/a3"}
        for metric in METRICS
        if comparison["delta_pct"][metric] > 5
    ]
    triggers = read(packet, TRIGGERS)
    if triggers != expected_triggers:
        fail("marker-control >5% flags are not a raw recomputation")
    for trigger in triggers:
        if trigger["delta_pct"] <= 5:
            fail("marker-control trigger is not strictly above 5%")


def audit_marker_controls(packet: Path | None = None) -> dict[str, Any]:
    packet = (packet or P).resolve()
    entries = validate_cases(packet)
    parsed_rows = validate_runs(packet, entries)
    validate_derived(packet, parsed_rows, entries)
    result = {
        "schema": "0701-marker-control-audit-v1",
        "cases": len(entries),
        "legs": len(LEGS),
        "runs": len(parsed_rows),
        "samples_per_run": SAMPLES,
        "profile": PROFILE,
        "cpu": CPU,
        "trigger_count": len(read(packet, TRIGGERS)),
    }
    return result


def main() -> None:
    packet = Path(sys.argv[1]) if len(sys.argv) == 2 else P
    if len(sys.argv) > 2:
        raise SystemExit("usage: audit_marker_controls.py [packet-directory]")
    result = audit_marker_controls(packet)
    print("PASS: marker-control cases, oracle identities, raw samples, statistics and >5% flags", flush=True)
    print(json.dumps(result, sort_keys=True), flush=True)


if __name__ == "__main__":
    main()
