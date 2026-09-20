#!/usr/bin/env python3
"""Compare or capture the source-bound TRACE0721 DOCX diagnostics.

The compare action consumes two public-oracle reports and the stderr files
emitted by the temporary tracer. It checks public output identity, document
boundary identity, ordered MCE input/output vectors, and structured chunk/range
metadata. Missing versus empty early-failure scopes normalize to the same empty
value. Completion flags are retained in the emitted summary but are not a
cross-lane requirement.

The capture action only runs an already-built oracle binary. It does not build,
edit, restore, or clean a checkout.
"""

from __future__ import annotations

import argparse
from collections import Counter
import hashlib
import json
from pathlib import Path
import subprocess
import sys
import time
from typing import Any


TRACE_PREFIX = "TRACE0721 "
CASE_PREFIX = "CASE0721 "


class TraceError(RuntimeError):
    """A malformed diagnostic or semantic mismatch."""


def fail(message: str) -> None:
    raise TraceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def canonical(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest(value: Any) -> str:
    return hashlib.sha256(canonical(value)).hexdigest()


def write_new(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def as_int(value: Any, label: str) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"{label} is not an integer")
    return value


def as_list(value: Any, label: str) -> list[Any]:
    require(isinstance(value, list), f"{label} is not a list")
    return value


def parse_trace(path: Path) -> dict[str, Any]:
    """Parse TRACE lines while retaining the CASE label that preceded a scope."""
    require(path.is_file() and not path.is_symlink(), f"missing trace stderr: {path}")
    current_case: str | None = None
    starts: dict[int, dict[str, Any]] = {}
    documents: list[dict[str, Any]] = []
    try:
        lines = path.read_text(encoding="utf-8", errors="strict").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"cannot read trace stderr {path}: {error}")

    for line_number, line in enumerate(lines, start=1):
        if line.startswith(CASE_PREFIX):
            current_case = line[len(CASE_PREFIX):]
            continue
        if not line.startswith(TRACE_PREFIX):
            continue
        raw = line[len(TRACE_PREFIX):]
        try:
            value = json.loads(raw)
        except json.JSONDecodeError as error:
            fail(f"{path}:{line_number}: malformed TRACE0721 JSON: {error}")
        require(isinstance(value, dict), f"{path}:{line_number}: trace record is not an object")
        kind = value.get("kind")
        sequence = as_int(value.get("sequence"), f"{path}:{line_number}.sequence")
        if kind == "document_start":
            require(sequence not in starts, f"{path}:{line_number}: duplicate document sequence {sequence}")
            starts[sequence] = {"case": current_case, "start": value, "line": line_number}
        elif kind == "document_end":
            info = starts.pop(sequence, None)
            require(info is not None, f"{path}:{line_number}: end without start for sequence {sequence}")
            start = info["start"]
            for field in ("xml_bytes", "xml_sha256"):
                if field in start and field in value:
                    require(
                        start[field] == value[field],
                        f"{path}: sequence {sequence} boundary field {field} changed",
                    )
            documents.append({
                "case": info["case"],
                "sequence": sequence,
                "start": start,
                "end": value,
            })
        else:
            fail(f"{path}:{line_number}: unknown TRACE0721 kind {kind!r}")

    require(not starts, f"{path}: unterminated TRACE0721 document scopes: {sorted(starts)}")
    occurrences: Counter[str] = Counter()
    for document in documents:
        key = document["case"] if document["case"] is not None else "<uncased>"
        document["occurrence"] = occurrences[key]
        occurrences[key] += 1
    return {
        "path": str(path),
        "sha256": sha(path),
        "documents": documents,
        "case_count": len(occurrences),
        "line_count": len(lines),
    }


def document_key(document: dict[str, Any]) -> tuple[str, int]:
    case = document["case"] if document["case"] is not None else "<uncased>"
    return case, as_int(document["occurrence"], f"{case}.occurrence")


def validate_boundary(document: dict[str, Any], lane: str) -> dict[str, Any]:
    start = document["start"]
    end = document["end"]
    bytes_start = as_int(start.get("xml_bytes"), f"{lane}.start.xml_bytes")
    bytes_end = as_int(end.get("xml_bytes"), f"{lane}.end.xml_bytes")
    sha_start = start.get("xml_sha256")
    sha_end = end.get("xml_sha256")
    require(isinstance(sha_start, str) and len(sha_start) == 64,
            f"{lane}: malformed start XML SHA")
    require(isinstance(sha_end, str) and len(sha_end) == 64,
            f"{lane}: malformed end XML SHA")
    require(bytes_start == bytes_end and sha_start == sha_end,
            f"{lane}: document boundary identity changed inside one scope")
    return {"xml_bytes": bytes_start, "xml_sha256": sha_start}


def normalized_scope_values(scans: Any, field: str, label: str) -> list[Any]:
    """Treat absent scopes and scopes with only empty values as empty output."""
    if scans is None:
        return []
    scans = as_list(scans, f"{label}.{field}.scans")
    nonempty: list[list[Any]] = []
    for index, scan in enumerate(scans):
        require(isinstance(scan, dict), f"{label}.{field}[{index}] is not an object")
        value = scan.get(field)
        if value is None:
            continue
        value = as_list(value, f"{label}.{field}[{index}]")
        if value:
            nonempty.append(value)
    flattened: list[Any] = []
    for value in nonempty:
        flattened.extend(value)
    return flattened


def normalized_active_calls(end: dict[str, Any], label: str) -> list[dict[str, Any]]:
    calls = as_list(end.get("active_calls", []), f"{label}.active_calls")
    result: list[dict[str, Any]] = []
    for index, call in enumerate(calls):
        require(isinstance(call, dict), f"{label}.active_calls[{index}] is not an object")
        inputs = as_list(call.get("input_offsets"), f"{label}.active_calls[{index}].input_offsets")
        output = call.get("output_offsets")
        if output is not None:
            output = as_list(output, f"{label}.active_calls[{index}].output_offsets")
        result.append({
            "stage": call.get("stage"),
            "xml_bytes": call.get("xml_bytes"),
            "xml_sha256": call.get("xml_sha256"),
            "input_offsets": inputs,
            "output_offsets": output,
            "result_kind": call.get("result_kind"),
            # Error detail is useful for refusal parity; successful Debug text
            # is an implementation rendering and is intentionally excluded.
            "result_debug": call.get("result_debug") if call.get("result_kind") != "ok" else None,
        })
    return result


def normalized_document(document: dict[str, Any], lane: str) -> dict[str, Any]:
    end = document["end"]
    label = f"{lane}:{document_key(document)}"
    boundary = validate_boundary(document, label)
    alt = end.get("alt_scans", [])
    ranges = end.get("range_scans", [])
    reads = as_list(end.get("reader_reads", []), f"{label}.reader_reads")
    owner_counts: Counter[str] = Counter()
    for index, read in enumerate(reads):
        require(isinstance(read, dict), f"{label}.reader_reads[{index}] is not an object")
        owner = read.get("owner")
        require(owner in {"alt", "range"}, f"{label}: unknown reader owner {owner!r}")
        as_int(read.get("start"), f"{label}.reader_reads[{index}].start")
        as_int(read.get("end"), f"{label}.reader_reads[{index}].end")
        owner_counts[str(owner)] += 1

    range_scans = as_list(ranges, f"{label}.range_scans")
    observer_events = 0
    for index, scan in enumerate(range_scans):
        require(isinstance(scan, dict), f"{label}.range_scans[{index}] is not an object")
        count = as_int(scan.get("observer_events", 0),
                       f"{label}.range_scans[{index}].observer_events")
        require(count >= 0, f"{label}.range_scans[{index}].observer_events is negative")
        observer_events += count

    raw_chunks = normalized_scope_values(alt, "raw_chunks", label)
    selected_chunks = normalized_scope_values(alt, "selected_chunks", label)
    raw_ranges = normalized_scope_values(ranges, "raw_ranges", label)
    selected_ranges = normalized_scope_values(ranges, "selected_ranges", label)

    return {
        **boundary,
        "active_calls": normalized_active_calls(end, label),
        "raw_chunks": raw_chunks,
        "selected_chunks": selected_chunks,
        "raw_ranges": raw_ranges,
        "selected_ranges": selected_ranges,
        "reader_counts": {
            "alt": owner_counts["alt"],
            "range": owner_counts["range"],
            "total": len(reads),
        },
        "observer_events": observer_events,
        # Retain completion state for review, but no comparison requires it.
        "alt_completed": [
            bool(scan.get("completed")) for scan in as_list(alt, f"{label}.alt_scans")
            if isinstance(scan, dict)
        ],
        "range_completed": [
            bool(scan.get("completed")) for scan in range_scans
        ],
    }


def compare_document(
    baseline: dict[str, Any],
    candidate: dict[str, Any],
) -> dict[str, Any]:
    key = document_key(baseline)
    require(key == document_key(candidate), f"document key changed: {key} vs {document_key(candidate)}")
    left = normalized_document(baseline, "baseline")
    right = normalized_document(candidate, "candidate")

    semantic_fields = (
        "xml_bytes", "xml_sha256", "active_calls",
        "raw_chunks", "selected_chunks", "raw_ranges", "selected_ranges",
    )
    for field in semantic_fields:
        require(
            left[field] == right[field],
            f"{key}: {field} differs\nbaseline={left[field]!r}\ncandidate={right[field]!r}",
        )

    baseline_reads = left["reader_counts"]
    candidate_reads = right["reader_counts"]
    require(
        candidate_reads["range"] == 0,
        f"{key}: candidate performed {candidate_reads['range']} range reader calls",
    )
    if baseline_reads["range"] > 0:
        require(
            baseline_reads["alt"] == candidate_reads["alt"],
            f"{key}: fused alt read count changed: "
            f"baseline={baseline_reads['alt']} candidate={candidate_reads['alt']}",
        )
        require(
            candidate_reads["total"] == candidate_reads["alt"],
            f"{key}: candidate reader total is not the fused alt count",
        )

    # On a normal fused scope, the range observer sees the same event stream
    # that the single reader exposed. Early alternative-format errors may stop
    # before the observer, so this is a diagnostic relation rather than a
    # cross-lane completion requirement.
    observer_relation = {
        "baseline_range_reads": baseline_reads["range"],
        "candidate_range_reads": candidate_reads["range"],
        "baseline_observer_events": left["observer_events"],
        "candidate_observer_events": right["observer_events"],
        "candidate_alt_reads": candidate_reads["alt"],
        "candidate_observer_matches_alt_reads": (
            right["observer_events"] == candidate_reads["alt"]
        ),
    }
    return {
        "case": key[0],
        "occurrence": key[1],
        "boundary": {"xml_bytes": left["xml_bytes"], "xml_sha256": left["xml_sha256"]},
        "mce_call_count": len(left["active_calls"]),
        "reader_counts": {
            "baseline": baseline_reads,
            "candidate": candidate_reads,
        },
        "observer_relation": observer_relation,
        "completion_flags": {
            "baseline_alt": left["alt_completed"],
            "candidate_alt": right["alt_completed"],
            "baseline_range": left["range_completed"],
            "candidate_range": right["range_completed"],
        },
    }


def compare_trace_documents(
    baseline_trace: dict[str, Any],
    candidate_trace: dict[str, Any],
) -> list[dict[str, Any]]:
    left = {document_key(item): item for item in baseline_trace["documents"]}
    right = {document_key(item): item for item in candidate_trace["documents"]}
    require(len(left) == len(baseline_trace["documents"]), "baseline has duplicate trace document keys")
    require(len(right) == len(candidate_trace["documents"]), "candidate has duplicate trace document keys")
    require(set(left) == set(right),
            f"trace document keys differ: baseline-only={sorted(set(left)-set(right))}, "
            f"candidate-only={sorted(set(right)-set(left))}")
    return [compare_document(left[key], right[key]) for key in sorted(left)]


def capture(args: argparse.Namespace) -> int:
    binary = Path(args.binary)
    report = Path(args.report)
    stdout = Path(args.stdout)
    stderr = Path(args.stderr)
    require(binary.is_file() and not binary.is_symlink(), f"missing oracle binary: {binary}")
    for path in (report, stdout, stderr):
        require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    cwd = Path(args.cwd).resolve()
    require(cwd.is_dir(), f"missing capture cwd: {cwd}")
    started = time.monotonic()
    with stdout.open("wb") as out, stderr.open("wb") as err:
        completed = subprocess.run(
            [str(binary), "--output", str(report)],
            cwd=cwd,
            stdout=out,
            stderr=err,
            check=False,
        )
    receipt = {
        "schema": "litchi.docx-trace-capture-0721.v1",
        "command": [str(binary), "--output", str(report)],
        "cwd": str(cwd),
        "exit_code": completed.returncode,
        "elapsed_seconds": time.monotonic() - started,
        "report_sha256": sha(report) if report.is_file() else None,
        "stdout_sha256": sha(stdout),
        "stderr_sha256": sha(stderr),
    }
    if args.receipt is not None:
        write_new(Path(args.receipt), receipt)
    require(completed.returncode == 0, f"oracle exited with {completed.returncode}")
    print(json.dumps(receipt, sort_keys=True))
    return 0


def compare(args: argparse.Namespace) -> int:
    baseline_report = Path(args.baseline_report)
    candidate_report = Path(args.candidate_report)
    baseline_report_bytes = baseline_report.read_bytes()
    candidate_report_bytes = candidate_report.read_bytes()
    require(
        baseline_report_bytes == candidate_report_bytes,
        "public oracle reports are not byte-identical",
    )
    baseline_trace = parse_trace(Path(args.baseline_stderr))
    candidate_trace = parse_trace(Path(args.candidate_stderr))
    comparisons = compare_trace_documents(baseline_trace, candidate_trace)
    result = {
        "schema": "litchi.docx-trace-analysis-0721.v1",
        "status": "pass",
        "public_report_sha256": sha(baseline_report),
        "public_report_bytes": len(baseline_report_bytes),
        "baseline_trace_sha256": baseline_trace["sha256"],
        "candidate_trace_sha256": candidate_trace["sha256"],
        "document_count": len(comparisons),
        "documents": comparisons,
        "reader_proof": {
            "candidate_range_reader_documents": [
                item["case"] for item in comparisons
                if item["reader_counts"]["candidate"]["range"] != 0
            ],
            "baseline_total_reads": sum(
                item["reader_counts"]["baseline"]["total"] for item in comparisons
            ),
            "candidate_total_reads": sum(
                item["reader_counts"]["candidate"]["total"] for item in comparisons
            ),
            "baseline_observer_events": sum(
                item["observer_relation"]["baseline_observer_events"] for item in comparisons
            ),
            "candidate_observer_events": sum(
                item["observer_relation"]["candidate_observer_events"] for item in comparisons
            ),
        },
    }
    if args.output is None:
        print(json.dumps(result, indent=2, sort_keys=True))
    else:
        write_new(Path(args.output), result)
        print(f"wrote {args.output}")
    return 0


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="action", required=True)

    capture_parser = subparsers.add_parser("capture", help="run one already-built oracle")
    capture_parser.add_argument("--binary", required=True)
    capture_parser.add_argument("--report", required=True)
    capture_parser.add_argument("--stdout", required=True)
    capture_parser.add_argument("--stderr", required=True)
    capture_parser.add_argument("--cwd", default=".")
    capture_parser.add_argument("--receipt")
    capture_parser.set_defaults(handler=capture)

    compare_parser = subparsers.add_parser("compare", help="compare reports and trace stderr")
    compare_parser.add_argument("--baseline-report", required=True)
    compare_parser.add_argument("--candidate-report", required=True)
    compare_parser.add_argument("--baseline-stderr", required=True)
    compare_parser.add_argument("--candidate-stderr", required=True)
    compare_parser.add_argument("--output")
    compare_parser.set_defaults(handler=compare)

    args = parser.parse_args()
    try:
        return args.handler(args)
    except (OSError, TraceError) as error:
        print(f"trace-analyze.py: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())

