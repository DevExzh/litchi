#!/usr/bin/env python3
"""Independently summarize the retained whole-process DOCX CPU profile.

The capture uses ``perf record`` with a cycles event.  ``perf script`` prints a
period on every sample header; this parser weights each frame ancestry and leaf
by that period.  It keeps the target executable separate from the ``perf``
helper process, and compares the resulting period sum with ``perf report``.
Nothing in this file reruns the workload or invokes a profiler.
"""

from __future__ import annotations

import csv
import hashlib
import json
import re
from pathlib import Path
from typing import Any, Iterable


ROOT = Path(__file__).resolve().parent
SUMMARY_PATH = ROOT / "profile-summary.json"
SUMMARY_SCHEMA = "docx-plain-paragraph-tail-append-profile-summary-v1"
PERF_EVENT = "cpu/cycles/P"
STAT_EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-misses",
    "page-faults",
)

# ``perf`` truncates a process comm to 15 bytes.  The executable name is read
# from the retained binary identity rather than silently assuming this value.
HEADER_RE = re.compile(
    r"^(?P<comm>\S+)\s+(?P<pid>\d+)\s+(?P<timestamp>\S+):\s+"
    r"(?P<period>\d+)\s+(?P<event>cpu/cycles/P):\s*$",
    re.MULTILINE,
)
REPORT_ROW_RE = re.compile(
    r"^\s+(?P<percent>\d+(?:\.\d+)?)%\s+(?P<command>\S+)\s+"
    r"(?P<object>\S+)\s+\[[^\]]+\]\s+(?P<symbol>.+?)\s*$"
)
REPORT_EVENT_RE = re.compile(r"^# Event count \(approx\.\):\s*(\d+)\s*$", re.MULTILINE)
REPORT_LOST_RE = re.compile(r"^# Total Lost Samples:\s*(\d+)\s*$", re.MULTILINE)
PERF_STAT_COMMENT_RE = re.compile(r"^#")
STRACE_ROW_RE = re.compile(
    r"^\s*(?P<percent>\d+(?:\.\d+)?)\s+"
    r"(?P<seconds>\d+(?:\.\d+)?)\s+"
    r"(?P<usecs>\d+)\s+(?P<calls>\d+)\s+"
    r"(?:(?P<errors>\d+)\s+)?"
    r"(?P<syscall>\S+)\s*$"
)
STRACE_TOTAL_RE = re.compile(
    r"^\s*(?P<percent>\d+(?:\.\d+)?)\s+"
    r"(?P<seconds>\d+(?:\.\d+)?)\s+"
    r"(?P<usecs>\d+)\s+(?P<calls>\d+)\s+(?P<errors>\d+)\s+total\s*$"
)

MARKERS = {
    "lifecycle": "run_total_iteration",
    # The narrow marker below avoids attributing the semantic-oracle
    # transaction scanner to the source-backed paragraph scanner.  The
    # ``scan_document_any`` row remains as a diagnostic for the raw function
    # name because perf inlines direct calls without printing their module.
    "scan_document": "litchi_docx::source_backed::paragraph_copy::scan_document",
    "scan_document_any": "scan_document",
    "copy_plain_paragraph": "copy_plain_paragraph",
    "publish_plain_paragraph": "publish_plain_paragraph",
    "build_corpus": "build_corpus",
    "verify_with_policy": "verify_with_policy",
    "semantic_record": "semantic_record",
    "deflate_medium": "deflate_medium",
}


class AnalysisError(ValueError):
    """Retained profile evidence does not satisfy the parser contract."""


def fail(message: str) -> None:
    raise AnalysisError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=_pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"cannot read JSON {path}: {error}")


def _pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            fail(f"duplicate JSON key {key!r}")
        value[key] = item
    return value


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def metadata(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing regular file {path}")
    return {"bytes": path.stat().st_size, "sha256": sha256_file(path)}


def receipt_artifact(receipt_path: Path, artifact_name: str, path: Path,
                     root: Path) -> dict[str, Any]:
    receipt = read_json(receipt_path)
    require(isinstance(receipt, dict), f"{receipt_path}: receipt is not an object")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and artifact_name in artifacts,
            f"{receipt_path}: missing artifact {artifact_name}")
    expected = artifacts[artifact_name]
    require(isinstance(expected, dict), f"{receipt_path}: malformed artifact {artifact_name}")
    actual = metadata(path)
    require(expected == actual,
            f"{receipt_path}: artifact {artifact_name} metadata differs")
    return {"path": str(path.relative_to(root)), **actual}


def parse_frame(line: str) -> str:
    """Return a stable frame label while retaining the printed symbol text."""
    text = line.strip()
    require(text != "", "empty perf stack frame")
    # The first token is an address for normal perf output.  Keep the rest,
    # including the DSO annotation, because it is useful for review.  Remove
    # only the final DSO parenthesis from the leaf key later.
    fields = text.split(None, 1)
    require(len(fields) == 2, f"malformed perf stack frame: {line!r}")
    return text


def frame_symbol(frame: str) -> str:
    fields = frame.split(None, 1)
    symbol = fields[1] if len(fields) == 2 else frame
    # DSO annotations are not part of the symbol identity used for leaf
    # grouping.  The annotation can contain spaces, so strip only the final
    # parenthesized suffix when it is present.
    symbol = re.sub(r"\s+\([^\n]*\)$", "", symbol)
    return re.sub(r"\+0x[0-9a-fA-F]+$", "", symbol)


def marker_matches(name: str, marker: str, frames: list[str]) -> bool:
    if name != "scan_document":
        return any(marker in frame for frame in frames)
    if any(marker in frame for frame in frames):
        return True
    # Direct inlined frames are printed as ``scan_document+0x...``.  They are
    # source-backed when the same stack carries its paragraph-copy scanner;
    # the semantic transaction scanner carries the transaction module's full
    # closure name instead and is intentionally excluded here.
    return (
        any(re.search(r"\bscan_document\+0x[0-9a-fA-F]+", frame) for frame in frames)
        and any(
            "plain_paragraph_copy_snapshot_with_limits" in frame
            or "copy_plain_paragraph" in frame
            or "source_backed::paragraph_copy" in frame
            for frame in frames
        )
        and not any("document::transaction::scan_document" in frame for frame in frames)
    )


def parse_perf_script(path: Path, expected_comm: str) -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"cannot read perf script {path}: {error}")
    matches = list(HEADER_RE.finditer(text))
    require(matches, "perf script contains no cpu/cycles/P sample headers")
    samples: list[dict[str, Any]] = []
    for index, match in enumerate(matches):
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        body = text[match.end():end]
        frames = [parse_frame(line) for line in body.splitlines() if line.startswith("\t")]
        require(frames, f"perf sample {index} has no stack frames")
        samples.append({
            "comm": match.group("comm"),
            "pid": int(match.group("pid")),
            "timestamp": match.group("timestamp"),
            "period": int(match.group("period")),
            "event": match.group("event"),
            "frames": frames,
        })

    require(all(row["event"] == PERF_EVENT for row in samples),
            "perf script contains an unexpected event")
    target = [row for row in samples if row["comm"] == expected_comm]
    helper = [row for row in samples if row["comm"] != expected_comm]
    require(target, f"perf script contains no samples for {expected_comm}")
    require(all(row["period"] > 0 for row in samples), "perf sample period is not positive")

    target_period = sum(row["period"] for row in target)
    helper_period = sum(row["period"] for row in helper)
    all_period = target_period + helper_period

    marker_rows: dict[str, list[dict[str, Any]]] = {}
    weighted_markers: dict[str, dict[str, Any]] = {}
    for name, marker in MARKERS.items():
        matched = [row for row in target if marker_matches(name, marker, row["frames"])]
        marker_rows[name] = matched
        period = sum(row["period"] for row in matched)
        best = max(matched, key=lambda row: row["period"], default=None)
        weighted_markers[name] = {
            "marker": marker,
            "sample_blocks": len(matched),
            "period": period,
            "percent_of_target_period": percent(period, target_period),
            "representative_stack": best["frames"][:12] if best else [],
        }

    lifecycle_rows = marker_rows["lifecycle"]
    scan_rows = marker_rows["scan_document"]
    lifecycle_ids = {id(row) for row in lifecycle_rows}
    scan_in_lifecycle = [row for row in scan_rows if id(row) in lifecycle_ids]
    scan_outside_lifecycle = [row for row in scan_rows if id(row) not in lifecycle_ids]
    require(len(scan_in_lifecycle) + len(scan_outside_lifecycle) == len(scan_rows),
            "scan_document ancestry partition is incomplete")
    require(
        sum(row["period"] for row in scan_in_lifecycle)
        + sum(row["period"] for row in scan_outside_lifecycle)
        == weighted_markers["scan_document"]["period"],
        "scan_document ancestry period partition is incomplete",
    )

    def ancestry_row(rows: list[dict[str, Any]], marker: str) -> dict[str, Any]:
        period = sum(row["period"] for row in rows)
        best = max(rows, key=lambda row: row["period"], default=None)
        return {
            "marker": marker,
            "sample_blocks": len(rows),
            "period": period,
            "percent_of_target_period": percent(period, target_period),
            "percent_of_lifecycle_period": percent(period, weighted_markers["lifecycle"]["period"]),
            "representative_stack": best["frames"][:12] if best else [],
        }

    leaf_periods: dict[str, int] = {}
    for row in target:
        leaf = frame_symbol(row["frames"][0])
        leaf_periods[leaf] = leaf_periods.get(leaf, 0) + row["period"]
    leaves = [
        {
            "symbol": symbol,
            "period": period,
            "percent_of_target_period": percent(period, target_period),
        }
        for symbol, period in sorted(leaf_periods.items(), key=lambda item: (-item[1], item[0]))[:20]
    ]

    unresolved = [row for row in target if any("[unknown]" in frame for frame in row["frames"])]
    return {
        "event": PERF_EVENT,
        "sample_blocks": len(samples),
        "target_comm": expected_comm,
        "target_sample_blocks": len(target),
        "helper_sample_blocks": len(helper),
        "target_period": target_period,
        "helper_period": helper_period,
        "all_period": all_period,
        "target_percent_of_all_period": percent(target_period, all_period),
        "helper_percent_of_all_period": percent(helper_period, all_period),
        "unresolved_target_sample_blocks": len(unresolved),
        "unresolved_target_period": sum(row["period"] for row in unresolved),
        "unresolved_percent_of_target_period": percent(sum(row["period"] for row in unresolved), target_period),
        "target_non_lifecycle_sample_blocks": len(target) - len(lifecycle_rows),
        "target_non_lifecycle_period": target_period - weighted_markers["lifecycle"]["period"],
        "target_non_lifecycle_percent": percent(
            target_period - weighted_markers["lifecycle"]["period"], target_period
        ),
        "weighted_callstack_markers": weighted_markers,
        "scan_document_partition": {
            "source_backed_scan_in_lifecycle": ancestry_row(
                scan_in_lifecycle,
                "litchi_docx::source_backed::paragraph_copy::scan_document within run_total_iteration",
            ),
            "source_backed_scan_outside_lifecycle": ancestry_row(
                scan_outside_lifecycle,
                "litchi_docx::source_backed::paragraph_copy::scan_document outside run_total_iteration",
            ),
        },
        "top_leaf_symbols": leaves,
    }


def percent(numerator: int, denominator: int) -> float:
    require(denominator > 0, "percentage denominator is not positive")
    return numerator * 100.0 / denominator


def parse_report(path: Path) -> dict[str, Any]:
    try:
        text = path.read_text(encoding="utf-8")
    except (OSError, UnicodeError) as error:
        fail(f"cannot read perf report {path}: {error}")
    event_match = REPORT_EVENT_RE.search(text)
    lost_match = REPORT_LOST_RE.search(text)
    require(event_match is not None, "perf report event count is missing")
    require(lost_match is not None, "perf report lost sample count is missing")
    rows = []
    for line in text.splitlines():
        match = REPORT_ROW_RE.match(line)
        if match:
            rows.append({
                "reported_percent": float(match.group("percent")),
                "command": match.group("command"),
                "shared_object": match.group("object"),
                "symbol": match.group("symbol"),
            })
    target = [row for row in rows if row["command"] == "docx_plain_para"]
    require(target, "perf report has no target rows")
    return {
        "event": PERF_EVENT,
        "event_count_approx": int(event_match.group(1)),
        "total_lost_samples": int(lost_match.group(1)),
        "top_level_row_count": len(rows),
        "top_level_reported_percent_sum": sum(row["reported_percent"] for row in rows),
        "target_top_level_row_count": len(target),
        "target_top_rows": target[:20],
    }


def parse_stat(path: Path) -> dict[str, Any]:
    values: dict[str, dict[str, Any]] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"cannot read perf stat {path}: {error}")
    for row in csv.reader(line for line in lines if not PERF_STAT_COMMENT_RE.match(line)):
        if len(row) < 5 or row[2] not in STAT_EVENTS:
            continue
        require(row[0].isdigit(), f"perf stat {row[2]} is not numeric: {row[0]!r}")
        require(row[3].isdigit(), f"perf stat {row[2]} runtime field is not numeric")
        try:
            running_percent = float(row[4])
        except ValueError:
            fail(f"perf stat {row[2]} running percentage is not numeric")
        values[row[2]] = {
            "value": int(row[0]),
            "unit": row[1],
            "runtime_field": int(row[3]),
            "running_percent": running_percent,
            "metric": row[5] if len(row) > 5 else "",
            "metric_unit": row[6] if len(row) > 6 else "",
        }
    require(tuple(values) == STAT_EVENTS,
            f"perf stat event set differs: {tuple(values)}")
    cycles = values["cycles"]["value"]
    instructions = values["instructions"]["value"]
    branches = values["branches"]["value"]
    require(cycles > 0 and branches > 0, "perf stat denominator is zero")
    return {
        "events": values,
        "derived": {
            "instructions_per_cycle": instructions / cycles,
            "branch_misses_percent_of_branches": (
                values["branch-misses"]["value"] * 100.0 / branches
            ),
            "cache_misses_per_cycle": values["cache-misses"]["value"] / cycles,
        },
        "scope": "separate perf stat whole-process run including corpus, warmups, lifecycles, report serialization and teardown",
    }


def parse_strace(path: Path) -> dict[str, Any]:
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except (OSError, UnicodeError) as error:
        fail(f"cannot read strace {path}: {error}")
    rows = []
    total = None
    for line in lines:
        match = STRACE_TOTAL_RE.match(line)
        if match:
            total = {
                "percent": float(match.group("percent")),
                "seconds": float(match.group("seconds")),
                "usecs_per_call": int(match.group("usecs")),
                "calls": int(match.group("calls")),
                "errors": int(match.group("errors")),
            }
            continue
        match = STRACE_ROW_RE.match(line)
        if match:
            rows.append({
                "syscall": match.group("syscall"),
                "percent": float(match.group("percent")),
                "seconds": float(match.group("seconds")),
                "usecs_per_call": int(match.group("usecs")),
                "calls": int(match.group("calls")),
                "errors": int(match.group("errors") or 0),
            })
    require(total is not None, "strace total row is missing")
    require(rows, "strace contains no syscall rows")
    require(sum(row["calls"] for row in rows) == total["calls"],
            "strace syscall call counts do not equal total")
    require(sum(row["errors"] for row in rows) == total["errors"],
            "strace syscall error counts do not equal total")
    rounded_seconds_error = abs(sum(row["seconds"] for row in rows) - total["seconds"])
    require(rounded_seconds_error <= max(0.000020, len(rows) * 0.0000005),
            "strace syscall seconds do not equal rounded total")
    return {
        "total": total,
        "rows": sorted(rows, key=lambda row: (-row["seconds"], row["syscall"])),
        "scope": "whole process including corpus, warmups, lifecycles, report JSON serialization and teardown; syscall time is not sink attribution",
    }


def build_summary(root: Path) -> dict[str, Any]:
    root = Path(root).resolve()
    binaries = read_json(root / "binaries.json")
    require(isinstance(binaries, dict) and isinstance(binaries.get("normal"), dict),
            "binaries.json normal identity is missing")
    binary_path = Path(binaries["normal"]["path"])
    expected_comm = binary_path.name[:15]
    require(expected_comm, "normal binary name is empty")

    script_path = root / "validation/profile-perf-script.stdout"
    report_path = root / "validation/profile-perf-report.stdout"
    stat_path = root / "profiles/perf-stat.csv"
    strace_path = root / "profiles/strace.txt"
    script_receipt = root / "validation/profile-perf-script.json"
    report_receipt = root / "validation/profile-perf-report.json"
    stat_receipt = root / "validation/profile-perf-stat.json"
    strace_receipt = root / "validation/profile-strace.json"
    for receipt in (script_receipt, report_receipt, stat_receipt, strace_receipt):
        value = read_json(receipt)
        require(value.get("exit_code") == 0, f"profile receipt failed: {receipt.name}")

    script = parse_perf_script(script_path, expected_comm)
    report = parse_report(report_path)
    require(report["event_count_approx"] == script["all_period"],
            "perf report event count differs from weighted perf script periods")
    require(report["total_lost_samples"] == 0, "perf report reports lost samples")
    stat = parse_stat(stat_path)
    strace = parse_strace(strace_path)

    input_artifacts = {
        "perf_script": receipt_artifact(script_receipt, "profile-perf-script.stdout", script_path, root),
        "perf_report": receipt_artifact(report_receipt, "profile-perf-report.stdout", report_path, root),
        "perf_stat": receipt_artifact(stat_receipt, "profile-perf-stat.stdout", root / "validation/profile-perf-stat.stdout", root),
        "strace": receipt_artifact(strace_receipt, "profile-strace.stdout", root / "validation/profile-strace.stdout", root),
        "perf_stat_csv": metadata(stat_path) | {"path": str(stat_path.relative_to(root))},
        "strace_txt": metadata(strace_path) | {"path": str(strace_path.relative_to(root))},
    }
    return {
        "schema": SUMMARY_SCHEMA,
        "status": "pass",
        "scope": "whole process including corpus generation, independent oracles, warmup, measured lifecycles, report JSON serialization and teardown",
        "binary": {
            "path": str(binary_path),
            "comm": expected_comm,
            "sha256": binaries["normal"].get("sha256"),
        },
        "inputs": input_artifacts,
        "perf_script_weighted": script,
        "perf_report_no_children": report,
        "perf_stat": stat,
        "strace": {
            "total": strace["total"],
            "top_rows": strace["rows"][:20],
            "row_count": len(strace["rows"]),
            "scope": strace["scope"],
        },
        "reconciliation": {
            "script_all_period_equals_report_event_count": True,
            "target_period_excludes_perf_helper": True,
            "perf_stat_is_separate_run": True,
            "perf_report_and_perf_stat_must_not_be_subtracted": True,
        },
    }


def derive(root: Path = ROOT) -> dict[str, Any]:
    """Derive the summary using only files rooted under ``root``.

    The absolute binary path retained in ``binaries.json`` is provenance only;
    this function never opens it.  That makes the same parser usable against a
    copied portable bundle after temporary executables have been removed.
    """
    return build_summary(Path(root))


def main() -> int:
    try:
        summary = derive(ROOT)
        with SUMMARY_PATH.open("x", encoding="utf-8") as stream:
            json.dump(summary, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except AnalysisError as error:
        print(f"profile-analysis.py: FAIL: {error}")
        return 1
    except OSError as error:
        print(f"profile-analysis.py: FAIL: {error}")
        return 1
    print("profile-analysis.py: PASS: wrote profile-summary.json")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
