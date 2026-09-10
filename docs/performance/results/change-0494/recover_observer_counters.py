#!/usr/bin/env python3
"""Derive corrected perf-stat and strace summaries from profiling-r1 raw files.

The original profile workload and observer processes are retained.  This helper
does not rerun either observer.  It corrects two parser-boundary mistakes in a
new immutable receipt: GNU perf's CSV field 3 is the running time in
nanoseconds and field 4 is the running percentage, while strace omits its
errors column for zero-error syscall rows.  Raw files and the original profile
summary remain the authoritative inputs and are bound by hash.
"""

from __future__ import annotations

import argparse
from collections import Counter
import datetime as _datetime
import json
from pathlib import Path
import sys
from typing import Any

sys.path.insert(0, str(Path(__file__).resolve().parent))
from support import ROOT, meta, sha, snapshot  # noqa: E402
import measure as canonical_measure  # noqa: E402
import recover_profile as recovery  # noqa: E402


SCHEMA = "docx-edit-provider-observer-correction-v1"
VERSION = 1
PROVIDERS = ("owned", "file-warm", "short-read")


class CorrectionError(RuntimeError):
    """A fail-closed raw observer correction error."""


def fail(message: str) -> None:
    raise CorrectionError(message)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")


def _write_new(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError:
        fail(f"refusing to replace immutable correction artifact: {path}")


def _file_meta(path: Path) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        fail(f"{path}: expected a regular non-symlink file")
    return {"path": str(path), **meta(path)}


def _numeric(value: str) -> int | float | None:
    value = value.strip().replace(" ", "")
    if not value or value.startswith("<") or value in {"-", "<notcounted>"}:
        return None
    try:
        return int(value)
    except ValueError:
        try:
            return float(value)
        except ValueError:
            return None


def parse_perf_stat_corrected(text: str) -> dict[str, Any]:
    """Parse perf ``-x,`` output with the GNU field positions preserved."""

    events: dict[str, dict[str, Any]] = {}
    raw_lines: list[str] = []
    ignored_comment_lines: list[str] = []
    unparsed_lines: list[str] = []
    for raw in text.splitlines():
        line = raw.rstrip("\r")
        if not line or line.lstrip().startswith("#"):
            if line.lstrip().startswith("#"):
                ignored_comment_lines.append(line)
            continue
        fields = [item.strip() for item in line.split(",")]
        if len(fields) < 3 or not fields[2]:
            unparsed_lines.append(line)
            continue
        event = fields[2]
        if any(char not in "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789_.-"
               for char in event):
            unparsed_lines.append(line)
            continue
        if event in events:
            fail(f"perf stat contains duplicate event {event!r}")
        running_time_ns = _numeric(fields[3]) if len(fields) > 3 else None
        running_percent = _numeric(fields[4]) if len(fields) > 4 else None
        if isinstance(running_percent, (int, float)) and not 0 <= running_percent <= 100:
            fail(f"perf event {event}: running percentage is outside 0..100")
        events[event] = {
            "value": _numeric(fields[0]),
            "value_text": fields[0],
            "unit": fields[1] or None,
            "event": event,
            "running_time_ns": running_time_ns,
            "running_percent": running_percent,
            "raw": line,
        }
        raw_lines.append(line)
    if unparsed_lines:
        fail(f"perf stat has unexplained rows: {unparsed_lines[:3]}")

    def value(name: str) -> int | float | None:
        item = events.get(name)
        return None if item is None else item["value"]

    cycles = value("cycles")
    instructions = value("instructions")
    branches = value("branches")
    branch_misses = value("branch-misses")
    derived: dict[str, float] = {}
    if isinstance(cycles, (int, float)) and cycles > 0 and isinstance(instructions, (int, float)):
        derived["ipc"] = instructions / cycles
    if isinstance(branches, (int, float)) and branches > 0 and isinstance(branch_misses, (int, float)):
        derived["branch_miss_rate"] = branch_misses / branches
    for numerator, denominator, name in (
        ("L1-dcache-load-misses", "L1-dcache-loads", "l1_dcache_load_miss_rate"),
        ("LLC-load-misses", "LLC-loads", "llc_load_miss_rate"),
    ):
        left, right = value(numerator), value(denominator)
        if isinstance(left, (int, float)) and isinstance(right, (int, float)) and right > 0:
            derived[name] = left / right
    return {
        "events": events,
        "derived": derived,
        "raw_lines": raw_lines,
        "ignored_comment_lines": ignored_comment_lines,
        "unparsed_lines": unparsed_lines,
        "field_contract": {
            "value": 0,
            "unit": 1,
            "event": 2,
            "running_time_ns": 3,
            "running_percent": 4,
        },
    }


def parse_strace_corrected(text: str) -> dict[str, Any]:
    """Parse strace summaries with an optional zero-error column."""

    syscalls: dict[str, dict[str, Any]] = {}
    ignored_header_lines: list[str] = []
    unparsed_lines: list[str] = []
    reported_total: dict[str, Any] | None = None
    for raw in text.splitlines():
        line = raw.rstrip("\r")
        stripped = line.strip()
        if not stripped:
            continue
        if stripped.startswith("-") or stripped.startswith("% time"):
            ignored_header_lines.append(line)
            continue
        fields = stripped.split()
        if fields[-1:] == ["total"]:
            if len(fields) != 6:
                unparsed_lines.append(line)
                continue
            reported_total = {
                "raw": line,
                "pct": _numeric(fields[0]),
                "seconds": _numeric(fields[1]),
                "usec": _numeric(fields[2]),
                "calls": _numeric(fields[3]),
                "errors": _numeric(fields[4]),
            }
            continue
        if len(fields) == 6:
            pct, seconds, usec, calls, errors, syscall = fields
            error_source = "explicit"
        elif len(fields) == 5:
            pct, seconds, usec, calls, syscall = fields
            errors = "0"
            error_source = "implicit-zero"
        else:
            unparsed_lines.append(line)
            continue
        if syscall == "total" or not syscall:
            unparsed_lines.append(line)
            continue
        syscalls[syscall] = {
            "pct": _numeric(pct),
            "seconds": _numeric(seconds),
            "usec": _numeric(usec),
            "calls": _numeric(calls),
            "errors": _numeric(errors),
            "error_source": error_source,
            "raw": line,
        }
    total_calls = sum(int(item["calls"]) for item in syscalls.values()
                      if isinstance(item.get("calls"), (int, float)))
    total_errors = sum(int(item["errors"]) for item in syscalls.values()
                       if isinstance(item.get("errors"), (int, float)))
    if reported_total is None:
        fail("strace summary has no total row")
    if unparsed_lines:
        fail(f"strace summary has unexplained rows: {unparsed_lines[:3]}")
    if total_calls != reported_total["calls"] or total_errors != reported_total["errors"]:
        fail(
            f"strace totals do not conserve: rows calls/errors {total_calls}/{total_errors}, "
            f"reported {reported_total['calls']}/{reported_total['errors']}"
        )
    return {
        "syscalls": syscalls,
        "total_calls": total_calls,
        "total_errors": total_errors,
        "reported_total": reported_total,
        "ignored_header_lines": ignored_header_lines,
        "unparsed_lines": unparsed_lines,
        "field_contract": {
            "with_errors": "pct seconds usec calls errors syscall",
            "without_errors": "pct seconds usec calls syscall; errors=0",
        },
    }


def classify_stack_corrected(frames: list[str]) -> tuple[str, str]:
    """Classify short and qualified symbols while retaining ambiguity."""

    qualified_run = any("docx_edit_provider::run_sample" in frame for frame in frames)
    bare_run = any(frame == "run_sample" for frame in frames)
    has_run_sample = qualified_run or bare_run
    has_publish = any("publish_docx_source_edit" in frame for frame in frames)
    has_prepare = any(
        frame == "prepare" or "docx_edit_provider::prepare" in frame for frame in frames
    )
    has_output_oracle = any(
        marker in frame
        for frame in frames
        for marker in ("verify_docx_source_edit_output", "sha256_hex")
    )
    if has_run_sample and has_publish:
        classification = "run_sample_publish_ancestor"
    elif has_run_sample and has_output_oracle:
        classification = "run_sample_output_oracle"
    elif has_run_sample:
        classification = "run_sample_setup_or_teardown"
    elif has_prepare:
        classification = "preflight"
    else:
        classification = "process_setup_or_unclassified"
    if bare_run and not qualified_run:
        qualification = "bare_run_sample_ambiguous"
    elif qualified_run:
        qualification = "qualified_run_sample"
    else:
        qualification = "no_run_sample_marker"
    return classification, qualification


def summarize_stacks_corrected(parsed: dict[str, Any]) -> dict[str, Any]:
    """Summarize the retained stack export with bare-symbol diagnostics."""

    samples = parsed.get("samples")
    if not isinstance(samples, list) or not samples:
        fail("retained perf script has no parsed samples")
    unparsed = [str(line) for line in parsed.get("unparsed_lines", [])]
    ignored_comment_lines = [line for line in unparsed if line.lstrip().startswith("#")]
    unparsed_payload_lines = [line for line in unparsed if not line.lstrip().startswith("#")]
    if unparsed_payload_lines:
        fail(f"perf script has unexplained payload lines: {unparsed_payload_lines[:3]}")
    classes: Counter[str] = Counter()
    class_samples: Counter[str] = Counter()
    class_leafs: dict[str, Counter[str]] = {}
    folded: Counter[str] = Counter()
    qualifications: Counter[str] = Counter()
    run_sample_period = 0
    publish_period = 0
    for sample in samples:
        frames = sample.get("frames")
        if not isinstance(frames, list) or not frames:
            fail("retained perf script contains an invalid frame list")
        period = int(sample.get("period", 1))
        if period < 1:
            fail("retained perf script contains a non-positive sample period")
        classification, qualification = classify_stack_corrected(frames)
        classes[classification] += period
        class_samples[classification] += 1
        class_leafs.setdefault(classification, Counter())[frames[0]] += period
        qualifications[qualification] += period
        folded[";".join(reversed(frames))] += period
        if classification.startswith("run_sample_"):
            run_sample_period += period
        if classification == "run_sample_publish_ancestor":
            publish_period += period
    total_period = sum(classes.values())
    class_rows = []
    for name in sorted(classes):
        class_rows.append({
            "class": name,
            "sample_count": class_samples[name],
            "period": classes[name],
            "share_of_all_period": classes[name] / total_period if total_period else None,
            "share_of_run_sample_period": (
                classes[name] / run_sample_period if run_sample_period else None
            ),
            "top_leaf_symbols": [
                {"symbol": symbol, "period": period}
                for symbol, period in sorted(
                    class_leafs[name].items(), key=lambda item: (-item[1], item[0])
                )[:20]
            ],
        })
    return {
        "schema": "docx-edit-provider-perf-stacks-corrected-v1",
        "sample_count": len(samples),
        "total_period": total_period,
        "run_sample_period": run_sample_period,
        "run_sample_publish_period": publish_period,
        "classes": class_rows,
        "qualification_periods": dict(sorted(qualifications.items())),
        "ambiguous_bare_run_sample": qualifications.get("bare_run_sample_ambiguous", 0) > 0,
        "ignored_comment_lines": ignored_comment_lines,
        "ignored_comment_count": len(ignored_comment_lines),
        "unparsed_lines": unparsed_payload_lines,
        "folded_stack_count": len(folded),
        "folded_stacks": [
            {"stack": stack, "period": period}
            for stack, period in sorted(folded.items(), key=lambda item: (-item[1], item[0]))
        ],
        "interpretation": (
            "The authenticated owned executable and 100-row report establish the selected "
            "harness context. The demangled export uses bare run_sample symbols, so those "
            "candidate phase classes remain explicitly ambiguous rather than being treated "
            "as qualified Rust ownership. Periods are statistical sample weights."
        ),
    }


def _provider_map(summary: dict[str, Any]) -> dict[str, dict[str, Any]]:
    rows = summary.get("providers")
    if not isinstance(rows, list):
        fail("profile summary provider list is missing")
    result = {row.get("provider"): row for row in rows if isinstance(row, dict)}
    if set(result) != set(PROVIDERS):
        fail("profile summary does not contain exactly the three warm profile providers")
    return result


def _recovery_helper_archive(path: Path) -> dict[str, Any]:
    """Bind recovery-r3 to its archived helper before the future helper fix."""

    custody_path = path / "helper-custody.json"
    custody = _json(custody_path)
    helpers = custody.get("helpers") if isinstance(custody, dict) else None
    if not isinstance(helpers, dict):
        fail(f"{custody_path}: recovery helper map is missing")
    archived: dict[str, dict[str, Any]] = {}
    current: dict[str, dict[str, Any]] = {}
    matches: dict[str, bool] = {}
    for name in ("recover_profile.py", "test_recover_profile.py"):
        expected = helpers.get(name)
        if not isinstance(expected, dict):
            fail(f"{custody_path}: {name} binding is missing")
        archived_meta = _file_meta(path / name)
        for key in ("bytes", "sha256"):
            if archived_meta.get(key) != expected.get(key):
                fail(f"{custody_path}: archived {name} differs")
        current_meta = _file_meta(ROOT / name)
        archived[name] = archived_meta
        current[name] = current_meta
        matches[name] = all(
            current_meta.get(key) == expected.get(key) for key in ("bytes", "sha256")
        )
    return {
        "custody": {"path": str(custody_path), **meta(custody_path)},
        "archived": archived,
        "current": current,
        "current_matches_archived": matches,
        "label": custody.get("label"),
    }


def _validate_recovery_summary(path: Path, summary_meta: dict[str, Any],
                               record_data: dict[str, Any]) -> dict[str, Any]:
    value = _json(path)
    if value.get("schema") != "docx-edit-provider-profile-recovery-v1" or value.get("status") != "pass":
        fail(f"{path}: retained stack recovery is not a passing recovery")
    source_record = value.get("source_attempt", {}).get("profile_summary")
    if not isinstance(source_record, dict) or source_record.get("sha256") != summary_meta["sha256"]:
        fail(f"{path}: stack recovery is not bound to this profile summary")
    recovered_data = value.get("source_attempt", {}).get("record", {}).get("perf_data", {})
    if recovered_data.get("sha256") != record_data["sha256"] or recovered_data.get("bytes") != record_data["bytes"]:
        fail(f"{path}: stack recovery is not bound to retained perf.data")
    artifacts = value.get("artifacts")
    if not isinstance(artifacts, dict):
        fail(f"{path}: recovery artifact inventory is missing")
    if not isinstance(artifacts.get("stdout"), dict):
        fail(f"{path}: recovery perf-script stdout artifact is missing")
    return {"path": str(path), **meta(path),
            "stack_summary": value.get("stack_summary"),
            "artifacts": artifacts,
            "command": value.get("command"),
            "recovery_helper": value.get("custody", {}).get("recovery_helper"),
            "source_matches_retained": value.get("custody", {}).get("source_matches_retained"),
            "source_unchanged_during_recovery": value.get("custody", {}).get("source_unchanged_during_recovery"),
        }


def _corrected_stack_summary(recovery_record: dict[str, Any], script_path: Path) -> dict[str, Any]:
    """Reclassify the retained export and bind it to recovery-r3's artifact map."""

    expected = recovery_record["artifacts"]["stdout"]
    script_meta = _file_meta(script_path)
    for key in ("path", "bytes", "sha256"):
        if script_meta.get(key) != expected.get(key):
            fail(f"recovery perf-script export {key} differs from recovery receipt")
    parsed = recovery.profiler.parse_perf_script_text(
        script_path.read_text(encoding="utf-8", errors="replace")
    )
    corrected = summarize_stacks_corrected(parsed)
    return {
        "raw": script_meta,
        "old_recovery_summary": recovery_record.get("stack_summary"),
        "corrected": corrected,
        "parser": {
            "name": "profile.parse_perf_script_text",
            "source": str((ROOT / "profile.py").resolve()),
            "note": "retained perf.data export was parsed without a CPU field",
        },
    }


def run(args: argparse.Namespace) -> Path:
    profile_dir = args.profile_dir.resolve()
    output = args.output_dir.resolve()
    if output.exists():
        fail(f"refusing to reuse correction output directory: {output}")
    summary_path = profile_dir / "profile-summary.json"
    summary_meta = _file_meta(summary_path)
    summary = _json(summary_path)
    if summary.get("schema") != "docx-edit-provider-profile-v1":
        fail(f"{summary_path}: profile summary schema differs")
    builds, protocol, custody = recovery._build_protocol(args.build_record.resolve(), summary)
    archive = recovery._helper_archive(args.helper_archive.resolve())
    original_gate = recovery._profile_gate(args.gate_receipt.resolve(), archive["profile"])
    record_inputs = recovery._record_inputs(profile_dir, summary, builds)
    stack_recovery = _validate_recovery_summary(
        args.recovery_summary.resolve(), summary_meta, record_inputs["perf_data"]
    )
    recovery_archive = _recovery_helper_archive(args.recovery_helper_archive.resolve())
    retained_recovery_helper = stack_recovery.get("recovery_helper")
    archived_recovery_helper = recovery_archive["archived"]["recover_profile.py"]
    if not isinstance(retained_recovery_helper, dict):
        fail("recovery summary does not bind its recovery helper")
    for key in ("bytes", "sha256"):
        if retained_recovery_helper.get(key) != archived_recovery_helper.get(key):
            fail(f"recovery summary {key} differs from archived recovery helper")
    corrected_stack = _corrected_stack_summary(
        stack_recovery, args.recovery_script.resolve()
    )
    old_providers = _provider_map(summary)
    source_before = snapshot()
    output.mkdir(parents=True, exist_ok=False)
    corrected: dict[str, Any] = {}
    for provider in PROVIDERS:
        directory = profile_dir / provider
        perf_path = directory / "perf-stat.csv"
        strace_path = directory / "strace-summary.txt"
        perf_raw = _file_meta(perf_path)
        strace_raw = _file_meta(strace_path)
        perf = parse_perf_stat_corrected(perf_path.read_text(encoding="utf-8", errors="replace"))
        strace = parse_strace_corrected(strace_path.read_text(encoding="utf-8", errors="replace"))
        old = old_providers[provider]
        corrected[provider] = {
            "perf_stat": {
                "raw": perf_raw,
                "old_profile_summary": {
                    "path": str(summary_path),
                    "sha256": summary_meta["sha256"],
                    "parsed": old["perf"].get("parsed"),
                },
                "corrected": perf,
                "interpretation": (
                    "running_time_ns uses perf CSV field 3; running_percent uses field 4. "
                    "L1 zero values are retained as raw counters only and do not prove zero "
                    "cache misses or a functioning cache event."
                ),
            },
            "strace": {
                "raw": strace_raw,
                "old_profile_summary": {
                    "path": str(summary_path),
                    "sha256": summary_meta["sha256"],
                    "parsed": old["strace"].get("parsed"),
                },
                "corrected": strace,
                "interpretation": (
                    "strace -f -c is whole-child observer scope. Syscall write rows include "
                    "report serialization and observer-visible writes; they are not output "
                    "sink byte counts or operation-local publication costs."
                ),
            },
        }
    source_after = snapshot()
    correction = {
        "schema": SCHEMA,
        "version": VERSION,
        "status": "pass",
        "created_utc": _now(),
        "source_attempt": {
            "profile_summary": summary_meta,
            "profile_gate": original_gate,
            "build_record": {role: custody["builds"][role] for role in canonical_measure.ROLES},
            "protocol": custody["protocol"],
            "retained_record": record_inputs,
            "stack_recovery": stack_recovery,
            "corrected_stack": corrected_stack,
        },
        "custody": {
            "helper_archive": archive,
            "recovery_helper_archive": recovery_archive,
            "recovery_helper": {"path": str(Path(__file__).resolve()), **meta(Path(__file__))},
            "build_protocol": custody,
            "source_before": source_before,
            "source_after": source_after,
            "source_unchanged_during_correction": source_before == source_after,
        },
        "providers": corrected,
        "scope": {
            "perf_stat": "retained whole-child perf stat output; corrected CSV column interpretation only",
            "strace": "retained whole-child strace -f -c output; corrected optional errors column only",
            "stack": (
                "retained perf.data re-exported by recovery-r3 and reclassified from the raw "
                "export; no workload rerun. Bare run_sample symbols remain ambiguous."
            ),
            "cpu": "original workload and perf stat were taskset CPU 2; perf script recovery was taskset CPU 2 but perf.data samples have no CPU attribute",
        },
        "interpretation": (
            "This receipt corrects parser field positions and optional strace columns from "
            "immutable raw observer files. It preserves the original raw files and old "
            "parsed summary, rejects unexplained rows and total mismatches, and makes no "
            "cache-miss, sink-byte, or operation-local timing claim from these whole-child "
            "observer counters. The corrected stack receipt reports statistical callchain "
            "weights and marks bare run_sample attribution as ambiguous."
        ),
        "stack": corrected_stack,
    }
    output_path = output / "observer-correction.json"
    _write_new(output_path, correction)
    return output_path


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--profile-dir", type=Path, default=ROOT / "profiling-r1")
    parser.add_argument("--recovery-summary", type=Path,
                        default=ROOT / "profiling-recovery-r3" / "recovery-summary.json")
    parser.add_argument("--recovery-script", type=Path,
                        default=ROOT / "profiling-recovery-r3" / "perf-script.txt")
    parser.add_argument("--build-record", type=Path, default=ROOT / "build-normal.json")
    parser.add_argument("--gate-receipt", type=Path, default=ROOT / "validation" / "profile-r1.json")
    parser.add_argument("--helper-archive", type=Path,
                        default=ROOT / "profiling-r1-helper-sources")
    parser.add_argument("--recovery-helper-archive", type=Path,
                        default=ROOT / "profiling-recovery-r3-helper-sources")
    parser.add_argument("--output-dir", type=Path, required=True)
    return parser


def main(argv: list[str] | None = None) -> int:
    try:
        args = _parser().parse_args(argv)
        print(run(args))
        return 0
    except (CorrectionError, OSError, ValueError, canonical_measure.ProviderMatrixError) as error:
        print(f"recover_observer_counters.py: FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
