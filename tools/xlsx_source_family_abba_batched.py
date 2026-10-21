#!/usr/bin/env python3
"""Reproducible batched e63->W4 source-family ABBA evidence.

This adapter invokes already-built litchi-perf-baseline children.  Each child
runs a complete operation for --samples N iterations in one process, so the
measured span is the sum of N operation timings rather than process startup.
Every child result, failure, and short leg is appended to an fsync'd journal
before the next child starts.  No source or benchmark binary is built here.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
from pathlib import Path
import re
import subprocess
import sys
import time
from typing import Any, Iterable

FAMILIES: tuple[tuple[str, str, str, tuple[str, ...]], ...] = (
    ("page_margin", "xlsx_eager_page_margin_edit_save", "xlsx_source_backed_page_margin_edit_save", ()),
    ("defined_names", "xlsx_eager_defined_names_edit_save", "xlsx_source_backed_defined_names_edit_save", ()),
    ("data_validation", "xlsx_eager_data_validation_edit_save", "xlsx_source_backed_data_validation_edit_save", ()),
    ("auto_filter", "xlsx_eager_auto_filter_edit_save", "xlsx_source_backed_auto_filter_edit_save", ()),
    ("conditional_formatting", "xlsx_eager_conditional_formatting_edit_save", "xlsx_source_backed_conditional_formatting_edit_save", ()),
    ("row_visibility", "xlsx_eager_row_visibility_edit_save", "xlsx_source_backed_row_visibility_edit_save", ("--xlsx-row-visibility-shape", "medium")),
)
LEG_ORDER = ("before", "after", "after", "before")
MAX_SAMPLES = 100_000
MAX_CHILD_OUTPUT = 64 * 1024 * 1024
HEAVY_PROCESS_NAMES = {
    "cargo", "rustc", "rustdoc", "rustfmt", "perf", "heaptrack", "valgrind",
    "clang", "clang++", "gcc", "cc", "ld.lld", "scoped_patch_probe", "ole_w5_probe",
}
HEAVY_PROCESS_WORD = re.compile(
    r"(?:^|[\s/'\"])(?:cargo|rustc|rustdoc|rustfmt|perf|heaptrack|valgrind|"
    r"clang(?:\+\+)?|gcc|cc|ld\.lld)(?:$|[\s/'\"])",
)


class RunFailure(RuntimeError):
    pass


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def canonical_sha(value: Any) -> str:
    return hashlib.sha256(json.dumps(value, sort_keys=True, separators=(",", ":")).encode()).hexdigest()


def percentile(values: list[int | float], fraction: float) -> int | float:
    ordered = sorted(values)
    return ordered[round((len(ordered) - 1) * fraction)]


def statistics(values: list[int | float]) -> dict[str, Any]:
    if not values:
        raise ValueError("statistics requires at least one value")
    ordered = sorted(values)
    return {
        "count": len(values),
        "min_ns": ordered[0],
        "p50_ns": percentile(values, 0.50),
        "p95_ns": percentile(values, 0.95),
        "max_ns": ordered[-1],
        "mean_ns": sum(values) / len(values),
        "sum_ns": sum(values),
    }


def bounded_text(value: bytes, limit: int = 64 * 1024) -> str:
    text = value.decode("utf-8", errors="replace")
    return text if len(text) <= limit else text[:limit] + "\n[truncated]"


class Recorder:
    """Append-only evidence journal with durable per-event writes."""

    def __init__(self, out_dir: Path) -> None:
        self.out_dir = out_dir
        if out_dir.exists():
            if not out_dir.is_dir():
                raise RunFailure(f"result path is not a directory: {out_dir}")
            if any(out_dir.iterdir()):
                raise RunFailure(f"result directory must be new and empty: {out_dir}")
        else:
            out_dir.mkdir(parents=True)
        self.raw_dir = out_dir / "raw"
        self.raw_dir.mkdir(parents=True, exist_ok=True)
        self.journal_path = out_dir / "events.ndjson"
        self.stream = self.journal_path.open("a", encoding="utf-8")
        self.index = 0

    def append(self, event: dict[str, Any]) -> int:
        event["event_index"] = self.index
        event["recorded_unix_ns"] = time.time_ns()
        line = json.dumps(event, sort_keys=True, separators=(",", ":"))
        self.stream.write(line + "\n")
        self.stream.flush()
        os.fsync(self.stream.fileno())
        self.index += 1
        return self.index - 1

    def checkpoint(self, name: str, value: dict[str, Any]) -> None:
        destination = self.out_dir / name
        temporary = destination.with_suffix(destination.suffix + ".tmp")
        temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        with temporary.open("rb") as stream:
            os.fsync(stream.fileno())
        os.replace(temporary, destination)
        directory_fd = os.open(destination.parent, os.O_RDONLY | getattr(os, "O_DIRECTORY", 0))
        try:
            os.fsync(directory_fd)
        finally:
            os.close(directory_fd)

    def close(self) -> None:
        self.stream.close()


def extract_result(report: dict[str, Any], expected_case: str, expected_samples: int) -> dict[str, Any]:
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1 or not isinstance(results[0], dict):
        raise RunFailure(f"{expected_case}: child report must contain exactly one object result")
    result = results[0]
    if result.get("case") != expected_case:
        raise RunFailure(f"{expected_case}: child reported case {result.get('case')!r}")
    elapsed = result.get("elapsed_ns")
    samples = elapsed.get("samples") if isinstance(elapsed, dict) else None
    if (
        not isinstance(samples, list)
        or len(samples) != expected_samples
        or not samples
        or any(type(value) is not int or value <= 0 for value in samples)
    ):
        raise RunFailure(f"{expected_case}: child must have exactly {expected_samples} positive elapsed samples")
    if not isinstance(result.get("output_sha256"), str):
        raise RunFailure(f"{expected_case}: child has no output_sha256")
    sink = result.get("sink")
    if isinstance(sink, dict):
        largest_write = sink.get("largest_write")
        if isinstance(largest_write, list):
            if len(largest_write) != expected_samples or any(type(value) is not int for value in largest_write):
                raise RunFailure(f"{expected_case}: sink largest_write vector has the wrong shape")
            largest_write = max(largest_write, default=0)
        if type(largest_write) is int and largest_write > 64 * 1024:
            raise RunFailure(f"{expected_case}: largest sink write exceeds 64 KiB")
    return result


def _metric_vector(value: Any, expected_samples: int, label: str) -> int | list[int] | None:
    """Validate a scalar/vector metric without changing its reported values."""
    if value is None:
        return None
    if type(value) is int:
        if value < 0:
            raise RunFailure(f"{label} must be nonnegative")
        return value
    if not isinstance(value, list) or len(value) != expected_samples:
        raise RunFailure(f"{label} must have exactly {expected_samples} samples")
    if any(type(item) is not int or item < 0 for item in value):
        raise RunFailure(f"{label} samples must be nonnegative integers")
    return list(value)


def source_signature(result: dict[str, Any], expected_samples: int) -> dict[str, Any]:
    source = result.get("source")
    if not isinstance(source, dict):
        return {}
    calls = source.get("read_calls")
    calls_name = "read_calls"
    if calls is None:
        calls = source.get("source_read_calls")
        calls_name = "source_read_calls"
    bytes_read = source.get("read_bytes")
    bytes_name = "read_bytes"
    if bytes_read is None:
        bytes_read = source.get("source_read_bytes")
        bytes_name = "source_read_bytes"
    return {
        "read_calls": _metric_vector(calls, expected_samples, f"source.{calls_name}"),
        "read_bytes": _metric_vector(bytes_read, expected_samples, f"source.{bytes_name}"),
    }


def sink_signature(result: dict[str, Any], expected_samples: int) -> dict[str, Any]:
    sink = result.get("sink")
    if not isinstance(sink, dict):
        return {}
    return {
        "accepted_bytes": _metric_vector(sink.get("accepted_bytes"), expected_samples, "sink.accepted_bytes"),
        "largest_write": _metric_vector(sink.get("largest_write"), expected_samples, "sink.largest_write"),
    }


def paired_block_mean_ratio(legs: list[dict[str, Any]]) -> float:
    """Return mean(after leg means) / mean(before leg means) for one ABBA block."""
    means = {
        revision: [leg["mean_elapsed_ns"] for leg in legs if leg["revision"] == revision]
        for revision in ("before", "after")
    }
    if not means["before"] or not means["after"]:
        raise RunFailure("an ABBA block must contain before and after legs")
    before = sum(means["before"]) / len(means["before"])
    after = sum(means["after"]) / len(means["after"])
    if before <= 0:
        raise RunFailure("before ABBA mean must be positive")
    return after / before


def elapsed_summary(samples: list[int], expected_samples: int) -> dict[str, Any]:
    """Summarize child per-operation samples without scaling them twice."""
    if len(samples) != expected_samples or any(type(value) is not int or value <= 0 for value in samples):
        raise RunFailure(f"child must provide exactly {expected_samples} positive elapsed samples")
    return {
        "sum_elapsed_ns": sum(samples),
        "mean_elapsed_ns": sum(samples) / len(samples),
        "per_op_elapsed_ns": list(samples),
    }


def _process_scan() -> list[dict[str, Any]]:
    """Capture compiler/profiler processes that would contaminate a timing run."""
    processes: list[dict[str, Any]] = []
    proc_root = Path("/proc")
    for entry in proc_root.glob("[0-9]*"):
        try:
            pid = int(entry.name)
            if pid == os.getpid():
                continue
            raw = (entry / "cmdline").read_bytes()
        except (FileNotFoundError, PermissionError, ValueError):
            continue
        parts = [part.decode("utf-8", errors="replace") for part in raw.split(b"\0") if part]
        if not parts:
            continue
        executable = Path(parts[0]).name
        command = " ".join(parts)
        if not is_heavy_process(executable, command):
            continue
        processes.append({"pid": pid, "executable": executable, "command": parts})
    return sorted(processes, key=lambda process: process["pid"])


def is_heavy_process(executable: str, command: str) -> bool:
    """Recognize build/profiling work and every protected benchmark child."""
    if executable in HEAVY_PROCESS_NAMES:
        return True
    if executable.startswith(("litchi-perf-", "xlsx_", "snappy-")):
        return True
    # A probe may be a script launched by a generic interpreter.  Inspect
    # token basenames for those protected probe names without treating this
    # adapter's own xlsx_ filename as a child process.
    for token in command.split():
        basename = Path(token.strip("'\"")).name
        if basename.startswith(("litchi-perf-", "snappy-")) or basename in {
            "scoped_patch_probe", "ole_w5_probe",
        }:
            return True
    return bool(HEAVY_PROCESS_WORD.search(command))


def record_quiet_scan(recorder: Recorder, phase: str) -> dict[str, Any]:
    processes = _process_scan()
    scan = {"phase": phase, "checked_unix_ns": time.time_ns(), "heavy_processes": processes}
    recorder.append({"kind": "process_scan", "status": "quiet" if not processes else "busy", **scan})
    if processes:
        raise RunFailure(f"heavy processes present during {phase}: {[process['pid'] for process in processes]}")
    return scan


def binary_identity(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise RunFailure(f"binary is not a regular file: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_file(path)}


def run_family_child(
    recorder: Recorder,
    binary: Path,
    case: str,
    extra: tuple[str, ...],
    revision: str,
    implementation: str,
    phase: str,
    family: str,
    samples: int,
    warmup: int,
    block: int | None = None,
    leg: int | None = None,
) -> tuple[int, dict[str, Any]]:
    if not 1 <= samples <= MAX_SAMPLES:
        raise ValueError(f"samples outside 1..{MAX_SAMPLES}: {samples}")
    report_path = recorder.raw_dir / f"family-{recorder.index:08d}.json"
    command = [str(binary), "--warmup", str(warmup), "--samples", str(samples), "--case", case, *extra, "--json", str(report_path)]
    started = time.monotonic_ns()
    try:
        completed = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=900)
    except Exception as error:
        event = {
            "kind": "family_child", "status": "spawn_failed", "phase": phase,
            "family": family, "implementation": implementation, "revision": revision,
            "case": case, "samples": samples, "warmup": warmup, "block": block, "leg": leg,
            "command": command, "parent_wall_ns": time.monotonic_ns() - started,
            "error": repr(error),
        }
        recorder.append(event)
        raise RunFailure(f"{case} {revision}: child spawn failed: {error}") from error
    base = {
        "kind": "family_child", "phase": phase, "family": family,
        "implementation": implementation, "revision": revision, "case": case,
        "samples": samples, "warmup": warmup, "block": block, "leg": leg,
        "command": command, "returncode": completed.returncode,
        "parent_wall_ns": time.monotonic_ns() - started,
        "stdout": bounded_text(completed.stdout), "stderr": bounded_text(completed.stderr),
    }
    if completed.returncode != 0:
        base["status"] = "failed"
        recorder.append(base)
        raise RunFailure(f"{case} {revision}: child exited {completed.returncode}")
    try:
        raw = report_path.read_bytes()
        if len(raw) > MAX_CHILD_OUTPUT:
            raise RunFailure(f"{case}: child JSON exceeds {MAX_CHILD_OUTPUT} bytes")
        report = json.loads(raw)
        if not isinstance(report, dict):
            raise RunFailure(f"{case}: child JSON is not an object")
        result = extract_result(report, case, samples)
        sink = sink_signature(result, samples)
        source = source_signature(result, samples)
        if implementation == "source" and (
            sink.get("accepted_bytes") is None or sink.get("largest_write") is None
        ):
            raise RunFailure(f"{case}: source child must report sink accepted_bytes and largest_write")
    except Exception as error:
        base["status"] = "invalid_report"
        base["error"] = repr(error)
        if report_path.exists():
            base["raw_report_sha256"] = sha256_file(report_path)
        recorder.append(base)
        raise RunFailure(f"{case} {revision}: invalid child report: {error}") from error
    elapsed = result["elapsed_ns"]["samples"]
    base.update({
        "status": "ok", "result": result, "report_sha256": canonical_sha(report),
        "output_sha256": result["output_sha256"], "elapsed_samples_ns": elapsed,
        # The child already reports one elapsed value per complete operation.
        # Preserve those values; only the leg mean is normalized.
        **elapsed_summary(elapsed, samples), "sink": sink, "source": source,
    })
    event_index = recorder.append(base)
    return event_index, base


def run_merge_child(
    recorder: Recorder,
    binary: Path,
    revision: str,
    operation: str,
    phase: str,
    repetitions: int,
    block: int | None = None,
    leg: int | None = None,
) -> tuple[list[int], list[dict[str, Any]]]:
    timings: list[int] = []
    events: list[dict[str, Any]] = []
    for repetition in range(repetitions):
        command = [str(binary), "--child", "--implementation", "source", "--operation", operation]
        started = time.monotonic_ns()
        try:
            completed = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=900)
        except Exception as error:
            event = {
                "kind": "merge_child", "status": "spawn_failed", "phase": phase,
                "revision": revision, "operation": operation, "repetition": repetition,
                "block": block, "leg": leg, "command": command,
                "parent_wall_ns": time.monotonic_ns() - started, "error": repr(error),
            }
            recorder.append(event)
            raise RunFailure(f"merge {operation} {revision}: spawn failed: {error}") from error
        base = {
            "kind": "merge_child", "phase": phase, "revision": revision,
            "operation": operation, "repetition": repetition, "block": block, "leg": leg,
            "command": command, "returncode": completed.returncode,
            "parent_wall_ns": time.monotonic_ns() - started,
            "stdout": bounded_text(completed.stdout), "stderr": bounded_text(completed.stderr),
        }
        if completed.returncode != 0:
            base["status"] = "failed"
            recorder.append(base)
            raise RunFailure(f"merge {operation} {revision}: child exited {completed.returncode}")
        try:
            value = json.loads(completed.stdout)
            if not isinstance(value, dict):
                raise RunFailure("child output is not an object")
            primary_ns = value.get("primary_ns")
            if not isinstance(primary_ns, int) or primary_ns <= 0:
                raise RunFailure("child primary_ns is not positive")
            patch = value.get("patch")
            if not isinstance(patch, dict):
                raise RunFailure("child has no patch object")
            required = ("forward_applied", "forward_exact_one_part_overlay", "inverse_restored_complete_package", "foreign_refused")
            if any(patch.get(key) is not True for key in required) or value.get("stale_refused") is not True:
                raise RunFailure("child replay safety flags are not all true")
            output_sha = value.get("output_sha256")
            if not isinstance(output_sha, str):
                raise RunFailure("child has no output_sha256")
        except Exception as error:
            base["status"] = "invalid_report"
            base["error"] = repr(error)
            recorder.append(base)
            raise RunFailure(f"merge {operation} {revision}: invalid child output: {error}") from error
        base.update({"status": "ok", "result": value, "primary_ns": primary_ns, "output_sha256": output_sha})
        recorder.append(base)
        timings.append(primary_ns)
        events.append(base)
    return timings, events


def write_json(path: Path, value: Any) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    with temporary.open("rb") as stream:
        os.fsync(stream.fileno())
    os.replace(temporary, path)


def family_spec(name: str) -> tuple[str, str, str, tuple[str, ...]]:
    for family in FAMILIES:
        if family[0] == name:
            return family
    raise RunFailure(f"unknown family {name}")


def calibrate_family(
    recorder: Recorder,
    binary_by_revision: dict[str, Path],
    spec: tuple[str, str, str, tuple[str, ...]],
    implementation: str,
    target_ns: int,
    minimum_ns: int,
) -> dict[str, Any]:
    family, eager_case, source_case, extra = spec
    case = eager_case if implementation == "eager" else source_case
    normalization: dict[str, Any] = {}
    for revision in LEG_ORDER[:2]:
        _, first = run_family_child(recorder, binary_by_revision[revision], case, extra, revision, implementation, "calibration_one", family, 1, 0)
        first_values = list(first["elapsed_samples_ns"])
        n = max(1, math.ceil(target_ns / max(1, min(first_values))))
        attempts: list[dict[str, Any]] = []
        while True:
            _, check = run_family_child(recorder, binary_by_revision[revision], case, extra, revision, implementation, "calibration_batch", family, n, 0)
            attempts.append({"event_index": check["event_index"], "samples": n, "sum_elapsed_ns": check["sum_elapsed_ns"], "min_elapsed_ns": min(check["elapsed_samples_ns"])})
            if check["sum_elapsed_ns"] >= minimum_ns:
                break
            if n >= MAX_SAMPLES:
                raise RunFailure(f"{family}/{implementation}/{revision}: could not reach minimum leg duration")
            n = min(MAX_SAMPLES, n * 2)
        normalization[revision] = {"samples": n, "target_ns": target_ns, "minimum_ns": minimum_ns, "first_event_index": first["event_index"], "attempts": attempts}
    return normalization


def family_leg(
    recorder: Recorder,
    binary_by_revision: dict[str, Path],
    spec: tuple[str, str, str, tuple[str, ...]],
    implementation: str,
    normalization: dict[str, Any],
    blocks: int,
    warmup: int,
    minimum_ns: int,
) -> dict[str, Any]:
    family, eager_case, source_case, extra = spec
    case = eager_case if implementation == "eager" else source_case
    warmup_events: list[int] = []
    for revision in ("before", "after"):
        for index in range(warmup):
            event_index, _ = run_family_child(recorder, binary_by_revision[revision], case, extra, revision, implementation, "warmup", family, normalization[revision]["samples"], 0, block=-1, leg=index)
            warmup_events.append(event_index)
    observations: list[dict[str, Any]] = []
    checkpoints: list[dict[str, Any]] = []
    for block in range(blocks):
        process_scan_before = record_quiet_scan(recorder, f"before_{family}_{implementation}_block_{block}")
        legs: list[dict[str, Any]] = []
        for leg, revision in enumerate(LEG_ORDER):
            event_index, event = run_family_child(recorder, binary_by_revision[revision], case, extra, revision, implementation, "sample", family, normalization[revision]["samples"], 0, block=block, leg=leg)
            if event["sum_elapsed_ns"] < minimum_ns:
                # The short result is already durable in the journal.
                raise RunFailure(f"{family}/{implementation}/{revision} block {block} leg {leg}: measured {event['sum_elapsed_ns']} ns below {minimum_ns}")
            legs.append({"event_index": event_index, "revision": revision, "leg": leg, "sum_elapsed_ns": event["sum_elapsed_ns"], "mean_elapsed_ns": event["mean_elapsed_ns"], "elapsed_samples_ns": event["elapsed_samples_ns"], "per_op_elapsed_ns": event["per_op_elapsed_ns"], "output_sha256": event["output_sha256"], "sink": event["sink"], "source": event["source"]})
            observations.append(event)
        checkpoint = {
            "family": family,
            "implementation": implementation,
            "block": block,
            "legs": legs,
            "paired_block_mean_ratio": paired_block_mean_ratio(legs),
            "process_scan_before_block": process_scan_before,
        }
        checkpoint["mean_elapsed_ns_by_revision"] = {
            revision: sum(leg["mean_elapsed_ns"] for leg in legs if leg["revision"] == revision)
            / len([leg for leg in legs if leg["revision"] == revision])
            for revision in ("before", "after")
        }
        checkpoint["process_scan"] = record_quiet_scan(recorder, f"after_{family}_{implementation}_block_{block}")
        recorder.checkpoint(f"checkpoint-{family}-{implementation}-block-{block:03d}.json", checkpoint)
        checkpoints.append(checkpoint)
    validate_family_observations(family, implementation, observations)
    per_revision = {revision: [value for event in observations if event["revision"] == revision for value in event["per_op_elapsed_ns"]] for revision in ("before", "after")}
    source_by_revision = {
        revision: next(event["source"] for event in observations if event["revision"] == revision)
        for revision in ("before", "after")
    }
    sink_by_revision = {
        revision: next(event["sink"] for event in observations if event["revision"] == revision)
        for revision in ("before", "after")
    }
    ratios = [checkpoint["paired_block_mean_ratio"] for checkpoint in checkpoints]
    return {
        "family": family,
        "implementation": implementation,
        "normalization": normalization,
        "warmup_event_indices": warmup_events,
        "blocks": checkpoints,
        "paired_block_mean_ratios": ratios,
        "statistics_paired_block_mean_ratio": statistics(ratios),
        "statistics_per_op": {revision: statistics(values) for revision, values in per_revision.items()},
        "output_sha256": observations[0]["output_sha256"],
        "sink_by_revision": sink_by_revision,
        "source_by_revision": source_by_revision,
    }


def validate_family_observations(family: str, implementation: str, observations: list[dict[str, Any]]) -> None:
    if not observations:
        raise RunFailure(f"{family}/{implementation}: no observations")
    output = observations[0]["output_sha256"]
    sink_by_revision: dict[str, dict[str, Any]] = {}
    source_by_revision: dict[str, dict[str, Any]] = {}
    for event in observations:
        if event["output_sha256"] != output:
            raise RunFailure(f"{family}/{implementation}: output hash changed")
        revision = event["revision"]
        if revision not in sink_by_revision:
            sink_by_revision[revision] = event["sink"]
            source_by_revision[revision] = event["source"]
        elif event["sink"] != sink_by_revision[revision]:
            raise RunFailure(f"{family}/{implementation}/{revision}: sink geometry changed")
        elif event["source"] != source_by_revision[revision]:
            raise RunFailure(f"{family}/{implementation}/{revision}: source read evidence changed")
        if implementation == "source":
            source = event["source"]
            if not source.get("read_calls") or not source.get("read_bytes"):
                raise RunFailure(f"{family}/{implementation}: source read evidence absent")
            sink = event["sink"]
            if sink.get("accepted_bytes") is None or sink.get("largest_write") is None:
                raise RunFailure(f"{family}/{implementation}: sink geometry absent")


def run_compact(
    recorder: Recorder,
    binaries: dict[str, Path],
    target_ns: int,
    minimum_ns: int,
    blocks: int,
    warmup: int,
) -> dict[str, Any]:
    normalized: dict[str, int] = {}
    calibration: dict[str, Any] = {}
    for operation in ("merge", "unmerge"):
        for revision in ("before", "after"):
            first_times, _ = run_merge_child(recorder, binaries[revision], revision, operation, "calibration_one", 1)
            reps = max(1, math.ceil(target_ns / max(1, first_times[0])))
            attempts: list[dict[str, Any]] = []
            while True:
                times, _ = run_merge_child(recorder, binaries[revision], revision, operation, "calibration_batch", reps)
                total = sum(times)
                attempts.append({"repetitions": reps, "sum_primary_ns": total, "min_primary_ns": min(times)})
                if total >= minimum_ns:
                    break
                if reps >= MAX_SAMPLES:
                    raise RunFailure(f"compact/{operation}/{revision}: could not reach minimum duration")
                reps = min(MAX_SAMPLES, reps * 2)
            normalized[f"{operation}/{revision}"] = reps
            calibration[f"{operation}/{revision}"] = {"repetitions": reps, "attempts": attempts}
    warmup_events: list[dict[str, Any]] = []
    for revision in ("before", "after"):
        for operation in ("merge", "unmerge"):
            for _ in range(warmup):
                _, events = run_merge_child(recorder, binaries[revision], revision, operation, "warmup", normalized[f"{operation}/{revision}"])
                warmup_events.extend(events)
    summaries: dict[str, Any] = {"calibration": calibration, "normalization": normalized, "warmup_event_count": len(warmup_events), "operations": {}}
    for operation in ("merge", "unmerge"):
        blocks_out: list[dict[str, Any]] = []
        measured: list[dict[str, Any]] = []
        for block in range(blocks):
            process_scan_before = record_quiet_scan(recorder, f"before_compact_{operation}_block_{block}")
            legs: list[dict[str, Any]] = []
            for leg, revision in enumerate(LEG_ORDER):
                repetitions = normalized[f"{operation}/{revision}"]
                times, events = run_merge_child(recorder, binaries[revision], revision, operation, "sample", repetitions, block, leg)
                total = sum(times)
                if total < minimum_ns:
                    raise RunFailure(f"compact/{operation}/{revision} block {block} leg {leg}: measured {total} ns below {minimum_ns}")
                first_hash = events[0]["output_sha256"]
                if any(event["output_sha256"] != first_hash for event in events):
                    raise RunFailure(f"compact/{operation}/{revision} block {block}: output hash changed")
                legs.append({"revision": revision, "leg": leg, "event_indices": [event["event_index"] for event in events], "sum_primary_ns": total, "mean_primary_ns": total / len(times), "primary_ns": times, "output_sha256": first_hash})
                measured.extend(events)
            before_means = [leg["mean_primary_ns"] for leg in legs if leg["revision"] == "before"]
            after_means = [leg["mean_primary_ns"] for leg in legs if leg["revision"] == "after"]
            paired_ratio = (sum(after_means) / len(after_means)) / (sum(before_means) / len(before_means))
            checkpoint = {
                "operation": operation,
                "block": block,
                "legs": legs,
                "paired_block_mean_ratio": paired_ratio,
                "process_scan_before_block": process_scan_before,
                "process_scan": record_quiet_scan(recorder, f"after_compact_{operation}_block_{block}"),
            }
            recorder.checkpoint(f"checkpoint-compact-{operation}-block-{block:03d}.json", checkpoint)
            blocks_out.append(checkpoint)
        hashes = {event["output_sha256"] for event in measured}
        if len(hashes) != 1:
            raise RunFailure(f"compact/{operation}: output hash changed across ABBA legs")
        leg_totals = {
            revision: [
                leg["sum_primary_ns"]
                for block in blocks_out
                for leg in block["legs"]
                if leg["revision"] == revision
            ]
            for revision in ("before", "after")
        }
        summaries["operations"][operation] = {
            "blocks": blocks_out,
            "statistics_primary_ns_per_child": {revision: statistics([event["result"]["primary_ns"] for event in measured if event["revision"] == revision]) for revision in ("before", "after")},
            "statistics_primary_sum_ns_per_leg": {revision: statistics(values) for revision, values in leg_totals.items()},
            "paired_block_mean_ratios": [block["paired_block_mean_ratio"] for block in blocks_out],
            "statistics_paired_block_mean_ratio": statistics([block["paired_block_mean_ratio"] for block in blocks_out]),
            "output_sha256": measured[0]["output_sha256"],
        }
    return summaries


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--before-bin", type=Path, required=True)
    parser.add_argument("--after-bin", type=Path, required=True)
    parser.add_argument("--merge-before-bin", type=Path)
    parser.add_argument("--merge-after-bin", type=Path)
    parser.add_argument("--out-dir", type=Path, required=True)
    parser.add_argument("--family", action="append", choices=[family[0] for family in FAMILIES])
    parser.add_argument("--blocks", type=int, default=5)
    parser.add_argument("--warmup", type=int, default=2)
    parser.add_argument("--target-ms", type=float, default=500.0)
    parser.add_argument("--minimum-ms", type=float, default=150.0)
    parser.add_argument("--compact", action="store_true")
    parser.add_argument("--plan", dest="plan_only", action="store_true", help="write the complete route/normalization plan without invoking children")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.blocks < 1 or args.warmup < 0:
        raise SystemExit("blocks must be positive and warmup nonnegative")
    if args.compact != (args.merge_before_bin is not None and args.merge_after_bin is not None):
        raise SystemExit("--compact requires both merge binaries and no compact flag is invalid without both")
    if args.merge_before_bin is not None and args.merge_after_bin is None:
        raise SystemExit("merge binaries must be supplied as a pair")
    target_ns = round(args.target_ms * 1_000_000)
    minimum_ns = round(args.minimum_ms * 1_000_000)
    if target_ns < minimum_ns or minimum_ns <= 0:
        raise SystemExit("target must be >= positive minimum")
    selected = [family_spec(name) for name in args.family] if args.family else list(FAMILIES)
    binaries = {"before": args.before_bin.resolve(), "after": args.after_bin.resolve()}
    binary_metadata = {revision: binary_identity(path) for revision, path in binaries.items()}
    merge_binaries = None
    if args.compact:
        merge_binaries = {
            "before": args.merge_before_bin.resolve(),
            "after": args.merge_after_bin.resolve(),
        }
    merge_metadata = {
        revision: binary_identity(path) for revision, path in merge_binaries.items()
    } if merge_binaries is not None else None
    runner = Path(__file__).resolve()
    try:
        cpu_affinity = sorted(os.sched_getaffinity(0))
    except (AttributeError, OSError):
        cpu_affinity = None
    out = args.out_dir.resolve()
    recorder = Recorder(out)
    metadata = {
        "schema": "litchi.xlsx.source-family-abba.batched.v1", "claim": "paired complete-operation evidence with batched child samples",
        "before_binary": binary_metadata["before"], "after_binary": binary_metadata["after"],
        "merge_before_binary": merge_metadata["before"] if merge_metadata is not None else None,
        "merge_after_binary": merge_metadata["after"] if merge_metadata is not None else None,
        "runner": binary_identity(runner), "cpu_affinity": cpu_affinity,
        "families": [family[0] for family in selected], "blocks": args.blocks, "warmup": args.warmup,
        "target_ns": target_ns, "minimum_ns": minimum_ns, "started_unix_ns": time.time_ns(),
        "plan_only": args.plan_only,
    }
    write_json(out / "run.json", metadata)
    if args.plan_only:
        metadata.update({"status": "plan_only", "finished_unix_ns": time.time_ns(), "planned_route_count": len(selected), "planned_family_routes": [{"family": f[0], "eager_case": f[1], "source_case": f[2], "extra_args": list(f[3])} for f in selected], "planned_compact": args.compact})
        write_json(out / "plan.json", metadata)
        recorder.close()
        return 0
    try:
        metadata["pre_run_process_scan"] = record_quiet_scan(recorder, "pre_run")
        write_json(out / "run.json", metadata)
        route_summaries: list[dict[str, Any]] = []
        for spec in selected:
            family = spec[0]
            route: dict[str, Any] = {"family": family}
            normalizations: dict[str, Any] = {}
            for implementation in ("eager", "source"):
                normalizations[implementation] = calibrate_family(recorder, binaries, spec, implementation, target_ns, minimum_ns)
            route["normalization"] = normalizations
            route["eager"] = family_leg(recorder, binaries, spec, "eager", normalizations["eager"], args.blocks, args.warmup, minimum_ns)
            route["source"] = family_leg(recorder, binaries, spec, "source", normalizations["source"], args.blocks, args.warmup, minimum_ns)
            if route["eager"]["output_sha256"] != route["source"]["output_sha256"]:
                raise RunFailure(f"{family}: eager/source output hashes differ")
            write_json(out / f"route-{family}.json", route)
            route_summaries.append(route)
        compact = None
        if args.compact and not args.plan_only:
            compact = run_compact(recorder, binaries, target_ns, minimum_ns, args.blocks, args.warmup)
            write_json(out / "compact.json", compact)
        metadata.update({"status": "complete" if not args.plan_only else "preflight_complete", "finished_unix_ns": time.time_ns(), "routes": route_summaries, "compact": compact})
        write_json(out / "report.json", metadata)
        return 0
    except Exception as error:
        metadata.update({"status": "failed", "finished_unix_ns": time.time_ns(), "error": repr(error)})
        write_json(out / "failure.json", metadata)
        print(f"error: {error}", file=sys.stderr)
        return 1
    finally:
        recorder.close()


if __name__ == "__main__":
    raise SystemExit(main())
