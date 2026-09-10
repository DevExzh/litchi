#!/usr/bin/env python3
"""Run bounded whole-child observers for the 0495 managed edit provider.

This helper observes the retained *after* normal build.  It does not build the
target or run the formal matrix.  Each perf/strace/perf-record child includes
the complete Rust command: process startup, corpus/preflight work, the timed
edit/commit/publication path, post-clock output validation, and report
serialization.  The counters therefore describe the whole child.  The Rust
report remains the authority for operation-local latency, allocation, budget,
source, sink, and output invariants.

Observer failures are retained as typed ``unavailable`` or ``failed`` results;
unsupported PMU/ptrace access is never represented as zero counters.  Every
attempt keeps its command, raw output hashes, process terminal receipt, source
binding, and private-TMP cleanup receipt.
"""

from __future__ import annotations

import argparse
from collections import Counter
from contextlib import contextmanager
import datetime as _datetime
import fcntl
import json
import math
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
from typing import Any, Iterable, Iterator

sys.path.insert(0, str(Path(__file__).resolve().parent))
from support import ENV, REPO, ROOT, TEMP, meta, sha  # noqa: E402
import measure as canonical_measure  # noqa: E402


SCHEMA = "docx-edit-provider-managed-profile-v1"
VERSION = 1
PHASE = "after"
ROLE = "normal"
APIS = tuple(canonical_measure.APIS)
PROFILE_PROVIDERS = ("owned", "file-warm", "short")
PROFILE_TO_ARM = {"owned": "owned", "file-warm": "file-warm", "short": "short"}
FULL_EVENTS = (
    "task-clock", "cycles", "instructions", "branches", "branch-misses",
    "L1-dcache-loads", "L1-dcache-load-misses", "LLC-loads",
    "LLC-load-misses", "minor-faults", "major-faults", "context-switches",
    "cpu-migrations", "page-faults",
)
CORE_EVENTS = (
    "task-clock", "cycles", "instructions", "branches", "branch-misses",
    "minor-faults", "major-faults",
)
PERF_RECORD_FREQUENCY_HZ = 199
PERF_RECORD_SAMPLES = 100
PERF_RECORD_WARMUP = 3
TASKSET = "/usr/bin/taskset"
PERF = "/usr/bin/perf"
STRACE = "/usr/bin/strace"
REVISION_RE = re.compile(r"^[0-9a-f]{40}$")
EVENT_RE = re.compile(r"^[A-Za-z0-9_.:-]+$")
LABEL_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.-]*$")

WHOLE_CHILD_SCOPE = (
    "perf/strace/perf-record observes one complete child process including "
    "process startup, fixture/preflight setup, warmups, the measured opened "
    "document edit/commit/publication, post-clock output hashing and semantic "
    "validation, and report serialization; counters are not operation-local"
)
REPORT_SCOPE = (
    "operation-local timing, allocation, resource-budget, source, sink, and "
    "output invariants are accepted only through measure.validate_report"
)
STACK_SCOPE = (
    "perf-record periods are statistical whole-child samples; run_sample and "
    "publication markers are call-stack subsets and are not phase durations; "
    "bare symbols are accepted for this selected harness executable and may be "
    "ambiguous when DWARF symbolization is incomplete"
)


class ProfileError(RuntimeError):
    """A fail-closed profile custody, parser, or artifact error."""


def fail(message: str) -> None:
    raise ProfileError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def _now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _json(path: Path) -> Any:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError, json.JSONDecodeError) as error:
        fail(f"{path}: invalid JSON: {error}")
    _finite(value, str(path))
    return value


def _finite(value: Any, path: str) -> None:
    if isinstance(value, float):
        require(math.isfinite(value), f"{path}: non-finite number")
    elif isinstance(value, dict):
        for key, child in value.items():
            require(isinstance(key, str), f"{path}: non-string key")
            _finite(child, f"{path}.{key}")
    elif isinstance(value, list):
        for index, child in enumerate(value):
            _finite(child, f"{path}[{index}]")


def _write_new(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as stream:
            json.dump(value, stream, indent=2, sort_keys=True, allow_nan=False)
            stream.write("\n")
    except FileExistsError:
        fail(f"refusing to replace immutable profile artifact: {path}")


def _numeric(value: str) -> int | float | None:
    text = value.strip().replace(" ", "")
    if not text or text.startswith("<") or text in {"-", "<notcounted>", "?"}:
        return None
    try:
        return int(text)
    except ValueError:
        try:
            return float(text)
        except ValueError:
            return None


def parse_perf_text(text: str) -> dict[str, Any]:
    """Parse ``perf stat -x,`` without confusing runtime and run percentage.

    Current perf CSV output places runtime in field 3 (zero based index 3) and
    the multiplexing percentage in field 4.  Both are retained explicitly so
    a later report cannot silently treat nanoseconds as a percentage.
    """

    events: dict[str, dict[str, Any]] = {}
    unparsed: list[str] = []
    comments: list[str] = []
    parsed_lines: list[str] = []
    for raw in text.splitlines():
        line = raw.rstrip("\r")
        if not line.strip():
            continue
        if line.lstrip().startswith("#"):
            comments.append(line)
            continue
        fields = [item.strip() for item in line.split(",")]
        if len(fields) < 3:
            unparsed.append(line)
            continue
        event = fields[2]
        if not event or EVENT_RE.fullmatch(event) is None:
            unparsed.append(line)
            continue
        value_text = fields[0]
        running_time_text = fields[3] if len(fields) > 3 else ""
        running_percent_text = fields[4] if len(fields) > 4 else ""
        running_percent = _numeric(running_percent_text)
        if isinstance(running_percent, (int, float)) and not 0 <= running_percent <= 100:
            unparsed.append(line)
            continue
        events[event] = {
            "value": _numeric(value_text),
            "value_text": value_text,
            "unit": fields[1] or None,
            "event": event,
            "running_time_ns": _numeric(running_time_text),
            "running_percent": running_percent,
            "raw": line,
        }
        parsed_lines.append(line)

    def value(name: str) -> int | float | None:
        item = events.get(name)
        return None if item is None else item["value"]

    cycles, instructions = value("cycles"), value("instructions")
    branches, branch_misses = value("branches"), value("branch-misses")
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
        "raw_lines": parsed_lines,
        "unparsed_lines": unparsed,
        "ignored_comment_lines": comments,
    }


def parse_strace_text(text: str) -> dict[str, Any]:
    """Parse ``strace -f -c`` with either a present or blank errors column."""

    syscalls: dict[str, dict[str, Any]] = {}
    unparsed: list[str] = []
    comments: list[str] = []
    reported_total: dict[str, Any] | None = None
    for raw in text.splitlines():
        line = raw.rstrip("\r")
        stripped = line.strip()
        if not stripped or stripped.startswith("-") or stripped.startswith("% time"):
            continue
        if stripped.startswith("#"):
            comments.append(line)
            continue
        parts = stripped.split()
        if stripped.endswith(" total"):
            if len(parts) >= 6:
                reported_total = {
                    "raw": line,
                    "calls": _numeric(parts[-3]),
                    "errors": _numeric(parts[-2]),
                }
            elif len(parts) == 5:
                reported_total = {
                    "raw": line,
                    "calls": _numeric(parts[-2]),
                    "errors": None,
                }
            else:
                unparsed.append(line)
            continue
        # A zero-error row may omit the errors field entirely.  Because split()
        # removes the blank column, len==5 is the unambiguous no-errors form.
        if len(parts) == 5:
            pct, seconds, usec, calls, syscall = parts
            errors = None
        elif len(parts) >= 6:
            pct, seconds, usec, calls, errors = parts[:5]
            syscall = " ".join(parts[5:])
        else:
            unparsed.append(line)
            continue
        if not syscall or any(_numeric(token) is None for token in (pct, seconds, usec, calls)):
            unparsed.append(line)
            continue
        record = {
            "pct": _numeric(pct),
            "seconds": _numeric(seconds),
            "usec": _numeric(usec),
            "calls": _numeric(calls),
            "errors": _numeric(errors) if errors is not None else None,
            "raw": line,
        }
        syscalls[syscall] = record
    total_calls = sum(
        int(item["calls"]) for item in syscalls.values()
        if isinstance(item.get("calls"), (int, float))
    )
    error_values = [item["errors"] for item in syscalls.values()
                    if isinstance(item.get("errors"), (int, float))]
    total_errors = sum(int(item) for item in error_values)
    return {
        "syscalls": syscalls,
        "total_calls": total_calls,
        "total_errors": total_errors,
        "reported_total": reported_total,
        "unparsed_lines": unparsed,
        "ignored_comment_lines": comments,
    }


_PERF_SCRIPT_HEADER_RE = re.compile(
    r"^\S.*:\s+(?:(?P<period>\d+)\s+)?"
    r"(?P<event>[^\s:]+(?::[^\s:]+)*):\s*$"
)


def _perf_frame_symbol(line: str) -> str | None:
    if not line or not line[0].isspace():
        return None
    fields = line.strip().split()
    if not fields:
        return None
    if len(fields) >= 2 and re.fullmatch(r"(?:0x)?[0-9a-f]+", fields[0], re.IGNORECASE):
        symbol = fields[1]
    else:
        symbol = fields[0]
    symbol = symbol.split("+0x", 1)[0]
    return symbol or None


def parse_perf_script_text(text: str) -> dict[str, Any]:
    """Parse a retained perf script; comments are metadata, not samples."""

    samples: list[dict[str, Any]] = []
    unparsed: list[str] = []
    comments: list[str] = []
    current: dict[str, Any] | None = None

    def flush() -> None:
        nonlocal current
        if current is not None:
            if current["frames"]:
                samples.append(current)
            else:
                unparsed.append(current["header"])
        current = None

    for raw in text.splitlines():
        line = raw.rstrip("\r")
        if not line.strip():
            flush()
            continue
        if line.lstrip().startswith("#"):
            comments.append(line)
            continue
        header = _PERF_SCRIPT_HEADER_RE.match(line)
        if header is not None and "cycles" in header.group("event"):
            flush()
            period_text = header.group("period")
            current = {
                "header": line,
                "period": int(period_text) if period_text is not None else 1,
                "frames": [],
            }
            continue
        if current is None:
            unparsed.append(line)
            continue
        symbol = _perf_frame_symbol(line)
        if symbol is None:
            unparsed.append(line)
        else:
            current["frames"].append(symbol)
    flush()
    return {
        "samples": samples,
        "sample_count": len(samples),
        "total_period": sum(int(item["period"]) for item in samples),
        "unparsed_lines": unparsed,
        "ignored_comment_lines": comments,
    }


def _symbol_is(frame: str, name: str) -> bool:
    symbol = frame.split("+0x", 1)[0]
    return symbol == name or symbol.endswith(f"::{name}")


def classify_perf_stack(frames: Iterable[str]) -> str:
    frame_list = list(frames)
    has_run_sample = any(_symbol_is(frame, "run_sample") for frame in frame_list)
    has_publish = any(_symbol_is(frame, "publish_docx_source_edit") for frame in frame_list)
    has_prepare = any(_symbol_is(frame, "prepare") for frame in frame_list)
    has_output_oracle = any(
        _symbol_is(frame, "verify_docx_source_edit_output") or
        _symbol_is(frame, "sha256_hex")
        for frame in frame_list
    )
    if has_run_sample and has_publish:
        return "run_sample_publish_ancestor"
    if has_run_sample and has_output_oracle:
        return "run_sample_output_oracle"
    if has_run_sample:
        return "run_sample_setup_or_teardown"
    if has_prepare:
        return "preflight"
    return "process_setup_or_unclassified"


def summarize_perf_stacks(parsed: dict[str, Any]) -> dict[str, Any]:
    samples = parsed.get("samples")
    require(isinstance(samples, list), "perf script parser did not return samples")
    folded: Counter[str] = Counter()
    classes: Counter[str] = Counter()
    class_samples: Counter[str] = Counter()
    class_leafs: dict[str, Counter[str]] = {}
    run_sample_period = 0
    publish_period = 0
    for item in samples:
        if not isinstance(item, dict) or not isinstance(item.get("frames"), list):
            continue
        frames = [str(frame) for frame in item["frames"] if str(frame)]
        if not frames:
            continue
        weight = max(1, int(item.get("period", 1)))
        folded[";".join(reversed(frames))] += weight
        stack_class = classify_perf_stack(frames)
        classes[stack_class] += weight
        class_samples[stack_class] += 1
        class_leafs.setdefault(stack_class, Counter())[frames[0]] += weight
        if any(_symbol_is(frame, "run_sample") for frame in frames):
            run_sample_period += weight
        if stack_class == "run_sample_publish_ancestor":
            publish_period += weight
    total_period = sum(classes.values())
    rows = []
    for name in sorted(classes):
        period = classes[name]
        rows.append({
            "class": name,
            "sample_count": class_samples[name],
            "period": period,
            "share_of_all_period": period / total_period if total_period else None,
            "share_of_run_sample_period": period / run_sample_period if run_sample_period else None,
            "top_leaf_symbols": [
                {"symbol": symbol, "period": count}
                for symbol, count in sorted(class_leafs[name].items(), key=lambda pair: (-pair[1], pair[0]))[:20]
            ],
        })
    return {
        "schema": "docx-edit-provider-perf-stacks-v1",
        "sample_count": len(samples),
        "total_period": total_period,
        "run_sample_period": run_sample_period,
        "run_sample_publish_period": publish_period,
        "classes": rows,
        "folded_stack_count": len(folded),
        "folded_stacks": [
            {"stack": stack, "period": period}
            for stack, period in sorted(folded.items(), key=lambda pair: (-pair[1], pair[0]))
        ],
        "scope": STACK_SCOPE,
    }


def provider_args(provider: str) -> list[str]:
    if provider not in PROFILE_PROVIDERS:
        fail(f"unknown provider {provider!r}; expected {', '.join(PROFILE_PROVIDERS)}")
    arm = dict(canonical_measure.ARM_BY_NAME[PROFILE_TO_ARM[provider]])
    result = ["--provider", arm["provider"]]
    if arm["max_range_bytes"] is not None:
        result += ["--max-range", str(arm["max_range_bytes"])]
    if arm["trace_ranges"]:
        result += ["--trace-ranges"]
    if arm["short_read_bytes"] is not None and arm["max_range_bytes"] is None:
        result += ["--short-read-bytes", str(arm["short_read_bytes"])]
    return result


def _target_command(binary: Path, api: str, provider: str, samples: int, warmup: int,
                    revision: str, report: Path, cpu: int) -> list[str]:
    require(api in APIS, f"unknown API {api}")
    return [
        TASKSET, "-c", str(cpu), str(binary), "docx-managed-edit",
        "--edit-api", api, *provider_args(provider),
        "--samples", str(samples), "--warmup", str(warmup),
        "--source-revision", revision, "--output", str(report),
    ]


def _observer_command(tool: str, binary: Path, api: str, provider: str,
                      samples: int, warmup: int, revision: str, report: Path,
                      cpu: int, observer_output: Path, events: Iterable[str] | None = None) -> list[str]:
    target = _target_command(binary, api, provider, samples, warmup, revision, report, cpu)
    target_args = target[3:]
    if tool == "perf":
        selected = tuple(events or FULL_EVENTS)
        return [TASKSET, "-c", str(cpu), PERF, "stat", "--no-big-num", "-x,",
                "-e", ",".join(selected), "-o", str(observer_output), "--", *target_args]
    if tool == "strace":
        return [TASKSET, "-c", str(cpu), STRACE, "-f", "-c", "-o",
                str(observer_output), "--", *target_args]
    fail(f"unknown observer {tool}")


def _record_command(binary: Path, api: str, samples: int, warmup: int,
                    revision: str, report: Path, cpu: int, data: Path,
                    frequency_hz: int) -> list[str]:
    return [
        TASKSET, "-c", str(cpu), PERF, "record", "-F", str(frequency_hz),
        "-e", "cycles:u", "--call-graph", "dwarf", "-o", str(data), "--",
        *_target_command(binary, api, "owned", samples, warmup, revision, report, cpu)[3:],
    ]


def _script_command(data: Path, cpu: int) -> list[str]:
    # Do not request `cpu`: on some perf versions cycles:u records do not carry
    # a CPU field and perf script then fails before exporting any stacks.
    return [
        TASKSET, "-c", str(cpu), PERF, "script", "--header", "--demangle",
        "-F", "comm,pid,tid,time,period,event,ip,sym,dso", "-i", str(data),
    ]


def _provider_report_identity(path: Path, provider: str, api: str, *, binary: Path,
                              revision: str, samples: int, warmup: int) -> dict[str, Any]:
    if not path.is_file() or path.is_symlink():
        fail(f"{path}: observer did not produce a regular report")
    try:
        value = canonical_measure.validate_report(
            path, role=ROLE, api=api, arm_name=PROFILE_TO_ARM[provider],
            samples=samples, warmups=warmup, source_revision=revision,
            binary_sha256=sha(binary), binary_bytes=binary.stat().st_size,
        )
    except (KeyError, TypeError, canonical_measure.ProviderMatrixError, OSError, ValueError) as error:
        fail(f"{path}: canonical report validation failed: {error}")
    reported = value["provider"]
    return {
        "path": str(path), "sha256": sha(path), "bytes": path.stat().st_size,
        "schema": value["schema"], "version": value["version"],
        "api": value["api"], "provider_reported": reported["name"],
        "provider_kind": reported["kind"], "rows": len(value["rows"]),
        "source_archive_sha256": value["source_archive_sha256"],
        "source_archive_bytes": value["source_archive_bytes"],
        "source_revision": value["source_revision"],
        "binary_sha256": value["binary_sha256"], "binary_bytes": value["binary_bytes"],
    }


def _artifact(path: Path) -> dict[str, Any] | None:
    if not path.is_file() or path.is_symlink():
        return None
    return {"path": str(path), **meta(path)}


def _artifact_map(directory: Path) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    if not directory.is_dir():
        return result
    for path in sorted(directory.iterdir(), key=lambda item: item.name):
        item = _artifact(path)
        if item is not None:
            result[path.name] = item
    return result


def _source_snapshot() -> dict[str, Any]:
    try:
        return canonical_measure._normalized_snapshot()
    except (canonical_measure.ProviderMatrixError, OSError, ValueError, KeyError, TypeError) as error:
        fail(f"canonical source snapshot failed: {error}")


def _source_equal(expected: dict[str, Any], actual: dict[str, Any], label: str) -> None:
    require(expected == actual, f"{label}: source manifest changed")


def _run_process(argv: list[str], *, stdout: Path, stderr: Path, env: dict[str, str],
                 timeout: int) -> dict[str, Any]:
    start = _now()
    process: subprocess.Popen[bytes] | None = None
    timed_out = False
    termination: str | None = None
    try:
        stdout.parent.mkdir(parents=True, exist_ok=True)
        stderr.parent.mkdir(parents=True, exist_ok=True)
        with stdout.open("xb") as out, stderr.open("xb") as err:
            try:
                process = subprocess.Popen(
                    argv, cwd=REPO, env=env, stdin=subprocess.DEVNULL,
                    stdout=out, stderr=err, start_new_session=True,
                )
            except OSError as error:
                return {
                    "argv": argv, "started_utc": start, "finished_utc": _now(),
                    "pid": None, "process_group_id": None, "new_session": True,
                    "exit_code": None, "timed_out": False, "termination": None,
                    "launch_error": f"{type(error).__name__}: {error}",
                }
            try:
                process.wait(timeout=timeout)
            except subprocess.TimeoutExpired:
                timed_out = True
                termination = "SIGTERM"
                try:
                    os.killpg(process.pid, signal.SIGTERM)
                except ProcessLookupError:
                    pass
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    termination = "SIGKILL"
                    try:
                        os.killpg(process.pid, signal.SIGKILL)
                    except ProcessLookupError:
                        pass
                    process.wait()
    except FileExistsError:
        # Existing raw paths indicate an attempt collision.  Keep this a hard
        # refusal instead of misclassifying it as an unavailable observer.
        raise
    except OSError as error:
        return {
            "argv": argv, "started_utc": start, "finished_utc": _now(),
            "pid": None, "process_group_id": None, "new_session": True,
            "exit_code": None, "timed_out": False, "termination": None,
            "launch_error": f"{type(error).__name__}: {error}",
        }
    return {
        "argv": argv, "started_utc": start, "finished_utc": _now(),
        "pid": process.pid if process is not None else None,
        "process_group_id": process.pid if process is not None else None,
        "new_session": True, "exit_code": process.returncode if process is not None else None,
        "timed_out": timed_out, "termination": termination,
    }


def _tool_probe(path: str) -> dict[str, Any]:
    if not Path(path).is_file() or not os.access(path, os.X_OK):
        return {"path": path, "available": False, "reason": "tool not found or not executable"}
    try:
        completed = subprocess.run(
            [path, "--version"], cwd=REPO, env=ENV, stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE, stderr=subprocess.STDOUT, timeout=10, check=False,
        )
    except OSError as error:
        return {"path": path, "available": False, "reason": f"{type(error).__name__}: {error}"}
    except subprocess.TimeoutExpired:
        return {"path": path, "available": False, "reason": "version probe timed out"}
    output = completed.stdout.decode("utf-8", errors="replace") if completed.stdout else ""
    return {
        "path": path, "available": completed.returncode == 0,
        "exit_code": completed.returncode, "version_output": output[:4096],
    }


def _stderr(path: Path) -> str:
    if not path.is_file():
        return ""
    return path.read_text(encoding="utf-8", errors="replace")


def _perf_failure_status(process: dict[str, Any], stderr: Path, *, classify_pmu: bool = True) -> tuple[str, str | None]:
    if process.get("launch_error"):
        return "unavailable", "perf observer could not be launched"
    if process.get("timed_out"):
        return "failed", "perf process group timed out"
    if process.get("exit_code") == 0:
        return "pass", None
    lower = _stderr(stderr).lower()
    if classify_pmu and any(marker in lower for marker in (
        "no permission to enable", "event syntax error", "unknown tracepoint",
        "failed to open event", "failed to parse event", "invalid or unsupported event",
        "cannot find pmu", "access to performance monitoring",
        "perf_event_open", "kernel.perf_event_paranoid",
    )):
        return "unavailable", "perf PMU/event unavailable"
    return "failed", "perf observer or target exited nonzero"


def _strace_failure_status(process: dict[str, Any], stderr: Path) -> tuple[str, str | None]:
    if process.get("launch_error"):
        return "unavailable", "strace observer could not be launched"
    if process.get("timed_out"):
        return "failed", "strace process group timed out"
    if process.get("exit_code") == 0:
        return "pass", None
    lower = _stderr(stderr).lower()
    if any(marker in lower for marker in (
        "strace: attach:", "strace: ptrace(", "strace: ptrace ",
        "strace: invalid system call", "strace: unknown syscall",
        "ptrace(ptrace_traceme", "ptrace(ptrace_seize",
    )):
        return "unavailable", "strace/ptrace setup or syscall support unavailable"
    return "failed", "strace observer or target exited nonzero"


def _perf_script_failure_status(process: dict[str, Any], stderr: Path) -> tuple[str, str | None]:
    """Classify post-processing failures separately from a recorded target."""

    if process.get("launch_error"):
        return "unavailable", "perf script could not be launched"
    if process.get("timed_out"):
        return "failed", "perf script process timed out"
    if process.get("exit_code") == 0:
        return "pass", None
    lower = _stderr(stderr).lower()
    if any(marker in lower for marker in (
        "failed to open", "cannot open", "no samples", "permission denied",
        "data file is corrupted", "bad event",
    )):
        return "unavailable", "perf script stack export unavailable"
    return "failed", "perf script exited nonzero"


def _check_perf_parse(parsed: dict[str, Any], *, expected_events: Iterable[str]) -> str | None:
    if parsed["unparsed_lines"]:
        return f"perf stat has unexplained rows: {parsed['unparsed_lines'][:3]}"
    if not parsed["events"]:
        return "perf stat produced no event rows"
    unsupported = [
        name for name, item in parsed["events"].items()
        if isinstance(item.get("value_text"), str)
        and "not supported" in item["value_text"].lower()
    ]
    if unsupported:
        return f"UNAVAILABLE: perf events unsupported: {unsupported}"
    missing = [event for event in expected_events if event not in parsed["events"]]
    if missing:
        return f"perf stat omitted requested events: {missing}"
    return None


def _check_strace_parse(parsed: dict[str, Any]) -> str | None:
    if parsed["unparsed_lines"]:
        return f"strace summary has unexplained rows: {parsed['unparsed_lines'][:3]}"
    total = parsed.get("reported_total")
    if total is not None and not isinstance(total.get("calls"), (int, float)):
        return "strace reported total calls are malformed"
    if total is not None and isinstance(total.get("calls"), (int, float)):
        if int(total["calls"]) != int(parsed["total_calls"]):
            return "strace reported total calls do not equal parsed syscall rows"
    return None


def _write_started(path: Path, *, role: str, argv: list[str], source: dict[str, Any],
                   metadata: dict[str, Any], env: dict[str, str]) -> None:
    _write_new(path, {
        "schema": "docx-edit-provider-profile-process-start-v1",
        "role": role, "argv": argv, "cwd": str(REPO), "new_session": True,
        "environment": {key: env.get(key) for key in ("RUSTUP_TOOLCHAIN", "CARGO_TARGET_DIR", "TMPDIR", "LC_ALL")},
        "scope": WHOLE_CHILD_SCOPE, "started_utc": _now(), "source_before": source,
        "metadata": metadata,
    })


def _write_terminal(path: Path, *, role: str, process: dict[str, Any], source_before: dict[str, Any],
                    source_after: dict[str, Any], status: str, reason: str | None,
                    directory: Path) -> None:
    _write_new(path, {
        "schema": "docx-edit-provider-profile-process-terminal-v1",
        "role": role, "status": status, "reason": reason, "process": process,
        "source_before": source_before, "source_after": source_after,
        "source_unchanged": source_before == source_after, "finished_utc": _now(),
        "artifacts": _artifact_map(directory),
    })


def _attempt_observer(*, tool: str, name: str, binary: Path, api: str, provider: str,
                     args: argparse.Namespace, directory: Path, revision: str,
                     env: dict[str, str], source_before: dict[str, Any],
                     expected_events: Iterable[str] | None = None) -> dict[str, Any]:
    observer_output = directory / f"{name}.csv" if tool == "perf" else directory / f"{name}.summary.txt"
    report = directory / f"{name}.report.json"
    stdout, stderr = directory / f"{name}.stdout", directory / f"{name}.stderr"
    started, terminal = directory / f"{name}.started.json", directory / f"{name}.terminal.json"
    events = tuple(expected_events or FULL_EVENTS)
    argv = _observer_command(tool, binary, api, provider, args.samples, args.warmup,
                             revision, report, args.cpu, observer_output, events)
    _write_started(
        started, role=f"{tool}-{name}", argv=argv, source=source_before, env=env,
        metadata={"api": api, "provider": provider, "samples": args.samples,
                  "warmup": args.warmup, "events": list(events) if tool == "perf" else None},
    )
    process = _run_process(argv, stdout=stdout, stderr=stderr, env=env, timeout=args.timeout)
    source_after = _source_snapshot()
    _source_equal(source_before, source_after, f"{tool} source after {name}")
    parsed: dict[str, Any] | None = None
    report_meta: dict[str, Any] | None = None
    if tool == "perf" and observer_output.is_file():
        parsed = parse_perf_text(observer_output.read_text(encoding="utf-8", errors="replace"))
    elif tool == "strace" and observer_output.is_file():
        parsed = parse_strace_text(observer_output.read_text(encoding="utf-8", errors="replace"))
    if tool == "perf":
        status, reason = _perf_failure_status(process, stderr, classify_pmu=True)
    else:
        status, reason = _strace_failure_status(process, stderr)
    if status == "pass" and parsed is None:
        status, reason = "failed", f"{tool} observer produced no summary"
    if status == "pass" and tool == "perf" and parsed is not None:
        parse_error = _check_perf_parse(parsed, expected_events=events)
        if parse_error is not None:
            status = "unavailable" if parse_error.startswith("UNAVAILABLE:") else "failed"
            reason = parse_error.removeprefix("UNAVAILABLE: ").strip()
    if status == "pass" and tool == "strace" and parsed is not None:
        parse_error = _check_strace_parse(parsed)
        if parse_error is not None:
            status, reason = "failed", parse_error
    if status == "pass" and not report.is_file():
        status, reason = "failed", "target produced no report under observer"
    if status == "pass":
        try:
            report_meta = _provider_report_identity(
                report, provider, api, binary=binary, revision=revision,
                samples=args.samples, warmup=args.warmup,
            )
        except ProfileError as error:
            status, reason = "failed", str(error)
    _write_terminal(
        terminal, role=f"{tool}-{name}", process=process, source_before=source_before,
        source_after=source_after, status=status, reason=reason, directory=directory,
    )
    return {
        "tool": tool, "status": {"pass": "available"}.get(status, status),
        "reason": reason, "process": process, "parsed": parsed,
        "report": report_meta, "selected_events": list(events) if tool == "perf" else None,
        "started": {"path": str(started), **meta(started)},
        "terminal": {"path": str(terminal), **meta(terminal)},
        "artifacts": _artifact_map(directory),
    }


def _run_perf(binary: Path, api: str, provider: str, args: argparse.Namespace,
              directory: Path, revision: str, env: dict[str, str],
              source_before: dict[str, Any], *, tools_available: bool) -> dict[str, Any]:
    if not tools_available:
        return {"tool": "perf", "status": "unavailable", "reason": "perf or taskset unavailable",
                "attempts": [], "selected_attempt": None, "artifacts": _artifact_map(directory)}
    full = _attempt_observer(
        tool="perf", name="perf-full", binary=binary, api=api, provider=provider,
        args=args, directory=directory, revision=revision, env=env,
        source_before=source_before, expected_events=FULL_EVENTS,
    )
    attempts = [full]
    if full["status"] == "unavailable":
        core = _attempt_observer(
            tool="perf", name="perf-core", binary=binary, api=api, provider=provider,
            args=args, directory=directory, revision=revision, env=env,
            source_before=source_before, expected_events=CORE_EVENTS,
        )
        attempts.append(core)
    selected = attempts[-1]
    selected_name = "full" if selected is full else "core"
    if selected["status"] == "available" and selected_name == "core":
        status = "degraded"
    else:
        status = selected["status"]
    return {
        "tool": "perf", "status": status,
        "reason": selected.get("reason"), "selected_attempt": selected_name,
        "selected_events": selected.get("selected_events"), "attempts": attempts,
        "parsed": selected.get("parsed"), "report": selected.get("report"),
        "artifacts": _artifact_map(directory),
    }


def _run_strace(binary: Path, api: str, provider: str, args: argparse.Namespace,
                directory: Path, revision: str, env: dict[str, str],
                source_before: dict[str, Any], *, tools_available: bool) -> dict[str, Any]:
    if not tools_available:
        return {"tool": "strace", "status": "unavailable", "reason": "strace or taskset unavailable",
                "process": None, "parsed": None, "report": None, "artifacts": _artifact_map(directory)}
    return _attempt_observer(
        tool="strace", name="strace", binary=binary, api=api, provider=provider,
        args=args, directory=directory, revision=revision, env=env,
        source_before=source_before,
    )


def _run_record(binary: Path, api: str, provider: str, args: argparse.Namespace,
                directory: Path, revision: str, env: dict[str, str],
                source_before: dict[str, Any], *, tools_available: bool) -> dict[str, Any]:
    if provider != "owned" or not args.record_owned:
        return {"status": "skipped", "reason": "perfrecord is bounded to the selected owned route"}
    if not tools_available:
        return {"status": "unavailable", "reason": "perf or taskset unavailable", "record": None, "script": None}
    data = directory / "perf-record.data"
    report = directory / "perf-record.report.json"
    stdout, stderr = directory / "perf-record.stdout", directory / "perf-record.stderr"
    started, terminal = directory / "perf-record.started.json", directory / "perf-record.terminal.json"
    argv = _record_command(binary, api, args.record_samples, args.record_warmup,
                           revision, report, args.cpu, data, args.record_frequency)
    _write_started(
        started, role="owned-perf-record", argv=argv, source=source_before, env=env,
        metadata={"api": api, "provider": "owned", "samples": args.record_samples,
                  "warmup": args.record_warmup, "frequency_hz": args.record_frequency,
                  "event": "cycles:u", "callgraph": "dwarf"},
    )
    process = _run_process(argv, stdout=stdout, stderr=stderr, env=env, timeout=args.timeout)
    source_after = _source_snapshot()
    _source_equal(source_before, source_after, "perf record source after")
    status, reason = _perf_failure_status(process, stderr, classify_pmu=True)
    report_meta = None
    if status == "pass" and not data.is_file():
        status, reason = "failed", "perf record exited successfully without perf.data"
    if status == "pass" and not report.is_file():
        status, reason = "failed", "owned workload produced no report under perf record"
    if status == "pass":
        try:
            report_meta = _provider_report_identity(
                report, provider, api, binary=binary, revision=revision,
                samples=args.record_samples, warmup=args.record_warmup,
            )
        except ProfileError as error:
            status, reason = "failed", str(error)
    _write_terminal(
        terminal, role="owned-perf-record", process=process,
        source_before=source_before, source_after=source_after,
        status=status, reason=reason, directory=directory,
    )
    record = {
        "status": {"pass": "available"}.get(status, status), "reason": reason,
        "process": process, "report": report_meta,
        "data": _artifact(data), "started": {"path": str(started), **meta(started)},
        "terminal": {"path": str(terminal), **meta(terminal)},
    }
    result: dict[str, Any] = {"status": record["status"], "reason": reason, "record": record, "script": None}
    if status != "pass":
        result["artifacts"] = _artifact_map(directory)
        return result

    script_stdout, script_stderr = directory / "perf-script.txt", directory / "perf-script.stderr"
    script_started, script_terminal = directory / "perf-script.started.json", directory / "perf-script.terminal.json"
    script_argv = _script_command(data, args.cpu)
    _write_started(
        script_started, role="owned-perf-script", argv=script_argv, source=source_after, env=env,
        metadata={"input": {"path": str(data), **meta(data)}, "scope": STACK_SCOPE},
    )
    script_process = _run_process(script_argv, stdout=script_stdout, stderr=script_stderr,
                                  env=env, timeout=args.timeout)
    script_source_after = _source_snapshot()
    _source_equal(source_after, script_source_after, "perf script source after")
    script_status, script_reason = _perf_script_failure_status(script_process, script_stderr)
    parsed = None
    stack_summary = None
    folded = directory / "perf-folded.txt"
    summary_path = directory / "perf-stack-summary.json"
    if script_status == "pass":
        parsed = parse_perf_script_text(script_stdout.read_text(encoding="utf-8", errors="replace"))
        if parsed["unparsed_lines"]:
            script_status, script_reason = "failed", f"perf script has unexplained lines: {parsed['unparsed_lines'][:3]}"
        elif parsed["sample_count"] == 0:
            script_status, script_reason = "unavailable", "perf script exported no symbolized samples (DWARF attribution unavailable)"
        else:
            stack_summary = summarize_perf_stacks(parsed)
            stack_summary["report_identity"] = {
                "api": api, "provider": "owned", "source_revision": revision,
                "binary_sha256": sha(binary), "binary_bytes": binary.stat().st_size,
                "samples": args.record_samples, "warmup": args.record_warmup,
            }
            with folded.open("x", encoding="utf-8") as stream:
                for item in stack_summary["folded_stacks"]:
                    stream.write(f"{item['stack']} {item['period']}\n")
            _write_new(summary_path, stack_summary)
    _write_terminal(
        script_terminal, role="owned-perf-script", process=script_process,
        source_before=source_after, source_after=script_source_after,
        status=script_status, reason=script_reason, directory=directory,
    )
    result["script"] = {
        "status": script_status, "reason": script_reason, "process": script_process,
        "parsed": parsed,
        "stack_summary": None if not summary_path.is_file() else {"path": str(summary_path), **meta(summary_path)},
        "folded": _artifact(folded),
        "started": {"path": str(script_started), **meta(script_started)},
        "terminal": {"path": str(script_terminal), **meta(script_terminal)},
    }
    result["status"] = script_status if script_status != "pass" else "available"
    result["reason"] = script_reason
    result["artifacts"] = _artifact_map(directory)
    return result


def _profile_custody(build_path: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    """Authenticate after-normal retained build, protocol, source, and helpers."""

    build_path = build_path.resolve()
    require(build_path.is_file() and not build_path.is_symlink(), f"{build_path}: build receipt is not regular")
    try:
        builds = canonical_measure.load_builds(build_path.parent)
        protocol, protocol_digest = canonical_measure._load_protocol(builds)
    except (canonical_measure.ProviderMatrixError, OSError, ValueError, KeyError, TypeError) as error:
        fail(f"canonical build/protocol custody validation failed: {error}")
    build = builds["after-normal"]
    require(Path(build["path"]).resolve() == build_path,
            f"{build_path}: profile requires canonical after-normal receipt")
    binary = Path(build["binary"]["path"])
    require(binary.is_file() and not binary.is_symlink() and os.access(binary, os.X_OK),
            f"{binary}: retained executable is not regular/executable")
    require(sha(binary) == build["binary"]["sha256"] and binary.stat().st_size == build["binary"]["bytes"],
            f"{build_path}: retained binary digest/size differs")
    require(build["phase"] == PHASE and build["role"] == ROLE, f"{build_path}: build phase/role differs")
    protocol_build = protocol["builds"]["after-normal"]
    require(protocol_build["receipt_sha256"] == build["receipt_sha256"]
            and protocol_build["binary"] == build["binary"]
            and protocol_build["source"] == build["source"],
            f"{build_path}: protocol does not bind after-normal receipt")
    custody = {
        "validator": "measure.load_builds + measure._load_protocol + measure.validate_report",
        "phase": PHASE, "role": ROLE,
        "build": {
            "path": build["path"], "receipt_sha256": build["receipt_sha256"],
            "binary": build["binary"], "source": build["source"],
            "gate": build["gate"], "git_revision": build["git_revision"],
        },
        "protocol": {
            "path": str(ROOT / "protocol.json"), "sha256": protocol_digest,
            "schema": protocol["schema"], "version": protocol["version"],
            "change": protocol["change"], "source": protocol["source"],
            "builds": protocol["builds"],
        },
        "driver_bindings": {name: sha(ROOT / name) for name in canonical_measure.DRIVER_FILES},
        "harness_binding": protocol["shared_harness"],
        "profile": {"path": str(Path(__file__).resolve()), "sha256": sha(Path(__file__)),
                    "bytes": Path(__file__).stat().st_size},
    }
    return builds, protocol, custody


def _prepare_scratch(attempt: str, label: str) -> tuple[Path, Path]:
    require(LABEL_RE.fullmatch(attempt) is not None, f"unsafe profile attempt: {attempt!r}")
    require(LABEL_RE.fullmatch(label) is not None, f"unsafe profile label: {label!r}")
    managed = TEMP / "managed"
    if managed.exists():
        require(managed.is_dir() and not managed.is_symlink(), f"managed scratch owner is unsafe: {managed}")
    managed.mkdir(parents=True, exist_ok=True)
    run_root = managed / attempt / label
    require(not run_root.exists(), f"refusing to reuse private profile scratch: {run_root}")
    run_root.parent.mkdir(parents=True, exist_ok=False)
    run_root.mkdir()
    tmp_root = run_root / "tmp"
    tmp_root.mkdir()
    return run_root, tmp_root


def _cleanup_scratch(run_root: Path, tmp_root: Path) -> dict[str, Any]:
    try:
        return canonical_measure._cleanup_private(run_root, tmp_root)
    except (canonical_measure.ProviderMatrixError, OSError) as error:
        return {
            "schema": "docx-edit-provider-private-cleanup-v1", "status": "failed",
            "root": str(run_root), "tmpdir": str(tmp_root), "removed": [],
            "remaining": [f"{type(error).__name__}: {error}"],
        }


@contextmanager
def _cpu_lock(held: bool) -> Iterator[None]:
    """Use the same lock as measure.py unless the caller already owns it."""

    if held:
        yield
        return
    path = Path(canonical_measure.CPU_LOCK)
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("a+") as stream:
        fcntl.flock(stream.fileno(), fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(stream.fileno(), fcntl.LOCK_UN)


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--build-receipt", "--build-record", dest="build_receipt", type=Path, required=True,
                        help="canonical build-after-normal.json retained receipt")
    parser.add_argument("--api", choices=APIS, action="append", required=True,
                        help="API route to observe; repeat for both routes")
    parser.add_argument("--provider", choices=PROFILE_PROVIDERS, action="append", required=True,
                        help="provider arm; repeat for selected arms")
    parser.add_argument("--samples", type=int, default=3)
    parser.add_argument("--warmup", type=int, default=1)
    parser.add_argument("--record-owned", action="store_true",
                        help="run one 199 Hz owned perf-record stack profile per API")
    parser.add_argument("--record-samples", type=int, default=PERF_RECORD_SAMPLES)
    parser.add_argument("--record-warmup", type=int, default=PERF_RECORD_WARMUP)
    parser.add_argument("--record-frequency", type=int, default=PERF_RECORD_FREQUENCY_HZ)
    parser.add_argument("--cpu", type=int, default=canonical_measure.CPU)
    parser.add_argument("--timeout", type=int, default=canonical_measure.MAX_TIMEOUT_SECONDS)
    parser.add_argument("--attempt", default="profile-r1")
    parser.add_argument("--cpu-lock-held", action="store_true",
                        help="caller already owns the canonical shared CPU lock")
    parser.add_argument("--output-dir", type=Path, required=True)
    return parser


def _validate_args(args: argparse.Namespace) -> None:
    require(args.cpu == canonical_measure.CPU,
            f"cpu must be the protocol CPU {canonical_measure.CPU}")
    require(1 <= args.samples <= canonical_measure.DRIVER_LIMITS["max_samples"],
            "samples must be in 1..10000")
    require(0 <= args.warmup <= canonical_measure.DRIVER_LIMITS["max_warmup"],
            "warmup must be in 0..1000")
    require(1 <= args.timeout <= canonical_measure.MAX_TIMEOUT_SECONDS,
            "timeout must be in 1..3600 seconds")
    require(1 <= args.record_samples <= canonical_measure.DRIVER_LIMITS["max_samples"],
            "record-samples must be in 1..10000")
    require(0 <= args.record_warmup <= canonical_measure.DRIVER_LIMITS["max_warmup"],
            "record-warmup must be in 0..1000")
    require(1 <= args.record_frequency <= 10_000, "record-frequency must be in 1..10000")
    require(len(set(args.api)) == len(args.api), "API selectors must be unique")
    require(len(set(args.provider)) == len(args.provider), "provider selectors must be unique")
    require(LABEL_RE.fullmatch(args.attempt) is not None, f"unsafe profile attempt: {args.attempt!r}")


def run(args: argparse.Namespace) -> Path:
    _validate_args(args)
    output = args.output_dir.resolve() if args.output_dir.is_absolute() else (REPO / args.output_dir).resolve()
    require(not output.exists(), f"refusing to reuse profile output directory: {output}")
    build_path = args.build_receipt.resolve()
    builds, protocol, custody = _profile_custody(build_path)
    build = builds["after-normal"]
    binary = Path(build["binary"]["path"]).resolve()
    revision = build["git_revision"]
    require(REVISION_RE.fullmatch(revision) is not None, "authenticated build revision is malformed")
    source_before = _source_snapshot()
    _source_equal(build["source"], source_before, "profile source before")
    _source_equal(protocol["source"][PHASE], source_before, "profile protocol source before")

    output.mkdir(parents=True, exist_ok=False)
    base_env = dict(ENV)
    base_env["PYTHONDONTWRITEBYTECODE"] = "1"
    tool_versions = {name: _tool_probe(path) for name, path in (
        ("taskset", TASKSET), ("perf", PERF), ("strace", STRACE),
    )}
    taskset_ok = bool(tool_versions["taskset"].get("available"))
    perf_ok = taskset_ok and bool(tool_versions["perf"].get("available"))
    strace_ok = taskset_ok and bool(tool_versions["strace"].get("available"))
    cases: list[dict[str, Any]] = []
    cleanup_failures: list[str] = []

    def collect() -> None:
        for api in args.api:
            for provider in args.provider:
                label = f"{provider}-{api}"
                case_dir = output / label
                case_dir.mkdir()
                run_root, tmp_root = _prepare_scratch(args.attempt, label)
                env = dict(base_env)
                env["TMPDIR"] = str(tmp_root)
                case_source_before = _source_snapshot()
                try:
                    perf_result = _run_perf(
                        binary, api, provider, args, case_dir, revision, env,
                        case_source_before, tools_available=perf_ok,
                    )
                    strace_result = _run_strace(
                        binary, api, provider, args, case_dir, revision, env,
                        case_source_before, tools_available=strace_ok,
                    )
                    record_result = _run_record(
                        binary, api, provider, args, case_dir, revision, env,
                        case_source_before, tools_available=perf_ok,
                    )
                    case_source_after = _source_snapshot()
                    _source_equal(case_source_before, case_source_after, f"profile source after {label}")
                finally:
                    cleanup = _cleanup_scratch(run_root, tmp_root)
                    cleanup_path = case_dir / "profile-cleanup.json"
                    _write_new(cleanup_path, cleanup)
                    cleanup_artifact = {"path": str(cleanup_path), **meta(cleanup_path)}
                if cleanup["status"] != "pass":
                    cleanup_failures.append(label)
                cases.append({
                    "attempt": args.attempt, "label": label, "api": api,
                    "provider": provider, "arm": PROFILE_TO_ARM[provider],
                    "scope": WHOLE_CHILD_SCOPE, "report_scope": REPORT_SCOPE,
                    "environment": {key: env.get(key) for key in ("RUSTUP_TOOLCHAIN", "CARGO_TARGET_DIR", "TMPDIR", "LC_ALL")},
                    "perf": perf_result, "strace": strace_result, "perf_record": record_result,
                    "source_before": case_source_before, "source_after": case_source_after,
                    "cleanup": cleanup, "cleanup_artifact": cleanup_artifact,
                })

    with _cpu_lock(args.cpu_lock_held):
        collect()

    summary = {
        "schema": SCHEMA, "version": VERSION, "created_utc": _now(),
        "phase": PHASE, "role": ROLE, "attempt": args.attempt,
        "apis": list(args.api), "providers": list(args.provider),
        "samples": args.samples, "warmup": args.warmup, "cpu": args.cpu,
        "binary": {"path": str(binary), **meta(binary)}, "source_revision": revision,
        "build_receipt": {"path": str(build_path), "sha256": sha(build_path)},
        "custody": custody,
        "protocol": {"path": str(ROOT / "protocol.json"), "sha256": sha(ROOT / "protocol.json"),
                      "schema": protocol["schema"], "change": protocol["change"]},
        "helper_bindings": {name: sha(ROOT / name) for name in canonical_measure.DRIVER_FILES},
        "tool_versions": tool_versions,
        "environment": {key: base_env.get(key) for key in ("RUSTUP_TOOLCHAIN", "CARGO_TARGET_DIR", "TMPDIR", "LC_ALL")},
        "cpu_lock": {"path": canonical_measure.CPU_LOCK, "caller_held": args.cpu_lock_held},
        "observer_scope": {"perf": WHOLE_CHILD_SCOPE, "strace": WHOLE_CHILD_SCOPE,
                            "perf_record": WHOLE_CHILD_SCOPE, "stacks": STACK_SCOPE,
                            "report_validator": "measure.validate_report", "latency": REPORT_SCOPE},
        "cases": cases,
    }
    summary_path = output / "profile-summary.json"
    _write_new(summary_path, summary)
    if cleanup_failures:
        fail(f"private scratch cleanup failed for: {', '.join(cleanup_failures)}")
    return summary_path


def main(argv: list[str] | None = None) -> int:
    try:
        args = _parser().parse_args(argv)
        print(run(args))
        return 0
    except (ProfileError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"profile.py: FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
