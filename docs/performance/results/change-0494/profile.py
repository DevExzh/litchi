#!/usr/bin/env python3
"""Run whole-child perf/strace profiles for the 0494 DOCX edit selector.

The 0494 measurement driver owns formal timing custody.  This helper is a
small, separate observer runner: it authenticates the retained normal binary,
runs one fresh process per provider and observer, preserves raw observer
output, and emits parsed counters without turning unsupported events into
zeros.  It does not build, capture the formal matrix, or mutate source files.
"""

from __future__ import annotations

import argparse
from collections import Counter
import datetime as _datetime
import json
import os
from pathlib import Path
import re
import shutil
import signal
import subprocess
import sys
from typing import Any, Iterable

sys.path.insert(0, str(Path(__file__).resolve().parent))
from support import ENV, REPO, ROOT, TEMP, meta, sha, snapshot  # noqa: E402
import measure as canonical_measure  # noqa: E402


SCHEMA = "docx-edit-provider-profile-v1"
# Report identity, route scopes, corpus identity, and limits are owned by the
# canonical measurement driver.  These aliases are only used to describe the
# retained stack report; report acceptance calls ``measure.validate_report``.
REPORT_SCHEMA = canonical_measure.SCHEMA
REPORT_VERSION = canonical_measure.VERSION
REPORT_CASE_NAME = canonical_measure.CASE
EXPECTED_SOURCE_ARCHIVE_SHA256 = canonical_measure.CORPUS["archive_sha256"]
EXPECTED_SOURCE_ARCHIVE_BYTES = canonical_measure.CORPUS["archive_bytes"]
EXPECTED_ARCHIVE_MEMBERS = canonical_measure.CORPUS["archive_member_count"]
EXPECTED_TIMING_SCOPE = canonical_measure.TIMING_SCOPE
EXPECTED_SETUP_SCOPE = canonical_measure.SETUP_SCOPE
EXPECTED_PHYSICAL_SCOPE = canonical_measure.PHYSICAL_SCOPE
EXPECTED_RANGE_SCOPE = canonical_measure.RANGE_SCOPE
EXPECTED_FILE_SCOPE = canonical_measure.FILE_SCOPE
EXPECTED_ZERO_LENGTH_SCOPE = canonical_measure.ZERO_LENGTH_SCOPE
READ_LIMIT_KEYS = tuple(canonical_measure.READ_LIMITS)
LIMIT_KEYS = tuple(canonical_measure.LIMITS)
PROVIDERS = ("owned", "file-warm", "short-read")
PROFILE_ARM_NAMES = {
    "owned": "owned",
    "file-warm": "file-warm",
    "short-read": "short",
}
FULL_EVENTS = (
    "task-clock",
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "L1-dcache-loads",
    "L1-dcache-load-misses",
    "LLC-loads",
    "LLC-load-misses",
    "minor-faults",
    "major-faults",
    "context-switches",
    "cpu-migrations",
    "page-faults",
)
CORE_EVENTS = (
    "task-clock",
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "minor-faults",
    "major-faults",
)
EVENT_RE = re.compile(r"^[A-Za-z0-9_.-]+$")
HEX40_RE = re.compile(r"^[0-9a-f]{40}$")
HEX64_RE = re.compile(r"^[0-9a-f]{64}$")
PERF_RECORD_FREQUENCY_HZ = 199
PERF_RECORD_SAMPLES = 100
PERF_RECORD_WARMUP = 3

# ``perf script`` has changed its default field order between distributions.
# Keep the header matcher permissive, but require a non-indented event line so
# a symbol containing ``cycles`` cannot start a false sample.
_PERF_SCRIPT_HEADER_RE = re.compile(
    r"^\S.*:\s+(?:(?P<period>\d+)\s+)?(?P<event>[^\s:]+(?::[^\s:]+)*):\s*$"
)
_PERF_FRAME_RE = re.compile(
    r"^\s+(?:[0-9a-f]+\s+)?(?P<symbol>[^\s(]+)(?:\+0x[0-9a-f]+)?(?:\s+\([^)]*\))?.*$",
    re.IGNORECASE,
)


class ProfileError(RuntimeError):
    """A fail-closed profile custody or parsing error."""


def fail(message: str) -> None:
    raise ProfileError(message)


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
        fail(f"refusing to replace immutable profile artifact: {path}")


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


def parse_perf_text(text: str) -> dict[str, Any]:
    """Parse GNU perf's ``-x,`` output while retaining unsupported rows."""

    events: dict[str, dict[str, Any]] = {}
    raw_lines: list[str] = []
    for raw in text.splitlines():
        line = raw.strip("\r")
        if not line or line.lstrip().startswith("#"):
            continue
        fields = [item.strip() for item in line.split(",")]
        if len(fields) < 3:
            continue
        event = fields[2]
        if not event or not EVENT_RE.fullmatch(event):
            continue
        value_text = fields[0]
        running_text = fields[3] if len(fields) > 3 else ""
        record = {
            "value": _numeric(value_text),
            "value_text": value_text,
            "unit": fields[1] or None,
            "event": event,
            "running_percent": _numeric(running_text),
            "raw": line,
        }
        events[event] = record
        raw_lines.append(line)

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
    return {"events": events, "derived": derived, "raw_lines": raw_lines}


_STRACE_RE = re.compile(
    r"^\s*(?P<pct>\S+)\s+(?P<seconds>\S+)\s+(?P<usec>\S+)\s+"
    r"(?P<calls>\S+)\s+(?P<errors>\S+)\s+(?P<syscall>\S+)\s*$"
)


def parse_strace_text(text: str) -> dict[str, Any]:
    """Parse ``strace -f -c`` output, retaining the unparsed lines."""

    syscalls: dict[str, dict[str, Any]] = {}
    unparsed: list[str] = []
    total: dict[str, Any] | None = None
    for raw in text.splitlines():
        line = raw.rstrip("\r")
        if not line.strip() or line.lstrip().startswith("-") or line.lstrip().startswith("% time"):
            continue
        if line.strip().endswith(" total"):
            parts = line.split()
            if len(parts) >= 5:
                total = {"raw": line, "calls": _numeric(parts[-3]), "errors": _numeric(parts[-2])}
                continue
        match = _STRACE_RE.match(line)
        if match is None:
            unparsed.append(line)
            continue
        fields = match.groupdict()
        syscall = fields.pop("syscall")
        record = {key: _numeric(value) for key, value in fields.items()}
        record["raw"] = line
        syscalls[syscall] = record
    call_total = sum(
        int(item["calls"])
        for item in syscalls.values()
        if isinstance(item.get("calls"), (int, float))
    )
    error_total = sum(
        int(item["errors"])
        for item in syscalls.values()
        if isinstance(item.get("errors"), (int, float))
    )
    return {
        "syscalls": syscalls,
        "total_calls": call_total,
        "total_errors": error_total,
        "reported_total": total,
        "unparsed_lines": unparsed,
    }


def _perf_frame_symbol(line: str) -> str | None:
    """Extract one symbol from an indented ``perf script`` frame line."""

    if not line or not line[0].isspace():
        return None
    stripped = line.strip()
    if not stripped:
        return None
    # The normal form is ``address symbol+offset (dso)``.  Keep unknown and
    # kernel frames too; dropping them would bias the class totals toward Rust
    # frames merely because symbolization was incomplete.
    fields = stripped.split()
    if len(fields) >= 2 and re.fullmatch(r"(?:0x)?[0-9a-f]+", fields[0], re.IGNORECASE):
        symbol = fields[1]
    else:
        symbol = fields[0]
    symbol = symbol.split("+0x", 1)[0]
    return symbol or None


def parse_perf_script_text(text: str) -> dict[str, Any]:
    """Parse a retained ``perf script`` export into weighted call stacks.

    Perf normally prints frames leaf first.  Each returned sample preserves
    that order and carries the sample period when the selected output fields
    include it; otherwise period one is used.  Period is a sampling weight,
    not elapsed time for a Rust phase.
    """

    samples: list[dict[str, Any]] = []
    unparsed: list[str] = []
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
        header = _PERF_SCRIPT_HEADER_RE.match(line)
        if header is not None and "cycles" in header.group("event"):
            flush()
            period_text = header.group("period")
            period = int(period_text) if period_text is not None else 1
            current = {
                "header": line,
                "period": period,
                "frames": [],
            }
            continue
        if current is None:
            # ``perf script --header`` metadata and lost-event diagnostics are
            # retained for audit, but they are not samples.
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
    }


def classify_perf_stack(frames: Iterable[str]) -> str:
    """Classify a sampled stack without pretending samples are phase timers."""

    frame_list = list(frames)
    has_run_sample = any("docx_edit_provider::run_sample" in frame for frame in frame_list)
    has_publish = any("publish_docx_source_edit" in frame for frame in frame_list)
    has_prepare = any("docx_edit_provider::prepare" in frame for frame in frame_list)
    has_output_oracle = any(
        marker in frame
        for frame in frame_list
        for marker in ("verify_docx_source_edit_output", "sha256_hex")
    )
    if has_run_sample and has_publish:
        # A publish ancestor under run_sample is the useful sampled subset for
        # the operation body.  It is still a statistical call-stack count.
        return "run_sample_publish_ancestor"
    if has_run_sample and has_output_oracle:
        # The output oracle runs after the lifecycle clock in the current
        # implementation, despite sharing the run_sample function.
        return "run_sample_output_oracle"
    if has_run_sample:
        return "run_sample_setup_or_teardown"
    if has_prepare:
        return "preflight"
    return "process_setup_or_unclassified"


def summarize_perf_stacks(parsed: dict[str, Any]) -> dict[str, Any]:
    """Create deterministic folded stacks and phase marker counts."""

    samples = parsed.get("samples")
    if not isinstance(samples, list):
        fail("perf script parser did not return samples")
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
        weight = int(item.get("period", 1))
        if weight < 1:
            weight = 1
        # perf's frame order is leaf -> caller; folded stack tools conventionally
        # consume root -> leaf, hence the reverse here.
        folded[";".join(reversed(frames))] += weight
        stack_class = classify_perf_stack(frames)
        classes[stack_class] += weight
        class_samples[stack_class] += 1
        class_leafs.setdefault(stack_class, Counter())[frames[0]] += weight
        if any("docx_edit_provider::run_sample" in frame for frame in frames):
            run_sample_period += weight
        if stack_class == "run_sample_publish_ancestor":
            publish_period += weight

    total_period = sum(classes.values())
    class_rows = []
    for name in sorted(classes):
        period = classes[name]
        class_rows.append({
            "class": name,
            "sample_count": class_samples[name],
            "period": period,
            "share_of_all_period": period / total_period if total_period else None,
            "share_of_run_sample_period": (
                period / run_sample_period if run_sample_period else None
            ),
            "top_leaf_symbols": [
                {"symbol": symbol, "period": value}
                for symbol, value in sorted(
                    class_leafs[name].items(), key=lambda pair: (-pair[1], pair[0])
                )[:20]
            ],
        })
    return {
        "schema": "docx-edit-provider-perf-stacks-v1",
        "sample_count": len(samples),
        "total_period": total_period,
        "run_sample_period": run_sample_period,
        "run_sample_publish_period": publish_period,
        "classes": class_rows,
        "folded_stack_count": len(folded),
        "folded_stacks": [
            {"stack": stack, "period": period}
            for stack, period in sorted(folded.items(), key=lambda pair: (-pair[1], pair[0]))
        ],
        "interpretation": (
            "Periods are statistical perf samples. A publish ancestor under "
            "run_sample is an actionable callchain subset, not an exact phase "
            "duration. Whole-child setup and report/oracle work remain in scope."
        ),
    }


def provider_args(provider: str) -> list[str]:
    if provider == "owned":
        return ["--provider", "owned"]
    if provider == "file-warm":
        return ["--provider", "file"]
    if provider == "short-read":
        return ["--provider", "short", "--short-range", "4096"]
    fail(f"unknown provider {provider!r}; expected {', '.join(PROVIDERS)}")


def _profile_custody(build_path: Path) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    """Load the canonical build pair and frozen protocol as one binding."""

    build_path = build_path.resolve()
    if not build_path.is_file() or build_path.is_symlink():
        fail(f"{build_path}: retained build receipt is not a regular file")
    try:
        builds = canonical_measure.load_builds(build_path.parent)
        protocol, protocol_digest = canonical_measure._load_protocol(builds)
    except (canonical_measure.ProviderMatrixError, OSError, ValueError) as error:
        fail(f"canonical custody validation failed: {error}")
    normal = builds["normal"]
    if Path(normal["path"]).resolve() != build_path:
        fail(f"{build_path}: profile requires the canonical retained normal build")
    binary = normal["binary"]
    binary_path = Path(binary["path"])
    if not binary_path.is_file() or binary_path.is_symlink() or not os.access(binary_path, os.X_OK):
        fail(f"{binary_path}: retained binary is not a regular executable")
    if sha(binary_path) != binary["sha256"] or binary_path.stat().st_size != binary["bytes"]:
        fail(f"{build_path}: retained binary hash or size does not match canonical receipt")
    custody = {
        "validator": "measure.load_builds + measure._load_protocol",
        "source": normal["source"],
        "git_revision": normal["git_revision"],
        "protocol": {
            "path": str(ROOT / "protocol.json"),
            "sha256": protocol_digest,
            "schema": protocol["schema"],
            "version": protocol["version"],
            "change": protocol["change"],
            "case": protocol["case"],
            "source": protocol["source"],
            "builds": protocol["builds"],
        },
        "builds": {
            role: {
                "path": build["path"],
                "receipt_sha256": build["receipt_sha256"],
                "role": build["role"],
                "binary": build["binary"],
                "source": build["source"],
                "gate": build["gate"],
                "git_revision": build["git_revision"],
            }
            for role, build in builds.items()
        },
    }
    return builds, protocol, custody


def _prepare_private_scratch(label: str) -> tuple[Path, Path]:
    """Create a private scratch pair below the canonical managed owner."""

    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_.-]*", label):
        fail(f"unsafe profile scratch label: {label!r}")
    managed = TEMP / "managed"
    if managed.exists() and (managed.is_symlink() or not managed.is_dir()):
        fail(f"managed scratch owner is unsafe: {managed}")
    managed.mkdir(parents=True, exist_ok=True)
    attempt_parent = managed / f"profile-r1-{label}"
    run_root = attempt_parent / label
    tmp_root = run_root / "tmp"
    if attempt_parent.exists() or run_root.exists() or tmp_root.exists():
        fail(f"refusing to reuse private profile scratch: {run_root}")
    attempt_parent.mkdir()
    run_root.mkdir()
    tmp_root.mkdir()
    return run_root, tmp_root


def _cleanup_private_scratch(run_root: Path, tmp_root: Path) -> dict[str, Any]:
    """Reuse measure.py's exact managed-path cleanup receipt."""

    try:
        return canonical_measure._cleanup_private(run_root, tmp_root)
    except (canonical_measure.ProviderMatrixError, OSError) as error:
        return {
            "schema": "docx-edit-provider-private-cleanup-v1",
            "status": "failed",
            "root": str(run_root),
            "tmpdir": str(tmp_root),
            "removed": [],
            "remaining": [str(error)],
        }


def _write_cleanup_receipt(directory: Path, run_root: Path, tmp_root: Path) -> dict[str, Any]:
    cleanup = _cleanup_private_scratch(run_root, tmp_root)
    path = directory / "profile-cleanup.json"
    _write_new(path, cleanup)
    cleanup["artifact"] = {"path": str(path), **meta(path)}
    return cleanup


def _source_equal(expected: Any, actual: Any, label: str) -> None:
    if expected != actual:
        fail(f"{label}: source manifest changed")


def _canonical_source_snapshot() -> dict[str, Any]:
    try:
        return canonical_measure._normalized_snapshot()
    except (canonical_measure.ProviderMatrixError, OSError, ValueError) as error:
        fail(f"canonical source snapshot failed: {error}")


def _run_process(argv: list[str], *, stdout: Path, stderr: Path, env: dict[str, str],
                 timeout: int) -> dict[str, Any]:
    stdout.parent.mkdir(parents=True, exist_ok=True)
    start = _now()
    timed_out = False
    termination: str | None = None
    process: subprocess.Popen[bytes] | None = None
    try:
        with stdout.open("xb") as out, stderr.open("xb") as err:
            process = subprocess.Popen(
                argv,
                cwd=REPO,
                env=env,
                stdin=subprocess.DEVNULL,
                stdout=out,
                stderr=err,
                start_new_session=True,
            )
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
    except OSError as error:
        return {
            "argv": argv,
            "started_utc": start,
            "finished_utc": _now(),
            "pid": None,
            "process_group_id": None,
            "new_session": True,
            "exit_code": None,
            "timed_out": False,
            "termination": None,
            "launch_error": f"{type(error).__name__}: {error}",
        }
    return {
        "argv": argv,
        "started_utc": start,
        "finished_utc": _now(),
        "pid": process.pid if process is not None else None,
        "process_group_id": process.pid if process is not None else None,
        "new_session": True,
        "exit_code": process.returncode if process is not None else None,
        "timed_out": timed_out,
        "termination": termination,
    }


def _tool_version(path: str) -> dict[str, Any]:
    """Capture a short observer version receipt without running the target."""

    try:
        completed = subprocess.run(
            [path, "--version"], cwd=REPO, env=ENV,
            stdin=subprocess.DEVNULL, stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
            timeout=10, check=False,
        )
    except OSError as error:
        return {"path": path, "available": False, "error": f"{type(error).__name__}: {error}"}
    except subprocess.TimeoutExpired:
        return {"path": path, "available": False, "error": "version probe timed out"}
    output = completed.stdout.decode("utf-8", errors="replace") if completed.stdout else ""
    return {
        "path": path,
        "available": completed.returncode == 0,
        "exit_code": completed.returncode,
        "version_output": output[:4096],
    }


def _provider_command(binary: Path, provider: str, samples: int, warmup: int,
                      revision: str, report: Path, cpu: int) -> list[str]:
    return [
        "/usr/bin/taskset",
        "-c",
        str(cpu),
        str(binary),
        "docx-edit-provider",
        *provider_args(provider),
        "--samples",
        str(samples),
        "--warmup",
        str(warmup),
        "--source-revision",
        revision,
        "--output",
        str(report),
    ]


def _observer_command(tool: str, binary: Path, provider: str, samples: int, warmup: int,
                      revision: str, report: Path, cpu: int, observer_output: Path) -> list[str]:
    target = _provider_command(binary, provider, samples, warmup, revision, report, cpu)
    if tool == "perf":
        return [
            "/usr/bin/taskset", "-c", str(cpu), "/usr/bin/perf", "stat",
            "--no-big-num", "-x,", "-e", ",".join(FULL_EVENTS), "-o", str(observer_output), "--",
            str(binary), "docx-edit-provider", *provider_args(provider),
            "--samples", str(samples), "--warmup", str(warmup),
            "--source-revision", revision, "--output", str(report),
        ]
    if tool == "strace":
        return [
            "/usr/bin/taskset", "-c", str(cpu), "/usr/bin/strace", "-f", "-c",
            "-o", str(observer_output), "--", *target[3:],
        ]
    fail(f"unknown observer {tool}")


def _core_observer_command(binary: Path, provider: str, samples: int, warmup: int,
                           revision: str, report: Path, cpu: int, observer_output: Path) -> list[str]:
    return [
        "/usr/bin/taskset", "-c", str(cpu), "/usr/bin/perf", "stat",
        "--no-big-num", "-x,", "-e", ",".join(CORE_EVENTS), "-o", str(observer_output), "--",
        str(binary), "docx-edit-provider", *provider_args(provider),
        "--samples", str(samples), "--warmup", str(warmup),
        "--source-revision", revision, "--output", str(report),
    ]


def _report_identity(
    path: Path,
    provider: str,
    *,
    expected_binary_sha256: str,
    expected_binary_bytes: int,
    expected_source_revision: str,
    expected_samples: int,
    expected_warmup: int,
) -> dict[str, Any]:
    """Validate a child report with the canonical measurement contract."""

    if not path.is_file() or path.is_symlink():
        fail(f"{path}: observer did not produce a regular report")
    try:
        value = canonical_measure.validate_report(
            path,
            role="normal",
            arm_name=PROFILE_ARM_NAMES[provider],
            samples=expected_samples,
            warmups=expected_warmup,
            source_revision=expected_source_revision,
            binary_sha256=expected_binary_sha256,
            binary_bytes=expected_binary_bytes,
        )
    except KeyError:
        fail(f"unknown profile provider {provider!r}")
    except (canonical_measure.ProviderMatrixError, OSError, ValueError) as error:
        fail(f"{path}: canonical report validation failed: {error}")
    reported = value["provider"]
    return {
        "sha256": sha(path),
        "bytes": path.stat().st_size,
        "rows": len(value["rows"]),
        "schema": value["schema"],
        "version": value["version"],
        "case_name": value["case_name"],
        "provider_reported": reported["name"],
        "provider_kind": reported["kind"],
        "source_archive_sha256": value["source_archive_sha256"],
        "source_archive_bytes": value["source_archive_bytes"],
        "source_revision": value["source_revision"],
        "binary_sha256": value["binary_sha256"],
        "binary_bytes": value["binary_bytes"],
        "provider_match": True,
    }


def _artifact_map(paths: Iterable[Path]) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for path in paths:
        if path.is_file() and not path.is_symlink():
            result[path.name] = {"path": str(path), **meta(path)}
    return result


def _write_process_started(path: Path, *, role: str, argv: list[str], env: dict[str, str],
                           source_before: dict[str, Any], metadata: dict[str, Any]) -> None:
    """Retain the immutable command receipt before starting a process group."""

    _write_new(path, {
        "schema": "docx-edit-provider-profile-process-start-v1",
        "role": role,
        "argv": argv,
        "cwd": str(REPO),
        "environment": {
            key: env.get(key)
            for key in ("RUSTUP_TOOLCHAIN", "CARGO_TARGET_DIR", "TMPDIR", "LC_ALL")
        },
        "new_session": True,
        "process_group_policy": "the observer and target share this fresh process group; timeout kills the group",
        "started_utc": _now(),
        "source_before": source_before,
        "metadata": metadata,
    })


def _write_process_terminal(path: Path, *, role: str, process: dict[str, Any],
                            source_before: dict[str, Any], source_after: dict[str, Any],
                            status: str, reason: str | None,
                            artifacts: Iterable[Path]) -> None:
    """Retain terminal state, source custody, and hashes for one process group."""

    _write_new(path, {
        "schema": "docx-edit-provider-profile-process-terminal-v1",
        "role": role,
        "status": status,
        "reason": reason,
        "process": process,
        "source_before": source_before,
        "source_after": source_after,
        "source_unchanged": source_before == source_after,
        "finished_utc": _now(),
        "artifacts": _artifact_map(artifacts),
    })


def _run_perf(binary: Path, provider: str, args: argparse.Namespace, directory: Path,
              revision: str, env: dict[str, str]) -> dict[str, Any]:
    full_stat = directory / "perf-stat.csv"
    full_stderr = directory / "perf.stderr"
    full_report = directory / "perf-report.json"
    full = _observer_command("perf", binary, provider, args.samples, args.warmup,
                             revision, full_report, args.cpu, full_stat)
    result = _run_process(full, stdout=directory / "perf.stdout", stderr=full_stderr,
                          env=env, timeout=args.timeout)
    selected = "full"
    attempts = [{"events": list(FULL_EVENTS), "process": result}]
    # perf stat's exit code is also the target's exit code. Classify only
    # perf-specific PMU diagnostics as unavailable; ordinary target failures
    # stay failed and do not trigger the reduced-event retry.
    full_status, full_reason = _record_failure_status(
        result, full_stderr, classify_pmu=True,
    )
    if full_status == "pass" and not full_stat.is_file():
        full_status, full_reason = "failed", "perf stat exited successfully without its CSV"
    if full_status == "pass" and not full_report.is_file():
        full_status, full_reason = "failed", "target produced no report under perf stat"
    if full_status == "pass":
        selected_report = full_report
        selected_stat = full_stat
    elif full_status == "unavailable":
        # Retry a smaller event set only for a PMU/event availability failure.
        # A generic nonzero target status is retained as failed and never
        # relabelled as an unavailable profiler.
        core_stat = directory / "perf-core-stat.csv"
        core_stderr = directory / "perf-core.stderr"
        core_report = directory / "perf-core-report.json"
        core = _core_observer_command(binary, provider, args.samples, args.warmup,
                                      revision, core_report, args.cpu, core_stat)
        core_result = _run_process(core, stdout=directory / "perf-core.stdout",
                                   stderr=core_stderr, env=env, timeout=args.timeout)
        attempts.append({"events": list(CORE_EVENTS), "process": core_result})
        core_status, core_reason = _record_failure_status(
            core_result, core_stderr, classify_pmu=True,
        )
        if core_status == "pass" and not core_stat.is_file():
            core_status, core_reason = "failed", "core perf stat exited without its CSV"
        if core_status == "pass" and not core_report.is_file():
            core_status, core_reason = "failed", "target produced no report under core perf stat"
        if core_status == "pass":
            selected = "core"
            selected_report, selected_stat = core_report, core_stat
        else:
            selected = "failed" if core_status == "failed" else "unavailable"
            full_reason = core_reason
            selected_report = None
            selected_stat = None
    else:
        selected = "failed"
        selected_report = full_report if full_report.is_file() else None
        selected_stat = full_stat if full_stat.is_file() else None
    report_meta = None
    if selected_report is not None:
        try:
            report_meta = _report_identity(
                selected_report,
                provider,
                expected_binary_sha256=sha(binary),
                expected_binary_bytes=binary.stat().st_size,
                expected_source_revision=revision,
                expected_samples=args.samples,
                expected_warmup=args.warmup,
            )
        except ProfileError as error:
            full_reason = str(error)
            selected = "failed"
    parsed = None
    if selected_stat is not None:
        parsed = parse_perf_text(selected_stat.read_text(encoding="utf-8", errors="replace"))
    status = {
        "full": "available",
        "core": "degraded",
        "unavailable": "unavailable",
        "failed": "failed",
    }[selected]
    if report_meta is not None and not report_meta.get("provider_match", False):
        status = "failed"
        full_reason = "target report provider does not match requested provider"
    return {
        "tool": "perf",
        "status": status,
        "reason": full_reason,
        "selected_events": (
            list(FULL_EVENTS) if selected == "full" else
            list(CORE_EVENTS) if selected == "core" else None
        ),
        "selected_attempt": selected,
        "attempts": attempts,
        "report": report_meta,
        "parsed": parsed,
        "artifacts": _artifact_map(
            path for path in directory.iterdir() if path.name.startswith("perf")
        ),
    }


def _run_strace(binary: Path, provider: str, args: argparse.Namespace, directory: Path,
                revision: str, env: dict[str, str]) -> dict[str, Any]:
    summary = directory / "strace-summary.txt"
    report = directory / "strace-report.json"
    command = _observer_command("strace", binary, provider, args.samples, args.warmup,
                                revision, report, args.cpu, summary)
    process = _run_process(command, stdout=directory / "strace.stdout",
                           stderr=directory / "strace.stderr", env=env,
                           timeout=args.timeout)
    parsed = None
    report_meta = None
    reason: str | None = None
    if summary.is_file():
        parsed = parse_strace_text(summary.read_text(encoding="utf-8", errors="replace"))
    if report.is_file() and process.get("exit_code") == 0:
        try:
            report_meta = _report_identity(
                report,
                provider,
                expected_binary_sha256=sha(binary),
                expected_binary_bytes=binary.stat().st_size,
                expected_source_revision=revision,
                expected_samples=args.samples,
                expected_warmup=args.warmup,
            )
        except ProfileError as error:
            reason = str(error)
    if process.get("exit_code") != 0 or process.get("timed_out") or process.get("launch_error"):
        status, failure_reason = _strace_failure_status(process, directory / "strace.stderr")
        reason = reason or failure_reason
    elif parsed is None or report_meta is None:
        status = "failed"
        reason = reason or "strace output or target report is missing"
    elif not report_meta.get("provider_match", False):
        status = "failed"
        reason = "target report provider does not match requested provider"
    else:
        status = "available"
    return {
        "tool": "strace",
        "status": status,
        "reason": reason,
        "process": process,
        "report": report_meta,
        "parsed": parsed,
        "artifacts": _artifact_map(path for path in directory.iterdir() if path.name.startswith("strace")),
    }


def _record_command(binary: Path, samples: int, warmup: int, revision: str,
                    report: Path, cpu: int, data: Path, frequency_hz: int) -> list[str]:
    return [
        "/usr/bin/taskset", "-c", str(cpu), "/usr/bin/perf", "record",
        "-F", str(frequency_hz), "-e", "cycles:u", "--call-graph", "dwarf",
        "-o", str(data), "--", str(binary), "docx-edit-provider",
        "--provider", "owned", "--samples", str(samples), "--warmup", str(warmup),
        "--source-revision", revision, "--output", str(report),
    ]


def _script_command(data: Path, cpu: int) -> list[str]:
    # Explicit fields make period parsing stable across perf versions while
    # retaining the event and symbolized callchain needed for classification.
    return [
        "/usr/bin/taskset", "-c", str(cpu), "/usr/bin/perf", "script", "--header",
        "--demangle", "-F", "comm,pid,tid,cpu,time,period,event,ip,sym,dso",
        "-i", str(data),
    ]


def _record_failure_status(process: dict[str, Any], stderr: Path, *, classify_pmu: bool = True) -> tuple[str, str | None]:
    """Classify observer launch/terminal failures without hiding target errors."""

    if process.get("launch_error"):
        return "unavailable", "perf record could not be launched"
    if process.get("timed_out"):
        return "failed", "perf record process group timed out"
    if process.get("exit_code") == 0:
        return "pass", None
    diagnostic = ""
    if stderr.is_file():
        diagnostic = stderr.read_text(encoding="utf-8", errors="replace")
    lower = diagnostic.lower()
    # PMU access/configuration failures are a real unavailable outcome.  A
    # generic nonzero status remains failed so a normal benchmark failure is
    # never relabelled as an unavailable profiler.
    if classify_pmu and any(marker in lower for marker in (
        # These are perf event setup diagnostics. Keep generic "permission
        # denied" out: a target is allowed to report that ordinary error.
        "no permission to enable", "event syntax error",
        "unknown tracepoint", "failed to open event", "failed to parse event",
        "invalid or unsupported event", "event not supported", "cannot find pmu",
    )):
        return "unavailable", "perf PMU/event unavailable"
    return "failed", "perf record or target exited nonzero"


def _strace_failure_status(process: dict[str, Any], stderr: Path) -> tuple[str, str | None]:
    """Separate strace setup failures from failures emitted by the target.

    A nonzero strace wrapper exit alone is insufficient evidence that ptrace is
    unavailable: the traced DOCX process can fail for an ordinary application
    reason and its stderr is forwarded by strace. Only diagnostics that name
    the strace/ptrace setup or unsupported syscall selection are classified as
    an unavailable observer. Target permission errors remain ``failed``.
    """

    if process.get("launch_error"):
        return "unavailable", "strace could not be launched"
    if process.get("timed_out"):
        return "failed", "strace process group timed out"
    if process.get("exit_code") == 0:
        return "pass", None
    diagnostic = ""
    if stderr.is_file():
        diagnostic = stderr.read_text(encoding="utf-8", errors="replace")
    lower = diagnostic.lower()
    observer_markers = (
        "strace: attach:",
        "strace: ptrace(",
        "strace: ptrace ",
        "strace: invalid system call",
        "strace: unknown syscall",
        "ptrace(ptrace_traceme",
        "ptrace(ptrace_seize",
    )
    if any(marker in lower for marker in observer_markers):
        return "unavailable", "strace/ptrace setup or syscall support unavailable"
    return "failed", "strace or target exited nonzero"


def _run_owned_record(binary: Path, args: argparse.Namespace, directory: Path,
                      revision: str, env: dict[str, str],
                      source_before: dict[str, Any]) -> dict[str, Any]:
    """Run and retain one owned-source DWARF stack profile."""

    directory.mkdir(parents=True, exist_ok=False)
    data = directory / "perf.data"
    report = directory / "perf-record-report.json"
    record_stdout = directory / "perf-record.stdout"
    record_stderr = directory / "perf-record.stderr"
    record_started = directory / "perf-record.started.json"
    record_terminal = directory / "perf-record.terminal.json"
    if shutil.which("/usr/bin/perf") is None:
        _write_new(directory / "perf-record.unavailable.json", {
            "schema": "docx-edit-provider-perf-record-unavailable-v1",
            "status": "unavailable",
            "reason": "tool not found",
            "created_utc": _now(),
            "source_before": source_before,
        })
        return {
            "status": "unavailable",
            "reason": "tool not found",
            "artifacts": _artifact_map(directory.iterdir()),
        }

    record_argv = _record_command(
        binary, args.record_samples, args.record_warmup, revision, report, args.cpu,
        data, args.record_frequency,
    )
    _write_process_started(
        record_started,
        role="owned-perf-record",
        argv=record_argv,
        env=env,
        source_before=source_before,
        metadata={
            "binary": {"path": str(binary), **meta(binary)},
            "source_revision": revision,
            "samples": args.record_samples,
            "warmup": args.record_warmup,
            "frequency_hz": args.record_frequency,
            "event": "cycles:u",
            "callgraph": "dwarf",
            "scope": "whole child including setup, warmups, measured rows, and report serialization",
        },
    )
    record_process = _run_process(
        record_argv, stdout=record_stdout, stderr=record_stderr, env=env,
        timeout=args.timeout,
    )
    source_after_record = snapshot()
    _source_equal(source_before, source_after_record, "owned perf record source after")
    record_status, record_reason = _record_failure_status(record_process, record_stderr)
    if record_status == "pass" and not data.is_file():
        record_status, record_reason = "failed", "perf record exited successfully without perf.data"
    if record_status == "pass" and not report.is_file():
        record_status, record_reason = "failed", "owned workload produced no report"
    report_meta = None
    if record_status == "pass":
        try:
            report_meta = _report_identity(
                report,
                "owned",
                expected_binary_sha256=sha(binary),
                expected_binary_bytes=binary.stat().st_size,
                expected_source_revision=revision,
                expected_samples=args.record_samples,
                expected_warmup=args.record_warmup,
            )
        except ProfileError as error:
            record_status, record_reason = "failed", str(error)
        else:
            if not report_meta.get("provider_match", False):
                record_status, record_reason = "failed", "target report provider does not match requested provider"
    _write_process_terminal(
        record_terminal,
        role="owned-perf-record",
        process=record_process,
        source_before=source_before,
        source_after=source_after_record,
        status=record_status,
        reason=record_reason,
        artifacts=(record_stdout, record_stderr, data, report),
    )
    result: dict[str, Any] = {
        "status": record_status,
        "reason": record_reason,
        "record": {
            "samples": args.record_samples,
            "warmup": args.record_warmup,
            "frequency_hz": args.record_frequency,
            "event": "cycles:u",
            "callgraph": "dwarf",
            "process": record_process,
            "source_before": source_before,
            "source_after": source_after_record,
            "data": None if not data.is_file() else {"path": str(data), **meta(data)},
            "report": report_meta,
        },
    }
    if record_status != "pass":
        result["artifacts"] = _artifact_map(directory.iterdir())
        return result

    script_stdout = directory / "perf-script.txt"
    script_stderr = directory / "perf-script.stderr"
    script_started = directory / "perf-script.started.json"
    script_terminal = directory / "perf-script.terminal.json"
    script_argv = _script_command(data, args.cpu)
    _write_process_started(
        script_started,
        role="owned-perf-script",
        argv=script_argv,
        env=env,
        source_before=source_after_record,
        metadata={
            "binary": {"path": str(binary), **meta(binary)},
            "source_revision": revision,
            "input": {"path": str(data), **meta(data)},
        },
    )
    script_process = _run_process(
        script_argv, stdout=script_stdout, stderr=script_stderr, env=env,
        timeout=args.timeout,
    )
    source_after_script = snapshot()
    _source_equal(source_after_record, source_after_script, "owned perf script source after")
    script_status, script_reason = _record_failure_status(script_process, script_stderr)
    if script_status == "pass" and not script_stdout.is_file():
        script_status, script_reason = "failed", "perf script produced no export"
    parsed = None
    stack_summary = None
    folded_path = directory / "perf-folded.txt"
    stack_summary_path = directory / "perf-stack-summary.json"
    if script_status == "pass":
        parsed = parse_perf_script_text(script_stdout.read_text(encoding="utf-8", errors="replace"))
        if parsed["sample_count"] < 1:
            script_status, script_reason = "failed", "perf script exported no symbolized samples"
        else:
            stack_summary = summarize_perf_stacks(parsed)
            stack_summary["report_identity"] = {
                "schema": REPORT_SCHEMA,
                "version": REPORT_VERSION,
                "case_name": REPORT_CASE_NAME,
                "provider": "owned",
                "source_archive_sha256": report_meta["source_archive_sha256"],
                "source_archive_bytes": report_meta["source_archive_bytes"],
                "source_revision": report_meta["source_revision"],
                "binary_sha256": report_meta["binary_sha256"],
                "binary_bytes": report_meta["binary_bytes"],
                "samples": args.record_samples,
                "warmup": args.record_warmup,
                "timing_scope": EXPECTED_TIMING_SCOPE,
            }
            folded_lines = [
                f"{item['stack']} {item['period']}\n"
                for item in stack_summary["folded_stacks"]
            ]
            folded_path.open("x", encoding="utf-8").writelines(folded_lines)
            _write_new(stack_summary_path, stack_summary)
    _write_process_terminal(
        script_terminal,
        role="owned-perf-script",
        process=script_process,
        source_before=source_after_record,
        source_after=source_after_script,
        status=script_status,
        reason=script_reason,
        artifacts=(script_stdout, script_stderr, folded_path, stack_summary_path),
    )
    result["script"] = {
        "status": script_status,
        "reason": script_reason,
        "process": script_process,
        "source_before": source_after_record,
        "source_after": source_after_script,
        "parsed": parsed,
        "stack_summary": None if stack_summary is None else {
            "path": str(stack_summary_path),
            "sha256": sha(stack_summary_path),
            "bytes": stack_summary_path.stat().st_size,
        },
        "folded": None if not folded_path.is_file() else {"path": str(folded_path), **meta(folded_path)},
    }
    result["status"] = script_status if script_status != "pass" else record_status
    result["reason"] = script_reason if script_reason != "" else record_reason
    result["artifacts"] = _artifact_map(directory.iterdir())
    return result


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--build-record", type=Path,
                        default=ROOT / "build-normal.json")
    parser.add_argument("--binary", type=Path)
    parser.add_argument("--binary-sha256")
    parser.add_argument("--source-revision")
    parser.add_argument("--provider", choices=PROVIDERS, action="append", required=True)
    parser.add_argument("--samples", type=int, default=5)
    parser.add_argument("--warmup", type=int, default=2)
    parser.add_argument("--cpu", type=int, default=2)
    parser.add_argument("--timeout", type=int, default=600)
    parser.add_argument(
        "--record-owned", action="store_true",
        help="also run one owned-source perf record and export sampled call stacks",
    )
    parser.add_argument("--record-samples", type=int, default=PERF_RECORD_SAMPLES)
    parser.add_argument("--record-warmup", type=int, default=PERF_RECORD_WARMUP)
    parser.add_argument("--record-frequency", type=int, default=PERF_RECORD_FREQUENCY_HZ)
    parser.add_argument("--output-dir", type=Path, required=True)
    return parser


def run(args: argparse.Namespace) -> Path:
    if args.samples < 1 or args.samples > 10_000 or args.warmup < 0 or args.warmup > 1_000:
        fail("samples must be in 1..10000 and warmup in 0..1000")
    if args.cpu < 0 or args.timeout < 1:
        fail("cpu must be non-negative and timeout must be positive")
    if args.record_samples < 1 or args.record_samples > 10_000:
        fail("record-samples must be in 1..10000")
    if args.record_warmup < 0 or args.record_warmup > 1_000:
        fail("record-warmup must be in 0..1000")
    if args.record_frequency < 1 or args.record_frequency > 10_000:
        fail("record-frequency must be in 1..10000 Hz")
    if len(set(args.provider)) != len(args.provider):
        fail("provider selectors must be unique")
    output = (REPO / args.output_dir).resolve() if not args.output_dir.is_absolute() else args.output_dir.resolve()
    if output.exists():
        fail(f"refusing to reuse profile output directory: {output}")
    build_path = args.build_record.resolve()
    builds, protocol, custody = _profile_custody(build_path)
    normal = builds["normal"]
    retained_binary = Path(normal["binary"]["path"]).resolve()
    binary = retained_binary if args.binary is None else args.binary.resolve()
    if binary != retained_binary:
        fail("--binary does not name the authenticated retained normal executable")
    if args.binary is not None:
        if not args.binary_sha256 or HEX64_RE.fullmatch(args.binary_sha256) is None:
            fail("--binary-sha256 is required with --binary")
        if args.binary_sha256 != normal["binary"]["sha256"]:
            fail("--binary-sha256 differs from the authenticated retained executable")
    revision = normal["git_revision"]
    if args.source_revision is not None:
        if HEX40_RE.fullmatch(args.source_revision) is None or args.source_revision != revision:
            fail("--source-revision differs from authenticated build revision")
    if shutil.which("/usr/bin/taskset") is None:
        fail("/usr/bin/taskset is unavailable")
    source_before = snapshot()
    canonical_source_before = _canonical_source_snapshot()
    _source_equal(custody["source"], canonical_source_before, "profile canonical source before")
    _source_equal(custody["source"], protocol["source"], "profile protocol source")
    output.mkdir(parents=True, exist_ok=False)
    base_env = dict(ENV)
    base_env["PYTHONDONTWRITEBYTECODE"] = "1"
    tool_versions = {
        name: _tool_version(path)
        for name, path in (
            ("taskset", "/usr/bin/taskset"),
            ("perf", "/usr/bin/perf"),
            ("strace", "/usr/bin/strace"),
        )
    }
    results: list[dict[str, Any]] = []
    cleanup_failures: list[str] = []
    for provider in args.provider:
        directory = output / provider
        directory.mkdir()
        run_root, tmp_root = _prepare_private_scratch(provider)
        profile_env = dict(base_env)
        profile_env["TMPDIR"] = str(tmp_root)
        try:
            perf_path = shutil.which("/usr/bin/perf")
            strace_path = shutil.which("/usr/bin/strace")
            # The helper uses absolute paths in argv.  Record an unavailable result
            # if a tool disappears after the support probe rather than failing with
            # a fabricated zero counter.
            if perf_path is None:
                perf_result: dict[str, Any] = {"tool": "perf", "status": "unavailable",
                                               "reason": "tool not found", "artifacts": {}}
            else:
                perf_result = _run_perf(binary, provider, args, directory, revision, profile_env)
            if strace_path is None:
                strace_result: dict[str, Any] = {"tool": "strace", "status": "unavailable",
                                                 "reason": "tool not found", "artifacts": {}}
            else:
                strace_result = _run_strace(binary, provider, args, directory, revision, profile_env)
            source_after = snapshot()
            _source_equal(source_before, source_after, f"profile source after {provider}")
            canonical_source_after = _canonical_source_snapshot()
            _source_equal(custody["source"], canonical_source_after, f"profile canonical source after {provider}")
        finally:
            cleanup = _write_cleanup_receipt(directory, run_root, tmp_root)
        if cleanup["status"] != "pass":
            cleanup_failures.append(f"{provider}: private scratch cleanup failed")
        results.append({
            "provider": provider,
            "environment": {key: profile_env.get(key) for key in ("RUSTUP_TOOLCHAIN", "CARGO_TARGET_DIR", "TMPDIR", "LC_ALL")},
            "scope": "whole child for perf/strace; operation-local lifecycle counters remain in the Rust report",
            "perf": perf_result,
            "strace": strace_result,
            "source_before": source_before,
            "source_after": source_after,
            "cleanup": cleanup,
        })
    owned_record = None
    if args.record_owned:
        owned_directory = output / "owned"
        owned_directory.mkdir(exist_ok=True)
        record_directory = owned_directory / "record"
        record_run_root, record_tmp_root = _prepare_private_scratch("owned-record")
        record_env = dict(base_env)
        record_env["TMPDIR"] = str(record_tmp_root)
        try:
            owned_record = _run_owned_record(
                binary, args, record_directory, revision, record_env, source_before,
            )
            source_after_record = snapshot()
            _source_equal(source_before, source_after_record, "profile source after owned record")
            _source_equal(custody["source"], _canonical_source_snapshot(), "profile canonical source after owned record")
        finally:
            record_cleanup = _write_cleanup_receipt(record_directory, record_run_root, record_tmp_root)
        owned_record["cleanup"] = record_cleanup
        if record_cleanup["status"] != "pass":
            cleanup_failures.append("owned-record: private scratch cleanup failed")
        if owned_record.get("status") == "failed":
            # Write the enclosing summary below before surfacing the failure so
            # the retained terminal/raw artifacts remain reviewable.
            record_failure = owned_record.get("reason") or "owned perf record failed"
        else:
            record_failure = None
    else:
        record_failure = None
    summary = {
        "schema": SCHEMA,
        "version": 1,
        "created_utc": _now(),
        "binary": {"path": str(binary), "bytes": binary.stat().st_size, "sha256": sha(binary), "executable": True},
        "build_record": {"path": str(build_path), "sha256": sha(build_path)},
        "custody": custody,
        "source_revision": revision,
        "cwd": str(REPO),
        "tool_versions": tool_versions,
        "environment": {key: base_env.get(key) for key in ("RUSTUP_TOOLCHAIN", "CARGO_TARGET_DIR", "TMPDIR", "LC_ALL")},
        "providers": results,
        "owned_record": owned_record,
        "observer_scope": {
            "perf": "whole child including setup, process startup, warmups, measured rows, and report serialization",
            "strace": "whole child syscall summary including setup and report serialization; no byte counts",
            "perf_record": "one owned whole-child DWARF call-stack sample; period counts are statistical samples, not phase durations",
            "latency": "use the Rust report's operation-local elapsed vectors, not observer process elapsed time",
            "canonical_report_validator": "measure.validate_report; operation rows are accepted only after the canonical schema and invariant checks",
            "source_build_gate_protocol": "measure.load_builds and measure._load_protocol authenticate the retained normal/allocator receipts, source manifest, build gate receipts, and frozen protocol together",
            "private_scratch": "each observer uses TEMP/managed/profile-r1-<label>/<label>/tmp and retains the exact measure._cleanup_private receipt",
        },
    }
    summary_path = output / "profile-summary.json"
    _write_new(summary_path, summary)
    if cleanup_failures:
        fail("; ".join(cleanup_failures))
    if record_failure is not None:
        fail(record_failure)
    return summary_path
def main(argv: list[str] | None = None) -> int:
    try:
        args = _parser().parse_args(argv)
        path = run(args)
    except (ProfileError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"profile.py: FAIL: {error}", file=sys.stderr)
        return 1
    print(path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
