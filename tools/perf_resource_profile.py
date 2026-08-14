#!/usr/bin/env python3
"""Small, process-isolated resource evidence runner for the perf harness.

This module deliberately uses only the Python standard library.  It wraps the
existing ``litchi-perf-baseline`` selectors with the host tools that happen to
be available and turns their output into one compact, content-addressed JSON
record.  Missing or permission-denied profilers are represented explicitly;
they are never converted into zeroes.

The external traces are temporary.  The aggregate report retains their SHA
and size, the command, and parsed counters, but not the potentially large raw
trace.  Source/sink counters are logical harness counters.  ``strace`` values
are whole-process syscall observations and must not be read as decompressed or
recompressed byte counts.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import platform
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Iterable, Sequence


SCHEMA_VERSION = 1
TOOL_NAME = "litchi-resource-profile"
TOOL_VERSION = "0.1.1"
REPO_ROOT = Path(__file__).resolve().parents[1]
HARNESS_MANIFEST = REPO_ROOT / "tools" / "perf-baseline" / "Cargo.toml"
DEFAULT_BINARY = (
    REPO_ROOT / "tools" / "perf-baseline" / "target" / "release" / "litchi-perf-baseline"
)
DEFAULT_OUTPUT = REPO_ROOT / "docs" / "performance" / "results" / "resource-profile-current-head-0115.json"
MAX_UNTRACKED_FILES = 4096
MAX_UNTRACKED_FILE_BYTES = 16 * 1024 * 1024
MAX_UNTRACKED_BYTES = 64 * 1024 * 1024

DEFAULT_WORKLOADS: tuple[dict[str, Any], ...] = (
    {
        "id": "opc-source-one-part",
        "purpose": "OPC source-backed one-Part publication",
        "args": (
            "--case",
            "opc_source_overlay_one_part_save",
            "--shape",
            "few-large",
            "--payload",
            "incompressible",
        ),
        "profile_external": True,
    },
    {
        "id": "xlsx-managed-batch",
        "purpose": "managed source-backed XLSX scalar batch edit/save",
        "args": (
            "--case",
            "xlsx_source_backed_managed_cell_values_batch_edit_save",
            "--xlsx-cell-crud-shape",
            "medium",
        ),
        "profile_external": True,
    },
    {
        "id": "rtf-streaming",
        "purpose": "bounded RTF streaming creation",
        "args": (
            "--case",
            "rtf_streaming_create",
            "--semantic-shape",
            "medium",
        ),
        "profile_external": True,
    },
    {
        "id": "cfb-selective",
        "purpose": "CFB selective MiniFAT/FAT reads",
        "args": (
            "--case",
            "cfb_selective_mini_shared_read,cfb_selective_fat_shared_read",
            "--shape",
            "many-small",
            "--payload",
            "incompressible",
        ),
        "profile_external": True,
    },
    {
        "id": "cfb-save",
        "purpose": "CFB same-length source-overlay atomic save",
        "args": (
            "--case",
            "cfb_file_same_length_overlay_atomic_save",
            "--filesystem-cache",
            "warm",
        ),
        "profile_external": True,
    },
    {
        "id": "opc-scaling",
        "purpose": "explicit OPC open-session scaling",
        "args": (
            "--case",
            "opc_open_session_scaling",
            "--shape",
            "many-small",
            "--payload",
            "incompressible",
            "--workers",
            "1,2,4,8,available",
        ),
        "profile_external": False,
    },
    {
        "id": "cfb-scaling",
        "purpose": "explicit CFB bulk-read scaling",
        "args": (
            "--case",
            "cfb_bulk_read_scaling",
            "--shape",
            "many-small",
            "--payload",
            "incompressible",
            "--workers",
            "1,2,4,8,available",
        ),
        "profile_external": False,
    },
)

EVENTS = ("cycles", "instructions", "branches", "branch-misses", "cache-misses", "page-faults")
TIME_FIELDS = {
    "Maximum resident set size (kbytes)": "max_rss_kib",
    "User time (seconds)": "user_seconds",
    "System time (seconds)": "system_seconds",
    "Voluntary context switches": "voluntary_context_switches",
    "Involuntary context switches": "involuntary_context_switches",
    "Major (requiring I/O) page faults": "major_page_faults",
    "Minor (reclaiming a frame) page faults": "minor_page_faults",
}
CORPUS_IDENTITY_KEYS = (
    "name",
    "generator",
    "package_format",
    "shape",
    "payload_kind",
    "compression",
    "entry_count",
    "archive_member_count",
    "entry_bytes",
    "uncompressed_payload_bytes",
    "archive_bytes",
    "archive_sha256",
    "target_entry",
    "target_payload_bytes",
    "target_payload_sha256",
    "rtf_variant",
    "xlsx",
)


def _reject_nonfinite(value: str) -> None:
    raise ValueError(f"non-finite JSON value {value!r}")


def load_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle, parse_constant=_reject_nonfinite)


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        while True:
            block = handle.read(1024 * 1024)
            if not block:
                return digest.hexdigest()
            digest.update(block)


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def artifact(path: Path, *, retained: bool = False) -> dict[str, Any]:
    if not path.is_file():
        return {
            "present": False,
            "retained": retained,
            "sha256": None,
            "bytes": None,
        }
    return {
        "present": True,
        "retained": retained,
        "sha256": sha256_file(path),
        "bytes": path.stat().st_size,
    }


def short_text(path: Path, limit: int = 512) -> str | None:
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    text = "\n".join(line.strip() for line in text.splitlines() if line.strip())
    return text[:limit] or None


def command_record(command: Sequence[str]) -> list[str]:
    return [str(item) for item in command]


def run_command(
    command: Sequence[str],
    *,
    stdout_path: Path,
    stderr_path: Path,
    timeout_seconds: float,
) -> dict[str, Any]:
    started = time.monotonic_ns()
    try:
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            completed = subprocess.run(
                list(command),
                stdout=stdout,
                stderr=stderr,
                check=False,
                timeout=timeout_seconds,
            )
        status = completed.returncode
        timed_out = False
    except subprocess.TimeoutExpired as error:
        status = None
        timed_out = True
        if not stdout_path.exists():
            stdout_path.write_bytes(b"")
        if not stderr_path.exists():
            stderr_path.write_bytes(str(error).encode())
    except OSError as error:
        status = None
        timed_out = False
        stdout_path.write_bytes(b"")
        stderr_path.write_text(str(error), encoding="utf-8")
    finished = time.monotonic_ns()
    return {
        "command": command_record(command),
        "returncode": status,
        "timed_out": timed_out,
        "wall_ns": finished - started,
        "stdout": artifact(stdout_path),
        "stderr": artifact(stderr_path),
        "stderr_excerpt": short_text(stderr_path),
    }


def tool_path(name: str) -> Path | None:
    candidate = Path(name)
    if candidate.is_absolute() and candidate.is_file():
        return candidate
    found = shutil.which(name)
    return Path(found) if found else None


def probe_tool(name: str, version_args: Sequence[str]) -> dict[str, Any]:
    path = tool_path(name)
    if path is None:
        return {
            "available": False,
            "path": None,
            "version": None,
            "binary_sha256": None,
            "returncode": None,
        }
    try:
        completed = subprocess.run(
            [str(path), *version_args],
            capture_output=True,
            check=False,
            timeout=10,
        )
        text = (completed.stdout + completed.stderr).decode("utf-8", errors="replace")
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        version = lines[0][:256] if lines else None
        return {
            "available": completed.returncode == 0,
            "path": str(path),
            "version": version,
            "binary_sha256": sha256_file(path) if path.is_file() else None,
            "returncode": completed.returncode,
        }
    except (OSError, subprocess.TimeoutExpired) as error:
        return {
            "available": False,
            "path": str(path),
            "version": None,
            "binary_sha256": sha256_file(path) if path.is_file() else None,
            "returncode": None,
            "error": str(error),
        }


def git_command_bytes(*args: str) -> bytes | None:
    try:
        completed = subprocess.run(
            ["git", "-C", str(REPO_ROOT), *args],
            capture_output=True,
            check=False,
            timeout=30,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    if completed.returncode != 0:
        return None
    return completed.stdout


def git_output(*args: str) -> str | None:
    output = git_command_bytes(*args)
    return output.decode("utf-8", errors="replace").strip() if output is not None else None


def git_output_sha256(*args: str) -> dict[str, Any]:
    output = git_command_bytes(*args)
    if output is None:
        return {"available": False, "sha256": None, "bytes": None}
    return {"available": True, "sha256": sha256_bytes(output), "bytes": len(output)}


def file_content_identity(path: Path) -> dict[str, Any]:
    if not path.is_file():
        return {"path": str(path), "present": False, "sha256": None, "bytes": None}
    return {
        "path": str(path),
        "present": True,
        "sha256": sha256_file(path),
        "bytes": path.stat().st_size,
    }


def _untracked_status_paths(status_bytes: bytes) -> list[str]:
    paths: list[str] = []
    for record in status_bytes.split(b"\0"):
        if record.startswith(b"?? "):
            paths.append(os.fsdecode(record[3:]))
    return sorted(set(paths))


def untracked_content_identity(
    status_bytes: bytes | None,
    *,
    repo_root: Path = REPO_ROOT,
) -> dict[str, Any]:
    """Hash every untracked file named by ``git status -z`` within bounded limits.

    The result is explicitly ``error`` when a path cannot be read, escapes the
    repository, or exceeds a file/count/aggregate budget.  A partial list is
    retained for diagnosis, but callers must not treat it as a complete source
    snapshot unless the status is ``ok``.
    """
    limits = {
        "max_files": MAX_UNTRACKED_FILES,
        "max_file_bytes": MAX_UNTRACKED_FILE_BYTES,
        "max_total_bytes": MAX_UNTRACKED_BYTES,
    }
    if status_bytes is None:
        return {
            "status": "error",
            "error": "git status bytes unavailable",
            "entries": [],
            "total_bytes": 0,
            "limits": limits,
        }
    entries: list[dict[str, Any]] = []
    total_bytes = 0
    error: str | None = None

    def fail(message: str) -> None:
        nonlocal error
        if error is None:
            error = message

    def add_path(path: Path, relative: str) -> None:
        nonlocal total_bytes
        if error is not None:
            return
        try:
            relative_path = Path(relative)
            if relative_path.is_absolute() or ".." in relative_path.parts:
                fail(f"untracked path escapes repository: {relative}")
                return
            path = repo_root.joinpath(*relative_path.parts)
            path.relative_to(repo_root)
            ancestor = repo_root
            for component in relative_path.parts[:-1]:
                ancestor /= component
                ancestor_stat = ancestor.lstat()
                if stat.S_ISLNK(ancestor_stat.st_mode):
                    fail(f"untracked path crosses symlinked ancestor: {relative}")
                    return
                if not stat.S_ISDIR(ancestor_stat.st_mode):
                    fail(f"untracked path ancestor is not a directory: {relative}")
                    return
            stat_result = path.lstat()
        except (OSError, ValueError) as exc:
            fail(f"cannot inspect untracked path {relative}: {exc}")
            return
        if stat.S_ISDIR(stat_result.st_mode):
            try:
                children = sorted(path.iterdir(), key=lambda child: child.name)
            except OSError as exc:
                fail(f"cannot enumerate untracked directory {relative}: {exc}")
                return
            for child in children:
                child_relative = f"{relative.rstrip('/')}/{child.name}"
                add_path(child, child_relative)
            return
        if len(entries) >= MAX_UNTRACKED_FILES:
            fail(f"untracked file limit exceeded ({MAX_UNTRACKED_FILES})")
            return
        if stat.S_ISLNK(stat_result.st_mode):
            try:
                target = os.readlink(path)
                target_bytes = os.fsencode(target)
            except OSError as exc:
                fail(f"cannot read untracked symlink {relative}: {exc}")
                return
            if len(target_bytes) > MAX_UNTRACKED_FILE_BYTES:
                fail(f"untracked symlink target exceeds limit: {relative}")
                return
            if total_bytes + len(target_bytes) > MAX_UNTRACKED_BYTES:
                fail(f"untracked file aggregate limit exceeded at symlink: {relative}")
                return
            total_bytes += len(target_bytes)
            entries.append(
                {
                    "path": relative,
                    "kind": "symlink",
                    "target_sha256": sha256_bytes(target_bytes),
                    "target_bytes": len(target_bytes),
                }
            )
            return
        if not stat.S_ISREG(stat_result.st_mode):
            fail(f"unsupported untracked path kind: {relative}")
            return
        digest = hashlib.sha256()
        file_bytes = 0
        try:
            with path.open("rb") as handle:
                while True:
                    block = handle.read(1024 * 1024)
                    if not block:
                        break
                    file_bytes += len(block)
                    if file_bytes > MAX_UNTRACKED_FILE_BYTES:
                        fail(f"untracked file exceeds per-file limit: {relative}")
                        return
                    total_candidate = total_bytes + file_bytes
                    if total_candidate > MAX_UNTRACKED_BYTES:
                        fail(f"untracked file aggregate limit exceeded at: {relative}")
                        return
                    digest.update(block)
        except OSError as exc:
            fail(f"cannot read untracked file {relative}: {exc}")
            return
        total_bytes += file_bytes
        entries.append(
            {
                "path": relative,
                "kind": "file",
                "sha256": digest.hexdigest(),
                "bytes": file_bytes,
            }
        )

    for relative in _untracked_status_paths(status_bytes):
        add_path(repo_root / relative, relative)
        if error is not None:
            break
    return {
        "status": "error" if error is not None else "ok",
        "error": error,
        "entries": sorted(entries, key=lambda entry: entry["path"]),
        "total_bytes": total_bytes,
        "limits": limits,
    }


def source_content_identity() -> dict[str, Any]:
    status_bytes = git_command_bytes(
        "status", "--porcelain=v1", "-z", "--untracked-files=all"
    )
    revision = git_output("rev-parse", "HEAD")
    head_tree = git_output("rev-parse", "HEAD^{tree}")
    unstaged_diff = git_output_sha256("diff", "--binary", "--no-ext-diff")
    staged_diff = git_output_sha256("diff", "--binary", "--no-ext-diff", "--cached")
    head_to_worktree_diff = git_output_sha256(
        "diff", "--binary", "--no-ext-diff", "HEAD"
    )
    harness_manifest = file_content_identity(HARNESS_MANIFEST)
    harness_lock = file_content_identity(HARNESS_MANIFEST.with_name("Cargo.lock"))
    untracked = untracked_content_identity(status_bytes)
    snapshot_complete = all(
        (
            status_bytes is not None,
            revision is not None,
            head_tree is not None,
            unstaged_diff["available"],
            staged_diff["available"],
            head_to_worktree_diff["available"],
            harness_manifest["present"],
            harness_lock["present"],
            untracked["status"] == "ok",
        )
    )
    return {
        "revision": revision,
        "head_tree": head_tree,
        "git_worktree_dirty": bool(status_bytes),
        "snapshot_status": "complete" if snapshot_complete else "incomplete",
        "status_z": {
            "available": status_bytes is not None,
            "sha256": sha256_bytes(status_bytes) if status_bytes is not None else None,
            "bytes": len(status_bytes) if status_bytes is not None else None,
        },
        "unstaged_diff": unstaged_diff,
        "staged_diff": staged_diff,
        "head_to_worktree_diff": head_to_worktree_diff,
        "harness_manifest": harness_manifest,
        "harness_lock": harness_lock,
        "untracked_content": untracked,
        "snapshot_atomic": False,
        "scope": (
            "HEAD tree plus exact tracked diffs, status path identity, and bounded untracked "
            "content; pre/post snapshots are non-atomic and are not cryptographic "
            "source-to-binary proof"
        ),
    }


def environment() -> dict[str, Any]:
    cpu_model = None
    try:
        for line in Path("/proc/cpuinfo").read_text(encoding="utf-8", errors="replace").splitlines():
            if line.lower().startswith("model name"):
                cpu_model = line.split(":", 1)[-1].strip()
                break
    except OSError:
        pass
    try:
        page_size = os.sysconf("SC_PAGE_SIZE")
    except (AttributeError, OSError, ValueError):
        page_size = None
    return {
        "os": platform.system(),
        "kernel": platform.release(),
        "machine": platform.machine(),
        "python": platform.python_version(),
        "cpu_model": cpu_model,
        "logical_cpus_available": os.cpu_count(),
        "page_size_bytes": page_size,
        "perf_event_paranoid": (
            Path("/proc/sys/kernel/perf_event_paranoid").read_text().strip()
            if Path("/proc/sys/kernel/perf_event_paranoid").is_file()
            else None
        ),
        "rustc_version": tool_version("rustc", ("--version",)),
        "cargo_version": tool_version("cargo", ("--version",)),
    }


def tool_version(name: str, args: Sequence[str]) -> str | None:
    path = tool_path(name)
    if path is None:
        return None
    try:
        completed = subprocess.run([str(path), *args], capture_output=True, check=False, timeout=10)
    except (OSError, subprocess.TimeoutExpired):
        return None
    if completed.returncode != 0:
        return None
    output = (completed.stdout + completed.stderr).decode("utf-8", errors="replace")
    return next((line.strip() for line in output.splitlines() if line.strip()), None)


def parse_time_report(path: Path) -> dict[str, Any]:
    parsed: dict[str, Any] = {"status": "missing"}
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError as error:
        return {"status": "missing", "error": str(error)}
    for raw_line in lines:
        line = raw_line.strip()
        if line.startswith("Elapsed (wall clock") and ": " in line:
            parsed["elapsed_wall_seconds"] = parse_elapsed_clock(line.rpartition(": ")[2].strip())
            continue
        if ":" not in line:
            continue
        label, raw = line.split(":", 1)
        label = label.strip()
        raw = raw.strip()
        key = TIME_FIELDS.get(label)
        if key is not None:
            try:
                parsed[key] = float(raw) if "time" in label.lower() else int(raw.replace(",", ""))
            except ValueError:
                parsed[key] = None
    parsed["status"] = "ok" if "max_rss_kib" in parsed else "unparsed"
    parsed["artifact"] = artifact(path)
    return parsed


def parse_elapsed_clock(value: str) -> float | None:
    try:
        parts = value.split(":")
        if len(parts) == 3:
            hours, minutes, seconds = parts
            return int(hours) * 3600 + int(minutes) * 60 + float(seconds)
        if len(parts) == 2:
            minutes, seconds = parts
            return int(minutes) * 60 + float(seconds)
        return float(value)
    except ValueError:
        return None


def _numeric_token(value: str) -> int | None:
    value = value.strip().replace(",", "")
    try:
        return int(value)
    except ValueError:
        try:
            number = float(value)
        except ValueError:
            return None
        return int(number) if math.isfinite(number) else None


def _bytes_token(value: str) -> int | None:
    match = re.fullmatch(r"([0-9]+(?:\.[0-9]+)?)\s*([KMGT]?)(?:i?B)?", value.strip(), re.IGNORECASE)
    if not match:
        return _numeric_token(value)
    number = float(match.group(1))
    unit = match.group(2).upper()
    scale = {"": 1, "K": 1024, "M": 1024**2, "G": 1024**3, "T": 1024**4}[unit]
    return int(number * scale)


def parse_heaptrack_print(path: Path) -> dict[str, Any]:
    """Parse stable process-total fields from heaptrack_print -H output."""
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        return {"status": "missing", "error": str(error), "artifact": artifact(path)}
    patterns: tuple[tuple[str, str, Any], ...] = (
        ("allocation_calls", r"calls to allocation functions:\s*([0-9,]+)", _numeric_token),
        (
            "temporary_allocations",
            r"([0-9,]+) temporary allocations of [0-9,]+ allocations in total",
            _numeric_token,
        ),
        ("peak_heap_bytes", r"peak heap memory consumption:\s*([0-9.,]+\s*[KMGT]?(?:i?B)?)", _bytes_token),
        (
            "peak_rss_bytes",
            r"peak RSS(?:\s*\([^)]*\))?:\s*([0-9.,]+\s*[KMGT]?(?:i?B)?)",
            _bytes_token,
        ),
    )
    parsed: dict[str, Any] = {"status": "ok", "artifact": artifact(path)}
    for key, pattern, converter in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        parsed[key] = converter(match.group(1)) if match else None
    if all(parsed[key] is None for key, _, _ in patterns):
        parsed["status"] = "unparsed"
    return parsed


def parse_heaptrack_histogram(path: Path) -> int | None:
    """Return total allocated bytes from heaptrack's ``size<TAB>count`` TSV."""
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    total = 0
    found = False
    for line in text.splitlines():
        fields = line.split("\t")
        if len(fields) != 2:
            continue
        size = _numeric_token(fields[0])
        count = _numeric_token(fields[1])
        if size is None or count is None or size < 0 or count < 0:
            continue
        total += size * count
        found = True
    return total if found else None


def parse_perf_stat(path: Path, returncode: int | None, stderr_path: Path) -> dict[str, Any]:
    counters = {
        event: {"value": None, "unit": None, "available": False} for event in EVENTS
    }
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        text = ""
    for line in text.splitlines():
        if not line or line.startswith("#"):
            continue
        fields = line.split(",")
        event = next((candidate for candidate in EVENTS if candidate in fields), None)
        if event is None:
            continue
        index = fields.index(event)
        value = fields[0].strip() if fields else ""
        try:
            numeric = float(value.replace(" ", "").replace(",", ""))
        except ValueError:
            numeric = None
        counters[event] = {
            "value": numeric,
            "unit": fields[1] if len(fields) > 1 and fields[1] else None,
            "available": numeric is not None,
            "raw": fields[: index + 1],
        }
    available = any(item["available"] for item in counters.values())
    if available:
        status = "ok"
    else:
        status = "unsupported" if returncode not in (0, None) else "unavailable"
    return {
        "status": status,
        "returncode": returncode,
        "counters": counters,
        "stderr_excerpt": short_text(stderr_path),
        "artifact": artifact(path),
    }


SYSCALL_RE = re.compile(r"\b(read|write)\([^=]*\)\s+=\s+(-?\d+)")
SIZE_BUCKETS = (
    ("0", 0, 0),
    ("1-511", 1, 511),
    ("512-4095", 512, 4095),
    ("4096-16383", 4096, 16383),
    ("16384-65535", 16384, 65535),
    ("65536-262143", 65536, 262143),
    ("262144-1048575", 262144, 1048575),
    ("1048576+", 1048576, None),
)


def syscall_bucket(size: int) -> str:
    for name, lower, upper in SIZE_BUCKETS:
        if size >= lower and (upper is None or size <= upper):
            return name
    return "1048576+"


def parse_strace(path: Path, returncode: int | None) -> dict[str, Any]:
    summary: dict[str, Any] = {}
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        text = ""
    for syscall in ("read", "write"):
        buckets = {name: 0 for name, _, _ in SIZE_BUCKETS}
        calls = 0
        failures = 0
        total_bytes = 0
        largest = 0
        for match in SYSCALL_RE.finditer(text):
            if match.group(1) != syscall:
                continue
            size = int(match.group(2))
            if size < 0:
                failures += 1
                continue
            calls += 1
            total_bytes += size
            largest = max(largest, size)
            buckets[syscall_bucket(size)] += 1
        summary[syscall] = {
            "calls": calls,
            "failed_calls": failures,
            "returned_bytes": total_bytes,
            "largest_returned_bytes": largest,
            "size_buckets": buckets,
        }
    return {
        "status": "ok" if summary["read"]["calls"] or summary["write"]["calls"] else "unparsed",
        "returncode": returncode,
        "scope": "whole process; read/write only; not logical/decompressed I/O",
        "syscalls": summary,
        "artifact": artifact(path),
    }


def corpus_identity(corpus: Any) -> dict[str, Any]:
    if not isinstance(corpus, dict):
        return {}
    return {key: corpus[key] for key in CORPUS_IDENTITY_KEYS if key in corpus}


def compact_filesystem_evidence(value: Any) -> Any:
    """Drop raw per-request arrays while retaining logical totals and buckets."""
    if not isinstance(value, list):
        return value
    compact: list[Any] = []
    for item in value:
        if not isinstance(item, dict):
            compact.append(item)
            continue
        row: dict[str, Any] = {}
        for key, field in item.items():
            if key == "corpus":
                row[key] = corpus_identity(field)
            elif key == "samples" and isinstance(field, list):
                row[key] = [
                    {
                        sample_key: sample_value
                        for sample_key, sample_value in sample.items()
                        if sample_key
                        not in {
                            "logical_read_request_sizes",
                            "read_request_sizes",
                            "range_sizes",
                        }
                    }
                    if isinstance(sample, dict)
                    else sample
                    for sample in field
                ]
            else:
                row[key] = field
        compact.append(row)
    return compact


def logical_measurements(report: dict[str, Any]) -> list[dict[str, Any]]:
    results = report.get("results")
    if not isinstance(results, list):
        return []
    measurements: list[dict[str, Any]] = []
    for result in results:
        if not isinstance(result, dict):
            continue
        row: dict[str, Any] = {
            "case": result.get("case"),
            "corpus": corpus_identity(result.get("corpus")),
            "elapsed_ns": {
                key: result.get("elapsed_ns", {}).get(key)
                for key in ("unit", "p50", "p95", "p99", "mean", "standard_deviation")
                if isinstance(result.get("elapsed_ns"), dict) and key in result["elapsed_ns"]
            },
        }
        for key in ("source", "sink", "execution", "output_sha256", "cache_state"):
            if key in result:
                row[key] = result[key]
        measurements.append(row)
    filesystem = report.get("filesystem_evidence")
    if filesystem:
        measurements.append({"filesystem_evidence": compact_filesystem_evidence(filesystem)})
    return measurements


def scaling_analysis(measurements: Iterable[dict[str, Any]]) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    for measurement in measurements:
        execution = measurement.get("execution")
        elapsed = measurement.get("elapsed_ns")
        if not isinstance(execution, dict) or not isinstance(elapsed, dict):
            continue
        workers = execution.get("worker_count")
        p50 = elapsed.get("p50")
        if isinstance(workers, int) and workers > 0 and isinstance(p50, (int, float)) and p50 > 0:
            rows.append(
                {
                    "workers": workers,
                    "p50_ns": p50,
                    "logical_tasks": execution.get("logical_tasks"),
                    "logical_bytes": execution.get("logical_bytes"),
                }
            )
    rows.sort(key=lambda row: row["workers"])
    if not rows:
        return {
            "status": "not_applicable",
            "metric": "elapsed_ns.p50",
            "rows": [],
            "classification": None,
        }
    baseline = next((row for row in rows if row["workers"] == 1), rows[0])
    baseline_time = float(baseline["p50_ns"])
    analysed: list[dict[str, Any]] = []
    for row in rows:
        workers = row["workers"]
        speedup = baseline_time / float(row["p50_ns"])
        efficiency = speedup / (workers / baseline["workers"])
        serial_fraction = None
        serial_fraction_raw = None
        serial_fraction_valid: bool | None = None
        serial_fraction_reason = "baseline"
        if workers != baseline["workers"]:
            serial_fraction_raw = (1.0 / speedup - 1.0 / workers) / (1.0 - 1.0 / workers)
            serial_fraction_valid = math.isfinite(serial_fraction_raw) and 0.0 <= serial_fraction_raw <= 1.0
            if serial_fraction_valid:
                serial_fraction = serial_fraction_raw
                serial_fraction_reason = None
            else:
                serial_fraction_reason = (
                    "nonfinite" if not math.isfinite(serial_fraction_raw) else "outside_0_1"
                )
        analysed.append(
            {
                **row,
                "speedup_vs_baseline": speedup,
                "efficiency_vs_baseline": efficiency,
                "amdahl_serial_fraction": serial_fraction,
                "amdahl_serial_fraction_raw": serial_fraction_raw,
                "amdahl_serial_fraction_valid": serial_fraction_valid,
                "amdahl_serial_fraction_reason": serial_fraction_reason,
            }
        )
    nonbaseline = [row for row in analysed if row["workers"] != baseline["workers"]]
    serial_values = [
        row["amdahl_serial_fraction"]
        for row in nonbaseline
        if row["amdahl_serial_fraction_valid"] is True
    ]
    invalid_amdahl = [row for row in nonbaseline if row["amdahl_serial_fraction_valid"] is False]
    if not nonbaseline:
        classification = "single_width_only"
    elif invalid_amdahl:
        classification = "nonideal_or_measurement_noise"
    elif all(row["speedup_vs_baseline"] <= 1.05 for row in nonbaseline):
        classification = "no_measured_speedup"
    elif any(row["speedup_vs_baseline"] > row["workers"] * 1.05 for row in nonbaseline):
        classification = "superlinear_or_measurement_noise"
    elif serial_values and max(serial_values) >= 0.75:
        classification = "serial_dominated"
    elif serial_values and max(serial_values) >= 0.25:
        classification = "mixed_serial_parallel"
    else:
        classification = "parallel_friendly_within_measured_widths"
    return {
        "status": "observed",
        "metric": "elapsed_ns.p50",
        "baseline_workers": baseline["workers"],
        "rows": analysed,
        "classification": classification,
        "model": "Amdahl estimate from p50 only; descriptive, not a physical limit",
    }


def _profile_command(
    binary: Path,
    args: Sequence[str],
    *,
    warmup: int,
    samples: int,
    output: Path,
) -> list[str]:
    return [
        str(binary),
        "--warmup",
        str(warmup),
        "--samples",
        str(samples),
        *args,
        "--json",
        str(output),
    ]


def _tool_command(name: str, fallback: str | None = None) -> str | None:
    path = tool_path(name)
    return str(path) if path is not None else fallback


def _profile_external(
    spec: dict[str, Any],
    binary: Path,
    args: Sequence[str],
    workdir: Path,
    tools: dict[str, dict[str, Any]],
    *,
    timeout_seconds: float,
) -> dict[str, Any]:
    result: dict[str, Any] = {}
    profile_args = _profile_command(binary, args, warmup=0, samples=1, output=workdir / "profile-harness.json")

    time_tool = tools.get("time", {})
    if time_tool.get("available"):
        time_report = workdir / "time-v.txt"
        stdout = workdir / "time-stdout.txt"
        stderr = workdir / "time-stderr.txt"
        command = [str(time_tool["path"]), "-v", "-o", str(time_report), "--", *profile_args]
        run = run_command(command, stdout_path=stdout, stderr_path=stderr, timeout_seconds=timeout_seconds)
        result["time"] = {**run, "parsed": parse_time_report(time_report)}
    else:
        result["time"] = {"status": "unsupported", "reason": "GNU /usr/bin/time unavailable"}

    perf_tool = tools.get("perf", {})
    if perf_tool.get("available"):
        perf_report = workdir / "perf-stat.txt"
        stdout = workdir / "perf-stdout.txt"
        stderr = workdir / "perf-stderr.txt"
        command = [
            str(perf_tool["path"]),
            "stat",
            "-x,",
            "-o",
            str(perf_report),
            "-e",
            ",".join(EVENTS),
            "--",
            *profile_args,
        ]
        run = run_command(command, stdout_path=stdout, stderr_path=stderr, timeout_seconds=timeout_seconds)
        result["perf"] = {
            **run,
            "parsed": parse_perf_stat(perf_report, run["returncode"], stderr),
        }
    else:
        result["perf"] = {
            "status": "unsupported",
            "reason": "perf unavailable",
            "counters": {event: {"value": None, "available": False} for event in EVENTS},
        }

    strace_tool = tools.get("strace", {})
    if strace_tool.get("available"):
        strace_report = workdir / "strace.log"
        stdout = workdir / "strace-stdout.txt"
        stderr = workdir / "strace-stderr.txt"
        command = [
            str(strace_tool["path"]),
            "-f",
            "-qq",
            "-e",
            "trace=read,write",
            "-o",
            str(strace_report),
            "--",
            *profile_args,
        ]
        run = run_command(command, stdout_path=stdout, stderr_path=stderr, timeout_seconds=timeout_seconds)
        result["strace"] = {**run, "parsed": parse_strace(strace_report, run["returncode"])}
    else:
        result["strace"] = {"status": "unsupported", "reason": "strace unavailable"}

    heaptrack_tool = tools.get("heaptrack", {})
    print_tool = tools.get("heaptrack_print", {})
    if heaptrack_tool.get("available") and print_tool.get("available"):
        prefix = workdir / "heaptrack-profile"
        stdout = workdir / "heaptrack-stdout.txt"
        stderr = workdir / "heaptrack-stderr.txt"
        command = [str(heaptrack_tool["path"]), "--record-only", "-o", str(prefix), *profile_args]
        run = run_command(command, stdout_path=stdout, stderr_path=stderr, timeout_seconds=timeout_seconds)
        captures = sorted(
            path
            for path in workdir.glob(prefix.name + "*")
            if path.is_file() and path not in (stdout, stderr)
        )
        capture = captures[-1] if captures else None
        if capture is not None:
            printed = workdir / "heaptrack-print.txt"
            print_stderr = workdir / "heaptrack-print-stderr.txt"
            histogram = workdir / "heaptrack-histogram.tsv"
            print_run = run_command(
                [str(print_tool["path"]), "-f", str(capture), "-H", str(histogram)],
                stdout_path=printed,
                stderr_path=print_stderr,
                timeout_seconds=timeout_seconds,
            )
            parsed = parse_heaptrack_print(printed)
            parsed["allocated_bytes"] = parse_heaptrack_histogram(histogram)
            parsed["histogram_artifact"] = artifact(histogram)
            result["heaptrack"] = {
                **run,
                "capture": artifact(capture),
                "print": {**print_run, "parsed": parsed},
            }
        else:
            result["heaptrack"] = {
                **run,
                "status": "unparsed",
                "capture": {"present": False, "retained": False},
            }
    else:
        result["heaptrack"] = {
            "status": "unsupported",
            "reason": "heaptrack or heaptrack_print unavailable",
        }
    return result


def profile_workload(
    spec: dict[str, Any],
    binary: Path,
    *,
    warmup: int,
    samples: int,
    tools: dict[str, dict[str, Any]],
    root: Path,
    timeout_seconds: float,
) -> dict[str, Any]:
    workload_dir = root / str(spec["id"])
    workload_dir.mkdir(parents=True, exist_ok=True)
    args = tuple(str(item) for item in spec["args"])
    harness_json = workload_dir / "harness.json"
    command = _profile_command(binary, args, warmup=warmup, samples=samples, output=harness_json)
    time_report = workload_dir / "baseline.time.txt"
    stdout = workload_dir / "baseline.stdout.txt"
    stderr = workload_dir / "baseline.stderr.txt"
    time_tool = tools.get("time", {})
    if time_tool.get("available"):
        timed_command = [str(time_tool["path"]), "-v", "-o", str(time_report), "--", *command]
    else:
        timed_command = command
    run = run_command(timed_command, stdout_path=stdout, stderr_path=stderr, timeout_seconds=timeout_seconds)
    if run["returncode"] != 0 or not harness_json.is_file():
        raise RuntimeError(
            f"harness failed for {spec['id']}: returncode={run['returncode']} stderr={run['stderr_excerpt']}"
        )
    report = load_json(harness_json)
    measurements = logical_measurements(report)
    output: dict[str, Any] = {
        "id": spec["id"],
        "purpose": spec["purpose"],
        "harness": {
            "command": command_record(command),
            "returncode": run["returncode"],
            "report": artifact(harness_json),
            "logical_measurements": measurements,
        },
        "time": {
            "command": command_record(timed_command),
            "run": run,
            "parsed": parse_time_report(time_report) if time_tool.get("available") else {"status": "unsupported"},
        },
        "external_profiles": {},
        "physical_io_claim_scope": (
            "strace is a whole-process read/write syscall trace only; it is not a source-byte, "
            "decompressed-byte, recompressed-byte, or memory-copy measurement"
        ),
    }
    if spec.get("profile_external", False):
        external_dir = workload_dir / "external"
        external_dir.mkdir(exist_ok=True)
        output["external_profiles"] = _profile_external(
            spec,
            binary,
            args,
            external_dir,
            tools,
            timeout_seconds=timeout_seconds,
        )
    return output


def build_identity(
    binary: Path,
    *,
    build_command: Sequence[str] | None,
    build_result: dict[str, Any] | None,
    source_identity: dict[str, Any],
    pre_build_source_identity: dict[str, Any] | None,
    rerun_command: Sequence[str],
) -> dict[str, Any]:
    build_executed = build_result is not None and build_result.get("executed") is True
    build_succeeded = (
        build_executed
        and build_result.get("returncode") == 0
        and build_result.get("timed_out") is False
    )
    snapshots_captured = pre_build_source_identity is not None
    snapshots_complete = (
        snapshots_captured
        and source_identity.get("snapshot_status") == "complete"
        and pre_build_source_identity.get("snapshot_status") == "complete"
    )
    snapshots_match = (
        pre_build_source_identity == source_identity if snapshots_captured else None
    )
    post_build_dirty = source_identity.get("git_worktree_dirty") is True
    if not build_executed:
        provenance_status = "prebuilt_binary_hash_only"
    elif not build_succeeded:
        provenance_status = "build_failed"
    elif not snapshots_captured or not snapshots_complete or not snapshots_match or post_build_dirty:
        provenance_status = "build_succeeded_source_snapshot_only"
    else:
        provenance_status = "build_succeeded_matching_source_snapshots"
    normalized_build_result = dict(build_result or {})
    normalized_build_result.update(
        {
            "build_succeeded": build_succeeded,
            "provenance_status": provenance_status,
            "pre_build_snapshot_captured": snapshots_captured,
            "post_build_snapshot_captured": True,
            "snapshot_complete": snapshots_complete,
            "snapshot_match": snapshots_match,
        }
    )
    rerun_reasons: list[str] = []
    if not build_executed:
        rerun_reasons.append("binary was prebuilt and not built in this invocation")
    elif not build_succeeded:
        rerun_reasons.append("build did not complete successfully without timeout")
    if build_succeeded and not snapshots_captured:
        rerun_reasons.append("pre-build source snapshot was not captured")
    if build_succeeded and not snapshots_complete:
        rerun_reasons.append("source snapshot was incomplete")
    if build_succeeded and snapshots_captured and not snapshots_match:
        rerun_reasons.append("pre-build and post-build source snapshots differ")
    if build_succeeded and post_build_dirty:
        rerun_reasons.append("post-build worktree is dirty")
    rerun_required = provenance_status in {
        "prebuilt_binary_hash_only",
        "build_failed",
        "build_succeeded_source_snapshot_only",
    }
    identity = {
        "binary": str(binary),
        "binary_sha256": sha256_file(binary) if binary.is_file() else None,
        "binary_bytes": binary.stat().st_size if binary.is_file() else None,
        "build_command": command_record(build_command) if build_command else None,
        "build_result": normalized_build_result,
        "source_content_identity": source_identity,
        "source_content_identity_pre_build": pre_build_source_identity,
        "source_content_identity_post_build": source_identity,
        "provenance": {
            "status": provenance_status,
            "binary_hash_is_exact": binary.is_file(),
            "build_succeeded": build_succeeded,
            "snapshots_captured": snapshots_captured,
            "snapshots_complete": snapshots_complete,
            "snapshots_match": snapshots_match,
            "post_build_worktree_dirty": post_build_dirty,
            "source_binary_binding": (
                "non-atomic matching source snapshots only; no cryptographic source-to-binary proof"
                if provenance_status == "build_succeeded_matching_source_snapshots"
                else "not established: binary hash is exact, but the build/source binding conditions were not all met"
            ),
            "snapshot_atomic": False,
            "cryptographic_source_binary_binding": False,
            "rerun_required": rerun_required,
            "rerun_reasons": rerun_reasons,
            "rerun_command": command_record(rerun_command) if rerun_required else None,
        },
    }
    return identity


def run_profile(arguments: argparse.Namespace) -> int:
    binary = Path(arguments.binary).resolve() if arguments.binary else DEFAULT_BINARY
    build_command: list[str] = [
        "cargo",
        "build",
        "--release",
        "--locked",
        "--manifest-path",
        str(HARNESS_MANIFEST),
    ]
    build_result: dict[str, Any] = {
        "executed": False,
        "provenance_status": "prebuilt_binary_hash_only",
        "source_binding": "not_established",
        "reason": "prebuilt binary supplied; command recorded for reproducibility",
        "rerun_required": True,
        "rerun_after": "current production batches settle",
    }
    pre_build_source_identity = source_content_identity()
    if arguments.build:
        build_stdout_fd, build_stdout_name = tempfile.mkstemp(
            prefix="litchi-resource-build-", suffix=".stdout"
        )
        build_stderr_fd, build_stderr_name = tempfile.mkstemp(
            prefix="litchi-resource-build-", suffix=".stderr"
        )
        os.close(build_stdout_fd)
        os.close(build_stderr_fd)
        build_stdout = Path(build_stdout_name)
        build_stderr = Path(build_stderr_name)
        try:
            build_run = run_command(
                build_command,
                stdout_path=build_stdout,
                stderr_path=build_stderr,
                timeout_seconds=arguments.timeout,
            )
        finally:
            build_stdout.unlink(missing_ok=True)
            build_stderr.unlink(missing_ok=True)
        build_result = {
            "executed": True,
            "provenance_status": "build_succeeded",
            "source_binding": "captured_pre_and_post_build",
            **build_run,
        }
        if build_run["returncode"] != 0:
            raise RuntimeError(f"cargo build failed: {build_run['stderr_excerpt']}")
    if not binary.is_file():
        raise FileNotFoundError(f"performance binary does not exist: {binary}")

    tools = {
        "time": probe_tool("/usr/bin/time", ("--version",)),
        "heaptrack": probe_tool("heaptrack", ("--version",)),
        "heaptrack_print": probe_tool("heaptrack_print", ("--version",)),
        "perf": probe_tool("perf", ("--version",)),
        "strace": probe_tool("strace", ("--version",)),
        "taskset": probe_tool("taskset", ("--version",)),
    }
    selected = list(DEFAULT_WORKLOADS)
    if arguments.only:
        wanted = set(arguments.only.split(","))
        unknown = wanted - {str(spec["id"]) for spec in selected}
        if unknown:
            raise ValueError(f"unknown workload(s): {', '.join(sorted(unknown))}")
        selected = [spec for spec in selected if spec["id"] in wanted]
    source_identity = source_content_identity()
    revision = source_identity["revision"]
    status = git_output("status", "--porcelain")
    output_path = Path(arguments.output).resolve() if arguments.output else DEFAULT_OUTPUT
    rerun_command = (
        "python3",
        "tools/perf_resource_profile.py",
        "run",
        "--build",
        "--binary",
        str(binary),
        "--output",
        str(output_path),
        "--warmup",
        str(arguments.warmup),
        "--samples",
        str(arguments.samples),
        "--timeout",
        str(arguments.timeout),
    )
    with tempfile.TemporaryDirectory(prefix="litchi-resource-profile-") as temporary:
        root = Path(temporary)
        workloads = [
            profile_workload(
                spec,
                binary,
                warmup=arguments.warmup,
                samples=arguments.samples,
                tools=tools,
                root=root,
                timeout_seconds=arguments.timeout,
            )
            for spec in selected
        ]
        for workload in workloads:
            workload["scaling"] = scaling_analysis(workload["harness"]["logical_measurements"])
        report = {
            "schema_version": SCHEMA_VERSION,
            "tool": {"name": TOOL_NAME, "version": TOOL_VERSION, "python_standard_library_only": True},
            "scope": {
                "revision": revision,
                "git_worktree_dirty": bool(status),
                "git_status_sha256": hashlib.sha256((status or "").encode()).hexdigest(),
                "claim": "current-HEAD evidence only; no before/after optimization comparison",
                "excluded_formats": ["iWork"],
            },
            "environment": environment(),
            "binary_identity": build_identity(
                binary,
                build_command=build_command,
                build_result=build_result,
                source_identity=source_identity,
                pre_build_source_identity=pre_build_source_identity,
                rerun_command=rerun_command,
            ),
            "configuration": {
                "warmup_iterations": arguments.warmup,
                "samples": arguments.samples,
                "external_profile_samples": 1,
                "timeout_seconds": arguments.timeout,
                "workloads": [str(spec["id"]) for spec in selected],
            },
            "tools": tools,
            "workloads": workloads,
            "limitations": [
                "The selected synthetic corpora are generated inside the existing harness.",
                "External tool runs use one sample and are process-total evidence with profiler overhead.",
                "No cold-cache, remote-range, allocation attribution, decompressed-byte, recompressed-byte, or memory-copy claim is made.",
                "Amdahl rows use harness p50 elapsed time and are descriptive at the measured widths only; fractions outside [0,1] remain raw but are null/invalidated in the estimate.",
            ],
        }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, sort_keys=True, allow_nan=False)
        handle.write("\n")
    return 0


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    run = subparsers.add_parser("run", help="run current-HEAD resource evidence")
    run.add_argument("--binary", type=Path, default=DEFAULT_BINARY)
    run.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    run.add_argument("--build", action="store_true", help="build the isolated release harness first")
    run.add_argument("--only", help="comma-separated workload IDs")
    run.add_argument("--warmup", type=int, default=1)
    run.add_argument("--samples", type=int, default=3)
    run.add_argument("--timeout", type=float, default=600.0)
    run.set_defaults(function=run_profile)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    if arguments.command == "run":
        if arguments.warmup < 0 or arguments.samples < 1 or arguments.timeout <= 0:
            raise SystemExit("--warmup must be non-negative, --samples and --timeout must be positive")
    try:
        return int(arguments.function(arguments))
    except (OSError, RuntimeError, ValueError) as error:
        print(f"{TOOL_NAME}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
