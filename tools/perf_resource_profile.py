#!/usr/bin/env python3
"""Small, process-isolated resource evidence runner for the perf harness.

This module deliberately uses only the Python standard library.  It wraps the
existing ``litchi-perf-baseline`` selectors with the host tools that happen to
be available and turns their output into one compact, content-addressed JSON
record.  Missing or permission-denied profilers are represented explicitly;
they are never converted into zeroes.

The ordinary current-HEAD mode uses temporary external traces and retains
their SHA and size, command, and parsed counters, but not the potentially
large raw trace.  Its XLSX managed-batch ABBA mode instead retains per-leg
harness, ``/usr/bin/time``, and optional heaptrack artifacts under an explicit
artifact directory.  The DOCX semantic/full-text ABBA mode (optionally
including the one-paragraph text case) uses explicit control/candidate
binaries, requires matching deterministic corpus manifests,
and labels all harness elapsed values as instrumented resource observations.
The XLSX borrowed-parser ABBA mode applies the same retained-artifact contract
to its fixed tiny-read and medium edit/save selector tuple.
Source/sink counters are logical harness counters.
``strace`` values are whole-process syscall observations and must not be read
as decompressed or recompressed byte counts.
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
import statistics
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from typing import Any, Callable, Iterable, Sequence


SCHEMA_VERSION = 1
TOOL_NAME = "litchi-resource-profile"
TOOL_VERSION = "0.1.2"
ABBA_SCHEMA_VERSION = 1
REPO_ROOT = Path(__file__).resolve().parents[1]
HARNESS_MANIFEST = REPO_ROOT / "tools" / "perf-baseline" / "Cargo.toml"
DEFAULT_BINARY = (
    REPO_ROOT / "tools" / "perf-baseline" / "target" / "release" / "litchi-perf-baseline"
)
DEFAULT_OUTPUT = REPO_ROOT / "docs" / "performance" / "results" / "resource-profile-current-head-0115.json"
DEFAULT_ABBA_OUTPUT = (
    REPO_ROOT
    / "docs"
    / "performance"
    / "results"
    / "xlsx-managed-batch-resource-abba-current.json"
)
DEFAULT_DOCX_ABBA_OUTPUT = (
    REPO_ROOT
    / "docs"
    / "performance"
    / "results"
    / "docx-semantic-resource-abba-current.json"
)
DEFAULT_XLSX_XML_BORROWED_ABBA_OUTPUT = (
    REPO_ROOT
    / "docs"
    / "performance"
    / "results"
    / "xlsx-xml-borrowed-resource-abba-current.json"
)
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

XLSX_MANAGED_BATCH_ID = "xlsx-managed-batch"
XLSX_MANAGED_BATCH_CASE = "xlsx_source_backed_managed_cell_values_batch_edit_save"
XLSX_MANAGED_BATCH_SHAPE = "medium"
XLSX_MANAGED_BATCH_ARGS: tuple[str, ...] = (
    "--case",
    XLSX_MANAGED_BATCH_CASE,
    "--xlsx-cell-crud-shape",
    XLSX_MANAGED_BATCH_SHAPE,
)
XLSX_XML_BORROWED_ID = "xlsx-xml-borrowed"
XLSX_XML_BORROWED_ID_ALIASES: tuple[str, ...] = (
    XLSX_XML_BORROWED_ID,
    "xlsx-borrowed",
)
XLSX_XML_BORROWED_CASES: tuple[str, ...] = (
    "xlsx_first_cell",
    "xlsx_source_first_cell",
    "xlsx_eager_cell_values_one_edit_save",
    "xlsx_source_backed_cell_values_one_edit_save",
)
XLSX_XML_BORROWED_SHAPE = "tiny"
XLSX_XML_BORROWED_CELL_CRUD_SHAPE = "medium"
XLSX_XML_BORROWED_ARGS: tuple[str, ...] = (
    "--case",
    ",".join(XLSX_XML_BORROWED_CASES),
    "--xlsx-shape",
    XLSX_XML_BORROWED_SHAPE,
    "--xlsx-cell-crud-shape",
    XLSX_XML_BORROWED_CELL_CRUD_SHAPE,
)
DOCX_SEMANTIC_ID = "docx-semantic"
DOCX_SEMANTIC_ID_ALIASES: tuple[str, ...] = (
    DOCX_SEMANTIC_ID,
    "docx-semantic-full-text",
)
DOCX_SEMANTIC_CASES: tuple[str, ...] = (
    "docx_semantic_open",
    "docx_semantic_full_text",
)
DOCX_SEMANTIC_ONE_PARAGRAPH_TEXT_CASE = "docx_semantic_one_paragraph_text"
DOCX_SEMANTIC_CASES_WITH_ONE_PARAGRAPH_TEXT: tuple[str, ...] = (
    *DOCX_SEMANTIC_CASES,
    DOCX_SEMANTIC_ONE_PARAGRAPH_TEXT_CASE,
)
DOCX_SEMANTIC_SHAPE = "large"
DOCX_SEMANTIC_ARGS: tuple[str, ...] = (
    "--case",
    ",".join(DOCX_SEMANTIC_CASES),
    "--semantic-shape",
    DOCX_SEMANTIC_SHAPE,
)
ABBA_LEG_ORDER: tuple[str, ...] = ("A1", "B1", "B2", "A2")
ABBA_LEG_VARIANTS: dict[str, str] = {
    "A1": "control",
    "B1": "candidate",
    "B2": "candidate",
    "A2": "control",
}
RESOURCE_METRIC_SPECS: tuple[tuple[str, str], ...] = (
    (
        "harness.elapsed_ns.p50",
        "instrumented harness operation-summary elapsed time (ns); not latency evidence",
    ),
    (
        "harness.elapsed_ns.p95",
        "instrumented harness operation-summary elapsed time (ns); not latency evidence",
    ),
    (
        "harness.elapsed_ns.p99",
        "instrumented harness operation-summary elapsed time (ns); not latency evidence",
    ),
    (
        "harness.elapsed_ns.mean",
        "instrumented harness operation-summary elapsed time (ns); not latency evidence",
    ),
    (
        "harness.elapsed_ns.standard_deviation",
        "instrumented harness operation-summary elapsed time (ns); not latency evidence",
    ),
    ("time.max_rss_kib", "whole-process maximum resident set size (KiB)"),
    ("time.user_seconds", "whole-process user CPU time (s)"),
    ("time.system_seconds", "whole-process system CPU time (s)"),
    (
        "time.voluntary_context_switches",
        "whole-process voluntary context switches",
    ),
    (
        "time.involuntary_context_switches",
        "whole-process involuntary context switches",
    ),
    ("time.major_page_faults", "whole-process major page faults"),
    ("time.minor_page_faults", "whole-process minor page faults"),
    ("heaptrack.allocation_calls", "whole-process heaptrack allocation calls"),
    (
        "heaptrack.temporary_allocations",
        "whole-process heaptrack temporary allocations",
    ),
    ("heaptrack.allocated_bytes", "whole-process heaptrack allocated bytes"),
    ("heaptrack.peak_heap_bytes", "whole-process heaptrack peak heap bytes"),
    ("heaptrack.peak_rss_bytes", "whole-process heaptrack peak RSS bytes"),
)
def _docx_resource_metric_specs(
    cases: Sequence[str],
) -> tuple[tuple[str, str], ...]:
    return tuple(
        [
            (
                f"harness.{case}.elapsed_ns.{statistic}",
                f"instrumented {case} operation-summary elapsed time ({statistic}); "
                "not latency evidence",
            )
            for case in cases
            for statistic in ("p50", "p95", "p99", "mean", "standard_deviation")
        ]
        + [
            spec
            for spec in RESOURCE_METRIC_SPECS
            if not spec[0].startswith("harness.elapsed_ns.")
        ]
    )


DOCX_RESOURCE_METRIC_SPECS: tuple[tuple[str, str], ...] = _docx_resource_metric_specs(
    DOCX_SEMANTIC_CASES
)


def _normalize_xlsx_xml_borrowed_cases(
    cases: Sequence[str] | None = None,
) -> tuple[str, ...]:
    selected = XLSX_XML_BORROWED_CASES if cases is None else tuple(cases)
    if selected != XLSX_XML_BORROWED_CASES:
        raise ResourceProfileInputError(
            "XLSX XML borrowed resource cases must be the fixed four-case tuple"
        )
    return selected


def xlsx_xml_borrowed_cases() -> tuple[str, ...]:
    """Return the fixed borrowed-parser case tuple used by resource ABBA."""
    return XLSX_XML_BORROWED_CASES


def xlsx_xml_borrowed_args(
    cases: Sequence[str] = XLSX_XML_BORROWED_CASES,
) -> tuple[str, ...]:
    """Build the exact tiny/medium harness arguments for borrowed-parser ABBA."""
    selected = _normalize_xlsx_xml_borrowed_cases(cases)
    if selected == XLSX_XML_BORROWED_CASES:
        return XLSX_XML_BORROWED_ARGS
    # Keep the construction explicit if the fixed tuple is ever reviewed; the
    # normalizer currently rejects this branch for strict identity.
    return (
        "--case",
        ",".join(selected),
        "--xlsx-shape",
        XLSX_XML_BORROWED_SHAPE,
        "--xlsx-cell-crud-shape",
        XLSX_XML_BORROWED_CELL_CRUD_SHAPE,
    )


XLSX_XML_BORROWED_RESOURCE_METRIC_SPECS: tuple[tuple[str, str], ...] = (
    _docx_resource_metric_specs(XLSX_XML_BORROWED_CASES)
)


def _normalize_docx_cases(cases: Sequence[str] | None) -> tuple[str, ...]:
    selected = DOCX_SEMANTIC_CASES if cases is None else tuple(cases)
    allowed = (
        DOCX_SEMANTIC_CASES,
        DOCX_SEMANTIC_CASES_WITH_ONE_PARAGRAPH_TEXT,
    )
    if selected not in allowed:
        raise ResourceProfileInputError(
            "DOCX resource cases must be the default open/full-text pair or "
            "that pair plus docx_semantic_one_paragraph_text"
        )
    return selected


def docx_semantic_cases(*, include_one_paragraph_text: bool = False) -> tuple[str, ...]:
    """Return the fixed DOCX resource case tuple selected by the CLI flag."""
    return (
        DOCX_SEMANTIC_CASES_WITH_ONE_PARAGRAPH_TEXT
        if include_one_paragraph_text
        else DOCX_SEMANTIC_CASES
    )


def docx_semantic_args(cases: Sequence[str] = DOCX_SEMANTIC_CASES) -> tuple[str, ...]:
    """Build the harness arguments for one of the fixed DOCX case tuples."""
    selected = _normalize_docx_cases(cases)
    return (
        "--case",
        ",".join(selected),
        "--semantic-shape",
        DOCX_SEMANTIC_SHAPE,
    )


def docx_resource_metric_specs(
    cases: Sequence[str] = DOCX_SEMANTIC_CASES,
) -> tuple[tuple[str, str], ...]:
    """Return resource metric definitions for the selected fixed DOCX cases."""
    return _docx_resource_metric_specs(_normalize_docx_cases(cases))
NOT_MEASURED_RESOURCE_DIMENSIONS: dict[str, str] = {
    "memory_copy_bytes": (
        "not measured: /usr/bin/time and heaptrack report process totals, not bytes copied"
    ),
    "decompressed_bytes": (
        "not measured: no profiler counter is interpreted as decompressed payload bytes"
    ),
    "physical_cold_io": (
        "not measured: this mode does not flush or otherwise establish cold physical I/O"
    ),
}
LATENCY_SEPARATION = {
    "status": "not_measured",
    "reason": (
        "resource legs are instrumented by /usr/bin/time and optionally heaptrack; "
        "their harness elapsed_ns values are not latency evidence"
    ),
    "required_source": (
        "run a separate uninstrumented latency harness with the same binary and corpus identity"
    ),
}


class ResourceProfileInputError(ValueError):
    """Raised when resource-profile comparison inputs are not comparable."""


SHA256_HEX_RE = re.compile(r"[0-9a-fA-F]{64}\Z")
HARNESS_TOOL_NAME = "litchi-perf-baseline"
HARNESS_TOOL_VERSION = "0.1.0"
HARNESS_TOOL_PROFILE = "release"
HARNESS_REVISION_RE = re.compile(r"[0-9a-f]{40}\Z")
# These are the fields emitted by tools/perf-baseline/src/main.rs.  They are
# deliberately required rather than inferred from the Python runner's host:
# the binary that produced a leg is the authority for compiler, target, host,
# allocator, and build-environment identity.
HARNESS_ENVIRONMENT_FIELDS: tuple[str, ...] = (
    "rustc_version",
    "logical_cpus_available",
    "allocator",
    "rustflags",
    "cargo_build_target",
    "perf_event_paranoid",
    "os",
    "kernel",
    "cpu_model",
    "total_memory_bytes",
    "page_size_bytes",
    "filesystem_type",
    "source_destination_same_device",
    "cpu_affinity",
    "storage_identifier",
)
HARNESS_ENVIRONMENT_EXCLUDED_FIELDS = {"git_revision", "git_worktree_dirty"}
ELAPSED_STATISTIC_FIELDS: tuple[str, ...] = (
    "min",
    "p50",
    "p95",
    "p99",
    "max",
    "mean",
    "standard_deviation",
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
TIME_EXPECTED_FIELDS: tuple[str, ...] = tuple(TIME_FIELDS.values()) + (
    "elapsed_wall_seconds",
)
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
    # Keep this allow-list deliberately small.  It records knobs that can
    # alter scheduling/allocation without serializing ambient credentials,
    # paths, or arbitrary user configuration into a retained report.
    environment_keys = (
        "LANG",
        "LC_ALL",
        "LC_CTYPE",
        "RUSTFLAGS",
        "CARGO_BUILD_JOBS",
        "RAYON_NUM_THREADS",
        "MALLOC_CONF",
    )
    selected_environment = {
        key: os.environ[key] for key in environment_keys if key in os.environ
    }
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
        "selected_environment": selected_environment,
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


def parse_time_report(path: Path, *, retained: bool = False) -> dict[str, Any]:
    parsed: dict[str, Any] = {"status": "missing"}
    malformed = False
    try:
        lines = path.read_text(encoding="utf-8", errors="replace").splitlines()
    except OSError as error:
        return {"status": "missing", "error": str(error)}
    for raw_line in lines:
        line = raw_line.strip()
        if line.startswith("Elapsed (wall clock") and ": " in line:
            elapsed = parse_elapsed_clock(line.rpartition(": ")[2].strip())
            parsed["elapsed_wall_seconds"] = elapsed
            malformed = malformed or elapsed is None
            continue
        if ":" not in line:
            continue
        label, raw = line.split(":", 1)
        label = label.strip()
        raw = raw.strip()
        key = TIME_FIELDS.get(label)
        if key is not None:
            try:
                value: float | int
                if "time" in label.lower():
                    value = float(raw)
                    if not math.isfinite(value):
                        raise ValueError("non-finite time value")
                else:
                    value = int(raw.replace(",", ""))
                parsed[key] = value
            except (OverflowError, ValueError):
                parsed[key] = None
                malformed = True
    missing = [key for key in TIME_EXPECTED_FIELDS if key not in parsed]
    parsed["expected_fields"] = list(TIME_EXPECTED_FIELDS)
    if malformed:
        parsed["status"] = "unparsed"
    elif not any(key in parsed for key in TIME_EXPECTED_FIELDS):
        parsed["status"] = "unparsed"
    elif missing:
        parsed["status"] = "unavailable"
        parsed["reason"] = (
            "incomplete GNU time -v report; missing expected fields: "
            + ", ".join(missing)
        )
        parsed["missing_fields"] = missing
    else:
        parsed["status"] = "ok"
    parsed["artifact"] = artifact(path, retained=retained)
    return parsed


def parse_elapsed_clock(value: str) -> float | None:
    try:
        parts = value.split(":")
        if len(parts) == 3:
            hours, minutes, seconds = parts
            result = int(hours) * 3600 + int(minutes) * 60 + float(seconds)
            return result if math.isfinite(result) else None
        if len(parts) == 2:
            minutes, seconds = parts
            result = int(minutes) * 60 + float(seconds)
            return result if math.isfinite(result) else None
        result = float(value)
        return result if math.isfinite(result) else None
    except (OverflowError, ValueError):
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


def parse_heaptrack_print(path: Path, *, retained: bool = False) -> dict[str, Any]:
    """Parse stable process-total fields from heaptrack_print -H output."""
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        return {
            "status": "missing",
            "error": str(error),
            "artifact": artifact(path, retained=retained),
        }
    patterns: tuple[tuple[str, str, Any], ...] = (
        (
            "allocation_calls",
            r"^calls to allocation functions:\s*([0-9,]+)",
            _numeric_token,
        ),
        (
            "temporary_allocations",
            r"^temporary memory allocations:\s*([0-9,]+)(?:\s|$)",
            _numeric_token,
        ),
        (
            "peak_heap_bytes",
            r"^peak heap memory consumption:\s*([0-9.,]+\s*[KMGT]?(?:i?B)?)",
            _bytes_token,
        ),
        (
            "peak_rss_bytes",
            r"^peak RSS(?:\s*\([^)]*\))?:\s*([0-9.,]+\s*[KMGT]?(?:i?B)?)",
            _bytes_token,
        ),
    )
    parsed: dict[str, Any] = {
        "status": "ok",
        "artifact": artifact(path, retained=retained),
    }
    for key, pattern, converter in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE)
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


def _require_retained_artifact_identity(
    expected: Any,
    path: Path,
    *,
    location: str,
) -> dict[str, Any]:
    """Require retained bytes to match the identity published by an earlier report."""
    if not isinstance(expected, dict):
        raise ResourceProfileInputError(f"{location} must be an artifact object")
    observed = artifact(path, retained=True)
    for key in ("present", "bytes", "sha256"):
        if expected.get(key) != observed.get(key):
            raise ResourceProfileInputError(
                f"{location}.{key} does not match retained artifact {path}"
            )
    if observed["present"] is not True:
        raise ResourceProfileInputError(f"{location} retained artifact is missing: {path}")
    return observed


def run_reprocess_docx_heaptrack(arguments: argparse.Namespace) -> int:
    """Reparse hash-verified DOCX heaptrack text and refresh derived statistics."""
    input_path = Path(arguments.input).expanduser().resolve()
    output_path = Path(arguments.output).expanduser().resolve()
    if input_path == output_path:
        raise ResourceProfileInputError("reprocessed output must differ from the input report")
    if os.path.lexists(output_path):
        raise ResourceProfileInputError(f"reprocessed output already exists: {output_path}")
    report = load_json(input_path)
    if not isinstance(report, dict):
        raise ResourceProfileInputError("input report must be an object")
    tool = report.get("tool")
    if not isinstance(tool, dict) or tool.get("mode") != "docx-semantic-abba-resource-profile":
        raise ResourceProfileInputError("input is not a DOCX semantic ABBA resource report")
    scope = report.get("scope")
    if not isinstance(scope, dict):
        raise ResourceProfileInputError("input report scope must be an object")
    cases = _normalize_docx_cases(scope.get("cases"))
    legs = report.get("legs")
    if not isinstance(legs, list):
        raise ResourceProfileInputError("input report legs must be an array")
    validate_abba_order([leg.get("leg") if isinstance(leg, dict) else None for leg in legs])
    artifact_root_value = report.get("artifact_directory")
    if not isinstance(artifact_root_value, str) or not artifact_root_value:
        raise ResourceProfileInputError("input report artifact_directory must be a path")
    artifact_root = Path(artifact_root_value).expanduser().resolve()

    for leg in legs:
        label = leg["leg"]
        leg_dir = artifact_root / label.lower()
        declared_leg_dir = leg.get("artifact_directory")
        if not isinstance(declared_leg_dir, str) or Path(declared_leg_dir).expanduser().resolve() != leg_dir:
            raise ResourceProfileInputError(
                f"legs.{label}.artifact_directory does not match the canonical artifact path"
            )
        heaptrack = leg.get("heaptrack")
        printed = heaptrack.get("print") if isinstance(heaptrack, dict) else None
        parsed_before = printed.get("parsed") if isinstance(printed, dict) else None
        if not isinstance(parsed_before, dict):
            raise ResourceProfileInputError(f"legs.{label}.heaptrack.print.parsed is required")
        summary_path = leg_dir / "heaptrack-print.txt"
        histogram_path = leg_dir / "heaptrack-histogram.tsv"
        _require_retained_artifact_identity(
            printed.get("artifact"),
            summary_path,
            location=f"legs.{label}.heaptrack.print.artifact",
        )
        _require_retained_artifact_identity(
            parsed_before.get("histogram_artifact"),
            histogram_path,
            location=f"legs.{label}.heaptrack.print.parsed.histogram_artifact",
        )
        parsed = parse_heaptrack_print(summary_path, retained=True)
        parsed["allocated_bytes"] = parse_heaptrack_histogram(histogram_path)
        parsed["histogram_artifact"] = artifact(histogram_path, retained=True)
        parsed["scope"] = "whole process; heaptrack instrumentation overhead is included"
        printed["parsed"] = parsed

    report["statistics"] = abba_statistics(
        legs,
        metric_specs=docx_resource_metric_specs(cases),
    )
    report["reprocessing"] = {
        "source_report": artifact(input_path, retained=True),
        "raw_heaptrack_artifacts_verified": True,
        "operation": "reparse process-total heaptrack fields and refresh derived statistics",
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("x", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, sort_keys=True, allow_nan=False)
        handle.write("\n")
    return 0


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


def instrumented_harness_metrics(
    report: dict[str, Any], cases: Sequence[str]
) -> dict[str, Any]:
    """Extract per-case elapsed summaries with an explicit instrumented label.

    These values are retained only to align resource legs.  They are collected
    inside the same process invocation as ``/usr/bin/time``/heaptrack and must
    not be consumed as uninstrumented latency evidence.
    """
    results = report.get("results") if isinstance(report, dict) else None
    if not isinstance(results, list):
        return {}
    wanted = set(cases)
    metrics: dict[str, Any] = {}
    for result in results:
        if not isinstance(result, dict) or result.get("case") not in wanted:
            continue
        elapsed = result.get("elapsed_ns")
        if not isinstance(elapsed, dict):
            continue
        case = str(result["case"])
        for statistic in ("p50", "p95", "p99", "mean", "standard_deviation"):
            if statistic in elapsed:
                metrics[f"harness.{case}.elapsed_ns.{statistic}"] = elapsed[statistic]
    return metrics


def _canonical_json(value: Any, location: str) -> str:
    try:
        return json.dumps(value, sort_keys=True, separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        raise ResourceProfileInputError(f"{location} is not canonical JSON: {error}") from error


def validate_abba_order(labels: Sequence[str]) -> tuple[str, ...]:
    """Require the fixed control/candidate/control ABBA execution order."""
    observed = tuple(str(label) for label in labels)
    if observed != ABBA_LEG_ORDER:
        raise ResourceProfileInputError(
            f"ABBA leg order must be {list(ABBA_LEG_ORDER)!r}; got {list(observed)!r}"
        )
    return observed


def _validate_sha256(value: Any, location: str) -> str:
    if not isinstance(value, str) or SHA256_HEX_RE.fullmatch(value) is None:
        raise ResourceProfileInputError(
            f"{location} must contain exactly 64 hexadecimal SHA-256 characters"
        )
    return value.lower()


def binary_identity(path: Path, *, label: str) -> dict[str, Any]:
    """Return exact identity for a release binary, failing before any run."""
    resolved = path.expanduser().resolve()
    if not resolved.is_file():
        raise ResourceProfileInputError(f"{label} binary does not exist: {resolved}")
    if not os.access(resolved, os.X_OK):
        raise ResourceProfileInputError(f"{label} binary is not executable: {resolved}")
    try:
        digest = sha256_file(resolved)
        binary_stat = resolved.stat()
        size = binary_stat.st_size
        mode_bits = stat.S_IMODE(binary_stat.st_mode)
    except OSError as error:
        raise ResourceProfileInputError(f"cannot inspect {label} binary {resolved}: {error}") from error
    _validate_sha256(digest, f"{label} binary SHA-256")
    return {
        "label": label,
        "path": str(resolved),
        "binary_sha256": digest,
        "binary_bytes": size,
        "mode_bits": mode_bits,
        "executable": True,
    }


def _validate_binary_descriptor(
    binary: Any,
    *,
    location: str,
    expected_label: str | None = None,
) -> dict[str, Any]:
    """Verify a run-produced binary descriptor against the current executable."""
    if not isinstance(binary, dict):
        raise ResourceProfileInputError(f"{location} must be an object")
    label = binary.get("label")
    if expected_label is not None and label != expected_label:
        raise ResourceProfileInputError(
            f"{location}.label must be {expected_label!r}; got {label!r}"
        )
    path_value = binary.get("path")
    if not isinstance(path_value, str) or not path_value:
        raise ResourceProfileInputError(f"{location}.path must be a non-empty string")
    path = Path(path_value)
    if not path.is_absolute():
        raise ResourceProfileInputError(f"{location}.path must be absolute")
    try:
        path_stat = path.lstat()
    except OSError as error:
        raise ResourceProfileInputError(f"{location}.path is not readable: {error}") from error
    if stat.S_ISLNK(path_stat.st_mode) or not stat.S_ISREG(path_stat.st_mode):
        raise ResourceProfileInputError(f"{location}.path must reference a regular file")
    if not os.access(path, os.X_OK):
        raise ResourceProfileInputError(f"{location}.path must reference an executable file")
    if binary.get("executable") is not True:
        raise ResourceProfileInputError(f"{location}.executable must be true")
    expected_mode = binary.get("mode_bits")
    if (
        isinstance(expected_mode, bool)
        or not isinstance(expected_mode, int)
        or expected_mode < 0
        or expected_mode > 0o7777
    ):
        raise ResourceProfileInputError(f"{location}.mode_bits must be permission bits")
    expected_size = binary.get("binary_bytes")
    if (
        isinstance(expected_size, bool)
        or not isinstance(expected_size, int)
        or expected_size < 0
    ):
        raise ResourceProfileInputError(f"{location}.binary_bytes must be a non-negative integer")
    expected_digest = _validate_sha256(binary.get("binary_sha256"), f"{location}.binary_sha256")
    try:
        actual_size = path_stat.st_size
        actual_mode = stat.S_IMODE(path_stat.st_mode)
        actual_digest = sha256_file(path)
    except OSError as error:
        raise ResourceProfileInputError(f"cannot hash {location}.path: {error}") from error
    if actual_size != expected_size:
        raise ResourceProfileInputError(
            f"{location}.binary_bytes does not match executable: "
            f"{expected_size} != {actual_size}"
        )
    if actual_mode != expected_mode:
        raise ResourceProfileInputError(
            f"{location}.mode_bits does not match executable: "
            f"{expected_mode:o} != {actual_mode:o}"
        )
    if actual_digest.lower() != expected_digest:
        raise ResourceProfileInputError(
            f"{location}.binary_sha256 does not match executable"
        )
    return {
        **binary,
        "path": str(path),
        "binary_sha256": expected_digest,
        "binary_bytes": actual_size,
        "mode_bits": actual_mode,
        "executable": True,
    }


def reserve_abba_paths(output_path: Path, artifact_root: Path) -> tuple[Path, Path]:
    """Reserve fresh ABBA output/artifact paths before launching any workload."""
    output = output_path.expanduser().resolve()
    root = artifact_root.expanduser().resolve()
    if output == root:
        raise ResourceProfileInputError("ABBA output path and artifact root must differ")
    try:
        output.relative_to(root)
        root_contains_output = True
    except ValueError:
        root_contains_output = False
    try:
        root.relative_to(output)
        output_contains_root = True
    except ValueError:
        output_contains_root = False
    if root_contains_output or output_contains_root:
        raise ResourceProfileInputError(
            "ABBA output path and artifact root must not contain one another"
        )
    if os.path.lexists(output):
        raise ResourceProfileInputError(f"ABBA output path already exists: {output}")
    if os.path.lexists(root):
        raise ResourceProfileInputError(
            f"ABBA artifact root already exists (possible stale capture): {root}"
        )
    try:
        output.parent.mkdir(parents=True, exist_ok=True)
        root.mkdir(parents=True, exist_ok=False)
    except OSError as error:
        raise ResourceProfileInputError(f"cannot reserve ABBA output paths: {error}") from error
    return output, root


def _harness_tool_identity(report: Any, location: str) -> dict[str, Any]:
    tool = report.get("tool") if isinstance(report, dict) else None
    if not isinstance(tool, dict):
        raise ResourceProfileInputError(f"{location}.tool must be an object")
    if tool.get("name") != HARNESS_TOOL_NAME:
        raise ResourceProfileInputError(
            f"{location}.tool.name must be {HARNESS_TOOL_NAME!r}"
        )
    if tool.get("version") != HARNESS_TOOL_VERSION:
        raise ResourceProfileInputError(
            f"{location}.tool.version must be {HARNESS_TOOL_VERSION!r}"
        )
    if tool.get("profile") != HARNESS_TOOL_PROFILE:
        raise ResourceProfileInputError(
            f"{location}.tool.profile must be {HARNESS_TOOL_PROFILE!r}"
        )
    for key in ("target_os", "target_arch"):
        value = tool.get(key)
        if not isinstance(value, str) or not value:
            raise ResourceProfileInputError(
                f"{location}.tool.{key} must be a non-empty string"
            )
    _canonical_json(tool, f"{location}.tool")
    return dict(tool)


def _harness_environment_identity(report: Any, location: str) -> dict[str, Any]:
    environment = report.get("environment") if isinstance(report, dict) else None
    if not isinstance(environment, dict):
        raise ResourceProfileInputError(f"{location}.environment must be an object")
    for key in HARNESS_ENVIRONMENT_FIELDS:
        if key not in environment:
            raise ResourceProfileInputError(
                f"{location}.environment.{key} is required for stable harness identity"
            )
    identity = {
        key: value
        for key, value in environment.items()
        if key not in HARNESS_ENVIRONMENT_EXCLUDED_FIELDS
    }
    _canonical_json(identity, f"{location}.environment")
    return identity


def _requested_sample_count(configuration: dict[str, Any], location: str) -> int:
    samples = configuration.get("samples_per_case")
    if isinstance(samples, bool) or not isinstance(samples, int) or samples <= 0:
        raise ResourceProfileInputError(
            f"{location}.samples_per_case must be a positive integer"
        )
    warmup = configuration.get("warmup_iterations_per_case")
    if isinstance(warmup, bool) or not isinstance(warmup, int) or warmup < 0:
        raise ResourceProfileInputError(
            f"{location}.warmup_iterations_per_case must be a non-negative integer"
        )
    return samples


def _finite_nonnegative(value: Any) -> bool:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return False
    try:
        number = float(value)
    except (OverflowError, ValueError):
        return False
    return math.isfinite(number) and number >= 0


def _validate_elapsed_statistics(
    elapsed: Any, *, sample_count: int, location: str
) -> None:
    if not isinstance(elapsed, dict):
        raise ResourceProfileInputError(f"{location} is required")
    if elapsed.get("unit") != "ns":
        raise ResourceProfileInputError(f"{location}.unit must be 'ns'")
    samples = elapsed.get("samples")
    if not isinstance(samples, list) or len(samples) != sample_count:
        observed = len(samples) if isinstance(samples, list) else None
        raise ResourceProfileInputError(
            f"{location}.sample_count must equal requested {sample_count}; got {observed!r}"
        )
    for index, value in enumerate(samples):
        if (
            isinstance(value, bool)
            or not isinstance(value, int)
            or value < 0
            or not _finite_nonnegative(value)
        ):
            raise ResourceProfileInputError(
                f"{location}.samples[{index}] must be a finite non-negative integer"
            )
    for key in ELAPSED_STATISTIC_FIELDS:
        value = elapsed.get(key)
        if not _finite_nonnegative(value):
            raise ResourceProfileInputError(
                f"{location}.{key} must be a finite non-negative number"
            )
    confidence = elapsed.get("confidence_interval_95")
    if not isinstance(confidence, dict):
        raise ResourceProfileInputError(
            f"{location}.confidence_interval_95 is required"
        )
    if not isinstance(confidence.get("method"), str) or not confidence["method"]:
        raise ResourceProfileInputError(
            f"{location}.confidence_interval_95.method must be a non-empty string"
        )
    for key in ("lower", "upper"):
        if not _finite_nonnegative(confidence.get(key)):
            raise ResourceProfileInputError(
                f"{location}.confidence_interval_95.{key} must be finite"
            )
    reported_count = elapsed.get("sample_count")
    if reported_count is not None and (
        isinstance(reported_count, bool)
        or not isinstance(reported_count, int)
        or reported_count != sample_count
    ):
        raise ResourceProfileInputError(
            f"{location}.sample_count must equal requested {sample_count}"
        )


def _harness_result(report: Any, location: str) -> tuple[dict[str, Any], dict[str, Any]]:
    if not isinstance(report, dict):
        raise ResourceProfileInputError(f"{location} must be an object")
    if report.get("schema_version") != SCHEMA_VERSION:
        raise ResourceProfileInputError(
            f"{location}.schema_version must be {SCHEMA_VERSION!r}"
        )
    configuration = report.get("configuration")
    if not isinstance(configuration, dict):
        raise ResourceProfileInputError(f"{location}.configuration must be an object")
    sample_count = _requested_sample_count(configuration, f"{location}.configuration")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != 1:
        raise ResourceProfileInputError(
            f"{location}.results must contain exactly one XLSX managed-batch result"
        )
    result = results[0]
    if not isinstance(result, dict):
        raise ResourceProfileInputError(f"{location}.results[0] must be an object")
    if result.get("case") != XLSX_MANAGED_BATCH_CASE:
        raise ResourceProfileInputError(
            f"{location}.results[0].case must be {XLSX_MANAGED_BATCH_CASE!r}"
        )
    corpus = result.get("corpus")
    if not isinstance(corpus, dict) or not corpus:
        raise ResourceProfileInputError(f"{location}.results[0].corpus must be a non-empty object")
    if corpus.get("shape") != XLSX_MANAGED_BATCH_SHAPE:
        raise ResourceProfileInputError(
            f"{location}.results[0].corpus.shape must be {XLSX_MANAGED_BATCH_SHAPE!r}"
        )
    _validate_elapsed_statistics(
        result.get("elapsed_ns"),
        sample_count=sample_count,
        location=f"{location}.results[0].elapsed_ns",
    )
    return configuration, result


def _docx_harness_results(
    report: Any,
    location: str,
    *,
    cases: Sequence[str] = DOCX_SEMANTIC_CASES,
) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    """Validate the selected fixed DOCX semantic rows before comparison.

    The resource runner deliberately does not infer a corpus from a command
    line.  Every leg must carry the selected semantic rows and a complete
    deterministic corpus manifest, including archive and target hashes.  This
    keeps an accidentally changed generator, shape, or fixture from becoming a
    seemingly comparable resource result.
    """
    selected_cases = _normalize_docx_cases(cases)
    if not isinstance(report, dict):
        raise ResourceProfileInputError(f"{location} must be an object")
    if report.get("schema_version") != SCHEMA_VERSION:
        raise ResourceProfileInputError(
            f"{location}.schema_version must be {SCHEMA_VERSION!r}"
        )
    configuration = report.get("configuration")
    if not isinstance(configuration, dict):
        raise ResourceProfileInputError(f"{location}.configuration must be an object")
    sample_count = _requested_sample_count(configuration, f"{location}.configuration")
    results = report.get("results")
    if not isinstance(results, list) or len(results) != len(selected_cases):
        raise ResourceProfileInputError(
            f"{location}.results must contain exactly the selected DOCX semantic rows"
        )
    observed_cases = [
        result.get("case") if isinstance(result, dict) else None for result in results
    ]
    if tuple(observed_cases) != selected_cases:
        raise ResourceProfileInputError(
            f"{location}.results cases must be {list(selected_cases)!r}; "
            f"got {observed_cases!r}"
        )
    validated: list[dict[str, Any]] = []
    required_manifest_keys = (
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
    )
    for index, result in enumerate(results):
        result_location = f"{location}.results[{index}]"
        if not isinstance(result, dict):
            raise ResourceProfileInputError(f"{result_location} must be an object")
        _validate_elapsed_statistics(
            result.get("elapsed_ns"),
            sample_count=sample_count,
            location=f"{result_location}.elapsed_ns",
        )
        corpus = result.get("corpus")
        if not isinstance(corpus, dict) or not corpus:
            raise ResourceProfileInputError(
                f"{result_location}.corpus must be a non-empty object"
            )
        for key in required_manifest_keys:
            if key not in corpus:
                raise ResourceProfileInputError(
                    f"{result_location}.corpus.{key} is required for DOCX identity"
                )
        if corpus.get("shape") != DOCX_SEMANTIC_SHAPE:
            raise ResourceProfileInputError(
                f"{result_location}.corpus.shape must be {DOCX_SEMANTIC_SHAPE!r}"
            )
        if corpus.get("generator") != "litchi-docx-semantic-v1":
            raise ResourceProfileInputError(
                f"{result_location}.corpus.generator is not the fixed DOCX generator"
            )
        if corpus.get("package_format") != "DOCX/OPC/ZIP":
            raise ResourceProfileInputError(
                f"{result_location}.corpus.package_format is not DOCX/OPC/ZIP"
            )
        for key in (
            "entry_count",
            "archive_member_count",
            "entry_bytes",
            "uncompressed_payload_bytes",
            "archive_bytes",
            "target_payload_bytes",
        ):
            value = corpus.get(key)
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise ResourceProfileInputError(
                    f"{result_location}.corpus.{key} must be a non-negative integer"
                )
        _validate_sha256(
            corpus.get("archive_sha256"), f"{result_location}.corpus.archive_sha256"
        )
        _validate_sha256(
            corpus.get("target_payload_sha256"),
            f"{result_location}.corpus.target_payload_sha256",
        )
        validated.append(result)
    within_report_identity = _canonical_json(
        validated[0]["corpus"], f"{location}.results[0].corpus"
    )
    if any(
        _canonical_json(result["corpus"], f"{location}.corpus")
        != within_report_identity
        for result in validated[1:]
    ):
        raise ResourceProfileInputError(
            f"{location}.results DOCX corpus identities do not match"
        )
    return configuration, validated


XLSX_XML_BORROWED_RESULT_IDENTITY: dict[str, tuple[bool, bool, bool]] = {
    # source, sink, and output are deliberately separate identity channels;
    # an absent channel is meaningful and must remain absent in every leg.
    "xlsx_first_cell": (False, False, False),
    "xlsx_source_first_cell": (True, False, False),
    "xlsx_eager_cell_values_one_edit_save": (False, True, True),
    "xlsx_source_backed_cell_values_one_edit_save": (True, True, True),
}


def _xlsx_xml_borrowed_harness_results(
    report: Any,
    location: str,
    *,
    cases: Sequence[str] = XLSX_XML_BORROWED_CASES,
) -> tuple[dict[str, Any], list[dict[str, Any]], list[dict[str, Any]]]:
    """Validate the exact four-row XLSX borrowed-parser harness report.

    The tiny read selectors and medium cell-edit selectors intentionally use
    different deterministic corpora.  Identity is therefore checked per
    case, then compared across ABBA legs, rather than incorrectly requiring
    all four rows in one report to share one archive.
    """
    selected_cases = _normalize_xlsx_xml_borrowed_cases(cases)
    if not isinstance(report, dict):
        raise ResourceProfileInputError(f"{location} must be an object")
    if report.get("schema_version") != SCHEMA_VERSION:
        raise ResourceProfileInputError(
            f"{location}.schema_version must be {SCHEMA_VERSION!r}"
        )
    configuration = report.get("configuration")
    if not isinstance(configuration, dict):
        raise ResourceProfileInputError(f"{location}.configuration must be an object")
    sample_count = _requested_sample_count(configuration, f"{location}.configuration")
    expected_configuration = {
        "cases": list(selected_cases),
        "xlsx_shapes": [XLSX_XML_BORROWED_SHAPE],
        "xlsx_cell_crud_shapes": [XLSX_XML_BORROWED_CELL_CRUD_SHAPE],
    }
    for key, expected in expected_configuration.items():
        if configuration.get(key) != expected:
            raise ResourceProfileInputError(
                f"{location}.configuration.{key} must be {expected!r}; "
                f"got {configuration.get(key)!r}"
            )
    results = report.get("results")
    if not isinstance(results, list) or len(results) != len(selected_cases):
        raise ResourceProfileInputError(
            f"{location}.results must contain exactly the four XLSX borrowed-parser rows"
        )
    observed_cases = [
        result.get("case") if isinstance(result, dict) else None for result in results
    ]
    if tuple(observed_cases) != selected_cases:
        raise ResourceProfileInputError(
            f"{location}.results cases must be {list(selected_cases)!r}; "
            f"got {observed_cases!r}"
        )

    required_manifest_keys = (
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
        "xlsx",
    )
    expected_corpus = {
        "xlsx_first_cell": {
            "name": "xlsx-tiny",
            "generator": "litchi-xlsx-synthetic-v1",
            "shape": XLSX_XML_BORROWED_SHAPE,
            "payload_kind": "deterministic-integer-grid",
        },
        "xlsx_source_first_cell": {
            "name": "xlsx-tiny",
            "generator": "litchi-xlsx-synthetic-v1",
            "shape": XLSX_XML_BORROWED_SHAPE,
            "payload_kind": "deterministic-integer-grid",
        },
        "xlsx_eager_cell_values_one_edit_save": {
            "name": "xlsx-cell-values-medium",
            "generator": "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
            "shape": XLSX_XML_BORROWED_CELL_CRUD_SHAPE,
            "payload_kind": "deterministic-multi-sheet-scalar-grid-with-media",
        },
        "xlsx_source_backed_cell_values_one_edit_save": {
            "name": "xlsx-cell-values-medium",
            "generator": "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1",
            "shape": XLSX_XML_BORROWED_CELL_CRUD_SHAPE,
            "payload_kind": "deterministic-multi-sheet-scalar-grid-with-media",
        },
    }
    validated: list[dict[str, Any]] = []
    result_identities: list[dict[str, Any]] = []
    for index, result in enumerate(results):
        result_location = f"{location}.results[{index}]"
        if not isinstance(result, dict):
            raise ResourceProfileInputError(f"{result_location} must be an object")
        case = result["case"]
        _validate_elapsed_statistics(
            result.get("elapsed_ns"),
            sample_count=sample_count,
            location=f"{result_location}.elapsed_ns",
        )
        corpus = result.get("corpus")
        if not isinstance(corpus, dict) or not corpus:
            raise ResourceProfileInputError(
                f"{result_location}.corpus must be a non-empty object"
            )
        for key in required_manifest_keys:
            if key not in corpus:
                raise ResourceProfileInputError(
                    f"{result_location}.corpus.{key} is required for XLSX identity"
                )
        expected = expected_corpus[case]
        for key, expected_value in expected.items():
            if corpus.get(key) != expected_value:
                raise ResourceProfileInputError(
                    f"{result_location}.corpus.{key} must be {expected_value!r}"
                )
        if corpus.get("package_format") != "XLSX/OPC/ZIP":
            raise ResourceProfileInputError(
                f"{result_location}.corpus.package_format must be 'XLSX/OPC/ZIP'"
            )
        if corpus.get("compression") != "deflate":
            raise ResourceProfileInputError(
                f"{result_location}.corpus.compression must be 'deflate'"
            )
        if corpus.get("target_entry") != "Sheet1!A1":
            raise ResourceProfileInputError(
                f"{result_location}.corpus.target_entry must be 'Sheet1!A1'"
            )
        for key in (
            "entry_count",
            "archive_member_count",
            "entry_bytes",
            "uncompressed_payload_bytes",
            "archive_bytes",
            "target_payload_bytes",
        ):
            value = corpus.get(key)
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise ResourceProfileInputError(
                    f"{result_location}.corpus.{key} must be a non-negative integer"
                )
        _validate_sha256(
            corpus.get("archive_sha256"), f"{result_location}.corpus.archive_sha256"
        )
        _validate_sha256(
            corpus.get("target_payload_sha256"),
            f"{result_location}.corpus.target_payload_sha256",
        )
        xlsx = corpus["xlsx"]
        if not isinstance(xlsx, dict):
            raise ResourceProfileInputError(
                f"{result_location}.corpus.xlsx must be an object"
            )
        for key in (
            "sheet_count",
            "rows_per_sheet",
            "columns_per_sheet",
            "one_percent_update_count",
        ):
            value = xlsx.get(key)
            if isinstance(value, bool) or not isinstance(value, int) or value < 0:
                raise ResourceProfileInputError(
                    f"{result_location}.corpus.xlsx.{key} must be a non-negative integer"
                )
        source_members = xlsx.get("source_members")
        if not isinstance(source_members, dict):
            raise ResourceProfileInputError(
                f"{result_location}.corpus.xlsx.source_members must be an object"
            )
        for key in ("workbook", "styles"):
            value = source_members.get(key)
            if not isinstance(value, str) or not value:
                raise ResourceProfileInputError(
                    f"{result_location}.corpus.xlsx.source_members.{key} must be a string"
                )
        worksheets = source_members.get("worksheets")
        if (
            not isinstance(worksheets, list)
            or not worksheets
            or any(not isinstance(value, str) or not value for value in worksheets)
        ):
            raise ResourceProfileInputError(
                f"{result_location}.corpus.xlsx.source_members.worksheets must be non-empty strings"
            )
        shared_strings = source_members.get("shared_strings")
        if shared_strings is not None and (
            not isinstance(shared_strings, str) or not shared_strings
        ):
            raise ResourceProfileInputError(
                f"{result_location}.corpus.xlsx.source_members.shared_strings must be a string or null"
            )

        expected_source, expected_sink, expected_output = XLSX_XML_BORROWED_RESULT_IDENTITY[case]
        source = result.get("source")
        sink = result.get("sink")
        output = result.get("output_sha256")
        if (source is not None) != expected_source:
            raise ResourceProfileInputError(
                f"{result_location}.source presence does not match fixed {case} identity"
            )
        if (sink is not None) != expected_sink:
            raise ResourceProfileInputError(
                f"{result_location}.sink presence does not match fixed {case} identity"
            )
        if (output is not None) != expected_output:
            raise ResourceProfileInputError(
                f"{result_location}.output_sha256 presence does not match fixed {case} identity"
            )
        if source is not None:
            if not isinstance(source, dict) or not source:
                raise ResourceProfileInputError(
                    f"{result_location}.source must be a non-empty object"
                )
            _canonical_json(source, f"{result_location}.source")
        if sink is not None:
            if not isinstance(sink, dict) or not sink:
                raise ResourceProfileInputError(
                    f"{result_location}.sink must be a non-empty object"
                )
            _canonical_json(sink, f"{result_location}.sink")
        if output is not None:
            _validate_sha256(output, f"{result_location}.output_sha256")
        result_identity = {
            "case": case,
            "source_present": source is not None,
            "source": source,
            "sink_present": sink is not None,
            "sink": sink,
            "output_present": output is not None,
            "output_sha256": output,
        }
        _canonical_json(result_identity, f"{result_location} identity")
        validated.append(result)
        result_identities.append(result_identity)
    return configuration, validated, result_identities


def _require_revision(report: dict[str, Any], location: str) -> str:
    environment = report.get("environment")
    if not isinstance(environment, dict):
        raise ResourceProfileInputError(f"{location}.environment must be an object")
    if environment.get("git_worktree_dirty") is not False:
        raise ResourceProfileInputError(
            f"{location}.environment.git_worktree_dirty must be false"
        )
    revision = environment.get("git_revision")
    if not isinstance(revision, str) or HARNESS_REVISION_RE.fullmatch(revision) is None:
        raise ResourceProfileInputError(
            f"{location}.environment.git_revision must be exactly 40 lowercase hexadecimal characters"
        )
    return revision


def _xlsx_xml_borrowed_harness_identity(
    report: Any,
    location: str,
    *,
    cases: Sequence[str] = XLSX_XML_BORROWED_CASES,
) -> dict[str, Any]:
    """Return the identity channels shared by timed and heaptrack harnesses.

    Elapsed samples are deliberately excluded: the second process is
    instrumented and therefore has different resource-observation timing.
    Configuration, corpus rows, source/sink/output channels, revision, tool,
    and stable environment remain exact identity requirements.
    """
    configuration, results, result_identities = _xlsx_xml_borrowed_harness_results(
        report, location, cases=cases
    )
    revision = _require_revision(report, location)
    environment = _harness_environment_identity(report, location)
    tool = _harness_tool_identity(report, location)
    return {
        "schema_version": report["schema_version"],
        "tool": tool,
        "environment": environment,
        "git_revision": revision,
        "git_worktree_dirty": False,
        "configuration": configuration,
        "corpus": [
            {"case": result["case"], "corpus": result["corpus"]}
            for result in results
        ],
        "result_identities": result_identities,
    }


def validate_abba_inputs(
    legs: Sequence[dict[str, Any]],
    *,
    expected_configuration: dict[str, Any] | None = None,
    workload: str = "xlsx",
    docx_cases: Sequence[str] | None = None,
    xlsx_cases: Sequence[str] | None = None,
) -> dict[str, Any]:
    """Validate clean binary/revision/corpus/config identity for four ABBA legs.

    ``legs`` intentionally contains the parsed harness report and the exact
    binary descriptor used for that leg.  Validation is performed before any
    statistics are interpreted, and a missing identity is an error rather than
    an implicit match.
    """
    if workload not in {"xlsx", DOCX_SEMANTIC_ID, XLSX_XML_BORROWED_ID}:
        raise ResourceProfileInputError(f"unsupported ABBA workload: {workload!r}")
    selected_docx_cases = (
        _normalize_docx_cases(docx_cases)
        if workload == DOCX_SEMANTIC_ID
        else DOCX_SEMANTIC_CASES
    )
    selected_xlsx_cases = (
        _normalize_xlsx_xml_borrowed_cases(xlsx_cases)
        if workload == XLSX_XML_BORROWED_ID
        else XLSX_XML_BORROWED_CASES
    )
    if len(legs) != len(ABBA_LEG_ORDER):
        raise ResourceProfileInputError(
            f"ABBA requires exactly {len(ABBA_LEG_ORDER)} legs; got {len(legs)}"
        )
    if any(not isinstance(leg, dict) for leg in legs):
        raise ResourceProfileInputError("ABBA legs must be objects")
    labels = [leg.get("leg") for leg in legs]
    validate_abba_order(labels)
    revisions: dict[str, str] = {}
    binary_hashes: dict[str, str] = {}
    binary_paths: dict[str, str] = {}
    binary_modes: dict[str, int] = {}
    configurations: list[dict[str, Any]] = []
    corpora: list[Any] = []
    tools: list[Any] = []
    environments: list[dict[str, Any]] = []
    harness_identities: list[dict[str, Any]] = []
    result_identity_sets: list[list[dict[str, Any]]] = []
    for index, leg in enumerate(legs):
        location = f"legs[{index}]"
        if not isinstance(leg, dict):
            raise ResourceProfileInputError(f"{location} must be an object")
        label = leg.get("leg")
        if label not in ABBA_LEG_VARIANTS:
            raise ResourceProfileInputError(f"{location}.leg is not a recognized ABBA leg")
        expected_variant = ABBA_LEG_VARIANTS[label]
        if leg.get("variant") != expected_variant:
            raise ResourceProfileInputError(
                f"{location}.variant must be {expected_variant!r} for {label!r}"
            )
        binary = _validate_binary_descriptor(
            leg.get("binary_identity"),
            location=f"{location}.binary_identity",
            expected_label=expected_variant,
        )
        binary_hashes[label] = binary["binary_sha256"]
        binary_paths[label] = binary["path"]
        binary_modes[label] = binary["mode_bits"]
        report = leg.get("harness_report")
        if report is None:
            report = leg.get("report")
        if workload == DOCX_SEMANTIC_ID:
            configuration, results = _docx_harness_results(
                report,
                f"{location}.harness_report",
                cases=selected_docx_cases,
            )
            corpus_value: Any = [
                {"case": result["case"], "corpus": result["corpus"]}
                for result in results
            ]
            result_identities: list[dict[str, Any]] = []
        elif workload == XLSX_XML_BORROWED_ID:
            configuration, results, result_identities = _xlsx_xml_borrowed_harness_results(
                report,
                f"{location}.harness_report",
                cases=selected_xlsx_cases,
            )
            corpus_value = [
                {"case": result["case"], "corpus": result["corpus"]}
                for result in results
            ]
        else:
            configuration, result = _harness_result(
                report, f"{location}.harness_report"
            )
            corpus_value = result["corpus"]
            result_identities = []
        configurations.append(configuration)
        corpora.append(corpus_value)
        result_identity_sets.append(result_identities)
        report_location = f"{location}.harness_report"
        revision = _require_revision(report, report_location)
        revisions[label] = revision
        environment_identity = _harness_environment_identity(report, report_location)
        environments.append(environment_identity)
        tool = _harness_tool_identity(report, report_location)
        tools.append(tool)
        harness_identity = {
            "leg": label,
            "variant": expected_variant,
            "schema_version": report["schema_version"],
            "tool": tool,
            "environment": environment_identity,
            "git_revision": revision,
            "git_worktree_dirty": False,
            "configuration": configuration,
        }
        if workload == XLSX_XML_BORROWED_ID:
            harness_identity["result_identities"] = result_identities
        harness_identities.append(harness_identity)

    configuration_identity = _canonical_json(configurations[0], "ABBA configuration")
    if any(_canonical_json(item, "ABBA configuration") != configuration_identity for item in configurations[1:]):
        raise ResourceProfileInputError("ABBA harness configurations do not match")
    corpus_identity_json = _canonical_json(corpora[0], "ABBA corpus")
    if any(
        _canonical_json(item, "ABBA corpus") != corpus_identity_json
        for item in corpora[1:]
    ):
        if workload == XLSX_XML_BORROWED_ID:
            raise ResourceProfileInputError(
                "ABBA XLSX borrowed-parser corpus identities do not match"
            )
        label = "DOCX semantic corpora" if workload == DOCX_SEMANTIC_ID else "XLSX corpora"
        raise ResourceProfileInputError(f"ABBA {label} do not match")
    if workload == XLSX_XML_BORROWED_ID:
        result_identity_json = _canonical_json(
            result_identity_sets[0], "ABBA XLSX result identities"
        )
        if any(
            _canonical_json(item, "ABBA XLSX result identities") != result_identity_json
            for item in result_identity_sets[1:]
        ):
            raise ResourceProfileInputError(
                "ABBA XLSX borrowed-parser source/sink/output identities do not match"
            )
    tool_identity = _canonical_json(tools[0], "ABBA tool identity")
    if any(_canonical_json(item, "ABBA tool identity") != tool_identity for item in tools[1:]):
        raise ResourceProfileInputError("ABBA harness tool identities do not match")
    environment_identity = _canonical_json(environments[0], "ABBA environment identity")
    if any(
        _canonical_json(item, "ABBA environment identity") != environment_identity
        for item in environments[1:]
    ):
        raise ResourceProfileInputError(
            "ABBA harness stable environment identities do not match"
        )
    if binary_hashes["A1"] != binary_hashes["A2"]:
        raise ResourceProfileInputError("control binary changed between A1 and A2")
    if binary_hashes["B1"] != binary_hashes["B2"]:
        raise ResourceProfileInputError("candidate binary changed between B1 and B2")
    if binary_paths["A1"] != binary_paths["A2"]:
        raise ResourceProfileInputError("control binary path changed between A1 and A2")
    if binary_paths["B1"] != binary_paths["B2"]:
        raise ResourceProfileInputError("candidate binary path changed between B1 and B2")
    if binary_modes["A1"] != binary_modes["A2"]:
        raise ResourceProfileInputError("control binary mode changed between A1 and A2")
    if binary_modes["B1"] != binary_modes["B2"]:
        raise ResourceProfileInputError("candidate binary mode changed between B1 and B2")
    if binary_hashes["A1"] == binary_hashes["B1"]:
        raise ResourceProfileInputError("control and candidate binary hashes are identical")
    if revisions["A1"] != revisions["A2"]:
        raise ResourceProfileInputError("control revision changed between A1 and A2")
    if revisions["B1"] != revisions["B2"]:
        raise ResourceProfileInputError("candidate revision changed between B1 and B2")
    if revisions["A1"] == revisions["B1"]:
        raise ResourceProfileInputError("control and candidate revisions are identical")
    if expected_configuration is not None:
        for key, expected in expected_configuration.items():
            if configurations[0].get(key) != expected:
                raise ResourceProfileInputError(
                    f"ABBA configuration.{key} does not match fixed configuration: "
                    f"{configurations[0].get(key)!r} != {expected!r}"
                )
    docx_corpus_identities = (
        [
            {
                "case": item["case"],
                "corpus": item["corpus"],
                "identity": corpus_identity(item["corpus"]),
                "identity_sha256": sha256_bytes(
                    _canonical_json(item["corpus"], "DOCX corpus").encode("utf-8")
                ),
            }
            for item in corpora[0]
        ]
        if workload == DOCX_SEMANTIC_ID
        else None
    )
    xlsx_corpus_identities = (
        [
            {
                "case": item["case"],
                "corpus": item["corpus"],
                "identity": corpus_identity(item["corpus"]),
                "identity_sha256": sha256_bytes(
                    _canonical_json(item["corpus"], "XLSX corpus").encode("utf-8")
                ),
            }
            for item in corpora[0]
        ]
        if workload == XLSX_XML_BORROWED_ID
        else None
    )
    validation = {
        "status": "validated",
        "workload": workload,
        "leg_order": list(ABBA_LEG_ORDER),
        "control_revision": revisions["A1"],
        "candidate_revision": revisions["B1"],
        "control_binary_sha256": binary_hashes["A1"],
        "candidate_binary_sha256": binary_hashes["B1"],
        "configuration": configurations[0],
        "corpus": corpora[0],
        "corpus_identities": (
            docx_corpus_identities
            if workload == DOCX_SEMANTIC_ID
            else xlsx_corpus_identities
            if workload == XLSX_XML_BORROWED_ID
            else corpus_identity(corpora[0])
        ),
        "tool": tools[0],
        "environment": environments[0],
        "harness_identities": harness_identities,
        "claim": "identity validation only; no performance or speedup claim",
    }
    if workload == XLSX_XML_BORROWED_ID:
        validation["result_identities"] = result_identity_sets[0]
    return validation


def validate_docx_abba_inputs(
    legs: Sequence[dict[str, Any]],
    *,
    expected_configuration: dict[str, Any] | None = None,
    cases: Sequence[str] = DOCX_SEMANTIC_CASES,
) -> dict[str, Any]:
    """DOCX-named wrapper for callers that do not need workload dispatch."""
    return validate_abba_inputs(
        legs,
        expected_configuration=expected_configuration,
        workload=DOCX_SEMANTIC_ID,
        docx_cases=cases,
    )


def validate_xlsx_xml_borrowed_abba_inputs(
    legs: Sequence[dict[str, Any]],
    *,
    expected_configuration: dict[str, Any] | None = None,
    cases: Sequence[str] = XLSX_XML_BORROWED_CASES,
) -> dict[str, Any]:
    """Validate the strict four-case XLSX borrowed-parser resource workload."""
    selected = _normalize_xlsx_xml_borrowed_cases(cases)
    return validate_abba_inputs(
        legs,
        expected_configuration=expected_configuration,
        workload=XLSX_XML_BORROWED_ID,
        xlsx_cases=selected,
    )


def _finite_resource_value(value: Any) -> float | int | None:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        return None
    try:
        number = float(value)
    except (OverflowError, ValueError):
        return None
    if not math.isfinite(number) or number < 0:
        return None
    return value


def _leg_resource_metric(leg: dict[str, Any], metric: str) -> float | int | None:
    direct = leg.get("resource_metrics")
    if isinstance(direct, dict) and metric in direct:
        return _finite_resource_value(direct[metric])
    if metric.startswith("harness.") and ".elapsed_ns." in metric:
        report = leg.get("harness_report") or leg.get("report")
        if isinstance(report, dict):
            results = report.get("results")
            if isinstance(results, list):
                fields = metric.split(".")
                statistic = fields[-1]
                case = ".".join(fields[1:-2]) if len(fields) >= 4 else None
                case = case or None
                for result in results:
                    if not isinstance(result, dict):
                        continue
                    if case is not None and result.get("case") != case:
                        continue
                    elapsed = result.get("elapsed_ns")
                    if isinstance(elapsed, dict):
                        return _finite_resource_value(elapsed.get(statistic))
        return None
    if metric.startswith("time."):
        timed = leg.get("time")
        parsed = timed.get("parsed") if isinstance(timed, dict) else None
        if isinstance(parsed, dict):
            return _finite_resource_value(parsed.get(metric.split(".", 1)[1]))
        return None
    if metric.startswith("heaptrack."):
        heaptrack = leg.get("heaptrack")
        parsed: Any = None
        if isinstance(heaptrack, dict):
            printed = heaptrack.get("print")
            if isinstance(printed, dict):
                parsed = printed.get("parsed")
            if parsed is None:
                parsed = heaptrack.get("parsed")
        if isinstance(parsed, dict):
            return _finite_resource_value(parsed.get(metric.split(".", 1)[1]))
    return None


def _value_summary(values: Sequence[float | int | None]) -> dict[str, Any]:
    observed = [float(value) for value in values if _finite_resource_value(value) is not None]
    if not observed:
        return {
            "status": "not_measured",
            "count": 0,
            "mean": None,
            "median": None,
            "minimum": None,
            "maximum": None,
        }
    try:
        mean = statistics.fmean(observed)
    except (OverflowError, ValueError):
        mean = None
    try:
        median = statistics.median(observed)
    except (OverflowError, ValueError):
        median = None
    if mean is not None and not math.isfinite(mean):
        mean = None
    if median is not None and not math.isfinite(median):
        median = None
    overflowed = mean is None or median is None
    return {
        "status": "observed_with_overflow" if overflowed else "observed",
        "count": len(observed),
        "mean": mean,
        "median": median,
        "minimum": min(observed),
        "maximum": max(observed),
        "overflow_fields": [
            field
            for field, value in (("mean", mean), ("median", median))
            if value is None
        ],
    }


def _paired_resource_statistic(
    control: float | int | None,
    candidate: float | int | None,
    *,
    execution_order: str,
) -> dict[str, Any]:
    control_value = _finite_resource_value(control)
    candidate_value = _finite_resource_value(candidate)
    result: dict[str, Any] = {
        "execution_order": execution_order,
        "control": control_value,
        "candidate": candidate_value,
        "relative_delta_percent": None,
        "ratio_candidate_to_control": None,
        "status": "not_measured",
    }
    if control_value is None or candidate_value is None:
        return result
    control_number = float(control_value)
    candidate_number = float(candidate_value)
    if control_number == 0:
        result["status"] = "undefined_zero_control" if candidate_number != 0 else "observed_equal_zero"
        return result
    try:
        relative_delta = (candidate_number - control_number) / control_number * 100.0
    except (OverflowError, ZeroDivisionError):
        relative_delta = None
    try:
        ratio = candidate_number / control_number
    except (OverflowError, ZeroDivisionError):
        ratio = None
    if relative_delta is not None and not math.isfinite(relative_delta):
        relative_delta = None
    if ratio is not None and not math.isfinite(ratio):
        ratio = None
    result.update(
        {
            "relative_delta_percent": relative_delta,
            "ratio_candidate_to_control": ratio,
            "status": "observed" if relative_delta is not None and ratio is not None else "overflow",
        }
    )
    return result


def abba_statistics(
    legs: Sequence[dict[str, Any]],
    *,
    metric_specs: Sequence[tuple[str, str]] = RESOURCE_METRIC_SPECS,
) -> dict[str, Any]:
    """Emit descriptive paired resource statistics without accepting a speedup."""
    if any(not isinstance(leg, dict) for leg in legs):
        raise ResourceProfileInputError("ABBA legs must be objects")
    labels = [leg.get("leg") for leg in legs]
    validate_abba_order(labels)
    metrics: dict[str, Any] = {}
    for metric, description in metric_specs:
        values = {label: _leg_resource_metric(leg, metric) for label, leg in zip(labels, legs)}
        controls = [values["A1"], values["A2"]]
        candidates = [values["B1"], values["B2"]]
        metrics[metric] = {
            "description": description,
            "values_by_leg": values,
            "control": _value_summary(controls),
            "candidate": _value_summary(candidates),
            "paired": {
                "A1_control_to_B1_candidate": _paired_resource_statistic(
                    values["A1"], values["B1"], execution_order="A1, B1"
                ),
                "A2_control_to_B2_candidate": _paired_resource_statistic(
                    values["A2"], values["B2"], execution_order="B2, A2"
                ),
            },
            "claim": "descriptive paired process statistics only; no automatic speedup claim",
        }
    return {
        "status": "observed",
        "leg_order": list(ABBA_LEG_ORDER),
        "metrics": metrics,
        "not_measured": dict(NOT_MEASURED_RESOURCE_DIMENSIONS),
        "interpretation": (
            "Values are descriptive summaries of four fresh process legs. A relative delta is "
            "not an accepted optimization or speedup claim."
        ),
    }


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


def _retain_run_artifacts(
    run: dict[str, Any], stdout_path: Path, stderr_path: Path
) -> dict[str, Any]:
    retained = dict(run)
    retained["stdout"] = artifact(stdout_path, retained=True)
    retained["stderr"] = artifact(stderr_path, retained=True)
    return retained


def _profile_abba_heaptrack(
    binary: Path,
    args: Sequence[str],
    workdir: Path,
    tools: dict[str, dict[str, Any]],
    *,
    warmup: int,
    samples: int,
    timeout_seconds: float,
    expected_harness_identity: dict[str, Any] | None = None,
    harness_identity_validator: Callable[[Any, str], dict[str, Any]] | None = None,
) -> dict[str, Any]:
    """Collect an optional retained heaptrack process-total profile for one leg."""
    heaptrack_tool = tools.get("heaptrack", {})
    print_tool = tools.get("heaptrack_print", {})
    if not heaptrack_tool.get("available"):
        return {
            "status": "unsupported",
            "reason": "heaptrack unavailable",
            "scope": "whole process when available; profiler overhead is included",
        }
    profile_json = workdir / "heaptrack-harness.json"
    profile_args = _profile_command(
        binary,
        args,
        warmup=warmup,
        samples=samples,
        output=profile_json,
    )
    prefix = workdir / "heaptrack-profile"
    stdout = workdir / "heaptrack-stdout.txt"
    stderr = workdir / "heaptrack-stderr.txt"
    command = [str(heaptrack_tool["path"]), "--record-only", "-o", str(prefix), *profile_args]
    run = run_command(command, stdout_path=stdout, stderr_path=stderr, timeout_seconds=timeout_seconds)
    run = _retain_run_artifacts(run, stdout, stderr)
    captures = sorted(
        path
        for path in workdir.glob(prefix.name + "*")
        if path.is_file() and path not in (stdout, stderr)
    )
    result: dict[str, Any] = {
        "status": "ok" if captures and run["returncode"] == 0 else "failed",
        "scope": "whole process; heaptrack instrumentation overhead is included",
        "command": command_record(command),
        "run": run,
        "harness": artifact(profile_json, retained=True),
        "captures": [artifact(path, retained=True) for path in captures],
    }
    if expected_harness_identity is not None:
        if harness_identity_validator is None:
            raise ResourceProfileInputError(
                "heaptrack harness identity validator is required when timed identity is supplied"
            )
        if profile_json.is_file():
            try:
                heaptrack_report = load_json(profile_json)
                observed_identity = harness_identity_validator(
                    heaptrack_report, "heaptrack.harness_report"
                )
            except (OSError, ValueError) as error:
                raise ResourceProfileInputError(
                    f"heaptrack harness report is invalid: {error}"
                ) from error
            expected_json = _canonical_json(
                expected_harness_identity, "timed harness identity"
            )
            observed_json = _canonical_json(
                observed_identity, "heaptrack harness identity"
            )
            expected_sha256 = sha256_bytes(expected_json.encode("utf-8"))
            observed_sha256 = sha256_bytes(observed_json.encode("utf-8"))
            if observed_json != expected_json:
                raise ResourceProfileInputError(
                    "heaptrack harness identity does not match timed leg identity: "
                    f"timed_sha256={expected_sha256}, heaptrack_sha256={observed_sha256}"
                )
            result["harness_identity"] = {
                "status": "validated",
                "sha256": observed_sha256,
                "scope": (
                    "exact configuration, corpus, source/sink/output, revision, "
                    "tool, and stable environment identity"
                ),
            }
        else:
            raise ResourceProfileInputError(
                "heaptrack harness report was not produced; cannot validate timed leg identity"
            )
    if not captures:
        result["capture"] = artifact(workdir / "heaptrack-profile", retained=True)
        if run["returncode"] == 0:
            result["status"] = "unparsed"
        return result
    capture = captures[-1]
    result["capture"] = artifact(capture, retained=True)
    if not print_tool.get("available"):
        result["print"] = {
            "status": "unsupported",
            "reason": "heaptrack_print unavailable; raw capture retained",
        }
        if run["returncode"] == 0:
            result["status"] = "unsupported"
            result["failure_stage"] = "heaptrack_print_unavailable"
        else:
            result["status"] = "failed"
            result["failure_stage"] = "heaptrack"
        return result
    printed = workdir / "heaptrack-print.txt"
    print_stderr = workdir / "heaptrack-print-stderr.txt"
    histogram = workdir / "heaptrack-histogram.tsv"
    print_command = [str(print_tool["path"]), "-f", str(capture), "-H", str(histogram)]
    print_run = run_command(
        print_command,
        stdout_path=printed,
        stderr_path=print_stderr,
        timeout_seconds=timeout_seconds,
    )
    print_run = _retain_run_artifacts(print_run, printed, print_stderr)
    parsed = parse_heaptrack_print(printed, retained=True)
    parsed["allocated_bytes"] = parse_heaptrack_histogram(histogram)
    parsed["histogram_artifact"] = artifact(histogram, retained=True)
    parsed["scope"] = "whole process; heaptrack instrumentation overhead is included"
    print_status = "ok"
    if print_run["returncode"] != 0 or parsed.get("status") != "ok":
        print_status = "failed"
        result["status"] = "failed"
        result["failure_stage"] = "heaptrack_print"
    result["print"] = {
        "status": print_status,
        "command": command_record(print_command),
        "run": print_run,
        "parsed": parsed,
        "artifact": artifact(printed, retained=True),
    }
    return result


def profile_xlsx_abba_leg(
    *,
    leg: str,
    variant: str,
    binary: Path,
    binary_descriptor: dict[str, Any],
    artifact_root: Path,
    warmup: int,
    samples: int,
    tools: dict[str, dict[str, Any]],
    timeout_seconds: float,
) -> dict[str, Any]:
    """Run one fixed-configuration XLSX managed-batch process and retain artifacts."""
    if leg not in ABBA_LEG_VARIANTS or ABBA_LEG_VARIANTS[leg] != variant:
        raise ResourceProfileInputError(f"invalid ABBA leg/variant pair: {leg!r}/{variant!r}")
    leg_dir = artifact_root / leg.lower()
    leg_dir.mkdir(parents=True, exist_ok=True)
    harness_json = leg_dir / "harness.json"
    stdout = leg_dir / "harness-stdout.txt"
    stderr = leg_dir / "harness-stderr.txt"
    command = _profile_command(
        binary,
        XLSX_MANAGED_BATCH_ARGS,
        warmup=warmup,
        samples=samples,
        output=harness_json,
    )
    time_tool = tools.get("time", {})
    if time_tool.get("available"):
        time_report = leg_dir / "time-v.txt"
        timed_command = [str(time_tool["path"]), "-v", "-o", str(time_report), "--", *command]
    else:
        time_report = leg_dir / "time-v.txt"
        timed_command = command
    run = run_command(timed_command, stdout_path=stdout, stderr_path=stderr, timeout_seconds=timeout_seconds)
    run = _retain_run_artifacts(run, stdout, stderr)
    if run["returncode"] != 0 or not harness_json.is_file():
        raise RuntimeError(
            f"harness failed for {leg}: returncode={run['returncode']} stderr={run['stderr_excerpt']}"
        )
    report = load_json(harness_json)
    time_result: dict[str, Any]
    if time_tool.get("available"):
        parsed_time = parse_time_report(time_report, retained=True)
        time_result = {
            "status": parsed_time.get("status", "unparsed"),
            "command": command_record(timed_command),
            "run": run,
            "parsed": parsed_time,
        }
    else:
        time_result = {
            "status": "unsupported",
            "reason": "GNU /usr/bin/time unavailable",
            "scope": "whole process when available",
        }
    return {
        "leg": leg,
        "variant": variant,
        "binary_identity": dict(binary_descriptor),
        "harness": {
            "command": command_record(command),
            "run": run,
            "report": artifact(harness_json, retained=True),
            "logical_measurements": logical_measurements(report),
        },
        # This private field is removed before JSON publication.  Keeping the
        # parsed report until validation prevents a compacted summary from
        # accidentally hiding revision/corpus/config mismatches.
        "harness_report": report,
        "time": time_result,
        "heaptrack": _profile_abba_heaptrack(
            binary,
            XLSX_MANAGED_BATCH_ARGS,
            leg_dir,
            tools,
            warmup=warmup,
            samples=samples,
            timeout_seconds=timeout_seconds,
        ),
        "artifact_directory": str(leg_dir),
    }


def profile_docx_semantic_abba_leg(
    *,
    leg: str,
    variant: str,
    binary: Path,
    binary_descriptor: dict[str, Any],
    artifact_root: Path,
    warmup: int,
    samples: int,
    tools: dict[str, dict[str, Any]],
    timeout_seconds: float,
    cases: Sequence[str] = DOCX_SEMANTIC_CASES,
) -> dict[str, Any]:
    """Run one fixed DOCX semantic resource leg.

    The selected DOCX cases share one fresh process so their corpus and tool
    identities are aligned.  The process is run once under GNU time and once
    under heaptrack when those optional tools are available; neither run is a
    latency sample.
    """
    selected_cases = _normalize_docx_cases(cases)
    if leg not in ABBA_LEG_VARIANTS or ABBA_LEG_VARIANTS[leg] != variant:
        raise ResourceProfileInputError(f"invalid ABBA leg/variant pair: {leg!r}/{variant!r}")
    leg_dir = artifact_root / leg.lower()
    leg_dir.mkdir(parents=True, exist_ok=True)
    harness_json = leg_dir / "harness.json"
    stdout = leg_dir / "harness-stdout.txt"
    stderr = leg_dir / "harness-stderr.txt"
    command = _profile_command(
        binary,
        docx_semantic_args(selected_cases),
        warmup=warmup,
        samples=samples,
        output=harness_json,
    )
    time_tool = tools.get("time", {})
    time_report = leg_dir / "time-v.txt"
    timed_command = (
        [str(time_tool["path"]), "-v", "-o", str(time_report), "--", *command]
        if time_tool.get("available")
        else command
    )
    run = run_command(
        timed_command,
        stdout_path=stdout,
        stderr_path=stderr,
        timeout_seconds=timeout_seconds,
    )
    run = _retain_run_artifacts(run, stdout, stderr)
    if run["returncode"] != 0 or not harness_json.is_file():
        raise RuntimeError(
            f"harness failed for {leg}: returncode={run['returncode']} stderr={run['stderr_excerpt']}"
        )
    report = load_json(harness_json)
    if time_tool.get("available"):
        parsed_time = parse_time_report(time_report, retained=True)
        time_result: dict[str, Any] = {
            "status": parsed_time.get("status", "unparsed"),
            "command": command_record(timed_command),
            "run": run,
            "parsed": parsed_time,
            "scope": "whole process; instrumented resource observation, not latency evidence",
        }
    else:
        time_result = {
            "status": "unsupported",
            "reason": "GNU /usr/bin/time unavailable",
            "scope": "whole process when available; not latency evidence",
        }
    return {
        "leg": leg,
        "variant": variant,
        "binary_identity": dict(binary_descriptor),
        "harness": {
            "command": command_record(command),
            "run": run,
            "report": artifact(harness_json, retained=True),
            "logical_measurements": logical_measurements(report),
            "instrumented_resource_metrics": instrumented_harness_metrics(
                report, selected_cases
            ),
        },
        # This private field is removed before JSON publication.  It remains
        # available through validation so a compact summary cannot hide a
        # changed case order, corpus, configuration, or revision.
        "harness_report": report,
        "resource_metrics": instrumented_harness_metrics(
            report, selected_cases
        ),
        "latency_evidence": dict(LATENCY_SEPARATION),
        "time": time_result,
        "heaptrack": _profile_abba_heaptrack(
            binary,
            docx_semantic_args(selected_cases),
            leg_dir,
            tools,
            warmup=warmup,
            samples=samples,
            timeout_seconds=timeout_seconds,
        ),
        "artifact_directory": str(leg_dir),
    }


def profile_xlsx_xml_borrowed_abba_leg(
    *,
    leg: str,
    variant: str,
    binary: Path,
    binary_descriptor: dict[str, Any],
    artifact_root: Path,
    warmup: int,
    samples: int,
    tools: dict[str, dict[str, Any]],
    timeout_seconds: float,
    cases: Sequence[str] = XLSX_XML_BORROWED_CASES,
) -> dict[str, Any]:
    """Run one retained four-case XLSX borrowed-parser resource leg."""
    selected_cases = _normalize_xlsx_xml_borrowed_cases(cases)
    if leg not in ABBA_LEG_VARIANTS or ABBA_LEG_VARIANTS[leg] != variant:
        raise ResourceProfileInputError(f"invalid ABBA leg/variant pair: {leg!r}/{variant!r}")
    leg_dir = artifact_root / leg.lower()
    leg_dir.mkdir(parents=True, exist_ok=True)
    harness_json = leg_dir / "harness.json"
    stdout = leg_dir / "harness-stdout.txt"
    stderr = leg_dir / "harness-stderr.txt"
    harness_args = xlsx_xml_borrowed_args(selected_cases)
    command = _profile_command(
        binary,
        harness_args,
        warmup=warmup,
        samples=samples,
        output=harness_json,
    )
    time_tool = tools.get("time", {})
    time_report = leg_dir / "time-v.txt"
    timed_command = (
        [str(time_tool["path"]), "-v", "-o", str(time_report), "--", *command]
        if time_tool.get("available")
        else command
    )
    run = run_command(
        timed_command,
        stdout_path=stdout,
        stderr_path=stderr,
        timeout_seconds=timeout_seconds,
    )
    run = _retain_run_artifacts(run, stdout, stderr)
    if run["returncode"] != 0 or not harness_json.is_file():
        raise RuntimeError(
            f"harness failed for {leg}: returncode={run['returncode']} stderr={run['stderr_excerpt']}"
        )
    report = load_json(harness_json)
    harness_identity = _xlsx_xml_borrowed_harness_identity(
        report, "timed.harness_report", cases=selected_cases
    )
    resource_metrics = instrumented_harness_metrics(report, selected_cases)
    if time_tool.get("available"):
        parsed_time = parse_time_report(time_report, retained=True)
        time_result: dict[str, Any] = {
            "status": parsed_time.get("status", "unparsed"),
            "command": command_record(timed_command),
            "run": run,
            "parsed": parsed_time,
            "scope": "whole process; instrumented resource observation, not latency evidence",
        }
    else:
        time_result = {
            "status": "unsupported",
            "reason": "GNU /usr/bin/time unavailable",
            "scope": "whole process when available; not latency evidence",
        }
    return {
        "leg": leg,
        "variant": variant,
        "binary_identity": dict(binary_descriptor),
        "harness": {
            "command": command_record(command),
            "run": run,
            "report": artifact(harness_json, retained=True),
            "logical_measurements": logical_measurements(report),
            "instrumented_resource_metrics": resource_metrics,
        },
        # Keep the complete report private until strict identity validation has
        # consumed every result row and its source/sink/output channels.
        "harness_report": report,
        "resource_metrics": resource_metrics,
        "latency_evidence": dict(LATENCY_SEPARATION),
        "time": time_result,
        "heaptrack": _profile_abba_heaptrack(
            binary,
            harness_args,
            leg_dir,
            tools,
            warmup=warmup,
            samples=samples,
            timeout_seconds=timeout_seconds,
            expected_harness_identity=harness_identity,
            harness_identity_validator=lambda value, location: (
                _xlsx_xml_borrowed_harness_identity(
                    value, location, cases=selected_cases
                )
            ),
        ),
        "artifact_directory": str(leg_dir),
    }


def _fixed_xlsx_abba_configuration(warmup: int, samples: int) -> dict[str, Any]:
    return {
        "warmup_iterations": warmup,
        "samples": samples,
        "case": XLSX_MANAGED_BATCH_CASE,
        "xlsx_cell_crud_shape": XLSX_MANAGED_BATCH_SHAPE,
        "harness_expected": {
            "samples_per_case": samples,
            "warmup_iterations_per_case": warmup,
            "cases": [XLSX_MANAGED_BATCH_CASE],
            "xlsx_cell_crud_shapes": [XLSX_MANAGED_BATCH_SHAPE],
        },
        "leg_order": list(ABBA_LEG_ORDER),
    }


def _fixed_xlsx_xml_borrowed_abba_configuration(
    warmup: int,
    samples: int,
    cases: Sequence[str] = XLSX_XML_BORROWED_CASES,
) -> dict[str, Any]:
    selected_cases = _normalize_xlsx_xml_borrowed_cases(cases)
    return {
        "warmup_iterations": warmup,
        "samples": samples,
        "cases": list(selected_cases),
        "xlsx_shape": XLSX_XML_BORROWED_SHAPE,
        "xlsx_cell_crud_shape": XLSX_XML_BORROWED_CELL_CRUD_SHAPE,
        "harness_expected": {
            "samples_per_case": samples,
            "warmup_iterations_per_case": warmup,
            "cases": list(selected_cases),
            "xlsx_shapes": [XLSX_XML_BORROWED_SHAPE],
            "xlsx_cell_crud_shapes": [XLSX_XML_BORROWED_CELL_CRUD_SHAPE],
        },
        "leg_order": list(ABBA_LEG_ORDER),
        "latency_evidence": dict(LATENCY_SEPARATION),
        "resource_tools": ["/usr/bin/time", "heaptrack", "heaptrack_print"],
    }


def _fixed_docx_abba_configuration(
    warmup: int,
    samples: int,
    cases: Sequence[str] = DOCX_SEMANTIC_CASES,
) -> dict[str, Any]:
    selected_cases = _normalize_docx_cases(cases)
    return {
        "warmup_iterations": warmup,
        "samples": samples,
        "cases": list(selected_cases),
        "include_one_paragraph_text": (
            DOCX_SEMANTIC_ONE_PARAGRAPH_TEXT_CASE in selected_cases
        ),
        "semantic_shape": DOCX_SEMANTIC_SHAPE,
        "harness_expected": {
            "samples_per_case": samples,
            "warmup_iterations_per_case": warmup,
            "cases": list(selected_cases),
            "semantic_shapes": [DOCX_SEMANTIC_SHAPE],
        },
        "leg_order": list(ABBA_LEG_ORDER),
        "latency_evidence": dict(LATENCY_SEPARATION),
        "resource_tools": ["/usr/bin/time", "heaptrack", "heaptrack_print"],
    }


def run_xlsx_managed_batch_abba(arguments: argparse.Namespace) -> int:
    control = binary_identity(Path(arguments.control_binary), label="control")
    candidate = binary_identity(Path(arguments.candidate_binary), label="candidate")
    if control["binary_sha256"] == candidate["binary_sha256"]:
        raise ResourceProfileInputError("control and candidate binary hashes are identical")
    output_path = Path(arguments.output).expanduser().resolve()
    artifact_root = (
        Path(arguments.artifact_dir).expanduser().resolve()
        if arguments.artifact_dir
        else output_path.with_name(output_path.stem + "-artifacts")
    )
    output_path, artifact_root = reserve_abba_paths(output_path, artifact_root)
    tools = {
        "time": probe_tool("/usr/bin/time", ("--version",)),
        "heaptrack": probe_tool("heaptrack", ("--version",)),
        "heaptrack_print": probe_tool("heaptrack_print", ("--version",)),
    }
    legs: list[dict[str, Any]] = []
    for leg in ABBA_LEG_ORDER:
        variant = ABBA_LEG_VARIANTS[leg]
        descriptor = control if variant == "control" else candidate
        legs.append(
            profile_xlsx_abba_leg(
                leg=leg,
                variant=variant,
                binary=Path(descriptor["path"]),
                binary_descriptor=descriptor,
                artifact_root=artifact_root,
                warmup=arguments.warmup,
                samples=arguments.samples,
                tools=tools,
                timeout_seconds=arguments.timeout,
            )
        )
    control_after = binary_identity(Path(arguments.control_binary), label="control")
    candidate_after = binary_identity(Path(arguments.candidate_binary), label="candidate")
    if control_after["binary_sha256"] != control["binary_sha256"]:
        raise ResourceProfileInputError("control binary changed during ABBA execution")
    if control_after["mode_bits"] != control["mode_bits"]:
        raise ResourceProfileInputError("control binary mode changed during ABBA execution")
    if candidate_after["binary_sha256"] != candidate["binary_sha256"]:
        raise ResourceProfileInputError("candidate binary changed during ABBA execution")
    if candidate_after["mode_bits"] != candidate["mode_bits"]:
        raise ResourceProfileInputError("candidate binary mode changed during ABBA execution")
    fixed_configuration = _fixed_xlsx_abba_configuration(arguments.warmup, arguments.samples)
    validation = validate_abba_inputs(
        legs,
        expected_configuration=fixed_configuration["harness_expected"],
    )
    statistics_report = abba_statistics(legs)
    published_legs = []
    for leg, harness_identity in zip(legs, validation["harness_identities"]):
        published = {
            key: value for key, value in leg.items() if key != "harness_report"
        }
        published["harness_identity"] = harness_identity
        published_legs.append(published)
    canonical_harness_identity = {
        "schema_version": SCHEMA_VERSION,
        "tool": validation["tool"],
        "environment": validation["environment"],
        "configuration": validation["configuration"],
        "leg_revisions": {
            identity["leg"]: identity["git_revision"]
            for identity in validation["harness_identities"]
        },
        "clean_worktrees": True,
    }
    report = {
        "schema_version": SCHEMA_VERSION,
        "abba_schema_version": ABBA_SCHEMA_VERSION,
        "tool": {
            "name": TOOL_NAME,
            "version": TOOL_VERSION,
            "mode": "xlsx-managed-batch-abba-resource-profile",
            "python_standard_library_only": True,
        },
        "scope": {
            "claim": (
                "four fresh process legs in fixed A1/B1/B2/A2 order; descriptive resource "
                "comparison only and no automatic speedup claim"
            ),
            "excluded_formats": ["iWork"],
            "resource_scope": (
                "harness elapsed summaries plus process-total /usr/bin/time and heaptrack "
                "observations when available"
            ),
        },
        "host_environment": environment(),
        "binary_identity": {"control": control, "candidate": candidate},
        "configuration": fixed_configuration,
        "canonical_harness_identity": canonical_harness_identity,
        "tools": tools,
        "validation": validation,
        "legs": published_legs,
        "statistics": statistics_report,
        "perf_counters": {
            "status": "not_measured",
            "reason": "perf counters are not collected or synthesized by this mode",
            "counters": {event: {"value": None, "available": False} for event in EVENTS},
        },
        "not_measured": dict(NOT_MEASURED_RESOURCE_DIMENSIONS),
        "artifact_directory": str(artifact_root),
        "limitations": [
            "Control and candidate revisions must come from clean harness worktrees.",
            "Missing optional tools remain unsupported/null; no zero is substituted.",
            "Allocation and RSS values are whole-process observations, including profiler overhead where applicable.",
            "Copy bytes, decompressed bytes, and physical-cold I/O are not measured.",
        ],
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("x", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, sort_keys=True, allow_nan=False)
        handle.write("\n")
    return 0


def run_xlsx_xml_borrowed_abba(arguments: argparse.Namespace) -> int:
    """Run strict retained A1/B1/B2/A2 XLSX borrowed-parser resource evidence."""
    selected_cases = _normalize_xlsx_xml_borrowed_cases(
        getattr(arguments, "cases", XLSX_XML_BORROWED_CASES)
    )
    control = binary_identity(Path(arguments.control_binary), label="control")
    candidate = binary_identity(Path(arguments.candidate_binary), label="candidate")
    if control["binary_sha256"] == candidate["binary_sha256"]:
        raise ResourceProfileInputError("control and candidate binary hashes are identical")
    output_path = Path(arguments.output).expanduser().resolve()
    artifact_root = (
        Path(arguments.artifact_dir).expanduser().resolve()
        if arguments.artifact_dir
        else output_path.with_name(output_path.stem + "-artifacts")
    )
    output_path, artifact_root = reserve_abba_paths(output_path, artifact_root)
    tools = {
        "time": probe_tool("/usr/bin/time", ("--version",)),
        "heaptrack": probe_tool("heaptrack", ("--version",)),
        "heaptrack_print": probe_tool("heaptrack_print", ("--version",)),
    }
    legs: list[dict[str, Any]] = []
    for leg in ABBA_LEG_ORDER:
        variant = ABBA_LEG_VARIANTS[leg]
        descriptor = control if variant == "control" else candidate
        legs.append(
            profile_xlsx_xml_borrowed_abba_leg(
                leg=leg,
                variant=variant,
                binary=Path(descriptor["path"]),
                binary_descriptor=descriptor,
                artifact_root=artifact_root,
                warmup=arguments.warmup,
                samples=arguments.samples,
                tools=tools,
                timeout_seconds=arguments.timeout,
                cases=selected_cases,
            )
        )
    control_after = binary_identity(Path(arguments.control_binary), label="control")
    candidate_after = binary_identity(Path(arguments.candidate_binary), label="candidate")
    if control_after["binary_sha256"] != control["binary_sha256"]:
        raise ResourceProfileInputError(
            "control binary changed during XLSX borrowed-parser ABBA execution"
        )
    if control_after["mode_bits"] != control["mode_bits"]:
        raise ResourceProfileInputError("control binary mode changed during ABBA execution")
    if candidate_after["binary_sha256"] != candidate["binary_sha256"]:
        raise ResourceProfileInputError(
            "candidate binary changed during XLSX borrowed-parser ABBA execution"
        )
    if candidate_after["mode_bits"] != candidate["mode_bits"]:
        raise ResourceProfileInputError("candidate binary mode changed during ABBA execution")

    fixed_configuration = _fixed_xlsx_xml_borrowed_abba_configuration(
        arguments.warmup,
        arguments.samples,
        selected_cases,
    )
    validation = validate_abba_inputs(
        legs,
        expected_configuration=fixed_configuration["harness_expected"],
        workload=XLSX_XML_BORROWED_ID,
    )
    statistics_report = abba_statistics(
        legs,
        metric_specs=XLSX_XML_BORROWED_RESOURCE_METRIC_SPECS,
    )
    published_legs = []
    for leg, harness_identity in zip(legs, validation["harness_identities"]):
        published = {
            key: value for key, value in leg.items() if key != "harness_report"
        }
        published["harness_identity"] = harness_identity
        published_legs.append(published)
    canonical_harness_identity = {
        "schema_version": SCHEMA_VERSION,
        "tool": validation["tool"],
        "environment": validation["environment"],
        "configuration": validation["configuration"],
        "leg_revisions": {
            identity["leg"]: identity["git_revision"]
            for identity in validation["harness_identities"]
        },
        "clean_worktrees": True,
    }
    report = {
        "schema_version": SCHEMA_VERSION,
        "abba_schema_version": ABBA_SCHEMA_VERSION,
        "tool": {
            "name": TOOL_NAME,
            "version": TOOL_VERSION,
            "mode": "xlsx-xml-borrowed-abba-resource-profile",
            "python_standard_library_only": True,
        },
        "scope": {
            "workload": XLSX_XML_BORROWED_ID,
            "cases": list(selected_cases),
            "xlsx_shape": XLSX_XML_BORROWED_SHAPE,
            "xlsx_cell_crud_shape": XLSX_XML_BORROWED_CELL_CRUD_SHAPE,
            "claim": (
                "four fresh process legs in fixed A1/B1/B2/A2 order over exact tiny "
                "and medium XLSX corpora; descriptive allocation/RSS/resource "
                "comparison only and no automatic latency or speedup claim"
            ),
            "resource_scope": (
                "instrumented harness summaries plus whole-process /usr/bin/time and "
                "optional heaptrack allocation/peak-heap/peak-RSS observations"
            ),
            "instrumented_elapsed": "retained for resource-leg alignment only; not latency evidence",
            "physical_io": (
                "no physical I/O claim; no cache flush, source-byte, decompressed-byte, "
                "recompressed-byte, or memory-copy counter is collected"
            ),
        },
        "latency_evidence": dict(LATENCY_SEPARATION),
        "host_environment": environment(),
        "binary_identity": {"control": control, "candidate": candidate},
        "configuration": fixed_configuration,
        "corpus_identities": validation["corpus_identities"],
        "result_identities": validation["result_identities"],
        "canonical_harness_identity": canonical_harness_identity,
        "tools": tools,
        "validation": validation,
        "legs": published_legs,
        "statistics": statistics_report,
        "perf_counters": {
            "status": "not_measured",
            "reason": "perf counters are not collected or synthesized by this mode",
            "counters": {event: {"value": None, "available": False} for event in EVENTS},
        },
        "not_measured": dict(NOT_MEASURED_RESOURCE_DIMENSIONS),
        "physical_io_claim_scope": (
            "whole-process profiler observations do not identify physical source reads, "
            "decompression, recompression, or memory copies"
        ),
        "artifact_directory": str(artifact_root),
        "limitations": [
            "Control and candidate revisions must come from clean release harness worktrees.",
            "The four case names, order, tiny/medium shapes, and identity channels are fixed.",
            "Missing optional tools remain unsupported/null; no zero is substituted.",
            "Instrumented harness elapsed values are not latency evidence.",
            "Allocation and RSS values are whole-process observations, including profiler overhead.",
            "Source, sink, and output identities are compared exactly per case; absent values remain absent.",
            "Physical I/O, copy bytes, decompressed bytes, and recompressed bytes are not measured.",
        ],
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("x", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, sort_keys=True, allow_nan=False)
        handle.write("\n")
    return 0


def run_docx_semantic_abba(arguments: argparse.Namespace) -> int:
    """Run reproducible DOCX semantic resource evidence.

    Control and candidate are intentionally explicit executable paths.  This
    mode never builds either binary and never substitutes a current working
    tree or an implicit Cargo target for a missing path.
    """
    selected_cases = docx_semantic_cases(
        include_one_paragraph_text=getattr(arguments, "include_one_paragraph_text", False)
    )
    control = binary_identity(Path(arguments.control_binary), label="control")
    candidate = binary_identity(Path(arguments.candidate_binary), label="candidate")
    if control["binary_sha256"] == candidate["binary_sha256"]:
        raise ResourceProfileInputError("control and candidate binary hashes are identical")
    output_path = Path(arguments.output).expanduser().resolve()
    artifact_root = (
        Path(arguments.artifact_dir).expanduser().resolve()
        if arguments.artifact_dir
        else output_path.with_name(output_path.stem + "-artifacts")
    )
    output_path, artifact_root = reserve_abba_paths(output_path, artifact_root)
    tools = {
        "time": probe_tool("/usr/bin/time", ("--version",)),
        "heaptrack": probe_tool("heaptrack", ("--version",)),
        "heaptrack_print": probe_tool("heaptrack_print", ("--version",)),
    }
    legs: list[dict[str, Any]] = []
    for leg in ABBA_LEG_ORDER:
        variant = ABBA_LEG_VARIANTS[leg]
        descriptor = control if variant == "control" else candidate
        legs.append(
            profile_docx_semantic_abba_leg(
                leg=leg,
                variant=variant,
                binary=Path(descriptor["path"]),
                binary_descriptor=descriptor,
                artifact_root=artifact_root,
                warmup=arguments.warmup,
                samples=arguments.samples,
                tools=tools,
                timeout_seconds=arguments.timeout,
                cases=selected_cases,
            )
        )
    control_after = binary_identity(Path(arguments.control_binary), label="control")
    candidate_after = binary_identity(Path(arguments.candidate_binary), label="candidate")
    if control_after["binary_sha256"] != control["binary_sha256"]:
        raise ResourceProfileInputError("control binary changed during DOCX ABBA execution")
    if control_after["mode_bits"] != control["mode_bits"]:
        raise ResourceProfileInputError("control binary mode changed during DOCX ABBA execution")
    if candidate_after["binary_sha256"] != candidate["binary_sha256"]:
        raise ResourceProfileInputError("candidate binary changed during DOCX ABBA execution")
    if candidate_after["mode_bits"] != candidate["mode_bits"]:
        raise ResourceProfileInputError("candidate binary mode changed during DOCX ABBA execution")
    fixed_configuration = _fixed_docx_abba_configuration(
        arguments.warmup,
        arguments.samples,
        selected_cases,
    )
    validation = validate_abba_inputs(
        legs,
        expected_configuration=fixed_configuration["harness_expected"],
        workload=DOCX_SEMANTIC_ID,
        docx_cases=selected_cases,
    )
    statistics_report = abba_statistics(
        legs,
        metric_specs=docx_resource_metric_specs(selected_cases),
    )
    published_legs = []
    for leg, harness_identity in zip(legs, validation["harness_identities"]):
        published = {
            key: value for key, value in leg.items() if key != "harness_report"
        }
        # Keep the canonical identity needed to audit a published result even
        # though the raw harness JSON is intentionally not copied into it.
        published["harness_identity"] = harness_identity
        published_legs.append(published)
    canonical_harness_identity = {
        "schema_version": SCHEMA_VERSION,
        "tool": validation["tool"],
        "environment": validation["environment"],
        "configuration": validation["configuration"],
        "leg_revisions": {
            identity["leg"]: identity["git_revision"]
            for identity in validation["harness_identities"]
        },
        "clean_worktrees": True,
    }
    report = {
        "schema_version": SCHEMA_VERSION,
        "abba_schema_version": ABBA_SCHEMA_VERSION,
        "tool": {
            "name": TOOL_NAME,
            "version": TOOL_VERSION,
            "mode": "docx-semantic-abba-resource-profile",
            "python_standard_library_only": True,
        },
        "scope": {
            "claim": (
                "four fresh process legs in fixed A1/B1/B2/A2 order over the same "
                "deterministic DOCX semantic corpus; descriptive resource "
                "comparison only and no automatic latency or speedup claim"
            ),
            "workload": DOCX_SEMANTIC_ID,
            "cases": list(selected_cases),
            "semantic_shape": DOCX_SEMANTIC_SHAPE,
            "resource_scope": (
                "process-total /usr/bin/time and optional heaptrack observations, plus "
                "per-case harness summaries collected inside those instrumented runs"
            ),
            "physical_io": (
                "no physical I/O claim; no cache flush, source-byte, decompressed-byte, "
                "recompressed-byte, or memory-copy counter is collected"
            ),
        },
        "latency_evidence": dict(LATENCY_SEPARATION),
        "host_environment": environment(),
        "binary_identity": {"control": control, "candidate": candidate},
        "configuration": fixed_configuration,
        "corpus_identities": validation["corpus_identities"],
        "canonical_harness_identity": canonical_harness_identity,
        "tools": tools,
        "validation": validation,
        "legs": published_legs,
        "statistics": statistics_report,
        "perf_counters": {
            "status": "not_measured",
            "reason": "perf counters are not collected or synthesized by this mode",
            "counters": {event: {"value": None, "available": False} for event in EVENTS},
        },
        "not_measured": dict(NOT_MEASURED_RESOURCE_DIMENSIONS),
        "physical_io_claim_scope": (
            "whole-process profiler observations do not identify physical source reads, "
            "decompression, recompression, or memory copies"
        ),
        "artifact_directory": str(artifact_root),
        "limitations": [
            "Control and candidate revisions must come from clean release harness worktrees.",
            "Missing optional tools remain unsupported/null; no zero is substituted.",
            "Instrumented harness elapsed values are retained for workload alignment only, not latency evidence.",
            "Allocation and RSS values are whole-process observations, including profiler overhead where applicable.",
            "Physical I/O, copy bytes, decompressed bytes, and recompressed bytes are not measured.",
        ],
    }
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("x", encoding="utf-8") as handle:
        json.dump(report, handle, indent=2, sort_keys=True, allow_nan=False)
        handle.write("\n")
    return 0


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


def _apply_run_sampling_defaults(arguments: argparse.Namespace) -> None:
    """Apply mode-specific run defaults while keeping ordinary run lightweight."""
    requested = getattr(arguments, "workload", None)
    only = getattr(arguments, "only", None)
    borrowed = requested in XLSX_XML_BORROWED_ID_ALIASES or (
        requested is None and only in XLSX_XML_BORROWED_ID_ALIASES
    )
    if getattr(arguments, "warmup", None) is None:
        arguments.warmup = 30 if borrowed else 1
    if getattr(arguments, "samples", None) is None:
        arguments.samples = 500 if borrowed else 3


def run_profile(arguments: argparse.Namespace) -> int:
    _apply_run_sampling_defaults(arguments)
    abba_control = getattr(arguments, "control_binary", None)
    abba_candidate = getattr(arguments, "candidate_binary", None)
    if getattr(arguments, "include_one_paragraph_text", False) and (
        abba_control is None or abba_candidate is None
    ):
        raise ResourceProfileInputError(
            "--include-one-paragraph-text requires explicit control and candidate "
            "binaries for the DOCX semantic ABBA workload"
        )
    if getattr(arguments, "workload", None) is not None and (
        abba_control is None and abba_candidate is None
    ):
        raise ResourceProfileInputError(
            "--workload is only valid with explicit --control-binary and --candidate-binary"
        )
    if abba_control is not None or abba_candidate is not None:
        if abba_control is None or abba_candidate is None:
            raise ResourceProfileInputError(
                "--control-binary and --candidate-binary must be supplied together"
            )
        if arguments.build:
            raise ResourceProfileInputError(
                "ABBA mode accepts already-built release binaries; omit --build"
            )
        requested_workload = getattr(arguments, "workload", None)
        if arguments.only and arguments.only not in {
            XLSX_MANAGED_BATCH_ID,
            *DOCX_SEMANTIC_ID_ALIASES,
            *XLSX_XML_BORROWED_ID_ALIASES,
        }:
            raise ResourceProfileInputError(
                "ABBA mode only supports --only xlsx-managed-batch, "
                "xlsx-xml-borrowed, or docx-semantic"
            )
        if requested_workload in DOCX_SEMANTIC_ID_ALIASES:
            requested_workload = DOCX_SEMANTIC_ID
        elif requested_workload in XLSX_XML_BORROWED_ID_ALIASES:
            requested_workload = XLSX_XML_BORROWED_ID
        if requested_workload is None:
            requested_workload = (
                DOCX_SEMANTIC_ID
                if arguments.only in DOCX_SEMANTIC_ID_ALIASES
                else XLSX_XML_BORROWED_ID
                if arguments.only in XLSX_XML_BORROWED_ID_ALIASES
                else XLSX_MANAGED_BATCH_ID
            )
        if getattr(arguments, "include_one_paragraph_text", False) and (
            requested_workload != DOCX_SEMANTIC_ID
        ):
            raise ResourceProfileInputError(
                "--include-one-paragraph-text requires the DOCX semantic ABBA workload"
            )
        if arguments.only in DOCX_SEMANTIC_ID_ALIASES:
            only_workload = DOCX_SEMANTIC_ID
        elif arguments.only in XLSX_XML_BORROWED_ID_ALIASES:
            only_workload = XLSX_XML_BORROWED_ID
        else:
            only_workload = arguments.only
        if only_workload and only_workload != requested_workload:
            raise ResourceProfileInputError(
                "--only and --workload must select the same ABBA workload"
            )
        abba_arguments = argparse.Namespace(
            control_binary=abba_control,
            candidate_binary=abba_candidate,
            output=(
                DEFAULT_ABBA_OUTPUT
                if arguments.output == DEFAULT_OUTPUT
                else arguments.output
            ),
            artifact_dir=getattr(arguments, "artifact_dir", None),
            warmup=arguments.warmup,
            samples=arguments.samples,
            timeout=arguments.timeout,
            include_one_paragraph_text=getattr(
                arguments, "include_one_paragraph_text", False
            ),
        )
        if requested_workload == DOCX_SEMANTIC_ID:
            if arguments.output == DEFAULT_OUTPUT:
                abba_arguments.output = DEFAULT_DOCX_ABBA_OUTPUT
            return run_docx_semantic_abba(abba_arguments)
        if requested_workload == XLSX_XML_BORROWED_ID:
            if arguments.output == DEFAULT_OUTPUT:
                abba_arguments.output = DEFAULT_XLSX_XML_BORROWED_ABBA_OUTPUT
            return run_xlsx_xml_borrowed_abba(abba_arguments)
        if arguments.output == DEFAULT_OUTPUT:
            abba_arguments.output = DEFAULT_ABBA_OUTPUT
        return run_xlsx_managed_batch_abba(abba_arguments)
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
    run.add_argument(
        "--workload",
        choices=(
            XLSX_MANAGED_BATCH_ID,
            *XLSX_XML_BORROWED_ID_ALIASES,
            *DOCX_SEMANTIC_ID_ALIASES,
        ),
        help="resource ABBA workload when explicit control/candidate binaries are supplied",
    )
    run.add_argument(
        "--control-binary",
        "--control",
        "--before-binary",
        dest="control_binary",
        type=Path,
        help="switch to the retained resource ABBA mode selected by --workload/--only",
    )
    run.add_argument(
        "--candidate-binary",
        "--candidate",
        "--after-binary",
        dest="candidate_binary",
        type=Path,
        help="switch to the retained resource ABBA mode selected by --workload/--only",
    )
    run.add_argument("--artifact-dir", type=Path)
    run.add_argument(
        "--include-one-paragraph-text",
        "--include-docx-one-paragraph-text",
        dest="include_one_paragraph_text",
        action="store_true",
        help=(
            "include docx_semantic_one_paragraph_text in the DOCX resource ABBA "
            "case tuple"
        ),
    )
    run.add_argument("--warmup", type=int)
    run.add_argument("--samples", type=int)
    run.add_argument("--timeout", type=float, default=600.0)
    run.set_defaults(function=run_profile)
    abba = subparsers.add_parser(
        "compare-xlsx-managed-batch",
        aliases=("abba-xlsx-managed-batch", "abba", "compare"),
        help="run a retained A1/B1/B2/A2 resource comparison for XLSX managed batch",
    )
    abba.add_argument(
        "--control-binary",
        "--control",
        "--before-binary",
        dest="control_binary",
        type=Path,
        required=True,
    )
    abba.add_argument(
        "--candidate-binary",
        "--candidate",
        "--after-binary",
        dest="candidate_binary",
        type=Path,
        required=True,
    )
    abba.add_argument("--output", type=Path, default=DEFAULT_ABBA_OUTPUT)
    abba.add_argument(
        "--artifact-dir",
        type=Path,
        help="directory for retained per-leg harness/time/heaptrack artifacts",
    )
    abba.add_argument("--warmup", type=int, default=3)
    abba.add_argument("--samples", type=int, default=30)
    abba.add_argument("--timeout", type=float, default=600.0)
    abba.set_defaults(function=run_xlsx_managed_batch_abba)

    borrowed = subparsers.add_parser(
        "compare-xlsx-xml-borrowed",
        aliases=("abba-xlsx-xml-borrowed", "compare-xlsx-borrowed"),
        help=(
            "run strict retained A1/B1/B2/A2 resource evidence for the fixed "
            "XLSX borrowed-parser four-case workload"
        ),
    )
    borrowed.add_argument(
        "--control-binary",
        "--control",
        "--before-binary",
        dest="control_binary",
        type=Path,
        required=True,
    )
    borrowed.add_argument(
        "--candidate-binary",
        "--candidate",
        "--after-binary",
        dest="candidate_binary",
        type=Path,
        required=True,
    )
    borrowed.add_argument("--output", type=Path, default=DEFAULT_XLSX_XML_BORROWED_ABBA_OUTPUT)
    borrowed.add_argument(
        "--artifact-dir",
        type=Path,
        help="directory for retained per-leg harness/time/heaptrack artifacts",
    )
    borrowed.add_argument("--warmup", type=int, default=30)
    borrowed.add_argument("--samples", type=int, default=500)
    borrowed.add_argument("--timeout", type=float, default=600.0)
    borrowed.set_defaults(function=run_xlsx_xml_borrowed_abba)

    docx = subparsers.add_parser(
        "compare-docx-semantic",
        aliases=("abba-docx-semantic", "compare-docx", "abba-docx"),
        help=(
            "run a retained A1/B1/B2/A2 resource comparison for DOCX semantic "
            "open/full-text workloads (optionally one-paragraph text)"
        ),
    )
    docx.add_argument(
        "--control-binary",
        "--control",
        "--before-binary",
        dest="control_binary",
        type=Path,
        required=True,
    )
    docx.add_argument(
        "--candidate-binary",
        "--candidate",
        "--after-binary",
        dest="candidate_binary",
        type=Path,
        required=True,
    )
    docx.add_argument("--output", type=Path, default=DEFAULT_DOCX_ABBA_OUTPUT)
    docx.add_argument(
        "--artifact-dir",
        type=Path,
        help="directory for retained per-leg harness/time/heaptrack artifacts",
    )
    docx.add_argument(
        "--include-one-paragraph-text",
        "--include-docx-one-paragraph-text",
        dest="include_one_paragraph_text",
        action="store_true",
        help="include docx_semantic_one_paragraph_text in the resource case tuple",
    )
    docx.add_argument("--warmup", type=int, default=3)
    docx.add_argument("--samples", type=int, default=30)
    docx.add_argument("--timeout", type=float, default=600.0)
    docx.set_defaults(function=run_docx_semantic_abba)
    reprocess_docx = subparsers.add_parser(
        "reprocess-docx-heaptrack",
        help="reparse hash-verified heaptrack artifacts from a retained DOCX report",
    )
    reprocess_docx.add_argument("--input", type=Path, required=True)
    reprocess_docx.add_argument("--output", type=Path, required=True)
    reprocess_docx.set_defaults(function=run_reprocess_docx_heaptrack)
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    arguments = build_parser().parse_args(argv)
    if arguments.command == "run":
        _apply_run_sampling_defaults(arguments)
        if arguments.warmup < 0 or arguments.samples < 1 or arguments.timeout <= 0:
            raise SystemExit("--warmup must be non-negative, --samples and --timeout must be positive")
    elif arguments.command in {
        "compare-xlsx-managed-batch",
        "abba-xlsx-managed-batch",
        "abba",
        "compare",
        "compare-xlsx-xml-borrowed",
        "abba-xlsx-xml-borrowed",
        "compare-xlsx-borrowed",
        "compare-docx-semantic",
        "abba-docx-semantic",
        "compare-docx",
        "abba-docx",
    }:
        if arguments.warmup < 0 or arguments.samples < 1 or arguments.timeout <= 0:
            raise SystemExit("--warmup must be non-negative, --samples and --timeout must be positive")
    try:
        return int(arguments.function(arguments))
    except (OSError, RuntimeError, ValueError) as error:
        print(f"{TOOL_NAME}: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
