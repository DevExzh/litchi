#!/usr/bin/env python3
"""Fail-closed quiet-host ABBA driver for the OPC managed read-ahead lane.

This driver is deliberately separate from the build/correctness prefix and
from the historical all-in-one shell recipe.  It owns only latency captures:
the source pair, probe and binaries are supplied as immutable inputs.  Every
child is run in its own process group and CPU affinity, with a bounded address
space, output streams, report, target and deadline.  Full ``/proc`` snapshots
are retained before, during and after every child.  A busy host, a surviving
descendant, a changed source, malformed report, or any resource violation
fails closed.

The normal binary's allocation field is intentionally not interpreted.  This
is a latency-only lane; allocator qualification requires a separate binary
and protocol.
"""

from __future__ import annotations

import argparse
import fcntl
import hashlib
import json
import os
from pathlib import Path
import resource
import signal
import subprocess
import sys
import threading
import time
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any, Iterable, Mapping


SCHEMA = "litchi.opc.managed-read-ahead.timing.v1"
REPORT_SCHEMA = "docx_managed_read_ahead_v1"
ARCHIVE_BYTES = 16_793_036
ARCHIVE_SHA256 = "a4a2e4921235a6da6b38e31d26ddcca1301909885e37330ab4f83ecc0c4e04f4"
TEXT_BYTES = 10_000
TEXT_SHA256 = "ad4fe690f0ef2281ad8e64a78d1f4d64e7c8625d672b3ac2f7e9fbc28a82f4af"
EXACT_RANGES = [
    (16_793_014, 22),
    (16_791_709, 46),
    (16_791_709, 1_305),
    (0, 30),
    (49, 363),
    (412, 16),
    (412, 16),
    (428, 30),
    (469, 234),
    (703, 16),
    (703, 16),
    (2_907, 30),
    (2_965, 324),
    (3_289, 16),
    (3_289, 16),
    (1_420, 30),
    (1_467, 1_424),
    (2_891, 16),
    (2_891, 16),
]
FORWARD_RANGES = [(16_793_014, 22), (16_791_709, 1_327), (0, 4_096)]
HEAVY_NAMES = frozenset(
    {
        "cargo",
        "cargo-fuzz",
        "rustc",
        "rustdoc",
        "rustfmt",
        "clippy-driver",
        "valgrind",
        "perf",
        "heaptrack",
    }
)
DEFAULT_MAX_OUTPUT_BYTES = 4 * 1024 * 1024
DEFAULT_MAX_REPORT_BYTES = 4 * 1024 * 1024
DEFAULT_MAX_RSS_BYTES = 20_000_000_000
DEFAULT_MAX_TARGET_BYTES = 20_000_000_000
DEFAULT_TIMEOUT_SECONDS = 120.0
DEFAULT_POLL_SECONDS = 0.1
ABBA_ORDER = ("a1", "b1", "b2", "a2")


class TimingError(RuntimeError):
    """A source, host, child, output, or report contract failed."""


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="milliseconds").replace("+00:00", "Z")


def sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise TimingError(f"expected regular file, not symlink: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def write_json_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    if path.exists() or path.is_symlink():
        raise TimingError(f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise TimingError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def read_json(path: Path, max_bytes: int) -> dict[str, Any]:
    if path.is_symlink() or not path.is_file():
        raise TimingError(f"missing or symlinked JSON file: {path}")
    if path.stat().st_size > max_bytes:
        raise TimingError(f"JSON file exceeds cap: {path}")
    try:
        value = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_duplicate_keys)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise TimingError(f"invalid JSON {path}: {error}") from error
    if not isinstance(value, dict):
        raise TimingError(f"JSON root must be an object: {path}")
    return value


def require(condition: bool, message: str) -> None:
    if not condition:
        raise TimingError(message)


def require_equal(actual: Any, expected: Any, location: str) -> None:
    if actual != expected:
        raise TimingError(f"{location}: expected {expected!r}, got {actual!r}")


def range_dicts(ranges: Iterable[tuple[int, int]]) -> list[dict[str, int]]:
    return [{"offset": offset, "requested": length, "returned": length} for offset, length in ranges]


def directory_size(path: Path) -> int:
    if not path.exists():
        return 0
    if path.is_symlink():
        raise TimingError(f"target is a symlink: {path}")
    if path.is_file():
        return path.stat().st_size
    total = 0
    for root, dirs, files in os.walk(path, followlinks=False):
        dirs[:] = [name for name in dirs if not (Path(root) / name).is_symlink()]
        for name in files:
            item = Path(root) / name
            if item.is_symlink():
                raise TimingError(f"target contains a symlink: {item}")
            try:
                total += item.stat().st_size
            except FileNotFoundError:
                continue
    return total


def _proc_stat_fields(raw: str) -> list[str]:
    right = raw.rsplit(") ", 1)
    if len(right) != 2:
        raise ValueError("malformed /proc stat")
    return right[1].split()


def read_proc(pid: int) -> dict[str, Any] | None:
    proc = Path("/proc") / str(pid)
    try:
        comm = (proc / "comm").read_text(encoding="utf-8", errors="replace").strip()
        cmdline = [part.decode("utf-8", "replace") for part in (proc / "cmdline").read_bytes().split(b"\0") if part]
        exe = os.readlink(proc / "exe")
        cwd = os.readlink(proc / "cwd")
        fields = _proc_stat_fields((proc / "stat").read_text(encoding="utf-8", errors="replace"))
        status = (proc / "status").read_text(encoding="utf-8", errors="replace")
    except (FileNotFoundError, PermissionError, OSError, ValueError):
        return None
    try:
        rss = next(int(line.split()[1]) * os.sysconf("SC_PAGE_SIZE") for line in status.splitlines() if line.startswith("VmRSS:"))
    except (StopIteration, ValueError):
        rss = 0
    try:
        return {
            "pid": pid,
            "ppid": int(fields[1]),
            "pgrp": int(fields[2]),
            "state": fields[0],
            "utime": int(fields[11]),
            "stime": int(fields[12]),
            "psr": int(fields[36]),
            "rss_bytes": rss,
            "comm": comm,
            "exe": exe,
            "cwd": cwd,
            "cmdline": cmdline,
        }
    except (IndexError, ValueError):
        return None


def proc_snapshot() -> dict[int, dict[str, Any]]:
    result: dict[int, dict[str, Any]] = {}
    for item in Path("/proc").iterdir():
        if item.name.isdigit():
            entry = read_proc(int(item.name))
            if entry is not None:
                result[entry["pid"]] = entry
    return result


def load_average() -> list[float]:
    try:
        return [float(value) for value in os.getloadavg()]
    except (AttributeError, OSError):
        return []


def descendants(snapshot: Mapping[int, Mapping[str, Any]], root: int) -> set[int]:
    owned = {root}
    changed = True
    while changed:
        changed = False
        for pid, entry in snapshot.items():
            if pid not in owned and int(entry.get("ppid", -1)) in owned:
                owned.add(pid)
                changed = True
    # A detached descendant can be reparented while retaining the process
    # group.  Keep it owned for cleanup and do not mistake it for host work.
    owned.update(pid for pid, entry in snapshot.items() if int(entry.get("pgrp", -1)) == root)
    return owned


def is_heavy(entry: Mapping[str, Any]) -> bool:
    names = {Path(str(entry.get("exe", ""))).name, str(entry.get("comm", ""))}
    if names & HEAVY_NAMES:
        return True
    command = entry.get("cmdline", [])
    return bool(command) and Path(str(command[0])).name in HEAVY_NAMES


def snapshot_json(snapshot: Mapping[int, Mapping[str, Any]], load: list[float]) -> dict[str, Any]:
    return {
        "utc": utc_now(),
        "loadavg": load,
        "processes": [snapshot[pid] for pid in sorted(snapshot)],
    }


def quiet_violations(
    previous: Mapping[int, Mapping[str, Any]],
    current: Mapping[int, Mapping[str, Any]],
    *,
    allowed_root: int | None,
    cpu: int,
    load_before: float,
    max_load_delta: float,
) -> list[str]:
    # The monitor itself is expected to consume a small amount of CPU between
    # snapshots.  Never classify that work as an unrelated selected-CPU user.
    owned = {os.getpid()}
    if allowed_root is not None:
        owned.update(descendants(current, allowed_root))
    violations: list[str] = []
    heavy = [
        entry
        for pid, entry in current.items()
        if pid not in owned and is_heavy(entry)
    ]
    if heavy:
        violations.append("unrelated heavy process: " + repr(heavy))
    busy: list[dict[str, Any]] = []
    for pid, entry in current.items():
        if pid in owned or int(entry.get("psr", -1)) != cpu:
            continue
        if str(entry.get("comm", "")).startswith("["):
            continue
        old = previous.get(pid)
        if old is not None and int(entry.get("utime", 0)) + int(entry.get("stime", 0)) > int(old.get("utime", 0)) + int(old.get("stime", 0)):
            busy.append({"pid": pid, "comm": entry.get("comm"), "psr": entry.get("psr")})
    if busy:
        violations.append("another process used selected CPU: " + repr(busy))
    load = load_average()
    if load and load[0] > load_before + max_load_delta:
        violations.append(f"load1 increased beyond quiet bound: before={load_before}, now={load[0]}")
    return violations


def process_rss(snapshot: Mapping[int, Mapping[str, Any]], root: int) -> int:
    return sum(int(snapshot[pid].get("rss_bytes", 0)) for pid in descendants(snapshot, root) if pid in snapshot)


def process_group_exists(pgid: int) -> bool:
    try:
        os.killpg(pgid, 0)
    except ProcessLookupError:
        return False
    except PermissionError:
        return True
    return True


def terminate_group(process: subprocess.Popen[bytes], *, grace_seconds: float = 5.0) -> bool:
    forced = False
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except (ProcessLookupError, PermissionError):
        process.terminate()
    deadline = time.monotonic() + grace_seconds
    while time.monotonic() < deadline and process.poll() is None:
        time.sleep(0.05)
    if process.poll() is None or process_group_exists(process.pid):
        forced = True
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except (ProcessLookupError, PermissionError):
            process.kill()
    try:
        process.wait(timeout=grace_seconds)
    except subprocess.TimeoutExpired as error:
        raise TimingError(f"child did not terminate after process-group cleanup: {process.pid}") from error
    deadline = time.monotonic() + grace_seconds
    while time.monotonic() < deadline and process_group_exists(process.pid):
        time.sleep(0.05)
    if process_group_exists(process.pid):
        raise TimingError(f"descendant process group survived cleanup: {process.pid}")
    return forced


class BoundedCapture(threading.Thread):
    def __init__(self, pipe: Any, path: Path, maximum: int):
        super().__init__(daemon=True)
        self.pipe = pipe
        self.path = path
        self.maximum = maximum
        self.bytes = 0
        self.overflow = False
        self.error: str | None = None

    def run(self) -> None:
        try:
            with self.path.open("xb") as output:
                while True:
                    chunk = self.pipe.read(64 * 1024)
                    if not chunk:
                        break
                    remaining = self.maximum - self.bytes
                    if remaining > 0:
                        output.write(chunk[:remaining])
                        self.bytes += min(len(chunk), remaining)
                    if len(chunk) > remaining:
                        self.overflow = True
                output.flush()
                os.fsync(output.fileno())
        except BaseException as error:  # retain the failure for the parent loop
            self.error = f"{type(error).__name__}: {error}"


def validate_report(path: Path, endpoint: str, revision: str, policy: str, warmup: int, samples: int, max_bytes: int) -> dict[str, Any]:
    report = read_json(path, max_bytes)
    require_equal(report.get("schema"), REPORT_SCHEMA, "report.schema")
    require_equal(report.get("version"), 1, "report.version")
    require_equal(report.get("requested_source_revision"), revision, "report.requested_source_revision")
    require_equal(report.get("source_bytes"), ARCHIVE_BYTES, "report.source_bytes")
    require_equal(report.get("source_sha256"), ARCHIVE_SHA256, "report.source_sha256")
    require_equal(report.get("expected_text_bytes"), TEXT_BYTES, "report.expected_text_bytes")
    require_equal(report.get("expected_text_sha256"), TEXT_SHA256, "report.expected_text_sha256")
    require_equal(report.get("expected_archive_members"), 20, "report.expected_archive_members")
    require_equal(report.get("warmup"), warmup, "report.warmup")
    require_equal(report.get("samples"), samples, "report.samples")
    provider = report.get("provider")
    require(isinstance(provider, dict), "report.provider must be an object")
    require_equal(provider.get("name"), "PptxRangeSource", "report.provider.name")
    require_equal(provider.get("max_range_bytes"), 65_536, "report.provider.max_range_bytes")
    require_equal(provider.get("delay_us"), 0, "report.provider.delay_us")
    require_equal(provider.get("transfer_bytes_per_second"), None, "report.provider.transfer_bytes_per_second")
    corpus = report.get("corpus")
    require(isinstance(corpus, dict), "report.corpus must be an object")
    require_equal(corpus.get("archive_bytes"), ARCHIVE_BYTES, "report.corpus.archive_bytes")
    require_equal(corpus.get("archive_sha256"), ARCHIVE_SHA256, "report.corpus.archive_sha256")
    require_equal(corpus.get("target_payload_bytes"), TEXT_BYTES, "report.corpus.target_payload_bytes")
    require_equal(corpus.get("target_payload_sha256"), TEXT_SHA256, "report.corpus.target_payload_sha256")
    media_ranges = report.get("media_ranges")
    require(isinstance(media_ranges, list) and len(media_ranges) == 8, "report.media_ranges must contain eight members")
    expected = FORWARD_RANGES if policy == "forward" else EXACT_RANGES
    policy_record = report.get("policy")
    require(isinstance(policy_record, dict), "report.policy must be an object")
    require_equal(policy_record.get("name"), "forward_start" if policy == "forward" else "exact", "report.policy.name")
    require_equal(policy_record.get("enabled"), policy == "forward", "report.policy.enabled")
    require_equal(policy_record.get("configured_window_bytes"), 4_096 if policy == "forward" else None, "report.policy.configured_window_bytes")
    rows = report.get("rows")
    require(isinstance(rows, list) and len(rows) == samples, "report.rows length does not match samples")
    for row_index, row in enumerate(rows):
        require(isinstance(row, dict), f"report.rows[{row_index}] must be an object")
        require_equal(row.get("sample_index"), warmup + row_index, f"rows[{row_index}].sample_index")
        require(isinstance(row.get("elapsed_ns"), int) and row["elapsed_ns"] > 0, f"rows[{row_index}].elapsed_ns")
        require_equal(row.get("actual_text_verified"), True, f"rows[{row_index}].actual_text_verified")
        require_equal(row.get("actual_text_bytes"), TEXT_BYTES, f"rows[{row_index}].actual_text_bytes")
        require_equal(row.get("actual_text_sha256"), TEXT_SHA256, f"rows[{row_index}].actual_text_sha256")
        require_equal(row.get("source_version_unchanged"), True, f"rows[{row_index}].source_version_unchanged")
        require_equal(row.get("policy"), policy_record, f"rows[{row_index}].policy")
        cache = row.get("cache")
        require(isinstance(cache, dict), f"rows[{row_index}].cache")
        require_equal(cache.get("open_successful_loads"), 0, f"rows[{row_index}].cache.open_successful_loads")
        require_equal(cache.get("successful_loads"), 1, f"rows[{row_index}].cache.successful_loads")
        require_equal(cache.get("failed_loads"), 0, f"rows[{row_index}].cache.failed_loads")
        require_equal(cache.get("budget_managed"), True, f"rows[{row_index}].cache.budget_managed")
        physical = row.get("physical")
        require(isinstance(physical, dict), f"rows[{row_index}].physical")
        require_equal(physical.get("calls"), len(expected), f"rows[{row_index}].physical.calls")
        expected_bytes = sum(length for _, length in expected)
        require_equal(physical.get("requested_bytes"), expected_bytes, f"rows[{row_index}].physical.requested_bytes")
        require_equal(physical.get("returned_bytes"), expected_bytes, f"rows[{row_index}].physical.returned_bytes")
        require_equal(physical.get("short_reads"), 0, f"rows[{row_index}].physical.short_reads")
        require_equal(physical.get("ranges"), range_dicts(expected), f"rows[{row_index}].physical.ranges")
        source_read = row.get("source_read")
        if policy == "forward":
            require(isinstance(source_read, dict), f"rows[{row_index}].source_read")
            for key, value in {"enabled": True, "configured_window_bytes": 4_096, "requests": 19, "hits": 16, "misses": 3, "fills": 3, "requested_bytes": 3_966, "returned_bytes": 5_445}.items():
                require_equal(source_read.get(key), value, f"rows[{row_index}].source_read.{key}")
        else:
            require_equal(source_read, None, f"rows[{row_index}].source_read")
        budget = row.get("budget")
        require(isinstance(budget, dict), f"rows[{row_index}].budget")
        require_equal(budget.get("memory_after_drop"), 0, f"rows[{row_index}].budget.memory_after_drop")
        require_equal(budget.get("objects_after_drop"), 0, f"rows[{row_index}].budget.objects_after_drop")
        require_equal(budget.get("managed"), True, f"rows[{row_index}].budget.managed")
        require_equal(budget.get("input_bytes_match_physical_returned"), True, f"rows[{row_index}].budget.input_bytes_match_physical_returned")
        allocation = row.get("allocation")
        require(isinstance(allocation, dict), f"rows[{row_index}].allocation")
        require(allocation.get("sample") in (None, {}), f"rows[{row_index}].allocation.sample must be unavailable in latency-only mode")
    return {
        "endpoint": endpoint,
        "revision": revision,
        "policy": policy,
        "samples": samples,
        "warmup": warmup,
        "report_sha256": sha256_file(path)[0],
        "report_bytes": path.stat().st_size,
    }


def safe_target_check(path: Path, maximum: int) -> int:
    size = directory_size(path)
    if size > maximum:
        raise TimingError(f"target exceeds cap: {path} ({size} > {maximum})")
    return size


def verify_source(root: Path, revision: str, tree: str) -> dict[str, str]:
    """Verify that a timing endpoint is the exact clean source checkout."""
    if not root.is_dir() or root.is_symlink():
        raise TimingError(f"source checkout is missing or symlinked: {root}")

    def git(*arguments: str) -> str:
        completed = subprocess.run(
            ["git", "-C", str(root), *arguments],
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        if completed.returncode != 0:
            detail = completed.stderr.strip() or completed.stdout.strip()
            raise TimingError(f"git {' '.join(arguments)} failed for {root}: {detail}")
        return completed.stdout.strip()

    actual_revision = git("rev-parse", "HEAD")
    actual_tree = git("rev-parse", "HEAD^{tree}")
    status = git("status", "--porcelain=v1", "--untracked-files=all")
    require_equal(actual_revision, revision, f"{root} HEAD")
    require_equal(actual_tree, tree, f"{root} HEAD^{{tree}}")
    require_equal(status, "", f"{root} status")
    return {"path": str(root), "revision": actual_revision, "tree": actual_tree}


def acquire_cpu_lock(path: Path) -> Any:
    """Hold the protocol's external CPU lock for the complete ABBA sequence."""
    if path.is_symlink():
        raise TimingError(f"CPU lock is symlinked: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    handle = path.open("a+b")
    try:
        fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
    except (BlockingIOError, OSError) as error:
        handle.close()
        raise TimingError(f"CPU lock is busy: {path}: {error}") from error
    return handle


def run_child(
    *,
    output: Path,
    endpoint: str,
    binary: Path,
    revision: str,
    source_tree: str,
    policy: str,
    label: str,
    source_cwd: Path,
    target: Path,
    tmpdir: Path,
    cpu: int,
    samples: int,
    warmup: int,
    timeout_seconds: float,
    poll_seconds: float,
    max_rss: int,
    max_target: int,
    max_output: int,
    max_report: int,
    max_load_delta: float,
) -> dict[str, Any]:
    report_path = output / "reports" / policy / f"{label}.json"
    stdout_path = output / "logs" / f"{policy}-{label}.stdout"
    stderr_path = output / "logs" / f"{policy}-{label}.stderr"
    event_path = output / "events" / f"{policy}-{label}.json"
    for path in (report_path, stdout_path, stderr_path, event_path):
        if path.exists() or path.is_symlink():
            raise TimingError(f"output path is not fresh: {path}")
    if not binary.is_file() or not os.access(binary, os.X_OK) or binary.is_symlink():
        raise TimingError(f"binary is missing, non-executable, or symlinked: {binary}")
    source_before = verify_source(source_cwd, revision, source_tree)
    if cpu not in os.sched_getaffinity(0):
        raise TimingError(f"selected CPU is not allowed: {cpu}")
    tmpdir.mkdir(parents=True, exist_ok=True)
    report_path.parent.mkdir(parents=True, exist_ok=True)
    stdout_path.parent.mkdir(parents=True, exist_ok=True)
    event_path.parent.mkdir(parents=True, exist_ok=True)
    before = proc_snapshot()
    before_load = load_average()
    time.sleep(poll_seconds)
    quiet_probe = proc_snapshot()
    probe_load = load_average()
    pre_violations = quiet_violations(
        before,
        quiet_probe,
        allowed_root=None,
        cpu=cpu,
        load_before=before_load[0] if before_load else 0.0,
        max_load_delta=max_load_delta,
    )
    if pre_violations:
        raise TimingError(f"quiet-host preflight rejected {endpoint}/{policy}/{label}: {pre_violations}")
    safe_target_check(target, max_target)
    env = os.environ.copy()
    env.pop("LITCHI_CAPPED", None)
    env.pop("LITCHI_BLESS", None)
    env.pop("RUSTFLAGS", None)
    env["TMPDIR"] = str(tmpdir)
    command = [
        str(binary),
        "docx-managed-read-ahead",
        "--policy",
        "forward-start" if policy == "forward" else "exact",
    ]
    if policy == "forward":
        command.extend(["--window-bytes", "4096"])
    command.extend(
        [
            "--max-range",
            "65536",
            "--delay-us",
            "0",
            "--samples",
            str(samples),
            "--warmup",
            str(warmup),
            "--source-revision",
            revision,
            "--output",
            str(report_path),
        ]
    )

    def preexec() -> None:
        os.sched_setaffinity(0, {cpu})
        resource.setrlimit(resource.RLIMIT_AS, (max_rss, max_rss))
        resource.setrlimit(resource.RLIMIT_CORE, (0, 0))

    process: subprocess.Popen[bytes] | None = None
    stdout_capture: BoundedCapture | None = None
    stderr_capture: BoundedCapture | None = None
    violations: list[str] = []
    timeout = False
    forced_cleanup = False
    peak_rss = 0
    started_utc: str | None = None
    target_before = directory_size(target)
    target_after = target_before
    during_path = output / "host" / f"{policy}-{label}-during.jsonl"
    during_path.parent.mkdir(parents=True, exist_ok=True)
    if during_path.exists() or during_path.is_symlink():
        raise TimingError(f"host output path is not fresh: {during_path}")
    with during_path.open("x", encoding="utf-8") as during:
        try:
            started_utc = utc_now()
            process = subprocess.Popen(
                command,
                cwd=str(source_cwd),
                env=env,
                stdin=subprocess.DEVNULL,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                start_new_session=True,
                preexec_fn=preexec,
            )
            require(process.stdout is not None and process.stderr is not None, "child pipes were not created")
            stdout_capture = BoundedCapture(process.stdout, stdout_path, max_output)
            stderr_capture = BoundedCapture(process.stderr, stderr_path, max_output)
            stdout_capture.start()
            stderr_capture.start()
            previous = quiet_probe
            started = time.monotonic()
            while process.poll() is None:
                time.sleep(poll_seconds)
                current = proc_snapshot()
                current_load = load_average()
                during.write(json.dumps(snapshot_json(current, current_load), sort_keys=True) + "\n")
                during.flush()
                violations.extend(
                    quiet_violations(
                        previous,
                        current,
                        allowed_root=process.pid,
                        cpu=cpu,
                        load_before=before_load[0] if before_load else 0.0,
                        max_load_delta=max_load_delta,
                    )
                )
                previous = current
                peak_rss = max(peak_rss, process_rss(current, process.pid))
                if peak_rss > max_rss:
                    violations.append(f"RSS exceeded cap: {peak_rss} > {max_rss}")
                target_after = safe_target_check(target, max_target)
                if report_path.exists() and report_path.stat().st_size > max_report:
                    violations.append(f"report exceeded cap: {report_path}")
                if stdout_capture.overflow or stderr_capture.overflow:
                    violations.append("stdout/stderr exceeded output cap")
                if violations:
                    break
                if time.monotonic() - started > timeout_seconds:
                    timeout = True
                    violations.append(f"child exceeded timeout: {timeout_seconds}s")
                    break
            if process.poll() is None or violations or timeout:
                forced_cleanup = terminate_group(process)
            else:
                process.wait(timeout=5)
        finally:
            if process is not None and process.poll() is None:
                forced_cleanup = terminate_group(process) or forced_cleanup
    if stdout_capture is not None:
        stdout_capture.join(timeout=5)
    if stderr_capture is not None:
        stderr_capture.join(timeout=5)
    after = proc_snapshot()
    after_load = load_average()
    source_after: dict[str, str] | None = None
    try:
        source_after = verify_source(source_cwd, revision, source_tree)
    except TimingError as error:
        violations.append(f"source changed during child: {error}")
    post_violations = quiet_violations(
        quiet_probe,
        after,
        allowed_root=None,
        cpu=cpu,
        load_before=before_load[0] if before_load else 0.0,
        max_load_delta=max_load_delta,
    )
    violations.extend(post_violations)
    if process is None:
        raise TimingError("child was not started")
    descendants_left = sorted(descendants(after, process.pid) - {process.pid})
    if descendants_left:
        violations.append(f"descendants survived child exit: {descendants_left}")
    if forced_cleanup:
        violations.append("forced process-group cleanup was required")
    if stdout_capture is not None and stdout_capture.error:
        violations.append(stdout_capture.error)
    if stderr_capture is not None and stderr_capture.error:
        violations.append(stderr_capture.error)
    target_after = safe_target_check(target, max_target)
    validation: dict[str, Any] | None = None
    if process.returncode == 0 and not violations and not timeout:
        validation = validate_report(report_path, endpoint, revision, policy, warmup, samples, max_report)
    elif process.returncode != 0:
        violations.append(f"child exited with status {process.returncode}")
    event = {
        "schema": SCHEMA,
        "endpoint": endpoint,
        "policy": policy,
        "label": label,
        "argv": command,
        "cwd": str(source_cwd),
        "revision": revision,
        "source": {"before": source_before, "after": source_after, "expected_tree": source_tree},
        "started_utc": started_utc,
        "ended_utc": utc_now(),
        "returncode": process.returncode,
        "timeout": timeout,
        "forced_cleanup": forced_cleanup,
        "cpu": cpu,
        "rss_peak_bytes": peak_rss,
        "rss_cap_bytes": max_rss,
        "target": {"path": str(target), "bytes_before": target_before, "bytes_after": target_after, "cap_bytes": max_target},
        "stdout": {"path": str(stdout_path), "bytes": stdout_capture.bytes if stdout_capture else 0, "overflow": stdout_capture.overflow if stdout_capture else True},
        "stderr": {"path": str(stderr_path), "bytes": stderr_capture.bytes if stderr_capture else 0, "overflow": stderr_capture.overflow if stderr_capture else True},
        "report": {"path": str(report_path), "bytes": report_path.stat().st_size if report_path.exists() else 0},
        "host": {
            "before": snapshot_json(before, before_load),
            "probe": snapshot_json(quiet_probe, probe_load),
            "during_jsonl": str(during_path),
            "after": snapshot_json(after, after_load),
        },
        "violations": violations,
        "validation": validation,
    }
    write_json_exclusive(event_path, event)
    if violations:
        raise TimingError(f"{endpoint}/{policy}/{label} rejected: {violations}")
    require(validation is not None, f"{endpoint}/{policy}/{label} has no validation result")
    return event


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-binary", required=True, type=Path)
    parser.add_argument("--candidate-binary", required=True, type=Path)
    parser.add_argument("--baseline-revision", required=True)
    parser.add_argument("--candidate-revision", required=True)
    parser.add_argument("--baseline-source", required=True, type=Path)
    parser.add_argument("--candidate-source", required=True, type=Path)
    parser.add_argument("--baseline-tree", required=True)
    parser.add_argument("--candidate-tree", required=True)
    parser.add_argument("--target-baseline", required=True, type=Path)
    parser.add_argument("--target-candidate", required=True, type=Path)
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--tmpdir", required=True, type=Path)
    parser.add_argument("--cpu-lock", required=True, type=Path)
    parser.add_argument("--cpu", required=True, type=int)
    parser.add_argument("--samples", type=int, default=30)
    parser.add_argument("--warmup", type=int, default=3)
    parser.add_argument("--timeout-seconds", type=float, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--poll-seconds", type=float, default=DEFAULT_POLL_SECONDS)
    parser.add_argument("--max-rss-bytes", type=int, default=DEFAULT_MAX_RSS_BYTES)
    parser.add_argument("--max-target-bytes", type=int, default=DEFAULT_MAX_TARGET_BYTES)
    parser.add_argument("--max-output-bytes", type=int, default=DEFAULT_MAX_OUTPUT_BYTES)
    parser.add_argument("--max-report-bytes", type=int, default=DEFAULT_MAX_REPORT_BYTES)
    parser.add_argument("--max-load-delta", type=float, default=1.0)
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        require(
            len(args.baseline_revision) == 40
            and len(args.candidate_revision) == 40
            and len(args.baseline_tree) == 40
            and len(args.candidate_tree) == 40,
            "revisions and trees must be 40-character hashes",
        )
        require(args.samples > 0 and args.warmup >= 0, "samples/warmup are invalid")
        require(args.timeout_seconds > 0 and args.poll_seconds > 0, "timeouts are invalid")
        require(args.max_rss_bytes > 0 and args.max_target_bytes > 0, "resource caps are invalid")
        require(args.max_output_bytes > 0 and args.max_report_bytes > 0, "output caps are invalid")
        require(args.output_dir.is_absolute() and not args.output_dir.exists() and not args.output_dir.is_symlink(), "output-dir must be a fresh absent absolute path")
        sources = {"baseline": args.baseline_source.resolve(), "candidate": args.candidate_source.resolve()}
        trees = {"baseline": args.baseline_tree, "candidate": args.candidate_tree}
        revisions = {"baseline": args.baseline_revision, "candidate": args.candidate_revision}
        source_identities = {endpoint: verify_source(sources[endpoint], revisions[endpoint], trees[endpoint]) for endpoint in ("baseline", "candidate")}
        args.output_dir.mkdir(parents=True)
        for directory in (args.target_baseline, args.target_candidate):
            require(not directory.is_symlink(), f"target is symlinked: {directory}")
            safe_target_check(directory, args.max_target_bytes)
        args.tmpdir.mkdir(parents=True, exist_ok=True)
        args.output_dir = args.output_dir.resolve()
        args.tmpdir = args.tmpdir.resolve()
        binaries = {"baseline": args.baseline_binary.resolve(), "candidate": args.candidate_binary.resolve()}
        targets = {"baseline": args.target_baseline.resolve(), "candidate": args.target_candidate.resolve()}
        events: list[dict[str, Any]] = []
        cpu_lock = acquire_cpu_lock(args.cpu_lock.resolve())
        try:
            for policy in ("forward", "exact"):
                for label in ABBA_ORDER:
                    endpoint = "baseline" if label.startswith("a") else "candidate"
                    events.append(
                        run_child(
                            output=args.output_dir,
                            endpoint=endpoint,
                            binary=binaries[endpoint],
                            revision=revisions[endpoint],
                            source_tree=trees[endpoint],
                            policy=policy,
                            label=label,
                            source_cwd=sources[endpoint],
                            target=targets[endpoint],
                            tmpdir=args.tmpdir,
                            cpu=args.cpu,
                            samples=args.samples,
                            warmup=args.warmup,
                            timeout_seconds=args.timeout_seconds,
                            poll_seconds=args.poll_seconds,
                            max_rss=args.max_rss_bytes,
                            max_target=args.max_target_bytes,
                            max_output=args.max_output_bytes,
                            max_report=args.max_report_bytes,
                            max_load_delta=args.max_load_delta,
                        )
                    )
        finally:
            fcntl.flock(cpu_lock.fileno(), fcntl.LOCK_UN)
            cpu_lock.close()
        receipt = {
            "schema": SCHEMA,
            "classification": "latency-only capture; allocator is not applicable",
            "sources": source_identities,
            "source_revisions": revisions,
            "source_trees": trees,
            "binaries": {endpoint: {"path": str(path), "sha256": sha256_file(path)[0], "bytes": sha256_file(path)[1]} for endpoint, path in binaries.items()},
            "cpu": args.cpu,
            "cpu_lock": str(args.cpu_lock.resolve()),
            "abba_order": list(ABBA_ORDER),
            "policies": ["forward", "exact"],
            "samples": args.samples,
            "warmup": args.warmup,
            "resource_caps": {"rss_bytes": args.max_rss_bytes, "target_bytes": args.max_target_bytes, "output_bytes": args.max_output_bytes, "report_bytes": args.max_report_bytes, "timeout_seconds": args.timeout_seconds},
            "events": [event["endpoint"] + "/" + event["policy"] + "/" + event["label"] for event in events],
            "qualification_status": "capture_complete_pending_analysis",
        }
        write_json_exclusive(args.output_dir / "receipt.json", receipt)
        return 0
    except (OSError, TimingError, ValueError) as error:
        print(f"opc read-ahead timing rejected: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
