#!/usr/bin/env python3
"""Run the native XLS F3 probe in a strict, calibrated A1/B1/B2/A2 lane.

The driver never builds the probe. It validates the two clean, pinned source
snapshots and their copied probe sources, calibrates each endpoint separately,
then runs every requested route/corpus cell in a fixed ABBA order. A normal
run is the complete three-route by two-corpus matrix. A subset is permitted
only with ``--diagnostic`` and is explicitly unqualified in the receipt.

Every child is started with an actual Linux CPU affinity and RLIMIT_AS scope;
the driver's process/load snapshots before and after each child are retained in
the external receipt. There are no retries or reordered legs.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import platform
import resource
import signal
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path
from typing import Any


SCHEMA = "litchi.xls.f3.abba.v2"
ROLES = ("A1", "B1", "B2", "A2")
ROUTES = ("write-only", "first-read", "read-after-write")
CORPORA = ("number", "rk-mulrk")
HEX_DIGEST_LENGTH = 64
MIN_LEG_NS = 250_000_000
DEFAULT_MEMORY_LIMIT_BYTES = 2 * 1024**3
HARD_MEMORY_LIMIT_BYTES = 2 * 1024**3
MAX_CHILD_OUTPUT_BYTES = 16 * 1024 * 1024
DEFAULT_CALIBRATION_START_BATCH = 1
DEFAULT_CALIBRATION_MAX_BATCH = 1 << 20
BASELINE_LINEAGE = "9b1f96d6dce1506dc9d0080f0696b42ba592c397"
CANDIDATE_LINEAGE = "8acd229235e3c01c6e529f6d1f055d84b8e03de4"
PROBE_COMMITS = (
    "d8fe3bcacb0c51cf354314393f93cef0a393e157",
    "b03aaf1aace43fab45666dc3a1507f3da5922d5a",
)
PROBE_PATHS = (
    "tools/perf-baseline/src/xls_f3_qualification.rs",
    "tools/perf-baseline/examples/xls_f3_qualification.rs",
    "tools/perf-baseline/src/xls_numeric.rs",
    "tools/perf-baseline/src/lib.rs",
    "tools/xls_f3_qualification_abba.py",
)
REQUIRED_COLUMNS = (
    "route",
    "corpus",
    "sample",
    "nanoseconds",
    "edit_ns",
    "set_ns",
    "commit_ns",
    "write_ns",
    "read_ns",
    "output_bytes",
    "output_sha256",
    "batch",
)
NUMERIC_COLUMNS = (
    "nanoseconds",
    "edit_ns",
    "set_ns",
    "commit_ns",
    "write_ns",
    "read_ns",
    "output_bytes",
)


class RunnerError(RuntimeError):
    """A fail-closed runner, identity, child, or CSV protocol error."""


def canonical_json(value: Any) -> bytes:
    return json.dumps(
        value,
        sort_keys=True,
        separators=(",", ":"),
        ensure_ascii=True,
        allow_nan=False,
    ).encode("utf-8")


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> dict[str, Any]:
    resolved = path.expanduser().resolve(strict=True)
    if not resolved.is_file():
        raise RunnerError(f"path is not a regular file: {resolved}")
    digest = hashlib.sha256()
    size = 0
    with resolved.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
            size += len(chunk)
    return {"path": str(resolved), "bytes": size, "sha256": digest.hexdigest()}


def run_capture(command: list[str]) -> str:
    try:
        completed = subprocess.run(
            command,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            encoding="utf-8",
            errors="strict",
        )
    except (OSError, subprocess.CalledProcessError, UnicodeError) as error:
        raise RunnerError(f"command failed: {' '.join(command)}: {error}") from error
    return completed.stdout.strip()


def hash_git_diff(source: Path) -> str:
    try:
        completed = subprocess.run(
            ["git", "-C", str(source), "diff", "--no-ext-diff", "--binary", "HEAD", "--"],
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        raise RunnerError(f"cannot hash git diff for {source}: {error}") from error
    return sha256_bytes(completed.stdout)


def require_ancestor(source: Path, ancestor: str) -> None:
    try:
        completed = subprocess.run(
            ["git", "-C", str(source), "merge-base", "--is-ancestor", ancestor, "HEAD"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.PIPE,
            check=False,
        )
    except OSError as error:
        raise RunnerError(f"cannot check source lineage for {source}: {error}") from error
    if completed.returncode != 0:
        detail = completed.stderr.decode("utf-8", "replace").strip()
        raise RunnerError(
            f"{source} does not descend from pinned lineage {ancestor}"
            + (f": {detail}" if detail else "")
        )


def probe_identities(repo: Path) -> dict[str, str]:
    identities: dict[str, str] = {}
    for relative in PROBE_PATHS:
        identities[relative] = sha256_file(repo / relative)["sha256"]
    return identities


def source_identity(
    path: Path,
    *,
    lineage: str,
    expected_probe: dict[str, str],
) -> dict[str, Any]:
    source = path.expanduser().resolve(strict=True)
    if not source.is_dir():
        raise RunnerError(f"source snapshot is not a directory: {source}")
    revision = run_capture(["git", "-C", str(source), "rev-parse", "HEAD"])
    tree = run_capture(["git", "-C", str(source), "rev-parse", "HEAD^{tree}"])
    status = run_capture(
        ["git", "-C", str(source), "status", "--porcelain=v1", "--untracked-files=all"]
    )
    if status:
        raise RunnerError(f"source snapshot is dirty: {source}: {status}")
    require_ancestor(source, lineage)
    actual_probe = probe_identities(source)
    if actual_probe != expected_probe:
        mismatches = [
            relative
            for relative in PROBE_PATHS
            if actual_probe.get(relative) != expected_probe.get(relative)
        ]
        raise RunnerError(
            f"{source} has probe identity mismatches: {', '.join(mismatches)}"
        )
    identity = {
        "path": str(source),
        "revision": revision,
        "tree": tree,
        "lineage": lineage,
        "worktree_clean": True,
        "status_sha256": sha256_bytes(status.encode("utf-8")),
        "diff_sha256": hash_git_diff(source),
        "probe_commits": list(PROBE_COMMITS),
        "probe_sha256": actual_probe,
    }
    identity["source_sha256"] = sha256_bytes(canonical_json(identity))
    return identity


def host_identity() -> dict[str, Any]:
    uname = platform.uname()._asdict()
    affinity = sorted(os.sched_getaffinity(0))
    cpuinfo = Path("/proc/cpuinfo")
    meminfo = Path("/proc/meminfo")
    facts = {
        "uname": uname,
        "python": sys.version,
        "cpu_count": os.cpu_count(),
        "affinity": affinity,
        "cpuinfo_sha256": sha256_file(cpuinfo)["sha256"] if cpuinfo.is_file() else None,
        "meminfo_sha256": sha256_file(meminfo)["sha256"] if meminfo.is_file() else None,
    }
    facts["host_sha256"] = sha256_bytes(canonical_json(facts))
    return facts


def parse_cpu_list(value: str) -> set[int]:
    result: set[int] = set()
    for part in value.strip().split(","):
        if not part:
            continue
        if "-" in part:
            left, right = part.split("-", 1)
            start, end = int(left), int(right)
            if start > end:
                raise ValueError(value)
            result.update(range(start, end + 1))
        else:
            result.add(int(part))
    if not result:
        raise ValueError(value)
    return result


def process_scan(cpu: int) -> dict[str, Any]:
    proc = Path("/proc")
    load_path = proc / "loadavg"
    if not proc.is_dir() or not load_path.is_file():
        raise RunnerError("quiet process/load scans require Linux /proc")
    try:
        load_values = load_path.read_text(encoding="ascii").split()[:3]
        loadavg = [float(value) for value in load_values]
    except (OSError, UnicodeError, ValueError) as error:
        raise RunnerError(f"cannot read /proc/loadavg: {error}") from error
    processes: list[dict[str, Any]] = []
    for entry in proc.iterdir():
        if not entry.name.isdigit():
            continue
        try:
            status = (entry / "status").read_text(encoding="ascii")
            comm = (entry / "comm").read_text(encoding="utf-8").strip()
        except (FileNotFoundError, PermissionError, OSError, UnicodeError):
            continue
        allowed_text = next(
            (
                line.split(":", 1)[1].strip()
                for line in status.splitlines()
                if line.startswith("Cpus_allowed_list:")
            ),
            None,
        )
        if allowed_text is None:
            continue
        try:
            allowed = sorted(parse_cpu_list(allowed_text))
        except ValueError:
            continue
        processes.append({"pid": int(entry.name), "comm": comm, "affinity": allowed})
    processes.sort(key=lambda item: item["pid"])
    scan = {"cpu": cpu, "loadavg": loadavg, "processes": processes}
    scan["scan_sha256"] = sha256_bytes(canonical_json(scan))
    return scan


def child_scope(pid: int) -> dict[str, Any]:
    status_path = Path("/proc") / str(pid) / "status"
    limits_path = Path("/proc") / str(pid) / "limits"
    try:
        status = status_path.read_text(encoding="ascii")
        limits = limits_path.read_text(encoding="ascii")
    except (FileNotFoundError, OSError, UnicodeError) as error:
        raise RunnerError(f"cannot inspect child scope for PID {pid}: {error}") from error
    allowed_text = next(
        (
            line.split(":", 1)[1].strip()
            for line in status.splitlines()
            if line.startswith("Cpus_allowed_list:")
        ),
        None,
    )
    if allowed_text is None:
        raise RunnerError(f"child PID {pid} did not report Cpus_allowed_list")
    try:
        affinity = sorted(parse_cpu_list(allowed_text))
    except ValueError as error:
        raise RunnerError(f"child PID {pid} reported invalid affinity {allowed_text!r}") from error
    max_address = None
    for line in limits.splitlines():
        if line.startswith("Max address space"):
            fields = line.split()
            if len(fields) >= 3 and fields[-1] == "bytes":
                try:
                    max_address = int(fields[-2])
                except ValueError:
                    pass
            break
    if max_address is None:
        raise RunnerError(f"child PID {pid} did not report a finite Max address space")
    return {"pid": pid, "affinity": affinity, "max_address_space_bytes": max_address}


def child_preexec(cpu: int, memory_limit_bytes: int):
    def apply_scope() -> None:
        os.sched_setaffinity(0, {cpu})
        resource.setrlimit(resource.RLIMIT_AS, (memory_limit_bytes, memory_limit_bytes))

    return apply_scope


def terminate_child(process: subprocess.Popen[bytes]) -> None:
    try:
        if process.poll() is None:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except (AttributeError, OSError):
                process.kill()
        process.wait(timeout=5.0)
    except (AttributeError, OSError, subprocess.SubprocessError, subprocess.TimeoutExpired):
        pass


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--control-binary", type=Path, required=True)
    parser.add_argument("--candidate-binary", type=Path, required=True)
    parser.add_argument("--control-source", type=Path, required=True)
    parser.add_argument("--candidate-source", type=Path, required=True)
    parser.add_argument("--control-lineage", default=BASELINE_LINEAGE)
    parser.add_argument("--candidate-lineage", default=CANDIDATE_LINEAGE)
    parser.add_argument("--output-dir", type=Path, required=True)
    parser.add_argument("--cpu", type=int, required=True)
    parser.add_argument("--memory-limit-bytes", type=int, default=DEFAULT_MEMORY_LIMIT_BYTES)
    parser.add_argument("--warmup", type=int, default=20)
    parser.add_argument("--samples", type=int, default=40)
    parser.add_argument("--calibration-start-batch", type=int, default=DEFAULT_CALIBRATION_START_BATCH)
    parser.add_argument("--calibration-max-batch", type=int, default=DEFAULT_CALIBRATION_MAX_BATCH)
    parser.add_argument("--timeout-seconds", type=float, default=3600.0)
    parser.add_argument("--route", choices=ROUTES, action="append", dest="routes")
    parser.add_argument("--corpus", choices=CORPORA, action="append", dest="corpora")
    parser.add_argument(
        "--diagnostic",
        action="store_true",
        help="permit a subset; the receipt is explicitly not qualification-eligible",
    )
    return parser.parse_args()


def validate_options(options: argparse.Namespace) -> tuple[list[str], list[str]]:
    if os.name != "posix" or not hasattr(os, "sched_setaffinity"):
        raise RunnerError("qualification requires POSIX CPU affinity")
    if not hasattr(resource, "RLIMIT_AS"):
        raise RunnerError("qualification requires RLIMIT_AS")
    if options.cpu < 0 or options.cpu not in os.sched_getaffinity(0):
        raise RunnerError(f"CPU {options.cpu} is not in the driver's affinity")
    if options.memory_limit_bytes <= 0 or options.memory_limit_bytes > HARD_MEMORY_LIMIT_BYTES:
        raise RunnerError(f"memory limit must be in 1..{HARD_MEMORY_LIMIT_BYTES}")
    hard = resource.getrlimit(resource.RLIMIT_AS)[1]
    infinity = getattr(resource, "RLIM_INFINITY", -1)
    if hard != infinity and hard < options.memory_limit_bytes:
        raise RunnerError("parent RLIMIT_AS hard limit is below the child memory scope")
    if options.warmup < 0 or options.samples <= 0:
        raise RunnerError("warmup must be nonnegative and samples must be positive")
    if options.calibration_start_batch <= 0 or options.calibration_max_batch < options.calibration_start_batch:
        raise RunnerError("calibration batch bounds are invalid")
    if options.timeout_seconds <= 0:
        raise RunnerError("timeout must be positive")
    control_path = options.control_binary.expanduser().resolve(strict=True)
    candidate_path = options.candidate_binary.expanduser().resolve(strict=True)
    if control_path == candidate_path:
        raise RunnerError("control and candidate binaries must have distinct paths")
    routes = options.routes or list(ROUTES)
    corpora = options.corpora or list(CORPORA)
    if len(routes) != len(set(routes)) or len(corpora) != len(set(corpora)):
        raise RunnerError("duplicate route or corpus selectors are not permitted")
    if not routes or not corpora:
        raise RunnerError("the matrix must contain at least one route and corpus")
    complete = tuple(routes) == ROUTES and tuple(corpora) == CORPORA
    if not complete and not options.diagnostic:
        raise RunnerError(
            "only the complete 3-route x 2-corpus matrix qualifies; use --diagnostic for subsets"
        )
    return routes, corpora


def parse_probe_output(
    stdout_path: Path,
    route: str,
    corpus: str,
    samples: int,
    batch: int,
) -> dict[str, Any]:
    raw = stdout_path.read_bytes()
    if len(raw) > MAX_CHILD_OUTPUT_BYTES:
        raise RunnerError(f"{stdout_path} exceeds the child output limit")
    try:
        text = raw.decode("utf-8")
        reader = csv.DictReader(
            (line for line in text.splitlines() if line.strip()), strict=True
        )
        fieldnames = reader.fieldnames
        if fieldnames != list(REQUIRED_COLUMNS):
            raise RunnerError(f"{stdout_path} has unexpected CSV columns: {fieldnames!r}")
        if len(fieldnames) != len(set(fieldnames)):
            raise RunnerError(f"{stdout_path} has duplicate CSV columns")
        rows = list(reader)
    except UnicodeDecodeError as error:
        raise RunnerError(f"{stdout_path} is not UTF-8: {error}") from error
    except csv.Error as error:
        raise RunnerError(f"{stdout_path} is malformed CSV: {error}") from error
    if not rows:
        raise RunnerError(f"{stdout_path} has no CSV rows")
    if len(rows) != samples:
        raise RunnerError(f"{stdout_path} has {len(rows)} rows, expected {samples}")
    digests: set[str] = set()
    output_bytes: set[int] = set()
    parsed_samples: list[int] = []
    total_ns = 0
    for row in rows:
        if len(row) != len(REQUIRED_COLUMNS) or set(row) != set(REQUIRED_COLUMNS):
            raise RunnerError(f"{stdout_path} contains an extra or missing CSV field: {row!r}")
        if any(key is None or value is None or value == "" for key, value in row.items()):
            raise RunnerError(f"{stdout_path} contains a null or empty CSV field")
        if row["route"] != route or row["corpus"] != corpus:
            raise RunnerError(f"{stdout_path} contains a route/corpus mismatch")
        try:
            sample = int(row["sample"])
            numeric = {
                name: int(row[name])
                for name in NUMERIC_COLUMNS
            }
            reported_batch = int(row["batch"])
        except ValueError as error:
            raise RunnerError(f"{stdout_path} contains a non-integer CSV value") from error
        if sample < 0 or any(value < 0 for value in numeric.values()) or reported_batch <= 0:
            raise RunnerError(f"{stdout_path} contains a negative or zero measurement")
        if reported_batch != batch:
            raise RunnerError(f"{stdout_path} reports batch {reported_batch}, expected {batch}")
        digest = row["output_sha256"]
        if len(digest) != HEX_DIGEST_LENGTH or any(
            character not in "0123456789abcdef" for character in digest
        ):
            raise RunnerError(f"{stdout_path} contains an invalid output digest")
        parsed_samples.append(sample)
        digests.add(digest)
        output_bytes.add(numeric["output_bytes"])
        total_ns += numeric["nanoseconds"]
    if parsed_samples != list(range(samples)):
        raise RunnerError(f"{stdout_path} sample indexes are not contiguous")
    if len(digests) != 1 or len(output_bytes) != 1:
        raise RunnerError(f"{stdout_path} output identity is not stable")
    return {
        "rows": len(rows),
        "elapsed_ns": total_ns,
        "output_sha256": next(iter(digests)),
        "output_bytes": next(iter(output_bytes)),
        "stdout_sha256": sha256_bytes(raw),
    }


def run_leg(
    *,
    role: str,
    binary: Path,
    route: str,
    corpus: str,
    warmup: int,
    samples: int,
    batch: int,
    options: argparse.Namespace,
    output_dir: Path,
    minimum_ns: int,
    phase: str,
) -> dict[str, Any]:
    stem = f"{phase}-{route}-{corpus}-{role}"
    stdout_path = output_dir / f"{stem}.stdout.csv"
    stderr_path = output_dir / f"{stem}.stderr.log"
    command = [
        str(binary.resolve()),
        "--route",
        route,
        "--corpus",
        corpus,
        "--warmup",
        str(warmup),
        "--samples",
        str(samples),
        "--batch",
        str(batch),
    ]
    before_scan = process_scan(options.cpu)
    process: subprocess.Popen[bytes] | None = None
    observed_scope: dict[str, Any] | None = None
    after_scan: dict[str, Any] | None = None
    try:
        with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
            try:
                process = subprocess.Popen(
                    command,
                    stdin=subprocess.DEVNULL,
                    stdout=stdout,
                    stderr=stderr,
                    cwd=str(binary.resolve().parent),
                    env=os.environ.copy(),
                    close_fds=True,
                    start_new_session=True,
                    preexec_fn=child_preexec(options.cpu, options.memory_limit_bytes),
                )
            except (OSError, subprocess.SubprocessError) as error:
                raise RunnerError(f"{role} child failed to launch: {error}") from error
            observation_deadline = time.monotonic() + min(2.0, options.timeout_seconds)
            while observed_scope is None:
                if process.poll() is not None:
                    raise RunnerError(f"{role} child exited before scope verification")
                try:
                    observed_scope = child_scope(process.pid)
                except RunnerError:
                    if time.monotonic() >= observation_deadline:
                        raise
                    time.sleep(0.01)
            if observed_scope["affinity"] != [options.cpu]:
                raise RunnerError(
                    f"{role} child affinity was {observed_scope['affinity']}, expected {[options.cpu]}"
                )
            if observed_scope["max_address_space_bytes"] != options.memory_limit_bytes:
                raise RunnerError(
                    f"{role} child memory scope was {observed_scope['max_address_space_bytes']}, "
                    f"expected {options.memory_limit_bytes}"
                )
            try:
                returncode = process.wait(timeout=options.timeout_seconds)
            except subprocess.TimeoutExpired as error:
                terminate_child(process)
                raise RunnerError(f"{role} child timed out") from error
    except BaseException:
        if process is not None:
            terminate_child(process)
        raise
    finally:
        if process is not None and process.poll() is None:
            terminate_child(process)
        after_scan = process_scan(options.cpu)
    if returncode != 0:
        raise RunnerError(f"{role} child returned {returncode}: {stderr_path}")
    if stdout_path.stat().st_size > MAX_CHILD_OUTPUT_BYTES:
        raise RunnerError(f"{role} stdout exceeds the child output limit")
    parsed = parse_probe_output(stdout_path, route, corpus, samples, batch)
    if parsed["elapsed_ns"] < minimum_ns:
        raise RunnerError(
            f"{role} {route}/{corpus} leg was {parsed['elapsed_ns']} ns, below {minimum_ns} ns"
        )
    parsed.update(
        {
            "phase": phase,
            "role": role,
            "route": route,
            "corpus": corpus,
            "command": command,
            "returncode": returncode,
            "stdout": sha256_file(stdout_path),
            "stderr": sha256_file(stderr_path),
            "memory_limit_bytes": options.memory_limit_bytes,
            "child_scope": observed_scope,
            "before_scan": before_scan,
            "after_scan": after_scan,
        }
    )
    return parsed


def calibrate(
    *,
    side: str,
    binary: Path,
    route: str,
    corpus: str,
    options: argparse.Namespace,
    output_dir: Path,
) -> tuple[int, list[dict[str, Any]]]:
    batch = options.calibration_start_batch
    attempts: list[dict[str, Any]] = []
    while batch <= options.calibration_max_batch:
        attempt = run_leg(
            role=f"calibration-{side}-{len(attempts) + 1}",
            binary=binary,
            route=route,
            corpus=corpus,
            warmup=0,
            samples=1,
            batch=batch,
            options=options,
            output_dir=output_dir,
            minimum_ns=0,
            phase="calibration",
        )
        attempts.append(attempt)
        if attempt["elapsed_ns"] >= MIN_LEG_NS:
            return batch, attempts
        if batch > options.calibration_max_batch // 2:
            break
        batch *= 2
    raise RunnerError(
        f"{side} calibration for {route}/{corpus} did not reach {MIN_LEG_NS} ns"
    )


def main() -> int:
    options = parse_args()
    try:
        routes, corpora = validate_options(options)
        control_binary = sha256_file(options.control_binary)
        candidate_binary = sha256_file(options.candidate_binary)
        if control_binary["sha256"] == candidate_binary["sha256"]:
            raise RunnerError("control and candidate binaries have the same SHA-256")
        repo = Path(__file__).resolve().parent.parent
        expected_probe = probe_identities(repo)
        control_source = source_identity(
            options.control_source,
            lineage=options.control_lineage,
            expected_probe=expected_probe,
        )
        candidate_source = source_identity(
            options.candidate_source,
            lineage=options.candidate_lineage,
            expected_probe=expected_probe,
        )
        if control_source["path"] == candidate_source["path"]:
            raise RunnerError("control and candidate sources must have distinct paths")
        if control_source["revision"] == candidate_source["revision"]:
            raise RunnerError("control and candidate sources must have distinct revisions")
        if control_source["tree"] == candidate_source["tree"]:
            raise RunnerError("control and candidate sources must have distinct trees")
        driver = sha256_file(Path(__file__))
        host = host_identity()
        output_dir = options.output_dir.expanduser().resolve()
        if output_dir.exists() and any(output_dir.iterdir()):
            raise RunnerError(f"output directory is not empty: {output_dir}")
        output_dir.mkdir(parents=True, exist_ok=True)
        legs: list[dict[str, Any]] = []
        calibrations: dict[str, Any] = {}
        expected_outputs: dict[str, str] = {}
        for route in routes:
            for corpus in corpora:
                key = f"{route}/{corpus}"
                control_batch, control_attempts = calibrate(
                    side="control",
                    binary=options.control_binary,
                    route=route,
                    corpus=corpus,
                    options=options,
                    output_dir=output_dir,
                )
                candidate_batch, candidate_attempts = calibrate(
                    side="candidate",
                    binary=options.candidate_binary,
                    route=route,
                    corpus=corpus,
                    options=options,
                    output_dir=output_dir,
                )
                common_batch = max(control_batch, candidate_batch)
                calibrations[key] = {
                    "control_admitted_batch": control_batch,
                    "candidate_admitted_batch": candidate_batch,
                    "common_batch": common_batch,
                    "control": control_attempts,
                    "candidate": candidate_attempts,
                }
                binaries = {
                    "A1": options.control_binary,
                    "B1": options.candidate_binary,
                    "B2": options.candidate_binary,
                    "A2": options.control_binary,
                }
                cell_legs: list[dict[str, Any]] = []
                for role in ROLES:
                    leg = run_leg(
                        role=role,
                        binary=binaries[role],
                        route=route,
                        corpus=corpus,
                        warmup=options.warmup,
                        samples=options.samples,
                        batch=common_batch,
                        options=options,
                        output_dir=output_dir,
                        minimum_ns=MIN_LEG_NS,
                        phase="measurement",
                    )
                    legs.append(leg)
                    cell_legs.append(leg)
                selected = [leg["output_sha256"] for leg in cell_legs]
                if len(set(selected)) != 1:
                    raise RunnerError(f"{key} output digest differs across ABBA legs")
                expected_outputs[key] = selected[0]
        complete = tuple(routes) == ROUTES and tuple(corpora) == CORPORA
        receipt = {
            "schema": SCHEMA,
            "created_utc": datetime.now(timezone.utc).isoformat(),
            "qualification_eligible": bool(complete and not options.diagnostic),
            "diagnostic": options.diagnostic,
            "matrix_complete": complete,
            "order": list(ROLES),
            "routes": routes,
            "corpora": corpora,
            "configuration": {
                "warmup": options.warmup,
                "samples": options.samples,
                "minimum_leg_ns": MIN_LEG_NS,
                "calibration_start_batch": options.calibration_start_batch,
                "calibration_max_batch": options.calibration_max_batch,
                "cpu": options.cpu,
                "memory_limit_bytes": options.memory_limit_bytes,
                "timeout_seconds": options.timeout_seconds,
            },
            "sources": {"control": control_source, "candidate": candidate_source},
            "binaries": {"control": control_binary, "candidate": candidate_binary},
            "driver": driver,
            "host": host,
            "calibrations": calibrations,
            "outputs": expected_outputs,
            "legs": legs,
        }
        receipt_path = output_dir / "xls-f3-abba-receipt.json"
        receipt_path.write_bytes(canonical_json(receipt) + b"\n")
        print(receipt_path)
        return 0
    except RunnerError as error:
        print(f"xls F3 ABBA: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
