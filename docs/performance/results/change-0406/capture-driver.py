#!/usr/bin/env python3
"""Capture one pinned, current-state hardware profile for the Litchi goal.

The driver accepts an already-built release harness and an expected Git
revision.  It never invokes Cargo build or any other build command.  The
repository may retain the intentionally untracked ``docs/GOAL.md`` file; that
one path is hash-bound and every other source-tree change is rejected.  All
workload profilers run sequentially against the same
opc_source_materialize/few-large/incompressible selector.  Raw reports and
profiler outputs are written below a temporary directory first; the final
0406 directory is published only after report validation succeeds.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import platform
import re
import shlex
import shutil
import stat
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Sequence


CASE = "opc_source_materialize"
SHAPE = "few-large"
PAYLOAD = "incompressible"
CHANGE = "0406"
RUSTUP_TOOLCHAIN = "1.98.1"
ALLOWED_UNTRACKED_PATH = "docs/GOAL.md"
DEFAULT_REPO = Path("/home/zhuhe/code/litchi")
DEFAULT_BINARY = DEFAULT_REPO / "tools/perf-baseline/target/release/litchi-perf-baseline"
DEFAULT_TEMP = Path("/tmp/litchi-goal-profile-output")
DEFAULT_FINAL = DEFAULT_REPO / "docs/performance/results/change-0406"
EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-misses",
    "page-faults",
)
REQUIRED_TOOLS = (
    "taskset",
    "perf",
    "heaptrack",
    "heaptrack_print",
    "strace",
    "fincore",
    "gzip",
    "zstd",
    "readelf",
)


class CaptureError(RuntimeError):
    """Raised when the pinned capture cannot be safely completed."""


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def json_dump(path: Path, value: Any) -> None:
    path.write_text(
        json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n",
        encoding="utf-8",
    )


def reject_nonfinite(value: str) -> None:
    raise ValueError(f"non-finite JSON constant {value!r}")


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"), parse_constant=reject_nonfinite)
    except (OSError, UnicodeError, json.JSONDecodeError, ValueError) as error:
        raise CaptureError(f"cannot read JSON report {path}: {error}") from error
    if not isinstance(value, dict):
        raise CaptureError(f"JSON report {path} must contain an object")
    return value


def require_tool(name: str) -> str:
    path = shutil.which(name)
    if path is None:
        raise CaptureError(f"required tool is unavailable: {name}")
    return str(Path(path).resolve())


class Capture:
    def __init__(self, arguments: argparse.Namespace) -> None:
        self.repo = arguments.repo.resolve()
        self.binary = arguments.binary.resolve()
        self.expected_revision = arguments.revision
        self.cpu = str(arguments.cpu)
        self.warmup = arguments.warmup
        self.samples = arguments.samples
        self.heaptrack_samples = arguments.heaptrack_samples
        self.temp = arguments.temp.resolve()
        self.final = arguments.final.resolve()
        self.skip_strace = arguments.skip_strace
        self.commands: list[dict[str, Any]] = []
        self.tools: dict[str, str] = {}
        self.start_time = utc_now()
        self.allowed_untracked: dict[str, Any] | None = None
        self.source_before: dict[str, Any] | None = None

    def rel(self, path: Path) -> str:
        try:
            return str(path.resolve().relative_to(self.temp))
        except ValueError:
            return str(path)

    def command_text(self, argv: Sequence[str]) -> str:
        return shlex.join(str(item) for item in argv)

    def prune_compressed_intermediate(
        self,
        raw_name: str,
        compressed_name: str,
        command_label: str,
        *,
        stdout_pruned: bool = True,
    ) -> None:
        raw = self.temp / raw_name
        if raw.is_file():
            raw.unlink()
        for command in reversed(self.commands):
            if command.get("label") == command_label:
                if stdout_pruned:
                    command["stdout_pruned_after_compression"] = True
                else:
                    command["derived_artifact_pruned_after_compression"] = raw_name
                command["compressed_artifact"] = compressed_name
                return

    def run(
        self,
        label: str,
        argv: Sequence[str],
        *,
        stdout_name: str | None = None,
        stderr_name: str | None = None,
        check: bool = True,
        cwd: Path | None = None,
        env: dict[str, str] | None = None,
    ) -> dict[str, Any]:
        command = [str(item) for item in argv]
        stdout_path = self.temp / stdout_name if stdout_name else None
        stderr_path = self.temp / stderr_name if stderr_name else None
        stdout_handle = stdout_path.open("wb") if stdout_path else subprocess.PIPE
        stderr_handle = stderr_path.open("wb") if stderr_path else subprocess.PIPE
        started_wall = utc_now()
        started_ns = time.monotonic_ns()
        command_environment = os.environ.copy()
        command_environment["RUSTUP_TOOLCHAIN"] = RUSTUP_TOOLCHAIN
        if env is not None:
            command_environment.update(env)
        try:
            completed = subprocess.run(
                command,
                cwd=str(cwd or self.repo),
                stdin=subprocess.DEVNULL,
                stdout=stdout_handle,
                stderr=stderr_handle,
                env=command_environment,
                check=False,
            )
        finally:
            if stdout_path:
                stdout_handle.close()
            if stderr_path:
                stderr_handle.close()
        elapsed_ns = time.monotonic_ns() - started_ns
        record: dict[str, Any] = {
            "label": label,
            "argv": command,
            "command": self.command_text(command),
            "cwd": str(cwd or self.repo),
            "started_utc": started_wall,
            "finished_utc": utc_now(),
            "wall_ns": elapsed_ns,
            "returncode": completed.returncode,
            "stdout": self.rel(stdout_path) if stdout_path else None,
            "stderr": self.rel(stderr_path) if stderr_path else None,
            "checked": check,
            "environment_overrides": {"RUSTUP_TOOLCHAIN": RUSTUP_TOOLCHAIN},
        }
        self.commands.append(record)
        if check and completed.returncode != 0:
            raise CaptureError(
                f"{label} failed with exit {completed.returncode}: "
                f"{self.command_text(command)}"
            )
        return record

    def read_command_output(self, record: dict[str, Any], stream: str) -> str:
        name = record.get(stream)
        if not isinstance(name, str):
            return ""
        return (self.temp / name).read_text(encoding="utf-8", errors="replace")

    def git(self, *args: str, check: bool = True) -> str:
        command = ["git", "-C", str(self.repo), *args]
        started_wall = utc_now()
        started_ns = time.monotonic_ns()
        result = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
        )
        elapsed_ns = time.monotonic_ns() - started_ns
        self.commands.append(
            {
                "label": "git-" + (args[0] if args else "status"),
                "argv": command,
                "command": self.command_text(command),
                "cwd": str(self.repo),
                "started_utc": started_wall,
                "finished_utc": utc_now(),
                "wall_ns": elapsed_ns,
                "returncode": result.returncode,
                "stdout": None,
                "stderr": None,
                "stdout_bytes": len(result.stdout.encode()),
                "stdout_sha256": sha256_bytes(result.stdout.encode()),
                "stderr_bytes": len(result.stderr.encode()),
                "stderr_sha256": sha256_bytes(result.stderr.encode()),
                "checked": check,
                "environment_overrides": {},
            }
        )
        if check and result.returncode != 0:
            raise CaptureError(
                f"git {' '.join(args)} failed with exit {result.returncode}: "
                f"{result.stderr.strip()}"
            )
        return result.stdout

    def git_state(self) -> dict[str, Any]:
        status = self.git("status", "--porcelain=v1")
        status_z = self.git("status", "--porcelain=v1", "-z")
        entries: list[dict[str, str]] = []
        for line in status.splitlines():
            if not line:
                continue
            entries.append({"code": line[:2], "path": line[3:]})
        allowed = [
            entry
            for entry in entries
            if entry["code"] == "??" and entry["path"] == ALLOWED_UNTRACKED_PATH
        ]
        unexpected = [entry for entry in entries if entry not in allowed]
        allowed_identity: dict[str, Any] | None = None
        if not unexpected and len(allowed) <= 1 and allowed:
            goal = self.repo / ALLOWED_UNTRACKED_PATH
            if not goal.is_file():
                raise CaptureError(f"status lists missing {ALLOWED_UNTRACKED_PATH}")
            allowed_identity = {
                "path": ALLOWED_UNTRACKED_PATH,
                "bytes": goal.stat().st_size,
                "sha256": sha256_file(goal),
            }
        return {
            "head": self.git("rev-parse", "--verify", "HEAD").strip(),
            "tree": self.git("rev-parse", "--verify", "HEAD^{tree}").strip(),
            "status_porcelain": status,
            "status_entries": entries,
            "status_sha256": sha256_bytes(status.encode()),
            "status_z_sha256": sha256_bytes(status_z.encode()),
            "dirty": bool(status),
            "allowed_untracked": allowed_identity,
        }

    def assert_source_binding(self) -> dict[str, Any]:
        state = self.git_state()
        if state["head"] != self.expected_revision:
            raise CaptureError(
                f"HEAD {state['head']} does not match expected revision "
                f"{self.expected_revision}"
            )
        if state["status_entries"] and state["allowed_untracked"] is None:
            raise CaptureError(
                "source tree has tracked changes or an unexpected untracked path; "
                f"only {ALLOWED_UNTRACKED_PATH} may remain untracked"
            )
        if self.allowed_untracked is None:
            self.allowed_untracked = state["allowed_untracked"]
        elif self.allowed_untracked != state["allowed_untracked"]:
            raise CaptureError(
                f"{ALLOWED_UNTRACKED_PATH} changed during capture; refusing mixed source evidence"
            )
        return state

    def harness_args(self, json_name: str, *, samples: int | None = None) -> list[str]:
        return [
            str(self.binary),
            "--warmup",
            str(self.warmup),
            "--samples",
            str(self.samples if samples is None else samples),
            "--case",
            CASE,
            "--shape",
            SHAPE,
            "--payload",
            PAYLOAD,
            "--json",
            str(self.temp / json_name),
        ]

    def validate_report(
        self,
        name: str,
        *,
        expected_samples: int,
        expected_warmup: int,
    ) -> dict[str, Any]:
        path = self.temp / name
        report = load_json(path)
        try:
            sys.path.insert(0, str(self.repo))
            from tools import perf_compare

            perf_compare.validate_parallel_metrics(report, name)
        except CaptureError:
            raise
        except Exception as error:
            raise CaptureError(f"existing validator rejected {name}: {error}") from error

        if report.get("schema_version") != 1:
            raise CaptureError(f"{name}.schema_version is not 1")
        tool = report.get("tool")
        if (
            not isinstance(tool, dict)
            or tool.get("name") != "litchi-perf-baseline"
            or tool.get("binary") != "litchi-perf-baseline"
            or tool.get("profile") != "release"
        ):
            raise CaptureError(f"{name}.tool.name is not litchi-perf-baseline")
        binary_identity = report.get("binary_identity")
        if not isinstance(binary_identity, dict):
            raise CaptureError(f"{name}.binary_identity is missing")
        if binary_identity.get("binary_sha256") != self.binary_sha256:
            raise CaptureError(f"{name}.binary_identity hash does not match captured binary")
        if binary_identity.get("binary_bytes") != self.binary_bytes:
            raise CaptureError(f"{name}.binary_identity size does not match captured binary")
        environment = report.get("environment")
        if not isinstance(environment, dict):
            raise CaptureError(f"{name}.environment is missing")
        if environment.get("git_revision") != self.expected_revision:
            raise CaptureError(f"{name}.environment.git_revision does not match expected revision")
        if self.source_before is None:
            raise CaptureError("source binding was not captured before report validation")
        if environment.get("git_worktree_dirty") != self.source_before["dirty"]:
            raise CaptureError(
                f"{name}.environment.git_worktree_dirty does not match source binding"
            )
        if RUSTUP_TOOLCHAIN not in str(environment.get("rustc_version", "")):
            raise CaptureError(
                f"{name}.environment.rustc_version does not identify Rust {RUSTUP_TOOLCHAIN}"
            )

        configuration = report.get("configuration")
        if not isinstance(configuration, dict):
            raise CaptureError(f"{name}.configuration is missing")
        exact_fields = {
            "samples_per_case": expected_samples,
            "warmup_iterations_per_case": expected_warmup,
            "cases": [CASE],
            "corpus_shapes": [SHAPE],
            "payload_kinds": [PAYLOAD],
        }
        for field, expected in exact_fields.items():
            if configuration.get(field) != expected:
                raise CaptureError(
                    f"{name}.configuration.{field}={configuration.get(field)!r}; "
                    f"expected {expected!r}"
                )
        results = report.get("results")
        if not isinstance(results, list) or len(results) != 1:
            raise CaptureError(f"{name}.results must contain exactly one result")
        result = results[0]
        if not isinstance(result, dict) or result.get("case") != CASE:
            raise CaptureError(f"{name}.results[0].case is not {CASE}")
        corpus = result.get("corpus")
        if not isinstance(corpus, dict):
            raise CaptureError(f"{name}.results[0].corpus is missing")
        for field, expected in (
            ("shape", SHAPE),
            ("payload_kind", PAYLOAD),
            ("entry_count", 4),
            ("entry_bytes", 4 * 1024 * 1024),
            ("uncompressed_payload_bytes", 16 * 1024 * 1024),
        ):
            if corpus.get(field) != expected:
                raise CaptureError(
                    f"{name}.results[0].corpus.{field}={corpus.get(field)!r}; "
                    f"expected {expected!r}"
                )
        elapsed = result.get("elapsed_ns")
        if not isinstance(elapsed, dict) or not isinstance(elapsed.get("samples"), list):
            raise CaptureError(f"{name}.results[0].elapsed_ns.samples is missing")
        if len(elapsed["samples"]) != expected_samples:
            raise CaptureError(
                f"{name}.results[0].elapsed_ns.samples has {len(elapsed['samples'])}; "
                f"expected {expected_samples}"
            )
        return report

    def collect_versions(self) -> dict[str, Any]:
        versions: dict[str, Any] = {}
        commands = {
            "rustc": ["rustc", f"+{RUSTUP_TOOLCHAIN}", "--version", "--verbose"],
            "cargo": ["cargo", f"+{RUSTUP_TOOLCHAIN}", "--version", "--verbose"],
            "perf": [self.tools["perf"], "--version"],
            "heaptrack": [self.tools["heaptrack"], "--version"],
            "heaptrack_print": [self.tools["heaptrack_print"], "--version"],
            "strace": [self.tools["strace"], "--version"],
            "fincore": [self.tools["fincore"], "--version"],
            "taskset": [self.tools["taskset"], "--version"],
            "time": ["/usr/bin/time", "--version"],
            "gzip": [self.tools["gzip"], "--version"],
            "zstd": [self.tools["zstd"], "--version"],
            "python3": [sys.executable, "--version"],
        }
        for name, command in commands.items():
            record = self.run(
                f"version-{name}",
                command,
                stdout_name=f"versions/{name}.stdout.txt",
                stderr_name=f"versions/{name}.stderr.txt",
            )
            versions[name] = {
                "stdout": self.read_command_output(record, "stdout"),
                "stderr": self.read_command_output(record, "stderr"),
                "command_index": len(self.commands) - 1,
            }
        return versions

    def initialize(self) -> dict[str, Any]:
        if self.warmup < 0 or self.samples < 1 or self.heaptrack_samples < 1:
            raise CaptureError("warmup must be non-negative and sample counts positive")
        if not self.repo.is_dir():
            raise CaptureError(f"repository does not exist: {self.repo}")
        if not self.binary.is_file() or not os.access(self.binary, os.X_OK):
            raise CaptureError(f"binary is not an executable file: {self.binary}")
        if self.temp.exists():
            raise CaptureError(
                f"temporary output already exists: {self.temp}; remove it only after "
                "checking that no capture is running"
            )
        if self.final.exists():
            raise CaptureError(
                f"final output already exists: {self.final}; choose a fresh 0406 path"
            )
        self.temp.mkdir(parents=True)
        (self.temp / "versions").mkdir()
        driver_copy = self.temp / "capture-driver.py"
        shutil.copy2(Path(__file__).resolve(), driver_copy)
        self.driver_artifact = {
            "path": driver_copy.name,
            "bytes": driver_copy.stat().st_size,
            "sha256": sha256_file(driver_copy),
        }
        self.tools = {name: require_tool(name) for name in REQUIRED_TOOLS}
        self.tools["time"] = "/usr/bin/time"
        self.binary_sha256 = sha256_file(self.binary)
        self.binary_bytes = self.binary.stat().st_size
        self.binary_mode = stat.S_IMODE(self.binary.stat().st_mode)
        source = self.assert_source_binding()
        self.source_before = source
        self.run(
            "uname",
            ["uname", "-a"],
            stdout_name="uname.txt",
            stderr_name="uname.stderr.txt",
        )
        self.run(
            "lscpu",
            ["lscpu"],
            stdout_name="lscpu.txt",
            stderr_name="lscpu.stderr.txt",
        )
        self.run(
            "free",
            ["free", "-h"],
            stdout_name="memory-free.txt",
            stderr_name="memory-free.stderr.txt",
        )
        self.run(
            "taskset-self",
            [self.tools["taskset"], "-pc", str(os.getpid())],
            stdout_name="taskset-self.txt",
            stderr_name="taskset-self.stderr.txt",
        )
        self.run(
            "readelf-build-id",
            [self.tools["readelf"], "-n", str(self.binary)],
            stdout_name="binary-readelf.txt",
            stderr_name="binary-readelf.stderr.txt",
        )
        versions = self.collect_versions()
        environment = {
            "captured_utc": self.start_time,
            "repo": str(self.repo),
            "expected_revision": self.expected_revision,
            "source_before": source,
            "binary": {
                "path": str(self.binary),
                "sha256": self.binary_sha256,
                "bytes": self.binary_bytes,
                "mode_bits": self.binary_mode,
            },
            "tools": self.tools,
            "cpu_argument": self.cpu,
            "python_affinity": sorted(os.sched_getaffinity(0)),
            "kernel": platform.release(),
            "rustup_toolchain": RUSTUP_TOOLCHAIN,
            "versions": versions,
            "perf_event_paranoid": Path("/proc/sys/kernel/perf_event_paranoid")
            .read_text(encoding="utf-8")
            .strip(),
            "protocol": {
                "selector": CASE,
                "shape": SHAPE,
                "payload": PAYLOAD,
                "warmup_iterations": self.warmup,
                "samples": self.samples,
                "heaptrack_warmup_iterations": self.warmup,
                "heaptrack_samples": self.heaptrack_samples,
                "profiler_runs_sequential": True,
                "external_scope": (
                    "whole process: executable startup, deterministic corpus construction, "
                    "source-backed open, warmups, measured iterations, verification, and "
                    "profiler instrumentation where applicable"
                ),
                "harness_scope": (
                    "the harness elapsed_ns and operation_metrics rows; "
                    "opc_source_materialize starts its operation timer after source-backed "
                    "open and reset, then verifies each materialized package after timing"
                ),
                "allocation_scope": (
                    "Heaptrack process totals at 20 warmups and 100 measured samples; "
                    "includes source construction/open/verification and Heaptrack overhead; "
                    "not operation-local allocation attribution"
                ),
                "heaptrack_record_mode": (
                    "heaptrack --record-only records and interprets the capture; "
                    "heaptrack 1.5.0 has no separate --interpret command"
                ),
                "heaptrack_symbolized_trace": (
                    "heaptrack_print -F writes the symbolized allocation stacks; "
                    "the uncompressed intermediate is gzip-compressed and pruned"
                ),
                "perf_storage": (
                    "perf.data is retained for standard report reruns; raw perf script "
                    "is gzip-compressed and its uncompressed intermediate is pruned"
                ),
                "claim": (
                    "descriptive current-state baseline only; no speedup, regression, "
                    "throughput, RSS, or operation-local allocation claim"
                ),
                "clean_abba_claim_eligible": False,
                "clean_abba_claim_reason": (
                    f"the run retains intentionally untracked {ALLOWED_UNTRACKED_PATH}"
                ),
            },
        }
        json_dump(self.temp / "environment.json", environment)
        return environment

    def fincore(self, name: str) -> None:
        self.run(
            name,
            [
                self.tools["fincore"],
                "--json",
                "--bytes",
                "--output-all",
                str(self.binary),
            ],
            stdout_name=f"{name}.json",
            stderr_name=f"{name}.stderr.txt",
            check=False,
        )

    def run_harness(self, label: str, json_name: str, *, samples: int | None = None) -> None:
        self.run(
            label,
            [self.tools["taskset"], "-c", self.cpu, *self.harness_args(json_name, samples=samples)],
            stdout_name=f"{label}.stdout.txt",
            stderr_name=f"{label}.stderr.txt",
        )

    def compress_zstd(self, name: str) -> None:
        source = self.temp / name
        destination = self.temp / f"{name}.zst"
        self.run(
            f"zstd-{name}",
            [self.tools["zstd"], "-q", "-19", "-f", str(source), "-o", str(destination)],
            stdout_name=f"zstd-{name}.stdout.txt",
            stderr_name=f"zstd-{name}.stderr.txt",
        )

    def capture(self) -> dict[str, Any]:
        self.fincore("fincore-before")

        self.run_harness("baseline", "baseline.json")

        time_report = self.temp / "time-v.txt"
        self.run(
            "time",
            [
                self.tools["taskset"],
                "-c",
                self.cpu,
                "/usr/bin/time",
                "-v",
                "-o",
                str(time_report),
                "--",
                *self.harness_args("time.json"),
            ],
            stdout_name="time.stdout.txt",
            stderr_name="time.stderr.txt",
        )

        perf_stat_path = self.temp / "perf-stat.csv"
        self.run(
            "perf-stat",
            [
                self.tools["taskset"],
                "-c",
                self.cpu,
                self.tools["perf"],
                "stat",
                "-x,",
                "-o",
                str(perf_stat_path),
                "-e",
                ",".join(EVENTS),
                "--",
                *self.harness_args("perf-stat.json"),
            ],
            stdout_name="perf-stat.stdout.txt",
            stderr_name="perf-stat.stderr.txt",
        )

        perf_data = self.temp / "perf.data"
        self.run(
            "perf-record",
            [
                self.tools["taskset"],
                "-c",
                self.cpu,
                self.tools["perf"],
                "record",
                "-e",
                "cycles:u",
                "-F",
                "99",
                "--call-graph",
                "dwarf",
                "-o",
                str(perf_data),
                "--",
                *self.harness_args("perf-record.json"),
            ],
            stdout_name="perf-record.stdout.txt",
            stderr_name="perf-record.stderr.txt",
        )

        self.run(
            "perf-report-self",
            [
                self.tools["perf"],
                "report",
                "--stdio",
                "--no-children",
                "--call-graph=none",
                "--percent-limit=0.5",
                "-i",
                str(perf_data),
            ],
            stdout_name="perf-report-self.txt",
            stderr_name="perf-report-self.stderr.txt",
        )
        self.run(
            "perf-report-inclusive",
            [
                self.tools["perf"],
                "report",
                "--stdio",
                "--children",
                "--call-graph=none",
                "--percent-limit=0.5",
                "-i",
                str(perf_data),
            ],
            stdout_name="perf-report-inclusive.txt",
            stderr_name="perf-report-inclusive.stderr.txt",
        )
        self.run(
            "perf-script",
            [self.tools["perf"], "script", "-i", str(perf_data)],
            stdout_name="perf-script.txt",
            stderr_name="perf-script.stderr.txt",
        )
        self.run(
            "perf-script-gzip",
            [self.tools["gzip"], "-9", "-c", str(self.temp / "perf-script.txt")],
            stdout_name="perf-script.txt.gz",
            stderr_name="perf-script-gzip.stderr.txt",
        )
        self.prune_compressed_intermediate(
            "perf-script.txt", "perf-script.txt.gz", "perf-script"
        )

        heap_prefix = self.temp / "heaptrack-profile"
        self.run(
            "heaptrack",
            [
                self.tools["taskset"],
                "-c",
                self.cpu,
                self.tools["heaptrack"],
                "--record-only",
                "-o",
                str(heap_prefix),
                *self.harness_args("heaptrack.json", samples=self.heaptrack_samples),
            ],
            stdout_name="heaptrack.stdout.txt",
            stderr_name="heaptrack.stderr.txt",
        )
        captures = sorted(
            path
            for path in self.temp.glob("heaptrack-profile*")
            if path.is_file()
            and path.name not in {"heaptrack.stdout.txt", "heaptrack.stderr.txt"}
        )
        captures = [
            path
            for path in captures
            if path.name not in {"heaptrack-histogram.tsv", "heaptrack-print.txt"}
        ]
        if not captures:
            raise CaptureError("Heaptrack completed without a capture artifact")
        heap_capture = captures[-1]
        (self.temp / "heaptrack-capture-name.txt").write_text(
            heap_capture.name + "\n", encoding="utf-8"
        )
        self.run(
            "heaptrack-print",
            [
                self.tools["heaptrack_print"],
                "-f",
                str(heap_capture),
                "-H",
                str(self.temp / "heaptrack-histogram.tsv"),
                "-F",
                str(self.temp / "heaptrack-flamegraph.txt"),
            ],
            stdout_name="heaptrack-print.txt",
            stderr_name="heaptrack-print.stderr.txt",
        )
        self.run(
            "heaptrack-flamegraph-gzip",
            [self.tools["gzip"], "-9", "-c", str(self.temp / "heaptrack-flamegraph.txt")],
            stdout_name="heaptrack-flamegraph.txt.gz",
            stderr_name="heaptrack-flamegraph-gzip.stderr.txt",
        )
        self.prune_compressed_intermediate(
            "heaptrack-flamegraph.txt",
            "heaptrack-flamegraph.txt.gz",
            "heaptrack-print",
            stdout_pruned=False,
        )

        if not self.skip_strace:
            self.run(
                "strace",
                [
                    self.tools["taskset"],
                    "-c",
                    self.cpu,
                    self.tools["strace"],
                    "-f",
                    "-qq",
                    "-e",
                    "trace=read,write",
                    "-o",
                    str(self.temp / "strace.log"),
                    "--",
                    *self.harness_args("strace.json"),
                ],
                stdout_name="strace.stdout.txt",
                stderr_name="strace.stderr.txt",
                check=False,
            )
            if (self.temp / "strace.log").is_file():
                self.run(
                    "strace-gzip",
                    [self.tools["gzip"], "-9", "-c", str(self.temp / "strace.log")],
                    stdout_name="strace.log.gz",
                    stderr_name="strace-gzip.stderr.txt",
                )
                self.prune_compressed_intermediate(
                    "strace.log", "strace.log.gz", "strace"
                )
        self.fincore("fincore-after")

        reports = {
            "baseline": ("baseline.json", self.samples, self.warmup),
            "time": ("time.json", self.samples, self.warmup),
            "perf-stat": ("perf-stat.json", self.samples, self.warmup),
            "perf-record": ("perf-record.json", self.samples, self.warmup),
            "heaptrack": (
                "heaptrack.json",
                self.heaptrack_samples,
                self.warmup,
            ),
        }
        if not self.skip_strace and (self.temp / "strace.json").is_file():
            reports["strace"] = ("strace.json", self.samples, self.warmup)
        for label, (name, sample_count, warmup_count) in reports.items():
            self.validate_report(
                name,
                expected_samples=sample_count,
                expected_warmup=warmup_count,
            )
            self.compress_zstd(name)

        self.write_resource_summary()
        self.write_perf_summary()
        after = self.assert_source_binding()
        return {
            "reports": list(reports),
            "source_after": after,
            "heaptrack_capture": heap_capture.name,
        }

    def write_resource_summary(self) -> None:
        text = (self.temp / "time-v.txt").read_text(
            encoding="utf-8", errors="replace"
        )
        keys = {
            "user_seconds": r"^\s*User time \(seconds\):\s*(.+?)\s*$",
            "system_seconds": r"^\s*System time \(seconds\):\s*(.+?)\s*$",
            "elapsed": r"^\s*Elapsed \(wall clock\) time .*?:\s*(\S+)\s*$",
            "max_rss_kib": r"^\s*Maximum resident set size \(kbytes\):\s*(.+?)\s*$",
            "minor_faults": r"^\s*Minor \(reclaiming a frame\) page faults:\s*(.+?)\s*$",
            "major_faults": r"^\s*Major \(requiring I/O\) page faults:\s*(.+?)\s*$",
            "voluntary_context_switches": r"^\s*Voluntary context switches:\s*(.+?)\s*$",
            "involuntary_context_switches": r"^\s*Involuntary context switches:\s*(.+?)\s*$",
        }
        result: dict[str, Any] = {}
        for key, pattern in keys.items():
            match = re.search(pattern, text, re.MULTILINE)
            if match:
                result[key] = match.group(1).strip()
        result["scope"] = (
            "whole process under /usr/bin/time -v, including corpus construction, "
            "source open, warmups, measured iterations, and verification"
        )
        json_dump(self.temp / "resource-summary.json", result)

    def write_perf_summary(self) -> None:
        rows: list[dict[str, Any]] = []
        path = self.temp / "perf-stat.csv"
        if path.is_file():
            for line in path.read_text(encoding="utf-8", errors="replace").splitlines():
                fields = line.split(",")
                if len(fields) >= 3:
                    rows.append(
                        {
                            "value": fields[0],
                            "unit": fields[1],
                            "event": fields[2],
                            "raw_fields": fields,
                        }
                    )
        json_dump(
            self.temp / "perf-summary.json",
            {
                "events": list(EVENTS),
                "scope": (
                    "whole process under perf stat, including corpus construction, "
                    "source open, warmups, measured iterations, and verification"
                ),
                "rows": rows,
            },
        )

    def make_manifest(self, capture_result: dict[str, Any], environment: dict[str, Any]) -> None:
        artifacts: list[dict[str, Any]] = []
        for path in sorted(self.temp.rglob("*")):
            if not path.is_file() or path.name == "artifact-manifest.json":
                continue
            artifacts.append(
                {
                    "path": self.rel(path),
                    "bytes": path.stat().st_size,
                    "sha256": sha256_file(path),
                }
            )
        manifest = {
            "schema_version": 1,
            "change": CHANGE,
            "kind": "current_head_hardware_profile",
            "claim_authorized": False,
            "outcome": "descriptive_baseline",
            "source": {
                "repo": str(self.repo),
                "expected_revision": self.expected_revision,
                "source_before": environment["source_before"],
                "source_after": capture_result["source_after"],
                "binary": environment["binary"],
                "binary_readelf_artifact": "binary-readelf.txt",
                "capture_driver": self.driver_artifact,
                "allowed_untracked": environment["source_before"].get(
                    "allowed_untracked"
                ),
                "clean_abba_claim_eligible": False,
            },
            "selector": {
                "case": CASE,
                "shape": SHAPE,
                "payload": PAYLOAD,
                "entry_count": 4,
                "entry_bytes": 4194304,
                "uncompressed_payload_bytes": 16777216,
            },
            "protocol": environment["protocol"],
            "profiler_order": [
                "fincore-before",
                "baseline",
                "time",
                "perf-stat",
                "perf-record",
                "perf-report-self",
                "perf-report-inclusive",
                "perf-script",
                "heaptrack",
                "heaptrack-print",
                "strace",
                "fincore-after",
            ],
            "profiler_order_is_sequential": True,
            "retained_profile_artifacts": {
                "perf_data": "perf.data",
                "perf_script": "perf-script.txt.gz",
                "heaptrack_capture": capture_result["heaptrack_capture"],
                "heaptrack_symbolized_stacks": "heaptrack-flamegraph.txt.gz",
                "strace": None if self.skip_strace else "strace.log.gz",
            },
            "parallel_metrics_validation": (
                "tools.perf_compare.validate_parallel_metrics; report identity, "
                "revision, corpus, and sample checks are explicit in the driver"
            ),
            "limitations": [
                "One current-state capture does not establish a speedup or regression.",
                f"The intentionally untracked {ALLOWED_UNTRACKED_PATH} is retained "
                "and hash-bound; this run is ineligible for a clean ABBA claim.",
                "External counters and profiles are process-total observations.",
                "Uncompressed perf-script and strace intermediates are pruned after gzip; "
                "perf.data remains retained for standard perf report reruns.",
                "Heaptrack uses 20 warmups and 100 measured samples and includes "
                "Heaptrack instrumentation overhead.",
                "The harness operation_metrics allocator status is unavailable in "
                "the normal release binary.",
                "fincore is retained as residency context; no cold-cache claim is made.",
                "strace read/write counts are whole-process syscall observations, not "
                "decompressed or recompressed byte counts.",
            ],
            "commands_artifact": "commands.json",
            "environment_artifact": "environment.json",
            "capture_driver_artifact": "capture-driver.py",
            "artifacts": artifacts,
        }
        json_dump(self.temp / "artifact-manifest.json", manifest)

    def publish(self) -> None:
        self.final.parent.mkdir(parents=True, exist_ok=True)
        staging = self.final.with_name(self.final.name + f".tmp-{os.getpid()}")
        if staging.exists():
            shutil.rmtree(staging)
        shutil.copytree(self.temp, staging)
        os.replace(staging, self.final)

    def write_commands(self) -> None:
        json_dump(
            self.temp / "commands.json",
            {
                "schema_version": 1,
                "captured_utc": self.start_time,
                "commands": self.commands,
            },
        )


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--repo", type=Path, default=DEFAULT_REPO)
    parser.add_argument("--binary", type=Path, default=DEFAULT_BINARY)
    parser.add_argument("--revision", required=True, help="expected Git HEAD")
    parser.add_argument("--temp", type=Path, default=DEFAULT_TEMP)
    parser.add_argument("--final", type=Path, default=DEFAULT_FINAL)
    parser.add_argument("--cpu", default="2", help="taskset CPU list; default: 2")
    parser.add_argument("--warmup", type=int, default=20)
    parser.add_argument("--samples", type=int, default=500)
    parser.add_argument("--heaptrack-samples", type=int, default=100)
    parser.add_argument(
        "--skip-strace",
        action="store_true",
        help="omit optional whole-process read/write syscall capture",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    arguments = parse_args(argv)
    capture = Capture(arguments)
    try:
        environment = capture.initialize()
        capture_result = capture.capture()
        capture.write_commands()
        capture.make_manifest(capture_result, environment)
        capture.publish()
        print(f"published {capture.final}")
        print(f"binary sha256 {capture.binary_sha256}")
        print(f"revision {capture.expected_revision}")
        return 0
    except Exception as error:
        if capture.temp.exists():
            failure = {
                "captured_utc": utc_now(),
                "error": str(error),
                "commands_completed": len(capture.commands),
            }
            try:
                json_dump(capture.temp / "failure.json", failure)
            except OSError:
                pass
            print(f"capture failed; partial artifacts retained at {capture.temp}", file=sys.stderr)
        print(f"litchi-goal-profile: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
