#!/usr/bin/env python3
"""Run the source-matched W5 OLE owner/history qualification.

This runner is deliberately separate from the generic 31-round lanes.  W5's
reviewed protocol calibrates every selector and clone size to a 500 ms span,
requires five ``baseline,candidate,candidate,baseline`` blocks, and then runs
the retained-state allocator capture as a non-claimable paired phase.  The
build phase uses external standalone drivers prepared by
``prepare_ole_owner_inputs.py``; no implementation source is overlaid on the
pre-owner checkout.

All output and targets are caller-supplied external directories.  A build
receipt and append-only event journal bind source, driver, lock, executable,
toolchain and fixed resource identities before ABBA can start.  This module
does not make a performance or acceptance decision.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import math
import os
from pathlib import Path
import resource
import shutil
import signal
import subprocess
import sys
import time
from typing import Any, Mapping, Sequence

TOOLS_ROOT = Path(__file__).resolve().parents[1]
if str(TOOLS_ROOT) not in sys.path:
    sys.path.insert(0, str(TOOLS_ROOT))

import native_qualification_runner as generic  # noqa: E402
from native_qualification import prepare_ole_owner_inputs as preparation  # noqa: E402


SCHEMA = "litchi.ole-owner-qualification.runner.v1"
MANIFEST_SCHEMA = "litchi.ole-owner-qualification.manifest.v1"
ROLES = ("before", "after")
VARIANTS = ("normal", "allocator")
BUILD_KEYS = tuple(f"{role}/{variant}" for role in ROLES for variant in VARIANTS)
ABBA_ORDER = ("before", "after", "after", "before")
ROUTES = (
    "open",
    "begin_edit",
    "noop_commit",
    "replace_stream",
    "read",
    "edit",
    "doc_tracked_revision",
    "xls_cells",
    "xls_cells_stage",
    "doc_history",
    "xls_history",
)
ALLOCATOR_ROUTES = ("read", "edit", "xls_cells", "xls_cells_stage", "doc_history", "xls_history")
CLONE_SIZES = ("small", "large")
WARMUP = 2
SAMPLES = 3
BLOCKS = 5
CAL_WARMUP = 1
CAL_SAMPLES = 1
MIN_SPAN_NS = 250_000_000
TARGET_SPAN_NS = 500_000_000
TARGET_PREFLIGHT_BYTES = 19_200_000_000
TARGET_HARD_LIMIT_BYTES = 20_000_000_000
MEMORY_HARD_LIMIT_BYTES = 20_000_000_000
ROUTE_SEEDS = {
    "open": 1_000,
    "begin_edit": 1_000,
    "noop_commit": 100,
    "replace_stream": 100,
    "read": 50_000,
    "edit": 60_000,
    "doc_tracked_revision": 10,
    "xls_cells": 800,
    "xls_cells_stage": 800,
    "doc_history": 1,
    "xls_history": 1,
}
CLONE_SEED = 1_000
ALLOCATOR_FIELDS = (
    "delta_calls",
    "delta_bytes",
    "live_calls",
    "live_bytes",
    "peak_calls",
    "peak_bytes",
    "peak_delta_calls",
    "peak_delta_bytes",
    "base_live_calls",
    "base_live_bytes",
    "base_peak_calls",
    "base_peak_bytes",
)


class RunnerError(RuntimeError):
    """A fail-closed manifest, source, process, or receipt error."""


def _now() -> str:
    return _dt.datetime.now(_dt.timezone.utc).isoformat(timespec="milliseconds")


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> tuple[str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"expected a regular file: {path}")
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while block := stream.read(1024 * 1024):
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _resolve(value: str | os.PathLike[str]) -> Path:
    try:
        return generic._resolve_path(str(value))
    except generic.RunnerError as error:
        raise RunnerError(str(error)) from error


def _under(path: Path, parent: Path) -> bool:
    return generic._under(path, parent)


def _git(root: Path, *args: str) -> str:
    result = subprocess.run(
        ["git", "-C", str(root), *args],
        check=False,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    if result.returncode:
        raise RunnerError(f"git {' '.join(args)} failed in {root}: {result.stderr.strip()}")
    return result.stdout


def _status_paths(status: str) -> list[str]:
    result: list[str] = []
    for line in status.splitlines():
        if len(line) < 4:
            raise RunnerError(f"malformed git status line: {line!r}")
        path = line[3:]
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        result.append(path)
    return result


def _source_state(root: Path) -> dict[str, Any]:
    root = _resolve(root)
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    diff = _git(root, "diff", "--no-ext-diff", "--binary")
    return {
        "root": str(root),
        "head": _git(root, "rev-parse", "HEAD").strip(),
        "tree": _git(root, "rev-parse", "HEAD^{tree}").strip(),
        "status": status,
        "status_paths": _status_paths(status),
        "clean": not bool(status),
        "status_sha256": _sha256_bytes(status.encode()),
        "diff_sha256": _sha256_bytes(diff.encode()),
        "production_tree_sha256": _sha256_bytes(
            _git(root, "ls-tree", "-r", "--full-tree", "HEAD").encode()
        ),
    }


def _json_hash(path: Path) -> tuple[str, int]:
    return _sha256_file(path)


def _write_json(path: Path, value: Any, *, replace: bool = False) -> None:
    if path.is_symlink() or (path.exists() and not replace):
        raise RunnerError(f"refusing to overwrite output: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def _require_mapping(value: Any, label: str) -> Mapping[str, Any]:
    if not isinstance(value, Mapping):
        raise RunnerError(f"{label} must be an object")
    return value


def _load_json(path: Path, label: str) -> tuple[dict[str, Any], str, int]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"{label} is missing or symlinked: {path}")
    try:
        raw = path.read_bytes()
        document = json.loads(raw)
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read {label}: {error}") from error
    if not isinstance(document, dict):
        raise RunnerError(f"{label} root must be an object")
    return document, _sha256_bytes(raw), len(raw)


def _load_manifest(path: Path) -> tuple[dict[str, Any], str]:
    document, digest, _ = _load_json(path, "manifest")
    if document.get("schema") != MANIFEST_SCHEMA or document.get("lane") != "w5-ole-owner":
        raise RunnerError("OLE manifest schema or lane differs")
    if document.get("candidate_commit") is not None:
        raise RunnerError("candidate_commit must remain runtime-parameterized")
    expected_routes = list(ROUTES)
    if document.get("routes") != expected_routes:
        raise RunnerError("manifest route set/order differs from the reviewed eleven-route workload")
    paired = _require_mapping(document.get("paired"), "manifest.paired")
    if paired.get("abba_order") != list(ABBA_ORDER) or paired.get("blocks") != BLOCKS:
        raise RunnerError("manifest ABBA order or block count differs")
    if paired.get("warmup") != WARMUP or paired.get("samples") != SAMPLES:
        raise RunnerError("manifest warmup/sample count differs")
    if paired.get("calibration_target_ns") != TARGET_SPAN_NS or paired.get("minimum_span_ns") != MIN_SPAN_NS:
        raise RunnerError("manifest calibration/span bounds differ")
    clone = _require_mapping(document.get("clone"), "manifest.clone")
    if clone.get("sizes") != list(CLONE_SIZES) or clone.get("small_siblings") != 1 or clone.get("large_siblings") != 1024:
        raise RunnerError("manifest clone workload differs")
    fixture = _require_mapping(document.get("fixture"), "manifest.fixture")
    if fixture.get("kind") != "probe-owned deterministic OLE/DOC/XLS corpus" or fixture.get("generator_identity") != "bound to the reviewed probe pins; no untracked fixture path":
        raise RunnerError("manifest fixture policy differs")
    allocator = _require_mapping(document.get("allocator"), "manifest.allocator")
    if allocator.get("routes") != list(ALLOCATOR_ROUTES) or allocator.get("history_steps") != 128:
        raise RunnerError("manifest allocator routes/history differ")
    if allocator.get("retained_packages") != 129 or allocator.get("retained_patches") != 128:
        raise RunnerError("manifest retained history cardinality differs")
    if allocator.get("elapsed_claimable") is not False:
        raise RunnerError("allocator elapsed values must be non-claimable")
    quiet = _require_mapping(document.get("quiet"), "manifest.quiet")
    if quiet.get("process_guard") != "tools.native_qualification_runner.ProcessGuard":
        raise RunnerError("manifest process guard differs")
    if quiet.get("cpu_affinity_required") is not True:
        raise RunnerError("manifest requires CPU affinity")
    if quiet.get("target_preflight_bytes") != TARGET_PREFLIGHT_BYTES or quiet.get("target_hard_bytes") != TARGET_HARD_LIMIT_BYTES or quiet.get("memory_bytes") != MEMORY_HARD_LIMIT_BYTES:
        raise RunnerError("manifest resource limits differ from fixed runner limits")
    if quiet.get("cargo_jobs") != 1 or quiet.get("incremental") is not False:
        raise RunnerError("manifest Cargo resource policy differs")
    pins = _require_mapping(document.get("input_pins"), "manifest.input_pins")
    for key, expected in {
        "reviewed_probe_sha256": preparation.REVIEWED_PROBE_SHA256,
        "candidate_adapter_sha256": preparation.CANDIDATE_ADAPTER_SHA256,
        "clone_probe_sha256": preparation.CLONE_PROBE_SHA256,
        "lock_sha256": preparation.LOCK_SHA256,
    }.items():
        if pins.get(key) != expected:
            raise RunnerError(f"manifest input pin differs: {key}")
    return document, digest


def _event_hash(events: Sequence[Mapping[str, Any]]) -> str:
    return _sha256_bytes(json.dumps(list(events), sort_keys=True, separators=(",", ":")).encode())


def _event_lines(path: Path) -> list[dict[str, Any]]:
    if path.is_symlink() or not path.is_file():
        raise RunnerError(f"event journal is missing or symlinked: {path}")
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
        events = [json.loads(line) for line in lines]
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"invalid event journal {path}: {error}") from error
    if not all(isinstance(event, dict) for event in events):
        raise RunnerError(f"event journal contains a non-object event: {path}")
    return events


class PairedRunner:
    """Source-bound build and five-block ABBA runner."""

    def __init__(
        self,
        *,
        manifest_path: Path,
        preparation_path: Path,
        output: Path,
        target_root: Path,
        phase: str,
        cpu: int | None,
    ) -> None:
        if phase not in ("build", "abba"):
            raise RunnerError(f"unknown phase: {phase}")
        self.manifest_path = _resolve(manifest_path)
        self.preparation_path = _resolve(preparation_path)
        self.output = _resolve(output)
        self.target_root = _resolve(target_root)
        self.phase = phase
        self.cpu = cpu
        self.manifest, self.manifest_sha256 = _load_manifest(self.manifest_path)
        self.preparation, self.preparation_sha256, _ = _load_json(self.preparation_path, "preparation")
        try:
            preparation.validate_preparation_receipt(self.preparation_path)
        except preparation.PreparationError as error:
            raise RunnerError(f"preparation validation failed: {error}") from error
        self.runner_sha256, _ = _sha256_file(Path(__file__).resolve())
        self.guard = generic.ProcessGuard()
        self.events_path = self.output / f"events-{phase}.jsonl"
        self.receipt_path = self.output / f"{phase}-receipt.json"
        self.recorded_events: list[dict[str, Any]] = []
        self.host_dir = self.output / "host"
        self.driver_roots = {
            "before": _resolve(str(self.preparation["baseline"]["driver"]["root"])),
            "after": _resolve(str(self.preparation["candidate"]["driver"]["root"])),
        }
        self.source_roots = {
            "before": _resolve(str(self.preparation["baseline"]["root"])),
            "after": _resolve(str(self.preparation["candidate"]["root"])),
        }
        self.artifacts: dict[str, dict[str, dict[str, Path]]] = {
            role: {variant: {} for variant in VARIANTS} for role in ROLES
        }
        if self.phase == "build":
            self._prepare_build_destinations()
        else:
            if self.cpu is None:
                raise RunnerError("ABBA requires an explicit CPU affinity")
            if self.cpu not in os.sched_getaffinity(0):
                raise RunnerError(f"CPU {self.cpu} is not in the runner affinity")
            self._prepare_abba_handoff()

    def _prepare_build_destinations(self) -> None:
        if self.output.is_symlink() or self.output.exists() and not self.output.is_dir():
            raise RunnerError(f"build output must be a directory: {self.output}")
        if self.output.exists() and any(self.output.iterdir()):
            raise RunnerError(f"build output must be fresh and empty: {self.output}")
        self.output.mkdir(parents=True, exist_ok=True)
        if self.target_root.is_symlink() or self.target_root.exists() and not self.target_root.is_dir():
            raise RunnerError(f"target root must be a directory: {self.target_root}")
        self.target_root.mkdir(parents=True, exist_ok=True)
        if any(self.target_root.iterdir()):
            raise RunnerError(f"target root must be fresh and empty: {self.target_root}")
        for source in self.source_roots.values():
            if _under(self.output, source) or _under(self.target_root, source):
                raise RunnerError("build output and targets must be outside source worktrees")
        for driver in self.driver_roots.values():
            if _under(self.output, driver) or _under(self.target_root, driver):
                raise RunnerError("build output and targets must be outside prepared driver trees")
        if _under(self.output, self.target_root) or _under(self.target_root, self.output):
            raise RunnerError("build output and target root must be disjoint")
        self.host_dir.mkdir(exist_ok=True)
        (self.output / "commands").mkdir()
        if self.events_path.exists() or self.receipt_path.exists():
            raise RunnerError("build journal/receipt already exists")

    def _prepare_abba_handoff(self) -> None:
        if self.output.is_symlink() or not self.output.is_dir():
            raise RunnerError("ABBA output directory is missing or symlinked")
        if self.receipt_path.exists() or self.events_path.exists():
            raise RunnerError("ABBA journal/receipt already exists")
        if not self.target_root.is_dir() or self.target_root.is_symlink():
            raise RunnerError("ABBA target root is missing or symlinked")
        for source in self.source_roots.values():
            if _under(self.output, source) or _under(self.target_root, source):
                raise RunnerError("ABBA output and targets must be outside source worktrees")
        for driver in self.driver_roots.values():
            if _under(self.output, driver) or _under(self.target_root, driver):
                raise RunnerError("ABBA output and targets must be outside prepared driver trees")
        if _under(self.output, self.target_root) or _under(self.target_root, self.output):
            raise RunnerError("ABBA output and target root must be disjoint")
        self.host_dir.mkdir(exist_ok=True)
        (self.output / "runs").mkdir()
        (self.output / "calibration").mkdir()
        self._load_build_handoff()

    def _endpoint(self, role: str) -> Mapping[str, Any]:
        if role == "before":
            return _require_mapping(self.preparation["baseline"], "preparation.baseline")
        if role == "after":
            return _require_mapping(self.preparation["candidate"], "preparation.candidate")
        raise RunnerError(f"unknown endpoint role {role}")

    def _source_identity(self, role: str) -> dict[str, Any]:
        endpoint = self._endpoint(role)
        recorded = _require_mapping(endpoint, f"preparation.{role}")
        current = _source_state(self.source_roots[role])
        for key in (
            "root",
            "head",
            "tree",
            "status",
            "status_paths",
            "clean",
            "status_sha256",
            "diff_sha256",
            "production_tree_sha256",
        ):
            if current.get(key) != recorded.get(key):
                raise RunnerError(f"{role} source identity changed: {key}")
        driver = _require_mapping(recorded.get("driver"), f"preparation.{role}.driver")
        files = driver.get("files")
        if not isinstance(files, list) or len(files) != 4:
            raise RunnerError(f"{role} driver file identity is incomplete")
        driver_files: list[dict[str, Any]] = []
        for record in files:
            item = _require_mapping(record, f"{role} driver file")
            path = _resolve(str(item.get("path", "")))
            digest, size = _sha256_file(path)
            if digest != item.get("sha256") or size != item.get("bytes"):
                raise RunnerError(f"{role} driver input changed: {path}")
            driver_files.append({"path": str(path), "bytes": size, "sha256": digest})
        identity = {"source": current, "driver_root": str(self.driver_roots[role]), "driver_files": driver_files}
        identity["identity_sha256"] = _sha256_bytes(
            json.dumps(identity, sort_keys=True, separators=(",", ":")).encode()
        )
        return identity

    def _target(self, role: str, variant: str) -> Path:
        if role not in ROLES or variant not in VARIANTS:
            raise RunnerError("invalid target role/variant")
        return self.target_root / f"{role}-{variant}"

    def _binary_paths(self, role: str, variant: str) -> dict[str, Path]:
        target = self._target(role, variant)
        return {
            "probe": target / "release" / "examples" / "ole_w5_probe",
            "clone": target / "release" / "examples" / "ole_package_clone_probe",
        }

    def _toolchain(self, env: Mapping[str, str]) -> dict[str, Any]:
        values: dict[str, Any] = {}
        for name, args in (("cargo", ["-V"]), ("rustc", ["-Vv"])):
            executable = shutil.which(name, path=env.get("PATH"))
            if executable is None:
                raise RunnerError(f"{name} is not available in build PATH")
            path = _resolve(executable)
            digest, size = _sha256_file(path)
            result = subprocess.run(
                [str(path), *args],
                cwd=str(self.driver_roots["before"]),
                env=dict(env),
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            if result.returncode:
                raise RunnerError(f"{name} identity command failed: {result.stderr.strip()}")
            values[name] = {
                "path": str(path),
                "bytes": size,
                "sha256": digest,
                "version": result.stdout,
            }
        return values

    def _environment(self, role: str, variant: str) -> dict[str, str]:
        env = dict(os.environ)
        env["CARGO_TARGET_DIR"] = str(self._target(role, variant))
        env["CARGO_BUILD_JOBS"] = "1"
        env["CARGO_INCREMENTAL"] = "0"
        env["LITCHI_MEM_MAX"] = "20G"
        return env

    def _preexec(self):
        cpu = self.cpu

        def configure() -> None:
            if cpu is not None:
                os.sched_setaffinity(0, {cpu})
            resource.setrlimit(resource.RLIMIT_AS, (MEMORY_HARD_LIMIT_BYTES, MEMORY_HARD_LIMIT_BYTES))

        return configure

    def _host_snapshot(self, label: str, child: int | None = None) -> list[dict[str, Any]]:
        snapshot = generic._proc_snapshot()
        offenders = self.guard.offenders(child)
        serializable = {str(pid): entry for pid, entry in snapshot.items()}
        path = self.host_dir / (label.replace("/", "-") + ".json")
        _write_json(
            path,
            {
                "label": label,
                "time_ns": time.time_ns(),
                "cpu": self.cpu,
                "offenders": offenders,
                "processes": serializable,
            },
        )
        return offenders

    def _target_size(self, target: Path, *, preflight: bool = False) -> int:
        if target.is_symlink():
            raise RunnerError(f"target is symlinked: {target}")
        size = generic._directory_size(target)
        limit = TARGET_PREFLIGHT_BYTES if preflight else TARGET_HARD_LIMIT_BYTES
        if size > limit:
            raise RunnerError(f"target exceeds {limit} bytes: {target}")
        return size

    def _terminate(self, process: subprocess.Popen[Any]) -> None:
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except (ProcessLookupError, PermissionError):
            process.terminate()
        try:
            process.wait(timeout=5)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(process.pid, signal.SIGKILL)
            except (ProcessLookupError, PermissionError):
                process.kill()
            process.wait()

    def _wait_descendants_gone(self, pid: int) -> list[int]:
        deadline = time.monotonic() + 5.0
        while time.monotonic() < deadline:
            snapshot = generic._proc_snapshot()
            descendants = generic._descendants(snapshot, pid)
            remaining = [child for child in descendants if child != pid]
            if not remaining:
                return []
            time.sleep(0.1)
        return remaining

    def _stabilize_target(self, target: Path) -> int:
        previous: int | None = None
        stable = 0
        for _ in range(10):
            size = self._target_size(target)
            if size == previous:
                stable += 1
                if stable >= 2:
                    return size
            else:
                stable = 0
            previous = size
            time.sleep(0.2)
        raise RunnerError(f"target did not stabilize: {target}")

    def _new_log(self, path: Path) -> Path:
        if path.is_symlink() or path.exists():
            raise RunnerError(f"refusing to overwrite child log: {path}")
        path.parent.mkdir(parents=True, exist_ok=True)
        if path.parent.is_symlink():
            raise RunnerError(f"child log parent is symlinked: {path.parent}")
        return path

    def _append_event(self, event: Mapping[str, Any]) -> None:
        if self.events_path.is_symlink():
            raise RunnerError(f"event journal is symlinked: {self.events_path}")
        self.events_path.parent.mkdir(parents=True, exist_ok=True)
        mode = "x" if not self.events_path.exists() else "a"
        with self.events_path.open(mode, encoding="utf-8") as stream:
            stream.write(json.dumps(event, sort_keys=True) + "\n")
            stream.flush()
            os.fsync(stream.fileno())
        self.recorded_events.append(dict(event))

    def _run_process(
        self,
        *,
        event_id: str,
        role: str,
        variant: str,
        argv: Sequence[str],
        cwd: Path,
        env: Mapping[str, str],
        stdout_path: Path,
        stderr_path: Path,
        expected_outputs: Sequence[Path] = (),
        require_stdout: bool = False,
    ) -> dict[str, Any]:
        self.guard.require_quiet()
        identity_before = self._source_identity(role)
        self._host_snapshot(event_id + "-before")
        target = self._target(role, variant)
        target_before = self._target_size(target, preflight=self.phase == "build")
        stdout_path = self._new_log(stdout_path)
        stderr_path = self._new_log(stderr_path)
        start = time.monotonic_ns()
        process: subprocess.Popen[Any] | None = None
        returncode: int | None = None
        violation: list[dict[str, Any]] = []
        runner_error: str | None = None
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            try:
                process = subprocess.Popen(
                    list(argv),
                    cwd=str(cwd),
                    env=dict(env),
                    stdin=subprocess.DEVNULL,
                    stdout=stdout,
                    stderr=stderr,
                    start_new_session=True,
                    preexec_fn=self._preexec(),
                )
                checks = 0
                while True:
                    returncode = process.poll()
                    if returncode is not None:
                        break
                    checks += 1
                    offenders = self._host_snapshot(f"{event_id}-during-{checks:05d}", process.pid)
                    if offenders:
                        violation.extend(offenders)
                        runner_error = "process guard rejected an unrelated Litchi process"
                        self._terminate(process)
                        returncode = process.returncode
                        break
                    self._target_size(target)
                    time.sleep(0.1)
            except BaseException as error:
                runner_error = str(error)
                if process is not None and process.poll() is None:
                    self._terminate(process)
                returncode = process.returncode if process is not None else None
        post_descendants = self._wait_descendants_gone(process.pid) if process is not None else []
        post_offenders = self.guard.offenders()
        if post_descendants:
            runner_error = runner_error or f"child descendants survived exit: {post_descendants}"
        if post_offenders:
            runner_error = runner_error or "unrelated Litchi process remained after child exit"
        self._host_snapshot(event_id + "-after")
        try:
            target_after = self._stabilize_target(target)
        except RunnerError as error:
            target_after = -1
            runner_error = runner_error or str(error)
        try:
            stdout_hash, stdout_bytes = _sha256_file(stdout_path)
            stderr_hash, stderr_bytes = _sha256_file(stderr_path)
        except RunnerError as error:
            runner_error = runner_error or str(error)
            stdout_hash, stdout_bytes = _sha256_bytes(b""), 0
            stderr_hash, stderr_bytes = _sha256_bytes(b""), 0
        outputs: list[dict[str, Any]] = []
        for path in expected_outputs:
            if path.is_symlink() or not path.is_file():
                runner_error = runner_error or f"declared output is missing or symlinked: {path}"
                outputs.append({"path": str(path), "exists": False})
                continue
            digest, size = _sha256_file(path)
            outputs.append({"path": str(path), "exists": True, "bytes": size, "sha256": digest})
        try:
            identity_after = self._source_identity(role)
        except RunnerError as error:
            identity_after = None
            runner_error = runner_error or str(error)
        if identity_after is not None and identity_after["identity_sha256"] != identity_before["identity_sha256"]:
            runner_error = runner_error or f"{role} source/driver identity changed during {event_id}"
        event = {
            "schema": SCHEMA,
            "event_id": event_id,
            "phase": self.phase,
            "role": role,
            "variant": variant,
            "argv": list(argv),
            "cwd": str(cwd),
            "environment": {
                key: env.get(key)
                for key in ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "LITCHI_MEM_MAX")
            },
            "source_identity_before": identity_before,
            "source_identity_sha256": identity_before["identity_sha256"],
            "source_identity_after": identity_after,
            "start_utc": _now(),
            "duration_ns": time.monotonic_ns() - start,
            "returncode": returncode,
            "stdout": {"path": str(stdout_path), "bytes": stdout_bytes, "sha256": stdout_hash},
            "stderr": {"path": str(stderr_path), "bytes": stderr_bytes, "sha256": stderr_hash},
            "outputs": outputs,
            "target": {"path": str(target), "bytes_before": target_before, "bytes_after": target_after},
            "process_guard_violations": violation,
            "post_exit_descendants": post_descendants,
            "post_exit_offenders": post_offenders,
            "runner_error": runner_error,
        }
        self._append_event(event)
        if runner_error:
            raise RunnerError(f"{runner_error} during {event_id}")
        if returncode != 0:
            raise RunnerError(f"child failed during {event_id}: exit {returncode}")
        if require_stdout and stdout_bytes == 0:
            raise RunnerError(f"child emitted no stdout during {event_id}")
        if any(not item.get("exists") for item in outputs):
            raise RunnerError(f"child omitted a declared output during {event_id}")
        return event

    def _features(self, role: str, variant: str) -> str:
        if role == "before":
            features = "legacy-owner-api,legacy-property-api"
        else:
            features = "current-api"
        if variant == "allocator":
            features += ",w5-counting-allocator"
        return features

    def _build_command(self, role: str, variant: str) -> list[str]:
        cargo = os.environ.get("CARGO", "cargo")
        return [
            cargo,
            "build",
            "--release",
            "--offline",
            "--locked",
            "--manifest-path",
            str(self.driver_roots[role] / "Cargo.toml"),
            "--no-default-features",
            "--features",
            self._features(role, variant),
            "--example",
            "ole_w5_probe",
            "--example",
            "ole_package_clone_probe",
        ]

    def _write_receipt(self, status: str, *, error: str | None = None, extra: Mapping[str, Any] | None = None) -> None:
        journal_record: dict[str, Any] = {
            "path": str(self.events_path),
            "count": len(self.recorded_events),
            "event_ids": [event.get("event_id") for event in self.recorded_events],
            "sha256": _event_hash(self.recorded_events),
        }
        if self.events_path.is_file() and not self.events_path.is_symlink():
            journal_digest, journal_bytes = _sha256_file(self.events_path)
            journal_record.update({"bytes": journal_bytes, "file_sha256": journal_digest})
        receipt: dict[str, Any] = {
            "schema": SCHEMA,
            "phase": self.phase,
            "status": status,
            "manifest": {"path": str(self.manifest_path), "sha256": self.manifest_sha256},
            "preparation": {"path": str(self.preparation_path), "sha256": self.preparation_sha256},
            "runner": {"path": str(Path(__file__).resolve()), "sha256": self.runner_sha256},
            "limits": {
                "target_preflight_bytes": TARGET_PREFLIGHT_BYTES,
                "target_hard_bytes": TARGET_HARD_LIMIT_BYTES,
                "memory_bytes": MEMORY_HARD_LIMIT_BYTES,
            },
            "abba_order": list(ABBA_ORDER),
            "event_journal": journal_record,
            "claim_scope": "orchestration and source-bound evidence only; no performance acceptance claim",
        }
        if error is not None:
            receipt["error"] = error
        if extra:
            receipt.update(extra)
        _write_json(self.receipt_path, receipt)

    def build(self) -> None:
        try:
            env = dict(os.environ)
            toolchain = self._toolchain(env)
            builds: dict[str, Any] = {}
            for role in ROLES:
                for variant in VARIANTS:
                    key = f"{role}/{variant}"
                    target = self._target(role, variant)
                    target.mkdir()
                    binaries = self._binary_paths(role, variant)
                    command = self._build_command(role, variant)
                    build_env = self._environment(role, variant)
                    event = self._run_process(
                        event_id=f"build-{role}-{variant}",
                        role=role,
                        variant=variant,
                        argv=command,
                        cwd=self.driver_roots[role],
                        env=build_env,
                        stdout_path=self.output / "commands" / f"{key.replace('/', '-')}.stdout",
                        stderr_path=self.output / "commands" / f"{key.replace('/', '-')}.stderr",
                        expected_outputs=tuple(binaries.values()),
                        require_stdout=False,
                    )
                    artifacts = {
                        name: {"path": str(path), "sha256": _sha256_file(path)[0], "bytes": _sha256_file(path)[1]}
                        for name, path in binaries.items()
                    }
                    self.artifacts[role][variant] = {name: path for name, path in binaries.items()}
                    builds[key] = {
                        "role": role,
                        "variant": variant,
                        "features": self._features(role, variant),
                        "target": str(target),
                        "command": command,
                        "event_id": event["event_id"],
                        "artifacts": artifacts,
                    }
            self._write_receipt(
                "passed",
                extra={
                    "toolchain": toolchain,
                    "identities": {role: self._source_identity(role) for role in ROLES},
                    "builds": builds,
                    "target_root": str(self.target_root),
                },
            )
        except BaseException as error:
            try:
                self._write_receipt("failed", error=str(error))
            except BaseException:
                pass
            raise

    def _load_build_handoff(self) -> None:
        receipt, receipt_hash, _ = _load_json(self.output / "build-receipt.json", "build receipt")
        if receipt.get("schema") != SCHEMA or receipt.get("phase") != "build" or receipt.get("status") != "passed":
            raise RunnerError("build receipt is not a successful OLE build receipt")
        if receipt.get("manifest", {}).get("sha256") != self.manifest_sha256:
            raise RunnerError("build receipt manifest identity differs")
        if receipt.get("preparation", {}).get("sha256") != self.preparation_sha256:
            raise RunnerError("build receipt preparation identity differs")
        if receipt.get("limits") != {
            "target_preflight_bytes": TARGET_PREFLIGHT_BYTES,
            "target_hard_bytes": TARGET_HARD_LIMIT_BYTES,
            "memory_bytes": MEMORY_HARD_LIMIT_BYTES,
        }:
            raise RunnerError("build receipt resource limits differ")
        if receipt.get("target_root") != str(self.target_root):
            raise RunnerError("ABBA target root differs from build receipt")
        journal = _require_mapping(receipt.get("event_journal"), "build receipt event journal")
        event_path = _resolve(str(journal.get("path", "")))
        if event_path != self.output / "events-build.jsonl":
            raise RunnerError("build event journal path differs")
        events = _event_lines(event_path)
        journal_digest, journal_bytes = _sha256_file(event_path)
        if (
            journal.get("count") != len(events)
            or journal.get("sha256") != _event_hash(events)
            or journal.get("file_sha256") != journal_digest
            or journal.get("bytes") != journal_bytes
        ):
            raise RunnerError("build event journal hash/count differs")
        expected_ids = [f"build-{key.replace('/', '-')}" for key in BUILD_KEYS]
        if [event.get("event_id") for event in events] != expected_ids:
            raise RunnerError("build event journal order differs")
        builds = _require_mapping(receipt.get("builds"), "build receipt builds")
        identities = _require_mapping(receipt.get("identities"), "receipt.identities")
        for role in ROLES:
            identity = self._source_identity(role)
            recorded_identity = _require_mapping(identities.get(role), f"receipt.identities.{role}")
            if identity != recorded_identity:
                raise RunnerError(f"{role} source/driver identity changed since build")
            for variant in VARIANTS:
                key = f"{role}/{variant}"
                build = _require_mapping(builds.get(key), f"receipt.builds.{key}")
                if build.get("event_id") != f"build-{role}-{variant}" or build.get("target") != str(self._target(role, variant)):
                    raise RunnerError(f"build endpoint contract differs for {key}")
                event = events[BUILD_KEYS.index(key)]
                if (
                    event.get("returncode") != 0
                    or event.get("runner_error") is not None
                    or event.get("source_identity_sha256") != recorded_identity.get("identity_sha256")
                    or event.get("source_identity_before") != recorded_identity
                    or event.get("source_identity_after") != recorded_identity
                ):
                    raise RunnerError(f"build event identity or result differs for {key}")
                artifacts = _require_mapping(build.get("artifacts"), f"receipt.builds.{key}.artifacts")
                expected = self._binary_paths(role, variant)
                for name, path in expected.items():
                    item = _require_mapping(artifacts.get(name), f"artifact {key}/{name}")
                    if _resolve(str(item.get("path", ""))) != path:
                        raise RunnerError(f"artifact path differs for {key}/{name}")
                    if not path.is_file() or not os.access(path, os.X_OK):
                        raise RunnerError(f"artifact is missing/non-executable: {path}")
                    digest, size = _sha256_file(path)
                    if item.get("sha256") != digest or item.get("bytes") != size:
                        raise RunnerError(f"artifact hash changed: {path}")
                    self.artifacts[role][variant][name] = path
        self.build_receipt_hash = receipt_hash

    def _parse_route(self, path: Path, route: str, expected_rows: int) -> list[dict[str, int]]:
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except (OSError, UnicodeDecodeError) as error:
            raise RunnerError(f"cannot read probe output {path}: {error}") from error
        if not lines or lines[0] != "route,sample,nanoseconds,output_bytes":
            raise RunnerError(f"{path}: route header differs")
        rows: list[dict[str, int]] = []
        for line in lines[1:]:
            fields = line.split(",")
            if len(fields) != 4 or fields[0] != route:
                raise RunnerError(f"{path}: unexpected route output row: {line!r}")
            if not all(field.isdigit() for field in fields[1:]):
                raise RunnerError(f"{path}: malformed route row: {line!r}")
            rows.append({"sample": int(fields[1]), "nanoseconds": int(fields[2]), "output_bytes": int(fields[3])})
        if len(rows) != expected_rows or [row["sample"] for row in rows] != list(range(expected_rows)):
            raise RunnerError(f"{path}: expected ordered {expected_rows}-row {route} output")
        if any(row["output_bytes"] <= 0 for row in rows):
            raise RunnerError(f"{path}: route output extent is empty")
        return rows

    def _parse_clone(
        self,
        path: Path,
        expected_rows: int,
        *,
        samples: int | None = None,
    ) -> list[dict[str, int | str]]:
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except (OSError, UnicodeDecodeError) as error:
            raise RunnerError(f"cannot read clone output {path}: {error}") from error
        if not lines or lines[0] != "route,size,sample,nanoseconds,batch,archive_bytes":
            raise RunnerError(f"{path}: clone header differs")
        rows: list[dict[str, int | str]] = []
        for line in lines[1:]:
            fields = line.split(",")
            if len(fields) != 6 or fields[0] != "clone_drop" or fields[1] not in CLONE_SIZES:
                raise RunnerError(f"{path}: malformed clone row: {line!r}")
            if not all(field.isdigit() for field in fields[2:]):
                raise RunnerError(f"{path}: malformed clone numeric field: {line!r}")
            rows.append({"size": fields[1], "sample": int(fields[2]), "nanoseconds": int(fields[3]), "batch": int(fields[4]), "archive_bytes": int(fields[5])})
        if len(rows) != expected_rows:
            raise RunnerError(f"{path}: expected {expected_rows} clone rows")
        sample_count = SAMPLES if samples is None else samples
        if expected_rows != len(CLONE_SIZES) * sample_count:
            raise RunnerError(f"{path}: clone row count does not match sample count")
        for size in CLONE_SIZES:
            selected = [row for row in rows if row["size"] == size]
            if len(selected) != sample_count or [row["sample"] for row in selected] != list(range(sample_count)):
                raise RunnerError(f"{path}: clone {size} rows are incomplete")
        return rows

    def _parse_allocator(self, path: Path, route: str) -> dict[str, Any]:
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except (OSError, UnicodeDecodeError) as error:
            raise RunnerError(f"cannot read allocator output {path}: {error}") from error
        if len(lines) < 4 or lines[0] != "route,sample,nanoseconds,output_bytes" or not lines[1].startswith("w5_alloc,phase,route,"):
            raise RunnerError(f"{path}: allocator headers are missing")
        normal_rows: list[dict[str, int]] = []
        for line in lines[2:]:
            fields = line.split(",")
            if fields and fields[0] == "w5_alloc":
                continue
            if len(fields) != 4 or fields[0] != route or not all(field.isdigit() for field in fields[1:]):
                raise RunnerError(f"{path}: malformed route row: {line!r}")
            normal_rows.append({"sample": int(fields[1]), "nanoseconds": int(fields[2]), "output_bytes": int(fields[3])})
        if len(normal_rows) != 1 or normal_rows[0]["sample"] != 0 or normal_rows[0]["output_bytes"] <= 0:
            raise RunnerError(f"{path}: allocator route output is incomplete")
        normal = normal_rows
        header = lines[1].split(",")
        if len(header) != 16 or header[0] != "w5_alloc":
            raise RunnerError(f"{path}: allocator header differs")
        rows: list[dict[str, Any]] = []
        for line in lines[2:]:
            fields = line.split(",")
            if fields and fields[0] != "w5_alloc":
                continue
            if len(fields) != len(header) or fields[0] != "w5_alloc" or fields[2] != route:
                raise RunnerError(f"{path}: malformed allocator row: {line!r}")
            numeric = fields[3:]
            if not all(value.isdigit() for value in numeric):
                raise RunnerError(f"{path}: allocator counter is not an integer")
            row = dict(zip(header[1:], fields[1:]))
            row.update({key: int(row[key]) for key in header[3:]})
            rows.append(row)
        expected_phases = ("stage", "full") if route == "xls_cells_stage" else ("full",)
        if [str(row["phase"]) for row in rows] != list(expected_phases) or any(row["iteration"] != 0 for row in rows):
            raise RunnerError(f"{path}: allocator phase rows differ for {route}")
        for row in rows:
            for field in ALLOCATOR_FIELDS:
                if not isinstance(row.get(field), int) or row[field] < 0:
                    raise RunnerError(f"{path}: allocator field {field} is invalid")
        return {"route": route, "output_bytes": normal[0]["output_bytes"], "rows": rows}

    @staticmethod
    def _route_stats(rows: Sequence[Mapping[str, int]]) -> dict[str, Any]:
        values = [int(row["nanoseconds"]) for row in rows]
        return {
            "count": len(values),
            "minimum_ns": min(values),
            "maximum_ns": max(values),
            "sum_ns": sum(values),
            "output_bytes": sorted({int(row["output_bytes"]) for row in rows}),
        }

    def _run_probe(self, role: str, variant: str, label: str, args: Sequence[str]) -> tuple[Path, dict[str, Any]]:
        binary = self.artifacts[role][variant]["probe"]
        env = self._environment(role, variant)
        stdout = self.output / "runs" / f"{label}.stdout"
        stderr = self.output / "runs" / f"{label}.stderr"
        event = self._run_process(
            event_id=f"run-{label}",
            role=role,
            variant=variant,
            argv=[str(binary), *args],
            cwd=self.driver_roots[role],
            env=env,
            stdout_path=stdout,
            stderr_path=stderr,
            require_stdout=True,
        )
        return stdout, event

    def _run_clone(
        self,
        role: str,
        variant: str,
        label: str,
        batch: int,
        *,
        warmup: int = WARMUP,
        samples: int = SAMPLES,
    ) -> tuple[Path, dict[str, Any]]:
        binary = self.artifacts[role][variant]["clone"]
        env = self._environment(role, variant)
        stdout = self.output / "runs" / f"{label}.stdout"
        stderr = self.output / "runs" / f"{label}.stderr"
        event = self._run_process(
            event_id=f"run-{label}",
            role=role,
            variant=variant,
            argv=[str(binary), "--warmup", str(warmup), "--samples", str(samples), "--batch", str(batch)],
            cwd=self.driver_roots[role],
            env=env,
            stdout_path=stdout,
            stderr_path=stderr,
            require_stdout=True,
        )
        return stdout, event

    def _calibrate_route(self, role: str, route: str) -> int:
        batch = ROUTE_SEEDS[route]
        attempt = 0
        while batch <= 1_000_000_000:
            attempt += 1
            label = f"cal-{role}-{route}-{attempt:03d}"
            stdout, _ = self._run_probe(role, "normal", label, ["--route", route, "--warmup", str(CAL_WARMUP), "--samples", str(CAL_SAMPLES), "--batch", str(batch)])
            rows = self._parse_route(stdout, route, CAL_SAMPLES)
            self._write_measurement(
                f"calibration/{label}.json",
                {"role": role, "route": route, "batch": batch, "rows": rows},
            )
            span = rows[0]["nanoseconds"]
            if span >= TARGET_SPAN_NS:
                return batch
            batch = max(batch + 1, math.ceil(batch * TARGET_SPAN_NS / max(1, span) * 1.2))
        raise RunnerError(f"calibration exceeded batch limit for {role}/{route}")

    def _calibrate_clone(self, role: str) -> int:
        batch = CLONE_SEED
        attempt = 0
        while batch <= 1_000_000_000:
            attempt += 1
            label = f"cal-{role}-clone-{attempt:03d}"
            stdout, _ = self._run_clone(
                role,
                "normal",
                label,
                batch,
                warmup=CAL_WARMUP,
                samples=CAL_SAMPLES,
            )
            rows = self._parse_clone(stdout, 2 * CAL_SAMPLES, samples=CAL_SAMPLES)
            self._write_measurement(f"calibration/{label}.json", {"role": role, "batch": batch, "rows": rows})
            if min(int(row["nanoseconds"]) for row in rows) >= TARGET_SPAN_NS:
                return batch
            span = min(int(row["nanoseconds"]) for row in rows)
            batch = max(batch + 1, math.ceil(batch * TARGET_SPAN_NS / max(1, span) * 1.2))
        raise RunnerError(f"clone calibration exceeded batch limit for {role}")

    def _write_measurement(self, relative: str, value: Any) -> None:
        _write_json(self.output / relative, value, replace=True)

    def _check_span(self, rows: Sequence[Mapping[str, Any]], label: str) -> None:
        if any(int(row["nanoseconds"]) < MIN_SPAN_NS for row in rows):
            raise RunnerError(f"accepted leg is shorter than 250 ms: {label}")

    def _timing_phase(self) -> dict[str, Any]:
        batches: dict[str, Any] = {"clone": {}}
        batches["clone"]["before"] = self._calibrate_clone("before")
        batches["clone"]["after"] = self._calibrate_clone("after")
        batches["clone"]["run"] = max(batches["clone"]["before"], batches["clone"]["after"])
        timing: dict[str, Any] = {"clone": [], "routes": {}, "batches": batches}
        for block in range(BLOCKS):
            legs: list[dict[str, Any]] = []
            for index, role in enumerate(ABBA_ORDER):
                label = f"clone-b{block:02d}-{index}-{role}"
                stdout, event = self._run_clone(role, "normal", label, batches["clone"]["run"])
                rows = self._parse_clone(stdout, 2 * SAMPLES)
                self._check_span(rows, label)
                stats = {
                    size: self._route_stats([row for row in rows if row["size"] == size])
                    for size in CLONE_SIZES
                }
                legs.append({"role": role, "rows": rows, "stats": stats, "event_id": event["event_id"]})
            for size in CLONE_SIZES:
                extents = {tuple(leg["stats"][size]["output_bytes"]) for leg in legs}
                if len(extents) != 1:
                    raise RunnerError(f"clone archive extent differs in block {block}/{size}")
            timing["clone"].append({"block": block, "legs": legs})
            self._write_measurement("measurements.json", {"status": "timing", "timing": timing})
        for route in ROUTES:
            batches[route] = {
                "before": self._calibrate_route("before", route),
                "after": self._calibrate_route("after", route),
            }
            batches[route]["run"] = max(batches[route]["before"], batches[route]["after"])
        for block in range(BLOCKS):
            for route in ROUTES:
                legs = []
                for index, role in enumerate(ABBA_ORDER):
                    label = f"route-{route}-b{block:02d}-{index}-{role}"
                    stdout, event = self._run_probe(role, "normal", label, ["--route", route, "--warmup", str(WARMUP), "--samples", str(SAMPLES), "--batch", str(batches[route]["run"])])
                    rows = self._parse_route(stdout, route, SAMPLES)
                    self._check_span(rows, label)
                    legs.append({"role": role, "rows": rows, "stats": self._route_stats(rows), "event_id": event["event_id"]})
                extents = {tuple(leg["stats"]["output_bytes"]) for leg in legs}
                if len(extents) != 1:
                    raise RunnerError(f"route output extent differs in block {block}/{route}")
                timing["routes"].setdefault(route, []).append({"block": block, "legs": legs})
                self._write_measurement("measurements.json", {"status": "timing", "timing": timing})
        return timing

    def _allocator_phase(self) -> dict[str, Any]:
        allocation: dict[str, Any] = {
            "routes": {},
            "abba_order": list(ABBA_ORDER),
            "elapsed_claimable": False,
            "history_retention": {"doc_steps": 128, "xls_steps": 128, "packages": 129, "patches": 128},
            "allocator_paths": {"deallocation": "record_dealloc", "reallocation": "record_realloc"},
        }
        for route in ALLOCATOR_ROUTES:
            legs = []
            for index, role in enumerate(ABBA_ORDER):
                label = f"alloc-{route}-{index}-{role}"
                stdout, event = self._run_probe(role, "allocator", label, ["--route", route, "--warmup", "0", "--samples", "1", "--batch", "1"])
                report = self._parse_allocator(stdout, route)
                legs.append({"role": role, "report": report, "event_id": event["event_id"]})
            extents = {leg["report"]["output_bytes"] for leg in legs}
            if len(extents) != 1:
                raise RunnerError(f"allocator output extent differs for {route}")
            allocation["routes"][route] = {"legs": legs}
            self._write_measurement("measurements.json", {"status": "allocator", "allocation": allocation})
        return allocation

    def abba(self) -> None:
        try:
            timing = self._timing_phase()
            allocation = self._allocator_phase()
            measurements = {
                "schema": "litchi.ole-owner-qualification.measurements.v1",
                "status": "complete; coordinator review pending",
                "routes": list(ROUTES),
                "abba_order": list(ABBA_ORDER),
                "blocks": BLOCKS,
                "warmup": WARMUP,
                "samples": SAMPLES,
                "minimum_span_ns": MIN_SPAN_NS,
                "timing": timing,
                "allocation": allocation,
                "claim_scope": "raw paired evidence only; no performance acceptance claim",
            }
            self._write_measurement("measurements.json", measurements)
            digest, size = _sha256_file(self.output / "measurements.json")
            self._write_receipt(
                "passed",
                extra={
                    "build_receipt_sha256": self.build_receipt_hash,
                    "measurements": {"path": str(self.output / "measurements.json"), "sha256": digest, "bytes": size},
                },
            )
        except BaseException as error:
            try:
                self._write_receipt("failed", error=str(error))
            except BaseException:
                pass
            raise


def _parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--preparation", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--target-root", type=Path, required=True)
    parser.add_argument("--phase", choices=("build", "abba"), required=True)
    parser.add_argument("--cpu", type=int)
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv or sys.argv[1:])
    try:
        runner = PairedRunner(
            manifest_path=args.manifest,
            preparation_path=args.preparation,
            output=args.output,
            target_root=args.target_root,
            phase=args.phase,
            cpu=args.cpu,
        )
        if args.phase == "build":
            runner.build()
        else:
            runner.abba()
    except (OSError, RunnerError, preparation.PreparationError) as error:
        print(f"error: {error}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
