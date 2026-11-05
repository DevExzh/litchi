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
import struct
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
ALLOCATOR_EVIDENCE_FIELDS = (
    "route",
    "alloc_calls",
    "alloc_bytes",
    "dealloc_calls",
    "dealloc_bytes",
    "realloc_calls",
    "realloc_old_bytes",
    "realloc_new_bytes",
    "base_live_calls",
    "base_live_bytes",
    "live_calls",
    "live_bytes",
    "peak_calls",
    "peak_bytes",
    "retained_packages",
    "retained_patches",
)
EVENT_ENVIRONMENT_KEYS = (
    "CARGO_TARGET_DIR",
    "CARGO_BUILD_JOBS",
    "CARGO_INCREMENTAL",
    "LITCHI_MEM_MAX",
    "PATH",
    "CARGO",
    "RUSTC",
    "RUSTC_WRAPPER",
    "RUSTC_WORKSPACE_WRAPPER",
    "CARGO_BUILD_RUSTC",
    "RUSTUP_TOOLCHAIN",
)
ORACLE_FILES = (
    "ole-source.bin",
    "ole-noop.bin",
    "ole-replacement.bin",
    "properties.tsv",
    "doc-source.bin",
    "doc-edit.bin",
    "doc-proof.tsv",
    "xls-source.bin",
    "xls-edit.bin",
    "xls-proof.tsv",
    "xls-opaque-source.bin",
    "xls-opaque-edit.bin",
    "doc-history-final.bin",
    "xls-history-final.bin",
    "history-proof.tsv",
    "ole-inventory.tsv",
    "oracle.tsv",
    "allocator.tsv",
)
TARGET_NAME = "ole_w5_edit_target.bin"


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
    if allocator.get("oracle_evidence_file") != "allocator.tsv" or allocator.get("oracle_evidence_fields") != list(ALLOCATOR_EVIDENCE_FIELDS):
        raise RunnerError("manifest allocator oracle evidence schema differs")
    if allocator.get("conservation") != (
        "live_calls=base_live_calls+alloc_calls-dealloc_calls; live_bytes=base_live_bytes+alloc_bytes-dealloc_bytes-realloc_old_bytes+realloc_new_bytes; peak is at least live; retained history is 129 packages and 128 patches only on history routes"
    ):
        raise RunnerError("manifest allocator conservation policy differs")
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
        "reviewed_lock_sha256": preparation.REVIEWED_LOCK_SHA256,
        "baseline_lock_sha256": preparation.BASELINE_LOCK_SHA256,
        "candidate_lock_sha256": preparation.CANDIDATE_LOCK_SHA256,
        "oracle_source_sha256": preparation.ORACLE_SOURCE_SHA256,
    }.items():
        if pins.get(key) != expected:
            raise RunnerError(f"manifest input pin differs: {key}")
    if pins.get("reviewed_probe_provenance") != preparation.REVIEWED_PROBE_SOURCE:
        raise RunnerError("manifest reviewed probe provenance differs")
    if pins.get("candidate_adapter_policy") != "CFB limits arity only; no implementation overlay":
        raise RunnerError("manifest candidate adapter policy differs")
    if pins.get("clone_probe_provenance") != preparation.CLONE_PROBE_SOURCE:
        raise RunnerError("manifest clone probe provenance differs")
    if pins.get("lock_provenance") != preparation.LOCK_SOURCE:
        raise RunnerError("manifest lock provenance differs")
    if pins.get("lock_patch_policy") != preparation.LOCK_PATCH_POLICY:
        raise RunnerError("manifest lock patch policy differs")
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
        # The build receipt is the authority for both phases.  ABBA must not
        # rediscover compiler tools from the ambient environment, because a
        # changed RUSTC/CARGO or wrapper would otherwise silently change the
        # code being measured.
        self.toolchain: dict[str, Any] | None = None
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
        if not isinstance(files, list) or len(files) != 5:
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
            "oracle": target / "release" / "examples" / "ole_owner_oracle",
        }

    @staticmethod
    def _path_executable(name: str, env: Mapping[str, str]) -> tuple[Path, Path]:
        """Return the absolute PATH entry and its resolved file identity.

        Rustup installs ``cargo`` and ``rustc`` as symlinked dispatchers.  The
        dispatcher path (including its basename) must be retained when it is
        put in Cargo's environment; invoking the resolved ``rustup`` target as
        ``RUSTC`` would change dispatch semantics.  The resolved target is
        recorded separately and hashed, so both the child command and the
        underlying executable are source-bound.
        """

        executable = shutil.which(name, path=env.get("PATH"))
        if executable is None:
            raise RunnerError(f"{name} is not available in build PATH")
        path = Path(executable)
        if not path.is_absolute():
            path = Path.cwd() / path
        path = path.absolute()
        try:
            resolved = path.resolve(strict=True)
        except OSError as error:
            raise RunnerError(f"{name} PATH executable cannot be resolved: {path}: {error}") from error
        if not resolved.is_file():
            raise RunnerError(f"{name} PATH executable is not a regular file: {resolved}")
        return path, resolved

    def _toolchain(self, env: Mapping[str, str]) -> dict[str, Any]:
        selected = {
            name: self._path_executable(name, env)
            for name in ("cargo", "rustc")
        }
        # Deliberately ignore ambient CARGO/RUSTC and both wrapper variables:
        # PATH selection above is the reviewed input, and these explicit child
        # values are what the build receipt and every event attest.
        version_env = dict(env)
        version_env["CARGO"] = str(selected["cargo"][0])
        version_env["RUSTC"] = str(selected["rustc"][0])
        for key in (
            "RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER", "CARGO_BUILD_RUSTC",
            "RUSTUP_TOOLCHAIN",
        ):
            version_env.pop(key, None)
        values: dict[str, Any] = {"path_env": env.get("PATH", "")}
        if not values["path_env"]:
            raise RunnerError("build PATH is empty")
        for name, args in (("cargo", ["-V"]), ("rustc", ["-Vv"])):
            path, resolved = selected[name]
            digest, size = _sha256_file(resolved)
            result = subprocess.run(
                [str(path), *args],
                cwd=str(self.driver_roots["before"]),
                env=version_env,
                check=False,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            if result.returncode:
                raise RunnerError(f"{name} identity command failed: {result.stderr.strip()}")
            values[name] = {
                "path": str(path),
                "resolved_path": str(resolved),
                "bytes": size,
                "sha256": digest,
                "version": result.stdout,
            }
        return values

    def _environment(self, role: str, variant: str) -> dict[str, str]:
        if self.toolchain is None:
            raise RunnerError("toolchain identity is not loaded")
        env = dict(os.environ)
        env["CARGO_TARGET_DIR"] = str(self._target(role, variant))
        env["CARGO_BUILD_JOBS"] = "1"
        env["CARGO_INCREMENTAL"] = "0"
        env["LITCHI_MEM_MAX"] = "20G"
        env["PATH"] = str(self.toolchain["path_env"])
        env["CARGO"] = str(self.toolchain["cargo"]["path"])
        env["RUSTC"] = str(self.toolchain["rustc"]["path"])
        for key in (
            "RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER", "CARGO_BUILD_RUSTC",
            "RUSTUP_TOOLCHAIN",
        ):
            env.pop(key, None)
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
                key: env.get(key) for key in EVENT_ENVIRONMENT_KEYS
            },
            "toolchain": self.toolchain,
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

    def _build_command(self, role: str, variant: str, *, cargo: str) -> list[str]:
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
            "--example",
            "ole_owner_oracle",
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
            self.toolchain = toolchain
            builds: dict[str, Any] = {}
            for role in ROLES:
                for variant in VARIANTS:
                    key = f"{role}/{variant}"
                    target = self._target(role, variant)
                    target.mkdir()
                    binaries = self._binary_paths(role, variant)
                    command = self._build_command(role, variant, cargo=str(toolchain["cargo"]["path"]))
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
        self.toolchain = dict(_require_mapping(receipt.get("toolchain"), "build receipt toolchain"))
        if not isinstance(self.toolchain.get("path_env"), str) or not self.toolchain["path_env"]:
            raise RunnerError("build receipt toolchain PATH is missing")
        for name in ("cargo", "rustc"):
            record = _require_mapping(self.toolchain.get(name), f"build receipt toolchain.{name}")
            path = _resolve(str(record.get("path", "")))
            resolved = _resolve(str(record.get("resolved_path", "")))
            if path != resolved:
                raise RunnerError(f"build receipt toolchain.{name} resolved path differs")
            digest, size = _sha256_file(path)
            if record.get("sha256") != digest or record.get("bytes") != size:
                raise RunnerError(f"build receipt toolchain.{name} identity changed")
            if not isinstance(record.get("version"), str) or not record["version"].strip():
                raise RunnerError(f"build receipt toolchain.{name} version is missing")
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
                    or event.get("phase") != "build"
                    or event.get("role") != role
                    or event.get("variant") != variant
                    or event.get("argv") != build.get("command")
                    or event.get("cwd") != str(self.driver_roots[role])
                    or event.get("process_guard_violations")
                    or event.get("post_exit_descendants")
                    or event.get("post_exit_offenders")
                ):
                    raise RunnerError(f"build event identity or result differs for {key}")
                expected_environment = {
                    name: self._environment(role, variant).get(name)
                    for name in EVENT_ENVIRONMENT_KEYS
                }
                if event.get("environment") != expected_environment:
                    raise RunnerError(f"build environment differs for {key}")
                if event.get("toolchain") != self.toolchain:
                    raise RunnerError(f"build toolchain identity differs for {key}")
                artifacts = _require_mapping(build.get("artifacts"), f"receipt.builds.{key}.artifacts")
                expected = self._binary_paths(role, variant)
                event_outputs = event.get("outputs")
                if (
                    not isinstance(event_outputs, list)
                    or not all(isinstance(item, Mapping) for item in event_outputs)
                    or [item.get("path") for item in event_outputs] != [str(path) for path in expected.values()]
                ):
                    raise RunnerError(f"build event outputs differ for {key}")
                for name, path in expected.items():
                    item = _require_mapping(artifacts.get(name), f"artifact {key}/{name}")
                    if _resolve(str(item.get("path", ""))) != path:
                        raise RunnerError(f"artifact path differs for {key}/{name}")
                    if not path.is_file() or not os.access(path, os.X_OK):
                        raise RunnerError(f"artifact is missing/non-executable: {path}")
                    digest, size = _sha256_file(path)
                    if item.get("sha256") != digest or item.get("bytes") != size:
                        raise RunnerError(f"artifact hash changed: {path}")
                    event_output = _require_mapping(event_outputs[list(expected).index(name)], f"build event output {key}/{name}")
                    if event_output.get("sha256") != digest or event_output.get("bytes") != size or event_output.get("exists") is not True:
                        raise RunnerError(f"build event artifact hash differs: {path}")
                for stream_name in ("stdout", "stderr"):
                    stream = _require_mapping(event.get(stream_name), f"build event {key}.{stream_name}")
                    stream_path = _resolve(str(stream.get("path", "")))
                    if not _under(stream_path, self.output / "commands"):
                        raise RunnerError(f"build event log escapes commands directory: {stream_path}")
                    digest, size = _sha256_file(stream_path)
                    if stream.get("sha256") != digest or stream.get("bytes") != size:
                        raise RunnerError(f"build event log hash differs: {stream_path}")
                self.artifacts[role][variant] = {name: path for name, path in expected.items()}
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

    def _oracle_paths(self, role: str, variant: str) -> tuple[Path, list[Path]]:
        directory = self.output / "oracle" / f"{role}-{variant}"
        if directory.exists():
            raise RunnerError(f"oracle output already exists: {directory}")
        directory.mkdir(parents=True)
        return directory, [directory / name for name in ORACLE_FILES]

    def _run_oracle(self, role: str, variant: str) -> dict[str, Any]:
        binary = self.artifacts[role][variant]["oracle"]
        directory, expected_outputs = self._oracle_paths(role, variant)
        env = self._environment(role, variant)
        stdout = self.output / "runs" / f"oracle-{role}-{variant}.stdout"
        stderr = self.output / "runs" / f"oracle-{role}-{variant}.stderr"
        event = self._run_process(
            event_id=f"oracle-{role}-{variant}",
            role=role,
            variant=variant,
            argv=[str(binary), "--output", str(directory)],
            cwd=self.driver_roots[role],
            env=env,
            stdout_path=stdout,
            stderr_path=stderr,
            expected_outputs=expected_outputs,
            require_stdout=True,
        )
        actual = {
            path.name: {"path": str(path), "bytes": _sha256_file(path)[1], "sha256": _sha256_file(path)[0]}
            for path in expected_outputs
        }
        try:
            stdout_lines = Path(event["stdout"]["path"]).read_text(encoding="utf-8").splitlines()
        except (OSError, UnicodeDecodeError) as error:
            raise RunnerError(f"cannot read oracle invocation manifest: {directory}") from error
        if not stdout_lines or stdout_lines[0] != f"ole_owner_oracle,{directory}":
            raise RunnerError(f"oracle invocation did not identify its selected output: {directory}")
        invocation: dict[str, tuple[int, str]] = {}
        for line in stdout_lines[1:]:
            fields = line.split("\t")
            if len(fields) != 4 or fields[0] != "oracle_output" or fields[1] in invocation:
                raise RunnerError(f"oracle invocation output manifest is malformed: {directory}")
            try:
                invocation[fields[1]] = (int(fields[2]), fields[3])
            except ValueError as error:
                raise RunnerError(f"oracle invocation output size is not numeric: {directory}") from error
        if set(invocation) != set(ORACLE_FILES):
            raise RunnerError(f"oracle invocation output manifest set differs: {directory}")
        for name in ORACLE_FILES:
            if invocation[name] != (actual[name]["bytes"], actual[name]["sha256"]):
                raise RunnerError(f"oracle invocation output hash differs: {directory}/{name}")
        if {
            path.name
            for path in directory.iterdir()
            if path.is_file() and not path.is_symlink()
        } != set(ORACLE_FILES):
            raise RunnerError(f"oracle output contains an unexpected file: {directory}")
        return {"role": role, "variant": variant, "directory": str(directory), "files": actual, "event_id": event["event_id"]}

    @staticmethod
    def _oracle_metadata(directory: Path) -> dict[str, str]:
        path = directory / "oracle.tsv"
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except (OSError, UnicodeDecodeError) as error:
            raise RunnerError(f"cannot read oracle metadata {path}: {error}") from error
        if lines != [
            "schema\t1",
            "owner_streams\t65",
            f"owner_target\t{TARGET_NAME}",
            "owner_original_bytes\t15",
            "owner_replacement_bytes\t18",
            "properties\tvalidated",
            "doc\trevision\tInsertion\t0\t24\tW5Oracle",
            "xls\tCells\t0\t0\t4242",
        ]:
            raise RunnerError(f"oracle metadata is not the explicit expected-value record: {path}")
        properties = (directory / "properties.tsv").read_text(encoding="utf-8").splitlines()
        if properties != [
            "link_base\thttps://example.invalid/base/",
            "hyperlinks\t1",
            "hyperlink\t7\t11\t0\thttps://example.invalid/target\tbookmark",
            "opaque_property\t42\t128",
        ]:
            raise RunnerError(f"oracle property values are not the explicit expected record: {directory}")
        inventory = (directory / "ole-inventory.tsv").read_text(encoding="utf-8").splitlines()
        if not inventory or inventory[0].split("\t") != [
            "kind", "name", "entry_type", "sid", "sid_left", "sid_right", "sid_child",
            "clsid", "state_bits", "creation_time", "modified_time", "size", "payload_hex",
        ]:
            raise RunnerError(f"oracle OLE inventory header differs: {directory}")
        if len(inventory) != 1 + 2 * (64 + 2):
            raise RunnerError(f"oracle OLE inventory cardinality differs: {directory}")
        rows: dict[str, dict[str, tuple[str, ...]]] = {"source": {}, "replacement": {}}
        for line in inventory[1:]:
            fields = line.split("\t")
            if len(fields) != 13 or fields[0] not in rows or fields[1] in rows[fields[0]]:
                raise RunnerError(f"oracle OLE inventory row differs: {line!r}")
            rows[fields[0]][fields[1]] = tuple(fields[2:])
        if len(rows["source"]) != 66 or len(rows["replacement"]) != 66:
            raise RunnerError(f"oracle OLE inventory does not include root/storage and all streams: {directory}")
        root_name = "Root Entry"
        for kind, side in rows.items():
            if root_name not in side or side[root_name][0] != "5":
                raise RunnerError(f"oracle OLE inventory omits the {kind} Root Entry storage identity: {directory}")
        source_nonstreams = {name for name, fields in rows["source"].items() if fields[0] != "2"}
        replacement_nonstreams = {name for name, fields in rows["replacement"].items() if fields[0] != "2"}
        if source_nonstreams != replacement_nonstreams or not source_nonstreams:
            raise RunnerError(f"oracle OLE inventory omitted non-stream metadata: {directory}")
        if any(rows["source"][name] != rows["replacement"][name] for name in source_nonstreams):
            raise RunnerError(f"oracle OLE non-stream metadata changed: {directory}")
        expected_streams = {f"ole_w5_sibling_{index:03}" for index in range(64)} | {TARGET_NAME}
        if not expected_streams.issubset(rows["source"]) or not expected_streams.issubset(rows["replacement"]):
            raise RunnerError(f"oracle OLE inventory omitted an expected stream: {directory}")
        for name in expected_streams:
            before = rows["source"][name]
            after = rows["replacement"][name]
            if any(fields[-2] != str(len(bytes.fromhex(fields[-1]))) for fields in (before, after)):
                raise RunnerError(f"oracle stream size differs from its payload: {directory}/{name}")
            if name == TARGET_NAME:
                if bytes.fromhex(before[-1]) != b"ole-w5-original" or bytes.fromhex(after[-1]) != b"ole-w5-replacement":
                    raise RunnerError(f"oracle replacement target proof differs: {directory}")
            elif before != after or bytes.fromhex(before[-1]) != f"opaque sibling payload {name[-3:]}".encode():
                raise RunnerError(f"oracle untouched sibling proof differs: {name}")
        doc_proof = (directory / "doc-proof.tsv").read_text(encoding="utf-8").splitlines()
        if len(doc_proof) != 1 + 2 * 256 or doc_proof[0] != "kind\tindex\ttext":
            raise RunnerError(f"oracle DOC paragraph proof cardinality differs: {directory}")
        doc_indices: dict[str, list[int]] = {"source": [], "edit": []}
        for line in doc_proof[1:]:
            fields = line.split("\t")
            if len(fields) != 3 or fields[0] not in {"source", "edit"} or not fields[1].isdigit():
                raise RunnerError(f"oracle DOC paragraph proof row differs: {line!r}")
            index = int(fields[1])
            doc_indices[fields[0]].append(index)
            expected = f"ole-w5 tracked paragraph {index:03}"
            if fields[0] == "edit" and index == 0:
                expected = f"ole-w5-tracked-revision {expected}"
            if index >= 256 or fields[2] != expected:
                raise RunnerError(f"oracle DOC paragraph proof value differs: {line!r}")
        if any(doc_indices[kind] != list(range(256)) for kind in ("source", "edit")):
            raise RunnerError(f"oracle DOC paragraph indices are incomplete or duplicated: {directory}")
        xls_proof = (directory / "xls-proof.tsv").read_text(encoding="utf-8").splitlines()
        if len(xls_proof) != 1 + 2 * 256 + 1 or xls_proof[0] != "kind\trow\tvalue_bits":
            raise RunnerError(f"oracle XLS cell proof cardinality differs: {directory}")
        xls_indices: dict[str, list[int]] = {"source": [], "edit": []}
        for line in xls_proof[1:-1]:
            fields = line.split("\t")
            if len(fields) != 3 or fields[0] not in {"source", "edit"} or not fields[1].isdigit() or not fields[2].isdigit():
                raise RunnerError(f"oracle XLS cell proof row differs: {line!r}")
            row, bits = int(fields[1]), int(fields[2])
            xls_indices[fields[0]].append(row)
            # Compare the serialized IEEE bits without trusting a decimal
            # rendering.  The explicit values have stable bit patterns.
            expected_value = float(row) if fields[0] == "source" or row != 0 else 4242.0
            expected_bits = struct.unpack("<Q", struct.pack("<d", expected_value))[0]
            if row >= 256 or bits != expected_bits:
                raise RunnerError(f"oracle XLS cell proof value differs: {line!r}")
        if any(xls_indices[kind] != list(range(256)) for kind in ("source", "edit")):
            raise RunnerError(f"oracle XLS cell indices are incomplete or duplicated: {directory}")
        if xls_proof[-1] != "opaque\tbytes\t98304":
            raise RunnerError(f"oracle XLS opaque proof differs: {directory}")
        opaque_source = directory / "xls-opaque-source.bin"
        opaque_output = directory / "xls-opaque-edit.bin"
        if opaque_source.stat().st_size != 98304 or opaque_output.read_bytes() != opaque_source.read_bytes():
            raise RunnerError(f"oracle XLS opaque stream proof differs: {directory}")
        history = (directory / "history-proof.tsv").read_text(encoding="utf-8").splitlines()
        if history != [
            "kind\tpackages\tpatches\tforward\tinverse",
            "doc\t129\t128\tvalidated\tvalidated",
            "xls\t129\t128\tvalidated\tvalidated",
        ]:
            raise RunnerError(f"oracle retained-history proof differs: {directory}")
        return {line.split("\t", 1)[0]: line for line in lines}

    @staticmethod
    def _parse_oracle_allocator(path: Path) -> list[dict[str, int | str]]:
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except (OSError, UnicodeDecodeError) as error:
            raise RunnerError(f"cannot read oracle allocator evidence {path}: {error}") from error
        if lines == ["status\tnative-uninstrumented"]:
            raise RunnerError("allocator oracle must be the instrumented consumer variant")
        if not lines or lines[0].split("\t") != list(ALLOCATOR_EVIDENCE_FIELDS):
            raise RunnerError(f"oracle allocator header differs: {path}")
        if len(lines) != len(ALLOCATOR_ROUTES) + 1:
            raise RunnerError(f"oracle allocator route count differs: {path}")
        rows: list[dict[str, int | str]] = []
        for line in lines[1:]:
            fields = line.split("\t")
            if len(fields) != len(ALLOCATOR_EVIDENCE_FIELDS) or fields[0] not in ALLOCATOR_ROUTES:
                raise RunnerError(f"oracle allocator route row differs: {line!r}")
            row: dict[str, int | str] = {"route": fields[0]}
            try:
                row.update({key: int(value) for key, value in zip(ALLOCATOR_EVIDENCE_FIELDS[1:], fields[1:])})
            except ValueError as error:
                raise RunnerError(f"oracle allocator counter is not an integer: {line!r}") from error
            if any(int(row[key]) < 0 for key in ALLOCATOR_EVIDENCE_FIELDS[1:]):
                raise RunnerError(f"oracle allocator contains a negative counter: {line!r}")
            expected_calls = int(row["base_live_calls"]) + int(row["alloc_calls"]) - int(row["dealloc_calls"])
            expected_bytes = (
                int(row["base_live_bytes"]) + int(row["alloc_bytes"]) - int(row["dealloc_bytes"])
                - int(row["realloc_old_bytes"]) + int(row["realloc_new_bytes"])
            )
            if int(row["live_calls"]) != expected_calls or int(row["live_bytes"]) != expected_bytes:
                raise RunnerError(f"oracle allocator live-state conservation failed: {line!r}")
            if int(row["peak_calls"]) < int(row["live_calls"]) or int(row["peak_bytes"]) < int(row["live_bytes"]):
                raise RunnerError(f"oracle allocator peak is below live state: {line!r}")
            expected_retained = (129, 128) if str(row["route"]) in {"doc_history", "xls_history"} else (0, 0)
            if (int(row["retained_packages"]), int(row["retained_patches"])) != expected_retained:
                raise RunnerError(f"oracle allocator retained history differs: {line!r}")
            rows.append(row)
        if [str(row["route"]) for row in rows] != list(ALLOCATOR_ROUTES):
            raise RunnerError("oracle allocator routes are missing or reordered")
        return rows

    def _oracle_phase(self) -> dict[str, Any]:
        reports: dict[str, Any] = {}
        for role in ROLES:
            for variant in VARIANTS:
                report = self._run_oracle(role, variant)
                directory = _resolve(str(report["directory"]))
                self._oracle_metadata(directory)
                if variant == "allocator":
                    report["allocation"] = self._parse_oracle_allocator(directory / "allocator.tsv")
                else:
                    try:
                        marker = (directory / "allocator.tsv").read_text(encoding="utf-8").splitlines()
                    except (OSError, UnicodeDecodeError) as error:
                        raise RunnerError(f"cannot read native allocator marker {directory}: {error}") from error
                    if marker != ["status\tnative-uninstrumented"]:
                        raise RunnerError(f"normal oracle unexpectedly contains allocator instrumentation: {directory}")
                reports[f"{role}/{variant}"] = report
        required_equal = tuple(name for name in ORACLE_FILES if name != "allocator.tsv")
        baseline = reports["before/normal"]["files"]
        for key in ("before/allocator", "after/normal", "after/allocator"):
            selected = reports[key]["files"]
            for name in required_equal:
                if selected[name]["sha256"] != baseline[name]["sha256"] or selected[name]["bytes"] != baseline[name]["bytes"]:
                    raise RunnerError(f"oracle cross-endpoint output digest differs for {name}: {key}")
        for key in ("before/normal", "before/allocator", "after/normal", "after/allocator"):
            selected = reports[key]["files"]
            if selected["ole-source.bin"]["sha256"] != selected["ole-noop.bin"]["sha256"]:
                raise RunnerError(f"oracle no-op digest differs from source: {key}")
        return {
            "status": "preflight-validated",
            "files": reports,
            "cross_endpoint_digest_files": list(required_equal),
            "timing_boundary": "oracle runs before calibration and ABBA; no oracle checks occur inside timed probe loops",
            "claim_scope": "correctness and identity preflight only; no performance acceptance claim",
        }

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
            oracle = self._oracle_phase()
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
                "oracle": oracle,
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
