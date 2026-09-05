#!/usr/bin/env python3
"""Verify and project the complete 0418 PPTX evidence bundle.

Every preflight and formal record is checked before deterministic summary and
allocation/RSS projections are published.  Manifest paths remain relative to
the capture root; absolute paths inside captured provenance are opaque.  A
missing raw artifact may be supplied by a neighboring zstd or zstd sidecar.
"""

from __future__ import annotations

import argparse
import contextvars
import datetime as dt
import hashlib
import importlib
import json
import math
import os
import re
import selectors
import subprocess
import sys
import time
from pathlib import Path, PureWindowsPath
from typing import Any


CHANGE_ROOT = Path(__file__).resolve().parent


def find_repo_root() -> Path | None:
    here = Path(__file__).resolve()
    for candidate in (here, *here.parents, Path.cwd().resolve()):
        if (candidate / "tools" / "perf_abba_summary.py").is_file():
            return candidate
    return None


REPO_ROOT = find_repo_root()
if REPO_ROOT is not None and str(REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(REPO_ROOT))
EXPECTED_CHANGE = 418
ABBA_LEGS = ("A1", "B1", "B2", "A2")
ABBA_ROLES = ("control", "candidate", "candidate", "control")
LEG_TO_ROLE = dict(zip(ABBA_LEGS, ABBA_ROLES))
SHA256 = re.compile(r"^[0-9a-f]{64}$")
REVISION = re.compile(r"^[0-9a-f]{40}$")
SELECTOR = re.compile(r"^[A-Za-z0-9_.-]+$")
EXPECTED_FLAGS = [
    "--shape", "many-small", "--payload", "compressible",
    "--writer-shape", "large", "--xlsx-shape", "medium",
    "--xlsx-cell-crud-shape", "medium",
    "--xlsx-row-visibility-shape", "medium",
    "--semantic-shape", "medium", "--workers", "1",
    "--filesystem-cache", "warm",
]
ALLOCATION_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after",
)
STATISTICS = ("p50", "mean", "p95", "p99")
REVIEW_THRESHOLD_PERCENT = 5
MAX_DECOMPRESSED_MEMBER_BYTES = 512 * 1024 * 1024
MAX_DECOMPRESSED_TOTAL_BYTES = 2 * 1024 * 1024 * 1024
ZSTD_DECOMPRESSION_TIMEOUT_SECONDS = 120
ZSTD_READ_BLOCK_BYTES = 1024 * 1024
TIME_LABELS = (
    "User time (seconds)",
    "System time (seconds)",
    "Elapsed (wall clock) time (h:mm:ss or m:ss)",
    "Maximum resident set size (kbytes)",
    "Major (requiring I/O) page faults",
    "Minor (reclaiming a frame) page faults",
    "Voluntary context switches",
    "Involuntary context switches",
)
TIME_VALUE_PATTERNS = {
    "User time (seconds)": r"(?:0|[0-9]+(?:\.[0-9]+)?)",
    "System time (seconds)": r"(?:0|[0-9]+(?:\.[0-9]+)?)",
    "Elapsed (wall clock) time (h:mm:ss or m:ss)": (
        r"(?:[0-9]+:)?[0-9]+:[0-9]{2}(?:\.[0-9]+)?"
    ),
    "Maximum resident set size (kbytes)": r"[0-9]+(?:,[0-9]+)*",
    "Major (requiring I/O) page faults": r"[0-9]+",
    "Minor (reclaiming a frame) page faults": r"[0-9]+",
    "Voluntary context switches": r"[0-9]+",
    "Involuntary context switches": r"[0-9]+",
}


class VerificationError(ValueError):
    """A captured artifact or projection failed the 0418 contract."""


def fail(path: str, message: str) -> None:
    raise VerificationError(f"{path}: {message}")


class ArtifactBudget:
    """Bound decoded artifact bytes retained by one verification run."""

    def __init__(self) -> None:
        self.total_bytes = 0

    def member_limit(self, path: str) -> int:
        remaining = MAX_DECOMPRESSED_TOTAL_BYTES - self.total_bytes
        if remaining <= 0:
            fail(path, "decompressed artifact aggregate exceeds 2 GiB")
        return min(MAX_DECOMPRESSED_MEMBER_BYTES, remaining)

    def add(self, size: int, path: str) -> None:
        if size > MAX_DECOMPRESSED_MEMBER_BYTES:
            fail(path, "decompressed artifact exceeds 512 MiB member ceiling")
        if self.total_bytes + size > MAX_DECOMPRESSED_TOTAL_BYTES:
            fail(path, "decompressed artifact aggregate exceeds 2 GiB")
        self.total_bytes += size


_ACTIVE_ARTIFACT_BUDGET: contextvars.ContextVar[ArtifactBudget | None] = (
    contextvars.ContextVar("active_artifact_budget", default=None)
)


def strict_pairs(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key {key!r}")
        result[key] = value
    return result


def reject_constant(value: str) -> None:
    raise ValueError(f"non-finite JSON value {value!r}")


def canonical(value: Any) -> str:
    return json.dumps(
        value, sort_keys=True, separators=(",", ":"),
        ensure_ascii=False, allow_nan=False,
    )


def digest_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def digest_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def digest_json(value: Any) -> str:
    return digest_bytes(canonical(value).encode("utf-8"))


def load_json_bytes(value: bytes, label: str) -> Any:
    try:
        return json.loads(
            value.decode("utf-8"),
            object_pairs_hook=strict_pairs,
            parse_constant=reject_constant,
        )
    except (UnicodeDecodeError, json.JSONDecodeError, ValueError) as error:
        fail(label, f"invalid strict JSON: {error}")
    raise AssertionError("unreachable")


def string(value: Any, path: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        fail(path, "must be a non-empty string" if nonempty else "must be a string")
    return value


def integer(value: Any, path: str, *, minimum: int = 0) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < minimum:
        fail(path, f"must be an integer >= {minimum}")
    return value


def positive_integer(value: Any, path: str) -> int:
    return integer(value, path, minimum=1)


def boolean(value: Any, path: str) -> bool:
    if not isinstance(value, bool):
        fail(path, "must be a boolean")
    return value


def mapping(value: Any, path: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(path, "must be an object")
    return value


def sequence(value: Any, path: str) -> list[Any]:
    if not isinstance(value, list):
        fail(path, "must be a list")
    return value


def hash_value(value: Any, path: str) -> str:
    value = string(value, path)
    if SHA256.fullmatch(value) is None:
        fail(path, "must be a lowercase SHA-256")
    return value


def same(left: Any, right: Any, path: str) -> None:
    if canonical(left) != canonical(right):
        fail(path, "does not match the bound value")


def relative_manifest_path(value: Any, path: str) -> str:
    value = string(value, path)
    if Path(value).is_absolute() or PureWindowsPath(value).is_absolute():
        fail(path, "must be relative")
    if "\\" in value:
        fail(path, "must use normalized POSIX separators")
    parts = Path(value).parts
    if not parts or any(part in ("", ".", "..") for part in parts):
        fail(path, "must be normalized and must not contain '.' or '..'")
    normalized = Path(*parts).as_posix()
    if normalized != value:
        fail(path, f"must be normalized as {normalized!r}")
    return normalized


def portable_path(value: Any, expected: str, path: str) -> None:
    value = string(value, path).replace("\\", "/")
    expected = expected.replace("\\", "/")
    if value != expected and not value.endswith("/" + expected):
        fail(path, f"must be {expected!r} or an absolute path ending in it")


def parse_timestamp(value: Any, path: str) -> dt.datetime:
    value = string(value, path)
    try:
        result = dt.datetime.fromisoformat(value)
    except ValueError as error:
        fail(path, f"invalid ISO-8601 timestamp: {error}")
    if result.tzinfo is None or result.utcoffset() is None:
        fail(path, "must include an explicit UTC offset")
    return result


class Artifact:
    """Read one manifest artifact, optionally through a zstd sidecar."""

    def __init__(
        self, root: Path, relative: str, label: str, *, nonempty: bool = True
    ):
        self.root = root
        self.relative = relative_manifest_path(relative, label)
        budget = _ACTIVE_ARTIFACT_BUDGET.get() or ArtifactBudget()
        raw = root / self.relative
        self.path: Path | None = None
        self.compressed = False
        if raw.is_file():
            self.path = raw
            self.data = self._read_file(raw, label, budget)
        else:
            compressed: Path | None = None
            if raw.name.endswith((".zst", ".zstd")):
                if raw.is_file():
                    compressed = raw
            else:
                for suffix in (".zst", ".zstd"):
                    candidate = Path(str(raw) + suffix)
                    if candidate.is_file():
                        compressed = candidate
                        break
            if compressed is None:
                fail(label, "artifact is missing")
            self.path = compressed
            self.compressed = True
            self.data = self._decode_zstd(compressed, label, budget)
        if nonempty and not self.data:
            fail(label, "artifact must not be empty")
        self.sha256 = digest_bytes(self.data)
        self.bytes = len(self.data)

    @staticmethod
    def _read_file(path: Path, label: str, budget: ArtifactBudget) -> bytes:
        limit = budget.member_limit(label)
        try:
            declared_size = path.stat().st_size
        except OSError as error:
            fail(label, f"cannot stat artifact: {error}")
        if declared_size > limit:
            if limit < MAX_DECOMPRESSED_MEMBER_BYTES:
                fail(label, "decompressed artifact aggregate exceeds 2 GiB")
            fail(label, "decompressed artifact exceeds 512 MiB member ceiling")
        try:
            with path.open("rb") as stream:
                data = stream.read(limit + 1)
        except OSError as error:
            fail(label, f"cannot read artifact: {error}")
        if len(data) > limit:
            if limit < MAX_DECOMPRESSED_MEMBER_BYTES:
                fail(label, "decompressed artifact aggregate exceeds 2 GiB")
            fail(label, "decompressed artifact exceeds 512 MiB member ceiling")
        budget.add(len(data), label)
        return data

    @staticmethod
    def _decode_zstd(path: Path, label: str, budget: ArtifactBudget) -> bytes:
        limit = budget.member_limit(label)
        try:
            process = subprocess.Popen(
                ["zstd", "-q", "-d", "-c", str(path)],
                stdout=subprocess.PIPE,
                stderr=subprocess.DEVNULL,
            )
        except OSError as error:
            fail(label, f"cannot decode zstd sidecar: {error}")
        assert process.stdout is not None
        output = process.stdout
        poller = selectors.DefaultSelector()
        poller.register(output, selectors.EVENT_READ)
        chunks: list[bytes] = []
        total = 0
        deadline = time.monotonic() + ZSTD_DECOMPRESSION_TIMEOUT_SECONDS

        def reap() -> None:
            if process.poll() is None:
                try:
                    process.kill()
                except OSError:
                    pass
            try:
                process.wait(timeout=5)
            except (OSError, subprocess.TimeoutExpired):
                pass

        try:
            while poller.get_map():
                remaining_time = deadline - time.monotonic()
                if remaining_time <= 0:
                    reap()
                    fail(label, "zstd sidecar decompression timed out")
                events = poller.select(remaining_time)
                if not events:
                    reap()
                    fail(label, "zstd sidecar decompression timed out")
                for key, _ in events:
                    remaining_bytes = limit - total
                    block = os.read(
                        key.fd,
                        min(ZSTD_READ_BLOCK_BYTES, remaining_bytes + 1),
                    )
                    if not block:
                        poller.unregister(key.fileobj)
                        continue
                    total += len(block)
                    chunks.append(block)
                    if total > limit:
                        reap()
                        if limit < MAX_DECOMPRESSED_MEMBER_BYTES:
                            fail(label, "decompressed artifact aggregate exceeds 2 GiB")
                        fail(label, "decompressed artifact exceeds 512 MiB member ceiling")
            remaining_time = max(0.0, deadline - time.monotonic())
            try:
                return_code = process.wait(timeout=remaining_time)
            except subprocess.TimeoutExpired:
                reap()
                fail(label, "zstd sidecar decompression timed out")
        except VerificationError:
            reap()
            raise
        except OSError as error:
            reap()
            fail(label, f"cannot read zstd sidecar: {error}")
        finally:
            poller.close()
        if return_code != 0:
            fail(label, f"zstd sidecar failed with exit status {return_code}")
        data = b"".join(chunks)
        budget.add(len(data), label)
        return data

    def json(self, label: str) -> Any:
        return load_json_bytes(self.data, label)

    def text(self, label: str) -> str:
        try:
            return self.data.decode("utf-8")
        except UnicodeDecodeError as error:
            fail(label, f"must be UTF-8 text: {error}")
        raise AssertionError("unreachable")


def document(root: Path, relative: str, label: str) -> tuple[Artifact, dict[str, Any]]:
    artifact = Artifact(root, relative, label)
    value = mapping(artifact.json(label), label)
    return artifact, value


def validate_source(value: Any, path: str) -> dict[str, Any]:
    item = mapping(value, path)
    required = {"worktree", "revision", "git_status_porcelain", "clean", "source_files"}
    if set(item) != required:
        fail(path, f"must contain exactly {sorted(required)!r}")
    worktree = string(item["worktree"], f"{path}.worktree")
    if not (Path(worktree).is_absolute() or PureWindowsPath(worktree).is_absolute()):
        fail(f"{path}.worktree", "must be absolute")
    revision = string(item["revision"], f"{path}.revision")
    if REVISION.fullmatch(revision) is None:
        fail(f"{path}.revision", "must be a 40-character lowercase revision")
    if item["git_status_porcelain"] != "" or item["clean"] is not True:
        fail(path, "source must be clean")
    files = sequence(item["source_files"], f"{path}.source_files")
    if not files:
        fail(f"{path}.source_files", "must not be empty")
    normalized: list[dict[str, Any]] = []
    names: list[str] = []
    for index, raw in enumerate(files):
        entry = mapping(raw, f"{path}.source_files[{index}]")
        if set(entry) != {"path", "bytes", "sha256"}:
            fail(f"{path}.source_files[{index}]", "has an unexpected schema")
        name = relative_manifest_path(entry["path"], f"{path}.source_files[{index}].path")
        size = positive_integer(entry["bytes"], f"{path}.source_files[{index}].bytes")
        digest = hash_value(entry["sha256"], f"{path}.source_files[{index}].sha256")
        names.append(name)
        normalized.append({"path": name, "bytes": size, "sha256": digest})
    if names != sorted(names) or len(names) != len(set(names)):
        fail(f"{path}.source_files", "paths must be unique and sorted")
    return {
        "worktree": worktree, "revision": revision,
        "git_status_porcelain": "", "clean": True,
        "source_files": normalized,
    }


def validate_binary(value: Any, path: str) -> dict[str, Any]:
    item = mapping(value, path)
    required = {
        "label", "path", "sha256", "binary_sha256", "bytes",
        "binary_bytes", "mode_bits", "executable", "profile",
    }
    if set(item) != required:
        fail(path, f"must contain exactly {sorted(required)!r}")
    label = string(item["label"], f"{path}.label")
    binary_path = string(item["path"], f"{path}.path")
    if not (Path(binary_path).is_absolute() or PureWindowsPath(binary_path).is_absolute()):
        fail(f"{path}.path", "must be absolute")
    digest = hash_value(item["sha256"], f"{path}.sha256")
    if item["binary_sha256"] != digest:
        fail(f"{path}.binary_sha256", "must equal sha256")
    size = positive_integer(item["bytes"], f"{path}.bytes")
    if item["binary_bytes"] != size:
        fail(f"{path}.binary_bytes", "must equal bytes")
    mode = integer(item["mode_bits"], f"{path}.mode_bits")
    if mode > 0o7777 or mode & 0o111 == 0:
        fail(f"{path}.mode_bits", "must contain executable permission bits")
    if item["executable"] is not True:
        fail(f"{path}.executable", "must be true")
    profile = string(item["profile"], f"{path}.profile")
    if profile != "release-debug1-frame-pointer":
        fail(f"{path}.profile", "must identify the release frame-pointer build")
    if Path(binary_path).is_file():
        actual_path = Path(binary_path)
        if digest_file(actual_path) != digest or actual_path.stat().st_size != size:
            fail(f"{path}.path", "on-disk binary identity differs")
    return {
        "label": label, "path": binary_path, "sha256": digest,
        "binary_sha256": digest, "bytes": size, "binary_bytes": size,
        "mode_bits": mode, "executable": True, "profile": profile,
    }


def validate_protocol(root: Path) -> tuple[Artifact, dict[str, Any], list[dict[str, Any]]]:
    artifact, protocol = document(root, "protocol.json", "protocol.json")
    if protocol.get("change") != EXPECTED_CHANGE:
        fail("protocol.change", f"must be {EXPECTED_CHANGE}")
    if protocol.get("order") != list(ABBA_ROLES):
        fail("protocol.order", f"must be {list(ABBA_ROLES)!r}")
    if protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol", "must pin cpu=2 and workers=1")
    jobs = sequence(protocol.get("jobs"), "protocol.jobs")
    if len(jobs) != 4:
        fail("protocol.jobs", "must contain exactly four selectors")
    normalized: list[dict[str, Any]] = []
    names: set[str] = set()
    for index, raw in enumerate(jobs):
        item = mapping(raw, f"protocol.jobs[{index}]")
        if set(item) != {"selector", "role", "samples", "warmups"}:
            fail(f"protocol.jobs[{index}]", "has an unexpected schema")
        selector = string(item["selector"], f"protocol.jobs[{index}].selector")
        if SELECTOR.fullmatch(selector) is None or selector in names:
            fail(f"protocol.jobs[{index}].selector", "is unsafe or duplicated")
        names.add(selector)
        role = string(item["role"], f"protocol.jobs[{index}].role")
        samples = positive_integer(item["samples"], f"protocol.jobs[{index}].samples")
        warmups = integer(item["warmups"], f"protocol.jobs[{index}].warmups")
        expected = (100, 10) if selector == "pptx_cross_copy_media_rich" else (500, 20)
        if (samples, warmups) != expected:
            fail(f"protocol.jobs[{index}]", f"must use counts {expected!r}")
        normalized.append({
            "selector": selector, "role": role,
            "samples": samples, "warmups": warmups,
        })
    expected_names = {
        "pptx_cross_copy_media_rich_lifecycle",
        "pptx_cross_copy_plain_lifecycle",
        "pptx_cross_copy_plain",
        "pptx_cross_copy_media_rich",
    }
    if names != expected_names:
        fail("protocol.jobs", f"selectors must be {sorted(expected_names)!r}")
    allocator = mapping(protocol.get("allocator"), "protocol.allocator")
    if allocator.get("samples") != 30 or allocator.get("warmups") != 3:
        fail("protocol.allocator", "must use samples=30 and warmups=3")
    if allocator.get("order") != "same ABBA":
        fail("protocol.allocator.order", "must be 'same ABBA'")
    if allocator.get("latency_comparison") != "excluded":
        fail("protocol.allocator.latency_comparison", "must be excluded")
    if allocator.get("selectors") != [job["selector"] for job in normalized]:
        fail("protocol.allocator.selectors", "must cover all four selectors in job order")
    if protocol.get("common_flags") != EXPECTED_FLAGS:
        fail("protocol.common_flags", "does not match frozen command flags")
    profile = mapping(protocol.get("profile"), "protocol.profile")
    expected_profile = {
        "samples": 20, "warmups": 3, "event": "cycles:u",
        "frequency": 999, "call_graph": "fp,127",
        "scope": "wholecommand plus descendants; no phaseelapsed percentages",
    }
    for key, expected in expected_profile.items():
        if profile.get(key) != expected:
            fail(f"protocol.profile.{key}", f"must be {expected!r}")
    latency = mapping(protocol.get("latency_policy"), "protocol.latency_policy")
    ceilings = mapping(latency.get("drift_ceilings_percent"), "protocol.latency_policy.drift_ceilings_percent")
    if ceilings != {name: 5 for name in STATISTICS}:
        fail("protocol.latency_policy.drift_ceilings_percent", "must be five percent for all statistics")
    if "no pooled speedup claim" not in str(latency.get("acceptance", "")):
        fail("protocol.latency_policy.acceptance", "must retain the no pooled speedup claim")
    memory = mapping(protocol.get("memory_policy"), "protocol.memory_policy")
    if memory.get("metrics") != [
        "allocation_calls", "allocated_bytes", "live_bytes_before",
        "live_bytes_after", "peak_live_bytes_before",
        "peak_live_bytes_after", "whole_process_max_rss",
    ]:
        fail("protocol.memory_policy.metrics", "does not retain required metrics")
    if protocol.get("runtime_toolchain") != "RUSTUP_TOOLCHAIN=1.98.1 set during build and capture":
        fail("protocol.runtime_toolchain", "does not bind toolchain 1.98.1")
    return artifact, protocol, normalized


def validate_roles(root: Path) -> tuple[Artifact, dict[str, Any]]:
    artifact, roles = document(root, "roles.json", "roles.json")
    if set(roles) != {"control", "candidate"}:
        fail("roles.json", "must contain exactly control and candidate")
    normalized: dict[str, Any] = {}
    for role in ("control", "candidate"):
        item = mapping(roles[role], f"roles.{role}")
        if set(item) != {"revision", "worktree"}:
            fail(f"roles.{role}", "has an unexpected schema")
        revision = string(item["revision"], f"roles.{role}.revision")
        if REVISION.fullmatch(revision) is None:
            fail(f"roles.{role}.revision", "must be a lowercase 40-character revision")
        worktree = string(item["worktree"], f"roles.{role}.worktree")
        if not (Path(worktree).is_absolute() or PureWindowsPath(worktree).is_absolute()):
            fail(f"roles.{role}.worktree", "must be absolute")
        normalized[role] = {"revision": revision, "worktree": worktree}
    return artifact, normalized


def validate_build(
    root: Path, protocol_artifact: Artifact, roles: dict[str, Any]
) -> tuple[Artifact, dict[str, Any]]:
    artifact, build = document(root, "build-identity.json", "build-identity.json")
    if build.get("schema_version") != 1 or build.get("change") != EXPECTED_CHANGE:
        fail("build-identity", "must be schema 1 for change 0418")
    if build.get("status") != "complete":
        fail("build-identity.status", "must be complete")
    protocol = mapping(build.get("protocol"), "build-identity.protocol")
    if hash_value(protocol.get("sha256"), "build-identity.protocol.sha256") != protocol_artifact.sha256:
        fail("build-identity.protocol.sha256", "does not match protocol.json")
    portable_path(protocol.get("path"), "protocol.json", "build-identity.protocol.path")
    environment = mapping(build.get("environment"), "build-identity.environment")
    for key, expected in {
        "RUSTUP_TOOLCHAIN": "1.98.1",
        "CARGO_BUILD_JOBS": "4",
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "RUSTFLAGS": "-C force-frame-pointers=yes -C force-unwind-tables=yes",
    }.items():
        if environment.get(key) != expected:
            fail(f"build-identity.environment.{key}", f"must be {expected!r}")
    if not isinstance(environment.get("CARGO_TARGET_DIR"), str) or not environment["CARGO_TARGET_DIR"]:
        fail("build-identity.environment.CARGO_TARGET_DIR", "must be non-empty")
    if build.get("common_flags") != EXPECTED_FLAGS:
        fail("build-identity.common_flags", "does not match protocol common_flags")
    build_roles = mapping(build.get("roles"), "build-identity.roles")
    if set(build_roles) != {"control", "candidate"}:
        fail("build-identity.roles", "must contain exactly control and candidate")
    normalized_roles: dict[str, Any] = {}
    for role in ("control", "candidate"):
        item = mapping(build_roles[role], f"build-identity.roles.{role}")
        if item.get("role") != role:
            fail(f"build-identity.roles.{role}.role", f"must be {role!r}")
        source = validate_source(item.get("source"), f"build-identity.roles.{role}.source")
        if source["revision"] != roles[role]["revision"] or source["worktree"] != roles[role]["worktree"]:
            fail(f"build-identity.roles.{role}.source", "does not match roles.json")
        if item.get("source_after") != source:
            fail(f"build-identity.roles.{role}.source_after", "must equal post-build source")
        if item.get("exit_code") != 0:
            fail(f"build-identity.roles.{role}.exit_code", "must be zero")
        binaries = mapping(item.get("binaries"), f"build-identity.roles.{role}.binaries")
        if set(binaries) != {"normal", "allocator"}:
            fail(f"build-identity.roles.{role}.binaries", "must contain normal and allocator")
        normalized_binaries: dict[str, Any] = {}
        for phase in ("normal", "allocator"):
            binary = validate_binary(
                binaries[phase], f"build-identity.roles.{role}.binaries.{phase}"
            )
            if binary["label"] != f"{role}/{phase}":
                fail(f"build-identity.roles.{role}.binaries.{phase}.label", "does not bind role and phase")
            normalized_binaries[phase] = binary
        normalized_roles[role] = {
            "role": role, "source": source, "binaries": normalized_binaries,
        }
    return artifact, {"raw": build, "roles": normalized_roles}


def expected_plan(
    jobs: list[dict[str, Any]], protocol: dict[str, Any]
) -> list[dict[str, Any]]:
    selectors = [job["selector"] for job in jobs]
    ordered: list[dict[str, Any]] = []
    for preflight, phase, allocator in (
        (True, "normal", False), (True, "allocator", True),
        (False, "normal", False), (False, "allocator", True),
    ):
        for selector in selectors:
            job = next(job for job in jobs if job["selector"] == selector)
            samples, warmups = (
                (1, 0) if preflight else
                ((30, 3) if allocator else (job["samples"], job["warmups"]))
            )
            for leg, role in zip(ABBA_LEGS, ABBA_ROLES):
                prefix = "preflight" if preflight else phase
                ordered.append({
                    "key": f"{prefix}:{phase}:{leg}:{selector}",
                    "preflight": preflight, "phase": phase, "leg": leg,
                    "role": role, "selector": selector,
                    "samples": samples, "warmups": warmups,
                })
    return ordered


def expected_artifact_paths(item: dict[str, Any]) -> dict[str, str]:
    leg = item["leg"].lower()
    if item["preflight"]:
        stem, directory = f"{leg}-{item['phase']}-{item['selector']}", "preflight"
    else:
        stem, directory = f"{leg}-{item['selector']}", item["phase"]
    return {
        "report": f"{directory}/{stem}.json",
        "catalog": f"{directory}/{stem}.catalog.json",
        "time_v": f"{directory}/{stem}.time.txt",
        "stdout": f"{directory}/{stem}.stdout.txt",
        "stderr": f"{directory}/{stem}.stderr.txt",
    }


def expected_configuration(
    selector: str, samples: int, warmups: int
) -> dict[str, Any]:
    return {
        "samples_per_case": samples,
        "warmup_iterations_per_case": warmups,
        "filesystem_cache_states": ["warm"],
        "filesystem_fresh_child_per_sample": True,
        "filesystem_process_isolated": True,
        "filesystem_root_selected": False,
        "cases": [selector],
        "corpus_shapes": ["many-small"],
        "payload_kinds": ["compressible"],
        "writer_shapes": ["large"],
        "xlsx_shapes": ["medium"],
        "xlsb_shapes": ["tiny", "medium", "large", "sparse"],
        "xlsx_cell_crud_shapes": ["medium"],
        "xlsx_row_visibility_shapes": ["medium"],
        "semantic_shapes": ["medium"],
        "rtf_variants": ["plain"],
        "range_simulation": {
            "bandwidth_bytes_per_second": 52428800,
            "fixed_latency_us": 100,
            "max_physical_range_bytes": 4096,
            "request_overhead_us": 25,
        },
        "execution_workers": [1],
        "opc_cache_lock_diagnostics": False,
    }


def validate_elapsed(
    value: Any, path: str, expected_samples: int, *, strict_stats: bool
) -> dict[str, Any]:
    elapsed = mapping(value, path)
    if elapsed.get("unit") != "ns":
        fail(f"{path}.unit", "must be 'ns'")
    samples = sequence(elapsed.get("samples"), f"{path}.samples")
    if len(samples) != expected_samples:
        fail(f"{path}.samples", f"must contain {expected_samples} samples")
    normalized = [
        positive_integer(item, f"{path}.samples[{index}]")
        for index, item in enumerate(samples)
    ]
    if normalized != sorted(normalized):
        fail(f"{path}.samples", "must be sorted ascending")
    order = sequence(elapsed.get("sample_order"), f"{path}.sample_order")
    if len(order) != expected_samples or any(
        isinstance(item, bool) or not isinstance(item, int) or item < 0
        for item in order
    ) or set(order) != set(range(expected_samples)):
        fail(f"{path}.sample_order", "must be the complete original-index permutation")
    if strict_stats:
        try:
            from tools import perf_abba_summary
            return perf_abba_summary.recompute_statistics(elapsed, path)
        except Exception as error:
            fail(path, f"statistics rejected by perf_abba_summary: {error}")
    for name in ("min", "p50", "p95", "p99", "max"):
        if elapsed.get(name) != normalized[0]:
            fail(f"{path}.{name}", "does not match its one retained sample")
    for name in ("mean", "standard_deviation"):
        number = elapsed.get(name)
        if (
            isinstance(number, bool)
            or not isinstance(number, (int, float))
            or not math.isfinite(float(number))
        ):
            fail(f"{path}.{name}", "must be finite")
    if (
        float(elapsed["mean"]) != float(normalized[0])
        or float(elapsed["standard_deviation"]) != 0.0
    ):
        fail(path, "one-sample floating statistics are inconsistent")
    confidence = mapping(
        elapsed.get("confidence_interval_95"),
        f"{path}.confidence_interval_95",
    )
    if confidence.get("method") != "two-sided Student's t interval for the mean":
        fail(f"{path}.confidence_interval_95.method", "does not match the harness")
    if (
        confidence.get("lower") != elapsed["mean"]
        or confidence.get("upper") != elapsed["mean"]
    ):
        fail(f"{path}.confidence_interval_95", "does not match one-sample statistics")
    return {
        "sample_count": expected_samples,
        "min": normalized[0],
        "p50": normalized[0],
        "p95": normalized[0],
        "p99": normalized[0],
        "max": normalized[0],
        "mean": float(normalized[0]),
        "standard_deviation": 0.0,
        "confidence_interval_95": {
            "method": confidence["method"],
            "lower": float(normalized[0]),
            "upper": float(normalized[0]),
        },
    }


def validate_environment(
    value: Any, path: str, role: str, revision: str, phase: str
) -> dict[str, Any]:
    environment = mapping(value, path)
    required = (
        "rustc_version", "git_revision", "git_worktree_dirty",
        "logical_cpus_available", "allocator", "rustflags",
        "cargo_build_target", "perf_event_paranoid", "os", "kernel",
        "cpu_model", "total_memory_bytes", "page_size_bytes",
        "filesystem_type", "source_destination_same_device",
        "cpu_affinity", "storage_identifier",
    )
    if any(key not in environment for key in required):
        fail(path, "is missing a required harness environment field")
    rustc = string(environment["rustc_version"], f"{path}.rustc_version")
    if not rustc.startswith("rustc 1.98.1 "):
        fail(f"{path}.rustc_version", "must be from Rust 1.98.1")
    if environment["git_revision"] != revision:
        fail(f"{path}.git_revision", f"must match {role} source revision")
    if environment["git_worktree_dirty"] is not False:
        fail(f"{path}.git_worktree_dirty", "must be false")
    if environment["logical_cpus_available"] != 1 or environment["cpu_affinity"] != "2":
        fail(path, "must bind one logical CPU and affinity 2")
    if environment["rustflags"] != "-C force-frame-pointers=yes -C force-unwind-tables=yes":
        fail(f"{path}.rustflags", "does not match the build")
    allocator = (
        "CountingSystemAllocator(std::alloc::System)"
        if phase == "allocator" else "Rust system allocator"
    )
    if environment["allocator"] != allocator:
        fail(f"{path}.allocator", "does not match the binary lane")
    if environment["os"] != "linux":
        fail(f"{path}.os", "must be linux for the captured host")
    positive_integer(environment["total_memory_bytes"], f"{path}.total_memory_bytes")
    positive_integer(environment["page_size_bytes"], f"{path}.page_size_bytes")
    if environment["cargo_build_target"] is not None:
        string(environment["cargo_build_target"], f"{path}.cargo_build_target")
    if environment["filesystem_type"] is not None:
        string(environment["filesystem_type"], f"{path}.filesystem_type")
    if environment["storage_identifier"] is not None:
        string(environment["storage_identifier"], f"{path}.storage_identifier")
    if environment["source_destination_same_device"] is not None:
        boolean(
            environment["source_destination_same_device"],
            f"{path}.source_destination_same_device",
        )
    string(environment["kernel"], f"{path}.kernel")
    string(environment["cpu_model"], f"{path}.cpu_model")
    string(environment["perf_event_paranoid"], f"{path}.perf_event_paranoid")
    return environment


def validate_time(
    artifact: Artifact, path: str, *, argv: list[str] | None = None
) -> int:
    """Require complete GNU time -v output and return maximum RSS in KiB."""

    exit_lines = [
        line.strip()
        for line in artifact.text(path).splitlines()
        if line.strip().startswith("Exit status:")
    ]
    if exit_lines != ["Exit status: 0"]:
        fail(path, "must contain exactly one zero exit status")

    if artifact.path is not None and not artifact.compressed:
        try:
            from tools import perf_resource_profile
            parsed = perf_resource_profile.parse_time_report(artifact.path)
        except Exception as error:
            fail(path, f"existing time parser failed: {error}")
        if parsed.get("status") != "ok":
            fail(path, f"GNU time record is incomplete: {parsed!r}")
        text = artifact.text(path)
        command_lines = [
            line for line in text.splitlines()
            if line.lstrip().startswith("Command being timed:")
        ]
        if len(command_lines) != 1:
            fail(path, "must contain exactly one timed command")
        if argv is not None:
            command = command_lines[0]
            if len(argv) < 8 or argv[7] not in command:
                fail(path, "timed command does not bind the captured binary")
            for token in (
                "--case",
                argv[9] if len(argv) > 9 else "",
                argv[-3] if len(argv) >= 4 else "",
                argv[-1] if len(argv) >= 1 else "",
            ):
                if token and token not in command:
                    fail(path, "timed command does not bind capture arguments")
        rss = parsed.get("max_rss_kib")
        if isinstance(rss, bool) or not isinstance(rss, int) or rss <= 0:
            fail(path, "maximum RSS must be a positive integer")
        return rss
    text = artifact.text(path)
    lines = text.splitlines()
    for label in TIME_LABELS:
        matching = [
            line.strip()
            for line in lines
            if line.strip().startswith(label + ":")
        ]
        if len(matching) != 1:
            fail(path, f"must contain exactly one {label!r} field")
        raw_value = matching[0][len(label) + 1:].strip()
        pattern = TIME_VALUE_PATTERNS[label]
        if re.fullmatch(pattern, raw_value) is None:
            fail(path, f"{label!r} must contain a finite numeric value")
        if label in {
            "User time (seconds)",
            "System time (seconds)",
            "Elapsed (wall clock) time (h:mm:ss or m:ss)",
        }:
            try:
                numeric_values = (
                    [float(raw_value)]
                    if label != "Elapsed (wall clock) time (h:mm:ss or m:ss)"
                    else [float(part) for part in raw_value.split(":")]
                )
            except (OverflowError, ValueError):
                fail(path, f"{label!r} must contain a finite numeric value")
            if any(not math.isfinite(value) for value in numeric_values):
                fail(path, f"{label!r} must contain a finite numeric value")
    if argv is not None:
        command_lines = [
            line for line in lines
            if line.lstrip().startswith("Command being timed:")
        ]
        if len(command_lines) != 1 or len(argv) < 8 or argv[7] not in command_lines[0]:
            fail(path, "timed command does not bind the captured binary")
        command = command_lines[0]
        for token in (
            "--case",
            argv[9] if len(argv) > 9 else "",
            argv[-3] if len(argv) >= 4 else "",
            argv[-1] if len(argv) >= 1 else "",
        ):
            if token and token not in command:
                fail(path, "timed command does not bind capture arguments")
    match = re.search(
        r"(?m)^\s*Maximum resident set size \(kbytes\):\s*([0-9,]+)\s*$",
        text,
    )
    if match is None:
        fail(path, "must contain maximum resident set size")
    rss = int(match.group(1).replace(",", ""))
    if rss <= 0:
        fail(path, "maximum RSS must be positive")
    return rss


def validate_operation(
    result: dict[str, Any],
    path: str,
    *,
    lifecycle: bool,
    allocator_lane: bool,
) -> tuple[dict[str, Any], str]:
    operation = mapping(result.get("operation_metrics"), f"{path}.operation_metrics")
    elapsed = mapping(result.get("elapsed_ns"), f"{path}.elapsed_ns")
    try:
        from tools import perf_compare
        perf_compare._validate_operation_metrics(
            operation,
            f"{path}.operation_metrics",
            elapsed["samples"],
            1,
            elapsed_sample_order=elapsed["sample_order"],
        )
    except Exception as error:
        fail(path, f"operation metrics rejected by perf_compare: {error}")
    allocation = operation.get("allocation")
    if lifecycle:
        if allocation is None:
            fail(f"{path}.operation_metrics", "lifecycle rows must carry allocation status")
        allocation = mapping(allocation, f"{path}.operation_metrics.allocation")
        expected = "measured" if allocator_lane else "unavailable"
        if allocation.get("status") != expected:
            fail(f"{path}.operation_metrics.allocation.status", f"must be {expected!r}")
        if expected == "measured":
            try:
                from tools import perf_compare
                perf_compare._validate_allocator_operation_evidence(result, path, 1)
            except Exception as error:
                fail(path, f"allocator operation evidence rejected: {error}")
        return operation, expected
    if allocation is not None:
        fail(f"{path}.operation_metrics.allocation", "phase rows must omit allocation")
    return operation, "legacy_unavailable"


def validate_report(
    report: dict[str, Any],
    catalog: dict[str, Any],
    path: str,
    *,
    item: dict[str, Any],
    binary: dict[str, Any],
    revision: str,
) -> dict[str, Any]:
    expected_top = {
        "schema_version", "tool", "binary_identity", "environment",
        "configuration", "parallel_metrics", "results", "corpus_catalog",
    }
    if set(report) != expected_top:
        fail(path, "report has an unexpected top-level schema")
    if report["schema_version"] != 1:
        fail(f"{path}.schema_version", "must be 1")
    allocator_lane = item["phase"] == "allocator"
    expected_tool = {
        "name": "litchi-perf-baseline",
        "version": "0.1.0",
        "binary": (
            "litchi-perf-baseline-alloc"
            if allocator_lane else "litchi-perf-baseline"
        ),
        "profile": "release",
        "target_os": "linux",
        "target_arch": "x86_64",
        "instrumentation": (
            "system_allocator_operation_scoped" if allocator_lane else "none"
        ),
    }
    if report["tool"] != expected_tool:
        fail(f"{path}.tool", "does not match the bound binary lane")
    identity = mapping(report["binary_identity"], f"{path}.binary_identity")
    if set(identity) != {
        "path", "binary_sha256", "binary_bytes", "mode_bits",
        "executable", "profile",
    }:
        fail(f"{path}.binary_identity", "has an unexpected schema")
    if identity["path"] != binary["path"]:
        fail(f"{path}.binary_identity.path", "does not match build identity")
    if hash_value(
        identity["binary_sha256"], f"{path}.binary_identity.binary_sha256"
    ) != binary["sha256"]:
        fail(f"{path}.binary_identity.binary_sha256", "does not match build identity")
    if identity["binary_bytes"] != binary["bytes"] or identity["profile"] != "release":
        fail(f"{path}.binary_identity", "size/profile does not match binary")
    integer(identity["mode_bits"], f"{path}.binary_identity.mode_bits")
    if identity["mode_bits"] != binary["mode_bits"]:
        fail(f"{path}.binary_identity.mode_bits", "does not match build identity")
    if identity["executable"] != binary["executable"]:
        fail(
            f"{path}.binary_identity.executable",
            "does not match build identity",
        )
    if identity["executable"] is not True:
        fail(f"{path}.binary_identity.executable", "must be true")
    environment = validate_environment(
        report["environment"], f"{path}.environment",
        item["role"], revision, item["phase"],
    )
    try:
        from tools import perf_compare
        perf_compare.validate_parallel_metrics(report, path)
    except Exception as error:
        fail(path, f"parallel metrics rejected: {error}")
    configuration = expected_configuration(
        item["selector"], item["samples"], item["warmups"]
    )
    if report["configuration"] != configuration:
        fail(f"{path}.configuration", "does not match protocol flags/counts")
    catalog_reference = mapping(report["corpus_catalog"], f"{path}.corpus_catalog")
    expected_reference = {
        key: catalog[key]
        for key in (
            "manifest_version", "catalog_id",
            "catalog_sha256", "content_set_sha256",
        )
    }
    if catalog_reference != expected_reference:
        fail(f"{path}.corpus_catalog", "does not match catalog sidecar")
    try:
        from tools import validate_perf_corpus_binding
        validate_perf_corpus_binding.validate_binding(report, catalog)
    except Exception as error:
        fail(path, f"report/catalog binding rejected: {error}")
    results = sequence(report["results"], f"{path}.results")
    if len(results) != 1:
        fail(f"{path}.results", "must contain exactly one result")
    result = mapping(results[0], f"{path}.results[0]")
    expected_result_keys = {
        "case", "corpus", "elapsed_ns", "sink", "source",
        "output_sha256", "operation_metrics",
    }
    if set(result) != expected_result_keys:
        fail(f"{path}.results[0]", "has an unexpected schema")
    if result["case"] != item["selector"]:
        fail(f"{path}.results[0].case", "does not match capture selector")
    elapsed_stats = validate_elapsed(
        result["elapsed_ns"], f"{path}.results[0].elapsed_ns",
        item["samples"], strict_stats=not item["preflight"],
    )
    try:
        from tools import perf_abba_summary
        indexed = perf_abba_summary._index_results(report, path)
        perf_abba_summary._validate_pptx_cross_copy_result_rows(
            indexed, configuration, path
        )
        if not item["preflight"] and not allocator_lane:
            perf_abba_summary._validate_report(
                report, path, profile="current-v1",
                report_role=item["leg"].lower(),
            )
    except Exception as error:
        fail(path, f"strict PPTX/report validation rejected report: {error}")
    sink = mapping(result["sink"], f"{path}.results[0].sink")
    if set(sink) != {
        "accepted_bytes", "write_calls", "largest_write", "write_size_buckets",
    }:
        fail(f"{path}.results[0].sink", "has an unexpected schema")
    accepted = positive_integer(
        sink["accepted_bytes"], f"{path}.results[0].sink.accepted_bytes"
    )
    calls = positive_integer(
        sink["write_calls"], f"{path}.results[0].sink.write_calls"
    )
    largest = positive_integer(
        sink["largest_write"], f"{path}.results[0].sink.largest_write"
    )
    if largest > 65536:
        fail(f"{path}.results[0].sink.largest_write", "must not exceed 65536")
    buckets = mapping(
        sink["write_size_buckets"], f"{path}.results[0].sink.write_size_buckets"
    )
    bucket_names = {
        "bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
        "bytes_4097_to_16384", "bytes_16385_to_65536",
        "bytes_over_65536",
    }
    if set(buckets) != bucket_names:
        fail(f"{path}.results[0].sink.write_size_buckets", "has an unexpected schema")
    bucket_total = 0
    for name in bucket_names:
        bucket_total += integer(
            buckets[name],
            f"{path}.results[0].sink.write_size_buckets.{name}",
        )
    if bucket_total != calls or buckets["bytes_over_65536"] != 0:
        fail(f"{path}.results[0].sink.write_size_buckets", "does not reconcile")
    source = mapping(result["source"], f"{path}.results[0].source")
    for name in (
        "read_calls", "read_bytes", "ordinary_payload_read_calls",
        "ordinary_payload_read_bytes", "max_in_flight_reads",
    ):
        if name not in source:
            fail(f"{path}.results[0].source", f"missing {name!r}")
        values = sequence(
            source[name], f"{path}.results[0].source.{name}"
        )
        for index, value in enumerate(values):
            integer(
                value,
                f"{path}.results[0].source.{name}[{index}]",
            )
    operation, allocation_status = validate_operation(
        result, f"{path}.results[0]",
        lifecycle=item["selector"].endswith("_lifecycle"),
        allocator_lane=allocator_lane,
    )
    special = mapping(
        source.get("pptx_cross_copy"),
        f"{path}.results[0].source.pptx_cross_copy",
    )
    for name in (
        "source_archive_sha256", "destination_archive_sha256",
        "expected_output_sha256",
    ):
        hash_value(
            special.get(name),
            f"{path}.results[0].source.pptx_cross_copy.{name}",
        )
    output_oracle = hash_value(
        result["output_sha256"], f"{path}.results[0].output_sha256"
    )
    output_vector = sequence(
        special.get("output_sha256"),
        f"{path}.results[0].source.pptx_cross_copy.output_sha256",
    )
    if len(output_vector) != item["samples"]:
        fail(
            f"{path}.results[0].source.pptx_cross_copy.output_sha256",
            "does not match sample count",
        )
    output_vector = [
        hash_value(
            value,
            f"{path}.results[0].source.pptx_cross_copy.output_sha256[{index}]",
        )
        for index, value in enumerate(output_vector)
    ]
    if any(value != special["expected_output_sha256"] for value in output_vector):
        fail(
            f"{path}.results[0].source.pptx_cross_copy.output_sha256",
            "contains an unexpected digest",
        )
    if output_oracle != special["expected_output_sha256"]:
        fail(
            f"{path}.results[0].output_sha256",
            "does not match expected_output_sha256",
        )
    if special.get("output_sha256") != output_vector:
        fail(
            f"{path}.results[0].source.pptx_cross_copy.output_sha256",
            "does not match its validated output vector",
        )
    gates = mapping(
        special.get("gates"),
        f"{path}.results[0].source.pptx_cross_copy.gates",
    )
    if not gates or any(value is not True for value in gates.values()):
        fail(
            f"{path}.results[0].source.pptx_cross_copy.gates",
            "every correctness gate must be true",
        )
    vectors: dict[str, Any] = {}
    vector_names = (
        "plan_ns", "commit_ns", "publication_ns", "reopen_ns", "output_sha256",
    )
    if item["selector"].endswith("_lifecycle"):
        vector_names = (*vector_names, "lifecycle_ns")
    for name in vector_names:
        values = sequence(
            special.get(name),
            f"{path}.results[0].source.pptx_cross_copy.{name}",
        )
        if len(values) != item["samples"]:
            fail(
                f"{path}.results[0].source.pptx_cross_copy.{name}",
                "does not match sample count",
            )
        if name != "output_sha256":
            for index, value in enumerate(values):
                positive_integer(
                    value,
                    f"{path}.results[0].source.pptx_cross_copy.{name}[{index}]",
                )
        vectors[name] = {"length": len(values), "sha256": digest_json(values)}
    if item["selector"].endswith("_lifecycle"):
        if special["lifecycle_ns"] != result["elapsed_ns"]["samples"]:
            fail(
                f"{path}.results[0].source.pptx_cross_copy.lifecycle_ns",
                "must match elapsed samples",
            )
    else:
        if "lifecycle_ns" in special:
            fail(
                f"{path}.results[0].source.pptx_cross_copy.lifecycle_ns",
                "is forbidden for phase selectors",
            )
        phase_total = [
            special["plan_ns"][index]
            + special["commit_ns"][index]
            + special["publication_ns"][index]
            for index in range(item["samples"])
        ]
        sample_order = result["elapsed_ns"]["sample_order"]
        elapsed_samples = result["elapsed_ns"]["samples"]
        if any(
            phase_total[original_index] != elapsed_samples[sorted_index]
            for sorted_index, original_index in enumerate(sample_order)
        ):
            fail(
                f"{path}.results[0].source.pptx_cross_copy",
                "phase vectors do not align with elapsed sample_order",
            )
    return {
        "report": report,
        "catalog": catalog,
        "result": result,
        "environment": environment,
        "elapsed_stats": elapsed_stats,
        "operation": operation,
        "allocation_status": allocation_status,
        "vectors": vectors,
        "source_special": special,
        "output_oracle": output_oracle,
        "sink": {
            "accepted_bytes": accepted,
            "write_calls": calls,
            "largest_write": largest,
        },
    }


def validate_capture(
    root: Path,
    protocol_artifact: Artifact,
    protocol: dict[str, Any],
    jobs: list[dict[str, Any]],
    build_artifact: Artifact,
    build: dict[str, Any],
) -> tuple[
    Artifact,
    dict[tuple[bool, str, str, str], dict[str, Any]],
    list[dict[str, Any]],
]:
    capture_artifact, capture = document(root, "capture.json", "capture.json")
    if capture.get("schema_version") != 1 or capture.get("change") != EXPECTED_CHANGE:
        fail("capture", "must be schema 1 for change 0418")
    if capture.get("status") != "complete":
        fail("capture.status", "must be complete before projection")
    protocol_ref = mapping(capture.get("protocol"), "capture.protocol")
    if hash_value(
        protocol_ref.get("sha256"), "capture.protocol.sha256"
    ) != protocol_artifact.sha256:
        fail("capture.protocol.sha256", "does not match protocol.json")
    portable_path(protocol_ref.get("path"), "protocol.json", "capture.protocol.path")
    build_ref = mapping(capture.get("build_identity"), "capture.build_identity")
    if hash_value(
        build_ref.get("sha256"), "capture.build_identity.sha256"
    ) != build_artifact.sha256:
        fail("capture.build_identity.sha256", "does not match build-identity.json")
    portable_path(
        build_ref.get("path"), "build-identity.json",
        "capture.build_identity.path",
    )
    expected_sources = {
        role: build["roles"][role]["source"] for role in ("control", "candidate")
    }
    expected_binaries = {
        role: build["roles"][role]["binaries"]
        for role in ("control", "candidate")
    }
    if capture.get("source_roles") != expected_sources:
        fail("capture.source_roles", "does not match build source identities")
    if capture.get("binaries") != expected_binaries:
        fail("capture.binaries", "does not match build binary identities")
    execution = mapping(capture.get("execution"), "capture.execution")
    expected_execution = {
        "cpu": 2,
        "workers": 1,
        "abba_legs": list(ABBA_LEGS),
        "abba_roles": list(ABBA_ROLES),
        "common_flags": EXPECTED_FLAGS,
        "common_flags_source": "protocol.common_flags",
        "filesystem_cache": "warm",
        "fresh_process_per_selector_per_leg": True,
        "allocator_latency_comparison": "excluded by protocol",
    }
    if execution != expected_execution:
        fail("capture.execution", "does not match frozen execution contract")
    if capture.get("environment_overrides") != {"RUSTUP_TOOLCHAIN": "1.98.1"}:
        fail("capture.environment_overrides", "must bind Rust 1.98.1")
    plan = expected_plan(jobs, protocol)
    if capture.get("expected") != plan:
        fail("capture.expected", "does not match deterministic 64-run plan")
    runs = sequence(capture.get("runs"), "capture.runs")
    if len(runs) != len(plan):
        fail("capture.runs", f"must contain exactly {len(plan)} records")
    expected_completed = sorted(item["key"] for item in plan)
    if capture.get("completed_run_keys") != expected_completed:
        fail("capture.completed_run_keys", "must list all completed keys sorted")
    evidence: dict[tuple[bool, str, str, str], dict[str, Any]] = {}
    for index, (raw_run, expected) in enumerate(zip(runs, plan)):
        run = mapping(raw_run, f"capture.runs[{index}]")
        required_keys = {
            "key", "preflight", "phase", "leg", "role", "selector",
            "samples", "warmups", "argv", "cwd", "source_revision",
            "source_files", "binary_sha256", "binary_bytes",
            "environment_overrides", "started_utc", "finished_utc",
            "elapsed_seconds", "exit_code", "report", "catalog", "time_v",
            "stdout", "stderr", "artifact_sha256", "source_after",
        }
        if set(run) != required_keys:
            fail(f"capture.runs[{index}]", "has an unexpected schema")
        for key, expected_value in expected.items():
            if run.get(key) != expected_value:
                fail(
                    f"capture.runs[{index}].{key}",
                    "does not match deterministic run plan",
                )
        role = expected["role"]
        phase = expected["phase"]
        source = build["roles"][role]["source"]
        binary = build["roles"][role]["binaries"][phase]
        if (
            run["cwd"] != source["worktree"]
            or run["source_revision"] != source["revision"]
        ):
            fail(f"capture.runs[{index}]", "cwd/revision does not match role")
        if (
            run["source_files"] != source["source_files"]
            or run["source_after"] != source
        ):
            fail(f"capture.runs[{index}]", "source identity changed during capture")
        if (
            run["binary_sha256"] != binary["sha256"]
            or run["binary_bytes"] != binary["bytes"]
        ):
            fail(f"capture.runs[{index}]", "binary identity does not match build")
        if run["environment_overrides"] != {"RUSTUP_TOOLCHAIN": "1.98.1"}:
            fail(f"capture.runs[{index}].environment_overrides", "does not match capture")
        if run["exit_code"] != 0:
            fail(f"capture.runs[{index}].exit_code", "must be zero")
        elapsed_seconds = run["elapsed_seconds"]
        if (
            isinstance(elapsed_seconds, bool)
            or not isinstance(elapsed_seconds, (int, float))
            or not math.isfinite(float(elapsed_seconds))
            or elapsed_seconds <= 0
        ):
            fail(f"capture.runs[{index}].elapsed_seconds", "must be finite and positive")
        started = parse_timestamp(
            run["started_utc"], f"capture.runs[{index}].started_utc"
        )
        finished = parse_timestamp(
            run["finished_utc"], f"capture.runs[{index}].finished_utc"
        )
        if finished <= started:
            fail(f"capture.runs[{index}]", "finished_utc must be after started_utc")
        paths = expected_artifact_paths(expected)
        for key, relative in paths.items():
            if run[key] != relative:
                fail(f"capture.runs[{index}].{key}", f"must be {relative!r}")
        argv = sequence(run["argv"], f"capture.runs[{index}].argv")
        if (
            len(argv) < 8
            or argv[:6] != ["taskset", "-c", "2", "/usr/bin/time", "-v", "-o"]
        ):
            fail(f"capture.runs[{index}].argv", "does not start with taskset/time")
        portable_path(
            argv[6], paths["time_v"], f"capture.runs[{index}].argv[6]"
        )
        if argv[7] != binary["path"]:
            fail(f"capture.runs[{index}].argv[7]", "does not bind build binary")
        expected_tail = [
            "--case", expected["selector"], *EXPECTED_FLAGS,
            "--samples", str(expected["samples"]),
            "--warmup", str(expected["warmups"]),
        ]
        if argv[8:8 + len(expected_tail)] != expected_tail:
            fail(f"capture.runs[{index}].argv", "case/flag/sample contract differs")
        output_index = 8 + len(expected_tail)
        if (
            len(argv) != output_index + 4
            or argv[output_index] != "--json"
            or argv[output_index + 2] != "--corpus-manifest"
        ):
            fail(f"capture.runs[{index}].argv", "has unexpected output arguments")
        portable_path(
            argv[output_index + 1], paths["report"],
            f"capture.runs[{index}].argv.report",
        )
        portable_path(
            argv[output_index + 3], paths["catalog"],
            f"capture.runs[{index}].argv.catalog",
        )
        hashes = mapping(
            run["artifact_sha256"], f"capture.runs[{index}].artifact_sha256"
        )
        if set(hashes) != {"report", "catalog", "time_v", "stdout", "stderr"}:
            fail(f"capture.runs[{index}].artifact_sha256", "must hash five artifacts")
        artifacts: dict[str, Artifact] = {}
        for key, relative in paths.items():
            artifact = Artifact(
                root, relative, f"capture.runs[{index}].{key}",
                nonempty=key in {"report", "catalog", "time_v"},
            )
            if hash_value(
                hashes[key], f"capture.runs[{index}].artifact_sha256.{key}"
            ) != artifact.sha256:
                fail(
                    f"capture.runs[{index}].artifact_sha256.{key}",
                    "does not match artifact bytes",
                )
            artifacts[key] = artifact
        rss_kib = validate_time(
            artifacts["time_v"], f"capture.runs[{index}].time_v", argv=argv
        )
        report = mapping(
            artifacts["report"].json(f"capture.runs[{index}].report"),
            f"capture.runs[{index}].report",
        )
        catalog = mapping(
            artifacts["catalog"].json(f"capture.runs[{index}].catalog"),
            f"capture.runs[{index}].catalog",
        )
        parsed = validate_report(
            report, catalog, f"capture.runs[{index}].report",
            item=expected, binary=binary, revision=source["revision"],
        )
        key = (expected["preflight"], phase, expected["leg"], expected["selector"])
        evidence[key] = {
            **parsed,
            "run": run,
            "paths": paths,
            "report_sha256": artifacts["report"].sha256,
            "catalog_sha256": artifacts["catalog"].sha256,
            "time_sha256": artifacts["time_v"].sha256,
            "rss_kib": rss_kib,
        }
    return capture_artifact, evidence, plan


_DYNAMIC_SPECIAL_FIELDS = frozenset({
    "plan_ns", "commit_ns", "publication_ns", "reopen_ns",
    "lifecycle_ns", "output_sha256",
})
_PHASE_SPECIFIC_SPECIAL_FIELDS = frozenset({
    "timing_scope", "performance_claim", "implementation",
})


def stable_special(value: dict[str, Any]) -> dict[str, Any]:
    """Keep contract metadata while removing per-run timing/output vectors."""

    return {
        key: child
        for key, child in value.items()
        if key not in _DYNAMIC_SPECIAL_FIELDS
    }


def cross_phase_special(value: dict[str, Any]) -> dict[str, Any]:
    """Compare lifecycle and phase identity without phase-specific metadata."""

    ignored = _DYNAMIC_SPECIAL_FIELDS | _PHASE_SPECIFIC_SPECIAL_FIELDS
    return {key: child for key, child in value.items() if key not in ignored}


def stable_corpus(value: dict[str, Any], *, cross_selector: bool = False) -> str:
    """Canonicalize corpus identity, ignoring only selector naming when needed."""

    projected = dict(value)
    if cross_selector:
        projected.pop("name", None)
    return canonical(projected)


def validate_cross_identity(
    evidence: dict[tuple[bool, str, str, str], dict[str, Any]],
    selectors: list[str],
) -> None:
    for preflight in (True, False):
        for phase in ("normal", "allocator"):
            for selector in selectors:
                rows = [
                    evidence[(preflight, phase, leg, selector)]
                    for leg in ABBA_LEGS
                ]
                first = rows[0]
                stable = canonical(stable_special(first["source_special"]))
                corpus = stable_corpus(first["result"]["corpus"])
                output = first["output_oracle"]
                output_vector = first["vectors"]["output_sha256"]
                hashes = (
                    first["source_special"]["source_archive_sha256"],
                    first["source_special"]["destination_archive_sha256"],
                    first["source_special"]["expected_output_sha256"],
                )
                for row in rows[1:]:
                    if canonical(stable_special(row["source_special"])) != stable:
                        fail(
                            f"{phase}/{selector}",
                            "source identity differs across ABBA legs",
                        )
                    if stable_corpus(row["result"]["corpus"]) != corpus:
                        fail(
                            f"{phase}/{selector}",
                            "corpus identity differs across ABBA legs",
                        )
                    if row["output_oracle"] != output:
                        fail(
                            f"{phase}/{selector}",
                            "output oracle differs across ABBA legs",
                        )
                    if row["vectors"]["output_sha256"] != output_vector:
                        fail(
                            f"{phase}/{selector}",
                            "output digest vector differs across ABBA legs",
                        )
                    current = row["source_special"]
                    if (
                        current["source_archive_sha256"],
                        current["destination_archive_sha256"],
                        current["expected_output_sha256"],
                    ) != hashes:
                        fail(
                            f"{phase}/{selector}",
                            "source/output archive hashes differ across ABBA legs",
                        )
            for leg in ABBA_LEGS:
                for lifecycle, phase_selector in (
                    (
                        "pptx_cross_copy_media_rich_lifecycle",
                        "pptx_cross_copy_media_rich",
                    ),
                    ("pptx_cross_copy_plain_lifecycle", "pptx_cross_copy_plain"),
                ):
                    lifecycle_row = evidence[(preflight, phase, leg, lifecycle)]
                    phase_row = evidence[(preflight, phase, leg, phase_selector)]
                    if cross_phase_special(
                        lifecycle_row["source_special"]
                    ) != cross_phase_special(phase_row["source_special"]):
                        fail(f"{phase}/{leg}", "lifecycle and phase identity differs")
                    if stable_corpus(
                        lifecycle_row["result"]["corpus"], cross_selector=True
                    ) != stable_corpus(
                        phase_row["result"]["corpus"], cross_selector=True
                    ):
                        fail(f"{phase}/{leg}", "lifecycle and phase corpus differs")
                    if (
                        lifecycle_row["source_special"]["expected_output_sha256"]
                        != phase_row["source_special"]["expected_output_sha256"]
                    ):
                        fail(f"{phase}/{leg}", "lifecycle/phase output oracle differs")
                    if lifecycle_row["output_oracle"] != phase_row["output_oracle"]:
                        fail(f"{phase}/{leg}", "lifecycle/phase output oracle differs")
        for leg in ABBA_LEGS:
            for selector in selectors:
                normal_row = evidence[(preflight, "normal", leg, selector)]
                allocator_row = evidence[(preflight, "allocator", leg, selector)]
                if canonical(stable_special(normal_row["source_special"])) != canonical(
                    stable_special(allocator_row["source_special"])
                ):
                    fail(
                        f"{leg}/{selector}",
                        "normal and allocator source identity differs",
                    )
                if stable_corpus(normal_row["result"]["corpus"]) != stable_corpus(
                    allocator_row["result"]["corpus"]
                ):
                    fail(f"{leg}/{selector}", "normal and allocator corpus differs")
                if normal_row["output_oracle"] != allocator_row["output_oracle"]:
                    fail(f"{leg}/{selector}", "normal and allocator output oracle differs")
        if preflight:
            for phase in ("normal", "allocator"):
                for leg in ABBA_LEGS:
                    for selector in selectors:
                        preflight_row = evidence[(True, phase, leg, selector)]
                        formal_row = evidence[(False, phase, leg, selector)]
                        if canonical(
                            stable_special(preflight_row["source_special"])
                        ) != canonical(stable_special(formal_row["source_special"])):
                            fail(
                                f"{phase}/{leg}/{selector}",
                                "preflight and formal source identity differs",
                            )
                        if canonical(preflight_row["catalog"]) != canonical(
                            formal_row["catalog"]
                        ):
                            fail(
                                f"{phase}/{leg}/{selector}",
                                "preflight and formal catalog identity differs",
                            )
                        if stable_corpus(
                            preflight_row["result"]["corpus"]
                        ) != stable_corpus(formal_row["result"]["corpus"]):
                            fail(
                                f"{phase}/{leg}/{selector}",
                                "preflight and formal corpus identity differs",
                            )
                        if preflight_row["output_oracle"] != formal_row["output_oracle"]:
                            fail(
                                f"{phase}/{leg}/{selector}",
                                "preflight and formal output oracle differs",
                            )


def vector_projection(values: list[int], scope: str) -> dict[str, Any]:
    total = sum(values)
    return {
        "status": "measured",
        "scope": scope,
        "sample_count": len(values),
        "values_sha256": digest_json(values),
        "minimum": min(values),
        "maximum": max(values),
        "sum": total,
        "mean": total / len(values),
    }


def percent(delta: int, denominator: int) -> float | None:
    if denominator == 0:
        return None
    value = delta / denominator * 100.0
    if not math.isfinite(value):
        raise VerificationError("computed percentage is non-finite")
    return value


def matched_metric(
    control: list[int] | None,
    candidate: list[int] | None,
    scope: str,
    *,
    reviewable: bool,
) -> dict[str, Any]:
    if control is None or candidate is None:
        return {
            "status": "legacy_unavailable",
            "scope": "legacy_phase_selector_no_operation_allocation",
            "reviewable": reviewable,
        }
    if len(control) != len(candidate):
        raise VerificationError("matched allocator vectors have different cardinality")
    control_sum = sum(control)
    candidate_sum = sum(candidate)
    control_mean = control_sum / len(control)
    candidate_mean = candidate_sum / len(candidate)
    delta_sum = candidate_sum - control_sum
    delta_mean = candidate_mean - control_mean
    review = reviewable and (
        control_sum > 0
        and candidate_sum > control_sum * (1 + REVIEW_THRESHOLD_PERCENT / 100)
    )
    return {
        "status": "measured",
        "scope": scope,
        "sample_count": len(control),
        "reviewable": reviewable,
        "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
        "alignment": (
            "process-leg aggregate comparison over elapsed-sorted vectors; "
            "no per-rank pairing"
        ),
        "control_sum": control_sum,
        "control_mean": control_mean,
        "candidate_sum": candidate_sum,
        "candidate_mean": candidate_mean,
        "candidate_minus_control_sum": delta_sum,
        "candidate_minus_control_mean": delta_mean,
        "candidate_minus_control_percent": percent(delta_sum, control_sum),
        "review_required": review,
        "review_reason": (
            "candidate aggregate exceeds control by more than five percent"
            if review else None
        ),
    }


def rss_pair(control: int, candidate: int) -> dict[str, Any]:
    delta = candidate - control
    review = candidate > control * (1 + REVIEW_THRESHOLD_PERCENT / 100)
    return {
        "control_kib": control,
        "candidate_kib": candidate,
        "candidate_minus_control_kib": delta,
        "candidate_minus_control_percent": percent(delta, control),
        "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
        "review_required": review,
        "review_reason": (
            "candidate RSS exceeds control by more than five percent"
            if review else None
        ),
    }


def build_rss_projection(
    evidence: dict[tuple[bool, str, str, str], dict[str, Any]],
    selectors: list[str],
) -> dict[str, Any]:
    lanes: dict[str, Any] = {}
    for phase in ("normal", "allocator"):
        lane: dict[str, Any] = {}
        for selector in selectors:
            legs: dict[str, Any] = {}
            for leg in ABBA_LEGS:
                row = evidence[(False, phase, leg, selector)]
                legs[leg] = {
                    "role": LEG_TO_ROLE[leg],
                    "rss_kib": row["rss_kib"],
                    "time_v": row["paths"]["time_v"],
                    "time_v_sha256": row["time_sha256"],
                }
            lane[selector] = {
                "legs": legs,
                "a1_to_b1": rss_pair(
                    legs["A1"]["rss_kib"], legs["B1"]["rss_kib"]
                ),
                "a2_to_b2": rss_pair(
                    legs["A2"]["rss_kib"], legs["B2"]["rss_kib"]
                ),
            }
        lanes[phase] = lane
    return {
        "source": "GNU /usr/bin/time -v Maximum resident set size (kbytes)",
        "scope": (
            "whole-process maximum resident set size, including warmups "
            "and harness setup"
        ),
        "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
        "lanes": lanes,
    }


def build_allocation_projection(
    evidence: dict[tuple[bool, str, str, str], dict[str, Any]],
    selectors: list[str],
    roles: dict[str, Any],
    build: dict[str, Any],
    protocol_artifact: Artifact,
    capture_artifact: Artifact,
) -> dict[str, Any]:
    selector_projection: dict[str, Any] = {}
    measured_count = 0
    unavailable_count = 0
    for selector in selectors:
        legs: dict[str, Any] = {}
        for leg in ABBA_LEGS:
            row = evidence[(False, "allocator", leg, selector)]
            operation = row["operation"]
            allocation = operation.get("allocation")
            record: dict[str, Any] = {
                "role": LEG_TO_ROLE[leg],
                "report": row["paths"]["report"],
                "report_sha256": row["report_sha256"],
                "sample_count": operation["sample_count"],
                "sample_indices": operation["sample_indices"],
                "alignment": operation["alignment"],
                "status": row["allocation_status"],
                "metrics": {},
                "rss_kib": row["rss_kib"],
            }
            if allocation is None:
                for field in ALLOCATION_FIELDS:
                    record["metrics"][field] = {
                        "status": "legacy_unavailable",
                        "scope": "legacy_phase_selector_no_operation_allocation",
                    }
                    unavailable_count += 1
            else:
                scope = allocation["scope"]
                for field in ALLOCATION_FIELDS:
                    values = allocation[field]["values"]
                    record["metrics"][field] = vector_projection(values, scope)
                    measured_count += 1
            legs[leg] = record
        selector_projection[selector] = {"legs": legs}
    matched: dict[str, Any] = {}
    for selector in selectors:
        selector_pairs: dict[str, Any] = {}
        for pair, control_leg, candidate_leg in (
            ("a1_to_b1", "A1", "B1"),
            ("a2_to_b2", "A2", "B2"),
        ):
            left = selector_projection[selector]["legs"][control_leg]
            right = selector_projection[selector]["legs"][candidate_leg]
            metrics: dict[str, Any] = {}
            for field in ALLOCATION_FIELDS:
                left_metric = left["metrics"][field]
                right_metric = right["metrics"][field]
                if (
                    left_metric["status"] == "measured"
                    and right_metric["status"] == "measured"
                ):
                    left_values = evidence[
                        (False, "allocator", control_leg, selector)
                    ]["operation"]["allocation"][field]["values"]
                    right_values = evidence[
                        (False, "allocator", candidate_leg, selector)
                    ]["operation"]["allocation"][field]["values"]
                    metrics[field] = matched_metric(
                        left_values,
                        right_values,
                        left_metric["scope"],
                        reviewable=field == "allocated_bytes",
                    )
                else:
                    metrics[field] = matched_metric(
                        None,
                        None,
                        "legacy",
                        reviewable=field == "allocated_bytes",
                    )
            selector_pairs[pair] = {
                "control_leg": control_leg,
                "candidate_leg": candidate_leg,
                "metrics": metrics,
                "rss": rss_pair(
                    left["rss_kib"], right["rss_kib"]
                ),
            }
        matched[selector] = selector_pairs
    return {
        "schema_version": 1,
        "change": "0418",
        "timing_status": "observational_only",
        "latency_comparison": "excluded",
        "review_threshold_percent": REVIEW_THRESHOLD_PERCENT,
        "review_scope": (
            "aggregate allocated_bytes and whole-process RSS growth over five "
            "percent requires explicit tradeoff review"
        ),
        "nonclaimable_metrics": [
            "live_bytes_before", "live_bytes_after",
            "peak_live_bytes_before", "peak_live_bytes_after",
        ],
        "scope": (
            "operation_global_system_allocator for lifecycle selectors; "
            "legacy phase selectors retain unavailable status"
        ),
        "protocol": {
            "path": "protocol.json",
            "sha256": protocol_artifact.sha256,
            "order": list(ABBA_LEGS),
            "roles": list(ABBA_ROLES),
            "allocator_samples": 30,
            "allocator_warmups": 3,
            "metric_fields": list(ALLOCATION_FIELDS),
        },
        "provenance": {
            "capture": {"path": "capture.json", "sha256": capture_artifact.sha256},
            "roles": {
                role: {
                    "revision": roles[role]["revision"],
                    "worktree_basename": Path(roles[role]["worktree"]).name,
                }
                for role in ("control", "candidate")
            },
            "binaries": {
                role: {
                    phase: {
                        "sha256": build["roles"][role]["binaries"][phase]["sha256"],
                        "bytes": build["roles"][role]["binaries"][phase]["bytes"],
                    }
                    for phase in ("normal", "allocator")
                }
                for role in ("control", "candidate")
            },
        },
        "selectors": selector_projection,
        "matched_pairs": matched,
        "rss": build_rss_projection(evidence, selectors),
        "validation": {
            "measured_vector_count": measured_count,
            "legacy_unavailable_vector_count": unavailable_count,
            "expected_measured_vector_count": 80,
            "expected_legacy_unavailable_vector_count": 80,
            "all_vectors_aligned": True,
            "all_allocator_statuses_bound": True,
            "no_allocator_latency_claim": True,
        },
    }


def source_projection(source: dict[str, Any]) -> dict[str, Any]:
    return {
        "worktree_basename": Path(source["worktree"]).name,
        "revision": source["revision"],
        "clean": source["clean"],
        "source_files": source["source_files"],
    }


def binary_projection(binary: dict[str, Any]) -> dict[str, Any]:
    return {
        "sha256": binary["sha256"],
        "bytes": binary["bytes"],
        "mode_bits": binary["mode_bits"],
        "executable": binary["executable"],
        "profile": binary["profile"],
    }


def evidence_projection(
    evidence: dict[tuple[bool, str, str, str], dict[str, Any]],
    selectors: list[str],
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for phase in ("normal", "allocator"):
        for selector in selectors:
            for leg in ABBA_LEGS:
                row = evidence[(False, phase, leg, selector)]
                special = row["source_special"]
                rows.append({
                    "phase": phase,
                    "leg": leg,
                    "role": LEG_TO_ROLE[leg],
                    "selector": selector,
                    "samples": row["run"]["samples"],
                    "warmups": row["run"]["warmups"],
                    "report": row["paths"]["report"],
                    "catalog": row["paths"]["catalog"],
                    "report_sha256": row["report_sha256"],
                    "catalog_sha256": row["catalog_sha256"],
                    "source_archive_sha256": special["source_archive_sha256"],
                    "destination_archive_sha256": special["destination_archive_sha256"],
                    "expected_output_sha256": special["expected_output_sha256"],
                    "output_sha256": row["result"]["output_sha256"],
                    "gates": special["gates"],
                    "vectors": row["vectors"],
                    "allocation_status": row["allocation_status"],
                    "rss_kib": row["rss_kib"],
                })
    return rows


def tool_provenance() -> dict[str, Any]:
    modules = (
        ("tools/perf_abba_summary.py", "tools.perf_abba_summary"),
        ("tools/perf_compare.py", "tools.perf_compare"),
        ("tools/perf_resource_profile.py", "tools.perf_resource_profile"),
        (
            "tools/validate_perf_corpus_binding.py",
            "tools.validate_perf_corpus_binding",
        ),
    )
    result: dict[str, Any] = {}
    for name, module_name in modules:
        try:
            module = importlib.import_module(module_name)
            source_name = module.__file__
        except (ImportError, AttributeError) as error:
            fail(name, f"verifier source/tool cannot be imported: {error}")
        if source_name is None:
            fail(name, "imported tool has no source file")
        path = Path(source_name).resolve()
        if not path.is_file():
            fail(name, "imported verifier source/tool is missing")
        result[name] = {
            "bytes": path.stat().st_size,
            "sha256": digest_file(path),
        }
    verifier_path = Path(__file__).resolve()
    if not verifier_path.is_file():
        fail(
            "docs/performance/results/change-0418/verify.py",
            "verifier source is missing",
        )
    result["docs/performance/results/change-0418/verify.py"] = {
        "bytes": verifier_path.stat().st_size,
        "sha256": digest_file(verifier_path),
    }
    return result


def build_summary(
    protocol_artifact: Artifact,
    protocol: dict[str, Any],
    roles_artifact: Artifact,
    roles: dict[str, Any],
    build_artifact: Artifact,
    build: dict[str, Any],
    capture_artifact: Artifact,
    evidence: dict[tuple[bool, str, str, str], dict[str, Any]],
    plan: list[dict[str, Any]],
    selectors: list[str],
    allocation: dict[str, Any],
) -> dict[str, Any]:
    normal_summaries: dict[str, Any] = {}
    try:
        from tools import perf_abba_summary
        ceilings = {name: 5 for name in STATISTICS}
        for selector in selectors:
            reports = [
                evidence[(False, "normal", leg, selector)]["report"]
                for leg in ABBA_LEGS
            ]
            normal_summaries[selector] = perf_abba_summary.summarize_reports(
                reports=reports,
                cases=[selector],
                drift_ceilings=ceilings,
            )
    except Exception as error:
        fail("normal_abba", f"perf_abba_summary rejected formal reports: {error}")
    preflight_count = sum(1 for item in plan if item["preflight"])
    formal_count = len(plan) - preflight_count
    return {
        "schema_version": 1,
        "change": "0418",
        "timing_status": (
            "normal ABBA elapsed statistics validated from retained samples"
        ),
        "allocator_timing_status": "excluded",
        "protocol": {
            "path": "protocol.json",
            "sha256": protocol_artifact.sha256,
            "order": list(ABBA_LEGS),
            "roles": list(ABBA_ROLES),
            "cpu": protocol["cpu"],
            "workers": protocol["workers"],
            "common_flags": EXPECTED_FLAGS,
            "normal_counts": {
                job["selector"]: {
                    "samples": job["samples"], "warmups": job["warmups"],
                }
                for job in protocol["jobs"]
            },
            "allocator_counts": {"samples": 30, "warmups": 3},
            "latency_drift_ceilings_percent": {
                name: 5 for name in STATISTICS
            },
        },
        "provenance": {
            "roles": {
                "path": "roles.json",
                "sha256": roles_artifact.sha256,
                "control": source_projection(build["roles"]["control"]["source"]),
                "candidate": source_projection(build["roles"]["candidate"]["source"]),
            },
            "build_identity": {
                "path": "build-identity.json",
                "sha256": build_artifact.sha256,
            },
            "capture": {
                "path": "capture.json",
                "sha256": capture_artifact.sha256,
            },
            "binaries": {
                role: {
                    phase: binary_projection(
                        build["roles"][role]["binaries"][phase]
                    )
                    for phase in ("normal", "allocator")
                }
                for role in ("control", "candidate")
            },
            "validator_sources": tool_provenance(),
        },
        "capture": {
            "expected_run_count": len(plan),
            "preflight_run_count": preflight_count,
            "formal_run_count": formal_count,
            "validated_run_count": len(evidence),
            "selectors": selectors,
        },
        "normal_abba": normal_summaries,
        "correctness_evidence": {
            "formal_runs": evidence_projection(evidence, selectors),
            "cross_leg_source_identity_verified": True,
            "cross_phase_lifecycle_identity_verified": True,
            "source_output_hashes_verified": True,
            "all_correctness_gates_true": True,
            "lifecycle_raw_vectors_retained_in_reports": True,
        },
        "rss": allocation["rss"],
        "allocation": {
            "projection": "allocation-metrics.json",
            "measured_lifecycle_selector_count": 2,
            "legacy_phase_selector_count": 2,
            "latency_comparison": "excluded",
        },
        "verification": {
            "protocol_identity_verified": True,
            "role_source_identities_verified": True,
            "binary_identities_verified": True,
            "catalog_bindings_verified": True,
            "command_arguments_and_order_verified": True,
            "preflight_verified": True,
            "formal_sample_contract_verified": True,
            "normal_statistics_recomputed": True,
            "allocator_vectors_and_statuses_verified": True,
            "rss_from_time_v_verified": True,
            "deterministic_projection": True,
        },
    }


def serialized(value: Any) -> bytes:
    try:
        return (
            json.dumps(
                value, indent=2, sort_keys=True,
                ensure_ascii=False, allow_nan=False,
            ) + "\n"
        ).encode("utf-8")
    except (TypeError, ValueError, OverflowError) as error:
        raise VerificationError(
            f"projection is not deterministically serializable: {error}"
        ) from error


def publish_or_replay(
    root: Path, name: str, value: dict[str, Any], *, write: bool
) -> None:
    expected = serialized(value)
    path = root / name
    if write:
        temporary = root / f".{name}.tmp"
        temporary.write_bytes(expected)
        temporary.replace(path)
    else:
        if not path.is_file():
            fail(name, "published projection is missing; run with --write first")
        if path.read_bytes() != expected:
            fail(name, "published projection differs from deterministic replay")


def _verify(root: Path, *, write: bool) -> dict[str, Any]:
    root = root.resolve()
    protocol_artifact, protocol, jobs = validate_protocol(root)
    roles_artifact, roles = validate_roles(root)
    build_artifact, build = validate_build(root, protocol_artifact, roles)
    capture_artifact, evidence, plan = validate_capture(
        root, protocol_artifact, protocol, jobs, build_artifact, build,
    )
    selectors = [job["selector"] for job in jobs]
    validate_cross_identity(evidence, selectors)
    allocation = build_allocation_projection(
        evidence, selectors, roles, build,
        protocol_artifact, capture_artifact,
    )
    summary = build_summary(
        protocol_artifact, protocol, roles_artifact, roles,
        build_artifact, build, capture_artifact, evidence,
        plan, selectors, allocation,
    )
    publish_or_replay(root, "summary.json", summary, write=write)
    publish_or_replay(
        root, "allocation-metrics.json", allocation, write=write
    )
    return {
        "status": "pass",
        "change": EXPECTED_CHANGE,
        "write": write,
        "reports": len(evidence),
        "preflight_reports": sum(key[0] for key in evidence),
        "formal_reports": sum(not key[0] for key in evidence),
        "selectors": selectors,
        "outputs": ["summary.json", "allocation-metrics.json"],
    }


def verify(root: Path, *, write: bool) -> dict[str, Any]:
    token = _ACTIVE_ARTIFACT_BUDGET.set(ArtifactBudget())
    try:
        return _verify(root, write=write)
    finally:
        _ACTIVE_ARTIFACT_BUDGET.reset(token)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=CHANGE_ROOT)
    parser.add_argument(
        "--write", "--write-summary", action="store_true", dest="write",
        help="publish deterministic summary.json and allocation-metrics.json",
    )
    args = parser.parse_args()
    try:
        result = verify(args.root, write=args.write)
    except (VerificationError, OSError, subprocess.SubprocessError) as error:
        print(f"0418 verification failed: {error}", file=sys.stderr)
        return 1
    print(json.dumps(result, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
