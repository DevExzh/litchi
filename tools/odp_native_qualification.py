#!/usr/bin/env python3
"""Run a matched native ODP correctness preflight and calibrated ABBA.

The runner never builds a source revision.  An execution manifest supplies two
already-built probe binaries, their source worktrees and requested frozen Git
refs, plus the fixture digest and external target directories.  The runner
resolves and records those identities at execution time, runs the probe's
untimed correctness mode first, then calibrates each route independently and
collects one-process-per-leg ABBA samples.  Any source, fixture, binary,
process, load, output, or timing-contract mismatch fails closed.

All artifacts are written to the caller-provided directory.  The runner
rejects an artifact or target directory below either source worktree so a
qualification cannot accidentally publish binaries or raw evidence into the
checkout.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
import math
import os
from pathlib import Path
import platform
import re
import shutil
import subprocess
import sys
import time
from typing import Any, Iterable, Mapping
import zipfile


INPUT_SCHEMA = "litchi.odp.native-qualification.input.v1"
REPORT_SCHEMA = "litchi.odp.native-qualification.v1"
PROBE_SCHEMA = "litchi.odp.native-qualification.probe.v1"
LEGACY_PROBE_SCHEMA = "litchi.odp.native-qualification.legacy-probe.v1"
LEGACY_PROFILE = "legacy-e4a-plan"
CURRENT_PROFILE = "current-commit-patch"
PROBE_PROFILES = {
    LEGACY_PROFILE: LEGACY_PROBE_SCHEMA,
    CURRENT_PROFILE: PROBE_SCHEMA,
}
# These are the only routes whose owner lifecycle and publication result are
# shared by the pinned pre-Patch baseline and the current candidate.  Replay
# remains a candidate-only correctness proof and is never a timed route.
ROUTES = ("open", "read", "write", "append")
ABBA_ORDER = ("a1", "b1", "b2", "a2")
ROLE_TO_ENDPOINT = {"a1": "baseline", "a2": "baseline", "b1": "candidate", "b2": "candidate"}
HEX64 = re.compile(r"[0-9a-f]{64}\Z")
REVISION = re.compile(r"[0-9a-f]{40}\Z")
MAX_MANIFEST_BYTES = 256 * 1024
MAX_CHILD_OUTPUT = 4 * 1024 * 1024

HEAVY_EXECUTABLES = {
    "cargo",
    "rustc",
    "rustdoc",
    "rustfmt",
    "clippy-driver",
    "perf",
    "heaptrack",
    "valgrind",
    "callgrind",
    "clang",
    "clang++",
    "gcc",
    "cc",
    "ld.lld",
}
HEAVY_WORD = re.compile(
    r"(?:^|[\s/'\"])(?:cargo|rustc|rustdoc|rustfmt|clippy-driver|perf|heaptrack|"
    r"valgrind|callgrind|clang(?:\+\+)?|gcc|cc|ld\.lld)(?:$|[\s/'\"])",
)
PROBE_PREFIXES = ("litchi-perf-", "odp_native_qualification_probe", "odp-native-")


class RunFailure(RuntimeError):
    """A fail-closed qualification input or execution failure."""


@dataclass(frozen=True)
class Endpoint:
    name: str
    root: Path
    requested_ref: str
    binary: Path
    target_dir: Path
    profile: str


@dataclass(frozen=True)
class Fixture:
    path: Path
    bytes: int
    sha256: str
    provenance: str


@dataclass(frozen=True)
class Protocol:
    cpu: int
    blocks: int
    warmups: int
    target_ns: int
    minimum_ns: int
    max_load: float
    max_iterations: int
    timeout_s: float


@dataclass(frozen=True)
class InputManifest:
    path: Path
    fixture: Fixture
    baseline: Endpoint
    candidate: Endpoint
    protocol: Protocol
    title: str
    body: str


@dataclass(frozen=True)
class ResolvedEndpoint:
    endpoint: Endpoint
    resolved_ref: str
    binary_bytes: int
    binary_sha256: str


def _object(value: Any, location: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise RunFailure(f"{location} must be an object")
    return value


def _string(value: Any, location: str, *, nonempty: bool = True) -> str:
    if not isinstance(value, str) or (nonempty and not value):
        raise RunFailure(f"{location} must be a non-empty string")
    return value


def _positive_int(value: Any, location: str) -> int:
    if type(value) is not int or value <= 0:
        raise RunFailure(f"{location} must be a positive integer")
    return value


def _nonnegative_int(value: Any, location: str) -> int:
    if type(value) is not int or value < 0:
        raise RunFailure(f"{location} must be a nonnegative integer")
    return value


def _positive_float(value: Any, location: str) -> float:
    if type(value) not in (int, float) or not math.isfinite(float(value)) or float(value) <= 0:
        raise RunFailure(f"{location} must be a positive finite number")
    return float(value)


def _path(value: Any, location: str, base: Path) -> Path:
    raw = _string(value, location)
    path = Path(raw)
    return (base / path).resolve() if not path.is_absolute() else path.resolve()


def _digest(value: Any, location: str) -> str:
    result = _string(value, location).lower()
    if HEX64.fullmatch(result) is None:
        raise RunFailure(f"{location} must be 64 lowercase hexadecimal characters")
    return result


def _load_json(path: Path) -> dict[str, Any]:
    try:
        raw = path.read_bytes()
    except OSError as error:
        raise RunFailure(f"cannot read manifest {path}: {error}") from error
    if len(raw) > MAX_MANIFEST_BYTES:
        raise RunFailure(f"manifest exceeds {MAX_MANIFEST_BYTES} bytes")

    def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise RunFailure(f"manifest contains duplicate key {key!r}")
            result[key] = value
        return result

    try:
        value = json.loads(raw.decode("utf-8"), object_pairs_hook=reject_duplicate_keys)
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunFailure(f"manifest is not UTF-8 JSON: {error}") from error
    return _object(value, "manifest")


def load_manifest(path: Path) -> InputManifest:
    """Parse the execution manifest without resolving stale source hashes."""

    path = path.resolve()
    value = _load_json(path)
    if value.get("schema") != INPUT_SCHEMA:
        raise RunFailure(f"manifest.schema must be {INPUT_SCHEMA!r}")
    base = path.parent
    fixture_value = _object(value.get("fixture"), "manifest.fixture")
    fixture = Fixture(
        path=_path(fixture_value.get("path"), "manifest.fixture.path", base),
        bytes=_positive_int(fixture_value.get("bytes"), "manifest.fixture.bytes"),
        sha256=_digest(fixture_value.get("sha256"), "manifest.fixture.sha256"),
        provenance=_string(fixture_value.get("provenance"), "manifest.fixture.provenance"),
    )

    def endpoint(name: str) -> Endpoint:
        endpoint_value = _object(value.get(name), f"manifest.{name}")
        profile = _string(endpoint_value.get("profile"), f"manifest.{name}.profile")
        if profile not in PROBE_PROFILES:
            raise RunFailure(
                f"manifest.{name}.profile must be one of {sorted(PROBE_PROFILES)}"
            )
        return Endpoint(
            name=name,
            root=_path(endpoint_value.get("root"), f"manifest.{name}.root", base),
            requested_ref=_string(endpoint_value.get("ref"), f"manifest.{name}.ref"),
            binary=_path(endpoint_value.get("binary"), f"manifest.{name}.binary", base),
            target_dir=_path(endpoint_value.get("target_dir"), f"manifest.{name}.target_dir", base),
            profile=profile,
        )

    protocol_value = _object(value.get("protocol"), "manifest.protocol")
    cpu = _nonnegative_int(protocol_value.get("cpu"), "manifest.protocol.cpu")
    blocks = _positive_int(protocol_value.get("blocks"), "manifest.protocol.blocks")
    warmups = _nonnegative_int(protocol_value.get("warmups"), "manifest.protocol.warmups")
    target_ns = round(_positive_float(protocol_value.get("target_ms"), "manifest.protocol.target_ms") * 1_000_000)
    minimum_ns = round(
        _positive_float(protocol_value.get("minimum_ms"), "manifest.protocol.minimum_ms") * 1_000_000
    )
    if target_ns <= 0 or minimum_ns <= 0:
        raise RunFailure("manifest protocol durations must be at least one nanosecond")
    if target_ns < minimum_ns:
        raise RunFailure("manifest.protocol.target_ms must be at least minimum_ms")
    protocol = Protocol(
        cpu=cpu,
        blocks=blocks,
        warmups=warmups,
        target_ns=target_ns,
        minimum_ns=minimum_ns,
        max_load=_positive_float(protocol_value.get("max_load"), "manifest.protocol.max_load"),
        max_iterations=_positive_int(
            protocol_value.get("max_iterations"), "manifest.protocol.max_iterations"
        ),
        timeout_s=_positive_float(protocol_value.get("timeout_s"), "manifest.protocol.timeout_s"),
    )
    title = _string(value.get("title", "Qualification appended title"), "manifest.title")
    body = _string(value.get("body", "Qualification appended body"), "manifest.body")
    return InputManifest(
        path=path,
        fixture=fixture,
        baseline=endpoint("baseline"),
        candidate=endpoint("candidate"),
        protocol=protocol,
        title=title,
        body=body,
    )


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except (OSError, ValueError) as error:
        raise RunFailure(f"cannot hash file {path}: {error}") from error
    return digest.hexdigest()


def _sha256_stream(stream: Any) -> dict[str, Any]:
    digest = hashlib.sha256()
    total = 0
    try:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            total += len(chunk)
            digest.update(chunk)
    except (OSError, ValueError) as error:
        raise RunFailure(f"cannot hash fixture content stream: {error}") from error
    return {"bytes": total, "sha256": digest.hexdigest()}


def _fixture_content_identity(path: Path) -> dict[str, Any]:
    try:
        with zipfile.ZipFile(path) as archive:
            entries = [entry for entry in archive.infolist() if entry.filename == "content.xml"]
            if len(entries) != 1:
                raise RunFailure(
                    f"fixture must contain exactly one content.xml entry, observed {len(entries)}"
                )
            with archive.open(entries[0], "r") as stream:
                identity = _sha256_stream(stream)
    except RunFailure:
        raise
    except (
        OSError,
        EOFError,
        KeyError,
        RuntimeError,
        ValueError,
        zipfile.BadZipFile,
    ) as error:
        raise RunFailure(f"cannot read fixture content.xml: {error}") from error
    if identity["bytes"] <= 0:
        raise RunFailure("fixture content.xml is empty")
    return identity


def is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _run_git(root: Path, *arguments: str) -> str:
    try:
        completed = subprocess.run(
            ["git", "-C", str(root), *arguments],
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            check=False,
            timeout=10,
        )
    except (OSError, subprocess.SubprocessError) as error:
        raise RunFailure(f"cannot invoke git for {root}: {error}") from error
    if completed.returncode != 0:
        detail = completed.stderr.strip() or completed.stdout.strip()
        raise RunFailure(f"git {' '.join(arguments)} failed for {root}: {detail}")
    return completed.stdout.strip()


def _validate_external_target(path: Path, source_roots: Iterable[Path], label: str) -> None:
    for root in source_roots:
        if is_relative_to(path, root):
            raise RunFailure(f"{label} {path} must be outside source worktree {root}")


def _reject_overlapping_paths(first: Path, second: Path, label: str) -> None:
    if is_relative_to(first, second) or is_relative_to(second, first):
        raise RunFailure(f"{label} paths must not overlap: {first} and {second}")


def resolve_endpoint(endpoint: Endpoint, source_roots: Iterable[Path]) -> ResolvedEndpoint:
    if not endpoint.root.is_dir():
        raise RunFailure(f"{endpoint.name}.root is not a directory: {endpoint.root}")
    status = _run_git(endpoint.root, "status", "--porcelain", "--untracked-files=all")
    if status:
        raise RunFailure(f"{endpoint.name}.root is dirty; qualification requires a clean worktree")
    resolved_ref = _run_git(endpoint.root, "rev-parse", "--verify", f"{endpoint.requested_ref}^{{commit}}")
    head = _run_git(endpoint.root, "rev-parse", "--verify", "HEAD")
    if REVISION.fullmatch(resolved_ref) is None or REVISION.fullmatch(head) is None:
        raise RunFailure(
            f"{endpoint.name}: ref/head resolved to invalid revisions {resolved_ref!r}/{head!r}"
        )
    if head != resolved_ref:
        raise RunFailure(
            f"{endpoint.name}: clean worktree HEAD {head} does not match requested ref {resolved_ref}"
        )
    if not endpoint.binary.is_file() or not os.access(endpoint.binary, os.X_OK):
        raise RunFailure(f"{endpoint.name}.binary is not an executable regular file: {endpoint.binary}")
    if not endpoint.target_dir.is_dir():
        raise RunFailure(f"{endpoint.name}.target_dir is not a directory: {endpoint.target_dir}")
    _validate_external_target(endpoint.target_dir, source_roots, f"{endpoint.name}.target_dir")
    if not is_relative_to(endpoint.binary, endpoint.target_dir):
        raise RunFailure(f"{endpoint.name}.binary must be below its declared target_dir")
    identity = binary_identity(endpoint.binary)
    return ResolvedEndpoint(
        endpoint=endpoint,
        resolved_ref=resolved_ref,
        binary_bytes=identity["bytes"],
        binary_sha256=identity["sha256"],
    )


def validate_fixture(fixture: Fixture) -> dict[str, Any]:
    if not fixture.path.is_file():
        raise RunFailure(f"fixture is not a regular file: {fixture.path}")
    try:
        actual_bytes = fixture.path.stat().st_size
    except (OSError, ValueError) as error:
        raise RunFailure(f"cannot stat fixture {fixture.path}: {error}") from error
    actual_sha256 = sha256_file(fixture.path)
    if actual_bytes != fixture.bytes:
        raise RunFailure(f"fixture byte count changed: manifest {fixture.bytes}, observed {actual_bytes}")
    if actual_sha256 != fixture.sha256:
        raise RunFailure(f"fixture SHA-256 changed: manifest {fixture.sha256}, observed {actual_sha256}")
    content = _fixture_content_identity(fixture.path)
    return {
        "path": str(fixture.path),
        "bytes": actual_bytes,
        "sha256": actual_sha256,
        "content": content,
        "provenance": fixture.provenance,
    }


def binary_identity(path: Path) -> dict[str, Any]:
    if not path.is_file():
        raise RunFailure(f"binary is not a regular file: {path}")
    try:
        bytes_value = path.stat().st_size
    except (OSError, ValueError) as error:
        raise RunFailure(f"cannot stat binary {path}: {error}") from error
    return {"path": str(path), "bytes": bytes_value, "sha256": sha256_file(path)}


def is_heavy_process(executable: str, command: str) -> bool:
    if executable in HEAVY_EXECUTABLES or executable.startswith(PROBE_PREFIXES):
        return True
    for token in command.split():
        token_name = Path(token.strip("'\"")).name
        if token_name.startswith(PROBE_PREFIXES):
            return True
    return bool(HEAVY_WORD.search(command))


def _read_process_cmdline(entry: Path, pid: int) -> bytes | None:
    try:
        return (entry / "cmdline").read_bytes()
    except FileNotFoundError as error:
        # A PID that vanished between enumeration and read is benign, but only
        # after an explicit second stat confirms that it really vanished.
        try:
            entry.stat()
        except FileNotFoundError:
            return None
        except OSError as stat_error:
            raise RunFailure(f"cannot recheck /proc/{pid} after cmdline race: {stat_error}") from stat_error
        raise RunFailure(f"/proc/{pid}/cmdline disappeared without the process vanishing") from error
    except PermissionError as error:
        raise RunFailure(f"cannot read /proc/{pid}/cmdline: permission denied") from error
    except OSError as error:
        raise RunFailure(f"cannot read /proc/{pid}/cmdline: {error}") from error


def process_scan() -> list[dict[str, Any]]:
    processes: list[dict[str, Any]] = []
    proc_root = Path("/proc")
    try:
        entries = list(proc_root.glob("[0-9]*"))
    except OSError as error:
        raise RunFailure(f"cannot enumerate /proc: {error}") from error
    for entry in entries:
        try:
            pid = int(entry.name)
        except ValueError as error:
            raise RunFailure(f"/proc contained a non-numeric process entry {entry.name!r}") from error
        if pid == os.getpid():
            continue
        raw = _read_process_cmdline(entry, pid)
        if raw is None:
            continue
        parts = [part.decode("utf-8", errors="replace") for part in raw.split(b"\0") if part]
        if not parts:
            continue
        command = " ".join(parts)
        executable = Path(parts[0]).name
        if is_heavy_process(executable, command):
            processes.append({"pid": pid, "executable": executable, "command": parts})
    return sorted(processes, key=lambda value: value["pid"])


def load_average() -> float:
    try:
        return float(Path("/proc/loadavg").read_text(encoding="ascii").split()[0])
    except (OSError, IndexError, ValueError) as error:
        raise RunFailure(f"cannot read /proc/loadavg: {error}") from error


def quiet_guard(protocol: Protocol, phase: str) -> dict[str, Any]:
    load = load_average()
    processes = process_scan()
    result = {
        "phase": phase,
        "load_1m": load,
        "max_load": protocol.max_load,
        "heavy_processes": processes,
        "checked_unix_ns": time.time_ns(),
    }
    if load > protocol.max_load:
        raise RunFailure(f"{phase}: one-minute load {load} exceeds {protocol.max_load}")
    if processes:
        raise RunFailure(f"{phase}: protected build/profiling processes are running: {processes}")
    return result


def _require_taskset(cpu: int) -> str:
    executable = shutil.which("taskset")
    if executable is None:
        raise RunFailure("taskset is required for pinned native qualification")
    try:
        affinity = os.sched_getaffinity(0)
    except (AttributeError, OSError) as error:
        raise RunFailure(f"cannot inspect current CPU affinity: {error}") from error
    if cpu not in affinity:
        raise RunFailure(f"requested CPU {cpu} is not in the runner affinity {sorted(affinity)}")
    return executable


def _probe_command(
    endpoint: ResolvedEndpoint,
    manifest: InputManifest,
    mode: str,
    *,
    route: str | None = None,
    iterations: int | None = None,
    samples: int | None = None,
    warmups: int | None = None,
    target_ns: int | None = None,
) -> list[str]:
    command = [
        str(endpoint.endpoint.binary),
        "--mode",
        mode,
        "--fixture",
        str(manifest.fixture.path),
        "--title",
        manifest.title,
        "--body",
        manifest.body,
    ]
    if route is not None:
        command.extend(("--route", route))
    if iterations is not None:
        command.extend(("--iterations", str(iterations)))
    if samples is not None:
        command.extend(("--samples", str(samples)))
    if warmups is not None:
        command.extend(("--warmups", str(warmups)))
    if target_ns is not None:
        command.extend(("--target-ns", str(target_ns)))
    command.extend(("--max-iterations", str(manifest.protocol.max_iterations)))
    return command


def _parse_child_json(stdout: bytes, label: str) -> dict[str, Any]:
    if len(stdout) > MAX_CHILD_OUTPUT:
        raise RunFailure(f"{label}: child output exceeds {MAX_CHILD_OUTPUT} bytes")

    def reject_duplicate_keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in pairs:
            if key in result:
                raise RunFailure(f"{label}: child JSON contains duplicate key {key!r}")
            result[key] = value
        return result

    try:
        value = json.loads(
            stdout.decode("utf-8"), object_pairs_hook=reject_duplicate_keys
        )
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunFailure(f"{label}: child did not produce one JSON object: {error}") from error
    if not isinstance(value, dict):
        raise RunFailure(f"{label}: child JSON is not an object")
    return value


def run_child(
    endpoint: ResolvedEndpoint,
    manifest: InputManifest,
    mode: str,
    *,
    route: str | None = None,
    iterations: int | None = None,
    samples: int | None = None,
    warmups: int | None = None,
    target_ns: int | None = None,
    tmp_dir: Path,
    guard: dict[str, Any] | None = None,
    expected_schema: str | None = None,
) -> tuple[dict[str, Any], dict[str, Any]]:
    command = _probe_command(
        endpoint,
        manifest,
        mode,
        route=route,
        iterations=iterations,
        samples=samples,
        warmups=warmups,
        target_ns=target_ns,
    )
    taskset = _require_taskset(manifest.protocol.cpu)
    command = [taskset, "--cpu-list", str(manifest.protocol.cpu), *command]
    environment = os.environ.copy()
    environment["TMPDIR"] = str(tmp_dir)
    started = time.monotonic_ns()
    try:
        completed = subprocess.run(
            command,
            cwd=endpoint.endpoint.root,
            env=environment,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=manifest.protocol.timeout_s,
            check=False,
        )
    except (OSError, subprocess.SubprocessError) as error:
        raise RunFailure(
            f"{endpoint.endpoint.name}/{mode}/{route}: child launch failed: {error}"
        ) from error
    wall_ns = time.monotonic_ns() - started
    if completed.returncode != 0:
        stderr = completed.stderr.decode("utf-8", errors="replace")[:16_384]
        raise RunFailure(
            f"{endpoint.endpoint.name}/{mode}/{route}: child exited {completed.returncode}: {stderr}"
        )
    report = _parse_child_json(
        completed.stdout, f"{endpoint.endpoint.name}/{mode}/{route}"
    )
    if expected_schema is None:
        expected_schema = PROBE_PROFILES[endpoint.endpoint.profile]
    if report.get("schema") != expected_schema:
        raise RunFailure(
            f"{endpoint.endpoint.name}/{mode}/{route}: probe schema mismatch; "
            f"expected {expected_schema!r}, observed {report.get('schema')!r}"
        )
    event = {
        "endpoint": endpoint.endpoint.name,
        "profile": endpoint.endpoint.profile,
        "probe_schema": expected_schema,
        "mode": mode,
        "route": route,
        "command": command,
        "returncode": completed.returncode,
        "wall_ns": wall_ns,
        "stdout_sha256": hashlib.sha256(completed.stdout).hexdigest(),
        "stderr": completed.stderr.decode("utf-8", errors="replace")[:16_384],
        "guard": guard,
        "report": report,
    }
    return report, event


def _digest_object(value: Any, location: str) -> dict[str, Any]:
    item = _object(value, location)
    bytes_value = _positive_int(item.get("bytes"), f"{location}.bytes")
    digest = _digest(item.get("sha256"), f"{location}.sha256")
    return {"bytes": bytes_value, "sha256": digest}


def _scoped_replay_receipt(report: Mapping[str, Any], label: str) -> dict[str, Any]:
    scoped = _object(report.get("scoped_replay"), f"{label}.scoped_replay")
    expected_keys = {
        "provenance",
        "member_name",
        "recipient_archive",
        "source_member",
        "recipient_member",
        "forward_member",
        "inverse_member",
        "recipient_content",
        "forward_content",
        "inverse_content",
    }
    if set(scoped) != expected_keys:
        raise RunFailure(f"{label}: scoped replay receipt fields are incomplete")
    provenance = _string(scoped.get("provenance"), f"{label}.scoped_replay.provenance")
    member_name = _string(scoped.get("member_name"), f"{label}.scoped_replay.member_name")
    source_value = scoped.get("source_member")
    source_member = (
        None
        if source_value is None
        else _digest_object(source_value, f"{label}.scoped_replay.source_member")
    )
    return {
        "provenance": provenance,
        "member_name": member_name,
        "recipient_archive": _digest_object(
            scoped.get("recipient_archive"), f"{label}.scoped_replay.recipient_archive"
        ),
        "source_member": source_member,
        "recipient_member": _digest_object(
            scoped.get("recipient_member"), f"{label}.scoped_replay.recipient_member"
        ),
        "forward_member": _digest_object(
            scoped.get("forward_member"), f"{label}.scoped_replay.forward_member"
        ),
        "inverse_member": _digest_object(
            scoped.get("inverse_member"), f"{label}.scoped_replay.inverse_member"
        ),
        "recipient_content": _digest_object(
            scoped.get("recipient_content"), f"{label}.scoped_replay.recipient_content"
        ),
        "forward_content": _digest_object(
            scoped.get("forward_content"), f"{label}.scoped_replay.forward_content"
        ),
        "inverse_content": _digest_object(
            scoped.get("inverse_content"), f"{label}.scoped_replay.inverse_content"
        ),
    }


def _assert_endpoint_unchanged(endpoint: ResolvedEndpoint) -> dict[str, Any]:
    status = _run_git(endpoint.endpoint.root, "status", "--porcelain", "--untracked-files=all")
    if status:
        raise RunFailure(f"{endpoint.endpoint.name}: source worktree became dirty before timing")
    current_ref = _run_git(
        endpoint.endpoint.root,
        "rev-parse",
        "--verify",
        f"{endpoint.endpoint.requested_ref}^{{commit}}",
    )
    current_head = _run_git(endpoint.endpoint.root, "rev-parse", "--verify", "HEAD")
    if current_ref != endpoint.resolved_ref or current_head != endpoint.resolved_ref:
        raise RunFailure(
            f"{endpoint.endpoint.name}: ref/head changed from {endpoint.resolved_ref} to {current_ref}/{current_head}"
        )
    identity = binary_identity(endpoint.endpoint.binary)
    expected = {
        "path": str(endpoint.endpoint.binary),
        "bytes": endpoint.binary_bytes,
        "sha256": endpoint.binary_sha256,
    }
    if identity != expected:
        raise RunFailure(f"{endpoint.endpoint.name}: executable changed after identity capture")
    return {"head": current_head, "binary": identity}


def _revalidate_run_identity(
    manifest: InputManifest,
    endpoints: Mapping[str, ResolvedEndpoint],
    phase: str,
) -> dict[str, Any]:
    fixture = validate_fixture(manifest.fixture)
    endpoint_identity = {
        name: _assert_endpoint_unchanged(endpoints[name])
        for name in ("baseline", "candidate")
    }
    return {
        "phase": phase,
        "checked_unix_ns": time.time_ns(),
        "fixture": fixture,
        "endpoints": endpoint_identity,
    }


def validate_correctness_report(
    report: Mapping[str, Any],
    fixture: Fixture,
    fixture_identity: Mapping[str, Any],
    label: str,
) -> dict[str, dict[str, Any]]:
    if report.get("schema") != PROBE_SCHEMA or report.get("mode") != "verify":
        raise RunFailure(f"{label}: invalid correctness report identity")
    fixture_report = _digest_object(report.get("fixture"), f"{label}.fixture")
    if fixture_report != {"bytes": fixture.bytes, "sha256": fixture.sha256}:
        raise RunFailure(f"{label}: probe fixture identity differs from manifest")
    content_report = _digest_object(
        _object(report.get("fixture"), f"{label}.fixture").get("content"),
        f"{label}.fixture.content",
    )
    expected_content = _digest_object(fixture_identity.get("content"), f"{label}.fixture.content")
    if content_report != expected_content:
        raise RunFailure(f"{label}: probe content.xml identity differs from the fixture")
    checks = _object(report.get("checks"), f"{label}.checks")
    expected_checks = {
        "no_op_exact_source",
        "read_semantics",
        "append_changed",
        "patch_source_matches_fixture",
        "append_output_changed",
        "scoped_recipient_content_matches",
        "scoped_recipient_member_changed",
        "scoped_forward_content_matches_patch",
        "scoped_forward_preserves_unrelated_member",
        "scoped_inverse_preserves_unrelated_member",
        "scoped_inverse_restores_content_scope",
        "same_owner_replay_exact",
        "content_matching_forward_exact",
        "whole_owner_inverse_exact_source",
        "inverse_content_exact_source",
        "changed_recipient_refused",
    }
    if set(checks) != expected_checks or any(value is not True for value in checks.values()):
        raise RunFailure(f"{label}: correctness checks are incomplete or false")
    patch = _object(report.get("patch"), f"{label}.patch")
    if patch.get("is_noop") is not False:
        raise RunFailure(f"{label}: append patch unexpectedly reports noop")
    source_content_sha256 = _digest(
        patch.get("source_content_sha256"), f"{label}.patch.source_content_sha256"
    )
    target_content_sha256 = _digest(
        patch.get("target_content_sha256"), f"{label}.patch.target_content_sha256"
    )
    if source_content_sha256 == target_content_sha256:
        raise RunFailure(f"{label}: append patch content proof did not change")
    if source_content_sha256 != expected_content["sha256"]:
        raise RunFailure(f"{label}: patch source hash differs from fixture content.xml")
    read = _object(report.get("read"), f"{label}.read")
    if _positive_int(read.get("slide_count"), f"{label}.read.slide_count") <= 0:
        raise RunFailure(f"{label}: read route reported no slides")
    if type(read.get("first_slide_present")) is not bool:
        raise RunFailure(f"{label}.read.first_slide_present must be a boolean")
    _nonnegative_int(read.get("text_bytes"), f"{label}.read.text_bytes")
    _digest(read.get("text_sha256"), f"{label}.read.text_sha256")
    outputs_value = _object(report.get("outputs"), f"{label}.outputs")
    expected_outputs = {
        "source",
        "no_op",
        "append",
        "same_owner_replay",
        "content_matching_replay",
        "whole_owner_inverse",
        "inverse",
        "inverse_content",
    }
    if set(outputs_value) != expected_outputs:
        raise RunFailure(f"{label}: output digest set differs from the contract")
    outputs = {
        name: _digest_object(value, f"{label}.outputs.{name}")
        for name, value in outputs_value.items()
    }
    if outputs["source"] != {"bytes": fixture.bytes, "sha256": fixture.sha256}:
        raise RunFailure(f"{label}: source output digest differs from the fixture")
    if outputs["no_op"] != outputs["source"]:
        raise RunFailure(f"{label}: no-op output is not the exact source")
    if outputs["whole_owner_inverse"] != outputs["source"]:
        raise RunFailure(f"{label}: same-owner inverse output is not the exact source")
    if outputs["inverse_content"] != expected_content:
        raise RunFailure(f"{label}: inverse replay did not restore content.xml")
    if outputs["append"] == outputs["source"]:
        raise RunFailure(f"{label}: appended output is identical to the source")
    for name in ("same_owner_replay", "content_matching_replay"):
        if outputs[name] != outputs["append"]:
            raise RunFailure(f"{label}: {name} output differs from append output")
    scoped = _scoped_replay_receipt(report, label)
    if scoped["recipient_content"] != expected_content:
        raise RunFailure(f"{label}: scoped recipient content differs from the fixture")
    if scoped["recipient_archive"] == outputs["source"]:
        raise RunFailure(f"{label}: scoped recipient archive was not changed")
    if scoped["source_member"] is not None and scoped["source_member"] == scoped["recipient_member"]:
        raise RunFailure(f"{label}: scoped recipient member was not changed")
    if scoped["forward_member"] != scoped["recipient_member"]:
        raise RunFailure(f"{label}: forward replay did not preserve the unrelated member")
    if scoped["inverse_member"] != scoped["recipient_member"]:
        raise RunFailure(f"{label}: inverse replay did not preserve the unrelated member")
    if scoped["forward_content"]["sha256"] != target_content_sha256:
        raise RunFailure(f"{label}: scoped forward content differs from the patch target")
    if scoped["inverse_content"] != expected_content:
        raise RunFailure(f"{label}: scoped inverse did not restore source content")
    return outputs


def validate_legacy_correctness_report(
    report: Mapping[str, Any],
    fixture: Fixture,
    fixture_identity: Mapping[str, Any],
    label: str,
) -> dict[str, Any]:
    """Validate the pre-Patch e4a adapter without claiming candidate proofs."""

    if report.get("schema") != LEGACY_PROBE_SCHEMA or report.get("mode") != "verify":
        raise RunFailure(f"{label}: invalid legacy correctness report identity")
    fixture_report = _digest_object(report.get("fixture"), f"{label}.fixture")
    expected_fixture = {"bytes": fixture.bytes, "sha256": fixture.sha256}
    if fixture_report != expected_fixture:
        raise RunFailure(f"{label}: probe fixture identity differs from manifest")
    content_report = _digest_object(
        _object(report.get("fixture"), f"{label}.fixture").get("content"),
        f"{label}.fixture.content",
    )
    expected_content = _digest_object(fixture_identity.get("content"), f"{label}.fixture.content")
    if content_report != expected_content:
        raise RunFailure(f"{label}: probe content.xml identity differs from the fixture")

    expected_capabilities = {
        "patch_replay": False,
        "inverse_replay": False,
        "scoped_replay": False,
        "timed_routes": list(ROUTES),
        "replay_scope": "independent_plan_not_patch",
    }
    capabilities = _object(report.get("capabilities"), f"{label}.capabilities")
    if capabilities != expected_capabilities:
        raise RunFailure(
            f"{label}: legacy capability boundary is incomplete or claims Patch replay"
        )

    checks = _object(report.get("checks"), f"{label}.checks")
    expected_checks = {
        "no_op_exact_source",
        "read_semantics",
        "append_changed",
        "patch_source_matches_fixture",
        "append_output_changed",
        "independent_plan_repeat_exact",
    }
    if set(checks) != expected_checks or any(value is not True for value in checks.values()):
        raise RunFailure(f"{label}: legacy correctness checks are incomplete or false")

    patch_value = _object(report.get("patch"), f"{label}.patch")
    if set(patch_value) != {
        "is_noop",
        "source_content_sha256",
        "target_content_sha256",
    }:
        raise RunFailure(f"{label}: legacy patch proof fields are incomplete")
    if patch_value["is_noop"] is not False:
        raise RunFailure(f"{label}: legacy append plan unexpectedly reports noop")
    patch = {
        "is_noop": False,
        "source_content_sha256": _digest(
            patch_value["source_content_sha256"],
            f"{label}.patch.source_content_sha256",
        ),
        "target_content_sha256": _digest(
            patch_value["target_content_sha256"],
            f"{label}.patch.target_content_sha256",
        ),
    }
    if patch["source_content_sha256"] != expected_content["sha256"]:
        raise RunFailure(f"{label}: legacy plan source hash differs from fixture content.xml")
    if patch["source_content_sha256"] == patch["target_content_sha256"]:
        raise RunFailure(f"{label}: legacy plan content proof did not change")

    read_value = _object(report.get("read"), f"{label}.read")
    slide_count = _positive_int(read_value.get("slide_count"), f"{label}.read.slide_count")
    if not read_value.get("first_slide_present") is True:
        raise RunFailure(f"{label}: legacy read route reported no first slide")
    read = {
        "slide_count": slide_count,
        "first_slide_present": True,
        "text_bytes": _nonnegative_int(read_value.get("text_bytes"), f"{label}.read.text_bytes"),
        "text_sha256": _digest(read_value.get("text_sha256"), f"{label}.read.text_sha256"),
    }

    outputs_value = _object(report.get("outputs"), f"{label}.outputs")
    expected_outputs = {"source", "no_op", "append", "independent_append"}
    if set(outputs_value) != expected_outputs:
        raise RunFailure(f"{label}: legacy output digest set differs from the contract")
    outputs = {
        name: _digest_object(value, f"{label}.outputs.{name}")
        for name, value in outputs_value.items()
    }
    if outputs["source"] != expected_fixture:
        raise RunFailure(f"{label}: legacy source output digest differs from the fixture")
    if outputs["no_op"] != outputs["source"]:
        raise RunFailure(f"{label}: legacy no-op output is not the exact source")
    if outputs["append"] == outputs["source"]:
        raise RunFailure(f"{label}: legacy appended output is identical to the source")
    if outputs["independent_append"] != outputs["append"]:
        raise RunFailure(f"{label}: independent legacy plan output differs from append output")
    return {
        "profile": LEGACY_PROFILE,
        "schema": LEGACY_PROBE_SCHEMA,
        "capabilities": capabilities,
        "outputs": outputs,
        "read": read,
        "patch": patch,
    }


def _route_proofs(
    correctness: Mapping[str, Any], endpoint_name: str
) -> dict[str, dict[str, Any]]:
    endpoints = _object(correctness.get("endpoints"), "correctness.endpoints")
    endpoint = _object(
        endpoints.get(endpoint_name), f"correctness.endpoints.{endpoint_name}"
    )
    outputs = _object(endpoint.get("outputs"), f"correctness.{endpoint_name}.outputs")
    read = _object(endpoint.get("read"), f"correctness.{endpoint_name}.read")
    return {
        "open": {"kind": "open", "source": outputs["source"]},
        "read": {
            "kind": "read",
            "slide_count": read["slide_count"],
            "text": {
                "bytes": read["text_bytes"],
                "sha256": read["text_sha256"],
            },
        },
        "write": {"kind": "package", "output": outputs["no_op"]},
        "append": {"kind": "package", "output": outputs["append"]},
    }


def _route_output_bytes(route: str, proofs: Mapping[str, Mapping[str, Any]]) -> int:
    if route in ("open", "read"):
        return 0
    output = _object(proofs[route].get("output"), f"proofs.{route}.output")
    return _positive_int(output.get("bytes"), f"proofs.{route}.output.bytes")


def _write_json(path: Path, value: Any) -> None:
    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    with temporary.open("rb") as stream:
        os.fsync(stream.fileno())
    os.replace(temporary, path)


def _stats(values: list[float]) -> dict[str, Any]:
    if not values or any(value <= 0 or not math.isfinite(value) for value in values):
        raise RunFailure("timing statistics require positive finite values")
    ordered = sorted(values)
    return {
        "count": len(values),
        "min_ns": ordered[0],
        "p50_ns": ordered[round((len(ordered) - 1) * 0.50)],
        "p95_ns": ordered[round((len(ordered) - 1) * 0.95)],
        "max_ns": ordered[-1],
        "mean_ns": sum(values) / len(values),
    }


def validate_calibration(
    report: Mapping[str, Any],
    route: str,
    protocol: Protocol,
    expected_proof: Mapping[str, Any],
    label: str,
) -> int:
    if report.get("mode") != "calibrate" or report.get("route") != route:
        raise RunFailure(f"{label}: calibration identity mismatch")
    if report.get("target_ns") != protocol.target_ns:
        raise RunFailure(f"{label}: calibration target differs from the manifest")
    if report.get("proof") != expected_proof:
        raise RunFailure(f"{label}: calibration output proof differs from correctness preflight")
    iterations = _positive_int(report.get("iterations"), f"{label}.iterations")
    elapsed_ns = _positive_int(report.get("elapsed_ns"), f"{label}.elapsed_ns")
    if iterations > protocol.max_iterations or elapsed_ns < protocol.target_ns:
        raise RunFailure(f"{label}: calibration did not satisfy its target")
    output_bytes = _nonnegative_int(report.get("output_bytes"), f"{label}.output_bytes")
    expected_bytes = _route_output_bytes(route, {route: expected_proof})
    if output_bytes != expected_bytes:
        raise RunFailure(f"{label}: calibration output byte proof differs from correctness preflight")
    attempts = report.get("attempts")
    if not isinstance(attempts, list) or not attempts:
        raise RunFailure(f"{label}: calibration attempts are missing")
    previous_iterations = 0
    for index, attempt_value in enumerate(attempts):
        attempt = _object(attempt_value, f"{label}.attempts[{index}]")
        attempt_iterations = _positive_int(
            attempt.get("iterations"), f"{label}.attempts[{index}].iterations"
        )
        attempt_elapsed = _positive_int(
            attempt.get("elapsed_ns"), f"{label}.attempts[{index}].elapsed_ns"
        )
        attempt_output_bytes = _nonnegative_int(
            attempt.get("output_bytes"), f"{label}.attempts[{index}].output_bytes"
        )
        if attempt_output_bytes != expected_bytes:
            raise RunFailure(
                f"{label}: calibration attempt output proof differs from correctness preflight"
            )
        if attempt_iterations <= previous_iterations or attempt_iterations > protocol.max_iterations:
            raise RunFailure(f"{label}: calibration iterations are not strictly increasing")
        previous_iterations = attempt_iterations
        if index == len(attempts) - 1 and (
            attempt_iterations != iterations or attempt_elapsed != elapsed_ns
        ):
            raise RunFailure(f"{label}: final calibration attempt differs from its summary")
    return iterations


def validate_timed(
    report: Mapping[str, Any],
    route: str,
    iterations: int,
    protocol: Protocol,
    expected_proof: Mapping[str, Any],
    label: str,
) -> tuple[float, int]:
    if report.get("mode") != "timed" or report.get("route") != route:
        raise RunFailure(f"{label}: timed identity mismatch")
    if report.get("iterations") != iterations or report.get("warmups") != protocol.warmups:
        raise RunFailure(f"{label}: timed normalization mismatch")
    if report.get("proof") != expected_proof:
        raise RunFailure(f"{label}: timed output proof differs from correctness preflight")
    samples = report.get("samples")
    if not isinstance(samples, list) or len(samples) != 1:
        raise RunFailure(f"{label}: timed child must return one sample")
    elapsed_ns = _positive_int(samples[0], f"{label}.samples[0]")
    if elapsed_ns < protocol.minimum_ns:
        raise RunFailure(f"{label}: timed batch is shorter than minimum_ns")
    output_bytes = _nonnegative_int(report.get("output_bytes"), f"{label}.output_bytes")
    expected_bytes = _route_output_bytes(route, {route: expected_proof})
    if output_bytes != expected_bytes:
        raise RunFailure(f"{label}: timed output byte proof differs from correctness preflight")
    return elapsed_ns / iterations, output_bytes


def _validate_artifact_layout(manifest: InputManifest, out_dir: Path) -> None:
    roots = (manifest.baseline.root, manifest.candidate.root)
    _validate_external_target(out_dir, roots, "artifact directory")
    _reject_overlapping_paths(
        out_dir,
        manifest.baseline.target_dir,
        "artifact and baseline target",
    )
    _reject_overlapping_paths(
        out_dir,
        manifest.candidate.target_dir,
        "artifact and candidate target",
    )
    _reject_overlapping_paths(
        manifest.baseline.target_dir,
        manifest.candidate.target_dir,
        "baseline and candidate target",
    )


def _runtime_identity(manifest: InputManifest, out_dir: Path) -> dict[str, Any]:
    _validate_artifact_layout(manifest, out_dir)
    roots = (manifest.baseline.root, manifest.candidate.root)
    fixture = validate_fixture(manifest.fixture)
    if manifest.baseline.root == manifest.candidate.root:
        raise RunFailure("baseline and candidate roots must be distinct")
    baseline = resolve_endpoint(manifest.baseline, roots)
    candidate = resolve_endpoint(manifest.candidate, roots)
    if baseline.endpoint.profile != LEGACY_PROFILE:
        raise RunFailure(
            f"baseline.profile must be {LEGACY_PROFILE!r}; "
            f"observed {baseline.endpoint.profile!r}"
        )
    if candidate.endpoint.profile != CURRENT_PROFILE:
        raise RunFailure(
            f"candidate.profile must be {CURRENT_PROFILE!r}; "
            f"observed {candidate.endpoint.profile!r}"
        )
    if baseline.resolved_ref == candidate.resolved_ref:
        raise RunFailure("baseline and candidate refs resolve to the same commit")
    return {
        "fixture": fixture,
        "baseline": {
            "root": str(baseline.endpoint.root),
            "requested_ref": baseline.endpoint.requested_ref,
            "resolved_ref": baseline.resolved_ref,
            "profile": baseline.endpoint.profile,
            "binary": binary_identity(baseline.endpoint.binary),
            "target_dir": str(baseline.endpoint.target_dir),
        },
        "candidate": {
            "root": str(candidate.endpoint.root),
            "requested_ref": candidate.endpoint.requested_ref,
            "resolved_ref": candidate.resolved_ref,
            "profile": candidate.endpoint.profile,
            "binary": binary_identity(candidate.endpoint.binary),
            "target_dir": str(candidate.endpoint.target_dir),
        },
        "endpoints": {"baseline": baseline, "candidate": candidate},
    }


def _endpoint_metadata(endpoint: ResolvedEndpoint) -> dict[str, Any]:
    return {
        "root": str(endpoint.endpoint.root),
        "requested_ref": endpoint.endpoint.requested_ref,
        "resolved_ref": endpoint.resolved_ref,
        "head": endpoint.resolved_ref,
        "profile": endpoint.endpoint.profile,
        "probe_schema": PROBE_PROFILES[endpoint.endpoint.profile],
        "binary": {
            "path": str(endpoint.endpoint.binary),
            "bytes": endpoint.binary_bytes,
            "sha256": endpoint.binary_sha256,
        },
        "target_dir": str(endpoint.endpoint.target_dir),
    }


def _run_correctness(
    manifest: InputManifest,
    endpoints: Mapping[str, ResolvedEndpoint],
    fixture_identity: Mapping[str, Any],
    out_dir: Path,
    tmp_dir: Path,
) -> dict[str, Any]:
    endpoint_records: dict[str, dict[str, Any]] = {}
    events = []
    for name in ("baseline", "candidate"):
        # Correctness-only is an untimed proof preflight.  It may run while
        # unrelated host work is present; strict quiet-host guards remain in
        # calibration and ABBA timing below.
        profile = endpoints[name].endpoint.profile
        expected_schema = PROBE_PROFILES[profile]
        report, event = run_child(
            endpoints[name],
            manifest,
            "verify",
            tmp_dir=tmp_dir,
            expected_schema=expected_schema,
        )
        if name == "baseline":
            record = validate_legacy_correctness_report(
                report,
                manifest.fixture,
                fixture_identity,
                f"correctness.{name}",
            )
        else:
            outputs = validate_correctness_report(
                report,
                manifest.fixture,
                fixture_identity,
                f"correctness.{name}",
            )
            read = _object(report["read"], f"correctness.{name}.read")
            patch = _object(report["patch"], f"correctness.{name}.patch")
            record = {
                "profile": CURRENT_PROFILE,
                "schema": PROBE_SCHEMA,
                "capabilities": {
                    "patch_replay": True,
                    "inverse_replay": True,
                    "scoped_replay": True,
                    "timed_routes": list(ROUTES),
                    "replay_scope": "content_matching_patch",
                },
                "outputs": outputs,
                "read": {
                    "slide_count": read["slide_count"],
                    "first_slide_present": read["first_slide_present"],
                    "text_bytes": read["text_bytes"],
                    "text_sha256": read["text_sha256"],
                },
                "patch": {
                    "is_noop": patch["is_noop"],
                    "source_content_sha256": patch["source_content_sha256"],
                    "target_content_sha256": patch["target_content_sha256"],
                },
                "scoped_replay": _scoped_replay_receipt(
                    report, f"correctness.{name}"
                ),
            }
        if record["profile"] != profile or record["schema"] != expected_schema:
            raise RunFailure(f"correctness.{name}: report capability profile mismatch")
        endpoint_records[name] = record
        events.append(event)
        _write_json(out_dir / f"correctness-{name}.json", report)

    baseline = endpoint_records["baseline"]
    candidate = endpoint_records["candidate"]
    common_output_names = ("source", "no_op", "append")
    common_outputs = {
        name: baseline["outputs"][name] for name in common_output_names
    }
    if any(
        baseline["outputs"][name] != candidate["outputs"][name]
        for name in common_output_names
    ):
        raise RunFailure("baseline and candidate common output digests differ")
    if baseline["read"] != candidate["read"]:
        raise RunFailure("baseline and candidate read receipts differ")
    if baseline["patch"] != candidate["patch"]:
        raise RunFailure("baseline and candidate patch receipts differ")
    return {
        "endpoints": endpoint_records,
        "common": {
            "outputs": common_outputs,
            "read": baseline["read"],
            "patch": baseline["patch"],
        },
        # Keep the candidate's richer proof at the top level for report
        # readers; timing consumes the endpoint records so it cannot mistake
        # a candidate-only replay for a baseline capability.
        "outputs": candidate["outputs"],
        "read": candidate["read"],
        "patch": candidate["patch"],
        "scoped_replay": candidate["scoped_replay"],
        "events": events,
    }


def _run_timing(
    manifest: InputManifest,
    endpoints: Mapping[str, ResolvedEndpoint],
    fixture_identity: Mapping[str, Any],
    correctness: Mapping[str, Any],
    out_dir: Path,
    tmp_dir: Path,
) -> dict[str, Any]:
    identity_before_timing = _revalidate_run_identity(
        manifest, endpoints, "before_calibration"
    )
    if identity_before_timing["fixture"] != fixture_identity:
        raise RunFailure("fixture identity changed before calibration")
    proofs = {
        name: _route_proofs(correctness, name)
        for name in ("baseline", "candidate")
    }
    calibration: dict[str, dict[str, int]] = {"baseline": {}, "candidate": {}}
    calibration_events = []
    for name in ("baseline", "candidate"):
        expected_schema = PROBE_PROFILES[endpoints[name].endpoint.profile]
        for route in ROUTES:
            guard = quiet_guard(manifest.protocol, f"calibration-{name}-{route}")
            report, event = run_child(
                endpoints[name],
                manifest,
                "calibrate",
                route=route,
                target_ns=manifest.protocol.target_ns,
                tmp_dir=tmp_dir,
                guard=guard,
                expected_schema=expected_schema,
            )
            calibration[name][route] = validate_calibration(
                report,
                route,
                manifest.protocol,
                proofs[name][route],
                f"calibration.{name}.{route}",
            )
            event["guard"] = guard
            calibration_events.append(event)
    identity_after_calibration = _revalidate_run_identity(
        manifest, endpoints, "after_calibration"
    )
    if identity_after_calibration["fixture"] != fixture_identity:
        raise RunFailure("fixture identity changed during calibration")
    _write_json(
        out_dir / "calibration.json",
        {
            "iterations": calibration,
            "events": calibration_events,
            "identity_after_calibration": identity_after_calibration,
        },
    )

    routes: dict[str, Any] = {}
    for route in ROUTES:
        blocks = []
        by_role: dict[str, list[float]] = {role: [] for role in ABBA_ORDER}
        for block in range(manifest.protocol.blocks):
            guard = quiet_guard(manifest.protocol, f"before-{route}-block-{block}")
            legs = []
            for role in ABBA_ORDER:
                endpoint_name = ROLE_TO_ENDPOINT[role]
                endpoint = endpoints[endpoint_name]
                expected_schema = PROBE_PROFILES[endpoint.endpoint.profile]
                leg_guard = quiet_guard(manifest.protocol, f"{route}-{role}-block-{block}")
                report, event = run_child(
                    endpoint,
                    manifest,
                    "timed",
                    route=route,
                    iterations=calibration[endpoint_name][route],
                    samples=1,
                    warmups=manifest.protocol.warmups,
                    tmp_dir=tmp_dir,
                    guard=leg_guard,
                    expected_schema=expected_schema,
                )
                per_op_ns, output_bytes = validate_timed(
                    report,
                    route,
                    calibration[endpoint_name][route],
                    manifest.protocol,
                    proofs[endpoint_name][route],
                    f"timed.{route}.{role}.{block}",
                )
                after_leg_guard = quiet_guard(
                    manifest.protocol, f"after-{route}-{role}-block-{block}"
                )
                by_role[role].append(per_op_ns)
                legs.append({
                    "role": role,
                    "endpoint": endpoint_name,
                    "block": block,
                    "iterations": calibration[endpoint_name][route],
                    "elapsed_ns": report["samples"][0],
                    "per_operation_ns": per_op_ns,
                    "output_bytes": output_bytes,
                    "event": event,
                    "guard_after": after_leg_guard,
                })
            output_lengths = {leg["output_bytes"] for leg in legs}
            if len(output_lengths) != 1:
                raise RunFailure(f"{route}/block-{block}: output length changed across ABBA legs")
            baseline_mean = (by_role["a1"][-1] + by_role["a2"][-1]) / 2
            candidate_mean = (by_role["b1"][-1] + by_role["b2"][-1]) / 2
            if baseline_mean <= 0:
                raise RunFailure(f"{route}/block-{block}: baseline mean is not positive")
            blocks.append({
                "block": block,
                "abba_order": list(ABBA_ORDER),
                "legs": legs,
                "baseline_mean_ns": baseline_mean,
                "candidate_mean_ns": candidate_mean,
                "candidate_baseline_ratio": candidate_mean / baseline_mean,
                "guard_before_block": guard,
            })
            _write_json(out_dir / f"{route}-block-{block:03d}.json", blocks[-1])
        routes[route] = {
            "blocks": blocks,
            "statistics_ns": {
                "baseline": _stats(by_role["a1"] + by_role["a2"]),
                "candidate": _stats(by_role["b1"] + by_role["b2"]),
            },
            "paired_ratios": [block["candidate_baseline_ratio"] for block in blocks],
            "paired_ratio_statistics": _stats(
                [block["candidate_baseline_ratio"] for block in blocks]
            ),
            "output_bytes": blocks[0]["legs"][0]["output_bytes"],
        }
    identity_after_abba = _revalidate_run_identity(manifest, endpoints, "after_abba")
    if identity_after_abba["fixture"] != fixture_identity:
        raise RunFailure("fixture identity changed during ABBA")
    return {
        "calibration": calibration,
        "proofs": proofs,
        "routes": routes,
        "identity": {
            "before_calibration": identity_before_timing,
            "after_calibration": identity_after_calibration,
            "after_abba": identity_after_abba,
        },
    }


def _prepare_output(path: Path, roots: Iterable[Path]) -> Path:
    path = path.resolve()
    _validate_external_target(path, roots, "artifact directory")
    if path.exists():
        if not path.is_dir():
            raise RunFailure(f"artifact path is not a directory: {path}")
        if any(path.iterdir()):
            raise RunFailure(f"artifact directory must be new and empty: {path}")
    else:
        path.mkdir(parents=True)
    return path


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--out-dir", type=Path, required=True)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument(
        "--plan",
        action="store_true",
        help="validate the manifest shape and write a plan without touching binaries or timing",
    )
    mode.add_argument(
        "--correctness-only",
        action="store_true",
        help="run the untimed correctness preflight and skip calibration and native timing",
    )
    return parser.parse_args(argv)


def _write_failure_receipt(path: Path, error: BaseException) -> None:
    try:
        out_dir = path.resolve()
        if out_dir.is_dir():
            _write_json(
                out_dir / "failure.json",
                {
                    "schema": REPORT_SCHEMA,
                    "status": "failed",
                    "error_type": type(error).__name__,
                    "error": str(error),
                },
            )
    except Exception:
        # A receipt cannot be written when the requested directory itself is
        # inaccessible.  The caller still returns a failed process status and
        # emits the original structured error text without a traceback.
        return


def _fail(args: argparse.Namespace, error: BaseException) -> int:
    _write_failure_receipt(args.out_dir, error)
    print(f"ODP native qualification failed: {error}", file=sys.stderr)
    return 1


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    try:
        manifest = load_manifest(args.manifest)
        out_dir = _prepare_output(args.out_dir, (manifest.baseline.root, manifest.candidate.root))
        _validate_artifact_layout(manifest, out_dir)
        if args.plan:
            _write_json(
                out_dir / "plan.json",
                {
                    "schema": REPORT_SCHEMA,
                    "status": "plan_only",
                    "manifest": str(manifest.path),
                    "fixture": {
                        "path": str(manifest.fixture.path),
                        "bytes": manifest.fixture.bytes,
                        "sha256": manifest.fixture.sha256,
                        "provenance": manifest.fixture.provenance,
                    },
                    "baseline": {
                        "root": str(manifest.baseline.root),
                        "requested_ref": manifest.baseline.requested_ref,
                        "profile": manifest.baseline.profile,
                        "probe_schema": PROBE_PROFILES[manifest.baseline.profile],
                        "binary": str(manifest.baseline.binary),
                        "target_dir": str(manifest.baseline.target_dir),
                    },
                    "candidate": {
                        "root": str(manifest.candidate.root),
                        "requested_ref": manifest.candidate.requested_ref,
                        "profile": manifest.candidate.profile,
                        "probe_schema": PROBE_PROFILES[manifest.candidate.profile],
                        "binary": str(manifest.candidate.binary),
                        "target_dir": str(manifest.candidate.target_dir),
                    },
                    "routes": list(ROUTES),
                    "abba_order": list(ABBA_ORDER),
                    "protocol": {
                        "cpu": manifest.protocol.cpu,
                        "blocks": manifest.protocol.blocks,
                        "warmups": manifest.protocol.warmups,
                        "target_ns": manifest.protocol.target_ns,
                        "minimum_ns": manifest.protocol.minimum_ns,
                        "max_load": manifest.protocol.max_load,
                        "max_iterations": manifest.protocol.max_iterations,
                        "timeout_s": manifest.protocol.timeout_s,
                    },
                },
            )
            return 0

        runtime = _runtime_identity(manifest, out_dir)
        endpoints = runtime["endpoints"]
        tmp_dir = out_dir / "tmp"
        tmp_dir.mkdir()
        identity_before_correctness = None
        if args.correctness_only:
            identity_before_correctness = _revalidate_run_identity(
                manifest, endpoints, "before_correctness"
            )
            if identity_before_correctness["fixture"] != runtime["fixture"]:
                raise RunFailure("fixture identity changed before correctness preflight")
        correctness = _run_correctness(
            manifest, endpoints, runtime["fixture"], out_dir, tmp_dir
        )
        if args.correctness_only:
            identity_after_correctness = _revalidate_run_identity(
                manifest, endpoints, "after_correctness"
            )
            if identity_after_correctness["fixture"] != runtime["fixture"]:
                raise RunFailure("fixture identity changed during correctness preflight")
            report = {
                "schema": REPORT_SCHEMA,
                "status": "correctness_only",
                "claim_scope": "untimed correctness preflight only; no calibration or native timing claim",
                "manifest": str(manifest.path),
                "fixture": runtime["fixture"],
                "baseline": _endpoint_metadata(endpoints["baseline"]),
                "candidate": _endpoint_metadata(endpoints["candidate"]),
                "protocol": {
                    "cpu": manifest.protocol.cpu,
                    "max_load": manifest.protocol.max_load,
                    "max_iterations": manifest.protocol.max_iterations,
                    "timeout_s": manifest.protocol.timeout_s,
                },
                "environment": {
                    "platform": platform.platform(),
                    "python": platform.python_version(),
                    "logical_cpus": os.cpu_count(),
                    "runner_affinity": sorted(os.sched_getaffinity(0)),
                },
                "correctness": correctness,
                "identity": {
                    "before_correctness": identity_before_correctness,
                    "after_correctness": identity_after_correctness,
                },
                "started_unix_ns": time.time_ns(),
            }
            _write_json(out_dir / "report.json", report)
            return 0
        timing = _run_timing(
            manifest,
            endpoints,
            runtime["fixture"],
            correctness,
            out_dir,
            tmp_dir,
        )
        report = {
            "schema": REPORT_SCHEMA,
            "status": "complete",
            "claim_scope": "native elapsed operation batches only; correctness and source/output identity are preconditions",
            "manifest": str(manifest.path),
            "fixture": runtime["fixture"],
            "baseline": _endpoint_metadata(endpoints["baseline"]),
            "candidate": _endpoint_metadata(endpoints["candidate"]),
            "protocol": {
                "cpu": manifest.protocol.cpu,
                "blocks": manifest.protocol.blocks,
                "warmups": manifest.protocol.warmups,
                "target_ns": manifest.protocol.target_ns,
                "minimum_ns": manifest.protocol.minimum_ns,
                "max_load": manifest.protocol.max_load,
                "max_iterations": manifest.protocol.max_iterations,
                "timeout_s": manifest.protocol.timeout_s,
                "abba_order": list(ABBA_ORDER),
            },
            "environment": {
                "platform": platform.platform(),
                "python": platform.python_version(),
                "logical_cpus": os.cpu_count(),
                "runner_affinity": sorted(os.sched_getaffinity(0)),
            },
            "correctness": correctness,
            "calibration": timing["calibration"],
            "timing_proofs": timing["proofs"],
            "routes": timing["routes"],
            "identity": timing["identity"],
            "started_unix_ns": time.time_ns(),
        }
        _write_json(out_dir / "report.json", report)
        return 0
    except RunFailure as error:
        return _fail(args, error)
    except Exception as error:
        # Filesystem, hashing, JSON, subprocess, and directory failures must
        # remain structured failed receipts rather than leaking tracebacks.
        return _fail(args, error)


if __name__ == "__main__":
    raise SystemExit(main())
