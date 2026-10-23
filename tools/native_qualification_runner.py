#!/usr/bin/env python3
"""Fail-closed runner for the W4b XLSB and W5 legacy native ABBA lanes.

The runner deliberately takes all source, target, binary, fixture, and output
paths from a JSON manifest.  It does not know a wave commit and therefore
cannot silently continue using an old source checkout.  A manifest records the
source identity it observes at execution time instead.

There are two phases:

``build``
    Runs the manifest's Cargo commands one at a time, injecting ``--offline``
    and ``--locked`` only when those flags are absent before Cargo's ``--``
    separator.  It records source/lock/probe/fixture and binary hashes.

``abba``
    Requires the build receipt, then runs every program for the exact
    ``before, after, after, before`` order.  A child failure, missing output,
    source mutation, target overrun, or unrelated Litchi builder causes an
    immediate failed receipt.  Failed child logs are retained.

This is orchestration only.  It never deletes a target or an output and it
does not interpret benchmark numbers.  Keep large receipts and binaries in an
external directory; only reduced summaries belong in the repository.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
from pathlib import Path
import re
import resource
import signal
import subprocess
import sys
import time
from typing import Any, Iterable, Mapping, Sequence


SCHEMA = "litchi.native-qualification.v1"
ABBA_ORDER = ("before", "after", "after", "before")
MAX_TARGET_BYTES = 20 * 1024**3
DEFAULT_MEMORY_BYTES = 20 * 1024**3
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
LITCHI_ROOTS = (
    Path("/home/zhuhe/code/litchi"),
    Path("/home/zhuhe/code/litchi-worktrees"),
)
_TOKEN = re.compile(r"\{([a-z_][a-z0-9_]*)\}")


class RunnerError(RuntimeError):
    """A fail-closed manifest, guard, or child-command error."""


def _now() -> str:
    return _dt.datetime.now(_dt.timezone.utc).isoformat(timespec="milliseconds")


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _sha256_file(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while True:
            block = stream.read(1024 * 1024)
            if not block:
                break
            digest.update(block)
            size += len(block)
    return digest.hexdigest(), size


def _reject_3rdparty(path: Path) -> None:
    if "3rdparty" in path.parts:
        raise RunnerError(f"3rdparty paths are forbidden: {path}")


def _resolve_path(value: str | os.PathLike[str], *, base: Path | None = None) -> Path:
    path = Path(value)
    if not path.is_absolute():
        if base is None:
            raise RunnerError(f"path must be absolute: {value!r}")
        path = base / path
    probe = Path(path.anchor)
    for part in path.parts[1:]:
        probe /= part
        if probe.is_symlink():
            raise RunnerError(f"symlink path components are forbidden: {path}")
    path = path.expanduser().resolve(strict=False)
    _reject_3rdparty(path)
    return path


def _under(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
    except ValueError:
        return False
    return True


def _ensure_external(path: Path, source_roots: Iterable[Path], label: str) -> None:
    for root in source_roots:
        if _under(path, root):
            raise RunnerError(f"{label} must be outside source worktrees: {path}")


def _directory_size(path: Path) -> int:
    """Return a conservative size without following symlinks."""

    if not path.exists():
        return 0
    if path.is_symlink():
        raise RunnerError(f"target/output may not be a symlink: {path}")
    if path.is_file():
        return path.stat().st_size
    total = 0
    for root, dirs, files in os.walk(path, followlinks=False):
        dirs[:] = [name for name in dirs if not (Path(root) / name).is_symlink()]
        for name in files:
            item = Path(root) / name
            if item.is_symlink():
                continue
            try:
                total += item.stat().st_size
            except FileNotFoundError:
                # A compiler can replace a file during the walk.  The next
                # check will see it; never turn a missing file into deletion.
                continue
    return total


def _hash_paths(paths: Iterable[Path]) -> dict[str, Any]:
    """Hash files/directories deterministically, including relative names."""

    entries: list[tuple[str, Path]] = []
    for path in paths:
        path = path.resolve(strict=False)
        _reject_3rdparty(path)
        if not path.exists():
            raise RunnerError(f"identity path does not exist: {path}")
        if path.is_symlink():
            raise RunnerError(f"identity path may not be a symlink: {path}")
        if path.is_file():
            entries.append((path.name, path))
            continue
        for root, dirs, files in os.walk(path, followlinks=False):
            dirs[:] = sorted(name for name in dirs if not (Path(root) / name).is_symlink())
            for name in sorted(files):
                item = Path(root) / name
                if item.is_symlink():
                    raise RunnerError(f"identity directory contains symlink: {item}")
                entries.append((str(item.relative_to(path)), item))
    digest = hashlib.sha256()
    records: list[dict[str, Any]] = []
    for relative, item in sorted(entries, key=lambda pair: (str(pair[1].parent), pair[0])):
        file_hash, size = _sha256_file(item)
        digest.update(relative.encode("utf-8"))
        digest.update(b"\0")
        digest.update(str(size).encode("ascii"))
        digest.update(b"\0")
        digest.update(bytes.fromhex(file_hash))
        records.append({"path": str(item), "relative": relative, "bytes": size, "sha256": file_hash})
    return {"sha256": digest.hexdigest(), "files": records}


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


def _source_identity(spec: Mapping[str, Any], shared: Mapping[str, Any]) -> dict[str, Any]:
    root = _resolve_path(str(spec["source_root"]))
    if not root.is_dir():
        raise RunnerError(f"source root is not a directory: {root}")
    top = _resolve_path(_git(root, "rev-parse", "--show-toplevel").strip())
    if top != root:
        raise RunnerError(f"source_root is not the worktree top-level: {root} != {top}")
    head = _git(root, "rev-parse", "HEAD").strip()
    tree = _git(root, "rev-parse", "HEAD^{tree}").strip()
    status = _git(root, "status", "--porcelain=v1", "--untracked-files=all")
    if status and not bool(spec.get("allow_dirty", False)):
        raise RunnerError(f"source worktree is dirty: {root}\n{status}")
    diff = _git(root, "diff", "--no-ext-diff", "--binary")
    lock_values = spec.get("lock_paths", ["Cargo.lock"])
    if not isinstance(lock_values, list) or not lock_values:
        raise RunnerError("lock_paths must be a non-empty list")
    locks = [_resolve_path(str(value), base=root) for value in lock_values]
    for lock in locks:
        if not lock.is_file():
            raise RunnerError(f"lock file does not exist: {lock}")
    probe_values = [*shared.get("probe_paths", []), *spec.get("probe_paths", [])]
    fixture_values = [*shared.get("fixture_paths", []), *spec.get("fixture_paths", [])]
    probes = [_resolve_path(str(value), base=root) for value in probe_values]
    fixtures = [_resolve_path(str(value), base=root) for value in fixture_values]
    identity_paths = [*locks, *probes, *fixtures]
    files = _hash_paths(identity_paths) if identity_paths else {"sha256": _sha256_bytes(b""), "files": []}
    return {
        "source_root": str(root),
        "head": head,
        "tree": tree,
        "source_tree_sha256": _sha256_bytes((tree + "\0" + status + "\0" + diff).encode("utf-8")),
        "clean": not bool(status),
        "status_sha256": _sha256_bytes(status.encode("utf-8")),
        "diff_sha256": _sha256_bytes(diff.encode("utf-8")),
        "lock_paths": [str(path) for path in locks],
        "probe_paths": [str(path) for path in probes],
        "fixture_paths": [str(path) for path in fixtures],
        "identity_sha256": files["sha256"],
        "identity_files": files["files"],
    }


def _expand_environment(value: Any) -> Any:
    if isinstance(value, str):
        expanded = os.path.expandvars(value)
        if "$" in expanded:
            raise RunnerError(f"unexpanded environment variable in manifest value: {value!r}")
        return expanded
    if isinstance(value, list):
        return [_expand_environment(item) for item in value]
    if isinstance(value, dict):
        return {key: _expand_environment(item) for key, item in value.items()}
    return value


def load_manifest(path: Path) -> dict[str, Any]:
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise RunnerError(f"cannot read manifest {path}: {error}") from error
    if not isinstance(data, dict):
        raise RunnerError("manifest root must be an object")
    data = _expand_environment(data)
    if data.get("schema") != SCHEMA:
        raise RunnerError(f"manifest schema must be {SCHEMA!r}")
    if data.get("abba_order") != list(ABBA_ORDER):
        raise RunnerError(f"abba_order must be {list(ABBA_ORDER)!r}")
    if set(data.get("legs", {})) != {"before", "after"}:
        raise RunnerError("manifest legs must contain exactly before and after")
    if not isinstance(data.get("shared", {}), dict):
        raise RunnerError("shared must be an object")
    try:
        target_limit = int(data.get("target_limit_bytes", MAX_TARGET_BYTES))
        memory_limit = int(data.get("memory_limit_bytes", DEFAULT_MEMORY_BYTES))
    except (TypeError, ValueError) as error:
        raise RunnerError("target_limit_bytes and memory_limit_bytes must be integers") from error
    if target_limit <= 0 or target_limit > MAX_TARGET_BYTES:
        raise RunnerError(f"target_limit_bytes must be in (0, {MAX_TARGET_BYTES}]")
    if memory_limit < 1024**3 or memory_limit > DEFAULT_MEMORY_BYTES:
        raise RunnerError(f"memory_limit_bytes must be in [{1024**3}, {DEFAULT_MEMORY_BYTES}]")
    output = _resolve_path(str(data.get("output_dir", "")))
    source_roots: list[Path] = []
    for leg, spec in data["legs"].items():
        if not isinstance(spec, dict):
            raise RunnerError(f"legs.{leg} must be an object")
        source = _resolve_path(str(spec.get("source_root", "")))
        target = _resolve_path(str(spec.get("target_dir", "")))
        source_roots.append(source)
        if target == source or _under(target, source):
            raise RunnerError(f"legs.{leg}.target_dir must be external to source_root")
        if target == output or _under(target, output) or _under(output, target):
            raise RunnerError("target_dir and output_dir must be distinct trees")
        if not isinstance(spec.get("build", []), list) or not spec["build"]:
            raise RunnerError(f"legs.{leg}.build must be a non-empty list")
        for command in spec["build"]:
            _validate_command(command, f"legs.{leg}.build")
        programs = spec.get("programs", [])
        if not isinstance(programs, list) or not programs:
            raise RunnerError(f"legs.{leg}.programs must be a non-empty list")
        ids: set[str] = set()
        for program in programs:
            if not isinstance(program, dict) or not isinstance(program.get("id"), str):
                raise RunnerError(f"legs.{leg}.programs entries need an id")
            if program["id"] in ids:
                raise RunnerError(f"duplicate program id {program['id']!r} in {leg}")
            ids.add(program["id"])
            if not isinstance(program.get("argv"), list):
                raise RunnerError(f"program {leg}/{program['id']} needs argv")
            if "binary" not in program:
                raise RunnerError(f"program {leg}/{program['id']} needs binary")
    before_ids = [program["id"] for program in data["legs"]["before"]["programs"]]
    after_ids = [program["id"] for program in data["legs"]["after"]["programs"]]
    if before_ids != after_ids:
        raise RunnerError("before and after program order must match")
    if source_roots[0] == source_roots[1]:
        raise RunnerError("before and after source_root must be distinct worktrees")
    targets = [_resolve_path(str(data["legs"][leg]["target_dir"])) for leg in ("before", "after")]
    if targets[0] == targets[1]:
        raise RunnerError("before and after target_dir must be distinct")
    _ensure_external(output, source_roots, "output_dir")
    for leg, spec in data["legs"].items():
        _ensure_external(_resolve_path(str(spec["target_dir"])), source_roots, f"legs.{leg}.target_dir")
    shared = data["shared"]
    for key in ("probe_paths", "fixture_paths"):
        if not isinstance(shared.get(key, []), list):
            raise RunnerError(f"shared.{key} must be a list")
    return data


def _validate_command(command: Any, location: str) -> None:
    if not isinstance(command, dict):
        raise RunnerError(f"{location} command must be an object")
    argv = command.get("argv")
    if not isinstance(argv, list) or not argv or not all(isinstance(item, str) for item in argv):
        raise RunnerError(f"{location} command argv must be a non-empty string list")
    executable = Path(argv[0]).name
    if not isinstance(command.get("id"), str) or not command["id"]:
        raise RunnerError(f"{location} command needs a non-empty id")
    if executable != "cargo":
        raise RunnerError(f"{location} command must execute Cargo directly, not {argv[0]!r}")
    if "shell" in command or command.get("shell"):
        raise RunnerError(f"{location} may not use a shell")


def _render(value: str, context: Mapping[str, Any]) -> str:
    def replace(match: re.Match[str]) -> str:
        key = match.group(1)
        if key not in context:
            raise RunnerError(f"unknown manifest template {{{key}}}")
        return str(context[key])

    return _TOKEN.sub(replace, value)


def _render_argv(values: Sequence[str], context: Mapping[str, Any]) -> list[str]:
    return [_render(value, context) for value in values]


def normalize_cargo_argv(argv: Sequence[str]) -> list[str]:
    """Add Cargo flags before ``--`` without duplicating existing flags."""

    values = list(argv)
    if not values or Path(values[0]).name != "cargo":
        raise RunnerError(f"expected a Cargo executable, got {values!r}")
    separator = values.index("--") if "--" in values else len(values)
    before = values[:separator]
    for flag in ("--offline", "--locked"):
        if flag not in before and not any(item.startswith(flag + "=") for item in before):
            before.append(flag)
    return [*before, *values[separator:]]


def _read_proc(pid: int) -> dict[str, Any] | None:
    proc = Path("/proc") / str(pid)
    try:
        comm = (proc / "comm").read_text(encoding="utf-8", errors="replace").strip()
        cmdline = (proc / "cmdline").read_bytes().split(b"\0")
        cmdline = [part.decode("utf-8", "replace") for part in cmdline if part]
        exe = os.readlink(proc / "exe")
        cwd = os.readlink(proc / "cwd")
        stat = (proc / "stat").read_text(encoding="utf-8", errors="replace")
    except (FileNotFoundError, PermissionError, OSError):
        return None
    right = stat.rsplit(") ", 1)
    if len(right) != 2:
        return None
    fields = right[1].split()
    try:
        parent = int(fields[1])
    except (IndexError, ValueError):
        return None
    return {"pid": pid, "ppid": parent, "comm": comm, "exe": exe, "cwd": cwd, "cmdline": cmdline}


def _proc_snapshot() -> dict[int, dict[str, Any]]:
    entries: dict[int, dict[str, Any]] = {}
    try:
        pids = [int(item.name) for item in Path("/proc").iterdir() if item.name.isdigit()]
    except OSError:
        return entries
    for pid in pids:
        entry = _read_proc(pid)
        if entry is not None:
            entries[pid] = entry
    return entries


def _in_litchi_scope(entry: Mapping[str, Any]) -> bool:
    haystack = [str(entry.get("exe", "")), str(entry.get("cwd", "")), *entry.get("cmdline", [])]
    return any(any(value == str(root) or value.startswith(str(root) + "/") for root in LITCHI_ROOTS) for value in haystack)


def _is_heavy(entry: Mapping[str, Any]) -> bool:
    names = {Path(str(entry.get("exe", ""))).name, str(entry.get("comm", ""))}
    if names & HEAVY_NAMES:
        return True
    first = entry.get("cmdline", [""])[0] if entry.get("cmdline") else ""
    return Path(str(first)).name in HEAVY_NAMES


def _descendants(snapshot: Mapping[int, Mapping[str, Any]], root: int) -> set[int]:
    allowed = {root}
    changed = True
    while changed:
        changed = False
        for pid, entry in snapshot.items():
            if pid not in allowed and int(entry.get("ppid", -1)) in allowed:
                allowed.add(pid)
                changed = True
    return allowed


class ProcessGuard:
    """Reject unrelated compiler/profiler activity in any Litchi worktree."""

    def offenders(self, allowed_root: int | None = None) -> list[dict[str, Any]]:
        snapshot = _proc_snapshot()
        allowed = _descendants(snapshot, allowed_root) if allowed_root is not None else set()
        return [
            entry
            for pid, entry in snapshot.items()
            if pid not in allowed and _in_litchi_scope(entry) and _is_heavy(entry)
        ]

    def require_quiet(self) -> None:
        offenders = self.offenders()
        if offenders:
            raise RunnerError(f"unrelated Litchi builder/profiler is running: {offenders!r}")

    @staticmethod
    def loadavg() -> list[float] | None:
        try:
            return [float(value) for value in os.getloadavg()]
        except (AttributeError, OSError):
            return None


def _json_hash(path: Path) -> tuple[str, int]:
    return _sha256_file(path) if path.exists() else (_sha256_bytes(b""), 0)


class SafeRunner:
    def __init__(self, manifest: Mapping[str, Any], manifest_path: Path, phase: str, cpu: int | None):
        self.manifest = manifest
        self.manifest_path = manifest_path.resolve()
        self.phase = phase
        self.cpu = cpu
        self.guard = ProcessGuard()
        self.target_limit = int(manifest.get("target_limit_bytes", MAX_TARGET_BYTES))
        self.memory_limit = int(manifest.get("memory_limit_bytes", DEFAULT_MEMORY_BYTES))
        self.output = _resolve_path(str(manifest["output_dir"]))
        self.output.mkdir(parents=True, exist_ok=True)
        self.command_dir = self.output / "commands"
        self.command_dir.mkdir(parents=True, exist_ok=True)
        self.events_path = self.output / f"events-{phase}.jsonl"
        self.receipt_path = self.output / f"{phase}-receipt.json"
        if self.events_path.exists() or self.receipt_path.exists():
            raise RunnerError(
                f"output already contains {phase} receipt/journal; choose a fresh output_dir: {self.output}"
            )
        self.shared = manifest.get("shared", {})
        self.recorded_events: list[dict[str, Any]] = []

    def _context(self, leg: str, index: int, *, output: Path | None = None) -> dict[str, Any]:
        spec = self.manifest["legs"][leg]
        source = _resolve_path(str(spec["source_root"]))
        target = _resolve_path(str(spec["target_dir"]))
        fixture_values = [*self.shared.get("fixture_paths", []), *spec.get("fixture_paths", [])]
        fixture = ""
        if len(fixture_values) == 1:
            fixture = str(_resolve_path(str(fixture_values[0]), base=source))
        context: dict[str, Any] = {
            "source_root": source,
            "target_dir": target,
            "output_dir": self.output,
            "leg": leg,
            "index": index,
            "rounds": int(self.manifest.get("rounds", 31)),
            "fixture": fixture,
        }
        if output is not None:
            context["output"] = output
        return context

    def _environment(self, spec: Mapping[str, Any], context: Mapping[str, Any]) -> dict[str, str]:
        env = os.environ.copy()
        path_prepend = [_render(str(value), context) for value in self.manifest.get("path_prepend", [])]
        if path_prepend:
            for path in path_prepend:
                resolved = _resolve_path(path)
                if not resolved.is_dir():
                    raise RunnerError(f"path_prepend entry is not a directory: {resolved}")
            env["PATH"] = os.pathsep.join([*path_prepend, env.get("PATH", "")])
        custom = spec.get("environment", self.manifest.get("environment", {}))
        if not isinstance(custom, dict):
            raise RunnerError("environment must be an object")
        for key, value in custom.items():
            env[str(key)] = _render(str(value), context)
        target = str(context["target_dir"])
        if env.get("CARGO_TARGET_DIR", target) != target:
            raise RunnerError("manifest cannot override the explicit CARGO_TARGET_DIR")
        env["CARGO_TARGET_DIR"] = target
        env["CARGO_BUILD_JOBS"] = "1"
        env["CARGO_INCREMENTAL"] = "0"
        env["LITCHI_MEM_MAX"] = f"{(self.memory_limit + 1024**3 - 1) // 1024**3}G"
        return env

    @staticmethod
    def _recorded_environment(env: Mapping[str, str]) -> dict[str, str]:
        keys = (
            "PATH",
            "CARGO_TARGET_DIR",
            "CARGO_BUILD_JOBS",
            "CARGO_INCREMENTAL",
            "LITCHI_MEM_MAX",
            "LITCHI_PROBE_BATCH_ROUNDS",
            "TMPDIR",
        )
        return {key: env[key] for key in keys if key in env}

    def _identity(self, leg: str) -> dict[str, Any]:
        return _source_identity(self.manifest["legs"][leg], self.shared)

    def _check_identity(self, leg: str, expected: Mapping[str, Any]) -> dict[str, Any]:
        actual = self._identity(leg)
        for key in ("head", "tree", "source_tree_sha256", "clean", "status_sha256", "diff_sha256", "identity_sha256"):
            if actual.get(key) != expected.get(key):
                raise RunnerError(f"source identity changed for {leg}: {key}")
        return actual

    def _check_target(self, target: Path) -> int:
        if target.is_symlink():
            raise RunnerError(f"target_dir may not be a symlink: {target}")
        size = _directory_size(target)
        if size > self.target_limit:
            raise RunnerError(f"target exceeds byte limit: {target} ({size} > {self.target_limit})")
        return size

    def _append_event(self, event: Mapping[str, Any]) -> None:
        with self.events_path.open("a", encoding="utf-8") as stream:
            stream.write(json.dumps(event, sort_keys=True) + "\n")
            stream.flush()
            os.fsync(stream.fileno())

    @staticmethod
    def _terminate(process: subprocess.Popen[Any]) -> None:
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

    def _run_command(
        self,
        *,
        event_id: str,
        leg: str,
        argv: Sequence[str],
        cwd: Path,
        env: Mapping[str, str],
        identity: Mapping[str, Any],
        expected_outputs: Sequence[Path] = (),
        require_stdout: bool = False,
    ) -> dict[str, Any]:
        self.guard.require_quiet()
        target = _resolve_path(str(self.manifest["legs"][leg]["target_dir"]))
        target_before = self._check_target(target)
        stdout_path = self.command_dir / f"{event_id}.stdout"
        stderr_path = self.command_dir / f"{event_id}.stderr"
        if stdout_path.exists() or stderr_path.exists():
            raise RunnerError(f"refusing to overwrite existing command logs for {event_id}")
        start = time.monotonic_ns()
        started_at = _now()
        violation: list[dict[str, Any]] = []
        returncode: int | None = None
        process: subprocess.Popen[Any] | None = None
        runner_error: str | None = None
        last_target_check = 0.0
        with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
            try:
                process = subprocess.Popen(
                    list(argv),
                    cwd=str(cwd),
                    env=dict(env),
                    stdout=stdout,
                    stderr=stderr,
                    stdin=subprocess.DEVNULL,
                    start_new_session=True,
                    preexec_fn=self._preexec(),
                )
                while True:
                    returncode = process.poll()
                    if returncode is not None:
                        break
                    try:
                        offenders = self.guard.offenders(process.pid)
                        if offenders:
                            violation.extend(offenders)
                            runner_error = "process guard rejected an unrelated Litchi process"
                            self._terminate(process)
                            returncode = process.returncode
                            break
                        now = time.monotonic()
                        if now - last_target_check >= 1.0:
                            self._check_target(target)
                            last_target_check = now
                    except RunnerError as error:
                        runner_error = str(error)
                        self._terminate(process)
                        returncode = process.returncode
                        break
                    time.sleep(0.1)
            except BaseException as error:
                if process is not None and process.poll() is None:
                    self._terminate(process)
                runner_error = str(error)
                returncode = process.returncode if process is not None else None
        ended_at = _now()
        try:
            target_after = _directory_size(target)
        except RunnerError as error:
            target_after = -1
            runner_error = runner_error or str(error)
        if target_after > self.target_limit:
            runner_error = runner_error or (
                f"target exceeds byte limit: {target} ({target_after} > {self.target_limit})"
            )
        stdout_hash, stdout_bytes = _json_hash(stdout_path)
        stderr_hash, stderr_bytes = _json_hash(stderr_path)
        output_records: list[dict[str, Any]] = []
        for output in expected_outputs:
            if not output.exists() or not output.is_file():
                output_records.append({"path": str(output), "exists": False})
                continue
            output_hash, output_bytes = _sha256_file(output)
            output_records.append(
                {"path": str(output), "exists": True, "bytes": output_bytes, "sha256": output_hash}
            )
        try:
            identity_after = self._identity(leg)
        except RunnerError as error:
            identity_after = None
            runner_error = runner_error or str(error)
        if identity_after is not None:
            for key in ("head", "tree", "source_tree_sha256", "clean", "status_sha256", "diff_sha256", "identity_sha256"):
                if identity_after.get(key) != identity.get(key):
                    runner_error = runner_error or f"source identity changed for {leg}: {key}"
        event = {
            "schema": SCHEMA,
            "event_id": event_id,
            "leg": leg,
            "argv": list(argv),
            "cwd": str(cwd),
            "environment": self._recorded_environment(env),
            "start_utc": started_at,
            "end_utc": ended_at,
            "duration_ns": time.monotonic_ns() - start,
            "returncode": returncode,
            "stdout": {"path": str(stdout_path), "bytes": stdout_bytes, "sha256": stdout_hash},
            "stderr": {"path": str(stderr_path), "bytes": stderr_bytes, "sha256": stderr_hash},
            "outputs": output_records,
            "target": {"path": str(target), "bytes_before": target_before, "bytes_after": target_after},
            "loadavg": self.guard.loadavg(),
            "source_identity_sha256": identity["identity_sha256"],
            "source_head": identity["head"],
            "source_identity_after": identity_after,
            "process_guard_violations": violation,
            "runner_error": runner_error,
        }
        self._append_event(event)
        self.recorded_events.append(event)
        if runner_error is not None:
            raise RunnerError(f"{runner_error} during {event_id}")
        if returncode != 0:
            raise RunnerError(f"child failed for {event_id} with exit {returncode}; see {stderr_path}")
        if require_stdout and stdout_bytes == 0:
            raise RunnerError(f"child produced no stdout for {event_id}")
        missing = [record["path"] for record in output_records if not record.get("exists")]
        empty = [record["path"] for record in output_records if record.get("exists") and record.get("bytes", 0) == 0]
        if missing:
            raise RunnerError(f"child did not produce declared output(s) for {event_id}: {missing}")
        if empty:
            raise RunnerError(f"child produced empty declared output(s) for {event_id}: {empty}")
        return event

    def _preexec(self):
        cpu = self.cpu
        memory = self.memory_limit

        def configure() -> None:
            if cpu is not None:
                os.sched_setaffinity(0, {cpu})
            resource.setrlimit(resource.RLIMIT_AS, (memory, memory))

        return configure

    def _write_receipt(self, status: str, *, identities: Mapping[str, Any], events: Sequence[Mapping[str, Any]], error: str | None = None) -> None:
        receipt = {
            "schema": SCHEMA,
            "status": status,
            "phase": self.phase,
            "manifest": str(self.manifest_path),
            "manifest_sha256": _sha256_file(self.manifest_path)[0],
            "runner": {
                "path": str(Path(__file__).resolve()),
                "sha256": _sha256_file(Path(__file__).resolve())[0],
            },
            "created_utc": _now(),
            "target_limit_bytes": self.target_limit,
            "memory_limit_bytes": self.memory_limit,
            "cpu": self.cpu,
            "abba_order": list(ABBA_ORDER),
            "identities": identities,
            "event_count": len(self.recorded_events),
            "events_sha256": _sha256_file(self.events_path)[0] if self.events_path.exists() else _sha256_bytes(b""),
        }
        if error is not None:
            receipt["error"] = error
        self.receipt_path.write_text(json.dumps(receipt, indent=2, sort_keys=True) + "\n", encoding="utf-8")

    def build(self) -> None:
        identities: dict[str, Any] = {}
        try:
            for leg in ("before", "after"):
                identity = self._identity(leg)
                identities[leg] = identity
                source = Path(identity["source_root"])
                context = self._context(leg, 0)
                env = self._environment(self.manifest["legs"][leg], context)
                for command in self.manifest["legs"][leg]["build"]:
                    argv = normalize_cargo_argv(_render_argv(command["argv"], context))
                    cwd = _resolve_path(_render(str(command.get("cwd", ".")), context), base=source)
                    if not _under(cwd, source):
                        raise RunnerError(f"build cwd escapes source root: {cwd}")
                    outputs = [
                        _resolve_path(_render(str(value), context))
                        for value in command.get("artifacts", [])
                    ]
                    self._run_command(
                        event_id=f"build-{leg}-{command['id']}",
                        leg=leg,
                        argv=argv,
                        cwd=cwd,
                        env=env,
                        identity=identity,
                        expected_outputs=outputs,
                    )
                artifacts = []
                for command in self.manifest["legs"][leg]["build"]:
                    for value in command.get("artifacts", []):
                        path = _resolve_path(_render(str(value), context))
                        if not path.is_file() or not os.access(path, os.X_OK):
                            raise RunnerError(f"declared binary is missing or not executable: {path}")
                        digest, size = _sha256_file(path)
                        artifacts.append({"path": str(path), "bytes": size, "sha256": digest})
                identities[leg]["artifacts"] = artifacts
            self._write_receipt("passed", identities=identities, events=self.recorded_events)
        except (BaseException,) as error:
            self._write_receipt("failed", identities=identities, events=self.recorded_events, error=str(error))
            raise

    def _load_build_receipt(self) -> dict[str, Any]:
        path = self.output / "build-receipt.json"
        if not path.is_file():
            raise RunnerError(f"ABBA requires a successful build receipt: {path}")
        try:
            receipt = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as error:
            raise RunnerError(f"cannot read build receipt: {error}") from error
        if receipt.get("status") != "passed" or receipt.get("schema") != SCHEMA:
            raise RunnerError("build receipt is not a passed native-qualification receipt")
        if receipt.get("manifest_sha256") != _sha256_file(self.manifest_path)[0]:
            raise RunnerError("build receipt belongs to a different manifest")
        return receipt

    def abba(self) -> None:
        identities: dict[str, Any] = {}
        try:
            build_receipt = self._load_build_receipt()
            for leg in ("before", "after"):
                identity = self._identity(leg)
                recorded = build_receipt.get("identities", {}).get(leg, {})
                if (
                    identity.get("head") != recorded.get("head")
                    or identity.get("tree") != recorded.get("tree")
                    or identity.get("source_tree_sha256") != recorded.get("source_tree_sha256")
                    or identity.get("identity_sha256") != recorded.get("identity_sha256")
                ):
                    raise RunnerError(f"{leg} source identity differs from the build receipt")
                recorded_artifacts = recorded.get("artifacts", [])
                if not isinstance(recorded_artifacts, list) or not recorded_artifacts:
                    raise RunnerError(f"build receipt has no artifacts for {leg}")
                recorded_paths: set[str] = set()
                for artifact in recorded_artifacts:
                    path = _resolve_path(str(artifact.get("path", "")))
                    if not path.is_file() or not os.access(path, os.X_OK):
                        raise RunnerError(f"built artifact is missing or not executable: {path}")
                    digest, size = _sha256_file(path)
                    if digest != artifact.get("sha256") or size != artifact.get("bytes"):
                        raise RunnerError(f"built artifact changed since build receipt: {path}")
                    recorded_paths.add(str(path))
                for program in self.manifest["legs"][leg]["programs"]:
                    context = self._context(leg, 0)
                    program_path = _resolve_path(_render(str(program["binary"]), context))
                    if str(program_path) not in recorded_paths:
                        raise RunnerError(f"program binary was not recorded by the build phase: {program_path}")
                identities[leg] = identity
            for index, leg in enumerate(ABBA_ORDER, start=1):
                spec = self.manifest["legs"][leg]
                context = self._context(leg, index)
                env = self._environment(spec, context)
                for program in spec["programs"]:
                    program_id = program["id"]
                    binary = _resolve_path(_render(str(program["binary"]), context))
                    if not binary.is_file() or not os.access(binary, os.X_OK):
                        raise RunnerError(f"measurement binary is missing or not executable: {binary}")
                    output_values = [
                        _resolve_path(_render(str(value), context))
                        for value in program.get("outputs", [])
                    ]
                    for output in output_values:
                        if not _under(output, self.output):
                            raise RunnerError(f"declared child output escapes output_dir: {output}")
                        if output.exists():
                            raise RunnerError(f"refusing to overwrite existing child output: {output}")
                        output.parent.mkdir(parents=True, exist_ok=True)
                    program_context = dict(context)
                    if len(output_values) == 1:
                        program_context["output"] = output_values[0]
                    args = _render_argv([str(value) for value in program["argv"]], program_context)
                    argv = [str(binary), *args]
                    run_cwd = _resolve_path(str(spec.get("run_cwd", ".")), base=Path(identity["source_root"]))
                    if not _under(run_cwd, Path(identity["source_root"])):
                        raise RunnerError(f"run_cwd escapes source root: {run_cwd}")
                    self._run_command(
                        event_id=f"abba-{index:02d}-{leg}-{program_id}",
                        leg=leg,
                        argv=argv,
                        cwd=run_cwd,
                        env=env,
                        identity=identities[leg],
                        expected_outputs=output_values,
                        require_stdout=bool(program.get("require_stdout", True)),
                    )
            self._write_receipt("passed", identities=identities, events=self.recorded_events)
        except (BaseException,) as error:
            self._write_receipt("failed", identities=identities, events=self.recorded_events, error=str(error))
            raise


def _parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--phase", choices=("build", "abba"), required=True)
    parser.add_argument("--cpu", type=int, help="CPU affinity for child programs")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = _parse_args(argv or sys.argv[1:])
    try:
        manifest_path = args.manifest.resolve()
        manifest = load_manifest(manifest_path)
        runner = SafeRunner(manifest, manifest_path, args.phase, args.cpu)
        if args.phase == "build":
            runner.build()
        else:
            runner.abba()
    except RunnerError as error:
        print(f"native qualification refused: {error}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        print("native qualification interrupted; inspect retained logs", file=sys.stderr)
        return 130
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
