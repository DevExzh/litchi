#!/usr/bin/env python3
"""Small outer boundary for the final DOCX Ink qualification commands.

The committed ``ink-qualify.py``, ``ink-abba.py`` and ``ink-collect.py``
protocols remain responsible for provenance and semantic checks.  This
launcher owns only the process boundary: a systemd cgroup supplies the hard
20,000,000,000-byte memory limit and zero swap, while this parent supplies the
deadline, polls the newly-owned output and TMP trees, and kills the unit on a
cap breach. The output/TMP directories are fresh and are never reused.
"""

from __future__ import annotations

import argparse
from dataclasses import dataclass
import hashlib
import json
import os
from pathlib import Path
import resource
import shutil
import signal
import subprocess
import sys
import time
import re
from typing import Sequence


MEMORY_MAX_BYTES = 20_000_000_000
SWAP_MAX_BYTES = 0
OUTPUT_MAX_BYTES = 1_000_000_000
TMP_MAX_BYTES = 1_000_000_000
PHASE_DEFAULT_DEADLINES = {
    "allocation": 1_800,
    "abba": 7_200,
    "collect": 900,
}
_UNIT_CHARS = frozenset("abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789-_")
_UNIT_PROPERTIES = (
    "ActiveState",
    "SubState",
    "Result",
    "ControlGroup",
    "MemoryCurrent",
    "MemoryPeak",
    "MainPID",
    "ExecMainStatus",
)


class BoundaryError(RuntimeError):
    """A fail-closed outer-boundary or resource violation."""


@dataclass(frozen=True)
class Limits:
    deadline_seconds: float
    output_bytes: int = OUTPUT_MAX_BYTES
    tmp_bytes: int = TMP_MAX_BYTES
    memory_bytes: int = MEMORY_MAX_BYTES
    swap_bytes: int = SWAP_MAX_BYTES

    def __post_init__(self) -> None:
        if self.deadline_seconds <= 0:
            raise BoundaryError("deadline must be positive")
        if self.output_bytes <= 0 or self.tmp_bytes <= 0:
            raise BoundaryError("output and TMP caps must be positive")
        if self.output_bytes > OUTPUT_MAX_BYTES or self.tmp_bytes > TMP_MAX_BYTES:
            raise BoundaryError("output and TMP caps may not exceed 1,000,000,000 bytes")
        if self.memory_bytes != MEMORY_MAX_BYTES or self.swap_bytes != SWAP_MAX_BYTES:
            raise BoundaryError("Ink boundary memory policy is fixed at 20,000,000,000 bytes and zero swap")


def _sha256(path: Path) -> tuple[str, int]:
    digest = hashlib.sha256()
    size = 0
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
            size += len(chunk)
    return digest.hexdigest(), size


def tree_bytes(path: Path) -> int:
    """Count regular-file bytes without following symlinks."""

    if not path.exists():
        return 0
    if path.is_symlink():
        raise BoundaryError(f"resource tree may not be a symlink: {path}")
    if path.is_file():
        return path.stat().st_size
    total = 0
    for root, dirs, files in os.walk(path, followlinks=False):
        root_path = Path(root)
        kept_dirs = []
        for name in dirs:
            child = root_path / name
            if child.is_symlink():
                raise BoundaryError(f"resource tree contains a symlink: {child}")
            kept_dirs.append(name)
        dirs[:] = kept_dirs
        for name in files:
            child = root_path / name
            if child.is_symlink():
                raise BoundaryError(f"resource tree contains a symlink: {child}")
            total += child.stat().st_size
    return total


def _fresh_directory(path: Path, label: str, *, create: bool) -> None:
    if path.exists():
        if path.is_symlink() or not path.is_dir():
            raise BoundaryError(f"{label} must be a fresh directory: {path}")
        if any(path.iterdir()):
            raise BoundaryError(f"{label} is not empty: {path}")
    elif create:
        path.mkdir(parents=True)


def _disjoint(left: Path, right: Path) -> bool:
    try:
        left.relative_to(right)
        return False
    except ValueError:
        pass
    try:
        right.relative_to(left)
        return False
    except ValueError:
        return True


def _unit_name(value: str) -> str:
    if not value or len(value) > 180 or any(char not in _UNIT_CHARS for char in value):
        raise BoundaryError(f"invalid transient unit name: {value!r}")
    return value


def _scope_name(value: str) -> str:
    base = value.removesuffix(".scope")
    return f"{_unit_name(base)}.scope"


class SafeInkCommand:
    """Run one existing Ink command inside an owned, capped unit."""

    def __init__(
        self,
        *,
        phase: str,
        output: Path,
        tmp: Path,
        log_dir: Path,
        unit: str,
        limits: Limits,
        systemd_run: str = "/usr/bin/systemd-run",
        systemctl: str = "/usr/bin/systemctl",
    ) -> None:
        if phase not in PHASE_DEFAULT_DEADLINES:
            raise BoundaryError(f"unknown Ink phase: {phase}")
        self.phase = phase
        self.output = output.resolve()
        self.tmp = tmp.resolve()
        self.log_dir = log_dir.resolve()
        self.unit = _scope_name(unit)
        self.limits = limits
        self.systemd_run = systemd_run
        self.systemctl = systemctl
        if not _disjoint(self.output, self.tmp):
            raise BoundaryError("output and TMP trees must be disjoint")
        if not _disjoint(self.output, self.log_dir) or not _disjoint(self.tmp, self.log_dir):
            raise BoundaryError("output, TMP and logs must be pairwise disjoint")
        _fresh_directory(self.output, "output", create=False)
        _fresh_directory(self.tmp, "TMP", create=True)
        _fresh_directory(self.log_dir, "log", create=True)

    def systemd_argv(self, command: Sequence[str]) -> list[str]:
        if not command or any(not isinstance(word, str) or not word for word in command):
            raise BoundaryError("inner command must be a non-empty argv")
        child = [
            sys.executable,
            "-B",
            str(Path(__file__).resolve()),
            "--ink-child",
            f"--fsize={self.limits.output_bytes}",
            f"--tmp={self.tmp}",
            "--",
            *command,
        ]
        return [
            self.systemd_run,
            "--user",
            "--scope",
            f"--unit={self.unit}",
            f"--setenv=TMPDIR={self.tmp}",
            "--property=MemoryAccounting=yes",
            f"--property=MemoryMax={self.limits.memory_bytes}",
            f"--property=MemorySwapMax={self.limits.swap_bytes}",
            f"--property=RuntimeMaxSec={self.limits.deadline_seconds:g}",
            "--property=TasksMax=4096",
            "--property=KillMode=control-group",
            "--",
            *child,
        ]

    def _stop_owned(self, process: subprocess.Popen[bytes]) -> dict[str, object]:
        try:
            stop = subprocess.run(
                [self.systemctl, "--user", "kill", "--kill-who=all", self.unit],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.PIPE,
                check=False,
                timeout=10,
            )
            systemctl_exit = stop.returncode
            systemctl_stderr = stop.stderr.decode("utf-8", "replace")
        except subprocess.TimeoutExpired as error:
            systemctl_exit = None
            systemctl_stderr = f"systemctl stop timed out: {error}"
        group_signal = None
        # The unit kill is asynchronous.  A fast child may have completed
        # successfully while systemctl was issuing the kill request; do not
        # turn that completed result into a synthetic SIGTERM just because
        # the outer poll observed the cap first.
        if process.poll() is None:
            try:
                os.killpg(process.pid, signal.SIGTERM)
                group_signal = "SIGTERM"
            except ProcessLookupError:
                pass
        try:
            if process.poll() is None:
                process.wait(timeout=2)
        except subprocess.TimeoutExpired:
            try:
                os.killpg(process.pid, signal.SIGKILL)
                group_signal = "SIGKILL"
            except ProcessLookupError:
                pass
            process.wait(timeout=5)
        return {
            "systemctl_action": "kill",
            "systemctl_exit": systemctl_exit,
            "systemctl_stderr": systemctl_stderr,
            "group_signal": group_signal,
            "returncode": process.returncode,
        }

    def _post_unit_state(self) -> dict[str, object]:
        command = [
            self.systemctl,
            "--user",
            "show",
            "--no-pager",
            f"--property={','.join(_UNIT_PROPERTIES)}",
            self.unit,
        ]
        try:
            completed = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                timeout=5,
            )
        except (OSError, subprocess.TimeoutExpired) as error:
            return {
                "available": False,
                "unit_gone": False,
                "error": str(error),
                "command": command,
            }
        if completed.returncode != 0:
            stderr = completed.stderr[-4096:]
            unit_gone = bool(
                re.search(r"(?:not found|not loaded|no such unit|could not be found)", stderr, re.IGNORECASE)
            )
            return {
                "available": False,
                "unit_gone": unit_gone,
                "returncode": completed.returncode,
                "stderr": stderr,
                "command": command,
            }
        properties = dict(
            line.split("=", 1)
            for line in completed.stdout.splitlines()
            if "=" in line
        )
        if not properties:
            return {
                "available": False,
                "unit_gone": False,
                "command": command,
            }
        real_gc_shape = (
            properties.get("ActiveState") == "inactive"
            and properties.get("SubState") == "dead"
            and properties.get("Result") == "success"
            and properties.get("ControlGroup") == ""
            and properties.get("MemoryCurrent") == "[not set]"
            and properties.get("MemoryPeak") == "[not set]"
            and properties.get("MainPID", "[not set]") in {"0", "[not set]"}
            and properties.get("ExecMainStatus", "[not set]") in {"0", "[not set]"}
        )
        return {
            "available": not real_gc_shape,
            "unit_gone": real_gc_shape,
            "properties": properties,
            "command": command,
        }

    def _cleanup_tmp(self) -> bool:
        if not self.tmp.exists():
            return True
        if self.tmp.is_symlink() or not self.tmp.is_dir():
            raise BoundaryError(f"owned TMP path changed type: {self.tmp}")
        shutil.rmtree(self.tmp)
        return not self.tmp.exists()

    def run(self, command: Sequence[str]) -> dict[str, object]:
        argv = self.systemd_argv(command)
        stdout_path = self.log_dir / "outer.stdout"
        stderr_path = self.log_dir / "outer.stderr"
        started = time.time_ns()
        start_mono = time.monotonic()
        process: subprocess.Popen[bytes] | None = None
        violation: str | None = None
        cleanup: dict[str, object] | None = None
        try:
            with stdout_path.open("wb") as stdout, stderr_path.open("wb") as stderr:
                process = subprocess.Popen(
                    argv,
                    stdin=subprocess.DEVNULL,
                    stdout=stdout,
                    stderr=stderr,
                    start_new_session=True,
                )
                while process.poll() is None:
                    output_bytes = tree_bytes(self.output)
                    tmp_bytes = tree_bytes(self.tmp)
                    if output_bytes > self.limits.output_bytes:
                        violation = f"output cap exceeded: {output_bytes} > {self.limits.output_bytes}"
                    elif tmp_bytes > self.limits.tmp_bytes:
                        violation = f"TMP cap exceeded: {tmp_bytes} > {self.limits.tmp_bytes}"
                    elif time.monotonic() - start_mono > self.limits.deadline_seconds:
                        violation = f"deadline exceeded: {self.limits.deadline_seconds:g}s"
                    if violation is not None:
                        cleanup = self._stop_owned(process)
                        break
                    time.sleep(0.05)
                if process.poll() is None:
                    cleanup = self._stop_owned(process)
                returncode = process.returncode
                final_output_bytes = tree_bytes(self.output)
                final_tmp_bytes = tree_bytes(self.tmp)
                if violation is None:
                    if final_output_bytes > self.limits.output_bytes:
                        violation = (
                            f"output cap exceeded: {final_output_bytes} > "
                            f"{self.limits.output_bytes}"
                        )
                    elif final_tmp_bytes > self.limits.tmp_bytes:
                        violation = (
                            f"TMP cap exceeded: {final_tmp_bytes} > "
                            f"{self.limits.tmp_bytes}"
                        )
                if violation is not None and cleanup is None:
                    cleanup = self._stop_owned(process)
        except BaseException:
            if process is not None and process.poll() is None:
                cleanup = self._stop_owned(process)
            raise
        finally:
            tmp_before_cleanup = tree_bytes(self.tmp)
            tmp_cleaned = self._cleanup_tmp()
        ended = time.time_ns()
        post_unit = self._post_unit_state()
        output_bytes = tree_bytes(self.output)
        log_hashes = {}
        for path in (stdout_path, stderr_path):
            digest, size = _sha256(path)
            log_hashes[path.name] = {"path": str(path), "bytes": size, "sha256": digest}
        receipt = {
            "schema": "litchi.docx-durable-ink-outer.v1",
            "phase": self.phase,
            "status": "failed" if violation or returncode else "complete",
            "command": argv,
            "unit": self.unit,
            "output": {"path": str(self.output), "bytes": output_bytes, "cap": self.limits.output_bytes},
            "tmp": {"path": str(self.tmp), "bytes_before_cleanup": tmp_before_cleanup,
                    "cap": self.limits.tmp_bytes, "cleaned": tmp_cleaned},
            "limits": {"memory_bytes": self.limits.memory_bytes, "swap_bytes": self.limits.swap_bytes,
                       "deadline_seconds": self.limits.deadline_seconds},
            "started_ns": started,
            "finished_ns": ended,
            "returncode": returncode,
            "violation": violation,
            "cleanup": cleanup,
            "post_unit": post_unit,
            "logs": log_hashes,
        }
        receipt_path = self.log_dir / "outer-receipt.json"
        receipt_path.write_text(json.dumps(receipt, indent=2) + "\n", encoding="utf-8")
        return receipt


def _child_main(argv: Sequence[str]) -> int:
    """Install byte-valued file/core limits, then exec the owned command."""

    if len(argv) < 3 or not argv[0].startswith("--fsize=") or not argv[1].startswith("--tmp="):
        raise BoundaryError("invalid child launcher arguments")
    try:
        fsize = int(argv[0].split("=", 1)[1])
    except ValueError as error:
        raise BoundaryError("invalid child file-size limit") from error
    tmp = argv[1].split("=", 1)[1]
    if not tmp or argv[2] != "--" or len(argv) == 3:
        raise BoundaryError("child launcher requires -- followed by an argv")
    resource.setrlimit(resource.RLIMIT_FSIZE, (fsize, fsize))
    resource.setrlimit(resource.RLIMIT_CORE, (0, 0))
    os.environ["TMPDIR"] = tmp
    os.execvpe(argv[3], list(argv[3:]), os.environ.copy())
    raise AssertionError("execvpe returned")


def main(argv: Sequence[str] | None = None) -> int:
    raw = list(sys.argv[1:] if argv is None else argv)
    if raw[:1] == ["--ink-child"]:
        return _child_main(raw[1:])
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("phase", choices=tuple(PHASE_DEFAULT_DEADLINES))
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--tmp", type=Path, required=True)
    parser.add_argument("--log-dir", type=Path, required=True)
    parser.add_argument("--unit", required=True)
    parser.add_argument("--deadline", type=float)
    parser.add_argument("--output-cap", type=int, default=OUTPUT_MAX_BYTES)
    parser.add_argument("--tmp-cap", type=int, default=TMP_MAX_BYTES)
    parser.add_argument("--systemd-run", default="/usr/bin/systemd-run")
    parser.add_argument("--systemctl", default="/usr/bin/systemctl")
    help_args = raw[: raw.index("--")] if "--" in raw else raw
    if "--help" in help_args or "-h" in help_args:
        parser.parse_args(help_args)
    try:
        separator = raw.index("--")
    except ValueError:
        parser.error("an inner command is required after --")
    args = parser.parse_args(raw[:separator])
    command = raw[separator + 1 :]
    if not command:
        parser.error("an inner command is required after --")
    deadline = args.deadline or PHASE_DEFAULT_DEADLINES[args.phase]
    runner = SafeInkCommand(
        phase=args.phase,
        output=args.output,
        tmp=args.tmp,
        log_dir=args.log_dir,
        unit=args.unit,
        limits=Limits(deadline_seconds=deadline, output_bytes=args.output_cap, tmp_bytes=args.tmp_cap),
        systemd_run=args.systemd_run,
        systemctl=args.systemctl,
    )
    receipt = runner.run(command)
    print(json.dumps(receipt, sort_keys=True))
    return 0 if receipt["status"] == "complete" else 1


if __name__ == "__main__":
    raise SystemExit(main())
