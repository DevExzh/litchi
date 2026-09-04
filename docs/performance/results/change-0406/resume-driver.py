#!/usr/bin/env python3
"""Resume the paused 0406 profile without repeating completed workloads.

The first capture driver writes its command log only after all profiler legs
complete.  If it was terminated during a report leg, this wrapper reconstructs
the completed command prefix from the immutable output files and the failure
record, then loads that log into the immutable Capture implementation.  A
successful completed label is reused only when its required artifacts are
still present.  New subprocesses run with DEBUGINFOD_URLS empty so perf report
cannot block on external debuginfod services.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import importlib.util
import json
import os
import re
import shlex
import shutil
import subprocess
import sys
import time
from pathlib import Path
from typing import Any, Sequence


TEMP = Path("/tmp/litchi-goal-profile-output").resolve()
FINAL = Path("/home/zhuhe/code/litchi/docs/performance/results/change-0406").resolve()
IMMUTABLE_DRIVER = TEMP / "capture-driver.py"
RESUME_DRIVER_ARTIFACT = TEMP / "resume-driver.py"
RECOVERY_ARTIFACT = TEMP / "resume-recovery.json"
DEBUGINFOD_ENV = "DEBUGINFOD_URLS"


def load_json(path: Path) -> dict[str, Any]:
    value = json.loads(path.read_text(encoding="utf-8"))
    if not isinstance(value, dict):
        raise RuntimeError(f"{path} must contain a JSON object")
    return value


def dump_json(path: Path, value: Any) -> None:
    path.write_text(
        json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n",
        encoding="utf-8",
    )


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def utc_now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def import_immutable_driver() -> Any:
    if not IMMUTABLE_DRIVER.is_file():
        raise RuntimeError(f"immutable driver is missing: {IMMUTABLE_DRIVER}")
    spec = importlib.util.spec_from_file_location(
        "litchi_goal_profile_immutable", IMMUTABLE_DRIVER
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot import immutable driver: {IMMUTABLE_DRIVER}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def rel(temp: Path, path: Path | None) -> str | None:
    if path is None:
        return None
    try:
        return str(path.resolve().relative_to(temp))
    except ValueError:
        return str(path)


def harness_args(
    binary: Path,
    temp: Path,
    warmup: int,
    samples: int,
    case: str,
    shape: str,
    payload: str,
    json_name: str,
) -> list[str]:
    return [
        str(binary),
        "--warmup",
        str(warmup),
        "--samples",
        str(samples),
        "--case",
        case,
        "--shape",
        shape,
        "--payload",
        payload,
        "--json",
        str(temp / json_name),
    ]


def reconstructed_record(
    *,
    repo: Path,
    temp: Path,
    label: str,
    argv: Sequence[str],
    stdout: str | None = None,
    stderr: str | None = None,
    returncode: int = 0,
    checked: bool = True,
    environment_overrides: dict[str, str] | None = None,
    basis: str,
) -> dict[str, Any]:
    """Describe a completed pre-pause command without fabricating timing."""
    return {
        "label": label,
        "argv": [str(item) for item in argv],
        "command": shlex.join(str(item) for item in argv),
        "cwd": str(repo),
        "started_utc": None,
        "finished_utc": None,
        "wall_ns": None,
        "returncode": returncode,
        "stdout": stdout,
        "stderr": stderr,
        "checked": checked,
        "environment_overrides": (
            {"RUSTUP_TOOLCHAIN": "1.98.1"}
            if environment_overrides is None
            else environment_overrides
        ),
        "reconstructed_from_artifacts": True,
        "timing_unavailable_after_parent_exit": True,
        "reconstruction_basis": basis,
    }


def old_command_prefix(module: Any, environment: dict[str, Any]) -> list[dict[str, Any]]:
    """Reconstruct the 28 commands completed before inclusive report."""
    repo = Path(environment["repo"]).resolve()
    binary = Path(environment["binary"]["path"]).resolve()
    tools = {str(key): str(value) for key, value in environment["tools"].items()}
    temp = TEMP
    protocol = environment["protocol"]
    case = str(protocol["selector"])
    shape = str(protocol["shape"])
    payload = str(protocol["payload"])
    warmup = int(protocol["warmup_iterations"])
    samples = int(protocol["samples"])
    harness = lambda name: harness_args(
        binary, temp, warmup, samples, case, shape, payload, name
    )
    commands: list[dict[str, Any]] = []
    basis = "initial driver output and environment.json; driver exited before commands.json"

    # The original source binding performed these four direct Git calls before
    # the metadata commands.  Their values and hashes remain in environment.json.
    git_prefix = ["git", "-C", str(repo)]
    commands.extend(
        [
            reconstructed_record(
                repo=repo,
                temp=temp,
                label="git-status",
                argv=[*git_prefix, "status", "--porcelain=v1"],
                environment_overrides={},
                basis=basis,
            ),
            reconstructed_record(
                repo=repo,
                temp=temp,
                label="git-status",
                argv=[*git_prefix, "status", "--porcelain=v1", "-z"],
                environment_overrides={},
                basis=basis,
            ),
            reconstructed_record(
                repo=repo,
                temp=temp,
                label="git-rev-parse",
                argv=[*git_prefix, "rev-parse", "--verify", "HEAD"],
                environment_overrides={},
                basis=basis,
            ),
            reconstructed_record(
                repo=repo,
                temp=temp,
                label="git-rev-parse",
                argv=[*git_prefix, "rev-parse", "--verify", "HEAD^{tree}"],
                environment_overrides={},
                basis=basis,
            ),
        ]
    )

    taskset_text = (temp / "taskset-self.txt").read_text(
        encoding="utf-8", errors="replace"
    )
    taskset_match = re.search(r"\bpid\s+(\d+)'s current affinity", taskset_text)
    taskset_pid = taskset_match.group(1) if taskset_match else "unknown"
    metadata = [
        ("uname", ["uname", "-a"], "uname.txt", "uname.stderr.txt"),
        ("lscpu", ["lscpu"], "lscpu.txt", "lscpu.stderr.txt"),
        ("free", ["free", "-h"], "memory-free.txt", "memory-free.stderr.txt"),
        (
            "taskset-self",
            [tools["taskset"], "-pc", taskset_pid],
            "taskset-self.txt",
            "taskset-self.stderr.txt",
        ),
        (
            "readelf-build-id",
            [tools["readelf"], "-n", str(binary)],
            "binary-readelf.txt",
            "binary-readelf.stderr.txt",
        ),
    ]
    for label, argv, stdout, stderr in metadata:
        commands.append(
            reconstructed_record(
                repo=repo,
                temp=temp,
                label=label,
                argv=argv,
                stdout=stdout,
                stderr=stderr,
                basis=basis,
            )
        )

    versions = {
        "rustc": ["rustc", "+1.98.1", "--version", "--verbose"],
        "cargo": ["cargo", "+1.98.1", "--version", "--verbose"],
        "perf": [tools["perf"], "--version"],
        "heaptrack": [tools["heaptrack"], "--version"],
        "heaptrack_print": [tools["heaptrack_print"], "--version"],
        "strace": [tools["strace"], "--version"],
        "fincore": [tools["fincore"], "--version"],
        "taskset": [tools["taskset"], "--version"],
        "time": ["/usr/bin/time", "--version"],
        "gzip": [tools["gzip"], "--version"],
        "zstd": [tools["zstd"], "--version"],
        "python3": [sys.executable, "--version"],
    }
    for name, argv in versions.items():
        commands.append(
            reconstructed_record(
                repo=repo,
                temp=temp,
                label=f"version-{name}",
                argv=argv,
                stdout=f"versions/{name}.stdout.txt",
                stderr=f"versions/{name}.stderr.txt",
                basis=basis,
            )
        )

    commands.append(
        reconstructed_record(
            repo=repo,
            temp=temp,
            label="fincore-before",
            argv=[tools["fincore"], "--json", "--bytes", "--output-all", str(binary)],
            stdout="fincore-before.json",
            stderr="fincore-before.stderr.txt",
            checked=False,
            basis=basis,
        )
    )
    commands.append(
        reconstructed_record(
            repo=repo,
            temp=temp,
            label="baseline",
            argv=[tools["taskset"], "-c", str(environment["cpu_argument"]), *harness("baseline.json")],
            stdout="baseline.stdout.txt",
            stderr="baseline.stderr.txt",
            basis=basis,
        )
    )
    commands.append(
        reconstructed_record(
            repo=repo,
            temp=temp,
            label="time",
            argv=[
                tools["taskset"],
                "-c",
                str(environment["cpu_argument"]),
                "/usr/bin/time",
                "-v",
                "-o",
                str(temp / "time-v.txt"),
                "--",
                *harness("time.json"),
            ],
            stdout="time.stdout.txt",
            stderr="time.stderr.txt",
            basis=basis,
        )
    )
    commands.append(
        reconstructed_record(
            repo=repo,
            temp=temp,
            label="perf-stat",
            argv=[
                tools["taskset"],
                "-c",
                str(environment["cpu_argument"]),
                tools["perf"],
                "stat",
                "-x,",
                "-o",
                str(temp / "perf-stat.csv"),
                "-e",
                ",".join(module.EVENTS),
                "--",
                *harness("perf-stat.json"),
            ],
            stdout="perf-stat.stdout.txt",
            stderr="perf-stat.stderr.txt",
            basis=basis,
        )
    )
    commands.append(
        reconstructed_record(
            repo=repo,
            temp=temp,
            label="perf-record",
            argv=[
                tools["taskset"],
                "-c",
                str(environment["cpu_argument"]),
                tools["perf"],
                "record",
                "-e",
                "cycles:u",
                "-F",
                "99",
                "--call-graph",
                "dwarf",
                "-o",
                str(temp / "perf.data"),
                "--",
                *harness("perf-record.json"),
            ],
            stdout="perf-record.stdout.txt",
            stderr="perf-record.stderr.txt",
            basis=basis,
        )
    )
    commands.append(
        reconstructed_record(
            repo=repo,
            temp=temp,
            label="perf-report-self",
            argv=[
                tools["perf"],
                "report",
                "--stdio",
                "--no-children",
                "--call-graph=none",
                "--percent-limit=0.5",
                "-i",
                str(temp / "perf.data"),
            ],
            stdout="perf-report-self.txt",
            stderr="perf-report-self.stderr.txt",
            basis=basis,
        )
    )
    commands.append(
        reconstructed_record(
            repo=repo,
            temp=temp,
            label="perf-report-inclusive",
            argv=[
                tools["perf"],
                "report",
                "--stdio",
                "--children",
                "--call-graph=none",
                "--percent-limit=0.5",
                "-i",
                str(temp / "perf.data"),
            ],
            stdout="perf-report-inclusive.txt",
            stderr="perf-report-inclusive.stderr.txt",
            returncode=-15,
            basis="failure.json and preserved pre-resume report files",
        )
    )
    if len(commands) != 28:
        raise RuntimeError(f"reconstructed {len(commands)} commands; expected 28")
    return commands


REUSE_REQUIRED: dict[str, tuple[str, ...]] = {
    "fincore-before": ("fincore-before.json", "fincore-before.stderr.txt"),
    "baseline": ("baseline.json", "baseline.stdout.txt", "baseline.stderr.txt"),
    "time": ("time.json", "time-v.txt", "time.stdout.txt", "time.stderr.txt"),
    "perf-stat": (
        "perf-stat.json",
        "perf-stat.csv",
        "perf-stat.stdout.txt",
        "perf-stat.stderr.txt",
    ),
    "perf-record": (
        "perf-record.json",
        "perf.data",
        "perf-record.stdout.txt",
        "perf-record.stderr.txt",
    ),
    "perf-report-self": ("perf-report-self.txt", "perf-report-self.stderr.txt"),
}


class ResumeCapture:
    def __init__(self, module: Any, arguments: argparse.Namespace, prior: list[dict[str, Any]]) -> None:
        self.module = module
        self.capture = module.Capture(arguments)
        self.capture.commands = prior
        self.reused_labels: list[str] = []
        self.started_labels: list[str] = []

    def _artifact_exists(self, name: Any) -> bool:
        if not isinstance(name, str):
            return False
        path = Path(name)
        if not path.is_absolute():
            path = TEMP / path
        return path.is_file()

    def _reuse(self, label: str) -> dict[str, Any] | None:
        required = REUSE_REQUIRED.get(label)
        if required is None:
            return None
        for record in reversed(self.capture.commands):
            if (
                record.get("label") == label
                and record.get("returncode") == 0
                and all(self._artifact_exists(name) for name in required)
            ):
                record["reused_by_resume"] = True
                self.reused_labels.append(label)
                return record
        return None

    def run(self, label: str, argv: Sequence[str], **kwargs: Any) -> dict[str, Any]:
        reused = self._reuse(label)
        if reused is not None:
            return reused
        env = dict(kwargs.pop("env", {}) or {})
        env[DEBUGINFOD_ENV] = ""
        record = self.module.Capture.run(self.capture, label, argv, env=env, **kwargs)
        record.setdefault("environment_overrides", {})[DEBUGINFOD_ENV] = ""
        record["resume_debuginfod_override"] = ""
        self.started_labels.append(label)
        return record

    def git(self, *args: str, **kwargs: Any) -> str:
        result = self.module.Capture.git(self.capture, *args, **kwargs)
        self.capture.commands[-1].setdefault("environment_overrides", {})[DEBUGINFOD_ENV] = ""
        self.capture.commands[-1]["resume_debuginfod_override"] = ""
        return result

    def execute(self) -> dict[str, Any]:
        original_run = self.capture.run
        original_git = self.capture.git
        self.capture.run = self.run  # type: ignore[method-assign]
        self.capture.git = self.git  # type: ignore[method-assign]
        os.environ[DEBUGINFOD_ENV] = ""
        try:
            return self.capture.capture()
        finally:
            self.capture.run = original_run  # type: ignore[method-assign]
            self.capture.git = original_git  # type: ignore[method-assign]


def preserve_failed_report_files(temp: Path) -> dict[str, str]:
    preserved: dict[str, str] = {}
    for name in ("perf-report-inclusive.txt", "perf-report-inclusive.stderr.txt"):
        source = temp / name
        if source.is_file():
            target_name = name.replace(".txt", ".failed-before-resume.txt")
            target = temp / target_name
            shutil.copy2(source, target)
            preserved[name] = target_name
    return preserved


def load_or_reconstruct_commands(module: Any, environment: dict[str, Any]) -> tuple[list[dict[str, Any]], bool]:
    path = TEMP / "commands.json"
    if path.is_file():
        payload = load_json(path)
        commands = payload.get("commands")
        if not isinstance(commands, list):
            raise RuntimeError("commands.json.commands must be a list")
        copied = [dict(command) for command in commands]
        reconstructed = any(
            command.get("reconstructed_from_artifacts") is True for command in copied
        )
        return copied, reconstructed
    commands = old_command_prefix(module, environment)
    dump_json(
        path,
        {
            "schema_version": 1,
            "captured_utc": environment.get("captured_utc"),
            "commands": commands,
            "reconstructed_after_parent_exit": True,
        },
    )
    return commands, True


def main() -> int:
    if not TEMP.is_dir():
        raise RuntimeError(f"paused capture directory is missing: {TEMP}")
    module = import_immutable_driver()
    environment = load_json(TEMP / "environment.json")
    repo = Path(environment["repo"]).resolve()
    binary = Path(environment["binary"]["path"]).resolve()
    protocol = environment["protocol"]
    commands, reconstructed = load_or_reconstruct_commands(module, environment)
    preserved = preserve_failed_report_files(TEMP)
    for command in commands:
        if command.get("label") != "perf-report-inclusive" or command.get("returncode") == 0:
            continue
        old_stdout = command.get("stdout")
        old_stderr = command.get("stderr")
        if isinstance(old_stdout, str) and old_stdout in preserved:
            command["original_stdout"] = old_stdout
            command["stdout"] = preserved[old_stdout]
        if isinstance(old_stderr, str) and old_stderr in preserved:
            command["original_stderr"] = old_stderr
            command["stderr"] = preserved[old_stderr]
        command["failed_report_preserved"] = True
    shutil.copy2(Path(__file__).resolve(), RESUME_DRIVER_ARTIFACT)
    recovery = {
        "schema_version": 1,
        "resumed_utc": utc_now(),
        "immutable_driver": {
            "path": str(IMMUTABLE_DRIVER),
            "bytes": IMMUTABLE_DRIVER.stat().st_size,
            "sha256": sha256_file(IMMUTABLE_DRIVER),
        },
        "resume_driver": {
            "path": RESUME_DRIVER_ARTIFACT.name,
            "bytes": RESUME_DRIVER_ARTIFACT.stat().st_size,
            "sha256": sha256_file(RESUME_DRIVER_ARTIFACT),
        },
        "commands_source": "reconstructed from retained artifacts because the first driver exited before commands.json"
        if reconstructed
        else "retained commands.json",
        "commands_reconstructed": reconstructed,
        "prior_command_count": len(commands),
        "failed_report_preserved": preserved,
        "failure_artifact": "failure.json",
        "debuginfod_override": {DEBUGINFOD_ENV: ""},
        "resume_from": "perf-report-inclusive",
        "completed_workloads_reused": sorted(REUSE_REQUIRED),
        "workload_rerun": False,
        "source_and_binary_binding": {
            "expected_revision": environment["expected_revision"],
            "binary": environment["binary"],
        },
    }
    dump_json(RECOVERY_ARTIFACT, recovery)

    args = argparse.Namespace(
        repo=repo,
        binary=binary,
        revision=environment["expected_revision"],
        temp=TEMP,
        final=FINAL,
        cpu=str(environment["cpu_argument"]),
        warmup=int(protocol["warmup_iterations"]),
        samples=int(protocol["samples"]),
        heaptrack_samples=int(protocol["heaptrack_samples"]),
        skip_strace=False,
    )
    runner = ResumeCapture(module, args, commands)
    capture = runner.capture
    capture.start_time = str(environment.get("captured_utc", utc_now()))
    capture.tools = {str(key): str(value) for key, value in environment["tools"].items()}
    capture.binary_sha256 = str(environment["binary"]["sha256"])
    capture.binary_bytes = int(environment["binary"]["bytes"])
    capture.binary_mode = int(environment["binary"]["mode_bits"])
    capture.source_before = dict(environment["source_before"])
    capture.allowed_untracked = environment["source_before"].get("allowed_untracked")
    capture.driver_artifact = {
        "path": IMMUTABLE_DRIVER.name,
        "bytes": IMMUTABLE_DRIVER.stat().st_size,
        "sha256": sha256_file(IMMUTABLE_DRIVER),
    }

    # Validate the exact source/binary binding before any resumed profiler.
    os.environ[DEBUGINFOD_ENV] = ""
    precheck_command_count = len(capture.commands)
    current = capture.assert_source_binding()
    for command in capture.commands[precheck_command_count:]:
        command.setdefault("environment_overrides", {})[DEBUGINFOD_ENV] = ""
        command["resume_debuginfod_override"] = ""
    if current["head"] != environment["source_before"]["head"]:
        raise RuntimeError("HEAD changed since the paused profile")
    if current["allowed_untracked"] != environment["source_before"].get("allowed_untracked"):
        raise RuntimeError("allowed untracked source changed since the paused profile")
    if sha256_file(binary) != capture.binary_sha256 or binary.stat().st_size != capture.binary_bytes:
        raise RuntimeError("release binary changed since the paused profile")

    # Record the resume metadata in the retained environment before execution;
    # the final manifest will hash this updated provenance file.
    environment["resume"] = {
        "driver": RESUME_DRIVER_ARTIFACT.name,
        "recovery": RECOVERY_ARTIFACT.name,
        "resumed_utc": recovery["resumed_utc"],
        "commands_source": recovery["commands_source"],
        "commands_reconstructed": reconstructed,
        "debuginfod_override": {DEBUGINFOD_ENV: ""},
        "resume_from": "perf-report-inclusive",
        "workload_rerun": False,
    }
    dump_json(TEMP / "environment.json", environment)

    capture_result = runner.execute()
    capture.write_commands()
    commands_payload = load_json(TEMP / "commands.json")
    commands_payload["resume"] = {
        "driver": RESUME_DRIVER_ARTIFACT.name,
        "recovery": RECOVERY_ARTIFACT.name,
        "reused_labels": runner.reused_labels,
        "started_labels": runner.started_labels,
        "failed_report_preserved": preserved,
        "debuginfod_override": {DEBUGINFOD_ENV: ""},
        "workload_rerun": False,
    }
    dump_json(TEMP / "commands.json", commands_payload)

    capture.make_manifest(capture_result, environment)
    manifest_path = TEMP / "artifact-manifest.json"
    manifest = load_json(manifest_path)
    manifest["resume_driver_artifact"] = {
        "path": RESUME_DRIVER_ARTIFACT.name,
        "bytes": RESUME_DRIVER_ARTIFACT.stat().st_size,
        "sha256": sha256_file(RESUME_DRIVER_ARTIFACT),
    }
    manifest["resume_recovery_artifact"] = {
        "path": RECOVERY_ARTIFACT.name,
        "bytes": RECOVERY_ARTIFACT.stat().st_size,
        "sha256": sha256_file(RECOVERY_ARTIFACT),
    }
    manifest["resume"] = commands_payload["resume"]
    manifest["limitations"].append(
        "The first driver's pre-pause command timing was unavailable because it exited before commands.json; those records are explicitly marked reconstructed, while resumed commands retain exact timing."
    )
    dump_json(manifest_path, manifest)
    capture.publish()
    print(f"published {capture.final}")
    print(f"reused labels {runner.reused_labels}")
    print(f"started labels {runner.started_labels}")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as error:
        print(f"litchi-goal-profile-resume: {error}", file=sys.stderr)
        raise
