#!/usr/bin/env python3
"""Serial guarded source swaps, debug trace builds, and exact public parity.

This script is the only trace runner. It owns no native performance capture:
root invokes it after the native binaries are frozen, and all source changes
are restored in a finally block. New candidate files are deleted when the
baseline lane is installed and recreated from their exact snapshots when the
candidate lane returns.
"""

from __future__ import annotations

import json
from pathlib import Path
import shutil
import subprocess
import sys
import time

from custody import BIN, P, ROOT, TARGET, census, sha


BINARY_NAME = "docx-active-offset-oracle-0722"
LANES = {"baseline", "candidate"}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def load_source(lane: str) -> dict[str, str]:
    path = P / f"source-{lane}.json"
    require(path.is_file(), f"missing source manifest: {path}")
    value = json.loads(path.read_text(encoding="utf-8"))
    require(isinstance(value, dict), f"{path}: source manifest is not an object")
    require(all(isinstance(k, str) and isinstance(v, str) for k, v in value.items()),
            f"{path}: source manifest is not a hash map")
    return value


def install(
    lane: str,
    baseline: dict[str, str],
    candidate: dict[str, str],
    current: dict[str, str],
) -> dict[str, str]:
    expected = baseline if lane == "baseline" else candidate
    other = candidate if lane == "baseline" else baseline
    require(current == other or current == expected,
            f"source swap starts from neither frozen lane: {len(current)} files")
    changed = {
        relative
        for relative in set(expected) | set(other)
        if expected.get(relative) != other.get(relative)
    }
    for relative in sorted(changed):
        target = ROOT / relative
        if relative not in expected:
            # This is a candidate-only file, such as document_scan.rs or the
            # candidate-only test module. Delete only after the current census
            # proved its exact hash; arbitrary files are never touched.
            require(relative in current, f"cannot remove uncensused path {relative}")
            require(current[relative] == sha(target), f"source changed before removing {relative}")
            target.unlink()
            continue
        snapshot = P / lane / relative
        require(snapshot.is_file(), f"missing {lane} snapshot for {relative}")
        require(sha(snapshot) == expected[relative],
                f"{lane}/{relative}: snapshot hash differs from manifest")
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(snapshot, target)
    actual = census()
    require(actual == expected, f"{lane} source installation census mismatch")
    return actual


def run_checked(command: list[str], log: Path | None = None) -> subprocess.CompletedProcess[str]:
    if log is None:
        return subprocess.run(command, cwd=ROOT, text=True, check=False)
    with log.open("w", encoding="utf-8") as stream:
        return subprocess.run(
            command,
            cwd=ROOT,
            stdout=stream,
            stderr=subprocess.STDOUT,
            text=True,
            check=False,
        )


def main() -> int:
    require(len(sys.argv) == 2 and sys.argv[1] in LANES,
            "usage: trace-run.py baseline|candidate")
    lane = sys.argv[1]
    baseline = load_source("baseline")
    candidate = load_source("candidate")
    current = census()
    require(current == candidate, "trace lane must start from the frozen candidate source")
    out = P / "trace" / lane
    require(not out.exists(), f"refusing to replace existing trace output: {out}")
    out.mkdir(parents=True)
    installed = current
    scratch = ROOT.parent / "litchi-scratch-0722-trace"
    try:
        if lane == "baseline":
            installed = install("baseline", baseline, candidate, installed)
        apply_command = [sys.executable, str(P / "trace.py"), "apply", lane]
        applied = run_checked(apply_command)
        require(applied.returncode == 0, f"trace apply failed with {applied.returncode}")
        trace_source = census()
        (out / "source.json").write_text(
            json.dumps(trace_source, indent=2, sort_keys=True) + "\n", encoding="utf-8"
        )
        require(scratch.is_dir(), f"trace scratch was not created: {scratch}")
        shutil.copy2(scratch / "manifest.json", out / "patch.json")
        build_command = [
            "cargo",
            "build",
            "--locked",
            "--manifest-path",
            str(P / "oracle" / "Cargo.toml"),
            "--target-dir",
            str(TARGET),
            "-j",
            "2",
        ]
        started = time.monotonic()
        build = run_checked(build_command, out / "build.log")
        build_record: dict[str, object] = {
            "schema": "litchi.docx-trace-build-0722.v1",
            "command": build_command,
            "exit_code": build.returncode,
            "elapsed_seconds": time.monotonic() - started,
            "source_sha256": sha(out / "source.json"),
            "patch_sha256": sha(out / "patch.json"),
            "log_sha256": sha(out / "build.log"),
            "scripts": {
                name: sha(P / name)
                for name in ("trace-run.py", "trace.py", "trace.fragment", "trace-analyze.py")
            },
            "probe": {
                name: sha(P / "oracle" / name)
                for name in ("Cargo.toml", "Cargo.lock", "src/main.rs")
            },
        }
        (out / "build.json").write_text(
            json.dumps(build_record, indent=2, sort_keys=True) + "\n", encoding="utf-8"
        )
        require(build.returncode == 0, f"trace build failed with {build.returncode}")
        binary_source = TARGET / "debug" / BINARY_NAME
        require(binary_source.is_file(), f"trace build did not produce {binary_source}")
        binary = BIN / f"{lane}-trace"
        binary.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(binary_source, binary)
        build_record["binary"] = {
            "path": str(binary),
            "sha256": sha(binary),
            "bytes": binary.stat().st_size,
        }
        (out / "build.json").write_text(
            json.dumps(build_record, indent=2, sort_keys=True) + "\n", encoding="utf-8"
        )
        expected_report = P / "oracle" / lane / "report.json"
        require(expected_report.is_file(), f"missing untraced {lane} oracle report")
        for suffix in ("", "-repeat"):
            report = out / f"report{suffix}.json"
            stdout = out / f"stdout{suffix}"
            stderr = out / f"stderr{suffix}"
            receipt = out / f"capture{suffix}.json"
            command = [
                sys.executable,
                str(P / "trace-analyze.py"),
                "capture",
                "--binary",
                str(binary),
                "--report",
                str(report),
                "--stdout",
                str(stdout),
                "--stderr",
                str(stderr),
                "--receipt",
                str(receipt),
                "--cwd",
                str(ROOT),
            ]
            result = run_checked(command)
            require(result.returncode == 0, f"trace capture failed with {result.returncode}")
            require(census() == trace_source, "source changed during trace capture")
            require(report.read_bytes() == expected_report.read_bytes(),
                    f"{lane} traced public report differs from untraced report")
        require((out / "stderr").read_bytes() == (out / "stderr-repeat").read_bytes(),
                f"{lane} repeated trace stderr differs")
        build_record["captures"] = {
            "report_sha256": sha(out / "report.json"),
            "stderr_sha256": sha(out / "stderr"),
            "repeat_stderr_sha256": sha(out / "stderr-repeat"),
        }
        (out / "build.json").write_text(
            json.dumps(build_record, indent=2, sort_keys=True) + "\n", encoding="utf-8"
        )
    finally:
        if scratch.exists():
            restored = run_checked([sys.executable, str(P / "trace.py"), "restore"])
            require(restored.returncode == 0, f"trace restore failed with {restored.returncode}")
        if lane == "baseline":
            installed = install("candidate", baseline, candidate, census())
    require(census() == candidate, "trace runner did not restore the candidate source")
    print(f"{lane} trace build, repeated capture, public parity, and source restoration PASS")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, json.JSONDecodeError) as error:
        print(f"trace-run.py: {error}", file=sys.stderr)
        raise SystemExit(2)
