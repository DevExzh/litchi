#!/usr/bin/env python3
"""Replay small negative probes for the 0413 profile evidence checker.

Each probe copies one evidence input into a temporary directory, makes one
targeted mutation, and requires the checker to reject it.  The capture itself
is never modified.  This is intentionally a thin replay wrapper around the
frozen profile analyzer rather than a second validation framework.
"""

from __future__ import annotations

import argparse
import json
import shutil
import subprocess
import tempfile
from pathlib import Path


def materialize(path: Path, output: Path) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    if path.suffix == ".zst":
        with output.open("wb") as sink:
            result = subprocess.run(
                ["zstd", "-q", "-dc", str(path)],
                stdout=sink,
                stderr=subprocess.PIPE,
                check=False,
            )
        if result.returncode != 0:
            raise RuntimeError(f"cannot decompress {path}: {result.stderr.decode(errors='replace').strip()}")
    else:
        shutil.copyfile(path, output)


def first_existing(capture: Path, role: str, suffix: str) -> Path:
    for name in (f"{role}-profile.{suffix}", f"{role}-profile.{suffix}.zst"):
        path = capture / name
        if path.is_file():
            return path
    raise FileNotFoundError(capture / f"{role}-profile.{suffix}")


def run_checker(args: argparse.Namespace, extra: list[str], output: Path) -> subprocess.CompletedProcess[str]:
    command = [
        "python3",
        str(args.verifier),
        "--capture",
        str(args.capture),
        "--parser",
        str(args.parser),
        "--repo",
        str(args.repo),
        "--output",
        str(output),
    ]
    if args.control_data is not None:
        command.extend(["--control-data", str(args.control_data)])
    if args.candidate_data is not None:
        command.extend(["--candidate-data", str(args.candidate_data)])
    command.extend(extra)
    return subprocess.run(command, text=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, check=False)


def require_rejection(name: str, result: subprocess.CompletedProcess[str], marker: str) -> dict[str, object]:
    text = result.stdout + result.stderr
    if result.returncode != 2:
        raise RuntimeError(f"{name} expected rejection status 2, got {result.returncode}")
    if marker not in text:
        raise RuntimeError(f"{name} rejected for an unexpected reason: {text.strip()}")
    return {"name": name, "status": "rejected_as_expected", "marker": marker, "returncode": result.returncode}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--capture", type=Path, required=True)
    parser.add_argument("--verifier", type=Path, required=True)
    parser.add_argument("--parser", type=Path, required=True)
    parser.add_argument("--repo", type=Path, required=True)
    parser.add_argument("--control-data", type=Path)
    parser.add_argument("--candidate-data", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    args.capture = args.capture.resolve()
    args.verifier = args.verifier.resolve()
    args.parser = args.parser.resolve()
    args.repo = args.repo.resolve()
    with tempfile.TemporaryDirectory(prefix="litchi-goal-0413-probes-") as directory:
        temporary = Path(directory)
        control_source = first_existing(args.capture, "control", "script.txt")
        control_script = temporary / "control-profile.script.txt"
        materialize(control_source, control_script)
        original = control_script.read_text(encoding="utf-8")
        frame = next((line for line in original.splitlines(keepends=True) if line.startswith("\t")), None)
        if frame is None:
            raise RuntimeError("control profile has no frame to mutate")
        broken = original.replace(frame, "BROKEN PROFILE FRAME\n", 1)
        broken_script = temporary / "broken-frame" / "control-profile.script.txt"
        broken_script.parent.mkdir(parents=True, exist_ok=True)
        broken_script.write_text(broken, encoding="utf-8")
        result = run_checker(
            args,
            ["--control-script", str(broken_script)],
            temporary / "broken-frame.json",
        )
        checks = [require_rejection("unparsed_frame", result, "unparsed frame lines")]

        lost_script = temporary / "lost-sample" / "control-profile.script.txt"
        lost_script.parent.mkdir(parents=True, exist_ok=True)
        lost_script.write_text("lost 1 samples\n" + original, encoding="utf-8")
        result = run_checker(
            args,
            ["--control-script", str(lost_script)],
            temporary / "lost-sample.json",
        )
        checks.append(require_rejection("lost_sample", result, "lost perf samples"))

        candidate_identity = args.capture.parent / "checks" / "candidate-build-identity.json"
        if not candidate_identity.is_file():
            candidate_identity = Path("/tmp/litchi-goal-0413-candidate-binaries/identity.json")
        identity = json.loads(candidate_identity.read_text(encoding="utf-8"))
        normal = identity["binaries"]["litchi-perf-baseline"]
        normal["sha256"] = "0" * 64
        mutated_identity = temporary / "candidate-build-identity.json"
        mutated_identity.write_text(json.dumps(identity, indent=2) + "\n", encoding="utf-8")
        result = run_checker(
            args,
            ["--candidate-build-identity", str(mutated_identity)],
            temporary / "binary-scope.json",
        )
        checks.append(require_rejection("binary_identity", result, "binary_sha256"))

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps({"status": "pass", "checks": checks}, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "output": str(args.output), "checks": checks}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, json.JSONDecodeError) as error:
        print(f"litchi-goal-0413-profile-probes: FAIL: {error}", flush=True)
        raise SystemExit(2)
