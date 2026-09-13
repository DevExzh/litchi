#!/usr/bin/env python3
"""Capture the change-0558 evidence: ABBA latency legs and whole-child syscall counts.

The driver never edits production source. It runs one fresh child per cell and
retains a receipt, the child's JSON report, stdout and stderr for every child.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import pathlib
import platform
import shutil
import subprocess
import sys
import time

MODES = ("file-source", "owned-readat")
OPERATIONS = ("open", "list", "one-cell")
LEGS = (("A1", "baseline"), ("B1", "candidate"), ("B2", "candidate"), ("A2", "baseline"))
REPEATS = ("R1", "R2")


def sha256(path: pathlib.Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def child_env() -> dict[str, str]:
    env = dict(os.environ)
    env["RAYON_NUM_THREADS"] = "1"
    env["OMP_NUM_THREADS"] = "1"
    env["CARGO_INCREMENTAL"] = "0"
    return env


def run_child(argv: list[str], out_dir: pathlib.Path, name: str) -> dict:
    stdout_path = out_dir / f"{name}.stdout"
    stderr_path = out_dir / f"{name}.stderr"
    started = time.time()
    with stdout_path.open("wb") as out, stderr_path.open("wb") as err:
        completed = subprocess.run(argv, stdout=out, stderr=err, env=child_env(), check=False)
    receipt = {
        "name": name,
        "command": argv,
        "returncode": completed.returncode,
        "started_unix": started,
        "finished_unix": time.time(),
        "stdout": stdout_path.name,
        "stderr": stderr_path.name,
    }
    (out_dir / f"{name}.receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    return receipt


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--baseline-binary", required=True)
    parser.add_argument("--candidate-binary", required=True)
    parser.add_argument("--input", default="test-data/ole/xls/ConditionalFormattingSamples.xls")
    parser.add_argument("--output-dir", required=True)
    parser.add_argument("--cpu", type=int, default=17)
    parser.add_argument("--warmups", type=int, default=20)
    parser.add_argument("--samples", type=int, default=100)
    parser.add_argument("--stage", choices=("latency", "syscalls", "all"), default="all")
    args = parser.parse_args()

    out_root = pathlib.Path(args.output_dir)
    out_root.mkdir(parents=True, exist_ok=True)
    binaries = {
        "baseline": pathlib.Path(args.baseline_binary).resolve(),
        "candidate": pathlib.Path(args.candidate_binary).resolve(),
    }
    identity = {
        stage: {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}
        for stage, path in binaries.items()
    }
    if identity["baseline"]["sha256"] == identity["candidate"]["sha256"]:
        print("refusing to run: baseline and candidate binaries are identical", file=sys.stderr)
        return 2
    (out_root / "binary-identity.json").write_text(json.dumps(identity, indent=2) + "\n")

    source = pathlib.Path(args.input)
    (out_root / "input-identity.json").write_text(
        json.dumps(
            {"path": str(source), "bytes": source.stat().st_size, "sha256": sha256(source)},
            indent=2,
        )
        + "\n"
    )

    taskset = shutil.which("taskset")
    strace = shutil.which("strace")
    prefix = [taskset, "-c", str(args.cpu)] if taskset else []

    if args.stage in ("latency", "all"):
        lane = out_root / "latency"
        lane.mkdir(exist_ok=True)
        for repeat in REPEATS:
            for leg, stage in LEGS:
                for mode in MODES:
                    for operation in OPERATIONS:
                        name = f"{repeat}-{leg}-{stage}-{mode}-{operation}"
                        report = lane / f"{name}.json"
                        argv = prefix + [
                            str(binaries[stage]),
                            "--input", str(source),
                            "--mode", mode,
                            "--operation", operation,
                            "--warmups", str(args.warmups),
                            "--samples", str(args.samples),
                        ]
                        receipt = run_child(argv, lane, name)
                        if receipt["returncode"] != 0:
                            print(f"child {name} failed", file=sys.stderr)
                            return 1
                        shutil.move(str(lane / f"{name}.stdout"), str(report))
                        print(f"captured {name}")

    if args.stage in ("syscalls", "all"):
        if strace is None:
            print("strace unavailable; syscall lane skipped", file=sys.stderr)
        else:
            lane = out_root / "syscalls"
            lane.mkdir(exist_ok=True)
            for stage in ("baseline", "candidate"):
                for mode in MODES:
                    for operation in OPERATIONS:
                        name = f"{stage}-{mode}-{operation}"
                        table = lane / f"{name}.strace.txt"
                        argv = [
                            strace, "-f", "-c",
                            "-e", "trace=statx,fstat,newfstatat,pread64,read,write,openat,close,lseek,fsync,fdatasync",
                            "-o", str(table),
                        ] + ["--"] + prefix + [
                            str(binaries[stage]),
                            "--input", str(source),
                            "--mode", mode,
                            "--operation", operation,
                            "--warmups", "1",
                            "--samples", "5",
                        ]
                        receipt = run_child(argv, lane, name)
                        if receipt["returncode"] != 0:
                            print(f"strace child {name} failed", file=sys.stderr)
                            return 1
                        shutil.move(str(lane / f"{name}.stdout"), str(lane / f"{name}.json"))
                        print(f"traced {name}")

    (out_root / "host.json").write_text(
        json.dumps(
            {
                "platform": platform.platform(),
                "processor": platform.processor(),
                "python": sys.version,
                "cpu_pin": args.cpu,
                "quiescence": "not established; the host runs unrelated concurrent workloads",
            },
            indent=2,
        )
        + "\n"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
