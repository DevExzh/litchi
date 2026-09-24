#!/usr/bin/env python3
"""Capture focused matrix tests and their exact isolated source inputs."""

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import subprocess
import tarfile
import time


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--label", required=True)
    parser.add_argument("--test", action="append", required=True)
    args = parser.parse_args()
    workspace = args.workspace.resolve(strict=True)
    output = args.output.resolve()
    if output.exists():
        parser.error("output must be a new directory")
    if output == Path("/tmp") or Path("/tmp") in output.parents:
        parser.error("captures must be stored off tmpfs")
    if output == Path("/var/tmp") or Path("/var/tmp") in output.parents:
        parser.error("captures must be stored outside /var/tmp")

    runner = Path(__file__).resolve()
    gate_runner = runner.parent.parent / "ods-formula-array-reference-evaluation/gates/run.py"
    spec = importlib.util.spec_from_file_location("ods_gates", gate_runner)
    gates = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gates)
    before = gates.hashes(workspace)
    output.mkdir(parents=True)
    archive_path = output / "source.tar.gz"
    with tarfile.open(archive_path, "w:gz") as archive:
        for name in before:
            archive.add(workspace / name, arcname="source/" + name)
        archive.add(runner, arcname="capture.py")
        archive.add(gate_runner, arcname="source-hashes.py")
    with tarfile.open(archive_path, "r:gz") as archive:
        for name, expected in before.items():
            actual = hashlib.sha256(archive.extractfile("source/" + name).read()).hexdigest()
            if actual != expected:
                raise RuntimeError("source changed during archive: " + name)

    env = dict(os.environ, TMPDIR="/home/zhuhe/code/litchi-array-tmp",
               CARGO_TARGET_DIR="/home/zhuhe/code/litchi-array-target",
               CARGO_PROFILE_DEV_DEBUG="0", CARGO_PROFILE_TEST_DEBUG="0",
               CARGO_INCREMENTAL="0")
    command = ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods",
               "--all-features", "--no-fail-fast"]
    for target in args.test:
        command.extend(["--test", target])
    start = time.time()
    with (output / "test.log").open("w") as log:
        status = subprocess.run(command, cwd=workspace, env=env,
                                stdout=log, stderr=subprocess.STDOUT).returncode
    after = gates.hashes(workspace)
    receipt = {
        "label": args.label, "workspace": str(workspace),
        "head_context_only": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=workspace, text=True).strip(),
        "command": command, "status": status, "seconds": time.time() - start,
        "source_before": before, "source_after": after,
        "sources_unchanged": before == after,
        "archive_sha256": hashlib.sha256(archive_path.read_bytes()).hexdigest(),
        "runner_sha256": hashlib.sha256(runner.read_bytes()).hexdigest(),
        "log_sha256": hashlib.sha256((output / "test.log").read_bytes()).hexdigest(),
        "rustc": subprocess.check_output(["rustc", "-Vv"], cwd=workspace, text=True),
        "cargo": subprocess.check_output(["cargo", "-V"], cwd=workspace, text=True),
        "environment": {key: env[key] for key in (
            "TMPDIR", "CARGO_TARGET_DIR", "CARGO_PROFILE_DEV_DEBUG",
            "CARGO_PROFILE_TEST_DEBUG", "CARGO_INCREMENTAL")},
    }
    (output / "receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps({"status": status, "sources_unchanged": before == after,
                      "output": str(output)}), flush=True)
    raise SystemExit(status if before == after else 1)


if __name__ == "__main__":
    main()
