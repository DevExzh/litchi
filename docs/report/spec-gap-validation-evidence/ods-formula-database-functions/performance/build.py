#!/usr/bin/env python3
"""Build the database harness against an exact captured source closure."""
import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import time


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, required=True)
    parser.add_argument("--source-receipt", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--retain-binary", type=Path, required=True)
    args = parser.parse_args()
    workspace = args.workspace.resolve(strict=True)
    output = args.output.resolve()
    retained = args.retain_binary.resolve()
    for destination in (output, retained):
        if destination.exists():
            parser.error("destinations must not already exist")
        if any(base == destination or base in destination.parents
               for base in (Path("/tmp"), Path("/var/tmp"))):
            parser.error("build artifacts must remain off temporary filesystems")
    runner = Path(__file__).resolve()
    root = runner.parents[5]
    gate_runner = root / "docs/report/spec-gap-validation-evidence/ods-formula-array-reference-evaluation/gates/run.py"
    spec = importlib.util.spec_from_file_location("gates", gate_runner)
    gates = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(gates)
    source_receipt = args.source_receipt.resolve(strict=True)
    source = json.loads(source_receipt.read_text())["source_before"]
    if gates.hashes(workspace) != source:
        parser.error("workspace differs from captured source receipt")
    output.mkdir(parents=True)
    harness = runner.parent / "harness"
    relative = harness.relative_to(root)
    files = ("Cargo.toml", "Cargo.lock", "src/main.rs")
    for name in files:
        destination = workspace / relative / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(harness / name, destination)
        frozen = output / "harness" / name
        frozen.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(destination, frozen)
    before = {name: digest(workspace / relative / name) for name in files}
    shutil.copy2(runner, output / "build.py")
    env = dict(os.environ, TMPDIR="/home/zhuhe/code/litchi-array-tmp",
               CARGO_TARGET_DIR="/home/zhuhe/code/litchi-array-target",
               CARGO_INCREMENTAL="0")
    command = ["cargo", "build", "--locked", "--offline", "--release",
               "--manifest-path", str(workspace / relative / "Cargo.toml")]
    start = time.time()
    with (output / "build.log").open("w") as log:
        status = subprocess.run(command, cwd=workspace, env=env,
                                stdout=log, stderr=subprocess.STDOUT).returncode
    unchanged = gates.hashes(workspace) == source and before == {
        name: digest(workspace / relative / name) for name in files}
    binary_hash = None
    if status == 0 and unchanged:
        binary = Path(env["CARGO_TARGET_DIR"]) / "release/ods-formula-database-functions-performance"
        retained.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(binary, retained)
        binary_hash = digest(retained)
        assert binary_hash == digest(binary)
    receipt = {"command": command, "status": status,
               "seconds": time.time() - start, "sources_unchanged": unchanged,
               "source_receipt": str(source_receipt),
               "source_receipt_sha256": digest(source_receipt),
               "harness_sha256": before, "binary": str(retained),
               "binary_sha256": binary_hash, "runner_sha256": digest(runner),
               "log_sha256": digest(output / "build.log"),
               "rustc": subprocess.check_output(["rustc", "-Vv"], text=True),
               "cargo": subprocess.check_output(["cargo", "-V"], text=True),
               "environment": {key: env[key] for key in
                               ("TMPDIR", "CARGO_TARGET_DIR", "CARGO_INCREMENTAL")}}
    (output / "receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps({"status": status, "sources_unchanged": unchanged}))
    raise SystemExit(status if unchanged else 1)


if __name__ == "__main__":
    main()
