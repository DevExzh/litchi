#!/usr/bin/env python3
"""Run the full-cell scan harness gates and retain their exact output."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parents[4]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--output", type=Path, default=Path(__file__).resolve().parent / "final")
OUT = parser.parse_args().output.resolve()
OUT.mkdir(parents=True, exist_ok=False)
MANIFEST = "tools/perf-baseline/Cargo.toml"
COMMANDS = [
    ("fmt", ["cargo", "fmt", "--manifest-path", MANIFEST, "--check"]),
    ("check", ["cargo", "check", "--manifest-path", MANIFEST, "--locked", "--offline", "--all-features", "--all-targets", "-j", "2"]),
    ("scan-tests", ["cargo", "test", "--manifest-path", MANIFEST, "--locked", "--offline", "--lib", "producer_full_cell", "-j", "2"]),
    ("clippy", ["cargo", "clippy", "--manifest-path", MANIFEST, "--locked", "--offline", "--lib", "-j", "2", "--", "-D", "warnings"]),
]

def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()

for name, argv in COMMANDS:
    log = OUT / (name + ".log")
    receipt = {"argv": argv, "cwd": str(ROOT), "started_utc": now()}
    print(name + " starting", flush=True)
    with log.open("xb") as handle:
        code = subprocess.run(argv, cwd=ROOT, stdout=handle, stderr=subprocess.STDOUT).returncode
    receipt.update(exit_code=code, finished_utc=now(), output_sha256=hashlib.sha256(log.read_bytes()).hexdigest())
    with (OUT / (name + ".json")).open("x") as handle:
        json.dump(receipt, handle, indent=2)
        handle.write("\n")
    print(name + " exit " + str(code), flush=True)
    if code:
        raise SystemExit(code)
