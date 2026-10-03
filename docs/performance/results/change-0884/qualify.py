#!/usr/bin/env python3
"""Capture debug-profile report qualification; not a performance baseline."""
import argparse
import datetime
import hashlib
import json
from pathlib import Path
import subprocess

ROOT = Path(__file__).resolve().parents[4]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("--output", type=Path, default=Path(__file__).resolve().parent / "qualification")
OUT = parser.parse_args().output.resolve()
OUT.mkdir(parents=True, exist_ok=False)
for mode, binary in [("native", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")]:
    argv = ["cargo", "run", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--locked", "--offline", "-j", "2", "--bin", binary]
    if mode == "allocator":
        argv += ["--features", "allocator-metrics"]
    argv += ["--", "--case", "xlsx_producer_medium_source_full_cell_scan", "--warmup", "1", "--samples", "3", "--filesystem-cache", "warm", "--json", str(OUT / (mode + ".json")), "--producer-evidence", str(OUT / (mode + "-producer.json"))]
    receipt = {"argv": argv, "cwd": str(ROOT), "started_utc": datetime.datetime.now(datetime.timezone.utc).isoformat(), "purpose": "debug-profile semantic/report qualification only"}
    print(mode + " starting", flush=True)
    log = OUT / (mode + ".log")
    with log.open("xb") as handle:
        code = subprocess.run(argv, cwd=ROOT, stdout=handle, stderr=subprocess.STDOUT).returncode
    receipt.update(exit_code=code, finished_utc=datetime.datetime.now(datetime.timezone.utc).isoformat(), output_sha256=hashlib.sha256(log.read_bytes()).hexdigest())
    with (OUT / (mode + "-command.json")).open("x") as handle:
        json.dump(receipt, handle, indent=2)
        handle.write("\n")
    print(mode + " exit " + str(code), flush=True)
    if code:
        raise SystemExit(code)
