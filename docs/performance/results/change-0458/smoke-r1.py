#!/usr/bin/env python3
"""Check both instrumentation modes and clocks before freezing capture inputs."""
import importlib.util
import json
from pathlib import Path
import subprocess

import capture

ROOT = Path(__file__).resolve().parent


def main():
    spec = importlib.util.spec_from_file_location("oracle458", ROOT / "oracle.py")
    oracle = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(oracle)
    directory = ROOT / "smoke-r1"
    directory.mkdir()
    for instrumentation, binary in [("normal", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")]:
        for scope in ["lifecycle", "phases"]:
            lane = {"id": f"{instrumentation}-{scope}", "instrumentation": instrumentation, "scope": scope, "shape": "tiny", "repeat": "R1"}
            report = directory / (lane["id"] + ".json")
            argv = [str(capture.REPO / "tools/perf-baseline/target/release" / binary), "odp-append-attribution", "--mode", scope,
                    "--shape", "tiny", "--samples", "30", "--warmup", "3", "--repeat", "R1", "--output", str(report)]
            result = subprocess.run(argv, capture_output=True, text=True)
            assert result.returncode == 0, (argv, result.stderr)
            proof = oracle.validate(report, lane, {"samples": 30, "warmup": 3})
            print(json.dumps({"lane": lane, "oracle": proof}), flush=True)


if __name__ == "__main__":
    main()
