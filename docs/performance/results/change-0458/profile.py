#!/usr/bin/env python3
"""Capture separate whole-process counters and sampled stacks after timing."""
import importlib.util
import json
import os
from pathlib import Path
import subprocess

import capture

ROOT = Path(__file__).resolve().parent


def run(directory, name, argv, environment):
    with (directory / (name + ".stdout")).open("xb") as out, (directory / (name + ".stderr")).open("xb") as err:
        result = subprocess.run(argv, cwd=capture.REPO, env=environment, stdout=out, stderr=err)
    return {"argv": argv, "exit_code": result.returncode}


def main():
    protocol = json.loads((ROOT / "protocol.json").read_text())
    assert protocol["profile_sha256"] == capture.sha(Path(__file__))
    assert protocol["oracle_sha256"] == capture.sha(ROOT / "oracle.py")
    assert protocol["binary_binding_sha256"] == capture.sha(ROOT / "binary-binding.json")
    binding = json.loads((ROOT / "binary-binding.json").read_text())["binaries"]["normal"]
    assert capture.sha(Path(binding["path"])) == binding["sha256"]
    spec = importlib.util.spec_from_file_location("oracle458", ROOT / "oracle.py")
    oracle = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(oracle)
    profiles = ROOT / "profiling"
    profiles.mkdir()
    environment = os.environ | protocol["environment"]
    for kind in ("counters", "samples"):
        directory = profiles / kind
        directory.mkdir()
        report = directory / "report.json"
        lane = {"id": kind, "scope": "phases", "shape": "large", "instrumentation": "normal", "repeat": "diagnostic"}
        substitutions = dict(lane, binary=binding["path"], report=str(report))
        workload = [part.format_map(substitutions) for part in protocol["diagnostic_argv"]]
        if kind == "counters":
            prefix = ["perf", "stat", "-x,", "-o", str(directory / "counters.csv"), "-e", "cycles,instructions,branches,branch-misses,cache-misses,page-faults,context-switches"]
        else:
            prefix = ["perf", "record", "--no-buildid-cache", "-F", "99", "-e", "cycles:u", "--call-graph", "dwarf,8192", "-o", str(directory / "perf.data")]
        command = ["/usr/bin/time", "-v", "-o", str(directory / "resource.log"), "taskset", "-c", str(protocol["cpu"]), *prefix, "--", *workload]
        record = {"schema": "litchi-0458-profile-v1", "change": 458, "kind": kind,
                  "protocol_sha256": capture.sha(ROOT / "protocol.json"), "binary": binding,
                  "started_utc": capture.now(), "scope": "whole process including fixture setup, warmups, API calls, output checks, teardown and reporting"}
        record["workload"] = run(directory, "workload", command, environment)
        try:
            assert record["workload"]["exit_code"] == 0
            record["oracle"] = oracle.validate(report, lane, protocol | {"samples": 100})
            if kind == "samples":
                record["top_symbols"] = run(directory, "top-symbols", ["perf", "report", "--stdio", "--no-inline", "--no-children", "--call-graph", "none", "--percent-limit", "0", "-i", str(directory / "perf.data")], environment)
                record["script"] = run(directory, "perf-script", ["perf", "script", "--no-inline", "-i", str(directory / "perf.data")], environment)
                assert record["top_symbols"]["exit_code"] == record["script"]["exit_code"] == 0
            assert capture.sha(Path(binding["path"])) == binding["sha256"]
            record["status"] = "pass"
        except Exception as error:
            record["status"] = "failed"
            record["error"] = str(error)
        record["finished_utc"] = capture.now()
        record["artifacts"] = [capture.artifact(path) for path in sorted(directory.iterdir()) if path.is_file()]
        capture.write(directory / "receipt.json", record)
        print(kind, record["status"], flush=True)
        if record["status"] != "pass":
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
