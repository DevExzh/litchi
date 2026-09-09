#!/usr/bin/env python3
"""Retain separate PMU and syscall diagnostics for the 0487 executables."""
import argparse
from pathlib import Path
import sys

from support import ROOT, ENV, meta, now, read, write

SEALED = ROOT.parent / "change-0484"
BASELINE = ROOT.parent / "change-0485"
sys.path.insert(0, str(SEALED))
import measure_routes as routes
import profile_input_metadata as inputs
import profile_routes as process


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("attempt")
    args = parser.parse_args()
    if not args.attempt or "/" in args.attempt or args.attempt in (".", ".."):
        raise ValueError("invalid attempt")
    destination = ROOT / "profiles" / args.attempt
    destination.mkdir(parents=True, exist_ok=False)
    # The before executable is the retained 0485 after-build.  The 0484 tree
    # remains a sealed source of route construction and report validation.
    builds = {"before": BASELINE / "build-normal.json",
              "after": ROOT / "build-normal.json"}
    records = []
    for version in ("before", "after"):
        build = read(builds[version])
        binary = build["binary"]
        assert meta(binary["path"]) == {k: binary[k] for k in ("bytes", "sha256")}
        for workload in inputs.PROFILE_CASES:
            for input_mode in ("owned", "file"):
                arm = inputs._arm_for(workload, input_mode)
                case = routes._axis_case(arm)
                source = inputs._input_metadata(arm)
                for tool in (("perf",) if version == "before" else ("perf", "strace")):
                    label = f"{version}-{tool}-{input_mode}-{workload}"
                    directory = destination / label
                    directory.mkdir()
                    argv = routes._axis_argv(binary, case, arm, samples=1, warmups=1,
                                            report=directory / "report.json", resource=directory / "resource.txt")
                    if tool == "strace":
                        command = ["strace", "-f", "-c", "-e", inputs.TRACE_FILTER,
                                   "-o", str(directory / "profile.txt"), "--", *argv]
                    else:
                        command = ["perf", "stat", "-x,", "-o", str(directory / "profile.txt"),
                                   "-e", "cycles,instructions,branches,branch-misses,page-faults", "--", *argv]
                    started = {"label": label, "build": {"path": str(builds[version]), **meta(builds[version])},
                               "binary": binary, "arm": arm, "source": source, "argv": argv, "command": command,
                               "driver": meta(Path(__file__)), "validators": routes._script_hashes(),
                               "started_utc": now(), "samples": 1, "warmups": 1,
                               "scope": "Whole diagnostic child including setup and oracle work; profiler overhead included."}
                    write(directory / "started.json", started)
                    result = process._run_command(command, stdout_path=directory / "stdout.txt",
                                                  stderr_path=directory / "stderr.txt", timeout_seconds=300)
                    status = "pass"
                    error = None
                    try:
                        if result["returncode"] != 0 or result.get("timed_out"):
                            raise ValueError(f"profiler failed or unavailable: {result}")
                        for name in ("report.json", "resource.txt", "profile.txt"):
                            assert (directory / name).stat().st_size > 0
                        routes._check_axis_report(directory / "report.json", "normal", arm,
                                                  samples=1, warmups=1, binary=binary, argv=argv,
                                                  input_metadata=source if input_mode == "file" else None)
                        assert meta(binary["path"]) == {k: binary[k] for k in ("bytes", "sha256")}
                    except Exception as caught:
                        status, error = "failed_or_unavailable", str(caught)
                    receipt = dict(started, status=status, error=error, process=result, finished_utc=now(),
                                   artifacts={p.name: meta(p) for p in directory.iterdir() if p.is_file()})
                    write(directory / "receipt.json", receipt)
                    records.append({"label": label, "status": status,
                                    "receipt": {"path": str(directory / "receipt.json"), **meta(directory / "receipt.json")}})
                    print(label, status, flush=True)
    write(destination / "result.json", {"records": records, "finished_utc": now(),
                                        "status": "pass" if all(r["status"] == "pass" for r in records) else "incomplete"})
    if any(r["status"] != "pass" for r in records):
        raise SystemExit(1)


if __name__ == "__main__":
    main()
