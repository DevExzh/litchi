#!/usr/bin/env python3
"""Sample retained 0485 call stacks without rebuilding or editing sources."""
import argparse
import gzip
import hashlib
from pathlib import Path
import shutil
import sys

from support import ROOT, TEMP, meta, now, read, write

OLD = ROOT.parent / "change-0484"
sys.path.insert(0, str(OLD))
import measure_routes as routes
import profile_input_metadata as inputs
import profile_routes as process


def archive(path):
    original = meta(path)
    destination = path.with_name(path.name + ".gz")
    with path.open("rb") as source, destination.open("xb") as output:
        with gzip.GzipFile(filename="", mode="wb", fileobj=output, mtime=0) as packed:
            shutil.copyfileobj(source, packed)
    with gzip.open(destination, "rb") as source:
        assert hashlib.file_digest(source, "sha256").hexdigest() == original["sha256"]
    path.unlink()
    return {"original": original, "compressed": {"path": destination.name, **meta(destination)}}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("attempt")
    args = parser.parse_args()
    if not args.attempt or not all(c.isalnum() or c in "-_" for c in args.attempt):
        raise ValueError("invalid attempt")
    destination = ROOT / "callers" / args.attempt
    destination.mkdir(parents=True, exist_ok=False)
    build_path = ROOT.parent / "change-0485/build-normal.json"
    build = read(build_path)
    binary = build["binary"]
    assert meta(binary["path"]) == {k: binary[k] for k in ("bytes", "sha256")}
    records = []
    for workload in inputs.PROFILE_CASES:
        for input_mode in ("owned", "file"):
            label = f"{input_mode}-{workload}"
            directory = destination / label
            directory.mkdir()
            arm = inputs._arm_for(workload, input_mode)
            source = inputs._input_metadata(arm)
            argv = routes._axis_argv(binary, routes._axis_case(arm), arm, samples=3, warmups=1,
                                    report=directory / "report.json", resource=directory / "resource.txt")
            command = ["perf", "record", "-F", "499", "--call-graph", "dwarf,8192",
                       "-o", str(directory / "perf.data"), "--", *argv]
            started = {"label": label, "build": {"path": str(build_path), **meta(build_path)},
                       "binary": binary, "arm": arm, "source": source, "argv": argv,
                       "command": command, "driver": meta(Path(__file__)),
                       "validators": routes._script_hashes(), "started_utc": now(),
                       "samples": 3, "warmups": 1,
                       "scope": "Whole diagnostic child including setup and oracle work. Inclusive sampled stacks overlap; no syscall-count inference."}
            write(directory / "started.json", started)
            result = process._run_command(command, stdout_path=directory / "stdout.txt",
                                          stderr_path=directory / "stderr.txt", timeout_seconds=300)
            assert result["returncode"] == 0 and not result.get("timed_out"), result
            routes._check_axis_report(directory / "report.json", "normal", arm, samples=3, warmups=1,
                                      binary=binary, argv=argv,
                                      input_metadata=source if input_mode == "file" else None)
            export = ["perf", "script", "--header", "-i", str(directory / "perf.data"), "--demangle"]
            exported = process._run_command(export, stdout_path=directory / "perf-script.txt",
                                            stderr_path=directory / "perf-script.stderr", timeout_seconds=300)
            assert exported["returncode"] == 0 and not exported.get("timed_out"), exported
            for name in ("report.json", "resource.txt", "perf.data", "perf-script.txt"):
                assert (directory / name).stat().st_size > 0
            assert meta(binary["path"]) == {k: binary[k] for k in ("bytes", "sha256")}
            archives = {name: archive(directory / name) for name in ("perf.data", "perf-script.txt")}
            receipt = dict(started, status="pass", process=result, export_command=export,
                           export_process=exported, archives=archives, finished_utc=now(),
                           artifacts={p.name: meta(p) for p in directory.iterdir() if p.is_file()})
            write(directory / "receipt.json", receipt)
            records.append({"label": label, "receipt": {"path": str(directory / "receipt.json"), **meta(directory / "receipt.json")}})
            print(label, "pass", flush=True)
    write(destination / "result.json", {"status": "pass", "records": records})


if __name__ == "__main__":
    main()
