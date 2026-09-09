#!/usr/bin/env python3
"""Record host load around a measurement phase without changing affinity."""
import os
from pathlib import Path
import sys

from support import ROOT, meta, now, write


def main():
    label = sys.argv[1]
    if not label or not all(c.isalnum() or c in "-_" for c in label):
        raise ValueError("invalid observation label")
    processes = []
    for directory in Path("/proc").iterdir():
        if not directory.name.isdigit():
            continue
        try:
            name = (directory / "comm").read_text().strip()
            if name in {"cargo", "rustc", "rustdoc", "clippy-driver", "perf", "strace"}:
                processes.append({"pid": int(directory.name), "name": name,
                                  "affinity": sorted(os.sched_getaffinity(int(directory.name)))})
        except (OSError, ProcessLookupError):
            continue
    write(ROOT / "host-observations" / f"{label}.json", {
        "observed_utc": now(), "driver": meta(Path(__file__)),
        "loadavg": Path("/proc/loadavg").read_text(),
        "proc_stat": Path("/proc/stat").read_text(),
        "coordinator_affinity": sorted(os.sched_getaffinity(0)),
        "compiler_and_profiler_processes": processes,
        "scope": "Point observation only; CPU pinning and our shared lock do not guarantee host exclusivity.",
    })


if __name__ == "__main__":
    main()
