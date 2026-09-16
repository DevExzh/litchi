#!/usr/bin/env python3
"""Run one leg of the change-0653 semantic-read probe over the fixture corpus.

A thread pool drives subprocesses: this host's Python 3.14 breaks
`ProcessPoolExecutor` when the script has top-level code, and the probe is a
separate binary anyway.
"""

import os
import subprocess
import sys
from concurrent.futures import ThreadPoolExecutor

ROOT = "/home/zhuhe/code/litchi-worktrees/before-70d7768cc/test-data"
TIMEOUT_SECONDS = 180
WORKERS = 12


def fixtures():
    found = []
    for directory, _subdirectories, names in os.walk(ROOT):
        for name in names:
            if name.lower().endswith((".docx", ".xlsx", ".pptx")):
                found.append(os.path.join(directory, name))
    found.sort()
    return found


def run_one(binary, path, extra):
    display = os.path.relpath(path, ROOT)
    command = [binary, path, display] + extra
    try:
        completed = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            timeout=TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired:
        return display, f"{display}\tprobe.outcome\tTIMEOUT\n"
    text = completed.stdout.decode("utf-8", "replace")
    if completed.returncode != 0:
        text += f"{display}\tprobe.outcome\tEXIT={completed.returncode}\n"
    return display, text


def main():
    binary = sys.argv[1]
    destination = sys.argv[2]
    extra = sys.argv[3:]
    paths = fixtures()
    results = {}
    with ThreadPoolExecutor(max_workers=WORKERS) as pool:
        for display, text in pool.map(lambda path: run_one(binary, path, extra), paths):
            results[display] = text
    with open(destination, "w", encoding="utf-8") as handle:
        for display in sorted(results):
            handle.write(results[display])
    total_lines = sum(text.count("\n") for text in results.values())
    print(f"fixtures={len(paths)} lines={total_lines} out={destination}")


main()
