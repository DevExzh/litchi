#!/usr/bin/env python3
"""Run an existing PPT qualification command with a host-process guard.

The historical 0362g runner remains the measured runner and is invoked as an
owned child.  This outer adapter only rejects unrelated Litchi compiler,
profiler, probe, heaptrack, heaptrack_inter, and heaptrack_print processes
before, during, and after that child.  It does not change the workload, timing
loop, or probe arguments.
"""

from __future__ import annotations

import json
import os
from pathlib import Path
import re
import signal
import subprocess
import sys
import time


HEAVY_NAMES = frozenset(
    {
        "cargo",
        "rustc",
        "rustdoc",
        "rustfmt",
        "clippy-driver",
        "heaptrack",
        "heaptrack_inter",
        "heaptrack_print",
        "perf",
        "valgrind",
    }
)
PROBE_COMMAND = re.compile(
    r"(?:scoped_patch_probe|ppt_phase_abba|ppt_host_process_guard|heaptrack(?:_inter|_print)?)",
    re.IGNORECASE,
)


def process_table() -> dict[int, tuple[int, str, str]]:
    table = subprocess.check_output(
        ["ps", "-eo", "pid=,ppid=,comm=,args="], text=True
    )
    rows: dict[int, tuple[int, str, str]] = {}
    for line in table.splitlines():
        fields = line.split(None, 3)
        if len(fields) != 4:
            continue
        pid, parent = int(fields[0]), int(fields[1])
        rows[pid] = (parent, fields[2], fields[3])
    return rows


def descendants(rows: dict[int, tuple[int, str, str]], roots: set[int]) -> set[int]:
    owned = set(roots)
    changed = True
    while changed:
        changed = False
        for pid, (parent, _, _) in rows.items():
            if pid not in owned and parent in owned:
                owned.add(pid)
                changed = True
    return owned


def ancestors(rows: dict[int, tuple[int, str, str]], pid: int) -> set[int]:
    owned = {pid}
    while pid in rows:
        pid = rows[pid][0]
        if pid <= 0 or pid in owned:
            break
        owned.add(pid)
    return owned


def offenders(rows: dict[int, tuple[int, str, str]], allowed: set[int]) -> list[dict[str, object]]:
    result = []
    for pid, (parent, comm, command) in rows.items():
        if pid in allowed:
            continue
        name = Path(comm).name
        if name in HEAVY_NAMES or PROBE_COMMAND.search(command):
            result.append(
                {"pid": pid, "ppid": parent, "comm": comm, "command": command}
            )
    return result


def stop(process: subprocess.Popen[object]) -> None:
    try:
        os.killpg(process.pid, signal.SIGTERM)
    except ProcessLookupError:
        return
    try:
        process.wait(timeout=2)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(process.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        process.wait(timeout=5)


def check(label: str, child_pid: int | None) -> list[dict[str, object]]:
    rows = process_table()
    allowed = ancestors(rows, os.getpid())
    if child_pid is not None:
        allowed.update(descendants(rows, {child_pid}))
    found = offenders(rows, allowed)
    if found:
        print(
            json.dumps({"guard": label, "heaptrack_names": ["heaptrack", "heaptrack_inter", "heaptrack_print"], "offenders": found}),
            file=sys.stderr,
            flush=True,
        )
    return found


def run(command: list[str]) -> int:
    if not command or any(not isinstance(item, str) or not item for item in command):
        raise SystemExit("an argv is required after --")
    if check("before", None):
        return 125
    child = subprocess.Popen(command, start_new_session=True)
    violation = False
    while child.poll() is None:
        if check("during", child.pid):
            violation = True
            stop(child)
            break
        time.sleep(0.1)
    if child.poll() is None:
        child.wait()
    if check("after", child.pid):
        violation = True
    return 125 if violation else int(child.returncode)


def main() -> int:
    raw = sys.argv[1:]
    try:
        separator = raw.index("--")
    except ValueError as error:
        raise SystemExit("use: ppt_host_process_guard.py -- COMMAND [ARGS...]") from error
    return run(raw[separator + 1 :])


if __name__ == "__main__":
    raise SystemExit(main())
