#!/usr/bin/env python3
"""Run one retained validation command under the 0487 CPU lock.

The command is intentionally a small receipt wrapper.  It does not overwrite
an attempt: callers either provide a unique label or use ``--attempt`` to add
one to the label.  The source manifest is captured before and after the
command, so a build/test gate cannot silently accept a source mutation.
"""

from __future__ import annotations

import fcntl
from pathlib import Path
import subprocess
import sys

from support import ENV, ROOT, REPO, TEMP, environment, meta, now, snapshot, sha, write


def parse_args() -> tuple[str, str | None, list[str]]:
    values = list(sys.argv[1:])
    attempt: str | None = None
    if values and values[0] == "--attempt":
        if len(values) < 3:
            raise SystemExit("usage: gate.py [--attempt ATTEMPT] LABEL COMMAND ...")
        attempt = values[1]
        del values[:2]
    if len(values) < 2:
        raise SystemExit("usage: gate.py [--attempt ATTEMPT] LABEL COMMAND ...")
    label, *argv = values
    if attempt is not None:
        if not attempt or "/" in attempt or any(char.isspace() for char in attempt):
            raise SystemExit("gate.py: attempt must be a non-empty path-safe token")
        label = f"{label}-{attempt}"
    if not label or "/" in label or label in {".", ".."}:
        raise SystemExit("gate.py: label must be a path-safe token")
    return label, attempt, argv


def run() -> int:
    label, attempt, argv = parse_args()
    prefix = ROOT / "validation" / label
    started_path = prefix.with_suffix(".started.json")
    stdout_path = prefix.with_suffix(".stdout")
    stderr_path = prefix.with_suffix(".stderr")
    result_path = prefix.with_suffix(".json")
    if any(path.exists() for path in (started_path, stdout_path, stderr_path, result_path)):
        raise RuntimeError(f"validation attempt already exists for {label}")

    before = snapshot()
    started = {
        "schema": "docx-stream-append-gate-v1",
        "label": label,
        "attempt": attempt,
        "argv": argv,
        "cwd": str(REPO),
        "environment": environment(),
        "driver_sha256": sha(Path(__file__)),
        "common_sha256": sha(ROOT / "support.py"),
        "started_utc": now(),
        "source_before": before,
    }
    write(started_path, started)
    with stdout_path.open("xb") as out, stderr_path.open("xb") as err:
        completed = subprocess.run(argv, cwd=REPO, env=ENV, stdout=out, stderr=err)
    after = snapshot()
    receipt = dict(
        started,
        exit_code=completed.returncode,
        finished_utc=now(),
        source_after=after,
        source_unchanged=before == after,
        artifacts={
            path.name: meta(path)
            for path in (stdout_path, stderr_path)
        },
    )
    write(result_path, receipt)
    print(
        label,
        completed.returncode,
        "source unchanged",
        receipt["source_unchanged"],
        flush=True,
    )
    return completed.returncode


def main() -> None:
    TEMP.mkdir(parents=True, exist_ok=True)
    with (Path("/home/zhuhe/.cache/litchi-goal-0484") / "cpu.lock").open("a") as lock:
        fcntl.flock(lock, fcntl.LOCK_EX)
        raise SystemExit(run())


if __name__ == "__main__":
    main()
