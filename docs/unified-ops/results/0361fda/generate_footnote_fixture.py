#!/usr/bin/env python3
"""Create the external Pages footnote workload from the tracked graph builder.

The generator example includes the exact synthetic graph builder used by the
body-footnote integration test.  Cargo builds into an external target and the
output is written outside the checkout; the qualification runner replays this
command and byte-compares the result before accepting owner timings.
"""

from __future__ import annotations

import argparse
import os
from pathlib import Path
import signal
import subprocess


ROOT = Path(__file__).resolve().parents[4]
TIMEOUT_S = 900.0


def terminate_group(child: subprocess.Popen[bytes]) -> int:
    try:
        os.killpg(child.pid, signal.SIGTERM)
    except ProcessLookupError:
        pass
    try:
        return child.wait(timeout=5.0)
    except subprocess.TimeoutExpired:
        try:
            os.killpg(child.pid, signal.SIGKILL)
        except ProcessLookupError:
            pass
        return child.wait(timeout=5.0)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    output = args.output.resolve()
    if output.is_relative_to(ROOT):
        raise SystemExit(f"output must be outside the checkout: {output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    command = [
        "cargo",
        "run",
        "--offline",
        "--manifest-path",
        str(ROOT / "Cargo.toml"),
        "-p",
        "litchi-pages",
        "--example",
        "generate_footnote_fixture",
        "--",
        "--output",
        str(output),
    ]
    if (ROOT / "Cargo.lock").is_file():
        command[2:2] = ["--locked"]
    target = output.parent / ".footnote-generator-target"
    environment = dict(os.environ, CARGO_TARGET_DIR=str(target), CARGO_TERM_COLOR="never")
    child = subprocess.Popen(command, cwd=ROOT, env=environment, start_new_session=True)
    try:
        status = child.wait(timeout=TIMEOUT_S)
    except subprocess.TimeoutExpired as error:
        terminate_group(child)
        raise SystemExit(f"footnote generator exceeded {TIMEOUT_S:g}s") from error
    if status:
        raise SystemExit(f"footnote generator failed with status {status}")
    if not output.is_file():
        raise SystemExit(f"footnote generator did not produce {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
