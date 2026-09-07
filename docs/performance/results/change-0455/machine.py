#!/usr/bin/env python3
"""Record the host identity used by change-0455 captures.

The output is intentionally the small, text-preserving shape used by the
previous lifecycle bundle.  It records host context only; it does not launch a
benchmark, build a binary, or change filesystem cache state.
"""

from __future__ import annotations

import argparse
import datetime as dt
import json
import os
from pathlib import Path
import platform
import subprocess


ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def command(argv: list[str]) -> str | None:
    try:
        result = subprocess.run(
            argv,
            cwd=REPO,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
        )
    except (OSError, subprocess.CalledProcessError):
        return None
    return result.stdout


def text_file(path: Path) -> str | None:
    try:
        return path.read_text(encoding="utf-8")
    except (OSError, UnicodeError):
        return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, default=ROOT / "machine.json")
    args = parser.parse_args()
    if args.output.exists():
        raise SystemExit(f"refusing to replace existing machine record: {args.output}")
    uname = list(os.uname()) if hasattr(os, "uname") else list(platform.uname())
    affinity = sorted(os.sched_getaffinity(0)) if hasattr(os, "sched_getaffinity") else None
    record = {
        "schema": "change0455_machine_v1",
        "captured_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "uname": uname,
        "lscpu": command(["lscpu"]),
        "rustc": command(["rustc", "-Vv"]),
        "meminfo": text_file(Path("/proc/meminfo")),
        "logical_cpus": os.cpu_count(),
        "process_affinity": affinity,
        "cargo_config": text_file(REPO / ".cargo/config.toml"),
        "rust_toolchain": text_file(REPO / "rust-toolchain.toml"),
        "environment": {
            "RUSTUP_TOOLCHAIN": os.environ.get("RUSTUP_TOOLCHAIN"),
            "CARGO_BUILD_JOBS": os.environ.get("CARGO_BUILD_JOBS"),
            "CARGO_INCREMENTAL": os.environ.get("CARGO_INCREMENTAL"),
            "RUSTFLAGS": os.environ.get("RUSTFLAGS"),
        },
    }
    args.output.write_text(json.dumps(record, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"status": "pass", "output": str(args.output)}))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
