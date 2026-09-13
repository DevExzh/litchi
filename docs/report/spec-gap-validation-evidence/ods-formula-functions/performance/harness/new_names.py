#!/usr/bin/env python3
"""Capture absolute candidate-only samples for names absent from the baseline catalog."""

from __future__ import annotations

import argparse
import csv
import re
import subprocess
from pathlib import Path

FIELDS = (
    "workload",
    "name",
    "input_bytes",
    "repeat",
    "warmups",
    "iterations",
    "mean_ns",
    "p50_ns",
    "p95_ns",
    "p99_ns",
    "alloc_calls_p50",
    "alloc_calls_max",
    "dealloc_calls_p50",
    "dealloc_calls_max",
    "requested_bytes_p50",
    "requested_bytes_max",
    "released_bytes_p50",
    "released_bytes_max",
    "live_before_p50",
    "live_after_p50",
    "live_after_max",
    "peak_live_delta_p50",
    "peak_live_delta_max",
    "successes_p50",
    "successes_max",
    "checksum_p50",
    "checksum_max",
    "max_rss_kib",
    "status",
)

NAMES = ("BITAND", "BIN2DEC", "BINOM.DIST.RANGE", "MDETERM", "UNICODE", "DDE")


def kv(line: str) -> dict[str, str]:
    return dict(token.split("=", 1) for token in line.split()[1:])


def rss(path: Path) -> str:
    match = re.search(
        r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$",
        path.read_text(),
        re.M,
    )
    return match.group(1) if match else ""


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    binary = args.binary
    out = args.output
    out.mkdir(parents=True, exist_ok=True)
    rows: list[dict[str, str]] = []
    for workload, repeat in (("lookup", 10_000), ("parse", 1_000)):
        for name in NAMES:
            stem = f"{workload}-new-{name.replace('.', '_')}"
            stdout = out / f"{stem}.stdout"
            stderr = out / f"{stem}.stderr"
            time = out / f"{stem}.time"
            status = out / f"{stem}.status"
            command = [
                "taskset",
                "-c",
                "2",
                "/usr/bin/time",
                "-v",
                "-o",
                str(time),
                str(binary),
                "--workload",
                workload,
                "--case",
                "lookup-new-name" if workload == "lookup" else "parse-new-name",
                "--name",
                name,
                "--warmups",
                "3",
                "--iterations",
                "15",
                "--repeat",
                str(repeat),
            ]
            with (out / "new-name-commands.txt").open("a") as stream:
                stream.write(" ".join(command) + "\n")
            completed = subprocess.run(command, stdout=stdout.open("w"), stderr=stderr.open("w"))
            status.write_text(f"{completed.returncode}\n")
            lines = stdout.read_text().splitlines()
            config = kv(next(line for line in lines if line.startswith("config ")))
            result = kv(next(line for line in lines if line.startswith("result ")))
            row = {
                "workload": workload,
                "name": name,
                "input_bytes": config["input_bytes"],
                "repeat": config["repeat"],
                "warmups": config["warmups"],
                "iterations": config["iterations"],
                **result,
                "max_rss_kib": rss(time),
                "status": str(completed.returncode),
            }
            rows.append({field: row.get(field, "") for field in FIELDS})
    with (out / "new-names.csv").open("w", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)


if __name__ == "__main__":
    main()
