#!/usr/bin/env python3
import hashlib
import json
import os
import subprocess
import time
from datetime import datetime, timezone
from pathlib import Path

ROOT = Path("/home/zhuhe/litchi-performance-0932")
PACKET = ROOT / "heaptrack-packet-v1"
ANALYSIS = ROOT / "heaptrack-analysis-v1"
CHILDREN = sorted((PACKET / "children").iterdir())


def now():
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def sha(path):
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def desc(path):
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def write_xb(path, value):
    data = (json.dumps(value, indent=2, sort_keys=True) + "\\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def main():
    if ANALYSIS.exists():
        raise SystemExit(f"refusing existing analysis directory: {ANALYSIS}")
    ANALYSIS.mkdir()
    started = now()
    rows = []
    for child in CHILDREN:
        profile = next(child.glob("heaptrack-data*"))
        base = ANALYSIS / child.name
        base.mkdir()
        specs = [
            (
                "full-summary",
                [
                    "/usr/bin/heaptrack_print", "-f", str(profile),
                    "-m", "1", "-n", "1000000", "-s", "1000000",
                    "-a", "-p", "-T", "-l",
                    "-H", str(base / "allocation-size.histogram"),
                    "-F", str(base / "allocation-stacks.folded"),
                    "--flamegraph-cost-type", "allocations",
                ],
                base / "full-summary.txt",
            ),
            (
                "peak-stacks",
                [
                    "/usr/bin/heaptrack_print", "-f", str(profile),
                    "-m", "1", "-n", "1000000", "-s", "1000000",
                    "-F", str(base / "peak-stacks.folded"),
                    "--flamegraph-cost-type", "peak",
                ],
                base / "peak-stacks.txt",
            ),
            (
                "temporary-stacks",
                [
                    "/usr/bin/heaptrack_print", "-f", str(profile),
                    "-m", "1", "-n", "1000000", "-s", "1000000",
                    "-F", str(base / "temporary-stacks.folded"),
                    "--flamegraph-cost-type", "temporary",
                ],
                base / "temporary-stacks.txt",
            ),
        ]
        for label, argv, stdout_path in specs:
            stderr_path = stdout_path.with_suffix(".stderr")
            started_one = now()
            monotonic = time.monotonic()
            status = "completed"
            return_code = None
            error = None
            try:
                with stdout_path.open("xb") as stdout, stderr_path.open("xb") as stderr:
                    proc = subprocess.run(
                        argv,
                        cwd=str(ROOT),
                        stdout=stdout,
                        stderr=stderr,
                        timeout=600,
                        check=False,
                    )
                    return_code = proc.returncode
            except subprocess.TimeoutExpired as exc:
                status = "timeout"
                error = f"{type(exc).__name__}: {exc}"
            except Exception as exc:
                status = "runner_error"
                error = f"{type(exc).__name__}: {exc}"
            ended_one = now()
            receipt = {
                "schema": "litchi.performance.0932.heaptrack-print-terminal.v1",
                "status": status,
                "label": label,
                "child": child.name,
                "profile": desc(profile),
                "argv": argv,
                "started_utc": started_one,
                "ended_utc": ended_one,
                "wall_seconds": time.monotonic() - monotonic,
                "return_code": return_code,
                "error": error,
                "stdout": desc(stdout_path),
                "stderr": desc(stderr_path),
                "generated": [
                    desc(p) for p in sorted(base.iterdir())
                    if p not in (stdout_path, stderr_path) and p.is_file()
                ],
            }
            receipt_path = base / f"{label}.terminal.json"
            write_xb(receipt_path, receipt)
            rows.append(receipt)
    outer = {
        "schema": "litchi.performance.0932.heaptrack-print-run.v1",
        "status": "terminal",
        "started_utc": started,
        "ended_utc": now(),
        "analysis_started_after_capture_terminal": True,
        "commands": rows,
    }
    write_xb(ANALYSIS / "analysis-terminal.json", outer)
    print(json.dumps({"analysis": str(ANALYSIS), "commands": len(rows)}, sort_keys=True))
    return 0 if all(row["status"] == "completed" and row["return_code"] == 0 for row in rows) else 1


if __name__ == "__main__":
    raise SystemExit(main())
