#!/usr/bin/env python3
"""Capture alternating long-reference lanes for annotation-only comparison."""
from __future__ import annotations
import csv, hashlib, json, re, shlex, subprocess, sys, time
from pathlib import Path

CASES = (
    ("coverage-colon-sheet-1k", 1000),
    ("coverage-colon-sheet-4k", 1000),
    ("coverage-colon-sheet-16k", 128),
    ("coverage-source-iri-over-16k", 128),
)
FIELDS = (
    "group", "workload", "case", "input_bytes", "repeat", "warmups", "iterations",
    "expected_success", "mean_ns", "p50_ns", "p95_ns", "p99_ns",
    "alloc_calls_p50", "alloc_calls_max", "dealloc_calls_p50", "dealloc_calls_max",
    "requested_bytes_p50", "requested_bytes_max", "released_bytes_p50", "released_bytes_max",
    "live_before_p50", "live_after_p50", "live_after_max", "peak_live_delta_p50",
    "peak_live_delta_max", "successes_p50", "successes_max", "checksum_p50", "checksum_max",
    "max_rss_kib", "status",
)

def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()

def kv(line: str) -> dict[str, str]:
    return dict(token.split("=", 1) for token in line.split()[1:])

def rss(path: Path) -> str:
    match = re.search(r"^\s*Maximum resident set size \(kbytes\):\s*(\d+)\s*$", path.read_text(), re.M)
    return match.group(1) if match else ""

def capture(binary: Path, out: Path) -> list[dict[str, str]]:
    out.mkdir(parents=True, exist_ok=True)
    rows = []
    for case, repeat in CASES:
        stem = f"parse-{case}"
        stdout = out / f"{stem}.stdout"
        stderr = out / f"{stem}.stderr"
        timing = out / f"{stem}.time"
        status = out / f"{stem}.status"
        command = [
            "taskset", "-c", "2", "/usr/bin/time", "-v", "-o", str(timing), str(binary),
            "--workload", "parse", "--case", case, "--warmups", "3", "--iterations", "15",
            "--repeat", str(repeat),
        ]
        with (out / "commands.txt").open("a") as stream:
            stream.write(shlex.join(command) + "\n")
        with stdout.open("w") as so, stderr.open("w") as se:
            completed = subprocess.run(command, stdout=so, stderr=se)
        status.write_text(f"{completed.returncode}\n")
        lines = stdout.read_text().splitlines()
        config = kv(next(line for line in lines if line.startswith("config ")))
        result = kv(next(line for line in lines if line.startswith("result ")))
        row = {
            "group": "annotation-long",
            "workload": "parse",
            "case": case,
            "input_bytes": config["input_bytes"],
            "repeat": config["repeat"],
            "warmups": config["warmups"],
            "iterations": config["iterations"],
            "expected_success": config["expected_success"],
            **result,
            "max_rss_kib": rss(timing),
            "status": str(completed.returncode),
        }
        rows.append({field: row.get(field, "") for field in FIELDS})
    with (out / "raw.csv").open("w", newline="") as stream:
        writer = csv.DictWriter(stream, fieldnames=FIELDS, lineterminator="\n")
        writer.writeheader()
        writer.writerows(rows)
    return rows

def main() -> int:
    if len(sys.argv) != 4:
        raise SystemExit(f"usage: {sys.argv[0]} LABEL BINARY OUTPUT")
    label, binary_text, output_text = sys.argv[1:]
    binary = Path(binary_text); output = Path(output_text)
    start = time.time(); started = time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(start))
    output.mkdir(parents=True, exist_ok=True)
    bh = sha(binary)
    (output / "run-start.json").write_text(json.dumps({
        "label": label, "binary": str(binary), "binary_sha256": bh,
        "cases": [{"case": case, "repeat": repeat} for case, repeat in CASES],
        "warmups": 3, "iterations": 15, "started_at": started,
    }, indent=2) + "\n")
    rows = capture(binary, output)
    finished = time.time()
    (output / "run-end.json").write_text(json.dumps({
        "label": label, "binary": str(binary), "binary_sha256": bh,
        "status": 0 if all(r["status"] == "0" for r in rows) else 1,
        "row_count": len(rows), "raw_sha256": sha(output / "raw.csv"),
        "elapsed_seconds": finished - start,
        "finished_at": time.strftime("%Y-%m-%dT%H:%M:%SZ", time.gmtime(finished)),
    }, indent=2) + "\n")
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
