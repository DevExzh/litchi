#!/usr/bin/env python3
"""Validate retained caller profiles and summarize inclusive sample-period shares."""
import gzip
import hashlib
import re
import sys
from collections import Counter
from pathlib import Path

from support import ROOT, meta, read, sha, write

sys.path.insert(0, str(ROOT.parent / "change-0484"))
import measure_routes as routes

HEADER = re.compile(r"^(.+?)\s+(\d+)\s+([\d.]+):\s+(\d+)\s+cpu/cycles/P:\s*$")
SYMBOLS = ("statx", "ensure_current", "fill_fragment_buffer",
           "write_consumed_range", "flush_pending_consumed", "audit_splice")


def require(condition, message):
    if not condition:
        raise ValueError(message)


def summarize(script):
    samples = []
    current = None
    for line in script.splitlines():
        match = HEADER.match(line)
        if match:
            current = (match[1], int(match[4]), [])
            samples.append(current)
        elif line.startswith("\t"):
            require(current is not None, "frame without sample")
            current[2].append(line)
        elif line.strip() and not line.startswith("#"):
            raise ValueError(f"unrecognized perf-script line: {line[:160]}")
    require(bool(samples), "empty sample set")
    total = sum(period for _, period, _ in samples)
    require(total > 0, "empty period total")
    shares = {}
    for symbol in SYMBOLS:
        pattern = re.compile(r"(?<![A-Za-z0-9_])" + symbol + r"(?![A-Za-z0-9_])")
        selected = [(period, frames) for _, period, frames in samples
                    if any(pattern.search(frame) for frame in frames)]
        period = sum(value for value, _ in selected)
        shares[symbol] = {"samples": len(selected), "period": period,
                          "percent": 100 * period / total}
    return {"samples": len(samples), "period": total,
            "comm_samples": dict(Counter(comm for comm, _, _ in samples)),
            "inclusive": shares}


def main():
    result_path = ROOT / "callers/baseline1/result.json"
    result = read(result_path)
    require(result["status"] == "pass", "capture failed")
    expected = {f"{mode}-{case}" for mode in ("owned", "file")
                for case in ("s131072-a64-short-c64", "s64-a16384-short-c64")}
    require(len(result["records"]) == 4 and
            {r["label"] for r in result["records"]} == expected, "profile set mismatch")
    rows = []
    for record in result["records"]:
        path = Path(record["receipt"]["path"])
        require(meta(path) == {k: record["receipt"][k] for k in ("bytes", "sha256")},
                "receipt hash mismatch")
        receipt = read(path)
        require(receipt["status"] == "pass", "profile failed")
        require(receipt["driver"] == meta(ROOT / "profile_callers.py"), "driver mismatch")
        require(receipt["validators"] == routes._script_hashes(), "validator mismatch")
        for key in ("binary", "build"):
            entry = receipt[key]
            require(meta(entry["path"]) == {k: entry[k] for k in ("bytes", "sha256")},
                    f"{key} mismatch")
        for key in ("process", "export_process"):
            require(receipt[key]["returncode"] == 0 and not receipt[key].get("timed_out"),
                    f"{key} failed")
        for name, expected_meta in receipt["artifacts"].items():
            require(meta(path.parent / name) == expected_meta, f"artifact mismatch: {name}")
        for name, archive in receipt["archives"].items():
            packed = path.parent / archive["compressed"]["path"]
            require(meta(packed) == {k: archive["compressed"][k] for k in ("bytes", "sha256")},
                    f"archive mismatch: {name}")
            unpacked = gzip.decompress(packed.read_bytes())
            require({"bytes": len(unpacked), "sha256": hashlib.sha256(unpacked).hexdigest()}
                    == archive["original"], f"decompressed mismatch: {name}")
        arm = receipt["arm"]
        routes._check_axis_report(path.parent / "report.json", "normal", arm,
                                  samples=3, warmups=1, binary=receipt["binary"],
                                  argv=receipt["argv"],
                                  input_metadata=receipt["source"] if arm["input_mode"] == "file" else None)
        script = gzip.decompress((path.parent / "perf-script.txt.gz").read_bytes()).decode()
        rows.append({"label": record["label"], "receipt": record["receipt"], **summarize(script)})
    summary = {"schema": "docx-caller-sample-summary-v1", "status": "pass",
               "driver_sha256": sha(Path(__file__)), "capture": meta(result_path),
               "scope": "Whole diagnostic child, including setup/oracles and launcher samples; inclusive symbols overlap. Period-weighted sampled shares are not syscall counts or paired speedups. Unknown kernel frames and addr2line warnings remain in raw artifacts.",
               "profiles": rows}
    write(ROOT / "caller-summary.json", summary)
    for row in rows:
        print(row["label"], row["samples"], "samples; statx period share",
              f'{row["inclusive"]["statx"]["percent"]:.3f}%')


if __name__ == "__main__":
    main()
