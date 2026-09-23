#!/usr/bin/env python3
"""Reduce sealed-0734-probe reports to their timing and oracle verdicts.

The 0734 probe writes about 1 MB per process (per-sample stream inventories
and oracle detail). This keeps, per report: the top-level identity and
expected-oracle verdict, and per sample its index, whole_ns, output digest and
oracle verdict. It records the SHA-256 and size of every original report in
`reduction-manifest.json` before replacing it, so the reduction is auditable
and the full reports are reproducible from the recorded commands.

Usage: reduce_p0734.py MATRIX_DIR
"""

import glob
import hashlib
import json
import os
import sys

KEEP = (
    "schema_version", "case", "format", "operation", "input", "warmups",
    "samples_requested", "source_sha256", "expected_output_sha256",
)


def main():
    matrix = sys.argv[1]
    manifest = {}
    for path in sorted(glob.glob(f"{matrix}/p0734-*/r*-[ABC].json")):
        with open(path, "rb") as handle:
            raw = handle.read()
        report = json.loads(raw)
        reduced = {key: report[key] for key in KEEP if key in report}
        reduced["expected_oracle"] = {"oracle_ok": report["expected_oracle"]["oracle_ok"]}
        reduced["oracle_controls_count"] = len(report.get("oracle_controls", []))
        reduced["samples"] = [
            {
                "index": sample["index"],
                "phase_ns": {"whole_ns": sample["phase_ns"]["whole_ns"]},
                "output_sha256": sample["output_sha256"],
                "oracle": {"oracle_ok": sample["oracle"]["oracle_ok"]},
            }
            for sample in report["samples"]
        ]
        manifest[os.path.relpath(path, matrix)] = {
            "original_sha256": hashlib.sha256(raw).hexdigest(),
            "original_bytes": len(raw),
        }
        with open(path, "w") as handle:
            json.dump(reduced, handle, separators=(",", ":"))
    with open(f"{matrix}/reduction-manifest.json", "w") as handle:
        json.dump(manifest, handle, indent=1)
    print(f"reduced {len(manifest)} reports")


if __name__ == "__main__":
    main()
