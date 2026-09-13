#!/usr/bin/env python3
"""Capture each matrix case in a separate pinned release process."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess
import time


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--binary", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--cpu", type=int, default=6)
    parser.add_argument("--rounds", type=int, default=3)
    args = parser.parse_args()
    binary = args.binary.resolve(strict=True)
    output = args.output.resolve()
    if output.exists() or args.rounds < 1:
        parser.error("output must be new and rounds must be positive")
    if any(base == output or base in output.parents for base in (Path("/tmp"), Path("/var/tmp"))):
        parser.error("use disk-backed capture storage")
    before = digest(binary)
    cases = subprocess.check_output([str(binary), "--list"], text=True).splitlines()
    if not cases or len(cases) != len(set(cases)):
        parser.error("binary case list is empty or duplicated")
    output.mkdir(parents=True)
    records = []
    for round_number in range(1, args.rounds + 1):
        for case in cases:
            if not all(c.isalnum() or c in "-_" for c in case):
                raise ValueError("unsafe case name")
            prefix = f"{round_number:02}-{case}"
            command = ["/usr/bin/time", "-v", "-o", str(output / (prefix + ".time")),
                       "taskset", "-c", str(args.cpu), str(binary), "--case", case,
                       "--phase", "evaluate", "--warmups", "2", "--iterations", "31"]
            start = time.time()
            with (output / (prefix + ".jsonl")).open("w") as stdout:
                with (output / (prefix + ".stderr")).open("w") as stderr:
                    status = subprocess.run(command, stdout=stdout, stderr=stderr).returncode
            record = {"round": round_number, "case": case, "command": command,
                      "status": status, "started_unix": start, "seconds": time.time() - start}
            if status == 0:
                rows = [json.loads(line) for line in (output / (prefix + ".jsonl")).read_text().splitlines()]
                if len(rows) != 1:
                    raise ValueError("expected one result row")
                record["result"] = rows[0]
            records.append(record)
        print(f"round {round_number}: {len(cases)} cases", flush=True)
    after = digest(binary)
    receipt = {"binary": str(binary), "binary_before": before, "binary_after": after,
               "binary_unchanged": before == after, "cpu": args.cpu, "cases": cases,
               "runner_sha256": digest(Path(__file__).resolve()), "records": records,
               "artifacts": {path.name: digest(path) for path in sorted(output.iterdir())}}
    (output / "receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    raise SystemExit(0 if before == after and all(row["status"] == 0 for row in records) else 1)


if __name__ == "__main__":
    main()
