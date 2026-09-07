#!/usr/bin/env python3
"""Exercise inner evidence checks after resealing deliberately altered copies."""

import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

ROOT = Path(__file__).resolve().parent


def seal(root):
    rows = []
    for path in sorted(root.rglob("*")):
        if path.is_symlink():
            raise ValueError(f"unexpected symlink: {path}")
        if path.is_file() and path != root / "SHA256SUMS":
            with path.open("rb") as stream:
                digest = hashlib.file_digest(stream, "sha256").hexdigest()
            rows.append(f"{digest}  {path.relative_to(root).as_posix()}\n")
    (root / "SHA256SUMS").write_text("".join(rows))


def main():
    receipt = ROOT / "negative-verification.json"
    if receipt.exists():
        raise SystemExit("refusing to overwrite negative-verification.json")
    records = []
    for case in ("summary-statistic", "fuzz-input", "fuzz-extra-file"):
        with tempfile.TemporaryDirectory(prefix="litchi-0457-negative-") as directory:
            target = Path(directory) / ROOT.name
            shutil.copytree(ROOT, target)
            if case == "summary-statistic":
                path = target / "candidate-final/summary.json"
                value = json.loads(path.read_text())
                value["by_mode_shape"]["allocator"]["large"]["allocated_bytes"][0]["p50"] += 1
                path.write_text(json.dumps(value, indent=2) + "\n")
                expected = "summary differs"
            elif case == "fuzz-input":
                path = target / "fuzz/seeds/xml/good.xml"
                path.write_bytes(path.read_bytes() + b" ")
                expected = "hash differs for seeds/xml/good.xml"
            else:
                (target / "fuzz/unlisted-input.txt").write_text("unexpected input\n")
                expected = "checksum coverage differs"
            seal(target)
            command = [sys.executable, "-B", str(target / "verify.py")]
            result = subprocess.run(command, cwd=directory, capture_output=True, text=True)
            output = json.loads(result.stdout)
            failures = output.get("failures", [])
            passed = result.returncode == 1 and output.get("status") == "fail" and any(expected in item for item in failures)
            records.append({"case": case, "status": "pass" if passed else "failed", "exit_code": result.returncode,
                            "expected_failure": expected, "verifier_output": output, "stderr": result.stderr})
        records[-1]["temporary_directory_absent"] = not Path(directory).exists()
        print(f"{case}: {records[-1]['status']}", flush=True)
    value = {"schema": "litchi-0457-negative-verification-v1", "change": 457,
             "status": "pass" if all(row["status"] == "pass" and row["temporary_directory_absent"] for row in records) else "failed",
             "verifier_sha256": hashlib.sha256((ROOT / "verify.py").read_bytes()).hexdigest(),
             "driver_sha256": hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
             "scope": "Three altered copies with regenerated root seals; inner summary/input custody checks must reject each.",
             "records": records}
    with receipt.open("x") as stream:
        json.dump(value, stream, indent=2)
        stream.write("\n")
    return 0 if value["status"] == "pass" else 1


if __name__ == "__main__":
    raise SystemExit(main())
