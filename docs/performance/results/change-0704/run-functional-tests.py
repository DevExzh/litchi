#!/usr/bin/env python3
"""Run the bounded candidate-specific retention test inventory.

The ordinary package suite is useful integration evidence, but it cannot
prove the 1 MiB admission boundary.  The packet therefore carries an
explicit inventory for the new retention module and runs the broader bounded
filter that also includes the two policy intersection tests.  The receipt
records every passing test name; a green Cargo process with zero matching
tests is rejected.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import sys
import time
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
if len(sys.argv) != 1:
    raise SystemExit("usage: run-functional-tests.py")


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


candidate = json.loads((P / "source-census-candidate.json").read_text())
inventory = json.loads((P / "functional-test-inventory.json").read_text())
source_name = inventory["source"]
if source_name not in set(candidate["changed_paths"]) | set(candidate["added_paths"]):
    raise AssertionError(f"functional test source is not a candidate PPTX change: {source_name}")
source_path = ROOT / source_name
if not source_path.exists():
    raise AssertionError(f"functional test source is missing: {source_name}")
test_names = inventory["tests"]
if not test_names or len(set(test_names)) != len(test_names):
    raise AssertionError("functional test inventory must list unique tests")
source_text = source_path.read_text()
for test_name in test_names:
    if not re.search(rf"#\s*\[\s*test\s*\][\s\S]*?\bfn\s+{re.escape(test_name)}\b", source_text):
        raise AssertionError(f"listed functional test is absent from {source_name}: {test_name}")

target_dir = ROOT.parent / "litchi-target-0704"
command = ["cargo", "test", "-p", "litchi-pptx", "--locked"]
if source_name.startswith("crates/litchi-pptx/tests/"):
    command.extend(["--test", source_path.stem])
else:
    command.append("--lib")
command.extend([inventory["filter"], "--", "--test-threads=1"])
log = P / "functional-tests.log"
started = time.monotonic()
environment = dict(
    os.environ,
    CARGO_TARGET_DIR=str(target_dir),
    CARGO_BUILD_JOBS="2",
    RUSTFLAGS="-D warnings",
)
with log.open("w") as output:
    result = subprocess.run(
        command,
        cwd=ROOT,
        env=environment,
        stdout=output,
        stderr=subprocess.STDOUT,
    )
if result.returncode:
    raise SystemExit(result.returncode)
log_text = log.read_text()
passing = sorted(
    set(re.findall(r"^test\s+([^\s]+)\s+\.\.\.\s+ok$", log_text, flags=re.MULTILINE))
)
missing = [
    test_name
    for test_name in test_names
    if not any(name == test_name or name.endswith(f"::{test_name}") for name in passing)
]
missing_additional = [
    test_name
    for test_name in inventory.get("required_additional_tests", [])
    if not any(name == test_name or name.endswith(f"::{test_name}") for name in passing)
]
minimum = len(test_names) + inventory.get("minimum_additional_tests", 0)
if missing or missing_additional or len(passing) < minimum:
    raise AssertionError(
        f"functional filter did not execute every listed test or policy tests: "
        f"missing={missing}, missing_additional={missing_additional}, "
        f"passing={len(passing)}, required_minimum={minimum}"
    )
records = [{
    "source": source_name,
    "source_sha256": sha(source_path),
    "filter": inventory["filter"],
    "required_tests": test_names,
    "required_additional_tests": inventory.get("required_additional_tests", []),
    "passed_tests": passing,
    "passed_test_count": len(passing),
    "command": command,
    "exit_code": result.returncode,
    "seconds": round(time.monotonic() - started, 2),
    "log": log.name,
    "log_sha256": sha(log),
}]

(P / "functional-tests.json").write_text(
    json.dumps(
        {
            "candidate_source_file_count": candidate["source_file_count"],
            "required_policy": "the explicit changed PPTX retention module inventory must all pass",
            "tests": records,
        },
        indent=2,
    )
    + "\n"
)
print("candidate functional tests", len(passing), "passed")
