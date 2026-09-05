#!/usr/bin/env python3
"""Run serialized, source-bound checks for staged payload reuse."""
import argparse
import datetime
import hashlib
import json
import os
from pathlib import Path
import re
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument("group", choices=("rust", "pptx", "lint", "diagnostic", "repository", "harness"))
parser.add_argument("--tag", required=True)
args = parser.parse_args()
assert re.fullmatch(r"[A-Za-z0-9_-]+", args.tag)
receipt = ROOT / "checks" / f"{args.tag}.json"
assert not receipt.exists()
cargo = ["cargo", "+1.98.1"]
production = ["--locked", "-p", "litchi-opc", "-p", "litchi-pptx"]
checks = {
    "rust": [
        ("opc-topology", [*cargo, "test", "--locked", "-p", "litchi-opc", "--all-features", "--lib", "topology", "--", "--test-threads=1"], {}, True),
        ("pptx-reuse", [*cargo, "test", "--locked", "-p", "litchi-pptx", "--all-features", "--lib", "source_cross_copy", "--", "--test-threads=1"], {}, True),
        ("source-api", [*cargo, "test", "--locked", "-p", "litchi-pptx", "--all-features", "--test", "source_backed_cross_copy", "--test", "source_backed_cross_copy_adversarial", "--", "--test-threads=1"], {}, True),
    ],
    "lint": [
        ("clippy-strict", [*cargo, "clippy", *production, "--all-features", "--all-targets", "--no-deps", "--", "-D", "warnings"], {}, False),
        ("clippy-diagnostic", [*cargo, "clippy", *production, "--all-features", "--all-targets", "--no-deps", "--", "-A", "clippy::chunks_exact_to_as_chunks", "-A", "clippy::clone_on_copy", "-A", "clippy::needless_lifetimes"], {}, True),
        ("rustdoc", [*cargo, "doc", *production, "--all-features", "--no-deps"], {"RUSTDOCFLAGS": "-D warnings"}, True),
    ],
    "harness": [
        ("cross-copy", [*cargo, "test", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics", "--lib", "cross_copy", "--", "--test-threads=1"], {}, True),
    ],
    "repository": [
        ("boundaries", [sys.executable, "-B", "tools/check_crate_boundaries.py"], {}, True),
        ("crud-index", [sys.executable, "-B", "tools/validate_crud_coverage_index.py"], {}, True),
        ("claims", [sys.executable, "-B", "tools/check_perf_claims.py", "--registry", "docs/performance/claim-registry-v1.json", "--repo-root", ".", "--evidence-root", ".", "--mode", "strict"], {}, True),
    ],
}["rust" if args.group == "pptx" else "lint" if args.group == "diagnostic" else args.group]
if args.group in ("pptx", "diagnostic"):
    checks = checks[1:]


def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sources():
    roots = ("crates/litchi-opc/src", "crates/litchi-pptx/src", "crates/litchi-pptx/tests", "tools/perf-baseline/src")
    return {str(p.relative_to(REPO)): hashlib.sha256(p.read_bytes()).hexdigest()
            for directory in roots for p in sorted((REPO / directory).rglob("*.rs"))}


record = {"change": 424, "group": args.group, "status": "running", "source_before": sources(), "checks": []}
environment = {"RUSTUP_TOOLCHAIN": "1.98.1", "CARGO_BUILD_JOBS": "4", "CARGO_INCREMENTAL": "0", "PYTHONDONTWRITEBYTECODE": "1"}
record["environment"] = environment
for name, argv, extra, required in checks:
    path = ROOT / "checks" / f"{args.tag}-{name}.log"
    assert not path.exists()
    item = {"name": name, "argv": argv, "environment": extra, "required": required, "started_utc": now(), "log": str(path.relative_to(ROOT))}
    print(f"{name}: running", flush=True)
    with path.open("wb") as out:
        result = subprocess.run(argv, cwd=REPO, env=os.environ | environment | extra, stdout=out, stderr=subprocess.STDOUT)
    item.update(exit_code=result.returncode, finished_utc=now())
    if "test" in argv:
        counts = re.findall(r"test result: ok\. (\d+) passed", path.read_text())
        item["passed_tests"] = sum(map(int, counts))
        item["nonzero_test_selection"] = item["passed_tests"] > 0
    record["checks"].append(item)
    receipt.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n")
    print(f"{name}: exit {result.returncode}", flush=True)
    if required and (result.returncode != 0 or item.get("nonzero_test_selection") is False):
        break
record["source_after"] = sources()
record["source_unchanged"] = record["source_before"] == record["source_after"]
record["status"] = "pass" if record["source_unchanged"] and len(record["checks"]) == len(checks) and all(not c["required"] or (c["exit_code"] == 0 and c.get("nonzero_test_selection") is not False) for c in record["checks"]) else "failed"
receipt.write_text(json.dumps(record, indent=2, sort_keys=True) + "\n")
raise SystemExit(0 if record["status"] == "pass" else 1)
