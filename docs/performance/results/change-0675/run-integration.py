#!/usr/bin/env python3
"""Run the third-wave integration gates and retain each command and exit status."""
import json
import os
from pathlib import Path
import subprocess
import time

ROOT = Path.cwd()
OUT = Path(os.environ.get("LITCHI_GATE_OUTPUT", "/tmp/litchi-0675/integration"))
OUT.mkdir(parents=True, exist_ok=True)
PACKAGES = "litchi-core soapberry-zip litchi-cfb litchi-ole-common litchi-opc litchi-ooxml-common litchi-doc litchi-docx litchi-ppt litchi-pptx litchi-xls litchi-xlsx litchi-xlsb xml-minifier".split()
selected = [arg for package in PACKAGES for arg in ("-p", package)]
env = dict(os.environ, CARGO_BUILD_JOBS="2")
commands = [
    ("fmt", ["cargo", "fmt", "--all", "--check"]),
    ("check", ["cargo", "check", *selected, "--all-targets", "--locked"]),
    ("clippy", ["cargo", "clippy", *selected, "--lib", "--no-deps", "--locked", "--", "-D", "warnings"]),
    ("tests", ["cargo", "test", *selected, "--locked"]),
    ("facade", ["cargo", "test", "-p", "litchi", "--features", "doc,docx,ppt,pptx,xls,xlsx,xlsb,odt", "--locked"]),
    ("rustdoc", ["cargo", "doc", *selected, "--no-deps", "--locked"]),
    ("harness", ["cargo", "test", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--locked"]),
    ("allocator", ["cargo", "test", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--locked", "--features", "allocator-metrics", "--bin", "docx_bounded_tail_append_compare", "--", "--test-threads=1"]),
    ("facade-polyglot", ["cargo", "test", "-p", "litchi", "--locked", "--no-default-features", "--features", "docx,odt", "--lib", "--tests", "--", "--test-threads=1"]),
    ("claims", ["python3", "tools/check_perf_claims.py", "--registry", "docs/performance/claim-registry-v1.json", "--repo-root", ".", "--evidence-root", ".", "--mode", "strict"]),
    ("claims-structural", ["python3", "tools/check_perf_claims.py", "--registry", "docs/performance/claim-registry-v1.json", "--repo-root", ".", "--mode", "structural"]),
    ("gate-tests", ["python3", "-m", "unittest", "discover", "-s", "tools", "-p", "test_perf_claims.py"]),
    ("report", ["python3", "tools/check_report_claim_classification.py", "--registry", "docs/performance/report-claim-classification-v1.json", "--repo-root", "."]),
    ("coverage", ["python3", "tools/validate_crud_coverage_index.py", "--index", "docs/performance/crud-coverage-index-v1.json", "--catalog", "docs/performance/results/perf-corpus-manifest-v2.json", "--selector-source", "tools/perf-baseline/src/lib.rs", "--checklist", "docs/CRUD_Scenario_Checklist.md", "--repo-root", "."]),
    ("non-iwork", ["python3", "tools/non_iwork_gate.py", "verify"]),
    ("boundaries", ["python3", "tools/check_crate_boundaries.py"]),
]
results = []
requested = set(os.environ.get("LITCHI_GATE_ONLY", "").split(",")) - {""}
for name, command in commands:
    if requested and name not in requested:
        continue
    start = time.time()
    head = subprocess.check_output(["git", "rev-parse", "HEAD"], text=True).strip()
    with (OUT / (name + ".log")).open("w") as log:
        log.write("HEAD " + head + "\n$ " + " ".join(command) + "\n")
        log.flush()
        gate_env = dict(env)
        if name == "rustdoc":
            gate_env["RUSTDOCFLAGS"] = "-D warnings"
        result = subprocess.run(command, stdout=log, stderr=subprocess.STDOUT, env=gate_env)
        log.write("\nexit " + str(result.returncode) + "\n")
    results.append(dict(name=name, command=command, head=head, exit_code=result.returncode, seconds=round(time.time()-start, 2)))
    (OUT / "results.json").write_text(json.dumps(results, indent=2) + "\n")
    print(name, result.returncode, results[-1]["seconds"], flush=True)
    if result.returncode:
        raise SystemExit(result.returncode)
