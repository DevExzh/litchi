#!/usr/bin/env python3
"""Run scoped ODS checks and retain exact source provenance."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

ROOT = Path(__file__).resolve().parents[5]
OUT = Path(__file__).resolve().parent
FILES = [
    "crates/litchi-ods/src/codec/formula.rs",
    "crates/litchi-ods/src/codec/formula/functions.rs",
    "crates/litchi-ods/src/codec/formula/reference.rs",
    "crates/litchi-ods/src/codec/formula/reference/iri.rs",
    "crates/litchi-ods/src/model/cell.rs",
    "crates/litchi-ods/tests/ods_formula_references.rs",
    "crates/litchi-ods/tests/ods_formula_functions.rs",
    "crates/litchi-ods/tests/ods_formula_tokenizer_regression.rs",
    "crates/litchi-ods/tests/ods_formula_scan_regression.rs",
    "crates/litchi-ods/tests/ods_formula_string_regression.rs",
]
FILES += ["crates/litchi-ods/src/codec/formula/expression.rs", "crates/litchi-ods/tests/ods_formula_expressions.rs"]
FILES += sorted(str(path.relative_to(ROOT)) for path in (ROOT / "crates/litchi-ods/src/codec/formula/expression").rglob("*.rs"))
def hashes():
    return {name: hashlib.sha256((ROOT / name).read_bytes()).hexdigest() for name in FILES}

Path("/var/tmp/ods-expression-tmp").mkdir(exist_ok=True)
env = dict(os.environ, TMPDIR="/var/tmp/ods-expression-tmp", CARGO_TARGET_DIR="/var/tmp/ods-expression-target", RUSTDOCFLAGS="-D warnings")
commands = [
    ("test", ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--all-targets"]),
    ("clippy", ["cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--all-targets", "--", "-D", "warnings"]),
    ("doc", ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--no-deps"]),
    ("doctest", ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--doc"]),
    ("fmt", ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"]),
]
receipt = {"head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(), "source_before": hashes(), "commands": []}
for name, command in commands:
    start = time.time()
    with (OUT / (name + ".log")).open("w") as log:
        status = subprocess.run(command, cwd=ROOT, env=env, stdout=log, stderr=subprocess.STDOUT).returncode
    receipt["commands"].append({"name": name, "command": command, "status": status, "seconds": time.time() - start})
    print(name, status, flush=True)
receipt["environment"] = {key: env[key] for key in ["CARGO_TARGET_DIR", "TMPDIR", "RUSTDOCFLAGS"]}
receipt["source_after"] = hashes()
receipt["sources_unchanged"] = receipt["source_before"] == receipt["source_after"]
(OUT / "results.json").write_text(json.dumps(receipt, indent=2) + "\n")
raise SystemExit(0 if receipt["sources_unchanged"] and all(x["status"] == 0 for x in receipt["commands"]) else 1)
