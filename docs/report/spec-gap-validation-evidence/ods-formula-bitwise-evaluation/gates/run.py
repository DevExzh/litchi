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
# Include the complete ODS implementation and tests, plus shared execution inputs.
FILES = sorted(str(path.relative_to(ROOT)) for folder in ["src", "tests"]
    for path in (ROOT / "crates/litchi-ods" / folder).rglob("*.rs"))
FILES += ["Cargo.toml", "Cargo.lock", "crates/litchi-ods/Cargo.toml",
    "crates/litchi-core/src/execution.rs"]
FILES += sorted(str(path.relative_to(ROOT)) for path in (ROOT / "crates/litchi-core/src/budget").rglob("*.rs"))
if (ROOT / "crates/litchi-core/src/budget.rs").is_file():
    FILES.append("crates/litchi-core/src/budget.rs")
def hashes():
    return {name: hashlib.sha256((ROOT / name).read_bytes()).hexdigest() for name in FILES}

def main():
    Path("/home/zhuhe/code/litchi-bitwise-tmp").mkdir(exist_ok=True)
    env = dict(os.environ, TMPDIR="/home/zhuhe/code/litchi-bitwise-tmp", CARGO_TARGET_DIR="/home/zhuhe/code/litchi-bitwise-target", RUSTDOCFLAGS="-D warnings", CARGO_PROFILE_DEV_DEBUG="0", CARGO_PROFILE_TEST_DEBUG="0", CARGO_INCREMENTAL="0")
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
    receipt["environment"] = {key: env[key] for key in ["CARGO_TARGET_DIR", "TMPDIR", "RUSTDOCFLAGS", "CARGO_PROFILE_DEV_DEBUG", "CARGO_PROFILE_TEST_DEBUG", "CARGO_INCREMENTAL"]}
    receipt["source_after"] = hashes()
    receipt["sources_unchanged"] = receipt["source_before"] == receipt["source_after"]
    (OUT / "results.json").write_text(json.dumps(receipt, indent=2) + "\n")
    raise SystemExit(0 if receipt["sources_unchanged"] and all(x["status"] == 0 for x in receipt["commands"]) else 1)


if __name__ == "__main__":
    main()
