#!/usr/bin/env python3
"""Retain XLS retry admission and CFB dependency integration commands, source bindings, and exit statuses."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

root = Path(__file__).resolve().parents[4]
output = Path(os.environ.get("LITCHI_GATE_OUTPUT", str(Path(__file__).parent / "integration")))
output.mkdir(parents=True, exist_ok=True)
env = dict(os.environ, CARGO_BUILD_JOBS="2")
commands = [
    ("fmt", ["cargo", "fmt", "--all", "--check"]),
    ("check", ["cargo", "check", "-p", "litchi-cfb", "-p", "litchi-xls", "--all-features", "--all-targets", "--locked"]),
    ("clippy", ["cargo", "clippy", "-p", "litchi-cfb", "-p", "litchi-xls", "--all-features", "--lib", "--no-deps", "--locked", "--", "-D", "warnings"]),
    ("tests", ["cargo", "test", "-p", "litchi-cfb", "-p", "litchi-xls", "--all-features", "--locked", "--", "--test-threads=1"]),
    ("facade", ["cargo", "test", "-p", "litchi", "--locked", "--no-default-features", "--features", "xls,ods", "--lib", "--tests", "--", "--test-threads=1"]),
    ("rustdoc", ["cargo", "doc", "-p", "litchi-cfb", "-p", "litchi-xls", "--all-features", "--no-deps", "--locked"]),
]
selected = set(os.environ.get("LITCHI_GATE_ONLY", "").split(",")) - {""}
results = []
for name, command in commands:
    if selected and name not in selected:
        continue
    head = subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=root, text=True).strip()
    paths = subprocess.check_output(["git", "ls-files", "crates/litchi-cfb", "crates/litchi-xls"], cwd=root, text=True).splitlines()
    paths += [str(path.relative_to(root)) for owner in ["litchi-cfb", "litchi-xls"] for path in (root / "crates" / owner).rglob("*.rs")]
    hashes = {path: hashlib.sha256((root / path).read_bytes()).hexdigest() for path in sorted(set(paths)) if (root / path).is_file()}
    gate_env = dict(env)
    if name == "rustdoc":
        gate_env["RUSTDOCFLAGS"] = "-D warnings"
    start = time.monotonic()
    with (output / (name + ".log")).open("w") as log:
        log.write("HEAD " + head + "\n$ " + " ".join(command) + "\n")
        log.flush()
        result = subprocess.run(command, cwd=root, env=gate_env, stdout=log, stderr=subprocess.STDOUT)
        log.write("\nexit " + str(result.returncode) + "\n")
    results.append(dict(name=name, command=command, head=head, source_sha256=hashes, exit_code=result.returncode, seconds=round(time.monotonic() - start, 2)))
    (output / "results.json").write_text(json.dumps(results, indent=2) + "\n")
    print(name, result.returncode, results[-1]["seconds"], flush=True)
    if result.returncode:
        raise SystemExit(result.returncode)
