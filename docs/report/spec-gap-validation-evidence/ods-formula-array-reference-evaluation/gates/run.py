#!/usr/bin/env python3
"""Run scoped ODS checks and retain exact source provenance."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

ROOT = Path(__file__).resolve().parents[5]
OUT = Path(__file__).resolve().parent
# Local dependencies of litchi-ods, including proc-macro implementation inputs.
CRATES = ("litchi-ods", "litchi-core", "litchi-odf-common", "soapberry-zip",
          "xml-minifier", "xml-minifier-macros")


def hashes(root):
    files = {"Cargo.toml", "Cargo.lock", ".cargo/config.toml"}
    for crate in CRATES:
        crate_root = root / "crates" / crate
        files.add(str((crate_root / "Cargo.toml").relative_to(root)))
        for folder in ("src", "tests", "examples"):
            files.update(str(path.relative_to(root))
                         for path in (crate_root / folder).rglob("*") if path.is_file())
        if (crate_root / "build.rs").is_file():
            files.add(str((crate_root / "build.rs").relative_to(root)))
    return {name: hashlib.sha256((root / name).read_bytes()).hexdigest()
            for name in sorted(files)}

def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--workspace", type=Path, required=True,
                        help="isolated workspace containing the candidate sources")
    args = parser.parse_args()
    workspace = args.workspace.resolve(strict=True)
    if not (workspace / "crates/litchi-ods/Cargo.toml").is_file():
        parser.error("workspace does not contain litchi-ods")
    Path("/home/zhuhe/code/litchi-array-tmp").mkdir(exist_ok=True)
    env = dict(os.environ, TMPDIR="/home/zhuhe/code/litchi-array-tmp", CARGO_TARGET_DIR="/home/zhuhe/code/litchi-array-target", RUSTDOCFLAGS="-D warnings", CARGO_PROFILE_DEV_DEBUG="0", CARGO_PROFILE_TEST_DEBUG="0", CARGO_INCREMENTAL="0")
    commands = [
        ("test", ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--all-targets"]),
        ("clippy", ["cargo", "clippy", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--all-targets", "--", "-D", "warnings"]),
        ("doc", ["cargo", "doc", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--no-deps"]),
        ("doctest", ["cargo", "test", "--locked", "--offline", "-p", "litchi-ods", "--all-features", "--doc"]),
        ("fmt", ["cargo", "fmt", "-p", "litchi-ods", "--", "--check"]),
    ]
    receipt = {"head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=workspace, text=True).strip(), "workspace": str(workspace), "source_before": hashes(workspace), "commands": []}
    receipt["toolchain"] = {
        "rustc": subprocess.check_output(["rustc", "-Vv"], cwd=workspace, text=True),
        "cargo": subprocess.check_output(["cargo", "-V"], cwd=workspace, text=True),
    }
    receipt["runner_sha256"] = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    receipt["canonical_source"] = hashes(ROOT)
    receipt["candidate_matches_canonical"] = receipt["canonical_source"] == receipt["source_before"]
    if not receipt["candidate_matches_canonical"]:
        raise SystemExit("isolated candidate source differs from canonical source; sync before gating")
    for name, command in commands:
        start = time.time()
        with (OUT / (name + ".log")).open("w") as log:
            status = subprocess.run(command, cwd=workspace, env=env, stdout=log, stderr=subprocess.STDOUT).returncode
        receipt["commands"].append({"name": name, "command": command, "status": status, "seconds": time.time() - start})
        print(name, status, flush=True)
    receipt["environment"] = {key: env[key] for key in ["CARGO_TARGET_DIR", "TMPDIR", "RUSTDOCFLAGS", "CARGO_PROFILE_DEV_DEBUG", "CARGO_PROFILE_TEST_DEBUG", "CARGO_INCREMENTAL"]}
    receipt["source_after"] = hashes(workspace)
    receipt["canonical_source_after"] = hashes(ROOT)
    receipt["sources_unchanged"] = (
        receipt["source_before"] == receipt["source_after"]
        and receipt["canonical_source"] == receipt["canonical_source_after"]
    )
    (OUT / "results.json").write_text(json.dumps(receipt, indent=2) + "\n")
    raise SystemExit(0 if receipt["sources_unchanged"] and all(x["status"] == 0 for x in receipt["commands"]) else 1)


if __name__ == "__main__":
    main()
