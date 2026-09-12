#!/usr/bin/env python3
"""Replay the matched profile in clean checkouts; retain receipts in OUTPUT."""
import argparse
import gzip
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tempfile

BASE = "8e3ad310426a534c0bb17a789eb2c42a96e61310"
SOURCE = "crates/litchi-ods/src/advanced.rs"
HASHES = {
    "before": "0d1fad5bd849a0e34a1f60cdfec6dd62185bff2ac35673aa6148a63da3df3677",
    "after": "20eb69aa8402ac5c621efdf767a1069b89712f327e83d03d2f0a305a55e0822a",
}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("output", type=Path)
    parser.add_argument("--repeats", type=int, default=7)
    args = parser.parse_args()
    if args.repeats < 1:
        parser.error("repeats must be positive")
    output = args.output.resolve()
    output.mkdir(parents=True, exist_ok=False)
    evidence = Path(__file__).resolve().parent
    repo = Path(subprocess.check_output(
        ["git", "-C", str(evidence), "rev-parse", "--show-toplevel"], text=True
    ).strip())
    manifest = "docs/report/spec-gap-validation-evidence/ods-data-style-source-performance/harness"
    provenance = {"base": BASE, "repeats": args.repeats, "captures": {}}
    for label in ("before", "after"):
        with tempfile.TemporaryDirectory(prefix="ods-graph-preflight-replay-") as temporary:
            scratch = Path(temporary)
            checkout = scratch / "worktree"
            registered = False
            try:
                subprocess.run(["git", "-C", str(repo), "worktree", "add", "--detach", str(checkout), BASE], check=True)
                registered = True
                if label == "after":
                    subprocess.run(["git", "apply", "-"], cwd=checkout,
                                   input=gzip.decompress((evidence / "candidate.patch.gz").read_bytes()), check=True)
                source_hash = hashlib.sha256((checkout / SOURCE).read_bytes()).hexdigest()
                if source_hash != HASHES[label]:
                    raise RuntimeError("source hash mismatch")
                shutil.copytree(evidence / "harness", checkout / manifest, dirs_exist_ok=True)
                destination = output / label
                destination.mkdir()
                env = os.environ.copy()
                env.pop("RUSTFLAGS", None)
                env.pop("CARGO_ENCODED_RUSTFLAGS", None)
                env.update(CARGO_TARGET_DIR=str(scratch / "target"),
                           ODS_PROFILE_OUTPUT_DIR=str(destination),
                           ODS_PROFILE_REPEATS=str(args.repeats))
                with (output / (label + ".jsonl")).open("w") as stdout, (output / (label + "-build.log")).open("w") as stderr:
                    subprocess.run(["cargo", "run", "--manifest-path", str(checkout / manifest / "Cargo.toml"),
                                    "--locked", "--offline", "--release"], cwd=checkout, env=env,
                                   stdout=stdout, stderr=stderr, check=True)
                binary = scratch / "target/release/ods-data-style-source-performance"
                provenance["captures"][label] = {
                    "source_sha256": source_hash,
                    "binary_sha256": hashlib.sha256(binary.read_bytes()).hexdigest(),
                }
            finally:
                if registered:
                    subprocess.run(["git", "-C", str(repo), "worktree", "remove", "--force", str(checkout)], check=True)
    (output / "provenance.json").write_text(json.dumps(provenance, indent=2) + "\n")


if __name__ == "__main__":
    main()
