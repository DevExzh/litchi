#!/usr/bin/env python3
"""Run the scoped OPC/DOCX verification commands serially under retained gates."""
import argparse
import subprocess
import sys

from support import REPO, ROOT, TEMP


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("attempt")
    args = parser.parse_args()
    if not args.attempt or not all(c.isalnum() or c in "-_" for c in args.attempt):
        raise ValueError("invalid attempt")
    temporary = TEMP / "test-tmp"
    temporary.mkdir(parents=True, exist_ok=True)
    commands = [
        ("format", ["cargo", "fmt", "-p", "litchi-opc", "--", "--check"]),
        ("tests-all-features", ["cargo", "test", "-p", "litchi-opc", "-p", "litchi-docx", "--all-features"]),
        ("tests-no-default", ["cargo", "test", "-p", "litchi-opc", "-p", "litchi-docx", "--no-default-features"]),
        ("clippy", ["cargo", "clippy", "-p", "litchi-opc", "-p", "litchi-docx", "--all-features", "--all-targets", "--", "-D", "warnings"]),
        ("rustdoc", ["env", "RUSTDOCFLAGS=-D warnings", "cargo", "doc", "-p", "litchi-opc", "-p", "litchi-docx", "--all-features", "--no-deps"]),
        ("boundaries", ["python3", "-B", "tools/check_crate_boundaries.py"]),
        ("bench-tests-serial", ["cargo", "test", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", "docx_replayable_tail_append", "--features", "allocator-metrics", "--", "--test-threads=1"]),
        ("python", ["python3", "-B", "-m", "unittest", "discover", "-s", str(ROOT), "-p", "test_*.py", "-v"]),
    ]
    for label, command in commands:
        print(f"starting {label}-{args.attempt}", flush=True)
        subprocess.run([sys.executable, "-B", str(ROOT / "gate.py"),
                        f"{label}-{args.attempt}", "env", f"TMPDIR={temporary}", *command],
                       cwd=REPO, check=True)


if __name__ == "__main__":
    main()
