"""Serial quality gates for the matched CFB collector experiment."""

from __future__ import annotations

import argparse

from run import TARGET, configure, run


COMMANDS = [
    ("fmt", ["cargo", "fmt", "--all", "--check"]),
    ("harness-fmt", ["cargo", "fmt", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--all", "--check"]),
    ("cfb-tests", ["cargo", "test", "--locked", "-p", "litchi-cfb", "--all-features"]),
    ("cfb-no-default-tests", ["cargo", "test", "--locked", "-p", "litchi-cfb", "--no-default-features"]),
    ("xls-tests", ["cargo", "test", "--locked", "-p", "litchi-xls", "--all-features"]),
    ("doc-tests", ["cargo", "test", "--locked", "-p", "litchi-doc", "--all-features"]),
    ("ppt-tests", ["cargo", "test", "--locked", "-p", "litchi-ppt", "--all-features"]),
    ("workspace-check", ["cargo", "check", "--locked", "--workspace", "--all-features"]),
    ("ole-clippy", ["cargo", "clippy", "--locked", "-p", "litchi-cfb", "-p", "litchi-xls", "-p", "litchi-doc", "-p", "litchi-ppt", "--all-features", "--lib", "--", "-D", "warnings"]),
    ("harness-clippy", ["cargo", "clippy", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics", "--lib", "--", "-D", "warnings"]),
    ("ole-rustdoc", ["cargo", "doc", "--locked", "-p", "litchi-cfb", "-p", "litchi-xls", "-p", "litchi-doc", "-p", "litchi-ppt", "--all-features", "--no-deps"]),
    ("harness-rustdoc", ["cargo", "doc", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--features", "allocator-metrics", "--lib", "--no-deps"]),
    ("boundaries", ["python3", "-B", "tools/check_crate_boundaries.py"]),
    ("claims", ["python3", "-B", "tools/check_perf_claims.py", "--registry", "docs/performance/claim-registry-v1.json", "--repo-root", ".", "--evidence-root", ".", "--mode", "strict"]),
]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--stage", choices=("candidate", "final"), required=True)
    args = parser.parse_args()
    configure(args.stage)
    (TARGET / "test-tmp").mkdir(parents=True, exist_ok=True)
    prefix = [
        "env", "TMPDIR=" + str(TARGET / "test-tmp"),
        "CARGO_TARGET_DIR=" + str(TARGET), "CARGO_BUILD_JOBS=2",
        "CARGO_INCREMENTAL=0", "RUSTDOCFLAGS=-D warnings",
    ]
    for name, command in COMMANDS:
        run("check-" + name, prefix + command)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
