#!/usr/bin/env python3
"""Run `cargo check` over the workspace's per-format feature combinations.

The facade (`litchi`) is checked with `--no-default-features` alone, with each
format feature alone, and with each format aggregate (`legacy`, `ooxml`,
`odf`, ...). The format features are derived from `crates/litchi/Cargo.toml`
the way the facade's build script derives `cfg(litchi_unified)`: a feature
that turns on a `litchi-ops` format with a plain `"litchi-ops/<format>"` item
(record 0334). `litchi-cli` and `litchi-py` are checked the same way when
their manifests declare per-format features; today both pin every format, so
each is checked once with its own feature set. `--ops` adds `litchi-ops`
alone, with each of its format features and with each internal umbrella
(`odf-objects`, record 0358a).

Commands run one at a time, in a fixed order, through whatever `cargo` is
first on PATH (the coordinator's memory-capped shim on this host). The build
parallelism comes from the environment (`CARGO_BUILD_JOBS`); the script never
starts two cargo processes at once.

Usage:
    python3 tools/check_feature_matrix.py [--ops] [--all-targets] [--clippy]
        [--only GROUP[,GROUP...]] [--fail-fast] [--dry-run] [--list]

`--clippy` runs `cargo clippy ... -- -D warnings` instead of `cargo check`,
so code that one feature subset leaves unused fails too.

Exit status: 0 when every command succeeds, 1 when any fails, 2 on a usage or
manifest error.
"""

from __future__ import annotations

import argparse
import os
import shlex
import subprocess
import sys
import time
import tomllib
from dataclasses import dataclass
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
FACADE = ROOT / "crates" / "litchi" / "Cargo.toml"
OPS = ROOT / "crates" / "litchi-ops" / "Cargo.toml"
# litchi-ops features that are not formats but are checked alone with --ops.
OPS_UMBRELLAS = ("odf-objects",)
CLI = ROOT / "crates" / "litchi-cli" / "Cargo.toml"
PY = ROOT / "crates" / "pyo3-litchi" / "Cargo.toml"

# Facade aggregates made only of formats (and their substrates).
FACADE_AGGREGATES = ("legacy", "ooxml", "odf", "iwork", "word", "slides", "spreadsheets")

# Number of claimed formats (docs/FEATURE_MATRIX.md); a mismatch means the
# derivation and the manifest disagree.
CLAIMED_FORMATS = 21


@dataclass(frozen=True)
class Step:
    group: str
    package: str
    features: tuple[str, ...]
    no_default: bool

    def command(self, all_targets: bool, clippy: bool = False) -> list[str]:
        command = ["cargo", "clippy" if clippy else "check", "--locked", "-p", self.package]
        if self.no_default:
            command.append("--no-default-features")
        if self.features:
            command += ["--features", ",".join(self.features)]
        command.append("--all-targets" if all_targets else "--lib")
        if clippy:
            command += ["--", "-D", "warnings"]
        return command

    def label(self) -> str:
        features = ",".join(self.features) if self.features else "(none)"
        default = " --no-default-features" if self.no_default else ""
        return f"{self.package}{default} [{features}]"


def manifest_features(path: Path) -> dict[str, list[str]]:
    with path.open("rb") as handle:
        return tomllib.load(handle).get("features", {})


def format_features(features: dict[str, list[str]], dependency: str = "litchi-ops") -> list[str]:
    """Features that turn on a format of `dependency` with a plain item."""
    prefix = f"{dependency}/"
    return [
        name
        for name, items in features.items()
        if any(item.startswith(prefix) and len(item) > len(prefix) for item in items)
    ]


def ops_format_features(features: dict[str, list[str]]) -> list[str]:
    """litchi-ops features that are formats: the facade's format names."""
    facade = set(format_features(manifest_features(FACADE)))
    return [name for name in features if name in facade]


def plan(include_ops: bool) -> list[Step]:
    facade = manifest_features(FACADE)
    formats = format_features(facade)
    if len(formats) != CLAIMED_FORMATS:
        raise SystemExit(
            f"error: derived {len(formats)} facade format features, expected "
            f"{CLAIMED_FORMATS}: {formats}"
        )
    steps = [Step("facade", "litchi", (), True)]
    steps += [Step("facade", "litchi", (name,), True) for name in formats]
    missing = [name for name in FACADE_AGGREGATES if name not in facade]
    if missing:
        raise SystemExit(f"error: facade aggregates missing from the manifest: {missing}")
    steps += [Step("facade-aggregate", "litchi", (name,), True) for name in FACADE_AGGREGATES]

    for group, path, package in (("cli", CLI, "litchi-cli"), ("py", PY, "litchi-py")):
        own = format_features(manifest_features(path), "litchi")
        own += format_features(manifest_features(path), "litchi-ops")
        if own:
            steps += [Step(group, package, (name,), True) for name in sorted(set(own))]
        else:
            steps.append(Step(group, package, (), False))

    if include_ops:
        ops_features = manifest_features(OPS)
        ops = ops_format_features(ops_features)
        steps.append(Step("ops", "litchi-ops", (), True))
        steps += [Step("ops", "litchi-ops", (name,), True) for name in ops]
        # Internal umbrellas must build on their own too (record 0358a).
        steps += [Step("ops", "litchi-ops", (name,), True)
                  for name in OPS_UMBRELLAS if name in ops_features]
    return steps


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--ops", action="store_true", help="also check litchi-ops per format")
    parser.add_argument("--all-targets", action="store_true", help="check tests, examples and benches too")
    parser.add_argument("--only", help="comma-separated groups: facade, facade-aggregate, cli, py, ops")
    parser.add_argument(
        "--clippy", action="store_true", help="run warning-denied clippy instead of check"
    )
    parser.add_argument("--fail-fast", action="store_true", help="stop at the first failure")
    parser.add_argument("--dry-run", action="store_true", help="print the commands without running them")
    parser.add_argument("--list", action="store_true", help="print the derived facade format features")
    args = parser.parse_args()

    if args.list:
        print("\n".join(format_features(manifest_features(FACADE))))
        return 0

    steps = plan(args.ops or (args.only is not None and "ops" in args.only.split(",")))
    if args.only:
        groups = set(args.only.split(","))
        unknown = groups - {"facade", "facade-aggregate", "cli", "py", "ops"}
        if unknown:
            parser.error(f"unknown groups: {sorted(unknown)}")
        steps = [step for step in steps if step.group in groups]

    jobs = os.environ.get("CARGO_BUILD_JOBS", "(cargo default)")
    print(f"feature matrix: {len(steps)} cargo {'clippy' if args.clippy else 'check'} commands, CARGO_BUILD_JOBS={jobs}", flush=True)
    failures: list[Step] = []
    started = time.monotonic()
    for index, step in enumerate(steps, 1):
        command = step.command(args.all_targets, args.clippy)
        if args.dry_run:
            print(shlex.join(command))
            continue
        begin = time.monotonic()
        result = subprocess.run(
            command, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True
        )
        elapsed = time.monotonic() - begin
        status = "ok" if result.returncode == 0 else f"FAIL (exit {result.returncode})"
        print(f"[{index:2}/{len(steps)}] {status:<14} {elapsed:7.1f}s  {step.label()}", flush=True)
        if result.returncode != 0:
            failures.append(step)
            print(result.stdout[-6000:], flush=True)
            if args.fail_fast:
                break
    if args.dry_run:
        return 0
    total = time.monotonic() - started
    print(f"feature matrix: {len(steps) - len(failures)} ok, {len(failures)} failed, {total:.1f}s")
    for step in failures:
        print(f"  failed: {shlex.join(step.command(args.all_targets, args.clippy))}")
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
