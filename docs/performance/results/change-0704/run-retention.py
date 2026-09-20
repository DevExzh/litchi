#!/usr/bin/env python3
"""Run the candidate-only retention observer on its two bounded fixtures."""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
import sys
import time
from pathlib import Path


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
BINARY = ROOT.parent / "litchi-0704-bin" / "retention-observer"
BUILD_RECEIPT = PACKET / "retention-build.json"
CANDIDATE_CENSUS = PACKET / "source-census-candidate.json"
OUTPUT_DIR = PACKET / "retention-probe" / "results"
PRODUCTION_ROOTS = (
    ROOT / "crates" / "litchi-pptx",
    ROOT / "crates" / "litchi-ooxml-common",
    ROOT / "crates" / "litchi-opc",
)
MANIFEST = PACKET / "retention-probe" / "Cargo.toml"
LOCK = PACKET / "retention-probe" / "Cargo.lock"
BUILD_INPUTS = (
    ROOT / "Cargo.toml",
    ROOT / "Cargo.lock",
    ROOT / "crates" / "litchi-pptx" / "Cargo.toml",
    ROOT / "crates" / "litchi-ooxml-common" / "Cargo.toml",
    ROOT / "crates" / "litchi-opc" / "Cargo.toml",
    MANIFEST,
    LOCK,
    CANDIDATE_CENSUS,
)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def file_map(roots: tuple[Path, ...], relative_to: Path) -> dict[str, str]:
    result: dict[str, str] = {}
    for root in roots:
        if not root.is_dir():
            raise AssertionError(f"missing census root: {root}")
        for path in sorted(root.rglob("*")):
            if path.is_file() and not path.is_symlink():
                result[str(path.relative_to(relative_to))] = sha(path)
    return dict(sorted(result.items()))


def probe_map() -> dict[str, str]:
    root = PACKET / "retention-probe"
    return {
        name: digest
        for name, digest in file_map((root,), PACKET).items()
        if not name.startswith("retention-probe/results/")
    }


def input_map() -> dict[str, str]:
    return {str(path.relative_to(ROOT)): sha(path) for path in BUILD_INPUTS}


def rust_map(mapping: dict[str, str]) -> dict[str, str]:
    return {name: digest for name, digest in mapping.items() if name.endswith(".rs")}


def other_map(mapping: dict[str, str]) -> dict[str, str]:
    return {name: digest for name, digest in mapping.items() if not name.endswith(".rs")}


def parse_repeat() -> int:
    if len(sys.argv) > 2:
        raise SystemExit("usage: run-retention.py [repeats<=2]")
    value = int(sys.argv[1]) if len(sys.argv) == 2 else 2
    if value < 1 or value > 2:
        raise SystemExit("repeat count must be between 1 and 2")
    return value


def verify_bindings(build: dict[str, object]) -> None:
    production = file_map(PRODUCTION_ROOTS, ROOT)
    probe = probe_map()
    rust = rust_map(production)
    other = other_map(production)
    candidate = json.loads(CANDIDATE_CENSUS.read_text())["source_sha256"]
    if production != build["production_source_sha256"]:
        raise AssertionError("production source census differs from build receipt")
    if rust != build["production_rust_source_sha256"]:
        raise AssertionError("Rust production census differs from build receipt")
    if other != build["production_other_sha256"]:
        raise AssertionError("non-Rust production census differs from build receipt")
    if candidate != rust:
        raise AssertionError("source-census-candidate.json differs from Rust production census")
    if probe != build["probe_sha256"]:
        raise AssertionError("retention probe census differs from build receipt")
    if input_map() != build["build_inputs_sha256"]:
        raise AssertionError("build input census differs from build receipt")
    if not BINARY.is_file() or sha(BINARY) != build["binary_sha256"]:
        raise AssertionError("retention observer binary differs from build receipt")


def check_observer_output(output: str, source: str, repeats: int) -> None:
    lines = output.splitlines()
    required = {
        "probe\t0704-retention-observer",
        f"source\t{source}",
        f"repeats\t{repeats}",
        "parity\tdefault_vs_off\ttrue",
        "parity\tdefault_vs_tiny\ttrue",
        "run_complete\ttrue",
    }
    missing = sorted(required.difference(lines))
    if missing:
        raise AssertionError(f"observer omitted required receipts: {missing}")
    records = [line for line in lines if line.startswith("record\t")]
    if len(records) != repeats * 3:
        raise AssertionError(
            f"observer emitted {len(records)} records, expected {repeats * 3}"
        )
    if source == "generated:12x8":
        expected_target = "target1\tslide:0\tshape:0"
    else:
        expected_target = "target1\tslide:1\tshape:1"
    if expected_target not in lines:
        raise AssertionError(f"observer target receipt missing: {expected_target}")


def main() -> None:
    if not BUILD_RECEIPT.is_file():
        raise SystemExit(f"run build-retention.py first: {BUILD_RECEIPT}")
    build = json.loads(BUILD_RECEIPT.read_text())
    verify_bindings(build)
    repeats = parse_repeat()
    real = ROOT / "test-data" / "libreoffice-core" / "sd" / "qa" / "unit" / "data" / "pptx" / "slide-section-test.pptx"
    if not real.is_file():
        raise AssertionError(f"missing real fixture: {real}")
    cases = [
        ("real", str(real), sha(real)),
        ("generated", "generated:12x8", None),
    ]
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, object]] = []
    for label, source, source_sha256 in cases:
        output_path = OUTPUT_DIR / f"{label}.tsv"
        stderr_path = OUTPUT_DIR / f"{label}.stderr"
        command = [str(BINARY), source, str(repeats)]
        started = time.monotonic()
        with output_path.open("w") as stdout, stderr_path.open("w") as stderr:
            result = subprocess.run(
                command,
                cwd=ROOT,
                env={**os.environ},
                stdout=stdout,
                stderr=stderr,
            )
        output = output_path.read_text()
        if result.returncode != 0:
            raise SystemExit(
                f"{label} observer failed with {result.returncode}; see {stderr_path}"
            )
        check_observer_output(output, source, repeats)
        records.append(
            {
                "label": label,
                "source": source,
                "source_sha256": source_sha256,
                "command": command,
                "cwd": str(ROOT),
                "exit_code": result.returncode,
                "seconds": time.monotonic() - started,
                "output": str(output_path.relative_to(PACKET)),
                "output_sha256": sha(output_path),
                "stderr": str(stderr_path.relative_to(PACKET)),
                "stderr_sha256": sha(stderr_path),
                "binary": str(BINARY),
                "binary_sha256": sha(BINARY),
                "production_source_hash": build["production_source_hash"],
                "probe_source_hash": build["probe_source_hash"],
                "source_hash": build["source_hash"],
            }
        )
        print(label, "observer", result.returncode, flush=True)
    receipt = {
        "probe": "0704-retention-observer",
        "repeats": repeats,
        "binary": str(BINARY),
        "binary_sha256": sha(BINARY),
        "production_source_hash": build["production_source_hash"],
        "probe_source_hash": build["probe_source_hash"],
        "source_hash": build["source_hash"],
        "cases": records,
        "performance_claim": "none",
    }
    (PACKET / "retention-runs.json").write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps(receipt, indent=2))


if __name__ == "__main__":
    main()
