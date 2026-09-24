#!/usr/bin/env python3
"""Record the immutable inputs and tool identity for an XLSB profile run."""

from __future__ import annotations

import hashlib
import json
import os
import platform
import subprocess
import sys
from pathlib import Path


ROOT = Path(__file__).resolve().parents[5]
FIXTURE = ROOT / "test-data/poi/test-data/spreadsheet/testVarious.xlsb"

# Keep this list limited to the XLSB projection change, its harness, and the
# locked build inputs. The repository has other concurrent worktrees/changes.
PATHS = [
    "Cargo.lock",
    "tools/perf-baseline/Cargo.lock",
    "tools/perf-baseline/Cargo.toml",
    "rust-toolchain.toml",
    "crates/litchi-xlsb/src/calculation_chain/workbook.rs",
    "crates/litchi-xlsb/src/cell_values/drawing_transfer.rs",
    "crates/litchi-xlsb/src/cell_values/root.rs",
    "crates/litchi-xlsb/src/cell_values/workbook.rs",
    "crates/litchi-xlsb/src/cell_watches/workbook.rs",
    "crates/litchi-xlsb/src/sparkline/workbook.rs",
    "crates/litchi-xlsb/src/workbook/mod.rs",
    "crates/litchi-xlsb/src/workbook/model.rs",
    "crates/litchi-xlsb/src/workbook/package.rs",
    "crates/litchi-xlsb/src/workbook/tests.rs",
    "crates/litchi-xlsb/tests/drawing_skip_projection.rs",
    "tools/perf-baseline/src/bin/xlsb_crud.rs",
    "tools/perf-baseline/README.md",
]


def run(*args: str) -> str:
    return subprocess.check_output(args, cwd=ROOT, text=True).strip()


def digest(path: Path) -> str:
    hasher = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1024 * 1024), b""):
            hasher.update(block)
    return hasher.hexdigest()


def host_identity() -> dict[str, object]:
    cpu_model = None
    cpuinfo = Path("/proc/cpuinfo")
    if cpuinfo.is_file():
        for line in cpuinfo.read_text(errors="replace").splitlines():
            if line.lower().startswith("model name") and ":" in line:
                cpu_model = line.split(":", 1)[1].strip()
                break
    memory_bytes = None
    meminfo = Path("/proc/meminfo")
    if meminfo.is_file():
        for line in meminfo.read_text(errors="replace").splitlines():
            if line.startswith("MemTotal:"):
                memory_bytes = int(line.split()[1]) * 1024
                break
    return {
        "os": platform.platform(),
        "kernel": platform.release(),
        "machine": platform.machine(),
        "cpu_model": cpu_model,
        "logical_cpus": os.cpu_count(),
        "memory_bytes": memory_bytes,
    }


def record(phase: str, binary: Path | None) -> dict[str, object]:
    files: dict[str, dict[str, object]] = {}
    for relative in PATHS:
        path = ROOT / relative
        if not path.is_file():
            raise SystemExit(f"required provenance path is missing: {relative}")
        files[relative] = {
            "bytes": path.stat().st_size,
            "sha256": digest(path),
        }
    fixture = {
        "path": os.fspath(FIXTURE.relative_to(ROOT)),
        "bytes": FIXTURE.stat().st_size,
        "sha256": digest(FIXTURE),
    }
    candidate_status = subprocess.check_output(
        ["git", "status", "--short", "--untracked-files=all", "--", *PATHS],
        cwd=ROOT,
        text=True,
    ).splitlines()
    result: dict[str, object] = {
        "schema": "litchi-xlsb-drawing-projection-provenance-v1",
        "phase": phase,
        "repository": os.fspath(ROOT),
        "git_head": run("git", "rev-parse", "HEAD"),
        "candidate_status_short": candidate_status,
        "host": host_identity(),
        "rustc": run("rustc", "--version", "--verbose"),
        "cargo": run("cargo", "--version"),
        "rustup_toolchain": os.environ.get("RUSTUP_TOOLCHAIN"),
        "cargo_target_dir": os.environ.get("CARGO_TARGET_DIR"),
        "cargo_incremental": os.environ.get("CARGO_INCREMENTAL"),
        "fixture": fixture,
        "files": files,
    }
    if binary is not None:
        if not binary.is_file():
            raise SystemExit(f"binary is missing: {binary}")
        result["binary"] = {
            "path": os.fspath(binary),
            "bytes": binary.stat().st_size,
            "sha256": digest(binary),
        }
    return result


def main() -> None:
    if len(sys.argv) not in (2, 3):
        raise SystemExit("usage: record-provenance.py <phase> [binary]")
    phase = sys.argv[1]
    binary = Path(sys.argv[2]).resolve() if len(sys.argv) == 3 else None
    print(json.dumps(record(phase, binary), indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
