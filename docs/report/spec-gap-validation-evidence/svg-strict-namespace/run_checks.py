#!/usr/bin/env python3
"""Compile the retained Strict-namespace harness before running its probe."""

import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys


HERE = Path(__file__).resolve().parent
ROOT = next(path for path in HERE.parents if (path / "crates").is_dir())
RELATIVE = str(HERE.relative_to(ROOT))


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def source_paths():
    names = [
        "Cargo.lock",
        "Cargo.toml",
        "rust-toolchain.toml",
        "crates/litchi-opc/src/phys_pkg.rs",
        "crates/litchi-opc/src/constants.rs",
        "crates/litchi-pptx/src/presentation/source.rs",
        "crates/litchi-pptx/src/presentation/source/svg_lifecycle.rs",
        "crates/litchi-pptx/src/presentation/source/svg_lifecycle/owner.rs",
        "crates/litchi-pptx/tests/source_backed_svg_lifecycle.rs",
        "crates/litchi-drawingml/src/svg_blip.rs",
        "crates/litchi-ooxml-common/src/relationships/codec.rs",
        "3rdparty/specs/ECMA-376/ECMA-376-1_5th_edition_december_2016.zip",
        "3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip",
        "3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.26 http---schemas.microsoft.com-office-drawing-2016-SVG-main.md",
        "3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.24 http---schemas.microsoft.com-office-drawing-2016-SVG-main Schema.md",
        "3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx",
        "docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle/outputs/source.pptx",
    ]
    names.extend(
        str(path.relative_to(ROOT))
        for path in (HERE / "harness").rglob("*")
        if path.is_file()
    )
    names.extend(
        str((HERE / name).relative_to(ROOT))
        for name in (
            "README.md",
            "requirements.md",
            "run_checks.py",
            "run_probe.py",
            "validate_schema.py",
            "verify.py",
        )
    )
    paths = sorted({ROOT / name for name in names})
    missing = [path for path in paths if not path.is_file()]
    if missing:
        raise FileNotFoundError(", ".join(str(path) for path in missing))
    return paths


def snapshot():
    return {
        str(path.relative_to(ROOT)): digest(path)
        for path in source_paths()
    }


def tool(command):
    return subprocess.check_output(command, cwd=ROOT, text=True).strip()


def main():
    gates = HERE / "gates"
    raw = HERE / "raw-results"
    gates.mkdir(exist_ok=True)
    raw.mkdir(exist_ok=True)
    owned_target = HERE / "harness/target"
    if owned_target.exists():
        shutil.rmtree(owned_target)
    before = snapshot()
    (gates / "source-hashes.json").write_text(json.dumps(before, indent=2) + "\n")
    command = [
        "cargo",
        "check",
        "--locked",
        "--offline",
        "--manifest-path",
        f"{RELATIVE}/harness/Cargo.toml",
    ]
    environment = dict(os.environ)
    for name in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP"):
        environment.pop(name, None)
    environment["CARGO_TARGET_DIR"] = str(ROOT / "target")
    log = gates / "compile-first.log"
    with log.open("w") as sink:
        result = subprocess.run(
            command,
            cwd=ROOT,
            env=environment,
            stdout=sink,
            stderr=subprocess.STDOUT,
            check=False,
        )
    after = snapshot()
    receipt = {
        "passed": result.returncode == 0 and before == after,
        "source_unchanged": before == after,
        "command": command,
        "exit_code": result.returncode,
        "rustc": tool(["rustc", "-vV"]),
        "cargo": tool(["cargo", "--version"]),
        "source_files_checked": len(before),
        "source_hashes_sha256": digest(gates / "source-hashes.json"),
        "compile_log_sha256": digest(log),
    }
    (gates / "receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    if not receipt["passed"]:
        raise SystemExit(1)
    print(json.dumps(receipt, indent=2))


if __name__ == "__main__":
    main()
