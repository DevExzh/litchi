#!/usr/bin/env python3
"""Run the downstream example and offline schema checks with source receipts."""

import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys

HERE = Path(__file__).resolve().parent
ROOT = next(path for path in HERE.parents if (path / "crates").is_dir())


def digest(path):
    result = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1 << 20), b""):
            result.update(block)
    return result.hexdigest()


def main():
    manifest = HERE / "gates/source-hashes.json"
    expected_sources = json.loads(manifest.read_text())

    def check_sources():
        return all(digest(ROOT / name) == expected for name, expected in expected_sources.items())

    paths = [
        Path(__file__).resolve(), manifest, HERE / "validate_schema.py",
        HERE / "harness/Cargo.toml", HERE / "harness/Cargo.lock", HERE / "harness/main.rs",
        ROOT / "3rdparty/specs/[MS-ODRAWXML]/2 Structures/2.44 http---schemas.microsoft.com-office-drawing-2015-06-chart.md",
        ROOT / "3rdparty/specs/[MS-ODRAWXML]/5 Appendix A - Full XML Schemas/5.42 http---schemas.microsoft.com-office-drawing-2015-06-chart Schema.md",
        ROOT / "3rdparty/specs/[MS-OI29500]/2 Conformance Statements/2.1 Normative Variations.md",
        ROOT / "3rdparty/specs/ECMA-376/ECMA-376-4_5th_edition_december_2016.zip",
    ]
    inputs = {str(path.relative_to(ROOT)): digest(path) for path in paths}
    assert check_sources(), "source differs from compile-first gate inputs"
    relative = str(HERE.relative_to(ROOT))
    commands = [
        (["cargo", "run", "--locked", "--offline", "--manifest-path", f"{relative}/harness/Cargo.toml", "--", f"{relative}/outputs"], "probe.log"),
        ([sys.executable, f"{relative}/validate_schema.py", *[f"{relative}/outputs/{name}.xml" for name in ("source", "edited", "authored")]], "schema-validation.json"),
    ]
    environment = dict(os.environ)
    for name in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP"):
        environment.pop(name, None)
    environment["CARGO_TARGET_DIR"] = str(ROOT / "target")
    receipts = []
    for command, output in commands:
        with (HERE / output).open("w") as log:
            result = subprocess.run(command, cwd=ROOT, env=environment, stdout=log, stderr=subprocess.PIPE, text=True)
        if result.returncode:
            raise RuntimeError(result.stderr)
        receipts.append({"argv": command, "exit_code": result.returncode})
    assert check_sources(), "source changed during probe"
    assert inputs == {str(path.relative_to(ROOT)): digest(path) for path in paths}
    outputs = sorted((HERE / "outputs").glob("*.xml")) + [HERE / "probe.log", HERE / "schema-validation.json"]
    receipt = {
        "passed": True, "source_unchanged": True,
        "revision": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(),
        "binary_sha256": digest(ROOT / "target/debug/chart-formatcode2-evidence"),
        "gate_source_files_checked_before_and_after": len(expected_sources),
        "commands": receipts, "inputs": inputs,
        "outputs": {str(path.relative_to(ROOT)): digest(path) for path in outputs},
        "scope": "Downstream element and attribute examples; XSD only for the three element outputs",
    }
    (HERE / "probe-receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    print(json.dumps({"passed": True, "source_files_checked": len(expected_sources), "schema_outputs": 3}))


if __name__ == "__main__":
    main()
