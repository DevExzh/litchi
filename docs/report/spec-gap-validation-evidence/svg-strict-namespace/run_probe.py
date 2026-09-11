#!/usr/bin/env python3
"""Run the retained Strict-namespace downstream harness and schema checks."""

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
BASE = ROOT / "docs/report/spec-gap-validation-evidence/pptx-svg-lifecycle/outputs/source.pptx"


def digest(path):
    result = hashlib.sha256()
    with path.open("rb") as source:
        for block in iter(lambda: source.read(1 << 20), b""):
            result.update(block)
    return result.hexdigest()


def source_snapshot(manifest):
    values = json.loads(manifest.read_text())
    observed = {name: digest(ROOT / name) for name in values}
    return values, observed


def run(command, output, environment):
    with output.open("w") as sink:
        result = subprocess.run(
            command,
            cwd=ROOT,
            env=environment,
            stdout=sink,
            stderr=subprocess.STDOUT,
            check=False,
        )
    return {"argv": command, "exit_code": result.returncode, "output": str(output.relative_to(ROOT))}


def main():
    gates = HERE / "gates"
    raw = HERE / "raw-results"
    outputs = HERE / "outputs"
    manifest = gates / "source-hashes.json"
    compile_receipt = json.loads((gates / "receipt.json").read_text())
    if not compile_receipt["passed"]:
        raise RuntimeError("compile-first receipt is not passing")
    expected, before = source_snapshot(manifest)
    if before != expected:
        raise RuntimeError("source differs from compile-first manifest")
    if not BASE.is_file():
        raise FileNotFoundError(BASE)
    raw.mkdir(exist_ok=True)
    outputs.mkdir(exist_ok=True)
    owned_target = HERE / "harness/target"
    if owned_target.exists():
        shutil.rmtree(owned_target)

    environment = dict(os.environ)
    for name in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_BOOTSTRAP"):
        environment.pop(name, None)
    environment["CARGO_TARGET_DIR"] = str(ROOT / "target")
    source_arg = str(BASE.relative_to(ROOT))
    output_arg = str(outputs.relative_to(ROOT))
    commands = []
    commands.append(run(
        [
            "cargo",
            "run",
            "--locked",
            "--offline",
            "--manifest-path",
            f"{RELATIVE}/harness/Cargo.toml",
            "--",
            source_arg,
            output_arg,
        ],
        raw / "harness.log",
        environment,
    ))
    if commands[-1]["exit_code"] != 0:
        raise RuntimeError("downstream harness failed; see raw-results/harness.log")

    slide_names = ["strict-source", "strict-attached", "strict-detached"]
    controls = ["direct-transitional-valid", "direct-strict-invalid"]
    schema_command = [
        sys.executable,
        f"{RELATIVE}/validate_schema.py",
        *[f"{output_arg}/{name}.xml" for name in (*slide_names, *controls)],
    ]
    commands.append(run(schema_command, raw / "schema-validation.json", environment))
    if commands[-1]["exit_code"] != 0:
        raise RuntimeError("offline schema validator failed; see raw-results/schema-validation.json")
    schema = json.loads((raw / "schema-validation.json").read_text())
    if not schema["passed"]:
        raise RuntimeError("schema validator reported failure")
    after, observed_after = source_snapshot(manifest)
    if before != after or before != observed_after:
        raise RuntimeError("source changed during downstream probe")
    provenance = {
        "revision": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(),
        "python": sys.version,
        "rustc": subprocess.check_output(["rustc", "-vV"], cwd=ROOT, text=True),
        "cargo": subprocess.check_output(["cargo", "--version"], cwd=ROOT, text=True).strip(),
        "source_fixture": str(BASE.relative_to(ROOT)),
        "source_fixture_sha256": digest(BASE),
        "harness_binary": "target/debug/svg-strict-namespace-evidence",
        "harness_binary_sha256": digest(ROOT / "target/debug/svg-strict-namespace-evidence"),
        "validator_sha256": digest(HERE / "validate_schema.py"),
        "commands": commands,
        "source_hashes_sha256": digest(manifest),
    }
    (raw / "tool-provenance.json").write_text(json.dumps(provenance, indent=2) + "\n")
    generated = sorted(path for path in outputs.iterdir() if path.is_file())
    raw_results = sorted(path for path in raw.iterdir() if path.is_file())
    receipt = {
        "passed": True,
        "source_unchanged": True,
        "compile_first_receipt_sha256": digest(gates / "receipt.json"),
        "source_files_checked": len(expected),
        "source_fixture": provenance["source_fixture"],
        "source_fixture_sha256": provenance["source_fixture_sha256"],
        "commands": commands,
        "inputs": {
            "gates/source-hashes.json": digest(manifest),
            str(BASE.relative_to(ROOT)): digest(BASE),
        },
        "generated_outputs": {
            str(path.relative_to(ROOT)): digest(path) for path in generated
        },
        "raw_results": {
            str(path.relative_to(ROOT)): digest(path) for path in raw_results
        },
        "schema_summary": {
            "strict_slide_reports": sum(
                row["kind"] == "strict-presentation-slide" for row in schema["reports"]
            ),
            "direct_controls": [
                {"path": row["path"], "schema_valid": row["schema_valid"]}
                for row in schema["reports"]
                if row["kind"] == "direct-ms-odrawxml-svg-child"
            ],
            "xlsx_mixed_namespace": next(
                row for row in schema["reports"]
                if row["kind"] == "strict-native-xlsx-mixed-namespace-drawing"
            ),
        },
        "scope": "Bounded Strict OOXML core plus Transitional MS-ODRAWXML SVG child namespace proof and ordinary downstream PPTX lifecycle",
    }
    receipt_path = HERE / "probe-receipt.json"
    receipt_path.write_text(json.dumps(receipt, indent=2) + "\n")
    if owned_target.exists():
        shutil.rmtree(owned_target)
    print(json.dumps(receipt, indent=2))


if __name__ == "__main__":
    main()
