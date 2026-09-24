#!/usr/bin/env python3
"""Build and run the public caller in an owned temporary target directory."""

import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile

HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / "crates").is_dir())
sys.dont_write_bytecode = True
spec = importlib.util.spec_from_file_location("family_gates", HERE / "run_gates.py")
gates = importlib.util.module_from_spec(spec)
spec.loader.exec_module(gates)


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def main():
    source = gates.inputs()
    example = HERE / "caller-example"
    caller_inputs = {str(p.relative_to(HERE)): digest(p)
                     for p in [example / "Cargo.toml", example / "Cargo.lock", example / "main.rs"]}
    target = Path(tempfile.mkdtemp(prefix="litchi-xlsb-theme-family-caller-"))
    command = ["cargo", "run", "--locked", "--manifest-path", str(example / "Cargo.toml")]
    environment = dict(os.environ, CARGO_TARGET_DIR=str(target))
    try:
        with (HERE / "caller.log").open("w") as log:
            result = subprocess.run(command, cwd=ROOT, env=environment,
                                    stdout=log, stderr=subprocess.STDOUT, check=False)
        if result.returncode:
            raise RuntimeError(f"caller failed: see {HERE / 'caller.log'}")
        assert source == gates.inputs(), "production inputs changed during caller build/run"
        assert all(digest(HERE / name) == value for name, value in caller_inputs.items())
        xml_paths = [HERE / f"caller-{name}.xml" for name in ["updated", "removed", "added", "authored"]]
        with (HERE / "caller-schema-validation.json").open("w") as log:
            subprocess.run([sys.executable, str(HERE / "validate_schema.py"),
                            *[str(p.relative_to(ROOT)) for p in xml_paths]],
                           cwd=ROOT, stdout=log, check=True)
        report = {
            "passed": True, "source_unchanged": True,
            "production_source_manifest_sha256": hashlib.sha256(json.dumps(source, sort_keys=True).encode()).hexdigest(),
            "caller_inputs": caller_inputs,
            "binary_sha256": digest(target / "debug/xlsb-theme-family-caller-example"),
            "log_sha256": digest(HERE / "caller.log"),
            "outputs": {p.name: digest(p) for p in xml_paths},
            "command": ["cargo", "run", "--locked", "--manifest-path", str((example / "Cargo.toml").relative_to(ROOT))],
            "toolchain": subprocess.check_output(["rustc", "-vV"], cwd=ROOT, text=True),
        }
    finally:
        shutil.rmtree(target)
    report["owned_temporary_target_removed"] = not target.exists()
    (HERE / "caller-receipt.json").write_text(json.dumps(report, indent=2) + "\n")
    print(json.dumps({"passed": True, "outputs": len(xml_paths), "temporary_target_removed": True}))


if __name__ == "__main__":
    main()
