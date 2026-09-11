#!/usr/bin/env python3
"""Run the focused production gates and bind results to unchanged inputs."""

import hashlib
import json
import os
from pathlib import Path
import re
import subprocess


HERE = Path(__file__).resolve().parent
ROOT = next(p for p in HERE.parents if (p / "crates").is_dir())


def inputs():
    paths = [ROOT / "Cargo.toml", ROOT / "Cargo.lock", ROOT / "rust-toolchain.toml"]
    for name in ["litchi-drawingml", "litchi-ooxml-common", "litchi-opc", "litchi-core", "litchi-cfb", "litchi-sign", "soapberry-zip", "litchi-xlsb", "litchi-sheet", "litchi-spreadsheet-drawing", "litchi-xldm", "litchi-vba", "litchi-codepage", "litchi-crypto", "litchi-ole-common", "xml-minifier", "xml-minifier-macros"]:
        crate = ROOT / "crates" / name
        paths.append(crate / "Cargo.toml")
        if (crate / "build.rs").is_file():
            paths.append(crate / "build.rs")
        paths.extend(p for p in (crate / "src").rglob("*") if p.is_file())
    for name in ["litchi-drawingml", "litchi-xlsb"]:
        paths.extend(p for p in (ROOT / "crates" / name / "tests").rglob("*") if p.is_file())
    paths.extend(ROOT / name for name in ["tools/check_crate_boundaries.py", "tools/crate_boundaries.json", "tools/test_check_crate_boundaries.py"])
    return {str(p.relative_to(ROOT)): hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(paths)}


def main():
    directory = HERE / "gates"
    directory.mkdir(exist_ok=True)
    before = inputs()
    (directory / "source-hashes.json").write_text(json.dumps(before, indent=2) + "\n")
    commands = [
        ("compile", ["cargo", "check", "--locked", "-p", "litchi-drawingml", "-p", "litchi-xlsb", "--all-features", "--all-targets"]),
        ("tests", ["cargo", "test", "--locked", "-p", "litchi-drawingml", "-p", "litchi-xlsb", "--all-features", "--no-fail-fast"]),
        ("clippy", ["cargo", "clippy", "--locked", "-p", "litchi-drawingml", "-p", "litchi-xlsb", "--all-features", "--all-targets", "--", "-D", "warnings"]),
        ("rustdoc", ["cargo", "doc", "--locked", "-p", "litchi-drawingml", "-p", "litchi-xlsb", "--all-features", "--no-deps"]),
        ("format", ["cargo", "fmt", "-p", "litchi-drawingml", "-p", "litchi-xlsb", "--check"]),
        ("diff", ["git", "diff", "--check"]),
    ]
    receipt = {"base": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(), "toolchain": subprocess.check_output(["rustc", "-vV"], cwd=ROOT, text=True), "commands": []}
    for name, command in commands:
        environment = dict(os.environ)
        if name == "rustdoc":
            environment["RUSTDOCFLAGS"] = "-D warnings"
        with (directory / f"{name}.log").open("w") as log:
            result = subprocess.run(command, cwd=ROOT, env=environment, stdout=log, stderr=subprocess.STDOUT, check=False)
        receipt["commands"].append({"name": name, "argv": command, "exit_code": result.returncode, "RUSTDOCFLAGS": environment.get("RUSTDOCFLAGS", "")})
        print(f"{name}: {result.returncode}", flush=True)
        if result.returncode:
            break
    receipt["source_unchanged"] = before == inputs()
    test_log = (directory / "tests.log").read_text()
    totals = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", test_log)
    receipt["test_totals_including_doctests"] = dict(zip(["passed", "failed", "ignored"], [sum(int(row[i]) for row in totals) for i in range(3)]))
    receipt["logs"] = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(directory.glob("*.log"))}
    receipt["passed"] = len(receipt["commands"]) == len(commands) and all(row["exit_code"] == 0 for row in receipt["commands"]) and receipt["source_unchanged"]
    (directory / "receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    if not receipt["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
