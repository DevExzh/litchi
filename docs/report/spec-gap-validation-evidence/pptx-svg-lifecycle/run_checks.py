#!/usr/bin/env python3
"""Run compile-first PPTX SVG lifecycle gates against unchanged source inputs."""

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
    metadata = json.loads(subprocess.check_output(
        ["cargo", "metadata", "--locked", "--offline", "--format-version", "1", "--all-features"], cwd=ROOT))
    packages = {p["id"]: p for p in metadata["packages"]}
    nodes = {p["id"]: p for p in metadata["resolve"]["nodes"]}
    pending = [p["id"] for p in packages.values() if p["name"] in ("litchi-pptx", "litchi-drawingml")]
    seen = set()
    while pending:
        identity = pending.pop()
        if identity in seen:
            continue
        seen.add(identity)
        pending.extend(nodes[identity]["dependencies"])
    for identity in sorted(seen):
        package = packages[identity]
        if package["source"] is not None:
            continue
        manifest = Path(package["manifest_path"])
        crate = manifest.parent
        paths.append(manifest)
        if (crate / "build.rs").is_file():
            paths.append(crate / "build.rs")
        for subtree in ("src", "tests", "examples", "benches"):
            paths.extend(p for p in (crate / subtree).rglob("*") if p.is_file())
    for name in ["litchi-drawingml", "litchi-pptx", "litchi-opc"]:
        paths.extend(p for p in (ROOT / "crates" / name / "tests").rglob("*") if p.is_file())
    paths.extend(p for p in (ROOT / ".cargo").rglob("*") if p.is_file())
    paths.extend(p for p in (ROOT / "docs/adr").glob("*.md"))
    paths.extend(HERE / name for name in (
        "run_checks.py", "run_probe.py", "validate_schema.py", "verify.py",
        "requirements.md", "README.md", "corpus-scan.json", "scan_corpus.py", "worktree-input.patch",
        "harness/Cargo.toml", "harness/Cargo.lock", "harness/main.rs",
    ))
    paths.append(ROOT / "crates/litchi-pptx/docs/FEATURE_MATRIX.md")
    corpus = json.loads((HERE / "corpus-scan.json").read_text())
    paths.extend(ROOT / row[0] for report in corpus["reports"] for row in report["hits"])
    paths.extend(ROOT / name for name in ["tools/check_crate_boundaries.py", "tools/crate_boundaries.json", "tools/test_check_crate_boundaries.py"])
    return {str(p.relative_to(ROOT)): hashlib.sha256(p.read_bytes()).hexdigest() for p in sorted(set(paths))}


def main():
    directory = HERE / "gates"
    directory.mkdir(exist_ok=True)
    before = inputs()
    (directory / "source-hashes.json").write_text(json.dumps(before, indent=2) + "\n")
    commands = [
        ("compile", ["cargo", "check", "--locked", "-p", "litchi-drawingml", "-p", "litchi-pptx", "-p", "litchi-opc", "--all-features", "--all-targets"]),
        ("tests", ["cargo", "test", "--locked", "-p", "litchi-drawingml", "-p", "litchi-pptx", "-p", "litchi-opc", "--all-features", "--no-fail-fast"]),
        ("clippy", ["cargo", "clippy", "--locked", "-p", "litchi-drawingml", "-p", "litchi-pptx", "-p", "litchi-opc", "--all-features", "--all-targets", "--", "-D", "warnings"]),
        ("rustdoc", ["cargo", "doc", "--locked", "-p", "litchi-drawingml", "-p", "litchi-pptx", "-p", "litchi-opc", "--all-features", "--no-deps"]),
        ("format", ["cargo", "fmt", "-p", "litchi-drawingml", "-p", "litchi-pptx", "-p", "litchi-opc", "--check"]),
        ("diff", ["git", "diff", "--check"]),
    ]
    receipt = {"base": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(), "toolchain": subprocess.check_output(["rustc", "-vV"], cwd=ROOT, text=True), "commands": []}
    for name, command in commands:
        environment = dict(os.environ)
        environment.pop("RUSTFLAGS", None)
        environment.pop("CARGO_ENCODED_RUSTFLAGS", None)
        environment.pop("RUSTC_BOOTSTRAP", None)
        if name == "rustdoc":
            environment["RUSTDOCFLAGS"] = "-D warnings"
        with (directory / f"{name}.log").open("w") as log:
            result = subprocess.run(command, cwd=ROOT, env=environment, stdout=log, stderr=subprocess.STDOUT, check=False)
        receipt["commands"].append({"name": name, "argv": command, "exit_code": result.returncode, "RUSTDOCFLAGS": environment.get("RUSTDOCFLAGS", ""), "RUSTFLAGS": environment.get("RUSTFLAGS", ""), "RUSTC_BOOTSTRAP": environment.get("RUSTC_BOOTSTRAP", ""), "CARGO_ENCODED_RUSTFLAGS": environment.get("CARGO_ENCODED_RUSTFLAGS", "")})
        print(f"{name}: {result.returncode}", flush=True)
        if result.returncode:
            break
    receipt["source_unchanged"] = before == inputs()
    executed_names = {row["name"] for row in receipt["commands"]}
    test_path = directory / "tests.log"
    test_log = test_path.read_text() if "tests" in executed_names else ""
    totals = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", test_log)
    receipt["test_totals_including_doctests"] = dict(zip(["passed", "failed", "ignored"], [sum(int(row[i]) for row in totals) for i in range(3)]))
    receipt["logs"] = {
        f"{name}.log": hashlib.sha256((directory / f"{name}.log").read_bytes()).hexdigest()
        for name in sorted(executed_names)
    }
    receipt["passed"] = len(receipt["commands"]) == len(commands) and all(row["exit_code"] == 0 for row in receipt["commands"]) and receipt["source_unchanged"]
    (directory / "receipt.json").write_text(json.dumps(receipt, indent=2) + "\n")
    if not receipt["passed"]:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
