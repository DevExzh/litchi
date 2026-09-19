#!/usr/bin/env python3
"""Run final gates against an isolated checkout containing the selected batch."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import time


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def manifest(repo):
    paths = [repo / "Cargo.toml", repo / "Cargo.lock"]
    for crate in (repo / "crates").iterdir():
        for name in ("Cargo.toml", "build.rs"):
            if (crate / name).is_file():
                paths.append(crate / name)
        for directory in ("src", "tests"):
            paths.extend((crate / directory).rglob("*.rs"))
    paths.extend((repo / "crates/litchi-xlsx/tests/fixtures/form_control_properties").glob("*.xlsx"))
    paths.append(repo / "3rdparty/libreoffice-core/sc/qa/unit/data/xlsx/tdf169496_hidden_graphic.xlsx")
    paths.append(repo / "tools/check_crate_boundaries.py")
    return {str(p.relative_to(repo)): digest(p) for p in sorted(set(paths))}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("repo", type=Path)
    parser.add_argument("target", type=Path)
    args = parser.parse_args()
    repo = args.repo.resolve()
    output = Path(__file__).resolve().parent
    metadata = {"repo": str(repo), "target": str(args.target.resolve())}
    for label, command in (("head", ["git", "rev-parse", "HEAD"]),
                           ("rustc", ["rustc", "-Vv"]),
                           ("cargo", ["cargo", "-V"]),
                           ("kernel", ["uname", "-a"])):
        metadata[label] = subprocess.check_output(command, cwd=repo, text=True).strip()
    metadata["RUSTFLAGS"] = os.environ.get("RUSTFLAGS")
    metadata["RUSTDOCFLAGS"] = "-D warnings"
    (output / "environment.json").write_text(json.dumps(metadata, indent=2) + "\n")
    before = manifest(repo)
    (output / "source-before.json").write_text(json.dumps(before, indent=2) + "\n")
    shutil.copyfile(repo / "Cargo.lock", output / "Cargo.lock")
    env = dict(os.environ, CARGO_TARGET_DIR=str(args.target.resolve()))
    batch_files = json.loads((output / "batch-files.json").read_text())
    commands = [
        ("xlsx-tests", ["cargo", "test", "--locked", "-p", "litchi-xlsx"]),
        ("opc-tests", ["cargo", "test", "--locked", "-p", "litchi-opc"]),
        ("clippy", ["cargo", "clippy", "--locked", "-p", "litchi-xlsx", "-p", "litchi-opc", "--all-targets", "--", "-D", "warnings"]),
        ("rustdoc", ["cargo", "doc", "--locked", "-p", "litchi-xlsx", "-p", "litchi-opc", "--no-deps"]),
        ("format", ["cargo", "fmt", "-p", "litchi-xlsx", "-p", "litchi-opc", "--", "--check"]),
        ("batch-format", ["rustfmt", "--edition", "2024", "--check", "--config", "skip_children=true", *batch_files]),
        ("boundaries", ["python3", "tools/check_crate_boundaries.py"]),
        ("diff-check", ["git", "diff", "--check"]),
    ]
    results = []
    for name, command in commands:
        current_env = dict(env)
        if name == "rustdoc":
            current_env["RUSTDOCFLAGS"] = "-D warnings"
        started = time.time()
        log = output / (name + ".log")
        with log.open("w") as sink:
            run = subprocess.run(command, cwd=repo, env=current_env, stdout=sink, stderr=subprocess.STDOUT)
        # Strip display-only trailing spaces so retained diff logs are git-clean.
        log.write_text("\n".join(line.rstrip(" \t") for line in log.read_text().splitlines()) + ("\n" if log.stat().st_size else ""))
        results.append({"name": name, "command": command, "exit_code": run.returncode,
                        "seconds": time.time() - started, "log_sha256": digest(log)})
        print(f"{name}: {run.returncode}", flush=True)
        (output / "results.json").write_text(json.dumps(results, indent=2) + "\n")
    after = manifest(repo)
    (output / "source-after.json").write_text(json.dumps(after, indent=2) + "\n")
    stable = before == after
    baseline_format_file = "crates/litchi-xlsx/tests/drawing_svg_read.rs"
    baseline_bytes = subprocess.check_output(["git", "show", "HEAD:" + baseline_format_file], cwd=repo)
    format_diffs = [line.split(":", 1)[0].removeprefix("Diff in ")
                    for line in (output / "format.log").read_text().splitlines()
                    if line.startswith("Diff in ")]
    known_format_baseline = (bool(format_diffs)
        and set(format_diffs) == {str(repo / baseline_format_file)}
        and (repo / baseline_format_file).read_bytes() == baseline_bytes)
    required_passed = all(r["exit_code"] == 0 or
                         (r["name"] == "format" and r["exit_code"] == 1 and known_format_baseline)
                         for r in results)
    (output / "verification.json").write_text(json.dumps({"stable_sources": stable,
        "all_commands_passed": all(r["exit_code"] == 0 for r in results),
        "known_unchanged_baseline_format_failure": known_format_baseline,
        "all_required_checks_passed": required_passed}, indent=2) + "\n")
    return 0 if stable and required_passed else 1


if __name__ == "__main__":
    raise SystemExit(main())
