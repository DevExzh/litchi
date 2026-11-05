#!/usr/bin/env python3
"""Build the same standalone probe against an immutable workspace snapshot."""
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time


def artifact(path):
    path = Path(path).resolve()
    return dict(path=str(path), bytes=path.stat().st_size,
                sha256=hashlib.sha256(path.read_bytes()).hexdigest())


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("target", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("--policy", action="store_true")
    args = parser.parse_args()
    source, target, output = (path.resolve() for path in (args.source, args.target, args.output))
    output.mkdir(parents=True, exist_ok=False)
    probe = Path(__file__).with_name("char_style_probe.rs").resolve()
    manifest = dict(source=str(source), target=str(target), probe=artifact(probe),
                    builder=artifact(__file__), lock=artifact(source / "Cargo.lock"),
                    policy=args.policy, commands=[], artifacts={},
                    rustc=subprocess.check_output(["rustc", "-Vv"], text=True))
    env = os.environ.copy()
    env.update(CARGO_TARGET_DIR=str(target), CARGO_BUILD_JOBS="1", CARGO_INCREMENTAL="0")

    def execute(command, name):
        started = time.time_ns()
        with (output / f"{name}.stdout").open("w") as stdout, (output / f"{name}.stderr").open("w") as stderr:
            result = subprocess.run(command, cwd=source, env=env, stdout=stdout, stderr=stderr)
        manifest["commands"].append(dict(command=command, exit_code=result.returncode,
           started_ns=started, finished_ns=time.time_ns(), stdout=artifact(output / f"{name}.stdout"),
           stderr=artifact(output / f"{name}.stderr")))
        (output / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")
        result.check_returncode()

    execute(["cargo", "build", "--offline", "--locked", "--release", "--lib",
             "-p", "litchi-iwa-text-wire", "--message-format=json-render-diagnostics"], "library")
    libraries = {}
    for line in (output / "library.stdout").read_text().splitlines():
        try:
            item = json.loads(line)
        except json.JSONDecodeError:
            continue
        if item.get("reason") == "compiler-artifact":
            for path in item["filenames"]:
                if path.endswith(".rlib"):
                    libraries[item["target"]["name"]] = path
    for name in ("litchi_iwa_text_wire", "litchi_iwa_common"):
        manifest["artifacts"][name] = artifact(libraries[name])
    for name, allocation in (("native", False), ("allocation", True)):
        binary = output / f"char-style-{name}"
        command = ["rustc", "--edition=2024", "--crate-name", "char_style_probe",
                   "-C", "opt-level=3", "-C", "panic=abort", "-C", "lto=fat",
                   "-C", "embed-bitcode=yes", "--check-cfg", "cfg(char_style_policy)",
                   "--check-cfg", "cfg(allocation_probe)", "-L", f"dependency={target / 'release/deps'}",
                   "--extern", f"litchi_iwa_text_wire={libraries['litchi_iwa_text_wire']}",
                   "--extern", f"litchi_iwa_common={libraries['litchi_iwa_common']}"]
        if args.policy:
            command += ["--cfg", "char_style_policy"]
        if allocation:
            command += ["--cfg", "allocation_probe"]
        execute(command + [str(probe), "-o", str(binary)], name)
        manifest["artifacts"][name] = artifact(binary)
    manifest["status"] = "complete"
    (output / "manifest.json").write_text(json.dumps(manifest, indent=2) + "\n")
    print(str(output), flush=True)


if __name__ == "__main__":
    main()
