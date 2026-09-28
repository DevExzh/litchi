"""Root-only serial builds for the 0816 standalone execution harness.

Each failed attempt keeps its directory and logs.  A successful attempt also
gets an immutable top-level ``build.json`` used by the capture and audit
drivers.  This driver executes Cargo only when invoked by the root agent.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import time

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
assert PLAN["schema"] == "litchi.performance.0816.plan.v1"
assert ORIGIN["base"] == "c8ff2b9f65ecd011d12fd5b1ac352ea315785b93"
assert ORIGIN["target"] == str(c.TARGET)
assert ORIGIN["tool_allowlist"] == [
    "tools/perf-execution/src/main.rs",
    "tools/perf-execution/README.md",
]
assert not (P / "build.json").exists(), "refusing to overwrite build.json"

attempt = 0
while (P / f"build-{attempt}").exists():
    attempt += 1
OUT = P / f"build-{attempt}"
OUT.mkdir()

production = c.source()
assert production["revision"] == ORIGIN["base"]
assert len(production["files"]) == 9196
root_inputs = c.assert_root_inputs()
tool_lock = c.assert_tool_lock()
tool = c.tool_source()
architecture = c.architecture_hashes()
host = c.assert_host()
unrelated = c.assert_unrelated()
packet = c.packet_hashes()
drivers = c.driver_hashes()
frozen = {"production": production, "tool": tool}

source_path = OUT / "source.json"
c.write(source_path, frozen)
frozen_inputs_path = OUT / "frozen-inputs.json"
c.write(
    frozen_inputs_path,
    {
        "schema": "litchi.performance.0816.frozen-inputs.v1",
        "packet": packet,
        "drivers": drivers,
        "root_inputs": root_inputs,
        "tool_lock": tool_lock,
        "architecture": architecture,
        "host": host,
        "unrelated": unrelated,
    },
)

for name in (
    "RUSTUP_TOOLCHAIN",
    "RUSTFLAGS",
    "CARGO_ENCODED_RUSTFLAGS",
    "CARGO_BUILD_RUSTFLAGS",
    "CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS",
    "RUSTC_WRAPPER",
    "RUSTC_WORKSPACE_WRAPPER",
):
    assert not os.environ.get(name), f"unexpected inherited build override: {name}"

env = os.environ | {
    "CARGO_TARGET_DIR": str(c.TARGET),
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
}
manifest = c.TOOL / "Cargo.toml"
rows = []
binaries = {}
variants = (
    ("native", []),
    ("observer", ["--features", "source-metrics"]),
)

try:
    for name, features in variants:
        command = [
            "cargo",
            "build",
            "--offline",
            "--locked",
            "--release",
            "--manifest-path",
            str(manifest),
            *features,
        ]
        log = OUT / f"{name}.log"
        started = time.time()
        with log.open("w") as stream:
            result = subprocess.run(
                command,
                cwd=c.ROOT,
                env=env,
                stdout=stream,
                stderr=subprocess.STDOUT,
            )
        row = {
            "variant": name,
            "features": features,
            "command": command,
            "started": started,
            "ended": time.time(),
            "exit_code": result.returncode,
            "log": c.artifact(log),
            "environment": {
                key: env.get(key)
                for key in ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL")
            },
        }
        rows.append(row)
        c.write(OUT / "commands.json", rows)
        if result.returncode != 0:
            c.write(
                OUT / "failure.json",
                {
                    "schema": "litchi.performance.0816.build-failure.v1",
                    "attempt": attempt,
                    "rows": rows,
                    "source": c.artifact(source_path),
                    "frozen_inputs": c.artifact(frozen_inputs_path),
                },
            )
            raise RuntimeError(f"{name} build failed; retained {log}")

        c.unchanged(frozen)
        assert c.assert_root_inputs() == root_inputs
        assert c.assert_tool_lock() == tool_lock
        assert c.architecture_hashes() == architecture
        assert c.assert_host(host) == host
        assert c.packet_hashes() == packet
        assert c.assert_unrelated() == unrelated

        executable = c.TARGET / "release/litchi-perf-execution"
        assert executable.is_file(), f"missing release executable for {name}"
        destination = c.TARGET / f"attempt-{attempt}-{name}"
        assert not destination.exists(), f"refusing to overwrite {destination}"
        shutil.copy2(executable, destination)
        binaries[name] = c.artifact(destination)
        print(name, "built", flush=True)

    result = {
        "schema": "litchi.performance.0816.build.v1",
        "attempt": attempt,
        "source": c.artifact(source_path),
        "frozen_inputs": c.artifact(frozen_inputs_path),
        "root_inputs": root_inputs,
        "tool_lock": tool_lock,
        "architecture": architecture,
        "host": host,
        "unrelated": unrelated,
        "packet": packet,
        "drivers": drivers,
        "binaries": binaries,
        "rows": rows,
        "environment": {
            key: env.get(key)
            for key in ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL")
        },
        "profile": {
            "opt_level": 3,
            "debug": 1,
            "lto": "thin",
            "codegen_units": 1,
            "incremental": False,
            "panic": "unwind",
        },
    }
    c.write(OUT / "build.json", result)
    c.write(P / "build.json", result)
    print("0816 build complete", flush=True)
except Exception:
    raise
