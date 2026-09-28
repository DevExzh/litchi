"""Root-only release builds for the 0817 ordinary-save packet.

This driver is intentionally serial.  It preserves every failed Cargo
attempt and copies each successful executable to a stable target path used by
the capture driver.  It does not run a workload.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import time

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.assert_origin()
assert PLAN["schema"] == "litchi.performance.0817.plan.v1"
assert PLAN["base"] == ORIGIN["base"]
assert PLAN["target"] == str(c.TARGET)
assert PLAN["scratch"] == str(c.SCRATCH)
assert PLAN["affinity"] == [12, 13, 14, 15, 16, 17, 18, 19]
c.plan_cases(PLAN)
assert not (P / "build.json").exists(), "refusing to overwrite build.json"

attempt = 0
while (P / f"build-{attempt}").exists():
    attempt += 1
OUT = P / f"build-{attempt}"
OUT.mkdir()

production = c.source()
assert production["revision"] == ORIGIN["base"]
root_inputs = c.assert_root_inputs()
locks = c.lock_identity()
tool = c.tool_source()
architecture = c.architecture_hashes()
corpus = c.assert_corpus_inputs()
provenance = c.assert_provenance(corpus)
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
        "schema": "litchi.performance.0817.frozen-inputs.v1",
        "packet": packet,
        "drivers": drivers,
        "root_inputs": root_inputs,
        "locks": locks,
        "architecture": architecture,
        "corpus": corpus,
        "provenance": provenance,
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

build_config = PLAN["build"]
assert build_config["jobs"] == 2
assert build_config["rustflags"] is None
env = os.environ | {
    "CARGO_TARGET_DIR": str(c.TARGET),
    "CARGO_BUILD_JOBS": str(build_config["jobs"]),
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(build_config["opt_level"]),
    "CARGO_PROFILE_RELEASE_DEBUG": str(build_config["debug"]),
    "CARGO_PROFILE_RELEASE_LTO": str(build_config["lto"]),
    "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(build_config["codegen_units"]),
    "CARGO_PROFILE_RELEASE_INCREMENTAL": str(build_config["incremental"]).lower(),
    "CARGO_PROFILE_RELEASE_PANIC": str(build_config["panic"]),
}
manifest = c.TOOL / "Cargo.toml"
steps = (
    ("native", "litchi-perf-baseline", []),
    ("artifacts", "ordinary_save_artifacts", []),
    (
        "observer",
        "litchi-perf-baseline-alloc",
        ["allocator-metrics", "ordinary-save-process-metrics"],
    ),
)
declared_binaries = PLAN.get("binaries")
if declared_binaries is not None:
    assert isinstance(declared_binaries, dict)
    expected_binary_plan = {
        "native": ("litchi-perf-baseline", []),
        "artifacts": ("ordinary_save_artifacts", []),
        "observer": (
            "litchi-perf-baseline-alloc",
            ["allocator-metrics", "ordinary-save-process-metrics"],
        ),
    }
    for logical, (expected_name, expected_features) in expected_binary_plan.items():
        declared = declared_binaries[logical]
        assert declared.get("features") == expected_features
        declared_name = declared["cargo_bin"]
        assert declared_name == expected_name
rows = []
binaries = {}
c.TARGET.mkdir(parents=True, exist_ok=True)

try:
    for name, binary_name, features in steps:
        command = [
            "cargo",
            "build",
            "--offline",
            "--locked",
            "--release",
            "--manifest-path",
            str(manifest),
            "--bin",
            binary_name,
        ]
        if features:
            command.extend(["--features", ",".join(features)])
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
            "binary": name,
            "cargo_binary": binary_name,
            "features": features,
            "command": command,
            "started": started,
            "ended": time.time(),
            "exit_code": result.returncode,
            "log": c.artifact(log),
            "environment": {
                key: env.get(key)
                for key in (
                    "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                    "CARGO_PROFILE_RELEASE_OPT_LEVEL", "CARGO_PROFILE_RELEASE_DEBUG",
                    "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
                    "CARGO_PROFILE_RELEASE_INCREMENTAL", "CARGO_PROFILE_RELEASE_PANIC",
                )
            },
        }
        rows.append(row)
        c.write(OUT / "commands.json", rows)
        if result.returncode != 0:
            c.write(
                OUT / "failure.json",
                {
                    "schema": "litchi.performance.0817.build-failure.v1",
                    "attempt": attempt,
                    "rows": rows,
                    "source": c.artifact(source_path),
                    "frozen_inputs": c.artifact(frozen_inputs_path),
                },
            )
            raise RuntimeError(f"{name} build failed; retained {log}")

        c.check_stable(
            frozen, root_inputs, locks, architecture, corpus, host,
            packet, drivers, unrelated,
        )
        executable = c.TARGET / f"release/{binary_name}"
        assert executable.is_file(), f"missing release executable for {name}: {executable}"
        destination = c.TARGET / name
        assert not destination.exists(), f"refusing to overwrite {destination}"
        shutil.copy2(executable, destination)
        binaries[name] = {
            "cargo_name": binary_name,
            "features": features,
            "artifact": c.artifact(destination),
        }
        print(name, "built", flush=True)

    result = {
        "schema": "litchi.performance.0817.build.v1",
        "attempt": attempt,
        "source": c.artifact(source_path),
        "frozen_inputs": c.artifact(frozen_inputs_path),
        "root_inputs": root_inputs,
        "locks": locks,
        "architecture": architecture,
        "corpus": corpus,
        "provenance": provenance,
        "host": host,
        "unrelated": unrelated,
        "packet": packet,
        "drivers": drivers,
        "binaries": binaries,
        "rows": rows,
        "environment": {
            key: env.get(key)
            for key in (
                "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                "CARGO_PROFILE_RELEASE_OPT_LEVEL", "CARGO_PROFILE_RELEASE_DEBUG",
                "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
                "CARGO_PROFILE_RELEASE_INCREMENTAL", "CARGO_PROFILE_RELEASE_PANIC",
            )
        },
        "profile": {
            key: build_config[key]
            for key in ("opt_level", "debug", "lto", "codegen_units", "incremental", "panic")
        },
    }
    c.write(OUT / "build.json", result)
    c.write(P / "build.json", result)
    print("0817 build complete", flush=True)
except Exception:
    raise
