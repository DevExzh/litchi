"""Root-owned serial release builds for the two 0828 probe executables."""

from __future__ import annotations

import os
import shutil
import subprocess
import time

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
QUALITY = c.read(P / "quality.json")
OUT = P / "build-0"
assert PLAN["schema"] == "litchi.performance.0828.pptx-edit-profile.v1"
assert ORIGIN["base"] == PLAN["base"] == c.PREVIOUS_COMMIT
assert QUALITY["schema"] == "litchi.performance.0828.quality.v1"
assert QUALITY["status"] == "pass"
assert QUALITY["probe"]["gate_count"] == len(PLAN["quality"]["probe_gates"])
assert not OUT.exists(), f"refusing to overwrite {OUT}"
c.check_no_overrides()

frozen_path = c.Path(QUALITY["frozen_inputs"]["path"])
frozen = c.read(frozen_path)
c.stable(frozen)
assert c.probe_files() == frozen["probe"]
assert c.TARGET.is_dir() and not c.TARGET.is_symlink()

OUT.mkdir()
frozen_copy = OUT / "frozen-inputs.json"
c.write(frozen_copy, frozen)
source_path = OUT / "source.json"
c.write(source_path, {"production": frozen["source"], "tool": frozen["tool"]})
probe_path = OUT / "probe.json"
c.write(probe_path, frozen["probe"])

manifest = P / "probe-src/Cargo.toml"
assert manifest.is_file() and not manifest.is_symlink()
assert c.sha(manifest) == c.sha(P / "probe-src/Cargo.toml")
binary_name = PLAN["probe"]["cargo_binary"]
base_env = os.environ.copy()
base_env.update({
    "CARGO_TARGET_DIR": str(c.TARGET),
    "CARGO_BUILD_JOBS": str(PLAN["build"]["jobs"]),
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(PLAN["build"]["opt_level"]),
    "CARGO_PROFILE_RELEASE_DEBUG": str(PLAN["build"]["debug"]),
    "CARGO_PROFILE_RELEASE_LTO": str(PLAN["build"]["lto"]),
    "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(PLAN["build"]["codegen_units"]),
    "CARGO_PROFILE_RELEASE_INCREMENTAL": "false",
    "CARGO_PROFILE_RELEASE_PANIC": str(PLAN["build"]["panic"]),
    "PYTHONDONTWRITEBYTECODE": "1",
})

rows: list[dict[str, object]] = []
binaries: dict[str, dict[str, object]] = {}
for variant, rustflags in (("ordinary", None), ("fp", PLAN["build"]["fp_rustflags"])):
    env = base_env.copy()
    if rustflags is not None:
        env["RUSTFLAGS"] = rustflags
    command = [
        "cargo", "build", "--offline", "--locked", "--release",
        "--manifest-path", str(manifest), "--bin", binary_name,
    ]
    log = OUT / f"{variant}.log"
    assert not log.exists()
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
        "variant": variant,
        "cargo_binary": binary_name,
        "features": PLAN["build"]["features"],
        "rustflags": rustflags,
        "command": command,
        "environment": {
            key: env.get(key)
            for key in (
                "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                "CARGO_PROFILE_RELEASE_OPT_LEVEL", "CARGO_PROFILE_RELEASE_DEBUG",
                "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
                "CARGO_PROFILE_RELEASE_INCREMENTAL", "CARGO_PROFILE_RELEASE_PANIC",
                "RUSTFLAGS", "PYTHONDONTWRITEBYTECODE",
            )
        },
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "log": c.artifact(log),
    }
    rows.append(row)
    c.write(OUT / "commands.json", rows)
    if result.returncode != 0:
        c.write(OUT / "failure.json", {
            "schema": "litchi.performance.0828.build-failure.v1",
            "rows": rows,
            "source": c.artifact(source_path),
            "frozen_inputs": c.artifact(frozen_copy),
        })
        raise RuntimeError(f"0828 {variant} build failed; retained {log}")
    c.stable(frozen)
    executable = c.TARGET / f"release/{binary_name}"
    assert executable.is_file() and not executable.is_symlink(), executable
    destination = c.TARGET / variant
    assert not destination.exists(), f"refusing to overwrite {destination}"
    shutil.copy2(executable, destination)
    binaries[variant] = {
        "cargo_name": binary_name,
        "features": PLAN["build"]["features"],
        "rustflags": rustflags,
        "artifact": c.artifact(destination),
    }
    c.write(OUT / "commands.json", rows)
    print(variant, "built", flush=True)

current_source = c.source()
assert current_source["files"] == frozen["source"]["files"]
current_tool = c.tool_source()
assert current_tool == frozen["tool"]
result = {
    "schema": "litchi.performance.0828.build.v1",
    "source": c.artifact(source_path),
    "probe": c.artifact(probe_path),
    "frozen_inputs": c.artifact(frozen_copy),
    "root_inputs": frozen["root_inputs"],
    "locks": frozen["locks"],
    "architecture": frozen["architecture"],
    "corpus": frozen["corpus"],
    "provenance": frozen["provenance"],
    "host": frozen["host"],
    "unrelated": frozen["unrelated"],
    "source_revision": current_source["revision"],
    "probe_source": frozen["probe"],
    "binaries": binaries,
    "rows": rows,
    "target": str(c.TARGET),
    "profile": PLAN["build"],
    "environment_contract": {
        "offline": True,
        "locked": True,
        "release": True,
        "jobs": PLAN["build"]["jobs"],
        "preserve_binaries": True,
    },
}
c.write(OUT / "build.json", result)
c.write(P / "build.json", result)
c.stable(frozen)
print("0828 build complete: ordinary and fp binaries", flush=True)
