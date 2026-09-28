"""Root-owned serial release builds for the 0825 before/after legs.

The driver only builds and copies the three packet binaries.  It never runs a
workload, and every failed Cargo invocation leaves its log and partial receipt
in the leg directory.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time

import custody as c


assert len(sys.argv) == 2 and sys.argv[1] in {"before", "after"}
LEG = sys.argv[1]
PLAN = c.plan()
FROZEN = c.read(c.P / "freeze.json")
OUT = c.P / f"build-{LEG}"
assert not OUT.exists(), f"refusing to overwrite {OUT}"
assert FROZEN["schema"] == "litchi.performance.0825.freeze.v1"
c.assert_static()
c.assert_quality()
c.stable_inputs(FROZEN)
c.check_no_overrides()
current = c.assert_leg_source(LEG, FROZEN)

if LEG == "before":
    assert c.TARGET.is_dir(), f"owned target is missing: {c.TARGET}"
    assert (c.TARGET / "quality").is_dir()
    assert not any(path.name.startswith("before-") or path.name.startswith("after-")
                   for path in c.TARGET.iterdir()), (
        "baseline release binaries already exist in the owned target"
    )
else:
    assert c.TARGET.is_dir(), f"owned target is missing: {c.TARGET}"
    before_source_path = c.P / "build-before/source.json"
    assert before_source_path.is_file()
    before_source = c.read(before_source_path)
    assert c.changed_files(before_source, current) == set(c.ALLOWLIST)

OUT.mkdir()
source_path = OUT / "source.json"
c.write(source_path, current)
frozen_path = OUT / "frozen-inputs.json"
c.write(frozen_path, FROZEN)

build_config = PLAN["build"]
env = os.environ.copy()
env.update({
    "CARGO_TARGET_DIR": str(c.TARGET),
    "CARGO_BUILD_JOBS": str(build_config["jobs"]),
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(build_config["opt_level"]),
    "CARGO_PROFILE_RELEASE_DEBUG": str(build_config["debug"]),
    "CARGO_PROFILE_RELEASE_LTO": str(build_config["lto"]),
    "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(build_config["codegen_units"]),
    "CARGO_PROFILE_RELEASE_INCREMENTAL": "false",
    "CARGO_PROFILE_RELEASE_PANIC": str(build_config["panic"]),
    "PYTHONDONTWRITEBYTECODE": "1",
})
manifest = c.TOOL / "Cargo.toml"
steps = (
    ("native", PLAN["binaries"]["native"]),
    ("artifacts", PLAN["binaries"]["artifacts"]),
    ("observer", PLAN["binaries"]["observer"]),
)
rows: list[dict[str, object]] = []
binaries: dict[str, dict[str, object]] = {}

for logical, spec in steps:
    binary_name = spec["cargo_bin"]
    features = spec["features"]
    command = [
        "cargo", "build", "--offline", "--locked", "--release",
        "--manifest-path", str(manifest), "--bin", binary_name,
    ]
    if features:
        command.extend(["--features", ",".join(features)])
    log = OUT / f"{logical}.log"
    assert not log.exists()
    started = time.time()
    with log.open("x", encoding="utf-8") as stream:
        completed = subprocess.run(
            command, cwd=c.ROOT, env=env, stdout=stream, stderr=subprocess.STDOUT,
        )
    row = {
        "schema": "litchi.performance.0825.build-receipt.v1",
        "leg": LEG,
        "binary": logical,
        "cargo_bin": binary_name,
        "features": features,
        "command": command,
        "started": started,
        "ended": time.time(),
        "exit_code": completed.returncode,
        "log": c.artifact(log),
        "environment": {
            key: env.get(key) for key in (
                "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                "CARGO_PROFILE_RELEASE_OPT_LEVEL", "CARGO_PROFILE_RELEASE_DEBUG",
                "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
                "CARGO_PROFILE_RELEASE_INCREMENTAL", "CARGO_PROFILE_RELEASE_PANIC",
                "PYTHONDONTWRITEBYTECODE",
            )
        },
    }
    rows.append(row)
    c.write(OUT / "commands.json", rows)
    if completed.returncode != 0:
        c.write(OUT / "failure.json", {
            "schema": "litchi.performance.0825.build-failure.v1",
            "failed": row,
            "rows": rows,
            "source": c.artifact(source_path),
            "frozen_inputs": c.artifact(frozen_path),
        })
        raise RuntimeError(f"0825 {LEG}/{logical} build failed; retained {log}")
    executable = c.TARGET / "release" / binary_name
    assert executable.is_file() and not executable.is_symlink(), executable
    destination = c.TARGET / f"{LEG}-{logical}"
    assert not destination.exists(), f"refusing binary overwrite: {destination}"
    shutil.copy2(executable, destination)
    binaries[logical] = {
        "cargo_bin": binary_name,
        "features": features,
        "artifact": c.artifact(destination),
    }
    c.stable_inputs(FROZEN)
    assert c.assert_leg_source(LEG, FROZEN) == current
    print(f"0825 {LEG}/{logical} built", flush=True)

result = {
    "schema": f"litchi.performance.0825.build-{LEG}.v1",
    "leg": LEG,
    "source": c.artifact(source_path),
    "frozen_inputs": c.artifact(frozen_path),
    "quality": FROZEN["quality"],
    "root_inputs": FROZEN["root_inputs"],
    "architecture": FROZEN["architecture"],
    "corpus": FROZEN["corpus"],
    "unrelated": FROZEN["unrelated"],
    "candidate_archives": FROZEN["candidate_archives"],
    "binaries": binaries,
    "rows": rows,
    "target": str(c.TARGET),
    "profile": build_config,
    "environment_contract": {
        "offline": True, "locked": True, "release": True,
        "jobs": build_config["jobs"], "serial": True,
    },
}
c.write(OUT / "build.json", result)
c.stable_inputs(FROZEN)
print(f"0825 {LEG} build PASS: three serial native/artifact/observer binaries", flush=True)
