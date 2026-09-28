"""Root-only serial release builds for the unchanged 0814 source.

The first invocation writes the complete pre-build custody record before
starting Cargo. Three release executables are copied out of the one owned
target so later native and perf receipts can bind to immutable identities.
"""

import os
import shutil
import subprocess
import time

import custody as c


P = c.P
OUT = P / "build"
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
assert PLAN["schema"] == "litchi.performance.0814.current-production.v1"
assert PLAN["source_file_count"] == 9196
assert PLAN["source_allowlist"] == []
assert ORIGIN["base"] == PLAN["source_revision"]
assert ORIGIN["base"] == "5eb8629254fbf88009b988734af59eb90f8c6271"
assert not OUT.exists(), f"refusing to overwrite {OUT}"
assert not c.TARGET.exists(), f"refusing an existing target {c.TARGET}"

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

source = c.source()
assert source["revision"] == ORIGIN["base"]
assert len(source["files"]) == PLAN["source_file_count"]
sealed_0813_source = P.parent / "change-0813/build-after/source.json"
assert sealed_0813_source.is_file()
assert source["files"] == c.read(sealed_0813_source)["files"]
probe = c.assert_probe()
assert c.assert_root_inputs() == {
    "Cargo.lock": ORIGIN["root_inputs"]["root-Cargo.lock"],
    "rustfmt.toml": ORIGIN["root_inputs"]["rustfmt.toml"],
}
architecture = c.architecture_hashes()
unrelated = c.assert_unrelated()

template = P / "probe-src/Cargo.toml.template"
manifest = P / "probe-src/Cargo.toml"
assert manifest.read_text() == template.read_text().replace("@SRC@", str(c.ROOT))

OUT.mkdir()
frozen = {
    "schema": "litchi.performance.0814.frozen-inputs.v1",
    "source": source,
    "probe": probe,
    "root_inputs": c.root_input_hashes(),
    "architecture": architecture,
    "unrelated": unrelated,
    "sealed_0813_after_source": c.artifact(sealed_0813_source),
    "plan": c.sha(P / "plan.json"),
    "origin": c.sha(P / "origin.json"),
    "drivers": c.frozen_driver_hashes(),
}
c.write(OUT / "frozen-inputs.json", frozen)
c.write(OUT / "source.json", source)
c.write(OUT / "probe.json", probe)

base_env = os.environ.copy()
base_env.update(
    {
        "CARGO_TARGET_DIR": str(c.TARGET),
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
    }
)
rows = []
binaries = {}
variants = (
    ("control", []),
    ("profile", ["--features", "capture-profile"]),
    ("fp", ["--features", "capture-profile"]),
)
for name, features in variants:
    rustflags = None if name != "fp" else "-C force-frame-pointers=yes"
    env = base_env.copy()
    if rustflags is not None:
        env["RUSTFLAGS"] = rustflags
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
        "rustflags": rustflags,
        "command": command,
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "log": c.artifact(log),
        "environment": {
            key: env.get(key)
            for key in (
                "CARGO_TARGET_DIR",
                "CARGO_BUILD_JOBS",
                "CARGO_INCREMENTAL",
                "RUSTFLAGS",
            )
        },
    }
    rows.append(row)
    c.write(OUT / "commands.json", rows)
    assert result.returncode == 0, log
    assert c.source() == source
    assert c.assert_probe(probe) == probe
    assert c.assert_root_inputs() == frozen["root_inputs"]
    assert c.architecture_hashes() == architecture
    assert c.assert_unrelated() == unrelated
    executable = c.TARGET / "release/namespace-uri-probe"
    assert executable.is_file(), f"missing release executable for {name}"
    destination = c.TARGET / name
    assert not destination.exists()
    shutil.copy2(executable, destination)
    binary = c.artifact(destination)
    binaries[name] = binary
    c.write(OUT / "commands.json", rows)
    print(name, "built", flush=True)

c.write(
    OUT / "build.json",
    {
        "schema": "litchi.performance.0814.build.v1",
        "source": c.artifact(OUT / "source.json"),
        "probe": probe,
        "root_inputs": frozen["root_inputs"],
        "architecture": architecture,
        "unrelated": unrelated,
        "frozen_inputs": c.artifact(OUT / "frozen-inputs.json"),
        "binaries": binaries,
        "rows": rows,
        "target": str(c.TARGET),
        "environment_contract": {
            "jobs": 2,
            "offline": True,
            "locked": True,
            "release": True,
            "preserve_earlier_binaries": True
        }
    }
)
print("0814 build complete", flush=True)
