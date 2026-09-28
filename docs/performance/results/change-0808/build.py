"""Root-only serial builds with exact source, probe, and binary custody."""

import os
import shutil
import subprocess
import sys
import time

import custody as c


LEG = sys.argv[1]
assert LEG in ("before", "after")
P = c.P
OUT = P / f"build-{LEG}"
assert not OUT.exists(), f"refusing to overwrite {OUT}"
if LEG == "before":
    assert not c.TARGET.exists(), f"refusing an existing target {c.TARGET}"
else:
    assert c.TARGET.exists(), f"the before target is required for the after leg"
    prior = c.read(P / "build-before/build.json")
    for binary in prior["binaries"].values():
        assert c.artifact(binary["path"]) == binary

plan = c.read(P / "plan.json")
origin = c.read(P / "origin.json")
assert plan["schema"] == "litchi.performance.0808.v1"
assert plan["source_allowlist"] == ["crates/litchi-pptx/src/notes/codec.rs"]
assert origin["base"] == "d28e3dc702d84f8752a2e329c2d3fc15f5c94f06"

if LEG == "before":
    clean = subprocess.run(
        [
            "git",
            "diff",
            "--quiet",
            "HEAD",
            "--",
            "crates",
            "Cargo.toml",
            "clippy.toml",
            ".cargo/config.toml",
            "rust-toolchain.toml",
        ],
        cwd=c.ROOT,
    )
    assert clean.returncode == 0, "production source is modified before the baseline build"
else:
    before = c.read(P / "build-before/source.json")
    current = c.source()
    assert c.changed_files(before, current) == set(plan["source_allowlist"])

OUT.mkdir()
manifest = P / "probe-src/Cargo.toml"
template = P / "probe-src/Cargo.toml.template"
manifest.write_text(template.read_text().replace("@SRC@", str(c.ROOT)))
source = c.source()
assert len(source["files"]) == 9196
if LEG == "before":
    assert source["revision"] == origin["base"]
c.write(OUT / "source.json", source)

frozen_names = [
    "plan.json",
    "adoption-policy.json",
    "analysis-plan.json",
    "custody.py",
    "build.py",
    "capture.py",
    "profile.py",
    "quality.py",
    "probe_quality.py",
    "apply_candidate.py",
    "restore_candidate.py",
    "origin.json",
    "host.json",
    "inheritance.json",
    "architecture-inputs.json",
]
c.write(OUT / "frozen-inputs.json", {name: c.sha(P / name) for name in frozen_names})
probe_files = {
    str(path.relative_to(P)): c.sha(path)
    for path in (P / "probe-src").rglob("*")
    if path.is_file()
}
c.write(OUT / "probe.json", probe_files)

env = os.environ | {
    "CARGO_TARGET_DIR": str(c.TARGET),
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
}
rows = []
binaries = {}
for name, features in (
    ("native", []),
    ("allocation", ["--features", "allocator-metrics"]),
    ("profile", ["--features", "capture-profile"]),
):
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
        "name": name,
        "command": command,
        "exit_code": result.returncode,
        "started": started,
        "ended": time.time(),
        "log": c.artifact(log),
    }
    rows.append(row)
    c.write(OUT / "commands.json", rows)
    assert result.returncode == 0, log
    assert c.source() == source
    assert {
        str(path.relative_to(P)): c.sha(path)
        for path in (P / "probe-src").rglob("*")
        if path.is_file()
    } == probe_files
    executable = c.TARGET / "release/namespace-uri-probe"
    assert executable.is_file()
    destination = c.TARGET / f"{LEG}-{name}"
    assert not destination.exists()
    shutil.copy2(executable, destination)
    binaries[name] = c.artifact(destination)
    print(LEG, name, "built", flush=True)

c.write(
    OUT / "build.json",
    {
        "schema": f"litchi.performance.0808.build-{LEG}.v1",
        "source": c.artifact(OUT / "source.json"),
        "probe": probe_files,
        "lock": c.artifact(P / "probe-src/Cargo.lock"),
        "binaries": binaries,
        "rows": rows,
        "environment": {
            key: env.get(key)
            for key in ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL")
        },
    },
)
