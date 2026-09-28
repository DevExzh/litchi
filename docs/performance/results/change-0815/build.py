"""Root-only serial builds with exact source, probe, candidate, and binary custody."""

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
assert plan["schema"] == "litchi.performance.0815.v1"
assert plan["source_allowlist"] == ["crates/litchi-pptx/src/notes/codec.rs"]
assert origin["base"] == "55bb2ead3498043dd22b53555507b24402486248"
root_inputs = c.assert_root_inputs()

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

quality_reuse = c.read(P / "quality-reuse.json")
assert quality_reuse["schema"] == "litchi.performance.0815.quality-reuse.v1"
assert quality_reuse["source_files_equal"] is True
assert c.read(quality_reuse["source"]["path"])["files"] == c.read(
    P.parent / "change-0813/build-after/source.json"
)["files"]

architecture = c.architecture_hashes()
unrelated = c.assert_unrelated()
assert (P / "candidate/manifest.json").is_file()
assert (P / "candidate/candidate.patch").is_file()

OUT.mkdir()
manifest = P / "probe-src/Cargo.toml"
template = P / "probe-src/Cargo.toml.template"
assert manifest.read_text() == template.read_text().replace("@SRC@", str(c.ROOT))
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
    "inputs/root-Cargo.lock",
    "inputs/rustfmt.toml",
    "quality_reuse.py",
    "quality-reuse.json",
    "quality-reuse/reuse-inputs.json",
    "codegen.py",
    "toolchain.json",
    "codegen_analysis.py",
    "source-review.md",
    "protocol-review.md",
]
frozen_names.extend(
    str(path.relative_to(P))
    for path in sorted((P / "candidate").rglob("*"))
    if path.is_file()
)
c.write(
    OUT / "frozen-inputs.json",
    {
        "schema": "litchi.performance.0815.frozen-inputs.v1",
        "packet": {name: c.sha(P / name) for name in frozen_names},
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
    },
)
probe_files = c.assert_probe()
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
    assert c.assert_probe(probe_files) == probe_files
    assert c.assert_root_inputs() == root_inputs
    assert c.architecture_hashes() == architecture
    assert c.assert_unrelated() == unrelated
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
        "schema": f"litchi.performance.0815.build-{LEG}.v1",
        "source": c.artifact(OUT / "source.json"),
        "probe": probe_files,
        "lock": c.artifact(P / "probe-src/Cargo.lock"),
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
        "frozen_inputs": c.artifact(OUT / "frozen-inputs.json"),
        "binaries": binaries,
        "rows": rows,
        "environment": {
            key: env.get(key)
            for key in ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL")
        },
    },
)
