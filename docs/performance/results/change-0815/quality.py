"""Root-only six-gate PPTX production quality run.

The caller chooses the leg. Each leg is immutable after completion and keeps
its own logs so historical failed attempts cannot be silently recirculated.
"""

import os
import subprocess
import sys
import time

import custody as c


LEG = sys.argv[1]
assert LEG in ("before", "after")
P = c.P
OUT = P / f"quality-{LEG}"
assert not OUT.exists(), f"refusing to overwrite {OUT}"
OUT.mkdir()

plan = c.read(P / "plan.json")
assert plan["schema"] == "litchi.performance.0815.v1"
source = c.source()
root_inputs = c.assert_root_inputs()
architecture = c.architecture_hashes()
unrelated = c.assert_unrelated()
if LEG == "before":
    assert source["revision"] == c.read(P / "origin.json")["base"]
else:
    before = c.read(P / "build-before/source.json")
    assert c.changed_files(before, source) == set(plan["source_allowlist"])
c.write(OUT / "source.json", source)

packages = [
    "-p", "litchi-pptx",
]
commands = [
    ["cargo", "fmt", *packages, "--", "--check"],
    ["cargo", "check", "--offline", "--locked", *packages, "--all-features", "--all-targets"],
    ["cargo", "test", "--offline", "--locked", *packages, "--all-features", "--", "--test-threads=2"],
    ["cargo", "clippy", "--offline", "--locked", *packages, "--all-features", "--all-targets", "--", "-D", "warnings"],
    ["cargo", "doc", "--offline", "--locked", *packages, "--all-features", "--no-deps"],
    ["python3", "-B", "tools/check_crate_boundaries.py"],
]
env = os.environ | {
    "CARGO_TARGET_DIR": str(c.TARGET / f"quality-{LEG}"),
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_DEV_DEBUG": "0",
    "RUSTDOCFLAGS": "-D warnings",
    "PYTHONDONTWRITEBYTECODE": "1",
}
rows = []
for index, command in enumerate(commands):
    log = OUT / f"{index:02}.log"
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
        "gate": index + 1,
        "command": command,
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "log": c.artifact(log),
    }
    rows.append(row)
    c.write(OUT / "checks.json", rows)
    assert result.returncode == 0, log
    assert c.source() == source
    assert c.assert_root_inputs() == root_inputs
    assert c.architecture_hashes() == architecture
    assert c.assert_unrelated() == unrelated
    print(LEG, "quality gate", index + 1, "PASS", flush=True)

c.write(
    P / f"quality-{LEG}.json",
    {
        "schema": f"litchi.performance.0815.quality-{LEG}.v1",
        "source": c.artifact(OUT / "source.json"),
        "checks": c.artifact(OUT / "checks.json"),
        "rows": rows,
        "gate_count": 6,
        "environment": {
            key: env[key]
            for key in (
                "CARGO_TARGET_DIR",
                "CARGO_BUILD_JOBS",
                "CARGO_INCREMENTAL",
                "CARGO_PROFILE_DEV_DEBUG",
                "RUSTDOCFLAGS",
            )
        },
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
    },
)
