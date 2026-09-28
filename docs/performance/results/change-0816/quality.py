"""Root-only six-gate quality run for the standalone 0816 harness.

Attempts are numbered and retained.  A failed gate leaves its logs and a
failure receipt in place; a later invocation starts a new attempt rather than
overwriting that evidence.
"""

from __future__ import annotations

import os
import subprocess
import time

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
assert PLAN["schema"] == "litchi.performance.0816.plan.v1"
assert ORIGIN["base"] == "c8ff2b9f65ecd011d12fd5b1ac352ea315785b93"
assert not (P / "quality.json").exists(), "refusing to overwrite quality.json"

attempt = 0
while (P / f"quality-{attempt}").exists():
    attempt += 1
OUT = P / f"quality-{attempt}"
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

manifest = str(c.TOOL / "Cargo.toml")
commands = [
    ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
    [
        "cargo",
        "check",
        "--offline",
        "--locked",
        "--manifest-path",
        manifest,
        "--all-features",
        "--all-targets",
    ],
    [
        "cargo",
        "test",
        "--offline",
        "--locked",
        "--manifest-path",
        manifest,
        "--all-features",
        "--",
        "--test-threads=2",
    ],
    [
        "cargo",
        "clippy",
        "--offline",
        "--locked",
        "--manifest-path",
        manifest,
        "--all-features",
        "--all-targets",
        "--",
        "-D",
        "warnings",
    ],
    [
        "cargo",
        "doc",
        "--offline",
        "--locked",
        "--manifest-path",
        manifest,
        "--all-features",
        "--no-deps",
    ],
    ["python3", "-B", "tools/check_crate_boundaries.py"],
]
env = os.environ | {
    "CARGO_TARGET_DIR": str(c.TARGET / f"quality-{attempt}"),
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_DEV_DEBUG": "0",
    "RUSTDOCFLAGS": "-D warnings",
    "PYTHONDONTWRITEBYTECODE": "1",
}
rows = []


def retained_failure(index: int, row: dict[str, object]) -> None:
    c.write(
        OUT / "failure.json",
        {
            "schema": "litchi.performance.0816.quality-failure.v1",
            "attempt": attempt,
            "failed_gate": index + 1,
            "rows": rows,
            "source": c.artifact(source_path),
            "packet": packet,
            "drivers": drivers,
            "row": row,
        },
    )


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
    if result.returncode != 0:
        retained_failure(index, row)
        raise RuntimeError(f"quality gate {index + 1} failed; retained {log}")
    c.unchanged(frozen)
    assert c.assert_root_inputs() == root_inputs
    assert c.assert_tool_lock() == tool_lock
    assert c.architecture_hashes() == architecture
    assert c.assert_host(host) == host
    assert c.packet_hashes() == packet
    assert c.assert_unrelated() == unrelated
    print("quality gate", index + 1, "PASS", flush=True)

summary = {
    "schema": "litchi.performance.0816.quality.v1",
    "status": "pass",
    "attempt": attempt,
    "source": c.artifact(source_path),
    "checks": c.artifact(OUT / "checks.json"),
    "rows": rows,
    "gate_count": 6,
    "root_inputs": root_inputs,
    "tool_lock": tool_lock,
    "architecture": architecture,
    "host": host,
    "unrelated": unrelated,
    "packet": packet,
    "drivers": drivers,
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
}
c.write(OUT / "quality.json", summary)
c.write(P / "quality.json", summary)
print("0816 quality complete", flush=True)
