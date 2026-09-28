"""Root-only six-gate quality run for the 0817 baseline harness."""

from __future__ import annotations

import os
import subprocess
import time

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.assert_origin()
assert PLAN["schema"] == "litchi.performance.0817.plan.v1"
assert PLAN["base"] == ORIGIN["base"]
c.plan_cases(PLAN)
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
c.write(OUT / "frozen-inputs.json", {
    "schema": "litchi.performance.0817.frozen-inputs.v1",
    "packet": packet, "drivers": drivers, "root_inputs": root_inputs,
    "locks": locks, "architecture": architecture, "corpus": corpus,
    "host": host, "unrelated": unrelated,
})

manifest = str(c.TOOL / "Cargo.toml")
commands = [
    ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
    [
        "cargo", "check", "--offline", "--locked", "--manifest-path", manifest,
        "--all-features", "--all-targets",
    ],
    ["python3", "-B", str(P / "quality_tests.py"), str(OUT)],
    [
        "cargo", "clippy", "--offline", "--locked", "--manifest-path", manifest,
        "--all-features", "--all-targets", "--", "-D", "warnings",
    ],
    [
        "cargo", "doc", "--offline", "--locked", "--manifest-path", manifest,
        "--all-features", "--no-deps",
    ],
    ["python3", "-B", "tools/check_crate_boundaries.py"],
]

build_config = PLAN["build"]
assert build_config["jobs"] == 2
for name in (
    "RUSTUP_TOOLCHAIN",
    "RUSTFLAGS",
    "CARGO_ENCODED_RUSTFLAGS",
    "CARGO_BUILD_RUSTFLAGS",
    "CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS",
    "RUSTC_WRAPPER",
    "RUSTC_WORKSPACE_WRAPPER",
):
    assert not os.environ.get(name), f"unexpected inherited quality override: {name}"
env = os.environ | {
    "CARGO_TARGET_DIR": str(c.TARGET / "quality-1"),
    "CARGO_BUILD_JOBS": str(build_config["jobs"]),
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
            "schema": "litchi.performance.0817.quality-failure.v1",
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
    c.check_stable(
        frozen, root_inputs, locks, architecture, corpus, host,
        packet, drivers, unrelated,
    )
    print("quality gate", index + 1, "PASS", flush=True)

summary = {
    "schema": "litchi.performance.0817.quality.v1",
    "status": "pass",
    "attempt": attempt,
    "source": c.artifact(source_path),
    "checks": c.artifact(OUT / "checks.json"),
    "rows": rows,
    "gate_count": 6,
    "root_inputs": root_inputs,
    "locks": locks,
    "architecture": architecture,
    "corpus": corpus,
    "provenance": provenance,
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
print("0817 quality complete", flush=True)
