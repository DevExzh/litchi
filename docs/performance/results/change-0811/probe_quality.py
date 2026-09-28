"""Root-only three-gate quality run for the six-file public probe."""

import os
import re
import subprocess
import time

import custody as c


P = c.P
OUT = P / "probe-quality"
assert not OUT.exists(), f"refusing to overwrite {OUT}"
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
BUILD = c.read(P / "build/build.json")
FROZEN = c.read(BUILD["frozen_inputs"]["path"])
assert PLAN["schema"] == "litchi.performance.0811.current-production.v1"
assert BUILD["schema"] == "litchi.performance.0811.build.v1"
assert c.frozen_driver_hashes() == FROZEN["drivers"]
assert ORIGIN["base"] == PLAN["source_revision"]
source = c.source()
assert c.read(BUILD["source"]["path"]) == source
assert source["revision"] == ORIGIN["base"]
assert len(source["files"]) == PLAN["source_file_count"]
probe = c.assert_probe(BUILD["probe"])
root_inputs = c.assert_root_inputs()
architecture = c.architecture_hashes()
unrelated = c.assert_unrelated()
OUT.mkdir()
c.write(
    OUT / "inputs.json",
    {
        "source": source,
        "probe": probe,
        "driver": c.artifact(__file__),
        "build": c.artifact(P / "build/build.json"),
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
    },
)

manifest = str(P / "probe-src/Cargo.toml")
env = os.environ | {
    "CARGO_TARGET_DIR": str(c.TARGET / "probe-quality"),
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
    "PYTHONDONTWRITEBYTECODE": "1",
}
commands = [
    ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
    [
        "cargo",
        "test",
        "--offline",
        "--locked",
        "--release",
        "--manifest-path",
        manifest,
        "--all-features",
        "--",
        "--test-threads=1",
    ],
    [
        "cargo",
        "clippy",
        "--offline",
        "--locked",
        "--release",
        "--manifest-path",
        manifest,
        "--all-features",
        "--all-targets",
        "--",
        "-D",
        "warnings",
    ],
]
rows = []
test_count = None
for index, command in enumerate(commands):
    log = OUT / f"{index}.log"
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
    if index == 1 and result.returncode == 0:
        matches = re.findall(r"(\d+) passed; (\d+) failed", log.read_text())
        counts = [(int(passed), int(failed)) for passed, failed in matches]
        assert counts and max(passed for passed, _ in counts) == 36
        assert all(failed == 0 for _, failed in counts)
        test_count = 36
        row["test_counts"] = counts
    rows.append(row)
    c.write(OUT / "receipts.json", rows)
    assert result.returncode == 0, log
    assert c.source() == source
    assert c.assert_probe(probe) == probe
    assert c.assert_root_inputs() == root_inputs
    assert c.architecture_hashes() == architecture
    assert c.assert_unrelated() == unrelated
    print("probe quality gate", index + 1, "PASS", flush=True)

c.write(
    P / "probe-quality.json",
    {
        "schema": "litchi.performance.0811.probe-quality.v1",
        "inputs": c.artifact(OUT / "inputs.json"),
        "receipts": c.artifact(OUT / "receipts.json"),
        "gate_count": 3,
        "tests_passed": test_count,
        "target": str(c.TARGET / "probe-quality"),
    },
)
c.write(
    OUT / "complete.json",
    {
        "schema": "litchi.performance.0811.probe-quality.complete.v1",
        "inputs": c.artifact(OUT / "inputs.json"),
        "receipts": c.artifact(OUT / "receipts.json"),
        "gate_count": 3,
        "tests_passed": test_count,
    },
)
