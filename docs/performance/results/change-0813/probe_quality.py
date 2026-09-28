"""Root-only three-gate probe quality run for one frozen leg."""

import os
import re
import subprocess
import sys
import time

import custody as c


LEG = sys.argv[1]
assert LEG in ("before", "after")
P = c.P
OUT = P / f"probe-quality-{LEG}"
assert not OUT.exists(), f"refusing to overwrite {OUT}"
OUT.mkdir()
root_inputs = c.assert_root_inputs()
architecture = c.architecture_hashes()
unrelated = c.assert_unrelated()

build = c.read(P / f"build-{LEG}/build.json")
source = c.source()
assert c.read(build["source"]["path"]) == source
probe = {
    str(path.relative_to(P / "probe-src")): c.sha(path)
    for path in (P / "probe-src").rglob("*")
    if path.is_file()
}
assert c.assert_probe(probe) == probe
c.write(
    OUT / "inputs.json",
    {
        "source": source,
        "probe": probe,
        "driver": c.artifact(__file__),
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
    },
)

manifest = str(P / "probe-src/Cargo.toml")
env = os.environ | {
    "CARGO_TARGET_DIR": str(c.TARGET / f"probe-quality-{LEG}"),
    "CARGO_BUILD_JOBS": "2",
    "CARGO_INCREMENTAL": "0",
    "PYTHONDONTWRITEBYTECODE": "1",
}
commands = [
    ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
    ["cargo", "test", "--offline", "--locked", "--release", "--manifest-path", manifest, "--all-features", "--", "--test-threads=1"],
    ["cargo", "clippy", "--offline", "--locked", "--release", "--manifest-path", manifest, "--all-features", "--all-targets", "--", "-D", "warnings"],
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
        row["tests_passed"] = test_count
    rows.append(row)
    c.write(OUT / "receipts.json", rows)
    assert result.returncode == 0, log
    assert c.source() == source
    assert c.assert_root_inputs() == root_inputs
    assert c.architecture_hashes() == architecture
    assert c.assert_unrelated() == unrelated
    assert {
        str(path.relative_to(P / "probe-src")): c.sha(path)
        for path in (P / "probe-src").rglob("*")
        if path.is_file()
    } == probe
    print(LEG, "probe gate", index + 1, "PASS", flush=True)

c.write(
    OUT / "complete.json",
    {
        "schema": f"litchi.performance.0813.probe-quality-{LEG}.v1",
        "inputs": c.artifact(OUT / "inputs.json"),
        "receipts": c.artifact(OUT / "receipts.json"),
        "gate_count": 3,
        "tests_passed": test_count,
    },
)
