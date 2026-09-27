"""Root-only correctness tests for the frozen standalone probe oracle."""
import os
import subprocess
import time

import custody as c

out = c.P / "probe-tests"
assert not out.exists()
out.mkdir()
build = c.read(c.P / "build-before/build.json")
assert c.source() == c.read(c.P / "build-before/source.json")
assert all(c.sha(c.P / name) == digest for name, digest in build["probe"].items())
base = ["cargo", "test", "--release", "--offline", "--locked",
        "--manifest-path", str(c.P / "probe-src/Cargo.toml")]
# Synthetic COUNTERS tests must not also observe the real global allocator's
# allocations (including their own barriers/threads). Test those with default
# features, then exercise the real wrapper and oracle with all features.
commands = [base + ["--", "--test-threads=1"],
            base + ["--all-features", "--", "--test-threads=1",
                    "--skip", "allocation_metrics::tests::"]]
env = os.environ | {"CARGO_TARGET_DIR": str(c.TARGET),
                    "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0"}
rows = []
for index, command in enumerate(commands):
    log = out / f"{index:02}.log"
    started = time.time()
    with log.open("w") as stream:
        result = subprocess.run(command, cwd=c.ROOT, env=env,
                                stdout=stream, stderr=subprocess.STDOUT)
    rows.append({"command": command, "exit_code": result.returncode,
                 "started": started, "ended": time.time(), "log": c.artifact(log)})
    c.write(out / "checks.json", rows)
    assert result.returncode == 0, log
c.write(out / "receipt.json", {
    "rows": rows,
    "probe": build["probe"], "source": c.artifact(c.P / "build-before/source.json"),
    "lock": c.artifact(c.P / "probe-src/Cargo.lock"),
    "environment": {name: env[name] for name in
                    ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL")},
})
assert c.source() == c.read(c.P / "build-before/source.json")
assert all(c.sha(c.P / name) == digest for name, digest in build["probe"].items())
print("Probe oracle tests complete", flush=True)
