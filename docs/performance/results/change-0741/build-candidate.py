"""Build candidate lanes serially without touching the captured baseline."""
import os
import subprocess
import time

from build import P, ROOT, census, sha, write
import json


def guard_inputs():
    for name in ("constraints.json", "workspace-inputs.json"):
        record = json.loads((P / name).read_text())
        files = record["files"] if name == "workspace-inputs.json" else record
        for path, digest in files.items():
            assert sha(ROOT / path) == digest, path
    for row in json.loads((P / "build.json").read_text()):
        from pathlib import Path
        binary = Path(row["binary"])
        assert sha(binary) == row["binary_sha256"], binary
        assert binary.stat().st_size == row["binary_bytes"], binary


if __name__ == "__main__":
    guard_inputs()
    initial = census()
    # New owner sources must be staged so the inherited tracked-source census
    # cannot accidentally omit a newly introduced test or implementation module.
    untracked = subprocess.check_output(
        ["git", "ls-files", "--others", "--exclude-standard", "-z", "crates", "tools/perf-baseline"],
        cwd=ROOT,
    ).decode().split("\0")
    assert not [path for path in untracked if path.endswith(".rs")], untracked
    target = ROOT.parent / "litchi-target-0741-candidate"
    attempt = 0
    while (P / f"candidate-build-{attempt}").exists():
        attempt += 1
    folder = P / f"candidate-build-{attempt}"
    folder.mkdir()
    source = {
        "head": subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT, text=True).strip(),
        "files": initial,
    }
    (folder / "source.json").write_text(json.dumps(source, indent=2) + "\n")
    env = os.environ | {"CARGO_TARGET_DIR": str(target), "CARGO_BUILD_JOBS": "2"}
    rows = []
    for lane, extra in (("native", []), ("allocation", ["--features", "allocator-metrics"])):
        binary = "litchi-perf-baseline" + ("-alloc" if extra else "")
        command = ["cargo", "build", "--release", "--offline", "--locked",
                   "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", binary, *extra]
        log = folder / f"{lane}.log"
        start = time.time()
        with log.open("w") as output:
            result = subprocess.run(command, cwd=ROOT, env=env, stdout=output, stderr=subprocess.STDOUT)
        row = {"lane": lane, "command": command, "started": start, "ended": time.time(),
               "exit": result.returncode, "log": str(log.relative_to(P)), "log_sha256": sha(log)}
        if result.returncode == 0:
            exe = target / "release" / binary
            row |= {"binary": str(exe), "binary_sha256": sha(exe), "binary_bytes": exe.stat().st_size}
        rows.append(row)
        (folder / "build.json").write_text(json.dumps(rows, indent=2) + "\n")
        assert result.returncode == 0, row
        assert census() == initial, "source changed during candidate build"
        guard_inputs()
        print(f"PASS candidate {lane} build", flush=True)
    write("candidate-source.json", source)
    write("candidate-build.json", rows)
