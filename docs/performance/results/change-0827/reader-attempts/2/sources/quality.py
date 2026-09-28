"""Fresh harness quality and exact-source reuse of both 0824 PPTX gates."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OLD = P.parent / "change-0824"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def artifact(path):
    return {"path": str(path.resolve()), "bytes": path.stat().st_size, "sha256": sha(path)}


def write(path, value):
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n")


def state():
    names = subprocess.check_output([
        "git", "ls-files", "-z", "--", "crates", "tools/perf-baseline", "Cargo.toml",
        "Cargo.lock", "rustfmt.toml", "rust-toolchain.toml", "clippy.toml", ".cargo/config.toml",
    ], cwd=ROOT).decode().split("\0")
    names += list(json.loads((P / "architecture-inputs.json").read_text()))
    names += list(json.loads((P / "origin.json").read_text())["unrelated"])
    return {name: sha(ROOT / name) for name in sorted(set(names)) if name}


def main():
    assert not (P / "quality.json").exists()
    assert "0826 seal PASS" in (P / "prior-seal-check.log").read_text()
    origin = json.loads((P / "origin.json").read_text())
    assert subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT).decode().strip() == origin["base"]
    snapshot = state()
    oldfreeze = json.loads((OLD / "freeze.json").read_text())
    reuse = {}
    for leg in ("before", "after"):
        receipt_path = OLD / f"quality-{leg}.json"
        receipt = json.loads(receipt_path.read_text())
        assert receipt["status"] == "pass" and len(receipt["rows"]) == 6
        source_path = ROOT / receipt["source"]
        source = json.loads(source_path.read_text())
        archives = {}
        for name, baseline in oldfreeze["source"]["files"].items():
            expected = baseline
            if name in origin["source_allowlist"]:
                archive = P / "candidate" / leg / Path(name).name
                expected = sha(archive)
                archives[name] = expected
                assert snapshot[name] == sha(P / "candidate/after" / Path(name).name)
            else:
                assert snapshot[name] == baseline, name
            assert source[name] == expected, (leg, name)
        for row in receipt["rows"]:
            log = ROOT / row["log"]
            assert row["exit_code"] == 0 and sha(log) == row["log_sha256"]
            assert log.stat().st_size == row["log_bytes"]
        reuse[leg] = {"receipt": artifact(receipt_path), "source": artifact(source_path), "archives": archives}
    for name in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "CARGO_BUILD_RUSTFLAGS", "RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER", "RUSTUP_TOOLCHAIN"):
        assert not os.environ.get(name), name
    attempt = 0
    while (P / f"quality-{attempt}").exists():
        attempt += 1
    out = P / f"quality-{attempt}"
    out.mkdir()
    write(out / "source.json", snapshot)
    env_delta = {"CARGO_TARGET_DIR": str(ROOT.parent / "litchi-target-0827/quality"),
                 "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_DEV_DEBUG": "0",
                 "RUSTDOCFLAGS": "-D warnings", "PYTHONDONTWRITEBYTECODE": "1"}
    manifest = str(ROOT / "tools/perf-baseline/Cargo.toml")
    commands = [
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "check", "--offline", "--locked", "--manifest-path", manifest, "--all-features", "--all-targets"],
        ["cargo", "test", "--offline", "--locked", "--manifest-path", manifest, "--all-features", "--", "--test-threads=2"],
        ["cargo", "clippy", "--offline", "--locked", "--manifest-path", manifest, "--all-features", "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--offline", "--locked", "--manifest-path", manifest, "--all-features", "--no-deps"],
        ["python3", "-B", "tools/check_crate_boundaries.py"],
    ]
    result = {"schema": "litchi.performance.0827.quality.v1", "status": "running",
              "source": artifact(out / "source.json"), "environment": env_delta,
              "driver_sha256": sha(Path(__file__)), "pptx_reuse": reuse, "rows": [], "gate_count": 6}
    write(out / "receipt.json", result)
    for index, command in enumerate(commands):
        log = out / f"{index:02}.log"
        started = time.time()
        with log.open("x") as stream:
            completed = subprocess.run(command, cwd=ROOT, env=os.environ | env_delta, stdout=stream, stderr=subprocess.STDOUT)
        result["rows"].append({"command": command, "started": started, "ended": time.time(),
                               "exit_code": completed.returncode, "log": artifact(log)})
        result["status"] = "running" if completed.returncode == 0 else "failed"
        write(out / "receipt.json", result)
        assert completed.returncode == 0, log
        assert state() == snapshot, "quality source changed during gate"
        print("0827 quality", index + 1, "PASS", flush=True)
    result["status"] = "pass"
    write(out / "receipt.json", result)
    write(P / "quality.json", result)


if __name__ == "__main__":
    main()
