"""Root-owned serial quality gates; every attempt and failing log is retained."""
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / "litchi-target-0823" / "quality"
SOURCE = "crates/litchi-pptx/src/shape/reader.rs"


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write(path, data):
    path.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n")


def state():
    paths = subprocess.check_output(
        ["git", "ls-files", "-z", "crates", "Cargo.toml", "Cargo.lock", "rust-toolchain.toml", "rustfmt.toml", "clippy.toml"], cwd=ROOT
    ).decode().split("\0")
    paths += list(json.loads((P / "architecture-inputs.json").read_text()))
    paths += ["docs/FORMAT_IMPLEMENTATION_REVIEW.md", "docs/UNIFIED_OPS_API_DESIGN.md", "matrix-analysis.json"]
    paths += [str(x.relative_to(ROOT)) for family in ("synthetic", "real")
              for x in (P / f"{family}-probe-src").rglob("*") if x.is_file()]
    return {x: sha(ROOT / x) for x in sorted(set(paths)) if x and (ROOT / x).is_file()}


def main():
    assert len(sys.argv) == 2 and sys.argv[1] in ("before", "after", "probes")
    leg = sys.argv[1]
    plan = json.loads((P / "plan.json").read_text())
    assert subprocess.check_output(["git", "rev-parse", "HEAD"], cwd=ROOT).decode().strip() == plan["base"]
    expected = subprocess.check_output(["git", "show", f'{plan["base"]}:{SOURCE}'], cwd=ROOT) if leg != "after" else (P / "candidate/after/reader.rs").read_bytes()
    assert (ROOT / SOURCE).read_bytes() == expected
    for name in ("RUSTFLAGS", "CARGO_ENCODED_RUSTFLAGS", "RUSTC_WRAPPER", "RUSTC_WORKSPACE_WRAPPER", "RUSTUP_TOOLCHAIN"):
        assert not os.environ.get(name), name
    attempt = 0
    while (P / f"quality-{leg}-{attempt}").exists():
        attempt += 1
    out = P / f"quality-{leg}-{attempt}"
    out.mkdir()
    frozen = state()
    write(out / "source.json", frozen)
    env_delta = {"CARGO_TARGET_DIR": str(TARGET), "CARGO_BUILD_JOBS": "2",
                 "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_DEV_DEBUG": "0",
                 "RUSTDOCFLAGS": "-D warnings", "PYTHONDONTWRITEBYTECODE": "1"}
    env = os.environ | env_delta
    if leg == "probes":
        commands = []
        for family in ("synthetic", "real"):
            manifest = str(P / f"{family}-probe-src/Cargo.toml")
            commands.append(["cargo", "fmt", "--manifest-path", manifest, "--", "--check"])
            for features in ([], ["--all-features"]):
                for verb in ("check", "test", "clippy", "doc"):
                    command = ["cargo", verb, "--offline", "--locked", "--manifest-path", manifest, *features]
                    if verb in ("check", "clippy"):
                        command += ["--all-targets"]
                    if verb == "test":
                        command += ["--", "--test-threads=2"]
                    if verb == "clippy":
                        command += ["--", "-D", "warnings"]
                    if verb == "doc":
                        command += ["--no-deps"]
                    commands.append(command)
    else:
        package = ["-p", "litchi-pptx"]
        commands = [
            ["cargo", "fmt", *package, "--", "--check"],
            ["cargo", "check", "--offline", "--locked", *package, "--all-features", "--all-targets"],
            ["cargo", "test", "--offline", "--locked", *package, "--all-features", "--", "--test-threads=2"],
            ["cargo", "clippy", "--offline", "--locked", *package, "--all-features", "--all-targets", "--", "-D", "warnings"],
            ["cargo", "doc", "--offline", "--locked", *package, "--all-features", "--no-deps"],
            [sys.executable, "-B", "tools/check_crate_boundaries.py"],
        ]
    result = {"schema": "litchi.performance.0823.quality.v1", "leg": leg,
              "status": "running", "environment": env_delta,
              "driver_sha256": sha(Path(__file__)), "source": str((out / "source.json").relative_to(ROOT)), "rows": []}
    write(out / "receipt.json", result)
    for index, command in enumerate(commands):
        log = out / f"{index:02}.log"
        started = time.time()
        with log.open("x") as stream:
            completed = subprocess.run(command, cwd=ROOT, env=env, stdout=stream, stderr=subprocess.STDOUT)
        result["rows"].append({"command": command, "started": started, "ended": time.time(),
                               "exit_code": completed.returncode, "log": str(log.relative_to(ROOT)),
                               "log_bytes": log.stat().st_size, "log_sha256": sha(log)})
        result["status"] = "running" if completed.returncode == 0 else "failed"
        write(out / "receipt.json", result)
        assert completed.returncode == 0, str(log)
        assert state() == frozen, "quality inputs changed during gate"
        print(leg, index + 1, "PASS", flush=True)
    result["status"] = "pass"
    write(out / "receipt.json", result)
    final = P / f"quality-{leg}.json"
    assert not final.exists()
    write(final, result)


if __name__ == "__main__":
    main()
