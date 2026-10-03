#!/usr/bin/env python3
"""Build the exact reviewed candidate over the retained baseline source."""
import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import subprocess
import tarfile

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BASE = P.parent / "change-0885"
SOURCE = Path("/tmp/litchi-0887-source")
TARGET = ROOT / "target/litchi-0887-build"
CHANGED = ("crates/litchi-ooxml-common/src/mce/codec.rs", "crates/litchi-ooxml-common/src/mce/expanded_duplicates_tests.rs")

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def desc(path):
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}

def write(path, value):
    with path.open("x") as f:
        json.dump(value, f, indent=2, sort_keys=True)
        f.write("\n")

def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()

if SOURCE.exists() or TARGET.exists():
    raise ValueError("owned build paths already exist; inspect instead of restarting")
if subprocess.check_output(["git", "diff", "335a40350b", "HEAD", "--", "crates", "tools/perf-baseline"], cwd=ROOT):
    raise ValueError("committed runtime/harness differs from baseline source")
baseline = json.loads((BASE / "build.json").read_text())
manifest = json.loads((BASE / "source.json").read_text())
archive = BASE / "source.tar.gz"
if sha(archive) != baseline["source_archive"]["sha256"]:
    raise ValueError("baseline source archive changed")
patch = subprocess.check_output(["git", "diff", "--", *CHANGED], cwd=ROOT)
if patch != (P / "candidate.patch").read_bytes():
    raise ValueError("reviewed candidate patch changed")
SOURCE.mkdir()
TARGET.mkdir()
for path in (SOURCE, TARGET):
    (path / ".owner-0887").write_text("MCE attribute candidate 0887\n")
with tarfile.open(archive) as handle:
    handle.extractall(SOURCE, filter="data")
overlays = {}
for name in CHANGED:
    shutil.copy2(ROOT / name, SOURCE / name)
    overlays[name] = desc(SOURCE / name)
expected = {name: overlays.get(name, record)["sha256"] for name, record in manifest["files"].items()}
for name, digest in expected.items():
    if sha(SOURCE / name) != digest:
        raise ValueError("candidate source mismatch: " + name)
write(P / "candidate-source.json", {"baseline_revision": baseline["revision"], "baseline_manifest": desc(BASE / "source.json"),
    "baseline_archive": desc(archive), "patch": desc(P / "candidate.patch"), "overlays": overlays,
    "verified_source_files": len(expected), "source_root": str(SOURCE)})
env = {k: os.environ[k] for k in ("PATH", "HOME", "CARGO_HOME", "RUSTUP_HOME") if k in os.environ}
env.update({"LANG": "C", "LC_ALL": "C", "TZ": "UTC", "CARGO_TARGET_DIR": str(TARGET),
    "CARGO_BUILD_JOBS": "2", "CARGO_NET_OFFLINE": "true", "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3", "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "CARGO_PROFILE_RELEASE_LTO": "fat", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1"})
record = {"schema": "litchi.mce-inline-attrs-build-0887.v1", "status": "building", "environment": env,
    "source_root": str(SOURCE), "target_root": str(TARGET), "source": desc(P / "candidate-source.json"),
    "compiler": subprocess.check_output(["rustc", "--version", "--verbose"], env=env, text=True), "binaries": {}}
write(P / "build-start.json", record)
(P / "binaries").mkdir()
for role, name in (("native", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")):
    argv = ["cargo", "build", "--release", "--locked", "--offline", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", name, "-j", "2"]
    if role == "allocator":
        argv += ["--features", "allocator-metrics"]
    receipt = {"argv": argv, "cwd": str(SOURCE), "started_utc": now(), "environment": env}
    write(P / (role + "-build-start.json"), receipt)
    print(role + " candidate build starting", flush=True)
    log = P / (role + "-build.log")
    with log.open("xb") as out:
        code = subprocess.run(argv, cwd=SOURCE, env=env, stdout=out, stderr=subprocess.STDOUT).returncode
    receipt.update(exit_code=code, finished_utc=now(), log=desc(log))
    write(P / (role + "-build.json"), receipt)
    if code:
        raise SystemExit(code)
    target = P / "binaries" / role
    shutil.copy2(TARGET / "release" / name, target)
    shutil.copy2(TARGET / "release" / (name + ".d"), P / (role + ".d"))
    record["binaries"][role] = {**desc(target), "executable": True, "features": [] if role == "native" else ["allocator-metrics"]}
    print(role + " candidate retained", flush=True)
for name, digest in expected.items():
    if sha(SOURCE / name) != digest:
        raise ValueError("source changed during build: " + name)
record.update(status="pass", source_rechecked_files=len(expected))
write(P / "build.json", record)
