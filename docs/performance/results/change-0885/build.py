#!/usr/bin/env python3
"""Build retained release baseline binaries from an exact Git source snapshot."""
import datetime
import gzip
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
import subprocess
import tarfile

ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
SOURCE = Path("/tmp/litchi-0885-source")
TARGET = ROOT / "target/litchi-0885-build"
REVISION = "335a40350b"

def now():
    return datetime.datetime.now(datetime.timezone.utc).isoformat()

def sha(path):
    h = hashlib.sha256()
    with path.open("rb") as handle:
        while block := handle.read(1024 * 1024):
            h.update(block)
    return h.hexdigest()

def descriptor(path):
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}

def write(path, value):
    with path.open("x") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")

if SOURCE.exists() or TARGET.exists():
    raise SystemExit("owned source or target already exists; inspect prior execution")
SOURCE.mkdir()
TARGET.mkdir(parents=True)
for path in (SOURCE, TARGET):
    (path / ".owner-0885").write_text("performance baseline 0885\n")
revision = subprocess.check_output(["git", "rev-parse", REVISION], cwd=ROOT, text=True).strip()
paths = [".cargo", "Cargo.toml", "README.md", "clippy.toml", "deny.toml", "rust-toolchain.toml", "rustfmt.toml", "crates", "test-data", "tools/perf-baseline"]
temporary_tar = TARGET / "source.tar"
with temporary_tar.open("xb") as handle:
    subprocess.run(["git", "archive", "--format=tar", revision, "--", *paths], cwd=ROOT, stdout=handle, check=True)
archive = PACKET / "source.tar.gz"
with temporary_tar.open("rb") as src, archive.open("xb") as raw:
    with gzip.GzipFile(fileobj=raw, mode="wb", mtime=0) as dst:
        shutil.copyfileobj(src, dst)
with tarfile.open(temporary_tar) as handle:
    handle.extractall(SOURCE, filter="data")
temporary_tar.unlink()
files = {}
for path in sorted(SOURCE.rglob("*")):
    if path.is_file() and path.name != ".owner-0885":
        if path.is_symlink():
            raise RuntimeError("unexpected source symlink: " + str(path))
        files[path.relative_to(SOURCE).as_posix()] = {"bytes": path.stat().st_size, "sha256": sha(path), "mode": stat.S_IMODE(path.stat().st_mode)}
write(PACKET / "source.json", {"revision": revision, "selection": paths, "file_count": len(files), "files": files, "archive": descriptor(archive)})
env = {key: os.environ[key] for key in ("PATH", "HOME", "CARGO_HOME", "RUSTUP_HOME") if key in os.environ}
env.update({"LANG": "C", "LC_ALL": "C", "TZ": "UTC", "CARGO_TARGET_DIR": str(TARGET), "CARGO_BUILD_JOBS": "2", "CARGO_NET_OFFLINE": "true", "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3", "CARGO_PROFILE_RELEASE_DEBUG": "1", "CARGO_PROFILE_RELEASE_LTO": "fat", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1"})
build = {"schema": "litchi.xlsx-full-cell-release-build-0885.v1", "status": "building", "revision": revision, "source_root": str(SOURCE), "target_root": str(TARGET), "source_manifest": descriptor(PACKET / "source.json"), "source_archive": descriptor(archive), "environment": env, "compiler": subprocess.check_output(["rustc", "--version", "--verbose"], cwd=SOURCE, env=env, text=True), "cargo": subprocess.check_output(["cargo", "--version"], cwd=SOURCE, env=env, text=True).strip(), "binaries": {}}
write(PACKET / "build-start.json", build)
(PACKET / "binaries").mkdir()
for role, name in (("native", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")):
    argv = ["cargo", "build", "--release", "--locked", "--offline", "--manifest-path", "tools/perf-baseline/Cargo.toml", "--bin", name, "-j", "2"]
    if role == "allocator":
        argv += ["--features", "allocator-metrics"]
    record = {"argv": argv, "cwd": str(SOURCE), "environment": env, "started_utc": now()}
    write(PACKET / (role + "-build-start.json"), record)
    log = PACKET / (role + "-build.log")
    print(role + " build starting", flush=True)
    with log.open("xb") as handle:
        code = subprocess.run(argv, cwd=SOURCE, env=env, stdout=handle, stderr=subprocess.STDOUT).returncode
    record.update(exit_code=code, finished_utc=now(), log=descriptor(log))
    write(PACKET / (role + "-build.json"), record)
    if code:
        raise SystemExit(code)
    executable = TARGET / "release" / name
    retained = PACKET / "binaries" / role
    shutil.copy2(executable, retained)
    depfile = TARGET / "release" / (name + ".d")
    shutil.copy2(depfile, PACKET / (role + ".d"))
    build["binaries"][role] = {**descriptor(retained), "executable": True, "features": [] if role == "native" else ["allocator-metrics"], "build_receipt": descriptor(PACKET / (role + "-build.json")), "depfile": descriptor(PACKET / (role + ".d"))}
    print(role + " build retained", flush=True)
for name, expected in files.items():
    path = SOURCE / name
    if sha(path) != expected["sha256"] or path.stat().st_size != expected["bytes"]:
        raise RuntimeError("source changed during build: " + name)
build["status"] = "pass"
build["source_rechecked_files"] = len(files)
write(PACKET / "build.json", build)
print("both release binaries retained", flush=True)
