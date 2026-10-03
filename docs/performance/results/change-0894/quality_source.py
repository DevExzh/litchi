#!/usr/bin/env python3
"""Prepare an owned quality tree without consuming the shared ignored lock."""
from pathlib import Path
import datetime
import hashlib
import json
import os
import shutil
import subprocess
import sys
import tarfile

PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
DEST = ROOT / "target/litchi-0894-quality-src"
REVISION = "506aa2edae6666d9c76ae6827c2f5941a8e47896"
PATHS = (
    "crates/litchi-xlsx/src/raw/worksheet/selected.rs",
    "crates/litchi-xlsx/src/raw/worksheet/selected/producer_metadata.rs",
    "crates/litchi-xlsx/src/raw/worksheet/selected/producer_metadata_tests.rs",
    "crates/litchi-xlsx/src/raw/worksheet/tests.rs",
    "tools/perf-baseline/src/producer_shape/selected_route_tests.rs",
)
SUPPORT = (
    "tools/check_crate_boundaries.py",
    "tools/crate_boundaries.json",
    "docs/adr/0008-migration-and-verification.md",
    "docs/adr/0028-iwa-monolith-exit.md",
)


def identity(path):
    assert path.is_file() and not path.is_symlink(), path
    return {"path": str(path), "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}


def main():
    assert not DEST.exists(), DEST
    manifest = json.loads((PACKET / "source.json").read_text())
    archive = PACKET / "source.tar.gz"
    assert identity(archive)["sha256"] == manifest["archive"]["sha256"]
    DEST.mkdir()
    (DEST / ".owner-litchi-0894-quality-src").write_text("litchi 0894 isolated quality source\n")
    with tarfile.open(archive) as tar:
        for item in tar:
            assert not Path(item.name).is_absolute() and ".." not in Path(item.name).parts
            assert item.isdir() or item.isfile()
        tar.extractall(DEST, filter="data")
    for name, expected in manifest["files"].items():
        assert identity(DEST / name)["sha256"] == expected["sha256"], name
    overlays = {}
    for name in PATHS:
        target = DEST / name
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / name, target)
        overlays[name] = identity(target)
        assert overlays[name]["sha256"] == identity(ROOT / name)["sha256"]
    support = {}
    for name in SUPPORT:
        target = DEST / name
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(subprocess.check_output(["git", "show", f"{REVISION}:{name}"], cwd=ROOT))
        support[name] = identity(target)
    temporary = ROOT / "target/litchi-0894-quality/private-tmp"
    assert temporary.is_dir()
    env = {key: os.environ[key] for key in ("PATH", "HOME", "CARGO_HOME", "RUSTUP_HOME") if key in os.environ}
    env.update(RUSTUP_TOOLCHAIN="1.95.0", CARGO_NET_OFFLINE="true", TMPDIR=str(temporary), LANG="C", LC_ALL="C")
    argv = ["cargo", "generate-lockfile", "--offline", "--manifest-path", str(DEST / "Cargo.toml")]
    with (PACKET / "quality-lock-generation.log").open("x") as log:
        result = subprocess.run(argv, cwd=DEST, env=env, stdout=log, stderr=subprocess.STDOUT)
    receipt = {"schema": "litchi.isolated-quality-source-0894.v1", "status": "pass" if result.returncode == 0 else "fail",
               "source_root": str(DEST), "baseline_revision": REVISION, "archive": identity(archive),
               "overlays": overlays, "support": support, "argv": argv, "environment": env,
               "exit_code": result.returncode, "log": identity(PACKET / "quality-lock-generation.log"),
               "utc": datetime.datetime.now(datetime.timezone.utc).isoformat(),
               "shared_root_lock_consumed": False, "lock": identity(DEST / "Cargo.lock") if result.returncode == 0 else None}
    with (PACKET / "quality-source.json").open("x") as out:
        json.dump(receipt, out, indent=2); out.write("\n")
    assert result.returncode == 0, receipt
    print("isolated quality source and offline lock prepared")


def refresh(label):
    assert label.isalnum(), label
    assert (DEST / ".owner-litchi-0894-quality-src").read_text() == "litchi 0894 isolated quality source\n"
    current = PACKET / "quality-source.json"
    previous = PACKET / f"quality-source-before-{label}.json"
    assert not previous.exists()
    receipt = json.loads(current.read_text())
    assert identity(DEST / "Cargo.lock") == receipt["lock"]
    for name, expected in receipt["overlays"].items():
        assert identity(DEST / name) == expected, name
    shutil.copy2(current, previous)
    before = receipt["overlays"]
    for name in PATHS:
        shutil.copy2(ROOT / name, DEST / name)
    receipt["overlays"] = {name: identity(DEST / name) for name in PATHS}
    receipt["refresh"] = {"label": label, "previous_receipt": identity(previous),
                          "before": before, "runner": identity(Path(__file__).resolve()),
                          "utc": datetime.datetime.now(datetime.timezone.utc).isoformat()}
    temporary = current.with_suffix(".next")
    with temporary.open("x") as out:
        json.dump(receipt, out, indent=2); out.write("\n")
    temporary.replace(current)
    print(f"isolated quality overlay refreshed: {label}")


if __name__ == "__main__":
    if len(sys.argv) == 3 and sys.argv[1] == "refresh":
        refresh(sys.argv[2])
    else:
        assert len(sys.argv) == 1
        main()
