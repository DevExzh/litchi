#!/usr/bin/env python3
"""Build the 0890 MCE candidate over the retained 0885 source archive.

This helper deliberately builds only the candidate.  The before executables
are the retained 0885 release binaries.  The current 0889 diagnostic source
is excluded explicitly because it is a separate binary and is not linked into
the release workload.  Candidate edits are restricted to the OOXML MCE owner
and its tests; a changed manifest, dependency, harness, or diagnostic path
fails closed.

The helper is create-only: it refuses existing owned source/target/artifact
paths and must be run only after the candidate patch has been frozen.
"""

from __future__ import annotations

import datetime
import hashlib
import json
import os
from pathlib import Path
import shutil
import stat
import subprocess
import re
import tarfile


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
BASELINE = PACKET.parent / "change-0885"
BASELINE_REVISION = "335a40350bf06f49691745c14d310fd7a3ca9117"
SOURCE = Path("/tmp/litchi-0890-source")
TARGET = ROOT / "target/litchi-0890-build"
ARCHIVE = BASELINE / "source.tar.gz"
MANIFEST = BASELINE / "source.json"
FREEZE = PACKET / "candidate-freeze-v2.json"
PATCH = PACKET / "candidate-v2.patch"
CURRENT_BASE_REVISION = "f790c2228cbd9e89c2892ba3ffaeb49802fac43b"
SELECTION = (".cargo", "Cargo.toml", "README.md", "clippy.toml", "deny.toml", "rust-toolchain.toml", "rustfmt.toml", "crates", "test-data", "tools/perf-baseline")
DIAGNOSTIC_PATHS = frozenset(
    {
        "tools/perf-baseline/src/bin/xml_attribute_bounds.rs",
        "tools/perf-baseline/src/bin/xml_attribute_bounds/boundary_tests.rs",
        "tools/perf-baseline/src/bin/xml_attribute_bounds/input_tests.rs",
        "tools/perf-baseline/src/bin/xml_attribute_bounds/qualification_tests.rs",
    }
)
CANDIDATE_ROOT = "crates/litchi-ooxml-common/src/mce/"


def now() -> str:
    return datetime.datetime.now(datetime.timezone.utc).isoformat()


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, object]:
    info = path.stat()
    if not stat.S_ISREG(info.st_mode):
        raise ValueError(f"expected regular file: {path}")
    return {
        "path": str(path.resolve()),
        "bytes": info.st_size,
        "sha256": sha(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def write(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def git(*args: str, text: bool = True) -> str | bytes:
    return subprocess.check_output(["git", *args], cwd=ROOT, text=text)


def _patch_paths(value: bytes) -> set[str]:
    paths: set[str] = set()
    for line in value.decode("utf-8").splitlines():
        match = re.fullmatch(r"diff --git a/(.+) b/(.+)", line)
        if match is None:
            continue
        left, right = match.groups()
        path = right if left == "/dev/null" else left
        if left != right and left != "/dev/null":
            raise RuntimeError(f"candidate patch has asymmetric paths: {left!r}, {right!r}")
        if Path(path).is_absolute() or ".." in Path(path).parts:
            raise RuntimeError(f"candidate patch has an unsafe path: {path!r}")
        paths.add(path)
    if not paths:
        raise RuntimeError("candidate patch has no paths")
    return paths


def _load_frozen_candidate(manifest: dict[str, object]) -> tuple[dict[str, object], tuple[str, ...], dict[str, dict[str, object]]]:
    """Validate the already sealed v2 source inventory without writing it.

    The freeze is deliberately separate from this builder.  In particular,
    this function never derives a patch from the worktree and never creates a
    replacement source record.  The explicit ``files`` list is also the
    inventory for a future untracked regression file; relying only on
    ``git diff`` would silently omit such an overlay.
    """

    if not FREEZE.is_file() or not PATCH.is_file():
        raise RuntimeError("candidate-freeze-v2.json and candidate-v2.patch must be frozen before build")
    freeze = json.loads(FREEZE.read_text(encoding="utf-8"))
    if freeze.get("schema") != "litchi.mce-candidate-freeze-0890.v1":
        raise RuntimeError("candidate freeze schema is not the reviewed 0890 schema")
    head = str(git("rev-parse", "HEAD")).strip()
    if freeze.get("base_commit") != CURRENT_BASE_REVISION or freeze.get("base_commit") != head:
        raise RuntimeError("candidate freeze is not bound to the current f790c2228c base")
    paths = freeze.get("candidate_paths")
    entries = freeze.get("files")
    if not isinstance(paths, list) or not paths or len(set(paths)) != len(paths) or not all(isinstance(path, str) for path in paths):
        raise RuntimeError("candidate freeze has no unique explicit candidate path inventory")
    if any(not path.startswith(CANDIDATE_ROOT) for path in paths):
        raise RuntimeError("candidate freeze contains a path outside the MCE owner")
    if not isinstance(entries, list) or not entries or len(entries) != len(paths):
        raise RuntimeError("candidate freeze file inventory is incomplete")
    frozen_files: dict[str, dict[str, object]] = {}
    for entry in entries:
        if not isinstance(entry, dict) or not isinstance(entry.get("path"), str):
            raise RuntimeError("candidate freeze contains an invalid file descriptor")
        name = entry["path"]
        if name in frozen_files or name not in paths:
            raise RuntimeError(f"candidate freeze file inventory has an unexpected or duplicate path: {name}")
        if not isinstance(entry.get("bytes"), int) or entry["bytes"] < 0 or not isinstance(entry.get("sha256"), str) or len(entry["sha256"]) != 64:
            raise RuntimeError(f"candidate freeze file identity is incomplete: {name}")
        frozen_files[name] = entry
    if set(paths) != set(frozen_files):
        raise RuntimeError("candidate freeze path and file inventories differ")
    if set(paths) != _patch_paths(PATCH.read_bytes()):
        raise RuntimeError("candidate-v2.patch paths differ from the frozen candidate inventory")
    patch_record = freeze.get("patch")
    if not isinstance(patch_record, dict) or patch_record.get("path") != "docs/performance/results/change-0890/candidate-v2.patch":
        raise RuntimeError("candidate freeze does not bind candidate-v2.patch")
    if patch_record.get("sha256") != sha(PATCH) or patch_record.get("bytes") != PATCH.stat().st_size:
        raise RuntimeError("candidate-v2.patch identity differs from candidate-freeze-v2.json")
    for path in paths:
        current = ROOT / path
        if not current.is_file() or current.is_symlink():
            raise RuntimeError(f"candidate overlay is not a regular file: {path}")
        record = frozen_files[path]
        if record.get("sha256") != sha(current) or record.get("bytes") != current.stat().st_size:
            raise RuntimeError(f"current candidate overlay differs from frozen identity: {path}")
    tracked = {
        path
        for path in str(git("diff", "--name-only", BASELINE_REVISION, "--", *SELECTION)).splitlines()
        if path
    }
    untracked_mce = {
        path
        for path in str(git("ls-files", "--others", "--exclude-standard", "--", CANDIDATE_ROOT)).splitlines()
        if path
    }
    allowed = set(paths) | DIAGNOSTIC_PATHS
    unexpected = sorted(tracked - allowed)
    if unexpected:
        raise RuntimeError("source differs outside the explicit MCE candidate and diagnostic exclusion: " + ", ".join(unexpected))
    if untracked_mce - set(paths):
        raise RuntimeError("untracked MCE source is absent from frozen candidate inventory: " + ", ".join(sorted(untracked_mce - set(paths))))
    expected_files = set(manifest["files"]) | set(paths)
    if len(expected_files) != len(manifest["files"]) + len(set(paths) - set(manifest["files"])):
        raise RuntimeError("candidate source file census is not the baseline/overlay union")
    return freeze, tuple(sorted(paths)), frozen_files


def main() -> int:
    if SOURCE.exists() or TARGET.exists():
        raise RuntimeError("owned 0890 source or target already exists; inspect prior execution")
    if not ARCHIVE.is_file() or not MANIFEST.is_file():
        raise RuntimeError("retained 0885 source archive/manifest is missing")
    if sha(ARCHIVE) != "bdf6e1c815fbb33796cc779134d27a311cc25ba5043bb4f5e5a73fc4d5643bee":
        raise RuntimeError("retained 0885 source archive changed")
    manifest = json.loads(MANIFEST.read_text(encoding="utf-8"))
    if manifest.get("revision") != BASELINE_REVISION or manifest.get("file_count") != 10765:
        raise RuntimeError("retained 0885 source manifest identity mismatch")
    freeze, paths, frozen_files = _load_frozen_candidate(manifest)

    SOURCE.mkdir()
    TARGET.mkdir()
    for path in (SOURCE, TARGET):
        (path / ".owner-0890").write_text("MCE candidate 0890\n", encoding="utf-8")
    with tarfile.open(ARCHIVE) as handle:
        handle.extractall(SOURCE, filter="data")

    overlays: dict[str, dict[str, object]] = {}
    for name in paths:
        destination = SOURCE / name
        destination.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(ROOT / name, destination)
        overlays[name] = descriptor(destination)

    expected = {name: overlays.get(name, frozen_files.get(name, record))["sha256"] for name, record in manifest["files"].items()}
    for name in paths:
        expected[name] = overlays[name]["sha256"]
    for name, digest in expected.items():
        if sha(SOURCE / name) != digest:
            raise RuntimeError(f"candidate source mismatch: {name}")

    env = {key: os.environ[key] for key in ("PATH", "HOME", "CARGO_HOME", "RUSTUP_HOME") if key in os.environ}
    env.update(
        {
            "LANG": "C",
            "LC_ALL": "C",
            "TZ": "UTC",
            "CARGO_TARGET_DIR": str(TARGET),
            "CARGO_BUILD_JOBS": "2",
            "CARGO_NET_OFFLINE": "true",
            "CARGO_INCREMENTAL": "0",
            "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
            "CARGO_PROFILE_RELEASE_DEBUG": "1",
            "CARGO_PROFILE_RELEASE_LTO": "fat",
            "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1",
        }
    )
    build = {
        "schema": "litchi.mce-candidate-build-0890.v1",
        "status": "building",
        "baseline_revision": BASELINE_REVISION,
        "source": descriptor(FREEZE),
        "source_root": str(SOURCE),
        "target_root": str(TARGET),
        "environment": env,
        "compiler": subprocess.check_output(["rustc", "--version", "--verbose"], cwd=SOURCE, env=env, text=True),
        "cargo": subprocess.check_output(["cargo", "--version"], cwd=SOURCE, env=env, text=True).strip(),
        "binaries": {},
    }
    write(PACKET / "build-start.json", build)
    (PACKET / "binaries").mkdir()
    for role, name in (("native", "litchi-perf-baseline"), ("allocator", "litchi-perf-baseline-alloc")):
        argv = [
            "cargo",
            "build",
            "--release",
            "--locked",
            "--offline",
            "--manifest-path",
            "tools/perf-baseline/Cargo.toml",
            "--bin",
            name,
            "-j",
            "2",
        ]
        if role == "allocator":
            argv += ["--features", "allocator-metrics"]
        receipt = {"argv": argv, "cwd": str(SOURCE), "environment": env, "started_utc": now()}
        write(PACKET / f"{role}-build-start.json", receipt)
        log = PACKET / f"{role}-build.log"
        with log.open("xb") as handle:
            code = subprocess.run(argv, cwd=SOURCE, env=env, stdout=handle, stderr=subprocess.STDOUT).returncode
        receipt.update({"exit_code": code, "finished_utc": now(), "log": descriptor(log)})
        write(PACKET / f"{role}-build.json", receipt)
        if code:
            raise SystemExit(code)
        executable = TARGET / "release" / name
        retained = PACKET / "binaries" / role
        shutil.copy2(executable, retained)
        depfile = TARGET / "release" / f"{name}.d"
        shutil.copy2(depfile, PACKET / f"{role}.d")
        build["binaries"][role] = {
            **descriptor(retained),
            "features": [] if role == "native" else ["allocator-metrics"],
            "executable": True,
            "build_receipt": descriptor(PACKET / f"{role}-build.json"),
            "depfile": descriptor(PACKET / f"{role}.d"),
        }

    for name, digest in expected.items():
        if sha(SOURCE / name) != digest:
            raise RuntimeError(f"source changed during candidate build: {name}")
    build.update({"status": "pass", "source_rechecked_files": len(expected)})
    write(PACKET / "build.json", build)
    print("0890 candidate release binaries retained")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
