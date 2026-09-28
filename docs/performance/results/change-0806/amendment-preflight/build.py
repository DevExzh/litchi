"""Root-owned native-only build for the 0806 quality amendment preflight."""

from __future__ import annotations

import os
import shutil
import subprocess
import time
from pathlib import Path

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")


def source_files(leg: str) -> dict[str, str]:
    return c.archive_manifest(leg)


def probe_files() -> dict[str, str]:
    return c.probe_manifest()


def require_probe_matches_archive(leg: str) -> None:
    expected = P / "source" / leg / "litchi-opc-xml_attributes.rs"
    actual = P / "probe-src" / "src" / ("baseline.rs" if leg == "before" else "candidate.rs")
    if c.sha(expected) != c.sha(actual):
        raise AssertionError(f"{leg}: probe helper is not the source archive")


def main() -> None:
    leg = os.environ.get("AMENDMENT_LEG")
    if leg is None:
        import sys

        if len(sys.argv) != 2:
            raise SystemExit("usage: build.py before|after")
        leg = sys.argv[1]
    if leg not in ("before", "after"):
        raise SystemExit("leg must be before or after")
    out = P / f"build-{leg}"
    if out.exists():
        raise AssertionError(f"build output already exists: {out}")
    out.mkdir()

    workspace_source = c.source()
    c.write(P / "source.json", workspace_source) if not (P / "source.json").exists() else None
    manifest = P / "probe-src" / "Cargo.toml"
    template = P / "probe-src" / "Cargo.toml.template"
    expected_manifest = template.read_text(encoding="utf-8").replace("@SRC@", str(c.ROOT))
    if manifest.exists():
        if manifest.is_symlink() or manifest.read_text(encoding="utf-8") != expected_manifest:
            raise AssertionError(f"materialized manifest changed: {manifest}")
    else:
        manifest.write_text(expected_manifest, encoding="utf-8")

    require_probe_matches_archive(leg)
    archives = {name: source_files(name) for name in ("before", "after")}
    probe = probe_files()
    frozen = {
        "schema": "litchi.performance.0806.amendment-build-inputs.v1",
        "leg": leg,
        "plan": c.sha(P / "plan.json"),
        "build_driver": c.sha(Path(__file__)),
        "quality_driver": c.sha(P / "quality.py"),
        "capture_driver": c.sha(P / "capture.py"),
        "custody_driver": c.sha(P / "custody.py"),
        "archives": archives,
        "probe": probe,
        "workspace_source": workspace_source,
    }
    c.write(out / "frozen-inputs.json", frozen)
    env = os.environ | {
        "CARGO_TARGET_DIR": str(c.TARGET),
        "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0",
    }
    command = [
        "cargo", "build", "--offline", "--locked", "--release",
        "--manifest-path", str(manifest), "--bin", "attribute-boundary-probe",
    ]
    log = out / "native.log"
    started = time.time()
    with log.open("w", encoding="utf-8") as stream:
        result = subprocess.run(command, cwd=c.ROOT, env=env, stdout=stream, stderr=subprocess.STDOUT)
    row = {
        "command": command,
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "log": c.relative_artifact(log),
    }
    c.write(out / "commands.json", [row])
    if result.returncode != 0:
        raise SystemExit(f"native build failed: {log}")
    if probe != probe_files() or archives != {name: source_files(name) for name in ("before", "after")}:
        raise AssertionError("frozen source or probe input changed during build")
    built = c.TARGET / "release" / "attribute-boundary-probe"
    if not built.is_file():
        raise AssertionError(f"missing native binary: {built}")
    destination = c.TARGET / f"{leg}-native"
    if destination.exists():
        raise AssertionError(f"binary destination already exists: {destination}")
    shutil.copy2(built, destination)
    c.write(out / "build.json", {
        "schema": "litchi.performance.0806.amendment-build.v1",
        "leg": leg,
        "source": workspace_source,
        "archive": archives[leg],
        "probe": probe,
        "binary": c.artifact(destination),
        "lock": c.artifact(P / "probe-src" / "Cargo.lock"),
        "command": row,
        "environment": {key: env.get(key) for key in ("CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL", "RUSTFLAGS")},
        "profiles": False,
        "callgrind": False,
    })


if __name__ == "__main__":
    main()
