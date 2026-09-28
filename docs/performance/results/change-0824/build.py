"""Root-owned serial release builds for both 0824 probes and legs."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import custody as c


assert len(sys.argv) == 2 and sys.argv[1] in {"before", "after"}
LEG = sys.argv[1]
PLAN = c.read(c.P / "plan.json")
FROZEN = c.read(c.P / "freeze.json")
OUT = c.P / f"build-{LEG}"
assert not OUT.exists(), f"refusing to overwrite {OUT}"
c.assert_static()
c.stable_inputs(FROZEN)
c.check_no_overrides()

current = c.source()
if LEG == "before":
    assert c.TARGET.is_dir(), f"owned target is missing: {c.TARGET}"
    assert {path.name for path in c.TARGET.iterdir()} == {"quality"}, (
        "baseline build permits only the completed quality target before release artifacts"
    )
    c.assert_quality_before()
    assert current == FROZEN["source"], "baseline source differs from the frozen source"
    assert current["revision"] == c.BASE
    for production, archive in (
        (c.ALLOWLIST[0], "candidate/before/transaction.rs"),
        (c.ALLOWLIST[1], "candidate/before/xml.rs"),
    ):
        assert c.sha(c.ROOT / production) == FROZEN["candidate"][archive]
else:
    assert c.TARGET.exists(), f"before target is required: {c.TARGET}"
    before = c.read(c.P / "build-before/source.json")
    assert c.changed_files(before, current) == set(c.ALLOWLIST)
    for production, archive in (
        (c.ALLOWLIST[0], "candidate/after/transaction.rs"),
        (c.ALLOWLIST[1], "candidate/after/xml.rs"),
    ):
        assert c.sha(c.ROOT / production) == FROZEN["candidate"][archive]

OUT.mkdir()
source_path = OUT / "source.json"
c.write(source_path, current)
frozen_path = OUT / "frozen-inputs.json"
c.write(frozen_path, FROZEN)

env = os.environ.copy()
env.update({
    "CARGO_TARGET_DIR": str(c.TARGET),
    "CARGO_BUILD_JOBS": str(PLAN["build"]["jobs"]),
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(PLAN["build"]["opt_level"]),
    "CARGO_PROFILE_RELEASE_DEBUG": str(PLAN["build"]["debug"]),
    "CARGO_PROFILE_RELEASE_LTO": str(PLAN["build"]["lto"]),
    "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(PLAN["build"]["codegen_units"]),
    "CARGO_PROFILE_RELEASE_INCREMENTAL": "false",
    "CARGO_PROFILE_RELEASE_PANIC": str(PLAN["build"]["panic"]),
    "PYTHONDONTWRITEBYTECODE": "1",
})

rows: list[dict[str, object]] = []
binaries: dict[str, dict[str, object]] = {}
for probe in ("synthetic", "real"):
    spec = PLAN["probes"][probe]
    manifest = c.P / spec["manifest"].removeprefix("docs/performance/results/change-0824/")
    assert manifest.is_file() and not manifest.is_symlink()
    assert c.probe_files(probe) == FROZEN["probes"][probe]
    for variant in ("native", "allocation"):
        features = spec["features"][variant]
        command = [
            "cargo", "build", "--offline", "--locked", "--release",
            "--manifest-path", str(manifest), "--bin", spec["cargo_binary"],
        ]
        if features:
            command.extend(["--features", ",".join(features)])
        label = f"{probe}-{variant}"
        log = OUT / f"{label}.log"
        assert not log.exists()
        started = time.time()
        with log.open("w", encoding="utf-8") as stream:
            result = subprocess.run(
                command, cwd=c.ROOT, env=env, stdout=stream, stderr=subprocess.STDOUT,
            )
        row: dict[str, object] = {
            "schema": "litchi.performance.0824.build-receipt.v1",
            "leg": LEG,
            "probe": probe,
            "variant": variant,
            "features": features,
            "command": command,
            "environment": {
                key: env.get(key) for key in (
                    "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                    "CARGO_PROFILE_RELEASE_OPT_LEVEL", "CARGO_PROFILE_RELEASE_DEBUG",
                    "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
                    "CARGO_PROFILE_RELEASE_INCREMENTAL", "CARGO_PROFILE_RELEASE_PANIC",
                    "PYTHONDONTWRITEBYTECODE",
                )
            },
            "started": started,
            "ended": time.time(),
            "exit_code": result.returncode,
            "log": c.artifact(log),
        }
        rows.append(row)
        c.write(OUT / "commands.json", rows)
        if result.returncode != 0:
            c.write(OUT / "failure.json", {
                "schema": "litchi.performance.0824.build-failure.v1",
                "failed": row,
                "rows": rows,
                "source": c.artifact(source_path),
                "frozen_inputs": c.artifact(frozen_path),
            })
            raise RuntimeError(f"0824 {label} build failed; retained {log}")
        executable = c.TARGET / "release" / spec["cargo_binary"]
        assert executable.is_file() and not executable.is_symlink(), executable
        destination = c.TARGET / f"{LEG}-{probe}-{variant}"
        assert not destination.exists(), f"refusing binary overwrite: {destination}"
        shutil.copy2(executable, destination)
        binaries[label] = {
            "probe": probe,
            "variant": variant,
            "cargo_binary": spec["cargo_binary"],
            "features": features,
            "artifact": c.artifact(destination),
        }
        c.stable_inputs(FROZEN)
        assert c.source() == current, "production source changed during build"
        print(LEG, label, "built", flush=True)

result = {
    "schema": f"litchi.performance.0824.build-{LEG}.v1",
    "leg": LEG,
    "source": c.artifact(source_path),
    "frozen_inputs": c.artifact(frozen_path),
    "root_inputs": FROZEN["root_inputs"],
    "architecture": FROZEN["architecture"],
    "unrelated": FROZEN["unrelated"],
    "probes": FROZEN["probes"],
    "binaries": binaries,
    "rows": rows,
    "target": str(c.TARGET),
    "profile": PLAN["build"],
    "environment_contract": {"offline": True, "locked": True, "release": True,
                              "jobs": PLAN["build"]["jobs"], "serial": True},
}
c.write(OUT / "build.json", result)
c.stable_inputs(FROZEN)
print(f"0824 {LEG} build PASS: four serial native/allocation binaries", flush=True)
