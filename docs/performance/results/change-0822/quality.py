"""Root-owned quality gate for the 0822 packet.

The production result is replayed through only the three committed 0821
analysis loaders named in ``origin.json``.  The packet-local public probe gets
a fresh formatting, check, test, Clippy, and documentation run.  This driver
does not build a workload or run a timing capture.
"""

from __future__ import annotations

import importlib.util
import os
import re
import subprocess
import sys
import time
from pathlib import Path

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
assert PLAN["schema"] == "litchi.performance.0822.pptx-edit-profile.v1"
assert ORIGIN["schema"] == "litchi.performance.0822.origin.v1"
assert PLAN["base"] == ORIGIN["base"] == c.PREVIOUS_COMMIT
assert not (P / "quality.json").exists(), "refusing to overwrite quality.json"
assert not (P / "quality-0").exists(), "refusing to overwrite quality-0"
assert not c.TARGET.exists(), f"refusing stale target directory {c.TARGET}"
c.check_no_overrides()


def import_0821_analyze():
    """Load 0821's module with its own custody module, never this packet's."""
    previous = c.PREVIOUS_PACKET / "analyze.py"
    assert previous.is_file()
    old_module = sys.modules.pop("custody", None)
    sys.path.insert(0, str(c.PREVIOUS_PACKET))
    try:
        spec = importlib.util.spec_from_file_location("litchi_perf_0821_analyze", previous)
        assert spec is not None and spec.loader is not None
        module = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(module)
        return module
    finally:
        sys.path.remove(str(c.PREVIOUS_PACKET))
        if old_module is not None:
            sys.modules["custody"] = old_module


def production_reuse() -> dict[str, object]:
    """Replay exactly load_custody/load_build/load_quality from 0821."""
    module = import_0821_analyze()
    previous_custody = module.load_custody()
    previous_build = module.load_build(previous_custody)
    previous_quality = module.load_quality(previous_custody, previous_build)
    raw = previous_quality["raw"]
    assert raw["status"] == "pass" and raw["gate_count"] == 6
    reuse_receipts = previous_quality["reuse"]
    reuse = ORIGIN["quality_reuse"]
    assert reuse["loaders"] == ["load_custody", "load_build", "load_quality"]
    assert reuse["cargo_commands_executed"] is False
    return {
        "schema": "litchi.performance.0822.quality-reuse.v1",
        "mode": reuse["mode"],
        "loaders": reuse["loaders"],
        "cargo_commands_executed": False,
        "source": previous_custody["current_source"],
        "tool": previous_custody["current_tool"],
        "prior_quality": reuse_receipts["previous_quality"],
        "prior_build": previous_build["receipt"],
        "prior_quality_source": reuse_receipts["previous_source"],
        "prior_test_summary": reuse_receipts["previous_tests"],
        "prior_repair_inputs": reuse_receipts["previous_inputs"],
        "prior_seal": reuse_receipts["previous_seal"],
        "test_counts": ORIGIN["quality_reuse"]["required_test_counts"],
        "loader_module": c.artifact(c.PREVIOUS_PACKET / "analyze.py"),
    }


def test_counts(log: Path) -> list[dict[str, int]]:
    rows = []
    pattern = re.compile(r"(\d+) passed; (\d+) failed(?:; (\d+) ignored)?")
    for match in pattern.finditer(log.read_text(errors="replace")):
        rows.append({
            "passed": int(match.group(1)),
            "failed": int(match.group(2)),
            "ignored": int(match.group(3) or 0),
        })
    return rows


def main() -> None:
    production = production_reuse()
    frozen = c.freeze()
    out = P / "quality-0"
    out.mkdir()
    frozen_path = out / "frozen-inputs.json"
    c.write(frozen_path, frozen)
    source_path = out / "source.json"
    c.write(source_path, {"production": frozen["source"], "tool": frozen["tool"]})
    reuse_path = out / "production-reuse.json"
    c.write(reuse_path, production)

    manifest = P / "probe-src/Cargo.toml"
    assert manifest.is_file() and not manifest.is_symlink()
    assert c.probe_files()
    target = c.TARGET / "quality"
    env = os.environ.copy()
    env.update({
        "CARGO_TARGET_DIR": str(target),
        "CARGO_BUILD_JOBS": str(PLAN["build"]["jobs"]),
        "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(PLAN["build"]["opt_level"]),
        "CARGO_PROFILE_RELEASE_DEBUG": str(PLAN["build"]["debug"]),
        "CARGO_PROFILE_RELEASE_LTO": str(PLAN["build"]["lto"]),
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(PLAN["build"]["codegen_units"]),
        "CARGO_PROFILE_RELEASE_INCREMENTAL": "false",
        "CARGO_PROFILE_RELEASE_PANIC": str(PLAN["build"]["panic"]),
        "RUSTDOCFLAGS": "-Dwarnings",
        "PYTHONDONTWRITEBYTECODE": "1",
    })
    commands = [
        ("fmt", ["cargo", "fmt", "--manifest-path", str(manifest), "--", "--check"]),
        ("check", [
            "cargo", "check", "--offline", "--locked", "--release",
            "--manifest-path", str(manifest),
        ]),
        ("tests", [
            "cargo", "test", "--offline", "--locked", "--release",
            "--manifest-path", str(manifest), "--all-features", "--",
            "--test-threads=1",
        ]),
        ("clippy", [
            "cargo", "clippy", "--offline", "--locked", "--release",
            "--manifest-path", str(manifest), "--all-features", "--all-targets",
            "--", "-D", "warnings",
        ]),
        ("doc", [
            "cargo", "doc", "--offline", "--locked", "--release",
            "--manifest-path", str(manifest), "--no-deps",
        ]),
    ]
    rows: list[dict[str, object]] = []
    checks_path = out / "checks.json"
    for index, (name, command) in enumerate(commands, 1):
        log = out / f"{index:02d}-{name}.log"
        assert not log.exists()
        started = time.time()
        with log.open("w") as stream:
            result = subprocess.run(
                command,
                cwd=c.ROOT,
                env=env,
                stdout=stream,
                stderr=subprocess.STDOUT,
            )
        row = {
            "gate": index,
            "name": name,
            "command": command,
            "environment": {
                key: env.get(key)
                for key in (
                    "CARGO_TARGET_DIR", "CARGO_BUILD_JOBS", "CARGO_INCREMENTAL",
                    "CARGO_PROFILE_RELEASE_OPT_LEVEL", "CARGO_PROFILE_RELEASE_DEBUG",
                    "CARGO_PROFILE_RELEASE_LTO", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS",
                    "CARGO_PROFILE_RELEASE_INCREMENTAL", "CARGO_PROFILE_RELEASE_PANIC",
                    "RUSTDOCFLAGS",
                    "PYTHONDONTWRITEBYTECODE",
                )
            },
            "started": started,
            "ended": time.time(),
            "exit_code": result.returncode,
            "log": c.artifact(log),
        }
        if name == "tests" and result.returncode == 0:
            row["test_counts"] = test_counts(log)
        rows.append(row)
        c.write(checks_path, {
            "schema": "litchi.performance.0822.probe-quality-checks.v1",
            "rows": rows,
            "source": c.artifact(source_path),
            "frozen_inputs": c.artifact(frozen_path),
        })
        c.stable(frozen)
        if result.returncode != 0:
            raise RuntimeError(f"0822 probe quality gate failed: {name}; retained {log}")
        print("probe quality", name, "PASS", flush=True)

    summary = {
        "schema": "litchi.performance.0822.quality.v1",
        "status": "pass",
        "production_reuse": production,
        "probe": {
            "schema": "litchi.performance.0822.probe-quality.v1",
            "status": "pass",
            "gate_count": len(rows),
            "rows": rows,
            "tests": rows[2].get("test_counts", []),
        },
        "source": c.artifact(source_path),
        "frozen_inputs": c.artifact(frozen_path),
        "checks": c.artifact(checks_path),
        "root_inputs": frozen["root_inputs"],
        "locks": frozen["locks"],
        "architecture": frozen["architecture"],
        "corpus": frozen["corpus"],
        "provenance": frozen["provenance"],
        "host": frozen["host"],
        "unrelated": frozen["unrelated"],
        "probe_source": frozen["probe"],
        "target": str(c.TARGET),
        "started": rows[0]["started"],
        "ended": time.time(),
    }
    c.write(P / "quality.json", summary)
    c.stable(frozen)
    print("0822 quality PASS: committed production reuse plus five fresh probe gates", flush=True)


if __name__ == "__main__":
    main()
