"""Fresh probe gates plus exact reuse of the committed 0827 quality packet.

The production and perf-baseline trees are not rebuilt here. The previous
packet's committed seal is replayed byte-for-byte, and its source, tool,
normative-input, lock, and receipt identities are compared with the current
workspace before this packet runs the independent probe gates.
"""

from __future__ import annotations

import hashlib
import json
import os
import re
import subprocess
import time
from pathlib import Path

import custody as c


P = c.P
ROOT = c.ROOT
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
assert PLAN["schema"] == "litchi.performance.0829.pptx-edit-profile.v1"
assert ORIGIN["schema"] == "litchi.performance.0829.origin.v1"
assert PLAN["base"] == ORIGIN["base"] == c.BASE_COMMIT
assert not (P / "quality.json").exists(), "refusing to overwrite quality.json"
assert not (P / "quality-0").exists(), "refusing to overwrite quality-0"
assert not c.TARGET.exists(), f"refusing stale target directory {c.TARGET}"
c.check_no_overrides()


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


def previous_seal_blob() -> tuple[dict[str, object], bytes]:
    raw = subprocess.check_output([
        "git", "show", f"{c.PREVIOUS_COMMIT}:docs/performance/results/change-0827/seal.json",
    ], cwd=ROOT)
    value = json.loads(raw.decode())
    assert value["schema"] == "litchi.performance.0827.seal.v1"
    files = value.get("files")
    assert isinstance(files, dict) and len(files) >= 700
    for name, expected in files.items():
        relative = Path(name)
        assert not relative.is_absolute() and ".." not in relative.parts
        blob = subprocess.check_output([
            "git", "show", f"{c.PREVIOUS_COMMIT}:{name}"
        ], cwd=ROOT)
        assert hashlib.sha256(blob).hexdigest() == expected, f"prior seal mismatch: {name}"
    local = ROOT / "docs/performance/results/change-0827/seal.json"
    assert local.is_file() and c.sha(local) == hashlib.sha256(raw).hexdigest()
    return value, raw


def production_reuse() -> dict[str, object]:
    """Verify all committed 0827 source/quality/receipt identities."""
    reuse = ORIGIN["quality_reuse"]
    assert reuse["mode"] == "exact-source-0827-quality-receipts"
    assert reuse["cargo_commands_executed"] is False
    assert reuse["prior_quality"] == "docs/performance/results/change-0827/quality.json"
    assert reuse["prior_seal"] == "docs/performance/results/change-0827/seal.json"
    assert reuse["prior_seal_commit"] == c.PREVIOUS_COMMIT
    assert reuse["required_test_counts"] == {
        "passed": 641, "failed": 0, "ignored": 1, "suites": 28,
    }

    seal, seal_bytes = previous_seal_blob()
    prior_quality_path = ROOT / reuse["prior_quality"]
    prior_freeze_path = ROOT / "docs/performance/results/change-0827/freeze.json"
    prior_quality = c.read(prior_quality_path)
    prior_freeze = c.read(prior_freeze_path)
    assert prior_quality["schema"] == "litchi.performance.0827.quality.v1"
    assert prior_quality["status"] == "pass"
    assert prior_quality["gate_count"] == 6
    assert len(prior_quality["rows"]) == 6
    assert prior_freeze["schema"] == "litchi.performance.0827.freeze.v1"

    # The committed seal proves the historical files came from the pinned
    # commit.  Also bind the two local witnesses used by this reuse check to
    # the exact sealed blobs, so a copied or edited local packet cannot pass.
    sealed_files = seal["files"]
    for relative in (
        "docs/performance/results/change-0827/quality.json",
        "docs/performance/results/change-0827/freeze.json",
    ):
        assert relative in sealed_files
        assert c.sha(ROOT / relative) == sealed_files[relative]

    current_source = c.source()
    current_tool = c.tool_source()
    current_architecture = c.architecture_hashes()
    current_root_inputs = c.assert_root_inputs()
    current_locks = c.lock_identity()
    current_corpus = c.assert_corpus_inputs()
    current_provenance = c.assert_provenance(current_corpus)
    current_host = c.assert_host()
    current_unrelated = c.assert_unrelated()

    # The 0827 freeze records the source revision before its final report
    # commit.  The production file census is the reusable identity; the
    # revision itself advances to this packet's pinned base commit.
    assert current_source["files"] == prior_freeze["source"]["files"]
    assert current_tool == prior_freeze["tool"]
    assert current_architecture == prior_freeze["architecture"]
    assert current_root_inputs["Cargo.lock"] == prior_freeze["root_inputs"]["Cargo.lock"]
    assert current_root_inputs["rustfmt.toml"] == prior_freeze["root_inputs"]["rustfmt.toml"]
    assert current_locks["tool"]["sha256"] == prior_freeze["root_inputs"]["tools/perf-baseline/Cargo.lock"]
    prior_pptx = prior_freeze["corpus"]["test-data/ooxml/pptx/shapes.pptx"]
    assert {
        "bytes": current_corpus["source"]["bytes"],
        "sha256": current_corpus["source"]["sha256"],
    } == prior_pptx
    assert current_unrelated == prior_freeze["unrelated"]
    assert prior_quality["pptx_reuse"]

    # Retain the exact source receipt used by the historical quality run and
    # verify every entry against the current workspace.  The freeze census is
    # the authoritative production/tool split; the receipt additionally
    # includes the normative and unrelated files recorded by 0827.
    prior_source_path = c.verify_descriptor(prior_quality["source"])
    prior_source = c.read(prior_source_path)
    assert isinstance(prior_source, dict)
    expected_receipt = {**prior_freeze["source"]["files"], **prior_freeze["tool"]}
    for name, expected in expected_receipt.items():
        assert prior_source.get(name) == expected, f"prior source receipt changed: {name}"
    for name, expected in prior_source.items():
        path = ROOT / name
        assert path.is_file() and not path.is_symlink()
        assert c.sha(path) == expected, f"current source differs from prior receipt: {name}"

    manifest = str(ROOT / "tools/perf-baseline/Cargo.toml")
    expected_commands = [
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "check", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets"],
        ["cargo", "test", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--", "--test-threads=2"],
        ["cargo", "clippy", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--no-deps"],
        ["python3", "-B", "tools/check_crate_boundaries.py"],
    ]
    retained: list[dict[str, object]] = []
    for row, command in zip(prior_quality["rows"], expected_commands, strict=True):
        assert row["command"] == command, "prior quality command changed"
        log = c.verify_descriptor(row["log"])
        assert row["exit_code"] == 0
        retained.append(c.artifact(log))
    assert len(retained) == 6

    test_log = c.verify_descriptor(prior_quality["rows"][2]["log"])
    parsed = test_counts(test_log)
    assert len(parsed) == 28
    totals = {
        key: sum(row[key] for row in parsed)
        for key in ("passed", "failed", "ignored")
    }
    assert totals == {"passed": 641, "failed": 0, "ignored": 1}
    historical_counts = {
        **totals,
        "suites": len(parsed),
    }
    assert historical_counts == reuse["required_test_counts"]

    # The nested 0824 source/receipt witnesses are part of the committed
    # 0827 reuse claim; hash-check them before carrying that claim forward.
    for leg in ("before", "after"):
        leg_value = prior_quality["pptx_reuse"].get(leg)
        assert isinstance(leg_value, dict)
        c.verify_descriptor(leg_value["receipt"])
        leg_source_path = c.verify_descriptor(leg_value["source"])
        leg_source = c.read(leg_source_path)
        assert isinstance(leg_source, dict) and leg_source

    return {
        "schema": "litchi.performance.0829.quality-reuse.v1",
        "mode": reuse["mode"],
        "cargo_commands_executed": False,
        "prior_quality": c.artifact(prior_quality_path),
        "prior_freeze": c.artifact(prior_freeze_path),
        "prior_seal": c.artifact(ROOT / reuse["prior_seal"]),
        "prior_seal_commit": c.PREVIOUS_COMMIT,
        "prior_seal_bytes": len(seal_bytes),
        "prior_seal_files": len(seal["files"]),
        "source": current_source,
        "tool": current_tool,
        "architecture": current_architecture,
        "root_inputs": current_root_inputs,
        "locks": current_locks,
        "corpus": current_corpus,
        "provenance": current_provenance,
        "host": current_host,
        "unrelated": current_unrelated,
        "prior_quality_receipts": retained,
        "test_counts": historical_counts,
    }


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
                command, cwd=ROOT, env=env, stdout=stream, stderr=subprocess.STDOUT,
            )
        row: dict[str, object] = {
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
                    "RUSTDOCFLAGS", "PYTHONDONTWRITEBYTECODE",
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
            "schema": "litchi.performance.0829.probe-quality-checks.v1",
            "rows": rows,
            "source": c.artifact(source_path),
            "frozen_inputs": c.artifact(frozen_path),
        })
        c.stable(frozen)
        if result.returncode != 0:
            raise RuntimeError(f"0829 probe quality gate failed: {name}; retained {log}")
        print("0829 probe quality", name, "PASS", flush=True)

    summary = {
        "schema": "litchi.performance.0829.quality.v1",
        "status": "pass",
        "production_reuse": production,
        "probe": {
            "schema": "litchi.performance.0829.probe-quality.v1",
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
    print("0829 quality PASS: exact 0827 production reuse plus five fresh probe gates", flush=True)


if __name__ == "__main__":
    main()
