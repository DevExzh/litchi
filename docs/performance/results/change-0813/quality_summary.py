"""Independently validate the retained 0813 quality receipts.

This reader only consumes the production and probe quality artifacts.  It does
not invoke Cargo, rustfmt, a workload, or a profiler.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TEST_RESULT = re.compile(
    r"test result: (\w+)\.\s+(\d+) passed;\s+(\d+) failed;\s+"
    r"(\d+) ignored;\s+(\d+) measured;\s+(\d+) filtered out"
)
PROBE_FILES = {
    "Cargo.lock",
    "Cargo.toml",
    "Cargo.toml.template",
    "src/allocation_metrics.rs",
    "src/counting_allocator.rs",
    "src/main.rs",
}
UNRELATED = {
    "docs/FORMAT_IMPLEMENTATION_REVIEW.md": "bffd00f144c4c1bbb3b0805d21352b40e9ae46f7366c03d6e581ca61b1f27ce5",
    "docs/UNIFIED_OPS_API_DESIGN.md": "f5672c38393a2a6c52f028b2a501ddad93766ef974dbdb46cdc3c45e1db3ef6d",
    "matrix-analysis.json": "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
}


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def artifact(value: dict[str, Any], label: str) -> Path:
    require(isinstance(value, dict), f"{label}: artifact is not an object")
    path = Path(value["path"])
    if not path.is_absolute():
        path = P / path
    require(path.is_file() and not path.is_symlink(), f"{label}: missing {path}")
    require(path.stat().st_size == value["bytes"], f"{label}: byte count changed")
    require(sha(path) == value["sha256"], f"{label}: hash changed")
    return path


def packet_artifact(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing {path}")
    return {
        "path": str(path),
        "bytes": path.stat().st_size,
        "sha256": sha(path),
    }


def root_inputs() -> dict[str, str]:
    copies = {
        "Cargo.lock": P / "inputs/root-Cargo.lock",
        "rustfmt.toml": P / "inputs/rustfmt.toml",
    }
    result = {}
    for name, copy in copies.items():
        require(copy.is_file() and not copy.is_symlink(), f"missing frozen input: {copy}")
        digest = sha(copy)
        source = ROOT / name
        require(source.is_file() and not source.is_symlink(), f"missing root input: {source}")
        require(sha(source) == digest, f"root input changed: {name}")
        result[name] = digest
    return result


def architecture_inputs() -> dict[str, str]:
    declared = read_json(P / "architecture-inputs.json")
    require(len(declared) == 35, "architecture input count")
    actual = {name: sha(ROOT / name) for name in declared}
    require(actual == declared, "architecture input changed")
    return actual


def unrelated_inputs() -> dict[str, str]:
    actual = {name: sha(ROOT / name) for name in UNRELATED}
    require(actual == UNRELATED, "unrelated workspace input changed")
    return actual


def parse_tests(path: Path) -> list[dict[str, Any]]:
    rows = []
    for match in TEST_RESULT.finditer(path.read_text(encoding="utf-8", errors="replace")):
        status, passed, failed, ignored, measured, filtered = match.groups()
        rows.append(
            {
                "status": status,
                "passed": int(passed),
                "failed": int(failed),
                "ignored": int(ignored),
                "measured": int(measured),
                "filtered": int(filtered),
            }
        )
    return rows


def production_commands() -> list[list[str]]:
    return [
        ["cargo", "fmt", "-p", "litchi-pptx", "--", "--check"],
        [
            "cargo",
            "check",
            "--offline",
            "--locked",
            "-p",
            "litchi-pptx",
            "--all-features",
            "--all-targets",
        ],
        [
            "cargo",
            "test",
            "--offline",
            "--locked",
            "-p",
            "litchi-pptx",
            "--all-features",
            "--",
            "--test-threads=2",
        ],
        [
            "cargo",
            "clippy",
            "--offline",
            "--locked",
            "-p",
            "litchi-pptx",
            "--all-features",
            "--all-targets",
            "--",
            "-D",
            "warnings",
        ],
        [
            "cargo",
            "doc",
            "--offline",
            "--locked",
            "-p",
            "litchi-pptx",
            "--all-features",
            "--no-deps",
        ],
        ["python3", "-B", "tools/check_crate_boundaries.py"],
    ]


def probe_commands() -> list[list[str]]:
    manifest = str(P / "probe-src/Cargo.toml")
    return [
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        [
            "cargo",
            "test",
            "--offline",
            "--locked",
            "--release",
            "--manifest-path",
            manifest,
            "--all-features",
            "--",
            "--test-threads=1",
        ],
        [
            "cargo",
            "clippy",
            "--offline",
            "--locked",
            "--release",
            "--manifest-path",
            manifest,
            "--all-features",
            "--all-targets",
            "--",
            "-D",
            "warnings",
        ],
    ]


def candidate_source() -> dict[str, Any]:
    plan = read_json(P / "plan.json")
    allowlist = set(plan["source_allowlist"])
    before_build = read_json(P / "build-before/build.json")
    before_source = read_json(artifact(before_build["source"], "build-before source"))
    manifest = read_json(P / "candidate/manifest.json")
    require(manifest["schema"] == "litchi.performance.0813.candidate-manifest.v1", "candidate schema")
    require(manifest["base_commit"] == read_json(P / "origin.json")["base"], "candidate base")
    require(manifest["production_source_changed"] is False, "candidate archive changed production")
    rows = list(manifest.get("files", {}).values())
    if not rows:
        rows = [manifest]
    require(len(rows) == 1, "candidate manifest must contain one production file")
    row = rows[0]
    production_path = row["production_path"]
    require({production_path} == allowlist, "candidate scope differs from plan")
    before = row["before"]
    after = row["after"]
    before_path = artifact(before, "candidate before")
    after_path = artifact(after, "candidate after")
    patch = manifest["patch"]
    artifact(patch, "candidate patch")
    require(before_source["files"][production_path] == sha(before_path), "candidate before differs")
    expected = dict(before_source["files"])
    expected[production_path] = sha(after_path)
    return {
        "revision": before_source["revision"],
        "files": expected,
    }


def summarize_production(
    root_hashes: dict[str, str],
    architecture: dict[str, str],
    unrelated: dict[str, str],
    expected_source: dict[str, Any],
) -> dict[str, Any]:
    quality_path = P / "quality-after.json"
    quality = read_json(quality_path)
    require(quality["schema"] == "litchi.performance.0813.quality-after.v1", "quality schema")
    require(quality["gate_count"] == 6, "production gate count")
    require(quality["root_inputs"] == root_hashes, "production root-input hashes")
    require(quality["architecture"] == architecture, "production architecture hashes")
    require(quality["unrelated"] == unrelated, "production unrelated hashes")
    source_path = artifact(quality["source"], "quality-after source")
    source = read_json(source_path)
    require(source == expected_source, "quality-after source differs from candidate")
    checks_path = artifact(quality["checks"], "quality-after checks")
    checks = read_json(checks_path)
    require(checks == quality["rows"], "quality checks receipt differs from summary")
    expected = production_commands()
    require(len(checks) == len(expected), "production receipt row count")
    rows = []
    test_results = []
    for index, (row, command) in enumerate(zip(checks, expected), 1):
        require(row["gate"] == index, f"production gate numbering: {index}")
        require(row["command"] == command, f"production command: {index}")
        require(row["exit_code"] == 0, f"production gate failed: {index}")
        log = artifact(row["log"], f"production gate {index} log")
        result = {
            "gate": index,
            "command": command,
            "status": "pass",
            "exit_code": 0,
            "log": packet_artifact(log, f"production gate {index} log"),
        }
        if index == 3:
            test_results = parse_tests(log)
            require(len(test_results) == 85, "production test suite count")
            require(all(item["status"] == "ok" and item["failed"] == 0 for item in test_results), "production test status")
            require(sum(item["passed"] for item in test_results) == 1241, "production passed count")
            require(sum(item["failed"] for item in test_results) == 0, "production failed count")
            require(sum(item["ignored"] for item in test_results) == 3, "production ignored count")
            result["tests"] = {
                "suites": len(test_results),
                "passed": sum(item["passed"] for item in test_results),
                "failed": sum(item["failed"] for item in test_results),
                "ignored": sum(item["ignored"] for item in test_results),
            }
        rows.append(result)
    return {
        "schema": quality["schema"],
        "source": packet_artifact(source_path, "quality-after source"),
        "checks": packet_artifact(checks_path, "quality-after checks"),
        "root_inputs": root_hashes,
        "gates": rows,
        "tests": {
            "suites": len(test_results),
            "passed": sum(item["passed"] for item in test_results),
            "failed": sum(item["failed"] for item in test_results),
            "ignored": sum(item["ignored"] for item in test_results),
            "results": test_results,
        },
    }


def summarize_probe(
    leg: str, root_hashes: dict[str, str], architecture: dict[str, str], unrelated: dict[str, str]
) -> dict[str, Any]:
    directory = P / f"probe-quality-{leg}"
    complete = read_json(directory / "complete.json")
    require(complete["schema"] == f"litchi.performance.0813.probe-quality-{leg}.v1", f"probe {leg} schema")
    require(complete["gate_count"] == 3 and complete["tests_passed"] == 36, f"probe {leg} completion")
    inputs_path = artifact(complete["inputs"], f"probe {leg} inputs")
    inputs = read_json(inputs_path)
    require(inputs["root_inputs"] == root_hashes, f"probe {leg} root-input hashes")
    require(inputs["architecture"] == architecture, f"probe {leg} architecture hashes")
    require(inputs["unrelated"] == unrelated, f"probe {leg} unrelated hashes")
    driver_path = artifact(inputs["driver"], f"probe {leg} driver")
    require(driver_path == P / "probe_quality.py", f"probe {leg} driver path")
    build = read_json(P / f"build-{leg}/build.json")
    source_path = artifact(build["source"], f"build {leg} source")
    source = read_json(source_path)
    require(inputs["source"] == source, f"probe {leg} source")
    actual_probe = {
        str(path.relative_to(P / "probe-src")): sha(path)
        for path in (P / "probe-src").rglob("*")
        if path.is_file() and not path.is_symlink()
    }
    require(set(actual_probe) == PROBE_FILES, f"probe {leg} file set")
    require(inputs["probe"] == actual_probe, f"probe {leg} source files")
    receipts_path = artifact(complete["receipts"], f"probe {leg} receipts")
    receipts = read_json(receipts_path)
    require(len(receipts) == 3, f"probe {leg} receipt count")
    require(receipts == read_json(directory / "receipts.json"), f"probe {leg} receipt path")
    expected = probe_commands()
    rows = []
    test_results = []
    for index, (row, command) in enumerate(zip(receipts, expected), 1):
        require(row["gate"] == index, f"probe {leg} gate numbering")
        require(row["command"] == command, f"probe {leg} command: {index}")
        require(row["exit_code"] == 0, f"probe {leg} gate failed: {index}")
        log = artifact(row["log"], f"probe {leg} gate {index} log")
        result = {
            "gate": index,
            "command": command,
            "status": "pass",
            "exit_code": 0,
            "log": packet_artifact(log, f"probe {leg} gate {index} log"),
        }
        if index == 2:
            test_results = parse_tests(log)
            require(len(test_results) == 1, f"probe {leg} test suite count")
            test = test_results[0]
            require(test == {"status": "ok", "passed": 36, "failed": 0, "ignored": 0, "measured": 0, "filtered": 0}, f"probe {leg} test count")
            result["tests"] = {"passed": 36, "failed": 0, "ignored": 0}
        rows.append(result)
    return {
        "schema": complete["schema"],
        "inputs": packet_artifact(inputs_path, f"probe {leg} inputs"),
        "receipts": packet_artifact(receipts_path, f"probe {leg} receipts"),
        "source": packet_artifact(source_path, f"build {leg} source"),
        "probe": actual_probe,
        "root_inputs": root_hashes,
        "gates": rows,
        "tests": {"suites": 1, "passed": 36, "failed": 0, "ignored": 0},
    }


def summarize() -> dict[str, Any]:
    root_hash = root_inputs()
    architecture = architecture_inputs()
    unrelated = unrelated_inputs()
    expected_source = candidate_source()
    production = summarize_production(root_hash, architecture, unrelated, expected_source)
    return {
        "schema": "litchi.performance.0813.quality-summary.v1",
        "production_after": production,
        "probe": {
            "before": summarize_probe("before", root_hash, architecture, unrelated),
            "after": summarize_probe("after", root_hash, architecture, unrelated),
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser()
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--write", action="store_true")
    group.add_argument("--check", action="store_true")
    args = parser.parse_args()
    result = summarize()
    output = P / "quality-summary.json"
    encoded = json.dumps(result, indent=2, sort_keys=True) + "\n"
    if args.write:
        require(not output.exists(), "quality-summary.json already exists")
        output.write_text(encoded, encoding="utf-8")
    else:
        require(output.is_file(), "quality-summary.json is missing")
        require(output.read_text(encoding="utf-8") == encoded, "quality-summary.json is stale")
    print(json.dumps({"schema": result["schema"], "status": "PASS"}, sort_keys=True))


if __name__ == "__main__":
    main()
