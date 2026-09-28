"""Fail-closed validator for the 0820 allocator-test repair record.

The original durability packet stopped at its first failed quality run.  This
validator admits only the narrowly scoped test repair and its fresh six-gate
quality evidence.  It never runs Cargo, a benchmark, a workload, or a child
process.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
TOOL = ROOT / "tools" / "perf-baseline"
BASE = "096b810f23cc66fe01ec8c96c74f36888f954eff"
PRODUCTION_FILES = 9_197
TOOL_FILES = 87
REPAIR_NAME = "tools/perf-baseline/src/bin/support/counting_allocator.rs"
REPAIR_SCOPE = "cfg(test) allocator wrapper assertions only; original quality failure retained; no durability timing or release build admitted"
ORIGIN_SCHEMA = "litchi.performance.0820.origin.v1"
REPAIR_ORIGIN_SCHEMA = "litchi.performance.0820.repair-origin.v1"
REPAIR_INPUTS_SCHEMA = "litchi.performance.0820.repair-inputs.v1"
COMMAND_SCHEMA = "litchi.performance.0820.repair-command.v1"
QUALITY_SCHEMA = "litchi.performance.0820.repair-quality.v1"
TEST_SUMMARY_SCHEMA = "litchi.performance.0820.repair-tests.v1"
HEX = frozenset("0123456789abcdef")
QUALITY_GATES = ("fmt", "check", "tests", "clippy", "rustdoc", "boundaries")
COMMAND_NAMES = ("focused",) + QUALITY_GATES
TEST_SUMMARY_RE = re.compile(
    r"^test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored; "
    r"(\d+) measured; (\d+) filtered out;", re.MULTILINE)


class ValidationError(RuntimeError):
    """Retained repair evidence is missing, stale, or contradictory."""


def fail(message: str) -> None:
    raise ValidationError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def descriptor(value: Any, label: str, expected: Path | None = None) -> Path:
    require(isinstance(value, dict) and set(value) == {"path", "bytes", "sha256"},
            f"{label} descriptor changed")
    raw_path = value.get("path")
    require(isinstance(raw_path, str) and raw_path, f"{label}.path is invalid")
    path = Path(raw_path)
    if not path.is_absolute():
        path = ROOT / path
    path = path.resolve(strict=False)
    require(path.is_relative_to(ROOT.resolve()), f"{label}.path escaped workspace")
    require(expected is None or path == expected.resolve(), f"{label}.path changed")
    require(isinstance(value.get("bytes"), int) and not isinstance(value.get("bytes"), bool)
            and value["bytes"] >= 0, f"{label}.bytes is invalid")
    require(is_sha(value.get("sha256")), f"{label}.sha256 is invalid")
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    require(path.stat().st_size == value["bytes"], f"{label}.bytes changed")
    require(sha(path) == value["sha256"], f"{label}.sha256 changed")
    return path


def identity(path: Path) -> dict[str, Any]:
    path = path.resolve(strict=False)
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def descriptor_matches(value: Any, path: Path, label: str) -> None:
    descriptor(value, label, path)


def origin() -> dict[str, Any]:
    value = read_json(PACKET / "origin.json")
    require(value.get("schema") == ORIGIN_SCHEMA, "original origin schema changed")
    require(value.get("base") == BASE and value.get("production_changed") is False
            and value.get("runtime_harness_changed") is False
            and value.get("tool_changed") is False
            and value.get("tool_allowlist") == [],
            "original source boundary changed")
    require(value.get("target") == "/home/zhuhe/code/litchi-target-0820"
            and value.get("scratch") == "/home/zhuhe/code/litchi-fs-0820",
            "original target or scratch changed")
    unrelated = value.get("unrelated")
    require(isinstance(unrelated, dict) and set(unrelated) == {
        "docs/FORMAT_IMPLEMENTATION_REVIEW.md",
        "docs/UNIFIED_OPS_API_DESIGN.md", "matrix-analysis.json",
    } and all(is_sha(item) for item in unrelated.values()),
            "original unrelated-file witness changed")
    return value


def baseline_source() -> dict[str, Any]:
    value = read_json(PACKET / "quality-0" / "source.json")
    require(set(value) == {"production", "tool"}, "original source witness changed")
    production = value.get("production")
    tool = value.get("tool")
    require(isinstance(production, dict) and isinstance(tool, dict)
            and production.get("revision") == BASE
            and isinstance(production.get("files"), dict)
            and len(production["files"]) == PRODUCTION_FILES
            and len(tool) == TOOL_FILES
            and all(is_sha(item) for item in production["files"].values())
            and all(is_sha(item) for item in tool.values()),
            "original source census changed")
    return value


def frozen_inputs() -> dict[str, Any]:
    value = read_json(PACKET / "quality-0" / "frozen-inputs.json")
    require(value.get("schema") == "litchi.performance.0820.frozen-inputs.v1",
            "original frozen-input schema changed")
    required = {"schema", "packet", "drivers", "root_inputs", "locks", "architecture",
                "corpus", "host", "unrelated"}
    require(set(value) == required, "original frozen-input keys changed")
    for key in ("packet", "drivers", "root_inputs", "architecture", "corpus", "unrelated"):
        require(isinstance(value[key], dict), f"original frozen {key} changed")
    require(is_sha(value["host"]), "original frozen host hash changed")
    require(isinstance(value["locks"], dict)
            and set(value["locks"]) == {"packet_tool", "root", "tool"},
            "original frozen lock witnesses changed")
    for item in value["locks"].values():
        require(isinstance(item, dict) and set(item) == {"bytes", "path", "sha256"}
                and isinstance(item["bytes"], int) and item["bytes"] >= 0
                and isinstance(item["path"], str) and is_sha(item["sha256"]),
                "original frozen lock descriptor changed")
    require(all(is_sha(item) for item in value["packet"].values())
            and all(is_sha(item) for item in value["drivers"].values())
            and all(is_sha(item) for item in value["root_inputs"].values())
            and all(is_sha(item) for item in value["architecture"].values())
            and all(is_sha(item["sha256"]) for item in value["corpus"].values())
            and all(is_sha(item) for item in value["unrelated"].values()),
            "original frozen hashes changed")
    return value


def check_frozen_files(frozen: dict[str, Any], base: dict[str, Any]) -> None:
    packet = frozen["packet"]
    require(len(packet) == 23 and len(frozen["drivers"]) == 5,
            "original frozen packet cardinality changed")
    for name, wanted in packet.items():
        descriptor_matches({"path": str(PACKET / name), "bytes": (PACKET / name).stat().st_size,
                            "sha256": sha(PACKET / name)}, PACKET / name,
                           f"original frozen packet {name}")
        require(sha(PACKET / name) == wanted, f"original frozen packet changed: {name}")
    for name, wanted in frozen["drivers"].items():
        path = PACKET / name
        require(path.is_file() and sha(path) == wanted, f"original driver changed: {name}")
    for name, wanted in frozen["root_inputs"].items():
        packet_copy = PACKET / "inputs" / ("root-Cargo.lock" if name == "Cargo.lock" else name)
        live = ROOT / name
        require(packet_copy.is_file() and live.is_file()
                and sha(packet_copy) == wanted and sha(live) == wanted,
                f"root input changed: {name}")
    for name, wanted in frozen["architecture"].items():
        require((ROOT / name).is_file() and sha(ROOT / name) == wanted,
                f"architecture input changed: {name}")
    for name, item in frozen["corpus"].items():
        path = ROOT / name
        require(path.is_file() and path.stat().st_size == item["bytes"]
                and sha(path) == item["sha256"], f"corpus input changed: {name}")
    for name, wanted in frozen["unrelated"].items():
        require((ROOT / name).is_file() and sha(ROOT / name) == wanted,
                f"unrelated workspace file changed: {name}")
    require(sha(PACKET / "host.json") == frozen["host"],
            "original frozen host changed")
    lock_paths = {
        "packet_tool": PACKET / "inputs" / "tool-Cargo.lock",
        "root": ROOT / "Cargo.lock",
        "tool": TOOL / "Cargo.lock",
    }
    for name, expected in lock_paths.items():
        item = frozen["locks"][name]
        descriptor_matches(item, expected, f"original frozen lock {name}")
    require(base["production"]["files"] and len(base["production"]["files"]) == PRODUCTION_FILES,
            "baseline production census missing")


def check_original_failure(base: dict[str, Any], frozen: dict[str, Any]) -> list[dict[str, Any]]:
    checks_path = PACKET / "quality-0" / "checks.json"
    failure_path = PACKET / "quality-0" / "failure.json"
    checks = read_json(checks_path)
    failure = read_json(failure_path)
    require(isinstance(checks, list) and len(checks) == 3
            and [row.get("gate") for row in checks] == [1, 2, 3]
            and [row.get("exit_code") for row in checks] == [0, 0, 101],
            "original quality gate result changed")
    require(failure.get("schema") == "litchi.performance.0820.quality-failure.v1"
            and failure.get("attempt") == 0 and failure.get("failed_gate") == 3
            and failure.get("rows") == checks
            and failure.get("packet") == frozen["packet"]
            and failure.get("drivers") == frozen["drivers"],
            "original quality failure witness changed")
    descriptor_matches(failure.get("source"), PACKET / "quality-0" / "source.json",
                       "original failure source")
    descriptor_matches(failure["row"].get("log"), PACKET / "quality-0" / "02.log",
                       "original failed gate log")
    require(failure["row"] == checks[2], "original failed gate row changed")
    for index, row in enumerate(checks):
        descriptor_matches(row.get("log"), PACKET / "quality-0" / f"{index:02}.log",
                           f"original gate {index + 1} log")
        finite(row.get("started"), f"original gate {index + 1} start")
        finite(row.get("ended"), f"original gate {index + 1} end")
        require(row["started"] <= row["ended"], f"original gate {index + 1} timestamps reversed")
    manifest = str(TOOL / "Cargo.toml")
    expected = [
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "check", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets"],
        ["cargo", "test", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--", "--test-threads=2"],
    ]
    require([row.get("command") for row in checks] == expected,
            "original quality commands changed")
    require(failure["row"].get("exit_code") == 101, "original failure exit code changed")
    return checks


def current_source(base: dict[str, Any]) -> dict[str, Any]:
    production_names = base["production"]["files"]
    tool_names = base["tool"]
    production = {name: sha(ROOT / name) for name in production_names}
    tool = {name: sha(ROOT / name) for name in tool_names}
    require(production == base["production"]["files"], "production source changed")
    changed = {name for name in tool if tool[name] != base["tool"][name]}
    require(changed == {REPAIR_NAME} and len(tool) == TOOL_FILES,
            "repair tool source boundary changed")
    return {"production": {"revision": BASE, "files": production}, "tool": tool}


def check_repair_origin(base: dict[str, Any]) -> dict[str, Any]:
    value = read_json(PACKET / "repair" / "origin.json")
    require(set(value) == {"base", "before_sha256", "production_changed",
                            "runtime_harness_changed", "runtime_prefix_sha256",
                            "schema", "scope", "target", "tool_allowlist"},
            "repair origin keys changed")
    require(value.get("schema") == REPAIR_ORIGIN_SCHEMA and value.get("base") == BASE
            and value.get("production_changed") is False
            and value.get("runtime_harness_changed") is False
            and value.get("scope") == REPAIR_SCOPE
            and value.get("target") == "/home/zhuhe/code/litchi-target-0820/quality-1"
            and value.get("tool_allowlist") == [REPAIR_NAME]
            and is_sha(value.get("before_sha256"))
            and is_sha(value.get("runtime_prefix_sha256")),
            "repair origin changed")
    before_path = PACKET / "repair" / "before-counting_allocator.rs"
    after_path = ROOT / REPAIR_NAME
    require(sha(before_path) == value["before_sha256"] == base["tool"][REPAIR_NAME],
            "repair before-file identity changed")
    before = before_path.read_bytes()
    after = after_path.read_bytes()
    marker = b"#[cfg(test)]"
    require(before.count(marker) == 1 and after.count(marker) == 1,
            "repair test boundary changed")
    before_prefix, after_prefix = before.split(marker, 1)[0], after.split(marker, 1)[0]
    require(before_prefix == after_prefix
            and hashlib.sha256(after_prefix).hexdigest() == value["runtime_prefix_sha256"],
            "repair runtime prefix changed")
    require(b"during.live_bytes >= before.live_bytes + 256" not in after
            and b"assert_live_bytes_conservation" in after,
            "repair live-byte assertion boundary changed")
    return value


def check_repair_inputs(origin_value: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "repair" / "inputs.json"
    value = read_json(path)
    require(value.get("schema") == REPAIR_INPUTS_SCHEMA
            and set(value) == {"before", "origin", "original_frozen_inputs", "runner",
                               "schema", "source"},
            "repair inputs changed")
    descriptor_matches(value["runner"], PACKET / "repair" / "run.py", "repair runner")
    descriptor_matches(value["origin"], PACKET / "repair" / "origin.json", "repair origin")
    descriptor_matches(value["before"], PACKET / "repair" / "before-counting_allocator.rs",
                       "repair before file")
    descriptor_matches(value["source"], PACKET / "repair" / "source.json", "repair source")
    descriptor_matches(value["original_frozen_inputs"],
                       PACKET / "quality-0" / "frozen-inputs.json",
                       "original frozen inputs")
    require(read_json(PACKET / "repair" / "origin.json") == origin_value,
            "repair origin descriptor changed")
    return value


def check_repair_source(base: dict[str, Any], origin_value: dict[str, Any]) -> dict[str, Any]:
    value = read_json(PACKET / "repair" / "source.json")
    current = current_source(base)
    require(value == current, "repair source witness changed")
    require(value["production"]["revision"] == origin_value["base"]
            and len(value["production"]["files"]) == PRODUCTION_FILES
            and len(value["tool"]) == TOOL_FILES,
            "repair source census changed")
    return value


def expected_commands() -> dict[str, list[str]]:
    manifest = str(TOOL / "Cargo.toml")
    prefix = ["--offline", "--locked", "--manifest-path", manifest]
    return {
        "focused": ["cargo", "test", *prefix, "--all-features", "--bin",
                    "docx_replayable_tail_append", "--", "--test-threads=2"],
        "fmt": ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        "check": ["cargo", "check", *prefix, "--all-features", "--all-targets"],
        "tests": ["cargo", "test", *prefix, "--all-features", "--", "--test-threads=2"],
        "clippy": ["cargo", "clippy", *prefix, "--all-features", "--all-targets",
                   "--", "-D", "warnings"],
        "rustdoc": ["cargo", "doc", *prefix, "--all-features", "--no-deps"],
        "boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
    }


def check_command(name: str, origin_value: dict[str, Any], inputs: dict[str, Any],
                  source: dict[str, Any]) -> dict[str, Any]:
    directory = PACKET / "repair" / "commands" / name
    started_path, result_path, log_path = directory / "started.json", directory / "result.json", directory / "output.log"
    started = read_json(started_path)
    result = read_json(result_path)
    require(set(started) == {"schema", "name", "command", "started", "environment",
                             "inputs", "source"}, f"{name} start receipt keys changed")
    require(set(result) == {"schema", "name", "command", "started", "ended", "environment",
                            "inputs", "source", "exit_code", "log"},
            f"{name} result receipt keys changed")
    require(started.get("schema") == COMMAND_SCHEMA and result.get("schema") == COMMAND_SCHEMA
            and started.get("name") == name and result.get("name") == name
            and started.get("command") == expected_commands()[name]
            and result.get("command") == started.get("command")
            and result.get("started") == started.get("started")
            and result.get("environment") == started.get("environment")
            and result.get("inputs") == started.get("inputs")
            and result.get("source") == started.get("source")
            and result.get("exit_code") == 0,
            f"{name} command receipt changed")
    expected_environment = {
        "CARGO_TARGET_DIR": origin_value["target"], "CARGO_BUILD_JOBS": "2",
        "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_DEV_DEBUG": "0",
        "RUSTDOCFLAGS": "-D warnings", "PYTHONDONTWRITEBYTECODE": "1",
    }
    require(started.get("environment") == expected_environment,
            f"{name} environment changed")
    descriptor_matches(started["inputs"], PACKET / "repair" / "inputs.json",
                       f"{name} inputs")
    descriptor_matches(started["source"], PACKET / "repair" / "source.json",
                       f"{name} source")
    descriptor_matches(result["log"], log_path, f"{name} log")
    finite(started.get("started"), f"{name} start")
    finite(result.get("ended"), f"{name} end")
    require(started["started"] <= result["ended"], f"{name} timestamps reversed")
    return result


def check_test_summary(source: dict[str, Any]) -> dict[str, Any]:
    value = read_json(PACKET / "repair" / "test-summary.json")
    require(set(value) == {"schema", "suites", "passed", "failed", "ignored", "log", "scope"}
            and value.get("schema") == TEST_SUMMARY_SCHEMA
            and isinstance(value.get("suites"), int) and value["suites"] > 0
            and isinstance(value.get("passed"), int) and value["passed"] > 0
            and value.get("failed") == 0
            and isinstance(value.get("ignored"), int) and value["ignored"] >= 0
            and value.get("scope") ==
            "full perf-baseline all-features tests including doctest invocation",
            "repair test summary changed")
    descriptor_matches(value["log"], PACKET / "repair" / "commands" / "tests" / "output.log",
                       "repair test summary log")
    full_log = (PACKET / "repair" / "commands" / "tests" / "output.log").read_text(
        encoding="utf-8")
    require("test result: FAILED." not in full_log and "test result: ok." in full_log,
            "repair full-test log changed")
    summaries = TEST_SUMMARY_RE.findall(full_log)
    require(len(summaries) == 28
            and sum(int(row[0]) for row in summaries) == 641
            and sum(int(row[1]) for row in summaries) == 0
            and sum(int(row[2]) for row in summaries) == 1,
            "repair full-test aggregate changed")
    require(value["suites"] == len(summaries)
            and value["passed"] == sum(int(row[0]) for row in summaries)
            and value["failed"] == sum(int(row[1]) for row in summaries)
            and value["ignored"] == sum(int(row[2]) for row in summaries),
            "repair test summary does not match full-test log")
    focused_log = (PACKET / "repair" / "commands" / "focused" / "output.log").read_text(
        encoding="utf-8")
    focused_tests = (
        "global_allocator_records_successful_alloc_and_dealloc_with_process_live_accounting",
        "global_allocator_records_zeroed_alloc_and_zeroes_memory",
        "global_allocator_records_realloc_grow_and_shrink",
        "global_allocator_records_a_failed_allocation",
        "global_allocator_records_a_failed_realloc_without_losing_old_allocation",
    )
    require("test result: FAILED." not in focused_log
            and "PoisonError" not in focused_log
            and "test result: ok." in focused_log
            and all(name in focused_log for name in focused_tests),
            "focused allocator test log changed")
    focused_summaries = TEST_SUMMARY_RE.findall(focused_log)
    require(len(focused_summaries) == 1
            and tuple(int(item) for item in focused_summaries[0][:3]) == (5, 0, 0),
            "focused allocator test aggregate changed")
    return value


def check_chronology(original_rows: list[dict[str, Any]],
                     results: dict[str, dict[str, Any]],
                     cleanup: dict[str, Any] | None) -> None:
    timeline = [(f"original gate {index + 1}", row["started"], row["ended"])
                for index, row in enumerate(original_rows)]
    timeline.extend((name, results[name]["started"], results[name]["ended"])
                    for name in COMMAND_NAMES)
    for (previous_name, _, previous_end), (current_name, current_start, _) in zip(
            timeline, timeline[1:]):
        require(previous_end <= current_start,
                f"repair chronology overlaps {previous_name} and {current_name}")
    if cleanup is not None:
        require(results[QUALITY_GATES[-1]]["ended"] <= cleanup["started"],
                "repair cleanup started before the last quality gate ended")


def check_quality_summary(results: dict[str, dict[str, Any]], inputs: dict[str, Any],
                          source: dict[str, Any]) -> dict[str, Any]:
    value = read_json(PACKET / "repair" / "quality.json")
    require(value.get("schema") == QUALITY_SCHEMA and value.get("status") == "pass"
            and value.get("gate_count") == 6
            and set(value) == {"schema", "status", "gate_count", "gates", "focused",
                               "source", "inputs"},
            "repair quality summary changed")
    require(value.get("gates") == [results[name]["_identity"] for name in QUALITY_GATES]
            and value.get("focused") == results["focused"]["_identity"],
            "repair quality receipt set changed")
    descriptor_matches(value["source"], PACKET / "repair" / "source.json",
                       "repair quality source")
    descriptor_matches(value["inputs"], PACKET / "repair" / "inputs.json",
                       "repair quality inputs")
    return value


def check_no_downstream(final: bool) -> None:
    forbidden = (
        "quality.json", "build.json", "artifacts.complete.json", "artifacts-receipt.json",
        "artifact-admission.json", "artifact-audit.json", "zip-preservation.json",
        "qualification-admission.json", "artifacts", "qualification",
        "native", "observer",
    )
    if not final:
        forbidden += ("cleanup.json",)
    for name in forbidden:
        require(not (PACKET / name).exists(), f"downstream evidence exists: {name}")


def check_cleanup(final: bool) -> dict[str, Any] | None:
    path = PACKET / "cleanup.json"
    if not final:
        return None
    value = read_json(path)
    require(set(value) == {"schema", "target_removed", "scratch_removed", "release_binaries_built",
                           "source", "removed", "started", "ended"}
            and value.get("schema") == "litchi.performance.0820.cleanup.v1"
            and value.get("target_removed") is True
            and value.get("scratch_removed") is True
            and value.get("release_binaries_built") == 0,
            "repair cleanup witness changed")
    descriptor_matches(value["source"], PACKET / "repair" / "source.json",
                       "repair cleanup source")
    removed = value.get("removed")
    require(isinstance(removed, list) and len(removed) == 1
            and all(isinstance(row, dict) and set(row) == {"path", "files", "logical_bytes"}
                    and isinstance(row["path"], str)
                    and isinstance(row["files"], int) and row["files"] >= 0
                    and isinstance(row["logical_bytes"], int) and row["logical_bytes"] >= 0
                    for row in removed), "repair cleanup rows changed")
    paths = {row["path"] for row in removed}
    target = ROOT.parent / "litchi-target-0820"
    scratch = ROOT.parent / "litchi-fs-0820"
    require(paths == {str(target)} and not target.exists() and not scratch.exists(),
            "repair cleanup target witness changed")
    finite(value.get("started"), "repair cleanup start")
    finite(value.get("ended"), "repair cleanup end")
    require(value["started"] <= value["ended"], "repair cleanup timestamps reversed")
    return value


def check_seal() -> None:
    path = PACKET / "seal.json"
    if not path.is_file():
        return
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0820.seal.v1"
            and isinstance(value.get("files"), dict) and value["files"],
            "repair seal witness changed")
    required = {
        REPAIR_NAME,
        "docs/performance/0820-allocator-test-quality-repair.md",
        "docs/performance/BASELINE.md",
        "docs/performance/CRUD_COVERAGE.md",
        "docs/performance/GOAL_AUDIT.md",
        "docs/performance/HOTSPOTS.md",
        "docs/performance/REPORT.md",
    }
    require(required <= set(value["files"]), "repair seal additions changed")
    for name, wanted in value["files"].items():
        relative = Path(name)
        require(not relative.is_absolute() and ".." not in relative.parts and is_sha(wanted),
                f"malformed repair seal entry: {name}")
        target = (ROOT / relative).resolve(strict=False)
        require(target.is_relative_to(ROOT.resolve()) and target.is_file()
                and not target.is_symlink() and sha(target) == wanted,
                f"sealed repair payload changed: {name}")


def validate(*, final: bool = False) -> dict[str, Any]:
    check_no_downstream(final)
    original = origin()
    base = baseline_source()
    frozen = frozen_inputs()
    check_frozen_files(frozen, base)
    original_rows = check_original_failure(base, frozen)
    repair_origin = check_repair_origin(base)
    inputs = check_repair_inputs(repair_origin)
    source = check_repair_source(base, repair_origin)
    results: dict[str, dict[str, Any]] = {}
    for name in COMMAND_NAMES:
        result = check_command(name, repair_origin, inputs, source)
        result["_identity"] = identity(PACKET / "repair" / "commands" / name / "result.json")
        results[name] = result
    test_summary = check_test_summary(source)
    quality = check_quality_summary(results, inputs, source)
    cleanup = check_cleanup(final)
    check_chronology(original_rows, results, cleanup)
    check_seal()
    return {"status": "accepted", "original_failed_gate": 3,
            "focused": True, "quality_gates": 6, "cleanup_checked": cleanup is not None,
            "seal_checked": (PACKET / "seal.json").is_file(),
            "source": identity(PACKET / "repair" / "source.json"),
            "inputs": identity(PACKET / "repair" / "inputs.json"),
            "test_summary": identity(PACKET / "repair" / "test-summary.json"),
            "quality": identity(PACKET / "repair" / "quality.json"),
            "original": identity(PACKET / "quality-0" / "failure.json"),
            "repair_source": source, "repair_origin": repair_origin,
            "test_summary_value": test_summary, "quality_value": quality,
            "original_origin": original}


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--final", action="store_true",
                        help="require repair cleanup; replay seal.json when present")
    args = parser.parse_args(argv)
    try:
        value = validate(final=args.final)
        print(json.dumps({"status": value["status"], "quality_gates": value["quality_gates"],
                          "focused": value["focused"], "cleanup_checked": value["cleanup_checked"],
                          "seal_checked": value["seal_checked"]}, sort_keys=True))
    except (ValidationError, AssertionError, OSError, ValueError, KeyError, TypeError,
            IndexError) as error:
        print(f"0820 repair validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
