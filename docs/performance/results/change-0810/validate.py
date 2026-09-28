"""Fail-closed aggregate validator for the 0810 workflow packet.

This reader is intentionally independent of the numerical and profile
readers.  It verifies their retained schemas and ``--check`` paths, joins
their custody and policy results, checks stage chronology, and optionally
requires the final six-binary cleanup witness.  It never invokes Cargo,
rustfmt, a probe, a workload, Valgrind, or a profiler, and it writes no
artifact of its own.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / "litchi-target-0810"
BASE = "3677e31be5c9d5582a1f6d531ebb4d54db5a0acc"
SOURCE_ALLOWLIST = {"crates/litchi-pptx/src/notes/codec.rs"}
SOURCE_COUNT = 9196
CASES = tuple(
    (shape, mode)
    for shape in ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
    for mode in ("capture", "commit", "lifecycle")
)
ORDERS = {
    "qualification": (("before",),),
    "native": (
        ("before", "after"),
        ("after", "before"),
        ("before", "after"),
        ("after", "before"),
        ("after", "before"),
        ("before", "after"),
    ),
    "allocation": (("before", "after"), ("after", "before")),
}
LANE_SAMPLES = {"qualification": 1, "native": 30, "allocation": 3}
LANE_REPORTS = {"qualification": 18, "native": 216, "allocation": 72}
LANE_TOTALS = {"qualification": 18, "native": 6480, "allocation": 216}
PROFILE_JOBS = ((0, "before"), (0, "after"), (1, "after"), (1, "before"))
READER_CHECKS = (
    ("analysis.py", "analysis.json"),
    ("root_audit.py", "root-audit.json"),
    ("profile_analysis.py", "profile-analysis.json"),
    ("quality_summary.py", "quality-summary.json"),
)


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(
        character in "0123456789abcdef" for character in value
    )


def packet_path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path is missing")
    path = Path(value)
    if not path.is_absolute():
        path = P / path
    path = path.resolve()
    require(path.is_relative_to(P), f"{label}: path escapes packet")
    return path


def packet_artifact(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: artifact is malformed")
    path = packet_path(value.get("path"), label)
    require(type(value.get("bytes")) is int and value["bytes"] >= 0,
            f"{label}: byte count is invalid")
    require(is_sha(value.get("sha256")), f"{label}: SHA-256 is invalid")
    require(path.is_file() and not path.is_symlink(), f"{label}: file is missing")
    require(path.stat().st_size == value["bytes"], f"{label}: byte count changed")
    require(sha(path) == value["sha256"], f"{label}: SHA-256 changed")
    return path


def external_artifact(value: Any, label: str, cleanup: dict[str, Any] | None) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: binary is malformed")
    raw = value.get("path")
    require(isinstance(raw, str) and raw, f"{label}: binary path is missing")
    path = Path(raw).resolve()
    require(path.is_absolute() and path.parent == TARGET,
            f"{label}: binary path is outside owned target")
    size, digest = value.get("bytes"), value.get("sha256")
    require(type(size) is int and size > 0, f"{label}: binary byte count is invalid")
    require(is_sha(digest), f"{label}: binary SHA-256 is invalid")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == size and sha(path) == digest,
                f"{label}: live binary identity changed")
        return {"path": str(path), "bytes": size, "sha256": digest}
    require(cleanup is not None and cleanup.get("target_removed") is True,
            f"{label}: missing binary has no cleanup witness")
    require(not TARGET.exists(), f"{label}: target remains after cleanup")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list), f"{label}: cleanup binary list is malformed")
    matches = [item for item in removed if isinstance(item, dict)
               and item.get("path") == str(path)]
    require(len(matches) == 1 and matches[0].get("bytes") == size
            and matches[0].get("sha256") == digest,
            f"{label}: exact removed binary identity is missing")
    return {"path": str(path), "bytes": size, "sha256": digest}


def source_manifest(path: Path, label: str) -> dict[str, Any]:
    value = read(path)
    require(isinstance(value, dict), f"{label}: source manifest is malformed")
    revision, files = value.get("revision"), value.get("files")
    require(isinstance(revision, str) and len(revision) == 40,
            f"{label}: revision is malformed")
    require(isinstance(files, dict) and len(files) == SOURCE_COUNT,
            f"{label}: source census count changed")
    require(all(isinstance(name, str) and is_sha(digest)
                for name, digest in files.items()),
            f"{label}: source digest map is malformed")
    return {"revision": revision, "files": dict(files)}


def current_source() -> dict[str, str]:
    raw = subprocess.check_output(
        [
            "git",
            "ls-files",
            "-z",
            "--",
            "crates",
            "Cargo.toml",
            "clippy.toml",
            ".cargo/config.toml",
            "rust-toolchain.toml",
        ],
        cwd=ROOT,
    )
    names = [name for name in raw.decode().split("\0") if name]
    return {name: sha(ROOT / name) for name in names}


def verify_origin() -> dict[str, Any]:
    origin = read(P / "origin.json")
    require(origin.get("schema") == "litchi.performance.0810.origin.v1",
            "origin schema changed")
    require(origin.get("base") == BASE and origin.get("worktree") == str(ROOT),
            "origin base or worktree changed")
    production = origin.get("production_source")
    require(isinstance(production, dict)
            and production.get("revision") == BASE
            and production.get("tracked_file_count") == SOURCE_COUNT
            and production.get("candidate_allowlist") == sorted(SOURCE_ALLOWLIST),
            "origin production custody changed")
    unrelated = origin.get("unrelated")
    require(isinstance(unrelated, dict), "origin unrelated identities are missing")
    for name, digest in unrelated.items():
        path = ROOT / name
        require(path.is_file() and sha(path) == digest, f"unrelated file changed: {name}")
    architecture = read(P / "architecture-inputs.json")
    require(isinstance(architecture, dict), "architecture input map is malformed")
    for name, digest in architecture.items():
        path = ROOT / name
        require(path.is_file() and sha(path) == digest, f"architecture input changed: {name}")
    root_inputs = origin.get("root_inputs")
    require(isinstance(root_inputs, dict)
            and root_inputs.get("copies_captured_before_first_build") is True
            and is_sha(root_inputs.get("root-Cargo.lock"))
            and is_sha(root_inputs.get("rustfmt.toml")),
            "root input custody changed")
    for name, digest in {
        "Cargo.lock": root_inputs["root-Cargo.lock"],
        "rustfmt.toml": root_inputs["rustfmt.toml"],
    }.items():
        copy = P / "inputs" / ("root-Cargo.lock" if name == "Cargo.lock" else "rustfmt.toml")
        require(copy.is_file() and sha(copy) == digest, f"frozen {name} copy changed")
        require((ROOT / name).is_file() and sha(ROOT / name) == digest,
                f"live {name} changed")
    return {"origin": origin, "root_inputs": {
        "Cargo.lock": root_inputs["root-Cargo.lock"],
        "rustfmt.toml": root_inputs["rustfmt.toml"],
    }}


def verify_frozen_build_inputs() -> None:
    expected = {
        "plan.json", "adoption-policy.json", "analysis-plan.json", "custody.py",
        "build.py", "capture.py", "profile.py", "quality.py", "probe_quality.py",
        "apply_candidate.py", "restore_candidate.py", "origin.json", "host.json",
        "inheritance.json", "architecture-inputs.json", "inputs/root-Cargo.lock",
        "inputs/rustfmt.toml",
    }
    for leg in ("before", "after"):
        value = read(P / f"build-{leg}" / "frozen-inputs.json")
        require(set(value) == {"packet", "root_inputs"},
                f"{leg} frozen input envelope changed")
        require(set(value["packet"]) == expected, f"{leg} frozen input set changed")
        for name, digest in value["packet"].items():
            path = P / name
            require(is_sha(digest) and path.is_file() and sha(path) == digest,
                    f"{leg} frozen packet input changed: {name}")
        require(value["root_inputs"] == verify_origin()["root_inputs"],
                f"{leg} frozen root input receipt changed")


def verify_plan() -> dict[str, Any]:
    plan = read(P / "plan.json")
    require(plan.get("schema") == "litchi.performance.0810.v1", "plan schema changed")
    require(plan.get("source_allowlist") == sorted(SOURCE_ALLOWLIST), "plan scope changed")
    require(plan.get("cases") == [{"mode": mode, "shape": shape} for shape, mode in CASES],
            "plan case order changed")
    for lane, orders in ORDERS.items():
        row = plan.get(lane)
        require(isinstance(row, dict) and row.get("orders") == [list(order) for order in orders],
                f"{lane} order changed")
        require(row.get("samples") == LANE_SAMPLES[lane]
                and row.get("reports") == LANE_REPORTS[lane]
                and row.get("samples_total") == LANE_TOTALS[lane],
                f"{lane} cardinality changed")
    require(plan["qualification"].get("binary") == "allocation",
            "qualification binary changed")
    require(plan.get("totals") == {"reports": 310, "samples": 6718},
            "plan totals changed")
    boot = plan.get("bootstrap")
    require(boot == {
        "resamples": 10000,
        "seed": 810810,
        "statistic": "median",
        "sorted_zero_based_endpoints": [250, 9749],
    }, "bootstrap plan changed")
    policy = read(P / "adoption-policy.json")
    require(policy.get("schema") == "litchi.performance.0810.adoption-policy.v1",
            "adoption policy schema changed")
    require(policy.get("frozen_before_build") is True
            and policy.get("useful_public_workflow_benefit_required") is True
            and policy.get("allocation_count_alone_sufficient") is False
            and policy.get("latency", {}).get("seed") == 810810
            and policy.get("latency", {}).get("resamples") == 10000
            and policy.get("benefit", {}).get("eligible_modes") == ["capture", "lifecycle"]
            and policy.get("benefit", {}).get("minimum_improvement_percent") == 3.0,
            "adoption policy changed")
    return plan


def candidate_sources(before: dict[str, Any]) -> dict[str, Any]:
    manifest = read(P / "candidate/manifest.json")
    require(manifest.get("schema") == "litchi.performance.0810.candidate-manifest.v1",
            "candidate manifest schema changed")
    require(manifest.get("base_commit") == BASE
            and manifest.get("production_path") in SOURCE_ALLOWLIST
            and manifest.get("production_source_changed") is False,
            "candidate manifest scope changed")
    files = manifest.get("files")
    rows = list(files.values()) if isinstance(files, dict) else [manifest]
    require(len(rows) == 1 and {row.get("production_path") for row in rows} == SOURCE_ALLOWLIST,
            "candidate manifest file set changed")
    row = rows[0]
    before_path = packet_artifact(row["before"], "candidate before")
    after_path = packet_artifact(row["after"], "candidate after")
    require(before["files"][next(iter(SOURCE_ALLOWLIST))] == sha(before_path),
            "candidate before differs from build-before source")
    expected = dict(before["files"])
    expected[next(iter(SOURCE_ALLOWLIST))] = sha(after_path)
    packet_artifact(manifest["patch"], "candidate patch")
    return {"revision": before["revision"], "files": expected}


def verify_builds(cleanup: dict[str, Any] | None) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    builds: dict[str, Any] = {}
    for leg in ("before", "after"):
        directory = P / f"build-{leg}"
        manifest = read(directory / "build.json")
        require(manifest.get("schema") == f"litchi.performance.0810.build-{leg}.v1",
                f"{leg} build schema changed")
        source_path = packet_artifact(manifest["source"], f"{leg} source")
        require(source_path == (directory / "source.json").resolve(),
                f"{leg} source path changed")
        source = source_manifest(source_path, f"{leg} source")
        require(manifest.get("root_inputs") == verify_origin()["root_inputs"],
                f"{leg} root input receipt changed")
        rows = manifest.get("rows")
        require(isinstance(rows, list) and len(rows) == 3, f"{leg} build row count changed")
        last = -math.inf
        for row in rows:
            require(row.get("name") in {"native", "allocation", "profile"}
                    and row.get("exit_code") == 0, f"{leg} build command failed")
            require(isinstance(row.get("started"), (int, float))
                    and isinstance(row.get("ended"), (int, float))
                    and last <= row["started"] <= row["ended"],
                    f"{leg} build chronology changed")
            last = row["ended"]
            packet_artifact(row["log"], f"{leg} build log")
        binaries = manifest.get("binaries")
        require(isinstance(binaries, dict) and set(binaries) == {"native", "allocation", "profile"},
                f"{leg} binary matrix changed")
        checked = {name: external_artifact(value, f"{leg} {name}", cleanup)
                   for name, value in binaries.items()}
        builds[leg] = {"manifest": manifest, "source": source, "binaries": checked}
    before, after = builds["before"]["source"], builds["after"]["source"]
    require(before["revision"] == BASE, "before source revision changed")
    changed = {name for name in before["files"] | after["files"]
               if before["files"].get(name) != after["files"].get(name)}
    require(changed == SOURCE_ALLOWLIST, f"build source scope changed: {changed}")
    return builds, before, after


def verify_application(before: dict[str, Any], after: dict[str, Any]) -> None:
    application = read(P / "application.json")
    require(application.get("schema") == "litchi.performance.0810.application.v1",
            "application schema changed")
    require(application.get("source_before", {}).get("files") == before["files"]
            and application.get("source", {}).get("files") == after["files"]
            and application.get("allowlist") == sorted(SOURCE_ALLOWLIST),
            "application source custody changed")
    packet_artifact(application["manifest"], "application manifest")
    packet_artifact(application["patch"], "application patch")


def verify_lane(lane: str, plan: dict[str, Any], builds: dict[str, Any],
                expected_source: dict[str, Any], cleanup: dict[str, Any] | None) -> tuple[float, float]:
    directory = P / lane
    complete = read(directory / "complete.json")
    require(complete.get("schema") == f"litchi.performance.0810.{lane}.complete.v1",
            f"{lane} completion schema changed")
    require(complete.get("children") == LANE_REPORTS[lane]
            and complete.get("reports") == LANE_REPORTS[lane]
            and complete.get("samples") == LANE_TOTALS[lane]
            and complete.get("plan_sha256") == sha(P / "plan.json"),
            f"{lane} completion cardinality changed")
    source_path = packet_artifact(complete["source"], f"{lane} source")
    require(source_manifest(source_path, f"{lane} source") == expected_source,
            f"{lane} source differs from expected build")
    receipts = read(packet_artifact(complete["receipts"], f"{lane} receipts"))
    expected_jobs = [
        (block, shape, mode, leg)
        for block, order in enumerate(ORDERS[lane])
        for shape, mode in CASES
        for leg in order
    ]
    require(len(receipts) == len(expected_jobs), f"{lane} receipt count changed")
    first, last = math.inf, -math.inf
    previous = -math.inf
    for index, (row, identity) in enumerate(zip(receipts, expected_jobs)):
        block, shape, mode, leg = identity
        require(row.get("schema") == "litchi.performance.0810.capture-receipt.v1"
                and (row.get("lane"), row.get("block"), row.get("shape"),
                     row.get("mode"), row.get("leg"))
                == (lane, block, shape, mode, leg)
                and row.get("exit_code") == 0,
                f"{lane} receipt identity changed: {index}")
        require(row.get("root_inputs") == verify_origin()["root_inputs"],
                f"{lane} root input receipt changed: {index}")
        started, ended = row.get("started"), row.get("ended")
        require(isinstance(started, (int, float)) and isinstance(ended, (int, float))
                and previous <= started <= ended,
                f"{lane} receipt chronology changed: {index}")
        first, last = min(first, started), max(last, ended)
        previous = ended
        kind = "allocation" if lane == "qualification" else lane
        expected_binary = builds[leg]["binaries"][kind]
        require(row.get("binary") == expected_binary, f"{lane} binary changed: {index}")
        packet_artifact(row["report"], f"{lane} report {index}")
        packet_artifact(row["log"], f"{lane} log {index}")
        packet_artifact(row["rss"], f"{lane} RSS {index}")
    return first, last


def verify_quality_summary(expected_after: dict[str, Any], root_inputs: dict[str, str]) -> None:
    value = read(P / "quality-summary.json")
    require(value.get("schema") == "litchi.performance.0810.quality-summary.v1",
            "quality summary schema changed")
    production = value.get("production_after")
    require(isinstance(production, dict) and production.get("root_inputs") == root_inputs,
            "quality summary root-input custody changed")
    source_path = packet_artifact(production["source"], "quality summary source")
    require(source_manifest(source_path, "quality summary source") == expected_after,
            "quality summary source differs from candidate")
    gates = production.get("gates")
    require(isinstance(gates, list) and len(gates) == 6
            and all(row.get("status") == "pass" and row.get("exit_code") == 0 for row in gates),
            "quality summary production gates changed")
    tests = production.get("tests")
    require(tests.get("passed") == 1241 and tests.get("failed") == 0
            and tests.get("ignored") == 3, "production test summary changed")
    for leg in ("before", "after"):
        probe = value.get("probe", {}).get(leg)
        require(isinstance(probe, dict) and probe.get("root_inputs") == root_inputs
                and len(probe.get("gates", [])) == 3
                and all(row.get("status") == "pass" and row.get("exit_code") == 0
                        for row in probe["gates"])
                and probe.get("tests", {}).get("passed") == 36
                and probe.get("tests", {}).get("failed") == 0,
                f"probe {leg} quality summary changed")


def case_keys(rows: Iterable[dict[str, Any]], *, field: str = "case") -> set[str]:
    result = set()
    for row in rows:
        if field in row:
            result.add(row[field])
        else:
            result.add(f"{row.get('shape')}/{row.get('mode')}")
    return result


def median(values: list[int | float]) -> int | float:
    """Return the even-count median used by the independent raw audit."""
    require(values, "cannot take the median of an empty list")
    ordered = sorted(values)
    middle = len(ordered) // 2
    if len(ordered) % 2:
        return ordered[middle]
    return (ordered[middle - 1] + ordered[middle]) / 2


def verify_native_numerical_agreement(
    analysis: dict[str, Any], audit: dict[str, Any]
) -> None:
    """Require every retained native p50 pair to match the independent audit."""
    main_rows = analysis.get("native", {}).get("analysis", {}).get(
        "paired_by_block_before_after"
    )
    audit_rows = audit.get("native")
    require(isinstance(main_rows, dict), "workflow native paired rows are missing")
    require(isinstance(audit_rows, list), "root native paired rows are missing")
    main_cases = set(main_rows)
    audit_cases = {f"{row.get('shape')}/{row.get('mode')}" for row in audit_rows}
    require(main_cases == audit_cases,
            "main/root native numerical case set disagreement")
    require(len(main_cases) == len(CASES),
            "native numerical case count changed")
    audit_by_case = {
        f"{row.get('shape')}/{row.get('mode')}": row for row in audit_rows
    }
    require(len(audit_by_case) == len(audit_rows),
            "root native numerical cases are duplicated")
    for case in sorted(main_cases):
        main_case = main_rows[case]
        root_case = audit_by_case[case]
        metrics = main_case.get("metrics", {}).get("p50")
        require(isinstance(metrics, dict), f"{case}: native p50 metrics are missing")
        blocks = metrics.get("by_block")
        require(isinstance(blocks, list) and len(blocks) == 6,
                f"{case}: native block rows changed")
        root_blocks = root_case.get("blocks")
        require(isinstance(root_blocks, list) and len(root_blocks) == 6,
                f"{case}: root native block rows changed")
        root_by_block = {row.get("block"): row for row in root_blocks}
        require(set(root_by_block) == {0, 1, 2, 3, 4, 5},
                f"{case}: root native block identities changed")
        main_by_block = {row.get("block"): row for row in blocks}
        require(set(main_by_block) == set(root_by_block),
                f"{case}: main/root native block identities disagree")
        main_ratios = []
        for block in range(6):
            left, right = main_by_block[block], root_by_block[block]
            for field in ("before", "after", "ratio"):
                require(left.get(field) == right.get(field),
                        f"{case} block {block}: native {field} disagreement")
            main_ratios.append(left["ratio"])
        require(root_case.get("paired_ratios") == main_ratios,
                f"{case}: native paired ratio list disagreement")
        require(metrics.get("ratio_median") == root_case.get("ratio"),
                f"{case}: native ratio median disagreement")
        bootstrap = metrics.get("bootstrap")
        ci = root_case.get("ci95")
        require(isinstance(bootstrap, dict) and isinstance(ci, dict),
                f"{case}: native bootstrap CI is missing")
        require(bootstrap.get("ci_low") == ci.get("low")
                and bootstrap.get("ci_high") == ci.get("high")
                and bootstrap.get("resamples") == ci.get("resamples")
                and bootstrap.get("seed") == ci.get("seed"),
                f"{case}: native bootstrap CI disagreement")
        require(root_case.get("before_p50_ns") == median(
            [row["before"] for row in blocks]
        ) and root_case.get("after_p50_ns") == median(
            [row["after"] for row in blocks]
        ), f"{case}: native aggregate p50 disagreement")


ALLOCATION_METRICS = (
    "allocation_calls",
    "allocated_bytes",
    "net_live",
    "peak_above_entry",
)


def verify_allocation_numerical_agreement(
    analysis: dict[str, Any], audit: dict[str, Any]
) -> None:
    """Require all four per-block allocation guard metrics to match exactly."""
    main_rows = analysis.get("allocation", {}).get("analysis", {}).get(
        "paired_by_block_before_after"
    )
    audit_rows = audit.get("allocation")
    require(isinstance(main_rows, dict), "workflow allocation paired rows are missing")
    require(isinstance(audit_rows, list), "root allocation paired rows are missing")
    require(set(main_rows) == {f"{shape}/{mode}" for shape, mode in CASES},
            "workflow allocation numerical case set changed")
    expected = {
        (case, block, metric)
        for case in main_rows
        for block in (0, 1)
        for metric in ALLOCATION_METRICS
    }
    root_by_key = {
        (f"{row.get('shape')}/{row.get('mode')}", row.get("block"), row.get("metric")): row
        for row in audit_rows
    }
    require(len(root_by_key) == len(audit_rows),
            "root allocation numerical rows are duplicated")
    require(set(root_by_key) == expected,
            "main/root allocation numerical key set disagreement")
    for case, block, metric in sorted(expected):
        main_metric = main_rows[case].get("metrics", {}).get(metric)
        require(isinstance(main_metric, dict),
                f"{case} block {block}: allocation {metric} is missing")
        raw_blocks = main_metric.get("by_block", [])
        main_blocks = {row.get("block"): row for row in raw_blocks}
        require(len(main_blocks) == len(raw_blocks),
                f"{case}: allocation {metric} block rows are duplicated")
        require(set(main_blocks) == {0, 1},
                f"{case}: allocation {metric} block rows changed")
        left = main_blocks[block]
        right = root_by_key[(case, block, metric)]
        for field in ("before", "after"):
            require(left.get(field) == right.get(field),
                    f"{case} block {block} {metric}: {field} disagreement")
        require(right.get("increase") is (right.get("after") > right.get("before")),
                f"{case} block {block} {metric}: increase witness changed")


def verify_reader_agreement() -> dict[str, Any]:
    analysis = read(P / "analysis.json")
    audit = read(P / "root-audit.json")
    profile = read(P / "profile-analysis.json")
    require(analysis.get("schema") == "litchi.performance.0810.workflow-analysis.v1",
            "workflow analysis schema changed")
    require(audit.get("schema") == "litchi.performance.0810.root-audit.v1",
            "root audit schema changed")
    require(profile.get("schema") == "litchi-0810-callgrind-profile-analysis-v1",
            "profile analysis schema changed")
    require(analysis.get("counts") == {
        "reports": 306,
        "samples": 6714,
        "native_reports": 216,
        "allocation_reports": 72,
        "qualification_reports": 18,
    }, "workflow analysis counts changed")
    require(audit.get("reports") == 306 and audit.get("samples") == 6714,
            "root audit counts changed")
    require(profile.get("summary", {}).get("profile_count") == 4
            and profile.get("summary", {}).get("all_four_profiles_complete") is True,
            "profile analysis count changed")
    decision = analysis.get("decision_guards")
    require(isinstance(decision, dict), "workflow decision guards are missing")
    audit_latency = case_keys(audit.get("latency_violations", []))
    main_latency = case_keys(decision.get("latency_violations", []))
    audit_benefits = case_keys(audit.get("benefits", []))
    main_benefits = case_keys(decision.get("eligible_benefits", []))
    audit_resources = {
        (f"{row.get('shape')}/{row.get('mode')}", row.get("block"), row.get("metric"))
        for row in audit.get("resource_violations", [])
    }
    main_resources = {
        (row.get("case"), row.get("block"), row.get("metric"))
        for row in decision.get("resource_violations", [])
    }
    require(audit_latency == main_latency, "main/root latency policy disagreement")
    require(audit_benefits == main_benefits, "main/root benefit policy disagreement")
    require(audit_resources == main_resources, "main/root resource policy disagreement")
    require(audit.get("adoption_eligible") == decision.get("adoption_eligible"),
            "main/root adoption eligibility disagreement")
    verify_native_numerical_agreement(analysis, audit)
    verify_allocation_numerical_agreement(analysis, audit)
    require("production_adoption" not in audit,
            "raw audit must not claim production retention")
    require(analysis.get("source_disposition_contract", {}).get("retention_decision_external") is True,
            "analysis disposition boundary changed")
    return {"analysis": analysis, "audit": audit, "profile": profile}


def verify_decision_and_disposition(before: dict[str, Any], after: dict[str, Any],
                                   numeric: dict[str, Any]) -> dict[str, Any]:
    decision = read(P / "decision.json")
    require(decision.get("schema") == "litchi.performance.0810.decision.v1",
            "decision schema changed")
    eligible = numeric["analysis"]["decision_guards"]["adoption_eligible"]
    require(decision.get("adoption_eligible") == eligible
            and isinstance(decision.get("production_adoption"), bool),
            "decision policy result changed")
    disposition = read(P / "disposition.json")
    require(disposition.get("schema") == "litchi.performance.0810.disposition.v1",
            "disposition schema changed")
    retained = disposition.get("production_change_retained")
    require(isinstance(retained, bool)
            and disposition.get("status") in {"retained", "rejected"}
            and retained is (disposition["status"] == "retained")
            and decision["production_adoption"] is retained,
            "decision/disposition retention mismatch")
    if retained:
        require(decision.get("adoption_eligible") is True,
                "retained disposition is numerically ineligible")
    expected = after if retained else before
    require(current_source() == expected["files"],
            "live source does not match final disposition")
    if retained:
        require("restored_source" not in disposition,
                "retained disposition has a restoration witness")
    else:
        restored = packet_artifact(disposition.get("restored_source"), "restored source")
        require(source_manifest(restored, "restored source") == before,
                "restored source differs from baseline")
    return disposition


def verify_cleanup(final: bool, builds: dict[str, Any], disposition: dict[str, Any]) -> dict[str, Any] | None:
    path = P / "cleanup.json"
    if not path.is_file():
        require(not final, "final validation requires cleanup.json")
        require(TARGET.is_dir() and not TARGET.is_symlink(),
                "owned target disappeared before cleanup witness")
        return None
    cleanup = read(path)
    require(cleanup.get("schema") == "litchi.performance.0810.cleanup.v1"
            and cleanup.get("target") == str(TARGET)
            and cleanup.get("target_removed") is True
            and not TARGET.exists(), "cleanup witness is incomplete")
    require(type(cleanup.get("removed_files")) is int and cleanup["removed_files"] >= 6
            and type(cleanup.get("removed_logical_bytes")) is int
            and cleanup["removed_logical_bytes"] >= 0
            and cleanup.get("binaries_verified_before_removal") is True,
            "cleanup counts changed")
    removed = cleanup.get("removed_binaries")
    require(isinstance(removed, list) and len(removed) == 6,
            "cleanup binary witness count changed")
    expected = {
        (item["path"], item["bytes"], item["sha256"])
        for build in builds.values() for item in build["manifest"]["binaries"].values()
    }
    actual = {(item.get("path"), item.get("bytes"), item.get("sha256"))
              for item in removed if isinstance(item, dict)}
    require(actual == expected, "cleanup binary identities changed")
    for item in removed:
        require(not Path(item["path"]).exists(),
                f"removed binary remains: {item['path']}")
    source_leg = "after" if disposition["production_change_retained"] else "before"
    require(cleanup.get("source_leg") == source_leg,
            "cleanup source leg differs from disposition")
    source_path = packet_artifact(cleanup.get("source_manifest"), "cleanup source manifest")
    expected_source = builds[source_leg]["manifest"]["source"]
    require(source_path == packet_path(expected_source["path"], "cleanup expected source"),
            "cleanup source manifest path changed")
    return cleanup


def interval(rows: list[dict[str, Any]], label: str) -> tuple[float, float]:
    require(rows, f"{label}: no timestamped rows")
    first, last = math.inf, -math.inf
    previous = -math.inf
    for index, row in enumerate(rows):
        started, ended = row.get("started"), row.get("ended")
        require(isinstance(started, (int, float)) and math.isfinite(started)
                and isinstance(ended, (int, float)) and math.isfinite(ended)
                and previous <= started <= ended,
                f"{label}: chronology changed at {index}")
        first, last, previous = min(first, started), max(last, ended), ended
    return first, last


def timestamp_rows(path: Path, key: str) -> list[dict[str, Any]]:
    value = read(path)
    rows = value if key == "self" else value.get(key)
    require(isinstance(rows, list), f"{path.name}: {key} rows missing")
    return rows


def verify_chronology(
    intervals: dict[str, tuple[float, float]],
    chronology: dict[str, Any],
    cleanup: dict[str, Any] | None,
) -> None:
    """Validate receipt ordering against an immutable post-decision witness.

    Artifact mtimes are deliberately read from chronology.json.  A fresh
    checkout may rewrite all packet mtimes, so comparing them live would turn
    a valid committed packet into a false chronology failure.
    """
    require(chronology.get("schema") == "litchi.performance.0810.chronology.v1",
            "chronology schema changed")
    captured_at = chronology.get("captured_at")
    require(isinstance(captured_at, (int, float)) and not isinstance(captured_at, bool)
            and math.isfinite(captured_at) and captured_at > 0,
            "chronology capture time is invalid")
    files = chronology.get("files")
    expected_paths = {
        "application": P / "application.json",
        "qualification-audit": P / "qualification-audit.json",
        "analysis": P / "analysis.json",
        "root-audit": P / "root-audit.json",
        "profile-analysis": P / "profile-analysis.json",
        "decision": P / "decision.json",
        "disposition": P / "disposition.json",
    }
    require(isinstance(files, dict) and set(files) == set(expected_paths),
            "chronology artifact set changed")
    observed: dict[str, float] = {}
    for name, expected in expected_paths.items():
        value = files.get(name)
        require(isinstance(value, dict), f"chronology {name} witness is malformed")
        witnessed = packet_path(value.get("path"), f"chronology {name}")
        require(witnessed == expected,
                f"chronology {name} path changed")
        require(type(value.get("bytes")) is int and value["bytes"] >= 0,
                f"chronology {name} byte count is invalid")
        require(is_sha(value.get("sha256")),
                f"chronology {name} SHA-256 is invalid")
        require(witnessed.is_file() and not witnessed.is_symlink()
                and witnessed.stat().st_size == value["bytes"]
                and sha(witnessed) == value["sha256"],
                f"chronology {name} artifact identity changed")
        observed_ns = value.get("observed_mtime_ns")
        require(type(observed_ns) is int and observed_ns > 0,
                f"chronology {name} observed mtime is invalid")
        observed[name] = observed_ns / 1e9
        require(captured_at >= observed[name],
                f"chronology capture precedes {name} observation")

    def after(left: str, right: str) -> None:
        require(intervals[left][1] <= intervals[right][0],
                f"chronology overlap: {left} -> {right}")

    after("build-before", "probe-before")
    after("probe-before", "qualification")
    require(observed["application"] >= observed["qualification-audit"],
            "application precedes qualification audit")
    require(observed["application"] >= intervals["qualification"][1],
            "application precedes qualification completion")
    require(observed["application"] <= intervals["quality-after"][0],
            "after quality precedes candidate application")
    after("quality-after", "build-after")
    after("build-after", "probe-after")
    after("probe-after", "native")
    after("native", "allocation")
    after("allocation", "profiles")
    capture_end = max(intervals["native"][1], intervals["allocation"][1], intervals["profiles"][1])
    for name in ("analysis", "root-audit", "profile-analysis"):
        require(observed[name] >= capture_end,
                f"{name} precedes terminal capture")
    decision_time = observed["decision"]
    for name in ("analysis", "root-audit", "profile-analysis"):
        require(decision_time >= observed[name],
                f"decision precedes {name}")
    require(observed["disposition"] >= decision_time,
            "disposition precedes decision")
    if cleanup is not None:
        started, ended = cleanup.get("started"), cleanup.get("ended")
        require(isinstance(started, (int, float)) and not isinstance(started, bool)
                and math.isfinite(started)
                and isinstance(ended, (int, float)) and not isinstance(ended, bool)
                and math.isfinite(ended) and started > observed["disposition"]
                and ended >= started,
                "cleanup interval does not follow disposition witness")


def run_reader_checks() -> dict[str, str]:
    outputs = {}
    for script, output in READER_CHECKS:
        require((P / output).is_file(), f"reader output is missing: {output}")
        result = subprocess.run(
            [sys.executable, "-B", str(P / script), "--check"],
            cwd=ROOT,
            text=True,
            capture_output=True,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        )
        require(result.returncode == 0,
                f"{script} --check failed: {result.stdout}{result.stderr}")
        outputs[script] = (result.stdout + result.stderr).strip()
    return outputs


def validate(final: bool) -> dict[str, Any]:
    custody = verify_origin()
    verify_frozen_build_inputs()
    plan = verify_plan()
    cleanup = read(P / "cleanup.json") if (P / "cleanup.json").is_file() else None
    builds, before, after = verify_builds(cleanup)
    expected_after = candidate_sources(before)
    require(after == expected_after, "after build differs from candidate archive")
    verify_application(before, after)
    verify_lane("qualification", plan, builds, before, cleanup)
    verify_lane("native", plan, builds, after, cleanup)
    verify_lane("allocation", plan, builds, after, cleanup)
    intervals = {
        "build-before": interval(timestamp_rows(P / "build-before/build.json", "rows"), "build-before"),
        "probe-before": interval(timestamp_rows(P / "probe-quality-before/receipts.json", "self"), "probe-before"),
        "qualification": interval(timestamp_rows(P / "qualification/receipts.json", "self"), "qualification"),
        "quality-after": interval(timestamp_rows(P / "quality-after/checks.json", "self"), "quality-after"),
        "build-after": interval(timestamp_rows(P / "build-after/build.json", "rows"), "build-after"),
        "probe-after": interval(timestamp_rows(P / "probe-quality-after/receipts.json", "self"), "probe-after"),
        "native": interval(timestamp_rows(P / "native/receipts.json", "self"), "native"),
        "allocation": interval(timestamp_rows(P / "allocation/receipts.json", "self"), "allocation"),
        "profiles": interval(timestamp_rows(P / "profiles/receipts.json", "self"), "profiles"),
    }
    quality = read(P / "quality-summary.json")
    verify_quality_summary(after, custody["root_inputs"])
    numeric = verify_reader_agreement()
    disposition = verify_decision_and_disposition(before, after, numeric)
    cleanup = verify_cleanup(final, builds, disposition)
    chronology = read(P / "chronology.json")
    verify_chronology(intervals, chronology, cleanup)
    reader_outputs = run_reader_checks()
    return {
        "schema": "litchi.performance.0810.validation.v1",
        "final": final,
        "counts": {"reports": 310, "samples": 6718, "main_reports": 306,
                    "main_samples": 6714, "profile_reports": 4, "profile_samples": 4},
        "source": {"before_files": len(before["files"]), "after_files": len(after["files"]),
                    "allowlist": sorted(SOURCE_ALLOWLIST),
                    "retained": disposition["production_change_retained"]},
        "policy": {
            "adoption_eligible": numeric["analysis"]["decision_guards"]["adoption_eligible"],
            "main_root_agree": True,
        },
        "quality_summary_schema": quality["schema"],
        "cleanup": {"present": cleanup is not None, "required": final},
        "reader_checks": {name: "PASS" for name in reader_outputs},
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--final", action="store_true",
                        help="require the post-disposition six-binary cleanup witness")
    args = parser.parse_args()
    result = validate(args.final)
    print(json.dumps(result, indent=2, sort_keys=True))
    print("0810 aggregate validation PASS", flush=True)


if __name__ == "__main__":
    try:
        main()
    except AssertionError as error:
        print(f"0810 aggregate validation failed: {error}", file=sys.stderr)
        raise SystemExit(1)
