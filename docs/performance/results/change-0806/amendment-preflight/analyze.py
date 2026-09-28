"""Offline reader and protected-policy decision for the 0806 amendment lane.

This reader consumes only retained JSON, logs, source archives and binaries. It
never builds, runs a probe, invokes a profiler, or pools timing from 0805.
"""

from __future__ import annotations

import json
import math
import random
import re
import statistics
from pathlib import Path
from typing import Any

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
CASES = c.read(P / "cases.json")
EXPECTED_LABELS = [case["id"] for case in CASES]
PROTECTED = set(PLAN["policy"]["protected_consume_cases"])
HELPER_NAMES = [name for name in c.SOURCE_FILES if not name.endswith("-tests.rs")]
ORIGINAL_AFTER = P.parent / "candidate" / "after"
ORIGINAL_BEFORE = P.parent / "candidate" / "before"


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}: path missing")
    path = Path(raw)
    if not path.is_absolute():
        path = P / path
    path = path.resolve()
    require(path.is_relative_to(P.resolve()), f"{label}: path escapes packet")
    return path


def packet_artifact(descriptor: Any, label: str) -> Path:
    require(isinstance(descriptor, dict), f"{label}: descriptor missing")
    path = packet_path(descriptor.get("path"), label)
    require(isinstance(descriptor.get("bytes"), int) and descriptor["bytes"] >= 0,
            f"{label}: invalid byte count")
    require(isinstance(descriptor.get("sha256"), str) and len(descriptor["sha256"]) == 64,
            f"{label}: invalid digest")
    require(path.is_file() and not path.is_symlink(), f"{label}: missing {path}")
    require(path.stat().st_size == descriptor["bytes"] and c.sha(path) == descriptor["sha256"],
            f"{label}: artifact changed")
    return path


def cleanup_witness() -> dict[str, Any] | None:
    path = P / "cleanup.json"
    return c.read(path) if path.is_file() else None


def binary_artifact(descriptor: Any, label: str) -> None:
    require(isinstance(descriptor, dict), f"{label}: descriptor missing")
    path = Path(descriptor.get("path", ""))
    require(path.is_absolute(), f"{label}: binary path must be absolute")
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == descriptor.get("bytes")
                and c.sha(path) == descriptor.get("sha256"),
                f"{label}: binary changed")
        return
    witness = cleanup_witness()
    require(isinstance(witness, dict), f"{label}: missing binary without cleanup witness")
    require(witness.get("target") == str(c.TARGET)
            and witness.get("target_removed") is True
            and not c.TARGET.exists(),
            f"{label}: cleanup witness does not prove target removal")
    removed = witness.get("removed_binaries")
    require(isinstance(removed, list) and descriptor in removed,
            f"{label}: removed binary identity absent from cleanup witness")


def u64(value: int) -> int:
    return value & ((1 << 64) - 1)


def rotl(value: int, count: int) -> int:
    value = u64(value)
    return u64((value << count) | (value >> (64 - count)))


CONSTRUCT_STEP = 0x9E3779B97F4A7C15
CONSTRUCT_SEED = 0x6A09E667F3BCC909
VALUE_SEED = 0xBB67AE8584CAA73B
VALUE_FACTOR = 0xD6E8FEB86659FD93


def expected_error_marker(error: dict[str, Any]) -> int:
    kind = error.get("kind")
    if kind == "ExpectedEq":
        tag, second = 1, 0
        first = error["position"]
    elif kind == "ExpectedValue":
        tag, second = 2, 0
        first = error["position"]
    elif kind == "UnquotedValue":
        tag, second = 3, 0
        first = error["position"]
    elif kind == "ExpectedQuote":
        tag = 4
        first, second = error["position"], error["quote"]
    elif kind == "Duplicated":
        tag = 5
        first, second = error["position"], error["first_position"]
    else:
        raise AssertionError(f"unknown error kind {kind!r}")
    return u64(tag * CONSTRUCT_STEP) ^ rotl(first, 17) ^ rotl(second, 31)


def independent_expected(summary: dict[str, Any], mode: str, iterations: int) -> dict[str, int]:
    if mode == "construct":
        checksum = CONSTRUCT_SEED
        for index in range(iterations):
            checksum = u64(checksum + u64(index + 1) * CONSTRUCT_STEP)
        return {"checksum": checksum, "accepted": 0, "error_marker": 0}
    one_checksum = VALUE_SEED
    accepted = 0
    error_marker = 0
    for item in summary["items"]:
        if item["kind"] == "Attribute":
            accepted = u64(accepted + 1)
            checksum = (
                rotl(VALUE_SEED, 5)
                ^ u64(item["key_bytes"] * CONSTRUCT_STEP)
                ^ u64(item["value_bytes"] * VALUE_FACTOR)
            )
            one_checksum = rotl(one_checksum, 7) ^ checksum
        elif item["kind"] == "Error":
            error_marker = expected_error_marker(item["error"])
        else:
            raise AssertionError(f"unknown oracle item {item['kind']!r}")
    checksum = CONSTRUCT_SEED
    total_accepted = 0
    total_error = 0
    for index in range(iterations):
        checksum = u64(checksum + (one_checksum ^ u64(index + 1) * CONSTRUCT_STEP))
        total_accepted = u64(total_accepted + accepted)
        total_error = u64(total_error + error_marker)
    return {"checksum": checksum, "accepted": total_accepted, "error_marker": total_error}


def p50(values: list[float]) -> float:
    require(values, "empty process sample")
    ordered = sorted(values)
    return ordered[(len(ordered) + 1) // 2 - 1]


def bootstrap(values: list[float]) -> tuple[float, float, float]:
    require(len(values) == 6, "protected lane must have six paired ratios")
    rng = random.Random(PLAN["analysis"]["bootstrap_seed"])
    medians = [statistics.median(values[rng.randrange(len(values))] for _ in values)
               for _ in range(PLAN["analysis"]["bootstrap_resamples"])]
    medians.sort()
    low, high = PLAN["analysis"]["zero_based_endpoints"]
    return statistics.median(values), medians[low], medians[high]


def verify_archive_lineage() -> dict[str, Any]:
    before = c.archive_manifest("before")
    after = c.archive_manifest("after")
    for name in c.SOURCE_FILES:
        source_before = P / "source" / "before" / name
        source_after = P / "source" / "after" / name
        if name in c.SOURCE_FILES:
            require(source_before.is_file() and source_after.is_file(), f"missing archive {name}")
    # The original production leg must be the immutable 0806 before archive.
    for name in c.SOURCE_FILES:
        ref = ORIGINAL_BEFORE / name
        require(c.sha(P / "source" / "before" / name) == c.sha(ref),
                f"before archive drifted: {name}")
    # The shared tests are unchanged; only the five helper constructors are amended.
    test_name = "litchi-opc-xml_attributes-tests.rs"
    require(c.sha(P / "source" / "after" / test_name) == c.sha(ORIGINAL_AFTER / test_name),
            "amendment changed shared tests")
    for name in HELPER_NAMES:
        original = (ORIGINAL_AFTER / name).read_text(encoding="utf-8")
        amended = (P / "source" / "after" / name).read_text(encoding="utf-8")
        old = (
            "    #[allow(clippy::disallowed_methods)]\n"
            "    #[inline]\n"
            "    fn new(tag: &'a BytesStart<'a>) -> Self {\n"
            "        let mut attributes = tag.attributes();\n"
            "        attributes.with_checks(false);"
        )
        new = (
            "    #[inline]\n"
            "    fn new(tag: &'a BytesStart<'a>) -> Self {\n"
            "        let attributes = tag.unchecked_attributes();"
        )
        require(original.count(old) == 1, f"{name}: original constructor shape changed")
        require(amended == original.replace(old, new, 1), f"{name}: amendment is not mechanical")
        require(amended.count("let attributes = tag.unchecked_attributes();") == 1,
                f"{name}: constructor reuse missing")
        require(amended.count("attributes.with_checks(false);") == 1,
                f"{name}: unchecked helper implementation changed")
    return {"before": before, "after": after}


def verify_builds(archives: dict[str, Any]) -> dict[str, Any]:
    builds: dict[str, Any] = {}
    for leg in ("before", "after"):
        path = P / f"build-{leg}" / "build.json"
        require(path.is_file(), f"missing {leg} build")
        build = c.read(path)
        require(build.get("schema") == "litchi.performance.0806.amendment-build.v1",
                f"{leg}: build schema changed")
        require(build.get("archive") == archives[leg], f"{leg}: build archive mismatch")
        require(build.get("probe") == c.probe_manifest(), f"{leg}: probe custody mismatch")
        require(build.get("profiles") is False and build.get("callgrind") is False,
                f"{leg}: supplemental profile lane appeared")
        binary = build.get("binary")
        require(isinstance(binary, dict), f"{leg}: build binary missing")
        require(Path(binary.get("path", "")).resolve() == (c.TARGET / f"{leg}-native").resolve(),
                f"{leg}: binary escaped the dedicated target")
        binary_artifact(binary, f"{leg} build binary")
        packet_artifact(build["lock"], f"{leg} lock")
        command_row = build.get("command")
        require(isinstance(command_row, dict) and command_row.get("exit_code") == 0,
                f"{leg}: build command receipt missing")
        packet_artifact(command_row.get("log"), f"{leg} build log")
        expected_command = [
            "cargo", "build", "--offline", "--locked", "--release",
            "--manifest-path", str(P / "probe-src" / "Cargo.toml"),
            "--bin", "attribute-boundary-probe",
        ]
        require(command_row.get("command") == expected_command,
                f"{leg}: build command changed")
        commands_path = P / f"build-{leg}" / "commands.json"
        require(commands_path.is_file() and not commands_path.is_symlink(),
                f"{leg}: build command receipt missing")
        require(c.read(commands_path) == [command_row],
                f"{leg}: build command receipt disagrees")
        frozen_path = P / f"build-{leg}" / "frozen-inputs.json"
        require(frozen_path.is_file(), f"{leg}: frozen-inputs missing")
        frozen = c.read(frozen_path)
        require(frozen.get("leg") == leg, f"{leg}: frozen leg changed")
        require(frozen.get("archives") == {name: c.archive_manifest(name) for name in ("before", "after")},
                f"{leg}: frozen archives changed")
        require("analysis_driver" not in frozen and frozen.get("probe") == c.probe_manifest(),
                f"{leg}: reader was frozen into build inputs")
        expected_frozen = {
            "plan": c.sha(P / "plan.json"),
            "build_driver": c.sha(P / "build.py"),
            "quality_driver": c.sha(P / "quality.py"),
            "capture_driver": c.sha(P / "capture.py"),
            "custody_driver": c.sha(P / "custody.py"),
        }
        for key, value in expected_frozen.items():
            require(frozen.get(key) == value, f"{leg}: frozen input changed: {key}")
        source_path = P / "source.json"
        if source_path.is_file():
            require(frozen.get("workspace_source") == c.read(source_path),
                    f"{leg}: frozen workspace source changed")
        builds[leg] = build
    return builds


def verify_quality(archives: dict[str, Any]) -> dict[str, Any]:
    path = P / "quality" / "complete.json"
    require(path.is_file(), "quality completion missing")
    complete = c.read(path)
    require(complete.get("schema") == "litchi.performance.0806.amendment-quality.v1",
            "quality schema changed")
    require(complete.get("expected_test_counts") == {
        "before": PLAN["quality"]["before_test_count"],
        "after": PLAN["quality"]["after_test_count"],
    }, "quality expected test counts changed")
    require(complete.get("test_counts") == complete.get("expected_test_counts"),
            "quality test count gate did not pass")
    expected_archive_digests = {
        leg: {name: descriptor["sha256"] for name, descriptor in c.archive_manifest(leg).items()}
        for leg in ("before", "after")
    }
    require(complete.get("source_archives") == expected_archive_digests,
            "quality source archives changed")
    inputs_path = P / "quality" / "inputs.json"
    require(inputs_path.is_file(), "quality frozen inputs missing")
    inputs = c.read(inputs_path)
    require(inputs.get("archives") == {leg: {
        name: descriptor["sha256"] for name, descriptor in c.archive_manifest(leg).items()
    } for leg in ("before", "after")}, "quality frozen archives changed")
    require(inputs.get("plan") == c.sha(P / "plan.json"),
            "quality frozen plan changed")
    require(inputs.get("driver") == c.sha(P / "quality.py"),
            "quality frozen driver changed")
    for descriptor in inputs.get("source", {}).values():
        source_path = packet_artifact(descriptor, "quality frozen source")
        require(source_path == P / "source.json", "quality source custody changed")
    for leg in ("before", "after"):
        for name, source_path in c.SOURCE_FILES.items():
            mirror = P / "test-src" / leg / source_path
            require(mirror.is_file() and not mirror.is_symlink(),
                    f"quality {leg}: source mirror missing {source_path}")
            require(c.sha(mirror) == c.archive_manifest(leg)[name]["sha256"],
                    f"quality {leg}: source mirror changed {source_path}")
    receipt_path = P / "quality" / "receipts.json"
    require(receipt_path.is_file(), "quality receipts missing")
    receipts = c.read(receipt_path)
    require(complete.get("rows") == receipts and len(receipts) == 6,
            "quality receipt rows changed")
    expected_rows = []
    for leg in ("before", "after"):
        manifest = P / "test-src" / leg / "Cargo.toml"
        expected_rows.extend([
            (leg, ["cargo", "generate-lockfile", "--offline", "--manifest-path", str(manifest)]),
            (leg, ["cargo", "test", "--offline", "--locked", "--manifest-path", str(manifest),
                   "--workspace", "--", "--test-threads=2"]),
            (leg, ["cargo", "clippy", "--offline", "--locked", "--manifest-path", str(manifest),
                   "--workspace", "--all-targets", "--", "-D", "warnings"]),
        ])
    for receipt, (leg, command) in zip(receipts, expected_rows):
        require(receipt.get("leg") == leg and receipt.get("command") == command
                and receipt.get("exit_code") == 0, "quality command receipt changed")
        packet_artifact(receipt.get("log"), f"quality {leg} log")
        require(isinstance(receipt.get("started"), (int, float))
                and isinstance(receipt.get("ended"), (int, float))
                and receipt["ended"] >= receipt["started"],
                "quality timing metadata invalid")
    pattern = re.compile(r"test result: ok\. (\d+) passed; (\d+) failed;")
    previous_end = float("-inf")
    for receipt in receipts:
        require(receipt["started"] >= previous_end,
                "quality receipts overlap or are out of order")
        previous_end = receipt["ended"]
    for leg, index, expected in (("before", 1, 70), ("after", 4, 100)):
        log = packet_path(receipts[index]["log"]["path"], f"{leg} test log")
        passed = failed = 0
        for line in log.read_text(encoding="utf-8", errors="replace").splitlines():
            match = pattern.search(line)
            if match:
                passed += int(match.group(1))
                failed += int(match.group(2))
        require(passed == expected and failed == 0, f"{leg}: test log count mismatch")
    return complete


def report_for(receipt: dict[str, Any], case: dict[str, Any], leg: str, mode: str) -> dict[str, Any]:
    require(receipt.get("exit_code") == 0, f"failed child: {receipt}")
    report_descriptor = receipt.get("report")
    require(isinstance(report_descriptor, dict), "child omitted report descriptor")
    report_path = packet_path(report_descriptor["path"], "report")
    require(report_path.resolve().is_relative_to((P / "native").resolve()),
            "report path escapes native packet")
    expected_report = P / "native" / f"{receipt['block']}-{case['id']}-{mode}-{leg}.json"
    require(report_path == expected_report.resolve(), "report path changed")
    packet_artifact(report_descriptor, "report")
    report = c.read(report_path)
    require(report.get("schema") == "litchi.attribute-boundary-probe.v1", "report schema changed")
    require(report.get("tool") == "attribute-boundary-probe-0805", "probe lineage changed")
    require(report.get("binary") == f"{leg}-native", "report binary label changed")
    require(report.get("leg") == leg and report.get("mode") == mode, "report leg/mode mismatch")
    require(report.get("case") == case["id"], "report case mismatch")
    require(report.get("category") == case["category"]
            and report.get("attribute_count") == case["attribute_count"],
            "report case metadata changed")
    native = PLAN["native"]
    require(report.get("iterations") == native["iterations"]
            and report.get("warmup") == native["warmup"]
            and report.get("samples_requested") == native["samples"],
            "report native settings changed")
    require(report.get("source") == case["source"], "report source identity changed")
    oracle = report.get("semantic_oracle")
    require(isinstance(oracle, dict) and oracle.get("all_checks_passed") is True,
            "semantic oracle did not pass")
    require(oracle.get("baseline_matches_quick_xml") is True
            and oracle.get("candidate_matches_quick_xml") is True,
            "quick-xml oracle mismatch")
    require(oracle.get("quick_xml") == case["expected_baseline"], "quick-xml baseline drift")
    clone_checks = oracle.get("clone_checks")
    require(isinstance(clone_checks, list)
            and [clone.get("advance") for clone in clone_checks] == PLAN["clone_advances"],
            "clone oracle schedule missing or changed")
    for clone in clone_checks:
        require(clone.get("baseline_matches_quick_xml") is True
                and clone.get("candidate_matches_quick_xml") is True
                and clone.get("terminal_behavior_matches") is True,
                "clone oracle mismatch")
    sizes = report.get("iterator_sizes")
    require(sizes == {"baseline_checked_attributes": 120, "candidate_checked_attributes": 128},
            "iterator layout changed")
    expected = report.get("expected_result")
    samples = report.get("samples")
    require(isinstance(expected, dict) and isinstance(samples, list)
            and len(samples) == native["samples"], "sample shape changed")
    independent = independent_expected(case["expected_baseline"], mode, native["iterations"])
    require(expected == independent, "expected loop result is not independently reproducible")
    for index, sample in enumerate(samples):
        require(sample.get("index") == index, "sample indices are not contiguous")
        require(isinstance(sample.get("elapsed_ns"), int) and sample["elapsed_ns"] >= 0,
                "invalid elapsed sample")
        for key in ("checksum", "accepted", "error_marker"):
            require(sample.get(key) == expected.get(key), f"sample result drift: {key}")
    return report


def verify_capture(builds: dict[str, Any], archives: dict[str, Any]) -> dict[str, dict[str, dict[int, dict[str, Any]]]]:
    complete_path = P / "native" / "complete.json"
    require(complete_path.is_file(), "native completion missing")
    complete = c.read(complete_path)
    native = PLAN["native"]
    expected_children = native["blocks"] * PLAN["case_count"] * len(PLAN["modes"]) * 2
    require(complete.get("schema") == "litchi.performance.0806.amendment-native.v1",
            "native completion schema changed")
    require(complete.get("children") == expected_children
            and complete.get("expected_children") == expected_children,
            "native child count changed")
    require(complete.get("samples") == expected_children * native["samples"],
            "native sample count changed")
    packet_artifact(complete.get("receipts"), "native completion receipts")
    packet_artifact(complete.get("source"), "native completion source")
    require(packet_path(complete["receipts"]["path"], "native receipts")
            == (P / "native" / "receipts.json").resolve(),
            "native completion receipts path changed")
    require(packet_path(complete["source"]["path"], "native source")
            == (P / "native" / "source.json").resolve(),
            "native completion source path changed")
    require(complete.get("profiles") is False and complete.get("callgrind") is False,
            "native completion claims a forbidden profile")
    capture_source = c.read(P / "native" / "source.json")
    require(capture_source.get("schema") == "litchi.performance.0806.amendment-capture-source.v1",
            "native source schema changed")
    require(capture_source.get("archives") == archives
            and capture_source.get("probe") == c.probe_manifest(),
            "native source/probe witness changed")
    receipts_path = P / "native" / "receipts.json"
    receipts = c.read(receipts_path)
    require(len(receipts) == expected_children, "receipt count changed")
    case_map = {case["id"]: case for case in CASES}
    reports: dict[str, dict[str, dict[int, dict[str, Any]]]] = {
        case_id: {mode: {block: {} for block in range(native["blocks"])} for mode in PLAN["modes"]}
        for case_id in EXPECTED_LABELS
    }
    expected_order = []
    for block, order in enumerate(native["orders"]):
        for case in CASES:
            for mode in PLAN["modes"]:
                for leg in order:
                    expected_order.append((block, case["id"], mode, leg))
    actual_order = [(row.get("block"), row.get("case"), row.get("mode"), row.get("leg"))
                    for row in receipts]
    require(actual_order == expected_order, "receipt order or duplicate schedule changed")
    report_paths: set[str] = set()
    log_paths: set[str] = set()
    rss_paths: set[str] = set()
    previous_end = float("-inf")
    for receipt in receipts:
        require(receipt.get("schema") == "litchi.performance.0806.amendment-native-receipt.v1",
                "native receipt schema changed")
        require(isinstance(receipt.get("started"), (int, float))
                and isinstance(receipt.get("ended"), (int, float))
                and receipt["ended"] >= receipt["started"],
                "native receipt timing metadata invalid")
        require(receipt["started"] >= previous_end, "native receipts overlap or are out of order")
        previous_end = receipt["ended"]
        leg = receipt.get("leg")
        case_id = receipt.get("case")
        mode = receipt.get("mode")
        block = receipt.get("block")
        require(leg in ("before", "after") and case_id in case_map and mode in PLAN["modes"]
                and isinstance(block, int) and 0 <= block < native["blocks"],
                "receipt identity changed")
        binary = receipt.get("binary")
        require(binary == builds[leg]["binary"], f"{case_id}: binary custody mismatch")
        command = receipt.get("command")
        stem = f"{block}-{case_id}-{mode}-{leg}"
        expected_log = (P / "native" / f"{stem}.log").resolve()
        expected_rss = (P / "native" / f"{stem}.rss").resolve()
        expected_report = (P / "native" / f"{stem}.json").resolve()
        expected_command = [
            "/usr/bin/time", "-f", "%M", "-o", str(expected_rss),
            "taskset", "-c", str(PLAN["cpu"]), binary["path"],
            "--leg", leg, "--case", case_id, "--mode", mode,
            "--samples", str(native["samples"]), "--warmup", str(native["warmup"]),
            "--iterations", str(native["iterations"]),
            "--output", str(expected_report),
        ]
        require(command == expected_command, f"{case_id}/{mode}/{block}: command changed")
        for key, paths in (("log", log_paths), ("rss", rss_paths)):
            descriptor = receipt.get(key)
            packet_artifact(descriptor, f"{case_id}/{mode}/{block}/{leg} {key}")
            actual = packet_path(descriptor["path"], key)
            expected = expected_log if key == "log" else expected_rss
            require(actual == expected, f"{case_id}/{mode}/{block}/{leg}: {key} path changed")
            paths.add(str(actual))
        report = report_for(receipt, case_map[case_id], leg, mode)
        report_descriptor = receipt["report"]
        report_paths.add(str(packet_path(report_descriptor["path"], "report")))
        reports[case_id][mode][block][leg] = report
    require(len(report_paths) == expected_children, "duplicate report paths")
    require(len(log_paths) == expected_children and len(rss_paths) == expected_children,
            "duplicate log or RSS paths")
    for case_id in EXPECTED_LABELS:
        for mode in PLAN["modes"]:
            for block in range(native["blocks"]):
                require(set(reports[case_id][mode][block]) == {"before", "after"},
                        f"missing pair {case_id}/{mode}/{block}")
    return reports


def summarize(reports: dict[str, dict[str, dict[int, dict[str, Any]]]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    rows: list[dict[str, Any]] = []
    protected: list[dict[str, Any]] = []
    for case_id in EXPECTED_LABELS:
        for mode in PLAN["modes"]:
            p50s: dict[str, list[float]] = {"before": [], "after": []}
            ratios = []
            for block in range(PLAN["native"]["blocks"]):
                pair = reports[case_id][mode][block]
                before = p50([sample["elapsed_ns"] for sample in pair["before"]["samples"]])
                after = p50([sample["elapsed_ns"] for sample in pair["after"]["samples"]])
                p50s["before"].append(before)
                p50s["after"].append(after)
                require(before > 0, f"zero before p50 for {case_id}/{mode}/{block}")
                ratios.append(after / before)
            ratio, low, high = bootstrap(ratios)
            row = {
                "case": case_id,
                "mode": mode,
                "process_p50_before": p50s["before"],
                "process_p50_after": p50s["after"],
                "paired_ratios": ratios,
                "ratio_median": ratio,
                "bootstrap_ci_low": low,
                "bootstrap_ci_high": high,
                "change_percent_median": (ratio - 1.0) * 100.0,
                "diagnostic_regression": ratio > PLAN["analysis"]["diagnostic_ratio_above"] and low > PLAN["analysis"]["diagnostic_ci_low_above"],
            }
            rows.append(row)
            if mode == "consume" and case_id in PROTECTED and row["diagnostic_regression"]:
                protected.append(row)
    return rows, protected


def main() -> None:
    require(PLAN["schema"] == "litchi.performance.0806.amendment-preflight.v1", "plan schema changed")
    require(len(CASES) == 39 and EXPECTED_LABELS == [case["id"] for case in CASES], "case schedule changed")
    archives = verify_archive_lineage()
    quality = verify_quality(archives)
    builds = verify_builds(archives)
    reports = verify_capture(builds, archives)
    rows, protected = summarize(reports)
    benefits = {}
    for case_id in PLAN["policy"]["benefit_cases"]:
        row = next(row for row in rows if row["case"] == case_id and row["mode"] == "consume")
        benefits[case_id] = (
            row["ratio_median"] <= PLAN["policy"]["benefit_ratio_at_most"]
            and row["bootstrap_ci_high"] < PLAN["policy"]["benefit_ci_high_below"]
        )
    advance = bool(all(benefits.values()) and not protected)
    analysis_payload = {
        "schema": "litchi.performance.0806.amendment-analysis.v1",
        "rows": rows,
        "policy": PLAN["policy"],
        "claims": PLAN["scope"],
    }
    c.write(P / "analysis.json", analysis_payload)
    # Finalization is bound to the separate raw replay.  The raw auditor reads
    # the retained reports directly and independently reconstructs every row;
    # the small policy-only helper is not sufficient as final custody.
    audit_path = P / "root-native-audit.json"
    if not audit_path.is_file():
        c.write(P / "decision-pending.json", {
            "schema": "litchi.performance.0806.amendment-decision-pending.v1",
            "analysis": c.relative_artifact(P / "analysis.json"),
            "required_audit": "root-native-audit.json",
            "advance_to_workflow_trials": advance,
            "production_adoption": False,
            "protected_consume_regressions": protected,
            "dominant_class_benefits": {
                "distinct-1": bool(benefits["distinct-1"]),
                "distinct-2": bool(benefits["distinct-2"]),
            },
        })
        print("analysis written; decision pending root-native-audit.json")
        return
    audit = c.read(audit_path)
    require(audit.get("schema") == "litchi.performance.0806.amendment-root-native-audit.v1",
            "independent audit schema changed")
    require(audit.get("passed") is True
            and audit.get("production_adoption") is False
            and audit.get("matches_primary_analysis") is True,
            "independent audit did not pass")
    require(audit.get("native_reports") == 936 and audit.get("native_samples") == 28080,
            "independent audit cardinality changed")
    require(audit.get("advance_to_workflow_trials") is advance,
            "independent audit advance decision disagrees")
    require(audit.get("protected_consume_regressions") == protected,
            "independent audit protected rows disagree")
    require(audit.get("dominant_class_benefits") == {
        "distinct-1": bool(benefits["distinct-1"]),
        "distinct-2": bool(benefits["distinct-2"]),
    }, "independent audit benefit decision disagrees")
    require(audit.get("rows") == rows,
            "independent audit rows disagree with this analysis")
    decision = {
        "schema": "litchi.performance.0806.amendment-decision.v1",
        "legacy_schema": "litchi.performance.0805.preflight-decision.v1",
        "packet": "change-0806/amendment-preflight",
        "seed": PLAN["analysis"]["bootstrap_seed"],
        "advance_to_workflow_trials": advance,
        "production_adoption": False,
        "protected_consume_regressions": protected,
        "all_consume_regressions": [
            {
                "case": row["case"],
                "mode": row["mode"],
                "ratio_median": row["ratio_median"],
                "bootstrap_ci_low": row["bootstrap_ci_low"],
                "bootstrap_ci_high": row["bootstrap_ci_high"],
                "change_percent_median": row["change_percent_median"],
            }
            for row in rows
            if row["mode"] == "consume"
            and row["ratio_median"] > PLAN["analysis"]["diagnostic_ratio_above"]
            and row["bootstrap_ci_low"] > PLAN["analysis"]["diagnostic_ci_low_above"]
        ],
        "dominant_class_benefits": {
            "distinct-1": bool(benefits["distinct-1"]),
            "distinct-2": bool(benefits["distinct-2"]),
        },
        "benefit_policy_passed": advance,
        "independent_audit": c.relative_artifact(audit_path),
        "counts": {
            "case_count": len(CASES),
            "native_reports": len(rows) * PLAN["native"]["blocks"] * 2,
            "native_children": PLAN["native"]["blocks"] * len(CASES) * len(PLAN["modes"]) * 2,
            "native_samples": PLAN["native"]["blocks"] * len(CASES) * len(PLAN["modes"]) * 2 * PLAN["native"]["samples"],
        },
        "analysis": c.relative_artifact(P / "analysis.json"),
        "custody": {
            "archives": archives,
            "builds": {leg: c.relative_artifact(P / f"build-{leg}" / "build.json") for leg in ("before", "after")},
            "native": c.relative_artifact(P / "native" / "complete.json"),
            "probe": c.probe_manifest(),
            "source": c.relative_artifact(P / "source.json") if (P / "source.json").is_file() else None,
            "quality": c.relative_artifact(P / "quality" / "complete.json"),
            "profiles": False,
            "callgrind": False,
            "historical_timing_pooling": False,
        },
    }
    c.write(P / "decision.json", decision)


if __name__ == "__main__":
    main()
