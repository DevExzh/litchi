"""Root-owned ordinary-save artifact, qualification, and capture driver.

Usage is intentionally explicit:

    python3 -B capture.py artifacts before
    python3 -B capture.py qualification before
    python3 -B capture.py artifacts after
    python3 -B capture.py qualification after
    python3 -B capture.py native
    python3 -B capture.py observer

The first four commands are run while the matching source leg is installed.
The comparative lanes run both binaries against the restored shipped after
source; the binary and build source receipts identify which leg ran.
"""

from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path
from typing import Any

import custody as c


if len(sys.argv) == 3:
    KIND, LEG = sys.argv[1:]
    assert KIND in {"artifacts", "qualification"} and LEG in {"before", "after"}
    LANE = f"{KIND}-{LEG}"
elif len(sys.argv) == 2:
    KIND = sys.argv[1]
    assert KIND in {"native", "observer"}
    LEG = None
    LANE = KIND
else:
    raise AssertionError("use artifacts before|after, qualification before|after, native, or observer")

PLAN = c.plan()
FROZEN = c.read(c.P / "freeze.json")
c.assert_static()
c.assert_quality()
c.stable_inputs(FROZEN)
c.check_no_overrides()


def build_for(leg: str) -> dict[str, Any]:
    path = c.P / f"build-{leg}/build.json"
    assert path.is_file(), f"missing {leg} build receipt"
    value = c.read(path)
    assert value["schema"] == f"litchi.performance.0825.build-{leg}.v1"
    source_path = c.resolve_descriptor(value["source"])
    assert c.artifact(source_path) == value["source"]
    build_source = c.read(source_path)
    assert build_source["revision"] == c.BASE
    expected = c.candidate_files(leg)
    assert {name: build_source["files"][name] for name in c.ALLOWLIST} == expected
    frozen_source = FROZEN["source"]
    if leg == "before":
        assert c.changed_files(frozen_source, build_source) == set(c.ALLOWLIST)
    else:
        assert build_source == frozen_source
    assert value["quality"] == FROZEN["quality"]
    assert value["candidate_archives"] == FROZEN["candidate_archives"]
    for logical, spec in PLAN["binaries"].items():
        binary = value["binaries"][logical]
        assert binary["cargo_bin"] == spec["cargo_bin"]
        assert binary["features"] == spec["features"]
        descriptor = binary["artifact"]
        path = c.resolve_descriptor(descriptor)
        assert c.artifact(path) == descriptor
        assert path.name == f"{leg}-{logical}"
    return value


def stable_leg(leg: str) -> None:
    c.stable_inputs(FROZEN)
    assert c.assert_leg_source(leg, FROZEN)["revision"] == c.BASE


def stable_comparative() -> None:
    c.stable_inputs(FROZEN)
    assert c.assert_leg_source("after", FROZEN)["revision"] == c.BASE


def input_rows() -> list[dict[str, Any]]:
    corpus = c.corpus()
    result = []
    seen = set()
    for case in c.plan_cases(PLAN):
        if case["input"] in seen:
            continue
        seen.add(case["input"])
        result.append({
            "path": case["input"],
            "absolute": str(c.ROOT / case["input"]),
            **corpus[case["input"]],
        })
    assert len(result) == 3
    return result


INPUTS = input_rows()


def binary_descriptor(build: dict[str, Any], logical: str) -> dict[str, Any]:
    descriptor = build["binaries"][logical]["artifact"]
    assert c.artifact(c.resolve_descriptor(descriptor)) == descriptor
    return descriptor


def run_command(command: list[str], log: Path) -> tuple[int, float, float]:
    started = time.time()
    with log.open("x", encoding="utf-8") as stream:
        completed = subprocess.run(
            command, cwd=c.ROOT, env=os.environ | {"PYTHONDONTWRITEBYTECODE": "1"},
            stdout=stream, stderr=subprocess.STDOUT,
        )
    return completed.returncode, started, time.time()


def run_artifact_export(leg: str) -> None:
    stable_leg(leg)
    build = build_for(leg)
    output = c.P / PLAN["paths"]["artifact_output"][leg]
    complete_path = c.P / PLAN["paths"]["artifact_complete"][leg]
    receipt_path = c.P / f"artifacts-{leg}-receipt.json"
    source_path = c.P / f"artifacts-{leg}-source.json"
    log = c.P / f"artifacts-{leg}.log"
    assert not output.exists() and not complete_path.exists()
    assert not receipt_path.exists() and not source_path.exists() and not log.exists()
    scratch_marker = c.make_scratch()
    source = c.assert_leg_source(leg, FROZEN)
    c.write(source_path, source)
    binary = binary_descriptor(build, "artifacts")
    command = [
        "taskset", "-c", str(PLAN["cpu"]), binary["path"],
        "--output", str(output),
        "--filesystem-root", str(c.SCRATCH),
    ]
    for item in INPUTS:
        # Keep the manifest's caller-named path stable and packet-readable;
        # the exporter runs with the repository root as its working directory.
        command.extend(["--ooxml-file", item["path"]])
    exit_code, started, ended = run_command(command, log)
    row: dict[str, Any] = {
        "schema": "litchi.performance.0825.artifact-capture-receipt.v1",
        "lane": "artifacts",
        "leg": leg,
        "timing": "untimed_artifact_export",
        "command": command,
        "started": started,
        "ended": ended,
        "exit_code": exit_code,
        "binary": build["binaries"]["artifacts"],
        "source": c.artifact(source_path),
        "frozen_inputs": c.artifact(c.P / "freeze.json"),
        "scratch_marker": scratch_marker,
        "inputs": INPUTS,
        "log": c.artifact(log),
    }
    if output.is_dir() and (output / "manifest.json").is_file():
        row["manifest"] = c.artifact(output / "manifest.json")
    c.write(receipt_path, row)
    if exit_code != 0:
        raise RuntimeError(f"0825 artifact export failed; retained {log}")
    assert output.is_dir() and (output / "manifest.json").is_file()
    manifest = c.read(output / "manifest.json")
    assert manifest.get("schema_version") == 1
    assert manifest.get("kind") == "ordinary-save-artifact-export"
    exported = manifest.get("cases")
    assert isinstance(exported, list) and len(exported) == 6
    assert sum(row.get("origin") == "generated-harness-corpus" for row in exported) == 3
    assert sum(row.get("origin") == "caller-named-real-file" for row in exported) == 3
    expected_inputs = {item["sha256"] for item in INPUTS}
    real_inputs = set()
    for exported_case in exported:
        policies = exported_case.get("policy_outputs")
        assert isinstance(policies, list) and len(policies) == 4
        descriptors = [exported_case["source_archive"]]
        descriptors.extend(item["output"] for item in policies)
        descriptors.append(exported_case["stream_output"]["output"])
        for descriptor in descriptors:
            path = output / descriptor["path"]
            assert path.is_file() and not path.is_symlink()
            assert c.artifact(path)["sha256"] == descriptor["sha256"]
        if exported_case["origin"] == "caller-named-real-file":
            real_inputs.add(exported_case["source_archive_sha256"])
    assert real_inputs == expected_inputs
    complete = {
        "schema": f"litchi.performance.0825.artifacts-{leg}.complete.v1",
        "status": "pass",
        "leg": leg,
        "children": 1,
        "reports": 0,
        "samples": 0,
        "cases": len(exported),
        "policy_outputs_per_case": 5,
        "plan": c.artifact(c.P / "plan.json"),
        "build": c.artifact(c.P / f"build-{leg}/build.json"),
        "source": c.artifact(source_path),
        "receipt": c.artifact(receipt_path),
        "manifest": c.artifact(output / "manifest.json"),
    }
    c.write(complete_path, complete)
    stable_leg(leg)
    print(f"0825 artifacts {leg} PASS: six cases and five untimed policy outputs per case", flush=True)


def require_complete(leg: str) -> dict[str, Any]:
    path = c.P / PLAN["paths"]["qualification_complete"][leg]
    assert path.is_file(), f"qualification {leg} must pass before comparative capture"
    value = c.read(path)
    assert value["status"] == "pass"
    assert value["reports"] == PLAN["expected"]["qualification_reports_per_leg"]
    assert value["samples"] == PLAN["expected"]["qualification_samples_per_leg"]
    admission_path = c.P / f"qualification-admission-{leg}.json"
    assert admission_path.is_file(), f"independent qualification admission {leg} is required"
    admission = c.read(admission_path)
    assert admission["schema"] == "litchi.performance.0825.qualification-admission.v1"
    assert admission["accepted"] is True and admission["leg"] == leg
    assert admission["plan_sha256"] == c.sha(c.P / "plan.json")
    assert admission["reports"] == 12 and admission["samples"] == 12
    assert admission["qualification_complete"] == c.artifact(path)
    c.verify_descriptor(admission["qualification_receipts"])
    c.verify_descriptor(admission["artifact_admission"])
    return value


def require_admissions() -> dict[str, dict[str, Any]]:
    return {leg: c.require_admission(leg) for leg in ("before", "after")}


def run_qualification(leg: str) -> None:
    stable_leg(leg)
    c.make_scratch()
    build = build_for(leg)
    admission = c.require_admission(leg)
    output = c.P / PLAN["paths"]["qualification_output"][leg]
    assert not output.exists()
    output.mkdir()
    receipts_path = output / "receipts.json"
    source_path = output / "source.json"
    c.write(source_path, c.assert_leg_source(leg, FROZEN))
    lane_plan = PLAN["lanes"]["qualification"]
    binary = binary_descriptor(build, "observer")
    rows: list[dict[str, Any]] = []
    for index, case in enumerate(c.plan_cases(PLAN)):
        stem = f"{index:02d}-{case['format']}-{case['phase']}"
        report = output / f"{stem}.json"
        rss = output / f"{stem}.rss"
        log = output / f"{stem}.log"
        assert not report.exists() and not rss.exists() and not log.exists()
        command = [
            "/usr/bin/time", "-f", "%M", "-o", str(rss),
            "taskset", "-c", str(PLAN["cpu"]), binary["path"],
            "--warmup", str(lane_plan["warmup"]),
            "--samples", str(lane_plan["samples"]), "--case", case["case"],
            "--json", str(report), "--filesystem-root", str(c.SCRATCH),
            "--ooxml-file", str(c.ROOT / case["input"]),
        ]
        exit_code, started, ended = run_command(command, log)
        row = {
            "schema": "litchi.performance.0825.capture-receipt.v1",
            "lane": "qualification",
            "leg": leg,
            "block": 0,
            **case,
            "samples": lane_plan["samples"],
            "warmup": lane_plan["warmup"],
            "command": command,
            "started": started,
            "ended": ended,
            "exit_code": exit_code,
            "binary": binary,
            "source": c.artifact(source_path),
            "frozen_inputs": c.artifact(c.P / "freeze.json"),
            "artifact_admission": c.artifact(c.P / PLAN["paths"]["artifact_admission"][leg]),
            "log": c.artifact(log),
        }
        if report.exists():
            row["report"] = c.artifact(report)
        if rss.exists():
            row["rss"] = c.artifact(rss)
        rows.append(row)
        c.write(receipts_path, rows)
        if exit_code != 0:
            raise RuntimeError(f"0825 qualification {leg} failed; retained {log}")
        assert report.is_file() and rss.is_file()
        c.assert_rss(rss)
        c.validate_report(report, case, lane_plan["samples"], lane_plan["warmup"], "observer", admission)
        stable_leg(leg)
        print(f"0825 qualification {leg} {stem} PASS", flush=True)
    complete = {
        "schema": f"litchi.performance.0825.qualification-{leg}.complete.v1",
        "status": "pass",
        "lane": "qualification",
        "leg": leg,
        "blocks": 1,
        "reports": len(rows),
        "samples": sum(row["samples"] for row in rows),
        "expected_reports": PLAN["expected"]["qualification_reports_per_leg"],
        "expected_samples": PLAN["expected"]["qualification_samples_per_leg"],
        "plan": c.artifact(c.P / "plan.json"),
        "build": c.artifact(c.P / f"build-{leg}/build.json"),
        "source": c.artifact(source_path),
        "receipts": c.artifact(receipts_path),
    }
    assert complete["reports"] == complete["expected_reports"]
    assert complete["samples"] == complete["expected_samples"]
    c.write(output / "complete.json", complete)
    stable_leg(leg)
    print(f"0825 qualification {leg} PASS: {len(rows)} reports/{complete['samples']} samples", flush=True)


def run_comparative(lane: str) -> None:
    assert c.assert_leg_source("after", FROZEN)["revision"] == c.BASE
    c.make_scratch()
    before_build = build_for("before")
    after_build = build_for("after")
    require_admissions()
    require_complete("before")
    require_complete("after")
    output = c.P / lane
    assert not output.exists()
    output.mkdir()
    receipts_path = output / "receipts.json"
    source_path = output / "source.json"
    c.write(source_path, c.assert_leg_source("after", FROZEN))
    lane_plan = PLAN["lanes"][lane]
    admissions = require_admissions()
    builds = {"before": before_build, "after": after_build}
    logical_binary = lane_plan["binary"]
    rows: list[dict[str, Any]] = []
    for block, order in enumerate(lane_plan["orders"]):
        order_label = "BA" if order == ["before", "after"] else "AB"
        for case in c.plan_cases(PLAN):
            for leg in order:
                build = builds[leg]
                binary = binary_descriptor(build, logical_binary)
                stem = f"{block:02d}-{leg}-{case['format']}-{case['phase']}"
                report = output / f"{stem}.json"
                rss = output / f"{stem}.rss"
                log = output / f"{stem}.log"
                assert not report.exists() and not rss.exists() and not log.exists()
                command = [
                    "/usr/bin/time", "-f", "%M", "-o", str(rss),
                    "taskset", "-c", str(PLAN["cpu"]), binary["path"],
                    "--warmup", str(lane_plan["warmup"]),
                    "--samples", str(lane_plan["samples"]), "--case", case["case"],
                    "--json", str(report), "--filesystem-root", str(c.SCRATCH),
                    "--ooxml-file", str(c.ROOT / case["input"]),
                ]
                exit_code, started, ended = run_command(command, log)
                row = {
                    "schema": "litchi.performance.0825.capture-receipt.v1",
                    "lane": lane,
                    "leg": leg,
                    "block": block,
                    "order": order_label,
                    **case,
                    "samples": lane_plan["samples"],
                    "warmup": lane_plan["warmup"],
                    "command": command,
                    "started": started,
                    "ended": ended,
                    "exit_code": exit_code,
                    "binary": binary,
                    "source": c.artifact(source_path),
                    "build_source": build["source"],
                    "frozen_inputs": c.artifact(c.P / "freeze.json"),
                    "log": c.artifact(log),
                }
                if report.exists():
                    row["report"] = c.artifact(report)
                if rss.exists():
                    row["rss"] = c.artifact(rss)
                rows.append(row)
                c.write(receipts_path, rows)
                if exit_code != 0:
                    raise RuntimeError(f"0825 {lane} {leg} failed; retained {log}")
                assert report.is_file() and rss.is_file()
                c.assert_rss(rss)
                c.validate_report(
                    report, case, lane_plan["samples"], lane_plan["warmup"],
                    logical_binary, admissions[leg],
                )
                stable_comparative()
                print(f"0825 {lane} {stem} PASS", flush=True)
    expected_reports = PLAN["expected"][f"{lane}_reports"]
    expected_samples = PLAN["expected"][f"{lane}_samples"]
    complete = {
        "schema": f"litchi.performance.0825.{lane}.complete.v1",
        "status": "pass",
        "lane": lane,
        "blocks": len(lane_plan["orders"]),
        "reports": len(rows),
        "samples": sum(row["samples"] for row in rows),
        "expected_reports": expected_reports,
        "expected_samples": expected_samples,
        "plan": c.artifact(c.P / "plan.json"),
        "freeze": c.artifact(c.P / "freeze.json"),
        "source": c.artifact(source_path),
        "receipts": c.artifact(receipts_path),
    }
    assert complete["reports"] == expected_reports
    assert complete["samples"] == expected_samples
    c.write(output / "complete.json", complete)
    stable_comparative()
    print(f"0825 {lane} PASS: {len(rows)} reports/{complete['samples']} samples", flush=True)


if KIND == "artifacts":
    run_artifact_export(LEG)
elif KIND == "qualification":
    run_qualification(LEG)
else:
    run_comparative(KIND)
