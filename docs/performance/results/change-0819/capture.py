"""Root-only serial exporter and capture driver for the 0819 packet."""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

import custody as c


LANE = sys.argv[1] if len(sys.argv) == 2 else ""
assert LANE in ("artifacts", "qualification", "native", "observer"), (
    "choose exactly one lane: artifacts, qualification, native, or observer"
)
P = c.P
ROOT = c.ROOT
PLAN = c.read(P / "plan.json")
ORIGIN = c.assert_origin()
BUILD = c.read(P / "build.json")
QUALITY = c.read(P / "quality.json")
assert PLAN["schema"] == "litchi.performance.0819.plan.v1"
assert PLAN["base"] == ORIGIN["base"]
assert PLAN["target"] == str(c.TARGET)
assert PLAN["scratch"] == str(c.SCRATCH)
cases = c.plan_cases(PLAN)
expected_counts = PLAN["expected"]
assert BUILD["schema"] == "litchi.performance.0819.build.v1"
assert QUALITY["schema"] == "litchi.performance.0819.quality.v1"
assert QUALITY["status"] == "pass" and QUALITY["gate_count"] == 6
assert len(QUALITY["rows"]) == 6
assert all(row["exit_code"] == 0 for row in QUALITY["rows"])
assert BUILD["root_inputs"] == QUALITY["root_inputs"]
assert BUILD["locks"] == QUALITY["locks"]
assert BUILD["architecture"] == QUALITY["architecture"]
assert BUILD["corpus"] == QUALITY["corpus"]
assert BUILD["unrelated"] == QUALITY["unrelated"]

BUILD_SOURCE = c.read(BUILD["source"]["path"])
assert isinstance(BUILD_SOURCE, dict)
c.unchanged(BUILD_SOURCE)
root_inputs = BUILD["root_inputs"]
locks = BUILD["locks"]
architecture = BUILD["architecture"]
corpus = BUILD["corpus"]
host = BUILD["host"]
packet = BUILD["packet"]
drivers = BUILD["drivers"]
unrelated = BUILD["unrelated"]
c.check_stable(
    BUILD_SOURCE, root_inputs, locks, architecture, corpus, host,
    packet, drivers, unrelated,
)
QUALITY_SOURCE = c.read(QUALITY["source"]["path"])
assert QUALITY_SOURCE == BUILD_SOURCE
assert c.artifact(BUILD["binaries"]["native"]["artifact"]["path"]) == BUILD["binaries"]["native"]["artifact"]
assert c.artifact(BUILD["binaries"]["observer"]["artifact"]["path"]) == BUILD["binaries"]["observer"]["artifact"]
assert c.artifact(BUILD["binaries"]["artifacts"]["artifact"]["path"]) == BUILD["binaries"]["artifacts"]["artifact"]

def input_rows() -> list[dict[str, object]]:
    result = []
    seen = set()
    for case in cases:
        name = case["input"]
        if name in seen:
            continue
        seen.add(name)
        result.append(
            {
                "path": name,
                "absolute": str(ROOT / name),
                "bytes": corpus[name]["bytes"],
                "sha256": corpus[name]["sha256"],
            }
        )
    assert len(result) == 3
    return result


INPUTS = input_rows()


def stable() -> None:
    c.check_stable(
        BUILD_SOURCE, root_inputs, locks, architecture, corpus, host,
        packet, drivers, unrelated,
    )


def assert_rss(path: Path) -> None:
    text = path.read_text().strip()
    assert text.isdigit() and int(text) > 0, f"invalid RSS receipt {path}: {text!r}"


SCRATCH_MARKER = "litchi-performance-0819-owned-scratch-v1\n"


def prepare_scratch() -> dict[str, object]:
    """Create the requested filesystem root, or reject a foreign one."""
    marker = c.SCRATCH / ".litchi-performance-0819-owned"
    if c.SCRATCH.exists():
        assert c.SCRATCH.is_dir() and not c.SCRATCH.is_symlink()
        assert marker.is_file() and marker.read_text() == SCRATCH_MARKER, (
            f"refusing pre-existing unowned scratch root {c.SCRATCH}"
        )
    else:
        c.SCRATCH.mkdir(parents=True)
        marker.write_text(SCRATCH_MARKER)
    return c.artifact(marker)


def packet_descriptor(value: object, expected_path: Path) -> dict[str, object]:
    """Check a packet-relative or absolute descriptor emitted by an admission wrapper."""
    assert isinstance(value, dict)
    path_value = value.get("path")
    assert isinstance(path_value, str)
    path = Path(path_value)
    if not path.is_absolute():
        path = P / path
    assert path.resolve() == expected_path.resolve(), f"unexpected admission witness {path}"
    actual = c.artifact(path)
    assert value.get("sha256") == actual["sha256"]
    if "bytes" in value:
        assert value["bytes"] == actual["bytes"]
    return actual


def descriptor_path(value: object) -> Path:
    assert isinstance(value, dict) and isinstance(value.get("path"), str)
    path = Path(value["path"])
    if not path.is_absolute():
        path = P / path
    assert path.resolve().is_relative_to(P.resolve()), f"admission witness escapes packet: {path}"
    return path


def require_artifact_admission() -> dict[str, object]:
    path = P / "artifact-admission.json"
    assert path.is_file(), "artifact admission is required before qualification/native capture"
    admission = c.read(path)
    assert admission.get("schema") == "litchi.performance.0819.artifact-admission.v1"
    assert admission.get("accepted") is True
    assert admission.get("plan_sha256") == c.sha(P / "plan.json")
    packet_descriptor(admission["artifact_complete"], P / "artifacts.complete.json")
    packet_descriptor(admission["manifest"], P / "artifacts/manifest.json")
    audit_path = descriptor_path(admission["audit"])
    audit_descriptor = packet_descriptor(admission["audit"], audit_path)
    audit = c.read(audit_path)
    assert audit.get("schema") == "litchi.performance.0819.artifact-audit.v1"
    assert audit.get("ok") is True and not audit.get("errors")
    assert len(audit.get("cases", [])) == PLAN["expected"]["artifact_cases"]
    assert all(row.get("ok") is True for row in audit["cases"])
    packet_descriptor(admission["auditor"], P / "artifact_audit.py")
    preservation_path = P / "zip-preservation.json"
    packet_descriptor(admission["zip_preservation"], preservation_path)
    preservation = c.read(preservation_path)
    assert preservation.get("schema") == "litchi.performance.0819.zip-preservation.v1"
    assert isinstance(preservation.get("cases"), list)
    assert len(preservation["cases"]) == expected_counts["artifact_cases"]
    selectors = admission.get("selectors")
    assert isinstance(selectors, list) and len(selectors) == len(cases)
    expected = {
        (case["case"], case["format"], case["phase"], case["input"]): case
        for case in cases
    }
    observed = set()
    for selector in selectors:
        key = (selector.get("case"), selector.get("format"), selector.get("phase"), selector.get("input"))
        assert key in expected and key not in observed
        assert selector.get("source_sha256") == corpus[selector["input"]]["sha256"]
        assert isinstance(selector.get("published_sha256"), str) and selector["published_sha256"]
        assert selector.get("edit_outcome") == "admitted"
        observed.add(key)
    assert observed == set(expected)
    return admission


def require_qualification_admission(artifact_admission: dict[str, object]) -> dict[str, object]:
    path = P / "qualification-admission.json"
    assert path.is_file(), "qualification admission is required before native capture"
    admission = c.read(path)
    assert admission.get("schema") == "litchi.performance.0819.qualification-admission.v1"
    assert admission.get("accepted") is True
    assert admission.get("plan_sha256") == c.sha(P / "plan.json")
    assert admission.get("reports") == PLAN["expected"]["qualification_reports"]
    assert admission.get("samples") == PLAN["expected"]["qualification_samples"]
    packet_descriptor(admission["artifact_admission"], P / "artifact-admission.json")
    assert admission["artifact_admission"]["sha256"] == c.sha(P / "artifact-admission.json")
    packet_descriptor(admission["qualification_complete"], P / "qualification/complete.json")
    assert admission["qualification_complete"]["sha256"] == c.sha(P / "qualification/complete.json")
    assert artifact_admission.get("accepted") is True
    return admission


def require_complete(name: str, reports: int, samples: int) -> dict[str, object]:
    path = P / name / "complete.json"
    assert path.is_file(), f"{name} capture must complete before the next lane"
    complete = c.read(path)
    assert complete["reports"] == reports
    assert complete["samples"] == samples
    return complete


def validate_report(
    report_path: Path,
    case: dict[str, object],
    lane: str,
    samples: int,
    warmup: int,
    binary_name: str,
    admitted: dict[str, object] | None = None,
) -> None:
    report = c.read(report_path)
    config = report.get("configuration")
    assert isinstance(config, dict)
    assert config.get("samples_per_case") == samples
    assert config.get("warmup_iterations_per_case") == warmup
    results = report.get("results")
    assert isinstance(results, list) and len(results) == 1
    result = results[0]
    assert result.get("case") == case["case"]
    source = result.get("source")
    assert isinstance(source, dict)
    ordinary = source.get("ordinary_save")
    assert isinstance(ordinary, dict)
    assert ordinary.get("format") == case["format"].upper()
    assert ordinary.get("origin") == "caller-named-real-file"
    phase_name = {
        "lifecycle": "open+edit+save",
        "edit": "edit",
        "atomic_publish": "save-to-path",
        "counting_publish": "serialize-to-counting-sink",
    }[case["phase"]]
    assert ordinary.get("phase") == phase_name
    summary = ordinary.get("corpus")
    assert isinstance(summary, dict)
    real_file = summary.get("real_file")
    assert isinstance(real_file, dict)
    assert real_file.get("sha256") == corpus[case["input"]]["sha256"]
    assert real_file.get("bytes") == corpus[case["input"]]["bytes"]
    if admitted is not None:
        assert summary.get("source_archive_sha256") == admitted["source_sha256"]
        assert summary.get("published_sha256") == admitted["published_sha256"]
        assert summary.get("edit_outcome") == admitted["edit_outcome"] == "admitted"
    instrumented = binary_name == "observer"
    if instrumented:
        assert report.get("tool", {}).get("binary") == "litchi-perf-baseline-alloc"
        assert ordinary.get("process_probe", {}).get("fixed_count") == 32
        allocation = result.get("operation_metrics", {}).get("allocation")
        assert isinstance(allocation, dict) and allocation.get("status") == "measured"
    else:
        assert report.get("tool", {}).get("binary") == "litchi-perf-baseline"
        assert ordinary.get("process_probe") is None
        allocation = result.get("operation_metrics", {}).get("allocation")
        assert allocation is None or allocation.get("status") == "unavailable"


def run_artifact_export() -> None:
    output = P / "artifacts"
    log = P / "artifacts.log"
    rss = P / "artifacts.rss"
    receipt_path = P / "artifacts-receipt.json"
    source_path = P / "artifacts-source.json"
    assert not output.exists(), f"refusing to overwrite {output}"
    assert not log.exists() and not rss.exists() and not receipt_path.exists()
    assert not source_path.exists()
    scratch_marker = prepare_scratch()
    c.write(source_path, BUILD_SOURCE)
    binary = BUILD["binaries"]["artifacts"]["artifact"]
    command = [
        "/usr/bin/time", "-f", "%M", "-o", str(rss),
        "taskset", "-c", ",".join(str(cpu) for cpu in PLAN["affinity"]),
        binary["path"], "--output", str(output),
        "--filesystem-root", str(c.SCRATCH),
    ]
    for item in INPUTS:
        command.extend(["--ooxml-file", item["absolute"]])
    started = time.time()
    with log.open("w") as stream:
        result = subprocess.run(command, cwd=ROOT, stdout=stream, stderr=subprocess.STDOUT)
    row = {
        "schema": "litchi.performance.0819.artifact-receipt.v1",
        "lane": "artifacts",
        "command": command,
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "binary": BUILD["binaries"]["artifacts"],
        "source": c.artifact(source_path),
        "scratch_marker": scratch_marker,
        "inputs": INPUTS,
        "root_inputs": root_inputs,
        "locks": locks,
        "architecture": architecture,
        "corpus": corpus,
        "host": host,
        "packet": packet,
        "drivers": drivers,
        "unrelated": unrelated,
        "log": c.artifact(log),
    }
    if rss.exists():
        row["rss"] = c.artifact(rss)
    if output.exists():
        manifest_path = output / "manifest.json"
        if manifest_path.exists():
            row["manifest"] = c.artifact(manifest_path)
    c.write(receipt_path, row)
    if result.returncode != 0:
        raise RuntimeError(f"artifact export failed; retained {log}")
    assert output.is_dir() and (output / "manifest.json").is_file()
    assert rss.is_file()
    assert_rss(rss)
    manifest = c.read(output / "manifest.json")
    assert manifest.get("schema_version") == 1
    assert manifest.get("kind") == "ordinary-save-artifact-export"
    exported = manifest.get("cases")
    assert isinstance(exported, list) and len(exported) == PLAN["expected"]["artifact_cases"]
    assert sum(case.get("origin") == "generated-harness-corpus" for case in exported) == 3
    assert sum(case.get("origin") == "caller-named-real-file" for case in exported) == 3
    real_hashes = set()
    for exported_case in exported:
        assert len(exported_case.get("policy_outputs", [])) + 1 == PLAN["expected"]["artifact_policy_outputs_per_case"]
        descriptors = [exported_case["source_archive"]]
        descriptors.extend(item["output"] for item in exported_case["policy_outputs"])
        descriptors.append(exported_case["stream_output"]["output"])
        for descriptor in descriptors:
            path = output / descriptor["path"]
            assert path.is_file(), f"missing exported artifact {path}"
            assert c.artifact(path)["sha256"] == descriptor["sha256"]
        if exported_case["origin"] == "caller-named-real-file":
            real_hashes.add(exported_case["source_archive_sha256"])
    assert real_hashes == {item["sha256"] for item in INPUTS}
    stable()
    complete = {
        "schema": "litchi.performance.0819.artifacts.complete.v1",
        "children": 1,
        "reports": 0,
        "samples": 0,
        "cases": len(exported),
        "plan_sha256": c.sha(P / "plan.json"),
        "build_sha256": c.sha(P / "build.json"),
        "quality_sha256": c.sha(P / "quality.json"),
        "source": c.artifact(source_path),
        "receipt": c.artifact(receipt_path),
        "manifest": c.artifact(output / "manifest.json"),
    }
    c.write(P / "artifacts.complete.json", complete)
    print("0819 artifacts export complete", flush=True)


def run_capture(lane: str) -> None:
    lane_plan = PLAN["lanes"][lane]
    expected = {
        "blocks": lane_plan["blocks"],
        "samples": lane_plan["samples"],
        "warmup": lane_plan["warmup"],
        "binary": lane_plan["binary"],
    }
    assert len(lane_plan["orders"]) == expected["blocks"]
    assert all(order in ("forward", "reverse") for order in lane_plan["orders"])
    assert lane == "qualification" or lane in ("native", "observer")
    scratch_marker = prepare_scratch()
    artifact_admission = None
    if lane == "qualification":
        artifact_admission = require_artifact_admission()
    elif lane == "native":
        artifact_admission = require_artifact_admission()
        require_qualification_admission(artifact_admission)
    else:
        artifact_admission = require_artifact_admission()
        require_qualification_admission(artifact_admission)
        require_complete(
            "native", expected_counts["native_reports"], expected_counts["native_samples"]
        )
    admitted_by_case = {
        selector["case"]: selector for selector in artifact_admission["selectors"]
    }
    output = P / lane
    assert not output.exists(), f"refusing to overwrite {output}"
    output.mkdir()
    receipts_path = output / "receipts.json"
    binary_name = expected["binary"]
    binary = BUILD["binaries"][binary_name]["artifact"]
    assert c.artifact(binary["path"]) == binary
    source_path = output / "source.json"
    c.write(source_path, BUILD_SOURCE)
    rows = []
    for block, order in enumerate(lane_plan["orders"]):
        ordered = cases if order == "forward" else list(reversed(cases))
        for case in ordered:
            stem = f"{block:02d}-{case['format']}-{case['phase']}"
            report = output / f"{stem}.json"
            rss = output / f"{stem}.rss"
            log = output / f"{stem}.log"
            assert not report.exists() and not rss.exists() and not log.exists()
            command = [
                "/usr/bin/time", "-f", "%M", "-o", str(rss),
                "taskset", "-c", ",".join(str(cpu) for cpu in PLAN["affinity"]),
                binary["path"], "--warmup", str(expected["warmup"]),
                "--samples", str(expected["samples"]), "--case", case["case"],
                "--json", str(report), "--filesystem-root", str(c.SCRATCH),
                "--ooxml-file", str(ROOT / case["input"]),
            ]
            started = time.time()
            with log.open("w") as stream:
                result = subprocess.run(command, cwd=ROOT, stdout=stream, stderr=subprocess.STDOUT)
            row = {
                "schema": "litchi.performance.0819.capture-receipt.v1",
                "lane": lane,
                "block": block,
                "order": order,
                **case,
                "command": command,
                "started": started,
                "ended": time.time(),
                "exit_code": result.returncode,
                "binary": binary,
                "source": c.artifact(source_path),
                "scratch_marker": scratch_marker,
                "root_inputs": root_inputs,
                "locks": locks,
                "architecture": architecture,
                "corpus": corpus,
                "host": host,
                "packet": packet,
                "drivers": drivers,
                "unrelated": unrelated,
                "log": c.artifact(log),
            }
            if report.exists():
                row["report"] = c.artifact(report)
            if rss.exists():
                row["rss"] = c.artifact(rss)
            rows.append(row)
            c.write(receipts_path, rows)
            if result.returncode != 0:
                raise RuntimeError(f"{lane} capture failed; retained {log}")
            assert report.is_file() and rss.is_file()
            assert_rss(rss)
            validate_report(
                report, case, lane, expected["samples"], expected["warmup"], binary_name,
                admitted_by_case[case["case"]],
            )
            stable()
            print(stem, "PASS", flush=True)
    expected_reports = {
        "qualification": expected_counts["qualification_reports"],
        "native": expected_counts["native_reports"],
        "observer": expected_counts["observer_reports"],
    }[lane]
    expected_samples = {
        "qualification": expected_counts["qualification_samples"],
        "native": expected_counts["native_samples"],
        "observer": expected_counts["observer_samples"],
    }[lane]
    assert len(rows) == expected_reports
    assert len(rows) * expected["samples"] == expected_samples
    complete = {
        "schema": f"litchi.performance.0819.{lane}.complete.v1",
        "blocks": expected["blocks"],
        "children": len(rows),
        "reports": len(rows),
        "samples": len(rows) * expected["samples"],
        "plan_sha256": c.sha(P / "plan.json"),
        "build_sha256": c.sha(P / "build.json"),
        "quality_sha256": c.sha(P / "quality.json"),
        "source": c.artifact(source_path),
        "receipts": c.artifact(receipts_path),
    }
    c.write(output / "complete.json", complete)
    print(f"0819 {lane} capture complete", flush=True)


if LANE == "artifacts":
    run_artifact_export()
else:
    run_capture(LANE)
