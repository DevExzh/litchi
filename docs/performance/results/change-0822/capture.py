"""Root-owned serial qualification, native, and perf capture for 0822."""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import custody as c


LANE = sys.argv[1] if len(sys.argv) == 2 else ""
assert LANE in {"qualification", "native", "perf"}, (
    "choose exactly one lane: qualification, native, or perf"
)
P = c.P
PLAN = c.read(P / "plan.json")
QUALITY = c.read(P / "quality.json")
BUILD = c.read(P / "build.json")
assert QUALITY["status"] == "pass"
assert BUILD["schema"] == "litchi.performance.0822.build.v1"
FROZEN = c.read(BUILD["frozen_inputs"]["path"])
c.check_no_overrides()
c.stable(FROZEN)
assert c.probe_files() == FROZEN["probe"]
assert (P / "symbols/symbol.json").is_file(), (
    "run decode.py --symbols after build and before any qualification timing"
)

INPUT_ARG = PLAN["probe"]["input_path"]
REFERENCE_ARG = PLAN["probe"]["reference_path"]
assert INPUT_ARG == "test-data/ooxml/pptx/shapes.pptx"
assert REFERENCE_ARG == "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx"


def binary(name: str) -> dict[str, object]:
    value = BUILD["binaries"].get(name)
    assert isinstance(value, dict)
    descriptor = value["artifact"]
    assert c.artifact(descriptor["path"]) == descriptor
    return descriptor


def stable() -> None:
    c.stable(FROZEN)
    for name in ("ordinary", "fp"):
        descriptor = binary(name)
        assert c.artifact(descriptor["path"]) == descriptor


def run_probe(
    *,
    out: Path,
    label: str,
    arm: str,
    block: int,
    samples: int,
    warmup: int,
    lane: str,
) -> dict[str, object]:
    spec = PLAN["arms"][arm]
    binary_descriptor = binary(spec["binary"])
    report = out / f"{label}.json"
    rss = out / f"{label}.rss"
    log = out / f"{label}.log"
    assert not report.exists() and not rss.exists() and not log.exists()
    command = [
        "/usr/bin/time", "-f", "%M", "-o", str(rss),
        "taskset", "-c", str(PLAN["cpu"]), binary_descriptor["path"],
        "--mode", spec["mode"],
        "--input", INPUT_ARG,
        "--reference", REFERENCE_ARG,
        "--samples", str(samples),
        "--warmup", str(warmup),
        "--output", str(report),
    ]
    started = time.time()
    with log.open("w") as stream:
        result = subprocess.run(
            command,
            cwd=c.ROOT,
            env=os.environ | {"PYTHONDONTWRITEBYTECODE": "1"},
            stdout=stream,
            stderr=subprocess.STDOUT,
        )
    row: dict[str, object] = {
        "schema": "litchi.performance.0822.capture-receipt.v1",
        "lane": lane,
        "label": label,
        "block": block,
        "arm": arm,
        "mode": spec["mode"],
        "samples": samples,
        "warmup": warmup,
        "cpu": PLAN["cpu"],
        "command": command,
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "binary": binary_descriptor,
        "input": {
            "path": str(c.ROOT / INPUT_ARG),
            "bytes": 68822,
            "sha256": "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571",
        },
        "reference": {
            "path": str(c.ROOT / REFERENCE_ARG),
            "bytes": 68284,
            "sha256": "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf",
        },
        "source": c.artifact(P / "build-0/source.json"),
        "probe": c.artifact(P / "build-0/probe.json"),
        "frozen_inputs": c.artifact(BUILD["frozen_inputs"]["path"]),
        "log": c.artifact(log),
    }
    if report.exists():
        row["report"] = c.artifact(report)
    if rss.exists():
        row["rss"] = c.artifact(rss)
    return row


def validate_capture(row: dict[str, object]) -> None:
    assert row["exit_code"] == 0
    report = Path(row["report"]["path"])
    value = c.check_probe_report(
        report,
        mode=row["mode"],
        samples=row["samples"],
        warmup=row["warmup"],
    )
    # Keep the receipt boundary fail-closed even if the custody helper is
    # later relaxed for a reader compatibility change.  A successful process
    # exit alone is not a valid timing sample.
    assert value.get("warmup_verified") is True
    assert value.get("all_verified") is True
    records = value.get("samples")
    assert isinstance(records, list) and len(records) == row["samples"]
    elapsed = value.get("elapsed_ns")
    assert isinstance(elapsed, dict)
    elapsed_values = elapsed.get("samples")
    assert isinstance(elapsed_values, list)
    assert elapsed.get("sample_order") == list(range(row["samples"]))
    verification_fields = {
        "all_verified", "input_hash_verified", "reference_hash_verified",
        "output_hash_verified", "output_size_verified", "output_bytes_verified",
        "reopened", "marker_verified", "target_verified",
        "full_text_digest_verified", "slide_count_verified",
    }
    for index, record in enumerate(records):
        assert record.get("index") == index
        assert isinstance(elapsed_values[index], int) and elapsed_values[index] > 0
        assert record.get("elapsed_ns") == elapsed_values[index]
        checks = record.get("verification")
        assert isinstance(checks, dict)
        assert set(checks) == verification_fields and all(checks.values())
    output = value.get("output")
    assert isinstance(output, dict)
    assert output.get("bytes") == 68284
    assert output.get("sha256") == "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf"
    assert value.get("input_sha256", value.get("input_digest", row["input"]["sha256"])) == row["input"]["sha256"]
    assert value.get("reference_sha256", value.get("reference_digest", row["reference"]["sha256"])) == row["reference"]["sha256"]
    rss = Path(row["rss"]["path"]).read_text().strip()
    assert rss.isdigit() and int(rss) > 0


def native_lane(name: str, lane: dict[str, object], *, require: Path | None = None) -> None:
    out = P / name
    assert not out.exists(), f"refusing to overwrite {out}"
    if require is not None:
        assert require.is_file(), f"required prior lane missing: {require}"
    out.mkdir()
    rows: list[dict[str, object]] = []
    for block, order in enumerate(lane["orders"]):
        for arm in order:
            label = f"{block:02d}-{arm}"
            row = run_probe(
                out=out, label=label, arm=arm, block=block,
                samples=lane["samples"], warmup=lane["warmup"], lane=name,
            )
            rows.append(row)
            c.write(out / "receipts.json", rows)
            validate_capture(row)
            stable()
            print(name, label, "PASS", flush=True)
    complete = {
        "schema": f"litchi.performance.0822.{name}.complete.v1",
        "status": "pass",
        "blocks": lane["blocks"],
        "reports": len(rows),
        "samples": sum(row["samples"] for row in rows),
        "expected_reports": lane["reports"],
        "expected_samples": lane["samples_total"],
        "plan_sha256": c.sha(P / "plan.json"),
        "build_sha256": c.sha(P / "build.json"),
        "source": c.artifact(P / "build-0/source.json"),
        "probe": c.artifact(P / "build-0/probe.json"),
        "receipts": c.artifact(out / "receipts.json"),
    }
    assert complete["reports"] == complete["expected_reports"]
    assert complete["samples"] == complete["expected_samples"]
    c.write(out / "complete.json", complete)
    stable()


def perf_preflight() -> dict[str, object]:
    perf = shutil.which("perf")
    paranoid = Path("/proc/sys/kernel/perf_event_paranoid")
    result: dict[str, object] = {
        "schema": "litchi.performance.0822.perf-preflight.v1",
        "perf_path": perf,
        "perf_event_paranoid_path": str(paranoid),
        "perf_event_paranoid_readable": paranoid.is_file() and os.access(paranoid, os.R_OK),
        "started": time.time(),
    }
    if perf is None:
        result.update({"status": "unavailable", "reason": "perf executable not found"})
    elif not result["perf_event_paranoid_readable"]:
        result.update({"status": "unavailable", "reason": "perf_event_paranoid is unreadable"})
    else:
        value = paranoid.read_text().strip()
        result["perf_event_paranoid"] = value
        try:
            version = subprocess.run([perf, "--version"], capture_output=True, text=True)
            result["version_exit_code"] = version.returncode
            result["version"] = (version.stdout or version.stderr).strip()
            result["status"] = "ready" if version.returncode == 0 else "unavailable"
            if version.returncode != 0:
                result["reason"] = "perf --version failed"
        except OSError as error:
            result.update({"status": "unavailable", "reason": f"perf preflight error: {error}"})
    result["ended"] = time.time()
    return result


def permission_denial(text: str) -> bool:
    lowered = text.lower()
    return any(token in lowered for token in (
        "permission denied", "operation not permitted", "no permission",
        "perf_event_paranoid", "access to performance monitoring",
    ))


def perf_lane() -> None:
    out = P / "perf"
    assert not out.exists(), f"refusing to overwrite {out}"
    assert (P / "native/complete.json").is_file(), "perf follows native completion"
    out.mkdir()
    preflight = perf_preflight()
    c.write(out / "preflight.json", preflight)
    if preflight["status"] != "ready":
        c.write(out / "complete.json", {
            "schema": "litchi.performance.0822.perf.complete.v1",
            "status": "unavailable",
            "reason": preflight.get("reason", "perf preflight unavailable"),
            "typed_unavailable": True,
            "profile_fabricated": False,
            "expected_repeats": PLAN["perf"]["repeats"],
            "attempted_repeats": 0,
            "preflight": c.artifact(out / "preflight.json"),
            "plan_sha256": c.sha(P / "plan.json"),
            "build_sha256": c.sha(P / "build.json"),
        })
        stable()
        print("0822 perf unavailable (typed; no profile fabricated)", flush=True)
        return

    rows: list[dict[str, object]] = []
    binary_descriptor = binary(PLAN["perf"]["binary"])
    for repeat in range(PLAN["perf"]["repeats"]):
        report = out / f"{repeat}.json"
        raw = out / f"{repeat}.data"
        log = out / f"{repeat}.log"
        assert not report.exists() and not raw.exists() and not log.exists()
        command = [
            "taskset", "-c", str(PLAN["perf"]["cpu"]), "perf", "record",
            "--no-buildid-cache", "-e", PLAN["perf"]["event"],
            "-F", str(PLAN["perf"]["frequency_hz"]),
            "--call-graph", PLAN["perf"]["call_graph"], "-o", str(raw), "--",
            binary_descriptor["path"], "--mode", PLAN["perf"]["mode"],
            "--input", INPUT_ARG, "--reference", REFERENCE_ARG,
            "--samples", str(PLAN["perf"]["samples"]),
            "--warmup", str(PLAN["perf"]["warmup"]), "--output", str(report),
        ]
        started = time.time()
        with log.open("w") as stream:
            result = subprocess.run(
                command, cwd=c.ROOT, env=os.environ | {"PYTHONDONTWRITEBYTECODE": "1"},
                stdout=stream, stderr=subprocess.STDOUT,
            )
        text = log.read_text(errors="replace")
        row: dict[str, object] = {
            "schema": "litchi.performance.0822.perf-receipt.v1",
            "repeat": repeat,
            "command": command,
            "started": started,
            "ended": time.time(),
            "exit_code": result.returncode,
            "binary": binary_descriptor,
            "input": {
                "path": str(c.ROOT / INPUT_ARG),
                "bytes": 68822,
                "sha256": "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571",
            },
            "reference": {
                "path": str(c.ROOT / REFERENCE_ARG),
                "bytes": 68284,
                "sha256": "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf",
            },
            "samples": PLAN["perf"]["samples"],
            "warmup": PLAN["perf"]["warmup"],
            "cpu": PLAN["perf"]["cpu"],
            "log": c.artifact(log),
            "source": c.artifact(P / "build-0/source.json"),
            "probe": c.artifact(P / "build-0/probe.json"),
        }
        if report.exists():
            row["report"] = c.artifact(report)
        if raw.exists():
            row["raw"] = c.artifact(raw)
        rows.append(row)
        c.write(out / "receipts.json", rows)
        if result.returncode != 0 and permission_denial(text):
            c.write(out / "complete.json", {
                "schema": "litchi.performance.0822.perf.complete.v1",
                "status": "unavailable",
                "typed_unavailable": True,
                "profile_fabricated": False,
                "reason": "perf record permission denial",
                "preflight": c.artifact(out / "preflight.json"),
                "receipts": c.artifact(out / "receipts.json"),
                "expected_repeats": PLAN["perf"]["repeats"],
                "attempted_repeats": len(rows),
                "plan_sha256": c.sha(P / "plan.json"),
                "build_sha256": c.sha(P / "build.json"),
            })
            stable()
            print("0822 perf denied (typed; retained receipt, no profile fabricated)", flush=True)
            return
        assert result.returncode == 0, log
        assert report.is_file() and raw.is_file() and raw.stat().st_size > 0
        c.check_probe_report(report, mode=PLAN["perf"]["mode"], samples=PLAN["perf"]["samples"], warmup=0)
        stable()
        print("perf", repeat, "PASS", flush=True)
    c.write(out / "complete.json", {
        "schema": "litchi.performance.0822.perf.complete.v1",
        "status": "available",
        "typed_unavailable": False,
        "profile_fabricated": False,
        "reports": len(rows),
        "samples": sum(PLAN["perf"]["samples"] for _ in rows),
        "expected_reports": PLAN["perf"]["reports"],
        "expected_samples": PLAN["perf"]["samples_total"],
        "event": PLAN["perf"]["event"],
        "frequency_hz": PLAN["perf"]["frequency_hz"],
        "call_graph": PLAN["perf"]["call_graph"],
        "owner": PLAN["perf"]["owner"],
        "preflight": c.artifact(out / "preflight.json"),
        "receipts": c.artifact(out / "receipts.json"),
        "plan_sha256": c.sha(P / "plan.json"),
        "build_sha256": c.sha(P / "build.json"),
    })
    stable()


if LANE == "qualification":
    native_lane("qualification", PLAN["qualification"])
elif LANE == "native":
    native_lane("native", PLAN["native"], require=P / "qualification/complete.json")
else:
    perf_lane()
