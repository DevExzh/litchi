"""Root-owned serial qualification and paired native/allocation captures."""

from __future__ import annotations

import os
import subprocess
import sys
import time
from pathlib import Path

import custody as c


assert len(sys.argv) >= 2
LANE = sys.argv[1]
assert LANE in {"qualification", "native", "allocation"}
QUAL_LEG = sys.argv[2] if LANE == "qualification" and len(sys.argv) == 3 else None
assert (LANE == "qualification" and QUAL_LEG in {"before", "after"}) or (
    LANE in {"native", "allocation"} and len(sys.argv) == 2
), "use qualification before|after, native, or allocation"

PLAN = c.read(c.P / "plan.json")
FROZEN = c.read(c.P / "freeze.json")
c.assert_static()
c.stable_inputs(FROZEN)
c.check_no_overrides()

if LANE == "qualification":
    assert (c.P / f"build-{QUAL_LEG}/build.json").is_file()
    OUT = c.P / f"qualification-{QUAL_LEG}"
    legs = [QUAL_LEG]
    binary_variant = "allocation"
    lane_plan = PLAN["qualification"]
else:
    assert (c.P / "build-before/build.json").is_file()
    assert (c.P / "build-after/build.json").is_file()
    OUT = c.P / LANE
    legs = ["before", "after"]
    binary_variant = LANE
    lane_plan = PLAN[LANE]
assert not OUT.exists(), f"refusing to overwrite {OUT}"
OUT.mkdir()

builds = {leg: c.read(c.P / f"build-{leg}/build.json") for leg in legs}
expected_sources = {
    leg: c.read(build["source"]["path"])
    for leg, build in builds.items()
}
current_expected = expected_sources[QUAL_LEG] if QUAL_LEG else expected_sources["after"]
assert c.source() == current_expected, "current production source does not match capture leg"
if QUAL_LEG == "before":
    assert expected_sources["before"] == FROZEN["source"]
if QUAL_LEG == "after":
    assert c.changed_files(FROZEN["source"], expected_sources["after"]) == set(c.ALLOWLIST)
    assert c.sha(c.ROOT / c.ALLOWLIST[0]) == FROZEN["candidate"]["candidate/after/reader.rs"]
if LANE in {"native", "allocation"}:
    for leg in ("before", "after"):
        qualification = c.read(c.P / f"qualification-{leg}/complete.json")
        assert qualification["status"] == "pass"
        assert qualification["reports"] == 19 and qualification["samples"] == 19
        assert qualification["expected_reports"] == 19 and qualification["expected_samples"] == 19
for leg, build in builds.items():
    assert build["schema"] == f"litchi.performance.0823.build-{leg}.v1"
    assert c.probe_files("synthetic") == FROZEN["probes"]["synthetic"]
    assert c.probe_files("real") == FROZEN["probes"]["real"]
    for descriptor in build["binaries"].values():
        artifact = descriptor["artifact"]
        assert c.artifact(artifact["path"]) == artifact


def binary(leg: str, probe: str) -> dict[str, object]:
    build = builds[leg]
    key = f"{probe}-{binary_variant}"
    value = build["binaries"][key]["artifact"]
    assert c.artifact(value["path"]) == value
    return value


def run_child(*, leg: str, block: int, case: dict[str, object], samples: int,
              warmup: int) -> dict[str, object]:
    probe = case["probe"]
    label_bits = [str(block), leg, probe]
    args: list[str] = []
    if probe == "synthetic":
        shape = case["shape"]
        mode = case["mode"]
        label_bits.extend([str(shape), str(mode)])
        args = ["--mode", str(mode), "--shape", str(shape)]
    else:
        mode = case["mode"]
        label_bits.append(str(mode))
        args = [
            "--mode", str(mode),
            "--input", c.REAL_INPUT["path"],
            "--reference", c.REAL_REFERENCE["path"],
        ]
    label = "-".join(label_bits)
    report = OUT / f"{label}.json"
    rss = OUT / f"{label}.rss"
    log = OUT / f"{label}.log"
    assert not report.exists() and not rss.exists() and not log.exists()
    descriptor = binary(leg, probe)
    command = [
        "/usr/bin/time", "-f", "%M", "-o", str(rss),
        "taskset", "-c", str(PLAN["cpu"]), descriptor["path"],
        *args, "--samples", str(samples), "--warmup", str(warmup),
        "--output", str(report),
    ]
    started = time.time()
    with log.open("w", encoding="utf-8") as stream:
        result = subprocess.run(
            command, cwd=c.ROOT, env=os.environ | {"PYTHONDONTWRITEBYTECODE": "1"},
            stdout=stream, stderr=subprocess.STDOUT,
        )
    row: dict[str, object] = {
        "schema": "litchi.performance.0823.capture-receipt.v1",
        "lane": LANE,
        "leg": leg,
        "block": block,
        "case": case,
        "samples": samples,
        "warmup": warmup,
        "cpu": PLAN["cpu"],
        "command": command,
        "started": started,
        "ended": time.time(),
        "exit_code": result.returncode,
        "binary": descriptor,
        "source": c.artifact(c.P / f"build-{leg}/source.json"),
        "frozen_inputs": c.artifact(c.P / "freeze.json"),
        "log": c.artifact(log),
    }
    if report.is_file():
        row["report"] = c.artifact(report)
    if rss.is_file():
        row["rss"] = c.artifact(rss)
    rows.append(row)
    c.write(OUT / "receipts.json", rows)
    if result.returncode != 0 or "report" not in row or "rss" not in row:
        c.write(OUT / "failure.json", {
            "schema": "litchi.performance.0823.capture-failure.v1",
            "failed": row,
            "reason": "child exit or required report/RSS receipt missing",
        })
        raise RuntimeError(f"0823 child failed; retained {log}")
    rss_text = Path(row["rss"]["path"]).read_text(encoding="utf-8").strip()
    assert rss_text.isdigit() and int(rss_text) > 0
    report_path = Path(row["report"]["path"])
    if probe == "synthetic":
        c.check_synthetic_report(
            report_path, shape=str(case["shape"]), mode=str(case["mode"]),
            samples=samples, warmup=warmup, allocation=binary_variant == "allocation",
            binary_name=Path(descriptor["path"]).name,
        )
    else:
        c.check_real_report(
            report_path, samples=samples, warmup=warmup,
            allocation=binary_variant == "allocation",
            binary_name=Path(descriptor["path"]).name,
        )
    c.stable_inputs(FROZEN)
    assert c.source() == current_expected, "production source changed during capture"
    print(LANE, label, "PASS", flush=True)
    return row


rows: list[dict[str, object]] = []
cases = PLAN["cases"]
if LANE == "qualification":
    orders = [[QUAL_LEG]]
else:
    orders = lane_plan["orders"]
for block, order in enumerate(orders):
    for case in cases:
        for leg in order:
            run_child(
                leg=leg, block=block, case=case,
                samples=lane_plan["samples"], warmup=lane_plan["warmup"],
            )

expected_reports = len(cases) * len(orders) * len(orders[0])
expected_samples = expected_reports * lane_plan["samples"]
complete = {
    "schema": f"litchi.performance.0823.{LANE}.complete.v1",
    "status": "pass",
    "lane": LANE,
    "leg": QUAL_LEG,
    "blocks": len(orders),
    "reports": len(rows),
    "samples": sum(row["samples"] for row in rows),
    "expected_reports": expected_reports,
    "expected_samples": expected_samples,
    "plan": c.artifact(c.P / "plan.json"),
    "freeze": c.artifact(c.P / "freeze.json"),
    "receipts": c.artifact(OUT / "receipts.json"),
}
assert complete["reports"] == expected_reports and complete["samples"] == expected_samples
c.write(OUT / "complete.json", complete)
print(f"0823 {LANE} PASS: {len(rows)} reports/{complete['samples']} samples", flush=True)
