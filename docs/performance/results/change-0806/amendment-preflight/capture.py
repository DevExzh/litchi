"""Root-owned fresh native captures for the protected amendment preflight."""

from __future__ import annotations

import subprocess
import sys
import time
from pathlib import Path

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")


def main() -> None:
    if len(sys.argv) != 1:
        raise SystemExit("capture.py takes no arguments; it runs the frozen native lane")
    out = P / "native"
    if out.exists():
        raise AssertionError(f"capture output already exists: {out}")
    out.mkdir()
    builds = {}
    for leg in ("before", "after"):
        build = c.read(P / f"build-{leg}" / "build.json")
        if build.get("schema") != "litchi.performance.0806.amendment-build.v1":
            raise AssertionError(f"{leg}: build schema changed")
        binary = build.get("binary")
        if not isinstance(binary, dict):
            raise AssertionError(f"{leg}: binary descriptor missing")
        c.assert_artifact(binary, f"{leg} binary")
        builds[leg] = build
    frozen_archives = {leg: c.archive_manifest(leg) for leg in ("before", "after")}
    frozen_probe = c.probe_manifest()
    c.write(out / "source.json", {
        "schema": "litchi.performance.0806.amendment-capture-source.v1",
        "workspace": c.source(),
        "archives": frozen_archives,
        "probe": frozen_probe,
    })
    rows = []
    native = PLAN["native"]
    for block, order in enumerate(native["orders"]):
        for case in c.read(P / "cases.json"):
            case_id = case["id"]
            for mode in PLAN["modes"]:
                for leg in order:
                    stem = f"{block}-{case_id}-{mode}-{leg}"
                    report = out / f"{stem}.json"
                    log = out / f"{stem}.log"
                    rss = out / f"{stem}.rss"
                    binary = builds[leg]["binary"]
                    command = [
                        "/usr/bin/time", "-f", "%M", "-o", str(rss),
                        "taskset", "-c", str(PLAN["cpu"]), binary["path"],
                        "--leg", leg, "--case", case_id, "--mode", mode,
                        "--samples", str(native["samples"]),
                        "--warmup", str(native["warmup"]),
                        "--iterations", str(native["iterations"]),
                        "--output", str(report),
                    ]
                    started = time.time()
                    with log.open("w", encoding="utf-8") as stream:
                        result = subprocess.run(
                            command, cwd=c.ROOT, stdout=stream, stderr=subprocess.STDOUT
                        )
                    row = {
                        "schema": "litchi.performance.0806.amendment-native-receipt.v1",
                        "block": block,
                        "case": case_id,
                        "mode": mode,
                        "leg": leg,
                        "command": command,
                        "started": started,
                        "ended": time.time(),
                        "exit_code": result.returncode,
                        "binary": binary,
                        "log": c.relative_artifact(log),
                        "rss": c.relative_artifact(rss),
                    }
                    if report.exists():
                        row["report"] = c.relative_artifact(report)
                    rows.append(row)
                    c.write(out / "receipts.json", rows)
                    if result.returncode != 0:
                        raise SystemExit(f"native child failed: {log}")
                    if not report.is_file():
                        raise AssertionError(f"native child omitted report: {report}")
                    if c.archive_manifest(leg) != frozen_archives[leg] or c.probe_manifest() != frozen_probe:
                        raise AssertionError("source or probe archive changed during native capture")
                    if c.sha(Path(binary["path"])) != binary["sha256"]:
                        raise AssertionError(f"{leg}: native binary changed during capture")
                    data = c.read(report)
                    if data.get("schema") != "litchi.attribute-boundary-probe.v1":
                        raise AssertionError(f"{stem}: report schema changed")
                    if data.get("tool") != "attribute-boundary-probe-0805":
                        raise AssertionError(f"{stem}: report tool changed")
                    if len(data.get("samples", [])) != native["samples"]:
                        raise AssertionError(f"{stem}: sample count changed")
    c.write(out / "complete.json", {
        "schema": "litchi.performance.0806.amendment-native.v1",
        "children": len(rows),
        "expected_children": native["blocks"] * PLAN["case_count"] * len(PLAN["modes"]) * 2,
        "samples": native["blocks"] * PLAN["case_count"] * len(PLAN["modes"]) * 2 * native["samples"],
        "receipts": c.relative_artifact(out / "receipts.json"),
        "source": c.relative_artifact(out / "source.json"),
        "profiles": False,
        "callgrind": False,
        "historical_timing_pooling": False,
    })


if __name__ == "__main__":
    main()
