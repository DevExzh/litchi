"""Root-only serial captures for the frozen 0816 execution matrix."""

from __future__ import annotations

import subprocess
import sys
import time

import custody as c


LANE = sys.argv[1] if len(sys.argv) == 2 else ""
assert LANE in ("qualification", "native", "observer"), (
    "choose exactly one lane: qualification, native, or observer"
)
P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
BUILD = c.read(P / "build.json")
QUALITY = c.read(P / "quality.json")
assert PLAN["schema"] == "litchi.performance.0816.plan.v1"
assert ORIGIN["base"] == "c8ff2b9f65ecd011d12fd5b1ac352ea315785b93"
assert PLAN["counts"] == {"reports": 648, "samples": 13320}
assert BUILD["schema"] == "litchi.performance.0816.build.v1"
assert QUALITY["schema"] == "litchi.performance.0816.quality.v1"
assert QUALITY["status"] == "pass" and QUALITY["gate_count"] == 6
assert len(QUALITY["rows"]) == 6
assert all(row["exit_code"] == 0 for row in QUALITY["rows"])

OUT = P / LANE
assert not OUT.exists(), f"refusing to overwrite {OUT}"

BUILD_SOURCE = c.read(BUILD["source"]["path"])
assert isinstance(BUILD_SOURCE, dict)
c.unchanged(BUILD_SOURCE)
assert c.assert_root_inputs() == BUILD["root_inputs"]
assert c.assert_tool_lock() == BUILD["tool_lock"]
assert c.architecture_hashes() == BUILD["architecture"]
assert c.assert_host(BUILD["host"]) == BUILD["host"]
assert c.packet_hashes() == BUILD["packet"]
assert c.driver_hashes() == BUILD["drivers"]
assert c.assert_unrelated() == BUILD["unrelated"]
assert QUALITY["source"]
QUALITY_SOURCE = c.read(QUALITY["source"]["path"])
assert QUALITY_SOURCE == BUILD_SOURCE

assert PLAN["affinity"] == [12, 13, 14, 15, 16, 17, 18, 19]
CLI_KEYS = (
    "route",
    "shape",
    "state",
    "workers",
    "task_floor",
    "source_max_read_bytes",
    "source_delay_us",
)
lane_plan = PLAN[LANE]
expected = {
    "qualification": {"blocks": 1, "samples": 1, "warmup": 0, "binary": "observer"},
    "native": {"blocks": 6, "samples": 30, "warmup": 3, "binary": "native"},
    "observer": {"blocks": 2, "samples": 2, "warmup": 0, "binary": "observer"},
}[LANE]
for key in ("blocks", "samples", "warmup"):
    assert lane_plan[key] == expected[key]
assert lane_plan["orders"] == {
    "qualification": ["forward"],
    "native": ["forward", "reverse", "forward", "reverse", "reverse", "forward"],
    "observer": ["forward", "reverse"],
}[LANE]
assert len(PLAN["cases"]) == 72
assert len({tuple(case[key] for key in CLI_KEYS) for case in PLAN["cases"]}) == 72

binary = BUILD["binaries"][expected["binary"]]
assert c.artifact(binary["path"]) == binary
OUT.mkdir()
source_path = OUT / "source.json"
c.write(source_path, BUILD_SOURCE)


def check_stable() -> None:
    c.unchanged(BUILD_SOURCE)
    assert c.assert_root_inputs() == BUILD["root_inputs"]
    assert c.assert_tool_lock() == BUILD["tool_lock"]
    assert c.architecture_hashes() == BUILD["architecture"]
    assert c.assert_host(BUILD["host"]) == BUILD["host"]
    assert c.packet_hashes() == BUILD["packet"]
    assert c.driver_hashes() == BUILD["drivers"]
    assert c.assert_unrelated() == BUILD["unrelated"]
    assert c.artifact(binary["path"]) == binary


rows = []
for block, order in enumerate(lane_plan["orders"]):
    cases = PLAN["cases"] if order == "forward" else list(reversed(PLAN["cases"]))
    for case in cases:
        assert set(case) == set(CLI_KEYS), "case contains a non-CLI field"
        stem = (
            f"{block:02d}-{case['route']}-{case['shape']}-{case['state']}"
            f"-workers{case['workers']}-floor{case['task_floor']}"
            f"-maxread{case['source_max_read_bytes']}-delay{case['source_delay_us']}"
        )
        report = OUT / f"{stem}.json"
        rss = OUT / f"{stem}.rss"
        log = OUT / f"{stem}.log"
        assert not report.exists() and not rss.exists() and not log.exists()
        command = [
            "/usr/bin/time",
            "-f",
            "%M",
            "-o",
            str(rss),
            "taskset",
            "-c",
            ",".join(str(cpu) for cpu in PLAN["affinity"]),
            binary["path"],
        ]
        for key in CLI_KEYS:
            command.extend(["--" + key.replace("_", "-"), str(case[key])])
        command.extend(
            [
                "--samples",
                str(expected["samples"]),
                "--warmup",
                str(expected["warmup"]),
                "--output",
                str(report),
            ]
        )
        started = time.time()
        with log.open("w") as stream:
            result = subprocess.run(
                command,
                cwd=c.ROOT,
                stdout=stream,
                stderr=subprocess.STDOUT,
            )
        row = {
            "schema": "litchi.performance.0816.capture-receipt.v1",
            "lane": LANE,
            "block": block,
            **case,
            "command": command,
            "started": started,
            "ended": time.time(),
            "exit_code": result.returncode,
            "binary": binary,
            "source": c.artifact(source_path),
            "root_inputs": BUILD["root_inputs"],
            "tool_lock": BUILD["tool_lock"],
            "architecture": BUILD["architecture"],
            "host": BUILD["host"],
            "packet": BUILD["packet"],
            "drivers": BUILD["drivers"],
            "unrelated": BUILD["unrelated"],
            "log": c.artifact(log),
        }
        if report.exists():
            row["report"] = c.artifact(report)
        if rss.exists():
            row["rss"] = c.artifact(rss)
        rows.append(row)
        c.write(OUT / "receipts.json", rows)
        if result.returncode != 0:
            raise RuntimeError(f"{LANE} capture failed; retained {log}")
        assert report.is_file() and rss.is_file()
        report_value = c.read(report)
        expected_schema = (
            "litchi.execution-baseline.v1"
            if case["source_max_read_bytes"] == 0 and case["source_delay_us"] == 0
            else "litchi.execution-range-baseline.v1"
        )
        assert report_value.get("schema") == expected_schema
        assert len(report_value.get("samples", [])) == expected["samples"]
        config = report_value.get("config")
        assert isinstance(config, dict)
        assert config.get("samples") == expected["samples"]
        assert config.get("warmup") == expected["warmup"]
        check_stable()
        print(stem, "PASS", flush=True)

complete = {
    "schema": f"litchi.performance.0816.{LANE}.complete.v1",
    "blocks": expected["blocks"],
    "children": len(rows),
    "reports": len(rows),
    "samples": expected["samples"] * len(rows),
    "plan_sha256": c.sha(P / "plan.json"),
    "build_sha256": c.sha(P / "build.json"),
    "quality_sha256": c.sha(P / "quality.json"),
    "source": c.artifact(source_path),
    "receipts": c.artifact(OUT / "receipts.json"),
}
c.write(OUT / "complete.json", complete)
print(f"0816 {LANE} capture complete", flush=True)
