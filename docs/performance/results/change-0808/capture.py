"""Root-only serial workflow captures with a frozen schedule."""

import subprocess
import sys
import time

import custody as c


LANE = sys.argv[1]
assert LANE in ("qualification", "native", "allocation")
P = c.P
plan = c.read(P / "plan.json")
OUT = P / LANE
assert not OUT.exists(), f"refusing to overwrite {OUT}"
OUT.mkdir()

if LANE == "qualification":
    legs = ["before"]
    lane_plan = plan["qualification"]
    binary_kind = "allocation"
    blocks = lane_plan["blocks"]
else:
    legs = ["before", "after"]
    lane_plan = plan[LANE]
    binary_kind = LANE
    blocks = lane_plan["blocks"]

builds = {leg: c.read(P / f"build-{leg}/build.json") for leg in legs}
if LANE == "qualification":
    expected_sources = {"before": c.read(builds["before"]["source"]["path"])}
else:
    expected_sources = {
        leg: c.read(builds[leg]["source"]["path"])
        for leg in ("before", "after")
    }
    assert c.changed_files(
        expected_sources["before"], expected_sources["after"]
    ) == set(plan["source_allowlist"])
    assert c.source() == expected_sources["after"]
for leg, build in builds.items():
    assert expected_sources[leg] == c.read(build["source"]["path"])
    for binary in build["binaries"].values():
        assert c.artifact(binary["path"]) == binary

frozen = c.source()
if LANE == "qualification":
    assert frozen == expected_sources["before"]
else:
    assert frozen == expected_sources["after"]
c.write(OUT / "source.json", frozen)
rows = []
for block in range(blocks):
    cases = plan["cases"]
    order = ["before"] if LANE == "qualification" else plan["native"]["orders"][block]
    for case in cases:
        for leg in order:
            stem = f"{block}-{case['shape']}-{case['mode']}-{leg}"
            report = OUT / f"{stem}.json"
            rss = OUT / f"{stem}.rss"
            log = OUT / f"{stem}.log"
            samples = lane_plan["samples"]
            warmup = lane_plan["warmup"]
            binary = builds[leg]["binaries"][binary_kind]
            command = [
                "/usr/bin/time",
                "-f",
                "%M",
                "-o",
                str(rss),
                "taskset",
                "-c",
                str(plan["cpu"]),
                binary["path"],
                "--mode",
                case["mode"],
                "--shape",
                case["shape"],
                "--samples",
                str(samples),
                "--warmup",
                str(warmup),
                "--output",
                str(report),
            ]
            started = time.time()
            with log.open("w") as stream:
                result = subprocess.run(
                    command,
                    cwd=c.ROOT,
                    stdout=stream,
                    stderr=subprocess.STDOUT,
                )
            row = {
                "schema": "litchi.performance.0808.capture-receipt.v1",
                "lane": LANE,
                "block": block,
                **case,
                "leg": leg,
                "command": command,
                "started": started,
                "ended": time.time(),
                "exit_code": result.returncode,
                "binary": binary,
                "log": c.artifact(log),
                "rss": c.artifact(rss),
            }
            if report.exists():
                row["report"] = c.artifact(report)
            rows.append(row)
            c.write(OUT / "receipts.json", rows)
            assert result.returncode == 0, log
            assert report.is_file() and c.read(report)["samples"]
            assert len(c.read(report)["samples"]) == samples
            assert c.source() == frozen
            print(stem, "PASS", flush=True)

c.write(
    OUT / "complete.json",
    {
        "schema": f"litchi.performance.0808.{LANE}.complete.v1",
        "children": len(rows),
        "reports": len(rows),
        "samples": sum(lane_plan["samples"] for _ in rows),
        "plan_sha256": c.sha(P / "plan.json"),
        "receipts": c.artifact(OUT / "receipts.json"),
        "source": c.artifact(OUT / "source.json"),
    },
)
