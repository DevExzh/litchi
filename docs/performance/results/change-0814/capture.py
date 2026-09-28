"""Root-only serial native and sampled-perf captures for current production."""

import subprocess
import sys
import time

import custody as c


LANE = sys.argv[1] if len(sys.argv) == 2 else ""
assert LANE in ("native", "perf"), "choose exactly one lane: native or perf"
P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
BUILD = c.read(P / "build/build.json")
FROZEN = c.read(BUILD["frozen_inputs"]["path"])
QUALITY = c.read(P / "quality-reuse.json")
PROBE_QUALITY = c.read(P / "probe-quality.json")
assert PLAN["schema"] == "litchi.performance.0814.current-production.v1"
assert ORIGIN["base"] == PLAN["source_revision"]
assert BUILD["schema"] == "litchi.performance.0814.build.v1"
assert c.frozen_driver_hashes() == FROZEN["drivers"]
assert QUALITY["schema"] == "litchi.performance.0814.quality-reuse.v1"
assert QUALITY["gate_count"] == 6
assert QUALITY["cargo_executed"] is False
assert PROBE_QUALITY["schema"] == "litchi.performance.0814.probe-quality.v1"
assert PROBE_QUALITY["gate_count"] == 3
assert PROBE_QUALITY["tests_passed"] == 36
assert c.TARGET.is_dir()

source = c.source()
assert source["revision"] == ORIGIN["base"]
assert len(source["files"]) == PLAN["source_file_count"]
assert c.read(BUILD["source"]["path"]) == source
probe = c.assert_probe(BUILD["probe"])
root_inputs = c.assert_root_inputs()
architecture = c.architecture_hashes()
unrelated = c.assert_unrelated()
for variant, binary in BUILD["binaries"].items():
    assert c.artifact(binary["path"]) == {
        key: binary[key] for key in ("path", "bytes", "sha256")
    }


def check_stable() -> None:
    assert c.source() == source
    assert c.assert_probe(probe) == probe
    assert c.assert_root_inputs() == root_inputs
    assert c.architecture_hashes() == architecture
    assert c.assert_unrelated() == unrelated


if LANE == "native":
    OUT = P / "native"
    assert not OUT.exists(), f"refusing to overwrite {OUT}"
    OUT.mkdir()
    lane = PLAN["native"]
    assert lane["reports"] == 54
    assert lane["samples_total"] == 1620
    rows = []
    for block, order in enumerate(lane["orders"]):
        assert len(order) == 3
        for shape in lane["shapes"]:
            for variant in order:
                binary = BUILD["binaries"][variant]
                stem = f"{block}-{shape}-{variant}"
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
                    str(PLAN["cpu"]),
                    binary["path"],
                    "--mode",
                    "capture",
                    "--shape",
                    shape,
                    "--samples",
                    str(lane["samples"]),
                    "--warmup",
                    str(lane["warmup"]),
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
                    "schema": "litchi.performance.0814.native-receipt.v1",
                    "lane": "native",
                    "block": block,
                    "shape": shape,
                    "mode": "capture",
                    "variant": variant,
                    "command": command,
                    "started": started,
                    "ended": time.time(),
                    "exit_code": result.returncode,
                    "binary": binary,
                    "source": c.artifact(P / "build/source.json"),
                    "probe": probe,
                    "root_inputs": root_inputs,
                    "architecture": architecture,
                    "unrelated": unrelated,
                    "log": c.artifact(log),
                }
                if report.exists():
                    row["report"] = c.artifact(report)
                if rss.exists():
                    row["rss"] = c.artifact(rss)
                rows.append(row)
                c.write(OUT / "receipts.json", rows)
                assert result.returncode == 0, log
                assert report.is_file() and rss.is_file()
                report_value = c.read(report)
                assert report_value["schema"] == "litchi.pptx.public-workflow-probe-0806.v1"
                assert len(report_value["samples"]) == lane["samples"]
                assert report_value["warmup"] == lane["warmup"]
                check_stable()
                print(stem, "PASS", flush=True)
    c.write(
        OUT / "complete.json",
        {
            "schema": "litchi.performance.0814.native.complete.v1",
            "blocks": lane["blocks"],
            "children": len(rows),
            "reports": len(rows),
            "samples": sum(lane["samples"] for _ in rows),
            "plan_sha256": c.sha(P / "plan.json"),
            "build_sha256": c.sha(P / "build/build.json"),
            "source": c.artifact(P / "build/source.json"),
            "receipts": c.artifact(OUT / "receipts.json"),
        },
    )
else:
    OUT = P / "perf"
    assert (P / "native/complete.json").is_file(), "perf must follow native completion"
    native_complete = c.read(P / "native/complete.json")
    assert native_complete["reports"] == PLAN["native"]["reports"]
    assert native_complete["samples"] == PLAN["native"]["samples_total"]
    assert not OUT.exists(), f"refusing to overwrite {OUT}"
    OUT.mkdir()
    lane = PLAN["perf"]
    assert lane["reports"] == 2
    assert lane["samples_total"] == 200
    binary = BUILD["binaries"][lane["binary"]]
    rows = []
    for repeat in range(lane["repeats"]):
        stem = str(repeat)
        report = OUT / f"{stem}.json"
        raw = OUT / f"{stem}.data"
        log = OUT / f"{stem}.log"
        assert not report.exists() and not raw.exists() and not log.exists()
        command = [
            "taskset",
            "-c",
            str(lane["cpu"]),
            "perf",
            "record",
            "--no-buildid-cache",
            "-e",
            lane["event"],
            "-F",
            str(lane["frequency_hz"]),
            "--call-graph",
            lane["call_graph"],
            "-o",
            str(raw),
            "--",
            binary["path"],
            "--mode",
            lane["mode"],
            "--shape",
            lane["shape"],
            "--samples",
            str(lane["samples"]),
            "--warmup",
            str(lane["warmup"]),
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
            "schema": "litchi.performance.0814.perf-receipt.v1",
            "lane": "perf",
            "repeat": repeat,
            "command": command,
            "started": started,
            "ended": time.time(),
            "exit_code": result.returncode,
            "binary": binary,
            "source": c.artifact(P / "build/source.json"),
            "probe": probe,
            "root_inputs": root_inputs,
            "architecture": architecture,
            "unrelated": unrelated,
            "log": c.artifact(log),
        }
        if report.exists():
            row["report"] = c.artifact(report)
        if raw.exists():
            row["raw"] = c.artifact(raw)
        rows.append(row)
        c.write(OUT / "receipts.json", rows)
        assert result.returncode == 0, log
        assert report.is_file() and raw.is_file() and raw.stat().st_size > 0
        report_value = c.read(report)
        assert report_value["schema"] == "litchi.pptx.public-workflow-probe-0806.v1"
        assert len(report_value["samples"]) == lane["samples"]
        assert report_value["warmup"] == lane["warmup"]
        assert c.artifact(binary["path"]) == {
            key: binary[key] for key in ("path", "bytes", "sha256")
        }
        check_stable()
        print("perf", repeat, "PASS", flush=True)
    c.write(
        OUT / "complete.json",
        {
            "schema": "litchi.performance.0814.perf.complete.v1",
            "repeats": lane["repeats"],
            "processes": len(rows),
            "reports": len(rows),
            "samples": sum(lane["samples"] for _ in rows),
            "plan_sha256": c.sha(P / "plan.json"),
            "build_sha256": c.sha(P / "build/build.json"),
            "source": c.artifact(P / "build/source.json"),
            "receipts": c.artifact(OUT / "receipts.json"),
            "event": lane["event"],
            "frequency_hz": lane["frequency_hz"],
            "call_graph": lane["call_graph"],
            "owner": lane["owner"],
        },
    )
