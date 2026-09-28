"""Independent offline audit of frozen inputs and exact executed commands."""
import hashlib
import json
from pathlib import Path

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def read(path):
    return json.loads(path.read_text())


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def artifact(path):
    assert path.is_file() and not path.is_symlink(), path
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def verify(value):
    path = Path(value["path"])
    assert path.is_absolute()
    assert value == artifact(path), path
    return path


def check():
    frozen = read(P / "freeze.json")
    plan = read(P / "plan.json")
    assert frozen["schema"] == "litchi.performance.0827.freeze.v1"
    assert frozen["base"] == plan["base"] == "c5d375083c1624063837c0c57809bd3d371cf72b"
    assert frozen["source"] == frozen["after_source"]
    assert frozen["source"]["revision"] == plan["base"]
    for key in ("plan", "quality", "host", "toolchain", "root_input_manifest",
                "architecture_manifest", "corpus_manifest", "lock_parity", "origin", "provenance",
                "preflight", "preflight_suite", "allocation_schema"):
        verify(frozen[key])
    for group in ("drivers", "support"):
        for name, digest in frozen[group].items():
            assert sha(P / name) == digest, name
    for group in ("tool", "root_inputs", "architecture", "unrelated"):
        for name, digest in frozen[group].items():
            assert sha(ROOT / name) == digest, name
    assert len(frozen["architecture"]) == 35
    for name, expected in frozen["corpus"].items():
        assert {"bytes": (ROOT / name).stat().st_size, "sha256": sha(ROOT / name)} == expected
    for descriptor in frozen["candidate_archives"].values():
        verify(descriptor)
    after = frozen["source"]
    for name, digest in after["files"].items():
        assert sha(ROOT / name) == digest, name
    sources = {}
    for leg in ("before", "after"):
        sources[leg] = {"revision": plan["base"], "files": dict(after["files"])}
        for name in plan["source_allowlist"]:
            digest = sha(P / "candidate" / leg / Path(name).name)
            sources[leg]["files"][name] = digest
            assert frozen[f"{leg}_source_contract"][name] == digest
    quality = read(P / "quality.json")
    quality_source = read(verify(quality["source"]))
    assert all(quality_source[name] == digest for name, digest in after["files"].items())
    env = {
        "CARGO_TARGET_DIR": plan["target"], "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3", "CARGO_PROFILE_RELEASE_DEBUG": "1",
        "CARGO_PROFILE_RELEASE_LTO": "thin", "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1",
        "CARGO_PROFILE_RELEASE_INCREMENTAL": "false", "CARGO_PROFILE_RELEASE_PANIC": "unwind",
        "PYTHONDONTWRITEBYTECODE": "1",
    }
    spans = []
    for row in quality["rows"]:
        spans.append((row["started"], row["ended"], "quality"))
    builds = {}
    for leg in ("before", "after"):
        build = read(P / f"build-{leg}/build.json")
        builds[leg] = build
        assert build["schema"] == f"litchi.performance.0827.build-{leg}.v1"
        assert build["leg"] == leg and build["target"] == plan["target"]
        assert build["profile"] == plan["build"]
        assert read(verify(build["source"])) == sources[leg]
        assert read(verify(build["frozen_inputs"])) == frozen
        assert build["quality"] == frozen["quality"]
        for key in ("root_inputs", "architecture", "corpus", "unrelated", "candidate_archives"):
            assert build[key] == frozen[key]
        assert set(build["binaries"]) == {"native", "artifacts", "observer"}
        assert len(build["rows"]) == 3
        for logical, row in zip(("native", "artifacts", "observer"), build["rows"]):
            spec = plan["binaries"][logical]
            command = ["cargo", "build", "--offline", "--locked", "--release", "--manifest-path",
                       str(ROOT / "tools/perf-baseline/Cargo.toml"), "--bin", spec["cargo_bin"]]
            if spec["features"]:
                command += ["--features", ",".join(spec["features"])]
            assert row["schema"] == "litchi.performance.0827.build-receipt.v1"
            assert row["leg"] == leg and row["binary"] == logical
            assert row["cargo_bin"] == spec["cargo_bin"] and row["features"] == spec["features"]
            assert row["command"] == command and row["environment"] == env and row["exit_code"] == 0
            verify(row["log"])
            descriptor = build["binaries"][logical]
            assert descriptor["cargo_bin"] == spec["cargo_bin"] and descriptor["features"] == spec["features"]
            binary = descriptor["artifact"]
            assert binary["path"] == str(Path(plan["target"]) / f"{leg}-{logical}")
            if Path(binary["path"]).exists():
                verify(binary)
            else:
                assert binary in read(P / "cleanup.json")["binaries"]
            spans.append((row["started"], row["ended"], f"build-{leg}-{logical}"))
        transition = read(P / f"source-transition-{leg}.json")
        spans.append((transition["started"], transition["ended"], f"transition-{leg}"))
        export = read(P / f"artifacts-{leg}-receipt.json")
        assert export["binary"] == build["binaries"]["artifacts"]
        assert read(verify(export["source"])) == sources[leg]
        assert export["frozen_inputs"] == artifact(P / "freeze.json")
        command = ["taskset", "-c", "12", build["binaries"]["artifacts"]["artifact"]["path"],
                   "--output", str(P / plan["paths"]["artifact_output"][leg]),
                   "--filesystem-root", plan["scratch"]]
        for name in dict.fromkeys(case["input"] for case in plan["cases"]):
            command += ["--ooxml-file", name]
        assert export["command"] == command and export["exit_code"] == 0
        verify(export["manifest"])
        verify(export["log"])
        spans.append((export["started"], export["ended"], f"artifacts-{leg}"))
    for lane, leg_hint in (("qualification", "before"), ("qualification", "after"),
                           ("native", None), ("observer", None)):
        directory = P / (f"qualification-{leg_hint}" if leg_hint else lane)
        spec = plan["lanes"][lane]
        orders = [[leg_hint]] if leg_hint else spec["orders"]
        expected = [(block, case, leg) for block, order in enumerate(orders)
                    for case in plan["cases"] for leg in order]
        rows = read(directory / "receipts.json")
        assert len(rows) == len(expected)
        for row, (block, case, leg) in zip(rows, expected):
            assert row["schema"] == "litchi.performance.0827.capture-receipt.v1"
            assert row["lane"] == lane and row["leg"] == leg and row["block"] == block
            assert all(row[name] == value for name, value in case.items())
            assert row["samples"] == spec["samples"] and row["warmup"] == spec["warmup"]
            binary = builds[leg]["binaries"][spec["binary"]]["artifact"]
            assert row["binary"] == binary
            assert row["frozen_inputs"] == artifact(P / "freeze.json")
            assert read(verify(row["source"])) == sources[leg_hint or "after"]
            if leg_hint is None:
                assert row["build_source"] == builds[leg]["source"]
                assert row["order"] == ("BA" if spec["orders"][block] == ["before", "after"] else "AB")
            else:
                verify(row["artifact_admission"])
            report, rss = verify(row["report"]), verify(row["rss"])
            verify(row["log"])
            command = ["/usr/bin/time", "-f", "%M", "-o", str(rss), "taskset", "-c", "12", binary["path"],
                       "--warmup", str(spec["warmup"]), "--samples", str(spec["samples"]),
                       "--case", case["case"], "--json", str(report), "--filesystem-root", plan["scratch"],
                       "--ooxml-file", str(ROOT / case["input"])]
            assert row["command"] == command and row["exit_code"] == 0
            spans.append((row["started"], row["ended"], f"{lane}-{leg}-{block}-{case['case']}"))
    spans.sort()
    assert all(start <= end for start, end, _ in spans)
    assert all(left[1] <= right[0] for left, right in zip(spans, spans[1:])), "execution overlap"
    for leg in ("before", "after"):
        qualifier = read(P / f"qualification-admission-{leg}.json")
        assert qualifier["accepted"] is True
        assert qualifier["qualification_complete"] == artifact(P / f"qualification-{leg}/complete.json")
    print("0827 custody audit PASS: frozen full source, exact commands, six binaries, and serial capture order")


if __name__ == "__main__":
    check()
