#!/usr/bin/env python3
"""Verify durable Ink provenance, allocation receipts and every native attempt."""
import argparse
import hashlib
import json
from pathlib import Path
import statistics
import tomllib

ROUTES = ("forward", "inverse", "noop", "forward_write", "inverse_write", "noop_write")
METRICS = ("allocations", "allocated", "retained", "peak")
QUALIFY_SCHEMA = "litchi.docx-durable-ink.qualify.v2"
ABBA_SCHEMA = "litchi.docx-durable-ink-abba.v2"
STREAM_SCHEMA = "litchi.docx-durable-ink-stream.v1"
COMPARISON = "integrated-route"
SHARED_PREFIXES = (
    "crates/litchi-core/",
    "crates/litchi-opc/",
    "crates/litchi-ooxml-common/",
    "crates/litchi-drawingml/",
)
VALIDATION_NAMES = {
    "forward.docx",
    "forward.json",
    "inverse.docx",
    "inverse.json",
    "noop.docx",
    "noop.json",
    "source.docx",
    "streaming.json",
}


def load(path):
    return json.loads(Path(path).read_text())


def artifact(path):
    path = Path(path).resolve()
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        while chunk := stream.read(1024 * 1024):
            digest.update(chunk)
    return dict(path=str(path), bytes=path.stat().st_size, sha256=digest.hexdigest())


def verify(row):
    assert artifact(row["path"]) == row, row["path"]


def command(row):
    assert row["exit_code"] == 0, row
    assert row.get("overlap") is None, row
    verify(row["stdout"])
    verify(row["stderr"])


def validate_build_refs(build):
    assert set(build["endpoints"]) == {"base", "candidate"}
    assert set(build["resolved_refs"]) == {"base", "candidate"}
    assert build["resolved_refs"]["base"] != build["resolved_refs"]["candidate"]
    for label in ("base", "candidate"):
        endpoint = build["endpoints"][label]
        assert endpoint["source"]["commit"] == build["resolved_refs"][label]


def source_delta(base, candidate, prefixes, base_commit, candidate_commit):
    selected = lambda path: path.startswith(tuple(prefixes))
    paths = sorted({path for path in set(base) | set(candidate) if selected(path)})
    base_only, candidate_only, changed = [], [], []
    for path in paths:
        before, after = base.get(path), candidate.get(path)
        if before is None:
            candidate_only.append(dict(path=path, candidate=dict(bytes=after[0], sha256=after[1])))
        elif after is None:
            base_only.append(dict(path=path, base=dict(bytes=before[0], sha256=before[1])))
        elif before != after:
            changed.append(dict(path=path, base=dict(bytes=before[0], sha256=before[1]),
                                candidate=dict(bytes=after[0], sha256=after[1])))
    return dict(base_commit=base_commit, candidate_commit=candidate_commit,
                prefixes=list(prefixes), base_only=base_only,
                candidate_only=candidate_only, changed=changed)


def validate_shared_source_delta(build, source_maps):
    assert build["comparison"] == COMPARISON
    expected = source_delta(
        source_maps["base"], source_maps["candidate"], SHARED_PREFIXES,
        build["endpoints"]["base"]["source"]["commit"],
        build["endpoints"]["candidate"]["source"]["commit"],
    )
    assert build["shared_source_delta"] == expected


def validate_calibration_case(case, target_ns, minimum_ns):
    """Independently validate calibration and accepted ABBA sample bounds."""
    assert case["calibrations"]
    route, count = case["route"], case["count"]
    for calibration in case["calibrations"]:
        assert set(calibration) == {"A", "B"}
        values = list(calibration.values())
        for value in values:
            assert (value["route"], value["iterations"]) == (route, count)
            assert isinstance(value["ns"], int) and value["ns"] > 0
        for key in ("checksum", "output_bytes", "output_sha256"):
            assert len({value[key] for value in values}) == 1
    assert min(value["ns"] for value in case["calibrations"][-1].values()) >= target_ns
    for block in case["blocks"]:
        legs = block["legs"]
        assert [leg["side"] for leg in legs] == list("ABBA")
        for leg in legs:
            value = leg["value"]
            assert (value["route"], value["iterations"]) == (route, count)
            assert isinstance(value["ns"], int) and value["ns"] >= minimum_ns


def collect_build(directory):
    path = directory / "build.json"
    build = load(path)
    assert build["schema"] == QUALIFY_SCHEMA
    assert build["workload"] == "docx-durable-ink"
    assert build["status"] == "complete"
    validate_build_refs(build)
    for key in ("builder", "probe", "input_lock"):
        verify(build[key])
    source_maps = {}
    for label, endpoint in build["endpoints"].items():
        for key in ("project_manifest", "project_source", "lock"):
            verify(endpoint[key])
        profile = tomllib.loads(Path(endpoint["project_manifest"]["path"]).read_text())["profile"]["release"]
        assert profile == {"lto": True, "panic": "abort"}, profile
        assert endpoint["project_source"]["sha256"] == build["probe"]["sha256"]
        verify(endpoint["source"]["manifest"])
        source = load(endpoint["source"]["manifest"]["path"])
        assert source["commit"] == endpoint["source"]["commit"]
        verify(source["archive"])
        mapping = {}
        for row in source["files"]:
            actual = artifact(directory / f"source-{label}" / row["path"])
            assert (actual["bytes"], actual["sha256"]) == (row["bytes"], row["sha256"]), row
            mapping[row["path"]] = (row["bytes"], row["sha256"])
        source_maps[label] = mapping
        for row in endpoint["commands"]:
            command(row)
        for row in endpoint["binaries"].values():
            verify(row)
        for rows in endpoint["validation"].values():
            assert {Path(row["path"]).name for row in rows} == VALIDATION_NAMES
            for row in rows:
                verify(row)
    base, candidate = (build["endpoints"][key] for key in ("base", "candidate"))
    for key in ("registry", "rustc", "cargo"):
        assert base[key] == candidate[key], key
    assert base["lock"]["sha256"] == candidate["lock"]["sha256"]
    expected = {
        Path(row["path"]).name: row["sha256"] for row in base["validation"]["native"]
        if Path(row["path"]).name != "streaming.json"
    }
    base_stream = None
    for endpoint in (base, candidate):
        for rows in endpoint["validation"].values():
            actual = {
                Path(row["path"]).name: row["sha256"] for row in rows
                if Path(row["path"]).name != "streaming.json"
            }
            assert actual == expected
            stream = load(next(row["path"] for row in rows
                               if Path(row["path"]).name == "streaming.json"))
            assert stream["schema"] == STREAM_SCHEMA
            assert stream["read_chunk"] > 0
            assert stream["sink_window"] > 0
            assert stream["base_bytes"] > 0 and stream["target_bytes"] > 0
            assert [row["route"] for row in stream["routes"]] == ["forward", "inverse", "noop"]
            signature = [
                (row["route"], row["bytes"], row["sha256"]) for row in stream["routes"]
            ]
            current_stream = (
                stream["read_chunk"], stream["sink_window"], stream["base_bytes"],
                stream["target_bytes"], signature
            )
            if base_stream is None:
                base_stream = current_stream
            else:
                assert current_stream == base_stream
    validate_shared_source_delta(build, source_maps)
    return dict(receipt=artifact(path), **build)


def collect_allocations(directory, build):
    path = directory / "manifest.json"
    data = load(path)
    assert data["schema"] == "litchi.docx-durable-ink-allocation.v1"
    assert data["workload"] == "docx-durable-ink"
    assert data["status"] == "complete" and data["build"] == build["receipt"]
    assert [row["route"] for row in data["rows"]] == list(ROUTES)
    rows = []
    for row in data["rows"]:
        values = row["values"]
        assert [v["endpoint"] for v in values] == ["base", "candidate", "candidate", "base"]
        for value in values:
            verify(value["receipt"])
            receipt = load(value["receipt"]["path"])
            command(receipt)
            assert receipt["command"] == [build["endpoints"][value["endpoint"]]["binaries"]["allocation"]["path"], row["route"], "100"]
            result = value["value"]
            assert result == load(receipt["stdout"]["path"])
            assert (result["route"], result["iterations"]) == (row["route"], 100)
        for key in ("checksum", "output_bytes", "output_sha256"):
            assert len({v["value"][key] for v in values}) == 1, (row["route"], key)
        for metric in METRICS:
            assert values[0]["value"][metric] == values[3]["value"][metric]
            assert values[1]["value"][metric] == values[2]["value"][metric]
        rows.append(dict(route=row["route"], count=100,
                         base={key: values[0]["value"][key] for key in METRICS},
                         candidate={key: values[1]["value"][key] for key in METRICS}))
    return dict(receipt=artifact(path), timing="unqualified; allocation counts only", rows=rows)


def collect_native(directory, build):
    binaries = {side: build["endpoints"][label]["binaries"]["native"]
                for side, label in (("A", "base"), ("B", "candidate"))}
    attempts, files = [], []
    cases = {route: [] for route in ROUTES}
    minimum = None
    runner = None
    for attempt in sorted(directory.glob("native-a*")):
        if not attempt.is_dir():
            continue
        manifest = load(attempt / "manifest.json")
        assert manifest["schema"] == ABBA_SCHEMA
        assert manifest["workload"] == "docx-durable-ink"
        assert manifest["build_manifest"] == build["receipt"]
        assert manifest["binaries"] == binaries
        assert (manifest["target_ns"], manifest["minimum_ns"]) == (500_000_000, 150_000_000)
        verify(manifest["runner"])
        verify(manifest["build_manifest"])
        if runner is None:
            runner = manifest["runner"]
        assert runner == manifest["runner"]
        completed = 0
        for case in manifest["cases"]:
            assert case["route"] in cases
            validate_calibration_case(
                case, manifest["target_ns"], manifest["minimum_ns"]
            )
            for index, block in enumerate(case["blocks"]):
                legs = block["legs"]
                assert [leg["side"] for leg in legs] == list("ABBA")
                for leg in legs:
                    label, value = leg["label"], leg["value"]
                    assert (value["route"], value["iterations"]) == (case["route"], case["count"])
                    assert value["ns"] >= 150_000_000
                    minimum = value["ns"] if minimum is None else min(minimum, value["ns"])
                    receipt = load(attempt / f"{label}.receipt.json")
                    command(receipt)
                    assert receipt["command"] == ["taskset", "-c", str(manifest["cpu"]), binaries[leg["side"]]["path"], case["route"], str(case["count"])]
                    assert load(receipt["stdout"]["path"]) == value
                    during = list(attempt.glob(f"{label}-during-*.host.json"))
                    assert during
                    guarded = [attempt / f"{label}-before.host.json", attempt / f"{label}-after.host.json", *during]
                    assert all(load(path)["processes"] and not load(path)["blockers"] for path in guarded)
                for key in ("checksum", "output_bytes", "output_sha256"):
                    assert len({leg["value"][key] for leg in legs}) == 1
                ratio = sum(leg["value"]["ns"] for leg in legs[1:3]) / sum(legs[i]["value"]["ns"] for i in (0, 3))
                assert ratio == block["ratio"]
                cases[case["route"]].append(dict(attempt=attempt.name, block=index,
                                                 count=case["count"], ratio=ratio))
                completed += 1
        hosts = list(attempt.glob("*.host.json"))
        blocked = [artifact(path) for path in hosts if load(path)["blockers"]]
        if manifest["status"] == "complete":
            assert not blocked
            assert [case["route"] for case in manifest["cases"]] == list(ROUTES)
            assert all(len(case["blocks"]) == 5 for case in manifest["cases"])
        attempts.append(dict(directory=str(attempt), receipt=artifact(attempt / "manifest.json"),
                             raw_status=manifest["status"], qualified_blocks=completed,
                             host_samples=len(hosts), blocked_hosts=blocked,
                             runner_log=artifact(directory / f"{attempt.name}.log")))
        files.extend(artifact(path) for path in sorted(attempt.iterdir()) if path.is_file())
    summary = [dict(route=route, blocks=blocks,
                    median_ratio=statistics.median(block["ratio"] for block in blocks) if blocks else None)
               for route, blocks in cases.items()]
    return dict(attempts=attempts, cases=summary, minimum_sample_ns=minimum), files


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument(
        "--allocation",
        help="optional allocation receipt directory (for example allocation-a1)",
    )
    args = parser.parse_args()
    directory, out = args.directory.resolve(), args.out.resolve()
    build = collect_build(directory)
    allocations = (
        collect_allocations(directory / args.allocation, build)
        if args.allocation
        else dict(status="not-run; native timing is qualified independently")
    )
    native, files = collect_native(directory, build)
    if not native["attempts"] or not any(
        attempt["raw_status"] == "complete" for attempt in native["attempts"]
    ):
        raise RuntimeError("native durable Ink timing qualification has not completed")
    out.mkdir(parents=True, exist_ok=True)
    # Keep the unabridged inventory outside Git; the reduced evidence links it.
    inventory = directory / "native-artifacts.json"
    inventory.write_text(json.dumps(files, indent=2) + "\n")
    native["inventory"] = artifact(inventory)
    summary = dict(status="receipts verified; coordinator review required",
                   build=build, allocations=allocations, native=native)
    data = json.dumps(summary, indent=2) + "\n"
    assert len(data.encode()) < 1_048_576
    (out / "evidence.json").write_text(data)
    print(summary["status"])


if __name__ == "__main__":
    main()
