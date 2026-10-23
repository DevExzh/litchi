#!/usr/bin/env python3
"""Verify constructor provenance, allocation receipts and every native attempt."""
import argparse
import hashlib
import json
import re
from pathlib import Path
import statistics
import tomllib

ROUTES = ("field_path", "field_info", "message_info", "archive_info", "archive_object",
          "archive_object_reject", "parse", "validate", "encode",
          "retained_parse", "retained_validate", "retained_encoded_len", "retained_encode",
          "retained_validate_tight", "retained_encode_tight")
METRICS = ("allocations", "allocated", "retained", "peak")


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


def collect_build(directory):
    path = directory / "build.json"
    build = load(path)
    assert build["status"] == "complete"
    for key in ("builder", "probe", "input_lock"):
        verify(build[key])
    source_maps = []
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
            mapping[row["path"]] = row["sha256"]
        source_maps.append(mapping)
        for row in endpoint["commands"]:
            command(row)
        for row in endpoint["binaries"].values():
            verify(row)
        for rows in endpoint["validation"].values():
            assert len(rows) == 8
            for row in rows:
                verify(row)
    base, candidate = (build["endpoints"][key] for key in ("base", "candidate"))
    for key in ("registry", "rustc", "cargo"):
        assert base[key] == candidate[key], key
    assert base["lock"]["sha256"] == candidate["lock"]["sha256"]
    expected = {Path(row["path"]).name: row["sha256"] for row in base["validation"]["native"]}
    for endpoint in (base, candidate):
        for rows in endpoint["validation"].values():
            assert {Path(row["path"]).name: row["sha256"] for row in rows} == expected
    shared = ("crates/litchi-iwa-common/", "crates/litchi-iwa-protos/")
    assert {k: v for k, v in source_maps[0].items() if k.startswith(shared)} == {
        k: v for k, v in source_maps[1].items() if k.startswith(shared)}
    return dict(receipt=artifact(path), **build)


def collect_allocations(directory, build):
    path = directory / "manifest.json"
    data = load(path)
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
    planner = None
    attempt_dirs = [path for path in directory.glob("native-a*") if path.is_dir()]
    assert all(re.fullmatch(r"native-a[1-9][0-9]*", path.name) for path in attempt_dirs)
    attempt_dirs.sort(key=lambda path: int(path.name.removeprefix("native-a")))
    assert [int(path.name.removeprefix("native-a")) for path in attempt_dirs] == list(
        range(1, len(attempt_dirs) + 1)), "missing native attempt"
    for attempt in attempt_dirs:
        manifest = load(attempt / "manifest.json")
        assert manifest["schema"] == "constructor-abba.v2"
        assert manifest["binaries"] == binaries
        requested = [{"route": route, "blocks": 5 - len(blocks)}
                     for route, blocks in cases.items() if len(blocks) < 5]
        assert manifest["requested"] == requested, (attempt, "resume schedule")
        assert requested, (attempt, "attempt after all required blocks")
        verify(manifest["planner"])
        if planner is None:
            planner = manifest["planner"]
        assert planner == manifest["planner"]
        assert [case["route"] for case in manifest["cases"]] == [
            request["route"] for request in requested[:len(manifest["cases"])]]
        assert len(manifest["cases"]) <= len(requested)
        assert (manifest["target_ns"], manifest["minimum_ns"]) == (500_000_000, 150_000_000)
        verify(manifest["runner"])
        if runner is None:
            runner = manifest["runner"]
        assert runner == manifest["runner"]
        completed = 0
        for case_index, case in enumerate(manifest["cases"]):
            assert case["route"] in cases
            assert 0 <= len(case["blocks"]) <= requested[case_index]["blocks"]
            if case_index + 1 < len(manifest["cases"]):
                assert len(case["blocks"]) == requested[case_index]["blocks"]
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
            assert [case["route"] for case in manifest["cases"]] == [
                request["route"] for request in requested]
            assert all(len(case["blocks"]) == request["blocks"]
                       for case, request in zip(manifest["cases"], requested, strict=True))
        attempts.append(dict(directory=str(attempt), receipt=artifact(attempt / "manifest.json"),
                             raw_status=manifest["status"], qualified_blocks=completed,
                             host_samples=len(hosts), blocked_hosts=blocked,
                             runner_log=artifact(directory / f"{attempt.name}.log")))
        files.extend(artifact(path) for path in sorted(attempt.iterdir()) if path.is_file())
    summary = [dict(route=route, blocks=blocks,
                    median_ratio=statistics.median(block["ratio"] for block in blocks) if blocks else None)
               for route, blocks in cases.items()]
    return dict(attempts=attempts, cases=summary, minimum_sample_ns=minimum,
                required_blocks_per_route=5, qualified=all(len(blocks) == 5 for blocks in cases.values())), files


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("directory", type=Path)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--allocation", default="allocation-a1")
    args = parser.parse_args()
    directory, out = args.directory.resolve(), args.out.resolve()
    build = collect_build(directory)
    allocations = collect_allocations(directory / args.allocation, build)
    native, files = collect_native(directory, build)
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
