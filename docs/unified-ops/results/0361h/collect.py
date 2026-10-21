#!/usr/bin/env python3
"""Verify protected character-style evidence and emit a compact summary."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import statistics


def load(path):
    return json.loads(Path(path).read_text())


def artifact(path):
    path = Path(path).resolve()
    return dict(path=str(path), bytes=path.stat().st_size,
                sha256=hashlib.sha256(path.read_bytes()).hexdigest())


def verify(row):
    assert artifact(row["path"]) == row, row["path"]


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--coord", type=Path, required=True)
    parser.add_argument("--out", type=Path, required=True)
    parser.add_argument("--native", type=Path)
    args = parser.parse_args()
    coord, out = args.coord.resolve(), args.out.resolve()
    out.mkdir(parents=True, exist_ok=True)
    sources = {}
    source_evidence = []
    for name in ("base", "v2", "v3"):
        path = coord / f"source-manifest-{name}.json"
        manifest = load(path)
        assert manifest["base_commit"] == "61515148e1"
        entries = manifest["sources"]
        directory = coord / f"source-{name}"
        mapping = {}
        for entry in entries:
            data = artifact(directory / entry["path"])
            assert (data["bytes"], data["sha256"]) == (entry["bytes"], entry["sha256"]), entry
            mapping[entry["path"]] = entry["sha256"]
        sources[name] = mapping
        source_evidence.append(dict(candidate=name, manifest=artifact(path), file_count=len(entries),
            content_digest=hashlib.sha256(json.dumps(mapping, sort_keys=True).encode()).hexdigest()))
    changed = {}
    for name in ("v2", "v3"):
        assert sources[name].keys() == sources["base"].keys()
        changed[name] = [path for path in sources[name] if sources[name][path] != sources["base"][path]]
        assert len(changed[name]) == 7, changed[name]
    builds, allocations, profiles = {}, {}, {}
    for name in ("base", "v2", "v3"):
        path = coord / f"build-{name}/manifest.json"
        build = load(path)
        assert build["status"] == "complete"
        for row in (build["probe"], build["builder"], build["lock"], *build["artifacts"].values()):
            verify(row)
        for command in build["commands"]:
            assert command["exit_code"] == 0
            verify(command["stdout"])
            verify(command["stderr"])
        builds[name] = dict(manifest=artifact(path), **build)
        fingerprints = list((coord / f"target-{name}/release/.fingerprint").glob(
            "litchi-iwa-text-wire-*/lib-litchi_iwa_text_wire.json"))
        assert len(fingerprints) == 1
        fingerprint = load(fingerprints[0])
        profile = {key: fingerprint[key] for key in
                   ("rustc", "features", "declared_features", "profile", "rustflags", "config", "compile_kind")}
        profiles[name] = dict(artifact=artifact(fingerprints[0]), profile=profile)
        allocation_path = coord / f"allocation-{name}/manifest.json"
        allocation = load(allocation_path)
        verify(allocation["binary"])
        assert allocation["binary"] == build["artifacts"]["allocation"]
        assert len(allocation["rows"]) == 11
        for row in allocation["rows"]:
            value = row["result"]
            assert value["count"] == 100 and value["retained_bytes"] == 0
            raw = coord / f"allocation-{name}/{value['route']}-{value['runs']}.stdout"
            assert load(raw) == value
        allocations[name] = dict(manifest=artifact(allocation_path), **allocation)
    for key in ("probe", "builder", "lock"):
        assert len({builds[name][key]["sha256"] for name in builds}) == 1, key
    assert profiles["base"]["profile"] == profiles["v2"]["profile"] == profiles["v3"]["profile"]
    fidelity_path = coord / "native-fidelity/manifest.json"
    fidelity = load(fidelity_path)
    assert len(fidelity) == 6
    for row in fidelity:
        assert len(row["files"]) == 10
        for entry in row["files"]:
            path = coord / f"native-fidelity/{row['candidate']}-{row['size']}" / entry["name"]
            expected = coord / f"allocation-base/validation-{row['size']}" / entry["name"]
            assert path.read_bytes() == expected.read_bytes()
            measured = artifact(path)
            assert (measured["bytes"], measured["sha256"]) == (entry["bytes"], entry["sha256"])
    checks = []
    for name in ("initial-checks.json", "boundary-checks.json", "final-checks-v1.json", "repair-checks.json"):
        path = coord / name
        checks.append(dict(receipt=artifact(path), rows=load(path)))
    final_checks = load(coord / "repair-checks.json")
    assert {row["name"] for row in final_checks} == {"repair-clippy", "repair-owners", "repair-glossary"}
    assert all(row["exit_code"] == 0 for row in final_checks)
    raw = (coord / "raw-buffer-repair-v2.log").read_text()
    assert re.findall(r"test result: ok\. (\d+) passed; 0 failed", raw) == ["57", "8"]
    owner = (coord / "repair-owners.log").read_text()
    assert re.findall(r"test result: ok\. (\d+) passed; 0 failed", owner) == ["330", "477", "208"]
    logs = [artifact(coord / name) for name in (
        "raw-buffer-repair-v2.log", "repair-clippy.log", "repair-owners.log", "repair-glossary.log",
        "raw-initial.log", "consumer-initial.log", "boundary-tests.log", "boundary-tests-v2.log",
        "owners-boundary.log", "clippy.log", "raw-v3.log", "clippy-v2.log", "owners-all.log",
        "raw-buffer-repair.log", "glossary.log", "inventory.tsv")]
    summary = dict(base_commit="61515148e1", status="allocation/correctness checked; native pending",
        source_manifests=source_evidence, source_delta=changed, builds=builds, profiles=profiles,
        allocations=allocations, native_fidelity=dict(manifest=artifact(fidelity_path), rows=fidelity),
        validation=checks, validation_logs=logs)
    if args.native:
        native = args.native.resolve()
        manifest = load(native / "manifest.json")
        assert manifest["status"] == "complete" and len(manifest["cases"]) == 11
        assert manifest["binaries"]["A"] == builds["base"]["artifacts"]["native"]
        assert manifest["binaries"]["B"] == builds["v3"]["artifacts"]["native"]
        verify(manifest["runner"])
        hosts = list(native.glob("*.host.json"))
        assert hosts
        for path in hosts:
            host = load(path)
            assert host["blockers"] == [] and host["processes"], path
        cases = []
        minimum = None
        for case in manifest["cases"]:
            assert len(case["blocks"]) == 5
            ratios = []
            for block in case["blocks"]:
                legs = block["legs"]
                assert [leg["side"] for leg in legs] == list("ABBA")
                for leg in legs:
                    label, row = leg["label"], leg["value"]
                    assert (row["route"], row["runs"], row["count"]) == (case["route"], case["size"], case["count"])
                    assert row["elapsed_ns"] >= 150_000_000
                    minimum = row["elapsed_ns"] if minimum is None else min(minimum, row["elapsed_ns"])
                    receipt = load(native / f"{label}.receipt.json")
                    assert receipt["exit_code"] == 0 and receipt["overlap"] is None
                    verify(receipt["stdout"])
                    verify(receipt["stderr"])
                    assert load(receipt["stdout"]["path"]) == row
                    assert Path(receipt["command"][3]).resolve() == Path(manifest["binaries"][leg["side"]]["path"])
                    assert (native / f"{label}-before.host.json").exists()
                    assert (native / f"{label}-after.host.json").exists()
                    assert list(native.glob(f"{label}-during-*.host.json"))
                ratio = sum(leg["value"]["elapsed_ns"] for leg in legs[1:3]) / sum(legs[i]["value"]["elapsed_ns"] for i in (0, 3))
                assert ratio == block["ratio"]
                ratios.append(ratio)
            cases.append(dict(route=case["route"], size=case["size"], count=case["count"], ratios=ratios,
                              median_ratio=statistics.median(ratios)))
        all_cases = {(row["route"], row["size"]): {
            "route": row["route"], "size": row["size"], "blocks": []
        } for row in cases}
        attempts, files = [], []
        # A later interrupted route does not erase earlier clean, complete
        # blocks. Retain every attempt with these exact binaries/protocol.
        for attempt in sorted(coord.glob("native-v3-a*")):
            if not attempt.is_dir():
                continue
            prior = load(attempt / "manifest.json")
            assert prior["binaries"] == manifest["binaries"]
            assert prior["runner"] == manifest["runner"]
            assert (prior["cpu"], prior["target_ns"], prior["minimum_ns"]) == (
                manifest["cpu"], 500_000_000, 150_000_000)
            bad_hosts = [artifact(path) for path in sorted(attempt.glob("*.host.json"))
                         if load(path)["blockers"]]
            completed = 0
            for case in prior["cases"]:
                key = (case["route"], case["size"])
                assert key in all_cases
                for index, block in enumerate(case["blocks"]):
                    legs = block["legs"]
                    assert [leg["side"] for leg in legs] == list("ABBA")
                    for leg in legs:
                        label, value = leg["label"], leg["value"]
                        assert (value["route"], value["runs"], value["count"]) == (key[0], key[1], case["count"])
                        assert value["elapsed_ns"] >= 150_000_000
                        minimum = min(minimum, value["elapsed_ns"])
                        receipt = load(attempt / f"{label}.receipt.json")
                        assert receipt["exit_code"] == 0 and receipt["overlap"] is None
                        verify(receipt["stdout"])
                        verify(receipt["stderr"])
                        assert load(receipt["stdout"]["path"]) == value
                        assert Path(receipt["command"][3]).resolve() == Path(manifest["binaries"][leg["side"]]["path"])
                        guarded = [attempt / f"{label}-before.host.json", attempt / f"{label}-after.host.json"]
                        during = list(attempt.glob(f"{label}-during-*.host.json"))
                        assert during
                        assert all(load(path)["blockers"] == [] for path in guarded + during)
                    ratio = sum(leg["value"]["elapsed_ns"] for leg in legs[1:3]) / sum(legs[i]["value"]["elapsed_ns"] for i in (0, 3))
                    assert ratio == block["ratio"]
                    all_cases[key]["blocks"].append(dict(attempt=attempt.name, block=index,
                                                        calls_per_leg=case["count"], ratio=ratio))
                    completed += 1
            attempts.append(dict(directory=str(attempt), manifest=artifact(attempt / "manifest.json"),
                                 status=prior["status"], qualified_blocks=completed, blocked_hosts=bad_hosts,
                                 disposition="complete" if prior["status"] == "complete" else
                                     "stopped on host overlap; earlier clean blocks retained",
                                 runner_log=artifact(coord / f"{attempt.name}.log")))
            files.extend(artifact(path) for path in sorted(attempt.iterdir()) if path.is_file())
        assert len(attempts) >= 1
        for row in all_cases.values():
            assert len(row["blocks"]) >= 5
            row["median_ratio"] = statistics.median(block["ratio"] for block in row["blocks"])
        (out / "native-artifacts.json").write_text(json.dumps(files, indent=2) + "\n")
        summary["native"] = dict(manifest=artifact(native / "manifest.json"), cases=cases,
                                 host_samples=len(hosts), minimum_sample_ns=minimum,
                                 attempts=attempts, all_qualified_cases=list(all_cases.values()),
                                 inventory=artifact(out / "native-artifacts.json"))
        summary["status"] = "evidence checked; coordinator acceptance required"
    (out / "evidence.json").write_text(json.dumps(summary, indent=2) + "\n")
    print(summary["status"])


if __name__ == "__main__":
    main()
