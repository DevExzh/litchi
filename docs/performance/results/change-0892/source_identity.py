#!/usr/bin/env python3
"""Prove the retained 0891 executable's source matches the current workload."""
import argparse
import hashlib
import json
from pathlib import Path
import subprocess

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
REVISION = "0a0957ef7249b4b16ec26ed81d1fb32b7288bb0b"
DIAGNOSTICS = {
    "tools/perf-baseline/src/bin/xml_attribute_bounds.rs",
    "tools/perf-baseline/src/bin/xml_attribute_bounds/boundary_tests.rs",
    "tools/perf-baseline/src/bin/xml_attribute_bounds/input_tests.rs",
    "tools/perf-baseline/src/bin/xml_attribute_bounds/qualification_tests.rs",
}


def read(path):
    return json.loads(path.read_text())


def identity(path):
    h = hashlib.sha256()
    with path.open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return {"path": str(path.resolve()), "bytes": path.stat().st_size,
            "sha256": h.hexdigest()}


def git(*args):
    return subprocess.check_output(["git", *args], cwd=ROOT, text=True).splitlines()


def audit():
    baseline = P.parent / "change-0885"
    adopted = P.parent / "change-0891"
    manifest = read(baseline / "source.json")
    freeze = read(adopted / "candidate-freeze-v3.json")
    build = read(adopted / "build.json")
    expected = dict(manifest["files"])
    expected.update({x["path"]: x for x in freeze["files"]})
    assert len(expected) == build["source_file_count"] == build["source_rechecked_files"] == 10766
    assert build["status"] == "pass"
    assert not git("diff", "--name-only", REVISION, "--", *manifest["selection"])
    tracked = set(git("ls-files", "--", *manifest["selection"]))
    assert tracked == set(expected) | DIAGNOSTICS
    excluded = git("ls-files", "--others", "--exclude-standard", "--", *manifest["selection"])
    assert set(excluded) <= {"crates/litchi/examples/bench_word_parse.rs"}
    checked = 0
    for name, record in expected.items():
        if name in DIAGNOSTICS:
            continue
        actual = identity(ROOT / name)
        assert actual["bytes"] == record["bytes"] and actual["sha256"] == record["sha256"], name
        checked += 1
    binary = Path(build["binaries"]["native"]["path"])
    actual = identity(binary)
    assert all(actual[k] == build["binaries"]["native"][k] for k in ("bytes", "sha256"))
    for key, path in (("candidate_freeze", adopted / "candidate-freeze-v3.json"),
                      ("candidate_patch", adopted / "candidate-v3.patch"),
                      ("source_archive", baseline / "source.tar.gz"),
                      ("source_manifest", baseline / "source.json")):
        current = identity(path)
        assert all(current[k] == build[key][k] for k in ("bytes", "sha256")), key
    evidence = [baseline / "source.json", baseline / "source.tar.gz",
                adopted / "candidate-freeze-v3.json", adopted / "candidate-v3.patch",
                adopted / "build.json", adopted / "native-build.json", adopted / "native.d",
                adopted / "decision.json", adopted / "quality.json"]
    return {"schema": "litchi.current-profile-source-0892.v1", "status": "pass",
            "source_revision": REVISION, "observed_head": git("rev-parse", "HEAD")[0],
            "retained_source_union_files": len(expected), "current_tracked_selected_files": len(tracked),
            "live_files_hash_checked": checked, "binary": actual,
            "diagnostic_exceptions": sorted(DIAGNOSTICS),
            "exception_reason": "Only the separate xml_attribute_bounds diagnostic changed after 0885; it is not linked into the profiled binary.",
            "untracked_excluded": excluded, "evidence": [identity(x) for x in evidence],
            "production_changes": False, "performance_claim": None}


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    result = audit()
    with args.output.open("x") as f:
        json.dump(result, f, indent=2)
        f.write("\n")
    print("Retained binary and current workload source identity pass")
