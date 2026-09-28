"""Offline replay and source-custody checks for the 0826 checker repair."""
import hashlib
import json
from pathlib import Path
import re

import preflight

P = Path(__file__).resolve().parent
ROOT = P.parents[3]


def read(path):
    return json.loads(path.read_text())


def descriptor(value):
    path = ROOT / value["path"]
    assert path.is_file() and not path.is_symlink(), path
    assert preflight.descriptor(path) == value, path
    return path


def check():
    origin = read(P / "origin.json")
    for group in ("production", "runtime_tool", "architecture", "unrelated"):
        for name, digest in origin[group].items():
            assert preflight.sha(ROOT / name) == digest, name
    assert len(origin["architecture"]) == 35
    assert origin["base"] == "c8672fa1fa044c5016d64a4bba021f407829b2c8"
    value = preflight.derive()
    assert (P / "preflight.json").read_text() == json.dumps(value, indent=2, sort_keys=True) + "\n"
    quality = read(P / "quality.json")
    assert quality["schema"] == "litchi.performance.0826.quality.v1" and quality["status"] == "pass"
    for source in quality["sources"]:
        descriptor(source["original"])
        assert descriptor(source["snapshot"]).read_bytes() == (ROOT / source["original"]["path"]).read_bytes()
    assert len(quality["rows"]) == 2
    test_counts = []
    for row, flags in zip(quality["rows"], ([], ["-O"])):
        assert row["exit_code"] == 0
        assert row["command"] == ["python3", "-B", *flags, "-m", "unittest", "tools.test_perf_allocation_schema", "-v"]
        log = descriptor(row["log"]).read_text()
        assert log.rstrip().endswith("OK"), "regression suite not successful"
        matches = re.findall(r"^Ran (\d+) tests? in ", log, flags=re.M)
        assert len(matches) == 1 and int(matches[0]) > 0
        test_counts.append(int(matches[0]))
    assert test_counts[0] == test_counts[1]
    assert quality["rows"][0]["ended"] <= quality["rows"][1]["started"]
    for attempt in P.glob("quality-*/receipt.json"):
        receipt = read(attempt)
        for source in receipt["sources"]:
            snapshot = descriptor(source["snapshot"])
            assert hashlib.sha256(snapshot.read_bytes()).hexdigest() == source["original"]["sha256"]
        for row in receipt["rows"]:
            descriptor(row["log"])
    initial = read(P / "preflight-attempt-0/preflight.json")
    assert initial["validator"]["sha256"] == preflight.sha(P / "preflight-attempt-0/perf_allocation_schema.py")
    assert initial["driver"]["sha256"] == preflight.sha(P / "preflight-attempt-0/preflight.py")
    assert initial["reports"] == value["reports"] == 325
    assert not list(P.rglob("__pycache__"))
    cleanup = read(P / "cleanup.json")
    assert cleanup["schema"] == "litchi.performance.0826.cleanup.v1"
    assert cleanup["verified"] is True and cleanup["build_or_scratch_roots_created"] is False
    assert cleanup["absent_roots"] == [str(ROOT.parent / "litchi-target-0826"), str(ROOT.parent / "litchi-fs-0826")]
    assert not any(Path(name).exists() for name in cleanup["absent_roots"])
    print(f"0826 validation PASS: {test_counts[0]} regression tests in normal and optimized Python; 325 report schemas; runtime and normative sources unchanged")


if __name__ == "__main__":
    check()
