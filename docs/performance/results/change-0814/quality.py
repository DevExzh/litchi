"""Validate the sealed 0813-after six-gate production quality evidence.

The 0814 production source has the same 9,196 tracked source-file hashes as
the retained 0813 after source. Its commit revision is recorded separately.
This driver reuses the six retained production gates instead of running a
second full Cargo quality lane; the fresh 0814 quality lane is the three-gate,
36-test public probe check.
"""

import custody as c
from pathlib import Path


P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.read(P / "origin.json")
SEALED = P.parent / "change-0813"
SEALED_SUMMARY = SEALED / "quality-summary.json"
SEALED_SOURCE = SEALED / "build-after/source.json"
SEALED_SEAL = SEALED / "seal.json"
OUT = P / "quality-reuse"
assert not OUT.exists(), f"refusing to overwrite {OUT}"
assert PLAN["schema"] == "litchi.performance.0814.current-production.v1"
assert ORIGIN["base"] == PLAN["source_revision"]
assert SEALED_SUMMARY.is_file() and SEALED_SOURCE.is_file() and SEALED_SEAL.is_file()

source = c.source()
assert source["revision"] == ORIGIN["base"]
assert len(source["files"]) == PLAN["source_file_count"]
sealed_source = c.read(SEALED_SOURCE)
assert len(sealed_source["files"]) == PLAN["source_file_count"]
assert source["files"] == sealed_source["files"]
root_inputs = c.assert_root_inputs()
architecture = c.architecture_hashes()
unrelated = c.assert_unrelated()
probe = c.assert_probe()
summary = c.read(SEALED_SUMMARY)
seal_files = c.read(SEALED_SEAL)["files"]
assert summary["schema"] == "litchi.performance.0813.quality-summary.v1"


def sealed_artifact(value):
    path = Path(value["path"])
    assert c.artifact(path) == value
    relative = str(path.relative_to(c.ROOT))
    assert seal_files[relative] == value["sha256"], relative
    return relative


sealed_witnesses = [sealed_artifact(c.artifact(SEALED_SUMMARY)), sealed_artifact(c.artifact(SEALED_SOURCE))]
production = summary["production_after"]
assert production["schema"] == "litchi.performance.0813.quality-after.v1"
assert len(production["gates"]) == 6
assert all(
    gate["exit_code"] == 0 and gate["status"] == "pass"
    for gate in production["gates"]
)
assert production["tests"]["failed"] == 0
assert production["tests"]["ignored"] == 3
assert production["tests"]["passed"] == 1241
assert production["tests"]["suites"] == 85
assert production["root_inputs"] == root_inputs
sealed_witnesses.append(sealed_artifact(production["checks"]))
for gate in production["gates"]:
    sealed_witnesses.append(sealed_artifact(gate["log"]))

OUT.mkdir()
c.write(
    OUT / "reuse-inputs.json",
    {
        "current_source": source,
        "sealed_0813_after_source": c.artifact(SEALED_SOURCE),
        "sealed_0813_after_summary": c.artifact(SEALED_SUMMARY),
        "sealed_witnesses": sorted(set(sealed_witnesses)),
        "sealed_0813_after_source_files_equal": True,
        "probe": probe,
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
    },
)
c.write(
    P / "quality-reuse.json",
    {
        "schema": "litchi.performance.0814.quality-reuse.v1",
        "mode": "reuse-sealed-0813-after-production-gates",
        "reference": c.artifact(SEALED_SUMMARY),
        "source": c.artifact(SEALED_SOURCE),
        "source_files_equal": True,
        "gate_count": 6,
        "gates": production["gates"],
        "tests": production["tests"],
        "inputs": c.artifact(OUT / "reuse-inputs.json"),
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
        "cargo_executed": False,
        "reason": "0814 has no tracked production file delta from the retained 0813 after source census; the commit revision is recorded independently"
    }
)
print("0814 production quality reuse PASS", flush=True)
