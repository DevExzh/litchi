"""Reuse the exact-source sealed 0813 after-production quality gates."""

from pathlib import Path

import custody as c


P = c.P
OLD = P.parent / "change-0813"
OUT = P / "quality-reuse"
assert not OUT.exists() and not (P / "quality-reuse.json").exists()

source = c.source()
prior = c.read(OLD / "build-after/source.json")
assert source["files"] == prior["files"] and len(source["files"]) == 9196
assert source["revision"] == c.read(P / "origin.json")["base"]

seal = c.read(OLD / "seal.json")["files"]


def sealed(path: Path) -> dict[str, object]:
    row = c.artifact(path)
    key = str(path.relative_to(c.ROOT))
    assert seal[key] == row["sha256"], f"sealed 0813 identity changed: {key}"
    return row


summary = c.read(OLD / "quality-summary.json")["production_after"]
assert summary["schema"] == "litchi.performance.0813.quality-after.v1"
assert len(summary["gates"]) == 6
assert {
    key: summary["tests"][key]
    for key in ("passed", "failed", "ignored", "suites")
} == {"passed": 1241, "failed": 0, "ignored": 3, "suites": 85}
assert all(gate["exit_code"] == 0 and gate["status"] == "pass" for gate in summary["gates"])

root_inputs = c.assert_root_inputs()
assert summary["root_inputs"] == root_inputs
checks = sealed(OLD / "quality-after/checks.json")
for gate in summary["gates"]:
    sealed(Path(gate["log"]["path"]))

architecture = c.read(P / "architecture-inputs.json")
assert architecture == c.read(OLD / "architecture-inputs.json")
for name, digest in architecture.items():
    assert c.sha(c.ROOT / name) == digest

unrelated = c.read(P / "origin.json")["unrelated"]
for name, digest in unrelated.items():
    assert c.sha(c.ROOT / name) == digest

probe = {
    str(path.relative_to(P / "probe-src")): c.sha(path)
    for path in (P / "probe-src").rglob("*")
    if path.is_file()
}
sealed_probe = {
    str(path.relative_to(OLD / "probe-src")): c.sha(path)
    for path in (OLD / "probe-src").rglob("*")
    if path.is_file()
}
assert probe == sealed_probe

OUT.mkdir()
c.write(
    OUT / "reuse-inputs.json",
    {
        "current_source": source,
        "sealed_0813_source": sealed(OLD / "build-after/source.json"),
        "sealed_0813_summary": sealed(OLD / "quality-summary.json"),
        "sealed_0813_source_files_equal": True,
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
        "probe": probe,
        "checks": checks,
    },
)
c.write(
    P / "quality-reuse.json",
    {
        "schema": "litchi.performance.0815.quality-reuse.v1",
        "mode": "reuse-sealed-0813-after-production-gates",
        "reference": sealed(OLD / "quality-summary.json"),
        "source": sealed(OLD / "build-after/source.json"),
        "source_files_equal": True,
        "gate_count": 6,
        "gates": summary["gates"],
        "tests": summary["tests"],
        "inputs": c.artifact(OUT / "reuse-inputs.json"),
        "root_inputs": root_inputs,
        "architecture": architecture,
        "unrelated": unrelated,
        "cargo_executed": False,
        "reason": "0815 baseline equals all 9196 source hashes of the sealed 0813 after source",
    },
)
print("0815 baseline quality reuse PASS")
