"""Root-only restoration of the baseline after a rejected 0815 gate."""

from pathlib import Path

import custody as c


P = c.P
assert not (P / "restored-source.json").exists()
application = c.read(P / "application.json")
before = c.read(P / "build-before/source.json")
after = application["source"]
assert c.source() == after
assert application["allowlist"] == c.read(P / "plan.json")["source_allowlist"]

decision_path = P / "decision.json"
if decision_path.exists():
    decision = c.read(decision_path)
else:
    analysis = c.read(P / "analysis.json")
    decision = analysis.get("decision_guards", analysis.get("decision", analysis))
eligible = decision.get("adoption_eligible", decision.get("production_adoption"))
assert eligible is False, "restoration requires an explicit rejected adoption decision"

manifest = c.read(P / "candidate/manifest.json")
assert manifest["schema"] == "litchi.performance.0815.candidate-manifest.v1"
assert manifest["base_commit"] == c.read(P / "origin.json")["base"]
allowlist = set(c.read(P / "plan.json")["source_allowlist"])
if "files" in manifest:
    rows = list(manifest["files"].values())
    assert {row["production_path"] for row in rows} == allowlist
else:
    assert manifest["production_path"] in allowlist
    rows = [
        {
            "production_path": manifest["production_path"],
            "before": manifest["before"],
            "after": manifest["after"],
        }
    ]
assert len(rows) == 1
assert c.changed_files(before, after) == allowlist
for row in rows:
    baseline = P / row["before"]["path"]
    assert baseline.stat().st_size == row["before"]["bytes"]
    assert c.sha(baseline) == row["before"]["sha256"]
    assert row["before"]["sha256"] == before["files"][row["production_path"]]
    (c.ROOT / row["production_path"]).write_bytes(baseline.read_bytes())

restored = c.source()
assert restored == before
c.write(P / "restored-source.json", restored)
c.write(
    P / "disposition.json",
    {
        "schema": "litchi.performance.0815.disposition.v1",
        "status": "rejected",
        "production_change_retained": False,
        "reason": decision.get(
            "reason",
            "The candidate did not satisfy the frozen public workflow adoption policy.",
        ),
        "restored_source": c.artifact(P / "restored-source.json"),
    },
)
print("0815 candidate rejected; baseline source restored exactly", flush=True)
