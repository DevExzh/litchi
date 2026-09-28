"""Root-only quality adapter for the 0821 durability packet.

The 0821 source is the byte-identical 0820 allocator-repair source. Running
Cargo again here would create a second, redundant quality result, so this
adapter replays the committed 0820 repair receipt, every result/log
descriptor, its test summary, and the complete 0820 seal. It emits a fresh
0821 receipt bound to the current revision and packet inputs. It never starts
Cargo, a build, a workload, or a benchmark child.
"""

from __future__ import annotations

import time
from pathlib import Path

import custody as c


P = c.P
PLAN = c.read(P / "plan.json")
ORIGIN = c.assert_origin()
assert PLAN["schema"] == "litchi.performance.0821.plan.v1"
assert PLAN["base"] == ORIGIN["base"]
c.plan_cases(PLAN)
assert not (P / "quality.json").exists(), "refusing to overwrite quality.json"
assert not c.TARGET.exists(), f"refusing stale target directory {c.TARGET}"
assert not c.SCRATCH.exists(), f"refusing stale scratch directory {c.SCRATCH}"

attempt = 0
while (P / f"quality-{attempt}").exists():
    attempt += 1
OUT = P / f"quality-{attempt}"
OUT.mkdir()

started = time.time()
# Require the old packet to be intact in the worktree for the first replay.
# Later stability checks replay the seal from the committed base, allowing
# only the new packet's documentation indexes to change.
reuse = c.assert_quality_reuse(require_worktree_seal=True)
production = c.source()
assert production["revision"] == ORIGIN["base"]
c.repair_source_witness()
root_inputs = c.assert_root_inputs()
locks = c.lock_identity()
tool = c.tool_source()
architecture = c.architecture_hashes()
corpus = c.assert_corpus_inputs()
provenance = c.assert_provenance(corpus)
host = c.assert_host()
unrelated = c.assert_unrelated()
packet = c.packet_hashes()
drivers = c.driver_hashes()
frozen = {"production": production, "tool": tool}

source_path = OUT / "source.json"
c.write(source_path, frozen)
frozen_inputs_path = OUT / "frozen-inputs.json"
c.write(
    frozen_inputs_path,
    {
        "schema": "litchi.performance.0821.quality-reuse-inputs.v1",
        "packet": packet,
        "drivers": drivers,
        "root_inputs": root_inputs,
        "locks": locks,
        "architecture": architecture,
        "corpus": corpus,
        "provenance": provenance,
        "host": host,
        "unrelated": unrelated,
        "quality_reuse": reuse,
    },
)

rows = []
for index, item in enumerate(reuse["gates"], 1):
    result_descriptor = item["result"]
    result = c.read(Path(result_descriptor["path"]))
    rows.append(
        {
            "gate": index,
            "name": result["name"],
            "command": result["command"],
            "started": result["started"],
            "ended": result["ended"],
            "exit_code": result["exit_code"],
            "log": result["log"],
            "reused": True,
            "reused_result": result_descriptor,
        }
    )
assert [row["name"] for row in rows] == [
    "fmt", "check", "tests", "clippy", "rustdoc", "boundaries"
]
assert all(row["exit_code"] == 0 and row["reused"] is True for row in rows)

checks_path = OUT / "checks.json"
c.write(
    checks_path,
    {
        "schema": "litchi.performance.0821.quality-reuse-checks.v1",
        "mode": "committed-receipt-replay",
        "cargo_commands_executed": False,
        "rows": rows,
        "focused": reuse["focused"],
        "test_summary": reuse["test_summary"],
        "test_summary_log": reuse["test_summary_log"],
        "source": reuse["source"],
        "inputs": reuse["inputs"],
        "seal": reuse["seal"],
    },
)

summary = {
    "schema": "litchi.performance.0821.quality.v1",
    "status": "pass",
    "mode": "committed-receipt-replay",
    "attempt": attempt,
    "source": c.artifact(source_path),
    "checks": c.artifact(checks_path),
    "frozen_inputs": c.artifact(frozen_inputs_path),
    "rows": rows,
    "gate_count": 6,
    "quality_reuse": reuse,
    "reuse_provenance": PLAN["quality_reuse"],
    "root_inputs": root_inputs,
    "locks": locks,
    "architecture": architecture,
    "corpus": corpus,
    "provenance": provenance,
    "host": host,
    "unrelated": unrelated,
    "packet": packet,
    "drivers": drivers,
    "environment": {
        "mode": "committed-receipt-replay",
        "cargo_commands_executed": False,
        "source_packet": "docs/performance/results/change-0820/repair",
        "base": ORIGIN["base"],
    },
    "started": started,
    "ended": time.time(),
}
c.write(OUT / "quality.json", summary)
c.write(P / "quality.json", summary)
print("0821 quality reuse PASS: six committed repair gates; no Cargo commands", flush=True)
