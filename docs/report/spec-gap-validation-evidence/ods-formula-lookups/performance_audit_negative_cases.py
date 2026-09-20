#!/usr/bin/env python3
"""Exercise retained lookup audit rejection paths without rebuilding or timing."""

from copy import deepcopy
import importlib.util
from pathlib import Path


HERE = Path(__file__).resolve().parent
spec = importlib.util.spec_from_file_location("lookup_audit", HERE / "root_performance_audit.py")
assert spec is not None and spec.loader is not None
audit = importlib.util.module_from_spec(spec)
spec.loader.exec_module(audit)
profile, _ = audit.import_modules()
matrix = audit.matrix_contract(profile)
directory = HERE / "performance/results/candidate-final"
cases = ["scalar-control-arithmetic", "scalar-control-sin", "lookup-cancel-vlookup"]
records = [row for row in audit.rows(directory) if row["case"] in cases
           and row["phase"] == "evaluate" and row["sample_index"] == 1]


def check(rows):
    audit.verify_records(directory, rows, cases, ["evaluate"], 3, 1, matrix)


check(records)
for label, mutate in [
    ("missing row", lambda rows: rows.pop()),
    ("unknown case", lambda rows: rows[0].update(case="unknown")),
    ("wrong sample index", lambda rows: rows[0].update(sample_index=2)),
    ("wrong raw reads", lambda rows: rows[0]["samples"][0].update(reference_reads=99)),
    ("unbalanced allocation", lambda rows: rows[0]["samples"][0].update(live_after=-1)),
    ("cancellation per-repeat reads", lambda rows: next(row for row in rows
        if row["case"] == "lookup-cancel-vlookup").update(reference_reads_p50=4)),
]:
    changed = deepcopy(records)
    mutate(changed)
    try:
        check(changed)
    except audit.AuditError:
        continue
    raise AssertionError(f"accepted invalid retained evidence: {label}")

assert "crates/litchi-ods/src/codec/formula/reference/iri.rs" in audit.baseline_profile_paths(
    profile, audit.BASELINE_COMMIT
)
print("retained performance audit: positive fixture and six rejection cases passed")
