"""Replay the disjoint readback residual in the committed 0962 summary."""
import hashlib
import json
from pathlib import Path


def reduce(row):
    selected = row["selected_qualified"]
    rewritten = row["boundary_partition"]["rewritten_value_source"]
    validator = row["validate_xml"]["nested_under_rewritten"]
    result = {}
    for metric in ("samples", "period"):
        assert 0 <= validator[metric] <= rewritten[metric] <= selected[metric]
        result[metric] = rewritten[metric] - validator[metric]
    result["selected_period_fraction"] = result["period"] / selected["period"]
    return result


if __name__ == "__main__":
    source = Path(__file__).resolve().parents[1] / "change-0962/fp-wave-v1/boundary-attribution-summary-v1.json"
    raw = source.read_bytes()
    assert hashlib.sha256(raw).hexdigest() == "14fa5e02723da074f0cf06c494f2ff48e1389bb489e2f189fb09ad6ce08ea477"
    data = json.loads(raw)
    assert data["status"] == "passed"
    assert data["validate_xml"]["nested_subset_pass"]
    children = [{"index": row["index"], **reduce(row)} for row in data["children"]]
    pooled = reduce(data)
    for metric in ("samples", "period"):
        assert sum(row[metric] for row in children) == pooled[metric]
    print(json.dumps({
        "schema": "litchi.performance.0963.readback-residual.v1",
        "input_sha256": hashlib.sha256(raw).hexdigest(),
        "input_bytes": len(raw),
        "scope": "0962 diagnostic-build qualified stage ancestry; no fresh workload",
        "interpretation": "Rewritten readback excluding nested validator; includes retained merge and other work. Not removable cost, elapsed time, or a speedup bound.",
        "children": children,
        "pooled": pooled,
    }, indent=2, sort_keys=True))
