"""Independent raw-report arithmetic and frozen adoption-guard replay."""
import json
import math
from pathlib import Path
import random
import statistics

P = Path(__file__).resolve().parent


def read(path):
    return json.loads(path.read_text())


def audit():
    analysis = read(P / "analysis.json")
    plan = read(P / "plan.json")
    rows = []
    for case in plan["cases"]:
        key = f"{case['shape']}/{case['mode']}"
        ratios = []
        for block in range(6):
            values = {}
            for leg in ("before", "after"):
                report = read(P / "native" / f"{block}-{case['shape']}-{case['mode']}-{leg}.json")
                samples = sorted(sample["elapsed_ns"] for sample in report["samples"])
                assert len(samples) == 30
                values[leg] = samples[14]  # nearest-rank process p50
            ratios.append(values["after"] / values["before"])
        expected = analysis["native"]["analysis"]["paired_by_block_before_after"][key]["metrics"]["p50"]
        assert ratios == [row["ratio"] for row in expected["by_block"]]
        median = statistics.median(ratios)
        rng = random.Random(782078)
        boot = sorted(statistics.median([ratios[rng.randrange(6)] for _ in range(6)])
                      for _ in range(10000))
        low = boot[math.floor((1 - .95) / 2 * 10000)]
        high = boot[min(9999, math.ceil((1 + .95) / 2 * 10000) - 1)]
        assert median == expected["ratio_median"]
        assert low == expected["bootstrap"]["ci_low"]
        assert high == expected["bootstrap"]["ci_high"]
        rows.append({"case": key, "median_ratio": median,
                     "ci_low": low, "ci_high": high,
                     "all_blocks_regress": all(ratio > 1 for ratio in ratios),
                     "persistent_regression": median > 1.05 and low > 1})
    resource_violations = []
    for key, pair in analysis["allocation"]["analysis"]["paired_by_block_before_after"].items():
        for metric in ("net_live", "peak_above_entry"):
            for row in pair["metrics"][metric]["by_block"]:
                if row["after"] > row["before"]:
                    resource_violations.append({"case": key, "metric": metric, "block": row["block"]})
    failures = [row["case"] for row in rows if row["persistent_regression"]]
    assert failures == ["many/write", "many/lifecycle", "unicode/write", "unicode/lifecycle"]
    assert not resource_violations
    assert analysis["disposition"]["status"] == "rejected"
    return {"schema": "litchi.performance.0782.decision-audit.v1", "rows": rows,
            "latency_guard_failures": failures, "resource_guard_failures": resource_violations,
            "decision": "reject", "scope": "Within-run raw p50 arithmetic and frozen guard; no causal or historical timing attribution."}


if __name__ == "__main__":
    result = audit()
    assert result == read(P / "decision-audit.json")
    print("Raw p50 and frozen adoption guard replay passed: four latency violations.")
