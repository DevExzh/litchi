"""Independent arithmetic replay over raw process reports; no analysis imports."""
import hashlib
import json
import math
from pathlib import Path
import random
import statistics as st
import sys

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
METRICS = ("p50", "p95", "p99", "mean", "rss_kib")
COUNTERS = ("allocation_calls", "deallocation_calls", "reallocation_calls",
            "allocated_bytes", "deallocated_bytes", "net_live", "peak_above_entry")


def read(path):
    return json.loads(path.read_text())


def checked(descriptor):
    path = Path(descriptor["path"])
    if not path.is_absolute():
        path = ROOT / path
    data = path.read_bytes()
    assert len(data) == descriptor["bytes"]
    assert hashlib.sha256(data).hexdigest() == descriptor["sha256"]
    return data


def interval(values):
    rng = random.Random(823823)
    bootstrap = sorted(st.median([rng.choice(values) for _ in values]) for _ in range(10000))
    return [bootstrap[250], bootstrap[9749]]


def derive():
    native, observed, counts = {}, {}, {}
    for lane, reports, samples in (("qualification-before", 19, 1), ("qualification-after", 19, 1),
                                   ("native", 228, 30), ("allocation", 76, 3)):
        completion = read(P / lane / "complete.json")
        assert completion["status"] == "pass" and completion["reports"] == reports
        rows = json.loads(checked(completion["receipts"]))
        assert len(rows) == reports
        counts[lane] = {"reports": reports, "samples": reports * samples}
        for row in rows:
            assert row["exit_code"] == 0 and row["samples"] == samples
            report = json.loads(checked(row["report"]))
            values = [item["elapsed_ns"] for item in report["samples"]]
            assert len(values) == samples and all(isinstance(x, int) and x > 0 for x in values)
            if row["case"]["probe"] == "real":
                assert values == report["elapsed_ns"]["samples"]
            rss = int(checked(row["rss"]).decode().strip())
            assert rss > 0
            case = row["case"]
            key = (case["probe"], case.get("shape", "real"), case["mode"], row["block"], row["leg"])
            ordered = sorted(values)
            timing = {f"p{q}": ordered[math.ceil(samples * q / 100) - 1] for q in (50, 95, 99)}
            timing.update({"mean": sum(values) / samples, "rss_kib": rss})
            if lane == "native":
                assert key not in native
                native[key] = timing
            if lane == "allocation":
                per_sample = []
                for sample in report["samples"]:
                    a = sample["allocation"]
                    assert a["status"] == "measured" and a["failed_allocation_calls"] == 0
                    assert a["live_bytes_after"] - a["live_bytes_before"] == a["allocated_bytes"] - a["deallocated_bytes"]
                    assert a["region_peak_live_bytes"] >= max(a["live_bytes_before"], a["live_bytes_after"])
                    per_sample.append(a | {"net_live": a["live_bytes_after"] - a["live_bytes_before"],
                                           "peak_above_entry": a["region_peak_live_bytes"] - a["live_bytes_before"]})
                assert key not in observed
                observed[key] = {m: st.median(s[m] for s in per_sample) for m in COUNTERS} | {"rss_kib": rss}
    cases = sorted({key[:3] for key in native})
    assert len(cases) == 19 and len(native) == 228 and len(observed) == 76
    result = []
    for case in cases:
        row = {"case": list(case), "native": {}, "paired": {}, "allocation": {}}
        for leg in ("before", "after"):
            row["native"][leg] = {m: st.median(native[(*case, b, leg)][m] for b in range(6)) for m in METRICS}
            row["allocation"][leg] = {m: st.median(observed[(*case, b, leg)][m] for b in range(2)) for m in COUNTERS}
        for metric in METRICS:
            ratios = [native[(*case, b, "after")][metric] / native[(*case, b, "before")][metric] for b in range(6)]
            row["paired"][metric] = {"ratios": ratios, "median": st.median(ratios), "ci": interval(ratios)}
        result.append(row)
    return {"schema": "litchi.performance.0823.raw-audit.v1", "counts": counts, "rows": result}


def compare(value):
    other = {tuple(row["case"]): row for row in read(P / "analysis.json")["rows"]}
    for row in value["rows"]:
        actual = other[tuple(row["case"])]
        for leg in ("before", "after"):
            for metric in METRICS:
                wanted = actual["native"][leg]["rss_process_median"] if metric == "rss_kib" else actual["native"][leg]["process_median"][metric]
                assert row["native"][leg][metric] == wanted, (row["case"], leg, metric)
            for metric in COUNTERS:
                assert row["allocation"][leg][metric] == actual["allocation"][leg]["allocation_process_median"][metric]
        for metric, paired in row["paired"].items():
            wanted = actual["native_paired"]["metrics"][metric]
            assert paired["median"] == wanted["median_ratio"]
            assert paired["ci"] == [wanted["bootstrap"]["ci_low"], wanted["bootstrap"]["ci_high"]]


if __name__ == "__main__":
    assert sys.argv[1:] in (["--write"], ["--check"])
    value = derive()
    compare(value)
    encoded = json.dumps(value, indent=2, sort_keys=True) + "\n"
    path = P / "raw-audit.json"
    if sys.argv[1] == "--write":
        assert not path.exists()
        path.write_text(encoded)
    else:
        assert path.read_text() == encoded
    print("0823 independent raw audit PASS: 19 rows, 342 reports, 7106 samples")
