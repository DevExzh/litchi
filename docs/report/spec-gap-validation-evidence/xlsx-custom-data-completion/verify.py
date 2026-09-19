#!/usr/bin/env python3
"""Independently verify retained gate evidence and selected source hashes."""
import hashlib
import json
from pathlib import Path
import re
import statistics

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[3]


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def verify_gates():
    gates = ROOT / "gates"
    before = json.loads((gates / "source-before.json").read_text())
    after = json.loads((gates / "source-after.json").read_text())
    assert before == after, "gate sources changed"
    results = json.loads((gates / "results.json").read_text())
    verification = json.loads((gates / "verification.json").read_text())
    assert verification["stable_sources"] and verification["all_required_checks_passed"]
    counts = {}
    for result in results:
        log = gates / (result["name"] + ".log")
        assert sha(log) == result["log_sha256"], log
        if result["exit_code"]:
            assert result["name"] == "format" and result["exit_code"] == 1
            assert verification["known_unchanged_baseline_format_failure"]
            paths = re.findall(r"^Diff in (.*?):\d+:", log.read_text(), re.M)
            assert paths and all(p.endswith("/crates/litchi-xlsx/tests/drawing_svg_read.rs") for p in paths)
        if result["name"].endswith("-tests"):
            matches = re.findall(r"test result: ok\. (\d+) passed; (\d+) failed; (\d+) ignored", log.read_text())
            assert matches, log
            counts[result["name"]] = {
                "passed": sum(int(row[0]) for row in matches),
                "failed": sum(int(row[1]) for row in matches),
                "ignored": sum(int(row[2]) for row in matches),
            }
            assert counts[result["name"]]["failed"] == 0
    for relative in json.loads((gates / "batch-files.json").read_text()):
        assert sha(REPO / relative) == before[relative], relative
    assert sha(gates / "Cargo.lock") == before["Cargo.lock"]
    return counts


def verify_performance():
    captures = {}
    for name in ("baseline-1d39eea51", "candidate-final"):
        directory = ROOT / "performance/results" / name
        if not (directory / "provenance.json").is_file():
            continue
        provenance = json.loads((directory / "provenance.json").read_text())
        for line in (directory / "retained-files.sha256").read_text().splitlines():
            digest, relative = line.split("  ", 1)
            assert sha(directory / relative) == digest, (directory, relative)
        kind = provenance["capture_kind"]
        assert kind in ("baseline", "candidate") and kind not in captures
        assert provenance["samples"] == 15 and provenance["warmups"] == 3
        assert (directory / "source-manifest-before.sha256").read_bytes() == (directory / "source-manifest-after.sha256").read_bytes()
        for line in (directory / "profile-manifest.sha256").read_text().splitlines():
            digest, relative = line.split("  ", 1)
            path = ROOT / "performance" / relative
            assert sha(path) == digest, path
        groups = {}
        for raw in sorted(directory.glob("*.jsonl")):
            rows = [json.loads(line) for line in raw.read_text().splitlines()]
            assert len(rows) == 15 and {r["sample"] for r in rows} == set(range(1, 16))
            level, lane = rows[0]["level"], rows[0]["lane"]
            assert raw.stem == level + "-" + lane
            rss = []
            for row in rows:
                assert row["level"] == level and row["lane"] == lane and row["warmups"] == 3
                allocator, validation = row["allocator"], row["validation"]
                assert allocator["balanced"] and allocator["live_before"] == allocator["live_after"]
                assert validation["payloads_exact"] and validation["bindings_exact"]
                assert validation["source_exact_noop"] == (lane in ("read", "noop-commit-save"))
                assert validation["inverse_exact"] == (lane in ("read", "noop-commit-save", "remove-inverse"))
                time_log = directory / (raw.stem + "-sample" + str(row["sample"]) + ".time.txt")
                rss.append(int(re.search(r"Maximum resident set size \(kbytes\): (\d+)", time_log.read_text())[1]))
                assert "Exit status: 0" in time_log.read_text()
            assert len({json.dumps(r["fixture"], sort_keys=True) for r in rows}) == 1
            groups[raw.stem] = {
                "fixture": rows[0]["fixture"],
                "candidate_hashes": sorted({r["validation"]["candidate_sha256"] for r in rows}),
                "elapsed_ns": statistics.median(r["elapsed_ns"] for r in rows),
                "requested_bytes": statistics.median(r["allocator"]["requested_bytes"] for r in rows),
                "allocation_calls": statistics.median(r["allocator"]["allocation_calls"] for r in rows),
                "peak_live_delta": statistics.median(r["allocator"]["peak_live_delta"] for r in rows),
                "rss_kib": statistics.median(rss),
            }
        assert len(groups) == 15
        captures[kind] = {"provenance": provenance, "directory": directory, "groups": groups}
    assert set(captures) == {"baseline", "candidate"}
    baseline, candidate = captures["baseline"], captures["candidate"]
    for key in ("source_cargo_lock_sha256", "harness_cargo_lock_sha256", "rustflags", "rustdocflags", "cargo_incremental"):
        assert baseline["provenance"][key] == candidate["provenance"][key], key
    for filename in ("profile-manifest.sha256", "rustc-vv.txt", "cargo-version.txt"):
        assert (baseline["directory"] / filename).read_bytes() == (candidate["directory"] / filename).read_bytes(), filename
    comparisons = {}
    for group, before in baseline["groups"].items():
        after = candidate["groups"][group]
        assert before["fixture"] == after["fixture"], group
        assert before["candidate_hashes"] == after["candidate_hashes"], group
        comparisons[group] = {key: {"baseline": before[key], "candidate": after[key], "ratio": after[key] / before[key]}
            for key in ("elapsed_ns", "requested_bytes", "allocation_calls", "peak_live_delta", "rss_kib")}
    reported = json.loads((ROOT / "performance/results/matched-statistics.json").read_text())
    assert len(reported["groups"]) == len(comparisons)
    for group in reported["groups"]:
        actual = comparisons[group["level"] + "-" + group["lane"]]
        for metric, values in group["metrics"].items():
            scale = 1_000_000 if metric == "elapsed_ns" else 1
            assert values["baseline_median"] == actual[metric]["baseline"] / scale
            assert values["candidate_median"] == actual[metric]["candidate"] / scale
            assert abs(values["ratio"] - actual[metric]["ratio"]) < 1e-12
    diagnostic = ROOT / "performance/results/candidate-pre-reserve-fix"
    for line in (diagnostic / "retained-files.sha256").read_text().splitlines():
        digest, relative = line.split("  ", 1)
        assert sha(diagnostic / relative) == digest, relative
    return {"samples": 450, "reported_medians_verified": True,
            "diagnostic_hashes_verified": True, "comparisons": comparisons}


if __name__ == "__main__":
    print(json.dumps({"gates": verify_gates(), "performance": verify_performance()}, indent=2))
