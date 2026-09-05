#!/usr/bin/env python3
"""Verify retained journals/reports and render the 0420 resource diagnostic."""
import argparse
import hashlib
import json
from pathlib import Path
import re
import statistics
import subprocess
import sys

ROOT = Path(__file__).resolve().parent
VERIFY_SCRIPT = ROOT / "verify.py"
LEGS = ("A1", "B1", "B2", "A2")
ROLES = dict(zip(LEGS, ("control", "candidate", "candidate", "control")))
SELECTORS = ("pptx_cross_copy_media_rich_lifecycle", "pptx_cross_copy_plain_lifecycle")
METRICS = ("allocation_calls", "reallocation_calls", "allocated_bytes", "live_bytes_before", "live_bytes_after", "peak_live_bytes_before", "peak_live_bytes_after")


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load(path):
    return json.loads(path.read_text())


def delta(a, b):
    return (b - a) / a * 100 if a else None


def artifact(entry):
    path = ROOT / entry["path"]
    assert path.resolve().is_relative_to(ROOT.resolve())
    assert path.stat().st_size == entry["bytes"]
    assert sha(path) == entry["sha256"], path


def main():
    global ROOT
    parser = argparse.ArgumentParser()
    parser.add_argument("--replay", action="store_true")
    parser.add_argument("--root", type=Path, default=ROOT)
    args = parser.parse_args()
    ROOT = args.root.resolve()
    manifest = load(ROOT / "capture.json")
    assert manifest["status"] == "pass" and len(manifest["runs"]) == 16
    measurement = load(ROOT / "measurement-protocol.json")
    common = load(ROOT / "protocol.json")
    builds = {role: load(ROOT / f"build-{role}.json") for role in ("control", "candidate")}
    assert builds["control"]["environment"] == builds["candidate"]["environment"]
    assert builds["control"]["argv"] == builds["candidate"]["argv"]
    source_maps = {role: {item["path"]: item["sha256"] for item in build["source_before"]["source_files"]} for role, build in builds.items()}
    changed = sorted(path for path in source_maps["control"].keys() | source_maps["candidate"].keys() if source_maps["control"].get(path) != source_maps["candidate"].get(path))
    assert changed == ["crates/litchi-opc/src/package.rs", "crates/litchi-opc/src/package/payload_reuse_tests.rs", "crates/litchi-pptx/src/opened/cross_copy_plan.rs"], changed
    output = {"change": 420, "classification": "resource diagnostic; no latency claim", "runs": {}, "comparisons": {}}
    for index, run in enumerate(manifest["runs"]):
        mode = ("normal", "allocator")[index // 8]
        leg = LEGS[(index % 8) // 2]
        selector = SELECTORS[index % 2]
        role = ROLES[leg]
        assert (run["mode"], run["leg"], run["selector"], run["role"], run["index"]) == (mode, leg, selector, role, index + 1)
        assert run["status"] == "pass" and run["exit_code"] == 0
        assert run == load(ROOT / "runs" / mode / leg / selector / "journal.json")
        assert run["measurement_protocol_sha256"] == sha(ROOT / "measurement-protocol.json")
        assert run["common_protocol_sha256"] == sha(ROOT / "protocol.json")
        build_path = ROOT / f"build-{role}.json"
        assert run["build_sha256"] == sha(build_path)
        build = load(build_path)
        assert build["status"] == "pass" and build["source_before"] == build["source_after"]
        assert run["source_revision"] == build["source_before"]["revision"]
        assert run["source_after"]["clean"] is True
        assert run["source_after"]["revision"] == run["source_revision"]
        assert run["binary_sha256"] == build["binaries"][mode]["sha256"]
        assert run["contract"] == measurement[mode]
        argv = run["argv"]
        assert argv[:5] == ["taskset", "-c", "2", "/usr/bin/time", "-v"]
        assert argv[7] == build["binaries"][mode]["path"]
        assert argv[8:10] == ["--case", selector]
        assert argv[10:10 + len(common["common_flags"]) ] == common["common_flags"]
        for entry in run["artifacts"].values():
            artifact(entry)
        report_path = ROOT / run["artifacts"]["report"]["path"]
        report = load(report_path)
        assert report["environment"]["git_revision"] == run["source_revision"]
        assert report["environment"]["git_worktree_dirty"] is False
        assert report["binary_identity"]["binary_sha256"] == run["binary_sha256"]
        assert report["binary_identity"]["binary_bytes"] == run["binary_bytes"]
        assert report["binary_identity"]["path"] == run["binary"]["path"]
        row = report["results"][0]
        times = row["elapsed_ns"]["samples"]
        rss_text = (ROOT / run["artifacts"]["time_v"]["path"]).read_text()
        rss = re.findall(r"Maximum resident set size \(kbytes\): (\d+)", rss_text)
        assert len(rss) == 1 and "Exit status: 0" in rss_text
        result = {"report_sha256": sha(report_path), "rss_kib": int(rss[0]), "output_sha256": row["output_sha256"]}
        if mode == "normal":
            result["elapsed_ns"] = {key: row["elapsed_ns"][key] for key in ("mean", "p50", "p95", "p99", "min", "max")}
            result["elapsed_ns"]["sample_standard_deviation"] = statistics.stdev(times)
        else:
            allocation = row["operation_metrics"]["allocation"]
            assert allocation["status"] == "measured"
            result["allocation"] = {}
            for key in METRICS:
                values = allocation[key]["values"]
                assert len(values) == 30
                result["allocation"][key] = {"mean": statistics.mean(values), "min": min(values), "max": max(values)}
        output["runs"][f"{mode}/{leg}/{selector}"] = result

    for mode in ("normal", "allocator"):
        for selector in SELECTORS:
            command = [sys.executable, str(VERIFY_SCRIPT), "--root", str(ROOT), "--mode", mode, "--selector", selector]
            verified = subprocess.run(command, capture_output=True, text=True)
            if verified.returncode:
                raise RuntimeError(verified.stderr + verified.stdout)
            proof = json.loads(verified.stdout)
            assert proof["claim_authorized"] is False
            comparison = {"validation": proof, "pairs": {}, "same_revision_drift_percent": {}}
            def values(leg):
                row = output["runs"][f"{mode}/{leg}/{selector}"]
                values = {"rss_kib": row["rss_kib"]}
                if mode == "normal":
                    values.update({key: row["elapsed_ns"][key] for key in ("mean", "p50", "p95", "p99")})
                else:
                    values.update({key: row["allocation"][key]["mean"] for key in METRICS})
                return values
            for a, b in (("A1", "B1"), ("A2", "B2")):
                aa, bb = values(a), values(b)
                comparison["pairs"][f"{a}_to_{b}"] = {key: {"control": aa[key], "candidate": bb[key], "candidate_minus_control_percent": delta(aa[key], bb[key])} for key in aa}
            for a, b in (("A1", "A2"), ("B1", "B2")):
                aa, bb = values(a), values(b)
                comparison["same_revision_drift_percent"][f"{a}_to_{b}"] = {key: delta(aa[key], bb[key]) for key in aa}
            output["comparisons"][f"{mode}/{selector}"] = comparison
    output["limitations"] = ["100 normal samples are below the release latency claim threshold; allocator elapsed times excluded", "Allocation bytes count full successful realloc requests, not physical copying", "Live/peak values are process snapshots; no operation-local peak claim", "RSS is whole process; shared KVM background uncontrolled; no cold, scaling or native-producer generalization"]
    payload = json.dumps(output, indent=2, sort_keys=True, allow_nan=False) + "\n"
    path = ROOT / "summary.json"
    if args.replay:
        assert path.read_text() == payload, "summary differs from replay"
    else:
        path.write_text(payload)
    lines = ["# Matched resource diagnostic", "", "Normal runs: 100 samples / 10 warmups. Allocator runs: 30 / 3. Each row is one fresh process; allocator elapsed times are excluded. No release latency claim.", "", "## Normal timing and whole-process RSS", "", "| Scenario | Leg | p50 ms | Mean ms | p95 ms | p99 ms | RSS KiB |", "|---|---|---:|---:|---:|---:|---:|"]
    for selector in SELECTORS:
        name = "Media-rich" if "media_rich" in selector else "Plain"
        for leg in LEGS:
            row = output["runs"][f"normal/{leg}/{selector}"]
            ns = row["elapsed_ns"]
            numbers = " | ".join(f"{ns[key] / 1e6:.6f}" for key in ("p50", "mean", "p95", "p99"))
            lines.append(f"| {name} | {leg} | {numbers} | {row['rss_kib']:,} |")
    lines += ["", "## Operation allocation counters", "", "Requested bytes count full realloc requests. Live-after and high-water-after are process snapshots, not operation-local peaks. MB below is decimal; RSS remains KiB.", "", "| Scenario | Leg | Mean allocation calls | Mean requested MB | Live-after MB | High-water-after MB | RSS KiB |", "|---|---|---:|---:|---:|---:|---:|"]
    for selector in SELECTORS:
        name = "Media-rich" if "media_rich" in selector else "Plain"
        for leg in LEGS:
            row = output["runs"][f"allocator/{leg}/{selector}"]
            a = row["allocation"]
            values = " | ".join(f"{a[key]['mean'] / 1e6:.6f}" for key in ("allocated_bytes", "live_bytes_after", "peak_live_bytes_after"))
            lines.append(f"| {name} | {leg} | {a['allocation_calls']['mean']:.3f} | {values} | {row['rss_kib']:,} |")
    lines += ["", "A1/B1 and A2/B2 are the paired comparisons. The raw summary retains both directions, all four same-revision drift checks, and min/max allocation values. Each paired timing/resource direction must be reviewed; no universal speedup is inferred. Shared-host variability and the 100-sample timing scope limit generalization.", ""]
    table = "\n".join(lines)
    table_path = ROOT / "result-table.md"
    if args.replay:
        assert table_path.read_text() == table, "table differs from replay"
    else:
        table_path.write_text(table)
    print(json.dumps({"status": "pass", "runs": 16, "report_groups": 4, "replay": args.replay}))


if __name__ == "__main__":
    main()
