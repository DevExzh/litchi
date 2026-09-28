"""Offline replay of the 0807 alternating native wrapper controls."""

from __future__ import annotations

import argparse
import math
import random
import statistics
from pathlib import Path
from typing import Any

import analysis_common as c


def quantile(values: list[int], q: float) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * q) - 1)]


def bootstrap_interval(values: list[float], seed: int, resamples: int) -> list[float]:
    c.require(values, "cannot bootstrap an empty ratio vector")
    rng = random.Random(seed)
    medians = [statistics.median(rng.choices(values, k=len(values)))
               for _ in range(resamples)]
    medians.sort()
    return [medians[math.floor(0.025 * resamples)],
            medians[math.ceil(0.975 * resamples) - 1]]


def _frozen_inputs() -> None:
    frozen_path = c.PACKET / "build" / "frozen-inputs.json"
    if not frozen_path.is_file():
        return
    frozen = c.read_json(frozen_path)
    c.require(isinstance(frozen, dict), "build frozen-inputs is malformed")
    for name, digest in frozen.items():
        path = c.PACKET / name
        c.require(path.is_file() and c.sha256(path) == digest,
                  f"frozen input changed: {name}")


def _build() -> tuple[dict[str, Any], dict[str, Any]]:
    plan = c.read_json(c.PACKET / "plan.json")
    c.require(plan.get("schema") == "litchi.performance.0807.v1", "plan schema changed")
    c.require(plan.get("cpu") == 12, "native CPU changed")
    native = plan.get("native")
    c.require(isinstance(native, dict), "native plan is missing")
    c.require(native.get("samples") == 30 and native.get("warmup") == 3,
              "native sample policy changed")
    c.require(native.get("shapes") == list(c.SHAPES), "native shape order changed")
    c.require(native.get("orders") == [
        ["control", "profile"], ["profile", "control"],
        ["control", "profile"], ["profile", "control"],
        ["profile", "control"], ["control", "profile"],
    ], "native alternating order changed")
    bootstrap = plan.get("bootstrap")
    c.require(isinstance(bootstrap, dict)
              and bootstrap.get("seed") == c.BOOTSTRAP_SEED
              and bootstrap.get("resamples") == c.BOOTSTRAP_RESAMPLES,
              "bootstrap policy changed")
    build = c.read_json(c.PACKET / "build" / "build.json")
    binaries = build.get("binaries")
    c.require(isinstance(binaries, dict) and set(binaries) == {"control", "profile"},
              "native build binary matrix changed")
    cleanup = c.cleanup_witness()
    for leg in ("control", "profile"):
        c.external_artifact(binaries[leg], f"build {leg}", cleanup)
    all_binaries = dict(binaries)
    fp_receipt_path = c.PACKET / "build-fp" / "receipt.json"
    if fp_receipt_path.is_file():
        fp_receipt = c.read_json(fp_receipt_path)
        if isinstance(fp_receipt.get("binary"), dict):
            all_binaries["profile-fp"] = fp_receipt["binary"]
    c.validate_cleanup(cleanup, all_binaries)
    source_ref = build.get("source")
    c.require(isinstance(source_ref, dict), "build source artifact is missing")
    c.artifact(source_ref, "build source")
    c.source_identity()
    probe = build.get("probe")
    c.require(isinstance(probe, dict) and probe, "probe manifest is missing")
    for name, digest in probe.items():
        path = c.PACKET / name
        c.require(path.is_file() and c.sha256(path) == digest,
                  f"probe source changed: {name}")
    commands = build.get("commands")
    c.require(isinstance(commands, list) and len(commands) == 2,
              "native build command receipt cardinality changed")
    for row in commands:
        c.require(isinstance(row, dict) and row.get("exit_code") == 0,
                  "native build failed")
        c.artifact(row.get("log"), "native build log")
    return plan, build


def _report(path: Path, shape: str, leg: str, plan: dict[str, Any],
            build: dict[str, Any]) -> tuple[dict[str, Any], list[int], dict[str, Any]]:
    value = c.read_json(path)
    identity = c.check_report_identity(
        value, shape, samples=plan["native"]["samples"],
        warmup=plan["native"]["warmup"],
        binary=Path(build["binaries"][leg]["path"]).name,
    )
    return value, identity["values"], identity


def analyze() -> dict[str, Any]:
    plan, build = _build()
    complete = c.read_json(c.PACKET / "native" / "complete.json")
    expected_complete = {
        "processes": 36,
        "plan_sha256": c.sha256(c.PACKET / "plan.json"),
        "build_sha256": c.sha256(c.PACKET / "build" / "build.json"),
    }
    c.require(complete == expected_complete, "native completion receipt changed")
    rows = c.read_json(c.PACKET / "native" / "receipts.json")
    c.require(isinstance(rows, list) and len(rows) == 36,
              "native receipt cardinality changed")
    jobs = [(block, shape, leg)
            for block, order in enumerate(plan["native"]["orders"])
            for shape in c.SHAPES for leg in order]
    processes: list[dict[str, Any]] = []
    lookup: dict[tuple[int, str, str], dict[str, Any]] = {}
    previous_end = float("-inf")
    for row, (block, shape, leg) in zip(rows, jobs):
        label = f"native/{block}-{shape}-{leg}"
        c.require(isinstance(row, dict), f"{label}: receipt is malformed")
        c.require((row.get("block"), row.get("shape"), row.get("leg"))
                  == (block, shape, leg), f"{label}: matrix position changed")
        c.require(row.get("exit_code") == 0, f"{label}: process failed")
        started, ended = row.get("started"), row.get("ended")
        c.require(isinstance(started, (int, float)) and isinstance(ended, (int, float))
                  and previous_end <= started <= ended,
                  f"{label}: receipt timing order changed")
        previous_end = ended
        log = c.artifact(row.get("log"), f"{label} log")
        rss_path = c.artifact(row.get("rss"), f"{label} RSS")
        report_path = c.artifact(row.get("report"), f"{label} report")
        binary_path = build["binaries"][leg]["path"]
        expected_command = [
            "/usr/bin/time", "-f", "%M", "-o", row["rss"]["path"],
            "taskset", "-c", str(plan["cpu"]), binary_path,
            "--mode", "capture", "--shape", shape,
            "--samples", str(plan["native"]["samples"]),
            "--warmup", str(plan["native"]["warmup"]),
            "--output", row["report"]["path"],
        ]
        c.require(c.normalize_command(row.get("command"))
                  == c.normalize_command(expected_command),
                  f"{label}: command changed")
        c.require(not log.read_bytes(), f"{label}: stderr/stdout log is not empty")
        rss_lines = rss_path.read_text(encoding="utf-8").splitlines()
        c.require(len(rss_lines) == 1 and rss_lines[0].isdigit(),
                  f"{label}: RSS gauge is malformed")
        rss = int(rss_lines[0])
        c.require(rss > 0, f"{label}: RSS gauge is zero")
        _, values, identity = _report(report_path, shape, leg, plan, build)
        metrics = {
            "p50_ns": quantile(values, 0.5),
            "mean_ns": statistics.mean(values),
            "p95_ns": quantile(values, 0.95),
            "p99_ns": quantile(values, 0.99),
            "peak_process_rss_kib": rss,
        }
        item = {
            "block": block, "shape": shape, "leg": leg,
            **metrics, "sample_count": len(values),
            "report": str(report_path.relative_to(c.PACKET)),
            "report_sha256": c.sha256(report_path),
            "rss_artifact": str(rss_path.relative_to(c.PACKET)),
            "log_artifact": str(log.relative_to(c.PACKET)),
            "command": row["command"],
            "source_output_semantic_identity_matches": True,
        }
        processes.append(item)
        lookup[(block, shape, leg)] = item

    summaries: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    for shape in c.SHAPES:
        legs: dict[str, Any] = {}
        for leg in ("control", "profile"):
            metrics: dict[str, Any] = {}
            for key in ("p50_ns", "mean_ns", "p95_ns", "p99_ns", "peak_process_rss_kib"):
                values = [lookup[(block, shape, leg)][key] for block in range(6)]
                spread = (max(values) / min(values) - 1.0) * 100.0
                metrics[key] = {
                    "median": statistics.median(values), "values": values,
                    "spread_percent": spread,
                }
                if spread > 5.0:
                    spread_flags.append({"shape": shape, "leg": leg,
                                         "metric": key, "spread_percent": spread})
            legs[leg] = metrics
        ratios = {
            key: [lookup[(block, shape, "profile")][key]
                  / lookup[(block, shape, "control")][key]
                  for block in range(6)]
            for key in ("p50_ns", "mean_ns", "p95_ns", "p99_ns")
        }
        ratio_summary = {}
        for key, values in ratios.items():
            ratio_summary[key] = {
                "values": values,
                "median_ratio": statistics.median(values),
                "bootstrap95": bootstrap_interval(
                    values, plan["bootstrap"]["seed"], plan["bootstrap"]["resamples"]),
            }
        legs["wrapper_profile_over_control"] = ratio_summary
        summaries[shape] = legs
    return {
        "schema": "litchi-0807-native-analysis-v1",
        "packet": "change-0807",
        "plan": {"path": "plan.json", "sha256": c.sha256(c.PACKET / "plan.json"),
                 "schema": plan["schema"], "cpu": plan["cpu"]},
        "build": {"path": "build/build.json",
                   "sha256": c.sha256(c.PACKET / "build" / "build.json"),
                   "binaries": {name: c.external_artifact(value, f"build {name}", c.cleanup_witness())
                                for name, value in build["binaries"].items()}},
        "source": c.source_identity(),
        "fixture_oracle": {shape: c.oracle(shape) for shape in c.SHAPES},
        "processes": processes,
        "summaries": summaries,
        "spread_flags": spread_flags,
        "bootstrap": {"seed": plan["bootstrap"]["seed"],
                       "resamples": plan["bootstrap"]["resamples"],
                       "statistic": "median of six block profile/control ratios",
                       "interval": "percentile 95% bootstrap interval"},
        "verified_measured_samples": 36 * plan["native"]["samples"],
        "scope": "Native capture-only wrapper latency and whole-process /usr/bin/time RSS; wrapper perturbation controls only.",
        "claims": [
            "Profile/control ratios describe this probe's wrapper and code-generation perturbation.",
            "The bootstrap interval is a deterministic block-resampling diagnostic, not a production speedup or population confidence claim.",
            "The output and semantic identities are bound to sealed 0780 and 0785 capture fixtures.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    c.require(args.check ^ args.write, "choose exactly one of --write or --check")
    result = analyze()
    c.write_or_check(c.PACKET / "native-analysis.json", result, args.check)
    print("0807 native analysis PASS", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
