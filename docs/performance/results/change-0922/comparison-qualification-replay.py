#!/usr/bin/env python3
"""Replay the retained 0922 functional qualification reports offline.

The input result envelopes must already contain their 48 retained child rows.
This command reads each retained stdout, runs the unchanged 0917 validator, and
compares the matching baseline/candidate reports with the 0922 exact-capacity
comparator. It never starts a probe child and emits only a small receipt.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
RUNNER_PATH = HERE / "comparison.py"
MAX_PACKET_ARTIFACT = 1 << 20


class ReplayError(RuntimeError):
    pass


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReplayError(message)


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def load_runner():
    spec = importlib.util.spec_from_file_location("litchi_perf_0922_qualification_runner",
                                                  RUNNER_PATH)
    require(spec is not None and spec.loader is not None, f"cannot load runner: {RUNNER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def report_from_stdout(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"{label}: missing stdout: {path}")
    for line in reversed([line.strip() for line in path.read_text(encoding="utf-8").splitlines()
                          if line.strip()]):
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise ReplayError(f"{label}: stdout contains no JSON report")


def load_rows(path: Path, runner: Any, version: str) -> tuple[dict[str, Any], dict[str, Any]]:
    value = runner.read_json(path)
    require(isinstance(value, dict) and isinstance(value.get("runs"), list),
            f"{version}: qualification envelope is malformed")
    rows = value["runs"]
    require(len(rows) == 48, f"{version}: expected 48 retained rows, got {len(rows)}")
    reports: dict[str, Any] = {}
    summaries: dict[str, Any] = {}
    for row in rows:
        require(isinstance(row, dict), f"{version}: malformed qualification row")
        label = row.get("label")
        require(isinstance(label, str) and label not in reports,
                f"{version}: duplicate or missing row label")
        require(row.get("exit_code") == 0 and row.get("passed") is True,
                f"{version}/{label}: retained qualification did not pass")
        stdout = row.get("stdout")
        require(isinstance(stdout, dict) and isinstance(stdout.get("path"), str),
                f"{version}/{label}: stdout descriptor missing")
        report = report_from_stdout(Path(str(stdout["path"])), f"{version}/{label}")
        summary = runner.validate_probe_report(report, row, warmups=0, samples=1)
        require(row.get("validation") == summary,
                f"{version}/{label}: stored validation summary changed")
        reports[label] = report
        summaries[label] = summary
    return reports, summaries


def replay(baseline_path: Path, candidate_path: Path) -> dict[str, object]:
    runner = load_runner()
    baseline, baseline_summaries = load_rows(baseline_path, runner, "baseline")
    candidate, candidate_summaries = load_rows(candidate_path, runner, "candidate")
    require(set(baseline) == set(candidate), "baseline/candidate qualification labels differ")
    capacity = []
    for label in sorted(baseline):
        capacity.append(runner.compare_reports(baseline[label], candidate[label], label))
    require(all(bool(row["all_equal"]) for row in capacity),
            "qualification replay found a capacity change")
    result: dict[str, object] = {
        "schema": "litchi.performance.0922.qualification-replay.v1",
        "status": "verified",
        "performance_children_run": 0,
        "baseline_reports": len(baseline),
        "candidate_reports": len(candidate),
        "total_reports": len(baseline) + len(candidate),
        "capacity_pairs": len(capacity),
        "all_exact_capacities": True,
        "candidate_adoption": False,
        "validator": {"path": str(runner.VALIDATOR_PATH),
                       "bytes": runner.VALIDATOR_PATH.stat().st_size,
                       "sha256": sha256(runner.VALIDATOR_PATH)},
        "runner": {"path": str(RUNNER_PATH), "bytes": RUNNER_PATH.stat().st_size,
                   "sha256": sha256(RUNNER_PATH)},
        "baseline_results": {"path": str(baseline_path.resolve()),
                              "bytes": baseline_path.stat().st_size,
                              "sha256": sha256(baseline_path)},
        "candidate_results": {"path": str(candidate_path.resolve()),
                               "bytes": candidate_path.stat().st_size,
                               "sha256": sha256(candidate_path)},
        "reproduction": "Offline stdout replay only; no qualification or performance child is started.",
    }
    del baseline_summaries, candidate_summaries
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-results", type=Path, required=True)
    parser.add_argument("--candidate-results", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    try:
        result = replay(args.baseline_results.resolve(), args.candidate_results.resolve())
        data = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode("utf-8")
        require(len(data) < MAX_PACKET_ARTIFACT, "qualification replay receipt is at least 1 MiB")
        require(not args.output.exists(), f"refusing to replace retained output: {args.output}")
        args.output.parent.mkdir(parents=True, exist_ok=True)
        args.output.write_bytes(data)
        print(json.dumps({"status": result["status"], "receipt": str(args.output.resolve()),
                          "baseline_reports": result["baseline_reports"],
                          "candidate_reports": result["candidate_reports"],
                          "all_exact_capacities": result["all_exact_capacities"]},
                         sort_keys=True))
    except (ReplayError, OSError, KeyError, TypeError, ValueError, RuntimeError) as error:
        print(f"0922 qualification replay failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
