#!/usr/bin/env python3
"""Capture and compare the supplemental eager XLSX rewrite guard.

The guard uses the frozen primary driver's ``run`` function, so source and
binary custody, child serialization, RSS measurement, and the disk-backed
temporary directory stay identical to the primary lane.  Its separate plan
hash is recorded in a binding sidecar because the primary driver is frozen to
the main plan.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
EAGER_PLAN_PATH = HERE / "eager-guard-plan.json"
RUN_PATH = HERE / "run.py"
NUMERIC_PATH = HERE / "analyze.py"


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory guard artifact."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def sha256(path: Path) -> str:
    try:
        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def read_json(path: Path) -> Any:
    try:
        with path.open(encoding="utf-8") as stream:
            return json.load(stream)
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read JSON {path}: {error}") from error


def load_module(path: Path, name: str) -> Any:
    require(path.is_file(), f"missing helper {path}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


RUN = load_module(RUN_PATH, "xlsx_0525_frozen_run")
NUMERIC = load_module(NUMERIC_PATH, "xlsx_0525_numeric_for_eager_guard")
BASE = NUMERIC.BASE


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(HERE))
    except ValueError as error:
        raise EvidenceError(f"path is outside evidence directory: {path}") from error


def plans() -> tuple[dict[str, Any], dict[str, Any]]:
    primary = read_json(PLAN_PATH)
    eager = read_json(EAGER_PLAN_PATH)
    require(isinstance(primary, dict) and isinstance(eager, dict),
            "plan documents are not objects")
    require(eager.get("schema") == "litchi-0525-eager-guard-plan-v1",
            "eager plan schema differs")
    require(eager.get("primary_plan_sha256") == sha256(PLAN_PATH),
            "eager plan is bound to a different primary plan")
    require(eager.get("case") == "xlsx_eager_cell_values_one_percent_edit_save",
            "eager case differs")
    require(eager.get("shapes") == ["medium", "dense-sparse"],
            "eager shape matrix differs")
    require(eager.get("repeats") == 2 and eager.get("warmup") == 10
            and eager.get("samples") == 30 and eager.get("cpu") == primary["cpu"],
            "eager repeat/warmup/sample/CPU counts differ")
    return primary, eager


def eager_report_plan(primary: dict[str, Any], eager: dict[str, Any]) -> dict[str, Any]:
    copy = dict(primary)
    primary_copy = dict(primary["primary"])
    primary_copy.update(case=eager["case"], warmup=eager["warmup"],
                        samples=eager["samples"])
    copy["primary"] = primary_copy
    return copy


def eager_jobs(eager: dict[str, Any]) -> list[dict[str, Any]]:
    return [
        {
            "name": f"eager-r{repeat}-{shape}",
            "repeat": repeat,
            "shape": shape,
            "case": eager["case"],
            "warmup": eager["warmup"],
            "samples": eager["samples"],
        }
        for repeat in range(1, eager["repeats"] + 1)
        for shape in eager["shapes"]
    ]


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def capture(stage: str) -> None:
    primary, eager = plans()
    stage_dir = HERE / stage
    require(stage_dir.is_dir(), f"{stage} source stage is missing")
    binary = RUN.SCRATCH / f"{stage}-normal"
    identity_path = stage_dir / "binary-normal.json"
    identity = read_json(identity_path)
    require(isinstance(identity, dict) and identity.get("sha256") == RUN.sha(binary),
            f"{stage}: normal binary identity differs")
    for job in eager_jobs(eager):
        name = job["name"]
        output_path = stage_dir / f"{name}.json"
        rss_path = stage_dir / f"{name}.rss.json"
        command = [
            "taskset", "-c", str(eager["cpu"]), "/usr/bin/time", "-f",
            '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}',
            "-o", str(rss_path), str(binary), "--warmup", str(job["warmup"]),
            "--samples", str(job["samples"]), "--case", job["case"],
            "--xlsx-cell-crud-shape", job["shape"], "--json", str(output_path),
        ]
        RUN.run(stage, name, command, binary)
        receipt_path = stage_dir / f"{name}.receipt.json"
        binding = {
            "schema": "litchi-0525-eager-guard-binding-v1",
            "stage": stage,
            "name": name,
            "primary_plan_sha256": sha256(PLAN_PATH),
            "eager_plan_sha256": sha256(EAGER_PLAN_PATH),
            "run_script_sha256": sha256(RUN_PATH),
            "guard_script_sha256": sha256(Path(__file__)),
            "source_manifest_sha256": sha256(stage_dir / "source-manifest.json"),
            "binary_sha256": identity["sha256"],
            "receipt_sha256": sha256(receipt_path),
            "command_sha256": hashlib.sha256(
                json.dumps(command, separators=(",", ":")).encode()
            ).hexdigest(),
            "output": relative(output_path),
            "rss": relative(rss_path),
        }
        write_json(stage_dir / f"{name}.binding.json", binding)


def validate_binding(stage_dir: Path, name: str, primary: dict[str, Any],
                     eager: dict[str, Any], binary_sha: str,
                     manifest_sha: str) -> dict[str, Any]:
    path = stage_dir / f"{name}.binding.json"
    binding = read_json(path)
    require(isinstance(binding, dict), f"{relative(path)} is not an object")
    require(binding.get("primary_plan_sha256") == sha256(PLAN_PATH),
            f"{relative(path)} primary plan binding differs")
    require(binding.get("eager_plan_sha256") == sha256(EAGER_PLAN_PATH),
            f"{relative(path)} eager plan binding differs")
    require(binding.get("run_script_sha256") == sha256(RUN_PATH),
            f"{relative(path)} frozen run binding differs")
    require(binding.get("guard_script_sha256") == sha256(Path(__file__)),
            f"{relative(path)} guard script binding differs")
    require(binding.get("source_manifest_sha256") == manifest_sha,
            f"{relative(path)} source binding differs")
    require(binding.get("binary_sha256") == binary_sha,
            f"{relative(path)} binary binding differs")
    receipt_path = stage_dir / f"{name}.receipt.json"
    require(binding.get("receipt_sha256") == sha256(receipt_path),
            f"{relative(path)} receipt binding differs")
    return binding


def validate_guard_receipt(stage_dir: Path, name: str, job: dict[str, Any],
                           primary: dict[str, Any], eager: dict[str, Any],
                           binary: dict[str, Any], manifest_sha: str) -> dict[str, Any]:
    receipt_path = stage_dir / f"{name}.receipt.json"
    receipt = read_json(receipt_path)
    require(receipt.get("exit_code") == 0, f"{relative(receipt_path)} child failed")
    require(receipt.get("binary_sha256") == binary["sha256"],
            f"{relative(receipt_path)} binary binding differs")
    require(receipt.get("source_manifest_sha256") == manifest_sha,
            f"{relative(receipt_path)} source binding differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{relative(receipt_path)} primary plan binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{relative(receipt_path)} frozen run binding differs")
    command = receipt.get("command")
    require(isinstance(command, list), f"{relative(receipt_path)} command is not a list")
    require(command[:3] == ["taskset", "-c", str(eager["cpu"])],
            f"{name} CPU binding differs")
    require("/usr/bin/time" in command and "-o" in command,
            f"{name} RSS observer is missing")
    require(command[command.index("-o") + 1] ==
            str(stage_dir / f"{name}.rss.json"), f"{name} RSS path differs")
    for option, value in (
        ("--warmup", job["warmup"]),
        ("--samples", job["samples"]),
        ("--case", job["case"]),
        ("--xlsx-cell-crud-shape", job["shape"]),
        ("--json", str(stage_dir / f"{name}.json")),
    ):
        require(command.count(option) == 1 and
                command[command.index(option) + 1] == str(value),
                f"{name} command option {option} differs")
    expected = {f"{name}.json", f"{name}.stdout", f"{name}.stderr", f"{name}.rss.json"}
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict) and set(artifacts) == expected,
            f"{name} receipt artifact inventory differs")
    for filename, digest in artifacts.items():
        artifact = stage_dir / filename
        require(artifact.is_file() and not artifact.is_symlink() and
                sha256(artifact) == digest, f"{name} artifact digest differs: {filename}")
    validate_binding(stage_dir, name, primary, eager, binary["sha256"], manifest_sha)
    return receipt


def validate_stage(stage: str, primary: dict[str, Any], eager: dict[str, Any]) -> dict[str, Any]:
    stage_dir = HERE / stage
    manifest_sha = sha256(stage_dir / "source-manifest.json")
    binary_path = stage_dir / "binary-normal.json"
    binary_identity = read_json(binary_path)
    binary = {"sha256": binary_identity["sha256"], "path": binary_identity.get("path")}
    report_plan = eager_report_plan(primary, eager)
    rows = []
    for job in eager_jobs(eager):
        name = job["name"]
        validate_guard_receipt(stage_dir, name, job, primary, eager, binary, manifest_sha)
        report_path = stage_dir / f"{name}.json"
        rss_path = stage_dir / f"{name}.rss.json"
        raw = read_json(report_path)
        parsed = BASE.validate_result(raw, report_plan, {
            "name": name,
            "kind": "primary",
            "guard": None,
            "repeat": job["repeat"],
            "case": job["case"],
            "shape": job["shape"],
            "warmup": job["warmup"],
            "samples": job["samples"],
        }, binary, False)
        parsed["rss"] = {"scope": "whole_child_process", **BASE.validate_rss(rss_path)}
        parsed["report_sha256"] = sha256(report_path)
        rows.append(parsed)
    rows.sort(key=lambda row: (row["repeat"], row["shape"]))
    return {"stage": stage, "rows": rows, "manifest_sha256": manifest_sha,
            "binary_sha256": binary["sha256"],
            "validation": {"expected_matrix": True, "receipt_bindings": True,
                            "report_and_rss_valid": True}}


def change_percent(baseline: float, candidate: float) -> float:
    require(baseline > 0, "baseline metric must be positive")
    return (candidate / baseline - 1.0) * 100.0


def compare(baseline: dict[str, Any], candidate: dict[str, Any], eager: dict[str, Any]) -> dict[str, Any]:
    left = {(row["repeat"], row["shape"]): row for row in baseline["rows"]}
    right = {(row["repeat"], row["shape"]): row for row in candidate["rows"]}
    require(set(left) == set(right), "eager baseline/candidate matrices differ")
    comparisons = []
    adverse = []
    for key in sorted(left):
        before, after = left[key], right[key]
        require(before["identity"] == after["identity"],
                f"eager logical identity differs for {key}")
        metrics = {}
        for phase in ("elapsed_ns", "open_ns", "plan_ns", "commit_ns", "publication_ns"):
            metrics[phase] = {}
            for stat in ("p50", "p95", "p99", "mean"):
                before_value = before["timing"][phase][stat]
                after_value = after["timing"][phase][stat]
                change = change_percent(before_value, after_value)
                metrics[phase][stat] = {"baseline": before_value,
                                        "candidate": after_value,
                                        "change_percent": change,
                                        "adverse_over_five_percent": change > 5.0}
                if change > 5.0:
                    adverse.append({"repeat": key[0], "shape": key[1],
                                    "metric": phase, "stat": stat,
                                    **metrics[phase][stat]})
        before_rss, after_rss = before["rss"]["max_rss_kib"], after["rss"]["max_rss_kib"]
        rss_change = change_percent(before_rss, after_rss)
        metrics["max_rss_kib"] = {"baseline": before_rss, "candidate": after_rss,
                                   "change_percent": rss_change,
                                   "adverse_over_five_percent": rss_change > 5.0}
        if rss_change > 5.0:
            adverse.append({"repeat": key[0], "shape": key[1], "metric": "max_rss_kib",
                            "stat": "peak", **metrics["max_rss_kib"]})
        comparisons.append({"repeat": key[0], "shape": key[1],
                            "identity_equal": True, "metrics": metrics})
    return {
        "rows": comparisons,
        "adverse_flags_over_five_percent": adverse,
        "gate": eager["gate"],
        "all_adverse_metrics_retained": True,
        "no_primary_gain_claim": True,
    }


def analyze(stage: str | None) -> dict[str, Any]:
    primary, eager = plans()
    if stage in ("baseline", "candidate"):
        return {"status": "pass", "stage": stage, "plan_sha256": sha256(EAGER_PLAN_PATH),
                "evidence": validate_stage(stage, primary, eager),
                "scope": eager["purpose"]}
    baseline = validate_stage("baseline", primary, eager)
    candidate = validate_stage("candidate", primary, eager)
    return {"status": "pass", "stage": "compare", "plan_sha256": sha256(EAGER_PLAN_PATH),
            "baseline": baseline, "candidate": candidate,
            "comparison": compare(baseline, candidate, eager),
            "scope": eager["purpose"]}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("stage", choices=("baseline", "candidate", "compare"))
    parser.add_argument("action", choices=("capture", "analyze"), nargs="?", default="analyze")
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        if args.action == "capture":
            capture(args.stage)
            return 0
        stage = None if args.stage == "compare" else args.stage
        output = args.output or HERE / ("eager-guard-comparison.json" if stage is None
                                        else f"{stage}/eager-guard-analysis.json")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_text(json.dumps(analyze(stage), indent=2) + "\n", encoding="utf-8")
    except (EvidenceError, OSError, json.JSONDecodeError, AssertionError) as error:
        print(f"eager_guard.py: error: {error}", file=sys.stderr)
        return 2
    print(f"0525 eager guard verified: {output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
