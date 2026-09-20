#!/usr/bin/env python3
"""Run the frozen XLS worksheet-replay checkpoint evidence matrix.

This driver only executes already-built standalone probes and writes raw
captures below this packet's ``captures`` directory.  It never invokes Cargo
and never changes production sources.  The native and repeated-query tasks
have an explicit A/A stage so a baseline can be captured before a candidate
binary exists:

    python3 pilot.py freeze
    python3 pilot.py native --stage aa
    python3 pilot.py native --stage abba

With both binaries available, ``python3 pilot.py all`` runs the complete
matrix in A/A-before-ABBA order.  ``analyze.py`` is the authority for the
semantic and performance gates; this script validates only command shape and
raw report structure while preserving every stdout byte.
"""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import os
import platform
import subprocess
import sys
import time
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_PATH = PACKET / "plan.json"


def read_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write_json(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def rel(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def source_files() -> list[Path]:
    plan = read_json(PLAN_PATH)
    files: list[Path] = []
    for root in plan["source_roots"]:
        directory = ROOT / root
        files.extend(path for path in directory.rglob("*") if path.is_file())
    return sorted(files)


def current_source_hashes() -> dict[str, str]:
    return {rel(path): sha256_file(path) for path in source_files()}


def git_revision(revision: str) -> str:
    return subprocess.check_output(
        ["git", "rev-parse", revision], cwd=ROOT, text=True
    ).strip()


def git_source_hashes(revision: str) -> dict[str, str | None]:
    result: dict[str, str | None] = {}
    for path in source_files():
        name = rel(path)
        probe = subprocess.run(
            ["git", "cat-file", "-e", f"{revision}:{name}"],
            cwd=ROOT,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if probe.returncode:
            result[name] = None
            continue
        result[name] = sha256_bytes(
            subprocess.check_output(["git", "show", f"{revision}:{name}"], cwd=ROOT)
        )
    return result


def probe_hashes() -> dict[str, str]:
    plan = read_json(PLAN_PATH)
    result: dict[str, str] = {}
    for root in plan["probe_roots"]:
        directory = ROOT / root
        for path in sorted(directory.rglob("*")):
            if path.is_file():
                result[rel(path)] = sha256_file(path)
    return result


def tool_hashes() -> dict[str, str]:
    return {
        name: sha256_file(PACKET / name)
        for name in ("pilot.py", "analyze.py", "plan.json")
    }


def corpus_hashes(plan: dict[str, Any]) -> dict[str, dict[str, int | str]]:
    cases = list(plan["cases"])
    cases.extend(plan["budget_fence"]["cases"])
    result: dict[str, dict[str, int | str]] = {}
    for case in cases:
        path = ROOT / case["path"]
        if case["path"] in result:
            continue
        result[case["path"]] = {
            "bytes": path.stat().st_size,
            "sha256": sha256_file(path),
        }
    return result


def binary_path(plan: dict[str, Any], phase: str, kind: str) -> Path:
    return Path(plan["binary_root"]) / phase / plan["binaries"][kind]


def binary_records(plan: dict[str, Any]) -> dict[str, dict[str, dict[str, Any]]]:
    result: dict[str, dict[str, dict[str, Any]]] = {}
    for phase in ("baseline", "candidate"):
        result[phase] = {}
        for kind, name in plan["binaries"].items():
            path = binary_path(plan, phase, kind)
            result[phase][kind] = {
                "name": name,
                "path": str(path),
                "available": path.is_file(),
                "sha256": sha256_file(path) if path.is_file() else None,
                "bytes": path.stat().st_size if path.is_file() else None,
            }
    return result


def build_manifest_records(plan: dict[str, Any]) -> dict[str, Any]:
    records: dict[str, Any] = {}
    for phase in ("baseline", "candidate"):
        packet_path = PACKET / f"{phase}-builds.json"
        binary_root_path = Path(plan["binary_root"]) / phase / "builds.json"
        path = packet_path if packet_path.is_file() else binary_root_path
        if path.is_file():
            records[phase] = {
                "path": str(path),
                "sha256": sha256_file(path),
                "records": read_json(path),
            }
        else:
            records[phase] = None
    return records


def require_freeze(plan: dict[str, Any]) -> Path:
    path = PACKET / plan["capture_root"] / "freeze.json"
    if not path.is_file():
        raise RuntimeError(f"run pilot.py freeze before timing: missing {path}")
    try:
        frozen = read_json(path)
    except (OSError, json.JSONDecodeError) as error:
        raise RuntimeError(f"invalid freeze record {path}: {error}") from error
    if frozen.get("status") != "frozen":
        raise RuntimeError(f"freeze record is not final: {path}")
    if frozen.get("plan_sha256") != sha256_file(PLAN_PATH):
        raise RuntimeError("freeze record was made for a different plan")
    current_tools = tool_hashes()
    if frozen.get("tools_sha256_start") != current_tools:
        raise RuntimeError("tooling changed after freeze")
    if frozen.get("tools_sha256_end") != current_tools:
        raise RuntimeError("freeze tooling end binding is inconsistent")
    current_sources = current_source_hashes()
    if frozen.get("source_sha256_start") != current_sources:
        raise RuntimeError("freeze source start binding is inconsistent")
    if frozen.get("source_sha256_end") != current_sources:
        raise RuntimeError("production source changed after freeze")
    current_probes = probe_hashes()
    if frozen.get("probe_sha256") != current_probes or frozen.get("probe_sha256_end") != current_probes:
        raise RuntimeError("immutable probe changed after freeze")
    current_corpus = corpus_hashes(plan)
    if frozen.get("corpus") != current_corpus or frozen.get("corpus_end") != current_corpus:
        raise RuntimeError("fixture changed after freeze")
    case_manifest_sha = sha256_file(ROOT / plan["case_manifest"])
    if frozen.get("case_manifest_sha256") != case_manifest_sha or frozen.get("case_manifest_sha256_end") != case_manifest_sha:
        raise RuntimeError("case manifest changed after freeze")
    for phase in ("baseline", "candidate"):
        for kind in plan["binaries"]:
            binary = binary_path(plan, phase, kind)
            frozen_binary = frozen.get("binaries", {}).get(phase, {}).get(kind, {})
            if not binary.is_file() or frozen_binary.get("sha256") != sha256_file(binary):
                raise RuntimeError(f"binary changed or missing after freeze: {phase}/{kind}")
    return path


def base_manifest(plan: dict[str, Any], task: str, config: dict[str, Any]) -> dict[str, Any]:
    freeze_path = PACKET / plan["capture_root"] / "freeze.json"
    if task != "freeze":
        freeze = require_freeze(plan)
        freeze_sha256 = sha256_file(freeze_path)
    else:
        freeze_sha256 = None
    revision = plan["baseline_revision"]
    return {
        "schema_version": 1,
        "packet": "change-0725",
        "task": task,
        "status": "running",
        "started_utc": dt.datetime.now(dt.timezone.utc).isoformat(),
        "repository_root": str(ROOT),
        "repository_head": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "baseline_revision": revision,
        "baseline_revision_resolved": git_revision(revision),
        "cpu": plan["cpu"],
        "config": config,
        "plan_sha256": sha256_file(PLAN_PATH),
        "case_manifest_sha256": sha256_file(ROOT / plan["case_manifest"]),
        "freeze_sha256": freeze_sha256,
        "tools_sha256_start": tool_hashes(),
        "baseline_source_sha256": git_source_hashes(revision),
        "source_sha256_start": current_source_hashes(),
        "probe_sha256": probe_hashes(),
        "corpus": corpus_hashes(plan),
        "binaries": binary_records(plan),
        "build_manifests": build_manifest_records(plan),
        "commands": [],
        "raw_sha256": {},
    }


def refresh_binding_end(manifest: dict[str, Any], plan: dict[str, Any]) -> None:
    manifest["source_sha256_end"] = current_source_hashes()
    manifest["tools_sha256_end"] = tool_hashes()
    if manifest.get("tools_sha256_start") != manifest["tools_sha256_end"]:
        raise RuntimeError("tooling changed during capture")
    if manifest.get("source_sha256_start") != manifest["source_sha256_end"]:
        raise RuntimeError("production source changed during capture")
    manifest["probe_sha256_end"] = probe_hashes()
    if manifest.get("probe_sha256") != manifest["probe_sha256_end"]:
        raise RuntimeError("immutable probe changed during capture")
    manifest["corpus_end"] = corpus_hashes(plan)
    if manifest.get("corpus") != manifest["corpus_end"]:
        raise RuntimeError("fixture changed during capture")
    manifest["case_manifest_sha256_end"] = sha256_file(ROOT / plan["case_manifest"])
    if manifest.get("case_manifest_sha256") != manifest["case_manifest_sha256_end"]:
        raise RuntimeError("case manifest changed during capture")
    binary_end = binary_records(plan)
    if manifest.get("binaries") != binary_end:
        raise RuntimeError("binary changed during capture")
    manifest["binaries_end"] = binary_end
    build_end = build_manifest_records(plan)
    if manifest.get("build_manifests") != build_end:
        raise RuntimeError("build manifest changed during capture")
    manifest["build_manifests_end"] = build_end
    manifest["finished_utc"] = dt.datetime.now(dt.timezone.utc).isoformat()


def save_manifest(directory: Path, manifest: dict[str, Any]) -> None:
    write_json(directory / "manifest.json", manifest)


def require_binary(plan: dict[str, Any], phase: str, kind: str) -> Path:
    path = binary_path(plan, phase, kind)
    if not path.is_file():
        raise RuntimeError(f"missing {phase} {kind} probe: {path}")
    return path


def output_path(directory: Path, name: str) -> Path:
    path = directory / name
    if path.exists():
        raise RuntimeError(f"refusing to overwrite existing capture: {path}")
    return path


def run_command(
    command: list[str],
    output: Path,
    directory: Path,
    manifest: dict[str, Any],
    *,
    json_output: bool,
) -> Any:
    path = output_path(directory, output.name)
    started = time.monotonic()
    completed = subprocess.run(command, cwd=ROOT, capture_output=True)
    elapsed = time.monotonic() - started
    stderr_path = Path(str(path) + ".stderr")
    if completed.stderr:
        if stderr_path.exists():
            raise RuntimeError(f"refusing to overwrite existing stderr capture: {stderr_path}")
        stderr_path.write_bytes(completed.stderr)
        manifest["raw_sha256"][rel(stderr_path)] = sha256_file(stderr_path)
    record = {
        "command": command,
        "cwd": str(ROOT),
        "output": rel(path),
        "exit_code": completed.returncode,
        "seconds": elapsed,
    }
    manifest["commands"].append(record)
    if completed.returncode != 0:
        manifest["last_failure"] = {
            "command": command,
            "stderr": completed.stderr.decode("utf-8", errors="replace"),
        }
        raise RuntimeError(
            f"probe failed ({completed.returncode}): {' '.join(command)}\n"
            f"{completed.stderr.decode('utf-8', errors='replace')}"
        )
    path.write_bytes(completed.stdout)
    manifest["raw_sha256"][rel(path)] = sha256_file(path)
    if json_output:
        try:
            return json.loads(completed.stdout.decode("utf-8"))
        except json.JSONDecodeError as error:
            raise RuntimeError(f"probe produced invalid JSON: {path}: {error}") from error
    return completed.stdout.decode("utf-8")


def task_command(binary: Path, plan: dict[str, Any]) -> list[str]:
    return ["taskset", "-c", str(plan["cpu"]), str(binary)]


def case_map(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    return {case["case"]: case for case in plan["cases"]}


def legs_for_stage(stage: str) -> list[tuple[str, str]]:
    if stage == "aa":
        return [("aa1", "baseline"), ("aa2", "baseline")]
    if stage == "abba":
        return [
            ("a1", "baseline"),
            ("b1", "candidate"),
            ("b2", "candidate"),
            ("a2", "baseline"),
        ]
    if stage == "full":
        return legs_for_stage("aa") + legs_for_stage("abba")
    raise ValueError(f"unknown stage {stage}")


def new_or_existing_manifest(
    plan: dict[str, Any],
    task: str,
    config: dict[str, Any],
    directory: Path,
    stage: str,
) -> dict[str, Any]:
    path = directory / "manifest.json"
    if stage == "abba" and path.is_file():
        manifest = read_json(path)
        if manifest.get("task") != task or manifest.get("status") != "aa-complete":
            raise RuntimeError(f"{path} is not an A/A-complete manifest")
        if manifest.get("plan_sha256") != sha256_file(PLAN_PATH):
            raise RuntimeError(f"plan changed after A/A stage: {path}")
        return manifest
    if path.exists():
        raise RuntimeError(f"refusing to replace existing task manifest: {path}")
    directory.mkdir(parents=True, exist_ok=True)
    manifest = base_manifest(plan, task, config)
    save_manifest(directory, manifest)
    return manifest


def finish_task(
    manifest: dict[str, Any], plan: dict[str, Any], directory: Path, status: str
) -> None:
    refresh_binding_end(manifest, plan)
    manifest["status"] = status
    save_manifest(directory, manifest)


def native_command(binary: Path, case: dict[str, Any], mode: str, plan: dict[str, Any]) -> list[str]:
    native = plan["native"]
    return task_command(binary, plan) + [
        "--input",
        case["path"],
        "--budget",
        str(case["budget"]),
        "--mode",
        mode,
        "--worksheet",
        str(case["sheet"]),
        "--row",
        str(case["row"]),
        "--column",
        str(case["column"]),
        "--queries",
        str(native["queries"]),
        "--warmups",
        str(native["warmups"]),
        "--samples",
        str(native["samples"]),
    ]


def validate_native_shape(value: dict[str, Any], plan: dict[str, Any]) -> None:
    native = plan["native"]
    if value.get("schema_version") != 1 or value.get("probe") != "change-0686-xls-index-budget-retry":
        raise RuntimeError("native probe schema identity does not match the frozen probe")
    if value.get("queries") != native["queries"]:
        raise RuntimeError("native probe query count does not match plan")
    if value.get("warmups") != native["warmups"] or value.get("samples") != native["samples"]:
        raise RuntimeError("native probe warmup/sample count does not match plan")
    if value.get("fresh_owner_per_sample") is not True:
        raise RuntimeError("native probe must use a fresh owner per sample")
    records = value.get("records")
    if not isinstance(records, list) or len(records) != native["samples"]:
        raise RuntimeError("native probe returned the wrong record count")
    for record in records:
        queries = record.get("queries")
        if not isinstance(queries, list) or len(queries) != native["queries"]:
            raise RuntimeError("native probe returned the wrong query count")
        if [query.get("ordinal") for query in queries] != list(range(native["queries"])):
            raise RuntimeError("native probe query ordinals are not frozen")


def run_native(plan: dict[str, Any], stage: str) -> None:
    directory = PACKET / plan["capture_root"] / "native"
    directory.mkdir(parents=True, exist_ok=True)
    manifest = new_or_existing_manifest(
        plan,
        "native",
        {
            "groups": plan["native"]["groups"],
            "cases": len(plan["cases"]),
            "modes": ["owned", "file"],
            "legs": plan["native"]["legs"],
        },
        directory,
        stage,
    )
    cases = plan["cases"]
    for leg, phase in legs_for_stage(stage):
        binary = require_binary(plan, phase, "native")
        for case in cases:
            for mode in ("owned", "file"):
                name = f"{leg}-{case['case']}-{mode}.json"
                value = run_command(
                    native_command(binary, case, mode, plan),
                    directory / name,
                    directory,
                    manifest,
                    json_output=True,
                )
                validate_native_shape(value, plan)
                save_manifest(directory, manifest)
        print(f"native {stage} {leg} complete", flush=True)
    finish_task(manifest, plan, directory, "aa-complete" if stage == "aa" else "complete")


def repeat_command(binary: Path, case: dict[str, Any], mode: str, plan: dict[str, Any]) -> list[str]:
    return task_command(binary, plan) + [
        mode,
        case["path"],
        str(case["sheet"]),
        str(case["row"]),
        str(case["column"]),
        str(plan["repeat"]["repetitions"]),
    ]


def validate_repeat_shape(value: str, plan: dict[str, Any]) -> None:
    fields = dict(item.split("=", 1) for item in value.strip().split("\t") if "=" in item)
    required = {"repeats", "found", "nanos"}
    if set(fields) != required:
        raise RuntimeError(f"repeat probe fields differ from frozen CLI: {fields}")
    if int(fields["repeats"]) != plan["repeat"]["repetitions"]:
        raise RuntimeError("repeat probe repetition count does not match plan")
    if int(fields["found"]) < 0 or int(fields["nanos"]) < 0:
        raise RuntimeError("repeat probe emitted a negative counter")


def run_repeat(plan: dict[str, Any], stage: str) -> None:
    directory = PACKET / plan["capture_root"] / "repeat"
    directory.mkdir(parents=True, exist_ok=True)
    manifest = new_or_existing_manifest(
        plan,
        "repeat",
        {
            "groups": plan["repeat"]["groups"],
            "cases": plan["repeat_cases"],
            "modes": ["owned", "file"],
            "samples_per_leg": plan["repeat"]["samples_per_leg"],
            "legs": plan["native"]["legs"],
        },
        directory,
        stage,
    )
    cases = case_map(plan)
    for leg, phase in legs_for_stage(stage):
        binary = require_binary(plan, phase, "repeat")
        for case_name in plan["repeat_cases"]:
            case = cases[case_name]
            for mode in ("owned", "file"):
                for sample in range(plan["repeat"]["samples_per_leg"]):
                    name = f"{case_name}-{mode}-{leg}-{sample}.tsv"
                    value = run_command(
                        repeat_command(binary, case, mode, plan),
                        directory / name,
                        directory,
                        manifest,
                        json_output=False,
                    )
                    validate_repeat_shape(value, plan)
                    save_manifest(directory, manifest)
            print(f"repeat {stage} {leg} {case_name} complete", flush=True)
    finish_task(manifest, plan, directory, "aa-complete" if stage == "aa" else "complete")


def allocator_command(
    binary: Path, case: dict[str, Any], mode: str, operation: str, plan: dict[str, Any]
) -> list[str]:
    return task_command(binary, plan) + [
        mode,
        operation,
        case["path"],
        str(case["sheet"]),
        str(case["row"]),
        str(case["column"]),
        str(case["budget"]),
    ]


def run_allocator(plan: dict[str, Any]) -> None:
    directory = PACKET / plan["capture_root"] / "allocator"
    directory.mkdir(parents=True, exist_ok=True)
    manifest = base_manifest(
        plan,
        "allocator",
        {
            "groups_per_phase": plan["allocator"]["groups_per_phase"],
            "cases": len(plan["cases"]),
            "modes": ["owned", "file"],
            "operations": plan["allocator"]["operations"],
            "repeats": plan["allocator"]["repeats"],
        },
    )
    manifest_path = directory / "manifest.json"
    if manifest_path.exists():
        raise RuntimeError(f"refusing to replace existing task manifest: {manifest_path}")
    directory.mkdir(parents=True, exist_ok=True)
    save_manifest(directory, manifest)
    cases = plan["cases"]
    for phase in ("baseline", "candidate"):
        binary = require_binary(plan, phase, "allocator")
        for case in cases:
            for mode in ("owned", "file"):
                for operation in plan["allocator"]["operations"]:
                    for repeat in range(plan["allocator"]["repeats"]):
                        name = f"{phase}-{case['case']}-{mode}-{operation}-{repeat}.json"
                        value = run_command(
                            allocator_command(binary, case, mode, operation, plan),
                            directory / name,
                            directory,
                            manifest,
                            json_output=True,
                        )
                        if value.get("mode") != mode or value.get("operation") != operation:
                            raise RuntimeError(f"allocator report identity mismatch: {name}")
                        save_manifest(directory, manifest)
            print(f"allocator {phase} {case['case']} complete", flush=True)
    finish_task(manifest, plan, directory, "complete")


def budget_command(
    binary: Path, case: dict[str, Any], budget: int, plan: dict[str, Any], queries: int
) -> list[str]:
    return task_command(binary, plan) + [
        "--input",
        case["path"],
        "--worksheet",
        str(case["sheet"]),
        "--row",
        str(case["row"]),
        "--column",
        str(case["column"]),
        "--budget",
        str(budget),
        "--queries",
        str(queries),
    ]


def run_budget_fence(plan: dict[str, Any]) -> None:
    directory = PACKET / plan["capture_root"] / "budget-fence"
    manifest = base_manifest(
        plan,
        "budget-fence",
        {
            "cases": len(plan["budget_fence"]["cases"]),
            "budgets": sorted(
                {budget for case in plan["budget_fence"]["cases"] for budget in case["budgets"]}
            ),
            "queries": plan["budget_fence"]["queries"],
            "source_mode": "counted-owned",
        },
    )
    manifest_path = directory / "manifest.json"
    if manifest_path.exists():
        raise RuntimeError(f"refusing to replace existing task manifest: {manifest_path}")
    directory.mkdir(parents=True, exist_ok=True)
    save_manifest(directory, manifest)
    for phase in ("baseline", "candidate"):
        binary = require_binary(plan, phase, "budget")
        for case in plan["budget_fence"]["cases"]:
            for budget in case["budgets"]:
                name = f"{phase}-{case['case']}-{budget}.json"
                value = run_command(
                    budget_command(binary, case, budget, plan, plan["budget_fence"]["queries"]),
                    directory / name,
                    directory,
                    manifest,
                    json_output=True,
                )
                if value.get("max_query_index_bytes") != budget:
                    raise RuntimeError(f"budget report identity mismatch: {name}")
                save_manifest(directory, manifest)
            print(f"budget-fence {phase} {case['case']} complete", flush=True)
    finish_task(manifest, plan, directory, "complete")


def run_budget_primary(plan: dict[str, Any]) -> None:
    directory = PACKET / plan["capture_root"] / "budget-primary"
    manifest = base_manifest(
        plan,
        "budget-primary",
        {
            "cases": len(plan["cases"]),
            "queries": plan["budget_primary"]["queries"],
            "source_mode": "counted-owned",
        },
    )
    manifest_path = directory / "manifest.json"
    if manifest_path.exists():
        raise RuntimeError(f"refusing to replace existing task manifest: {manifest_path}")
    directory.mkdir(parents=True, exist_ok=True)
    save_manifest(directory, manifest)
    for phase in ("baseline", "candidate"):
        binary = require_binary(plan, phase, "budget")
        for case in plan["cases"]:
            name = f"{phase}-{case['case']}.json"
            value = run_command(
                budget_command(binary, case, case["budget"], plan, plan["budget_primary"]["queries"]),
                directory / name,
                directory,
                manifest,
                json_output=True,
            )
            if value.get("max_query_index_bytes") != case["budget"]:
                raise RuntimeError(f"budget primary report identity mismatch: {name}")
            save_manifest(directory, manifest)
        print(f"budget-primary {phase} complete", flush=True)
    finish_task(manifest, plan, directory, "complete")


def freeze(plan: dict[str, Any]) -> None:
    directory = PACKET / plan["capture_root"]
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / "freeze.json"
    if path.exists():
        raise RuntimeError(f"refusing to overwrite freeze record: {path}")
    records = binary_records(plan)
    for phase in ("baseline", "candidate"):
        for kind in plan["binaries"]:
            if not records[phase][kind]["available"]:
                raise RuntimeError(f"freeze requires both built binaries: {phase}/{kind}")
    builds = build_manifest_records(plan)
    if any(builds[phase] is None for phase in ("baseline", "candidate")):
        raise RuntimeError("freeze requires baseline and candidate build manifests")
    value = base_manifest(plan, "freeze", {"purpose": "pre-capture source/probe/corpus binding"})
    value["status"] = "frozen"
    refresh_binding_end(value, plan)
    write_json(path, value)
    print(f"wrote {path}")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "task",
        choices=("freeze", "native", "repeat", "allocator", "budget-primary", "budget-fence", "all"),
    )
    parser.add_argument(
        "--stage",
        choices=("aa", "abba", "full"),
        default="full",
        help="native/repeat stage; aa can be captured before candidate build",
    )
    args = parser.parse_args()
    plan = read_json(PLAN_PATH)
    if args.task == "freeze":
        freeze(plan)
    elif args.task == "native":
        run_native(plan, args.stage)
    elif args.task == "repeat":
        run_repeat(plan, args.stage)
    elif args.task == "allocator":
        run_allocator(plan)
    elif args.task == "budget-fence":
        run_budget_fence(plan)
    elif args.task == "budget-primary":
        run_budget_primary(plan)
    else:
        run_native(plan, "full")
        run_repeat(plan, "full")
        run_allocator(plan)
        run_budget_primary(plan)
        run_budget_fence(plan)
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, RuntimeError, subprocess.CalledProcessError) as error:
        print(f"pilot.py: {error}", file=sys.stderr)
        raise SystemExit(1)
