#!/usr/bin/env python3
"""Validate and attribute the baseline-only 0547 Callgrind profiles.

The four jobs each contain five positive timed constructor dumps.  CFB corpus
construction opens the generated compound file once before the timed loop, so
its first numbered dump is retained as setup evidence and excluded from the
timed totals.  The owner edge and its positive benchmark-runner/setup ancestry
classify every dump; dump ordinal alone is never used for that decision.

This report is an attribution diagnostic.  Callgrind Ir and call metadata do
not establish native latency, operation-local call counts, allocation counts,
RSS, or a speedup claim.
"""

from __future__ import annotations

import argparse
import datetime
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
PLAN_PATH = HERE / "plan.json"
BASELINE = HERE / "baseline"
RUN_PATH = HERE / "run.py"
OLD_PROFILE_PATH = HERE.parent / "change-0536" / "analyze_profiles.py"


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence item."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def load_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def sha256(path: Path) -> str:
    try:
        import hashlib

        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def relative(path: Path) -> str:
    try:
        return path.relative_to(HERE).as_posix()
    except ValueError:
        return str(path)


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


# 0536's parser and attribution routines are retained immutable evidence.  We
# use them as a library only; the globals that describe the old campaign are
# replaced with this campaign's paths before any routine is called.
LEGACY = load_module(OLD_PROFILE_PATH, "litchi_0547_retained_profile_parser")
LEGACY.HERE = HERE
LEGACY.PLAN_PATH = PLAN_PATH
LEGACY.RUN_PATH = RUN_PATH
LEGACY.STAGES = ("baseline",)
LEGACY.FOLDER = BASELINE

PROFILE_OWNER_XLS = (
    "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits"
)
PROFILE_OWNER_CFB = "litchi_cfb::file::OleFile<R>::open"
XLS_RUNNER = "litchi_perf_baseline::run_xls_owned_source_case"
CFB_RUNNER = "litchi_perf_baseline::run_cfb_open"
CFB_SETUP_CALLERS = (
    "litchi_perf_baseline::build_cfb_corpus",
    "litchi_perf_baseline::run::{{closure}}",
)
CHAIN_COLLECTOR = "litchi_cfb::file::SectorChainScratch::collect_exact"
CLAIM_SECTOR = "litchi_cfb::file::OleFile<R>::claim_sector"
VALIDATE_STREAM_ALLOCATIONS = (
    "litchi_cfb::file::OleFile<R>::validate_stream_allocations"
)
PHYSICAL_RECONCILIATION = (
    "litchi_cfb::file::OleFile<R>::validate_physical_sector_layout"
)
LOAD_FAT = "litchi_cfb::file::OleFile<R>::load_fat"
CHAIN_ERROR = "chain_error"
ATTRIBUTION_TARGETS = {
    "chain_collector": CHAIN_COLLECTOR,
    "claim_sector": CLAIM_SECTOR,
    "validate_stream_allocations": VALIDATE_STREAM_ALLOCATIONS,
    "physical_reconciliation": PHYSICAL_RECONCILIATION,
    "load_fat": LOAD_FAT,
    "chain_error": CHAIN_ERROR,
}


def plan_data() -> dict[str, Any]:
    plan = load_json(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("scope") ==
            "Fresh baseline-only OLE2 collector sub-operation attribution",
            "plan scope differs from the 0547 attribution pass")
    require(plan.get("status") == "frozen-before-build-and-capture",
            "plan is not frozen before build and capture")
    require(plan.get("cpu") == 2, "plan CPU differs from pinned CPU")
    profile = plan.get("profile")
    require(isinstance(profile, dict), "profile plan is missing")
    require(profile.get("repeats") == 1 and profile.get("warmup") == 0
            and profile.get("samples") == 5,
            "profile plan must be one repeat, zero warmup, five samples")
    require(profile.get("jobs") == [
        "xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large",
    ], "profile job matrix differs")
    require(profile.get("xls_owner") == PROFILE_OWNER_XLS,
            "XLS profile owner differs")
    require(profile.get("cfb_owner") == PROFILE_OWNER_CFB,
            "CFB profile owner differs")
    require(profile.get("instruction_flags") == [
        "--dump-instr=yes", "--dump-line=no", "--compress-pos=no",
        "--collect-jumps=yes",
    ], "Callgrind instruction flags differ")
    groups = plan.get("groups")
    require(isinstance(groups, dict), "plan groups are missing")
    require(groups.get("xls", {}).get("cases") and
            "xls_owned_source_open_one_cell" in groups["xls"]["cases"],
            "XLS group does not retain the profiled case")
    cfb = groups.get("cfb", {})
    require(cfb.get("cases") == ["cfb_open"], "CFB group case differs")
    require(cfb.get("shapes") == ["tiny", "many-small", "few-large"],
            "CFB group shape matrix differs")
    require(cfb.get("payload") == "incompressible",
            "CFB group payload differs")
    assembly = plan.get("assembly")
    require(isinstance(assembly, dict), "assembly plan is missing")
    require(assembly.get("owners") == [
        "SectorChainScratch", "CheckedBitSet", "chain_error",
    ], "assembly owner fragments differ")
    require(assembly.get("required") == ["collect_exact", "insert"],
            "assembly required fragments differ")
    owned = plan.get("owned_paths")
    require(owned == ["/home/zhuhe/litchi-goal-0547-target"],
            "0547 owned target differs")
    return plan


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    jobs: list[dict[str, Any]] = [{
        "name": "profile-r1-xls-owned",
        "repeat": 1,
        "group": "xls-owned",
        "kind": "xls",
        "shape": None,
        "owner": profile["xls_owner"],
        "runner": XLS_RUNNER,
        "setup_callers": [],
        "selection": {"cases": ["xls_owned_source_open_one_cell"]},
        "samples": profile["samples"],
        "warmup": profile["warmup"],
    }]
    for shape in plan["groups"]["cfb"]["shapes"]:
        jobs.append({
            "name": f"profile-r1-cfb-{shape}",
            "repeat": 1,
            "group": f"cfb-{shape}",
            "kind": "cfb",
            "shape": shape,
            "owner": profile["cfb_owner"],
            "runner": CFB_RUNNER,
            "setup_callers": list(CFB_SETUP_CALLERS),
            "selection": {
                "cases": ["cfb_open"],
                "shapes": [shape],
                "payload": "incompressible",
            },
            "samples": profile["samples"],
            "warmup": profile["warmup"],
        })
    return jobs


def _option(command: list[str], name: str) -> str:
    values = []
    for index, item in enumerate(command):
        if item == name:
            require(index + 1 < len(command),
                    f"command omits value for {name}")
            values.append(command[index + 1])
        elif item.startswith(name + "="):
            values.append(item.split("=", 1)[1])
    require(len(values) == 1, f"command omits or repeats {name}")
    return values[0]


def _validate_host(path: Path, label: str) -> None:
    value = load_json(path)
    require(isinstance(value, dict), f"{label}: host record is not an object")
    try:
        timestamp = datetime.datetime.fromisoformat(value["observed_utc"])
    except (KeyError, TypeError, ValueError) as error:
        raise EvidenceError(f"{label}: host timestamp is malformed") from error
    require(timestamp.tzinfo is not None, f"{label}: host timestamp has no timezone")
    processes = value.get("compiler_processes")
    require(isinstance(processes, list), f"{label}: compiler observation is not a list")
    for process in processes:
        require(isinstance(process, dict), f"{label}: compiler observation is malformed")
        require(isinstance(process.get("pid"), int) and process["pid"] > 0,
                f"{label}: compiler pid is invalid")
        require(process.get("comm") in {"cargo", "rustc"},
                f"{label}: compiler command is unexpected")
        require(isinstance(process.get("cwd"), str) and process["cwd"],
                f"{label}: compiler cwd is missing")
    require(value.get("scope") ==
            "Accessible compiler processes; no host quiescence guarantee",
            f"{label}: host scope differs")


def _validate_receipt(job: dict[str, Any], metadata: dict[str, Any]) -> dict[str, Any]:
    name = job["name"]
    path = BASELINE / f"{name}.receipt.json"
    receipt = load_json(path)
    require(receipt.get("exit_code") == 0, f"{name}: child exited unsuccessfully")
    require(receipt.get("execution_stage") == "baseline",
            f"{name}: execution stage differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            f"{name}: plan binding differs")
    source_sha = sha256(BASELINE / "source-manifest.json")
    require(receipt.get("source_manifest_sha256") == source_sha,
            f"{name}: source manifest binding differs")
    require(receipt.get("binary_sha256") == metadata["sha256"],
            f"{name}: binary binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            f"{name}: run-driver binding differs")
    command = receipt.get("command")
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            f"{name}: command is not an argv list")
    require(command[:3] == ["taskset", "-c", str(2)],
            f"{name}: command is not pinned to plan CPU")
    require(command[3:7] == ["valgrind", "--vgdb=no", "--tool=callgrind",
                              "--collect-atstart=no"],
            f"{name}: Callgrind mode differs")
    require(command.count("--vgdb=no") == 1, f"{name}: debugger argument repeats")
    require(f"--toggle-collect={job['owner']}" in command
            and f"--zero-before={job['owner']}" in command
            and f"--dump-after={job['owner']}" in command,
            f"{name}: exact owner controls are incomplete")
    require(_option(command, "--callgrind-out-file") ==
            str(BASELINE / f"{name}.callgrind"),
            f"{name}: Callgrind output path differs")
    require(_option(command, "--case") == ",".join(job["selection"]["cases"]),
            f"{name}: case option differs")
    require(_option(command, "--warmup") == str(job["warmup"]),
            f"{name}: warmup option differs")
    require(_option(command, "--samples") == str(job["samples"]),
            f"{name}: sample option differs")
    require(_option(command, "--json") == str(BASELINE / f"{name}.json"),
            f"{name}: report path differs")
    require(_option(command, "--corpus-manifest") ==
            str(BASELINE / f"{name}.catalog.json"),
            f"{name}: catalog path differs")
    if job["shape"] is None:
        require("--shape" not in command and "--payload" not in command,
                f"{name}: XLS command has CFB selectors")
    else:
        require(_option(command, "--shape") == job["shape"],
                f"{name}: shape option differs")
        require(_option(command, "--payload") == "incompressible",
                f"{name}: payload option differs")
    artifacts = receipt.get("artifacts")
    require(isinstance(artifacts, dict), f"{name}: receipt artifacts are missing")
    last_part = 5 if job["kind"] == "xls" else 6
    expected = {
        f"{name}.host.json", f"{name}.json", f"{name}.catalog.json",
        f"{name}.stdout", f"{name}.stderr", f"{name}.callgrind",
        *(f"{name}.callgrind.{number}" for number in range(1, last_part + 1)),
    }
    require(set(artifacts) == expected,
            f"{name}: receipt artifact set differs")
    for artifact, digest in artifacts.items():
        require(isinstance(digest, str) and len(digest) == 64,
                f"{name}: artifact hash is malformed: {artifact}")
        artifact_path = BASELINE / artifact
        require(artifact_path.is_file() and not artifact_path.is_symlink(),
                f"{name}: artifact is missing: {artifact}")
        require(sha256(artifact_path) == digest,
                f"{name}: artifact hash differs: {artifact}")
        if artifact.endswith(".host.json"):
            _validate_host(artifact_path, f"{name}/{artifact}")
    return receipt


def _validate_build(plan: dict[str, Any]) -> dict[str, Any]:
    source_manifest = BASELINE / "source-manifest.json"
    binary_meta_path = BASELINE / "binary-normal.json"
    receipt_path = BASELINE / "build-normal.receipt.json"
    require(source_manifest.is_file(), "baseline source manifest is missing")
    require(binary_meta_path.is_file(), "normal binary metadata is missing")
    require(receipt_path.is_file(), "normal build receipt is missing")
    source_sha = sha256(source_manifest)
    receipt = load_json(receipt_path)
    require(receipt.get("exit_code") == 0, "normal build failed")
    require(receipt.get("execution_stage") == "baseline",
            "normal build execution stage differs")
    require(receipt.get("plan_sha256") == sha256(PLAN_PATH),
            "normal build plan binding differs")
    require(receipt.get("source_manifest_sha256") == source_sha,
            "normal build source binding differs")
    require(receipt.get("script_sha256") == sha256(RUN_PATH),
            "normal build driver binding differs")
    command = receipt.get("command")
    require(isinstance(command, list), "normal build command is not an argv list")
    require("--release" in command and "--locked" in command,
            "normal build is not the locked release build")
    require("--target-dir" in command and
            command[command.index("--target-dir") + 1] == plan["owned_paths"][0],
            "normal build target directory differs")
    metadata = load_json(binary_meta_path)
    require(metadata.get("source_manifest_sha256") == source_sha,
            "normal binary source binding differs")
    require(metadata.get("build_receipt_sha256") == sha256(receipt_path),
            "normal binary build binding differs")
    binary_path = Path(metadata.get("path", ""))
    require(binary_path.is_file() and not binary_path.is_symlink(),
            "normal binary path is missing")
    require(metadata.get("sha256") == sha256(binary_path),
            "normal binary hash differs")
    require(metadata.get("bytes") == binary_path.stat().st_size,
            "normal binary size differs")
    return {
        "path": relative(binary_path),
        "absolute_path": str(binary_path),
        "sha256": metadata["sha256"],
        "bytes": metadata["bytes"],
        "metadata": relative(binary_meta_path),
        "metadata_sha256": sha256(binary_meta_path),
        "build_receipt": relative(receipt_path),
        "build_receipt_sha256": sha256(receipt_path),
        "source_manifest": relative(source_manifest),
        "source_manifest_sha256": source_sha,
    }


def _identity(row: dict[str, Any], label: str) -> dict[str, Any]:
    return {
        "case": row.get("case"),
        "corpus": row.get("corpus"),
        "sink": row.get("sink"),
        "source": LEGACY.normalize(row.get("source"), label + ".source")
        if row.get("source") is not None else None,
        "output_sha256": row.get("output_sha256"),
    }


def _validate_report(job: dict[str, Any], metadata: dict[str, Any]) -> dict[str, Any]:
    path = BASELINE / f"{job['name']}.json"
    require(path.is_file(), f"{job['name']}: benchmark report is missing")
    # The campaign-local numerical adapter performs the retained corpus,
    # elapsed-vector, operation-metric, and XLS source-contract checks.
    numeric = load_module(HERE / "analyze.py", "litchi_0547_numeric")
    rows = numeric.validate_report(path, job, "normal")
    require(len(rows) == 1, f"{job['name']}: expected one result row")
    row = rows[0]
    require(row.get("case") == job["selection"]["cases"][0],
            f"{job['name']}: result case differs")
    if job["shape"] is not None:
        require(row.get("corpus", {}).get("shape") == job["shape"],
                f"{job['name']}: result shape differs")
    report = load_json(path)
    require(report.get("binary_identity", {}).get("binary_sha256") ==
            metadata["sha256"], f"{job['name']}: report binary differs")
    return {
        "path": relative(path),
        "sha256": sha256(path),
        "identity": _identity(row, f"{job['name']}.row"),
        "row": row,
    }


def _attribution(parsed: dict[str, Any], dump_label: str) -> dict[str, Any]:
    functions = {
        label: LEGACY.attribution_for(
            parsed, target, dump_label,
            allow_inlined=(label in {"claim_sector", "chain_error"}),
        )
        for label, target in ATTRIBUTION_TARGETS.items()
    }
    require(functions["chain_collector"]["out_of_line"],
            f"{dump_label}: collector has no positive out-of-line edge")
    require(
        functions["physical_reconciliation"]["out_of_line"]
        and functions["physical_reconciliation"]["incoming_edge_count"] > 0,
        f"{dump_label}: physical reconciliation has no positive edge",
    )
    ranking = LEGACY.rank_exclusive_sectors(parsed, functions, dump_label)
    return {"functions": functions, "exclusive_sector_ranking": ranking}


def _profile_job(job: dict[str, Any], metadata: dict[str, Any],
                 plan: dict[str, Any]) -> dict[str, Any]:
    name = job["name"]
    receipt = _validate_receipt(job, metadata)
    report = _validate_report(job, metadata)
    numbered = []
    last_part = 5 if job["kind"] == "xls" else 6
    for number in range(1, last_part + 1):
        path = BASELINE / f"{name}.callgrind.{number}"
        # This routine classifies role from positive raw ancestry and validates
        # the exact selected-owner incoming edge.
        dump = LEGACY.validate_numbered_dump(path, number, {
            "name": name,
            "kind": job["kind"],
            "shape": job["shape"],
            "owner": job["owner"],
            "runner": job["runner"],
            "setup_callers": job["setup_callers"],
        }, plan)
        parsed = LEGACY.parse_raw_profile(path, job["owner"])
        attribution = _attribution(parsed, f"{name}.{number}")
        numbered.append({
            **dump,
            "attribution": attribution,
            # The retained 0536 parser intentionally keeps only the selected
            # Ir accounting and call graph.  These two fields are validated
            # by the independent instruction analyzer; retain their frozen
            # values here for a compact, stable profile envelope.
            "events": ["Ir"],
            "positions": ["instr"],
            "trigger": f"--dump-after={job['owner']}",
        })
    final_path = BASELINE / f"{name}.callgrind"
    require(final_path.is_file(), f"{name}: final Callgrind dump is missing")
    final = LEGACY.final_dump(final_path, last_part + 1, job["owner"])
    setup = [dump for dump in numbered if dump["role"] == "setup"]
    timed = [dump for dump in numbered if dump["role"] == "timed"]
    expected_setup = 0 if job["kind"] == "xls" else 1
    require(len(setup) == expected_setup,
            f"{name}: setup dump count differs")
    require(len(timed) == job["samples"],
            f"{name}: timed dump count differs")
    require(len(numbered) == job["samples"] + expected_setup,
            f"{name}: numbered dump count differs")
    if job["kind"] == "cfb":
        require(numbered[0]["role"] == "setup" and
                all(dump["role"] == "timed" for dump in numbered[1:]),
                f"{name}: CFB setup/timed order differs")
    else:
        require(all(dump["role"] == "timed" for dump in numbered),
                f"{name}: XLS dump role differs")
    return {
        "name": name,
        "repeat": job["repeat"],
        "group": job["group"],
        "kind": job["kind"],
        "shape": job["shape"],
        "owner": job["owner"],
        "runner": job["runner"],
        "receipt": relative(BASELINE / f"{name}.receipt.json"),
        "receipt_sha256": sha256(BASELINE / f"{name}.receipt.json"),
        "report": report,
        "raw_dumps": numbered,
        "setup_dumps": setup,
        "timed_dumps": timed,
        "final_process_dump": final,
        "validation": {
            "positive_owner_edge_classified_from_raw_callgrind": True,
            "runner_or_setup_ancestry_classified_from_positive_edges": True,
            "setup_excluded_from_timed_attribution": True,
            "single_constructor_call_per_timed_dump": all(
                dump["owner_calls"] == 1 for dump in timed
            ),
            "collector_instruction_attribution_present": True,
            "physical_reconciliation_positive_edge_present": True,
            "final_process_dump_zero_ir": True,
        },
    }


def _aggregate_attribution(profiles: list[dict[str, Any]],
                           role: str) -> dict[str, dict[str, int]]:
    """Aggregate named function costs without treating parents as additive."""

    totals: dict[str, dict[str, int]] = {}
    for profile in profiles:
        for dump in profile["raw_dumps"]:
            if dump["role"] != role:
                continue
            for label, attribution in dump["attribution"]["functions"].items():
                row = totals.setdefault(label, {
                    "self_ir": 0,
                    "direct_ir": 0,
                    "inclusive_ir": 0,
                    "calls": 0,
                    "dump_count": 0,
                })
                for field in ("self_ir", "direct_ir", "inclusive_ir", "calls"):
                    row[field] += int(attribution[field])
                row["dump_count"] += 1
    return totals


def analyze() -> dict[str, Any]:
    plan = plan_data()
    metadata = _validate_build(plan)
    profiles = [_profile_job(job, metadata, plan) for job in profile_jobs(plan)]
    require(len(profiles) == 4, "profile matrix is incomplete")
    require(sum(len(profile["timed_dumps"]) for profile in profiles) == 20,
            "expected twenty positive timed dumps")
    require(sum(len(profile["setup_dumps"]) for profile in profiles) == 3,
            "expected three CFB setup dumps")
    return {
        "schema": "cfb_ole2_constructor_callgrind_profile_analysis_0547_v1",
        "status": "pass",
        "stage": "baseline",
        "scope": plan["scope"],
        "performance_claim": "diagnostic-only",
        "selected_owners": {"xls": PROFILE_OWNER_XLS, "cfb": PROFILE_OWNER_CFB},
        "plan": relative(PLAN_PATH),
        "plan_sha256": sha256(PLAN_PATH),
        "metadata": metadata,
        "profile_count": len(profiles),
        "timed_constructor_dump_count": 20,
        "setup_dump_count": 3,
        "profiles": profiles,
        "aggregates": {
            "timed_named_attribution": _aggregate_attribution(profiles, "timed"),
            "setup_named_attribution": _aggregate_attribution(profiles, "setup"),
        },
        "helpers": {
            "change-0536/analyze_profiles.py": sha256(OLD_PROFILE_PATH),
            "analyze.py": sha256(HERE / "analyze.py"),
            "instruction_analysis.py": sha256(HERE / "instruction_analysis.py"),
        },
        "validation": {
            "source_binary_plan_and_driver_bindings": True,
            "all_receipt_artifacts_hashed_and_present": True,
            "all_reports_and_corpus_bindings_valid": True,
            "positive_owner_incoming_edges_valid": True,
            "timed_scope_uses_positive_runner_ancestry": True,
            "CFB_setup_scope_uses_positive_setup_ancestry": True,
            "all_timed_dumps_have_collector_and_physical_attribution": True,
            "all_final_process_dumps_zero_ir": True,
            "no_native_or_speedup_claim": True,
        },
        "limitations": [
            "Callgrind Ir and call metadata are mechanism diagnostics; collection-off call labels are not operation-local dynamic counts.",
            "This single-repeat baseline has no before/after comparison and does not establish native latency, hardware, allocation, RSS, or scaling behavior.",
        ],
    }


def write_report(report: dict[str, Any], output: Path) -> None:
    data = (json.dumps(report, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if output.exists():
        require(output.is_file() and not output.is_symlink(),
                f"refusing to read non-regular output {output}")
        require(output.read_bytes() == data,
                f"refusing to overwrite non-identical output {output}")
        return
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(data)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", "-o", type=Path,
                        default=HERE / "profile-analysis.json")
    args = parser.parse_args(argv)
    try:
        report = analyze()
        write_report(report, args.output)
    except (EvidenceError, AssertionError, KeyError, OSError, TypeError, ValueError) as error:
        print(f"analyze_profiles: error: {error}", file=sys.stderr)
        return 2
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
