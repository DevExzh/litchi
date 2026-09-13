#!/usr/bin/env python3
"""Validate the normal benchmark reports used by the 0547 profile lane.

0547 is a baseline-only Callgrind attribution pass.  It deliberately has no
native comparison or admission decision.  This module keeps the report and
corpus checks from the retained 0511 verifier, while binding them to the
single 0547 normal binary and the four five-sample profile jobs.
"""

from __future__ import annotations

import argparse
import importlib.util
import json
import math
import re
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
PLAN_PATH = HERE / "plan.json"
BASELINE = HERE / "baseline"
SHA256_RE = re.compile(r"^[0-9a-f]{64}$")


def _load(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


OLD = _load(HERE.parent / "change-0511" / "verify.py", "litchi_0511_verify_0547")


class EvidenceError(ValueError):
    """A missing, malformed, or contradictory evidence item."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise EvidenceError(message)


def read(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise EvidenceError(f"cannot read {path}: {error}") from error


def sha(path: Path) -> str:
    try:
        import hashlib

        digest = hashlib.sha256()
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(block)
        return digest.hexdigest()
    except OSError as error:
        raise EvidenceError(f"cannot hash {path}: {error}") from error


def check_hash(value: Any, label: str) -> str:
    require(isinstance(value, str) and SHA256_RE.fullmatch(value) is not None,
            f"{label} is not a lowercase SHA-256")
    return value


def plan_data() -> dict[str, Any]:
    plan = read(PLAN_PATH)
    require(isinstance(plan, dict), "plan is not an object")
    require(isinstance(plan.get("revision"), str)
            and len(plan["revision"]) == 40,
            "plan revision is missing")
    require(plan.get("scope") ==
            "Fresh baseline-only OLE2 collector sub-operation attribution",
            "plan scope differs from the 0547 attribution pass")
    require(plan.get("priority") ==
            "OLE2/OOXML active; ODF deferred until that goal completes; iWork excluded",
            "plan priority differs")
    require(plan.get("cpu") == 2, "plan CPU differs from the pinned CPU")

    groups = plan.get("groups")
    require(isinstance(groups, dict), "plan groups are missing")
    require(groups.get("xls", {}).get("cases") == [
        "xls_semantic_open",
        "xls_eager_open_list_worksheets",
        "xls_eager_open_one_cell",
        "xls_source_backed_open",
        "xls_source_backed_open_list_worksheets",
        "xls_source_backed_open_one_cell",
        "xls_owned_source_open",
        "xls_owned_source_open_list_worksheets",
        "xls_owned_source_open_one_cell",
    ], "plan XLS case matrix differs")
    cfb = groups.get("cfb", {})
    require(cfb.get("cases") == ["cfb_open"],
            "plan CFB case matrix differs")
    require(cfb.get("shapes") == ["tiny", "many-small", "few-large"],
            "plan CFB shape matrix differs")
    require(cfb.get("payload") == "incompressible",
            "plan CFB payload differs")

    profile = plan.get("profile")
    require(isinstance(profile, dict), "profile plan is missing")
    require(profile.get("repeats") == 1 and profile.get("warmup") == 0
            and profile.get("samples") == 5,
            "0547 profile counts must be one repeat, zero warmup, five samples")
    require(profile.get("jobs") == [
        "xls-owned", "cfb-tiny", "cfb-many-small", "cfb-few-large",
    ], "0547 profile job matrix differs")
    require(profile.get("xls_owner") ==
            "litchi_xls::workbook::source::SourceBackedWorkbook::from_read_at_with_limits",
            "plan XLS profile owner differs")
    require(profile.get("cfb_owner") == "litchi_cfb::file::OleFile<R>::open",
            "plan CFB profile owner differs")
    require(profile.get("instruction_flags") == [
        "--dump-instr=yes", "--dump-line=no", "--compress-pos=no",
        "--collect-jumps=yes",
    ], "Callgrind instruction flags differ from the frozen contract")

    owned = plan.get("owned_paths")
    require(owned == ["/home/zhuhe/litchi-goal-0547-target"],
            "0547 must have exactly one owned target path")
    require(plan.get("status") == "frozen-before-build-and-capture",
            "0547 plan is not frozen before build and capture")
    assembly = plan.get("assembly")
    require(isinstance(assembly, dict), "assembly plan is missing")
    require(assembly.get("owners") == [
        "SectorChainScratch", "CheckedBitSet", "chain_error",
    ], "assembly owners differ from the frozen contract")
    require(assembly.get("required") == ["collect_exact", "insert"],
            "assembly required symbols differ from the frozen contract")
    return plan


def profile_jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    profile = plan["profile"]
    jobs: list[dict[str, Any]] = [{
        "name": "profile-r1-xls-owned",
        "repeat": 1,
        "group": "xls-owned",
        "kind": "xls",
        "shape": None,
        "runner": "litchi_perf_baseline::run_xls_owned_source_case",
        "owner": profile["xls_owner"],
        "samples": profile["samples"],
        "warmup": profile["warmup"],
        "selection": {"cases": ["xls_owned_source_open_one_cell"]},
    }]
    for shape in plan["groups"]["cfb"]["shapes"]:
        jobs.append({
            "name": f"profile-r1-cfb-{shape}",
            "repeat": 1,
            "group": f"cfb-{shape}",
            "kind": "cfb",
            "shape": shape,
            "runner": "litchi_perf_baseline::run_cfb_open",
            "owner": profile["cfb_owner"],
            "samples": profile["samples"],
            "warmup": profile["warmup"],
            "selection": {
                "cases": ["cfb_open"],
                "shapes": [shape],
                "payload": "incompressible",
            },
        })
    return jobs


def normalized(value: Any, label: str) -> Any:
    if isinstance(value, list):
        require(value, f"{label}: identity vector is empty")
        values = [normalized(item, f"{label}[{index}]")
                  for index, item in enumerate(value)]
        require(all(item == values[0] for item in values),
                f"{label}: identity vector varies across samples")
        return values[0]
    if isinstance(value, dict):
        return {key: normalized(item, f"{label}.{key}")
                for key, item in sorted(value.items())}
    return value


def identity(row: dict[str, Any], label: str = "row") -> dict[str, Any]:
    require(isinstance(row, dict), f"{label}: result row is not an object")
    return {
        "case": row.get("case"),
        "corpus": row.get("corpus"),
        "sink": row.get("sink"),
        "source": normalized(row.get("source"), f"{label}.source")
        if row.get("source") is not None else None,
        "output_sha256": row.get("output_sha256"),
    }


def _validate_report_row(row: dict[str, Any], job: dict[str, Any]) -> dict[str, Any]:
    case = job["selection"]["cases"][0]
    corpus = row.get("corpus")
    require(isinstance(corpus, dict), f"{case}: corpus is missing")
    # XLS profile corpora intentionally retain the generator's concrete shape
    # (for example, ``256-comments-opaque-heavy``); the selection has no
    # --shape selector.  CFB shapes are selected explicitly and must match.
    shape = corpus.get("shape")
    require(isinstance(shape, str) and shape, f"{case}: corpus shape is missing")
    context = f"{case}/{shape}"
    require(row.get("case") == case, f"{context}: case differs")
    elapsed = row.get("elapsed_ns")
    require(isinstance(elapsed, dict), f"{context}: elapsed metrics are missing")
    # The retained helper checks positivity, order, statistics, and sample
    # count.  Keep its result instead of reimplementing statistical rules.
    checked_elapsed = OLD._validate_elapsed(row, job["samples"], context)
    if job["kind"] == "xls":
        OLD.verify_xls_row(row, job["samples"], context, False)
    else:
        require(row.get("source") is None and row.get("sink") is None,
                f"{context}: CFB row unexpectedly publishes source/sink data")
        require(row.get("output_sha256") is None,
                f"{context}: CFB row unexpectedly publishes output data")
        OLD.verify_operation_metrics(
            row, job["samples"], context, False, checked_elapsed["sample_order"]
        )
    operation = row.get("operation_metrics")
    require(isinstance(operation, dict), f"{context}: operation metrics are missing")
    allocation = operation.get("allocation")
    require(isinstance(allocation, dict), f"{context}: allocation envelope is missing")
    require(allocation.get("status") == "unavailable",
            f"{context}: normal allocation status is not unavailable")
    require(allocation.get("scope") == "operation_global_system_allocator",
            f"{context}: allocation scope differs")
    require("values" not in allocation,
            f"{context}: normal allocation envelope contains values")
    return row


def validate_report(path: Path, job: dict[str, Any], kind: str = "normal") -> list[dict[str, Any]]:
    """Validate one profile report and its corpus binding."""

    plan = plan_data()
    require(kind == "normal", "0547 has no allocator or native report lane")
    report = read(path)
    metadata = read(BASELINE / "binary-normal.json")
    require(isinstance(metadata, dict), "normal binary metadata is not an object")
    binary_sha = check_hash(metadata.get("sha256"), "normal binary metadata")
    OLD.verify_report_identity(
        report, {"binary_sha256": binary_sha}, job["samples"],
        job["warmup"], False, path.name,
    )
    environment = report.get("environment")
    require(isinstance(environment, dict), f"{path.name}: environment is missing")
    require(environment.get("git_revision") == plan["revision"],
            f"{path.name}: report revision differs from plan")
    require(environment.get("cpu_affinity") == str(plan["cpu"]),
            f"{path.name}: report CPU affinity differs from plan")
    binary_identity = report.get("binary_identity")
    require(isinstance(binary_identity, dict),
            f"{path.name}: binary identity is missing")
    require(binary_identity.get("path") == metadata.get("path"),
            f"{path.name}: report binary path differs from build identity")
    require(binary_identity.get("binary_bytes") == metadata.get("bytes"),
            f"{path.name}: report binary size differs")
    configuration = report.get("configuration")
    require(isinstance(configuration, dict), f"{path.name}: configuration is missing")
    selection = job["selection"]
    require(configuration.get("cases") == selection["cases"],
            f"{path.name}: selected cases differ")
    if "shapes" in selection:
        require(configuration.get("corpus_shapes") == selection["shapes"],
                f"{path.name}: selected CFB shape differs")
        require(configuration.get("payload_kinds") == [selection["payload"]],
                f"{path.name}: selected CFB payload differs")
    rows = report.get("results")
    require(isinstance(rows, list) and len(rows) == 1,
            f"{path.name}: expected one result row")
    catalog_path = path.with_name(path.name.replace(".json", ".catalog.json"))
    catalog = read(catalog_path)
    OLD.validate_binding(report, catalog)
    return [_validate_report_row(rows[0], job)]


def validate_profile_jobs() -> list[dict[str, Any]]:
    plan = plan_data()
    output: list[dict[str, Any]] = []
    metadata = read(BASELINE / "binary-normal.json")
    binary_sha = check_hash(metadata.get("sha256"), "normal binary metadata")
    for job in profile_jobs(plan):
        name = job["name"]
        report_path = BASELINE / f"{name}.json"
        receipt_path = BASELINE / f"{name}.receipt.json"
        require(report_path.is_file(), f"{name}: report is missing")
        require(receipt_path.is_file(), f"{name}: receipt is missing")
        receipt = read(receipt_path)
        require(receipt.get("exit_code") == 0, f"{name}: child failed")
        require(receipt.get("execution_stage") == "baseline",
                f"{name}: execution stage differs")
        require(receipt.get("plan_sha256") == sha(PLAN_PATH),
                f"{name}: receipt plan binding differs")
        require(receipt.get("source_manifest_sha256") ==
                sha(BASELINE / "source-manifest.json"),
                f"{name}: receipt source binding differs")
        require(receipt.get("binary_sha256") == binary_sha,
                f"{name}: receipt binary binding differs")
        require(receipt.get("script_sha256") == sha(HERE / "run.py"),
                f"{name}: receipt driver binding differs")
        rows = validate_report(report_path, job)
        require(len(rows) == 1, f"{name}: result row count differs")
        artifacts = receipt.get("artifacts")
        require(isinstance(artifacts, dict), f"{name}: receipt artifacts are missing")
        expected = {
            f"{name}.host.json", f"{name}.json", f"{name}.catalog.json",
            f"{name}.stdout", f"{name}.stderr", f"{name}.callgrind",
            *(f"{name}.callgrind.{number}"
              for number in range(1, 6 if job["kind"] == "xls" else 7)),
        }
        require(set(artifacts) == expected,
                f"{name}: receipt artifact set differs")
        for artifact, digest in artifacts.items():
            check_hash(digest, f"{name}/{artifact} receipt hash")
            artifact_path = BASELINE / artifact
            require(artifact_path.is_file() and not artifact_path.is_symlink(),
                    f"{name}: receipt artifact is missing: {artifact}")
            require(sha(artifact_path) == digest,
                    f"{name}: receipt artifact hash differs: {artifact}")
        output.append({
            "name": name,
            "group": job["group"],
            "kind": job["kind"],
            "shape": job["shape"],
            "report": str(report_path.relative_to(HERE)),
            "report_sha256": sha(report_path),
            "receipt": str(receipt_path.relative_to(HERE)),
            "receipt_sha256": sha(receipt_path),
            "identity": identity(rows[0], f"{name}.row"),
            "samples": job["samples"],
        })
    require(len(output) == 4, "0547 profile job matrix is incomplete")
    return output


def analyze() -> dict[str, Any]:
    plan = plan_data()
    profiles = validate_profile_jobs()
    return {
        "schema": "litchi-ole2-change-0547-profile-report-analysis-v1",
        "status": "pass",
        "stage": "baseline",
        "scope": plan["scope"],
        "performance_claim": "diagnostic-only",
        "plan": str(PLAN_PATH.relative_to(HERE)),
        "plan_sha256": sha(PLAN_PATH),
        "source_manifest_sha256": sha(BASELINE / "source-manifest.json"),
        "binary_sha256": read(BASELINE / "binary-normal.json")["sha256"],
        "profile_count": len(profiles),
        "profiles": profiles,
        "validation": {
            "normal_reports_valid": True,
            "corpus_bindings_valid": True,
            "source_binary_plan_and_driver_receipts_bound": True,
            "no_native_or_allocator_comparison": True,
            "no_speedup_claim": True,
        },
        "limitations": [
            "This report validates normal benchmark output only; it makes no native latency, RSS, hardware, allocation, or speedup claim.",
            "Callgrind Ir and call metadata remain mechanism evidence; operation-local calls must not be inferred from collection-off metadata.",
        ],
    }


def write_report(report: dict[str, Any], output: Path) -> None:
    data = (json.dumps(report, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if output.exists():
        require(output.is_file() and not output.is_symlink(),
                f"refusing to read non-regular output {output}")
        require(output.read_bytes() == data,
                f"refusing to overwrite non-identical report {output}")
        return
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(data)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", "-o", type=Path,
                        default=HERE / "analysis.json")
    args = parser.parse_args(argv)
    try:
        report = analyze()
        write_report(report, args.output)
    except (EvidenceError, AssertionError, KeyError, OSError, TypeError, ValueError) as error:
        print(f"analyze: error: {error}")
        return 2
    print(args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
