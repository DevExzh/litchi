"""Replay retained 0905 evidence without reopening cleaned source roots.

The replay is deliberately independent of the capture and reduction entry
points.  It hashes retained manifests, the prior source archive, role overlays,
build receipts and binaries, then asks the retained analyzer to review the raw
capture tree.  Paths recorded inside build receipts and the frozen plan are
treated as custody identifiers; the deleted isolated source roots are never
opened.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import random
import statistics
import tarfile
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
PLAN_PATH = PACKET / "plan.json"
RUNNER_PATH = PACKET / "run_v2.py"
ANALYZER_PATH = PACKET / "analyze_v3.py"
MICRO_PLAN_PATH = PACKET / "micro-plan.json"
MICRO_HELPER_PATH = PACKET / "micro.py"
MAIN_SCHEMA = "litchi.paired-audit-capture-0905.v1"
REDUCTION_SCHEMA = "litchi.paired-audit-reduction-0905.v1"
MICRO_PLAN_SCHEMA = "litchi.xml-attribute-micro-plan-0905.v1"
MICRO_SCHEMA = "litchi.xml-attribute-micro-0905.v1"
ROLES = ("baseline", "candidate")
MAIN_STAGES = ("qualification", "native", "allocator")
HEX = set("0123456789abcdef")
MICRO_PROTOCOL = {
    "cpu": 12,
    "blocks": 6,
    "samples": 30,
    "warmup": 5,
    "iterations": 1000,
    "bootstrap_draws": 10000,
    "seed": 905019,
    "p50_ratio_threshold": 1.05,
    "lower_95_threshold": 1.0,
    "review_ratio": 1.05,
    "spread_ratio": 1.2,
}
MICRO_CASES = [
    *(f"cardinality-{count}" for count in (0, 1, 2, 8, 32, 33, 128)),
    "xml-space-valid", "xml-space-invalid", "duplicate", "malformed-tail",
    "malformed-second",
]
MICRO_JOBS = [
    f"{case}/{policy}"
    for case in MICRO_CASES
    for policy in ("source", "compact", "authored")
]


class ReplayError(RuntimeError):
    """Retained evidence failed an offline replay check."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReplayError(message)


def read_json(path: Path, label: str) -> Any:
    try:
        value = json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        raise ReplayError(f"cannot read {label} {path}: {error}") from error
    return value


def path_value(value: Any, label: str, *, base: Path = PACKET) -> Path:
    if isinstance(value, (str, Path)):
        raw = Path(value)
    else:
        require(isinstance(value, dict) and isinstance(value.get("path"), str),
                f"{label} is not a path descriptor")
        raw = Path(value["path"])
    return raw if raw.is_absolute() else (base / raw).absolute()


def digest(path: Path, label: str = "artifact") -> dict[str, Any]:
    path = Path(path)
    require(not path.is_symlink(), f"{label} is a symlink: {path}")
    require(path.is_file(), f"{label} is not a retained regular file: {path}")
    data = path.read_bytes()
    mode = path.stat().st_mode & 0o777
    return {
        "path": str(path.resolve()),
        "bytes": len(data),
        "mode": mode,
        "executable": bool(mode & 0o111),
        "sha256": hashlib.sha256(data).hexdigest(),
    }


def simple_digest(path: Path, label: str = "artifact") -> dict[str, Any]:
    """Match the three-field descriptor emitted by the micro helper."""

    actual = digest(path, label)
    return {key: actual[key] for key in ("path", "bytes", "sha256")}


def simple_descriptor(value: Any, label: str) -> dict[str, Any]:
    actual = simple_digest(path_value(value, label), label)
    if isinstance(value, dict):
        for key in ("bytes", "sha256"):
            if key in value:
                require(actual[key] == value[key], f"{label}.{key} differs")
    return actual


def descriptor(value: Any, label: str) -> dict[str, Any]:
    actual = digest(path_value(value, label), label)
    if isinstance(value, dict):
        for key in ("bytes", "sha256", "mode", "executable"):
            if key in value:
                require(actual[key] == value[key], f"{label}.{key} differs")
    return actual


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite replay output {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as output:
            json.dump(value, output, indent=2, sort_keys=True)
            output.write("\n")
    except FileExistsError as error:
        raise ReplayError(f"refusing to overwrite replay output {path}") from error


def safe_member_name(name: str, label: str) -> str:
    path = Path(name)
    require(name and not path.is_absolute() and path.as_posix() == name
            and ".." not in path.parts,
            f"{label} contains an unsafe relative path {name!r}")
    return name


def check_sha(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def load_analyzer() -> Any:
    require(ANALYZER_PATH.is_file(), f"missing retained analyzer {ANALYZER_PATH}")
    spec = importlib.util.spec_from_file_location("litchi_0905_replay_analyzer",
                                                   ANALYZER_PATH)
    require(spec is not None and spec.loader is not None,
            "cannot load retained analyzer")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_main_plan(analyzer: Any) -> dict[str, Any]:
    plan = analyzer.load_plan()
    require(plan.get("schema") == MAIN_SCHEMA, "main plan schema is not 0905")
    require(plan.get("status") in {"frozen", "capture-ready", "ready"},
            "main plan is not frozen")
    require(PLAN_PATH.is_file(), f"missing main plan {PLAN_PATH}")
    return plan


def verify_main_plan_assets(plan: dict[str, Any]) -> None:
    helpers = plan.get("helpers")
    require(isinstance(helpers, dict), "main plan helper bindings are missing")
    for name, value in helpers.items():
        descriptor(value, f"main helper {name}")
    for key in ("method", "environment_receipt"):
        if key in plan:
            descriptor(plan[key], f"main plan {key}")


def manifest(value: Any, label: str) -> tuple[dict[str, Any], dict[str, Any], dict[str, dict[str, Any]]]:
    actual = descriptor(value, label)
    raw = read_json(Path(actual["path"]), label)
    require(isinstance(raw, dict) and isinstance(raw.get("files"), dict),
            f"{label} does not contain a files map")
    files: dict[str, dict[str, Any]] = {}
    for name, item in raw["files"].items():
        safe_member_name(name, label)
        require(isinstance(item, dict) and isinstance(item.get("bytes"), int)
                and item["bytes"] >= 0, f"{label} entry {name} has invalid bytes")
        check_sha(item.get("sha256"), f"{label} entry {name}.sha256")
        files[name] = item
    if "file_count" in raw:
        require(raw["file_count"] == len(files), f"{label}.file_count differs")
    return actual, raw, files


def archive_descriptor(source: dict[str, Any], raw: dict[str, Any], role: str) -> Any:
    for key in (f"{role}_archive", "archive"):
        if key in source:
            return source[key]
    for key in ("archive", "prior_archive"):
        if key in raw:
            return raw[key]
    raise ReplayError(f"no retained source archive descriptor for {role}")


def archive_records(value: Any, files: dict[str, dict[str, Any]], label: str) -> tuple[dict[str, Any], dict[str, dict[str, Any]], set[str]]:
    actual = descriptor(value, label)
    records: dict[str, dict[str, Any]] = {}
    try:
        archive = tarfile.open(actual["path"], "r:*")
    except (OSError, tarfile.TarError) as error:
        raise ReplayError(f"cannot open {label}: {error}") from error
    with archive:
        for member in archive:
            if member.isdir():
                continue
            require(member.isfile(), f"{label} contains a non-regular member {member.name!r}")
            name = safe_member_name(member.name, label)
            require(name not in records, f"{label} contains duplicate member {name}")
            stream = archive.extractfile(member)
            require(stream is not None, f"cannot read {label} member {name}")
            data = stream.read()
            records[name] = {
                "bytes": len(data),
                "sha256": hashlib.sha256(data).hexdigest(),
            }
    require(set(records) <= set(files),
            f"{label} contains files absent from its retained manifest")
    mismatched: set[str] = set()
    for name, expected in files.items():
        observed = records.get(name)
        if observed is None or (observed["bytes"], observed["sha256"]) != (
                expected["bytes"], expected["sha256"]):
            mismatched.add(name)
    return actual, records, mismatched


def overlay_roots(source: dict[str, Any], role: str) -> list[Path]:
    values: list[Any] = []
    for key in (f"{role}_overlay", f"{role}_overlay_root"):
        if key in source:
            values.append(source[key])
    if not values:
        values.append(PACKET / f"{role}-overlay")
    roots = [path_value(value, f"{role} overlay") for value in values]
    if role == "candidate":
        baseline_value = source.get("baseline_overlay", source.get("baseline_overlay_root"))
        roots.append(path_value(baseline_value, "baseline overlay")
                     if baseline_value is not None else PACKET / "baseline-overlay")
    return roots


def overlay_records(roots: list[Path], fallback: dict[str, Path], label: str) -> dict[str, dict[str, Any]]:
    records: dict[str, dict[str, Any]] = {}
    found_root = False
    for root in roots:
        if root.is_dir():
            found_root = True
            candidates = sorted(path for path in root.rglob("*") if path.is_file())
            for path in candidates:
                require(not path.is_symlink(), f"{label} contains a symlink {path}")
                name = safe_member_name(path.relative_to(root).as_posix(), label)
                actual = digest(path, f"{label} {name}")
                previous = records.get(name)
                require(previous is None or (previous["bytes"], previous["sha256"])
                        == (actual["bytes"], actual["sha256"]),
                        f"{label} duplicate overlay differs: {name}")
                records[name] = actual
        elif root.exists():
            require(root.is_file(), f"{label} overlay root is not a file or directory: {root}")
            found_root = True
            name = next((name for name, path in fallback.items() if path == root), None)
            require(name is not None, f"{label} standalone overlay has no bound name: {root}")
            records[name] = digest(root, f"{label} {name}")
    for name, path in fallback.items():
        if path.is_file():
            found_root = True
            actual = digest(path, f"{label} {name}")
            previous = records.get(name)
            require(previous is None or (previous["bytes"], previous["sha256"])
                    == (actual["bytes"], actual["sha256"]),
                    f"{label} fallback overlay differs: {name}")
            records[name] = actual
    require(found_root, f"missing retained {label} overlay")
    return records


def verify_overlay_set(records: dict[str, dict[str, Any]], expected_names: set[str],
                       files: dict[str, dict[str, Any]], label: str) -> None:
    require(set(records) == expected_names,
            f"{label} paths differ: expected {sorted(expected_names)}, got {sorted(records)}")
    for name in expected_names:
        expected = files[name]
        actual = records[name]
        require(actual["bytes"] == expected["bytes"]
                and actual["sha256"] == expected["sha256"],
                f"{label} differs from manifest: {name}")


def verify_sources(plan: dict[str, Any]) -> dict[str, Any]:
    source = plan.get("source")
    require(isinstance(source, dict), "main plan source is missing")
    baseline_spec = source.get("baseline_manifest")
    candidate_spec = source.get("candidate_manifest")
    require(baseline_spec is not None and candidate_spec is not None,
            "main plan source manifests are missing")
    baseline_desc, baseline_raw, baseline_files = manifest(baseline_spec, "baseline manifest")
    candidate_desc, candidate_raw, candidate_files = manifest(candidate_spec, "candidate manifest")
    require(set(baseline_files) == set(candidate_files),
            "baseline and candidate manifest path sets differ")
    expected_count = source.get("file_count")
    require(isinstance(expected_count, int) and expected_count > 0
            and len(baseline_files) == expected_count,
            "source manifest count is not the frozen positive count")
    names = source.get("candidate_files")
    require(isinstance(names, list) and len(names) == 1 and isinstance(names[0], str),
            "0905 source delta must name exactly one candidate file")
    candidate_names = {names[0]}
    changed = {
        name for name in baseline_files
        if (baseline_files[name]["bytes"], baseline_files[name]["sha256"]) !=
        (candidate_files[name]["bytes"], candidate_files[name]["sha256"])
    }
    require(changed == candidate_names,
            f"source manifest delta differs: {sorted(changed)}")

    baseline_archive_value = archive_descriptor(source, baseline_raw, "baseline")
    candidate_archive_value = archive_descriptor(source, candidate_raw, "candidate")
    baseline_archive, _, baseline_overlay_names = archive_records(
        baseline_archive_value, baseline_files, "baseline source archive")
    candidate_archive, _, candidate_overlay_names = archive_records(
        candidate_archive_value, candidate_files, "candidate source archive")
    baseline_fallback: dict[str, Path] = {}
    candidate_fallback = {names[0]: PACKET / "candidate-audit.rs"}
    baseline_overlays = overlay_records(overlay_roots(source, "baseline"),
                                       baseline_fallback, "baseline")
    verify_overlay_set(baseline_overlays, baseline_overlay_names, baseline_files,
                       "baseline overlay")
    candidate_overlays = overlay_records(overlay_roots(source, "candidate"),
                                        candidate_fallback, "candidate")
    verify_overlay_set(candidate_overlays, candidate_overlay_names, candidate_files,
                       "candidate overlay")
    return {
        "baseline_manifest": baseline_desc,
        "candidate_manifest": candidate_desc,
        "baseline_archive": baseline_archive,
        "candidate_archive": candidate_archive,
        "file_count": expected_count,
        "candidate_files": sorted(candidate_names),
        "baseline_overlay_files": sorted(baseline_overlay_names),
        "candidate_overlay_files": sorted(candidate_overlay_names),
        "source_roots_checked": False,
        "source_root_claim": "deleted isolated source roots are not reopened",
    }


def normalized_path_text(value: Any) -> str:
    return str(Path(str(value)).absolute())


def role_input_binding(raw: dict[str, Any], role: str, plan: dict[str, Any],
                       files: dict[str, dict[str, Any]], label: str) -> None:
    before, after = raw.get("inputs_before"), raw.get("inputs_after")
    require(isinstance(before, list) and before and before == after,
            f"{label} input census is not stable and nonempty")
    source = plan["source"]
    source_root = source.get("source_root")
    require(isinstance(source_root, str), f"{label} source_root is absent")
    names = source["candidate_files"]
    for name in names:
        expected_path = normalized_path_text(Path(source_root) / name)
        matches = [item for item in after if isinstance(item, dict)
                   and isinstance(item.get("path"), str)
                   and normalized_path_text(item["path"]) == expected_path]
        require(len(matches) == 1, f"{label} lacks one retained input for {name}")
        item = matches[0]
        require(item.get("bytes") == files[name]["bytes"]
                and item.get("sha256") == files[name]["sha256"],
                f"{label} {role} source input differs for {name}")
    for item in after:
        require(isinstance(item, dict) and isinstance(item.get("path"), str)
                and isinstance(item.get("bytes"), int) and item["bytes"] >= 0,
                f"{label} has a malformed retained input descriptor")
        check_sha(item.get("sha256"), f"{label} input sha256")


def verify_main_binaries(plan: dict[str, Any], source: dict[str, Any],
                         baseline_files: dict[str, dict[str, Any]],
                         candidate_files: dict[str, dict[str, Any]]) -> dict[str, Any]:
    roles = plan.get("roles")
    require(isinstance(roles, dict), "main plan roles are missing")
    output: dict[str, Any] = {}
    for role in ROLES:
        require(isinstance(roles.get(role), dict), f"missing main role {role}")
        output[role] = {}
        expected_files = baseline_files if role == "baseline" else candidate_files
        for kind in ("native", "allocator"):
            spec = roles[role].get(kind)
            require(isinstance(spec, dict), f"missing {role}.{kind} binary descriptor")
            binary = descriptor(spec, f"{role} {kind} binary")
            require(binary["executable"], f"{role} {kind} binary is not executable")
            receipt_spec = spec.get("build_receipt")
            require(receipt_spec is not None, f"{role} {kind} build receipt is missing")
            receipt = descriptor(receipt_spec, f"{role} {kind} build receipt")
            raw = read_json(Path(receipt["path"]), f"{role} {kind} build receipt")
            require(isinstance(raw, dict) and raw.get("exit_code") == 0
                    and raw.get("inputs_unchanged") is True,
                    f"{role} {kind} build receipt is unsuccessful")
            role_input_binding(raw, role, plan, expected_files,
                               f"{role} {kind} build receipt")
            if isinstance(raw.get("log"), dict):
                descriptor(raw["log"], f"{role} {kind} build log")
            output[role][kind] = {"binary": binary, "build_receipt": receipt}
    return output


def main_capture_replay(analyzer: Any, plan: dict[str, Any], binaries: dict[str, Any],
                        analysis_path: Path) -> dict[str, Any]:
    validator = analyzer.load_validator(plan)
    rows = {
        stage: analyzer.capture_review(plan, stage, binaries, validator)
        for stage in MAIN_STAGES
    }
    flattened = rows["qualification"] + rows["native"] + rows["allocator"]
    analyzer.compare_identities(flattened)
    reduction = analyzer.native_reduce(rows["native"], plan)
    canonical = read_json(analysis_path, "canonical main analysis")
    require(isinstance(canonical, dict)
            and canonical.get("schema") == REDUCTION_SCHEMA
            and canonical.get("status") == "pass"
            and canonical.get("stage") == "all",
            "canonical main analysis is not a passing 0905 all-stage result")
    for stage in MAIN_STAGES:
        require(canonical.get(stage) == rows[stage],
                f"canonical main analysis differs in {stage}")
    require(canonical.get("reduction") == reduction,
            "canonical main reduction differs from independent replay")
    return {
        "analysis": digest(analysis_path, "canonical main analysis"),
        "children": {stage: len(rows[stage]) for stage in MAIN_STAGES},
        "qualification": {"children": len(rows["qualification"])},
        "native": {"children": len(rows["native"]),
                   "primary_gate_pass": reduction["primary_gate_pass"],
                   "review_flags": len(reduction["review_flags"])},
        "allocator": {"children": len(rows["allocator"])},
        "reduction": reduction,
    }


def load_micro_plan() -> dict[str, Any]:
    require(MICRO_PLAN_PATH.is_file(), f"missing micro plan {MICRO_PLAN_PATH}")
    plan = read_json(MICRO_PLAN_PATH, "micro plan")
    require(isinstance(plan, dict) and plan.get("schema") == MICRO_PLAN_SCHEMA
            and plan.get("status") == "frozen", "micro plan is not frozen 0905")
    require(plan.get("protocol") == MICRO_PROTOCOL, "micro protocol changed")
    require(plan.get("jobs") == MICRO_JOBS, "micro job order changed")
    helper = simple_descriptor(plan.get("helper"), "micro helper")
    require(helper == simple_digest(MICRO_HELPER_PATH, "micro helper"),
            "micro helper descriptor does not bind current helper")
    for key in ("probe", "method", "semantic_parity"):
        require(key in plan, f"micro plan lacks {key} descriptor")
        descriptor(plan[key], f"micro {key}")
    return plan


def verify_micro_binaries(plan: dict[str, Any]) -> dict[str, Any]:
    output: dict[str, Any] = {}
    binaries = plan.get("binaries")
    receipts = plan.get("build_receipts")
    require(isinstance(binaries, dict) and isinstance(receipts, dict),
            "micro binaries or build receipts are missing")
    for role in ROLES:
        binary = descriptor(binaries.get(role), f"micro {role} binary")
        require(binary["executable"], f"micro {role} binary is not executable")
        receipt = descriptor(receipts.get(role), f"micro {role} build receipt")
        raw = read_json(Path(receipt["path"]), f"micro {role} build receipt")
        require(isinstance(raw, dict) and raw.get("exit_code") == 0
                and raw.get("inputs_unchanged") is True,
                f"micro {role} build receipt is unsuccessful")
        require(isinstance(raw.get("inputs_before"), list) and raw["inputs_before"]
                and raw["inputs_before"] == raw.get("inputs_after"),
                f"micro {role} build receipt input census is unstable")
        output[role] = {"binary": binary, "build_receipt": receipt}
    return output


def micro_start(plan: dict[str, Any], block: int, role: str) -> dict[str, Any]:
    binary = simple_descriptor(plan["binaries"][role], f"micro {role} binary")
    argv = ["taskset", "-c", str(plan["protocol"]["cpu"]), binary["path"],
            "--samples", str(plan["protocol"]["samples"]),
            "--warmup", str(plan["protocol"]["warmup"]),
            "--iterations", str(plan["protocol"]["iterations"])]
    if block % 2:
        argv.append("--reverse")
    return {
        "block": block,
        "role": role,
        "argv": argv,
        "binary": binary,
        "plan": simple_digest(MICRO_PLAN_PATH, "micro plan"),
        "helper": simple_digest(MICRO_HELPER_PATH, "micro helper"),
    }


def parse_micro_tsv(path: Path, plan: dict[str, Any], block: int,
                    identity: list[str] | None) -> tuple[list[str], dict[str, list[tuple[int, int]]]]:
    comments: list[str] = []
    cases: dict[str, list[tuple[int, int]]] = {}
    for line_number, line in enumerate(path.read_text().splitlines(), 1):
        if line.startswith(("# input\t", "# outcome\t")):
            comments.append(line)
        if line.startswith("#") or line.startswith("case\t"):
            continue
        fields = line.split("\t")
        require(len(fields) == 5, f"{path}:{line_number} does not have five TSV fields")
        case, policy, sample_text, elapsed_text, iterations_text = fields
        key = f"{case}/{policy}"
        require(key in MICRO_JOBS, f"{path}:{line_number} has an unexpected job {key}")
        try:
            sample = int(sample_text)
            elapsed = int(elapsed_text)
            iterations = int(iterations_text)
        except ValueError as error:
            raise ReplayError(f"{path}:{line_number} has a non-integer field") from error
        require(0 <= sample < plan["protocol"]["samples"] and elapsed > 0
                and iterations == plan["protocol"]["iterations"],
                f"{path}:{line_number} has invalid sample data")
        cases.setdefault(key, []).append((sample, elapsed))
    expected_order = list(reversed(MICRO_JOBS)) if block % 2 else MICRO_JOBS
    require(list(cases) == expected_order,
            f"{path} job order differs from the frozen block order")
    for key, values in cases.items():
        require([item[0] for item in values] == list(range(plan["protocol"]["samples"])),
                f"{path} samples are incomplete or duplicated for {key}")
    if identity is not None:
        require(comments == identity, f"{path} input/outcome identity differs")
    return comments, cases


def micro_capture_replay(plan: dict[str, Any], micro_analysis_path: Path) -> dict[str, Any]:
    verify_micro_binaries(plan)
    folder = PACKET / "micro"
    require(folder.is_dir(), f"missing micro capture directory {folder}")
    suffixes = (".start.json", ".tsv", ".stderr", ".receipt.json")
    expected_names = {
        f"b{block:02}-{role}{suffix}"
        for block in range(plan["protocol"]["blocks"])
        for role in ROLES
        for suffix in suffixes
    }
    actual_names = {path.name for path in folder.iterdir() if path.is_file()}
    require(actual_names == expected_names, "micro capture artifact set differs")
    rows: dict[str, dict[str, list[dict[str, Any]]]] = {}
    identity: list[str] | None = None
    for block in range(plan["protocol"]["blocks"]):
        for role in ROLES:
            stem = folder / f"b{block:02}-{role}"
            start = read_json(stem.with_suffix(".start.json"), f"micro {stem.name} start")
            require(start == micro_start(plan, block, role),
                    f"micro {stem.name} start custody differs")
            receipt = read_json(stem.with_suffix(".receipt.json"),
                                f"micro {stem.name} receipt")
            for key, value in start.items():
                require(receipt.get(key) == value,
                        f"micro {stem.name} receipt differs in {key}")
            tsv = stem.with_suffix(".tsv")
            stderr = stem.with_suffix(".stderr")
            require(receipt.get("exit_code") == 0
                    and receipt.get("stdout") == simple_digest(tsv, f"micro {stem.name} TSV")
                    and receipt.get("stderr") == simple_digest(stderr, f"micro {stem.name} stderr"),
                    f"micro {stem.name} receipt is unsuccessful")
            comments, cases = parse_micro_tsv(tsv, plan, block, identity)
            if identity is None:
                identity = comments
            for key, values in cases.items():
                ordered = sorted(item[1] for item in values)
                rows.setdefault(key, {}).setdefault(role, []).append({
                    "block": block,
                    "p50": statistics.median(ordered),
                    "p95": ordered[28],
                    "p99": ordered[29],
                })
    require(set(rows) == set(MICRO_JOBS), "micro case set is incomplete")
    results: dict[str, Any] = {}
    rng = random.Random(plan["protocol"]["seed"])
    for key in sorted(rows):
        roles = rows[key]
        require(set(roles) == set(ROLES), f"micro role pair is incomplete for {key}")
        baseline = roles["baseline"]
        candidate = roles["candidate"]
        require([row["block"] for row in baseline] == list(range(6))
                and [row["block"] for row in candidate] == list(range(6)),
                f"micro block order is incomplete for {key}")
        ratios = [right["p50"] / left["p50"]
                  for left, right in zip(baseline, candidate)]
        draws = sorted(
            statistics.median([ratios[rng.randrange(6)] for _ in range(6)])
            for _ in range(plan["protocol"]["bootstrap_draws"])
        )
        median_ratio = statistics.median(ratios)
        lower = draws[249]
        upper = draws[9749]
        material = (median_ratio > plan["protocol"]["p50_ratio_threshold"]
                    and lower > plan["protocol"]["lower_95_threshold"])
        results[key] = {
            "baseline": baseline,
            "candidate": candidate,
            "paired_p50_ratios": ratios,
            "median_paired_p50_ratio": median_ratio,
            "bootstrap_95": [lower, upper],
            "material_adverse_regression": material,
            "tail_review_flags": {
                metric: statistics.median(right[metric] / left[metric]
                                          for left, right in zip(baseline, candidate))
                > plan["protocol"]["review_ratio"]
                for metric in ("p95", "p99")
            },
            "spread_flags": {
                role_name: max(item["p50"] for item in values)
                / min(item["p50"] for item in values)
                > plan["protocol"]["spread_ratio"]
                for role_name, values in roles.items()
            },
        }
    canonical = read_json(micro_analysis_path, "canonical micro analysis")
    adverse = [key for key, value in results.items()
               if value["material_adverse_regression"]]
    expected = {
        "schema": MICRO_SCHEMA,
        "status": "pass",
        "plan": simple_digest(MICRO_PLAN_PATH, "micro plan"),
        "cases": results,
        "children": 12,
        "samples": 12 * len(MICRO_JOBS) * plan["protocol"]["samples"],
        "material_adverse_cases": adverse,
        "adoption_allowed": not adverse,
    }
    for key, value in expected.items():
        require(canonical.get(key) == value,
                f"canonical micro analysis differs in {key}")
    return {
        "analysis": digest(micro_analysis_path, "canonical micro analysis"),
        "children": 12,
        "samples": expected["samples"],
        "material_adverse_cases": adverse,
        "adoption_allowed": not adverse,
        "cases": results,
    }


def verify_optional_parity(plan: dict[str, Any]) -> dict[str, Any] | None:
    value = plan.get("semantic_parity")
    if value is None:
        return None
    actual = descriptor(value, "semantic parity comparison")
    raw = read_json(Path(actual["path"]), "semantic parity comparison")
    require(isinstance(raw, dict) and raw.get("status") == "pass"
            and raw.get("byte_identical") is True,
            "semantic parity comparison is not a passing byte-identical result")
    return {"comparison": actual, "status": raw["status"],
            "byte_identical": raw["byte_identical"]}


def main() -> dict[str, Any]:
    analyzer = load_analyzer()
    plan = load_main_plan(analyzer)
    verify_main_plan_assets(plan)
    source = verify_sources(plan)
    baseline_desc, _, baseline_files = manifest(plan["source"]["baseline_manifest"],
                                               "baseline manifest")
    candidate_desc, _, candidate_files = manifest(plan["source"]["candidate_manifest"],
                                                 "candidate manifest")
    del baseline_desc, candidate_desc
    binaries = verify_main_binaries(plan, source, baseline_files, candidate_files)
    main = main_capture_replay(analyzer, plan, binaries,
                               PACKET / "analysis.json")
    micro_plan = load_micro_plan()
    micro = micro_capture_replay(micro_plan, PACKET / "micro-analysis.json")
    parity = verify_optional_parity(plan)
    return {
        "schema": "litchi.paired-audit-replay-0905.v1",
        "status": "pass",
        "mode": "offline retained artifact/source/archive/overlay replay; no live source-root claim",
        "source": source,
        "binaries": binaries,
        "main": main,
        "micro": micro,
        "semantic_parity": parity,
        "plan": digest(PLAN_PATH, "main plan"),
        "capture_helper": digest(RUNNER_PATH, "main capture helper"),
        "analyzer": digest(ANALYZER_PATH, "main analyzer"),
        "replay_helper": digest(Path(__file__), "replay helper"),
    }


def cli() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path, required=True,
                        help="new JSON path; existing paths are refused")
    args = parser.parse_args()
    result = main()
    write_once(args.output, result)
    print("0905 retained replay PASS")


if __name__ == "__main__":
    try:
        cli()
    except (ReplayError, OSError, ValueError, KeyError, tarfile.TarError) as error:
        print(f"0905 replay failed: {error}")
        raise SystemExit(1)
