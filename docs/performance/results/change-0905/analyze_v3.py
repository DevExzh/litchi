"""Offline custody, correctness, and paired reduction for the 0905 packet.

This module never launches a workload.  It replays retained JSON, stdout,
stderr, resource, source, binary, and receipt metadata.  The qualification
action is intentionally independent from native reduction: it must pass before
the native population is admitted.  The standard XLSX primary uses the
individual validators retained in change-0705; the one-edit controls use the
case-specific checks below so their one-update corpus is never presented as a
one-percent corpus.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import random
import statistics
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
REPO = PACKET.parents[3]
RUNNER = PACKET / "run_v2.py"
SCHEMA = "litchi.paired-audit-capture-0905.v1"
ROLES = ("baseline", "candidate")
PHASES = ("open_ns", "plan_ns", "commit_ns", "publication_ns")
ALL_PHASES = PHASES + ("reopen_ns",)
ALLOCATOR_FIELDS = {
    "plan_allocation_metrics", "staging_allocation_metrics",
    "commit_allocation_metrics", "commit_core_allocation_metrics",
    "publication_allocation_metrics",
}
ALLOCATOR_SCOPE = "operation_global_system_allocator"
ALLOCATOR_INSTRUMENTATION = "system_allocator_operation_scoped"
ALLOCATOR_COUNTER_REVISION = "serialized_region_peak_v3"
ALLOCATOR_SAMPLE_FIELDS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
HEX = set("0123456789abcdef")
CASES = (
    ("primary-percent-medium", "primary", "xlsx_source_backed_cell_values_one_percent_edit_save", "xlsx_cell_crud_shape", "medium"),
    ("primary-percent-dense-sparse", "primary", "xlsx_source_backed_cell_values_one_percent_edit_save", "xlsx_cell_crud_shape", "dense-sparse"),
    ("control-xlsx-one-edit-medium", "control", "xlsx_source_backed_cell_values_one_edit_save", "xlsx_cell_crud_shape", "medium"),
    ("control-docx-one-edit-medium", "control", "docx_semantic_one_edit_save", "semantic_shape", "medium"),
)
CASE_BY_ID = {row[0]: row for row in CASES}
CASE_ORDER = tuple(row[0] for row in CASES)


class ReviewError(RuntimeError):
    """A retained packet failed an independent review check."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReviewError(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as error:
        raise ReviewError(f"cannot read JSON {path}: {error}") from error


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False).encode()


def digest(path: Path) -> dict[str, Any]:
    path = Path(path)
    require(not path.is_symlink(), f"retained artifact is a symlink: {path}")
    require(path.is_file(), f"retained artifact is missing: {path}")
    data = path.read_bytes()
    mode = path.stat().st_mode & 0o777
    return {"path": str(path.resolve()), "bytes": len(data), "mode": mode,
            "executable": bool(mode & 0o111),
            "sha256": hashlib.sha256(data).hexdigest()}


def path_value(value: Any, label: str, base: Path = PACKET) -> Path:
    if isinstance(value, (str, Path)):
        raw = Path(value)
    else:
        require(isinstance(value, dict) and isinstance(value.get("path"), str),
                f"{label} is not a path descriptor")
        raw = Path(value["path"])
    return raw if raw.is_absolute() else (base / raw).absolute()


def descriptor(value: Any, label: str) -> dict[str, Any]:
    path = path_value(value, label)
    actual = digest(path)
    if isinstance(value, dict):
        for key in ("bytes", "sha256"):
            if key in value:
                require(actual[key] == value[key], f"{label}.{key} does not match {path}")
    return actual


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict) and plan.get("schema") == SCHEMA,
            "plan schema is not the frozen 0905 schema")
    require(plan.get("status") in {"frozen", "capture-ready", "ready"},
            f"plan is not frozen: {plan.get('status')!r}")
    require(plan.get("plan_id") == "0905" and plan.get("seed") == 905009,
            "plan identity or seed changed")
    gate = plan.get("gate")
    require(isinstance(gate, dict), "plan.gate is missing")
    require(gate.get("bootstrap") == {
        "replicates": 10000, "seed": 905009, "statistic": "paired_median_ratio"
    }, "bootstrap gate changed")
    require(gate.get("primary") == {
        "p50_median_ratio_max": 0.98, "upper_95_ci_lt": 1.0
    }, "primary gate changed")
    require(gate.get("review") == {
        "regression_ratio": 1.05, "p50_spread": 1.20,
        "tail_metrics": ["p50", "p95", "p99"]
    }, "review gate changed")
    return plan


def case_rows(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    rows = plan.get("cases")
    require(isinstance(rows, list) and len(rows) == len(CASES),
            "plan must contain exactly four cases")
    seen: set[str] = set()
    output: dict[str, dict[str, Any]] = {}
    for row in rows:
        require(isinstance(row, dict) and isinstance(row.get("id"), str),
                "malformed plan case")
        case_id = row["id"]
        require(case_id in CASE_BY_ID and case_id not in seen, f"unexpected case id {case_id}")
        expected = CASE_BY_ID[case_id]
        require((case_id, row.get("kind"), row.get("case"),
                 "semantic_shape" if expected[3] == "semantic_shape" else "xlsx_cell_crud_shape",
                 row.get(expected[3])) == expected,
                f"case definition changed for {case_id}")
        seen.add(case_id)
        output[case_id] = row
    require(tuple(output) == CASE_ORDER, "plan case order changed")
    return output


def environment(plan: dict[str, Any]) -> dict[str, str]:
    value = plan.get("environment")
    if isinstance(value, dict) and "fixed" in value:
        value = value["fixed"]
    require(isinstance(value, dict), "plan.environment must be flat")
    result = {str(key): item for key, item in value.items()}
    require(all(isinstance(key, str) and isinstance(item, str)
                for key, item in result.items()), "environment contains a non-string")
    require(result == {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"},
            "fixed child environment changed")
    return dict(sorted(result.items()))


def source_specs(plan: dict[str, Any]) -> tuple[Any, Any, set[str], Path, Path]:
    source = plan.get("source")
    require(isinstance(source, dict), "plan.source is missing")
    baseline = source.get("baseline_manifest")
    candidate = source.get("candidate_manifest")
    names = source.get("candidate_files")
    require(baseline is not None and candidate is not None and isinstance(names, list)
            and len(names) == 1, "0905 requires exactly one candidate source file")
    candidate_names: set[str] = set()
    for name in names:
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and Path(name).as_posix() == name and ".." not in Path(name).parts,
                f"invalid candidate source path {name!r}")
        candidate_names.add(name)
    root = path_value(source.get("root_worktree", str(REPO)), "source.root_worktree")
    source_root = path_value(source.get("source_root"), "source.source_root")
    return baseline, candidate, candidate_names, root, source_root


def manifest(value: Any, label: str) -> tuple[Path, dict[str, dict[str, Any]]]:
    path = path_value(value, label)
    descriptor(value, label)
    raw = read_json(path)
    files = raw.get("files") if isinstance(raw, dict) else raw
    require(isinstance(files, dict) and files, f"{label} files map is empty")
    result: dict[str, dict[str, Any]] = {}
    for name, item in files.items():
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and Path(name).as_posix() == name and ".." not in Path(name).parts,
                f"{label} has an unsafe path {name!r}")
        require(isinstance(item, dict) and isinstance(item.get("bytes"), int)
                and isinstance(item.get("sha256"), str) and len(item["sha256"]) == 64,
                f"{label} entry {name} is malformed")
        result[name] = item
    return path, result


def source_check(plan: dict[str, Any]) -> dict[str, Any]:
    baseline_spec, candidate_spec, candidate_names, root, source_root = source_specs(plan)
    require(root.is_dir(), f"source root is missing: {root}")
    require(source_root.is_dir(), f"isolated source root is missing: {source_root}")
    _, baseline = manifest(baseline_spec, "baseline source manifest")
    _, candidate = manifest(candidate_spec, "candidate source manifest")
    expected_count = plan.get("source", {}).get("file_count")
    require(isinstance(expected_count, int) and expected_count > 0,
            "plan.source.file_count must be a positive frozen integer")
    require(len(candidate) == expected_count and set(baseline) == set(candidate),
            "source census count or file set changed")
    changed = {name for name in baseline
               if (baseline[name].get("bytes"), baseline[name].get("sha256"))
               != (candidate[name].get("bytes"), candidate[name].get("sha256"))}
    require(changed == candidate_names,
            f"source delta is not exactly the selected file: {sorted(changed)}")
    roots = ((root, candidate, "candidate worktree"),
             (source_root, candidate, "isolated candidate source"))
    checked = 0
    for current_root, expected_map, label in roots:
        for name, expected in expected_map.items():
            actual = digest(current_root / name)
            require(actual["bytes"] == expected["bytes"] and actual["sha256"] == expected["sha256"],
                    f"{label} differs at {name}")
            checked += 1
    return {"baseline_manifest": descriptor(baseline_spec, "baseline source manifest"),
            "candidate_manifest": descriptor(candidate_spec, "candidate source manifest"),
            "candidate_files": sorted(candidate_names), "file_count": len(candidate),
            "hashed_entries": checked, "root_worktree": str(root.resolve()),
            "source_root": str(source_root.resolve())}


def role_descriptors(plan: dict[str, Any], role: str) -> dict[str, Any]:
    roles = plan.get("roles", plan.get("binaries"))
    require(isinstance(roles, dict) and isinstance(roles.get(role), dict),
            f"missing {role} role descriptors")
    output = {}
    for kind in ("native", "allocator"):
        value = roles[role].get(kind)
        require(isinstance(value, dict), f"missing {role}.{kind} descriptor")
        output[kind] = value
    return output


def role_input_bindings(raw: dict[str, Any], plan: dict[str, Any], role: str,
                        label: str) -> None:
    before, after = raw.get("inputs_before"), raw.get("inputs_after")
    require(isinstance(before, list) and before and before == after,
            f"{label} input census is not nonempty and unchanged")
    baseline_spec, candidate_spec, names, _, source_root = source_specs(plan)
    manifest_spec = baseline_spec if role == "baseline" else candidate_spec
    _, expected_manifest = manifest(manifest_spec, f"{role} source manifest")
    indexed: dict[str, list[dict[str, Any]]] = {}
    for item in after:
        require(isinstance(item, dict) and isinstance(item.get("path"), str),
                f"{label} has a malformed input descriptor")
        indexed.setdefault(str(Path(item["path"]).resolve()), []).append(item)
    for name in sorted(names):
        expected_path = str((source_root / name).resolve())
        matches = indexed.get(expected_path, [])
        require(len(matches) == 1, f"{label} does not bind one input for {name}")
        require(matches[0].get("bytes") == expected_manifest[name]["bytes"]
                and matches[0].get("sha256") == expected_manifest[name]["sha256"],
                f"{label} {role} input differs from manifest: {name}")


def build_check(value: Any, label: str, role: str, plan: dict[str, Any]) -> dict[str, Any]:
    actual = descriptor(value, label)
    raw = read_json(Path(actual["path"]))
    require(isinstance(raw, dict) and raw.get("exit_code") == 0,
            f"{label} does not record exit_code=0")
    require(raw.get("inputs_unchanged") is True,
            f"{label} lacks inputs_unchanged=true")
    require(isinstance(raw.get("inputs_before"), list) and raw["inputs_before"],
            f"{label} inputs_before is empty")
    require(raw.get("inputs_before") == raw.get("inputs_after"),
            f"{label} inputs_before/inputs_after differ")
    # The isolated source root is reused sequentially. Bind the selected input
    # to the corresponding role manifest in each write-once receipt.
    role_input_bindings(raw, plan, role, label)
    return actual


def binaries_check(plan: dict[str, Any]) -> dict[str, Any]:
    output: dict[str, Any] = {}
    for role in ROLES:
        output[role] = {}
        for kind, value in role_descriptors(plan, role).items():
            actual = descriptor(value, f"{role} {kind} binary")
            require(bool(actual.get("executable")) or bool(Path(actual["path"]).stat().st_mode & 0o111),
                    f"{role} {kind} binary is not executable")
            receipt = value.get("build_receipt")
            require(receipt is not None, f"{role} {kind} build receipt is missing")
            output[role][kind] = {"binary": actual,
                                  "build_receipt": build_check(receipt, f"{role} {kind} build receipt", role, plan)}
    return output


def counts(plan: dict[str, Any], stage: str) -> tuple[int, int]:
    if stage == "qualification":
        value = plan.get("qualification")
        require(value == {"samples": 2, "warmup": 0}, "qualification counts changed")
        return 2, 0
    if stage == "native":
        value = plan.get("native")
        require(isinstance(value, dict) and value.get("blocks") == 6
                and value.get("samples") == 100 and value.get("warmup") == 10,
                "native counts changed")
        return 100, 10
    value = plan.get("allocator")
    require(isinstance(value, dict) and value.get("repeats") == 2
            and value.get("samples") == 3 and value.get("warmup") == 2,
            "allocator counts changed")
    return 3, 2


def jobs(stage: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    if stage == "qualification":
        for role in ROLES:
            for case_id in CASE_ORDER:
                rows.append({"stage": stage, "index": 0, "role": role, "case_id": case_id})
    elif stage == "native":
        for block in range(6):
            roles = ROLES if block % 2 == 0 else tuple(reversed(ROLES))
            order = CASE_ORDER if block % 2 == 0 else tuple(reversed(CASE_ORDER))
            for case_id in order:
                for role in roles:
                    rows.append({"stage": stage, "index": block, "role": role, "case_id": case_id})
    elif stage == "allocator":
        for repeat in range(2):
            for role in ROLES:
                for case_id in CASE_ORDER:
                    rows.append({"stage": stage, "index": repeat, "role": role, "case_id": case_id})
    else:
        raise ReviewError(f"unknown stage {stage}")
    return rows


def child_name(job: dict[str, Any]) -> str:
    prefix = {"qualification": "q", "native": "b", "allocator": "r"}[job["stage"]]
    return f"{prefix}{job['index']:02d}-{job['role']}-{job['case_id']}"


def command_for(plan: dict[str, Any], job: dict[str, Any], folder: Path) -> list[str]:
    samples, warmup = counts(plan, job["stage"])
    binary = path_value(role_descriptors(plan, job["role"])[
        "allocator" if job["stage"] == "allocator" else "native"]["path"], "binary")
    command = ["/usr/bin/taskset", "-c", str(plan["cpu"])]
    resource = plan.get("resource", {})
    if not isinstance(resource, dict) or resource.get("enabled", True):
        require(isinstance(resource, dict), "resource configuration is malformed")
        command += [resource.get("time_path", "/usr/bin/time"), "-f", resource.get(
            "format", '{"max_rss_kib":%M,"elapsed_seconds":%e,"user_seconds":%U,"system_seconds":%S}'
        ), "-o", str(folder / "resource.json")]
    row = CASE_BY_ID[job["case_id"]]
    command += [str(binary), "--warmup", str(warmup), "--samples", str(samples),
                "--case", row[2]]
    if row[3] == "semantic_shape":
        command += ["--semantic-shape", row[4]]
    else:
        command += ["--xlsx-cell-crud-shape", row[4]]
    command += ["--json", str(folder / "report.json")]
    return command


def resource_enabled(plan: dict[str, Any]) -> bool:
    value = plan.get("resource", {})
    return isinstance(value, dict) and value.get("enabled", True)


def artifact_check(folder: Path, recorded: Any, *, exclude: set[str]) -> None:
    require(isinstance(recorded, list), f"{folder} artifact descriptors are missing")
    actual = [digest(path) for path in sorted(folder.iterdir())
              if path.is_file() and path.name not in exclude]
    require(recorded == actual, f"artifact descriptor drift in {folder}")


def ci_critical(df: int) -> float:
    values = (12.706, 4.303, 3.182, 2.776, 2.571, 2.447, 2.365, 2.306, 2.262, 2.228,
              2.201, 2.179, 2.160, 2.145, 2.131, 2.120, 2.110, 2.101, 2.093, 2.086,
              2.080, 2.074, 2.069, 2.064, 2.060, 2.056, 2.052, 2.048, 2.045, 2.042)
    if df <= 0:
        return 0.0
    if df <= len(values):
        return values[df - 1]
    z = 1.959963984540054
    d = float(df)
    return (z + (z**3 + z) / (4 * d)
            + (5 * z**5 + 16 * z**3 + 3 * z) / (96 * d**2)
            + (3 * z**7 + 19 * z**5 + 17 * z**3 - 15 * z) / (384 * d**3))


def midpoint(left: int, right: int) -> int:
    return left // 2 + right // 2 + (left % 2 + right % 2) // 2


def nearest_rank(values: list[int], percentile: int) -> int:
    index = min((percentile * len(values) + 99) // 100 - 1, len(values) - 1)
    return sorted(values)[index]


def timing_check(elapsed: Any, samples: int, label: str) -> dict[str, Any]:
    require(isinstance(elapsed, dict) and elapsed.get("unit") == "ns", f"{label} elapsed unit")
    values = elapsed.get("samples")
    order = elapsed.get("sample_order")
    require(isinstance(values, list) and len(values) == samples and
            all(isinstance(v, int) and not isinstance(v, bool) and v > 0 for v in values),
            f"{label} elapsed samples")
    require(isinstance(order, list) and len(order) == samples and sorted(order) == list(range(samples)),
            f"{label} sample_order")
    require(values == sorted(values), f"{label} elapsed samples are not sorted")
    ordered = sorted(values)
    expected: dict[str, Any] = {
        "min": ordered[0], "p50": midpoint(ordered[(samples - 1) // 2], ordered[samples // 2]),
        "p95": nearest_rank(values, 95), "p99": nearest_rank(values, 99), "max": ordered[-1],
        "mean": statistics.mean(values),
    }
    for key, value in expected.items():
        reported = elapsed.get(key)
        require(reported == value if isinstance(value, int)
                else isinstance(reported, (int, float)) and math.isclose(float(reported), float(value), rel_tol=1e-12, abs_tol=1e-9),
                f"{label} elapsed.{key} differs from recomputation")
    deviation = statistics.stdev(values) if samples > 1 else 0.0
    reported_sd = elapsed.get("standard_deviation")
    require(isinstance(reported_sd, (int, float))
            and math.isclose(float(reported_sd), deviation, rel_tol=1e-12, abs_tol=1e-9),
            f"{label} standard deviation differs")
    interval = elapsed.get("confidence_interval_95")
    require(isinstance(interval, dict)
            and interval.get("method") == "two-sided Student's t interval for the mean",
            f"{label} confidence interval method changed")
    margin = ci_critical(samples - 1) * deviation / math.sqrt(samples) if samples > 1 else 0.0
    lower, upper = interval.get("lower"), interval.get("upper")
    require(isinstance(lower, (int, float)) and isinstance(upper, (int, float))
            and math.isclose(float(lower), max(0.0, float(expected["mean"]) - margin), rel_tol=1e-12, abs_tol=1e-9)
            and math.isclose(float(upper), float(expected["mean"]) + margin, rel_tol=1e-12, abs_tol=1e-9)
            and float(lower) <= float(upper), f"{label} confidence interval differs")
    return {**expected, "standard_deviation": deviation,
            "confidence_interval_95": {"lower": float(lower), "upper": float(upper)}}


def hex64(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def vector(value: Any, count: int, label: str, *, nonnegative: bool = True) -> list[Any]:
    require(isinstance(value, list) and len(value) == count, f"{label} vector length")
    if nonnegative:
        require(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0 for item in value),
                f"{label} contains an invalid integer")
    return value


def sink_check(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label} sink is missing")
    for key in ("accepted_bytes", "write_calls", "largest_write"):
        require(isinstance(value.get(key), int) and value[key] >= 0, f"{label} sink.{key}")
    require(value["accepted_bytes"] > 0 and value["largest_write"] <= 65536,
            f"{label} sink bounds")
    buckets = value.get("write_size_buckets")
    if buckets is not None:
        names = {"bytes_0", "bytes_1_to_512", "bytes_513_to_4096",
                 "bytes_4097_to_16384", "bytes_16385_to_65536", "bytes_over_65536"}
        require(isinstance(buckets, dict) and set(buckets) == names
                and all(isinstance(item, int) and item >= 0 for item in buckets.values())
                and sum(buckets.values()) == value["write_calls"]
                and buckets["bytes_over_65536"] == 0,
                f"{label} sink buckets")


def corpus_common(corpus: Any, shape: str, package: str, generator: str, label: str) -> dict[str, Any]:
    require(isinstance(corpus, dict), f"{label} corpus is missing")
    require(corpus.get("shape") == shape and corpus.get("package_format") == package
            and corpus.get("generator") == generator and corpus.get("compression") == "deflate",
            f"{label} corpus identity")
    for key in ("entry_count", "archive_member_count", "entry_bytes",
                "uncompressed_payload_bytes", "archive_bytes", "target_payload_bytes"):
        require(isinstance(corpus.get(key), int) and corpus[key] > 0, f"{label} corpus.{key}")
    for key in ("archive_sha256", "target_payload_sha256"):
        hex64(corpus.get(key), f"{label} corpus.{key}")
    return corpus


def xlsx_corpus(corpus: Any, shape: str, label: str) -> dict[str, Any]:
    corpus = corpus_common(corpus, shape, "XLSX/OPC/ZIP",
                           "litchi-xlsx-cell-values-source-edit-media-multi-sheet-v1", label)
    expected = {"medium": (9216, 48, 93), "dense-sparse": (17792, 128, 178)}[shape]
    require(corpus.get("name") == f"xlsx-cell-values-{shape}"
            and corpus.get("payload_kind") == "deterministic-multi-sheet-scalar-grid-with-media"
            and corpus.get("target_entry") == "Sheet1!A1"
            and corpus.get("target_payload_bytes") == 1
            and corpus.get("entry_count") == expected[0], f"{label} XLSX corpus fields")
    xlsx = corpus.get("xlsx")
    require(isinstance(xlsx, dict) and xlsx.get("sheet_count") == 4
            and xlsx.get("rows_per_sheet") == expected[1]
            and xlsx.get("columns_per_sheet") == expected[1]
            and xlsx.get("one_percent_update_count") == expected[2],
            f"{label} XLSX shape fields")
    members = xlsx.get("source_members")
    require(isinstance(members, dict) and members.get("workbook") == "xl/workbook.xml"
            and members.get("styles") == "xl/styles.xml"
            and members.get("shared_strings") is None
            and members.get("worksheets") == [f"xl/worksheets/sheet{i}.xml" for i in range(1, 5)],
            f"{label} XLSX member manifest")
    return corpus


def docx_corpus(corpus: Any, label: str) -> dict[str, Any]:
    corpus = corpus_common(corpus, "medium", "DOCX/OPC/ZIP", "litchi-docx-semantic-v1", label)
    require(corpus.get("name") == "docx-semantic-medium"
            and corpus.get("payload_kind") == "deterministic-semantic-text"
            and corpus.get("entry_count") == 200 and corpus.get("archive_member_count") == 12
            and corpus.get("entry_bytes") == 50 and corpus.get("uncompressed_payload_bytes") == 10000
            and corpus.get("target_entry") == "paragraph:0"
            and corpus.get("target_payload_bytes") == 50,
            f"{label} DOCX corpus fields")
    require(corpus.get("xlsx") is None, f"{label} DOCX corpus carries XLSX metadata")
    return corpus


def allocator_sample(value: Any, label: str, *, measured: bool) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} allocator sample is not an object")
    if not measured:
        require(value.get("status") == "unavailable" and value.get("scope") == ALLOCATOR_SCOPE
                and set(value) <= {"status", "scope"},
                f"{label} unexpectedly contains allocator measurements")
        return {"status": "unavailable", "scope": ALLOCATOR_SCOPE}
    require(value.get("status") == "measured" and value.get("scope") == ALLOCATOR_SCOPE,
            f"{label} allocator measurement scope")
    parsed: dict[str, Any] = {"status": "measured", "scope": ALLOCATOR_SCOPE}
    for field in ALLOCATOR_SAMPLE_FIELDS:
        item = value.get(field)
        require(isinstance(item, int) and not isinstance(item, bool) and item >= 0,
                f"{label}.{field} is not a nonnegative integer")
        parsed[field] = item
    require(parsed["failed_allocation_calls"] == 0,
            f"{label} recorded a failed allocation")
    require(parsed["peak_live_bytes_before"] >= parsed["live_bytes_before"]
            and parsed["peak_live_bytes_after"] >= parsed["live_bytes_after"]
            and parsed["peak_live_bytes_after"] >= parsed["peak_live_bytes_before"]
            and parsed["region_peak_live_bytes"] >= parsed["live_bytes_before"]
            and parsed["region_peak_live_bytes"] >= parsed["live_bytes_after"]
            and parsed["region_peak_live_bytes"] <= parsed["peak_live_bytes_after"],
            f"{label} allocator peak/live boundaries")
    return parsed


def operation_metrics_check(value: Any, samples: int, label: str, *, allocator: bool) -> Any:
    if value is None:
        require(not allocator, f"{label} allocator operation_metrics are missing")
        return None
    require(isinstance(value, dict), f"{label} operation_metrics is not an object")
    if allocator:
        require(value.get("sample_count") == samples
                and isinstance(value.get("sample_indices"), list)
                and sorted(value["sample_indices"]) == list(range(samples))
                and value.get("alignment") == "elapsed_ns.samples_by_elapsed_then_sample_index",
                f"{label} allocator operation_metrics alignment")
    def visit(item: Any, path: str) -> Any:
        if isinstance(item, dict):
            return {key: visit(child, f"{path}.{key}") for key, child in item.items()}
        if isinstance(item, list) and path.endswith(".values"):
            require(len(item) == samples, f"{label} operation metric vector length at {path}")
            require(all((isinstance(child, (int, float)) and not isinstance(child, bool)
                         and math.isfinite(float(child)) and child >= 0)
                        or isinstance(child, str) for child in item),
                    f"{label} operation metric vector values at {path}")
            return item
        return item
    return visit(value, label)


def source_identity(source: Any, samples: int, label: str, *, one_edit: bool,
                    allocator: bool, worksheet_count: int) -> dict[str, Any]:
    require(isinstance(source, dict), f"{label} source summary is missing")
    summary = source.get("xlsx_cell_values")
    require(isinstance(summary, dict) and summary.get("implementation") == "source-backed"
            and summary.get("cache_mode") == "unmanaged-control"
            and summary.get("cache_budget_managed") is False,
            f"{label} source-backed unmanaged summary")
    timing_scope = summary.get("timing_scope")
    require(isinstance(timing_scope, str)
            and all(word in timing_scope for word in ("open", "commit", "publication", "reopen")),
            f"{label} timing scope")
    require(isinstance(worksheet_count, int) and worksheet_count > 0,
            f"{label} worksheet count is invalid")
    if one_edit:
        require(summary.get("update_count") == 1 and summary.get("selected_worksheet_count") == 1,
                f"{label} one-edit selection counts")
    else:
        require(isinstance(summary.get("update_count"), int) and summary["update_count"] > 1
                and isinstance(summary.get("selected_worksheet_count"), int)
                and summary["selected_worksheet_count"] == worksheet_count,
                f"{label} percent selection counts")
    for key in ("payload_memory_limit", "publication_planning_memory_headroom", "cache_budget_memory_limit"):
        require(summary.get(key) is None, f"{label} unmanaged {key}")
    # All serialized source vectors must have a stable non-timing identity.  A
    # phase vector is separately validated and omitted from the identity.
    phases: dict[str, list[int]] = {}
    allocator_metrics: dict[str, list[dict[str, Any]]] = {}
    constants: dict[str, Any] = {}
    for key, value in summary.items():
        if key in ALLOCATOR_FIELDS:
            values = vector(value, samples, f"{label} source.xlsx_cell_values.{key}", nonnegative=False)
            allocator_metrics[key] = [allocator_sample(item, f"{label}.{key}[{index}]",
                                                        measured=allocator)
                                      for index, item in enumerate(values)]
            continue
        if isinstance(value, list):
            vector(value, samples, f"{label} source.xlsx_cell_values.{key}", nonnegative=False)
            if key in ALL_PHASES:
                require(all(isinstance(item, int) and not isinstance(item, bool) and item >= 0 for item in value),
                        f"{label} phase vector")
                phases[key] = value
            else:
                require(all(item == value[0] for item in value), f"{label} source identity varies: {key}")
                constants[key] = value[0]
        else:
            constants[key] = value
    if phases:
        require(set(phases) == set(ALL_PHASES), f"{label} incomplete phase vectors")
        elapsed = None
    for key in ("output_sha256", "semantic_sha256", "untouched_member_sha256"):
        hex64(constants.get(key), f"{label} {key}")
    require(constants.get("payload_materializations") == constants.get("cache_successful_loads")
            and constants.get("cache_cold_loads") == constants.get("cache_successful_loads")
            and constants.get("cache_retained_entries") == constants.get("cache_successful_loads")
            and constants.get("cache_retained_bytes", 0) > 0,
            f"{label} source cache counters")
    for key in ("cache_hits", "cache_waiter_joins", "cache_failed_loads", "cache_evictions",
                "cache_bypasses", "cache_oversized_bypasses", "cache_allocation_bypasses",
                "cache_in_flight_loads", "cache_budget_memory_used", "cache_budget_reserved_bytes",
                "cache_budget_reservation_failures"):
        require(constants.get(key) == 0, f"{label} source invariant {key}")
    if one_edit:
        expected_unselected = worksheet_count - summary["selected_worksheet_count"]
        require(expected_unselected > 0
                and constants.get("unselected_worksheet_read_calls") == expected_unselected
                and constants.get("unselected_worksheet_read_bytes", 0) > 0,
                f"{label} untouched worksheet compressed-copy reads")
    else:
        require(constants.get("unselected_worksheet_read_calls") == 0
                and constants.get("unselected_worksheet_read_bytes") == 0,
                f"{label} all-selected worksheet read accounting")
    require(constants.get("source_read_calls", 0) > 0 and constants.get("source_read_bytes", 0) > 0
            and constants.get("workbook_read_calls") == 1
            and constants.get("selected_worksheet_read_calls", 0) >= summary["selected_worksheet_count"]
            and constants.get("selected_worksheet_read_bytes", 0) > 0
            and constants.get("workbook_read_bytes", 0) > 0,
            f"{label} source read counters")
    for generic, specific in (("read_calls", "source_read_calls"),
                              ("read_bytes", "source_read_bytes"),
                              ("ordinary_payload_materializations", "payload_materializations")):
        require(source.get(generic) == summary.get(specific), f"{label} generic {generic}")
    for key, value in source.items():
        if isinstance(value, list):
            vector(value, samples, f"{label} source.{key}", nonnegative=False)
            require(all(item == value[0] for item in value), f"{label} generic source identity varies: {key}")
    return {"summary": summary, "constants": constants,
            "phases": phases, "allocator_metrics": allocator_metrics,
            "generic": {key: source[key] for key in source if key != "xlsx_cell_values"}}


def phase_alignment(result: dict[str, Any], phases: dict[str, list[int]], label: str) -> None:
    if not phases:
        return
    phase_sums = [sum(phases[phase][index] for phase in PHASES)
                  for index in range(len(phases[PHASES[0]]))]
    order = result["elapsed_ns"]["sample_order"]
    require([phase_sums[index] for index in order] == result["elapsed_ns"]["samples"]
            and sorted(range(len(phase_sums)), key=lambda index: (phase_sums[index], index)) == order,
            f"{label} phase vectors do not align with elapsed samples")


def primary_validate(result: dict[str, Any], shape: str, samples: int, validator: Any,
                     label: str, *, allocator: bool) -> dict[str, Any]:
    corpus = validator.validate_corpus(result.get("corpus", {}), shape)
    sink_check(result.get("sink"), label)
    validator.validate_sink(result.get("sink"), label)
    if samples > 30:
        validator.check_reported_statistics(result.get("elapsed_ns", {}), samples)
    source = result.get("source")
    if allocator:
        source_details = source_identity(
            source, samples, label, one_edit=False, allocator=True,
            worksheet_count=corpus["xlsx"]["sheet_count"])
        phases, constants = source_details["phases"], source_details["constants"]
        require(constants["output_sha256"] == result.get("output_sha256"),
                f"{label} output digest")
        phase_alignment(result, phases, label)
    else:
        phases, constants = validator.validate_source(result, corpus, shape, samples, compatibility=False)
        require(result.get("output_sha256") == constants["output_sha256"], f"{label} output digest")
        source_details = source_identity(
            source, samples, label, one_edit=False, allocator=False,
            worksheet_count=corpus["xlsx"]["sheet_count"])
    return {"corpus": corpus, "sink": result["sink"], "source": source,
            "phase_vectors": phases, "source_constants": constants,
            "allocator_metrics": source_details["allocator_metrics"]}


def xlsx_control_validate(result: dict[str, Any], shape: str, samples: int, label: str,
                          *, allocator: bool) -> dict[str, Any]:
    corpus = xlsx_corpus(result.get("corpus"), shape, label)
    sink_check(result.get("sink"), label)
    source = source_identity(
        result.get("source"), samples, label, one_edit=True,
        allocator=allocator, worksheet_count=corpus["xlsx"]["sheet_count"])
    require(isinstance(result.get("output_sha256"), str), f"{label} output digest missing")
    hex64(result["output_sha256"], f"{label} output digest")
    require(source["constants"]["output_sha256"] == result["output_sha256"],
            f"{label} source/output digest mismatch")
    phase_alignment(result, source["phases"], label)
    return {"corpus": corpus, "sink": result["sink"], "source": result["source"],
            "phase_vectors": source["phases"], "source_constants": source["constants"],
            "allocator_metrics": source["allocator_metrics"]}


def docx_control_validate(result: dict[str, Any], samples: int, label: str,
                          *, allocator: bool) -> dict[str, Any]:
    corpus = docx_corpus(result.get("corpus"), label)
    sink_check(result.get("sink"), label)
    require(result.get("source") is None and result.get("output_sha256") is None,
            f"{label} DOCX result has fabricated XLSX source/output identity")
    return {"corpus": corpus, "sink": result["sink"], "source": None,
            "phase_vectors": {}, "source_constants": {}, "allocator_metrics": {}}


def semantic_identity(result: dict[str, Any], validated: dict[str, Any]) -> dict[str, Any]:
    source = validated.get("source")
    identity: dict[str, Any] = {"corpus": validated["corpus"], "sink": validated["sink"]}
    if result.get("output_sha256") is not None:
        identity["output_sha256"] = result["output_sha256"]
    if isinstance(source, dict):
        # Keep source counters, cache/resource constants, and digests.  Timing
        # vectors and allocator measurements are deliberately excluded.
        def strip(value: Any, key: str = "") -> Any:
            if key.endswith("_ns") or "allocation" in key:
                return None
            if isinstance(value, dict):
                return {k: stripped for k, item in value.items()
                        if (stripped := strip(item, k)) is not None}
            if isinstance(value, list):
                items = [strip(item, key) for item in value]
                # Source evidence vectors are required to be constant by the
                # case validators.  Compare their semantic value rather than
                # their population-specific vector length (qualification=2,
                # native=100, allocator=3).
                if items and all(item == items[0] for item in items):
                    return items[0]
                return items
            return value
        identity["source"] = strip(source)
    return identity


def report_validate(plan: dict[str, Any], job: dict[str, Any], binary: dict[str, Any],
                    validator: Any, report: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    label = f"{job['stage']}/{child_name(job)}"
    samples, _ = counts(plan, job["stage"])
    require(report.get("schema_version") == 1, f"{label} report schema")
    require(report.get("binary_identity", {}).get("binary_sha256") == binary["sha256"],
            f"{label} report binary identity")
    tool = report.get("tool", {})
    require(tool.get("profile") == "release", f"{label} report profile")
    allocator = job["stage"] == "allocator"
    if allocator:
        require(tool.get("instrumentation") == ALLOCATOR_INSTRUMENTATION
                and tool.get("allocator_counter_revision") == ALLOCATOR_COUNTER_REVISION,
                f"{label} allocator instrumentation")
    else:
        require(tool.get("instrumentation") == "none", f"{label} native instrumentation")
    env = report.get("environment", {})
    require(env.get("cpu_affinity") == str(plan["cpu"])
            and env.get("git_revision") is None and env.get("git_worktree_dirty") is None
            and env.get("rustc_version") is None, f"{label} runtime metadata")
    config = report.get("configuration", {})
    row = CASE_BY_ID[job["case_id"]]
    require(config.get("cases") == [row[2]]
            and config.get("samples_per_case") == samples
            and config.get("warmup_iterations_per_case") == counts(plan, job["stage"])[1],
            f"{label} report configuration")
    if row[3] == "semantic_shape":
        require(config.get("semantic_shapes") == [row[4]], f"{label} semantic shape configuration")
    else:
        require(config.get("xlsx_cell_crud_shapes") == [row[4]], f"{label} XLSX shape configuration")
    results = report.get("results")
    require(isinstance(results, list) and len(results) == 1 and results[0].get("case") == row[2],
            f"{label} result case")
    result = results[0]
    timing = timing_check(result.get("elapsed_ns"), samples, label)
    operation_metrics = operation_metrics_check(result.get("operation_metrics"), samples, label,
                                                allocator=allocator)
    if row[1] == "primary":
        validated = primary_validate(result, row[4], samples, validator, label,
                                     allocator=allocator)
    elif row[2].startswith("xlsx_"):
        validated = xlsx_control_validate(result, row[4], samples, label,
                                           allocator=allocator)
    else:
        validated = docx_control_validate(result, samples, label, allocator=allocator)
    identity = semantic_identity(result, validated)
    return {"label": label, "stage": job["stage"], "index": job["index"],
            "role": job["role"], "case_id": job["case_id"], "samples": samples,
            "p50": timing["p50"], "p95": timing["p95"], "p99": timing["p99"],
            "mean": timing["mean"], "elapsed": timing,
            "identity": identity, "operation_metrics": operation_metrics,
            "operation_metrics_status": "available" if operation_metrics is not None else "not_applicable",
            "allocator_metrics": validated.get("allocator_metrics", {}),
            "allocator_metrics_status": "measured" if validated.get("allocator_metrics") else "not_applicable",
            "resource": None}, result


def leaf_review(plan: dict[str, Any], job: dict[str, Any], binaries: dict[str, Any],
                validator: Any) -> dict[str, Any]:
    folder = PACKET / job["stage"] / child_name(job)
    require(folder.is_dir(), f"missing child folder {folder}")
    start = read_json(folder / "start.json")
    command = read_json(folder / "command.json")
    receipt = read_json(folder / "receipt.json")
    require(not (folder / "validation-failure.json").exists(), f"retained child failure {folder}")
    binary_kind = "allocator" if job["stage"] == "allocator" else "native"
    binary = binaries[job["role"]][binary_kind]["binary"]
    expected_command = command_for(plan, job, folder)
    env = environment(plan)
    env_hash = hashlib.sha256(canonical(env)).hexdigest()
    for record, name in ((start, "start"), (command, "command"), (receipt, "receipt")):
        require(record.get("schema") == SCHEMA, f"{folder} {name} schema")
        require(record.get("argv") == expected_command and record.get("cwd") == str(REPO.resolve())
                and record.get("environment") == env and record.get("environment_sha256") == env_hash
                and record.get("binary") == binary,
                f"{folder} {name} invocation custody")
        require(record.get("plan") == descriptor(PACKET / "plan.json", "plan"),
                f"{folder} {name} plan custody")
        require(record.get("helper") == descriptor(RUNNER, "capture helper"),
                f"{folder} {name} capture helper custody")
    require(command.get("exit_code") == 0 and receipt.get("exit_code") == 0
            and receipt.get("status") == "pass" and not receipt.get("errors"),
            f"{folder} child receipt failed")
    require((folder / "stdout.txt").is_file() and (folder / "stderr.txt").is_file()
            and (folder / "report.json").is_file(), f"{folder} retained output missing")
    if resource_enabled(plan):
        resource = read_json(folder / "resource.json")
        require(isinstance(resource, dict) and isinstance(resource.get("max_rss_kib"), int)
                and resource["max_rss_kib"] >= 0, f"{folder} RSS receipt")
    artifact_check(folder, command.get("artifacts"), exclude={"command.json", "receipt.json"})
    artifact_check(folder, receipt.get("artifacts"), exclude={"receipt.json"})
    require(start.get("source") == command.get("source") == receipt.get("source_start")
            and command.get("source_end") == receipt.get("source_end")
            and receipt.get("source_error") is None and command.get("source_error") is None,
            f"{folder} source custody receipt")
    reviewed, result = report_validate(plan, job, binary, validator, read_json(folder / "report.json"))
    reviewed["folder"] = str(folder.resolve())
    reviewed["report"] = digest(folder / "report.json")
    reviewed["command"] = digest(folder / "command.json")
    reviewed["receipt"] = digest(folder / "receipt.json")
    reviewed["resource"] = read_json(folder / "resource.json") if resource_enabled(plan) else None
    return reviewed


def capture_review(plan: dict[str, Any], stage: str, binaries: dict[str, Any], validator: Any) -> list[dict[str, Any]]:
    capture = read_json(PACKET / stage / "capture.json")
    expected = jobs(stage)
    samples, warmup = counts(plan, stage)
    require(capture.get("schema") == SCHEMA and capture.get("status") == "pass"
            and capture.get("stage") == stage and capture.get("children") == len(expected)
            and capture.get("samples_per_child") == samples and capture.get("warmup_per_child") == warmup,
            f"{stage} capture manifest counts")
    require(capture.get("plan") == descriptor(PACKET / "plan.json", "plan")
            and capture.get("helper") == descriptor(RUNNER, "capture helper")
            and capture.get("environment") == environment(plan)
            and capture.get("environment_sha256") == hashlib.sha256(canonical(environment(plan))).hexdigest()
            and capture.get("binary_bindings") == binaries, f"{stage} capture custody")
    expected_jobs = [{"name": child_name(job), "stage": stage, "index": job["index"],
                      "role": job["role"], "case_id": job["case_id"]} for job in expected]
    require(capture.get("jobs") == expected_jobs, f"{stage} schedule changed")
    expected_names = {item["name"] for item in expected_jobs}
    actual_names = {path.name for path in (PACKET / stage).iterdir() if path.is_dir()}
    require(actual_names == expected_names, f"{stage} has extra or missing child folders")
    rows = [leaf_review(plan, job, binaries, validator) for job in expected]
    retained_receipts = capture.get("receipts")
    require(isinstance(retained_receipts, list) and len(retained_receipts) == len(expected),
            f"{stage} capture receipts are incomplete")
    for index, job in enumerate(expected):
        retained = read_json(PACKET / stage / child_name(job) / "receipt.json")
        require(retained_receipts[index] == retained,
                f"{stage} capture receipt differs from retained receipt for {child_name(job)}")
    return rows


def load_validator(plan: dict[str, Any]) -> Any:
    value = plan.get("semantic_validator") or plan.get("prior_semantic_validator")
    require(value is not None, "plan.semantic_validator descriptor is missing")
    actual = descriptor(value, "semantic validator")
    spec = importlib.util.spec_from_file_location("litchi_0905_semantic_validator", actual["path"])
    require(spec is not None and spec.loader is not None, "cannot load semantic validator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def compare_identities(rows: list[dict[str, Any]]) -> None:
    grouped: dict[str, dict[str, Any]] = {}
    for row in rows:
        key = row["case_id"]
        previous = grouped.setdefault(key, row["identity"])
        require(previous == row["identity"], f"semantic identity drift for {key}: {row['label']}")


def bootstrap(values: list[float], replicates: int, seed: int) -> dict[str, Any]:
    require(values, "paired bootstrap has no values")
    rng = random.Random(seed)
    draws: list[float] = []
    size = len(values)
    for _ in range(replicates):
        sample = [values[rng.randrange(size)] for _ in range(size)]
        draws.append(float(statistics.median(sample)))
    draws.sort()
    return {"replicates": replicates, "seed": seed, "statistic": "paired_median_ratio",
            "estimate": float(statistics.median(values)),
            "lower_95": draws[(25 * replicates + 999) // 1000 - 1],
            "upper_95": draws[(975 * replicates + 999) // 1000 - 1]}


def native_reduce(rows: list[dict[str, Any]], plan: dict[str, Any]) -> dict[str, Any]:
    native = [row for row in rows if row["stage"] == "native"]
    require(len(native) == 48, "native child count is not 48")
    pairs: dict[str, list[dict[str, Any]]] = {case_id: [] for case_id in CASE_ORDER}
    for block in range(6):
        for case_id in CASE_ORDER:
            pair = [row for row in native if row["index"] == block and row["case_id"] == case_id]
            require({row["role"] for row in pair} == set(ROLES) and len(pair) == 2,
                    f"native pair missing for block {block}/{case_id}")
            baseline = next(row for row in pair if row["role"] == "baseline")
            candidate = next(row for row in pair if row["role"] == "candidate")
            require(baseline["p50"] > 0, f"zero baseline p50 for {case_id}")
            pairs[case_id].append({"block": block, "baseline": baseline["p50"],
                                   "candidate": candidate["p50"],
                                   "ratio": candidate["p50"] / baseline["p50"]})
    gate = plan["gate"]
    summaries: dict[str, Any] = {}
    for case_id, values in pairs.items():
        ratios = [item["ratio"] for item in values]
        result = bootstrap(ratios, gate["bootstrap"]["replicates"], gate["bootstrap"]["seed"])
        result["blocks"] = values
        result["accepted"] = (CASE_BY_ID[case_id][1] == "primary"
                               and result["estimate"] <= gate["primary"]["p50_median_ratio_max"]
                               and result["upper_95"] < gate["primary"]["upper_95_ci_lt"])
        summaries[case_id] = result
    review_flags: list[dict[str, Any]] = []
    review_ratio = gate["review"]["regression_ratio"]
    for case_id in CASE_ORDER:
        for metric in gate["review"]["tail_metrics"]:
            base = [row[metric] for row in native if row["case_id"] == case_id and row["role"] == "baseline"]
            cand = [row[metric] for row in native if row["case_id"] == case_id and row["role"] == "candidate"]
            ratio = statistics.median(cand) / statistics.median(base)
            if ratio > review_ratio:
                review_flags.append({"case_id": case_id, "metric": metric,
                                     "ratio": ratio, "threshold": review_ratio})
        for role in ROLES:
            values = [row["p50"] for row in native if row["case_id"] == case_id and row["role"] == role]
            spread = max(values) / min(values)
            if spread > gate["review"]["p50_spread"]:
                review_flags.append({"case_id": case_id, "role": role,
                                     "metric": "p50_spread", "ratio": spread,
                                     "threshold": gate["review"]["p50_spread"]})
    primary = [summaries[case_id] for case_id in CASE_ORDER if CASE_BY_ID[case_id][1] == "primary"]
    return {"cases": summaries, "review_flags": review_flags,
            "primary_gate_pass": all(item["accepted"] for item in primary)}


def write_once(path: Path, value: Any) -> None:
    require(not path.exists(), f"refusing to overwrite {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    try:
        with path.open("x", encoding="utf-8") as handle:
            handle.write(json.dumps(value, indent=2, sort_keys=True) + "\n")
    except FileExistsError as error:
        raise ReviewError(f"refusing to overwrite {path}") from error


def analysis(plan: dict[str, Any], stage: str, output: Path) -> dict[str, Any]:
    case_rows(plan)
    source = source_check(plan)
    binaries = binaries_check(plan)
    validator = load_validator(plan)
    if stage == "qualification":
        rows = capture_review(plan, "qualification", binaries, validator)
        compare_identities(rows)
        result = {"schema": "litchi.paired-audit-reduction-0905.v1", "status": "pass",
                  "stage": "qualification", "source": source,
                  "rows": rows, "qualification_admitted": True,
                  "scope": "semantic/statistical qualification only; no native timing reduction"}
    elif stage == "native":
        qualification = capture_review(plan, "qualification", binaries, validator)
        compare_identities(qualification)
        require((PACKET / "qualification-analysis.json").is_file(),
                "native reduction requires qualification-analysis.json")
        native = capture_review(plan, "native", binaries, validator)
        compare_identities(qualification + native)
        result = {"schema": "litchi.paired-audit-reduction-0905.v1", "status": "pass",
                  "stage": "native", "source": source, "qualification": qualification,
                  "native": native, "reduction": native_reduce(native, plan),
                  "rss": "descriptive / executable bias retained"}
    elif stage == "allocator":
        qualification = capture_review(plan, "qualification", binaries, validator)
        compare_identities(qualification)
        allocator = capture_review(plan, "allocator", binaries, validator)
        compare_identities(qualification + allocator)
        require(len(allocator) == 16, "allocator child count is not 16")
        result = {"schema": "litchi.paired-audit-reduction-0905.v1", "status": "pass",
                  "stage": "allocator", "source": source, "allocator": allocator,
                  "scope": "allocator diagnostics; phase evidence is reported only when present"}
    else:
        qualification = capture_review(plan, "qualification", binaries, validator)
        compare_identities(qualification)
        require((PACKET / "qualification-analysis.json").is_file(),
                "full reduction requires qualification-analysis.json")
        native = capture_review(plan, "native", binaries, validator)
        allocator = capture_review(plan, "allocator", binaries, validator)
        compare_identities(qualification + native + allocator)
        result = {"schema": "litchi.paired-audit-reduction-0905.v1", "status": "pass",
                  "stage": "all", "source": source, "qualification": qualification,
                  "native": native, "allocator": allocator,
                  "reduction": native_reduce(native, plan),
                  "rss": "descriptive / executable bias retained"}
    write_once(output, result)
    return result


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=["qualification", "native", "allocator", "all"])
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    plan = load_plan()
    default = {"qualification": PACKET / "qualification-analysis.json",
               "native": PACKET / "analysis.json", "allocator": PACKET / "allocator-analysis.json",
               "all": PACKET / "analysis.json"}[args.action]
    result = analysis(plan, args.action, args.output or default)
    print(json.dumps({"status": result["status"], "stage": result["stage"]}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except (ReviewError, AssertionError, OSError, ValueError) as error:
        print(f"0905 reduction failed: {error}", file=sys.stderr)
        raise SystemExit(1)
