#!/usr/bin/env python3
"""Independently verify the retained 0500 DOCX performance evidence.

The verifier reads the raw CSV, stderr, receipts, frozen source manifests and
the generated comparison files.  It does not build or run the benchmark.
"""

import argparse
import csv
import hashlib
import json
import math
import re
from pathlib import Path


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
BASE_REVISION = "c797d04c7933e61356effcdccca8e81b430e0762"
HARNESS = "crates/litchi-docx/examples/managed_paragraph_batch_perf.rs"
PARAGRAPHS = (128, 512)
REPLACEMENTS = (1, 8, 32)
SOURCES = ("owned", "file")
BEFORE_MODES = ("repeated",)
AFTER_MODES = ("repeated", "batch")
KINDS = ("scalar_before_after", "after_batch_vs_scalar", "batch_after_vs_scalar_before")
PROFILE_NAMES = {
    "before": {"p512-k1-owned-repeated", "p512-k32-owned-repeated"},
    "after": {
        "p512-k1-owned-repeated",
        "p512-k1-owned-batch",
        "p512-k32-owned-repeated",
        "p512-k32-owned-batch",
    },
}
PERF_EVENTS = (
    "cycles",
    "instructions",
    "branches",
    "branch-misses",
    "cache-misses",
    "context-switches",
    "cpu-migrations",
    "page-faults",
)
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")
NAME_RE = re.compile(r"^p(128|512)-k(1|8|32)-(owned|file)-(repeated|batch)$")

BOOLEAN_KEYS = (
    "memory_released",
    "objects_released",
    "input_monotonic",
    "work_monotonic",
    "semantic_ok",
    "raw_untouched_ok",
    "raw_untouched_member_payloads_ok",
    "output_exact_ok",
    "managed_preflight_forward_ok",
    "managed_preflight_inverse_ok",
    "forward_ok",
    "inverse_ok",
    "unmanaged_preflight_forward_ok",
    "unmanaged_preflight_inverse_ok",
    "source_version_unchanged",
)
IDENTITY_KEYS = (
    "fixture_sha256",
    "fixture_bytes",
    "fixture_name",
    "expected_output_sha256",
    "expected_output_bytes",
    "output_sha256",
    "output_bytes",
)
TIMED_FIELDS = ("elapsed_ns", "edit_ns", "open_ns", "commit_ns", "publish_ns", "drop_ns")
CONSTANT_FIELDS = (
    "budget_after_work",
    "budget_after_input",
    "source_read_calls",
    "source_requested_bytes",
    "source_returned_bytes",
)


class VerificationError(RuntimeError):
    """A retained-evidence invariant failed."""


def require(condition, message):
    if not condition:
        raise VerificationError(message)


def load_json(path):
    require(path.is_file(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text())
    except (OSError, json.JSONDecodeError) as exc:
        raise VerificationError(f"cannot read JSON {path}: {exc}") from exc


def sha256(path):
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def read_csv(path):
    require(path.is_file(), f"missing CSV: {path}")
    with path.open(newline="") as stream:
        rows = list(csv.DictReader(stream))
    require(rows, f"empty CSV: {path}")
    return rows


def expected_child_names(modes):
    return {
        f"p{paragraphs}-k{count}-{source}-{mode}"
        for paragraphs in PARAGRAPHS
        for count in REPLACEMENTS
        for source in SOURCES
        for mode in modes
    }


def parse_name(name):
    match = NAME_RE.fullmatch(name)
    require(match is not None, f"unrecognised child name: {name}")
    return int(match.group(1)), int(match.group(2)), match.group(3), match.group(4)


def parse_bool(row, key, name):
    value = row.get(key)
    require(value in ("true", "false"), f"{name}: {key} is not boolean")
    return value == "true"


def validate_data_rows(name, rows, expected_mode=None):
    paragraphs, replacements, source, mode = parse_name(name)
    if expected_mode is not None:
        require(mode == expected_mode, f"{name}: unexpected mode")
    require(len(rows) == 66, f"{name}: expected 66 rows")
    required = set(IDENTITY_KEYS) | set(BOOLEAN_KEYS) | {
        "schema",
        "version",
        "api_path",
        "comparison_scope",
        "paragraphs",
        "replacements",
        "source",
        "mode",
        "repeat",
        "ordinal",
        "warmup",
        "elapsed_ns",
        "budget_before_memory",
        "budget_live_memory",
        "budget_after_memory",
        "budget_before_input",
        "budget_live_input",
        "budget_after_input",
        "budget_before_output",
        "budget_live_output",
        "budget_after_output",
        "budget_before_objects",
        "budget_live_objects",
        "budget_after_objects",
        "budget_before_work",
        "budget_live_work",
        "budget_after_work",
        "budget_managed",
    }
    require(required <= set(rows[0]), f"{name}: missing required columns")
    identities = set()
    warmup_identities = set()
    measured_identities = set()
    for row in rows:
        require(row["schema"] == "managed_paragraph_batch_perf_v1", f"{name}: schema")
        require(row["version"] == "1", f"{name}: schema version")
        expected_api_path = "replace_paragraph_text" if mode == "repeated" else "replace_body_paragraph_texts"
        require(row["api_path"] == expected_api_path, f"{name}: API path")
        require(int(row["paragraphs"]) == paragraphs and int(row["replacements"]) == replacements, f"{name}: fixture shape")
        require(row["source"] == source and row["mode"] == mode, f"{name}: route columns")
        repeat = int(row["repeat"])
        ordinal = int(row["ordinal"])
        warmup = row["warmup"]
        require(repeat in (0, 1) and warmup in ("true", "false"), f"{name}: sample identity")
        if warmup == "true":
            require(0 <= ordinal < 3, f"{name}: warmup ordinal")
            warmup_identities.add((repeat, ordinal))
        else:
            require(0 <= ordinal < 30, f"{name}: measured ordinal")
            measured_identities.add((repeat, ordinal))
        identity = tuple(row[key] for key in IDENTITY_KEYS)
        identities.add(identity)
        require(row["expected_output_sha256"] == row["output_sha256"], f"{name}: output digest differs from expected")
        require(row["expected_output_bytes"] == row["output_bytes"], f"{name}: output length differs from expected")
        for key in BOOLEAN_KEYS:
            require(parse_bool(row, key, name), f"{name}: {key} is false")
        require(row["budget_managed"] == "true", f"{name}: budget is not managed")
        for budget in ("memory", "objects"):
            require(int(row[f"budget_after_{budget}"]) == 0, f"{name}: released {budget} budget")
        for budget in ("input", "work"):
            before = int(row[f"budget_before_{budget}"])
            live = int(row[f"budget_live_{budget}"])
            after = int(row[f"budget_after_{budget}"])
            require(after >= live >= before, f"{name}: {budget} is not monotonic")
        require(int(row["output_bytes"]) >= 0, f"{name}: negative output length")
    require(warmup_identities == {(repeat, ordinal) for repeat in (0, 1) for ordinal in range(3)}, f"{name}: warmup identities")
    require(measured_identities == {(repeat, ordinal) for repeat in (0, 1) for ordinal in range(30)}, f"{name}: measured identities")
    require(len(identities) == 1, f"{name}: output identity changes within child")
    return {
        "name": name,
        "paragraphs": paragraphs,
        "replacements": replacements,
        "source": source,
        "mode": mode,
        "rows": rows,
        "measured": [row for row in rows if row["warmup"] == "false"],
        "identity": next(iter(identities)),
    }


def verify_receipt(phase, name, freeze_hash):
    folder = HERE / phase
    csv_path = folder / f"{name}.csv"
    receipt_path = folder / f"{name}.json"
    stderr_path = folder / f"{name}.time.stderr"
    stdout_path = folder / f"{name}.stdout"
    receipt = load_json(receipt_path)
    require(receipt.get("exit_code") == 0, f"{phase}/{name}: exit code")
    require(receipt.get("cleanup_verified") is True, f"{phase}/{name}: cleanup")
    require(receipt.get("binary_sha256") == freeze_hash, f"{phase}/{name}: binary hash")
    require(receipt.get("measured_samples") == 60 and receipt.get("warmup_samples") == 6, f"{phase}/{name}: receipt sample counts")
    hashes = {
        "csv": sha256(csv_path),
        "stderr": sha256(stderr_path),
        "stdout": sha256(stdout_path),
        "receipt": sha256(receipt_path),
    }
    require(hashes["csv"] == receipt.get("csv_sha256"), f"{phase}/{name}: CSV hash")
    require(hashes["stderr"] == receipt.get("stderr_sha256"), f"{phase}/{name}: stderr hash")
    data = validate_data_rows(name, read_csv(csv_path))
    return {**data, "receipt": receipt, "hashes": hashes, "stderr": stderr_path.read_text(errors="replace")}


def verify_phase(phase, modes, freeze_hash):
    folder = HERE / phase
    require(folder.is_dir(), f"missing phase directory: {phase}")
    names = expected_child_names(modes)
    require({path.stem for path in folder.glob("*.csv")} == names, f"{phase}: CSV inventory")
    require({path.stem for path in folder.glob("*.json")} == names, f"{phase}: receipt inventory")
    require({path.stem for path in folder.glob("*.stdout")} == names, f"{phase}: stdout inventory")
    require({path.name.removesuffix(".time.stderr") for path in folder.glob("*.time.stderr")} == names, f"{phase}: stderr inventory")
    return {name: verify_receipt(phase, name, freeze_hash) for name in sorted(names)}


def verify_route_identity(children):
    stems = {name.rsplit("-", 1)[0] for name in children}
    for stem in stems:
        routes = [children[f"{stem}-repeated"]]
        if f"{stem}-batch" in children:
            routes.append(children[f"{stem}-batch"])
        identities = {route["identity"] for route in routes}
        require(len(identities) == 1, f"{stem}: output identity differs across routes")


def verify_before_summary(children):
    summary_path = HERE / "before-summary.json"
    summary = load_json(summary_path)
    require(summary.get("children") == 12, "before summary: child count")
    require(summary.get("measured_samples") == 720 and summary.get("warmup_samples") == 72, "before summary: sample counts")
    expected = {}
    for name, child in children.items():
        values = sorted(int(row["elapsed_ns"]) for row in child["measured"])
        total = sum(values)
        edit = sum(int(row["edit_ns"]) for row in child["measured"])
        expected[name] = {
            "name": name,
            "p50_us": values[math.ceil(len(values) * 50 / 100) - 1] / 1000,
            "p95_us": values[math.ceil(len(values) * 95 / 100) - 1] / 1000,
            "p99_us": values[math.ceil(len(values) * 99 / 100) - 1] / 1000,
            "edit_share_of_total_timed_ns": edit / total,
            "ideal_upper_bound_if_edit_cost_zero": total / (total - edit),
            "charged_work": int(child["measured"][0]["budget_after_work"]),
        }
    observed = {row["name"]: row for row in summary.get("rows", [])}
    require(set(observed) == set(expected), "before summary: row inventory")
    for name, values in expected.items():
        for key, value in values.items():
            observed_value = observed[name][key]
            if isinstance(value, float):
                require(math.isclose(float(observed_value), value, rel_tol=1e-12), f"before summary {name}: {key}")
            else:
                require(observed_value == value, f"before summary {name}: {key}")
    return {"sha256": sha256(summary_path), "children": len(observed), "rows_recomputed": True}


def stats(rows):
    result = {}
    for field in TIMED_FIELDS:
        values = sorted(int(row[field]) for row in rows)
        prefix = field.removesuffix("_ns")
        result.update({f"{prefix}_p{p}_us": values[math.ceil(len(values) * p / 100) - 1] / 1000 for p in (50, 95, 99)})
        result[f"{prefix}_mean_us"] = sum(values) / len(values) / 1000
    result["throughput_output_bytes_s"] = sum(int(row["output_bytes"]) for row in rows) * 1e9 / sum(int(row["elapsed_ns"]) for row in rows)
    for field in CONSTANT_FIELDS:
        values = {int(row[field]) for row in rows}
        require(len(values) == 1, f"comparison: {field} varies within measured rows")
        result[field] = values.pop()
    return result


def delta(before, after):
    metrics = {}
    flags = []
    for key, value in before.items():
        pct = (after[key] / value - 1) * 100 if value else None
        metrics[key] = {"before": value, "after": after[key], "delta_pct": pct}
        if pct is not None:
            if (key.startswith(("elapsed_", "edit_")) or key == "rss_kib") and pct > 5:
                flags.append(key)
            if key == "throughput_output_bytes_s" and pct < -5:
                flags.append(key)
    return {"metrics": metrics, "adverse_flags": flags}


def assert_metric_equal(observed, expected, context):
    require(set(observed["metrics"]) == set(expected["metrics"]), f"{context}: metric inventory")
    for metric, values in expected["metrics"].items():
        for key in ("before", "after", "delta_pct"):
            actual = observed["metrics"][metric][key]
            wanted = values[key]
            if wanted is None:
                require(actual is None, f"{context}: {metric} {key}")
            else:
                require(math.isclose(float(actual), float(wanted), rel_tol=1e-12, abs_tol=1e-9), f"{context}: {metric} {key}")
    require(sorted(observed["adverse_flags"]) == sorted(expected["adverse_flags"]), f"{context}: adverse flags")


def verify_comparison(before, after):
    path = HERE / "comparison.json"
    comparison = load_json(path)
    require(comparison.get("children") == 36, "comparison: child count")
    require(comparison.get("measured_samples") == 2160 and comparison.get("warmup_samples") == 216, "comparison: sample counts")
    records = {(record["name"], record["kind"]): record for record in comparison.get("comparisons", [])}
    require(len(records) == 36, "comparison: duplicate records")
    expected_stems = {f"p{paragraphs}-k{count}-{source}" for paragraphs in PARAGRAPHS for count in REPLACEMENTS for source in SOURCES}
    require({name for name, _kind in records} == expected_stems, "comparison: workload inventory")
    require({kind for _name, kind in records} == set(KINDS), "comparison: comparison-kind inventory")
    route_pairs = {
        "scalar_before_after": ("before", "after", "repeated", "repeated"),
        "after_batch_vs_scalar": ("after", "after", "repeated", "batch"),
        "batch_after_vs_scalar_before": ("before", "after", "repeated", "batch"),
    }
    flags = {kind: {} for kind in KINDS}
    work_values = {}
    for stem in sorted(expected_stems):
        before_scalar = before[f"{stem}-repeated"]
        after_scalar = after[f"{stem}-repeated"]
        after_batch = after[f"{stem}-batch"]
        work_values[stem] = {
            "scalar_before": int(before_scalar["measured"][0]["budget_after_work"]),
            "scalar_after": int(after_scalar["measured"][0]["budget_after_work"]),
            "batch_after": int(after_batch["measured"][0]["budget_after_work"]),
        }
        routes = {"before": before_scalar, "after_scalar": after_scalar, "after_batch": after_batch}
        for kind in KINDS:
            left_phase, right_phase, left_mode, right_mode = route_pairs[kind]
            left = before_scalar if left_phase == "before" else (after_batch if left_mode == "batch" else after_scalar)
            right = before_scalar if right_phase == "before" else (after_batch if right_mode == "batch" else after_scalar)
            left_measured = left["measured"]
            right_measured = right["measured"]
            left_stats = stats(left_measured)
            right_stats = stats(right_measured)
            left_stats["rss_kib"] = int(RSS_RE.search(left["stderr"]).group(1))
            right_stats["rss_kib"] = int(RSS_RE.search(right["stderr"]).group(1))
            expected_aggregate = delta(left_stats, right_stats)
            observed = records[(stem, kind)]
            assert_metric_equal(observed["aggregate"], expected_aggregate, f"comparison {stem}/{kind} aggregate")
            for repeat in (0, 1):
                expected_repeat = delta(
                    stats([row for row in left_measured if int(row["repeat"]) == repeat]),
                    stats([row for row in right_measured if int(row["repeat"]) == repeat]),
                )
                assert_metric_equal(observed["repeats"][str(repeat)], expected_repeat, f"comparison {stem}/{kind} repeat {repeat}")
            for flag in expected_aggregate["adverse_flags"]:
                flags[kind][flag] = flags[kind].get(flag, 0) + 1
    require(comparison.get("aggregate_flags_by_comparison_kind") == flags, "comparison: flag counts")
    return {
        "sha256": sha256(path),
        "comparisons": len(records),
        "aggregate_flags_by_comparison_kind": flags,
        "work_by_route": work_values,
        "work_equality_not_assumed": True,
    }


def verify_profile_sample(name, path):
    rows = read_csv(path)
    data = validate_data_rows(name, rows)
    return {"sample_count": len(rows), "identity": data["identity"], "sha256": sha256(path)}


def verify_profiles(freezes, phases=("before", "after")):
    result = {}
    for phase in phases:
        folder = HERE / "profiles" / phase
        require(folder.is_dir(), f"missing profiles/{phase}")
        names = PROFILE_NAMES[phase]
        require({path.stem for path in folder.glob("*.json")} == names, f"profiles/{phase}: receipt inventory")
        require({path.name.removesuffix(".perf.csv") for path in folder.glob("*.perf.csv")} == names, f"profiles/{phase}: perf inventory")
        require({path.name.removesuffix(".samples.csv") for path in folder.glob("*.samples.csv")} == names, f"profiles/{phase}: sample inventory")
        phase_result = {}
        for name in sorted(names):
            receipt_path = folder / f"{name}.json"
            receipt = load_json(receipt_path)
            require(receipt.get("exit_code") == 0 and receipt.get("cleanup_verified") is True, f"profile {phase}/{name}: receipt")
            require(receipt.get("binary_sha256") == freezes[phase]["binary_sha256"], f"profile {phase}/{name}: binary")
            perf_path = folder / f"{name}.perf.csv"
            event_values = {}
            for line in perf_path.read_text(errors="replace").splitlines():
                if not line or line.startswith("#"):
                    continue
                values = next(csv.reader([line]))
                require(len(values) >= 4, f"profile {phase}/{name}: malformed perf row")
                event = values[2]
                require(event not in event_values, f"profile {phase}/{name}: duplicate event")
                require(values[0].isdigit(), f"profile {phase}/{name}: nonnumeric event count")
                event_values[event] = int(values[0])
            require(set(event_values) == set(PERF_EVENTS), f"profile {phase}/{name}: event inventory")
            sample = verify_profile_sample(name, folder / f"{name}.samples.csv")
            phase_result[name] = {
                "perf_sha256": sha256(perf_path),
                "samples_sha256": sample["sha256"],
                "sample_count": sample["sample_count"],
                "events": event_values,
                "receipt_sha256": sha256(receipt_path),
            }
        result[phase] = phase_result
    return result


def verify_freeze(phase):
    freeze_path = HERE / f"{phase}-freeze.json"
    source_path = HERE / f"{phase}-source.json"
    freeze = load_json(freeze_path)
    source = load_json(source_path)
    binary_path = Path(freeze["binary"])
    require(binary_path.is_file(), f"{phase}: frozen binary missing")
    binary_hash = sha256(binary_path)
    require(binary_hash == freeze["binary_sha256"], f"{phase}: frozen binary hash")
    require(sha256(ROOT / HARNESS) == freeze["harness_sha256"], f"{phase}: harness hash")
    if "source_receipt_sha256" in freeze:
        require(sha256(source_path) == freeze["source_receipt_sha256"], f"{phase}: source receipt hash")
    for key in ("candidate_source_sha256",):
        if key in freeze:
            require(sha256(source_path) == freeze[key], f"{phase}: source manifest hash")
    require(isinstance(source.get("files"), dict) and source["files"], f"{phase}: source file inventory")
    for path, expected in source["files"].items():
        require(re.fullmatch(r"[0-9a-f]{64}", expected) is not None, f"{phase}: malformed source hash {path}")
    if phase == "before":
        require(freeze.get("revision") == BASE_REVISION and source.get("revision") == BASE_REVISION, f"{phase}: base revision")
    else:
        require(freeze.get("base_revision", BASE_REVISION) == BASE_REVISION, f"{phase}: base revision")
        # The after source manifest is expected to describe the checked-out
        # candidate.  Historical before-source hashes cannot be checked
        # against the edited working tree, so only the after manifest is bound
        # to current bytes here.
        for path, expected in source["files"].items():
            require(sha256(ROOT / path) == expected, f"{phase}: source hash {path}")
    return {
        "sha256": sha256(freeze_path),
        "source_sha256": sha256(source_path),
        "binary": str(binary_path),
        "binary_sha256": binary_hash,
        "source_files": len(source["files"]),
    }


def verify_gates(allow_missing):
    commands = load_json(HERE / "gate-commands.json")
    checks = HERE / "checks"
    successful = {}
    missing = []
    attempts = {}
    for name, command in commands.items():
        receipt_path = checks / f"{name}.json"
        attempts[name] = len(list(checks.glob(f"{name}.attempt*.json")))
        if not receipt_path.is_file():
            missing.append(name)
            continue
        receipt = load_json(receipt_path)
        require(receipt.get("exit_code") == 0, f"gate {name}: failed current receipt")
        require(receipt.get("command") == ["taskset", "-c", "16-31", *command], f"gate {name}: command binding")
        log_path = checks / f"{name}.log"
        require(log_path.is_file(), f"gate {name}: missing log")
        log_hash = sha256(log_path)
        require(log_hash == receipt.get("log_sha256"), f"gate {name}: log hash")
        successful[name] = {
            "receipt": f"checks/{name}.json",
            "receipt_sha256": sha256(receipt_path),
            "log_sha256": log_hash,
            "attempts": attempts[name],
        }
    if missing and not allow_missing:
        raise VerificationError("missing gate receipts: " + ", ".join(missing))
    return {"successful": successful, "missing": missing, "attempts": attempts}


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--before-only", action="store_true", help="verify the available before evidence without after captures")
    parser.add_argument("--allow-missing-gates", action="store_true", help="record pending gates instead of requiring every receipt")
    parser.add_argument("--output", type=Path, default=HERE / "verification.json")
    args = parser.parse_args()

    freezes = {phase: verify_freeze(phase) for phase in ("before",) if (HERE / f"{phase}-freeze.json").is_file()}
    before = verify_phase("before", BEFORE_MODES, freezes["before"]["binary_sha256"])
    verify_route_identity(before)
    before_summary = verify_before_summary(before)
    if args.before_only:
        after = None
        comparison = None
        profiles = verify_profiles(freezes, phases=("before",))
        status = "before-only"
    else:
        freezes["after"] = verify_freeze("after")
        after = verify_phase("after", AFTER_MODES, freezes["after"]["binary_sha256"])
        verify_route_identity(after)
        for stem in {name.rsplit("-", 1)[0] for name in before}:
            require(before[f"{stem}-repeated"]["identity"] == after[f"{stem}-repeated"]["identity"], f"{stem}: scalar output identity")
            require(before[f"{stem}-repeated"]["identity"] == after[f"{stem}-batch"]["identity"], f"{stem}: batch output identity")
        comparison = verify_comparison(before, after)
        profiles = verify_profiles(freezes)
        status = "pass"
    gates = verify_gates(args.allow_missing_gates)
    if gates["missing"]:
        status = "pending-gates" if status == "pass" else status
    result = {
        "schema": "litchi-0500-independent-evidence-v1",
        "status": status,
        "frozen_phases": freezes,
        "before_children": len(before),
        "after_children": len(after) if after is not None else 0,
        "children": len(before) + (len(after) if after is not None else 0),
        "measured_samples": 2160 if after is not None else 720,
        "warmup_samples": 216 if after is not None else 72,
        "sample_identity": "repeat 0/1; warmup ordinal 0..2; measured ordinal 0..29",
        "output_identity_equal_across_routes": after is not None,
        "monotonic_work_input_and_released_budgets": True,
        "before_summary": before_summary,
        "comparison": comparison,
        "profiles": profiles,
        "raw_child_hashes": {
            "before": {name: child["hashes"] for name, child in sorted(before.items())},
            "after": {name: child["hashes"] for name, child in sorted(after.items())} if after is not None else {},
        },
        "gates": gates,
        "scope": "managed DOCX lifecycle and edit timings; warm synthetic selective workload; whole-child RSS/perf; shared host; no CRUD, cold-cache, remote-service, or native-producer claim",
    }
    args.output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"status": status, "before": len(before), "after": len(after) if after else 0, "missing_gates": gates["missing"]}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except (VerificationError, OSError, KeyError, ValueError) as exc:
        raise SystemExit(f"verification failed: {exc}") from exc
