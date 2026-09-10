#!/usr/bin/env python3
"""Independently verify the raw 0499 performance evidence.

This verifier reads only retained evidence and the frozen source files.  It
does not build, run a benchmark, or trust the summary values produced by the
capture/comparison scripts without recomputing them from the raw CSV and
stderr files.
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
PHASES = ("before", "after")
CORPORA = ("few-large", "many-small")
SOURCES = ("owned", "file", "instrumented")
APIS = (("serial", 1), ("batch", 1), ("batch", 2), ("batch", 4), ("batch", 8))
TRACE_CASES = (
    ("many-small", "serial", 1),
    ("many-small", "batch", 4),
    ("few-large", "batch", 4),
)
PROFILE_CASES = (("few-large", "owned", "batch", 4), ("many-small", "owned", "batch", 4))
GATE_CPUS = ("taskset", "-c", "16-31")
RSS_RE = re.compile(r"Maximum resident set size \(kbytes\):\s*(\d+)")

# These fields are operation and resource oracles.  Timing and derived
# percentile columns are deliberately excluded and are recomputed below.
ORACLE_KEYS = (
    "corpus_sha256",
    "logical_bytes",
    "digest",
    "source_calls",
    "source_requested_bytes",
    "source_returned_bytes",
    "source_max_request",
    "source_le_4k",
    "source_le_16k",
    "source_le_64k",
    "source_le_256k",
    "source_gt_256k",
    "cache_cold_loads",
    "cache_hits",
    "cache_retained_bytes",
    "cache_retained_entries",
    "cache_budget_memory_used",
    "cache_budget_objects_used",
    "budget_memory_after_read",
    "budget_memory_after_release",
    "budget_objects_after_read",
    "budget_objects_after_release",
    "budget_input_bytes",
    "budget_work",
    "accounting_deflated_read",
    "accounting_stored_read",
    "accounting_deflated_produced",
    "accounting_stored_accepted",
)
SUMMARY_ORACLE_KEYS = (
    "source_calls",
    "source_requested_bytes",
    "source_returned_bytes",
    "source_max_request",
    "cache_cold_loads",
    "cache_hits",
    "cache_retained_bytes",
    "budget_memory_after_release",
    "budget_objects_after_release",
)
SUMMARY_FIELD_NAMES = {
    "source_calls": "summary_source_calls",
    "source_requested_bytes": "summary_source_requested_bytes",
    "source_returned_bytes": "summary_source_returned_bytes",
    "source_max_request": "summary_source_max_request",
    "cache_cold_loads": "summary_cache_cold_loads",
    "cache_hits": "summary_cache_hits",
    "cache_retained_bytes": "summary_retained_bytes",
    "budget_memory_after_release": "summary_budget_memory_after_release",
    "budget_objects_after_release": "summary_budget_objects_after_release",
}


class VerificationError(RuntimeError):
    """A retained-evidence invariant was violated."""


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


def expected_name(corpus, source, api, workers):
    return f"{corpus}-{source}-{api}-w{workers}"


def expected_names():
    return {
        expected_name(corpus, source, api, workers)
        for corpus in CORPORA
        for source in SOURCES
        for api, workers in APIS
    }


def expected_parts(name):
    for corpus in CORPORA:
        prefix = corpus + "-"
        if not name.startswith(prefix):
            continue
        remainder = name[len(prefix):]
        for source in SOURCES:
            source_prefix = source + "-"
            if remainder.startswith(source_prefix):
                api_workers = remainder[len(source_prefix):]
                for api, workers in APIS:
                    if api_workers == f"{api}-w{workers}":
                        return corpus, source, api, workers
    raise VerificationError(f"unrecognised child name: {name}")


def source_value(source):
    return {"owned": "owned", "file": "file", "instrumented": "instrumented-short-delay"}[source]


def api_value(api):
    return {"serial": "serial-part-data", "batch": "read-parts-ordered"}[api]


def phase_paths(phase):
    directory = HERE / phase
    require(directory.is_dir(), f"missing phase directory: {directory}")
    names = expected_names()
    require({p.stem for p in directory.glob("*.csv")} == names, f"{phase}: CSV inventory mismatch")
    require({p.stem for p in directory.glob("*.json")} == names, f"{phase}: receipt inventory mismatch")
    require({p.stem for p in directory.glob("*.stdout")} == names, f"{phase}: stdout inventory mismatch")
    require(
        {p.name.removesuffix(".time.stderr") for p in directory.glob("*.time.stderr")} == names,
        f"{phase}: time-stderr inventory mismatch",
    )


def validate_summary(name, rows, samples, signature):
    summaries = [row for row in rows if row.get("record") == "summary"]
    require(len(summaries) == 1, f"{name}: expected one summary row")
    summary = summaries[0]
    require(summary.get("repeat", "") == "" and summary.get("index", "") == "", f"{name}: summary identity")
    measured = [row for row in samples if int(row["index"]) >= 3]
    nanoseconds = sorted(int(row["elapsed_ns"]) for row in measured)
    require(len(nanoseconds) == 60, f"{name}: measured sample count")
    expected = {
        "p50_ns": nanoseconds[math.ceil(len(nanoseconds) * 50 / 100) - 1],
        "p95_ns": nanoseconds[math.ceil(len(nanoseconds) * 95 / 100) - 1],
        "p99_ns": nanoseconds[math.ceil(len(nanoseconds) * 99 / 100) - 1],
        "mean_ns": int(sum(nanoseconds) / len(nanoseconds)),
    }
    for key, value in expected.items():
        require(int(summary[key]) == value, f"{name}: summary {key} differs from raw samples")
    throughput = sum(int(row["logical_bytes"]) for row in measured) * 1e9 / sum(nanoseconds)
    require(math.isclose(float(summary["throughput_bytes_s"]), throughput, rel_tol=1e-12), f"{name}: summary throughput")

    for key in SUMMARY_ORACLE_KEYS:
        summary_key = SUMMARY_FIELD_NAMES[key]
        value = signature[ORACLE_KEYS.index(key)]
        require(summary[summary_key] == value, f"{name}: summary oracle {summary_key}")


def load_child(phase, name, freeze_sha):
    corpus, source, api, workers = expected_parts(name)
    directory = HERE / phase
    csv_path = directory / f"{name}.csv"
    receipt_path = directory / f"{name}.json"
    stdout_path = directory / f"{name}.stdout"
    stderr_path = directory / f"{name}.time.stderr"
    receipt = load_json(receipt_path)
    require(receipt.get("exit_code") == 0, f"{name}: child exit code")
    require(receipt.get("cleanup_verified") is True, f"{name}: scratch cleanup")
    require(receipt.get("binary_sha256") == freeze_sha, f"{name}: binary receipt hash")
    require(receipt.get("measured_samples") == 60 and receipt.get("warmup_samples") == 6, f"{name}: receipt counts")
    hashes = {
        "csv": sha256(csv_path),
        "stderr": sha256(stderr_path),
        "stdout": sha256(stdout_path),
        "receipt": sha256(receipt_path),
    }
    require(hashes["csv"] == receipt.get("csv_sha256"), f"{name}: CSV hash")
    require(hashes["stderr"] == receipt.get("stderr_sha256"), f"{name}: stderr hash")

    rows = read_csv(csv_path)
    fields = set(rows[0])
    require(set(ORACLE_KEYS) <= fields, f"{name}: missing oracle columns")
    samples = [row for row in rows if row.get("record") == "sample"]
    require(len(samples) == 66, f"{name}: expected 66 sample rows")
    require(
        {(int(row["repeat"]), int(row["index"])) for row in samples}
        == {(repeat, index) for repeat in (0, 1) for index in range(33)},
        f"{name}: sample identities",
    )
    require(len(rows) == 67, f"{name}: unexpected CSV row count")
    for row in samples:
        require(row["corpus"] == corpus, f"{name}: corpus column")
        require(row["source"] == source_value(source), f"{name}: source column")
        require(row["api"] == api_value(api), f"{name}: API column")
        require(int(row["workers"]) == workers, f"{name}: worker column")
        require(int(row["budget_memory_after_release"]) == 0, f"{name}: released memory budget")
        require(int(row["budget_objects_after_release"]) == 0, f"{name}: released object budget")
        expected_cold_loads = 4 if corpus == "few-large" else 64
        require(int(row["cache_cold_loads"]) == expected_cold_loads, f"{name}: cold-load count")
    signatures = {tuple(row[key] for key in ORACLE_KEYS) for row in samples}
    require(len(signatures) == 1, f"{name}: sample oracle drift")
    signature = next(iter(signatures))
    semantic = [signature[ORACLE_KEYS.index(key)] for key in ("corpus_sha256", "logical_bytes", "digest")]
    require(receipt.get("semantic") == semantic, f"{name}: receipt semantic oracle")
    validate_summary(name, rows, samples, signature)
    return {
        "name": name,
        "corpus": corpus,
        "source": source,
        "api": api,
        "workers": workers,
        "rows": rows,
        "samples": samples,
        "signature": signature,
        "hashes": hashes,
        "receipt": receipt,
        "stderr": stderr_path.read_text(errors="replace"),
    }


def verify_freezes():
    candidate = load_json(HERE / "candidate-source.json")
    files = candidate.get("files")
    if isinstance(files, dict):
        source_files = files
    elif isinstance(files, list):
        source_files = {item["path"]: item["sha256"] for item in files}
    else:
        raise VerificationError("candidate-source.json: files must be a mapping or list")
    require(source_files, "candidate-source.json: empty files")
    source_hashes = {}
    for path, expected in source_files.items():
        file_path = ROOT / path
        require(file_path.is_file(), f"candidate source missing: {path}")
        actual = sha256(file_path)
        require(actual == expected, f"candidate source hash: {path}")
        source_hashes[path] = actual
    harness_path = "crates/litchi-opc/examples/source_backed_batch_perf.rs"
    require(harness_path in source_files, "candidate source does not freeze benchmark harness")
    freezes = {}
    for phase in PHASES:
        freeze_path = HERE / f"{phase}-freeze.json"
        freeze = load_json(freeze_path)
        binary_path = Path(freeze["binary"])
        require(binary_path.is_file(), f"{phase}: frozen binary missing")
        binary_hash = sha256(binary_path)
        require(binary_hash == freeze["binary_sha256"], f"{phase}: frozen binary hash")
        if phase == "before":
            require(freeze.get("revision") == "b27718fe8", "before: unexpected revision")
            require(freeze.get("harness_sha256") == source_hashes[harness_path], "before: harness hash")
        else:
            require(freeze.get("base_revision") == candidate.get("base_revision"), "after: base revision")
            require(freeze.get("candidate_source_sha256") == sha256(HERE / "candidate-source.json"), "after: source manifest hash")
            require(freeze.get("harness_sha256") == source_hashes[harness_path], "after: harness hash")
            build_receipt = HERE / freeze.get("build_receipt", "checks/benchmark-build.json")
            build = load_json(build_receipt)
            require(build.get("exit_code") == 0, "after: benchmark build gate")
        freezes[phase] = {
            "path": str(freeze_path.relative_to(HERE)),
            "sha256": sha256(freeze_path),
            "binary": str(binary_path),
            "binary_sha256": binary_hash,
        }
    return candidate, source_hashes, freezes


def parse_rss(stderr, name):
    match = RSS_RE.search(stderr)
    require(match is not None, f"{name}: missing maximum RSS")
    return int(match.group(1))


def stats(rows):
    nanoseconds = sorted(int(row["elapsed_ns"]) for row in rows)
    require(nanoseconds, "empty timing set")
    return {
        "p50_us": nanoseconds[math.ceil(len(nanoseconds) * 50 / 100) - 1] / 1000,
        "p95_us": nanoseconds[math.ceil(len(nanoseconds) * 95 / 100) - 1] / 1000,
        "p99_us": nanoseconds[math.ceil(len(nanoseconds) * 99 / 100) - 1] / 1000,
        "mean_us": sum(nanoseconds) / len(nanoseconds) / 1000,
        "throughput_bytes_s": sum(int(row["logical_bytes"]) for row in rows) * 1e9 / sum(nanoseconds),
    }


def compare_stats(before, after):
    metrics = {}
    flags = []
    for key, before_value in before.items():
        after_value = after[key]
        delta = (after_value / before_value - 1) * 100
        metrics[key] = {"before": before_value, "after": after_value, "delta_pct": delta}
        if (delta > 5 and key != "throughput_bytes_s") or (delta < -5 and key == "throughput_bytes_s"):
            flags.append(key)
    return {"metrics": metrics, "adverse_flags": flags}


def close(a, b):
    return math.isclose(float(a), float(b), rel_tol=1e-12, abs_tol=1e-9)


def verify_comparison(children):
    comparison = load_json(HERE / "comparison.json")
    require(comparison.get("children") == 60, "comparison: child count")
    require(comparison.get("measured_samples") == 3600, "comparison: measured count")
    require(comparison.get("warmup_samples") == 360, "comparison: warmup count")
    require(comparison.get("byte_counter_budget_oracles_equal") is True, "comparison: oracle flag")
    records = {record["name"]: record for record in comparison.get("comparisons", [])}
    require(set(records) == expected_names(), "comparison: name inventory")
    aggregate_counts = {}
    repeat_counts = {}
    for name in sorted(expected_names()):
        record = records[name]
        before = children["before"][name]
        after = children["after"][name]
        before_measured = [row for row in before["samples"] if int(row["index"]) >= 3]
        after_measured = [row for row in after["samples"] if int(row["index"]) >= 3]
        expected_aggregate = compare_stats(stats(before_measured) | {"rss_kib": parse_rss(before["stderr"], name)}, stats(after_measured) | {"rss_kib": parse_rss(after["stderr"], name)})
        observed_aggregate = record["aggregate"]
        for metric, values in expected_aggregate["metrics"].items():
            observed = observed_aggregate["metrics"][metric]
            for key in ("before", "after", "delta_pct"):
                require(close(observed[key], values[key]), f"comparison {name}: aggregate {metric} {key}")
        require(sorted(observed_aggregate["adverse_flags"]) == sorted(expected_aggregate["adverse_flags"]), f"comparison {name}: aggregate flags")
        for flag in expected_aggregate["adverse_flags"]:
            aggregate_counts[flag] = aggregate_counts.get(flag, 0) + 1
        for repeat in (0, 1):
            before_repeat = [row for row in before_measured if int(row["repeat"]) == repeat]
            after_repeat = [row for row in after_measured if int(row["repeat"]) == repeat]
            expected_repeat = compare_stats(stats(before_repeat), stats(after_repeat))
            observed_repeat = record["repeats"][str(repeat)]
            for metric, values in expected_repeat["metrics"].items():
                observed = observed_repeat["metrics"][metric]
                for key in ("before", "after", "delta_pct"):
                    require(close(observed[key], values[key]), f"comparison {name}: repeat {repeat} {metric} {key}")
            require(sorted(observed_repeat["adverse_flags"]) == sorted(expected_repeat["adverse_flags"]), f"comparison {name}: repeat {repeat} flags")
            for flag in expected_repeat["adverse_flags"]:
                repeat_counts[flag] = repeat_counts.get(flag, 0) + 1
    require(comparison.get("aggregate_flag_counts") == aggregate_counts, "comparison: aggregate flag counts")
    require(comparison.get("repeat_flag_counts") == repeat_counts, "comparison: repeat flag counts")
    return {
        "children": len(records),
        "aggregate_flag_counts": aggregate_counts,
        "repeat_flag_counts": repeat_counts,
        "comparison_sha256": sha256(HERE / "comparison.json"),
    }


def verify_gates(allow_missing):
    commands = load_json(HERE / "gate-commands.json")
    checks = HERE / "checks"
    successful = {}
    missing = []
    attempts = {}
    for name, command in commands.items():
        receipt_path = checks / f"{name}.json"
        attempt_paths = sorted(checks.glob(f"{name}.attempt*.json"))
        attempts[name] = len(attempt_paths)
        if not receipt_path.is_file():
            missing.append(name)
            continue
        receipt = load_json(receipt_path)
        require(receipt.get("exit_code") == 0, f"gate {name}: current receipt failed")
        expected_command = list(GATE_CPUS) + command
        require(receipt.get("command") == expected_command, f"gate {name}: command binding")
        log_path = checks / f"{name}.log"
        require(log_path.is_file(), f"gate {name}: missing log")
        log_hash = sha256(log_path)
        require(log_hash == receipt.get("log_sha256"), f"gate {name}: log hash")
        successful[name] = {"receipt": f"checks/{name}.json", "log_sha256": log_hash, "attempts": len(attempt_paths)}
    if missing and not allow_missing:
        raise VerificationError("missing gate receipts: " + ", ".join(missing))
    return {"successful": successful, "missing": missing, "attempts": attempts}


def parse_trace(path):
    calls = {}
    errors = {}
    for line in path.read_text(errors="replace").splitlines():
        parts = line.split()
        if not parts or parts[-1] not in {"clone", "clone3"}:
            continue
        require(len(parts) >= 5, f"trace {path}: malformed syscall row")
        syscall = parts[-1]
        numeric = parts[:-1]
        require(len(numeric) in {4, 5}, f"trace {path}: malformed columns")
        call_count = int(numeric[3])
        error_count = int(numeric[4]) if len(numeric) == 5 else 0
        calls[syscall] = calls.get(syscall, 0) + call_count
        errors[syscall] = errors.get(syscall, 0) + error_count
    return calls, errors


def verify_traces(freezes):
    expected_counts = {
        "before": {"many-small-serial-w1": 0, "many-small-batch-w4": 64, "few-large-batch-w4": 4},
        "after": {"many-small-serial-w1": 0, "many-small-batch-w4": 4, "few-large-batch-w4": 4},
    }
    result = {}
    for phase in PHASES:
        phase_result = {}
        directory = HERE / "thread-traces" / phase
        for corpus, api, workers in TRACE_CASES:
            name = f"{corpus}-{api}-w{workers}"
            receipt = load_json(directory / f"{name}.json")
            require(receipt.get("exit_code") == 0 and receipt.get("cleanup_verified") is True, f"trace {phase}/{name}: receipt")
            require(receipt.get("binary_sha256") == freezes[phase]["binary_sha256"], f"trace {phase}/{name}: binary hash")
            trace_path = directory / f"{name}.trace"
            calls, error_counts = parse_trace(trace_path)
            require(not error_counts or sum(error_counts.values()) == 0, f"trace {phase}/{name}: clone errors")
            total = sum(calls.values())
            require(total == expected_counts[phase][name], f"trace {phase}/{name}: expected {expected_counts[phase][name]} calls, found {total}")
            phase_result[name] = {"clone_calls": calls, "clone_errors": error_counts, "trace_sha256": sha256(trace_path)}
        result[phase] = phase_result
    return result


def verify_profiles(freezes):
    events = ("cycles", "instructions", "branches", "branch-misses", "cache-misses", "context-switches", "cpu-migrations", "page-faults")
    result = {}
    for phase in PHASES:
        phase_result = {}
        directory = HERE / "profiles" / phase
        for corpus, source, api, workers in PROFILE_CASES:
            name = f"{corpus}-{source}-{api}-w{workers}"
            receipt = load_json(directory / f"{name}.json")
            require(receipt.get("exit_code") == 0 and receipt.get("cleanup_verified") is True, f"profile {phase}/{name}: receipt")
            require(receipt.get("binary_sha256") == freezes[phase]["binary_sha256"], f"profile {phase}/{name}: binary hash")
            perf_path = directory / f"{name}.perf.csv"
            rows = {}
            for line in perf_path.read_text(errors="replace").splitlines():
                if not line or line.startswith("#"):
                    continue
                values = next(csv.reader([line]))
                require(len(values) >= 4, f"profile {phase}/{name}: malformed perf row")
                event = values[2]
                require(event not in rows, f"profile {phase}/{name}: duplicate event {event}")
                require(values[0].isdigit(), f"profile {phase}/{name}: nonnumeric {event} count")
                rows[event] = int(values[0])
            require(set(rows) == set(events), f"profile {phase}/{name}: event inventory")
            sample_path = directory / f"{name}.samples.csv"
            sample_rows = read_csv(sample_path)
            samples = [row for row in sample_rows if row.get("record") == "sample"]
            require(len(samples) == 66, f"profile {phase}/{name}: sample count")
            require({(int(row["repeat"]), int(row["index"])) for row in samples} == {(r, i) for r in (0, 1) for i in range(33)}, f"profile {phase}/{name}: sample identities")
            require(all(int(row["budget_memory_after_release"]) == 0 and int(row["budget_objects_after_release"]) == 0 for row in samples), f"profile {phase}/{name}: released budgets")
            expected_cold_loads = 4 if corpus == "few-large" else 64
            require(all(int(row["cache_cold_loads"]) == expected_cold_loads for row in samples), f"profile {phase}/{name}: cold loads")
            phase_result[name] = {
                "events": rows,
                "perf_sha256": sha256(perf_path),
                "samples_sha256": sha256(sample_path),
                "sample_count": len(samples),
            }
        result[phase] = phase_result
    return result


def verify_measurements(freezes):
    phase_data = {}
    for phase in PHASES:
        phase_paths(phase)
        phase_data[phase] = {name: load_child(phase, name, freezes[phase]["binary_sha256"]) for name in sorted(expected_names())}
    for name in sorted(expected_names()):
        before = phase_data["before"][name]
        after = phase_data["after"][name]
        require(before["signature"] == after["signature"], f"{name}: before/after oracle mismatch")
        require(
            {(int(row["repeat"]), int(row["index"])) for row in before["samples"]}
            == {(int(row["repeat"]), int(row["index"])) for row in after["samples"]},
            f"{name}: before/after sample identities",
        )
    return phase_data


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--allow-missing-gates", action="store_true", help="verify measurements while final gate receipts are still pending")
    args = parser.parse_args()
    candidate, source_hashes, freezes = verify_freezes()
    phase_data = verify_measurements(freezes)
    children = {phase: phase_data[phase] for phase in PHASES}
    comparison = verify_comparison(children)
    gates = verify_gates(args.allow_missing_gates)
    traces = verify_traces(freezes)
    profiles = verify_profiles(freezes)
    status = "pass" if not gates["missing"] else "pending-gates"
    result = {
        "schema": "litchi-0499-independent-evidence-v1",
        "status": status,
        "candidate_base_revision": candidate.get("base_revision"),
        "frozen_phases": freezes,
        "candidate_source_files": len(source_hashes),
        "children_per_phase": len(expected_names()),
        "sample_rows_per_child": 66,
        "measured_samples": 3600,
        "warmup_samples": 360,
        "exact_before_after_oracles": True,
        "raw_child_hashes": {
            phase: {name: phase_data[phase][name]["hashes"] for name in sorted(expected_names())} for phase in PHASES
        },
        "comparison": comparison,
        "gates": gates,
        "thread_traces": traces,
        "profiles": profiles,
        "scope": "warm synthetic selective OPC Part reads; whole-child RSS/profiles; shared host; no CRUD or remote-service claim",
    }
    (HERE / "verification.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print(json.dumps({"status": status, "children": 60, "aggregate_flags": comparison["aggregate_flag_counts"], "repeat_flags": comparison["repeat_flag_counts"], "missing_gates": gates["missing"]}, sort_keys=True))


if __name__ == "__main__":
    try:
        main()
    except (VerificationError, OSError, KeyError, ValueError) as exc:
        raise SystemExit(f"verification failed: {exc}") from exc
