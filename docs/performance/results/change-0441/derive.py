#!/usr/bin/env python3
"""Derive 0441 matched append evidence without starting any workload."""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import io
import json
import math
from pathlib import Path
import random
import statistics
import sys
import subprocess
from typing import Any


CHANGE = 441
ROLES = ("before", "after")
MODES = ("normal", "allocator")
SHAPES = ("tiny", "medium", "large")
PHASES = {"A1": ("before", "R1"), "B1": ("after", "R1"), "B2": ("after", "R2"), "A2": ("before", "R2")}
SAMPLES = 30
WARMUPS = 3
BOOTSTRAP = 2_000
SEED = 4_400_301
THRESHOLD = 5.0
MAX_BYTES = 512 * 1024 * 1024
HEX40 = r"[0-9a-fA-F]{40}"
HEX64 = r"[0-9a-fA-F]{64}"


class DeriveError(ValueError):
    pass


def fail(label: str, message: str) -> None:
    raise DeriveError(f"{label}: {message}")


def obj(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        fail(label, "expected object")
    return value


def parse(raw: bytes, label: str) -> Any:
    try:
        return json.loads(raw.decode("utf-8"), object_pairs_hook=_keys, parse_constant=_constant)
    except (UnicodeError, json.JSONDecodeError, DeriveError) as error:
        fail(label, f"invalid JSON: {error}")
    raise AssertionError("unreachable")


def _keys(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    value: dict[str, Any] = {}
    for key, item in pairs:
        if key in value:
            raise DeriveError(f"duplicate JSON key {key!r}")
        value[key] = item
    return value


def _constant(value: str) -> Any:
    raise DeriveError(f"non-finite JSON value {value!r}")


def load(path: Path, label: str) -> Any:
    try:
        raw = path.read_bytes()
    except OSError as error:
        fail(label, str(error))
    if len(raw) > MAX_BYTES:
        fail(label, "file exceeds bounded size")
    return parse(raw, label)


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def digest(value: Any, label: str) -> str:
    if not isinstance(value, str) or len(value) != 64 or any(char not in "0123456789abcdefABCDEF" for char in value):
        fail(label, "expected SHA-256")
    return value.lower()


def u64(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0 or value >= 1 << 64:
        fail(label, "expected u64")
    return value


def permutation(value: Any, label: str) -> list[int]:
    if not isinstance(value, list) or len(value) != SAMPLES or any(
        isinstance(item, bool) or not isinstance(item, int) for item in value
    ) or sorted(value) != list(range(SAMPLES)):
        fail(label, "expected a permutation of sample indices 0..29")
    return list(value)


def time_value(value: Any, label: str) -> dt.datetime:
    if not isinstance(value, str):
        fail(label, "expected timestamp")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        fail(label, str(error))
    if parsed.tzinfo is None:
        fail(label, "timestamp must include timezone")
    return parsed


def artifact(root: Path, value: Any, label: str) -> bytes:
    row = obj(value, label)
    name = row.get("path")
    if not isinstance(name, str) or not name or Path(name).is_absolute() or ".." in Path(name).parts:
        fail(label, "unsafe artifact path")
    path = (root / name).resolve()
    if not path.is_file() and Path(str(path) + ".gz").is_file():
        path = Path(str(path) + ".gz")
    if not path.is_file():
        fail(label, "artifact missing")
    stored = path.read_bytes()
    if len(stored) > MAX_BYTES:
        fail(label, "artifact exceeds bounded size")
    if path.suffix == ".gz":
        try:
            with gzip.GzipFile(fileobj=io.BytesIO(stored)) as stream:
                raw = stream.read(MAX_BYTES + 1)
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            fail(label, str(error))
    else:
        raw = stored
    if len(raw) > MAX_BYTES or row.get("bytes") != len(raw) or row.get("sha256") != sha(raw):
        fail(label, "artifact binding differs")
    return raw


def statistic(values: list[float], name: str) -> float:
    ordered = sorted(values)
    if name == "p50":
        return (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) / 2
    if name == "p95":
        return ordered[min(math.ceil(.95 * len(ordered)) - 1, len(ordered) - 1)]
    if name == "p99":
        return ordered[min(math.ceil(.99 * len(ordered)) - 1, len(ordered) - 1)]
    if name == "mean":
        return statistics.fmean(values)
    raise DeriveError(f"unknown statistic {name}")


def clean(value: float) -> int | float:
    return int(value) if value.is_integer() else value


def stats(values: list[int | float], label: str, bootstrap: bool = False) -> dict[str, Any]:
    if not values:
        fail(label, "empty vector")
    clean_values = [float(value) for value in values]
    result = {name: clean(statistic(clean_values, name)) for name in ("p50", "p95", "p99", "mean")}
    result.update({"count": len(values), "min": min(values), "max": max(values), "standard_deviation": clean(statistics.stdev(clean_values)) if len(values) > 1 else 0})
    if bootstrap:
        intervals: dict[str, Any] = {}
        for offset, name in enumerate(("p50", "p95", "p99", "mean")):
            rng = random.Random(SEED + offset)
            draws = [statistic([clean_values[rng.randrange(len(values))] for _ in values], name) for _ in range(BOOTSTRAP)]
            draws.sort()
            intervals[name] = {"lower": clean(draws[max(0, math.ceil(.025 * len(draws)) - 1)]), "upper": clean(draws[min(math.ceil(.975 * len(draws)) - 1, len(draws) - 1)]), "resamples": BOOTSTRAP, "seed": SEED + offset}
        result["bootstrap_95"] = intervals
    return result


def result_of(report: dict[str, Any], label: str) -> dict[str, Any]:
    values = report.get("results")
    if not isinstance(values, list) or len(values) != 1:
        fail(label, "expected one report result")
    return obj(values[0], f"{label}.results[0]")


def lane(root: Path, phase: str, role: str, mode: str, shape: str, repeat: str, attempt: str, protocol_sha: str) -> dict[str, Any]:
    name = f"{phase}-{role}-{mode}-{shape}-{repeat.lower()}"
    receipt_path = root / "runs" / phase / attempt / f"{name}-receipt.json"
    receipt = obj(load(receipt_path, str(receipt_path)), str(receipt_path))
    if receipt.get("schema") != "litchi-0441-capture-receipt-v1" or receipt.get("change") != CHANGE or receipt.get("status") != "pass" or receipt.get("exit_code") != 0 or receipt.get("oracle_exit_code") != 0:
        fail(str(receipt_path), "capture receipt is not passing")
    if receipt.get("phase") != phase or receipt.get("role") != role or receipt.get("selector") != "odp_existing_append_lifecycle" or receipt.get("source_field") != "odp_append":
        fail(str(receipt_path), "receipt lane identity differs")
    if receipt.get("protocol_sha256") != protocol_sha:
        fail(str(receipt_path), "protocol binding differs")
    if receipt.get('driver_sha256') != sha((root / 'capture.py').read_bytes()):
        fail(str(receipt_path), 'capture driver differs')
    build = obj(load(root / role / 'build.json', 'build'), 'build')
    if receipt.get('binary') != build.get('binaries', {}).get(mode) or receipt.get('revision') != build.get('revision'):
        fail(str(receipt_path), 'build binary/revision differs')
    if not (receipt.get('source_before') == receipt.get('source_after') == receipt.get('source_manifest') == build.get('source_manifest')):
        fail(str(receipt_path), 'build source differs')
    if receipt.get("source_unchanged") is not True or receipt.get("outside_bundle_status_unchanged") is not True:
        fail(str(receipt_path), "custody changed")
    started, finished = time_value(receipt.get("started_utc"), "started"), time_value(receipt.get("finished_utc"), "finished")
    if finished < started:
        fail(str(receipt_path), "interval reversed")
    artifacts = obj(receipt.get("artifacts"), f"{receipt_path}.artifacts")
    expected_artifacts = {"report", "catalog", "workload_log", "resource_log", "oracle_log"}
    if set(artifacts) != expected_artifacts:
        fail(str(receipt_path), "artifact inventory differs")
    for key, binding in artifacts.items():
        artifact(root, binding, f'{receipt_path}.{key}')
    oracle = subprocess.run([sys.executable, '-B', str(root / 'oracle/verify-report.py'), '--report', str(root / artifacts['report']['path']), '--mode', mode, '--shape', shape], capture_output=True, text=True)
    if oracle.returncode != 0 or oracle.stdout.strip() != 'VALID':
        fail(str(receipt_path), 'independent append oracle rejected report: ' + oracle.stderr)
    report = obj(parse(artifact(root, artifacts["report"], f"{receipt_path}.report"), "report"), "report")
    result = result_of(report, "report")
    if result.get("case") != "odp_existing_append_lifecycle" or obj(result.get("corpus"), "report.corpus").get("shape") != shape:
        fail("report", "selector or shape differs")
    configuration = obj(report.get("configuration"), "report.configuration")
    if configuration.get("samples_per_case") != SAMPLES or configuration.get("warmup_iterations_per_case") != WARMUPS or configuration.get("cases") != ["odp_existing_append_lifecycle"] or configuration.get("semantic_shapes") != [shape] or configuration.get("execution_workers") != [1]:
        fail("report.configuration", "known sample/selector/shape binding is missing")
    elapsed = obj(result.get("elapsed_ns"), "report.elapsed_ns")
    values = [u64(value, f"report.elapsed_ns.samples[{index}]") for index, value in enumerate(elapsed.get("samples", []))]
    sample_order = permutation(elapsed.get("sample_order"), "report.elapsed_ns.sample_order")
    if len(values) != SAMPLES:
        fail("report.elapsed_ns", "sample vector/order differs")
    report_binary = obj(report.get("binary_identity"), "report.binary_identity")
    receipt_binary = obj(receipt.get("binary"), f"{receipt_path}.binary")
    if any(report_binary.get(key) != receipt_binary.get(other) for key, other in (("path", "path"), ("binary_bytes", "bytes"), ("binary_sha256", "sha256"))):
        fail("report.binary_identity", "report binary differs from receipt binding")
    environment = obj(report.get("environment"), "report.environment")
    if environment.get("git_revision") != receipt.get("revision"):
        fail("report.environment.git_revision", "report revision differs from receipt")
    source = obj(obj(result.get("source"), "report.source").get("odp_append"), "report.source.odp_append")
    corpus = obj(result.get("corpus"), "report.corpus")
    identity = {"archive_sha256": digest(corpus.get("archive_sha256"), "corpus.archive_sha256"), "output_archive_sha256": digest(source.get("output_archive_sha256"), "source.output_archive_sha256"), "source_semantic_sha256": digest(source.get("source_semantic_sha256"), "source.source_semantic_sha256"), "output_semantic_sha256": digest(source.get("output_semantic_sha256"), "source.output_semantic_sha256"), "output_sha256": digest(result.get("output_sha256"), "result.output_sha256")}
    allocation = None
    if mode == "allocator":
        metrics = obj(obj(result.get("operation_metrics"), "operation_metrics").get("allocation"), "operation_metrics.allocation")
        if metrics.get("status") != "measured":
            fail("operation_metrics.allocation", "allocator vectors unavailable")
        names = ("allocation_calls", "deallocation_calls", "reallocation_calls", "failed_allocation_calls", "allocated_bytes", "deallocated_bytes", "live_bytes_before", "live_bytes_after", "region_peak_live_bytes")
        operation = obj(result.get("operation_metrics"), "operation_metrics")
        sample_indices = permutation(operation.get("sample_indices"), "operation_metrics.sample_indices")
        if sample_indices != sample_order:
            fail("allocation", "operation and elapsed sample order differs")
        vectors = {name: [u64(item, f"allocation.{name}") for item in obj(metrics.get(name), f"allocation.{name}").get("values", [])] for name in names}
        if any(len(values) != SAMPLES for values in vectors.values()):
            fail("allocation", "vector length differs")
        chronological = {
            name: [values[sample_indices.index(index)] for index in range(SAMPLES)]
            for name, values in vectors.items()
        }
        if any(value != 0 for value in chronological["failed_allocation_calls"]):
            fail("allocation.failed_allocation_calls", "allocation failure was recorded")
        balances = [
            before + allocated - deallocated - after
            for before, allocated, deallocated, after in zip(
                chronological["live_bytes_before"], chronological["allocated_bytes"],
                chronological["deallocated_bytes"], chronological["live_bytes_after"]
            )
        ]
        if any(value != 0 for value in balances):
            fail("allocation", "live-byte accounting does not balance")
        if any(
            peak < before or peak < after
            for peak, before, after in zip(
                chronological["region_peak_live_bytes"], chronological["live_bytes_before"], chronological["live_bytes_after"]
            )
        ):
            fail("allocation.region_peak_live_bytes", "region peak is below an endpoint")
        vectors['region_peak_above_entry'] = [peak - before for peak, before in zip(vectors['region_peak_live_bytes'], vectors['live_bytes_before'])]
        vectors['retained_live_delta'] = [after - before for after, before in zip(vectors['live_bytes_after'], vectors['live_bytes_before'])]
        chronological.update({name: [vectors[name][sample_indices.index(index)] for index in range(SAMPLES)] for name in ('region_peak_above_entry', 'retained_live_delta')})
        allocation = {
            "vectors": vectors, "chronological_vectors": chronological,
            "sample_indices": sample_indices,
            "statistics": {name: stats(values, f"allocation.{name}") for name, values in vectors.items()},
            "scope": metrics.get("scope"),
        }
    resource = artifact(root, artifacts["resource_log"], f"{receipt_path}.resource_log").decode("utf-8", errors="replace")
    rss = [int(line.split(":", 1)[1].strip()) * 1024 for line in resource.splitlines() if "Maximum resident set size (kbytes):" in line]
    if len(rss) != 1:
        fail("resource", "expected one maximum RSS")
    oracle_log = artifact(root, artifacts["oracle_log"], f"{receipt_path}.oracle_log").decode("utf-8", errors="replace")
    if not oracle_log.startswith("argv=") or "\nstdout=VALID" not in oracle_log:
        fail("oracle", "copied oracle did not leave the exact VALID marker")
    return {"phase": phase, "role": role, "mode": mode, "shape": shape, "repeat": repeat, "name": name, "started": started.isoformat(), "finished": finished.isoformat(), "elapsed_ns": stats(values, "elapsed_ns", bootstrap=mode == "normal"), "identity": identity, "process_memory": {"gnu_time_peak_rss_bytes": rss[0], "scope": "GNU time whole executable"}, "allocation": allocation or {"status": "unavailable"}, "binding": {"revision": receipt.get("revision"), "source_manifest": receipt.get("source_manifest"), "binary": {"path": receipt.get("binary", {}).get("path"), "bytes": receipt.get("binary", {}).get("bytes"), "sha256": receipt.get("binary", {}).get("sha256")}}}


def pct(before: float, after: float) -> float | None:
    return None if before == 0 else (after - before) * 100.0 / before


def compare(rows: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    by_key = {(row["role"], row["mode"], row["shape"], row["repeat"]): row for row in rows}
    pairs: list[dict[str, Any]] = []
    flags: list[dict[str, Any]] = []
    for mode in MODES:
        for shape in SHAPES:
            for repeat in ("R1", "R2"):
                before, after = by_key[("before", mode, shape, repeat)], by_key[("after", mode, shape, repeat)]
                metrics = [(f"elapsed_ns.{name}", before["elapsed_ns"].get(name), after["elapsed_ns"].get(name)) for name in ("p50", "p95", "p99", "mean")]
                metrics.append(("process_memory.gnu_time_peak_rss_bytes", before["process_memory"]["gnu_time_peak_rss_bytes"], after["process_memory"]["gnu_time_peak_rss_bytes"]))
                if mode == "allocator":
                    for name in ("allocation_calls", "allocated_bytes", "region_peak_live_bytes", "region_peak_above_entry", "retained_live_delta"):
                        metrics.append((f"allocation.{name}.p50", before["allocation"]["statistics"][name]["p50"], after["allocation"]["statistics"][name]["p50"]))
                for metric, old, new in metrics:
                    change = pct(float(old), float(new))
                    row = {"mode": mode, "shape": shape, "repeat": repeat, "metric": metric, "before": old, "after": new, "relative_percent": change, "flagged": change is not None and change > THRESHOLD, "scope": "matched descriptive comparison; every flag requires review"}
                    pairs.append(row)
                    if row["flagged"]:
                        flags.append(row)
    return pairs, flags


def repeats(rows):
    by_key = {(row['role'], row['mode'], row['shape'], row['repeat']): row for row in rows}
    result = []
    for role in ROLES:
        for mode in MODES:
            for shape in SHAPES:
                left, right = (by_key[(role, mode, shape, repeat)] for repeat in ('R1', 'R2'))
                metrics = [(f'elapsed_ns.{name}', left['elapsed_ns'][name], right['elapsed_ns'][name]) for name in ('p50', 'p95', 'p99', 'mean')]
                metrics.append(('rss_bytes', left['process_memory']['gnu_time_peak_rss_bytes'], right['process_memory']['gnu_time_peak_rss_bytes']))
                if mode == 'allocator':
                    metrics.extend((f'allocation.{name}.p50', left['allocation']['statistics'][name]['p50'], right['allocation']['statistics'][name]['p50']) for name in ('allocation_calls', 'allocated_bytes', 'region_peak_above_entry', 'retained_live_delta'))
                for metric, before, after in metrics:
                    change = pct(float(before), float(after))
                    result.append({'role':role,'mode':mode,'shape':shape,'metric':metric,'r1':before,'r2':after,'relative_percent':change,'flagged':change is not None and abs(change)>THRESHOLD})
    return result


def derive(root: Path, attempt: str) -> dict[str, Any]:
    protocol_path = root / "protocol.json"
    protocol = obj(load(protocol_path, "protocol.json"), "protocol.json")
    if protocol.get("change") != CHANGE or protocol.get("status") != "frozen" or protocol.get("samples") != SAMPLES or protocol.get("warmups") != WARMUPS or protocol.get("cpu") != 2 or protocol.get("workers") != 1:
        fail("protocol", "matrix binding differs")
    protocol_sha = sha(protocol_path.read_bytes())
    for key, name in (('verifier_sha256', 'verify-report.py'), ('protocol_sha256', 'protocol.json')):
        if sha((root / 'oracle' / name).read_bytes()) != protocol['oracle'][key]:
            fail('oracle', 'frozen oracle binding differs')
    rows = [lane(root, item['phase'], item['role'], item['mode'], item['shape'], item['repeat'], attempt, protocol_sha) for item in protocol['order']]
    if len(rows) != 24:
        fail("matrix", "expected 24 rows")
    for previous, current in zip(rows, rows[1:]):
        if time_value(previous['finished'], 'previous') > time_value(current['started'], 'current'):
            fail('matrix', 'capture order overlaps or differs')
    identity_fields = ("archive_sha256", "output_archive_sha256", "source_semantic_sha256", "output_semantic_sha256", "output_sha256")
    for mode in MODES:
        for shape in SHAPES:
            for repeat in ("R1", "R2"):
                pair = [row for row in rows if row["mode"] == mode and row["shape"] == shape and row["repeat"] == repeat]
                if len(pair) != 2 or any(pair[0]["identity"][field] != pair[1]["identity"][field] for field in identity_fields):
                    fail(f"identity.{mode}.{shape}.{repeat}", "before/after output or semantic identity differs")
    pairs, flags = compare(rows)
    repeat_rows = repeats(rows)
    gate_rows = [row for row in pairs if row["shape"] in {"medium", "large"} and row["repeat"] in {"R1", "R2"}]
    normal_gate = all(next(row["relative_percent"] for row in gate_rows if row["mode"] == "normal" and row["shape"] == shape and row["repeat"] == repeat and row["metric"] == "elapsed_ns.p50") <= -THRESHOLD for shape in ("medium", "large") for repeat in ("R1", "R2"))
    allocator_gate = all(any(row["relative_percent"] <= -THRESHOLD for row in gate_rows if row["mode"] == "allocator" and row["shape"] == shape and row["repeat"] == repeat and row["metric"] in ('allocation.allocated_bytes.p50', 'allocation.allocation_calls.p50')) for shape in ("medium", "large") for repeat in ("R1", "R2"))
    source_bindings = {(row["binding"]["revision"], json.dumps(row["binding"]["source_manifest"], sort_keys=True)) for row in rows}
    return {"schema": "litchi-0441-odp-append-abba-summary-v1", "change": CHANGE, "protocol_sha256": protocol_sha, "scope": "matched before/after owned append lifecycle; A1/B1/B2/A2 are serialized phase groups, so temporal grouping remains a limitation", "matrix": {"reports": 24, "retained_samples": 720, "samples": SAMPLES, "warmups": WARMUPS, "cpu": 2, "workers": 1}, "bindings": {"distinct_revision_source_pairs": len(source_bindings)}, "rows": rows, "comparisons": pairs, "review_flags": flags, "repeat_comparisons": repeat_rows, "repeat_flags": [row for row in repeat_rows if row["flagged"]], "acceptance": {"normal_p50_gate": normal_gate, "allocator_requested_or_call_gate": allocator_gate, "practical_gate_met": bool(normal_gate or allocator_gate), "manual_review_required": bool(flags), "threshold_percent": THRESHOLD}, "claims": ["no claim unless all identity/oracle gates pass", "no causal or universal memory claim from this matrix", "bootstrap intervals are descriptive and do not overcome temporal grouping"]}


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":"), ensure_ascii=False, allow_nan=False).encode("utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--root", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--attempt", default="formal")
    parser.add_argument("--output", type=Path)
    parser.add_argument("--write", action="store_true")
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    root = args.root.resolve()
    output = (args.output or root / "summary.json").resolve()
    try:
        value = derive(root, args.attempt)
        if args.check and (not output.is_file() or canonical(load(output, str(output))) != canonical(value)):
            fail(str(output), "retained summary differs")
        if args.write:
            output.write_bytes(canonical(value) + b"\n")
    except (OSError, KeyError, TypeError, ValueError, DeriveError) as error:
        print(f"INVALID: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
