#!/usr/bin/env python3
"""Validate and analyse the fixed 0716 DOCX baseline stationarity capture.

This packet is deliberately diagnostic.  It validates every retained child,
reconstructs acquisition order from ``elapsed_ns.sample_order``, and reports
within-child, between-child, and warmup variation.  It does not exclude a
sample, choose an acceptance threshold, or attribute a variation to hardware,
the scheduler, the allocator, or an implementation detail.
"""

from __future__ import annotations

import argparse
import datetime as datetime_module
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import statistics
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
LEGACY_PATH = REPO / "docs/performance/results/change-0709/analyze.py"
PAIR_HELPER_PATH = REPO / "docs/performance/results/change-0713/analyze.py"
HISTORICAL_PACKET = REPO / "docs/performance/results/change-0715"
HEX = set("0123456789abcdef")
WAIT4_FIELDS = (
    "ru_utime", "ru_stime", "ru_maxrss", "ru_minflt", "ru_majflt",
    "ru_inblock", "ru_oublock", "ru_nvcsw", "ru_nivcsw",
)
ALLOWED_CLEANUP_FILES = ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json")


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load helper: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


LEGACY = load_module(LEGACY_PATH, "ordinary_save_0709_for_0716")
PAIR_HELPER = load_module(PAIR_HELPER_PATH, "pair_helper_0713_for_0716")


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":")).encode()


def json_digest(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def write_new(path: Path, value: Any) -> bytes:
    require(not path.exists() and not path.is_symlink(), f"refusing to replace {path}")
    data = (json.dumps(value, indent=2) + "\n").encode()
    path.write_bytes(data)
    return data


def check_hex(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def float_delta(value: float, baseline: float) -> float:
    require(baseline > 0.0, "relative delta baseline must be positive")
    return (value / baseline - 1.0) * 100.0


def spread(values: list[float], label: str) -> float:
    require(values and all(value > 0.0 for value in values),
            f"{label} spread requires positive values")
    return (max(values) - min(values)) * 100.0 / min(values)


def load_plan() -> dict[str, Any]:
    plan = read(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema_version") == 1, "plan schema changed")
    require(plan.get("packet") == "change-0716-docx-baseline-stationarity",
            "packet identity changed")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("blocks") == 8 and plan.get("warmups") == [10, 100],
            "block/warmup plan changed")
    require(plan.get("samples") == 200 and plan.get("expected_children") == 32,
            "sample or child count changed")
    require(plan.get("analysis", {}).get("block_samples") == 50,
            "within-child block size changed")
    require(plan.get("analysis", {}).get("flag_percent") == 5,
            "diagnostic flag threshold changed")
    metrics = plan.get("analysis", {}).get("metrics")
    require(metrics == ["p50", "mean", "p95", "p99"], "analysis metrics changed")
    require(isinstance(plan.get("filesystem_root"), str)
            and plan["filesystem_root"], "filesystem root is missing")
    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and len(corpora) == 2, "corpus count changed")
    expected_ids = ["generated", "numbered-list"]
    require([item.get("id") for item in corpora] == expected_ids,
            "corpus order or identities changed")
    for corpus in corpora:
        require(isinstance(corpus, dict), "corpus entry is malformed")
        require(corpus.get("origin") in {
            "generated-harness-corpus", "caller-named-real-file",
        }, f"unsupported corpus origin: {corpus.get('origin')!r}")
        require(isinstance(corpus.get("expected_edit_admitted"), bool),
                f"{corpus.get('id')}: edit admission is missing")
        if corpus["origin"] == "generated-harness-corpus":
            require(corpus.get("path") is None and corpus.get("sha256") is None
                    and corpus.get("bytes") is None,
                    "generated corpus has fixture identity")
        else:
            require(isinstance(corpus.get("path"), str) and corpus["path"],
                    "real corpus path is missing")
            check_hex(corpus.get("sha256"), f"{corpus['id']}.sha256")
            positive_int(corpus.get("bytes"), f"{corpus['id']}.bytes")
    environment_keys = plan.get("environment_keys")
    expected_keys = [
        "LC_ALL", "LANG", "TZ", "RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
        "GLIBC_TUNABLES", "PERL_HASH_SEED", "PERL_PERTURB_KEYS",
    ]
    require(environment_keys == expected_keys, "environment key list changed")
    expected_environment = plan.get("expected_environment")
    require(isinstance(expected_environment, dict)
            and set(expected_environment) == set(expected_keys),
            "expected environment is malformed")
    overrides = plan.get("environment_overrides")
    require(overrides == {
        "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0",
    }, "environment overrides changed")
    context_paths = plan.get("context_paths")
    require(isinstance(context_paths, list) and context_paths
            and "/proc/stat" in context_paths,
            "context paths are incomplete")
    return plan


def validate_constraints() -> str:
    constraints_path = HERE / "constraints.json"
    constraints = read(constraints_path)
    require(isinstance(constraints, dict) and constraints, "constraints are missing")
    for name, digest in constraints.items():
        check_hex(digest, f"constraint {name}")
        path = REPO / name
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"constraint changed: {name}")
    return sha(constraints_path)


def validate_freeze(name: str) -> dict[str, str]:
    path = HERE / name
    value = read(path)
    require(isinstance(value, dict) and value, f"{name} is invalid")
    for raw, digest in value.items():
        check_hex(digest, f"{name}:{raw}")
        target = HERE / raw if "/" not in raw and not raw.startswith(".") else REPO / raw
        # The helper freeze names repository-relative paths; the capture freeze
        # names packet-local files.  Resolve both without accepting symlinks.
        if raw.startswith("docs/") or raw.startswith("tools/") or raw.startswith("crates/"):
            target = REPO / raw
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"{name} binding changed: {raw}")
    return value


def source_census() -> dict[str, str]:
    paths: list[Path] = [REPO / "Cargo.toml", REPO / "Cargo.lock"]
    paths.extend(path for path in (REPO / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path for path in (REPO / folder).rglob("*")
            if path.is_file() and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {str(path.relative_to(REPO)): sha(path) for path in sorted(set(paths))}


def validate_source() -> tuple[dict[str, str], str]:
    current = source_census()
    source_path = HERE / "source.json"
    source = read(source_path)
    require(isinstance(source, dict) and source, "source.json is invalid")
    require(current == source, "current checkout does not match source.json")
    historical_path = HISTORICAL_PACKET / "source-final.json"
    historical = read(historical_path)
    require(source == historical, "current source differs from 0715 source-final.json")
    return source, sha(source_path)


def fixture_binding(plan_corpus: dict[str, Any]) -> dict[str, Any] | None:
    if plan_corpus["origin"] == "generated-harness-corpus":
        return None
    raw = plan_corpus["path"]
    path = (REPO / raw).resolve() if not Path(raw).is_absolute() else Path(raw).resolve()
    require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
    digest = sha(path)
    require(digest == plan_corpus["sha256"] and path.stat().st_size == plan_corpus["bytes"],
            f"fixture identity changed: {plan_corpus['id']}")
    return {
        "path": raw,
        "bytes": path.stat().st_size,
        "sha256": digest,
    }


def jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    base = [(corpus, warmup) for corpus in plan["corpora"] for warmup in plan["warmups"]]
    result: list[dict[str, Any]] = []
    for block in range(plan["blocks"]):
        shift = block % len(base)
        for position, (corpus, warmup) in enumerate(base[shift:] + base[:shift]):
            prefix = ("docx_ordinary_save_" if corpus["id"] == "generated"
                      else "docx_real_file_ordinary_save_")
            result.append({
                "name": f"native-b{block + 1}-{corpus['id']}-w{warmup}",
                "block": block + 1,
                "position": position,
                "order_index": len(result),
                "corpus": corpus,
                "warmup": warmup,
                "samples": plan["samples"],
                "phase": "counting_publish",
                "case": prefix + "counting_publish",
            })
    require(len(result) == plan["expected_children"], "job count changed")
    return result


def expected_command(job: dict[str, Any], build: dict[str, Any], plan: dict[str, Any]) -> list[str]:
    argv = [
        "taskset", "-c", str(plan["cpu"]), build["binary"]["path"],
        "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
        "--case", job["case"], "--json", str(HERE / f"{job['name']}.json"),
        "--filesystem-root", plan["filesystem_root"],
    ]
    if job["corpus"]["path"]:
        argv += ["--ooxml-file", job["corpus"]["path"]]
    return argv


def walk_cleanup(value: Any, filename: str, output: list[dict[str, Any]]) -> None:
    if isinstance(value, dict):
        raw_path = value.get("path")
        digest = value.get("sha256", value.get("binary_sha256"))
        size = value.get("bytes", value.get("binary_bytes"))
        if isinstance(raw_path, str) and isinstance(digest, str):
            check_hex(digest, f"{filename}:{raw_path}")
            output.append({"path": raw_path, "sha256": digest, "bytes": size})
        for child in value.values():
            walk_cleanup(child, filename, output)
    elif isinstance(value, list):
        for child in value:
            walk_cleanup(child, filename, output)


def cleanup_witnesses() -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for filename in ALLOWED_CLEANUP_FILES:
        path = HERE / filename
        if path.is_file():
            walk_cleanup(read(path), filename, result)
    return result


def validate_binary_custody(build: dict[str, Any]) -> str:
    binary = build.get("binary")
    require(isinstance(binary, dict), "build binary record is missing")
    raw_path = binary.get("path")
    digest = binary.get("sha256")
    size = binary.get("bytes")
    require(isinstance(raw_path, str) and raw_path, "build binary path is missing")
    check_hex(digest, "build binary sha256")
    positive_int(size, "build binary bytes")
    path = Path(raw_path).resolve()
    if path.is_file() and not path.is_symlink():
        require(sha(path) == digest and path.stat().st_size == size,
                "live build binary identity changed")
        return "exact-live-binary-or-exact-post-cleanup-witness"
    for witness in cleanup_witnesses():
        candidate = Path(witness["path"])
        candidate = ((REPO / candidate).resolve() if not candidate.is_absolute()
                     else candidate.resolve())
        if (candidate == path and witness["sha256"] == digest
                and witness.get("bytes") == size):
            return "exact-live-binary-or-exact-post-cleanup-witness"
    fail("build binary is absent without an exact cleanup witness")


def validate_build(plan: dict[str, Any], source: dict[str, str]) -> tuple[dict[str, Any], str, str]:
    build_path = HERE / "build.json"
    build = read(build_path)
    require(isinstance(build, dict), "build.json is not an object")
    require(build.get("exit_code") == 0, "native build failed")
    require(build.get("source_sha256") == sha(HERE / "source.json"),
            "build/source manifest binding changed")
    log_sha = build.get("log_sha256")
    check_hex(log_sha, "build log sha256")
    log_path = HERE / "build.log"
    require(sha(log_path) == log_sha, "build log changed")
    command = build.get("command")
    expected_command_value = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", "litchi-perf-baseline",
        "--target-dir", str(REPO.parent / "litchi-target-0716"), "-j", "2",
    ]
    require(command == expected_command_value, "build command changed")
    custody = validate_binary_custody(build)
    binary = dict(build["binary"])
    build_info = {
        "binary": binary,
        "binary_sha256": binary["sha256"],
        "binary_bytes": binary["bytes"],
    }
    return build_info, custody, sha(build_path)


def expected_fixture_receipt(plan_corpus: dict[str, Any]) -> dict[str, Any] | None:
    binding = fixture_binding(plan_corpus)
    if binding is None:
        return None
    return binding


def validate_receipt(
    plan: dict[str, Any], job: dict[str, Any], build: dict[str, Any],
    source: dict[str, str], source_file_sha: str, build_file_sha: str,
    plan_file_sha: str, capture_file_sha: str, constraints_file_sha: str,
) -> tuple[dict[str, Any], dict[str, Any]]:
    name = job["name"]
    receipt_path = HERE / f"{name}.receipt.json"
    receipt = read(receipt_path)
    require(isinstance(receipt, dict), f"{name} receipt is not an object")
    require(receipt.get("job") == job, f"{name} job identity changed")
    require(receipt.get("exit_code") == 0, f"{name} child failed")
    require(receipt.get("command") == expected_command(job, build, plan),
            f"{name} argv changed")
    require(receipt.get("binary") == build["binary"], f"{name} binary binding changed")
    expected_source_digest = json_digest(source)
    require(receipt.get("source_before") == expected_source_digest
            and receipt.get("source_after") == expected_source_digest,
            f"{name} source custody digest changed")
    require(receipt.get("source_manifest_sha256") == source_file_sha,
            f"{name} source manifest binding changed")
    require(receipt.get("build_sha256") == build_file_sha,
            f"{name} build binding changed")
    require(receipt.get("plan_sha256") == plan_file_sha,
            f"{name} plan binding changed")
    require(receipt.get("script_sha256") == capture_file_sha,
            f"{name} capture script binding changed")
    environment = receipt.get("environment")
    require(environment == plan["expected_environment"],
            f"{name} environment binding changed")
    fixture = expected_fixture_receipt(job["corpus"])
    require(receipt.get("fixture_before") == fixture
            and receipt.get("fixture_after") == fixture,
            f"{name} fixture custody changed")
    seconds = receipt.get("seconds")
    finite(seconds, f"{name} child seconds")
    require(float(seconds) >= 0.0, f"{name} child seconds is negative")
    started = receipt.get("started_utc")
    require(isinstance(started, str) and started, f"{name} start timestamp is missing")
    try:
        datetime_module.datetime.fromisoformat(started)
    except ValueError:
        fail(f"{name} start timestamp is not ISO-8601")

    artifacts = receipt.get("artifacts")
    expected_names = {
        f"{name}.json", f"{name}.stdout", f"{name}.stderr", f"{name}.context.json",
    }
    require(isinstance(artifacts, dict) and set(artifacts) == expected_names,
            f"{name} artifact inventory changed")
    for filename, digest in artifacts.items():
        check_hex(digest, f"{name}:{filename}")
        path = HERE / filename
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"{name} artifact digest changed: {filename}")
    return receipt, read(HERE / f"{name}.json")


def parse_cpu_line(value: Any, cpu: int, label: str) -> list[int]:
    require(isinstance(value, dict) and isinstance(value.get("text"), str),
            f"{label} /proc/stat snapshot is unavailable")
    wanted = f"cpu{cpu} "
    line = next((row for row in value["text"].splitlines() if row.startswith(wanted)), None)
    require(line is not None, f"{label} CPU{cpu} counter line is missing")
    fields = line.split()
    require(fields[0] == f"cpu{cpu}" and len(fields) >= 2, f"{label} CPU line malformed")
    counters: list[int] = []
    for index, item in enumerate(fields[1:]):
        try:
            number = int(item)
        except ValueError:
            fail(f"{label} CPU counter {index} is not an integer")
        nonnegative_int(number, f"{label} CPU counter {index}")
        counters.append(number)
    require(counters, f"{label} CPU counter vector is empty")
    return counters


def validate_context(plan: dict[str, Any], job: dict[str, Any], receipt: dict[str, Any]) -> dict[str, Any]:
    name = job["name"]
    context_path = HERE / f"{name}.context.json"
    context = read(context_path)
    require(isinstance(context, dict), f"{name} context is not an object")
    require(set(context) == {"before", "after", "wait4", "scope"},
            f"{name} context fields changed")
    require(context.get("scope") == (
        "Whole child including setup, warmup, untimed opening, verification and reporting; "
        "system snapshots are not owner counters."
    ), f"{name} context scope changed")
    snapshots: dict[str, dict[str, Any]] = {}
    expected_snapshot_keys = {"monotonic_ns", *plan["context_paths"]}
    for phase in ("before", "after"):
        snapshot = context[phase]
        require(isinstance(snapshot, dict) and set(snapshot) == expected_snapshot_keys,
                f"{name} {phase} context keys changed")
        positive_int(snapshot["monotonic_ns"], f"{name} {phase} monotonic timestamp")
        for raw in plan["context_paths"]:
            value = snapshot[raw]
            require(isinstance(value, dict), f"{name} {phase} context {raw} malformed")
            if "text" in value:
                require(set(value) == {"text"} and isinstance(value["text"], str),
                        f"{name} {phase} context {raw} text malformed")
            else:
                require(set(value) == {"unavailable", "errno"}
                        and isinstance(value["unavailable"], str)
                        and isinstance(value["errno"], int),
                        f"{name} {phase} context {raw} unavailable malformed")
        snapshots[phase] = snapshot
    require(snapshots["after"]["monotonic_ns"] >= snapshots["before"]["monotonic_ns"],
            f"{name} context timestamps reversed")
    wait4 = context["wait4"]
    require(isinstance(wait4, dict) and set(wait4) == set(WAIT4_FIELDS),
            f"{name} wait4 fields changed")
    for field in WAIT4_FIELDS:
        value = wait4[field]
        finite(value, f"{name}.wait4.{field}")
        require(float(value) >= 0.0, f"{name}.wait4.{field} is negative")
    before_cpu = parse_cpu_line(snapshots["before"]["/proc/stat"], plan["cpu"],
                                f"{name} before")
    after_cpu = parse_cpu_line(snapshots["after"]["/proc/stat"], plan["cpu"],
                               f"{name} after")
    require(len(before_cpu) == len(after_cpu), f"{name} CPU counter width changed")
    deltas = [after - before for before, after in zip(before_cpu, after_cpu)]
    return {
        "artifact_sha256": sha(context_path),
        "before": snapshots["before"],
        "after": snapshots["after"],
        "wait4": wait4,
        "cpu_counter": {
            "cpu": plan["cpu"],
            "before": before_cpu,
            "after": after_cpu,
            "delta": deltas,
            "counter_units": "procfs jiffies; context snapshot delta only",
        },
        "causal_attribution": "none; whole-child counters do not isolate the measured serialization interval",
    }


def historical_report(corpus_id: str) -> tuple[Path, dict[str, Any]]:
    path = HISTORICAL_PACKET / f"baseline_A1-native-{corpus_id}-counting_publish.json"
    report = read(path)
    require(isinstance(report.get("results"), list) and len(report["results"]) == 1,
            f"historical report malformed: {path.name}")
    return path, report


def validate_report(
    plan: dict[str, Any], job: dict[str, Any], build: dict[str, Any], report: dict[str, Any],
) -> tuple[dict[str, Any], list[int], list[int], dict[str, Any]]:
    name = job["name"]
    LEGACY.check_report_metadata(report, {
        "binary_sha256": build["binary_sha256"], "binary_bytes": build["binary_bytes"],
    }, "native", job["case"], plan["samples"], job["warmup"], name)
    result = report["results"][0]
    require(result.get("case") == job["case"], f"{name} result case changed")
    elapsed, elapsed_expected = LEGACY.validate_elapsed(
        result.get("elapsed_ns"), plan["samples"], name)
    ordinary = LEGACY.validate_ordinary_save(
        result, job["corpus"], "counting_publish", "native", plan["samples"], name)
    LEGACY.validate_operation_metrics(
        result.get("operation_metrics"), result["elapsed_ns"], plan["samples"], "native", name)
    execution: list[int | None] = [None] * len(elapsed)
    order = result["elapsed_ns"]["sample_order"]
    require(len(order) == len(elapsed) and sorted(order) == list(range(len(elapsed))),
            f"{name} sample_order is not a complete permutation")
    for sorted_index, original_index in enumerate(order):
        execution[original_index] = elapsed[sorted_index]
    require(all(value is not None for value in execution), f"{name} execution reconstruction incomplete")
    execution_int = [int(value) for value in execution]
    historical_path, historical = historical_report(job["corpus"]["id"])
    historical_result = historical["results"][0]
    normalized = PAIR_HELPER.normalized_result(result)
    expected_normalized = PAIR_HELPER.normalized_result(historical_result)
    require(normalized == expected_normalized,
            f"{name} normalized output differs from 0715 baseline A1")
    return result, elapsed, execution_int, {
        "elapsed_stats": elapsed_expected,
        "normalized_sha256": json_digest(normalized),
        "historical_report": historical_path.name,
        "historical_report_sha256": sha(historical_path),
        "historical_normalized_sha256": json_digest(expected_normalized),
        "ordinary_published_sha256": ordinary["corpus"]["published_sha256"],
    }


def within_diagnostic(execution: list[int], plan: dict[str, Any]) -> dict[str, Any]:
    block_size = plan["analysis"]["block_samples"]
    require(len(execution) % block_size == 0 and len(execution) % 2 == 0,
            "sample count is incompatible with within-child diagnostic")
    block_means = [
        statistics.mean(execution[start:start + block_size])
        for start in range(0, len(execution), block_size)
    ]
    half = len(execution) // 2
    half_means = [statistics.mean(execution[:half]), statistics.mean(execution[half:])]
    threshold = float(plan["analysis"]["flag_percent"])
    return {
        "block_size": block_size,
        "block_means_ns": block_means,
        "block_span_percent": spread(block_means, "within block means"),
        "block_flag_over_5_percent": spread(block_means, "within block means") > threshold,
        "half_means_ns": half_means,
        "half_span_percent": spread(half_means, "within half means"),
        "half_flag_over_5_percent": spread(half_means, "within half means") > threshold,
        "interpretation": "descriptive relative spans; not a statistical stationarity test",
    }


def metric_stats(values: list[float], label: str) -> dict[str, Any]:
    value = spread(values, label)
    return {
        "values": values,
        "min": min(values),
        "max": max(values),
        "spread_percent": value,
        "flag_over_5_percent": value > 5.0,
    }


def group_diagnostics(rows: list[dict[str, Any]], plan: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    metrics = plan["analysis"]["metrics"]
    for corpus in plan["corpora"]:
        for warmup in plan["warmups"]:
            selected = [row for row in rows
                        if row["job"]["corpus"]["id"] == corpus["id"]
                        and row["job"]["warmup"] == warmup]
            require(len(selected) == plan["blocks"],
                    f"{corpus['id']}/w{warmup} does not have eight children")
            diagnostics = {}
            for metric in metrics:
                diagnostics[metric] = metric_stats(
                    [float(row["elapsed_stats"][metric]) for row in selected],
                    f"{corpus['id']}/w{warmup}/{metric}")
            result.append({
                "corpus_id": corpus["id"], "warmup": warmup,
                "repetitions": len(selected),
                "children": [row["job"]["name"] for row in selected],
                "metrics": diagnostics,
                "note": "Eight process-level summaries of in-process serialization samples; independence is not assumed, and spread is (max-min)/min.",
            })
    require(len(result) == 4, "group count changed")
    return result


def warmup_pairs(rows: list[dict[str, Any]], plan: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for block in range(1, plan["blocks"] + 1):
        for corpus in plan["corpora"]:
            selected = [row for row in rows
                        if row["job"]["block"] == block
                        and row["job"]["corpus"]["id"] == corpus["id"]]
            require(len(selected) == 2, f"block {block}/{corpus['id']} warmup pair is incomplete")
            by_warmup = {row["job"]["warmup"]: row for row in selected}
            require(set(by_warmup) == set(plan["warmups"]),
                    f"block {block}/{corpus['id']} warmup identities changed")
            cold = by_warmup[plan["warmups"][0]]
            warm = by_warmup[plan["warmups"][1]]
            metrics: dict[str, Any] = {}
            for metric in ("p50", "mean"):
                baseline = float(cold["elapsed_stats"][metric])
                observed = float(warm["elapsed_stats"][metric])
                metrics[metric] = {
                    "warmup_10": baseline, "warmup_100": observed,
                    "delta_percent": float_delta(observed, baseline),
                }
            result.append({
                "block": block, "corpus_id": corpus["id"],
                "warmup_10_child": cold["job"]["name"],
                "warmup_100_child": warm["job"]["name"],
                "warmup_10_position": cold["job"]["position"],
                "warmup_100_position": warm["job"]["position"],
                "warmup_10_order_index": cold["job"]["order_index"],
                "warmup_100_order_index": warm["job"]["order_index"],
                "metrics": metrics,
                "decision": "descriptive pair; no acceptance decision",
            })
    require(len(result) == 16, "warmup pair count changed")
    return result


def build_output(
    plan: dict[str, Any], source: dict[str, str], source_file_sha: str,
    build: dict[str, Any], build_custody: str, build_file_sha: str,
    constraints_file_sha: str, capture_file_sha: str, helper_freeze: dict[str, str],
    capture_freeze: dict[str, str], rows: list[dict[str, Any]],
) -> dict[str, Any]:
    groups = group_diagnostics(rows, plan)
    pairs = warmup_pairs(rows, plan)
    return {
        "schema_version": 1,
        "packet": plan["packet"],
        "revision": plan["revision"],
        "diagnostic_claim": (
            "Baseline-only stationarity observations. Whole-child wait4 and CPU12 counter "
            "deltas are retained as context and do not identify a cause for the measured interval."
        ),
        "scope": {
            "corpora": [corpus["id"] for corpus in plan["corpora"]],
            "phase": "counting_publish",
            "children": len(rows),
            "samples_per_child": plan["samples"],
            "cpu": plan["cpu"],
            "warmups": plan["warmups"],
            "blocks": plan["blocks"],
        },
        "rows": rows,
        "groups_by_corpus_warmup": groups,
        "warmup_pairs": pairs,
        "verification": {
            "current_source_matches_source_json": True,
            "current_source_matches_0715_source_final": True,
            "strict_receipts_verified": len(rows),
            "strict_0709_validators_used": [
                "check_report_metadata", "validate_elapsed",
                "validate_operation_metrics", "validate_ordinary_save",
            ],
            "normalized_parity_reference": "0715/baseline_A1-native-{corpus}-counting_publish.json",
            "normalized_parity_verified": True,
            "sample_order_reconstructed": True,
            "all_samples_retained": True,
            "within_child_flags_are_descriptive": True,
            "between_child_groups_are_descriptive": True,
            "warmup_pairs_are_descriptive": True,
            "whole_child_context_is_not_measured_interval": True,
            "causal_attribution": "none",
        },
        "binary_custody": {
            **build["binary"],
            "validation": "exact-live-binary-or-exact-post-cleanup-witness",
            "build_json_sha256": build_file_sha,
        },
        "source": {
            "manifest": "source.json",
            "sha256": source_file_sha,
            "entry_count": len(source),
            "historical_manifest": "../change-0715/source-final.json",
            "historical_manifest_sha256": sha(HISTORICAL_PACKET / "source-final.json"),
        },
        "freeze_bindings": {
            "constraints_json_sha256": constraints_file_sha,
            "capture_freeze": capture_freeze,
            "helper_freeze": helper_freeze,
            "capture_py_sha256": sha(HERE / "capture.py"),
            "plan_json_sha256": sha(HERE / "plan.json"),
            "analyzer_sha256": sha(HERE / "analyze.py"),
        },
        "limits": plan["limits"],
    }


def analyze() -> dict[str, Any]:
    """Run every validation and return the deterministic analysis object.

    Keeping this as a public, side-effect-free entry point lets the packet's
    refusal checks mutate one retained input at a time and prove that the
    analyzer rejects it without rewriting the evidence.
    """

    plan = load_plan()
    constraints_file_sha = validate_constraints()
    helper_freeze = validate_freeze("helper-freeze.json")
    capture_freeze = validate_freeze("capture-freeze.json")
    source, source_file_sha = validate_source()
    build, build_custody, build_file_sha = validate_build(plan, source)
    plan_file_sha = sha(HERE / "plan.json")
    capture_file_sha = sha(HERE / "capture.py")
    expected_jobs = jobs(plan)
    expected_receipts = {f"{job['name']}.receipt.json" for job in expected_jobs}
    actual_receipts = {path.name for path in HERE.glob("native-b*.receipt.json")}
    require(actual_receipts == expected_receipts,
            "receipt inventory is not exactly the frozen 32-child set")
    expected_child_files = set(expected_receipts)
    for job in expected_jobs:
        expected_child_files.update(
            f"{job['name']}.{suffix}"
            for suffix in ("json", "stdout", "stderr", "context.json")
        )
    actual_child_files = {path.name for path in HERE.glob("native-b*") if path.is_file()}
    require(actual_child_files == expected_child_files,
            "child artifact inventory is not exactly the frozen 32-child set")
    rows: list[dict[str, Any]] = []
    previous_context_after: int | None = None
    for job in expected_jobs:
        receipt, report = validate_receipt(
            plan, job, build, source, source_file_sha, build_file_sha,
            plan_file_sha, capture_file_sha, constraints_file_sha,
        )
        result, elapsed, execution, parity = validate_report(plan, job, build, report)
        context = validate_context(plan, job, receipt)
        before_ns = context["before"]["monotonic_ns"]
        after_ns = context["after"]["monotonic_ns"]
        if previous_context_after is not None:
            require(before_ns > previous_context_after,
                    f"{job['name']} context does not increase in job order")
        require(after_ns > before_ns, f"{job['name']} context interval is empty")
        previous_context_after = after_ns
        elapsed_stats = parity.pop("elapsed_stats")
        rows.append({
            "job": job,
            "receipt_sha256": sha(HERE / f"{job['name']}.receipt.json"),
            "report_sha256": sha(HERE / f"{job['name']}.json"),
            "artifact_sha256": receipt["artifacts"],
            "elapsed_samples_sorted_ns": elapsed,
            "sample_order": list(result["elapsed_ns"]["sample_order"]),
            "execution_samples_ns": execution,
            "elapsed_stats": elapsed_stats,
            "within_child": within_diagnostic(execution, plan),
            "normalized_parity": parity,
            "whole_child_context": context,
        })
    require(len(rows) == plan["expected_children"], "exact 32 rows were not retained")
    return build_output(
        plan, source, source_file_sha, build, build_custody, build_file_sha,
        constraints_file_sha, capture_file_sha, helper_freeze, capture_freeze, rows,
    )


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true",
                        help="recompute and compare without writing analysis.json")
    parser.add_argument("--output", default=str(HERE / "analysis.json"))
    args = parser.parse_args()
    try:
        output_path = Path(args.output)
        output = analyze()
        data = (json.dumps(output, indent=2) + "\n").encode()
        if args.check:
            require(output_path.is_file() and not output_path.is_symlink(),
                    f"missing analysis output: {output_path}")
            require(output_path.read_bytes() == data,
                    f"analysis output is not deterministic: {output_path}")
            print(f"PASS deterministic 0716 analysis check ({len(output['rows'])} rows)")
        else:
            require(output_path == HERE / "analysis.json",
                    "normal analysis output must be packet analysis.json")
            write_new(output_path, output)
            print(f"PASS verified {len(output['rows'])} children; wrote {output_path}")
    except (AssertionError, OSError, RuntimeError, ValueError, KeyError, TypeError) as error:
        print(f"analysis failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
