#!/usr/bin/env python3
"""Validate and describe the fixed 0717 DOCX process-counter capture.

This packet has two frozen executable lanes.  The native lane has no process
probe and the procfs lane has the opt-in ordinary-save probe.  The probe is
diagnostic evidence: its elapsed samples are never compared with native
elapsed samples, and its 32 empty controls are never subtracted from an
operation observation.

The module deliberately keeps the public ``read``, ``sha`` and ``analyze``
functions side-effect free.  ``--check`` recomputes the exact JSON output
without replacing it, which is used by the packet's refusal checks and final
audit.
"""

from __future__ import annotations

import argparse
import copy
import datetime as datetime_module
import hashlib
import importlib.util
import json
import math
from pathlib import Path
import statistics
from typing import Any


HERE = Path(__file__).resolve().parent
REPO = HERE.parents[3]
LEGACY_PATH = REPO / "docs/performance/results/change-0709/analyze.py"
PAIR_HELPER_PATH = REPO / "docs/performance/results/change-0713/analyze.py"
HISTORICAL_PACKET = REPO / "docs/performance/results/change-0716"
HEX = set("0123456789abcdef")
WAIT4_FIELDS = (
    "ru_utime", "ru_stime", "ru_maxrss", "ru_minflt", "ru_majflt",
    "ru_inblock", "ru_oublock", "ru_nvcsw", "ru_nivcsw",
)
LANES = ("native", "procfs")
CORPUS_IDS = ("generated", "numbered-list")
PROCESS_DELTA_FIELDS = (
    "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
    "syscr", "syscw", "minor_faults", "major_faults", "user_cpu_ticks",
    "system_cpu_ticks", "clock_ticks_per_second", "voluntary_context_switches",
    "nonvoluntary_context_switches", "rss_bytes", "peak_rss_bytes",
)
PROCESS_VECTOR_FIELDS = (
    "user_cpu_ticks", "system_cpu_ticks", "clock_ticks_per_second",
    "minor_faults", "major_faults", "voluntary_context_switches",
    "nonvoluntary_context_switches", "rss_delta_bytes", "peak_rss_bytes",
    "rchar", "wchar", "read_bytes", "write_bytes", "cancelled_write_bytes",
    "syscr", "syscw",
)
PROCESS_VECTOR_SOURCE_FIELDS = {
    "user_cpu_ticks": "user_cpu_ticks",
    "system_cpu_ticks": "system_cpu_ticks",
    "clock_ticks_per_second": "clock_ticks_per_second",
    "minor_faults": "minor_faults",
    "major_faults": "major_faults",
    "voluntary_context_switches": "voluntary_context_switches",
    "nonvoluntary_context_switches": "nonvoluntary_context_switches",
    "rss_delta_bytes": "rss_bytes",
    "peak_rss_bytes": "peak_rss_bytes",
    "rchar": "rchar",
    "wchar": "wchar",
    "read_bytes": "read_bytes",
    "write_bytes": "write_bytes",
    "cancelled_write_bytes": "cancelled_write_bytes",
    "syscr": "syscr",
    "syscw": "syscw",
}
PROCESS_VECTOR_UNITS = {
    "user_cpu_ticks": "procfs_clock_ticks",
    "system_cpu_ticks": "procfs_clock_ticks",
    "clock_ticks_per_second": "ticks_per_second",
    "minor_faults": "faults",
    "major_faults": "faults",
    "voluntary_context_switches": "context_switches",
    "nonvoluntary_context_switches": "context_switches",
    "rss_delta_bytes": "bytes",
    "peak_rss_bytes": "bytes",
    "rchar": "bytes",
    "wchar": "bytes",
    "read_bytes": "bytes",
    "write_bytes": "bytes",
    "cancelled_write_bytes": "bytes",
    "syscr": "system_calls",
    "syscw": "system_calls",
}
PROCESS_DELTA_UNITS = {
    **PROCESS_VECTOR_UNITS,
    "rss_bytes": "bytes",
}
PROCESS_VECTOR_SCOPES = {
    **{name: "procfs_in_process_operation_delta_including_procfs_probe_overhead"
       for name in (
           "user_cpu_ticks", "system_cpu_ticks", "clock_ticks_per_second",
           "minor_faults", "major_faults", "voluntary_context_switches",
           "nonvoluntary_context_switches", "rchar", "wchar", "read_bytes",
           "write_bytes", "cancelled_write_bytes", "syscr", "syscw",
       )},
    "rss_delta_bytes": "procfs_in_process_rss_delta_including_procfs_probe_overhead",
    "peak_rss_bytes": "process_lifetime_high_water_after_not_operation_peak",
}
EXPECTED_PROCESS_PROBE_SCOPE = (
    "ordinary_save_phase_interval_same_process_counters_including_procfs_probe_overhead"
)
EXPECTED_PROCESS_PROBE_LATENCY = "diagnostic_only_procfs_probe_instrumentation_latency"
EXPECTED_PROCESS_PROBE_CONTROL_SCOPE = (
    "fixed_32_empty_adjacent_procfs_snapshot_pairs_acquired_before_warmups_and_never_subtracted"
)
EXPECTED_PROCESS_SCOPE = "procfs_in_process_operation_delta_including_procfs_probe_overhead"
EXPECTED_PROCESS_IO_SCOPE = EXPECTED_PROCESS_SCOPE
EXPECTED_PROCESS_RSS_SCOPE = (
    "procfs_in_process_rss_delta_including_procfs_probe_overhead"
)
EXPECTED_PROCESS_HWM_SCOPE = "process_lifetime_high_water_after_not_operation_peak"
EXPECTED_ALIGNMENT = "elapsed_ns.samples_by_elapsed_then_sample_index"
ALLOWED_SOURCE_CHANGES = {
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/ordinary_save.rs",
}
ALLOWED_CLEANUP_FILES = ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json")


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load helper: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


LEGACY = load_module(LEGACY_PATH, "ordinary_save_0709_for_0717")
PAIR_HELPER = load_module(PAIR_HELPER_PATH, "pair_helper_0713_for_0717")


def fail(message: str) -> None:
    raise AssertionError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        fail(f"invalid JSON in {path}: {error}")


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical_json(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":")).encode()


def json_digest(value: Any) -> str:
    return hashlib.sha256(canonical_json(value)).hexdigest()


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


def load_plan() -> dict[str, Any]:
    plan = read(HERE / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    require(plan.get("schema_version") == 1, "plan schema changed")
    require(plan.get("packet") == "change-0717-docx-phase-process-probe",
            "packet identity changed")
    require(plan.get("cpu") == 12, "CPU binding changed")
    require(plan.get("blocks") == 16 and plan.get("warmup") == 100,
            "block/warmup plan changed")
    require(plan.get("samples") == 200 and plan.get("expected_children") == 64,
            "sample or child count changed")
    require(plan.get("lanes") == list(LANES), "lane order changed")
    analysis = plan.get("analysis")
    require(isinstance(analysis, dict), "analysis plan is missing")
    require(analysis.get("block_samples") == 50 and analysis.get("flag_percent") == 5,
            "analysis block or flag threshold changed")
    require(analysis.get("metrics") == ["p50", "mean", "p95", "p99"],
            "analysis metrics changed")
    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and [item.get("id") for item in corpora] == list(CORPUS_IDS),
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
    expected_keys = [
        "LC_ALL", "LANG", "TZ", "RUSTFLAGS", "LD_PRELOAD", "MALLOC_CONF",
        "GLIBC_TUNABLES", "PERL_HASH_SEED", "PERL_PERTURB_KEYS",
    ]
    require(plan.get("environment_keys") == expected_keys, "environment key list changed")
    require(isinstance(plan.get("expected_environment"), dict)
            and set(plan["expected_environment"]) == set(expected_keys),
            "expected environment is malformed")
    require(plan.get("environment_overrides") == {
        "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0",
    }, "environment overrides changed")
    context_paths = plan.get("context_paths")
    require(isinstance(context_paths, list) and context_paths == [
        "/proc/stat", "/proc/pressure/cpu", "/proc/pressure/io", "/proc/loadavg",
        "/proc/sys/kernel/perf_event_paranoid", "/proc/sys/kernel/randomize_va_space",
        "/sys/devices/system/cpu/cpu12/cpufreq/scaling_governor",
        "/sys/devices/system/cpu/cpu12/cpufreq/scaling_cur_freq",
    ], "context paths changed")
    require("process_probe" in analysis and isinstance(analysis["process_probe"], str),
            "process probe analysis description is missing")
    return plan


def validate_constraints() -> str:
    path = HERE / "constraints.json"
    constraints = read(path)
    require(isinstance(constraints, dict) and constraints, "constraints are missing")
    for name, digest in constraints.items():
        check_hex(digest, f"constraint {name}")
        target = REPO / name
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"constraint changed: {name}")
    return sha(path)


def validate_freeze(name: str) -> dict[str, str]:
    value = read(HERE / name)
    require(isinstance(value, dict) and value, f"{name} is invalid")
    for raw, digest in value.items():
        check_hex(digest, f"{name}:{raw}")
        if raw.startswith(("docs/", "tools/", "crates/")):
            target = REPO / raw
        else:
            target = HERE / raw
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"{name} binding changed: {raw}")
    return value


def validate_source() -> tuple[dict[str, str], str, dict[str, str], str]:
    baseline_path = HERE / "source-baseline.json"
    source_path = HERE / "source.json"
    baseline = read(baseline_path)
    source = read(source_path)
    require(isinstance(baseline, dict) and baseline, "source-baseline.json is invalid")
    require(isinstance(source, dict) and source, "source.json is invalid")
    current = source_census()
    require(current == source, "current checkout does not match source.json")
    changed = {name for name in set(source) | set(baseline)
               if source.get(name) != baseline.get(name)}
    require(changed == ALLOWED_SOURCE_CHANGES,
            f"source changed outside the three diagnostic files: {sorted(changed)}")
    return source, sha(source_path), baseline, sha(baseline_path)


def fixture_binding(corpus: dict[str, Any]) -> dict[str, Any] | None:
    if corpus["origin"] == "generated-harness-corpus":
        return None
    raw = corpus["path"]
    path = (REPO / raw).resolve() if not Path(raw).is_absolute() else Path(raw).resolve()
    require(path.is_file() and not path.is_symlink(), f"missing fixture: {path}")
    digest = sha(path)
    require(digest == corpus["sha256"] and path.stat().st_size == corpus["bytes"],
            f"fixture identity changed: {corpus['id']}")
    return {"path": raw, "bytes": path.stat().st_size, "sha256": digest}


def jobs(plan: dict[str, Any]) -> list[dict[str, Any]]:
    base = [(corpus, lane) for corpus in plan["corpora"] for lane in plan["lanes"]]
    result: list[dict[str, Any]] = []
    for block in range(plan["blocks"]):
        shift = block % len(base)
        for position, (corpus, lane) in enumerate(base[shift:] + base[:shift]):
            prefix = ("docx_ordinary_save_" if corpus["id"] == "generated"
                      else "docx_real_file_ordinary_save_")
            result.append({
                "name": f"{lane}-b{block + 1}-{corpus['id']}",
                "lane": lane,
                "block": block + 1,
                "position": position,
                "order_index": len(result),
                "corpus": corpus,
                "warmup": plan["warmup"],
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


def cleanup_witnesses() -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for filename in ALLOWED_CLEANUP_FILES:
        path = HERE / filename
        if not path.is_file():
            continue
        value = read(path)

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw_path, str) and isinstance(digest, str):
                    check_hex(digest, f"{filename}:{raw_path}")
                    result.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return result


def validate_binary(binary: dict[str, Any]) -> str:
    require(isinstance(binary, dict), "build binary record is missing")
    raw_path, digest, size = binary.get("path"), binary.get("sha256"), binary.get("bytes")
    require(isinstance(raw_path, str) and raw_path, "build binary path is missing")
    check_hex(digest, "build binary sha256")
    positive_int(size, "build binary bytes")
    path = Path(raw_path).resolve()
    if path.is_file() and not path.is_symlink():
        require(sha(path) == digest and path.stat().st_size == size,
                f"live build binary identity changed: {path}")
        return "exact-binary-identity-validated"
    for witness in cleanup_witnesses():
        candidate = Path(witness["path"])
        candidate = ((REPO / candidate).resolve() if not candidate.is_absolute()
                     else candidate.resolve())
        if candidate == path and witness["sha256"] == digest \
                and witness.get("bytes") == size:
            return "exact-binary-identity-validated"
    fail(f"build binary is absent without an exact cleanup witness: {path}")


def validate_builds(plan: dict[str, Any], source_file_sha: str) -> tuple[dict[str, Any], str]:
    path = HERE / "builds.json"
    builds = read(path)
    require(isinstance(builds, dict) and set(builds) == set(LANES),
            "builds.json must contain exactly native and procfs lanes")
    result: dict[str, Any] = {}
    target_dir = str(REPO.parent / "litchi-target-0717")
    for lane in LANES:
        row = builds[lane]
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"{lane} build failed")
        require(row.get("source_sha256") == source_file_sha,
                f"{lane} build/source binding changed")
        expected = [
            "cargo", "build", "--release", "--locked", "--manifest-path",
            "tools/perf-baseline/Cargo.toml", "--bin", "litchi-perf-baseline",
            "--target-dir", target_dir, "-j", "2",
        ]
        if lane == "procfs":
            expected += ["--features", "ordinary-save-process-metrics"]
        require(row.get("command") == expected, f"{lane} build command changed")
        log_name, log_digest = row.get("log"), row.get("log_sha256")
        require(isinstance(log_name, str) and log_name == f"build-{lane}.log",
                f"{lane} build log name changed")
        check_hex(log_digest, f"{lane} build log sha256")
        require(sha(HERE / log_name) == log_digest, f"{lane} build log changed")
        custody = validate_binary(row.get("binary"))
        result[lane] = {
            "command": row["command"],
            "exit_code": row["exit_code"],
            "source_sha256": row["source_sha256"],
            "log": log_name,
            "log_sha256": log_digest,
            "binary": row["binary"],
            "binary_custody": custody,
        }
    return result, sha(path)


def validate_receipt(
    plan: dict[str, Any], job: dict[str, Any], build: dict[str, Any], source_file_sha: str,
    builds_file_sha: str, plan_file_sha: str, capture_file_sha: str,
) -> tuple[dict[str, Any], dict[str, Any]]:
    name = job["name"]
    receipt = read(HERE / f"{name}.receipt.json")
    require(isinstance(receipt, dict), f"{name} receipt is not an object")
    require(receipt.get("job") == job, f"{name} job identity changed")
    require(receipt.get("exit_code") == 0, f"{name} child failed")
    require(receipt.get("command") == expected_command(job, build, plan),
            f"{name} command changed")
    require(receipt.get("binary") == build["binary"], f"{name} binary binding changed")
    source_digest = json_digest(source_census())
    require(receipt.get("source_before") == source_digest
            and receipt.get("source_after") == source_digest,
            f"{name} source custody digest changed")
    require(receipt.get("source_manifest_sha256") == source_file_sha,
            f"{name} source manifest binding changed")
    require(receipt.get("build_sha256") == builds_file_sha,
            f"{name} builds binding changed")
    require(receipt.get("plan_sha256") == plan_file_sha,
            f"{name} plan binding changed")
    require(receipt.get("script_sha256") == capture_file_sha,
            f"{name} capture script binding changed")
    require(receipt.get("environment") == plan["expected_environment"],
            f"{name} environment binding changed")
    fixture = fixture_binding(job["corpus"])
    require(receipt.get("fixture_before") == fixture and receipt.get("fixture_after") == fixture,
            f"{name} fixture custody changed")
    finite(receipt.get("seconds"), f"{name} child seconds")
    require(float(receipt["seconds"]) >= 0.0, f"{name} child seconds is negative")
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
        target = HERE / filename
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"{name} artifact digest changed: {filename}")
    return receipt, read(HERE / f"{name}.json")


def parse_cpu_line(value: Any, cpu: int, label: str) -> list[int]:
    require(isinstance(value, dict) and isinstance(value.get("text"), str),
            f"{label} /proc/stat snapshot is unavailable")
    prefix = f"cpu{cpu} "
    line = next((row for row in value["text"].splitlines() if row.startswith(prefix)), None)
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


def validate_context(plan: dict[str, Any], job: dict[str, Any]) -> dict[str, Any]:
    name = job["name"]
    path = HERE / f"{name}.context.json"
    context = read(path)
    require(isinstance(context, dict)
            and set(context) == {"before", "after", "wait4", "scope"},
            f"{name} context fields changed")
    require(context["scope"] == (
        "Whole child including setup, warmup, untimed opening, verification and reporting; "
        "system snapshots are not owner counters."
    ), f"{name} context scope changed")
    expected_keys = {"monotonic_ns", *plan["context_paths"]}
    snapshots: dict[str, dict[str, Any]] = {}
    for phase in ("before", "after"):
        snapshot = context[phase]
        require(isinstance(snapshot, dict) and set(snapshot) == expected_keys,
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
        finite(wait4[field], f"{name}.wait4.{field}")
        require(float(wait4[field]) >= 0.0, f"{name}.wait4.{field} is negative")
    before_cpu = parse_cpu_line(snapshots["before"]["/proc/stat"], plan["cpu"],
                                f"{name} before")
    after_cpu = parse_cpu_line(snapshots["after"]["/proc/stat"], plan["cpu"],
                               f"{name} after")
    require(len(before_cpu) == len(after_cpu), f"{name} CPU counter width changed")
    delta = [after - before for before, after in zip(before_cpu, after_cpu)]
    require(all(value >= 0 for value in delta), f"{name} CPU counter moved backwards")
    sysfs = {
        raw: {"before": snapshots["before"][raw], "after": snapshots["after"][raw]}
        for raw in plan["context_paths"] if raw.startswith("/sys/")
    }
    procfs = {
        raw: {"before": snapshots["before"][raw], "after": snapshots["after"][raw]}
        for raw in plan["context_paths"] if raw.startswith("/proc/")
    }
    return {
        "artifact_sha256": sha(path),
        "snapshots": snapshots,
        "wait4": wait4,
        "sysfs": sysfs,
        "procfs": procfs,
        "cpu_counter": {
            "cpu": plan["cpu"], "before": before_cpu, "after": after_cpu,
            "delta": delta, "counter_units": "procfs jiffies; context snapshot delta only",
        },
        "causal_attribution": "none; whole-child counters do not isolate the measured serialization interval",
    }


def validate_tool(report: dict[str, Any], lane: str, name: str) -> None:
    tool = report.get("tool")
    require(isinstance(tool, dict), f"{name} tool identity is missing")
    require(tool.get("name") == "litchi-perf-baseline"
            and tool.get("version") == "0.1.0"
            and tool.get("binary") == "litchi-perf-baseline"
            and tool.get("profile") == "release"
            and tool.get("target_os") == "linux"
            and tool.get("target_arch") == "x86_64",
            f"{name} tool identity changed")
    expected = "none" if lane == "native" else "ordinary_save_procfs_operation_scoped"
    require(tool.get("instrumentation") == expected,
            f"{name} instrumentation changed: {tool.get('instrumentation')!r}")


def validate_delta(value: Any, label: str) -> dict[str, int] | None:
    if value is None:
        return None
    require(isinstance(value, dict) and set(value) == set(PROCESS_DELTA_FIELDS),
            f"{label} delta fields changed")
    result: dict[str, int] = {}
    for field in PROCESS_DELTA_FIELDS:
        nonnegative_int(value[field], f"{label}.{field}")
        result[field] = value[field]
    positive_int(result["clock_ticks_per_second"], f"{label}.clock_ticks_per_second")
    return result


def process_probe_analysis(
    result: dict[str, Any], job: dict[str, Any], elapsed: list[int],
    sample_order: list[int],
) -> dict[str, Any]:
    ordinary = result.get("source", {}).get("ordinary_save")
    require(isinstance(ordinary, dict), f"{job['name']} ordinary-save source is missing")
    operation = result.get("operation_metrics")
    require(isinstance(operation, dict), f"{job['name']} operation metrics are missing")
    require(operation.get("sample_count") == job["samples"],
            f"{job['name']} process sample count changed")
    require(operation.get("sample_indices") == sample_order,
            f"{job['name']} process sample identity is not aligned to elapsed samples")
    require(operation.get("alignment") == EXPECTED_ALIGNMENT,
            f"{job['name']} process alignment changed")
    process = operation.get("process")
    require(isinstance(process, dict), f"{job['name']} process envelope is missing")
    require(set(process) == {"status", *PROCESS_VECTOR_FIELDS},
            f"{job['name']} process fields changed")
    status = process.get("status")
    require(status in {"measured", "unavailable"},
            f"{job['name']} process status is unsupported: {status!r}")
    if job["lane"] == "native":
        require("process_probe" not in ordinary,
                f"{job['name']} native result contains procfs probe evidence")
        require(status == "unavailable", f"{job['name']} native process status is not unavailable")
        probe = None
    else:
        probe = ordinary.get("process_probe")
        require(isinstance(probe, dict), f"{job['name']} procfs probe is missing")
        expected_scope = LEGACY.PHASE_TIMING_SCOPES[job["phase"]]
        expected_phase = LEGACY.PHASE_LABELS[job["phase"]]
        require(probe.get("phase") == expected_phase
                and probe.get("timing_scope") == expected_scope
                and probe.get("scope") == EXPECTED_PROCESS_PROBE_SCOPE
                and probe.get("latency_claim") == EXPECTED_PROCESS_PROBE_LATENCY
                and probe.get("control_scope") == EXPECTED_PROCESS_PROBE_CONTROL_SCOPE
                and probe.get("fixed_count") == 32,
                f"{job['name']} procfs probe metadata changed")
        controls = probe.get("empty_adjacent_snapshot_controls")
        samples = probe.get("sample_deltas")
        require(isinstance(controls, list) and len(controls) == 32,
                f"{job['name']} control count changed")
        require(isinstance(samples, list) and len(samples) == job["samples"],
                f"{job['name']} process sample vector length changed")
        controls_checked = [validate_delta(value, f"{job['name']}.control[{i}]")
                            for i, value in enumerate(controls)]
        sample_checked = [validate_delta(value, f"{job['name']}.sample[{i}]")
                          for i, value in enumerate(samples)]
        all_present = all(value is not None for value in sample_checked)
        all_absent = all(value is None for value in sample_checked)
        require(all_present or all_absent,
                f"{job['name']} process probe sample availability is asymmetric")
        require(status == ("measured" if all_present else "unavailable"),
                f"{job['name']} process status does not match probe availability")
    vectors: dict[str, Any] = {}
    for field in PROCESS_VECTOR_FIELDS:
        vector = process[field]
        require(isinstance(vector, dict), f"{job['name']}.{field} vector is missing")
        require(set(vector) == ({"values", "status", "scope"}
                                if status == "measured" else {"status", "scope"}),
                f"{job['name']}.{field} vector envelope changed")
        require(vector.get("status") == status,
                f"{job['name']}.{field} vector status changed")
        require(vector.get("scope") == PROCESS_VECTOR_SCOPES[field],
                f"{job['name']}.{field} vector scope changed")
        raw_field = PROCESS_VECTOR_SOURCE_FIELDS[field]
        entry: dict[str, Any] = {
            "unit": PROCESS_VECTOR_UNITS[field], "status": status,
            "scope": vector["scope"],
        }
        if status == "measured":
            values = vector.get("values")
            require(isinstance(values, list) and len(values) == job["samples"],
                    f"{job['name']}.{field} vector length changed")
            for index, value in enumerate(values):
                nonnegative_int(value, f"{job['name']}.{field}.values[{index}]")
            expected_values = [
                # sample_deltas are in original iteration order; operation
                # vectors are sorted by (elapsed, sample_index).
                probe["sample_deltas"][index][raw_field]  # type: ignore[index]
                for index in sample_order
            ]
            require(values == expected_values,
                    f"{job['name']}.{field} vector is not aligned to sample_order")
            entry["values"] = values
            entry["stats"] = LEGACY.integer_stats(values, f"{job['name']}.{field}")
        vectors[field] = entry
    control_summary: dict[str, Any] | None = None
    if probe is not None:
        controls = probe["empty_adjacent_snapshot_controls"]
        checked = [validate_delta(value, f"{job['name']}.control[{i}]")
                   for i, value in enumerate(controls)]
        control_fields: dict[str, Any] = {}
        for field in PROCESS_DELTA_FIELDS:
            values = [row[field] for row in checked if row is not None]
            control_fields[field] = {
                "unit": PROCESS_DELTA_UNITS.get(field, "procfs_counter"),
                "available_count": len(values),
                "stats": (LEGACY.integer_stats(values, f"{job['name']}.control.{field}")
                          if values else None),
            }
        control_summary = {
            "fixed_count": len(controls),
            "available_count": sum(value is not None for value in checked),
            "deltas": checked,
            "fields": control_fields,
            "subtraction_performed": False,
            "interpretation": (
                "counter-delta overhead only; no control durations or latency cost "
                "are measured, inferred, or subtracted"
            ),
        }
    fault = fault_diagnostic(elapsed, vectors, status, job["name"])
    return {
        "status": status,
        "sample_indices": sample_order,
        "vectors": vectors,
        "controls": control_summary,
        "fault_diagnostic": fault,
        "raw_probe": copy.deepcopy(probe),
        "raw_process": copy.deepcopy(process),
        "interpretation": (
            "descriptive counters for the same process (including any other threads); "
            "procfs probe overhead is included in the operation scope and no latency or "
            "causal claim is made"
        ),
    }


def fault_diagnostic(
    elapsed: list[int], vectors: dict[str, Any], status: str, label: str,
) -> dict[str, Any]:
    if status != "measured":
        return {"status": status, "interpretation": "fault observations unavailable"}
    result: dict[str, Any] = {"status": "measured", "distributions": {},
                              "minor_fault_elapsed_buckets": []}
    for field in ("minor_faults", "major_faults"):
        values = vectors[field]["values"]
        counts: dict[str, int] = {}
        for value in values:
            counts[str(value)] = counts.get(str(value), 0) + 1
        result["distributions"][field] = {
            "unit": "faults", "counts": {key: counts[key] for key in sorted(counts, key=int)},
            "stats": vectors[field]["stats"],
        }
    minor = vectors["minor_faults"]["values"]
    for bucket in sorted(set(minor)):
        selected = [elapsed[index] for index, value in enumerate(minor) if value == bucket]
        result["minor_fault_elapsed_buckets"].append({
            "minor_faults": bucket,
            "count": len(selected),
            "elapsed_ns": LEGACY.elapsed_stats(selected),
            "sample_positions_in_elapsed_order": [
                index for index, value in enumerate(minor) if value == bucket
            ],
        })
    result["interpretation"] = (
        "exact within-child fault distributions and elapsed buckets; controls are not "
        "subtracted and no causal or regression interpretation is made"
    )
    return result


def normalized_result(result: dict[str, Any]) -> dict[str, Any]:
    value = PAIR_HELPER.normalized_result(result)
    ordinary = value.get("source", {}).get("ordinary_save")
    if isinstance(ordinary, dict):
        # This is the only new diagnostic field removed in addition to the
        # already-reviewed 0713 normalization of timed envelopes.
        ordinary.pop("process_probe", None)
    return value


def validate_report(
    plan: dict[str, Any], job: dict[str, Any], build: dict[str, Any], report: dict[str, Any],
    source_file_sha: str,
) -> tuple[dict[str, Any], list[int], list[int], dict[str, Any], dict[str, Any]]:
    name = job["name"]
    validate_tool(report, job["lane"], name)
    LEGACY.check_report_metadata(
        report, {"binary_sha256": build["binary"]["sha256"],
                 "binary_bytes": build["binary"]["bytes"]},
        job["lane"], job["case"], plan["samples"], plan["warmup"], name,
    )
    result = report["results"][0]
    require(result.get("case") == job["case"], f"{name} result case changed")
    elapsed, elapsed_expected = LEGACY.validate_elapsed(
        result.get("elapsed_ns"), plan["samples"], name)
    ordinary = LEGACY.validate_ordinary_save(
        result, job["corpus"], "counting_publish", job["lane"], plan["samples"], name)
    LEGACY.validate_operation_metrics(
        result.get("operation_metrics"), result["elapsed_ns"], plan["samples"],
        job["lane"], name,
    )
    order = result["elapsed_ns"]["sample_order"]
    execution: list[int | None] = [None] * len(elapsed)
    for sorted_index, original_index in enumerate(order):
        execution[original_index] = elapsed[sorted_index]
    require(all(value is not None for value in execution), f"{name} execution reconstruction incomplete")
    execution_int = [int(value) for value in execution]
    historical_path = HISTORICAL_PACKET / f"native-b1-{job['corpus']['id']}-w100.json"
    historical = read(historical_path)
    historical_result = historical["results"][0]
    normalized = normalized_result(result)
    expected_normalized = normalized_result(historical_result)
    require(normalized == expected_normalized,
            f"{name} normalized output differs from 0716 native-b1 w100 baseline")
    process = process_probe_analysis(result, job, elapsed, order)
    return result, elapsed, execution_int, {
        "elapsed_stats": elapsed_expected,
        "normalized_sha256": json_digest(normalized),
        "historical_report": historical_path.name,
        "historical_report_sha256": sha(historical_path),
        "historical_normalized_sha256": json_digest(expected_normalized),
        "ordinary_published_sha256": ordinary["corpus"]["published_sha256"],
        "source_manifest_sha256": source_file_sha,
    }, process


def within_diagnostic(execution: list[int], plan: dict[str, Any]) -> dict[str, Any]:
    size = plan["analysis"]["block_samples"]
    require(len(execution) % size == 0 and len(execution) % 2 == 0,
            "sample count is incompatible with within-child diagnostic")
    blocks = [statistics.mean(execution[start:start + size])
              for start in range(0, len(execution), size)]
    halves = [statistics.mean(execution[:len(execution) // 2]),
              statistics.mean(execution[len(execution) // 2:])]
    block_span = spread(blocks, "within block means")
    half_span = spread(halves, "within half means")
    return {
        "block_size": size, "block_means_ns": blocks,
        "block_span_percent": block_span, "block_flag_over_5_percent": block_span > 5.0,
        "half_means_ns": halves, "half_span_percent": half_span,
        "half_flag_over_5_percent": half_span > 5.0,
        "interpretation": "descriptive relative spans; not a statistical stationarity test",
    }


def metric_stats(values: list[float], label: str) -> dict[str, Any]:
    value = spread(values, label)
    return {"values": values, "min": min(values), "max": max(values),
            "spread_percent": value, "flag_over_5_percent": value > 5.0}


def nonnegative_metric_stats(values: list[float], label: str) -> dict[str, Any]:
    """Describe a counter summary whose minimum may legitimately be zero."""

    require(values and all(value >= 0.0 for value in values),
            f"{label} requires non-negative values")
    minimum, maximum = min(values), max(values)
    if minimum == 0.0:
        relative_spread: float | None = 0.0 if maximum == 0.0 else None
        flag: bool | None = False if maximum == 0.0 else None
    else:
        relative_spread = (maximum - minimum) * 100.0 / minimum
        flag = relative_spread > 5.0
    return {
        "values": values,
        "min": minimum,
        "max": maximum,
        "spread_percent": relative_spread,
        "flag_over_5_percent": flag,
        "spread_note": (
            "relative spread is undefined when the minimum counter is zero"
            if relative_spread is None else "relative spread is (max-min)/min"
        ),
    }


def group_diagnostics(rows: list[dict[str, Any]], plan: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for corpus in plan["corpora"]:
        for lane in LANES:
            selected = [row for row in rows
                        if row["job"]["corpus"]["id"] == corpus["id"]
                        and row["job"]["lane"] == lane]
            require(len(selected) == plan["blocks"], f"{corpus['id']}/{lane} child count changed")
            metrics = {
                metric: metric_stats(
                    [float(row["elapsed_stats"][metric]) for row in selected],
                    f"{corpus['id']}/{lane}/{metric}",
                )
                for metric in plan["analysis"]["metrics"]
            }
            result.append({
                "corpus_id": corpus["id"], "lane": lane,
                "repetitions": len(selected),
                "children": [row["job"]["name"] for row in selected],
                "metrics": metrics,
                "note": "Sixteen process-level summaries of in-process serialization samples; spread is (max-min)/min and no independence is assumed.",
            })
    require(len(result) == 4, "group count changed")
    return result


def process_group_diagnostics(rows: list[dict[str, Any]], plan: dict[str, Any]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for corpus in plan["corpora"]:
        selected = [row for row in rows
                    if row["job"]["corpus"]["id"] == corpus["id"]
                    and row["job"]["lane"] == "procfs"]
        for field in ("minor_faults", "major_faults"):
            statuses = {row["process"]["status"] for row in selected}
            if statuses == {"measured"}:
                stats = {
                    key: nonnegative_metric_stats(
                        [float(row["process"]["vectors"][field]["stats"][key]) for row in selected],
                        f"{corpus['id']}/procfs/{field}/{key}",
                    )
                    for key in ("p50", "mean", "p95", "p99")
                }
            else:
                stats = {"status": sorted(statuses)}
            result.append({"corpus_id": corpus["id"], "lane": "procfs",
                           "metric": field, "children": [row["job"]["name"] for row in selected],
                           "stats": stats,
                           "note": "Descriptive process-counter spread; no controls are subtracted and no cause is inferred."})
    return result


def build_output(
    plan: dict[str, Any], source: dict[str, str], source_file_sha: str,
    baseline: dict[str, str], baseline_file_sha: str, builds: dict[str, Any],
    builds_file_sha: str, constraints_file_sha: str, helper_freeze: dict[str, str],
    capture_freeze: dict[str, str], rows: list[dict[str, Any]],
) -> dict[str, Any]:
    return {
        "schema_version": 1,
        "packet": plan["packet"],
        "revision": plan["revision"],
        "diagnostic_claim": (
            "Baseline-only process-counter observations. Native and procfs elapsed values "
            "are not a latency comparison; controls are retained and never subtracted."
        ),
        "scope": {
            "corpora": list(CORPUS_IDS), "phase": "counting_publish",
            "children": len(rows), "samples_per_child": plan["samples"],
            "cpu": plan["cpu"], "blocks": plan["blocks"], "warmup": plan["warmup"],
            "lanes": list(LANES),
        },
        "rows": rows,
        "groups_by_corpus_lane": group_diagnostics(rows, plan),
        "process_groups_by_corpus": process_group_diagnostics(rows, plan),
        "verification": {
            "current_source_matches_source_json": True,
            "source_changes_are_exactly_three_diagnostic_files": True,
            "strict_receipts_verified": len(rows),
            "strict_0709_validators_used": [
                "check_report_metadata", "validate_elapsed", "validate_operation_metrics",
                "validate_ordinary_save",
            ],
            "normalized_parity_reference": "0716/native-b1-{corpus}-w100.json",
            "normalized_parity_verified": True,
            "normalized_parity_extra_field_removed": "source.ordinary_save.process_probe only",
            "sample_order_reconstructed": True,
            "all_samples_retained": True,
            "all_children_retained": True,
            "within_child_flags_are_descriptive": True,
            "between_child_groups_are_descriptive": True,
            "fault_buckets_are_descriptive": True,
            "controls_subtracted": False,
            "whole_child_context_is_not_measured_interval": True,
            "wait4_and_sysfs_retained_separately": True,
            "causal_attribution": "none",
        },
        "binary_custody": {
            lane_builds: builds[lane_builds]["binary"] for lane_builds in LANES
        },
        "builds": builds,
        "source": {
            "manifest": "source.json", "sha256": source_file_sha,
            "entry_count": len(source), "baseline_manifest": "source-baseline.json",
            "baseline_sha256": baseline_file_sha, "baseline_entry_count": len(baseline),
            "allowed_changed_paths": sorted(ALLOWED_SOURCE_CHANGES),
        },
        "freeze_bindings": {
            "constraints_json_sha256": constraints_file_sha,
            "capture_freeze": capture_freeze, "helper_freeze": helper_freeze,
            "capture_py_sha256": sha(HERE / "capture.py"),
            "plan_json_sha256": sha(HERE / "plan.json"),
            "analyzer_sha256": sha(HERE / "analyze.py"),
            "builds_json_sha256": builds_file_sha,
        },
        "limits": plan["limits"],
    }


def analyze() -> dict[str, Any]:
    plan = load_plan()
    constraints_file_sha = validate_constraints()
    helper_freeze = validate_freeze("helper-freeze.json")
    capture_freeze = validate_freeze("capture-freeze.json")
    source, source_file_sha, baseline, baseline_file_sha = validate_source()
    builds, builds_file_sha = validate_builds(plan, source_file_sha)
    plan_file_sha = sha(HERE / "plan.json")
    capture_file_sha = sha(HERE / "capture.py")
    expected_jobs = jobs(plan)
    expected_receipts = {f"{job['name']}.receipt.json" for job in expected_jobs}
    actual_receipts = {path.name for path in HERE.glob("*-b*.receipt.json")}
    require(actual_receipts == expected_receipts,
            "receipt inventory is not exactly the frozen 64-child set")
    expected_artifacts = set(expected_receipts)
    for job in expected_jobs:
        expected_artifacts.update(
            f"{job['name']}.{suffix}"
            for suffix in ("json", "stdout", "stderr", "context.json")
        )
    actual_artifacts = {path.name for lane in LANES
                        for path in HERE.glob(f"{lane}-b*") if path.is_file()}
    require(actual_artifacts == expected_artifacts,
            "child artifact inventory is not exactly the frozen 64-child set")
    rows: list[dict[str, Any]] = []
    previous_context_after: int | None = None
    for job in expected_jobs:
        receipt, report = validate_receipt(
            plan, job, builds[job["lane"]], source_file_sha, builds_file_sha,
            plan_file_sha, capture_file_sha,
        )
        result, elapsed, execution, parity, process = validate_report(
            plan, job, builds[job["lane"]], report, source_file_sha,
        )
        context = validate_context(plan, job)
        before_ns = context["snapshots"]["before"]["monotonic_ns"]
        after_ns = context["snapshots"]["after"]["monotonic_ns"]
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
            "process": process,
            "normalized_parity": parity,
            "whole_child_context": context,
        })
    require(len(rows) == plan["expected_children"], "exact 64 rows were not retained")
    return build_output(
        plan, source, source_file_sha, baseline, baseline_file_sha, builds, builds_file_sha,
        constraints_file_sha, helper_freeze, capture_freeze, rows,
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
            print(f"PASS deterministic 0717 analysis check ({len(output['rows'])} rows)")
        else:
            require(output_path == HERE / "analysis.json",
                    "normal analysis output must be packet analysis.json")
            require(not output_path.exists() and not output_path.is_symlink(),
                    "refusing to replace existing analysis.json")
            output_path.write_bytes(data)
            print(f"PASS verified {len(output['rows'])} children; wrote {output_path}")
    except (AssertionError, OSError, RuntimeError, ValueError, KeyError, TypeError) as error:
        print(f"analysis failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
