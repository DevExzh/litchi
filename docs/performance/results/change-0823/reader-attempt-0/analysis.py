"""Fail-closed, offline replay for the 0823 PPTX scanner trial.

This reader consumes only packet receipts and probe JSON.  It never starts a
build, workload, profiler, decoder, or binary-inspection command.  The
drivers write one immutable receipt per process; this module checks those
receipts, validates both semantic oracles independently, and derives the
small deterministic artifacts used by the final validator.
"""

from __future__ import annotations

import csv
import hashlib
import json
import math
import random
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_SCHEMA = "litchi.performance.0823.pptx-scanner.v1"
ANALYSIS_SCHEMA = "litchi.performance.0823.scanner-analysis.v1"
BASE = "b76786208d04310b4a033b39d70de439a64fcf69"
BASE_SHORT = "b76786208d"
TARGET = Path("/home/zhuhe/code/litchi-target-0823")
REAL_INPUT = {"path": "test-data/ooxml/pptx/shapes.pptx", "bytes": 68822,
              "sha256": "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571"}
REAL_REFERENCE = {"path": "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx",
                  "bytes": 68284,
                  "sha256": "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf"}
REAL_FULL_TEXT_SHA256 = "5e9363235c4a1b158819a66be5aa46bf34d71b7da24592d855aad3ca90ce8b82"
REAL_MARKER = "litchi-perf-0638-ordinary-save"
SHAPES = ("tiny", "medium", "large", "vendor", "unicode-vendor", "valid-4attr")
MODES = ("capture", "commit", "lifecycle")
CASES = tuple(("synthetic", shape, mode) for shape in SHAPES for mode in MODES) + (
    ("real", "real", "direct"),
)
LEGS = ("before", "after")
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_SEED = 823823
BOOTSTRAP_LOW = 250
BOOTSTRAP_HIGH = 9749
RAW_ALLOC = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)
ALLOC_METRICS = RAW_ALLOC + ("net_live", "peak_above_entry")
TIMING_METRICS = ("p50", "mean", "p95", "p99")
HEX = frozenset("0123456789abcdef")


class ReplayError(RuntimeError):
    """Packet evidence is absent, malformed, stale, or contradictory."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def sha(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for block in iter(lambda: stream.read(1 << 20), b""):
                digest.update(block)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def integer(value: Any, label: str, *, positive: bool = False) -> None:
    require(isinstance(value, int) and not isinstance(value, bool)
            and (value > 0 if positive else value >= 0), f"{label}: invalid integer")


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label}: not finite")


def rel(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def packet_path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label}: path missing")
    raw = Path(value)
    path = (raw if raw.is_absolute() else PACKET / raw).resolve()
    require(path.is_relative_to(PACKET.resolve()), f"{label}: path escapes packet")
    return path


def artifact(value: Any, label: str) -> Path:
    require(isinstance(value, dict), f"{label}: artifact malformed")
    path = packet_path(value.get("path"), label)
    integer(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")), f"{label}.sha256 malformed")
    require(path.is_file() and not path.is_symlink(), f"{label}: file missing")
    require(path.stat().st_size == value["bytes"], f"{label}: byte count changed")
    require(sha(path) == value["sha256"], f"{label}: digest changed")
    return path


def external_artifact(value: Any, label: str) -> dict[str, Any]:
    """Check an absolute binary identity while allowing a later cleanup witness."""
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label}: binary identity malformed")
    path = Path(value["path"])
    require(path.is_absolute() and not path.is_symlink(), f"{label}: binary path malformed")
    integer(value.get("bytes"), f"{label}.bytes", positive=True)
    require(is_sha(value.get("sha256")), f"{label}.sha256 malformed")
    if path.is_file():
        require(path.stat().st_size == value["bytes"] and sha(path) == value["sha256"],
                f"{label}: binary changed")
    else:
        cleanup = PACKET / "cleanup.json"
        require(cleanup.is_file(), f"{label}: binary missing without cleanup witness")
        value2 = read(cleanup)
        require(value2.get("schema") == "litchi.performance.0823.cleanup.v1"
                and value2.get("verified") is True
                and value2.get("target_absent_after_removal") is True,
                f"{label}: cleanup witness is not verified")
        removed = value2.get("binaries", [])
        require(any(isinstance(row, dict) and row.get("path") == str(path)
                    and row.get("bytes") == value["bytes"]
                    and row.get("sha256") == value["sha256"] for row in removed),
                f"{label}: exact removed binary witness missing")
    return {"path": str(path), "bytes": value["bytes"], "sha256": value["sha256"]}


def source_manifest(value: Any, label: str) -> dict[str, Any]:
    path = artifact(value, label)
    document = read(path)
    require(isinstance(document, dict) and isinstance(document.get("files"), dict),
            f"{label}: source manifest malformed")
    files = document["files"]
    require(files and all(isinstance(name, str) and is_sha(digest)
                          for name, digest in files.items()),
            f"{label}: source census malformed")
    return {"revision": document.get("revision"), "files": dict(files),
            "descriptor": {"path": rel(path), "bytes": value["bytes"],
                           "sha256": value["sha256"]}}


def frozen_inputs() -> dict[str, Any]:
    plan = read(PACKET / "plan.json")
    require(plan.get("schema") == PLAN_SCHEMA and plan.get("base") == BASE,
            "plan schema or base changed")
    require(plan.get("target") == str(TARGET) and plan.get("cpu") == 12,
            "plan target or CPU changed")
    require(plan.get("source_allowlist") == ["crates/litchi-pptx/src/shape/reader.rs"],
            "source allowlist changed")
    expected_cases = [{"probe": "synthetic", "shape": shape, "mode": mode}
                      for shape in SHAPES for mode in MODES]
    expected_cases.append({"probe": "real", "mode": "direct",
                           "input_selector": "pinnedshapes",
                           "reference_selector": "0821realPPTXdefault"})
    require(plan.get("cases") == expected_cases, "case matrix changed")
    require(plan.get("qualification", {}).get("reports") == 38
            and plan["qualification"].get("samples_total") == 38,
            "qualification cardinality changed")
    require(plan.get("native", {}).get("reports") == 228
            and plan["native"].get("samples_total") == 6840,
            "native cardinality changed")
    require(plan.get("allocation", {}).get("reports") == 76
            and plan["allocation"].get("samples_total") == 228,
            "allocation cardinality changed")
    require(plan.get("totals") == {"qualification_reports": 38,
                                    "qualification_samples": 38,
                                    "native_reports": 228,
                                    "native_samples": 6840,
                                    "allocation_reports": 76,
                                    "allocation_samples": 228,
                                    "reports": 342, "samples": 7106},
            "total cardinality changed")
    policy = plan.get("policy")
    require(policy == {
        "bootstrap_resamples": 10000, "bootstrap_seed": 823823,
        "sorted_zero_based_endpoints": [250, 9749],
        "eligible_benefit_cases": "All twelve synthetic commit/lifecycle rows plus the real/direct row; synthetic capture rows are retained as negative controls because the timed capture may precede Scene::Scanner::scan.",
        "benefit_gate": "At least one eligible representative public workflow has median after/before <= 0.97 and bootstrap upper endpoint < 1.00.",
        "latency_veto": "Any of all nineteen measured rows with bootstrap lower endpoint > 1.05 vetoes adoption.",
        "allocation_gate": "Paired allocation medians must not increase calls, allocated bytes, net live bytes, or peak above entry.",
        "rss_review": "A paired process-RSS increase above 5 percent requires explicit review; RSS is diagnostic and cannot establish a memory saving.",
        "claims": "No universal, cross-format, cold-cache, tail-latency, RSS, or historical speedup claim.",
    }, "policy changed")
    return plan


def integrity(plan: dict[str, Any]) -> dict[str, Any]:
    origin = read(PACKET / "origin.json")
    require(origin.get("schema") == "litchi.performance.0823.origin.v1"
            and origin.get("base") == BASE
            and origin.get("production_changed") is False
            and origin.get("runtime_harness_changed") is False
            and origin.get("tool_changed") is False
            and origin.get("source_allowlist") == plan["source_allowlist"],
            "origin custody changed")
    unrelated = origin.get("unrelated")
    require(isinstance(unrelated, dict) and len(unrelated) == 3, "unrelated custody changed")
    for name, digest in unrelated.items():
        path = ROOT / name
        require(path.is_file() and not path.is_symlink() and sha(path) == digest,
                f"unrelated path changed: {name}")

    architecture = read(PACKET / "architecture-inputs.json")
    require(isinstance(architecture, dict) and len(architecture) == 35
            and all(isinstance(name, str) and is_sha(digest)
                    for name, digest in architecture.items()),
            "architecture input census changed")
    for name, digest in architecture.items():
        path = ROOT / name
        require(path.is_file() and sha(path) == digest, f"architecture input changed: {name}")

    root_inputs = read(PACKET / "root-inputs.json")
    require(root_inputs.get("schema") == "litchi.performance.0823.root-inputs.v1",
            "root-input schema changed")
    for key, row in root_inputs.items():
        if key == "schema":
            continue
        artifact(row, f"root input {key}")
        source_name = row.get("source_path")
        require(isinstance(source_name, str), f"root input {key}: source path missing")
        source = ROOT / source_name
        require(source.is_file() and sha(source) == row["sha256"],
                f"root input source changed: {source_name}")

    lock_parity = read(PACKET / "lock-parity.json")
    require(lock_parity.get("schema") == "litchi.performance.0823.lock-parity.v1"
            and lock_parity.get("cargo_lock_changes") ==
            "No lockfile generation or update is permitted during this trial.",
            "lock parity changed")
    for key in ("root_lock", "tool_lock"):
        artifact(lock_parity[key], f"lock parity {key}")
    for name, row in lock_parity.get("probe_locks", {}).items():
        artifact(row, f"probe lock {name}")

    host = read(PACKET / "host.json")
    require(host.get("schema") == "litchi.performance.0823.host.v1"
            and host.get("affinity_selected") == [12]
            and host.get("target") == str(TARGET)
            and host.get("scratch") is None, "host custody changed")
    toolchain = read(PACKET / "toolchain.json")
    contract = toolchain.get("build_contract", {})
    require(toolchain.get("schema") == "litchi.performance.0823.toolchain.v1"
            and contract.get("offline") is True and contract.get("locked") is True
            and contract.get("release") is True
            and contract.get("jobs") == plan["build"]["jobs"]
            and contract.get("opt_level") == plan["build"]["opt_level"]
            and contract.get("debug") == plan["build"]["debug"]
            and contract.get("codegen_units") == plan["build"]["codegen_units"]
            and contract.get("incremental") is False
            and contract.get("panic") == plan["build"]["panic"],
            "toolchain or build contract changed")
    return {"origin": origin, "architecture": architecture, "root_inputs": root_inputs,
            "lock_parity": lock_parity, "host": host, "toolchain": toolchain}


def source_state(leg: str) -> dict[str, Any]:
    directory = PACKET / f"build-{leg}"
    value = read(directory / "build.json")
    require(value.get("schema") == f"litchi.performance.0823.build-{leg}.v1"
            and value.get("leg") == leg, f"{leg}: build schema changed")
    require(isinstance(value.get("source"), dict), f"{leg}: build source missing")
    source = source_manifest(value["source"], f"{leg}: source")
    frozen_descriptor = value.get("frozen_inputs")
    require(isinstance(frozen_descriptor, dict)
            and artifact(frozen_descriptor, f"{leg}: frozen inputs")
            and frozen_descriptor.get("sha256") == sha(PACKET / "freeze.json"),
            f"{leg}: frozen-input custody changed")
    require(value.get("target") == str(TARGET), f"{leg}: target changed")
    freeze = read(PACKET / "freeze.json")
    require(value.get("root_inputs") == freeze.get("root_inputs"),
            f"{leg}: root-input custody changed")
    plan = read(PACKET / "plan.json")
    require(value.get("profile") == plan.get("build")
            and value.get("environment_contract") == {
                "offline": True, "locked": True, "release": True,
                "jobs": plan["build"]["jobs"], "serial": True,
            }, f"{leg}: build contract changed")
    require(value.get("architecture") == read(PACKET / "architecture-inputs.json"),
            f"{leg}: architecture receipt changed")
    require(value.get("unrelated") == read(PACKET / "origin.json").get("unrelated"),
            f"{leg}: unrelated receipt changed")
    require(value.get("probes") == {
        name: {str(path.relative_to(PACKET / f"{name}-probe-src")): sha(path)
               for path in sorted((PACKET / f"{name}-probe-src").rglob("*")) if path.is_file()}
        for name in ("synthetic", "real")}, f"{leg}: probe custody changed")
    binaries = value.get("binaries")
    require(isinstance(binaries, dict) and binaries, f"{leg}: binaries missing")
    checked: dict[str, dict[str, Any]] = {}
    for name, row in binaries.items():
        # build.py currently wraps each descriptor as {"artifact": ...}; keep
        # the reader compatible with the frozen direct-descriptor shape too.
        descriptor = (row["artifact"] if isinstance(row, dict) and "artifact" in row
                      else row)
        checked[name] = external_artifact(descriptor, f"{leg} binary {name}")
    require(set(checked) == {f"{probe}-{variant}" for probe in ("synthetic", "real")
                            for variant in ("native", "allocation")},
            f"{leg}: binary matrix changed")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 4, f"{leg}: build command count changed")
    seen_commands: set[tuple[str, str]] = set()
    seen_logs: set[str] = set()
    for index, row in enumerate(rows):
        probe = row.get("probe")
        variant = row.get("variant")
        require(probe in {"synthetic", "real"} and variant in {"native", "allocation"},
                f"{leg}: build {index} identity changed")
        spec = read(PACKET / "plan.json")["probes"][probe]
        manifest = PACKET / spec["manifest"].removeprefix("docs/performance/results/change-0823/")
        expected_command = ["cargo", "build", "--offline", "--locked", "--release",
                            "--manifest-path", str(manifest), "--bin", spec["cargo_binary"]]
        features = spec["features"][variant]
        if features:
            expected_command += ["--features", ",".join(features)]
        expected_environment = {
            "CARGO_TARGET_DIR": str(TARGET), "CARGO_BUILD_JOBS": "2",
            "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_RELEASE_OPT_LEVEL": "3",
            "CARGO_PROFILE_RELEASE_DEBUG": "1", "CARGO_PROFILE_RELEASE_LTO": "thin",
            "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": "1",
            "CARGO_PROFILE_RELEASE_INCREMENTAL": "false",
            "CARGO_PROFILE_RELEASE_PANIC": "unwind", "PYTHONDONTWRITEBYTECODE": "1",
        }
        command_key = (str(probe), str(variant))
        require(command_key not in seen_commands, f"{leg}: duplicate build identity {command_key}")
        seen_commands.add(command_key)
        require(isinstance(row, dict) and row.get("schema") ==
                "litchi.performance.0823.build-receipt.v1"
                and row.get("leg") == leg
                and row.get("features") == features
                and row.get("exit_code") == 0
                and row.get("command") == expected_command
                and row.get("environment") == expected_environment,
                f"{leg}: build command {index} changed")
        finite(row.get("started"), f"{leg} build {index} started")
        finite(row.get("ended"), f"{leg} build {index} ended")
        require(row["started"] <= row["ended"], f"{leg}: build times reversed")
        log = artifact(row.get("log"), f"{leg} build {index} log")
        require(str(log) not in seen_logs, f"{leg}: duplicate build log")
        seen_logs.add(str(log))
    return {"raw": value, "source": source, "binaries": checked}


def source_diff(before: dict[str, Any], after: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    freeze = read(PACKET / "freeze.json")
    require(freeze.get("schema") == "litchi.performance.0823.freeze.v1"
            and freeze.get("base") == BASE
            and freeze.get("source", {}).get("files") == before["source"]["files"],
            "freeze source custody changed")
    for name, digest in freeze.get("candidate", {}).items():
        path = PACKET / name
        require(path.is_file() and sha(path) == digest, f"candidate custody changed: {name}")
    for name, digest in freeze.get("drivers", {}).items():
        path = PACKET / name
        require(path.is_file() and sha(path) == digest, f"driver custody changed: {name}")
    left, right = before["source"]["files"], after["source"]["files"]
    require(set(left) == set(right), "before/after source file set changed")
    changed = sorted(name for name in left if left[name] != right[name])
    require(changed == plan["source_allowlist"], "candidate changed files outside allowlist")
    require(left != right, "candidate source is identical to before")
    candidate_patch = PACKET / "candidate/candidate.patch"
    require(candidate_patch.is_file() and not candidate_patch.is_symlink(),
            "candidate patch is missing")
    return {"before_revision": before["source"].get("revision"),
            "after_revision": after["source"].get("revision"),
            "changed_files": changed,
            "before_sha256": before["source"]["descriptor"]["sha256"],
            "after_sha256": after["source"]["descriptor"]["sha256"],
            "candidate_before_sha256": left[plan["source_allowlist"][0]],
            "candidate_after_sha256": right[plan["source_allowlist"][0]],
            "candidate_patch_sha256": sha(candidate_patch),
            "candidate_patch_bytes": candidate_patch.stat().st_size}


def quality_receipt() -> dict[str, Any]:
    """Require the three exact final quality receipts and their command custody."""
    expected_counts = {"before": 6, "after": 6, "probes": 18}
    freeze = read(PACKET / "freeze.json")
    # ``quality.py`` snapshots tracked production files plus the explicit
    # architecture/unrelated/probe inputs.  The root lock is a packet input
    # copied by custody.py, but it is not a tracked source entry in that map;
    # rustfmt.toml is tracked and is checked through the production census
    # below.  Freeze records .cargo/config.toml separately because the quality
    # driver's source enumerator intentionally omits it.
    expected_env = {
        "CARGO_TARGET_DIR": str(ROOT.parent / "litchi-target-0823" / "quality"),
        "CARGO_BUILD_JOBS": "2", "CARGO_INCREMENTAL": "0",
        "CARGO_PROFILE_DEV_DEBUG": "0", "RUSTDOCFLAGS": "-D warnings",
        "PYTHONDONTWRITEBYTECODE": "1",
    }

    def production_commands() -> list[list[str]]:
        package = ["-p", "litchi-pptx"]
        return [
            ["cargo", "fmt", *package, "--", "--check"],
            ["cargo", "check", "--offline", "--locked", *package,
             "--all-features", "--all-targets"],
            ["cargo", "test", "--offline", "--locked", *package,
             "--all-features", "--", "--test-threads=2"],
            ["cargo", "clippy", "--offline", "--locked", *package,
             "--all-features", "--all-targets", "--", "-D", "warnings"],
            ["cargo", "doc", "--offline", "--locked", *package,
             "--all-features", "--no-deps"],
            [sys.executable, "-B", "tools/check_crate_boundaries.py"],
        ]

    def probe_commands() -> list[list[str]]:
        commands: list[list[str]] = []
        for family in ("synthetic", "real"):
            manifest = str(PACKET / f"{family}-probe-src/Cargo.toml")
            commands.append(["cargo", "fmt", "--manifest-path", manifest, "--", "--check"])
            for features in ([], ["--all-features"]):
                for verb in ("check", "test", "clippy", "doc"):
                    command = ["cargo", verb, "--offline", "--locked",
                               "--manifest-path", manifest, *features]
                    if verb in ("check", "clippy"):
                        command.append("--all-targets")
                    if verb == "test":
                        command += ["--", "--test-threads=2"]
                    if verb == "clippy":
                        command += ["--", "-D", "warnings"]
                    if verb == "doc":
                        command.append("--no-deps")
                    commands.append(command)
        return commands

    def source_intersections(value: dict[str, Any], leg: str) -> None:
        source_name = value.get("source")
        require(isinstance(source_name, str), f"quality {leg}: source missing")
        source_path = (ROOT / source_name).resolve() if not Path(source_name).is_absolute() else Path(source_name)
        require(source_path.is_file() and not source_path.is_symlink(),
                f"quality {leg}: source snapshot missing")
        source = read(source_path)
        require(isinstance(source, dict), f"quality {leg}: source snapshot malformed")
        production = freeze.get("source", {}).get("files")
        require(isinstance(production, dict) and production,
                f"quality {leg}: freeze production census missing")
        omitted = ".cargo/config.toml"
        require(omitted in production,
                f"quality {leg}: documented config omission is absent from freeze")
        require(omitted not in source,
                f"quality {leg}: documented config omission unexpectedly covered")
        for name, digest in production.items():
            # quality.py's state() deliberately omits this config path; the
            # freeze receipt and integrity() bind its exact bytes separately.
            if name == omitted:
                continue
            expected = digest
            if leg == "after" and name == "crates/litchi-pptx/src/shape/reader.rs":
                expected = freeze["candidate"]["candidate/after/reader.rs"]
            require(source.get(name) == expected,
                    f"quality {leg}: production source intersection changed: {name}")
        for name, digest in read(PACKET / "architecture-inputs.json").items():
            require(source.get(name) == digest, f"quality {leg}: architecture intersection changed: {name}")
        for name, digest in read(PACKET / "origin.json")["unrelated"].items():
            require(source.get(name) == digest, f"quality {leg}: unrelated intersection changed: {name}")
        if leg == "probes":
            # Probe quality runs before candidate application, so the complete
            # production loop above is the frozen baseline as well as both
            # final packet-local probe trees.
            for family, files in freeze["probes"].items():
                prefix = f"docs/performance/results/change-0823/{family}-probe-src/"
                for name, digest in files.items():
                    require(source.get(prefix + name) == digest,
                            f"quality probes: probe source intersection changed: {family}/{name}")

    def one(leg: str, commands: list[list[str]]) -> dict[str, Any]:
        path = PACKET / f"quality-{leg}.json"
        require(len(commands) == expected_counts[leg],
                f"quality {leg}: expected command plan changed")
        value = read(path)
        require(value.get("schema") == "litchi.performance.0823.quality.v1"
                and value.get("leg") == leg and value.get("status") == "pass",
                f"quality {leg}: final receipt changed")
        require(value.get("environment") == expected_env,
                f"quality {leg}: environment changed")
        require(value.get("driver_sha256") == sha(PACKET / "quality.py"),
                f"quality {leg}: driver hash changed")
        require(value.get("rows") and len(value["rows"]) == len(commands),
                f"quality {leg}: command count changed")
        source_intersections(value, leg)
        seen_logs: set[str] = set()
        for index, (row, expected) in enumerate(zip(value["rows"], commands)):
            require(isinstance(row, dict) and row.get("command") == expected
                    and row.get("exit_code") == 0,
                    f"quality {leg}/{index}: command receipt changed")
            finite(row.get("started"), f"quality {leg}/{index} started")
            finite(row.get("ended"), f"quality {leg}/{index} ended")
            require(row["started"] <= row["ended"], f"quality {leg}/{index}: times reversed")
            log_name = row.get("log")
            require(isinstance(log_name, str), f"quality {leg}/{index}: log missing")
            log_path = (ROOT / log_name).resolve() if not Path(log_name).is_absolute() else Path(log_name)
            require(log_path.is_file() and not log_path.is_symlink(), f"quality {leg}/{index}: log missing")
            require(str(log_path) not in seen_logs,
                    f"quality {leg}/{index}: duplicate log path")
            seen_logs.add(str(log_path))
            require(log_path.stat().st_size == row.get("log_bytes")
                    and sha(log_path) == row.get("log_sha256"),
                    f"quality {leg}/{index}: log identity changed")
        return {"path": rel(path), "sha256": sha(path), "leg": leg,
                "commands": len(commands), "source": value["source"]}

    values = [one("before", production_commands()), one("after", production_commands()),
              one("probes", probe_commands())]
    # Verify retained failed attempts too; they do not satisfy the final gate,
    # but every failed command must retain its own log.
    for receipt_path in sorted(PACKET.glob("quality-*/receipt.json")):
        receipt = read(receipt_path)
        for index, row in enumerate(receipt.get("rows", [])):
            require(isinstance(row, dict) and isinstance(row.get("log"), str),
                    f"quality attempt {receipt_path.name}/{index}: log missing")
            log_path = (ROOT / row["log"]).resolve() if not Path(row["log"]).is_absolute() else Path(row["log"])
            require(log_path.is_file() and not log_path.is_symlink(),
                    f"quality attempt {receipt_path.name}/{index}: log missing")
            require(log_path.stat().st_size == row.get("log_bytes")
                    and sha(log_path) == row.get("log_sha256"),
                    f"quality attempt {receipt_path.name}/{index}: log identity changed")
    return {"receipts": values, "all_passed": True}


def chronology() -> dict[str, Any]:
    """Bind the documented serial execution order from every terminal receipt."""

    def bounds(rows: Any, label: str) -> dict[str, Any]:
        require(isinstance(rows, list) and rows, f"{label}: timestamp rows missing")
        previous_end: float | None = None
        starts: list[float] = []
        ends: list[float] = []
        for index, row in enumerate(rows):
            require(isinstance(row, dict), f"{label}/{index}: timestamp row malformed")
            started = row.get("started")
            ended = row.get("ended")
            finite(started, f"{label}/{index} started")
            finite(ended, f"{label}/{index} ended")
            require(started <= ended, f"{label}/{index}: timestamps reversed")
            if previous_end is not None:
                require(previous_end <= started,
                        f"{label}/{index}: serial receipt order overlaps")
            previous_end = float(ended)
            starts.append(float(started))
            ends.append(float(ended))
        return {"start": min(starts), "end": max(ends), "rows": len(rows)}

    def receipt_rows(directory: Path, label: str) -> list[dict[str, Any]]:
        complete = read(directory / "complete.json")
        require(complete.get("status") == "pass",
                f"{label}: complete receipt is not terminal")
        descriptor = complete.get("receipts")
        path = artifact(descriptor, f"{label}: receipts")
        rows = read(path)
        require(isinstance(rows, list), f"{label}: receipts malformed")
        return rows

    spans: dict[str, dict[str, Any]] = {}
    for leg in LEGS:
        quality = read(PACKET / f"quality-{leg}.json")
        spans[f"quality-{leg}"] = bounds(quality.get("rows"), f"quality-{leg}")
    spans["quality-probes"] = bounds(
        read(PACKET / "quality-probes.json").get("rows"), "quality-probes")
    for leg in LEGS:
        build = read(PACKET / f"build-{leg}/build.json")
        spans[f"build-{leg}"] = bounds(build.get("rows"), f"build-{leg}")
        spans[f"qualification-{leg}"] = bounds(
            receipt_rows(PACKET / f"qualification-{leg}", f"qualification-{leg}"),
            f"qualification-{leg}")
    for lane_name in ("native", "allocation"):
        spans[lane_name] = bounds(
            receipt_rows(PACKET / lane_name, lane_name), lane_name)

    order = ["quality-before", "quality-probes", "build-before",
             "qualification-before", "quality-after", "build-after",
             "qualification-after", "native", "allocation"]
    for left, right in zip(order, order[1:]):
        require(spans[left]["end"] <= spans[right]["start"],
                f"serial chronology overlaps: {left} before {right}")
    return {"order": order, "spans": spans, "serial": True}


def nearest(values: Iterable[float], percentile: float) -> float:
    ordered = sorted(values)
    require(ordered, "empty quantile vector")
    return ordered[max(1, math.ceil(len(ordered) * percentile)) - 1]


def stats(values: Iterable[float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty metric vector")
    for index, value in enumerate(vector):
        finite(value, f"metric[{index}]")
        require(value >= 0, f"metric[{index}]: negative")
    return {"count": len(vector), "min": min(vector), "p50": nearest(vector, .50),
            "mean": statistics.mean(vector), "p95": nearest(vector, .95),
            "p99": nearest(vector, .99), "max": max(vector)}


def spread(values: Iterable[float]) -> float:
    vector = list(values)
    require(vector, "empty spread vector")
    low, high = min(vector), max(vector)
    return 0.0 if low == high else (math.inf if low == 0 else (high - low) / abs(low))


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(values, "empty bootstrap vector")
    rng = random.Random(BOOTSTRAP_SEED)
    draws = sorted(statistics.median(values[rng.randrange(len(values))] for _ in values)
                   for _ in range(BOOTSTRAP_RESAMPLES))
    return {"seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
            "statistic": "median", "confidence": 0.95,
            "low_rank": BOOTSTRAP_LOW, "high_rank": BOOTSTRAP_HIGH,
            "ci_low": draws[BOOTSTRAP_LOW], "ci_high": draws[BOOTSTRAP_HIGH]}


def ratio(before: float, after: float) -> dict[str, Any]:
    if before <= 0:
        equal = after == before
        return {"before": before, "after": after, "ratio": 1.0 if equal else None,
                "change_percent": 0.0 if equal else None,
                "relative_change_defined": False, "zero_baseline_equal": equal,
                "zero_to_nonzero": not equal, "over_5_percent": not equal}
    value = after / before
    return {"before": before, "after": after, "ratio": value,
            "change_percent": (value - 1.0) * 100.0,
            "relative_change_defined": True, "zero_baseline_equal": False,
            "zero_to_nonzero": False, "over_5_percent": value > 1.05}


def historical_fixture(shape: str, mode: str) -> dict[str, Any]:
    path = ROOT / "docs/performance/results/change-0806/qualification" \
        / f"0-{shape}-{mode}-before.json"
    seal = read(ROOT / "docs/performance/results/change-0806/seal.json")
    relative = str(path.relative_to(ROOT / "docs/performance/results/change-0806"))
    require(seal.get("files", {}).get(relative) == sha(path),
            f"historical fixture seal changed: {relative}")
    value = read(path)
    require(value.get("schema") == "litchi.pptx.public-workflow-probe-0806.v1",
            f"historical fixture schema changed: {shape}/{mode}")
    return {"source": value["source"], "fixture": value["fixture"],
            "output": value["samples"][0]["output"],
            "verification": value["samples"][0]["verification"]}


def check_booleans(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label}: semantic verification malformed")
    booleans = [item for item in value.values() if isinstance(item, bool)]
    require(booleans and all(booleans), f"{label}: semantic proof failed")


def allocation_sample(value: Any, label: str) -> dict[str, float]:
    require(isinstance(value, dict) and value.get("status") == "measured"
            and value.get("scope") == "operation_global_system_allocator",
            f"{label}: allocation observer is not measured")
    result: dict[str, float] = {}
    for field in RAW_ALLOC:
        integer(value.get(field), f"{label}.{field}")
        result[field] = value[field]
    require(result["failed_allocation_calls"] == 0, f"{label}: failed allocation")
    require(result["live_bytes_after"] == result["live_bytes_before"]
            + result["allocated_bytes"] - result["deallocated_bytes"],
            f"{label}: live-byte conservation failed")
    require(result["region_peak_live_bytes"] >= result["live_bytes_before"]
            and result["region_peak_live_bytes"] >= result["live_bytes_after"]
            and result["peak_live_bytes_after"] >= result["peak_live_bytes_before"]
            and result["peak_live_bytes_after"] >= result["region_peak_live_bytes"],
            f"{label}: allocation peak ordering failed")
    result["net_live"] = result["live_bytes_after"] - result["live_bytes_before"]
    result["peak_above_entry"] = result["region_peak_live_bytes"] - result["live_bytes_before"]
    return result


def report_identity(entry: dict[str, Any], case: tuple[str, str, str], leg: str,
                    samples: int, warmup: int, allocation: bool,
                    label: str) -> dict[str, Any]:
    report_path = artifact(entry.get("report"), f"{label}: report")
    report = read(report_path)
    probe, shape, mode = case
    if probe == "synthetic":
        require(report.get("schema") == "litchi.pptx.public-workflow-probe-0806.v1"
                and report.get("tool") == "public-pptx-probe-0806"
                and report.get("shape") == shape and report.get("mode") == mode,
                f"{label}: synthetic probe identity changed")
        oracle = historical_fixture(shape, mode)
        require(report.get("source") == oracle["source"]
                and report.get("fixture") == oracle["fixture"],
                f"{label}: exact synthetic fixture changed")
        allocator = report.get("allocator")
        require(isinstance(allocator, dict)
                and allocator.get("allocator") ==
            ("CountingSystemAllocator(std::alloc::System)" if allocation
             else "Rust system allocator")
                and allocator.get("instrumentation") ==
            ("system_allocator_operation_scoped" if allocation else "none")
                and allocator.get("counter_revision") ==
            ("serialized_region_peak_v3" if allocation else None),
                f"{label}: allocator identity changed")
        require(report.get("warmup") == warmup
                and report.get("samples_requested") == samples,
                f"{label}: sample policy changed")
        rows = report.get("samples")
        require(isinstance(rows, list) and len(rows) == samples, f"{label}: samples changed")
        elapsed: list[float] = []
        alloc_values = {name: [] for name in ALLOC_METRICS}
        outputs = []
        expected_marker = mode in {"commit", "lifecycle"}
        vendor = shape in {"vendor", "unicode-vendor"}
        for index, row in enumerate(rows):
            require(isinstance(row, dict) and row.get("index") == index,
                    f"{label}: sample index changed")
            integer(row.get("elapsed_ns"), f"{label} elapsed", positive=True)
            elapsed.append(row["elapsed_ns"])
            require(row.get("source_sha256") == report["source"]["sha256"],
                    f"{label}: sample source changed")
            verification = row.get("verification")
            check_booleans(verification, f"{label} sample {index}")
            require(verification == oracle["verification"],
                    f"{label}: semantic verification oracle changed")
            require(row.get("output") == oracle["output"],
                    f"{label}: output identity changed")
            require(verification.get("marker_matches") is (True if expected_marker else None)
                    and verification.get("unknown_namespace_check") is
                    (True if vendor else None), f"{label}: marker/namespace oracle changed")
            outputs.append(row["output"])
            if allocation:
                values = allocation_sample(row.get("allocation"), f"{label} sample {index}")
                for name, item in values.items():
                    alloc_values[name].append(item)
            else:
                require(row.get("allocation") is None, f"{label}: native allocation present")
        require(len({json.dumps(item, sort_keys=True) for item in outputs}) == 1,
                f"{label}: nondeterministic output")
        return {"report": rel(report_path), "report_sha256": sha(report_path),
                "elapsed": elapsed, "stats": stats(elapsed),
                "allocation": alloc_values if allocation else None,
                "source": report["source"], "fixture": report["fixture"],
                "output": outputs[0]}

    require(report.get("schema") == "litchi.performance.0823.pptx-edit-trial.v1"
            and report.get("tool") == "pptx-edit-profile-0822"
            and report.get("base_revision") == BASE_SHORT
            and report.get("mode") == mode
            and report.get("input") == REAL_INPUT
            and report.get("reference") == REAL_REFERENCE
            and report.get("marker") == REAL_MARKER
            and report.get("target") == {"slide": 0, "shape": 0}
            and report.get("target_text") == REAL_MARKER
            and report.get("full_text_sha256") == REAL_FULL_TEXT_SHA256
            and report.get("full_text_digest") == REAL_FULL_TEXT_SHA256
            and report.get("slide_count") == 6,
            f"{label}: real probe identity changed")
    expected_verification = {
        "all_verified": True, "input_hash_verified": True,
        "reference_hash_verified": True, "output_hash_verified": True,
        "output_size_verified": True, "output_bytes_verified": True,
        "reopened": True, "marker_verified": True, "target_verified": True,
        "full_text_digest_verified": True, "slide_count_verified": True,
    }
    inputs = read(PACKET / "corpus-inputs.json")["real"]
    require(report.get("input", {}).get("bytes") == inputs["input"]["bytes"]
            and report.get("input", {}).get("sha256") == inputs["input"]["sha256"]
            and report.get("reference", {}).get("bytes") == inputs["reference"]["bytes"]
            and report.get("reference", {}).get("sha256") == inputs["reference"]["sha256"],
            f"{label}: real input/reference identity changed")
    output = report.get("output")
    require(isinstance(output, dict)
            and output.get("bytes") == inputs["reference"]["bytes"]
            and output.get("sha256") == inputs["reference"]["sha256"],
            f"{label}: real output identity changed")
    require(report.get("warmup") == warmup and report.get("samples_requested") == samples
            and report.get("warmup_verified") is True and report.get("all_verified") is True,
            f"{label}: real sample policy/oracle changed")
    rows = report.get("samples")
    elapsed_block = report.get("elapsed_ns")
    require(isinstance(rows, list) and len(rows) == samples
            and isinstance(elapsed_block, dict)
            and elapsed_block.get("sample_order") == list(range(samples)),
            f"{label}: real samples changed")
    elapsed: list[float] = []
    alloc_values = {name: [] for name in ALLOC_METRICS}
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and row.get("index") == index,
                f"{label}: real sample index changed")
        integer(row.get("elapsed_ns"), f"{label} elapsed", positive=True)
        require(row["elapsed_ns"] == elapsed_block["samples"][index]
                and row.get("output") == output,
                f"{label}: elapsed/output identity changed")
        elapsed.append(row["elapsed_ns"])
        require(row.get("verification") == expected_verification,
                f"{label}: semantic verification oracle changed")
        if allocation:
            values = allocation_sample(row.get("allocation"), f"{label} sample {index}")
            for name, item in values.items():
                alloc_values[name].append(item)
        else:
            require(row.get("allocation") is None, f"{label}: native allocation present")
        allocator = report.get("allocator")
        require(isinstance(allocator, dict)
                and allocator.get("instrumentation") ==
            ("system_allocator_operation_scoped" if allocation else "none")
                and allocator.get("allocator") ==
            ("CountingSystemAllocator(std::alloc::System)" if allocation
             else "Rust system allocator")
                and allocator.get("counter_revision") ==
            ("serialized_region_peak_v3" if allocation else None),
            f"{label}: allocator identity changed")
    return {"report": rel(report_path), "report_sha256": sha(report_path),
            "elapsed": elapsed, "stats": stats(elapsed),
            "allocation": alloc_values if allocation else None,
            "source": report["input"], "fixture": None, "output": output}


def expected_jobs(plan: dict[str, Any], lane: str, qualification_leg: str | None = None) -> list[dict[str, Any]]:
    section = plan[lane]
    if lane == "qualification":
        require(qualification_leg in LEGS, "qualification leg is missing")
        orders = [[qualification_leg]]
    else:
        orders = section["orders"]
    result = []
    for block, order in enumerate(orders):
        for case in CASES:
            for leg in order:
                result.append({"lane": lane, "block": block, "case": case, "leg": leg,
                               "samples": section["samples"], "warmup": section["warmup"]})
    return result


def receipt_case(row: dict[str, Any], label: str) -> tuple[str, str, str]:
    probe = row.get("probe_name", row.get("probe"))
    if isinstance(probe, dict):
        probe = row.get("case", {}).get("probe") if isinstance(row.get("case"), dict) else None
    case = row.get("case")
    if isinstance(case, dict):
        probe, shape, mode = case.get("probe"), case.get("shape", "real"), case.get("mode")
    else:
        shape, mode = row.get("shape", "real"), row.get("mode")
    if probe in {"synthetic", "real"}:
        return str(probe), str(shape if probe == "synthetic" else "real"), str(mode)
    if shape in SHAPES:
        return "synthetic", str(shape), str(mode)
    if row.get("input_selector") == "pinnedshapes":
        return "real", "real", "direct"
    fail(f"{label}: case identity missing")


def lane(plan: dict[str, Any], states: dict[str, Any], lane_name: str,
         *, directory_name: str | None = None,
         qualification_leg: str | None = None) -> list[dict[str, Any]]:
    directory = PACKET / (directory_name or lane_name)
    complete = read(directory / "complete.json")
    jobs = expected_jobs(plan, lane_name, qualification_leg)
    require(complete.get("schema") == f"litchi.performance.0823.{lane_name}.complete.v1"
            and complete.get("reports") == len(jobs)
            and complete.get("samples") == sum(job["samples"] for job in jobs)
            and complete.get("expected_reports") == len(jobs)
            and complete.get("expected_samples") == sum(job["samples"] for job in jobs)
            and complete.get("status") == "pass"
            and complete.get("lane") == lane_name,
            f"{lane_name}: completion receipt changed")
    if lane_name == "qualification":
        require(complete.get("leg") == qualification_leg, f"{lane_name}: qualification leg changed")
    else:
        require(complete.get("leg") is None, f"{lane_name}: unexpected qualification leg")
    plan_descriptor = complete.get("plan")
    freeze_descriptor = complete.get("freeze")
    require(artifact(plan_descriptor, f"{lane_name}: plan") and
            plan_descriptor.get("sha256") == sha(PACKET / "plan.json"),
            f"{lane_name}: plan custody changed")
    require(artifact(freeze_descriptor, f"{lane_name}: freeze") and
            freeze_descriptor.get("sha256") == sha(PACKET / "freeze.json"),
            f"{lane_name}: freeze custody changed")
    receipts_path = artifact(complete.get("receipts"), f"{lane_name}: receipts")
    receipts = read(receipts_path)
    require(isinstance(receipts, list) and len(receipts) == len(jobs),
            f"{lane_name}: receipt count changed")
    entries = []
    seen = set()
    seen_artifacts = set()
    for index, (row, job) in enumerate(zip(receipts, jobs)):
        label = f"{lane_name}/{index}"
        require(isinstance(row, dict)
                and row.get("schema") == "litchi.performance.0823.capture-receipt.v1"
                and row.get("lane") == lane_name
                and row.get("block") == job["block"]
                and row.get("leg") == job["leg"]
                and row.get("samples") == job["samples"]
                and row.get("warmup") == job["warmup"]
                and row.get("cpu") == plan["cpu"], f"{label}: receipt identity changed")
        case = receipt_case(row, label)
        require(case == job["case"], f"{label}: case order changed")
        key = (job["block"], case, job["leg"])
        require(key not in seen, f"{label}: duplicate process identity")
        seen.add(key)
        require(row.get("exit_code") == 0, f"{label}: failed process; retain its log")
        finite(row.get("started"), f"{label} started")
        finite(row.get("ended"), f"{label} ended")
        require(row["started"] <= row["ended"], f"{label}: receipt times reversed")
        log = artifact(row.get("log"), f"{label}: log")
        rss = artifact(row.get("rss"), f"{label}: RSS")
        report_descriptor = row.get("report")
        require(isinstance(report_descriptor, dict), f"{label}: report descriptor missing")
        artifact_paths = (str(packet_path(report_descriptor.get("path"), f"{label}: report")),
                          str(log), str(rss))
        require(not any(path in seen_artifacts for path in artifact_paths),
                f"{label}: duplicate report/log/RSS path")
        seen_artifacts.update(artifact_paths)
        rss_text = rss.read_text(encoding="utf-8").strip()
        require(rss_text.isdigit() and int(rss_text) > 0, f"{label}: RSS malformed")
        probe, shape, mode = case
        args = (["--mode", mode, "--shape", shape] if probe == "synthetic" else
                ["--mode", mode, "--input", REAL_INPUT["path"],
                 "--reference", REAL_REFERENCE["path"]])
        frozen_descriptor = row.get("frozen_inputs")
        require(isinstance(frozen_descriptor, dict)
                and artifact(frozen_descriptor, f"{label}: frozen inputs")
                and frozen_descriptor.get("sha256") == sha(PACKET / "freeze.json"),
                f"{label}: freeze receipt changed")
        expected_source = states[job["leg"]]["source"]["descriptor"]
        source = row.get("source")
        require(isinstance(source, dict)
                and source.get("bytes") == expected_source["bytes"]
                and source.get("sha256") == expected_source["sha256"],
                f"{label}: source receipt changed")
        binary = external_artifact(row.get("binary"), f"{label}: binary")
        variant = "native" if lane_name == "native" else "allocation"
        expected_binary = states[job["leg"]]["binaries"][f"{probe}-{variant}"]
        require(binary == expected_binary, f"{label}: binary is not from its leg build")
        expected_command = [
            "/usr/bin/time", "-f", "%M", "-o", str(rss), "taskset", "-c", "12",
            expected_binary["path"], *args, "--samples", str(job["samples"]),
            "--warmup", str(job["warmup"]), "--output", str(packet_path(
                report_descriptor.get("path"), f"{label}: report")),
        ]
        require(row.get("command") == expected_command,
                f"{label}: capture command changed")
        kind = lane_name == "native" and "native" or "allocation"
        outcome = report_identity(row, case, job["leg"], job["samples"], job["warmup"],
                                  kind == "allocation", label)
        entries.append({"identity": job, "case": case, "block": job["block"],
                        "leg": job["leg"], "report": outcome["report"],
                        "report_sha256": outcome["report_sha256"], "log": rel(log),
                        "rss_kib": int(rss_text), "outcome": outcome})
    require(len(seen) == len(jobs), f"{lane_name}: incomplete process identities")
    return entries


def group_stats(entries: list[dict[str, Any]], metrics: tuple[str, ...]) -> dict[str, Any]:
    grouped: dict[tuple[str, str, str], list[dict[str, Any]]] = {}
    for entry in entries:
        probe, shape, mode = entry["case"]
        grouped.setdefault((probe, shape, mode, entry["leg"]), []).append(entry)
    result: dict[str, Any] = {}
    for key, values in sorted(grouped.items()):
        values.sort(key=lambda item: item["block"])
        row: dict[str, Any] = {
            "processes": [{"block": item["block"], "report": item["report"],
                           "report_sha256": item["report_sha256"],
                           "stats": item["outcome"]["stats"], "rss_kib": item["rss_kib"]}
                          for item in values],
            "process_median": {metric: statistics.median(item["outcome"]["stats"][metric]
                                                          for item in values)
                               for metric in metrics},
            "spread_ratio": {metric: spread(item["outcome"]["stats"][metric] for item in values)
                             for metric in metrics},
            "rss_process_median": statistics.median(item["rss_kib"] for item in values),
            "rss_spread_ratio": spread(item["rss_kib"] for item in values),
        }
        row["flags"] = {
            "max_min_over_5_percent": {
                metric: row["spread_ratio"][metric] > .05 for metric in metrics
            },
            "rss_max_min_over_5_percent": row["rss_spread_ratio"] > .05,
            "p99_over_p50_over_5_percent": any(
                item["outcome"]["stats"]["p50"] > 0
                and item["outcome"]["stats"]["p99"] / item["outcome"]["stats"]["p50"] > 1.05
                for item in values),
        }
        if entries and values[0]["outcome"]["allocation"] is not None:
            row["allocation_process_median"] = {
                metric: statistics.median(
                    statistics.median(item["outcome"]["allocation"][metric])
                    for item in values)
                for metric in ALLOC_METRICS}
            row["allocation_spread_ratio"] = {
                metric: spread(statistics.median(item["outcome"]["allocation"][metric])
                               for item in values)
                for metric in ALLOC_METRICS}
            row["allocation_flags"] = {
                metric: row["allocation_spread_ratio"][metric] > .05
                for metric in ALLOC_METRICS
            }
        result["/".join(key)] = row
    return result


def paired(entries: list[dict[str, Any]], metrics: tuple[str, ...],
           *, allocation: bool = False) -> dict[str, Any]:
    lookup = {(item["case"], item["block"], item["leg"]): item for item in entries}
    result: dict[str, Any] = {}
    for case in CASES:
        block_values = sorted({item["block"] for item in entries if item["case"] == case})
        if not block_values:
            continue
        metric_rows: dict[str, Any] = {}
        for metric in metrics:
            rows = []
            ratios = []
            for block in block_values:
                before = lookup[(case, block, "before")]
                after = lookup[(case, block, "after")]
                if metric == "rss_kib":
                    left = before["rss_kib"]
                    right = after["rss_kib"]
                elif allocation:
                    left = statistics.median(before["outcome"]["allocation"][metric])
                    right = statistics.median(after["outcome"]["allocation"][metric])
                else:
                    left = before["outcome"]["stats"][metric]
                    right = after["outcome"]["stats"][metric]
                item = {"block": block, **ratio(left, right)}
                rows.append(item)
                if item["ratio"] is not None:
                    ratios.append(float(item["ratio"]))
            if not ratios:
                require(allocation, f"{case}/{metric}: no paired ratios")
            metric_rows[metric] = {
                "by_block": rows,
                "median_ratio": statistics.median(ratios) if ratios else None,
                "median_change_percent": ((statistics.median(ratios) - 1) * 100
                                           if ratios else None),
                "bootstrap": bootstrap(ratios) if ratios else None,
                "max_min_over_5_percent": any(item["over_5_percent"] for item in rows),
            }
        result["/".join(case)] = {"case": list(case), "blocks": len(block_values),
                                   "metrics": metric_rows,
                                   "comparison": "after/before paired by block"}
    return result


def rows_for(native: list[dict[str, Any]], allocation: list[dict[str, Any]]) -> list[dict[str, Any]]:
    native_groups = group_stats(native, TIMING_METRICS)
    allocation_groups = group_stats(allocation, TIMING_METRICS)
    native_pairs = paired(native, TIMING_METRICS + ("rss_kib",))
    allocation_pairs = paired(allocation, ALLOC_METRICS + ("rss_kib",), allocation=True)
    result = []
    for case in CASES:
        key = "/".join(case)
        result.append({"case": list(case), "native": {
            "before": native_groups.get(key + "/before"),
            "after": native_groups.get(key + "/after"),
        }, "allocation": {
            "before": allocation_groups.get(key + "/before"),
            "after": allocation_groups.get(key + "/after"),
        }, "native_paired": native_pairs.get(key),
           "allocation_paired": allocation_pairs.get(key)})
    return result


def guard(native_pairs: dict[str, Any], allocation_pairs: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    benefits, vetoes, resources, rss_reviews = [], [], [], []
    for key, group in sorted(native_pairs.items()):
        probe, _shape, mode = key.split("/", 2)
        timing = group["metrics"]["p50"]
        eligible = probe == "real" or mode in {"commit", "lifecycle"}
        if eligible and timing["median_ratio"] <= .97 and timing["bootstrap"]["ci_high"] < 1:
            benefits.append({"case": key, "median_ratio": timing["median_ratio"],
                             "bootstrap": timing["bootstrap"]})
        if timing["bootstrap"]["ci_low"] > 1.05:
            vetoes.append({"case": key, "median_ratio": timing["median_ratio"],
                           "bootstrap": timing["bootstrap"]})
        rss = group["metrics"]["rss_kib"]
        if rss["median_ratio"] > 1.05:
            rss_reviews.append({"case": key, "metric": "rss_kib",
                                "median_ratio": rss["median_ratio"],
                                "review_only": True})
    for key, group in sorted(allocation_pairs.items()):
        rss = group["metrics"]["rss_kib"]
        if rss["median_ratio"] is not None and rss["median_ratio"] > 1.05:
            rss_reviews.append({"case": key, "lane": "allocation", "metric": "rss_kib",
                                "median_ratio": rss["median_ratio"], "review_only": True})
        for metric in ("allocation_calls", "allocated_bytes", "net_live", "peak_above_entry"):
            rows = group["metrics"][metric]["by_block"]
            for item in rows:
                if item["after"] > item["before"]:
                    resources.append({"case": key, "metric": metric, "block": item["block"],
                                      "before": item["before"], "after": item["after"]})
    return {"benefits": benefits, "latency_vetoes": vetoes,
            "allocation_resource_violations": resources, "rss_reviews": rss_reviews,
            "benefit_satisfied": bool(benefits), "latency_veto_passed": not vetoes,
            "allocation_guard_passed": not resources,
            "adoption_eligible": bool(benefits) and not vetoes and not resources,
            "rss_is_review_only": True,
            "claims_excluded": ["universal", "cross-format", "cold-cache",
                                "tail-latency", "RSS saving", "historical speedup"]}


def native_csv(native: list[dict[str, Any]]) -> str:
    fields = ("case", "probe", "shape", "mode", "leg", "block", "report", "samples",
              "p50", "mean", "p95", "p99", "rss_kib", "p99_over_p50_over_5_percent")
    output = __import__("io").StringIO()
    writer = csv.DictWriter(output, fieldnames=fields, lineterminator="\n")
    writer.writeheader()
    for item in sorted(native, key=lambda value: (value["case"], value["block"], value["leg"])):
        probe, shape, mode = item["case"]
        values = item["outcome"]["stats"]
        writer.writerow({"case": "/".join(item["case"]), "probe": probe, "shape": shape,
                         "mode": mode, "leg": item["leg"], "block": item["block"],
                         "report": item["report"], "samples": values["count"],
                         "p50": values["p50"], "mean": values["mean"],
                         "p95": values["p95"], "p99": values["p99"],
                         "rss_kib": item["rss_kib"],
                         "p99_over_p50_over_5_percent":
                         values["p50"] > 0 and values["p99"] / values["p50"] > 1.05})
    return output.getvalue()


def markdown(rows: list[dict[str, Any]], guard_value: dict[str, Any], counts: dict[str, int]) -> str:
    lines = ["# 0823 scanner trial analysis", "", "Offline deterministic replay of retained packet evidence.", "",
             f"Reports: {counts['reports']}  Samples: {counts['samples']}", "",
             "## Native timing rows", "", "| case | before p50 | after p50 | after/before | CI95 | flags |",
             "|---|---:|---:|---:|---|---|"]
    for row in rows:
        group = row["native_paired"]
        timing = group["metrics"]["p50"]
        key = "/".join(row["case"])
        before = row["native"]["before"]["process_median"]["p50"]
        after = row["native"]["after"]["process_median"]["p50"]
        flags = {name: row["native"]["before"]["flags"][name]
                 or row["native"]["after"]["flags"][name]
                 for name in row["native"]["before"]["flags"]}
        lines.append(f"| {key} | {before} | {after} | {timing['median_ratio']:.9g} | "
                     f"[{timing['bootstrap']['ci_low']:.9g}, {timing['bootstrap']['ci_high']:.9g}] | "
                     f"spread={flags['p99_over_p50_over_5_percent']} |")
    lines.extend(["", "## Allocation guards", "",
                  "| case | calls before/after | bytes before/after | net live before/after | peak before/after |",
                  "|---|---:|---:|---:|---:|"])
    for row in rows:
        group = row["allocation_paired"]
        if group is None:
            continue
        metrics = group["metrics"]
        def cell(name: str) -> str:
            pair = metrics[name]["by_block"]
            return f"{statistics.median(x['before'] for x in pair):.9g}/{statistics.median(x['after'] for x in pair):.9g}"
        lines.append(f"| {'/'.join(row['case'])} | {cell('allocation_calls')} | {cell('allocated_bytes')} | "
                     f"{cell('net_live')} | {cell('peak_above_entry')} |")
    lines.extend(["", "## Disposition guards", "",
                  f"Benefit satisfied: `{guard_value['benefit_satisfied']}`; "
                  f"latency vetoes: `{len(guard_value['latency_vetoes'])}`; "
                  f"allocation violations: `{len(guard_value['allocation_resource_violations'])}`.", "",
                  "RSS increases above five percent are review flags only. The analysis makes no universal, cross-format, cold-cache, tail-latency, RSS-saving, or historical-speedup claim.", ""])
    return "\n".join(lines)


def analyze() -> dict[str, Any]:
    plan = frozen_inputs()
    custody = integrity(plan)
    quality = quality_receipt()
    before = source_state("before")
    after = source_state("after")
    diff = source_diff(before, after, plan)
    execution = chronology()
    states = {"before": before, "after": after}
    lanes = {
        "qualification": lane(plan, states, "qualification",
                               directory_name="qualification-before",
                               qualification_leg="before")
        + lane(plan, states, "qualification",
               directory_name="qualification-after", qualification_leg="after"),
        "native": lane(plan, states, "native"),
        "allocation": lane(plan, states, "allocation"),
    }
    capture_logs = [entry["log"] for name in ("qualification", "native", "allocation")
                    for entry in lanes[name]]
    require(len(capture_logs) == len(set(capture_logs)),
            "capture log path is reused across lanes")
    expected = plan["totals"]
    counts = {"qualification_reports": len(lanes["qualification"]),
              "qualification_samples": sum(len(x["outcome"]["elapsed"]) for x in lanes["qualification"]),
              "native_reports": len(lanes["native"]),
              "native_samples": sum(len(x["outcome"]["elapsed"]) for x in lanes["native"]),
              "allocation_reports": len(lanes["allocation"]),
              "allocation_samples": sum(len(x["outcome"]["elapsed"]) for x in lanes["allocation"])}
    counts["reports"] = sum(counts[key] for key in counts if key.endswith("_reports"))
    counts["samples"] = sum(counts[key] for key in counts if key.endswith("_samples"))
    require(counts == expected, "observed cardinality differs from plan")
    native_pairs = paired(lanes["native"], TIMING_METRICS + ("rss_kib",))
    allocation_pairs = paired(lanes["allocation"], ALLOC_METRICS + ("rss_kib",), allocation=True)
    row_values = rows_for(lanes["native"], lanes["allocation"])
    guard_value = guard(native_pairs, allocation_pairs, plan)
    return {"schema": ANALYSIS_SCHEMA, "plan_schema": PLAN_SCHEMA,
            "base": BASE, "counts": counts, "custody": custody,
            "quality": quality, "source": diff, "chronology": execution,
            "rows": row_values,
            "qualification": [{"case": list(x["case"]), "leg": x["leg"],
                               "block": x["block"], "report": x["report"],
                               "report_sha256": x["report_sha256"]}
                              for x in lanes["qualification"]],
            "native": {"reports": len(lanes["native"]), "samples": counts["native_samples"],
                       "processes": [{**x["identity"], "report": x["report"],
                                      "report_sha256": x["report_sha256"]} for x in lanes["native"]],
                       "paired": native_pairs},
            "allocation": {"reports": len(lanes["allocation"]), "samples": counts["allocation_samples"],
                           "processes": [{**x["identity"], "report": x["report"],
                                          "report_sha256": x["report_sha256"]} for x in lanes["allocation"]],
                           "paired": allocation_pairs},
            "decision": guard_value,
            "verification": {"all_raw_samples_checked": True,
                             "semantic_oracles_checked_for_both_families": True,
                             "exact_synthetic_fixture_oracles_checked": True,
                             "exact_real_file_oracles_checked": True,
                             "allocation_observer_status_checked": True,
                             "paired_cross_leg_expected": True,
                             "nearest_rank_quantiles": True,
                             "bootstrap_seed_and_endpoints_checked": True,
                             "no_aggregate_row_hidden": True,
                             "serial_chronology_checked": True,
                             "no_native_execution": True}}


def rendered(value: dict[str, Any], lanes: dict[str, list[dict[str, Any]]] | None = None) -> dict[str, str]:
    if lanes is None:
        plan = frozen_inputs()
        states = {"before": source_state("before"), "after": source_state("after")}
        lanes = {
            "qualification": lane(plan, states, "qualification",
                                   directory_name="qualification-before",
                                   qualification_leg="before")
            + lane(plan, states, "qualification",
                   directory_name="qualification-after", qualification_leg="after"),
            "native": lane(plan, states, "native"),
            "allocation": lane(plan, states, "allocation"),
        }
    return {"analysis.json": json.dumps(value, indent=2, sort_keys=True) + "\n",
            "native.csv": native_csv(lanes["native"]),
            "analysis.md": markdown(value["rows"], value["decision"], value["counts"])}


def replay(*, check: bool) -> dict[str, Any]:
    value = analyze()
    rendered_files = rendered(value)
    for name, encoded in rendered_files.items():
        path = PACKET / name
        if check:
            require(path.is_file() and path.read_text(encoding="utf-8") == encoded,
                    f"{name} does not replay byte-for-byte")
        else:
            require(not path.exists(), f"refusing to overwrite {name}")
            path.write_text(encoded, encoding="utf-8")
    return value


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    require(args in (["--write"], ["--check"]), "use exactly --write or --check")
    value = replay(check=args == ["--check"])
    print(f"0823 analysis PASS {value['counts']['reports']} reports/{value['counts']['samples']} samples")
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ReplayError, OSError, KeyError, TypeError, ValueError, AssertionError) as error:
        print(f"0823 analysis failed: {error}", file=sys.stderr)
        raise SystemExit(1)
