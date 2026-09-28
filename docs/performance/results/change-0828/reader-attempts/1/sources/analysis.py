"""Fail-closed offline replay for the 0828 PPTX edit profile.

The packet contains a deliberately small public-probe workload.  This module
only reads retained packet evidence: it never starts Cargo, a workload, the
probe, ``perf``, ``nm``, or ``objdump``.  The raw perf and decoded frame
artifacts remain the evidence; this reader derives deterministic summaries
from them and refuses stale or contradictory custody.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import importlib.util
import json
import math
import random
import re
import statistics
import subprocess
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any, Iterable

import custody


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_SCHEMA = "litchi.performance.0828.pptx-edit-profile.v1"
ANALYSIS_SCHEMA = "litchi.performance.0828.pptx-edit-profile-analysis.v1"
PROBE_REPORT_SCHEMA = "litchi.performance.0828.pptx-edit-profile.v1"
BASE_REVISION = "c990602492106d968898b7310bdd4bc17f3e4fcb"
SOURCE_COUNT = 9_197
TOOL_COUNT = 87
ARCHITECTURE_COUNT = 35
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_SEED = 828828
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9749
INPUT_BYTES = 68_822
INPUT_SHA256 = "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571"
REFERENCE_BYTES = 68_284
REFERENCE_SHA256 = "38c7fd3cc5037316e0a3a9c8cb6e6b013a95dd03b39bbd7fb357fc3cc3b9e5cf"
OWNER = "pptx_edit_profile_0828::edit_region_0828"
ARMS = ("control", "wrapped", "fp")
BINARY_BY_ARM = {"control": "ordinary", "wrapped": "ordinary", "fp": "fp"}
ARM_MODE = {"control": "direct", "wrapped": "wrapped", "fp": "wrapped"}
DSO_SUFFIXES = ("/ordinary", "/fp")
HEX = frozenset("0123456789abcdefABCDEF")

# The outer owner is deliberately broad enough to cover the complete edit
# helper.  These three nested wrappers are profiling markers only.  A sampled
# stack can contain none (unclassified) or more than one (ambiguous); the
# reader must retain those outcomes rather than forcing a phase assignment.
PHASE_WRAPPERS = {
    "capture": "pptx_edit_profile_0828::phase_opened_presentation_transaction_0828",
    "set_text": "pptx_edit_profile_0828::phase_set_shape_text_0828",
    "publish": "pptx_edit_profile_0828::phase_commit_apply_opened_presentation_commit_0828",
}
PHASE_ORDER = tuple(PHASE_WRAPPERS)


class ReplayError(RuntimeError):
    """Retained evidence is absent, malformed, stale, or contradictory."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
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
            and (value > 0 if positive else value >= 0), f"{label} is not an integer")


def finite(value: Any, label: str, *, positive: bool = False) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value))
            and (float(value) > 0 if positive else float(value) >= 0),
            f"{label} is not finite")


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def resolve_path(raw: Any, label: str, *, packet_bound: bool = False) -> Path:
    require(isinstance(raw, str) and raw, f"{label} path is missing")
    value = Path(raw)
    candidates = [value] if value.is_absolute() else [PACKET / value, ROOT / value]
    prefix = "docs/performance/results/change-0828/"
    if raw.startswith(prefix):
        candidates.insert(0, PACKET / raw[len(prefix):])
    for candidate in candidates:
        resolved = candidate.resolve(strict=False)
        if packet_bound and not resolved.is_relative_to(PACKET.resolve()):
            continue
        if resolved.is_file() and not resolved.is_symlink():
            return resolved
    resolved = candidates[0].resolve(strict=False)
    if packet_bound:
        require(resolved.is_relative_to(PACKET.resolve()),
                f"{label} escaped packet: {raw}")
    return resolved


def artifact(value: Any, label: str, *, packet_bound: bool = True,
             allow_missing: bool = False) -> Path:
    require(isinstance(value, dict), f"{label} descriptor is malformed")
    path = resolve_path(value.get("path"), label, packet_bound=packet_bound)
    integer(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")), f"{label}.sha256 is malformed")
    if not path.is_file() or path.is_symlink():
        require(allow_missing, f"missing {label}: {path}")
        return path
    require(path.stat().st_size == value["bytes"], f"{label}.bytes changed")
    require(sha256(path) == value["sha256"], f"{label}.sha256 changed")
    return path


def descriptor(value: Any, label: str, *, packet_bound: bool = True,
               allow_missing: bool = False) -> dict[str, Any]:
    path = artifact(value, label, packet_bound=packet_bound, allow_missing=allow_missing)
    return {"path": relative(path), "bytes": value["bytes"], "sha256": value["sha256"]}


def same_descriptor(left: Any, right: Any, label: str, *, packet_bound: bool = True) -> None:
    lp = artifact(left, f"{label} left", packet_bound=packet_bound)
    rp = artifact(right, f"{label} right", packet_bound=packet_bound)
    require(lp.resolve() == rp.resolve()
            and left.get("bytes") == right.get("bytes")
            and left.get("sha256") == right.get("sha256"),
            f"{label} identity changed")


def source_manifest(value: Any, label: str) -> dict[str, str]:
    require(isinstance(value, dict) and isinstance(value.get("files"), dict),
            f"{label} source manifest is malformed")
    files = value["files"]
    require(len(files) == SOURCE_COUNT
            and all(isinstance(name, str) and is_sha(digest)
                    for name, digest in files.items()),
            f"{label} source census changed")
    return dict(files)


def normalize_command(value: Any) -> list[str]:
    require(isinstance(value, list) and all(isinstance(item, str) for item in value),
            "command receipt is malformed")
    result = []
    marker = "/change-0828/"
    for item in value:
        if marker in item and item.startswith("/"):
            item = str(PACKET / item.split(marker, 1)[1])
        result.append(item)
    return result


def plan() -> dict[str, Any]:
    value = read_json(PACKET / "plan.json")
    require(value.get("schema") == PLAN_SCHEMA, "plan schema changed")
    require(value.get("base") == BASE_REVISION
            and value.get("target") == str(custody.TARGET)
            and value.get("scratch") is None
            and value.get("cpu") == 12, "plan base or execution host changed")
    expected_input = {
        "selector": "pinnedshapes",
        "path": "test-data/ooxml/pptx/shapes.pptx",
        "reference_selector": "0821realPPTXdefault",
        "reference_path": "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx",
        "reference_bytes": REFERENCE_BYTES,
        "reference_sha256": REFERENCE_SHA256,
        "source_path": "test-data/ooxml/pptx/shapes.pptx",
        "source_copy_path": "docs/performance/results/change-0821/artifacts/real-002-pptx/source.pptx",
        "source_bytes": INPUT_BYTES,
        "source_sha256": INPUT_SHA256,
    }
    require(value.get("input") == expected_input, "pinned PPTX input contract changed")
    require(value.get("probe") == {
        "package": "pptx-edit-profile-0828",
        "cargo_binary": "pptx-edit-profile-0828",
        "crate_owner": "pptx_edit_profile_0828",
        "owner_function": "edit_region_0828",
        "owner": OWNER,
        "phase_owners": dict(PHASE_WRAPPERS),
        "manifest": "docs/performance/results/change-0828/probe-src/Cargo.toml",
        "source_directory": "docs/performance/results/change-0828/probe-src",
        "modes": ["direct", "wrapped"],
        "input_selector": "pinnedshapes",
        "reference_selector": "0821realPPTXdefault",
        "input_path": "test-data/ooxml/pptx/shapes.pptx",
        "reference_path": "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx",
    }, "probe contract changed")
    require(value.get("arms") == {
        "control": {"binary": "ordinary", "mode": "direct",
                     "purpose": "direct public edit sequence without the wrapper"},
        "wrapped": {"binary": "ordinary", "mode": "wrapped",
                     "purpose": "the same public edit sequence through the named profiling wrapper"},
        "fp": {"binary": "fp", "mode": "wrapped",
               "purpose": "wrapped sequence in the frame-pointer build for native stack attribution"},
    }, "arm contract changed")
    require(value.get("qualification") == {
        "blocks": 1, "orders": [["control", "wrapped", "fp"]], "samples": 3,
        "warmup": 0, "reports": 3, "samples_total": 9,
        "purpose": "correctness and CLI-shape check before timed native blocks",
    }, "qualification contract changed")
    require(value.get("native") == {
        "blocks": 6,
        "orders": [["control", "wrapped", "fp"], ["wrapped", "fp", "control"],
                   ["fp", "control", "wrapped"], ["fp", "wrapped", "control"],
                   ["wrapped", "control", "fp"], ["control", "fp", "wrapped"]],
        "samples": 30, "warmup": 3, "reports": 18, "samples_total": 540,
        "rss": "maximum resident set size in KiB from /usr/bin/time",
    }, "native contract changed")
    require(value.get("perf") == {
        "binary": "fp", "mode": "wrapped", "input_selector": "pinnedshapes",
        "reference_selector": "0821realPPTXdefault", "repeats": 2, "samples": 2000,
        "warmup": 0, "cpu": 12, "event": "cycles:u", "frequency_hz": 997,
        "call_graph": "fp", "owner": OWNER, "phase_owners": dict(PHASE_WRAPPERS),
        "reports": 2, "samples_total": 4000,
        "input_path": "test-data/ooxml/pptx/shapes.pptx",
        "reference_path": "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx",
        "scope": "Whole-process sampled stacks and exact-owner attribution diagnostics only.",
    }, "perf contract changed")
    require(value.get("expected") == {
        "qualification_reports": 3, "qualification_samples": 9,
        "native_reports": 18, "native_samples": 540,
        "perf_reports": 2, "perf_samples": 4000,
        "total_reports": 23, "total_samples": 4549,
    }, "packet cardinality changed")
    stats = value.get("statistics")
    require(stats == {
        "native": ["p50", "p95", "p99", "mean", "spread", "rss"],
        "paired_ratios": ["wrapped/control", "fp/wrapped"],
        "bootstrap": {"resamples": BOOTSTRAP_RESAMPLES, "seed": BOOTSTRAP_SEED,
                       "statistic": "median", "sorted_zero_based_endpoints":
                       [BOOTSTRAP_LOW_RANK, BOOTSTRAP_HIGH_RANK]},
        "claims": ["descriptive native timing and RSS only",
                   "paired wrapped/control and fp/wrapped perturbation ratios",
                   "sampled cycles:u whole-process stack attribution with unresolved and lost-event diagnostics"],
        "adoption_threshold": None, "speedup_claim": False,
    }, "statistics contract changed")
    require(value.get("source") == {
        "production_changed": False, "runtime_harness_changed": False,
        "tool_changed": False, "production_file_count": SOURCE_COUNT,
        "tool_file_count": TOOL_COUNT, "probe_source_is_packet_local": True,
    }, "source contract changed")
    require(value.get("perf_denial") == {
        "status": "typed-unavailable", "must_retain_reason": True,
        "must_not_fabricate_profile": True,
    }, "perf denial contract changed")
    return value


def origin() -> dict[str, Any]:
    value = read_json(PACKET / "origin.json")
    require(value.get("schema") == "litchi.performance.0828.origin.v1"
            and value.get("base") == BASE_REVISION
            and value.get("previous_commit") == BASE_REVISION
            and value.get("production_changed") is False
            and value.get("runtime_harness_changed") is False
            and value.get("tool_changed") is False
            and value.get("tool_allowlist") == []
            and value.get("target") == str(custody.TARGET)
            and value.get("scratch") is None,
            "origin custody changed")
    require(value.get("unrelated") == custody.UNRELATED, "unrelated witness changed")
    reuse = value.get("quality_reuse")
    require(isinstance(reuse, dict)
            and reuse.get("schema") == "litchi.performance.0828.quality-reuse.v1"
            and reuse.get("mode") == "exact-source-0827-quality-receipts"
            and reuse.get("cargo_commands_executed") is False
            and reuse.get("prior_quality") == "docs/performance/results/change-0827/quality.json"
            and reuse.get("prior_seal") == "docs/performance/results/change-0827/seal.json"
            and reuse.get("prior_seal_commit") == BASE_REVISION
            and reuse.get("required_test_counts") ==
            {"passed": 641, "failed": 0, "ignored": 1, "suites": 28},
            "quality reuse contract changed")
    return value


def pinned_previous_seal() -> dict[str, Any]:
    """Check the prior seal from the exact 0828 base commit.

    Reading committed blobs keeps this proof independent of later rolling
    index edits in the live worktree.  It also makes the 0827 reader proof a
    custody input instead of a claim inferred from a current file.
    """
    previous = custody.PREVIOUS_PACKET
    commit = custody.PREVIOUS_COMMIT
    try:
        encoded = subprocess.check_output(
            ["git", "show", f"{commit}:docs/performance/results/change-0827/seal.json"],
            cwd=ROOT,
        )
    except (OSError, subprocess.CalledProcessError) as error:
        fail(f"cannot read pinned 0827 seal: {error}")
    seal = json.loads(encoded.decode("utf-8"))
    require(seal.get("schema") == "litchi.performance.0827.seal.v1",
            "pinned 0827 seal schema changed")
    files = seal.get("files")
    require(isinstance(files, dict) and len(files) >= 770,
            "pinned 0827 seal file map changed")
    for name, digest in files.items():
        path = Path(name)
        require(not path.is_absolute() and ".." not in path.parts and is_sha(digest),
                f"malformed pinned 0827 seal entry: {name}")
        try:
            blob = subprocess.check_output(["git", "show", f"{commit}:{name}"], cwd=ROOT)
        except (OSError, subprocess.CalledProcessError) as error:
            fail(f"missing pinned 0827 payload {name}: {error}")
        require(hashlib.sha256(blob).hexdigest() == digest,
                f"pinned 0827 payload changed: {name}")
    current = previous / "seal.json"
    require(current.is_file() and sha256(current) == hashlib.sha256(encoded).hexdigest(),
            "live 0827 seal differs from pinned base")
    analysis_blob = subprocess.check_output(
        ["git", "show", f"{commit}:docs/performance/results/change-0827/analysis.json"], cwd=ROOT)
    analysis = json.loads(analysis_blob.decode("utf-8"))
    require(analysis.get("counts") == {
                "native_reports": 144, "native_samples": 4320,
                "observer_reports": 48, "observer_samples": 144,
                "qualification_reports": 24, "qualification_samples": 24,
                "reports": 216, "samples": 4488,
            }, "pinned 0827 reader proof changed")
    return {"schema": seal["schema"], "commit": commit,
            "path": str(current), "bytes": len(encoded),
            "sha256": hashlib.sha256(encoded).hexdigest(),
            "files": len(files)}


def load_custody() -> dict[str, Any]:
    p = plan()
    o = origin()
    root_inputs = custody.assert_root_inputs()
    require(root_inputs == {
        "Cargo.lock": sha256(ROOT / "Cargo.lock"),
        "rustfmt.toml": sha256(ROOT / "rustfmt.toml"),
        "tools/perf-baseline/Cargo.lock": sha256(ROOT / "tools/perf-baseline/Cargo.lock"),
    }, "root input custody changed")
    locks = custody.lock_identity()
    architecture = custody.architecture_hashes()
    require(len(architecture) == ARCHITECTURE_COUNT, "architecture census changed")
    corpus = custody.assert_corpus_inputs()
    provenance = custody.assert_provenance(corpus)
    host = custody.assert_host()
    unrelated = custody.assert_unrelated()
    previous = pinned_previous_seal()
    source = custody.source()
    tool = custody.tool_source()
    require(len(source["files"]) == SOURCE_COUNT and len(tool) == TOOL_COUNT,
            "live source census changed")
    require(source["files"] == custody.read(custody.PREVIOUS_PACKET / "freeze.json")["source"]["files"],
            "production source differs from committed repair source")
    require(tool == custody.read(custody.PREVIOUS_PACKET / "freeze.json")["tool"],
            "tool source differs from committed repair source")
    ancestry = subprocess.run(["git", "merge-base", "--is-ancestor", BASE_REVISION,
                               source["revision"]], cwd=ROOT, check=False)
    require(ancestry.returncode == 0, "current HEAD is not based on measured revision")
    probe = custody.probe_files()
    return {"plan": p, "origin": o, "root_inputs": root_inputs,
            "locks": locks, "architecture": architecture, "corpus": corpus,
            "provenance": provenance, "host": host, "unrelated": unrelated,
            "previous_seal": previous, "source": source, "tool": tool,
            "probe": probe, "static_packet": custody.packet_hashes(),
            "drivers": custody.driver_hashes()}


def validate_source_receipt(value: Any, cv: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(value, dict)
            and isinstance(value.get("production"), dict)
            and isinstance(value.get("tool"), dict), f"{label} source receipt malformed")
    production = value["production"]
    files = source_manifest(production, f"{label} production")
    require(files == cv["source"]["files"], f"{label} production source changed")
    require(production.get("revision") in {BASE_REVISION, cv["source"]["revision"]},
            f"{label} production revision changed")
    tool = value["tool"]
    require(len(tool) == TOOL_COUNT and tool == cv["tool"],
            f"{label} tool source changed")
    return {"production": {"revision": production.get("revision"), "files": files},
            "tool": dict(tool)}


def external_descriptor(value: Any, label: str, *, base: Path | None = None) -> Path:
    require(isinstance(value, dict) and isinstance(value.get("path"), str),
            f"{label}: external descriptor malformed")
    raw = Path(value["path"])
    path = (base / raw if base is not None and not raw.is_absolute() else raw)
    path = path.resolve(strict=False)
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    integer(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")) and path.stat().st_size == value["bytes"]
            and sha256(path) == value["sha256"], f"{label}: identity changed")
    return path


def validate_nested_external_descriptors(value: Any, label: str) -> None:
    """Hash-check every retained descriptor in the imported 0821 witness."""
    if isinstance(value, dict):
        if {"path", "bytes", "sha256"}.issubset(value):
            base = custody.PREVIOUS_PACKET if label.endswith(".prior_build") else None
            external_descriptor(value, label, base=base)
            return
        for key, item in value.items():
            validate_nested_external_descriptors(item, f"{label}.{key}")
    elif isinstance(value, list):
        for index, item in enumerate(value):
            validate_nested_external_descriptors(item, f"{label}[{index}]")


def load_quality(cv: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0828.quality.v1"
            and value.get("status") == "pass", "quality result is not a pass")
    reuse = value.get("production_reuse")
    require(isinstance(reuse, dict)
            and reuse.get("schema") == "litchi.performance.0828.quality-reuse.v1"
            and reuse.get("mode") == "exact-source-0827-quality-receipts"
            and reuse.get("cargo_commands_executed") is False
            and isinstance(reuse.get("prior_quality"), dict)
            and isinstance(reuse.get("prior_freeze"), dict)
            and isinstance(reuse.get("prior_seal"), dict)
            and isinstance(reuse.get("prior_quality_receipts"), list)
            and len(reuse["prior_quality_receipts"]) == 6,
            "quality reuse witness changed")
    require(isinstance(reuse.get("source"), dict)
            and reuse["source"].get("files") == cv["source"]["files"]
            and reuse.get("tool") == cv["tool"]
            and reuse.get("architecture") == cv["architecture"]
            and reuse.get("root_inputs") == cv["root_inputs"]
            and reuse.get("locks") == cv["locks"]
            and reuse.get("corpus") == cv["corpus"]
            and reuse.get("provenance") == cv["provenance"]
            and reuse.get("host") == cv["host"]
            and reuse.get("unrelated") == cv["unrelated"],
            "quality reused custody witness changed")
    require(reuse.get("test_counts") ==
            {"passed": 641, "failed": 0, "ignored": 1, "suites": 28},
            "quality test count witness changed")
    validate_nested_external_descriptors(reuse, "quality.production_reuse")
    prior_quality_path = external_descriptor(reuse["prior_quality"],
                                             "quality reused prior quality")
    prior_freeze_path = external_descriptor(reuse["prior_freeze"],
                                            "quality reused prior freeze")
    prior_seal_path = external_descriptor(reuse["prior_seal"],
                                          "quality reused prior seal")
    require(str(prior_quality_path).endswith("docs/performance/results/change-0827/quality.json")
            and str(prior_freeze_path).endswith("docs/performance/results/change-0827/freeze.json")
            and str(prior_seal_path).endswith("docs/performance/results/change-0827/seal.json"),
            "quality prior packet paths changed")
    for index, item in enumerate(reuse["prior_quality_receipts"]):
        receipt_path = external_descriptor(item, f"quality prior receipt {index}")
        require(str(receipt_path).startswith(str(custody.PREVIOUS_PACKET)),
                f"quality prior receipt {index} escaped 0827 packet")
    tests = reuse["test_counts"]
    fresh = value.get("probe")
    require(isinstance(fresh, dict)
            and fresh.get("schema") == "litchi.performance.0828.probe-quality.v1"
            and fresh.get("status") == "pass"
            and fresh.get("gate_count") in {5, 6},
            "fresh probe-quality witness is missing")
    gates = fresh.get("rows")
    require(isinstance(gates, list) and len(gates) == fresh["gate_count"],
            "probe quality gate cardinality changed")
    expected_names = ["fmt", "check", "tests", "clippy", "doc"]
    require(fresh.get("gate_count") == len(expected_names)
            and [row.get("name") if isinstance(row, dict) else None for row in gates]
            == expected_names,
            "probe quality gate names changed")
    for index, row in enumerate(gates):
        require(isinstance(row, dict) and row.get("gate") == index + 1
                and row.get("name") == expected_names[index]
                and row.get("exit_code") == 0
                and isinstance(row.get("command"), list)
                and isinstance(row.get("environment"), dict),
                f"probe quality gate {index + 1} changed")
        if isinstance(row.get("log"), dict):
            artifact(row["log"], f"probe quality gate {index + 1} log")
    require(isinstance(fresh.get("tests"), list)
            and all(isinstance(item, dict) and item.get("failed") == 0
                    for item in fresh["tests"]),
            "fresh probe test-count witness changed")
    # Every current quality receipt carries these exact custody descriptors.
    # Do not accept the historical import/load aliases used by older packets.
    require(set(value) == {
        "schema", "status", "production_reuse", "probe", "source",
        "frozen_inputs", "checks", "root_inputs", "locks", "architecture",
        "corpus", "provenance", "host", "unrelated", "probe_source",
        "target", "started", "ended",
    }, "quality result fields changed")
    source_ref = value.get("source")
    frozen_ref = value.get("frozen_inputs")
    source_path = artifact(source_ref, "quality source")
    frozen_path = artifact(frozen_ref, "quality frozen inputs")
    validate_source_receipt(read_json(source_path), cv, "quality")
    frozen = read_json(frozen_path)
    require(isinstance(frozen.get("source"), dict)
            and frozen["source"].get("files") == cv["source"]["files"]
            and frozen.get("tool") == cv["tool"]
            and frozen.get("probe") == cv["probe"]
            and frozen.get("drivers") == cv["drivers"],
            "quality frozen probe/driver custody changed")
    require(value.get("probe_source") == frozen["probe"],
            "quality probe source custody changed")
    checks_path = artifact(value.get("checks"), "quality checks")
    checks = read_json(checks_path)
    require(checks.get("schema") == "litchi.performance.0828.probe-quality-checks.v1"
            and checks.get("rows") == gates
            and checks.get("source") == value["source"]
            and checks.get("frozen_inputs") == value["frozen_inputs"],
            "quality checks witness changed")
    require(value.get("target") == str(custody.TARGET),
            "quality target changed")
    finite(value.get("started"), "quality start")
    finite(value.get("ended"), "quality end")
    require(value["started"] <= value["ended"], "quality timestamps reversed")
    return {"path": relative(path), "sha256": sha256(path),
            "status": value["status"], "mode": reuse["mode"],
            "production_test_counts": tests, "probe_gate_count": len(gates),
            "fresh_probe_gate_count": len(gates), "raw": value,
            "source": source_path, "frozen": frozen_path}


def binary_identity(value: Any, label: str, cleanup: dict[str, Any] | None) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} binary identity malformed")
    raw = value.get("artifact") if isinstance(value.get("artifact"), dict) else value
    require(isinstance(raw, dict) and isinstance(raw.get("path"), str)
            and Path(raw["path"]).is_absolute() and is_sha(raw.get("sha256")),
            f"{label} binary descriptor malformed")
    integer(raw.get("bytes"), f"{label} binary bytes", positive=True)
    path = Path(raw["path"])
    if path.is_file() and not path.is_symlink():
        require(path.stat().st_size == raw["bytes"] and sha256(path) == raw["sha256"],
                f"{label} binary identity changed")
    else:
        require(cleanup is not None and cleanup.get("target_removed") is True,
                f"{label} binary missing without cleanup witness")
        removed = cleanup.get("removed_binaries", [])
        require(isinstance(removed, list)
                and any(isinstance(row, dict) and row.get("path") == str(path)
                        and row.get("bytes") == raw["bytes"]
                        and row.get("sha256") == raw["sha256"] for row in removed),
                f"{label} exact removed binary witness missing")
    return {"path": str(path), "bytes": raw["bytes"], "sha256": raw["sha256"]}


def cleanup_value() -> dict[str, Any] | None:
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return None
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0828.cleanup.v1"
            and value.get("target") == str(custody.TARGET)
            and value.get("target_removed") is True
            and value.get("scratch") is None,
            "cleanup witness changed")
    binaries = value.get("removed_binaries")
    require(isinstance(binaries, list) and len(binaries) == 2,
            "cleanup binary cardinality changed")
    for row in binaries:
        require(isinstance(row, dict) and set(row) == {"path", "bytes", "sha256"}
                and Path(row["path"]).is_absolute() and row["bytes"] > 0
                and is_sha(row["sha256"]), "cleanup binary descriptor malformed")
    artifact(value.get("source"), "cleanup source")
    artifact(value.get("frozen_inputs"), "cleanup frozen inputs")
    require(isinstance(value.get("removed"), list), "cleanup directory witness missing")
    finite(value.get("started"), "cleanup start")
    finite(value.get("ended"), "cleanup end")
    require(value["started"] <= value["ended"], "cleanup timestamps reversed")
    return value


def load_build(cv: dict[str, Any], quality: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "build.json"
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0828.build.v1",
            "build schema changed")
    source_ref = value.get("source")
    source_path = artifact(source_ref, "build source")
    source = validate_source_receipt(read_json(source_path), cv, "build")
    frozen_ref = value.get("frozen_inputs")
    frozen_path = artifact(frozen_ref, "build frozen inputs")
    frozen = read_json(frozen_path)
    require(isinstance(frozen, dict)
            and frozen.get("schema") == "litchi.performance.0828.frozen-inputs.v1",
            "build frozen-input schema changed")
    for key in ("root_inputs", "locks", "architecture", "corpus", "provenance",
                "host", "unrelated"):
        if key in frozen and key in cv:
            require(frozen[key] == cv[key], f"build frozen {key} custody changed")
    require(frozen.get("static_packet") == cv["static_packet"],
            "build frozen static packet changed")
    require(frozen.get("previous_seal") == cv["previous_seal"],
            "build frozen previous_seal custody changed")
    require(frozen.get("plan") == sha256(PACKET / "plan.json")
            and frozen.get("origin") == sha256(PACKET / "origin.json"),
            "build frozen plan/origin custody changed")
    require(frozen.get("probe") == cv["probe"]
            and frozen.get("drivers") == cv["drivers"]
            and isinstance(frozen.get("source"), dict)
            and frozen["source"].get("files") == cv["source"]["files"]
            and frozen.get("tool") == cv["tool"],
            "build frozen probe/driver custody changed")
    cleanup = cleanup_value()
    binaries = value.get("binaries")
    require(isinstance(binaries, dict) and set(binaries) == {"ordinary", "fp"},
            "build binary set changed")
    verified: dict[str, Any] = {}
    for name in ("ordinary", "fp"):
        row = binaries[name]
        require(isinstance(row, dict)
                and row.get("cargo_name") == "pptx-edit-profile-0828"
                and row.get("features", []) == []
                and row.get("rustflags") ==
                (None if name == "ordinary" else "-C force-frame-pointers=yes"),
                f"build binary metadata changed: {name}")
        verified[name] = {"cargo_name": row["cargo_name"],
                          "features": list(row.get("features", [])),
                          "artifact": binary_identity(row, f"build {name}", cleanup)}
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 2, "build row cardinality changed")
    seen = set()
    for row in rows:
        require(isinstance(row, dict) and row.get("variant") in {"ordinary", "fp"}
                and row["variant"] not in seen and row.get("exit_code") == 0,
                "build command result changed")
        seen.add(row["variant"])
        finite(row.get("started"), f"build {row['variant']} start")
        finite(row.get("ended"), f"build {row['variant']} end")
        require(row["started"] <= row["ended"], "build timestamps reversed")
        if isinstance(row.get("log"), dict):
            artifact(row["log"], f"build {row['variant']} log")
        flags = row.get("rustflags")
        expected_flags = None if row["variant"] == "ordinary" else "-C force-frame-pointers=yes"
        require(flags == expected_flags, f"build {row['variant']} RUSTFLAGS changed")
    require(seen == {"ordinary", "fp"}, "build rows incomplete")
    require(value.get("quality") == quality["raw"]
            or value.get("quality", {}).get("sha256") == quality["sha256"]
            or "quality" not in value, "build quality witness changed")
    return {"path": relative(path), "sha256": sha256(path), "source": source,
            "source_path": relative(source_path), "frozen": frozen,
            "frozen_path": relative(frozen_path), "binaries": verified,
            "raw": value, "cleanup": cleanup}


REPORT_KEYS = frozenset({
    "schema", "tool", "base_revision", "mode", "timing_scope", "input",
    "reference", "output", "marker", "target", "target_text", "full_text_sha256",
    "full_text_digest", "slide_count", "warmup", "samples_requested",
    "warmup_verified", "all_verified", "elapsed_ns", "samples",
})
SAMPLE_KEYS = frozenset({"index", "elapsed_ns", "output", "verification"})
VERIFICATION_KEYS = frozenset({
    "all_verified", "input_hash_verified", "reference_hash_verified",
    "output_hash_verified", "output_size_verified", "output_bytes_verified", "reopened",
    "marker_verified", "target_verified", "full_text_digest_verified",
    "slide_count_verified",
})


def identity(value: Any, label: str, *, bytes_: int, sha: str) -> dict[str, Any]:
    require(isinstance(value, dict) and set(value) == {"bytes", "sha256"},
            f"{label} identity malformed")
    require(value.get("bytes") == bytes_ and value.get("sha256") == sha,
            f"{label} identity changed")
    return {"bytes": bytes_, "sha256": sha}


def path_ends_with(value: Any, suffix: str, label: str) -> None:
    require(isinstance(value, str) and (value == suffix or value.endswith("/" + suffix)
                                        or value.endswith("\\" + suffix)),
            f"{label} path changed")


def report_values(report: dict[str, Any], *, arm: str, samples: int,
                  warmup: int, label: str) -> list[int]:
    # The probe report wire shape is frozen, while future diagnostic fields
    # may be appended by the root-owned harness.  Require every frozen field
    # and reject missing/unknown nested fields where they affect the oracle.
    require(REPORT_KEYS.issubset(report), f"{label}: report fields changed")
    require(report.get("schema") == PROBE_REPORT_SCHEMA
            and report.get("tool") == "pptx-edit-profile-0828"
            and report.get("base_revision") == BASE_REVISION
            and report.get("mode") == ARM_MODE[arm],
            f"{label}: probe identity changed")
    require(isinstance(report.get("timing_scope"), str)
            and "outside the clock" in report["timing_scope"]
            and "semantic readback" in report["timing_scope"],
            f"{label}: timing scope changed")
    source = report.get("input")
    require(isinstance(source, dict)
            and {"path", "bytes", "sha256"}.issubset(source),
            f"{label}: input identity malformed")
    path_ends_with(source.get("path"), "test-data/ooxml/pptx/shapes.pptx",
                   f"{label}: input")
    require(source.get("bytes") == INPUT_BYTES and source.get("sha256") == INPUT_SHA256,
            f"{label}: input identity changed")
    reference = report.get("reference")
    require(isinstance(reference, dict)
            and {"path", "bytes", "sha256"}.issubset(reference),
            f"{label}: reference identity malformed")
    path_ends_with(reference.get("path"),
                   "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx",
                   f"{label}: reference")
    require(reference.get("bytes") == REFERENCE_BYTES
            and reference.get("sha256") == REFERENCE_SHA256,
            f"{label}: reference identity changed")
    identity(report.get("output"), f"{label}: output",
             bytes_=REFERENCE_BYTES, sha=REFERENCE_SHA256)
    require(report.get("marker") == "litchi-perf-0638-ordinary-save"
            and report.get("target") == {"slide": 0, "shape": 0}
            and report.get("target_text") == report.get("marker")
            and report.get("slide_count") == 6,
            f"{label}: edit oracle changed")
    require(is_sha(report.get("full_text_sha256"))
            and report.get("full_text_digest") == report.get("full_text_sha256"),
            f"{label}: semantic digest changed")
    require(report.get("warmup") == warmup
            and report.get("samples_requested") == samples
            and report.get("warmup_verified") is True
            and report.get("all_verified") is True,
            f"{label}: sample policy or verification changed")
    elapsed = report.get("elapsed_ns")
    require(isinstance(elapsed, dict)
            and set(elapsed) == {"unit", "samples", "sample_order"}
            and elapsed.get("unit") == "ns"
            and elapsed.get("sample_order") == list(range(samples)),
            f"{label}: elapsed vector malformed")
    values = elapsed.get("samples")
    require(isinstance(values, list) and len(values) == samples,
            f"{label}: elapsed sample count changed")
    for index, value in enumerate(values):
        integer(value, f"{label}: elapsed[{index}]", positive=True)
    rows = report.get("samples")
    require(isinstance(rows, list) and len(rows) == samples,
            f"{label}: report sample count changed")
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and SAMPLE_KEYS.issubset(row)
                and row.get("index") == index
                and row.get("elapsed_ns") == values[index],
                f"{label}: sample {index} shape changed")
        identity(row.get("output"), f"{label}: sample {index} output",
                 bytes_=REFERENCE_BYTES, sha=REFERENCE_SHA256)
        verification = row.get("verification")
        require(isinstance(verification, dict) and VERIFICATION_KEYS.issubset(verification)
                and verification.get("all_verified") is True
                and all(verification.get(key) is True for key in VERIFICATION_KEYS
                        if key != "all_verified"),
                f"{label}: sample {index} verification changed")
    return list(values)


def arm_from_row(row: dict[str, Any], label: str) -> str:
    for key in ("arm", "variant", "leg", "kind"):
        value = row.get(key)
        if isinstance(value, str) and value in ARMS:
            return value
    raise ReplayError(f"{label}: arm is missing")


def binary_descriptor_from_receipt(row: dict[str, Any], label: str) -> dict[str, Any]:
    value = row.get("binary")
    require(isinstance(value, dict), f"{label}: binary receipt missing")
    raw = value.get("artifact") if isinstance(value.get("artifact"), dict) else value
    require(isinstance(raw, dict) and isinstance(raw.get("path"), str),
            f"{label}: binary receipt malformed")
    return {"path": str(Path(raw["path"])), "bytes": raw.get("bytes"),
            "sha256": raw.get("sha256")}


def validate_receipt_common(row: dict[str, Any], *, block: int, arm: str,
                            lane: str, samples: int, warmup: int,
                            build: dict[str, Any], label: str) -> tuple[Path, int]:
    require(isinstance(row, dict)
            and row.get("schema") == "litchi.performance.0828.capture-receipt.v1"
            and row.get("lane") == lane
            and row.get("block") == block and row.get("label") == f"{block:02d}-{arm}"
            and arm_from_row(row, label) == arm
            and row.get("mode") == ARM_MODE[arm]
            and row.get("samples") == samples and row.get("warmup") == warmup
            and row.get("cpu") == 12
            and row.get("exit_code") == 0, f"{label}: receipt matrix changed")
    finite(row.get("started"), f"{label}: start")
    finite(row.get("ended"), f"{label}: end")
    require(row["started"] <= row["ended"], f"{label}: receipt timestamps reversed")
    report_ref = row.get("report")
    report_path = artifact(report_ref, f"{label}: report")
    report = read_json(report_path)
    values = report_values(report, arm=arm, samples=samples, warmup=warmup, label=label)
    expected_binary = build["binaries"][BINARY_BY_ARM[arm]]["artifact"]
    observed_binary = binary_descriptor_from_receipt(row, label)
    require(observed_binary == expected_binary, f"{label}: binary identity changed")
    input_value = row.get("input")
    require(isinstance(input_value, dict)
            and input_value.get("bytes") == INPUT_BYTES
            and input_value.get("sha256") == INPUT_SHA256,
            f"{label}: input receipt identity changed")
    path_ends_with(input_value.get("path"), "test-data/ooxml/pptx/shapes.pptx",
                   f"{label}: input receipt")
    reference_value = row.get("reference")
    require(isinstance(reference_value, dict)
            and reference_value.get("bytes") == REFERENCE_BYTES
            and reference_value.get("sha256") == REFERENCE_SHA256,
            f"{label}: reference receipt identity changed")
    path_ends_with(reference_value.get("path"),
                   "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx",
                   f"{label}: reference receipt")
    for key in ("source", "probe", "frozen_inputs"):
        if isinstance(row.get(key), dict):
            artifact(row[key], f"{label}: {key}")
    if isinstance(row.get("rss"), dict):
        rss_path = artifact(row["rss"], f"{label}: RSS")
        rss_text = rss_path.read_text(encoding="utf-8").strip()
        require(rss_text.isdigit() and int(rss_text) > 0, f"{label}: RSS malformed")
        rss = int(rss_text)
    elif isinstance(row.get("rss_kib"), int):
        rss = row["rss_kib"]
        require(rss > 0, f"{label}: RSS malformed")
    else:
        fail(f"{label}: RSS receipt missing")
    if isinstance(row.get("log"), dict):
        log_path = artifact(row["log"], f"{label}: log")
        require(not log_path.read_bytes(), f"{label}: workload log is not empty")
    command = normalize_command(row.get("command"))
    require(len(command) >= 1 and command[0] == "/usr/bin/time"
            and "taskset" in command and expected_binary["path"] in command
            and "--input" in command and "--reference" in command
            and "--samples" in command and "--warmup" in command and "--mode" in command
            and "--output" in command,
            f"{label}: probe command boundary changed")
    require(command[command.index("--mode") + 1] == ARM_MODE[arm]
            and command[command.index("--input") + 1] == "test-data/ooxml/pptx/shapes.pptx"
            and command[command.index("--reference") + 1] ==
            "docs/performance/results/change-0821/artifacts/real-002-pptx/default.pptx"
            and int(command[command.index("--samples") + 1]) == samples
            and int(command[command.index("--warmup") + 1]) == warmup
            and command[command.index("--output") + 1] == str(report_path),
            f"{label}: probe command parameters changed")
    return report_path, rss


def load_lane(name: str, cv: dict[str, Any], build: dict[str, Any],
              *, expected_reports: int, expected_samples: int,
              samples: int, warmup: int, orders: list[list[str]]) -> list[dict[str, Any]]:
    directory = PACKET / name
    complete_path = directory / "complete.json"
    complete = read_json(complete_path)
    require(complete.get("schema") == f"litchi.performance.0828.{name}.complete.v1"
            and complete.get("status") == "pass"
            and complete.get("blocks") == len(orders)
            and complete.get("reports") == expected_reports
            and complete.get("samples") == expected_samples,
            f"{name} completion cardinality changed")
    require(complete.get("expected_reports") == expected_reports
            and complete.get("expected_samples") == expected_samples,
            f"{name} completion expectation changed")
    require(complete.get("plan_sha256") == sha256(PACKET / "plan.json")
            and complete.get("build_sha256") == sha256(PACKET / "build.json"),
            f"{name} completion custody changed")
    if isinstance(complete.get("receipts"), dict):
        receipts_path = artifact(complete["receipts"], f"{name} receipts")
    else:
        receipts_path = directory / "receipts.json"
    receipts = read_json(receipts_path)
    require(isinstance(receipts, list) and len(receipts) == expected_reports,
            f"{name} receipt cardinality changed")
    jobs = [(block, arm) for block, order in enumerate(orders) for arm in order]
    require(len(jobs) == expected_reports, f"{name} job matrix changed")
    rows: list[dict[str, Any]] = []
    previous = float("-inf")
    for receipt, (block, arm) in zip(receipts, jobs):
        label = f"{name}/{block}/{arm}"
        require(isinstance(receipt, dict), f"{label}: receipt malformed")
        require(previous <= receipt.get("started", float("-inf")),
                f"{label}: receipt order changed")
        previous = receipt["ended"]
        report_path, rss = validate_receipt_common(
            receipt, block=block, arm=arm, lane=name, samples=samples,
            warmup=warmup, build=build, label=label,
        )
        report = read_json(report_path)
        values = list(report["elapsed_ns"]["samples"])
        rows.append({
            "block": block, "arm": arm, "mode": ARM_MODE[arm],
            "binary": BINARY_BY_ARM[arm], "values": values,
            "stats": distribution(values), "rss_kib": rss,
            "report": relative(report_path), "report_sha256": sha256(report_path),
            "receipt": receipt,
        })
    require(sum(len(row["values"]) for row in rows) == expected_samples,
            f"{name} sample cardinality changed")
    return rows


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    ordered = sorted(values)
    require(ordered and 0 < quantile <= 1, "invalid quantile request")
    return ordered[max(1, math.ceil(len(ordered) * quantile)) - 1]


def distribution(values: Iterable[int | float]) -> dict[str, Any]:
    vector = list(values)
    require(vector, "empty timing vector")
    for index, value in enumerate(vector):
        finite(value, f"timing[{index}]", positive=True)
    return {
        "count": len(vector),
        "p50": nearest_rank(vector, .50),
        "p95": nearest_rank(vector, .95),
        "p99": nearest_rank(vector, .99),
        "mean": statistics.fmean(vector),
        "min": min(vector),
        "max": max(vector),
        "values": vector,
    }


def spread(values: Iterable[float]) -> float:
    vector = list(values)
    require(vector and all(value > 0 for value in vector), "invalid spread vector")
    return max(vector) / min(vector)


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(values and all(math.isfinite(value) and value > 0 for value in values),
            "invalid bootstrap vector")
    rng = random.Random(BOOTSTRAP_SEED)
    draws = sorted(statistics.median(rng.choice(values) for _ in values)
                   for _ in range(BOOTSTRAP_RESAMPLES))
    return {
        "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
        "statistic": "median", "low_rank": BOOTSTRAP_LOW_RANK,
        "high_rank": BOOTSTRAP_HIGH_RANK,
        "estimate": statistics.median(values),
        "ci_low": draws[BOOTSTRAP_LOW_RANK],
        "ci_high": draws[BOOTSTRAP_HIGH_RANK],
    }


def native_summary(rows: list[dict[str, Any]], p: dict[str, Any]) -> dict[str, Any]:
    lookup = {(row["block"], row["arm"]): row for row in rows}
    require(len(lookup) == 18, "native selector identity changed")
    arms: dict[str, Any] = {}
    spread_flags: list[dict[str, Any]] = []
    tail_flags: list[dict[str, Any]] = []
    for arm in ARMS:
        block_rows = [lookup[(block, arm)] for block in range(6)]
        metrics: dict[str, Any] = {}
        for metric in ("p50", "p95", "p99", "mean"):
            values = [row["stats"][metric] for row in block_rows]
            ratio = spread(values)
            metrics[metric] = {"values": values, "median": statistics.median(values),
                               "spread_ratio": ratio,
                               "spread_flag": ratio > 1.05}
            if ratio > 1.05:
                spread_flags.append({"arm": arm, "metric": metric,
                                     "spread_ratio": ratio, "descriptive_only": True})
        rss_values = [row["rss_kib"] for row in block_rows]
        rss_ratio = spread(rss_values)
        metrics["rss_kib"] = {"values": rss_values, "median": statistics.median(rss_values),
                               "spread_ratio": rss_ratio,
                               "spread_flag": rss_ratio > 1.05,
                               "source": "maximum resident set size in KiB from /usr/bin/time"}
        if rss_ratio > 1.05:
            spread_flags.append({"arm": arm, "metric": "rss_kib",
                                 "spread_ratio": rss_ratio, "descriptive_only": True})
        p99_p50 = metrics["p99"]["median"] / metrics["p50"]["median"]
        tail_flag = p99_p50 > 1.05
        if tail_flag:
            tail_flags.append({"arm": arm, "p99_to_p50_ratio": p99_p50,
                               "descriptive_only": True})
        arms[arm] = {"blocks": 6, "samples_per_report": p["native"]["samples"],
                     "metrics": metrics, "p99_to_p50_ratio": p99_p50,
                     "tail_flag": tail_flag,
                     "historical_timing_comparison": False,
                     "profile_perturbation_only": True,
                     "reports": [{"block": row["block"], "report": row["report"],
                                  "report_sha256": row["report_sha256"],
                                  "raw_quantiles": row["stats"]} for row in block_rows]}

    pair_rows: dict[str, Any] = {}
    for name, numerator, denominator in (("wrapped/control", "wrapped", "control"),
                                         ("fp/wrapped", "fp", "wrapped")):
        values: list[float] = []
        by_block: list[dict[str, Any]] = []
        for block in range(6):
            before = lookup[(block, denominator)]["stats"]["p50"]
            after = lookup[(block, numerator)]["stats"]["p50"]
            ratio = after / before
            values.append(ratio)
            by_block.append({"block": block, "numerator": numerator,
                             "denominator": denominator, "before": before,
                             "after": after, "ratio": ratio})
        pair_rows[name] = {
            "control": denominator, "numerator": numerator,
            "by_block": by_block, "ratio_values": values,
            "ratio_median": statistics.median(values),
            "bootstrap": bootstrap(values),
            "interpretation": "diagnostic wrapper/build perturbation only; no shipping latency or speedup claim",
        }
    return {
        "reports": len(rows), "samples": sum(len(row["values"]) for row in rows),
        "arms": arms, "paired_ratios": pair_rows,
        "spread_flags": spread_flags, "tail_flags": tail_flags,
        "quantile_definition": "nearest rank within each process; median over six process blocks",
        "raw_report_quantile_definition": "nearest-rank p50/p95/p99 and arithmetic mean",
        "historical_timing_comparison": False, "speedup_claim": False,
        "profile_perturbation_only": True,
    }


def load_qualification(cv: dict[str, Any], build: dict[str, Any], p: dict[str, Any]) -> dict[str, Any]:
    rows = load_lane("qualification", cv, build,
                     expected_reports=3, expected_samples=9,
                     samples=3, warmup=0,
                     orders=p["qualification"]["orders"])
    return {"reports": len(rows), "samples": sum(len(row["values"]) for row in rows),
            "rows": rows, "arms": list(ARMS),
            "purpose": p["qualification"]["purpose"],
            "timing_status": "qualification-only; excluded from native timing summaries"}


def load_native(cv: dict[str, Any], build: dict[str, Any], p: dict[str, Any]) -> dict[str, Any]:
    rows = load_lane("native", cv, build,
                     expected_reports=18, expected_samples=540,
                     samples=30, warmup=3,
                     orders=p["native"]["orders"])
    summary = native_summary(rows, p)
    summary["rows"] = rows
    return summary


def gzip_bytes(value: Any, label: str) -> tuple[bytes, dict[str, Any]]:
    packed_descriptor = value.get("compressed") if isinstance(value, dict) \
        and isinstance(value.get("compressed"), dict) else value
    path = artifact(packed_descriptor, label)
    packed = path.read_bytes()
    try:
        data = gzip.decompress(packed)
    except (OSError, EOFError) as error:
        fail(f"{label}: invalid deterministic gzip: {error}")
    original = value.get("original")
    require(isinstance(original, dict) and is_sha(original.get("sha256"))
            and isinstance(original.get("bytes"), int), f"{label}: original identity missing")
    require(len(data) == original["bytes"]
            and hashlib.sha256(data).hexdigest() == original["sha256"],
            f"{label}: decompressed identity changed")
    require(value.get("compression") == "gzip" and value.get("gzip_mtime") == 0,
            f"{label}: compression contract changed")
    if "decompressed_sha256" in value:
        require(value["decompressed_sha256"] == original["sha256"],
                f"{label}: decompressed digest changed")
    return data, original


def compressed_members(perf_dir: Path) -> dict[tuple[int, str], tuple[bytes, dict[str, Any], dict[str, Any]]]:
    path = perf_dir / "compression.json"
    value = read_json(path)
    require(isinstance(value, list) and value, "perf compression receipt is missing")
    members: dict[tuple[int, str], tuple[bytes, dict[str, Any], dict[str, Any]]] = {}
    for row in value:
        require(isinstance(row, dict) and isinstance(row.get("label"), str),
                "perf compression member malformed")
        match = re.fullmatch(r"([0-9]+)\.(data|raw|frames)", row["label"])
        require(match is not None, "perf compression label changed")
        repeat = int(match.group(1))
        kind = "raw" if match.group(2) in {"data", "raw"} else "frames"
        if "repeat" in row:
            require(row["repeat"] == repeat, "perf compression repeat changed")
        if "kind" in row:
            require(row["kind"] == kind, "perf compression kind changed")
        packed_value = dict(row)
        data, original = gzip_bytes(packed_value,
                                    f"perf {repeat} {kind} compressed")
        original_path = Path(row["original"].get("path", "")) if isinstance(row.get("original"), dict) else Path()
        require(original_path.name == f"{repeat}.data" if kind == "raw"
                else original_path.name == f"{repeat}.frames",
                f"perf {repeat} {kind} original filename changed")
        key = (repeat, kind)
        require(key not in members, f"duplicate perf compression member: {key}")
        members[key] = (data, original, row)
    for repeat in range(2):
        require((repeat, "raw") in members and (repeat, "frames") in members,
                f"perf repeat {repeat} raw/frame members are incomplete")
    return members


HEADER_RE = re.compile(
    r"^\s*(?P<comm>\S+)\s+(?P<pid>\d+)\s+(?:\[\d+\]\s+)?"
    r"(?P<timestamp>\d+(?:\.\d+)?):\s+(?P<period>\d+)\s+cycles:u:\s*$"
)
FRAME_RE = re.compile(
    r"^\s*(?P<address>[0-9A-Fa-f]+)\s+(?P<symbol>.+?)\s+"
    r"\((?P<dso>[^)]*)\)\s*$"
)


def parse_perf_frames(data: bytes, label: str) -> dict[str, Any]:
    text = data.decode("utf-8", errors="replace")
    lines = text.splitlines()
    lost_lines = [line for line in lines
                  if "PERF_RECORD_LOST" in line or
                  ("lost" in line.lower() and "sample" in line.lower())]
    truncation_lines = [line for line in lines if re.search(r"truncat|stack depth|callchain",
                                                              line, re.IGNORECASE)]
    blocks = text.strip().split("\n\n") if text.strip() else []
    samples: list[dict[str, Any]] = []
    status_lines: list[str] = []
    malformed_frames: list[str] = []
    for block_index, block in enumerate(blocks):
        block_lines = block.splitlines()
        if not block_lines:
            continue
        header = HEADER_RE.match(block_lines[0])
        if header is None:
            status_lines.extend(block_lines)
            continue
        period = int(header.group("period"))
        frames: list[dict[str, str]] = []
        for line in block_lines[1:]:
            match = FRAME_RE.match(line)
            if match is None:
                # perf may include a status line in a sample block. Retain it
                # as diagnostic data instead of silently dropping the block.
                malformed_frames.append(line)
                continue
            raw_symbol = match.group("symbol").strip()
            symbol = re.sub(r"\+0x[0-9A-Fa-f]+$", "", raw_symbol)
            frames.append({
                "address": match.group("address"),
                "symbol": symbol,
                "dso": match.group("dso").strip(),
                "raw": line,
            })
        samples.append({
            "index": len(samples),
            "header": block_lines[0],
            "command": header.group("comm"),
            "pid": int(header.group("pid")),
            "timestamp": header.group("timestamp"),
            "period": period,
            "frames": frames,
            "raw_lines": block_lines,
            "block_index": block_index,
            "malformed_frame_count": sum(1 for line in block_lines[1:]
                                          if line.strip() and FRAME_RE.match(line) is None),
            "explicit_truncation": any(re.search(r"truncat|stack depth|callchain", line,
                                                  re.IGNORECASE)
                                        for line in block_lines),
        })
    require(samples, f"{label}: decoded perf stream has no sample headers")
    require(all(sample["period"] > 0 for sample in samples),
            f"{label}: perf period is not positive")
    unknown_frames = [frame for sample in samples for frame in sample["frames"]
                      if frame["symbol"].lower() in {"[unknown]", "unknown", "??", "<unknown>"}
                      or "[unknown]" in frame["symbol"].lower()]
    unknown_sample_count = sum(
        any(frame["symbol"].lower() in {"[unknown]", "unknown", "??", "<unknown>"}
            or "[unknown]" in frame["symbol"].lower()
            for frame in sample["frames"])
        for sample in samples
    )
    return {
        "samples": samples,
        "whole_process_samples": len(samples),
        "whole_process_period": sum(sample["period"] for sample in samples),
        "header_count": len(samples),
        "lost_event_lines": len(lost_lines),
        "lost_event_text": lost_lines,
        "explicit_truncation_lines": truncation_lines,
        "explicit_truncation_count": len(truncation_lines),
        "status_line_count": len(status_lines),
        "status_lines": status_lines,
        "malformed_frame_count": len(malformed_frames),
        "malformed_frame_lines": malformed_frames,
        "unknown_frame_count": len(unknown_frames),
        "unknown_sample_count": unknown_sample_count,
        "raw_text_lines": len(lines),
        "raw_text_bytes": len(data),
    }


def binary_dso_matches(dso: str, binary: dict[str, Any]) -> bool:
    if not dso or dso in {"[unknown]", "unknown"}:
        return False
    expected = Path(binary["path"])
    # The decoder's live-binary descriptor is authoritative.  A basename or
    # suffix would permit a different DSO with the same executable name.
    return dso == str(expected)


def _unknown_symbol(symbol: str) -> bool:
    lowered = symbol.lower()
    return lowered in {"[unknown]", "unknown", "??", "<unknown>"} \
        or "[unknown]" in lowered


def classify_phase_frames(frames: list[dict[str, str]], binary: dict[str, Any]) -> dict[str, Any]:
    """Classify one exact-owner stack using only nested wrapper frames.

    The result is a partition label, never a cost estimate.  A phase is
    accepted only when its exact symbol occurs once in the exact measured DSO
    and no other phase marker (including a marker from another DSO) is present.
    This keeps missing, duplicated, and wrong-DSO markers visible instead of
    silently assigning a phase.
    """
    exact_dso = str(Path(binary["path"]))
    all_hits = []
    exact_hits = []
    other_dso_hits = []
    for index, frame in enumerate(frames):
        for phase, symbol in PHASE_WRAPPERS.items():
            if frame.get("symbol") != symbol:
                continue
            hit = {"phase": phase, "index": index, "dso": frame.get("dso", "")}
            all_hits.append(hit)
            if frame.get("dso") == exact_dso:
                exact_hits.append(hit)
            else:
                other_dso_hits.append(hit)
    by_phase = {phase: [hit for hit in exact_hits if hit["phase"] == phase]
                for phase in PHASE_ORDER}
    if len(all_hits) == 1 and len(exact_hits) == 1:
        status = "phase"
        phase = exact_hits[0]["phase"]
    elif len(all_hits) > 1 or len(exact_hits) > 1:
        status = "ambiguous"
        phase = None
    else:
        status = "unclassified"
        phase = None
    return {
        "status": status,
        "phase": phase,
        "all_hits": all_hits,
        "exact_hits": exact_hits,
        "other_dso_hits": other_dso_hits,
        "by_phase": by_phase,
    }


def _rank_leaf_counter(counter: Counter[tuple[str, str]], periods: Counter[tuple[str, str]]) -> list[dict[str, Any]]:
    """Render a complete (non-overlapping) leaf census for one partition."""
    return [
        {"symbol": symbol, "dso": dso, "samples": count, "period": periods[(symbol, dso)]}
        for (symbol, dso), count in sorted(counter.items(), key=lambda item: (-item[1], item[0]))
    ]


def exact_owner_summary(parsed: dict[str, Any], binary: dict[str, Any],
                        label: str) -> dict[str, Any]:
    owner_samples: list[dict[str, Any]] = []
    symbol_other_dso = 0
    owner_repeated = 0
    unknown_interior = 0
    no_descendant = 0
    leaf_counts: Counter[str] = Counter()
    leaf_period: Counter[str] = Counter()
    inclusive_counts: Counter[str] = Counter()
    inclusive_period: Counter[str] = Counter()
    callpaths: Counter[str] = Counter()
    callpath_period: Counter[str] = Counter()
    phase_counts: Counter[str] = Counter()
    phase_period: Counter[str] = Counter()
    phase_leaf_counts: dict[str, Counter[tuple[str, str]]] = {
        phase: Counter() for phase in PHASE_ORDER
    }
    phase_leaf_period: dict[str, Counter[tuple[str, str]]] = {
        phase: Counter() for phase in PHASE_ORDER
    }
    unclassified_leaf_counts: Counter[tuple[str, str]] = Counter()
    unclassified_leaf_period: Counter[tuple[str, str]] = Counter()
    ambiguous_leaf_counts: Counter[tuple[str, str]] = Counter()
    ambiguous_leaf_period: Counter[tuple[str, str]] = Counter()
    phase_other_dso = 0
    phase_marker_hits = 0
    phase_sample_rows: list[dict[str, Any]] = []
    for sample in parsed["samples"]:
        owner_indexes = [index for index, frame in enumerate(sample["frames"])
                         if frame["symbol"] == OWNER]
        if len(owner_indexes) > 1:
            owner_repeated += 1
        qualified_indexes = [index for index in owner_indexes
                             if binary_dso_matches(sample["frames"][index]["dso"], binary)]
        if owner_indexes and not qualified_indexes:
            symbol_other_dso += 1
        if len(qualified_indexes) != 1:
            continue
        owner_index = qualified_indexes[0]
        descendants = sample["frames"][:owner_index]
        if not descendants:
            no_descendant += 1
        names = [frame["symbol"] for frame in descendants]
        period = sample["period"]
        if any(name.lower() in {"[unknown]", "unknown", "??", "<unknown>"}
               or "[unknown]" in name.lower() for name in names):
            unknown_interior += 1
        leaf = names[0] if names else OWNER
        leaf_counts[leaf] += 1
        leaf_period[leaf] += period
        phase = classify_phase_frames(descendants, binary)
        phase_marker_hits += len(phase["all_hits"])
        phase_other_dso += len(phase["other_dso_hits"])
        phase_status = phase["status"]
        phase_name = phase["phase"]
        if phase_status == "phase":
            phase_counts[phase_name] += 1
            phase_period[phase_name] += period
            phase_leaf_counts[phase_name][(leaf, descendants[0]["dso"] if descendants else sample["frames"][owner_index]["dso"])] += 1
            phase_leaf_period[phase_name][(leaf, descendants[0]["dso"] if descendants else sample["frames"][owner_index]["dso"])] += period
        elif phase_status == "ambiguous":
            ambiguous_leaf_counts[(leaf, descendants[0]["dso"] if descendants else sample["frames"][owner_index]["dso"])] += 1
            ambiguous_leaf_period[(leaf, descendants[0]["dso"] if descendants else sample["frames"][owner_index]["dso"])] += period
        else:
            unclassified_leaf_counts[(leaf, descendants[0]["dso"] if descendants else sample["frames"][owner_index]["dso"])] += 1
            unclassified_leaf_period[(leaf, descendants[0]["dso"] if descendants else sample["frames"][owner_index]["dso"])] += period
        phase_sample_rows.append({
            "sample_index": sample["index"], "period": period,
            "phase_status": phase_status, "phase": phase_name,
            "phase_hits": phase["all_hits"],
            "phase_exact_hits": phase["exact_hits"],
            "phase_other_dso_hits": phase["other_dso_hits"],
        })
        inclusive_names = names + [OWNER]
        for name in inclusive_names:
            inclusive_counts[name] += 1
            inclusive_period[name] += period
        # Each root-to-owner prefix is an inclusive path.  Prefix rows overlap
        # by design and are diagnostic counts, never additive cost estimates.
        for end in range(1, len(inclusive_names) + 1):
            path = " -> ".join(inclusive_names[:end])
            callpaths[path] += 1
            callpath_period[path] += period
        owner_samples.append({
            "sample_index": sample["index"], "period": period,
            "timestamp": sample["timestamp"], "owner_frame_index": owner_index,
            "dso": sample["frames"][owner_index]["dso"],
            "leaf": names[0] if names else OWNER,
            "descendant_count": len(names),
            "phase_status": phase_status, "phase": phase_name,
            "phase_exact_hits": phase["exact_hits"],
            "phase_other_dso_hits": phase["other_dso_hits"],
            "unknown_interior": any(name.lower() in {"[unknown]", "unknown", "??", "<unknown>"}
                                     or "[unknown]" in name.lower() for name in names),
        })
    require(owner_samples, f"{label}: exact owner has zero qualified samples")
    partition_counts = {phase: phase_counts[phase] for phase in PHASE_ORDER}
    partition_counts["unclassified"] = sum(unclassified_leaf_counts.values())
    partition_counts["ambiguous"] = sum(ambiguous_leaf_counts.values())
    require(sum(partition_counts.values()) == len(owner_samples),
            f"{label}: phase partition is not exhaustive")
    partition_periods = {phase: phase_period[phase] for phase in PHASE_ORDER}
    partition_periods["unclassified"] = sum(unclassified_leaf_period.values())
    partition_periods["ambiguous"] = sum(ambiguous_leaf_period.values())
    require(sum(partition_periods.values()) == sum(item["period"] for item in owner_samples),
            f"{label}: phase period partition is not exhaustive")
    return {
        "whole_process_samples": parsed["whole_process_samples"],
        "whole_process_period": parsed["whole_process_period"],
        "owner_qualified_samples": len(owner_samples),
        "owner_qualified_period": sum(item["period"] for item in owner_samples),
        "unattributed_samples": parsed["whole_process_samples"] - len(owner_samples),
        "unattributed_period": parsed["whole_process_period"] -
        sum(item["period"] for item in owner_samples),
        "owner_symbol_other_dso_samples": symbol_other_dso,
        "owner_repeated_samples": owner_repeated,
        "qualified_stacks_with_unknown_interior": unknown_interior,
        "qualified_stacks_without_descendant": no_descendant,
        "phase_partition": {
            "phase_order": list(PHASE_ORDER),
            "phase_owners": dict(PHASE_WRAPPERS),
            "sample_counts": partition_counts,
            "periods": partition_periods,
            "owner_samples": len(owner_samples),
            "classified_samples": sum(phase_counts.values()),
            "unclassified_samples": partition_counts["unclassified"],
            "ambiguous_samples": partition_counts["ambiguous"],
            "phase_marker_hits": phase_marker_hits,
            "phase_marker_other_dso_hits": phase_other_dso,
            "leaf_census": {
                phase: _rank_leaf_counter(phase_leaf_counts[phase], phase_leaf_period[phase])
                for phase in PHASE_ORDER
            },
            "unclassified_leaf_census": _rank_leaf_counter(
                unclassified_leaf_counts, unclassified_leaf_period),
            "ambiguous_leaf_census": _rank_leaf_counter(
                ambiguous_leaf_counts, ambiguous_leaf_period),
            "sample_rows": phase_sample_rows,
            "additive_timing_claim": False,
            "causal_fraction_claim": False,
        },
        "ranked_leaf_within_owner": [
            {"symbol": name, "samples": count, "period": leaf_period[name]}
            for name, count in sorted(leaf_counts.items(), key=lambda item: (-item[1], item[0]))
        ],
        "callpath_inclusive": [
            {"path": name, "samples": count, "period": callpath_period[name]}
            for name, count in sorted(callpaths.items(), key=lambda item: (-item[1], item[0]))
        ],
        "frame_inclusive_through_owner": [
            {"symbol": name, "samples": count, "period": inclusive_period[name]}
            for name, count in sorted(inclusive_counts.items(), key=lambda item: (-item[1], item[0]))
        ],
        "owner_samples": owner_samples,
        "fraction_claim_authorized": False,
        "amdahl_fraction": None,
    }


def descriptor_identity(value: Any, label: str, *, allow_missing: bool = False) -> dict[str, Any]:
    """Return the identity fields of a retained (or archived) descriptor."""
    require(isinstance(value, dict) and isinstance(value.get("path"), str)
            and is_sha(value.get("sha256")), f"{label}: descriptor malformed")
    integer(value.get("bytes"), f"{label}.bytes")
    artifact(value, label, allow_missing=allow_missing)
    return {"path": str(Path(value["path"])), "bytes": value["bytes"],
            "sha256": value["sha256"]}


def same_descriptor_identity(left: Any, right: Any, label: str,
                            *, allow_missing: bool = False) -> None:
    """Compare descriptors even after the uncompressed source was removed."""
    a = descriptor_identity(left, f"{label} left", allow_missing=allow_missing)
    b = descriptor_identity(right, f"{label} right", allow_missing=allow_missing)
    left_path = resolve_path(a["path"], f"{label} left")
    right_path = resolve_path(b["path"], f"{label} right")
    require(left_path == right_path and a["bytes"] == b["bytes"]
            and a["sha256"] == b["sha256"],
            f"{label}: descriptor identity changed")


def validate_perf_report(path: Path, build: dict[str, Any], label: str) -> list[int]:
    report = read_json(path)
    return report_values(report, arm="fp", samples=2000, warmup=0, label=label)


def perf_command(row: dict[str, Any], binary: dict[str, Any], report: dict[str, Any],
                 raw: dict[str, Any], p: dict[str, Any], label: str) -> None:
    expected = [
        "taskset", "-c", str(p["perf"]["cpu"]), "perf", "record",
        "--no-buildid-cache", "-e", p["perf"]["event"], "-F",
        str(p["perf"]["frequency_hz"]), "--call-graph", p["perf"]["call_graph"],
        "-o", raw["path"], "--", binary["path"], "--mode", p["perf"]["mode"],
        "--input", p["perf"]["input_path"], "--reference", p["perf"]["reference_path"],
        "--samples", str(p["perf"]["samples"]), "--warmup", str(p["perf"]["warmup"]),
        "--output", report["path"],
    ]
    require(normalize_command(row.get("command")) == normalize_command(expected),
            f"{label}: perf command boundary changed")


def validate_perf_receipt(row: dict[str, Any], repeat: int, build: dict[str, Any],
                          p: dict[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(row, dict)
            and row.get("schema") == "litchi.performance.0828.perf-receipt.v1"
            and row.get("repeat") == repeat and row.get("exit_code") == 0,
            f"{label}: perf receipt identity changed")
    finite(row.get("started"), f"{label}: start")
    finite(row.get("ended"), f"{label}: end")
    require(row["started"] <= row["ended"], f"{label}: receipt timestamps reversed")
    binary = build["binaries"][p["perf"]["binary"]]["artifact"]
    require(row.get("binary") == binary, f"{label}: perf binary identity changed")
    raw = descriptor_identity(row.get("raw"), f"{label}: raw", allow_missing=True)
    report = descriptor_identity(row.get("report"), f"{label}: report")
    log = descriptor_identity(row.get("log"), f"{label}: log")
    require(Path(report["path"]).name == f"{repeat}.json"
            and Path(raw["path"]).name == f"{repeat}.data"
            and Path(log["path"]).name == f"{repeat}.log",
            f"{label}: perf artifact names changed")
    perf_command(row, binary, report, raw, p, label)
    for key in ("source", "probe"):
        if isinstance(row.get(key), dict):
            artifact(row[key], f"{label}: {key}")
    return {"raw": row["raw"], "report": row["report"], "log": row["log"],
            "report_path": resolve_path(report["path"], f"{label} report"),
            "raw_path": resolve_path(raw["path"], f"{label} raw"),
            "log_path": resolve_path(log["path"], f"{label} log")}


def decode_command(row: dict[str, Any], raw: dict[str, Any], label: str) -> None:
    expected = ["perf", "script", "--no-inline", "--ns", "-i", raw["path"]]
    require(normalize_command(row.get("command")) == normalize_command(expected),
            f"{label}: decoder command changed")


def symbol_evidence(perf_dir: Path, build: dict[str, Any], p: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / "symbols" / "symbol.json"
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0828.symbol-evidence.v1"
            and value.get("owner") == OWNER
            and value.get("owner_function") == "edit_region_0828",
            "perf owner symbol evidence changed")
    require(value.get("frame_pointer_and_call_verified") is True,
            "perf frame-pointer/call proof missing")
    finite(value.get("started"), "perf symbol evidence start")
    finite(value.get("ended"), "perf symbol evidence end")
    require(value["started"] <= value["ended"], "perf symbol evidence timestamps reversed")
    fp = build["binaries"]["fp"]["artifact"]
    require(value.get("binary") == fp and value.get("binary_sha256") == fp["sha256"],
            "perf symbol binary identity changed")
    for key in ("matched_mangled_symbol", "matched_address", "matched_size", "matched_type"):
        require(isinstance(value.get(key), str) and value[key],
                f"perf owner symbol {key} missing")
    lines = value.get("qualified_demangled_lines")
    require(isinstance(lines, list) and len(lines) == 1
            and isinstance(lines[0], str) and OWNER in lines[0],
            "perf exact demangled owner is ambiguous")
    artifacts = {}
    for key in ("nm", "nm_log", "nm_demangled", "nm_demangled_log",
                "phase_nm", "phase_nm_demangled", "assembly", "assembly_log"):
        artifacts[key] = artifact(value.get(key), f"perf symbol {key}")
    phase_assemblies = value.get("phase_assemblies")
    require(isinstance(phase_assemblies, dict)
            and set(phase_assemblies) == set(PHASE_WRAPPERS),
            "perf phase disassembly witness changed")
    for phase in PHASE_ORDER:
        phase_value = phase_assemblies[phase]
        require(isinstance(phase_value, dict)
                and set(phase_value) == {"assembly", "log"},
                f"perf phase {phase} disassembly descriptor changed")
        artifacts[f"{phase}_assembly"] = artifact(
            phase_value["assembly"], f"perf symbol {phase} assembly")
        artifacts[f"{phase}_assembly_log"] = artifact(
            phase_value["log"], f"perf symbol {phase} assembly log")
    nm_text = artifacts["nm"].read_text(encoding="utf-8", errors="replace")
    demangled_text = artifacts["nm_demangled"].read_text(encoding="utf-8", errors="replace")
    assembly_text = artifacts["assembly"].read_text(encoding="utf-8", errors="replace")
    phase_nm_text = artifacts["phase_nm"].read_text(encoding="utf-8", errors="replace")
    phase_demangled_text = artifacts["phase_nm_demangled"].read_text(
        encoding="utf-8", errors="replace")
    require(value.get("phase_owners") == dict(PHASE_WRAPPERS),
            "perf phase-owner symbol witness changed")
    matched_rows = value.get("matched_rows")
    owner_lines = value.get("owner_lines")
    require(isinstance(matched_rows, dict)
            and set(matched_rows) == {"edit_region", *PHASE_ORDER}
            and isinstance(owner_lines, dict)
            and set(owner_lines) == {"edit_region", *PHASE_ORDER},
            "perf exact owner rows changed")
    mangled = value["matched_mangled_symbol"]
    range_tuple = (value["matched_address"], value["matched_size"], value["matched_type"])
    nm_re = re.compile(r"^\s*([0-9A-Fa-f]+)\s+([0-9A-Fa-f]+)\s+(\S)\s+(\S+)\s*$")
    nm_rows = []
    for line in nm_text.splitlines():
        match = nm_re.match(line)
        if match:
            nm_rows.append((match.group(1), match.group(2), match.group(3), match.group(4)))
    require(len(nm_rows) == 1 and nm_rows[0] == (*range_tuple, mangled),
            "retained nm owner range does not match symbol receipt")
    demangled_rows = []
    for line in demangled_text.splitlines():
        match = nm_re.match(line)
        if match:
            demangled_rows.append((match.group(1), match.group(2), match.group(3), match.group(4)))
    require(len(demangled_rows) == 1 and demangled_rows[0][:3] == range_tuple
            and demangled_rows[0][3] == OWNER,
            "retained demangled owner range does not match symbol receipt")
    require(any(line.strip() == lines[0].strip() for line in demangled_text.splitlines())
            and mangled in assembly_text,
            "retained symbol/assembly output does not bind exact owner")
    phase_nm_rows = []
    phase_demangled_rows = []
    for line in phase_nm_text.splitlines():
        match = nm_re.match(line)
        if match:
            phase_nm_rows.append((match.group(1), match.group(2), match.group(3), match.group(4)))
    for line in phase_demangled_text.splitlines():
        match = nm_re.match(line)
        if match:
            phase_demangled_rows.append((match.group(1), match.group(2), match.group(3), match.group(4)))
    require(len(phase_nm_rows) == len(PHASE_ORDER)
            and len(phase_demangled_rows) == len(PHASE_ORDER),
            "retained phase nm output cardinality changed")
    expected_phase_raw = []
    expected_phase_demangled = []
    for phase in PHASE_ORDER:
        raw = matched_rows[phase]
        qualified = owner_lines[phase]
        require(isinstance(raw, dict)
                and tuple(raw.get(key) for key in ("address", "size", "type"))
                and isinstance(raw.get("symbol"), str),
                f"perf phase {phase} raw symbol row changed")
        require(isinstance(qualified, list) and len(qualified) == 1
                and isinstance(qualified[0], str)
                and qualified[0].strip().endswith(PHASE_WRAPPERS[phase]),
                f"perf phase {phase} demangled owner changed")
        expected_phase_raw.append((raw["address"], raw["size"], raw["type"], raw["symbol"]))
        expected_phase_demangled.append(
            (raw["address"], raw["size"], raw["type"], PHASE_WRAPPERS[phase]))
        require(raw["symbol"] in artifacts[f"{phase}_assembly"].read_text(
            encoding="utf-8", errors="replace"),
                f"retained phase disassembly omits exact owner: {phase}")
    require(sorted(phase_nm_rows) == sorted(expected_phase_raw)
            and sorted(phase_demangled_rows) == sorted(expected_phase_demangled),
            "retained phase nm ranges do not match symbol receipt")
    return {
        "schema": value["schema"], "owner": OWNER,
        "owner_function": value["owner_function"],
        "phase_owners": dict(PHASE_WRAPPERS),
        "matched_mangled_symbol": mangled,
        "matched_address": value["matched_address"],
        "matched_size": value["matched_size"], "matched_type": value["matched_type"],
        "binary": fp, "binary_sha256": fp["sha256"],
        "descriptor": {"path": relative(path), "bytes": path.stat().st_size,
                        "sha256": sha256(path)},
        "artifacts": {key: {"path": relative(item), "bytes": item.stat().st_size,
                             "sha256": sha256(item)} for key, item in artifacts.items()},
        "raw": value,
    }


def load_perf(cv: dict[str, Any], build: dict[str, Any], p: dict[str, Any]) -> dict[str, Any]:
    """Replay the terminal perf/decode receipts without invoking profiler tools."""
    perf_dir = PACKET / "perf"
    complete_path = perf_dir / "complete.json"
    complete = read_json(complete_path)
    require(complete.get("schema") == "litchi.performance.0828.perf.complete.v1",
            "perf completion schema changed")
    preflight = complete.get("preflight")
    if isinstance(preflight, dict):
        artifact(preflight, "perf preflight")
    status = complete.get("status")
    require(status in {"available", "unavailable"}, "perf status is not terminal")
    decode_path = perf_dir / "decode-complete.json"
    decode = read_json(decode_path)
    symbol = symbol_evidence(perf_dir, build, p)
    if status == "unavailable":
        require(complete.get("typed_unavailable") is True
                and complete.get("profile_fabricated") is False
                and isinstance(complete.get("reason"), str)
                and complete["reason"], "perf denial is not typed")
        require(decode.get("schema") == "litchi.performance.0828.decode.complete.v1"
                and decode.get("status") == "unavailable"
                and decode.get("typed_unavailable") is True
                and decode.get("profile_fabricated") is False
                and decode.get("symbol") == {
                    "path": str(PACKET / "symbols/symbol.json"),
                    "bytes": (PACKET / "symbols/symbol.json").stat().st_size,
                    "sha256": sha256(PACKET / "symbols/symbol.json"),
                }, "typed perf decode denial changed")
        require(isinstance(decode.get("reason"), str) and decode["reason"],
                "typed perf decode reason missing")
        require(not (perf_dir / "compression.json").exists(),
                "unavailable perf must not fabricate compression")
        return {"status": "unavailable", "reports": 0, "samples": 0,
                "planned_reports": p["perf"]["reports"],
                "planned_samples": p["perf"]["samples_total"],
                "reason": complete["reason"], "decode_reason": decode["reason"],
                "symbol": symbol, "complete": complete, "decode": decode,
                "profiles": [], "profile_fabricated": False}

    require(complete.get("typed_unavailable") is False
            and complete.get("profile_fabricated") is False
            and complete.get("reports") == 2
            and complete.get("samples") == 4000
            and complete.get("event") == "cycles:u"
            and complete.get("frequency_hz") == 997
            and complete.get("call_graph") == "fp"
            and complete.get("owner") == OWNER
            and complete.get("phase_owners") == dict(PHASE_WRAPPERS),
            "perf completion contract changed")
    require(complete.get("plan_sha256") == sha256(PACKET / "plan.json")
            and complete.get("build_sha256") == sha256(PACKET / "build.json"),
            "perf completion custody changed")
    receipts_path = artifact(complete.get("receipts"), "perf receipts")
    receipts = read_json(receipts_path)
    require(isinstance(receipts, list) and len(receipts) == 2,
            "perf receipt cardinality changed")
    binary = build["binaries"]["fp"]["artifact"]
    receipt_rows = []
    for repeat, row in enumerate(receipts):
        receipt_rows.append(validate_perf_receipt(row, repeat, build, p, f"perf/{repeat}"))
        report_path = receipt_rows[-1]["report_path"]
        validate_perf_report(report_path, build, f"perf/{repeat}/report")

    decode_requirements = (
        decode.get("schema") == "litchi.performance.0828.decode.complete.v1"
        and decode.get("status") == "available"
        and decode.get("typed_unavailable") is False
        and decode.get("profile_fabricated") is False
        and decode.get("reports") == 2 and decode.get("frames") == 2
        and decode.get("owner") == OWNER
        and decode.get("phase_owners") == dict(PHASE_WRAPPERS)
        and decode.get("command_contract") == ["perf", "script", "--no-inline", "--ns"]
        and decode.get("raw_and_frames_retained") is True
        and decode.get("retention") ==
        "lossless deterministic gzip; original descriptors bind decompressed bytes"
        and decode.get("uncompressed_copies_removed") is True
        and decode.get("plan_sha256") == sha256(PACKET / "plan.json")
        and decode.get("build_sha256") == sha256(PACKET / "build.json")
    )
    require(decode_requirements, "perf decode completion changed")
    same_descriptor_identity(decode.get("symbol"), symbol["descriptor"],
                             "perf decode symbol")
    decode_receipts_path = artifact(decode.get("receipts"), "perf decode receipts")
    decode_receipts = read_json(decode_receipts_path)
    require(isinstance(decode_receipts, list) and len(decode_receipts) == 2,
            "perf decode receipt cardinality changed")
    compression = compressed_members(perf_dir)
    owner_counts_path = artifact(decode.get("owner_counts"), "perf owner counts")
    owner_counts = read_json(owner_counts_path)
    require(isinstance(owner_counts, list) and len(owner_counts) == 2,
            "perf owner-count receipt cardinality changed")
    profiles = []
    for repeat, decoded in enumerate(decode_receipts):
        label = f"perf/decode/{repeat}"
        require(isinstance(decoded, dict)
                and decoded.get("schema") == "litchi.performance.0828.decode-receipt.v1"
                and decoded.get("repeat") == repeat and decoded.get("exit_code") == 0
                and decoded.get("binary") == binary,
                f"{label}: decode receipt changed")
        same_descriptor_identity(decoded.get("raw"), receipts[repeat]["raw"],
                                f"{label} raw", allow_missing=True)
        same_descriptor_identity(decoded.get("frames"), compression[(repeat, "frames")][1],
                                f"{label} frames", allow_missing=True)
        artifact(decoded.get("log"), f"{label} log")
        decode_command(decoded, receipts[repeat]["raw"], label)
        parsed = parse_perf_frames(compression[(repeat, "frames")][0], label)
        summary = exact_owner_summary(parsed, binary, label)
        summary["decode_diagnostics"] = {
            key: parsed[key] for key in (
                "header_count", "whole_process_samples", "whole_process_period",
                "lost_event_lines", "lost_event_text", "status_line_count", "status_lines",
                "malformed_frame_count", "malformed_frame_lines", "unknown_frame_count",
                "unknown_sample_count", "raw_text_lines", "raw_text_bytes",
            )
        }
        row = owner_counts[repeat]
        require(isinstance(row, dict) and row.get("repeat") == repeat
                and row.get("owner_present") is True
                and isinstance(row.get("owner_text_hits"), int)
                and row["owner_text_hits"] > 0,
                f"{label}: decoder owner witness changed")
        profiles.append({
            "repeat": repeat,
            "receipt": receipts[repeat],
            "decode_receipt": decoded,
            "report": relative(receipt_rows[repeat]["report_path"]),
            "report_sha256": sha256(receipt_rows[repeat]["report_path"]),
            "raw": compression[(repeat, "raw")][2]["compressed"],
            "raw_original": compression[(repeat, "raw")][1],
            "frames": compression[(repeat, "frames")][2]["compressed"],
            "frames_original": compression[(repeat, "frames")][1],
            "summary": summary,
        })
    return {"status": "available", "reports": 2, "samples": 4000,
            "planned_reports": 2, "planned_samples": 4000, "symbol": symbol,
            "complete": complete, "decode": decode, "profiles": profiles,
            "profile_fabricated": False}


def compare_root_audit(native: dict[str, Any]) -> dict[str, Any] | None:
    """Check the optional independent raw census without importing analysis."""
    path = PACKET / "root-audit.json"
    if not path.is_file():
        return None
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0828.root-audit.v1"
            and value.get("unprofiled_reports") == 21
            and value.get("unprofiled_samples") == 549,
            "independent root audit cardinality changed")
    groups = value.get("native_blocks")
    require(isinstance(groups, dict) and set(groups) == set(ARMS),
            "independent root audit arms changed")
    lookup = {(row["block"], row["arm"]): row for row in native["rows"]}
    for arm in ARMS:
        rows = groups[arm]
        require(isinstance(rows, list) and len(rows) == 6,
                f"independent root audit blocks changed: {arm}")
        for row in rows:
            require(isinstance(row, dict) and row.get("block") in range(6),
                    f"independent root audit row changed: {arm}")
            expected = lookup[(row["block"], arm)]
            for key in ("p50", "p95", "p99", "mean", "rss_kib"):
                require(row.get(key) == (expected["stats"][key]
                                         if key != "rss_kib" else expected["rss_kib"]),
                        f"independent root audit disagrees: {arm}/{row['block']}/{key}")
    medians = value.get("native_medians")
    require(isinstance(medians, dict), "independent root audit medians missing")
    for arm in ARMS:
        block_rows = [lookup[(block, arm)] for block in range(6)]
        expected = {metric: statistics.median(
            row["stats"][metric] if metric != "rss_kib" else row["rss_kib"]
            for row in block_rows)
                    for metric in ("p50", "p95", "p99", "mean", "rss_kib")}
        require(medians.get(arm) == expected,
                f"independent root audit median disagrees: {arm}")
    pairs = value.get("paired_p50")
    require(isinstance(pairs, dict), "independent root audit pair results missing")
    for name in ("wrapped/control", "fp/wrapped"):
        expected = native["paired_ratios"][name]
        observed = pairs.get(name)
        require(isinstance(observed, dict)
                and observed.get("block_ratios") == expected["ratio_values"]
                and observed.get("median") == expected["ratio_median"]
                and observed.get("ci95") == [expected["bootstrap"]["ci_low"],
                                               expected["bootstrap"]["ci_high"]],
                f"independent root audit pair disagrees: {name}")
    return {"path": relative(path), "bytes": path.stat().st_size, "sha256": sha256(path),
            "schema": value["schema"]}


def compact_build(build: dict[str, Any]) -> dict[str, Any]:
    return {
        "path": build["path"], "sha256": build["sha256"],
        "source_path": build["source_path"], "frozen_inputs_path": build["frozen_path"],
        "binaries": build["binaries"],
        "cleanup_policy": "target-only; exactly ordinary and fp binaries",
    }


def compact_quality(quality: dict[str, Any]) -> dict[str, Any]:
    return {
        "path": quality["path"], "sha256": quality["sha256"],
        "status": quality["status"], "mode": quality["mode"],
        "production_test_counts": quality["production_test_counts"],
        "probe_gate_count": quality["probe_gate_count"],
        "source": relative(quality["source"]), "frozen": relative(quality["frozen"]),
    }


def receipt_times(rows: list[dict[str, Any]], label: str) -> tuple[float, float]:
    require(rows, f"{label}: no timed receipts")
    previous = float("-inf")
    first = None
    last = None
    for index, row in enumerate(rows):
        finite(row.get("started"), f"{label}/{index} start")
        finite(row.get("ended"), f"{label}/{index} end")
        require(row["started"] <= row["ended"] and previous <= row["started"],
                f"{label}: receipt chronology changed")
        first = row["started"] if first is None else first
        last = row["ended"]
        previous = row["ended"]
    return first, last


def validate_chronology(quality: dict[str, Any], build: dict[str, Any],
                        perf: dict[str, Any], symbol: dict[str, Any],
                        qualification: dict[str, Any], native: dict[str, Any]) -> dict[str, Any]:
    quality_raw = quality["raw"]
    finite(quality_raw.get("started"), "quality start")
    finite(quality_raw.get("ended"), "quality end")
    require(quality_raw["started"] <= quality_raw["ended"], "quality timestamps reversed")
    quality_end = quality_raw["ended"]
    build_rows = build["raw"].get("rows")
    require(isinstance(build_rows, list) and len(build_rows) == 2,
            "build chronology rows missing")
    build_rows = sorted(build_rows, key=lambda row: (row.get("started", 0), row.get("variant", "")))
    build_start, build_end = receipt_times(build_rows, "build")
    require(quality_end <= build_start, "build began before quality completed")
    symbol_raw = symbol["raw"]
    finite(symbol_raw.get("started"), "symbols start")
    finite(symbol_raw.get("ended"), "symbols end")
    require(symbol_raw["started"] <= symbol_raw["ended"]
            and build_end <= symbol_raw["started"],
            "symbols chronology changed")
    qualification_start, qualification_end = receipt_times(
        [row["receipt"] for row in qualification["rows"]], "qualification")
    native_start, native_end = receipt_times(
        [row["receipt"] for row in native["rows"]], "native")
    require(symbol_raw["ended"] <= qualification_start
            and qualification_end <= native_start,
            "qualification/native chronology changed")
    preflight = perf["complete"].get("preflight")
    require(isinstance(preflight, dict), "perf preflight chronology descriptor missing")
    preflight_path = artifact(preflight, "perf preflight chronology")
    preflight_value = read_json(preflight_path)
    finite(preflight_value.get("started"), "perf preflight start")
    finite(preflight_value.get("ended"), "perf preflight end")
    require(preflight_value["started"] <= preflight_value["ended"]
            and native_end <= preflight_value["started"],
            "perf preflight chronology changed")
    if perf["status"] == "available":
        perf_rows = []
        for profile in perf["profiles"]:
            # The retained perf receipt is the authoritative process interval.
            perf_rows.append(read_json(PACKET / "perf" / "receipts.json")[profile["repeat"]])
        perf_start, perf_end = receipt_times(perf_rows, "perf")
        decode_receipts_path = artifact(perf["decode"].get("receipts"),
                                        "perf decode chronology receipts")
        decode_rows = read_json(decode_receipts_path)
        decode_start, decode_end = receipt_times(decode_rows, "perf decode")
        require(preflight_value["ended"] <= perf_start
                and perf_end <= decode_start,
                "perf/decode chronology changed")
        return {"quality_end": quality_end, "build_end": build_end,
                "symbols_end": symbol_raw["ended"], "qualification_end": qualification_end,
                "native_end": native_end, "preflight_end": preflight_value["ended"],
                "perf_end": perf_end, "decode_end": decode_end}
    return {"quality_end": quality_end, "build_end": build_end,
            "symbols_end": symbol_raw["ended"], "qualification_end": qualification_end,
            "native_end": native_end, "preflight_end": preflight_value["ended"]}


def build_analysis() -> dict[str, Any]:
    cv = load_custody()
    p = cv["plan"]
    quality = load_quality(cv)
    build = load_build(cv, quality)
    qualification = load_qualification(cv, build, p)
    native = load_native(cv, build, p)
    perf = load_perf(cv, build, p)
    chronology = validate_chronology(
        quality, build, perf, perf["symbol"], qualification, native)
    observed_reports = qualification["reports"] + native["reports"] + perf["reports"]
    observed_samples = qualification["samples"] + native["samples"] + perf["samples"]
    expected = p["expected"]
    require(qualification["reports"] == expected["qualification_reports"]
            and qualification["samples"] == expected["qualification_samples"]
            and native["reports"] == expected["native_reports"]
            and native["samples"] == expected["native_samples"],
            "native/qualification cardinality changed")
    if perf["status"] == "available":
        require(perf["reports"] == expected["perf_reports"]
                and perf["samples"] == expected["perf_samples"],
                "available perf cardinality changed")
    audit_present = (PACKET / "root-audit.json").is_file()
    verification = {
        "packet_custody_checked": True,
        "pinned_0827_reader_proof_checked": True,
        "production_source_checked": True,
        "probe_source_checked": True,
        "quality_checked": True,
        "build_checked": True,
        "qualification_checked": True,
        "native_checked": True,
        "perf_terminal_checked": True,
        "symbol_live_binary_checked": True,
        "perf_compression_checked": perf["status"] == "available",
        "raw_and_non_inline_frames_retained": perf["status"] == "available",
        "exact_owner_dso_filter_checked": perf["status"] == "available",
        "phase_owner_symbols_checked": perf["status"] == "available",
        "phase_partition_exhaustive": perf["status"] == "available",
        "phase_partition_mutually_exclusive": perf["status"] == "available",
        "phase_leaf_census_retained": perf["status"] == "available",
        "phase_timing_additivity_claim_omitted": True,
        "phase_causal_fraction_claim_omitted": True,
        "zero_owner_failure_guard_checked": perf["status"] == "available",
        "bootstrap_checked": True,
        "nearest_rank_checked": True,
        "paired_comparisons_checked": True,
        "rss_tail_spread_flags_checked": True,
        "historical_timing_comparison_omitted": True,
        "shipping_latency_claim_omitted": True,
        "amdahl_fraction_omitted": True,
        "independent_root_audit_present": audit_present,
    }
    custody_summary = {
        "base": cv["origin"]["base"],
        "target": str(custody.TARGET), "scratch": None,
        "production_file_count": len(cv["source"]["files"]),
        "tool_file_count": len(cv["tool"]),
        "probe_file_count": len(cv["probe"]),
        "architecture_file_count": len(cv["architecture"]),
        "production_changed": False, "runtime_harness_changed": False,
        "tool_changed": False, "tool_allowlist": [],
        "corpus": cv["corpus"], "unrelated": cv["unrelated"],
        "previous_seal": cv["previous_seal"],
        "plan": {"path": "plan.json", "sha256": sha256(PACKET / "plan.json")},
        "origin": {"path": "origin.json", "sha256": sha256(PACKET / "origin.json")},
    }
    return {
        "schema": ANALYSIS_SCHEMA,
        "plan_schema": PLAN_SCHEMA,
        "status": "accepted",
        "scope": p["scope"],
        "base": p["base"],
        "claims": p["statistics"]["claims"],
        "historical_timing_comparison": False,
        "speedup_claim": False,
        "adoption_threshold": None,
        "expected_counts": expected,
        "counts": {
            "reports": observed_reports, "samples": observed_samples,
            "qualification_reports": qualification["reports"],
            "qualification_samples": qualification["samples"],
            "native_reports": native["reports"], "native_samples": native["samples"],
            "perf_reports": perf["reports"], "perf_samples": perf["samples"],
            "planned_reports": expected["total_reports"],
            "planned_samples": expected["total_samples"],
        },
        "timing_status": {
            "native": {"reports": native["reports"], "samples": native["samples"],
                        "status": "descriptive perturbation only"},
            "qualification": {"reports": qualification["reports"],
                               "samples": qualification["samples"],
                               "status": qualification["timing_status"]},
            "perf": {"status": perf["status"], "reports": perf["reports"],
                     "samples": perf["samples"], "planned_reports": perf["planned_reports"],
                     "planned_samples": perf["planned_samples"],
                     "reason": perf.get("reason")},
        },
        "custody": custody_summary,
        "quality": compact_quality(quality),
        "build": compact_build(build),
        "qualification": {
            "reports": qualification["reports"], "samples": qualification["samples"],
            "arms": qualification["arms"], "purpose": qualification["purpose"],
            "rows": qualification["rows"],
        },
        "native": native,
        "perf": perf,
        "chronology": chronology,
        "statistics": p["statistics"],
        "verification": verification,
    }


def csv_text(value: dict[str, Any]) -> str:
    lines = ["arm,block,p50_ns,p95_ns,p99_ns,mean_ns,rss_kib,spread_p50,spread_p95,spread_p99,tail_flag"]
    for arm in ARMS:
        for row in sorted(value["native"]["rows"], key=lambda item: (item["block"], item["arm"])):
            if row["arm"] != arm:
                continue
            stats = row["stats"]
            lines.append(",".join(str(item) for item in (
                arm, row["block"], stats["p50"], stats["p95"], stats["p99"],
                stats["mean"], row["rss_kib"],
                value["native"]["arms"][arm]["metrics"]["p50"]["spread_ratio"],
                value["native"]["arms"][arm]["metrics"]["p95"]["spread_ratio"],
                value["native"]["arms"][arm]["metrics"]["p99"]["spread_ratio"],
                value["native"]["arms"][arm]["tail_flag"],
            )))
    return "\n".join(lines) + "\n"


def render_markdown(value: dict[str, Any]) -> str:
    lines = [
        "# 0828 PPTX edit profile",
        "",
        "Offline replay of the pinned real-file public PPTX edit transaction.",
        "",
        f"Observed reports/samples: {value['counts']['reports']} / {value['counts']['samples']}.",
        f"Planned reports/samples: {value['counts']['planned_reports']} / {value['counts']['planned_samples']}.",
        "Native values are descriptive perturbation evidence; no shipping latency, historical speedup, adoption, or Amdahl fraction is claimed.",
        "",
        "| Arm | p50 median (ns) | p95 median (ns) | p99 median (ns) | mean median (ns) | RSS median (KiB) | tail flag |",
        "|---|---:|---:|---:|---:|---:|---|",
    ]
    for arm in ARMS:
        metrics = value["native"]["arms"][arm]["metrics"]
        lines.append("| " + " | ".join((
            arm, f"{metrics['p50']['median']:.6g}", f"{metrics['p95']['median']:.6g}",
            f"{metrics['p99']['median']:.6g}", f"{metrics['mean']['median']:.6g}",
            f"{metrics['rss_kib']['median']:.6g}",
            "yes" if value["native"]["arms"][arm]["tail_flag"] else "no",
        )) + " |")
    lines.extend(["", "| Paired diagnostic | median ratio | CI95 |", "|---|---:|---|"])
    for name in ("wrapped/control", "fp/wrapped"):
        pair = value["native"]["paired_ratios"][name]
        ci = pair["bootstrap"]
        lines.append(f"| {name} | {pair['ratio_median']:.6g} | [{ci['ci_low']:.6g}, {ci['ci_high']:.6g}] |")
    perf = value["perf"]
    lines.extend(["", f"Perf status: {perf['status']}."])
    if perf["status"] == "unavailable":
        lines.append(f"Perf denial reason: {perf['reason']}.")
    else:
        lines.append("Perf frames retain whole-process periods, headers, lost/status/unknown diagnostics, and exact-owner leaf/callpath summaries.")
    lines.extend(["", "The reader never invokes Cargo, the probe, perf, nm, or objdump.", ""])
    return "\n".join(lines)


def output_text(value: dict[str, Any]) -> dict[str, str]:
    return {
        "analysis.json": json.dumps(value, indent=2, sort_keys=True) + "\n",
        "native.csv": csv_text(value),
        "analysis.md": render_markdown(value),
    }


def analyze(*, write: bool = False, check: bool = False) -> dict[str, Any]:
    value = build_analysis()
    outputs = output_text(value)
    if write:
        for name in outputs:
            require(not (PACKET / name).exists(), f"refusing to overwrite retained {name}")
        for name, text in outputs.items():
            (PACKET / name).write_text(text, encoding="utf-8")
    if check:
        for name, text in outputs.items():
            path = PACKET / name
            require(path.is_file() and path.read_text(encoding="utf-8") == text,
                    f"{name} does not replay deterministically")
    return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    mode = parser.add_mutually_exclusive_group()
    mode.add_argument("--write", action="store_true", help="write derived outputs")
    mode.add_argument("--check", action="store_true", help="replay retained outputs")
    args = parser.parse_args(argv)
    try:
        value = analyze(write=args.write or not args.check, check=args.check)
        print(json.dumps({"status": value["status"],
                          "reports": value["counts"]["reports"],
                          "samples": value["counts"]["samples"]}, sort_keys=True))
    except (ReplayError, AssertionError, OSError, ValueError, KeyError, TypeError, IndexError) as error:
        print(f"0828 analysis failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
