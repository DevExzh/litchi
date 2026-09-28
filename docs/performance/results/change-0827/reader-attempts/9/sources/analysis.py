"""Fail-closed offline replay for the 0827 ordinary-save comparison.

The 0827 packet measures the already committed 0824 PPTX change in the full
ordinary-save workflows.  This reader only consumes retained packet evidence;
it never starts Cargo, a probe, a profiler, a decoder, or a workload.  The
before arm is the two-file 0824 candidate archive and the after arm is the
current HEAD.  Timing and observer evidence stay in separate lanes.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import io
import json
import math
import random
import statistics
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN_SCHEMA = "litchi.performance.0827.plan.v1"
ANALYSIS_SCHEMA = "litchi.performance.0827.ordinary-save-analysis.v1"
CAPTURE_SCHEMA = "litchi.performance.0827.capture-receipt.v1"
BASE = "c5d375083c1624063837c0c57809bd3d371cf72b"
BASE_SHORT = BASE[:12]
TARGET = Path("/home/zhuhe/code/litchi-target-0827")
SCRATCH = Path("/home/zhuhe/code/litchi-fs-0827")
SOURCE_ALLOWLIST = (
    "crates/litchi-pptx/src/opened/transaction.rs",
    "crates/litchi-pptx/src/opened/xml.rs",
)
FORMATS = ("docx", "xlsx", "pptx")
PHASES = ("lifecycle", "edit", "atomic_publish", "counting_publish")
LEGS = ("before", "after")
CASES = tuple(
    {"case": f"{fmt}_real_file_ordinary_save_{phase}", "format": fmt,
     "input": {"docx": "test-data/ooxml/docx/documentProperties.docx",
               "xlsx": "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
               "pptx": "test-data/ooxml/pptx/shapes.pptx"}[fmt],
     "phase": phase}
    for fmt in FORMATS for phase in PHASES
)
BOOTSTRAP_SEED = 827827
BOOTSTRAP_RESAMPLES = 10_000
BOOTSTRAP_LOW_RANK = 250
BOOTSTRAP_HIGH_RANK = 9749
BOOTSTRAP_CONFIDENCE = 0.95
OBSERVER_CONTROLS = 32
REPORT_SCHEMA_VERSION = 1
HEX = frozenset("0123456789abcdefABCDEF")
EDIT_OUTCOME_SHA = hashlib.sha256(b"admitted").hexdigest()
SAVE_ENTRY_POINTS = {
    "DOCX": "litchi_docx::Package::save",
    "XLSX": "litchi_xlsx::Workbook::save",
    "PPTX": "litchi_pptx::Package::save",
}
SINK_ENTRY_POINTS = {
    "DOCX": "litchi_docx::Package::to_stream",
    "XLSX": "litchi_xlsx::Workbook::write_to",
    "PPTX": "litchi_pptx::Package::to_bytes",
}
FULL_ATOMIC_STEPS = (
    "litchi_opc::atomic::replace_with: destination permission probe, sibling temporary creation in the destination's own directory, the publication write, permission preservation, sync_all on the temporary, persist (rename) over the destination, parent-directory sync"
)
ALLOC_METRICS = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes", "net_live",
    "peak_above_entry",
)
PROCESS_METRICS = (
    "user_cpu_ticks", "system_cpu_ticks", "minor_faults", "major_faults",
    "voluntary_context_switches", "nonvoluntary_context_switches",
    "rss_delta_bytes", "peak_rss_bytes", "rchar", "wchar", "read_bytes",
    "write_bytes", "cancelled_write_bytes", "syscr", "syscw",
)
NATIVE_METRICS = ("p50", "p95", "p99", "mean", "rss_kib")
OBSERVER_PAIR_METRICS = ALLOC_METRICS + ("rss_kib",)

# ``custody.tracked_source_names`` is intentionally reproduced as a small
# immutable contract here.  The quality receipt is a full flat census (source,
# tools, packet-adjacent normative files, and unrelated workspace witnesses),
# while build custody is only the production tree used by Cargo.  Treating the
# two maps as the same state would reject every valid build receipt.
PRODUCTION_ROOT_FILES = frozenset({
    ".cargo/config.toml", "Cargo.toml", "clippy.toml", "rust-toolchain.toml",
})


class ReplayError(RuntimeError):
    """Packet evidence is absent, malformed, stale, or contradictory."""


def fail(message: str) -> None:
    raise ReplayError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def canonical_json(value: Any) -> str:
    """Encode JSON-compatible evidence independent of tuple/list spelling."""
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True,
                          separators=(",", ":"), allow_nan=False)
    except (TypeError, ValueError) as error:
        fail(f"evidence is not canonicalizable JSON: {error}")


def accepted_admission_attempt_root(leg: str, attempt_files: Any) -> Path:
    """Resolve the attempt directory named by an accepted admission binding.

    Admission retries are intentionally retained instead of being renamed.
    The accepted receipt's ``fresh_bound.attempt_files`` map is therefore the
    authority for the directory used by replay and chronology.
    """
    require(leg in LEGS, f"invalid admission leg: {leg}")
    require(isinstance(attempt_files, dict) and attempt_files,
            f"admission {leg} attempt binding is missing")
    roots: set[str] = set()
    prefix = f"admission-{leg}"
    for raw in attempt_files:
        require(isinstance(raw, str) and raw, f"admission {leg} attempt path is invalid")
        relative_path = Path(raw)
        require(not relative_path.is_absolute() and ".." not in relative_path.parts
                and relative_path.parts, f"admission {leg} attempt path escaped packet")
        root_name = relative_path.parts[0]
        valid_retry = root_name.startswith(prefix + "-retry") \
            and root_name[len(prefix) + len("-retry"):].isdigit()
        require(root_name == prefix or valid_retry,
                f"admission {leg} attempt root changed: {root_name}")
        roots.add(root_name)
    require(len(roots) == 1, f"admission {leg} has multiple attempt roots")
    return PACKET / next(iter(roots))


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, ValueError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(c in HEX for c in value)


def finite(value: Any, label: str) -> None:
    require(isinstance(value, (int, float)) and not isinstance(value, bool)
            and math.isfinite(float(value)), f"{label} is not finite")


def nonnegative_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= 0,
            f"{label} is not a non-negative integer")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def relative(path: Path) -> str:
    try:
        return str(path.relative_to(PACKET))
    except ValueError:
        return str(path)


def resolve_path(raw: Any, *, packet_bound: bool = False) -> Path:
    require(isinstance(raw, str) and raw, f"invalid artifact path: {raw!r}")
    value = Path(raw)
    candidates = [value] if value.is_absolute() else [PACKET / value, ROOT / value]
    prefix = "docs/performance/results/change-0827/"
    if raw.startswith(prefix):
        candidates.insert(0, PACKET / raw[len(prefix):])
    for candidate in candidates:
        candidate = candidate.resolve(strict=False)
        if packet_bound and not candidate.is_relative_to(PACKET.resolve()):
            continue
        if candidate.is_file() and not candidate.is_symlink():
            return candidate
    candidate = candidates[0].resolve(strict=False)
    if packet_bound:
        require(candidate.is_relative_to(PACKET.resolve()),
                f"artifact path escaped packet: {raw}")
    return candidate


def descriptor(value: Any, label: str, *, packet_bound: bool = False,
               allow_missing: bool = False) -> Path | None:
    require(isinstance(value, dict), f"{label} is not an artifact descriptor")
    path = resolve_path(value.get("path"), packet_bound=packet_bound)
    nonnegative_int(value.get("bytes"), f"{label}.bytes")
    require(is_sha(value.get("sha256")), f"{label}.sha256 is invalid")
    if not path.is_file():
        require(allow_missing, f"missing {label}: {path}")
        return None
    require(not path.is_symlink(), f"{label} is a symlink: {path}")
    require(path.stat().st_size == value["bytes"], f"{label}.bytes changed")
    require(sha256(path) == value["sha256"], f"{label}.sha256 changed")
    return path


def identity(path: Path, *, packet_bound: bool = False) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing artifact: {path}")
    if packet_bound:
        require(path.resolve().is_relative_to(PACKET.resolve()),
                f"artifact escaped packet: {path}")
    return {"path": relative(path), "bytes": path.stat().st_size,
            "sha256": sha256(path)}


def same_descriptor(value: Any, expected: dict[str, Any], label: str,
                    *, packet_bound: bool = True) -> Path:
    path = descriptor(value, label, packet_bound=packet_bound)
    expected_path = resolve_path(expected.get("path"), packet_bound=packet_bound)
    require(path is not None and path.resolve() == expected_path.resolve()
            and value.get("bytes") == expected.get("bytes")
            and value.get("sha256") == expected.get("sha256"),
            f"{label} identity changed")
    return path


def source_map(value: Any, label: str) -> dict[str, str]:
    """Read either a flat source census or the older {revision, files} shape."""
    require(isinstance(value, dict), f"{label} source is malformed")
    files = value.get("files", value)
    require(isinstance(files, dict) and files, f"{label} source files are missing")
    require(all(isinstance(name, str) and is_sha(digest)
                for name, digest in files.items()), f"{label} source digest is invalid")
    return dict(files)


def source_descriptor(value: Any, label: str) -> dict[str, Any]:
    path = descriptor(value, label, packet_bound=True)
    require(path is not None, f"{label} is missing")
    data = read_json(path)
    return {"descriptor": identity(path, packet_bound=True),
            "revision": data.get("revision"), "files": source_map(data, label)}


def plan_value() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(plan.get("schema") == PLAN_SCHEMA, "plan schema changed")
    require(plan.get("base") == BASE, "plan base changed")
    require(plan.get("cpu") == 12 or plan.get("affinity") == [12],
            "plan CPU affinity changed")
    require(tuple(plan.get("source_allowlist", ())) == SOURCE_ALLOWLIST,
            "source allowlist changed")
    require(plan.get("cases") == list(CASES), "case matrix changed")
    totals = plan.get("totals", plan.get("expected", {}))
    expected = {
        "qualification_reports": 24, "qualification_samples": 24,
        "native_reports": 144, "native_samples": 4320,
        "observer_reports": 48, "observer_samples": 144,
        "reports": 216, "samples": 4488,
    }
    require(set(totals) == set(expected)
            and all(totals.get(k) == v for k, v in expected.items()),
            "trial cardinality changed")
    expected_counts = plan.get("expected", {})
    require(set(expected_counts) == {"artifact_cases_per_leg", "artifact_policy_outputs_per_case",
                                    "qualification_reports_per_leg", "qualification_samples_per_leg",
                                    "native_reports", "native_samples", "observer_reports",
                                    "observer_samples", "total_reports", "total_samples"}
            and expected_counts.get("artifact_cases_per_leg") == 6
            and expected_counts.get("artifact_policy_outputs_per_case") == 5
            and expected_counts.get("qualification_reports_per_leg") == 12
            and expected_counts.get("qualification_samples_per_leg") == 12,
            "artifact or qualification cardinality changed")
    bootstrap = plan.get("bootstrap")
    require(isinstance(bootstrap, dict)
            and set(bootstrap) == {"seed", "resamples", "low_rank", "high_rank",
                                   "confidence", "statistic", "sorted_zero_based_endpoints"}
            and bootstrap.get("seed") == BOOTSTRAP_SEED
            and bootstrap.get("resamples") == BOOTSTRAP_RESAMPLES
            and bootstrap.get("low_rank") == BOOTSTRAP_LOW_RANK
            and bootstrap.get("high_rank") == BOOTSTRAP_HIGH_RANK
            and bootstrap.get("confidence") == BOOTSTRAP_CONFIDENCE
            and bootstrap.get("statistic") ==
            "paired after/before process p50 over six counterbalanced native blocks"
            and bootstrap.get("sorted_zero_based_endpoints") ==
            [BOOTSTRAP_LOW_RANK, BOOTSTRAP_HIGH_RANK],
            "bootstrap contract changed")
    lanes = plan.get("lanes", {})
    require(isinstance(lanes, dict), "lane plan missing")
    require(lanes.get("native", {}).get("legs") == list(LEGS)
            and lanes.get("native", {}).get("blocks") == 6
            and lanes.get("native", {}).get("orders") == [
                ["after", "before"], ["before", "after"], ["after", "before"],
                ["before", "after"], ["before", "after"], ["after", "before"]]
            and lanes.get("native", {}).get("samples") == 30
            and lanes.get("native", {}).get("warmup") == 3
            and lanes.get("native", {}).get("reports") == 144
            and lanes.get("native", {}).get("samples_total") == 4320,
            "native lane plan changed")
    require(lanes.get("observer", {}).get("legs") == list(LEGS)
            and lanes.get("observer", {}).get("blocks") == 2
            and lanes.get("observer", {}).get("orders") ==
            [["after", "before"], ["before", "after"]]
            and lanes.get("observer", {}).get("samples") == 3
            and lanes.get("observer", {}).get("warmup") == 0
            and lanes.get("observer", {}).get("reports") == 48
            and lanes.get("observer", {}).get("samples_total") == 144,
            "observer lane plan changed")
    require(lanes.get("qualification", {}).get("legs") == list(LEGS)
            and lanes.get("qualification", {}).get("blocks") == 1
            and lanes.get("qualification", {}).get("orders") ==
            [["before"], ["after"]]
            and lanes.get("qualification", {}).get("samples") == 1
            and lanes.get("qualification", {}).get("warmup") == 0
            and lanes.get("qualification", {}).get("reports") == 24
            and lanes.get("qualification", {}).get("samples_total") == 24,
            "qualification lane plan changed")
    require(plan.get("regression_policy") ==
            "Attribution only: no new adoption or rejection decision.",
            "ordinary-save packet must not encode an adoption decision")
    binaries = plan.get("binaries", {})
    require(isinstance(binaries, dict) and set(binaries) == {"native", "artifacts", "observer"}
            and binaries["native"] == {"cargo_bin": "litchi-perf-baseline",
                                        "features": [], "latency": "timed_native"}
            and binaries["artifacts"] == {"cargo_bin": "ordinary_save_artifacts",
                                           "features": [], "latency": "untimed_artifact_export"}
            and binaries["observer"] == {
                "cargo_bin": "litchi-perf-baseline-alloc",
                "features": ["allocator-metrics", "ordinary-save-process-metrics"],
                "latency": "diagnostic_observer"},
            "binary plan changed")
    require(plan.get("target") == str(TARGET) and plan.get("scratch") == str(SCRATCH),
            "owned target or scratch changed")
    return plan


def artifact_map(path: Path, label: str) -> dict[str, str]:
    return source_map(read_json(path), label)


def is_production_source_name(name: str) -> bool:
    return (name in PRODUCTION_ROOT_FILES or name.startswith("crates/"))


def full_live_source(names: Iterable[str]) -> dict[str, str]:
    result: dict[str, str] = {}
    for name in names:
        require(isinstance(name, str) and name and not Path(name).is_absolute()
                and ".." not in Path(name).parts,
                f"invalid live source census name: {name!r}")
        path = ROOT / name
        require(path.is_file() and not path.is_symlink(), f"live source missing: {name}")
        result[name] = sha256(path)
    require(result, "live source census is empty")
    return result


def live_source() -> dict[str, str]:
    """Hash only the frozen production source set, without invoking Git.

    The quality driver deliberately records a broader flat census.  Its
    production intersection is the same source set that ``custody.source``
    writes into build receipts; tool, docs, and unrelated witnesses remain
    checked by ``validate_quality`` as a separate full census.
    """
    for candidate in (PACKET / "quality-0" / "source.json",
                      PACKET / "quality" / "source.json"):
        if candidate.is_file():
            full = source_map(read_json(candidate), "quality source")
            production = {name: digest for name, digest in full.items()
                          if is_production_source_name(name)}
            require(production, "production source census is empty")
            # Hashing from the live tree also detects a changed or missing
            # production file; do not trust the frozen digest values here.
            return full_live_source(production)
    fail("quality source census names are missing")


def origin_and_inputs(plan: dict[str, Any]) -> dict[str, Any]:
    origin = read_json(PACKET / "origin.json")
    require(isinstance(origin, dict), "origin receipt malformed")
    require(origin.get("schema") == "litchi.performance.0827.origin.v1"
            and origin.get("base") == BASE
            and origin.get("production_changed_at_freeze") is False
            and origin.get("runtime_harness_changed") is False
            and origin.get("tool_changed") is False,
            "origin custody changed")
    require(tuple(origin.get("source_allowlist", ())) == SOURCE_ALLOWLIST,
            "origin allowlist changed")
    require(origin.get("target") == str(TARGET) and origin.get("scratch") == str(SCRATCH),
            "origin target or scratch changed")
    unrelated = origin.get("unrelated")
    if isinstance(unrelated, dict):
        for name, digest in unrelated.items():
            path = ROOT / name
            require(path.is_file() and sha256(path) == digest,
                    f"unrelated file changed: {name}")
    inputs: dict[str, Any] = {"origin": origin}
    root_manifest = PACKET / "root-inputs.json"
    if root_manifest.is_file():
        roots = read_json(root_manifest)
        require(roots.get("schema") == "litchi.performance.0827.root-inputs.v1",
                "root input schema changed")
        for name, item in roots.items():
            if name == "schema":
                continue
            require(isinstance(item, dict) and isinstance(item.get("source_path"), str),
                    f"root input {name} malformed")
            packet_file = descriptor(item, f"root input {name}", packet_bound=True)
            live_file = ROOT / item["source_path"]
            require(packet_file is not None and live_file.is_file()
                    and sha256(live_file) == item["sha256"],
                    f"root input {name} changed")
        inputs["root-inputs.json"] = roots
    architecture_path = PACKET / "architecture-inputs.json"
    if architecture_path.is_file():
        architecture = read_json(architecture_path)
        require(isinstance(architecture, dict) and len(architecture) == 35,
                "architecture input census changed")
        for name, digest in architecture.items():
            require(is_sha(digest) and (ROOT / name).is_file()
                    and sha256(ROOT / name) == digest,
                    f"architecture input changed: {name}")
        inputs["architecture-inputs.json"] = architecture
    lock_path = PACKET / "lock-parity.json"
    if lock_path.is_file():
        locks = read_json(lock_path)
        require(locks.get("schema") == "litchi.performance.0827.lock-parity.v1"
                and locks.get("cargo_lock_changes") ==
                "No lockfile generation or update is permitted during this trial.",
                "lock parity changed")
        for name in ("root_lock", "tool_lock"):
            item = locks.get(name)
            require(isinstance(item, dict) and is_sha(item.get("sha256")),
                    f"lock parity {name} malformed")
            lock_file = descriptor({"path": item["path"], "bytes":
                                     (PACKET / item["path"]).stat().st_size,
                                     "sha256": item["sha256"]},
                                    f"lock parity {name}", packet_bound=True)
            require(lock_file is not None, f"lock parity {name} missing")
        inputs["lock-parity.json"] = locks
    for filename in ("toolchain.json", "host.json", "provenance.json"):
        path = PACKET / filename
        if path.is_file():
            inputs[filename] = read_json(path)
    return inputs


def corpus_inputs(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    """Resolve the frozen real-file corpus identity used by every report."""
    value: Any = plan.get("corpus")
    if not isinstance(value, dict) or not all(isinstance(x, dict) for x in value.values()):
        for name in ("corpus-inputs.json", "corpus.json"):
            path = PACKET / name
            if path.is_file():
                value = read_json(path)
                break
    require(isinstance(value, dict), "corpus identity is missing")
    result = {}
    for case in CASES:
        row = value.get(case["input"])
        if row is None and isinstance(value.get(case["format"]), dict):
            row = value[case["format"]]
        require(isinstance(row, dict) and isinstance(row.get("path", case["input"]), str)
                and is_sha(row.get("sha256")) and isinstance(row.get("bytes"), int),
                f"corpus identity missing: {case['input']}")
        source = ROOT / row.get("path", case["input"])
        require(source.is_file() and source.stat().st_size == row["bytes"]
                and sha256(source) == row["sha256"],
                f"corpus input changed: {case['input']}")
        result[case["input"]] = {"path": row.get("path", case["input"]),
                                  "bytes": row["bytes"], "sha256": row["sha256"]}
    return result


def provenance_input(corpus: dict[str, dict[str, Any]]) -> dict[str, Any] | None:
    path = PACKET / "provenance.json"
    if not path.is_file():
        return None
    value = read_json(path)
    expected_corpus = {
        name: {"bytes": row["bytes"], "sha256": row["sha256"]}
        for name, row in corpus.items()
    }
    require(all(row["path"] == name for name, row in corpus.items()),
            "provenance corpus path changed")
    require(value.get("schema") == "litchi.performance.0827.provenance.v1"
            and value.get("corpus") == expected_corpus, "provenance corpus changed")
    reference = value.get("reference")
    require(isinstance(reference, dict) and isinstance(reference.get("path"), str)
            and is_sha(reference.get("sha256")), "provenance reference changed")
    reference_path = ROOT / reference["path"]
    require(reference_path.is_file() and sha256(reference_path) == reference["sha256"],
            "provenance reference bytes changed")
    return value


def bind_tree_files(root: Path, rows: Any, label: str) -> dict[str, dict[str, Any]]:
    """Bind a retained relative all-files census to the current artifact tree."""
    require(root.is_dir() and not root.is_symlink(), f"{label} root is missing")
    require(isinstance(rows, list) and rows, f"{label} all-files census is missing")
    expected: dict[str, dict[str, Any]] = {}
    for index, row in enumerate(rows):
        require(isinstance(row, dict) and isinstance(row.get("path"), str)
                and not Path(row["path"]).is_absolute()
                and not Path(row["path"]).is_symlink(),
                f"{label}[{index}] path is invalid")
        relative_path = Path(row["path"])
        require(relative_path != Path(".") and ".." not in relative_path.parts,
                f"{label}[{index}] path escapes root")
        require(relative_path.as_posix() not in expected,
                f"{label} contains duplicate file: {relative_path}")
        path = (root / relative_path).resolve(strict=False)
        require(path.is_relative_to(root.resolve()) and path.is_file()
                and not path.is_symlink(), f"{label} file missing: {relative_path}")
        require(isinstance(row.get("bytes"), int) and row["bytes"] >= 0
                and is_sha(row.get("sha256")), f"{label} descriptor malformed")
        actual = identity(path)
        require(actual["bytes"] == row["bytes"] and actual["sha256"] == row["sha256"],
                f"{label} file changed: {relative_path}")
        expected[relative_path.as_posix()] = {
            "path": relative_path.as_posix(), "bytes": row["bytes"],
            "sha256": row["sha256"],
        }
    actual_files: dict[str, dict[str, Any]] = {}
    for path in sorted(root.rglob("*")):
        require(not path.is_symlink(), f"{label} tree contains symlink: {path}")
        if path.is_file():
            relative_path = path.relative_to(root).as_posix()
            actual_files[relative_path] = identity(path)
            actual_files[relative_path]["path"] = relative_path
    require(actual_files == expected, f"{label} all-files census is not exact")
    return expected


def replay_independent_artifacts(leg: str, audit: dict[str, Any],
                                 preservation: dict[str, Any]) -> None:
    """Replay both independent packet oracles without invoking their CLIs."""
    require(isinstance(audit.get("artifact_directory"), str),
            f"artifact audit {leg} root is missing")
    artifact_root = Path(audit["artifact_directory"]).resolve()
    require(artifact_root.is_relative_to(PACKET.resolve()),
            f"artifact audit {leg} root escaped packet")
    # Imports are deliberately local: validate.py remains usable as a single
    # packet reader, while the oracle modules themselves stay standalone.
    import artifact_audit
    import preservation as preservation_oracle

    replayed_audit = artifact_audit.Audit(ROOT).run(artifact_root)
    require(canonical_json(replayed_audit) == canonical_json(audit),
            f"retained XML/ZIP audit {leg} does not replay exactly")
    replayed_preservation = preservation_oracle.analyze(leg)
    require(canonical_json(replayed_preservation) == canonical_json(preservation),
            f"retained ZIP preservation {leg} does not replay exactly")


def validate_admission() -> dict[str, Any] | None:
    """Validate both independent artifact and qualification admissions."""
    result: dict[str, Any] = {}
    for leg in LEGS:
        path = PACKET / f"artifact-admission-{leg}.json"
        require(path.is_file() and not path.is_symlink(),
                f"artifact admission {leg} is missing")
        value = read_json(path)
        require(value.get("schema") == "litchi.performance.0827.artifact-admission.v1"
                and value.get("accepted") is True and value.get("leg") == leg
                and value.get("plan_sha256") == sha256(PACKET / "plan.json"),
                f"artifact admission {leg} changed")
        descriptors: dict[str, Path | None] = {}
        for key in ("artifact_complete", "manifest", "audit", "auditor", "zip_preservation"):
            descriptors[key] = descriptor(value.get(key), f"admission {leg}/{key}",
                                           packet_bound=True)
        require(descriptors["artifact_complete"] is not None
                and descriptors["artifact_complete"].resolve() ==
                (PACKET / f"artifacts-{leg}.complete.json").resolve()
                and descriptors["manifest"] is not None
                and descriptors["audit"] is not None
                and descriptors["zip_preservation"] is not None
                and descriptors["auditor"] is not None
                and descriptors["auditor"].resolve() == (PACKET / "artifact_audit.py").resolve(),
                f"artifact admission {leg} descriptor paths changed")
        audit_path = descriptors["audit"]
        preservation_path = descriptors["zip_preservation"]
        require(audit_path is not None and preservation_path is not None,
                f"admission {leg} oracle descriptors are missing")
        audit = read_json(audit_path)
        require(audit.get("schema") == "litchi.performance.0827.artifact-audit.v1"
                and audit.get("ok") is True and audit.get("errors") == []
                and len(audit.get("cases", [])) == 6
                and isinstance(audit.get("artifact_directory"), str),
                f"independent artifact audit {leg} failed")
        preservation = read_json(preservation_path)
        require(preservation.get("schema") == "litchi.performance.0827.zip-preservation.v1"
                and preservation.get("leg") == leg and preservation.get("ok") is True
                and len(preservation.get("cases", [])) == 6,
                f"independent ZIP preservation {leg} failed")
        artifact_root = Path(audit["artifact_directory"]).resolve()
        require(artifact_root == Path(preservation.get("artifact_directory", "")).resolve()
                and artifact_root.is_relative_to(PACKET.resolve()),
                f"admission {leg} artifact root changed")
        fresh_bound = value.get("fresh_bound")
        require(isinstance(fresh_bound, dict)
                and set(fresh_bound) == {"artifact_files", "attempt_files", "manifest",
                                        "complete", "audit", "preservation"},
                f"admission {leg} fresh binding changed")
        bind_tree_files(artifact_root, fresh_bound["artifact_files"],
                        f"admission {leg}/fresh_bound/artifact_files")
        attempt_files = fresh_bound["attempt_files"]
        require(isinstance(attempt_files, dict) and attempt_files,
                f"admission {leg} attempt file binding missing")
        attempt_root = accepted_admission_attempt_root(leg, attempt_files)
        expected_attempt_files: dict[str, dict[str, Any]] = {}
        for relative_name, item in attempt_files.items():
            require(isinstance(relative_name, str),
                    f"admission {leg} attempt path is invalid")
            expected_path = PACKET / relative_name
            require(expected_path.resolve().is_relative_to(attempt_root.resolve()),
                    f"admission {leg} attempt path escaped accepted root")
            observed = descriptor(item, f"admission {leg}/attempt/{relative_name}",
                                  packet_bound=True)
            require(observed is not None and observed.resolve() == expected_path.resolve(),
                    f"admission {leg} attempt descriptor path changed")
            expected_attempt_files[relative_name] = identity(observed, packet_bound=True)
        require(attempt_root.is_dir() and not attempt_root.is_symlink(),
                f"admission {leg} attempt directory missing")
        require(all(not path.is_symlink() for path in attempt_root.rglob("*")),
                f"admission {leg} attempt tree contains a symlink")
        actual_attempt_files = {
            str(path.relative_to(PACKET)): identity(path, packet_bound=True)
            for path in sorted(attempt_root.rglob("*"))
            if path.is_file() and not path.is_symlink()
        }
        require(actual_attempt_files == expected_attempt_files,
                f"admission {leg} attempt files are not exactly bound")
        for key, expected_path in (("manifest", artifact_root / "manifest.json"),
                                   ("complete", PACKET / f"artifacts-{leg}.complete.json"),
                                   ("audit", audit_path),
                                   ("preservation", preservation_path)):
            observed = descriptor(fresh_bound[key],
                                  f"admission {leg}/fresh_bound/{key}", packet_bound=True)
            require(observed is not None and observed.resolve() == expected_path.resolve(),
                    f"admission {leg}/fresh_bound/{key} path changed")
        # Re-run both independent pure-Python oracles against the retained
        # artifact tree and require byte-for-byte JSON identity with the
        # admission evidence.  No child process or workload is started here.
        replay_independent_artifacts(leg, audit, preservation)
        selectors = value.get("selectors")
        require(isinstance(selectors, list) and len(selectors) == 12,
                f"artifact admission {leg} selectors changed")
        seen_selectors: set[str] = set()
        for index, selector in enumerate(selectors):
            require(isinstance(selector, dict), f"admission {leg} selector {index} malformed")
            for key in ("case", "format", "phase", "input", "edit_outcome"):
                require(isinstance(selector.get(key), str),
                        f"admission {leg} selector {index}/{key} missing")
            require(selector["case"] in {case["case"] for case in CASES}
                    and selector["case"] not in seen_selectors,
                    f"admission {leg} selector identity changed")
            expected_case = next(case for case in CASES if case["case"] == selector["case"])
            require(selector["format"] == expected_case["format"]
                    and selector["phase"] == expected_case["phase"]
                    and selector["input"] == expected_case["input"],
                    f"admission {leg} selector {index} case changed")
            seen_selectors.add(selector["case"])
            require(isinstance(selector.get("source_bytes"), int)
                    and selector["source_bytes"] > 0
                    and is_sha(selector.get("source_sha256"))
                    and isinstance(selector.get("published_bytes"), int)
                    and selector["published_bytes"] > 0
                    and is_sha(selector.get("published_sha256"))
                    and selector["edit_outcome"] == "admitted",
                    f"admission {leg} selector {index} oracle changed")
        require(seen_selectors == {case["case"] for case in CASES},
                f"admission {leg} selector coverage changed")
        checks = value.get("checks")
        require(isinstance(checks, dict) and set(checks) == {
            "fresh_manifest_identity", "fresh_all_files_bound", "independent_xml_zip_audit",
            "independent_zip_preservation", "historical_0819_0821_byte_identity",
            "source_inputs_exact", "five_policy_outputs_per_case", "chosen_edit_closures_exact",
            "qualification_not_started",
        } and all(item is True for item in checks.values()),
                f"artifact admission {leg} checks changed")
        qpath = PACKET / f"qualification-admission-{leg}.json"
        require(qpath.is_file() and not qpath.is_symlink(),
                f"qualification admission {leg} is missing")
        qvalue = read_json(qpath)
        require(qvalue.get("schema") == "litchi.performance.0827.qualification-admission.v1"
                and qvalue.get("accepted") is True and qvalue.get("leg") == leg
                and qvalue.get("plan_sha256") == sha256(PACKET / "plan.json")
                and qvalue.get("reports") == 12 and qvalue.get("samples") == 12,
                f"qualification admission {leg} changed")
        q_descriptors: dict[str, Path | None] = {}
        for key in ("artifact_admission", "qualification_complete", "qualification_receipts"):
            require(key in qvalue, f"qualification admission {leg}/{key} is missing")
            q_descriptors[key] = descriptor(qvalue[key],
                                            f"qualification admission {leg}/{key}",
                                            packet_bound=True)
        require(q_descriptors["artifact_admission"] is not None
                and q_descriptors["artifact_admission"].resolve() == path.resolve(),
                f"qualification admission {leg} artifact gate changed")
        require(q_descriptors["qualification_complete"] is not None
                and q_descriptors["qualification_complete"].resolve() ==
                (PACKET / f"qualification-{leg}/complete.json").resolve(),
                f"qualification admission {leg} completion changed")
        require(q_descriptors["qualification_receipts"] is not None
                and q_descriptors["qualification_receipts"].resolve() ==
                (PACKET / f"qualification-{leg}/receipts.json").resolve(),
                f"qualification admission {leg} receipts changed")
        qchecks = qvalue.get("checks")
        require(isinstance(qchecks, dict) and set(qchecks) == {
            "artifact_gate_consumed", "all_three_real_formats", "all_four_phases_per_format",
            "source_oracles", "publication_oracles", "edit_outcome_oracles",
            "repeated_output_oracles", "timed_sample_cardinality",
        } and all(item is True for item in qchecks.values()),
                f"qualification admission {leg} checks changed")
        qrows = qvalue.get("rows")
        require(isinstance(qrows, list) and len(qrows) == 12,
                f"qualification admission {leg} rows changed")
        for index, row in enumerate(qrows):
            require(isinstance(row, dict) and isinstance(row.get("case"), str)
                    and isinstance(row.get("report"), dict)
                    and isinstance(row.get("receipt"), dict),
                    f"qualification admission {leg} row {index} malformed")
            descriptor(row["report"], f"qualification admission {leg}/row{index}/report",
                       packet_bound=True)
            qreceipt = row["receipt"]
            require(qreceipt.get("schema") == CAPTURE_SCHEMA
                    and qreceipt.get("lane") == "qualification"
                    and qreceipt.get("leg") == leg and qreceipt.get("exit_code") == 0
                    and qreceipt.get("case") == row["case"],
                    f"qualification admission {leg} row {index} receipt changed")
        result[leg] = {"receipt": identity(path, packet_bound=True),
                       "qualification": identity(qpath, packet_bound=True),
                       "audit": identity(audit_path, packet_bound=True),
                       "preservation": identity(preservation_path, packet_bound=True),
                       "artifact_files": len(fresh_bound["artifact_files"]),
                       "attempt_files": len(attempt_files)}
    auditor = PACKET / "artifact_audit.py"
    require(auditor.is_file() and not auditor.is_symlink(),
            "admission auditor source is missing")
    result["auditor"] = identity(auditor, packet_bound=True)
    return result


def validate_quality(plan: dict[str, Any], live: dict[str, str]) -> dict[str, Any]:
    path = PACKET / "quality.json"
    value = read_json(path)
    require(value.get("schema") == "litchi.performance.0827.quality.v1"
            and value.get("status") == "pass" and value.get("gate_count") == 6,
            "quality receipt changed")
    require(value.get("environment", {}).get("CARGO_TARGET_DIR")
            == str(TARGET / "quality"), "quality target changed")
    require(value.get("driver_sha256") == sha256(PACKET / "quality.py"),
            "quality driver identity changed")
    source_desc = descriptor(value.get("source"), "quality source", packet_bound=True)
    require(source_desc is not None, "quality source missing")
    quality_source = source_map(read_json(source_desc), "quality source")
    current_full = full_live_source(quality_source)
    require(current_full == quality_source, "quality source does not match live after tree")
    production = {name: digest for name, digest in quality_source.items()
                  if is_production_source_name(name)}
    require(production == live, "quality production intersection does not match build source")
    rows = value.get("rows")
    require(isinstance(rows, list) and len(rows) == 6, "quality gate count changed")
    commands = []
    manifest = str(ROOT / "tools/perf-baseline/Cargo.toml")
    commands.extend([
        ["cargo", "fmt", "--manifest-path", manifest, "--", "--check"],
        ["cargo", "check", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets"],
        ["cargo", "test", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--", "--test-threads=2"],
        ["cargo", "clippy", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--offline", "--locked", "--manifest-path", manifest,
         "--all-features", "--no-deps"],
        ["python3", "-B", "tools/check_crate_boundaries.py"],
    ])
    logs: set[str] = set()
    for index, (row, command) in enumerate(zip(rows, commands)):
        require(row.get("command") == command and row.get("exit_code") == 0,
                f"quality command {index} changed")
        finite(row.get("started"), f"quality {index} start")
        finite(row.get("ended"), f"quality {index} end")
        require(row["started"] <= row["ended"], f"quality {index} timestamps reversed")
        log = descriptor(row.get("log"), f"quality {index} log", packet_bound=True)
        require(log is not None and str(log) not in logs, f"quality {index} log changed")
        logs.add(str(log))
    reused = value.get("pptx_reuse")
    require(isinstance(reused, dict) and set(reused) == {"before", "after"},
            "PPTX quality reuse witness missing")
    legacy_commands = [
        ["cargo", "fmt", "-p", "litchi-pptx", "--", "--check"],
        ["cargo", "check", "--offline", "--locked", "-p", "litchi-pptx",
         "--all-features", "--all-targets"],
        ["cargo", "test", "--offline", "--locked", "-p", "litchi-pptx",
         "--all-features", "--", "--test-threads=2"],
        ["cargo", "clippy", "--offline", "--locked", "-p", "litchi-pptx",
         "--all-features", "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--offline", "--locked", "-p", "litchi-pptx",
         "--all-features", "--no-deps"],
        ["/usr/bin/python3", "-B", "tools/check_crate_boundaries.py"],
    ]
    reuse_values = {}
    for leg in LEGS:
        item = reused[leg]
        require(isinstance(item, dict), f"PPTX reuse {leg} malformed")
        receipt_path = descriptor(item.get("receipt"), f"PPTX reuse {leg} receipt",
                                  packet_bound=False)
        source_path = descriptor(item.get("source"), f"PPTX reuse {leg} source",
                                 packet_bound=False)
        require(receipt_path is not None and source_path is not None,
                f"PPTX reuse {leg} artifacts missing")
        receipt = read_json(receipt_path)
        require(receipt.get("schema") == "litchi.performance.0824.quality.v1"
                and receipt.get("status") == "pass"
                and len(receipt.get("rows", [])) == 6,
                f"PPTX reused quality {leg} changed")
        old_rows = receipt["rows"]
        require(receipt.get("source") == str(source_path.relative_to(ROOT)),
                f"PPTX reused quality {leg} source binding changed")
        old_quality_driver = ROOT / "docs/performance/results/change-0824/quality.py"
        require(is_sha(receipt.get("driver_sha256")) and old_quality_driver.is_file()
                and receipt["driver_sha256"] == sha256(old_quality_driver),
                f"PPTX reused quality {leg} driver changed")
        old_logs: set[str] = set()
        for index, row in enumerate(old_rows):
            require(row.get("command") == legacy_commands[index]
                    and row.get("exit_code") == 0,
                    f"PPTX reused quality {leg} command {index} changed")
            finite(row.get("started"), f"PPTX reused quality {leg}/{index} start")
            finite(row.get("ended"), f"PPTX reused quality {leg}/{index} end")
            require(row["started"] <= row["ended"],
                    f"PPTX reused quality {leg}/{index} timestamps reversed")
            raw_log = row.get("log")
            require(isinstance(raw_log, str) and raw_log,
                    f"PPTX reused quality {leg}/{index} log path missing")
            log_path = Path(raw_log)
            if not log_path.is_absolute():
                log_path = ROOT / log_path
            log_path = log_path.resolve()
            require(log_path.is_relative_to(ROOT.resolve()) and log_path.is_file()
                    and not log_path.is_symlink()
                    and isinstance(row.get("log_bytes"), int)
                    and row["log_bytes"] >= 0
                    and row["log_bytes"] == log_path.stat().st_size
                    and is_sha(row.get("log_sha256"))
                    and row["log_sha256"] == sha256(log_path)
                    and str(log_path) not in old_logs,
                    f"PPTX reused quality {leg}/{index} log changed")
            old_logs.add(str(log_path))
        reused_source = source_map(read_json(source_path), f"PPTX reuse {leg} source")
        archives = item.get("archives")
        require(isinstance(archives, dict)
                and set(archives) == set(SOURCE_ALLOWLIST),
                f"PPTX reuse {leg} archives missing")
        for name in SOURCE_ALLOWLIST:
            expected = archives.get(name)
            if isinstance(expected, dict):
                archive_path = descriptor(expected, f"PPTX archive {leg}/{name}",
                                          packet_bound=True)
                require(archive_path is not None and sha256(archive_path) == expected["sha256"],
                        f"PPTX archive {leg}/{name} changed")
                expected_digest = expected["sha256"]
            else:
                require(is_sha(expected), f"PPTX archive {leg}/{name} digest missing")
                expected_digest = expected
            candidate = PACKET / "candidate" / leg / Path(name).name
            require(candidate.is_file() and sha256(candidate) == expected_digest,
                    f"PPTX archive {leg}/{name} is not the packet candidate")
            require(reused_source.get(name) == expected_digest,
                    f"PPTX reused source {leg}/{name} is not archive-bound")
        # The historical receipt has a larger source census than the current
        # build proof.  Every legacy entry must remain its retained 0824
        # value, while every overlapping production entry must also agree
        # with the current after quality census.
        legacy_source = source_map(read_json(source_path), f"PPTX reuse {leg} source")
        for name, digest in legacy_source.items():
            if name in SOURCE_ALLOWLIST:
                continue
            require(is_sha(digest), f"PPTX reused source {leg}/{name} digest invalid")
            if name in quality_source:
                require(digest == quality_source[name],
                        f"PPTX reused source {leg}/{name} differs from fresh quality")
        for name, digest in live.items():
            require(name in reused_source,
                    f"PPTX reused source {leg} omits production file {name}")
            if name in SOURCE_ALLOWLIST:
                expected = archives[name]
                expected = expected["sha256"] if isinstance(expected, dict) else expected
                require(reused_source[name] == expected,
                        f"PPTX reused source {leg}/{name} archive differs")
            else:
                require(reused_source[name] == digest,
                        f"PPTX reused source {leg}/{name} differs from live after source")
        reused_production = {name for name in reused_source
                             if is_production_source_name(name)}
        require(reused_production == set(live),
                f"PPTX reused source {leg} production census changed")
        reuse_values[leg] = {"receipt": identity(receipt_path),
                             "source": identity(source_path),
                             "files": reused_source, "archives": archives}
    require(reuse_values["before"]["files"].keys() == reuse_values["after"]["files"].keys(),
            "PPTX reused source file set changed")
    for name in reuse_values["before"]["files"]:
        if name in SOURCE_ALLOWLIST:
            continue
        require(reuse_values["before"]["files"][name] == reuse_values["after"]["files"][name],
                f"PPTX reused non-allowlisted source changed: {name}")
    return {"receipt": identity(path, packet_bound=True),
            "source": identity(source_desc, packet_bound=True),
            "rows": len(rows), "pptx_reuse": reuse_values,
            "environment": value.get("environment")}


def archive_source(leg: str) -> dict[str, str]:
    candidate = PACKET / "candidate" / leg
    if not candidate.is_dir():
        candidate = ROOT / "docs/performance/results/change-0824/candidate" / leg
    result = {}
    for name in SOURCE_ALLOWLIST:
        path = candidate / Path(name).name
        require(path.is_file() and not path.is_symlink(), f"candidate {leg} archive missing: {name}")
        result[name] = sha256(path)
    return result


def validate_source_transitions(before_hashes: dict[str, str],
                                after_hashes: dict[str, str]) -> dict[str, Any]:
    """Check both exact two-file installs and their frozen chronology witnesses."""
    transitions: dict[str, Any] = {}
    expected = {
        "before": {"before": after_hashes, "after": before_hashes},
        "after": {"before": before_hashes, "after": after_hashes},
    }
    for leg in LEGS:
        path = PACKET / f"source-transition-{leg}.json"
        value = read_json(path)
        require(value.get("schema") == "litchi.performance.0827.source-transition.v1"
                and value.get("leg") == leg and value.get("head") == BASE
                and value.get("only_allowlisted_files_changed") is True,
                f"source transition {leg} changed")
        for side in ("before", "after"):
            observed = value.get(side)
            require(isinstance(observed, dict)
                    and observed == expected[leg][side],
                    f"source transition {leg}/{side} changed")
        finite(value.get("started"), f"source transition {leg} start")
        finite(value.get("ended"), f"source transition {leg} end")
        require(value["started"] <= value["ended"],
                f"source transition {leg} timestamps reversed")
        transitions[leg] = identity(path, packet_bound=True)
    return transitions


def validate_source_states(plan: dict[str, Any], live: dict[str, str]) -> dict[str, Any]:
    before_hashes = archive_source("before")
    after_archives = archive_source("after") if (PACKET / "candidate/after").is_dir() else {}
    candidate = plan.get("candidate", {})
    for leg, key in (("before", "before_archives"), ("after", "after_archives")):
        expected_rows = candidate.get(key, {})
        if expected_rows:
            for name in SOURCE_ALLOWLIST:
                row = expected_rows.get(name)
                require(isinstance(row, dict) and is_sha(row.get("sha256"))
                        and row.get("sha256") == (before_hashes if leg == "before"
                                                  else after_archives).get(name),
                        f"candidate {leg} archive descriptor changed: {name}")
    after = dict(live)
    before = dict(after)
    for name, digest in before_hashes.items():
        before[name] = digest
    # The after arm must be the live HEAD, including both changed files.
    require(all(after.get(name) == live.get(name) for name in SOURCE_ALLOWLIST),
            "after source is not the live HEAD")
    if after_archives:
        for name in SOURCE_ALLOWLIST:
            require(after_archives[name] == after[name], f"after archive changed: {name}")
    require(set(before) == set(after), "before/after source file set changed")
    require(before != after and all(before[name] == after[name] for name in before
                                   if name not in SOURCE_ALLOWLIST),
            "source diff is outside the 0824 two-file allowlist")
    transitions = validate_source_transitions(
        before_hashes, {name: after[name] for name in SOURCE_ALLOWLIST})
    return {"before": before, "after": after,
            "changed_files": list(SOURCE_ALLOWLIST),
            "before_archives": before_hashes,
            "after_live_head": {name: after[name] for name in SOURCE_ALLOWLIST},
            "transitions": transitions}


def normalize_binary(row: Any) -> dict[str, Any]:
    if isinstance(row, dict) and isinstance(row.get("artifact"), dict):
        row = row["artifact"]
    require(isinstance(row, dict), "binary descriptor malformed")
    require(isinstance(row.get("path"), str) and is_sha(row.get("sha256")),
            "binary identity malformed")
    nonnegative_int(row.get("bytes"), "binary bytes")
    return {"path": row["path"], "bytes": row["bytes"], "sha256": row["sha256"]}


def load_build(leg: str, states: dict[str, Any]) -> dict[str, Any]:
    path = PACKET / f"build-{leg}" / "build.json"
    require(path.is_file() and not path.is_symlink(), f"{leg} build receipt missing")
    value = read_json(path)
    require(value.get("schema") == f"litchi.performance.0827.build-{leg}.v1"
            and value.get("leg") == leg,
            f"{leg} build schema changed")
    source = value.get("source")
    require(isinstance(source, dict) and "path" in source,
            f"{leg} build source missing")
    source_path = descriptor(source, f"{leg} build source", packet_bound=True)
    require(source_path is not None and source_path.resolve() ==
            (PACKET / f"build-{leg}/source.json").resolve(),
            f"{leg} build source path changed")
    source_value = source_descriptor(source, f"{leg} build source")
    require(source_value["files"] == states, f"{leg} build source is not its arm")
    require(source_value.get("revision") in (None, BASE),
            f"{leg} build source revision changed")
    require(value.get("target") == str(TARGET)
            and value.get("profile") == plan_value()["build"]
            and value.get("environment_contract") == {
                "offline": True, "locked": True, "release": True,
                "jobs": plan_value()["build"]["jobs"], "serial": True,
            }, f"{leg} build contract changed")
    for key in ("frozen_inputs", "quality"):
        require(key in value, f"{leg} build {key} witness missing")
        descriptor(value[key], f"{leg} build {key}", packet_bound=True)
    binaries_value = value.get("binaries", {})
    expected_specs = plan_value()["binaries"]
    require(isinstance(binaries_value, dict)
            and set(binaries_value) == set(expected_specs), f"{leg} binaries changed")
    binaries: dict[str, dict[str, Any]] = {}
    for name, spec in expected_specs.items():
        row = binaries_value[name]
        require(isinstance(row, dict) and row.get("cargo_bin") == spec["cargo_bin"]
                and row.get("features") == spec["features"]
                and set(row) == {"cargo_bin", "features", "artifact"},
                f"{leg} binary {name} contract changed")
        binary = normalize_binary(row)
        expected_path = (TARGET / f"{leg}-{name}").resolve()
        actual_path = Path(binary["path"]).resolve()
        require(actual_path == expected_path, f"{leg} binary path changed: {name}")
        if actual_path.is_file():
            descriptor(row["artifact"], f"{leg} binary {name}", packet_bound=False)
        else:
            require((PACKET / "cleanup.json").is_file(),
                    f"{leg} binary missing without cleanup witness")
        binaries[name] = binary
    rows = value.get("rows", [])
    require(isinstance(rows, list) and len(rows) == 3, f"{leg} build rows missing")
    logs: set[str] = set()
    manifest = str(ROOT / "tools/perf-baseline/Cargo.toml")
    build_config = plan_value()["build"]
    environment = {
        "CARGO_TARGET_DIR": str(TARGET), "CARGO_BUILD_JOBS": str(build_config["jobs"]),
        "CARGO_INCREMENTAL": "0", "CARGO_PROFILE_RELEASE_OPT_LEVEL": str(build_config["opt_level"]),
        "CARGO_PROFILE_RELEASE_DEBUG": str(build_config["debug"]),
        "CARGO_PROFILE_RELEASE_LTO": str(build_config["lto"]),
        "CARGO_PROFILE_RELEASE_CODEGEN_UNITS": str(build_config["codegen_units"]),
        "CARGO_PROFILE_RELEASE_INCREMENTAL": "false",
        "CARGO_PROFILE_RELEASE_PANIC": str(build_config["panic"]),
        "PYTHONDONTWRITEBYTECODE": "1",
    }
    expected_rows = []
    for name, spec in expected_specs.items():
        command = ["cargo", "build", "--offline", "--locked", "--release",
                   "--manifest-path", manifest, "--bin", spec["cargo_bin"]]
        if spec["features"]:
            command.extend(["--features", ",".join(spec["features"])])
        expected_rows.append((name, spec, command))
    for index, row in enumerate(rows):
        require(row.get("exit_code") == 0 and isinstance(row.get("command"), list),
                f"{leg} build row {index} changed")
        name, spec, command = expected_rows[index]
        require(row.get("schema") == "litchi.performance.0827.build-receipt.v1"
                and row.get("leg") == leg and row.get("binary") == name
                and row.get("cargo_bin") == spec["cargo_bin"]
                and row.get("features") == spec["features"]
                and row.get("command") == command
                and row.get("environment") == environment,
                f"{leg} build command {index} changed")
        finite(row.get("started"), f"{leg} build {index} start")
        finite(row.get("ended"), f"{leg} build {index} end")
        require(row["started"] <= row["ended"], f"{leg} build {index} timestamps reversed")
        log = descriptor(row.get("log"), f"{leg} build {index} log", packet_bound=True)
        require(log is not None and str(log) not in logs, f"{leg} build log reused")
        logs.add(str(log))
    return {"receipt": identity(path, packet_bound=True), "source": source_value,
            "binaries": binaries, "raw": value}


def nearest_rank(values: Iterable[float], quantile: float) -> float:
    ordered = sorted(values)
    require(ordered, "empty quantile vector")
    return ordered[max(1, math.ceil(len(ordered) * quantile)) - 1]


def integer_midpoint(left: int, right: int) -> int:
    return (left + right) // 2


def welford_mean(values: Iterable[int | float]) -> float:
    count = 0
    mean = 0.0
    for value in values:
        count += 1
        mean += (float(value) - mean) / count
    require(count > 0, "empty mean vector")
    return mean


def stats(values: Iterable[int]) -> dict[str, Any]:
    vector = list(values)
    require(vector and all(isinstance(x, int) and x > 0 for x in vector),
            "elapsed samples are invalid")
    ordered = sorted(vector)
    return {"count": len(vector), "min": ordered[0], "p50": integer_midpoint(
        ordered[(len(ordered) - 1) // 2], ordered[len(ordered) // 2]),
        "p95": nearest_rank(ordered, .95), "p99": nearest_rank(ordered, .99),
        "max": ordered[-1], "mean": welford_mean(ordered)}


def spread(values: Iterable[float]) -> float:
    vector = list(values)
    require(vector and all(float(x) > 0 for x in vector), "spread vector is invalid")
    return max(vector) / min(vector)


def counter_spread(values: Iterable[float]) -> float | None:
    """Return a multiplicative spread where defined; signed counters use null."""
    vector = list(values)
    require(vector and all(math.isfinite(float(x)) for x in vector),
            "counter spread vector is invalid")
    return spread(vector) if all(float(x) > 0 for x in vector) else None


def bootstrap(values: list[float]) -> dict[str, Any]:
    require(len(values) == 6, "paired bootstrap requires six native blocks")
    rng = random.Random(BOOTSTRAP_SEED)
    estimates = sorted(statistics.median(rng.choice(values) for _ in values)
                       for _ in range(BOOTSTRAP_RESAMPLES))
    return {"estimate": statistics.median(values),
            "ci_low": estimates[BOOTSTRAP_LOW_RANK],
            "ci_high": estimates[BOOTSTRAP_HIGH_RANK],
            "seed": BOOTSTRAP_SEED, "resamples": BOOTSTRAP_RESAMPLES,
            "low_rank": BOOTSTRAP_LOW_RANK, "high_rank": BOOTSTRAP_HIGH_RANK,
            "confidence": BOOTSTRAP_CONFIDENCE,
            "statistic": "median of six paired process ratios"}


def ratio(before: float, after: float) -> dict[str, Any]:
    finite(before, "ratio before")
    finite(after, "ratio after")
    require(before > 0, "ratio denominator is not positive")
    value = after / before
    return {"before": before, "after": after, "ratio": value,
            "change_percent": (value - 1.0) * 100.0,
            "over_5_percent": value > 1.05 or value < 0.95}


def signed_comparison(before: float, after: float) -> dict[str, Any]:
    """Retain signed counters even when a multiplicative ratio is undefined."""
    finite(before, "signed comparison before")
    finite(after, "signed comparison after")
    if before > 0:
        return ratio(before, after)
    return {"before": before, "after": after, "ratio": None,
            "change_percent": None, "over_5_percent": None,
            "ratio_defined": False}


def validate_delta(value: Any, label: str) -> None:
    if value is None:
        return
    require(isinstance(value, dict), f"{label} is malformed")
    for key, item in value.items():
        if key in {"status", "scope", "phase", "timing_scope", "latency_claim",
                   "control_scope", "alignment", "reason"}:
            continue
        if isinstance(item, (int, float)):
            require(item >= 0 and math.isfinite(float(item)), f"{label}.{key} changed")


def validate_metrics(result: dict[str, Any], count: int, lane: str,
                     sample_order: list[int], label: str) -> dict[str, Any]:
    metrics = result.get("operation_metrics")
    require(isinstance(metrics, dict) and metrics.get("sample_count") == count
            and metrics.get("sample_indices") == sample_order
            and metrics.get("alignment") ==
            "elapsed_ns.samples_by_elapsed_then_sample_index",
            f"{label} operation metrics alignment changed")
    allocation, process = metrics.get("allocation"), metrics.get("process")
    if lane == "native":
        require(isinstance(allocation, dict) and allocation.get("status") == "unavailable"
                and isinstance(process, dict) and process.get("status") == "unavailable",
                f"{label} native instrumentation changed")
    else:
        require(isinstance(allocation, dict) and allocation.get("status") == "measured"
                and isinstance(process, dict)
                and process.get("status") in {"measured", "unavailable"},
                f"{label} observer instrumentation changed")
        require(allocation.get("scope") == "operation_global_system_allocator",
                f"{label}.allocation scope changed")
        raw_vectors: dict[str, list[int]] = {}
        raw_names = ALLOC_METRICS[:11]
        for name in raw_names:
            item = allocation.get(name)
            require(isinstance(item, dict) and item.get("status") == "measured"
                    and item.get("scope") == "operation_global_system_allocator"
                    and isinstance(item.get("values"), list)
                    and len(item["values"]) == count,
                    f"{label}.allocation.{name} vector changed")
            vector: list[int] = []
            for index, number in enumerate(item["values"]):
                nonnegative_int(number, f"{label}.allocation.{name}[{index}]")
                vector.append(number)
            raw_vectors[name] = vector
        require(all(number == 0 for number in raw_vectors["failed_allocation_calls"]),
                f"{label}.allocation.failed_allocation_calls is nonzero")
        for index in range(count):
            require(raw_vectors["live_bytes_after"][index]
                    - raw_vectors["live_bytes_before"][index]
                    == raw_vectors["allocated_bytes"][index]
                    - raw_vectors["deallocated_bytes"][index],
                    f"{label}.allocation sample {index} live-byte conservation changed")
            require(raw_vectors["peak_live_bytes_before"][index]
                    >= raw_vectors["live_bytes_before"][index]
                    and raw_vectors["peak_live_bytes_after"][index]
                    >= raw_vectors["peak_live_bytes_before"][index]
                    and raw_vectors["region_peak_live_bytes"][index]
                    >= max(raw_vectors["live_bytes_before"][index],
                           raw_vectors["live_bytes_after"][index])
                    and raw_vectors["region_peak_live_bytes"][index]
                    <= raw_vectors["peak_live_bytes_after"][index],
                    f"{label}.allocation sample {index} peak bounds changed")
        if process.get("status") == "measured":
            for key, item in process.items():
                if not isinstance(item, dict) or "values" not in item:
                    continue
                require(isinstance(item["values"], list)
                        and len(item["values"]) == count,
                        f"{label}.process.{key} vector cardinality changed")
                for index, number in enumerate(item["values"]):
                    finite(number, f"{label}.process.{key}[{index}]")
                    require(float(number).is_integer()
                            and (int(number) >= 0 or key == "rss_delta_bytes"),
                            f"{label}.process.{key}[{index}] is invalid")
    return json.loads(json.dumps(metrics, sort_keys=True))


def read_rss(receipt: dict[str, Any], label: str) -> int:
    path = descriptor(receipt.get("rss"), f"{label} RSS", packet_bound=True)
    require(path is not None, f"{label} RSS missing")
    raw = path.read_text(encoding="utf-8").strip()
    require(raw.isdigit() and int(raw) > 0, f"{label} RSS is invalid")
    return int(raw)


def validate_report(report_path: Path, receipt: dict[str, Any], case: dict[str, Any],
                    lane: str, leg: str, plan: dict[str, Any],
                    build: dict[str, Any], corpus: dict[str, dict[str, Any]]) -> dict[str, Any]:
    report = read_json(report_path)
    label = f"{lane}/{leg}/{case['case']}/block{receipt.get('block')}"
    require(report.get("schema_version") == REPORT_SCHEMA_VERSION,
            f"{label} report schema changed")
    tool = report.get("tool")
    expected_binary = "litchi-perf-baseline" if lane == "native" else "litchi-perf-baseline-alloc"
    require(isinstance(tool, dict) and tool.get("binary") == expected_binary
            and tool.get("name") == "litchi-perf-baseline"
            and tool.get("profile") == "release"
            and tool.get("instrumentation") ==
            ("none" if lane == "native"
             else "ordinary_save_procfs_and_system_allocator_operation_scoped"),
            f"{label} tool identity changed")
    if lane != "native":
        require(tool.get("allocator_counter_revision") == "serialized_region_peak_v3",
                f"{label} allocator revision changed")
    else:
        require("allocator_counter_revision" not in tool,
                f"{label} native allocator revision present")
    identity_value = report.get("binary_identity")
    binary = build["binaries"]["native" if lane == "native" else "observer"]
    require(isinstance(identity_value, dict)
            and identity_value.get("binary_sha256") == binary["sha256"]
            and identity_value.get("binary_bytes") == binary["bytes"]
            and identity_value.get("executable") is True,
            f"{label} binary identity changed")
    config = report.get("configuration")
    lane_plan = plan["lanes"][lane]
    require(isinstance(config, dict)
            and config.get("samples_per_case") == lane_plan["samples"]
            and config.get("warmup_iterations_per_case") == lane_plan["warmup"]
            and config.get("cases") == [case["case"]]
            and config.get("filesystem_fresh_child_per_sample") is True
            and config.get("filesystem_process_isolated") is True,
            f"{label} capture configuration changed")
    result_list = report.get("results")
    require(isinstance(result_list, list) and len(result_list) == 1,
            f"{label} result cardinality changed")
    result = result_list[0]
    elapsed = result.get("elapsed_ns")
    count = lane_plan["samples"]
    require(result.get("case") == case["case"] and isinstance(elapsed, dict)
            and elapsed.get("unit") == "ns", f"{label} selector changed")
    samples = elapsed.get("samples")
    order = elapsed.get("sample_order")
    require(isinstance(samples, list) and len(samples) == count
            and all(isinstance(x, int) and x > 0 for x in samples)
            and samples == sorted(samples)
            and isinstance(order, list) and len(order) == count
            and all(type(x) is int and 0 <= x < count for x in order)
            and sorted(order) == list(range(count)),
            f"{label} elapsed samples changed")
    require(all(samples[index] != samples[index + 1] or order[index] < order[index + 1]
               for index in range(count - 1)),
            f"{label} elapsed tie order changed")
    raw_stats = stats(samples)
    require(elapsed.get("min") == raw_stats["min"]
            and elapsed.get("p50") == raw_stats["p50"]
            and elapsed.get("p95") == raw_stats["p95"]
            and elapsed.get("p99") == raw_stats["p99"]
            and elapsed.get("max") == raw_stats["max"],
            f"{label} raw quantiles changed")
    finite(elapsed.get("mean"), f"{label} raw mean")
    require(abs(float(elapsed["mean"]) - raw_stats["mean"]) < 1e-12,
            f"{label} raw mean changed")
    ordinary = (result.get("source") or {}).get("ordinary_save")
    require(isinstance(ordinary, dict)
            and ordinary.get("format") == case["format"].upper()
            and ordinary.get("origin") == "caller-named-real-file",
            f"{label} ordinary-save evidence changed")
    phase_name = {"lifecycle": "open+edit+save", "edit": "edit",
                  "atomic_publish": "save-to-path",
                  "counting_publish": "serialize-to-counting-sink"}[case["phase"]]
    require(ordinary.get("phase") == phase_name, f"{label} phase changed")
    timing_scope = ordinary.get("timing_scope")
    require(isinstance(timing_scope, str), f"{label} timing scope missing")
    require(ordinary.get("atomic_publication_steps") == FULL_ATOMIC_STEPS,
            f"{label} atomic boundary changed")
    summary = ordinary.get("corpus")
    require(isinstance(summary, dict), f"{label} corpus evidence missing")
    real = summary.get("real_file")
    require(isinstance(real, dict), f"{label} real input evidence missing")
    expected_input = corpus[case["input"]]
    require(real.get("bytes") == expected_input["bytes"]
            and real.get("sha256") == expected_input["sha256"],
            f"{label} real input identity changed")
    admission_path = PACKET / f"artifact-admission-{leg}.json"
    admission = read_json(admission_path)
    selectors = {item.get("case"): item for item in admission.get("selectors", [])}
    selector = selectors.get(case["case"])
    require(isinstance(selector, dict)
            and selector.get("input") == case["input"]
            and selector.get("source_sha256") == expected_input["sha256"]
            and selector.get("source_bytes") == expected_input["bytes"]
            and selector.get("published_sha256") is not None
            and selector.get("published_bytes", 0) > 0
            and selector.get("edit_outcome") == "admitted",
            f"{label} admission selector changed")
    require(summary.get("source_archive_bytes") == real.get("bytes")
            and summary.get("source_archive_sha256") == real.get("sha256")
            and summary.get("save_entry_point") == SAVE_ENTRY_POINTS[case["format"].upper()]
            and summary.get("sink_entry_point") == SINK_ENTRY_POINTS[case["format"].upper()]
            and summary.get("edit_admitted") is True
            and summary.get("edit_outcome") == "admitted",
            f"{label} admitted corpus outcome changed")
    require(summary.get("published_bytes") == selector["published_bytes"]
            and summary.get("published_sha256") == selector["published_sha256"],
            f"{label} published output identity changed")
    outcome_hashes = ordinary.get("edit_outcome_sha256")
    require(isinstance(outcome_hashes, list) and len(outcome_hashes) == count
            and all(x == EDIT_OUTCOME_SHA for x in outcome_hashes),
            f"{label} edit outcome vector changed")
    published = ordinary.get("published_sha256")
    if case["phase"] == "edit":
        require(published == [] and result.get("output_sha256") is None,
                f"{label} edit publication changed")
    else:
        require(isinstance(published, list) and len(published) == count
                and len(set(published)) == 1 and published[0] == selector["published_sha256"]
                and result.get("output_sha256") == published[0],
                f"{label} publication identity changed")
    metrics = validate_metrics(result, count, lane, order, label)
    process_probe = ordinary.get("process_probe")
    if lane == "native":
        require(process_probe is None, f"{label} native process probe present")
    elif process_probe is not None:
        require(isinstance(process_probe, dict)
                and process_probe.get("fixed_count") == OBSERVER_CONTROLS
                and isinstance(process_probe.get("empty_adjacent_snapshot_controls"), list)
                and len(process_probe["empty_adjacent_snapshot_controls"]) == OBSERVER_CONTROLS
                and isinstance(process_probe.get("sample_deltas"), list)
                and len(process_probe["sample_deltas"]) == count,
                f"{label} process probe changed")
        for index, item in enumerate(process_probe["empty_adjacent_snapshot_controls"]):
            validate_delta(item, f"{label}.control[{index}]")
        for index, item in enumerate(process_probe["sample_deltas"]):
            validate_delta(item, f"{label}.sample_delta[{index}]")
    return {"case": tuple(("", "", "")), "leg": leg, "lane": lane,
            "case_value": case["case"], "format": case["format"],
            "phase": case["phase"], "input": case["input"],
            "block": receipt["block"], "order": receipt.get("order"),
            "samples": list(samples), "sample_order": list(order),
            "reported": raw_stats, "stats": raw_stats,
            "rss_kib": read_rss(receipt, label), "operation_metrics": metrics,
            "process_probe": json.loads(json.dumps(process_probe, sort_keys=True))
            if process_probe else None,
            "report": relative(report_path), "report_sha256": sha256(report_path),
            "source_logical_bytes": summary["source_archive_bytes"],
            "published_logical_bytes": summary.get("published_bytes", real["bytes"]),
            "published_sha256": (published[0] if isinstance(published, list) and published else None)}


def order_legs(value: Any) -> list[str]:
    if isinstance(value, list):
        result = [str(x).lower() for x in value]
    elif isinstance(value, str):
        text = value.lower()
        result = {"ba": ["before", "after"], "ab": ["after", "before"],
                  "forward": ["before", "after"], "reverse": ["after", "before"]}.get(text)
        if result is None:
            result = [x.strip() for x in text.split(",")]
    else:
        fail(f"invalid arm order: {value!r}")
    require(result in (list(LEGS), list(reversed(LEGS))), "arm order changed")
    return result


def expected_jobs(plan: dict[str, Any], lane: str, qualification_leg: str | None = None) -> list[dict[str, Any]]:
    section = plan["lanes"][lane]
    if lane == "qualification":
        require(qualification_leg in LEGS, "qualification leg missing")
        orders = [[qualification_leg]]
    else:
        orders = [order_legs(x) for x in section.get("orders", [])]
    jobs = []
    for block, order in enumerate(orders):
        for case in CASES:
            for leg in order:
                jobs.append({"lane": lane, "block": block, "case": case,
                             "leg": leg, "samples": section["samples"],
                             "warmup": section["warmup"],
                             "order": "BA" if order == ["before", "after"] else "AB"})
    return jobs


def receipt_case(row: dict[str, Any]) -> dict[str, Any]:
    value = row.get("case")
    if isinstance(value, dict):
        name = value.get("case")
    else:
        name = value
    require(isinstance(name, str), "capture case identity missing")
    matches = [case for case in CASES if case["case"] == name]
    require(len(matches) == 1, f"unknown capture case: {name}")
    return matches[0]


def receipt_list(directory: Path, label: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    complete = read_json(directory / "complete.json")
    descriptor_value = complete.get("receipts")
    if isinstance(descriptor_value, dict):
        receipts_path = descriptor(descriptor_value, f"{label} receipts", packet_bound=True)
        require(receipts_path is not None, f"{label} receipts missing")
        rows = read_json(receipts_path)
    else:
        receipts_path = directory / "receipts.json"
        rows = read_json(receipts_path)
    require(isinstance(rows, list), f"{label} receipts malformed")
    return complete, rows


def load_lane(plan: dict[str, Any], states: dict[str, Any], builds: dict[str, Any],
              lane: str, *, qualification_leg: str | None = None,
              directory_name: str | None = None,
              corpus: dict[str, dict[str, Any]] | None = None) -> list[dict[str, Any]]:
    directory = PACKET / (directory_name or lane)
    complete, receipts = receipt_list(directory, lane)
    jobs = expected_jobs(plan, lane, qualification_leg)
    require(len(receipts) == len(jobs), f"{lane} receipt cardinality changed")
    expected_complete_schema = (f"litchi.performance.0827.qualification-{qualification_leg}.complete.v1"
                                if lane == "qualification" else
                                f"litchi.performance.0827.{lane}.complete.v1")
    require(complete.get("schema") == expected_complete_schema
            and complete.get("status") == "pass"
            and complete.get("lane") == lane
            and complete.get("blocks") == plan["lanes"][lane]["blocks"]
            and complete.get("reports") == len(jobs)
            and complete.get("samples") == sum(x["samples"] for x in jobs)
            and complete.get("expected_reports") == plan["expected"][
                "qualification_reports_per_leg" if lane == "qualification"
                else f"{lane}_reports"]
            and complete.get("expected_samples") == plan["expected"][
                "qualification_samples_per_leg" if lane == "qualification"
                else f"{lane}_samples"],
            f"{lane} completion changed")
    if lane == "qualification":
        require(complete.get("leg") == qualification_leg, f"{lane} completion leg changed")
    else:
        require(complete.get("leg") is None, f"{lane} completion leg changed")
    if "plan_sha256" in complete:
        require(complete["plan_sha256"] == sha256(PACKET / "plan.json"),
                f"{lane} plan custody changed")
    require(isinstance(complete.get("plan"), dict), f"{lane} completion plan missing")
    plan_path = descriptor(complete["plan"], f"{lane} completion plan", packet_bound=True)
    require(plan_path is not None and plan_path.resolve() == (PACKET / "plan.json").resolve(),
            f"{lane} completion plan changed")
    if lane == "qualification":
        build_path = descriptor(complete.get("build"), f"{lane} completion build",
                                packet_bound=True)
        require(build_path is not None, f"{lane} completion build missing")
        require(build_path is not None and build_path.resolve() ==
                (PACKET / f"build-{qualification_leg}/build.json").resolve(),
                f"{lane} completion build path changed")
    else:
        require(complete.get("build") is None, f"{lane} completion build unexpectedly present")
        require(isinstance(complete.get("freeze"), dict),
                f"{lane} completion freeze missing")
        freeze_path = descriptor(complete["freeze"], f"{lane} completion freeze",
                                 packet_bound=True)
        require(freeze_path is not None and freeze_path.resolve() ==
                (PACKET / "freeze.json").resolve(), f"{lane} completion freeze changed")
    source_path = descriptor(complete.get("source"), f"{lane} completion source",
                             packet_bound=True)
    require(source_path is not None and source_map(read_json(source_path),
                                                    f"{lane} completion source") ==
            (states[qualification_leg] if lane == "qualification" else states["after"]),
            f"{lane} completion source changed")
    receipts_path = descriptor(complete.get("receipts"), f"{lane} completion receipts",
                               packet_bound=True)
    require(receipts_path is not None and receipts_path.resolve() ==
            (directory / "receipts.json").resolve(), f"{lane} completion receipts changed")
    result = []
    seen: set[tuple[Any, ...]] = set()
    for row, job in zip(receipts, jobs):
        label = f"{lane}/{job['case']['case']}/block{job['block']}/{job['leg']}"
        require(row.get("schema") == CAPTURE_SCHEMA and row.get("lane") == lane
                and row.get("exit_code") == 0 and row.get("block") == job["block"]
                and row.get("leg") == job["leg"] and row.get("samples") == job["samples"]
                and row.get("warmup") == job["warmup"]
                and row.get("case") == job["case"]["case"]
                and row.get("format") == job["case"]["format"]
                and row.get("phase") == job["case"]["phase"]
                and row.get("input") == job["case"]["input"], f"{label} receipt changed")
        case = receipt_case(row)
        require(case == job["case"], f"{label} case order changed")
        key = (job["block"], job["case"]["case"], job["leg"])
        require(key not in seen, f"{label} duplicate process identity")
        seen.add(key)
        finite(row.get("started"), f"{label} start")
        finite(row.get("ended"), f"{label} end")
        require(row["started"] <= row["ended"], f"{label} timestamps reversed")
        report_path = descriptor(row.get("report"), f"{label} report", packet_bound=True)
        log_path = descriptor(row.get("log"), f"{label} log", packet_bound=True)
        rss_path = descriptor(row.get("rss"), f"{label} RSS", packet_bound=True)
        require(report_path is not None and log_path is not None and rss_path is not None,
                f"{label} artifacts missing")
        frozen_path = descriptor(row.get("frozen_inputs"), f"{label} frozen inputs",
                                 packet_bound=True)
        require(frozen_path is not None and frozen_path.resolve() ==
                (PACKET / "freeze.json").resolve(), f"{label} frozen inputs changed")
        source_path = descriptor(row.get("source"), f"{label} source", packet_bound=True)
        expected_receipt_source = states[job["leg"]] if lane == "qualification" else states["after"]
        require(source_path is not None
                and source_map(read_json(source_path), f"{label} source") == expected_receipt_source,
                f"{label} source custody changed")
        if lane == "qualification":
            require(row.get("build_source") is None,
                    f"{label} qualification build source unexpectedly present")
            admission_path = descriptor(row.get("artifact_admission"),
                                        f"{label} artifact admission", packet_bound=True)
            require(admission_path is not None
                    and admission_path.resolve() ==
                    (PACKET / f"artifact-admission-{job['leg']}.json").resolve()
                    and read_json(admission_path).get("accepted") is True
                    and read_json(admission_path).get("leg") == job["leg"],
                    f"{label} artifact admission changed")
        else:
            build_source_path = descriptor(row.get("build_source"), f"{label} build source",
                                          packet_bound=True)
            require(build_source_path is not None
                    and source_map(read_json(build_source_path), f"{label} build source")
                    == states[job["leg"]], f"{label} build source custody changed")
            require(row.get("artifact_admission") is None,
                    f"{label} comparative artifact admission unexpectedly present")
        binary_name = "native" if lane == "native" else "observer"
        binary = normalize_binary(row.get("binary"))
        require(binary == builds[job["leg"]]["binaries"][binary_name],
                f"{label} binary changed")
        expected_command = [
            "/usr/bin/time", "-f", "%M", "-o", str(rss_path),
            "taskset", "-c", "12", binary["path"],
            "--warmup", str(job["warmup"]), "--samples", str(job["samples"]),
            "--case", case["case"], "--json", str(report_path),
            "--filesystem-root", str(SCRATCH), "--ooxml-file",
            str(ROOT / case["input"]),
        ]
        require(row.get("command") == expected_command,
                f"{label} capture command changed")
        parsed = validate_report(report_path, row, case, lane, job["leg"], plan,
                                 builds[job["leg"]], corpus or {})
        parsed["case"] = (case["format"], case["phase"], case["case"])
        parsed["log"] = relative(log_path)
        result.append(parsed)
    require(len(seen) == len(jobs), f"{lane} process identities incomplete")
    return result


def grouped(entries: list[dict[str, Any]], lane: str) -> dict[tuple[str, str], dict[str, Any]]:
    groups: dict[tuple[str, str], list[dict[str, Any]]] = {}
    for entry in entries:
        groups.setdefault((entry["case"][2], entry["leg"]), []).append(entry)
    result = {}
    for (case, leg), values in sorted(groups.items()):
        values.sort(key=lambda x: x["block"])
        process = {"p50": [x["stats"]["p50"] for x in values],
                   "p95": [x["stats"]["p95"] for x in values],
                   "p99": [x["stats"]["p99"] for x in values],
                   "mean": [x["stats"]["mean"] for x in values]}
        row = {
            "case": case, "leg": leg, "blocks": len(values),
            "processes": [{"block": x["block"], "order": x["order"],
                           "report": x["report"], "report_sha256": x["report_sha256"],
                           "stats": x["stats"], "rss_kib": x["rss_kib"]} for x in values],
            "process_median": {name: statistics.median(vector)
                               for name, vector in process.items()},
            "rss_process_median": statistics.median(x["rss_kib"] for x in values),
            "spread_ratios": {name: spread(vector) for name, vector in process.items()},
            "rss_spread_ratio": spread(x["rss_kib"] for x in values),
            "spread_flags": [name for name, vector in process.items()
                             if spread(vector) > 1.05],
            "tail_ratio": statistics.median(process["p99"]) /
            statistics.median(process["p50"]),
        }
        row["tail_flag"] = row["tail_ratio"] > 1.05
        if lane == "observer":
            row["allocation"] = observer_allocation(values)
            row["process_diagnostics"] = observer_process(values)
        result[(case, leg)] = row
    return result


def metric_values(metrics: dict[str, Any], group: str, name: str) -> list[float] | None:
    value = metrics.get(group, {}).get(name)
    if not isinstance(value, dict) or not isinstance(value.get("values"), list):
        return None
    return [float(x) for x in value["values"]]


def observer_allocation(values: list[dict[str, Any]]) -> dict[str, Any]:
    per_process = []
    for entry in values:
        alloc = {}
        for name in ALLOC_METRICS[:11]:
            vector = metric_values(entry["operation_metrics"], "allocation", name)
            require(vector is not None and len(vector) == len(entry["samples"]),
                    f"observer allocation metric missing: {name}")
            alloc[name] = vector
        alloc["net_live"] = [a - b for a, b in zip(alloc["live_bytes_after"],
                                                   alloc["live_bytes_before"])]
        alloc["peak_above_entry"] = [a - b for a, b in zip(alloc["region_peak_live_bytes"],
                                                            alloc["live_bytes_before"])]
        for name, vector in alloc.items():
            if name == "net_live":
                require(all(math.isfinite(x) for x in vector),
                        f"observer allocation {name} is invalid")
            else:
                require(all(x >= 0 for x in vector),
                        f"observer allocation {name} is negative")
        per_process.append({"block": entry["block"], "values": alloc,
                            "medians": {name: statistics.median(vector)
                                        for name, vector in alloc.items()}})
    medians = {name: [x["medians"][name] for x in per_process]
               for name in ALLOC_METRICS}
    return {"processes": per_process, "process_medians": medians,
            "median_over_blocks": {name: statistics.median(vector)
                                    for name, vector in medians.items()},
            "spread_ratios": {name: counter_spread(vector) for name, vector in medians.items()},
            "spread_flags": [name for name, vector in medians.items()
                             if (counter_spread(vector) is not None
                                 and counter_spread(vector) > 1.05)]}


def observer_process(values: list[dict[str, Any]]) -> dict[str, Any]:
    result = {}
    for name in PROCESS_METRICS:
        vectors = [metric_values(x["operation_metrics"], "process", name) for x in values]
        if all(vector is not None for vector in vectors):
            result[name] = {"processes": [{"block": x["block"], "values": vector}
                                           for x, vector in zip(values, vectors)],
                            "median_over_blocks": statistics.median(
                                statistics.median(vector) for vector in vectors if vector is not None)}
    return {"available": bool(result), "metrics": result,
            "claim": "diagnostic process counters; no observer latency claim"}


def paired_native(entries: list[dict[str, Any]]) -> dict[str, Any]:
    lookup = {(x["case"], x["block"], x["leg"]): x for x in entries}
    result = {}
    for case_definition in CASES:
        case_name = (case_definition["format"], case_definition["phase"],
                     case_definition["case"])
        blocks = sorted({x["block"] for x in entries if x["case"] == case_name})
        require(len(blocks) == 6, f"native block coverage changed: {case_name[2]}")
        metrics = {}
        for name in NATIVE_METRICS:
            ratios = []
            by_block = []
            for block in blocks:
                before = lookup[(case_name, block, "before")]
                after = lookup[(case_name, block, "after")]
                left = before["rss_kib"] if name == "rss_kib" else before["stats"][name]
                right = after["rss_kib"] if name == "rss_kib" else after["stats"][name]
                row = {"block": block, **ratio(left, right)}
                by_block.append(row)
                ratios.append(row["ratio"])
            metrics[name] = {"by_block": by_block, "ratios": ratios,
                             "median_ratio": statistics.median(ratios),
                             "bootstrap": bootstrap(ratios),
                             "max_min_over_5_percent": any(x["over_5_percent"] for x in by_block)}
        result[case_name[2]] = {"case": list(case_name), "blocks": blocks,
                                "metrics": metrics,
                                "comparison": "after/before paired by native block"}
    return result


def paired_observer(entries: list[dict[str, Any]]) -> dict[str, Any]:
    lookup = {(x["case"], x["block"], x["leg"]): x for x in entries}
    result = {}
    for case in sorted({x["case"] for x in entries if x["leg"] == "before"}):
        blocks = sorted({x["block"] for x in entries if x["case"] == case})
        require(len(blocks) == 2, f"observer block coverage changed: {case}")
        rows = {}
        for name in OBSERVER_PAIR_METRICS:
            by_block = []
            ratios = []
            raw_before: list[Any] = []
            raw_after: list[Any] = []
            for block in blocks:
                before = lookup[(case, block, "before")]
                after = lookup[(case, block, "after")]
                if name == "rss_kib":
                    left = before["rss_kib"]
                    right = after["rss_kib"]
                    raw_before.append([before["rss_kib"]])
                    raw_after.append([after["rss_kib"]])
                else:
                    before_values = observer_allocation([before])["processes"][0]["values"][name]
                    after_values = observer_allocation([after])["processes"][0]["values"][name]
                    raw_before.append(before_values)
                    raw_after.append(after_values)
                    left = statistics.median(
                        before_values)
                    right = statistics.median(
                        after_values)
                item = {"block": block, **signed_comparison(left, right)}
                by_block.append(item)
                ratios.append(item["ratio"])
            defined_ratios = [item for item in ratios if item is not None]
            rows[name] = {"by_block": by_block,
                          "median_ratio": (statistics.median(defined_ratios)
                                            if defined_ratios else None),
                          "process_medians": {"before": [x["before"] for x in by_block],
                                               "after": [x["after"] for x in by_block]},
                          "all_values": {"before": raw_before, "after": raw_after},
                          "increase": statistics.median([x["after"] for x in by_block])
                          > statistics.median([x["before"] for x in by_block])}
        result[case[2]] = {"case": list(case), "blocks": blocks, "metrics": rows,
                        "comparison": "after/before paired by observer block"}
    return result


def decision_flags(native_pairs: dict[str, Any], observer_pairs: dict[str, Any]) -> dict[str, Any]:
    regressions = []
    rss_reviews = []
    allocation_increases = []
    for case, group in native_pairs.items():
        p50 = group["metrics"]["p50"]
        if p50["bootstrap"]["ci_low"] > 1.05:
            regressions.append({"case": case, "metric": "p50", "bootstrap": p50["bootstrap"]})
        rss = group["metrics"]["rss_kib"]
        if rss["median_ratio"] > 1.05:
            rss_reviews.append({"case": case, "lane": "native", "metric": "rss_kib",
                                "median_ratio": rss["median_ratio"], "review_only": True})
    for case, group in observer_pairs.items():
        for name in ("allocation_calls", "allocated_bytes", "net_live", "peak_above_entry"):
            if group["metrics"][name]["increase"]:
                allocation_increases.append({"case": case, "metric": name,
                                             "review_only": True})
        # Observer RSS is diagnostic; use the paired two-block median ratio,
        # exactly as for native RSS, rather than flagging one block in isolation.
        rss = group["metrics"].get("rss_kib")
        if isinstance(rss, dict) and isinstance(rss.get("median_ratio"), (int, float)) \
                and rss["median_ratio"] > 1.05:
            rss_reviews.append({"case": case, "lane": "observer", "metric": "rss_kib",
                                "median_ratio": rss["median_ratio"], "review_only": True})
    return {"p50_regression_flags": regressions,
            "allocation_increase_flags": allocation_increases,
            "rss_review_flags": rss_reviews,
            "adoption_decision": None,
            "claims_excluded": ["adoption", "general benefit", "unsupported general full-save claim",
                                "additive phase medians", "pooled native/observer latency",
                                "cross-format benefit", "historical speedup"]}


def chronology() -> dict[str, Any]:
    """Require the frozen quality, transitions, builds, gates, and lanes to serialize."""

    def bounds(path: Path, label: str) -> list[tuple[float, float]]:
        value = read_json(path)
        result: list[tuple[float, float]] = []

        def add(row: Any, row_label: str) -> None:
            if not isinstance(row, dict) or "started" not in row or "ended" not in row:
                return
            finite(row["started"], f"{row_label} start")
            finite(row["ended"], f"{row_label} end")
            require(row["started"] <= row["ended"], f"{row_label} timestamps reversed")
            result.append((float(row["started"]), float(row["ended"])))

        if isinstance(value, list):
            rows = value
            value = {}
        else:
            add(value, label)
            rows = value.get("rows")
        if isinstance(rows, list):
            for index, row in enumerate(rows):
                add(row, f"{label}/{index}")
        receipt_value = value.get("receipts")
        if isinstance(receipt_value, dict):
            receipt_path = descriptor(receipt_value, f"{label} receipts", packet_bound=True)
            require(receipt_path is not None, f"{label} receipts missing")
            receipt_rows = read_json(receipt_path)
            require(isinstance(receipt_rows, list), f"{label} receipts malformed")
            for index, row in enumerate(receipt_rows):
                add(row, f"{label}/receipt{index}")
            require(result, f"{label} has no timing bounds")
        return result

    def accepted_admission_receipts(leg: str) -> list[Path]:
        admission_path = PACKET / f"artifact-admission-{leg}.json"
        require(admission_path.is_file() and not admission_path.is_symlink(),
                f"chronology admission {leg} is missing")
        admission = read_json(admission_path)
        fresh_bound = admission.get("fresh_bound")
        require(isinstance(fresh_bound, dict),
                f"chronology admission {leg} fresh binding is missing")
        attempt_files = fresh_bound.get("attempt_files")
        attempt_root = accepted_admission_attempt_root(leg, attempt_files)
        receipts: list[Path] = []
        for raw, item in attempt_files.items():
            path = descriptor(item, f"chronology admission {leg}/{raw}",
                              packet_bound=True)
            require(path is not None and path.resolve().is_relative_to(attempt_root.resolve()),
                    f"chronology admission {leg} descriptor escaped accepted root")
            if path.name.endswith("-receipt.json"):
                receipts.append(path)
        require(receipts, f"chronology admission {leg} receipts are missing")
        return sorted(receipts)

    # Admission's top-level JSON records only its terminal timestamp.  Its
    # child receipts carry the complete interval and remain retained under the
    # admission directory, so use those receipts as the gate's bounds.
    event_files: list[tuple[str, list[Path]]] = [
        ("quality", [PACKET / "quality.json"]),
        ("source-transition-before", [PACKET / "source-transition-before.json"]),
        ("build-before", [PACKET / "build-before/build.json"]),
        ("artifacts-before", [PACKET / "artifacts-before-receipt.json"]),
        ("admission-before", accepted_admission_receipts("before")),
        ("qualification-before", [PACKET / "qualification-before/receipts.json"]),
        ("source-transition-after", [PACKET / "source-transition-after.json"]),
        ("build-after", [PACKET / "build-after/build.json"]),
        ("artifacts-after", [PACKET / "artifacts-after-receipt.json"]),
        ("admission-after", accepted_admission_receipts("after")),
        ("qualification-after", [PACKET / "qualification-after/receipts.json"]),
        ("native", [PACKET / "native/receipts.json"]),
        ("observer", [PACKET / "observer/receipts.json"]),
    ]
    spans = []
    for name, paths in event_files:
        require(paths, f"chronology {name} evidence is missing")
        all_bounds: list[tuple[float, float]] = []
        for path in paths:
            require(path.is_file() and not path.is_symlink(),
                    f"chronology {name} file is missing: {path}")
            all_bounds.extend(bounds(path, f"chronology {name}/{path.name}"))
        spans.append({"name": name, "start": min(x[0] for x in all_bounds),
                      "end": max(x[1] for x in all_bounds), "rows": len(all_bounds)})
    for left, right in zip(spans, spans[1:]):
        require(left["end"] <= right["start"],
                f"chronology overlaps: {left['name']} before {right['name']}")
    return {"serial": True, "order": [x["name"] for x in spans], "spans": spans}


def render_native_csv(rows: list[dict[str, Any]]) -> str:
    fields = ("case", "format", "phase", "leg", "block", "p50", "p95", "p99",
              "mean", "rss_kib", "report")
    output = io.StringIO(newline="")
    writer = csv.DictWriter(output, fieldnames=fields, lineterminator="\n")
    writer.writeheader()
    for entry in rows:
        item = entry["stats"]
        case = entry["case"]
        writer.writerow({"case": case[2], "format": case[0], "phase": case[1],
                         "leg": entry["leg"], "block": entry["block"],
                         "p50": item["p50"], "p95": item["p95"],
                         "p99": item["p99"], "mean": item["mean"],
                         "rss_kib": entry["rss_kib"], "report": entry["report"]})
    return output.getvalue()


def render_markdown(value: dict[str, Any]) -> str:
    lines = ["# 0827 ordinary-save comparison", "",
             "Offline replay of full ordinary-save before/after evidence.",
             "Native timing and observer counters are separate; this packet records flags and makes no adoption decision.", "",
             f"Reports: {value['counts']['reports']}  Samples: {value['counts']['samples']}", "",
             "| case | p50 before | p50 after | after/before | CI95 | RSS ratio |",
             "|---|---:|---:|---:|---|---:|"]
    for row in value["native_rows"]:
        pair = row["paired"]["metrics"]["p50"]
        lines.append(f"| {row['case']} | {row['before']['process_median']['p50']:.9g} | "
                     f"{row['after']['process_median']['p50']:.9g} | "
                     f"{pair['median_ratio']:.9g} | "
                     f"[{pair['bootstrap']['ci_low']:.9g}, {pair['bootstrap']['ci_high']:.9g}] | "
                     f"{row['paired']['metrics']['rss_kib']['median_ratio']:.9g} |")
    lines.extend(["", "Observer allocation vectors and diagnostic process counters are retained in analysis.json.",
                   "No unsupported general full-save, additive-phase, pooled-lane, or adoption claim is made.", ""])
    return "\n".join(lines)


def replay() -> dict[str, Any]:
    plan = plan_value()
    inputs = origin_and_inputs(plan)
    live = live_source()
    quality = validate_quality(plan, live)
    corpus = corpus_inputs(plan)
    provenance = provenance_input(corpus)
    admission = validate_admission()
    sources = validate_source_states(plan, live)
    builds = {leg: load_build(leg, sources[leg]) for leg in LEGS}
    qualification = []
    for leg in LEGS:
        qualification.extend(load_lane(plan, sources, builds, "qualification",
                                       qualification_leg=leg,
                                       directory_name=f"qualification-{leg}",
                                       corpus=corpus))
    # Admission is a source/output witness when this packet carries the
    # ordinary-save artifact audit; it remains explicit in the derived record.
    native = load_lane(plan, sources, builds, "native", corpus=corpus)
    observer = load_lane(plan, sources, builds, "observer", corpus=corpus)
    require(len(native) == 144 and len(observer) == 48 and len(qualification) == 24,
            "lane report cardinality changed")
    native_groups = grouped(native, "native")
    observer_groups = grouped(observer, "observer")
    native_pairs = paired_native(native)
    observer_pairs = paired_observer(observer)
    native_rows = []
    for case in CASES:
        name = case["case"]
        native_rows.append({"case": name, "format": case["format"], "phase": case["phase"],
                            "before": native_groups[(name, "before")],
                            "after": native_groups[(name, "after")],
                            "paired": native_pairs[name]})
    observer_rows = []
    for case in CASES:
        name = case["case"]
        observer_rows.append({"case": name, "format": case["format"], "phase": case["phase"],
                              "before": observer_groups[(name, "before")],
                              "after": observer_groups[(name, "after")],
                              "paired": observer_pairs[name]})
    counts = {"qualification_reports": len(qualification),
              "qualification_samples": sum(len(x["samples"]) for x in qualification),
              "native_reports": len(native),
              "native_samples": sum(len(x["samples"]) for x in native),
              "observer_reports": len(observer),
              "observer_samples": sum(len(x["samples"]) for x in observer)}
    counts.update({"reports": sum(counts[k] for k in counts if k.endswith("_reports")),
                   "samples": sum(counts[k] for k in counts if k.endswith("_samples"))})
    require(counts == {"qualification_reports": 24, "qualification_samples": 24,
                       "native_reports": 144, "native_samples": 4320,
                       "observer_reports": 48, "observer_samples": 144,
                       "reports": 216, "samples": 4488}, "observed cardinality changed")
    chronology_value = chronology()
    native_entries = [{**entry, "case_name": entry["case"]} for entry in native]
    return {"schema": ANALYSIS_SCHEMA, "plan_schema": PLAN_SCHEMA, "base": BASE,
            "counts": counts, "inputs": inputs, "quality": quality,
            "admission": admission,
            "corpus": corpus,
            "provenance": provenance,
            "source": sources, "build": {leg: builds[leg]["receipt"] for leg in LEGS},
            "chronology": chronology_value,
            "qualification": [{"case": list(x["case"]), "leg": x["leg"],
                               "block": x["block"], "report": x["report"],
                               "report_sha256": x["report_sha256"]} for x in qualification],
            "native_rows": native_rows, "observer_rows": observer_rows,
            "native": {"reports": len(native), "samples": counts["native_samples"],
                        "processes": native_entries, "paired": native_pairs},
            "observer": {"reports": len(observer), "samples": counts["observer_samples"],
                          "processes": observer, "paired": observer_pairs},
            "decision_flags": decision_flags(native_pairs, observer_pairs),
            "verification": {
                "full_source_census_checked": True,
                "two_file_before_archive_checked": True,
                "after_live_head_checked": True,
                "quality_and_pptx_reuse_checked": True,
                "build_receipts_checked": True,
                "capture_receipts_checked": True,
                "report_schema_checked": True,
                "input_output_admission_checked": True,
                "corpus_identity_checked": True,
                "raw_quantiles_checked": True,
                "native_timing_separated": True,
                "observer_counters_retained": True,
                "paired_ci_checked": True,
                "serial_chronology_checked": True,
                "no_adoption_decision": True,
                "no_additive_phase_medians": True,
            }}


def rendered(value: dict[str, Any]) -> dict[str, str]:
    return {"analysis.json": json.dumps(value, indent=2, sort_keys=True) + "\n",
            "native.csv": render_native_csv(value["native"]["processes"]),
            "analysis.md": render_markdown(value)}


def analyze(*, write: bool = False, check: bool = False) -> dict[str, Any]:
    value = replay()
    outputs = rendered(value)
    if write:
        for name in outputs:
            require(not (PACKET / name).exists(), f"refusing to overwrite retained {name}")
        for name, text in outputs.items():
            (PACKET / name).write_text(text, encoding="utf-8")
    if check:
        for name, text in outputs.items():
            path = PACKET / name
            require(path.is_file() and path.read_text(encoding="utf-8") == text,
                    f"{name} does not replay byte-for-byte")
    return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group()
    group.add_argument("--write", action="store_true")
    group.add_argument("--check", action="store_true")
    args = parser.parse_args(argv)
    try:
        value = analyze(write=args.write or not args.check, check=args.check)
        print(json.dumps({"status": "replayed", "reports": value["counts"]["reports"],
                          "samples": value["counts"]["samples"]}, sort_keys=True))
    except (ReplayError, AssertionError, OSError, ValueError, KeyError, TypeError,
            IndexError) as error:
        print(f"0827 analysis failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
