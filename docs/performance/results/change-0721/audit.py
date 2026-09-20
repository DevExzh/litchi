#!/usr/bin/env python3
"""Audit the terminal 0721 DOCX fusion packet.

The packet is deliberately incomplete while captures are being prepared.  The
default command is a terminal audit: it requires both source-bound analyzers,
all frozen child artifacts, the oracle and trace custody records, an explicit
source disposition, and cleanup witnesses for removed binaries.  ``--draft``
checks only the frozen preparation that exists so far and reports the missing
terminal inputs without describing a result.

This script never builds or executes a Rust binary. Its replay subprocesses
run the packet's Python analyzers, and correction diffs are applied only to
disposable temporary copies; those checks revalidate retained reports and
source/build identities after owned binaries have been removed.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
TARGET = ROOT.parent / "litchi-target-0721"
BINARY_ROOT = ROOT.parent / "litchi-0721-bin"
FILESYSTEM_ROOT = ROOT.parent / "litchi-0721-fs"
TRACE_SCRATCH = ROOT.parent / "litchi-scratch-0721-trace"

HEX = set("0123456789abcdef")
STAGES = (
    "baseline-A1", "candidate-B1", "candidate-B2", "baseline-A2",
    "baseline-A3", "candidate-B3", "candidate-B4", "baseline-A4",
)
ALLOCATOR_STAGES = STAGES[:4]
PHASES = ("edit", "lifecycle")
LANES = ("native", "allocator")
SOURCE_DELTA = {
    "crates/litchi-docx/src/alt/codec.rs",
    "crates/litchi-docx/src/alt/mod.rs",
    "crates/litchi-docx/src/namespace.rs",
    "crates/litchi-docx/src/parts/document_part.rs",
    "crates/litchi-docx/src/writer/doc/fusion_tests_0721.rs",
    "crates/litchi-docx/src/writer/doc/package.rs",
}
TRACE_BASELINE_PATCHED = {
    "crates/litchi-docx/src/lib.rs",
    "crates/litchi-docx/src/alt/codec.rs",
    "crates/litchi-docx/src/namespace.rs",
    "crates/litchi-docx/src/parts/document_part.rs",
    "crates/litchi-docx/src/writer/doc/package.rs",
}
TRACE_CANDIDATE_PATCHED = {
    "crates/litchi-docx/src/lib.rs",
    "crates/litchi-docx/src/alt/codec.rs",
    "crates/litchi-docx/src/namespace.rs",
    "crates/litchi-docx/src/writer/doc/package.rs",
}
OWNED_PATHS = [TARGET, BINARY_ROOT, FILESYSTEM_ROOT, TRACE_SCRATCH]
ORACLE_PROBE_FILES = ("Cargo.toml", "Cargo.lock", "src/main.rs")
EXPECTED_CORPUS = {
    "generated": {
        "id": "generated",
        "origin": "generated-harness-corpus",
        "path": None,
        "sha256": None,
        "bytes": None,
    },
    "numbered-list": {
        "id": "numbered-list",
        "origin": "caller-named-real-file",
        "path": "test-data/libreoffice-core/sw/qa/extras/ooxmlexport/data/NumberedList.docx",
        "sha256": "ebb078d791b6deb4d0a2dae69a15a5274baf2ade815f07b336b394bb43af4c88",
        "bytes": 55519,
    },
}


class AuditError(AssertionError):
    """A custody or final-evidence failure."""


class DraftPending(Exception):
    """The preparation check completed, but terminal artifacts are absent."""

    def __init__(self, pending: Iterable[str]):
        self.pending = tuple(dict.fromkeys(pending))
        super().__init__("terminal evidence is not available")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def canonical(value: Any) -> bytes:
    return (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()


def digest(value: Any) -> str:
    return hashlib.sha256(canonical(value)).hexdigest()


def check_digest(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def load_custody() -> Any:
    spec = importlib.util.spec_from_file_location("change0721_audit_custody", PACKET / "custody.py")
    require(spec is not None and spec.loader is not None, "cannot load custody.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def source_census(custody: Any) -> dict[str, str]:
    value = custody.census()
    require(isinstance(value, dict) and value, "source census is empty")
    require(all(isinstance(name, str) and isinstance(value_hash, str)
                for name, value_hash in value.items()),
            "source census is not a path-to-SHA map")
    return value


def load_plan() -> dict[str, Any]:
    plan = read_json(PACKET / "plan.json")
    require(isinstance(plan, dict) and plan.get("schema_version") == 1,
            "primary plan schema changed")
    require(plan.get("packet") == "change-0721-docx-structural-scan-fusion-pilot",
            "primary packet identity changed")
    revision = read_json(PACKET / "revision.json")
    require(plan.get("revision") == revision.get("revision"),
            "primary plan revision changed")
    require(plan.get("cpu") == 12 and plan.get("filesystem_root") == str(FILESYSTEM_ROOT),
            "primary execution binding changed")
    require(plan.get("phase_order") == list(PHASES), "primary phase order changed")
    require(plan.get("native") == {"samples": 200, "warmup": 100},
            "native sample plan changed")
    require(plan.get("allocator") == {"samples": 3, "warmup": 0},
            "allocator sample plan changed")
    require([item.get("label") for item in plan.get("stages", [])] == list(STAGES),
            "primary stage order changed")
    require(plan.get("allocator_stages") == list(ALLOCATOR_STAGES),
            "allocator stage scope changed")
    require(set(plan.get("source_delta_allowlist", [])) == SOURCE_DELTA,
            "source delta allowlist changed")
    require(plan.get("environment") == {
        "LC_ALL": "C", "LANG": "C", "TZ": "UTC",
        "RUSTFLAGS": None, "LD_PRELOAD": None, "MALLOC_CONF": None,
        "GLIBC_TUNABLES": None, "PERL_HASH_SEED": "0", "PERL_PERTURB_KEYS": "0",
    }, "child environment changed")
    require(plan.get("thresholds") == {
        "edit_improvement_percent": 3,
        "lifecycle_regression_percent": 3,
        "allocation_regression_percent": 3,
        "tail_flag_percent": 5,
        "repeat_drift_flag_percent": 5,
    }, "primary thresholds changed")
    corpora = plan.get("corpora")
    require(isinstance(corpora, list) and [item.get("id") for item in corpora]
            == ["generated", "numbered-list"], "corpus order changed")
    for corpus in corpora:
        expected = EXPECTED_CORPUS[corpus["id"]]
        for key, value in expected.items():
            require(corpus.get(key) == value, f"corpus binding changed: {corpus['id']}/{key}")
    return plan


def load_read_control_plan(plan: dict[str, Any]) -> dict[str, Any]:
    value = read_json(PACKET / "read-controls-plan.json")
    require(isinstance(value, dict) and value.get("schema_version") == 1,
            "read-control plan schema changed")
    require(value.get("revision") == plan["revision"]
            and value.get("packet") == "change-0721-docx-read-controls",
            "read-control plan identity changed")
    require(value.get("cpu") == 12 and value.get("filesystem_root") == str(FILESYSTEM_ROOT)
            and value.get("binary_root") == str(BINARY_ROOT),
            "read-control execution binding changed")
    require([item.get("label") for item in value.get("stages", [])] == list(STAGES),
            "read-control stage order changed")
    require(value.get("native") == {"samples": 200, "warmup": 100},
            "read-control sample plan changed")
    require([item.get("id") for item in value.get("controls", [])] == [
        "generated-medium-list-paragraphs", "pinned-media-eager-paragraph-count",
    ], "read-control list changed")
    require(value.get("source_maps") == {
        "baseline": "docs/performance/results/change-0721/source-baseline.json",
        "candidate": "docs/performance/results/change-0721/source-candidate.json",
    }, "read-control source map binding changed")
    require(value.get("constraints") == "docs/performance/results/change-0721/constraints.json",
            "read-control constraints binding changed")
    return value


def validate_constraints() -> str:
    constraints = read_json(PACKET / "constraints.json")
    require(isinstance(constraints, dict) and constraints, "constraints are empty")
    for raw, expected in constraints.items():
        check_digest(expected, f"constraint {raw}")
        target = ROOT / raw
        require(target.is_file() and not target.is_symlink() and sha(target) == expected,
                f"constraint changed: {raw}")
    return sha(PACKET / "constraints.json")


def validate_sources(custody: Any, plan: dict[str, Any], *, draft: bool) -> tuple[
    dict[str, str], dict[str, str], dict[str, str], str | None, dict[str, Any] | None
]:
    baseline = read_json(PACKET / "source-baseline.json")
    candidate = read_json(PACKET / "source-candidate.json")
    require(isinstance(baseline, dict) and baseline, "baseline source map is invalid")
    require(isinstance(candidate, dict) and candidate, "candidate source map is invalid")
    changed = sorted(name for name in set(baseline) | set(candidate)
                     if baseline.get(name) != candidate.get(name))
    require(set(changed) == SOURCE_DELTA, f"source delta changed: {changed}")
    current = source_census(custody)
    final_path = PACKET / "source-final.json"
    disposition_path = PACKET / "disposition.json"
    require(final_path.exists() == disposition_path.exists(),
            "source-final.json and disposition.json must be created together")
    if final_path.exists():
        final_source = read_json(final_path)
        disposition = read_json(disposition_path)
        require(final_source in (baseline, candidate),
                "source-final.json is neither the frozen baseline nor candidate")
        require(isinstance(disposition, dict)
                and set(disposition) == {"retained", "final_source"},
                "disposition schema changed")
        label = disposition.get("final_source")
        require(label in {"baseline", "candidate"}, "disposition source label changed")
        require(final_source == (baseline if label == "baseline" else candidate),
                "source-final.json does not match its disposition")
        require(isinstance(disposition.get("retained"), bool)
                and disposition["retained"] == (label == "candidate"),
                "disposition retained flag is inconsistent")
        require(current == final_source, "current checkout does not match source-final.json")
        return baseline, candidate, final_source, label, disposition
    require(draft and current == candidate,
            "terminal source disposition is missing or live source is not candidate")
    return baseline, candidate, candidate, None, None


def validate_source_snapshots(baseline: dict[str, str], candidate: dict[str, str], *, draft: bool) -> None:
    for lane, source, paths in (
        ("baseline", baseline, TRACE_BASELINE_PATCHED),
        ("candidate", candidate, TRACE_CANDIDATE_PATCHED),
    ):
        for relative in paths:
            snapshot = PACKET / lane / relative
            if not snapshot.is_file():
                if relative == "crates/litchi-docx/src/lib.rs":
                    # trace.py binds lib.rs to the unchanged checkout hash;
                    # there is intentionally no lane snapshot for this file.
                    continue
                if draft:
                    continue
                raise AuditError(f"missing {lane} trace source snapshot: {snapshot}")
            if sha(snapshot) != source[relative]:
                if draft:
                    continue
                raise AuditError(f"{lane} trace source snapshot changed: {relative}")
    candidate_test = PACKET / "candidate/crates/litchi-docx/src/writer/doc/fusion_tests_0721.rs"
    if candidate_test.is_file():
        if sha(candidate_test) != candidate["crates/litchi-docx/src/writer/doc/fusion_tests_0721.rs"] \
                and not draft:
            raise AuditError("candidate test source snapshot changed")
    elif not draft:
        raise AuditError(f"missing candidate test source snapshot: {candidate_test}")


def cleanup_witnesses() -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = PACKET / filename
        if not path.is_file() or path.is_symlink():
            continue
        value = read_json(path)

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest_value = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw_path, str) and isinstance(digest_value, str):
                    check_digest(digest_value, f"cleanup witness {raw_path}")
                    if size is not None:
                        positive_int(size, f"cleanup witness {raw_path}.bytes")
                    result.append({"path": raw_path, "sha256": digest_value, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return result


def resolve_artifact(raw: str, *, packet_relative: bool = True) -> Path:
    path = Path(raw)
    if path.is_absolute():
        return path.resolve()
    if packet_relative:
        return (PACKET / path).resolve()
    return (ROOT / path).resolve()


def custody_path(path: Path, expected_sha: str, expected_bytes: int | None,
                 witnesses: list[dict[str, Any]], label: str) -> str:
    check_digest(expected_sha, f"{label}.sha256")
    if expected_bytes is not None:
        positive_int(expected_bytes, f"{label}.bytes")
    if path.is_file() and not path.is_symlink():
        require(sha(path) == expected_sha, f"live {label} hash changed: {path}")
        if expected_bytes is not None:
            require(path.stat().st_size == expected_bytes, f"live {label} size changed: {path}")
        return "live"
    matches = [item for item in witnesses
               if resolve_artifact(item["path"]) == path
               and item["sha256"] == expected_sha
               and (expected_bytes is None or item.get("bytes") == expected_bytes)]
    require(len(matches) == 1, f"missing exact cleanup witness for {label}: {path}")
    return "cleanup-witness"


def build_command(stage: str, lane: str) -> list[str]:
    binary = "litchi-perf-baseline" if lane == "native" else "litchi-perf-baseline-alloc"
    command = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", binary,
        "--target-dir", str(TARGET), "-j", "2",
    ]
    if lane == "allocator":
        command += ["--features", "allocator-metrics"]
    return command


def validate_builds(baseline: dict[str, str], candidate: dict[str, str],
                    witnesses: list[dict[str, Any]], *, draft: bool) -> dict[tuple[str, str], dict[str, Any]]:
    builds: dict[tuple[str, str], dict[str, Any]] = {}
    for stage in ("baseline", "candidate"):
        path = PACKET / f"build-{stage}.json"
        if not path.is_file():
            if draft:
                continue
            raise AuditError(f"missing build record: {path}")
        rows = read_json(path)
        require(isinstance(rows, list), f"{path.name} is not a build record list")
        require({row.get("lane") for row in rows} == {"native", "alloc"},
                f"{path.name} must contain native and alloc exactly once")
        require(len(rows) == 2, f"{path.name} contains duplicate build rows")
        source = baseline if stage == "baseline" else candidate
        source_manifest = PACKET / f"source-{stage}.json"
        require(read_json(source_manifest) == source,
                f"source-{stage}.json differs from its source pair")
        for row in rows:
            raw_lane = row["lane"]
            lane = "native" if raw_lane == "native" else "allocator"
            key = (stage, lane)
            require(key not in builds, f"duplicate build identity: {key}")
            require(row.get("command") == build_command(stage, lane),
                    f"{stage}/{lane} build command changed")
            require(row.get("exit_code") == 0, f"{stage}/{lane} build failed")
            require(row.get("source_manifest_sha256") == sha(source_manifest),
                    f"{stage}/{lane} source manifest binding changed")
            log = PACKET / f"build-{stage}-{raw_lane}.log"
            check_digest(row.get("log_sha256"), f"{stage}/{lane} build log")
            require(log.is_file() and sha(log) == row["log_sha256"],
                    f"{stage}/{lane} build log changed")
            expected_binary = BINARY_ROOT / f"{stage}-{raw_lane}"
            require(row.get("binary") == str(expected_binary),
                    f"{stage}/{lane} binary path changed")
            binary_sha = row.get("binary_sha256")
            binary_bytes = row.get("binary_bytes")
            require(isinstance(binary_bytes, int) and binary_bytes > 0,
                    f"{stage}/{lane} binary size is malformed")
            mode = custody_path(expected_binary, binary_sha, binary_bytes, witnesses,
                                f"{stage}/{lane} binary")
            builds[key] = {
                "stage": stage, "lane": lane, "raw_lane": raw_lane,
                "record": row, "record_path": path, "record_sha256": sha(path),
                "source_manifest": source_manifest,
                "source_manifest_sha256": sha(source_manifest), "source": source,
                "binary": expected_binary, "binary_sha256": binary_sha,
                "binary_bytes": binary_bytes, "binary_custody": mode,
            }
    return builds


def validate_oracle(baseline: dict[str, str], candidate: dict[str, str],
                    witnesses: list[dict[str, Any]], *, draft: bool) -> list[Path]:
    oracle_root = PACKET / "oracle"
    required = [oracle_root / name for name in ("Cargo.toml", "Cargo.lock", "src/main.rs")]
    if not all(path.is_file() for path in required):
        if draft:
            return []
        raise AuditError("oracle sources are incomplete")
    outputs: list[Path] = []
    reports: list[bytes] = []
    for stage, source in (("baseline", baseline), ("candidate", candidate)):
        directory = oracle_root / stage
        names = ("source.json", "probe.json", "build.json", "build.log",
                 "result.json", "report.json", "stdout", "stderr")
        if not all((directory / name).is_file() for name in names):
            if draft:
                continue
            raise AuditError(f"oracle {stage} artifacts are incomplete")
        source_path = directory / "source.json"
        require(read_json(source_path) == source and sha(source_path) == sha(PACKET / f"source-{stage}.json"),
                f"oracle {stage} source binding changed")
        probe = read_json(directory / "probe.json")
        require(set(probe) == set(ORACLE_PROBE_FILES), f"oracle {stage} probe file set changed")
        for raw in ORACLE_PROBE_FILES:
            target = oracle_root / raw
            require(target.is_file() and sha(target) == probe[raw],
                    f"oracle {stage} probe binding changed: {raw}")
        build = read_json(directory / "build.json")
        expected_cmd = [
            "cargo", "build", "--release", "--locked", "--manifest-path",
            str(oracle_root / "Cargo.toml"), "--target-dir", str(TARGET), "-j", "2",
        ]
        require(build.get("command") == expected_cmd and build.get("exit_code") == 0,
                f"oracle {stage} build record changed")
        require(build.get("source_sha256") == sha(source_path)
                and build.get("probe_sha256") == sha(directory / "probe.json")
                and build.get("log_sha256") == sha(directory / "build.log"),
                f"oracle {stage} build custody changed")
        result = read_json(directory / "result.json")
        binary_record = result.get("binary")
        require(isinstance(binary_record, dict), f"oracle {stage} binary record is missing")
        expected_binary = BINARY_ROOT / f"{stage}-oracle"
        require(binary_record.get("path") == str(expected_binary),
                f"oracle {stage} binary path changed")
        mode = custody_path(expected_binary, binary_record.get("sha256"),
                            binary_record.get("bytes"), witnesses,
                            f"oracle {stage} binary")
        require(result.get("command") == [str(expected_binary), "--output", str(directory / "report.json")]
                and result.get("exit_code") == 0,
                f"oracle {stage} result command changed")
        require(result.get("source_sha256") == sha(source_path)
                and result.get("probe_sha256") == sha(directory / "probe.json"),
                f"oracle {stage} result custody changed")
        artifacts = result.get("artifacts")
        require(artifacts == {name: sha(directory / name) for name in ("report.json", "stdout", "stderr")},
                f"oracle {stage} artifact hashes changed")
        report = read_json(directory / "report.json")
        require(report.get("case_count") == 19 and len(report.get("cases", [])) == 19
                and len(report.get("packages", [])) == 2,
                f"oracle {stage} public matrix count changed")
        reports.append((directory / "report.json").read_bytes())
        outputs.append(expected_binary)
        # The build result itself does not carry a binary record; retain its
        # identity in the returned artifact list for the cleanup audit.
        _ = mode
    if len(reports) == 2:
        require(reports[0] == reports[1], "baseline and candidate oracle reports differ")
    return outputs


def validate_quality(final_label: str, final_source: dict[str, str]) -> None:
    quality_path = PACKET / "quality-docx.json"
    rows = read_json(quality_path)
    require(isinstance(rows, list)
            and [row.get("name") for row in rows]
            == ["fmt", "tests", "clippy", "doctests", "rustdoc"],
            "quality-docx.json does not contain the five scoped checks")
    target = str(TARGET)
    expected = {
        "fmt": ["cargo", "fmt", "-p", "litchi-docx", "--", "--check"],
        "tests": ["cargo", "test", "-p", "litchi-docx", "--locked", "--target-dir", target,
                  "-j", "2", "--all-features", "--all-targets"],
        "clippy": ["cargo", "clippy", "-p", "litchi-docx", "--locked", "--target-dir", target,
                   "-j", "2", "--all-features", "--all-targets", "--", "-D", "warnings"],
        "doctests": ["cargo", "test", "-p", "litchi-docx", "--locked", "--target-dir", target,
                     "-j", "2", "--all-features", "--doc"],
        "rustdoc": ["cargo", "doc", "-p", "litchi-docx", "--locked", "--target-dir", target,
                    "-j", "2", "--all-features", "--no-deps"],
    }
    source_manifest = PACKET / "quality-docx-source.json"
    # The DOCX quality pass is intentionally a candidate qualification pass;
    # a restored baseline is checked by the final harness/evidence rows below.
    require(read_json(source_manifest) == read_json(PACKET / "source-candidate.json"),
            "quality-docx must qualify the candidate source")
    for row in rows:
        name = row.get("name")
        require(row.get("exit_code") == 0 and row.get("command") == expected[name],
                f"quality command changed: {name}")
        require(row.get("source_manifest_sha256") == sha(source_manifest),
                f"quality source binding changed: {name}")
        log = PACKET / row.get("log", "")
        check_digest(row.get("log_sha256"), f"quality log {name}")
        require(log.is_file() and sha(log) == row["log_sha256"], f"quality log changed: {name}")
    _ = final_source


def validate_final_quality(final_source: dict[str, str]) -> None:
    harness = PACKET / "quality-harness.json"
    evidence = PACKET / "quality-evidence.json"
    require(harness.is_file() and evidence.is_file(),
            "final-source harness and evidence quality records are required")
    rows = read_json(harness)
    require(isinstance(rows, list) and [row.get("name") for row in rows] == ["tests"],
            "quality-harness.json does not contain its one check")
    expected_command = [
        "cargo", "test", "--release", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--locked", "--target-dir", str(TARGET), "-j", "2", "--lib",
    ]
    source_manifest = PACKET / "quality-harness-source.json"
    # The standalone harness qualification is captured once against the
    # candidate before the timed matrix.  Repository evidence below is the
    # final-source gate; a rejected candidate does not require recompiling the
    # unchanged harness solely to change this custody label.
    require(read_json(source_manifest) == read_json(PACKET / "source-candidate.json"),
            "quality-harness source is not the candidate qualification source")
    for row in rows:
        require(row.get("exit_code") == 0 and row.get("command") == expected_command,
                "final harness quality command changed")
        require(row.get("source_manifest_sha256") == sha(source_manifest),
                "final harness source binding changed")
        log = PACKET / row.get("log", "")
        check_digest(row.get("log_sha256"), "final harness log")
        require(log.is_file() and sha(log) == row["log_sha256"], "final harness log changed")
    expected_rows = read_json(ROOT / "docs/performance/results/change-0717/evidence/results.json")
    rows = read_json(evidence)
    require(isinstance(rows, list) and [row.get("name") for row in rows]
            == [row.get("name") for row in expected_rows],
            "quality-evidence gate names changed")
    source_manifest = PACKET / "quality-evidence-source.json"
    require(read_json(source_manifest) == final_source,
            "quality-evidence source is not the final source")
    for row, expected in zip(rows, expected_rows):
        require(row.get("exit_code") == 0 and row.get("command") == expected.get("command"),
                f"final evidence command changed: {row.get('name')}")
        require(row.get("source_manifest_sha256") == sha(source_manifest),
                f"final evidence source binding changed: {row.get('name')}")
        log = PACKET / row.get("log", "")
        check_digest(row.get("log_sha256"), f"final evidence log {row.get('name')}")
        require(log.is_file() and sha(log) == row["log_sha256"],
                f"final evidence log changed: {row.get('name')}")


def validate_capture_freeze(plan: dict[str, Any], *, draft: bool) -> None:
    path = PACKET / "capture-freeze.json"
    if not path.is_file():
        if draft:
            return
        raise AuditError("capture-freeze.json is missing")
    value = read_json(path)
    require(isinstance(value, dict), "capture-freeze.json is not an object")
    files = value.get("files")
    require(isinstance(files, dict) and files, "capture freeze file map is missing")
    packet_prefix = str(PACKET.relative_to(ROOT)) + "/"
    observed_packet = {
        raw.removeprefix(packet_prefix) for raw in files if raw.startswith(packet_prefix)
    }
    required = {
        "plan.json", "pilot.py", "read-controls.py", "read-controls-plan.json",
        "analyze.py", "capture.py", "custody.py", "source-baseline.json",
        "source-candidate.json", "constraints.json",
    }
    # Freeze paths are repository relative so a replay can verify them even
    # after packet cleanup.  The coordinator may retain additional helper
    # bindings, but the core recipe cannot be omitted.
    missing = sorted(required - observed_packet)
    if missing:
        if draft:
            raise DraftPending((
                "capture-freeze.json required recipe bindings: " + ", ".join(missing),
            ))
        raise AuditError(
            "capture freeze omits required recipe bindings: " + ", ".join(missing)
        )
    for raw, expected in files.items():
        require(isinstance(raw, str) and not Path(raw).is_absolute()
                and not raw.startswith("../"), f"capture freeze path is unsafe: {raw}")
        check_digest(expected, f"capture freeze {raw}")
        target = ROOT / raw
        if not (target.is_file() and not target.is_symlink() and sha(target) == expected):
            if draft:
                raise DraftPending(("capture-freeze.json final script hashes",))
            raise AuditError(f"capture freeze binding changed: {raw}")
    require(isinstance(value.get("utc_created"), str) and value["utc_created"],
            "capture freeze timestamp is missing")
    order = value.get("order")
    require(isinstance(order, str) and "native" in order and "allocator" in order,
            "capture freeze order is missing")


def validate_capture_orchestrator(*, draft: bool) -> None:
    path = PACKET / "capture.json"
    if not path.is_file():
        if draft:
            return
        raise AuditError("capture.json is missing")
    rows = read_json(path)
    if not (isinstance(rows, list) and len(rows) == 20):
        if draft:
            raise DraftPending(("complete 20-row capture orchestrator receipt",))
        raise AuditError("capture.json must retain exactly 20 orchestrator commands")
    expected: list[list[str]] = []
    for stage in STAGES:
        expected.append([sys.executable, str(PACKET / "pilot.py"), stage, "native"])
        expected.append([sys.executable, str(PACKET / "read-controls.py"), "capture", stage])
    for stage in ALLOCATOR_STAGES:
        expected.append([sys.executable, str(PACKET / "pilot.py"), stage, "allocator"])
    for index, (row, command) in enumerate(zip(rows, expected)):
        if not (isinstance(row, dict) and row.get("command") == command
                and row.get("exit_code") == 0):
            if draft:
                raise DraftPending((f"capture orchestrator row {index} completion",))
            raise AuditError(f"capture orchestrator row {index} changed")


def validate_analyzer_corrections() -> None:
    """Bind analysis-only replacements to their frozen originals and diffs."""

    path = PACKET / "analyzer-corrections.json"
    value = read_json(path)
    require(isinstance(value, dict)
            and isinstance(value.get("utc"), str) and value["utc"]
            and isinstance(value.get("reason"), str) and value["reason"],
            "analyzer correction metadata is incomplete")
    corrections = value.get("corrections")
    require(isinstance(corrections, list) and len(corrections) == 2,
            "analyzer correction inventory changed")
    expected = {
        "analyze.py": "analyze-final.py",
        "read-controls.py": "read-controls-analyze.py",
    }
    seen: set[str] = set()
    for item in corrections:
        require(isinstance(item, dict)
                and set(item) == {
                    "capture_time_script", "capture_time_sha256", "final_analyzer",
                    "final_analyzer_sha256", "diff", "diff_sha256",
                }, "analyzer correction row schema changed")
        capture_name = item["capture_time_script"]
        final_name = item["final_analyzer"]
        require(capture_name in expected and expected[capture_name] == final_name,
                "analyzer correction source mapping changed")
        require(capture_name not in seen, "duplicate analyzer correction source")
        seen.add(capture_name)
        capture_path = PACKET / capture_name
        final_path = PACKET / final_name
        diff_name = item["diff"]
        require(isinstance(diff_name, str)
                and Path(diff_name).name == diff_name
                and Path(diff_name).suffix == ".diff",
                "analyzer correction diff path is unsafe")
        diff_path = PACKET / diff_name
        require(capture_path.is_file() and final_path.is_file() and diff_path.is_file(),
                "analyzer correction artifact is missing")
        require(item["capture_time_sha256"] == sha(capture_path)
                and item["final_analyzer_sha256"] == sha(final_path)
                and item["diff_sha256"] == sha(diff_path),
                f"analyzer correction hashes changed: {capture_name}")
        diff_lines = diff_path.read_text(encoding="utf-8").splitlines()
        require(len(diff_lines) >= 2
                and diff_lines[0] == f"--- {capture_name}"
                and diff_lines[1] == f"+++ {final_name}",
                f"analyzer correction diff headers changed: {capture_name}")
        # Apply the retained unified diff to a disposable copy.  This checks
        # the claimed original-to-final relationship without touching either
        # packet script; the hashes above still bind the retained files.
        with tempfile.TemporaryDirectory(prefix="litchi-0721-correction-") as directory:
            scratch = Path(directory) / capture_name
            shutil.copy2(capture_path, scratch)
            result = subprocess.run(
                ["patch", "--silent", capture_name],
                cwd=directory,
                input=diff_path.read_bytes(),
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
            )
            require(result.returncode == 0,
                    f"analyzer correction diff does not apply: {capture_name}")
            require(scratch.read_bytes() == final_path.read_bytes(),
                    f"analyzer correction diff output changed: {capture_name}")
    require(seen == set(expected), "analyzer correction sources are incomplete")


def validate_trace(baseline: dict[str, str], candidate: dict[str, str],
                   witnesses: list[dict[str, Any]], *, draft: bool) -> list[Path]:
    """Validate the trace runner's source, probe, binary, and output custody."""
    path = PACKET / "trace-analysis.json"
    if not path.is_file():
        if draft:
            return []
        raise AuditError("trace-analysis.json is missing")
    value = read_json(path)
    require(isinstance(value, dict)
            and value.get("schema") == "litchi.docx-trace-analysis-0721.v1",
            "trace analysis schema changed")
    require(value.get("status") == "pass",
            "trace analysis does not record a passing untimed diagnostic")
    require(value.get("document_count") == 21
            and isinstance(value.get("documents"), list)
            and len(value["documents"]) == 21,
            "trace document matrix count changed")
    require(value.get("baseline_trace_sha256")
            == sha(PACKET / "trace/baseline/stderr")
            and value.get("candidate_trace_sha256")
            == sha(PACKET / "trace/candidate/stderr"),
            "trace stderr bindings changed")
    public_report = PACKET / "oracle/baseline/report.json"
    require(value.get("public_report_sha256") == sha(public_report)
            and value.get("public_report_bytes") == public_report.stat().st_size,
            "trace public report binding changed")
    reader_proof = value.get("reader_proof")
    require(isinstance(reader_proof, dict)
            and isinstance(reader_proof.get("baseline_observer_events"), int)
            and isinstance(reader_proof.get("baseline_total_reads"), int)
            and isinstance(reader_proof.get("candidate_observer_events"), int)
            and isinstance(reader_proof.get("candidate_total_reads"), int)
            and reader_proof.get("candidate_range_reader_documents") == [],
            "trace reader proof changed")
    expected_sources = {"baseline": baseline, "candidate": candidate}
    binary_paths: list[Path] = []
    for lane, source in expected_sources.items():
        directory = PACKET / "trace" / lane
        required = (
            "source.json", "patch.json", "build.json", "build.log",
            "capture.json", "capture-repeat.json", "report.json",
            "report-repeat.json", "stdout", "stdout-repeat", "stderr", "stderr-repeat",
        )
        if not all((directory / name).is_file() for name in required):
            if draft:
                continue
            raise AuditError(f"trace {lane} artifacts are incomplete")
        source_path = directory / "source.json"
        trace_source = read_json(source_path)
        require(trace_source != source,
                f"trace {lane} source did not retain transformed source")
        patch_set = TRACE_BASELINE_PATCHED if lane == "baseline" else TRACE_CANDIDATE_PATCHED
        patch = read_json(directory / "patch.json")
        require(isinstance(patch, dict)
                and patch.get("version") == 1
                and patch.get("lane") == lane
                and patch.get("status") == "applied"
                and patch.get("fragment_sha256") == sha(PACKET / "trace.fragment"),
                f"trace {lane} patch custody changed")
        instrument = patch
        require(instrument.get("fragment_sha256") == sha(PACKET / "trace.fragment"),
                f"trace {lane} fragment binding changed")
        manifest_sources = instrument.get("sources")
        require(isinstance(manifest_sources, list), f"trace {lane} source patch list is missing")
        by_relative = {item.get("relative"): item for item in manifest_sources
                       if isinstance(item, dict)}
        require(set(by_relative) == patch_set, f"trace {lane} patched source set changed")
        for relative, item in by_relative.items():
            require(item.get("original_sha256") == source[relative],
                    f"trace {lane} original source binding changed: {relative}")
            check_digest(item.get("transformed_sha256"), f"trace {lane} transformed source {relative}")
            require(trace_source.get(relative) == item["transformed_sha256"],
                    f"trace {lane} transformed source census changed: {relative}")
        # The temporary tracer also patches the unchanged litchi-docx lib.rs;
        # its original identity is carried by the patch record rather than a
        # source snapshot in the lane directory.
        for relative, value_hash in trace_source.items():
            if relative not in patch_set:
                expected = source.get(relative)
                if expected is not None:
                    require(value_hash == expected,
                            f"trace {lane} unrelated source changed: {relative}")
        build = read_json(directory / "build.json")
        expected_build = [
            "cargo", "build", "--locked", "--manifest-path",
            str(PACKET / "oracle/Cargo.toml"), "--target-dir", str(TARGET), "-j", "2",
        ]
        require(build.get("command") == expected_build and build.get("exit_code") == 0,
                f"trace {lane} build record changed")
        require(build.get("source_sha256") == sha(source_path)
                and build.get("patch_sha256") == sha(directory / "patch.json")
                and build.get("log_sha256") == sha(directory / "build.log"),
                f"trace {lane} build custody changed")
        scripts = build.get("scripts")
        require(scripts == {
            name: sha(PACKET / name)
            for name in ("trace-run.py", "trace.py", "trace.fragment", "trace-analyze.py")
        }, f"trace {lane} script custody changed")
        probe = build.get("probe")
        require(probe == {
            name: sha(PACKET / "oracle" / name) for name in ORACLE_PROBE_FILES
        }, f"trace {lane} probe custody changed")
        binary = build.get("binary")
        require(isinstance(binary, dict), f"trace {lane} binary custody is missing")
        binary_path = resolve_artifact(str(binary.get("path")))
        binary_paths.append(binary_path)
        custody_path(binary_path, binary.get("sha256"), binary.get("bytes"), witnesses,
                     f"trace {lane} binary")
        for capture_name in ("capture", "capture-repeat"):
            capture = read_json(directory / f"{capture_name}.json")
            report_name = "report.json" if capture_name == "capture" else "report-repeat.json"
            stdout_name = "stdout" if capture_name == "capture" else "stdout-repeat"
            stderr_name = "stderr" if capture_name == "capture" else "stderr-repeat"
            require(capture.get("schema") == "litchi.docx-trace-capture-0721.v1"
                    and capture.get("exit_code") == 0
                    and capture.get("command") == [binary["path"], "--output", str(directory / report_name)]
                    and capture.get("cwd") == str(ROOT),
                    f"trace {lane} {capture_name} custody changed")
            require(capture.get("report_sha256") == sha(directory / report_name)
                    and capture.get("stdout_sha256") == sha(directory / stdout_name)
                    and capture.get("stderr_sha256") == sha(directory / stderr_name),
                    f"trace {lane} {capture_name} artifact hashes changed")
        require((directory / "report.json").read_bytes()
                == (directory / "report-repeat.json").read_bytes(),
                f"trace {lane} repeated report changed")
        require((directory / "stderr").read_bytes()
                == (directory / "stderr-repeat").read_bytes(),
                f"trace {lane} repeated stderr changed")
        require((directory / "report.json").read_bytes()
                == public_report.read_bytes(),
                f"trace {lane} public report differs from oracle")
    # Recompute the structured differential from retained stderr/report files;
    # this is a Python-only replay and does not invoke a trace binary.
    with tempfile.TemporaryDirectory(prefix="litchi-0721-trace-audit-") as directory:
        output = Path(directory) / "trace-analysis.json"
        run_python_replay(
            PACKET / "trace-analyze.py",
            [
                "compare",
                "--baseline-report", str(PACKET / "trace/baseline/report.json"),
                "--candidate-report", str(PACKET / "trace/candidate/report.json"),
                "--baseline-stderr", str(PACKET / "trace/baseline/stderr"),
                "--candidate-stderr", str(PACKET / "trace/candidate/stderr"),
            ],
            output,
        )
        require(output.read_bytes() == path.read_bytes(),
                "trace analysis is not an exact replay")
    return binary_paths


def validate_cleanup(expected_binaries: list[Path], witnesses: list[dict[str, Any]]) -> None:
    path = PACKET / "cleanup.json"
    require(path.is_file(), "cleanup.json is required for the terminal audit")
    value = read_json(path)
    expected_paths = [str(item) for item in OWNED_PATHS]
    require(value.get("owned_paths") == expected_paths
            and value.get("owned_paths_absent") is True,
            "cleanup does not describe the exact owned roots")
    require(all(not item.exists() for item in OWNED_PATHS),
            "an owned target, binary, filesystem, or trace scratch root remains")
    require(isinstance(value.get("binaries"), list),
            "cleanup binary witness list is missing")
    for binary in expected_binaries:
        require(any(item.get("path") == str(binary) for item in value["binaries"]
                    if isinstance(item, dict)),
                f"cleanup omits binary witness: {binary}")
    for binary in expected_binaries:
        matches = [item for item in witnesses if resolve_artifact(item["path"]) == binary]
        require(matches, f"cleanup has no exact binary witness: {binary}")


def run_python_replay(script: Path, args: list[str], output: Path | None = None) -> None:
    command = [sys.executable, "-B", str(script), *args]
    if output is not None:
        command += ["--output", str(output)]
    try:
        subprocess.run(command, cwd=ROOT, check=True, stdout=subprocess.PIPE,
                       stderr=subprocess.PIPE, text=True)
    except subprocess.CalledProcessError as error:
        detail = (error.stderr or error.stdout or "").strip()
        raise AuditError(f"Python replay failed: {' '.join(command)}\n{detail}") from error


def validate_analysis(plan: dict[str, Any], final_label: str | None,
                      final_source: dict[str, str]) -> tuple[dict[str, Any], dict[str, Any]]:
    primary = read_json(PACKET / "analysis.json")
    read_control = read_json(PACKET / "read-controls-analysis.json")
    require(primary.get("schema_version") == 1
            and primary.get("packet") == plan["packet"], "primary analysis identity changed")
    require(read_control.get("schema_version") == 1
            and read_control.get("packet") == "change-0721-docx-read-controls",
            "read-control analysis identity changed")
    require(primary.get("scope", {}).get("children") == 48
            and primary.get("scope", {}).get("native_children") == 32
            and primary.get("scope", {}).get("allocator_children") == 16,
            "primary capture count changed")
    read_scope = read_control.get("scope", {})
    require(read_scope.get("children") == 16
            and read_scope.get("top_level_read_invocations") == 16
            and read_scope.get("pinned_filesystem_internal_children_per_invocation") == {
                "warmup": 100,
                "priming": 200,
                "measured": 200,
                "reported_total": 500,
            }
            and read_scope.get("pinned_filesystem_internal_children_across_stages") == 4000,
            "read-control capture scope changed")
    p_decision = primary.get("decision")
    r_decision = read_control.get("decision")
    require(isinstance(p_decision, dict) and isinstance(r_decision, dict),
            "analysis decision is missing")
    require(p_decision.get("deterministic_output_parity_pass") is True
            and r_decision.get("deterministic_output_parity_pass") is True,
            "deterministic output parity did not pass")
    p_gates = p_decision.get("hard_gates")
    r_gates = r_decision.get("hard_gates")
    require(isinstance(p_gates, list) and len(p_gates) == 64,
            "primary hard-gate count changed")
    require(isinstance(r_gates, list) and len(r_gates) == 16,
            "read-control hard-gate count changed")
    require(all(isinstance(item, dict) and isinstance(item.get("pass"), bool) for item in p_gates),
            "primary hard-gate rows are malformed")
    require(all(isinstance(item, dict) and isinstance(item.get("pass"), bool) for item in r_gates),
            "read-control hard-gate rows are malformed")
    p_all = all(item["pass"] for item in p_gates)
    r_all = all(item["pass"] for item in r_gates)
    require(p_decision.get("primary_hard_gates_pass") == p_all
            and p_decision.get("read_controls_pass") == r_all
            and p_decision.get("all_hard_gates_pass") == (p_all and r_all)
            and p_decision.get("accepted") == (p_all and r_all),
            "primary decision is inconsistent with its hard gates")
    require(r_decision.get("all_hard_gates_pass") == r_all
            and r_decision.get("accepted") == r_all,
            "read-control decision is inconsistent with its hard gates")
    require(primary.get("source", {}).get("final_sha256")
            == (sha(PACKET / "source-final.json") if final_label else None),
            "primary final source binding changed")
    require(read_control.get("scope", {}).get("final_source") == final_label,
            "read-control final source binding changed")
    primary_bindings = primary.get("freeze_bindings")
    require(primary_bindings == {
        "plan_sha256": sha(PACKET / "plan.json"),
        "capture_script_sha256": sha(PACKET / "pilot.py"),
        "analyzer_sha256": sha(PACKET / "analyze-final.py"),
        "analysis_script_sha256": sha(PACKET / "analyze-final.py"),
        "constraints_sha256": sha(PACKET / "constraints.json"),
    }, "primary analysis script or capture binding changed")
    require(read_control.get("plan_sha256") == sha(PACKET / "read-controls-plan.json")
            and read_control.get("capture_script_sha256") == sha(PACKET / "read-controls.py")
            and read_control.get("analysis_script_sha256")
            == sha(PACKET / "read-controls-analyze.py")
            and read_control.get("constraints_sha256") == sha(PACKET / "constraints.json"),
            "read-control analysis script or capture binding changed")
    if final_label is not None:
        retained = final_label == "candidate"
        require(p_decision.get("accepted") == retained,
                "final disposition disagrees with primary/read-control decision")
        require(primary.get("final_disposition") == read_json(PACKET / "disposition.json"),
                "primary disposition binding changed")
        require(read_control.get("verification", {}).get("final_source_matches_checkout") is True
                and read_control.get("verification", {}).get("final_source_label") == final_label,
                "read-control final source witness is missing")
    require(final_source in (read_json(PACKET / "source-baseline.json"),
                             read_json(PACKET / "source-candidate.json")),
            "final source is outside the frozen pair")
    # Replays are Python-only and therefore safe after binary cleanup.  They
    # independently recompute the same bytes rather than trusting analysis.json.
    run_python_replay(PACKET / "analyze-final.py", ["--check"])
    with tempfile.TemporaryDirectory(prefix="litchi-0721-audit-") as directory:
        output = Path(directory) / "read-controls-analysis.json"
        run_python_replay(PACKET / "read-controls-analyze.py", ["analyze"], output)
        require(output.read_bytes() == (PACKET / "read-controls-analysis.json").read_bytes(),
                "read-control analysis is not an exact replay")
    return primary, read_control


def draft_audit(custody: Any, plan: dict[str, Any]) -> None:
    pending: list[str] = []
    load_read_control_plan(plan)
    constraints_sha = validate_constraints()
    baseline, candidate, _final, label, _disposition = validate_sources(custody, plan, draft=True)
    validate_source_snapshots(baseline, candidate, draft=True)
    validate_capture_freeze(plan, draft=True)
    validate_capture_orchestrator(draft=True)
    witnesses = cleanup_witnesses()
    builds = validate_builds(baseline, candidate, witnesses, draft=True)
    if len(builds) != 4:
        pending.append("both baseline and candidate native/allocator build records")
    oracle = validate_oracle(baseline, candidate, witnesses, draft=True)
    if len(oracle) != 2:
        pending.append("baseline and candidate public oracle captures")
    if not (PACKET / "quality-docx.json").is_file():
        pending.append("candidate DOCX quality records")
    if not (PACKET / "analysis.json").is_file():
        pending.append("primary 48-child analysis")
    if not (PACKET / "read-controls-analysis.json").is_file():
        pending.append("read-control 16-child analysis")
    if not (PACKET / "trace-analysis.json").is_file():
        pending.append("trace-analysis.json")
    if not (PACKET / "capture-freeze.json").is_file():
        pending.append("capture-freeze.json")
    if not (PACKET / "capture.json").is_file():
        pending.append("capture.json orchestrator receipt")
    if label is None:
        pending.append("source-final.json and disposition.json")
    if not (PACKET / "cleanup.json").is_file():
        pending.append("terminal cleanup witnesses")
    if constraints_sha:
        print("DRAFT: frozen constraints and source pair validated")
    if pending:
        raise DraftPending(pending)


def validate_negative_checks() -> None:
    value = read_json(PACKET / "negative-checks.json")
    require(value.get("status") == "pass" and value.get("retained_inputs_unchanged") is True,
            "negative checks did not pass without mutating retained inputs")
    checks = value.get("checks", [])
    require(len(checks) == 9 and len({row["name"] for row in checks}) == 9
            and all(row.get("rejected") is True for row in checks),
            "negative check inventory changed")
    require(all(row.get("passed") is True for row in value.get("positive_replays", [])),
            "negative-check positive controls did not pass")
    for raw, item in value["inputs"].items():
        path = ROOT / raw
        require(path.is_file() and path.stat().st_size == item["bytes"]
                and sha(path) == item["sha256"], f"negative-check input changed: {raw}")


def terminal_audit(custody: Any, plan: dict[str, Any]) -> None:
    load_read_control_plan(plan)
    constraints_sha = validate_constraints()
    baseline, candidate, final_source, final_label, disposition = validate_sources(
        custody, plan, draft=False)
    require(final_label is not None and disposition is not None,
            "terminal source disposition is missing")
    validate_source_snapshots(baseline, candidate, draft=False)
    validate_capture_freeze(plan, draft=False)
    validate_capture_orchestrator(draft=False)
    witnesses = cleanup_witnesses()
    builds = validate_builds(baseline, candidate, witnesses, draft=False)
    require(len(builds) == 4, "terminal build inventory is incomplete")
    oracle_binaries = validate_oracle(baseline, candidate, witnesses, draft=False)
    trace_binaries = validate_trace(baseline, candidate, witnesses, draft=False)
    validate_quality(final_label, final_source)
    validate_final_quality(final_source)
    validate_analyzer_corrections()
    validate_negative_checks()
    primary, read_control = validate_analysis(plan, final_label, final_source)
    expected_binaries = [item["binary"] for item in builds.values()]
    expected_binaries += oracle_binaries + trace_binaries
    validate_cleanup(expected_binaries, witnesses)
    require(constraints_sha == sha(PACKET / "constraints.json"),
            "constraints manifest changed during audit")
    _ = (primary, read_control)
    print("PASS: exact 0721 source/build/oracle/trace custody, 48 primary (32 native + 16 allocator) and 16 read-control children, final disposition, replay, and cleanup")


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--draft", action="store_true",
                        help="validate preparation and report pending terminal evidence")
    args = parser.parse_args()
    try:
        plan = load_plan()
        custody = load_custody()
        if args.draft:
            try:
                draft_audit(custody, plan)
            except DraftPending as pending:
                print("DRAFT: terminal evidence remains pending:")
                for item in pending.pending:
                    print(f"  - {item}")
                return 0
            print("DRAFT: preparation is complete; terminal audit is still required")
            return 0
        terminal_audit(custody, plan)
        return 0
    except (AuditError, KeyError, OSError, TypeError, ValueError) as error:
        print(f"audit failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
