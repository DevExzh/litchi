#!/usr/bin/env python3
"""Audit the final 0717 source, build, evidence, and cleanup custody."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
LANES = ("native", "procfs")
ALLOWED_SOURCE_CHANGES = {
    "tools/perf-baseline/Cargo.toml",
    "tools/perf-baseline/src/lib.rs",
    "tools/perf-baseline/src/ordinary_save.rs",
}
SUPPORT_FILES = {
    "tools/perf-baseline/README.md",
    "tools/test_perf_abba_summary.py",
}
QUALITY_NAMES = (
    "fmt", "tests-native", "tests-procfs", "clippy", "rustdoc", "latency-guard",
)
EVIDENCE_NAMES = (
    "crate-boundaries", "claims", "claims-structural", "report", "coverage", "non-iwork",
)
OWNED_PATHS = [
    str(ROOT.parent / "litchi-target-0717"),
    str(ROOT.parent / "litchi-0717-bin"),
    str(ROOT.parent / "litchi-0717-fs"),
]
HEX = set("0123456789abcdef")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        raise AssertionError(f"invalid JSON in {path}: {error}") from error


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing evidence file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check_digest(value: Any, label: str) -> None:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")


def positive_int(value: Any, label: str) -> None:
    require(isinstance(value, int) and not isinstance(value, bool) and value > 0,
            f"{label} is not a positive integer")


def source_census() -> dict[str, str]:
    paths: list[Path] = [ROOT / "Cargo.toml", ROOT / "Cargo.lock"]
    paths.extend(path for path in (ROOT / ".cargo").rglob("*") if path.is_file())
    for folder in ("crates", "tools/perf-baseline"):
        paths.extend(
            path for path in (ROOT / folder).rglob("*")
            if path.is_file() and "target" not in path.parts
            and (path.suffix == ".rs" or path.name in {"Cargo.toml", "Cargo.lock"})
        )
    return {
        str(path.relative_to(ROOT)): sha(path)
        for path in sorted(set(paths))
        if path.is_file() and not path.is_symlink()
    }


def source_bindings() -> tuple[dict[str, str], dict[str, str], str]:
    baseline = read(PACKET / "source-baseline.json")
    source = read(PACKET / "source.json")
    historical = read(PACKET.parent / "change-0716/source.json")
    require(isinstance(baseline, dict) and baseline, "source baseline is invalid")
    require(isinstance(source, dict) and source, "source manifest is invalid")
    require(baseline == historical, "0717 baseline source differs from 0716 source")
    require(source_census() == source, "current checkout does not match source.json")
    changed = {name for name in set(source) | set(baseline)
               if source.get(name) != baseline.get(name)}
    require(changed == ALLOWED_SOURCE_CHANGES,
            f"source changes are not exactly the three harness files: {sorted(changed)}")

    final_source = read(PACKET / "source-final.json")
    require(final_source == source, "source-final.json does not equal source.json")

    preparation = read(PACKET / "source-preparation.json")
    require(preparation.get("source_baseline_sha256") == sha(PACKET / "source-baseline.json"),
            "source preparation baseline binding changed")
    require(preparation.get("source_sha256") == sha(PACKET / "source.json"),
            "source preparation source binding changed")
    require(set(preparation.get("changed_files", []))
            == ALLOWED_SOURCE_CHANGES | SUPPORT_FILES,
            "source preparation changed-file disclosure changed")
    require(preparation.get("production_crate_changes") == [],
            "production crate changes were disclosed")
    patch_digest = preparation.get("patch_sha256")
    if isinstance(patch_digest, str):
        check_digest(patch_digest, "source-preparation.patch_sha256")
        require(patch_digest == sha(PACKET / "source.patch"),
                "source patch binding changed")
    else:
        require(isinstance(patch_digest, dict) and patch_digest,
                "source-preparation.patch_sha256 is invalid")
        for raw, digest in patch_digest.items():
            check_digest(digest, f"source patch {raw}")
            require(sha(PACKET / raw) == digest, f"source patch binding changed: {raw}")
    return baseline, source, sha(PACKET / "source.json")


def validate_support() -> str:
    manifest = read(PACKET / "support-files.json")
    require(isinstance(manifest, dict) and set(manifest) == SUPPORT_FILES,
            "support-files.json does not cover exactly README and Python support")
    for raw, digest in manifest.items():
        check_digest(digest, f"support file {raw}")
        target = ROOT / raw
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"support file changed: {raw}")
    return sha(PACKET / "support-files.json")


def resolve_packet_or_root(raw: str, base: Path = PACKET) -> Path:
    path = Path(raw)
    if path.is_absolute():
        return path
    if raw.startswith(("docs/", "tools/", "crates/")):
        return ROOT / raw
    return base / raw


def validate_hash_map(filename: str, *, base: Path = PACKET) -> dict[str, str]:
    value = read(PACKET / filename)
    require(isinstance(value, dict) and value, f"{filename} is invalid")
    for raw, digest in value.items():
        check_digest(digest, f"{filename}:{raw}")
        target = resolve_packet_or_root(raw, base)
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"{filename} binding changed: {raw}")
    return value


def validate_freezes() -> None:
    for filename in (
        "constraints.json", "helper-freeze.json", "capture-freeze.json", "analysis-freeze.json",
    ):
        validate_hash_map(filename)


def validate_protocol() -> None:
    review = read(PACKET / "protocol-review.json")
    require(isinstance(review, dict), "protocol review is invalid")
    require(str(review.get("decision", "")).startswith("proceed"),
            "protocol review did not authorize diagnostic captures")
    resolutions = " ".join(str(item) for item in review.get("resolutions", []))
    text = f"{resolutions} {review.get('capture_rule', '')}".lower()
    require("same-process" in text or "same process" in text,
            "process scope is not disclosed as same-process")
    require("counter-delta" in text or "counter delta" in text,
            "control scope does not disclose counter-delta-only overhead")
    require("no native versus instrumented speedup" in text
            or "no latency" in text,
            "protocol review permits an invalid lane latency claim")


def expected_build_command(lane: str) -> list[str]:
    command = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", "litchi-perf-baseline",
        "--target-dir", str(ROOT.parent / "litchi-target-0717"), "-j", "2",
    ]
    if lane == "procfs":
        command += ["--features", "ordinary-save-process-metrics"]
    return command


def walk_witnesses() -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for filename in ("cleanup.json", "cleanup-witness.json", "cleanup-receipt.json"):
        path = PACKET / filename
        if not path.is_file():
            continue
        value = read(path)

        def walk(item: Any) -> None:
            if isinstance(item, dict):
                raw_path = item.get("path")
                digest = item.get("sha256", item.get("binary_sha256"))
                size = item.get("bytes", item.get("binary_bytes"))
                if isinstance(raw_path, str) and isinstance(digest, str):
                    check_digest(digest, f"{filename}:{raw_path}")
                    if size is not None:
                        positive_int(size, f"{filename}:{raw_path}.bytes")
                    result.append({"path": raw_path, "sha256": digest, "bytes": size})
                for child in item.values():
                    walk(child)
            elif isinstance(item, list):
                for child in item:
                    walk(child)

        walk(value)
    return result


def validate_binary(binary: dict[str, Any], witnesses: list[dict[str, Any]]) -> None:
    require(isinstance(binary, dict), "build binary record is missing")
    raw_path, digest, size = binary.get("path"), binary.get("sha256"), binary.get("bytes")
    require(isinstance(raw_path, str) and Path(raw_path).is_absolute(),
            "build binary path is not absolute")
    check_digest(digest, f"build binary {raw_path}")
    positive_int(size, f"build binary {raw_path}.bytes")
    path = Path(raw_path).resolve()
    if path.is_file() and not path.is_symlink():
        require(sha(path) == digest and path.stat().st_size == size,
                f"live binary identity changed: {path}")
        return
    matches = []
    for witness in witnesses:
        candidate = resolve_packet_or_root(witness["path"]).resolve()
        if candidate == path and witness["sha256"] == digest and witness.get("bytes") == size:
            matches.append(witness)
    require(len(matches) == 1, f"missing exact cleanup witness for binary: {path}")


def validate_builds(source_digest: str) -> tuple[dict[str, Any], str, list[dict[str, Any]]]:
    builds = read(PACKET / "builds.json")
    require(isinstance(builds, dict) and set(builds) == set(LANES),
            "builds.json must contain exactly native and procfs")
    witnesses = walk_witnesses()
    for lane in LANES:
        row = builds[lane]
        require(isinstance(row, dict) and row.get("exit_code") == 0,
                f"{lane} build did not pass")
        require(row.get("command") == expected_build_command(lane),
                f"{lane} build command changed")
        require(row.get("source_sha256") == source_digest,
                f"{lane} build source binding changed")
        log = row.get("log")
        require(log == f"build-{lane}.log", f"{lane} build log name changed")
        check_digest(row.get("log_sha256"), f"{lane} build log")
        require(sha(PACKET / log) == row["log_sha256"], f"{lane} build log changed")
        binary = row.get("binary")
        require(isinstance(binary, dict), f"{lane} binary record is missing")
        require(binary.get("path") == str(ROOT.parent / "litchi-0717-bin" / lane),
                f"{lane} final binary path changed")
        validate_binary(binary, witnesses)
    return builds, sha(PACKET / "builds.json"), witnesses


def validate_initial_archive(final_builds: dict[str, Any]) -> None:
    archive = PACKET / "build-attempts/01"
    if not archive.is_dir():
        return
    for path in archive.rglob("*"):
        require(not path.is_symlink(), f"initial build archive contains symlink: {path}")
    removal_path = archive / "archive-removal.json"
    if not removal_path.is_file():
        return
    removal = read(removal_path)
    require(removal.get("binaries_removed") is True,
            "initial binary archive does not record removal")
    archived = read(archive / "builds.json")
    archived_source = read(archive / "source.json")
    archived_prep = read(archive / "source-preparation.json")
    require(isinstance(archived, dict) and set(archived) == set(LANES),
            "initial archive build lanes changed")
    require(isinstance(archived_source, dict) and archived_source,
            "initial archive source is invalid")
    require(archived_prep.get("source_sha256") == sha(archive / "source.json"),
            "initial archive source binding changed")
    require(archived_prep.get("source_baseline_sha256")
            == sha(PACKET.parent / "change-0716/source.json"),
            "initial archive baseline binding changed")
    patch_digest = archived_prep.get("patch_sha256")
    if isinstance(patch_digest, str):
        require(patch_digest == sha(archive / "source.patch"),
                "initial archive patch binding changed")
    records = removal.get("binaries")
    require(isinstance(records, list) and len(records) == len(LANES),
            "initial archive does not retain both binary witnesses")
    by_path = {item.get("path"): item for item in records if isinstance(item, dict)}
    require(len(by_path) == len(LANES), "initial binary witnesses are not unique")
    for lane in LANES:
        row = archived[lane]
        require(row.get("exit_code") == 0, f"initial {lane} build did not pass")
        require(row.get("source_sha256") == sha(archive / "source.json"),
                f"initial {lane} source binding changed")
        log = archive / row.get("log", "")
        require(row.get("log") == f"build-{lane}.log" and sha(log) == row.get("log_sha256"),
                f"initial {lane} build log changed")
        binary = row.get("binary")
        witness = by_path.get(binary.get("path") if isinstance(binary, dict) else None)
        require(witness == binary, f"initial {lane} binary witness changed")
        # A feature-gated wording change may leave the native binary identical.
        # Source-bound build records, not unequal executable hashes, prove custody.


def quality_commands() -> dict[str, list[str]]:
    base = [
        "--release", "--locked", "--manifest-path", "tools/perf-baseline/Cargo.toml",
        "--target-dir", str(ROOT.parent / "litchi-target-0717"), "-j", "2",
    ]
    return {
        "fmt": ["cargo", "fmt", "--manifest-path", "tools/perf-baseline/Cargo.toml",
                "--", "--check"],
        "tests-native": ["cargo", "test", *base, "--lib"],
        "tests-procfs": ["cargo", "test", *base, "--lib", "--features",
                         "ordinary-save-process-metrics"],
        "clippy": ["cargo", "clippy", *base, "--all-targets", "--features",
                    "ordinary-save-process-metrics,allocator-metrics", "--", "-D", "warnings"],
        "rustdoc": ["cargo", "doc", *base, "--no-deps", "--lib", "--features",
                     "ordinary-save-process-metrics"],
        "latency-guard": ["python3", "-B", "-m", "unittest", "discover", "-s", "tools",
                          "-p", "test_perf_abba_summary.py"],
    }


def validate_quality(source_digest: str, support_digest: str) -> None:
    rows = read(PACKET / "quality.json")
    require(isinstance(rows, list) and [row.get("name") for row in rows] == list(QUALITY_NAMES),
            "quality.json does not contain the exact six fresh rows")
    expected = quality_commands()
    for row in rows:
        name = row["name"]
        require(row.get("exit_code") == 0, f"quality row failed: {name}")
        require(row.get("command") == expected[name], f"quality command changed: {name}")
        require(row.get("source_manifest_sha256") == source_digest,
                f"quality source binding changed: {name}")
        require(row.get("support_manifest_sha256") == support_digest,
                f"quality support binding changed: {name}")
        require(row.get("log") == f"quality-{name}.log", f"quality log name changed: {name}")
        check_digest(row.get("log_sha256"), f"quality log {name}")
        require(sha(PACKET / row["log"]) == row["log_sha256"], f"quality log changed: {name}")


def validate_evidence(final_source_digest: str) -> None:
    rows = read(PACKET / "evidence/results.json")
    require(isinstance(rows, list)
            and [row.get("name") for row in rows] == list(EVIDENCE_NAMES),
            "evidence does not contain the exact six repository gates")
    for row in rows:
        name = row["name"]
        require(row.get("exit_code") == 0, f"evidence gate failed: {name}")
        require(row.get("source_manifest_sha256") == final_source_digest,
                f"evidence source binding changed: {name}")
        log = PACKET / "evidence" / f"{name}.log"
        check_digest(row.get("log_sha256"), f"evidence log {name}")
        require(sha(log) == row["log_sha256"], f"evidence log changed: {name}")


def validate_reused_production(baseline: dict[str, str]) -> None:
    record = read(PACKET / "reused-production-verification.json")
    require(isinstance(record, dict), "reused production verification is invalid")
    prior = ROOT / record.get("packet", "")
    require(prior.is_dir() and not prior.is_symlink(), "reused production packet is missing")
    files = record.get("files")
    require(isinstance(files, dict) and files, "reused production file map is missing")
    for raw, digest in files.items():
        check_digest(digest, f"reused production {raw}")
        target = prior / raw
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"reused production artifact changed: {raw}")

    for raw in ("source-final.json", "quality-source.json", "source-candidate.json"):
        target = prior / raw
        if target.is_file():
            require(read(target) == baseline,
                    f"reused production source is not the unchanged baseline: {raw}")
    quality = read(prior / "quality.json")
    require(isinstance(quality, list) and quality, "reused production quality is invalid")
    for row in quality:
        require(row.get("exit_code") == 0, "reused production quality row failed")
        log = prior / row.get("log", f"{row.get('name')}.log")
        require(log.is_file() and sha(log) == row.get("log_sha256"),
                f"reused production quality log changed: {log}")
    evidence = read(prior / "evidence/results.json")
    require(isinstance(evidence, list) and len(evidence) == len(EVIDENCE_NAMES),
            "reused production evidence does not retain six gates")
    for row in evidence:
        require(row.get("exit_code") == 0, "reused production evidence gate failed")
        log = prior / "evidence" / f"{row.get('name')}.log"
        require(log.is_file() and sha(log) == row.get("log_sha256"),
                f"reused production evidence log changed: {log}")


def validate_negative_checks() -> None:
    checks = read(PACKET / "negative-checks.json")
    require(checks.get("status") == "pass"
            and checks.get("retained_inputs_unchanged") is True
            and checks.get("positive_exact_replay") is True,
            "negative checks did not pass exact replay")
    rows = checks.get("checks")
    require(isinstance(rows, list) and len(rows) >= 6 and all(row.get("rejected") for row in rows),
            "fewer than six rejecting negative checks were retained")
    require(checks.get("analysis_sha256") == sha(PACKET / "analysis.json"),
            "negative checks bind the wrong analysis")
    require(checks.get("analyzer_sha256") == sha(PACKET / "analyze.py"),
            "negative checks bind the wrong analyzer")


def validate_final_gate() -> None:
    gate = read(PACKET / "final-report-gate.json")
    require(gate.get("exit_code") == 0, "final report gate failed")
    require(gate.get("log_sha256") == sha(PACKET / "final-report-gate.log"),
            "final report gate log changed")
    docs = gate.get("docs")
    require(isinstance(docs, dict) and docs, "final report gate has no documentation bindings")
    for raw, digest in docs.items():
        check_digest(digest, f"final documentation {raw}")
        target = resolve_packet_or_root(raw)
        require(target.is_file() and not target.is_symlink() and sha(target) == digest,
                f"final documentation changed: {raw}")
    for field, filename in (
        ("analysis_sha256", "analysis.json"),
        ("negative_checks_sha256", "negative-checks.json"),
        ("artifact_manifest_sha256", "artifact-manifest.json"),
    ):
        if field in gate:
            require(gate[field] == sha(PACKET / filename), f"final gate binding changed: {field}")


def validate_artifact_manifest() -> None:
    path = PACKET / "artifact-manifest.json"
    if not path.is_file():
        return
    value = read(path)
    files = value.get("files") if isinstance(value, dict) else None
    require(isinstance(files, dict) and files, "artifact manifest is invalid")
    for raw, record in files.items():
        require(isinstance(record, dict), f"artifact manifest row is invalid: {raw}")
        target = PACKET / raw
        require(target.is_file() and not target.is_symlink(), f"artifact is missing: {raw}")
        require(record.get("bytes") == target.stat().st_size
                and record.get("sha256") == sha(target),
                f"artifact manifest binding changed: {raw}")


def validate_cleanup(builds: dict[str, Any], witnesses: list[dict[str, Any]]) -> None:
    cleanup = read(PACKET / "cleanup.json")
    require(cleanup.get("owned_paths") == OWNED_PATHS
            and cleanup.get("owned_paths_absent") is True,
            "cleanup roots are not the exact three owned paths")
    require(all(not Path(path).exists() for path in OWNED_PATHS),
            "an owned cleanup root remains on disk")
    expected = [builds[lane]["binary"] for lane in LANES]
    actual = cleanup.get("binaries")
    require(actual == expected, "cleanup does not retain both final binary witnesses")
    for binary in expected:
        validate_binary(binary, witnesses)


def validate_analysis_reconciliation() -> None:
    record = read(PACKET / "analysis-reconciliation.json")
    prior_path = PACKET / "analysis-attempts/02/analysis.json"
    require(record.get("prior_analysis_sha256") == sha(prior_path), "prior analysis binding changed")
    require(record.get("final_analysis_sha256") == sha(PACKET / "analysis.json"), "final analysis binding changed")
    prior, current = read(prior_path), read(PACKET / "analysis.json")
    fields = []
    for lane in LANES:
        require(prior["builds"][lane]["binary_custody"] == "live-binary", "prior custody mode changed")
        require(current["builds"][lane]["binary_custody"] == "exact-binary-identity-validated", "final custody mode changed")
        prior["builds"][lane]["binary_custody"] = current["builds"][lane]["binary_custody"]
        fields.append(f"builds.{lane}.binary_custody")
    prior["freeze_bindings"]["analyzer_sha256"] = current["freeze_bindings"]["analyzer_sha256"]
    fields.append("freeze_bindings.analyzer_sha256")
    require(record.get("changed_fields") == fields and prior == current,
            "analysis statistics or capture bindings changed during custody correction")
    require(record.get("all_statistics_and_capture_bindings_unchanged") is True,
            "analysis reconciliation did not pass")


def main() -> None:
    baseline, source, source_digest = source_bindings()
    support_digest = validate_support()
    validate_freezes()
    validate_protocol()
    builds, builds_digest, witnesses = validate_builds(source_digest)
    validate_initial_archive(builds)
    validate_quality(source_digest, support_digest)
    validate_evidence(sha(PACKET / "source-final.json"))
    validate_reused_production(baseline)
    validate_negative_checks()
    validate_analysis_reconciliation()
    validate_final_gate()
    validate_artifact_manifest()
    validate_cleanup(builds, witnesses)
    subprocess.run(
        ["python3", "-B", str(PACKET / "analyze.py"), "--check"],
        cwd=ROOT,
        check=True,
    )
    require(builds_digest == sha(PACKET / "builds.json"), "build manifest changed during audit")
    print("PASS: final 0717 source, archive, builds, quality, evidence, refusals, docs, and cleanup")


if __name__ == "__main__":
    main()
