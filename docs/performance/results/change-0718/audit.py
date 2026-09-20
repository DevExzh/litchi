#!/usr/bin/env python3
"""Final custody audit for the unchanged-source 0718 diagnostic packet."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import subprocess
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / "litchi-target-0718"
BIN = ROOT.parent / "litchi-0718-bin"
FS = ROOT.parent / "litchi-0718-fs"
HEX = set("0123456789abcdef")


def need(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read(path: Path) -> Any:
    need(path.is_file() and not path.is_symlink(), f"missing {path}")
    try:
        return json.loads(path.read_text())
    except json.JSONDecodeError as error:
        raise AssertionError(f"invalid JSON: {path}: {error}") from error


def sha(path: Path) -> str:
    need(path.is_file() and not path.is_symlink(), f"missing {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def digest(value: Any, label: str) -> None:
    need(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
         f"{label} is not a SHA-256 digest")


def resolve(raw: str, base: Path = P) -> Path:
    path = Path(raw)
    if path.is_absolute():
        return path
    if raw.startswith(("docs/", "tools/", "crates/", "Cargo.", ".cargo/")):
        return ROOT / raw
    return base / raw


def hash_map(path: Path, base: Path = P) -> dict[str, str]:
    value = read(path)
    need(isinstance(value, dict) and value, f"invalid hash map: {path}")
    for raw, value_hash in value.items():
        digest(value_hash, f"{path}:{raw}")
        target = resolve(raw, base)
        need(sha(target) == value_hash, f"hash binding changed: {path}:{raw}")
    return value


def census() -> dict[str, str]:
    paths = [ROOT / "Cargo.toml", ROOT / "Cargo.lock"]
    paths += [p for p in (ROOT / ".cargo").rglob("*") if p.is_file()]
    for folder in ("crates", "tools/perf-baseline"):
        paths += [p for p in (ROOT / folder).rglob("*")
                  if p.is_file() and "target" not in p.parts
                  and (p.suffix == ".rs" or p.name in {"Cargo.toml", "Cargo.lock"})]
    return {str(p.relative_to(ROOT)): sha(p) for p in sorted(set(paths))
            if p.is_file() and not p.is_symlink()}


def source_and_freezes() -> str:
    source = read(P / "source.json")
    prior = read(P.parent / "change-0717/source.json")
    need(source == prior == census(), "current source does not equal unchanged 0717 source")
    for name in ("constraints.json", "helper-freeze.json", "capture-freeze.json",
                 "analysis-freeze.json"):
        hash_map(P / name)
    return sha(P / "source.json")


def binary_record(record: dict[str, Any], cleanup: dict[str, Any] | None = None) -> None:
    path, value_hash, size = record.get("path"), record.get("sha256"), record.get("bytes")
    need(isinstance(path, str) and Path(path).is_absolute(), "binary path is not absolute")
    digest(value_hash, f"binary {path}")
    need(isinstance(size, int) and size > 0, f"binary size is invalid: {path}")
    live = Path(path)
    if live.is_file() and not live.is_symlink():
        need(sha(live) == value_hash and live.stat().st_size == size,
             f"live binary identity changed: {path}")
        return
    witnesses = (cleanup or {}).get("binaries", [])
    need(record in witnesses, f"missing cleanup binary witness: {path}")


def build_and_cleanup(source_hash: str) -> dict[str, Any]:
    build = read(P / "build.json")
    command = ["cargo", "build", "--release", "--locked", "--manifest-path",
               "tools/perf-baseline/Cargo.toml", "--bin", "litchi-perf-baseline",
               "--target-dir", str(TARGET), "-j", "2", "--features",
               "ordinary-save-process-metrics"]
    need(build.get("exit_code") == 0 and build.get("command") == command,
         "0718 procfs build record changed")
    need(build.get("source_sha256") == source_hash, "build source binding changed")
    need(build.get("log") in (None, "build.log"), "build log name changed")
    log = P / "build.log"
    digest(build.get("log_sha256"), "build log")
    need(sha(log) == build["log_sha256"], "build log changed")
    cleanup = read(P / "cleanup.json")
    owned = [str(TARGET), str(BIN), str(FS)]
    need(cleanup.get("owned_paths") == owned and cleanup.get("owned_paths_absent") is True,
         "cleanup roots changed")
    need(all(not Path(path).exists() for path in owned), "owned cleanup root remains")
    binary = build.get("binary")
    need(isinstance(binary, dict) and binary.get("path") == str(BIN / "procfs"),
         "build binary record changed")
    binary_record(binary, cleanup)
    need(cleanup.get("binaries") == [binary], "cleanup does not retain build binary witness")
    return build


def rows_and_logs(packet: Path, name: str, source_hash: str,
                  evidence: bool = False) -> None:
    rows = read(packet / name)
    need(isinstance(rows, list) and rows, f"invalid reused {name}")
    for row in rows:
        need(isinstance(row, dict) and row.get("exit_code") == 0,
             f"reused {name} contains a failed row")
        if row.get("source_manifest_sha256") is not None:
            need(row["source_manifest_sha256"] == source_hash,
                 f"reused {name} source binding changed")
        relative = (Path("evidence") / f"{row.get('name')}.log" if evidence
                    else Path(row.get("log", "")))
        log = packet / relative
        value_hash = row.get("log_sha256")
        digest(value_hash, f"{name} log")
        need(sha(log) == value_hash, f"reused {name} log changed: {log}")


def reused_verification(source_hash: str) -> None:
    record = read(P / "reused-verification.json")
    packet = ROOT / record.get("packet", "")
    need(record.get("packet") == "docs/performance/results/change-0717"
         and packet.is_dir() and not packet.is_symlink(), "0717 reused packet missing")
    files = record.get("files")
    need(isinstance(files, dict) and files, "0717 reused file map missing")
    for raw, value_hash in files.items():
        digest(value_hash, f"reused file {raw}")
        need(sha(packet / raw) == value_hash, f"reused file changed: {raw}")
    for name in ("source.json", "source-final.json", "quality-source.json"):
        target = packet / name
        if target.is_file():
            need(read(target) == read(P / "source.json"), f"reused source changed: {name}")
    support = packet / "support-files.json"
    support_map = hash_map(support)
    for raw, value_hash in support_map.items():
        digest(value_hash, f"support file {raw}")
        need(sha(ROOT / raw) == value_hash, f"support file changed: {raw}")
    rows_and_logs(packet, "quality.json", source_hash)
    rows_and_logs(packet, "evidence/results.json", source_hash, evidence=True)

    production = read(packet / "reused-production-verification.json")
    prod_packet = ROOT / production.get("packet", "")
    need(production.get("packet") == "docs/performance/results/change-0713"
         and prod_packet.is_dir() and production.get("tests_passed", 0) > 0
         and production.get("doctests_passed", 0) > 0
         and production.get("clippy_scope"), "0713 production disclosure is incomplete")
    prod_files = production.get("files")
    need(isinstance(prod_files, dict) and prod_files, "0713 production file map is missing")
    for raw, value_hash in prod_files.items():
        digest(value_hash, f"production file {raw}")
        need(sha(prod_packet / raw) == value_hash, f"production file changed: {raw}")
    rows_and_logs(prod_packet, "quality.json", sha(prod_packet / "quality-source.json"))
    rows_and_logs(prod_packet, "evidence/results.json", sha(prod_packet / "source-final.json"), evidence=True)


def motivation_and_mechanism() -> None:
    motivation = read(P / "motivation.json")
    base = ROOT / motivation.get("packet", "") if motivation.get("packet") else P
    for raw, value_hash in motivation.get("files", {}).items():
        digest(value_hash, f"motivation file {raw}")
        need(sha(resolve(raw, base)) == value_hash, f"motivation binding changed: {raw}")
    mechanism = read(P / "mechanism-review.json")
    for raw, value_hash in mechanism.get("source_files", {}).items():
        digest(value_hash, f"mechanism source {raw}")
        need(sha(ROOT / raw) == value_hash, f"mechanism source changed: {raw}")


def negative_and_report() -> None:
    checks = read(P / "negative-checks.json")
    rows = checks.get("checks")
    need(checks.get("status") == "pass" and checks.get("retained_inputs_unchanged") is True
         and checks.get("positive_exact_replay") is True
         and isinstance(rows, list) and len(rows) >= 6
         and all(row.get("rejected") is True for row in rows),
         "negative checks did not prove retained exact replay")
    need(checks.get("analysis_sha256") == sha(P / "analysis.json")
         and checks.get("analyzer_sha256") == sha(P / "analyze.py"),
         "negative checks bind the wrong analysis inputs")
    for key, filename in [("attribution_sha256", "attribution.py"),
                          ("brk_analysis_sha256", "brk-analysis.json"),
                          ("brk_analyzer_sha256", "brk-analysis.py")]:
        need(checks.get(key) == sha(P / filename), f"negative check binding changed: {key}")
    gate = read(P / "final-report-gate.json")
    need(gate.get("exit_code") == 0 and gate.get("log_sha256") == sha(P / "final-report-gate.log"),
         "final report gate failed")
    docs = gate.get("docs")
    need(isinstance(docs, dict) and docs, "final report gate has no docs")
    for raw, value_hash in docs.items():
        digest(value_hash, f"report doc {raw}")
        need(sha(resolve(raw)) == value_hash, f"report doc changed: {raw}")


def main() -> None:
    source_hash = source_and_freezes()
    build_and_cleanup(source_hash)
    reused_verification(source_hash)
    motivation_and_mechanism()
    negative_and_report()
    subprocess.run(["python3", "-B", str(P / "analyze.py"), "--check"],
                   cwd=ROOT, check=True)
    subprocess.run(["python3", "-B", str(P / "brk-analysis.py"), "--check"],
                   cwd=ROOT, check=True)
    print("PASS: 0718 source, freezes, build, reused verification, custody, refusals, and report gate")


if __name__ == "__main__":
    main()
