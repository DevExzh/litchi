#!/usr/bin/env python3
"""Independently replay the sealed 0730 DOC evidence.

This file intentionally does not import ``analyze.py``.  It rechecks custody,
matrix order, every public oracle and every named negative control, then
recomputes quantiles, pair deltas, and allocator-field deltas before comparing
the result with ``analysis.json``.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
import subprocess
from pathlib import Path
from typing import Any


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
CAPTURES = PACKET / "captures"
BEFORE = PACKET / "before"
CASES = ("docfloat", "docnohf")
NATIVE_SLOTS = ("baseline", "baseline", "baseline", "candidate", "candidate", "baseline")
ALLOCATION_SLOTS = ("baseline", "candidate", "candidate", "baseline")
META_FIELDS = [
    "entry_type", "name_utf16", "clsid", "bytes", "start_sector", "is_minifat",
    "raw_left_sibling", "raw_right_sibling", "raw_child", "raw_color",
    "raw_state_bits", "raw_creation_time", "raw_modification_time",
]
OWNERSHIP = (
    "Each allocation region reports boundary-relative allocations. The opened "
    "Editor is retained from open through replace_and_validate; the staged "
    "Editor is retained until finish; the returned Vec is retained across "
    "finish. Format output is also retained outside its region for validation. "
    "retained_bytes is live ownership at the region boundary, not RSS."
)
STAT_FIELDS = ("p50", "mean", "p95", "p99", "maximum")
ALLOC_FIELDS = ("allocated_bytes", "deallocated_bytes", "allocation_calls",
                "peak_live_bytes", "retained_bytes")
HEX = set("0123456789abcdef")
ANCESTOR_ORACLE_SHA = "9d1baa0a4978aa9f634545a09d4a594dfb2cfa3785c1895576680045af8b44f9"
LOCAL_ORACLE_SHA = "f870329cc076be509c2066e8ab654b8ce24ab63c3737c110710697a13eb0b4ea"


class AuditError(Exception):
    pass


def need(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def read(path: Path) -> Any:
    need(path.is_file() and not path.is_symlink(), f"missing file: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def sha(path: Path) -> str:
    need(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def digest(value: Any, label: str) -> str:
    need(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
         f"{label}: invalid SHA-256")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    need(isinstance(value, int) and not isinstance(value, bool) and value >= minimum,
         f"{label}: invalid integer")
    return value


def plan() -> dict[str, Any]:
    value = read(PACKET / "plan.json")
    expected = {
        "cpu": 12, "samples": 50, "warmups": 3, "cycles": 3,
        "cases": list(CASES), "native_order": list(NATIVE_SLOTS),
        "native_processes": 36, "allocation_order": list(ALLOCATION_SLOTS),
        "allocation_processes": 8, "before_native_processes": 2,
        "before_allocation_processes": 2,
    }
    need(isinstance(value, dict), "plan is not an object")
    for key, expected_value in expected.items():
        need(value.get(key) == expected_value, f"plan changed: {key}")
    return value


def cases() -> dict[str, dict[str, Any]]:
    rows = read(PACKET / "cases.json")
    need(isinstance(rows, list) and [row.get("case") for row in rows] == list(CASES),
         "case order changed")
    expected = {
        "docfloat": ("doc", "test-data/ole/doc/FloatingPictures.doc", 335360,
                     "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d"),
        "docnohf": ("doc", "test-data/ole/doc/NoHeadFoot.doc", 26112,
                    "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48"),
    }
    result = {}
    for row in rows:
        case = row.get("case")
        need(case in expected and case not in result, f"invalid case: {case!r}")
        fmt, path, size, source_hash = expected[case]
        fixture = ROOT / path
        need(row.get("format") == fmt and row.get("path") == path,
             f"{case}: fixture metadata changed")
        need(fixture.is_file() and fixture.stat().st_size == size and sha(fixture) == source_hash,
             f"{case}: fixture custody changed")
        need(row.get("bytes") == size and row.get("sha256") == source_hash,
             f"{case}: fixture receipt changed")
        result[case] = row
    need(set(result) == set(CASES), "case set changed")
    return result


def contract(case_rows: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    ancestor_path = ROOT / "docs/performance/results/change-0728/oracle-contract.json"
    need(sha(ancestor_path) == ANCESTOR_ORACLE_SHA, "sealed ancestor contract changed")
    need(sha(PACKET / "oracle-contract.json") == LOCAL_ORACLE_SHA,
         "0730 oracle contract copy changed")
    ancestor = read(ancestor_path)
    value = read(PACKET / "oracle-contract.json")
    need(isinstance(value, dict) and set(value) == set(case_rows), "contract case set changed")
    need(isinstance(ancestor, dict) and set(CASES) <= set(ancestor),
         "sealed ancestor contract is incomplete")
    required = {"directory_metadata_fields", "allocation_ownership_contract",
                "semantic_witness", "control_names", "identity"}
    for case, row in value.items():
        need(isinstance(row, dict) and set(row) == required, f"{case}: contract shape changed")
        need(row["directory_metadata_fields"] == META_FIELDS, f"{case}: metadata fields changed")
        need(row["allocation_ownership_contract"] == OWNERSHIP, f"{case}: ownership changed")
        need(row == ancestor[case], f"{case}: contract differs from sealed witness")
        need(isinstance(row["semantic_witness"], dict) and isinstance(row["control_names"], list)
             and row["control_names"] and isinstance(row["identity"], dict),
             f"{case}: contract incomplete")
    return value


def safe_relative(value: Any, label: str) -> str:
    need(isinstance(value, str) and value and not Path(value).is_absolute()
         and ".." not in Path(value).parts, f"{label}: unsafe path")
    return value


def custody(mapping: Any, base: Path, label: str) -> None:
    need(isinstance(mapping, dict) and mapping, f"{label}: empty custody")
    for raw, expected in mapping.items():
        relative = safe_relative(raw, f"{label} path")
        expected = digest(expected, f"{label} {relative}")
        target = base / relative
        if target.is_file() and not target.is_symlink() and sha(target) == expected:
            continue
        # Source guards can freeze a production source under an immutable
        # candidate-source or baseline-source archive after restoring the
        # rejected candidate.  The receipt still names the repository path.
        alternatives = (PACKET / "candidate-source" / relative,
                        PACKET / "baseline-source" / relative)
        need(any(p.is_file() and not p.is_symlink() and sha(p) == expected
                 for p in alternatives), f"{label}: missing archived {relative}")


def builds() -> dict[str, dict[str, Any]]:
    result = {}
    for variant in ("baseline", "candidate"):
        value = read(PACKET / f"{variant}-builds.json")
        need(isinstance(value, dict), f"{variant}: build receipt is not an object")
        receipt = safe_relative(value.get("receipt"), f"{variant}: receipt")
        log = read(PACKET / receipt)
        need(log.get("exit_code") == 0, f"{variant}: build failed")
        for key in ("source_sha256", "probe_sha256", "binaries"):
            need(value.get(key) == log.get(key), f"{variant}: {key} differs from build log")
        custody(value["source_sha256"], ROOT, f"{variant}: source")
        custody(value["probe_sha256"], PACKET, f"{variant}: probe")
        binaries = value["binaries"]
        need(isinstance(binaries, list) and len(binaries) == 2, f"{variant}: binary count")
        names = {Path(row.get("path", "")).name for row in binaries}
        need(names == {f"{variant}-ole_format_save_probe",
                       f"{variant}-ole_format_save_probe_alloc"},
             f"{variant}: binary names changed")
        for row in binaries:
            need(isinstance(row, dict) and isinstance(row.get("path"), str)
                 and Path(row["path"]).is_absolute(), f"{variant}: binary row invalid")
            integer(row.get("bytes"), f"{variant}: binary bytes", 1)
            digest(row.get("sha256"), f"{variant}: binary hash")
        result[variant] = value
    return result


def retention_build() -> dict[str, Any]:
    value = read(PACKET / "retention-builds.json")
    need(isinstance(value, dict), "retention build receipt is not an object")
    need(value.get("exit_code") == 0, "retention build did not pass")
    need(value.get("candidate_builds_sha256") == sha(PACKET / "candidate-builds.json"),
         "retention build is not bound to candidate builds")
    custody(value.get("probe_sha256"), PACKET, "retention: probe")
    binaries = value.get("binaries")
    need(isinstance(binaries, list) and len(binaries) == 2, "retention binary count")
    need({Path(row.get("path", "")).name for row in binaries}
         == {"doc_retention_probe", "doc_retention_probe_alloc"},
         "retention binary names changed")
    for row in binaries:
        need(isinstance(row, dict) and isinstance(row.get("path"), str)
             and Path(row["path"]).is_absolute(), "retention binary path invalid")
        integer(row.get("bytes"), "retention binary bytes", 1)
        digest(row.get("sha256"), "retention binary hash")
    return value


def binary_witness(build_receipts: dict[str, dict[str, Any]], retention: dict[str, Any]) -> None:
    expected = [row for value in build_receipts.values() for row in value["binaries"]]
    expected.extend(retention["binaries"])
    cleanup_path = PACKET / "cleanup.json"
    cleanup = read(cleanup_path) if cleanup_path.is_file() and not cleanup_path.is_symlink() else None
    witnesses = cleanup.get("identities", []) if isinstance(cleanup, dict) else []
    need(isinstance(witnesses, list), "cleanup identities are not a list")
    sort_key = lambda row: json.dumps(row, sort_keys=True)
    missing = []
    for row in expected:
        path = Path(row["path"])
        if path.exists():
            need(path.is_file() and not path.is_symlink() and path.stat().st_size == row["bytes"]
                 and sha(path) == row["sha256"], f"live binary changed: {path}")
        else:
            need(not path.is_symlink(), f"binary path is a symlink: {path}")
            missing.append(row)
    if missing:
        need(isinstance(cleanup, dict) and cleanup.get("removed") is True,
             "missing binaries have no completed cleanup receipt")
        need(sorted(witnesses, key=sort_key) == sorted(expected, key=sort_key),
             "cleanup identities are not the exact binary witness")
    elif witnesses:
        need(sorted(witnesses, key=sort_key) == sorted(expected, key=sort_key),
             "cleanup identities are not the exact binary witness")


def resolve(raw: str) -> Path:
    path = Path(raw)
    if path.is_absolute():
        return path
    if raw.startswith("docs/") or raw.startswith("crates/") or raw in ("Cargo.toml", "Cargo.lock"):
        return ROOT / raw
    return PACKET / raw


def freeze(build_receipts: dict[str, dict[str, Any]], retention: dict[str, Any]) -> None:
    value = read(PACKET / "freeze.json")
    need(isinstance(value, dict), "freeze is not an object")
    bindings = value.get("bindings", value)
    need(isinstance(bindings, dict) and bindings, "freeze bindings are empty")
    binaries = [row for receipt in build_receipts.values() for row in receipt["binaries"]]
    binaries.extend(retention["binaries"])
    for raw, expected in bindings.items():
        need(isinstance(raw, str), "freeze path is not a string")
        expected = digest(expected, f"freeze binding {raw}")
        path = resolve(raw)
        binary_path = any(raw == row["path"] or path == Path(row["path"])
                          for row in binaries)
        need(path == ROOT or ROOT in path.parents or path == PACKET
             or PACKET in path.parents or binary_path,
             f"freeze path is outside packet/workspace: {raw}")
        if path.is_file():
            need(not path.is_symlink() and sha(path) == expected, f"frozen path changed: {raw}")
        else:
            matching = [row for row in binaries if raw == row["path"] or path == Path(row["path"])]
            need(matching and expected == matching[0]["sha256"],
                 f"missing frozen path is not an exact binary witness: {raw}")
    candidate = build_receipts["candidate"]
    for relative, expected in candidate["source_sha256"].items():
        paths = [key for key in bindings if resolve(key) == ROOT / relative]
        archive_paths = [PACKET / "candidate-source" / relative,
                         PACKET / "baseline-source" / relative]
        found = any(key in bindings and bindings[key] == expected for key in paths)
        found = found or any(p.is_file() and sha(p) == expected for p in archive_paths)
        found = found or any(key.endswith("candidate-source/" + relative)
                             and bindings[key] == expected for key in bindings)
        need(found, f"candidate source is not frozen: {relative}")
    for relative, expected in candidate["probe_sha256"].items():
        paths = [key for key in bindings if resolve(key) == PACKET / relative]
        need(any(key in bindings and bindings[key] == expected for key in paths),
             f"candidate probe is not frozen: {relative}")
    if isinstance(value.get("binaries"), list):
        sort_key = lambda row: json.dumps(row, sort_keys=True)
        need(sorted(value["binaries"], key=sort_key) == sorted(binaries, key=sort_key),
             "freeze binary identities changed")


def ancestry() -> None:
    value = read(PACKET / "ancestry.json")
    need(isinstance(value, dict)
         and set(value) == {"packet", "artifact_manifest_sha256",
                             "oracle_contract_sha256", "baseline_head"},
         "ancestry receipt schema changed")
    need(value["packet"] == "change-0728", "ancestry packet changed")
    need(digest(value["artifact_manifest_sha256"], "ancestor artifact hash")
         == sha(ROOT / "docs/performance/results/change-0728/artifact-manifest.json"),
         "sealed artifact manifest changed")
    need(digest(value["oracle_contract_sha256"], "ancestor contract hash")
         == ANCESTOR_ORACLE_SHA, "ancestry contract hash changed")
    head = value["baseline_head"]
    need(isinstance(head, str) and len(head) == 40 and set(head) <= HEX,
         "baseline head is invalid")
    try:
        subprocess.run(["git", "cat-file", "-e", head + "^{commit}"], cwd=ROOT,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise AuditError("baseline head is not present") from error


def source_guard(build_receipts: dict[str, dict[str, Any]]) -> None:
    value = read(PACKET / "candidate-source.json")
    need(isinstance(value, dict) and isinstance(value.get("base_head"), str)
         and isinstance(value.get("files"), dict), "source guard receipt is incomplete")
    need(value["base_head"] == read(PACKET / "ancestry.json")["baseline_head"],
         "source guard base head differs from ancestry")
    baseline = build_receipts["baseline"]["source_sha256"]
    candidate = build_receipts["candidate"]["source_sha256"]
    changed = {relative for relative in set(baseline) | set(candidate)
               if baseline.get(relative) != candidate.get(relative)}
    need(set(value["files"]) == changed, "source guard changed-file set differs")
    for relative in changed:
        safe_relative(relative, "source guard path")
        need(value["files"][relative] == {"before_sha256": baseline.get(relative),
                                           "after_sha256": candidate.get(relative)},
             f"source guard row changed: {relative}")
        before, after = baseline.get(relative), candidate.get(relative)
        if before is not None:
            archive = PACKET / "baseline-source" / relative
            need(archive.is_file() and sha(archive) == before,
                 f"baseline archive changed: {relative}")
            try:
                original = subprocess.check_output(
                    ["git", "show", f"{value['base_head']}:{relative}"], cwd=ROOT)
            except (OSError, subprocess.CalledProcessError) as error:
                raise AuditError(f"baseline source cannot be recovered: {relative}") from error
            need(hashlib.sha256(original).hexdigest() == before,
                 f"baseline head source changed: {relative}")
        if after is not None:
            archive = PACKET / "candidate-source" / relative
            need(archive.is_file() and sha(archive) == after,
                 f"candidate archive changed: {relative}")
    disposition = read(PACKET / "disposition.json")
    need(isinstance(disposition, dict)
         and disposition.get("production") in ("candidate", "baseline"),
         "production disposition is missing")
    selected = candidate if disposition["production"] == "candidate" else baseline
    for relative, expected in selected.items():
        target = ROOT / relative
        need(target.is_file() and not target.is_symlink() and sha(target) == expected,
             f"selected source changed: {relative}")
    constraints = read(PACKET / "constraints.json")
    need(isinstance(constraints, dict), "constraints receipt is not an object")
    for relative, expected in constraints.items():
        safe_relative(relative, "constraint path")
        need(digest(expected, f"constraint {relative}") == sha(ROOT / relative),
             f"constraint changed: {relative}")


def quality_command_key(command: Any) -> list[str]:
    need(isinstance(command, list) and all(isinstance(item, str) for item in command),
         "quality command is not a string list")
    result = []
    for item in command:
        if item.endswith("/probe/Cargo.toml"):
            result.append("<probe-manifest>")
        elif item.endswith("/retention-probe/Cargo.toml"):
            result.append("<retention-manifest>")
        else:
            result.append(item)
    return result


def expected_quality_commands() -> list[list[str]]:
    common = ["--release", "--offline", "--locked"]
    return [
        ["cargo", "fmt", "-p", "litchi-doc", "-p", "litchi-ole-common", "--", "--check"],
        ["cargo", "test", "-p", "litchi-doc", "-p", "litchi-ole-common", *common,
         "--all-features", "--all-targets"],
        ["cargo", "clippy", "-p", "litchi-doc", "-p", "litchi-ole-common", *common,
         "--all-features", "--all-targets", "--", "-D", "warnings"],
        ["cargo", "test", "-p", "litchi-doc", "-p", "litchi-ole-common", *common,
         "--all-features", "--doc"],
        ["cargo", "doc", "-p", "litchi-doc", "-p", "litchi-ole-common", *common,
         "--all-features", "--no-deps"],
        ["cargo", "fmt", "--manifest-path", "<probe-manifest>", "--", "--check"],
        ["cargo", "test", "--manifest-path", "<probe-manifest>", *common, "--lib"],
        ["cargo", "clippy", "--manifest-path", "<probe-manifest>", *common,
         "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--manifest-path", "<probe-manifest>", *common, "--no-deps"],
        ["cargo", "fmt", "--manifest-path", "<retention-manifest>", "--", "--check"],
        ["cargo", "clippy", "--manifest-path", "<retention-manifest>", *common,
         "--all-targets", "--", "-D", "warnings"],
        ["cargo", "doc", "--manifest-path", "<retention-manifest>", *common, "--no-deps"],
        ["python3", "tools/check_crate_boundaries.py"],
    ]


def smoke_receipt(value: Any, case_rows: dict[str, dict[str, Any]],
                  contract_rows: dict[str, dict[str, Any]]) -> None:
    need(isinstance(value, dict) and value.get("status") == "pass"
         and value.get("kind") == "candidate-smoke", "candidate smoke failed")
    need(value.get("candidate_builds_sha256") == sha(PACKET / "candidate-builds.json"),
         "candidate smoke build hash changed")
    need(value.get("retention_builds_sha256") == sha(PACKET / "retention-builds.json"),
         "candidate smoke retention hash changed")
    runs = value.get("runs")
    need(isinstance(runs, list) and len(runs) == len(CASES), "candidate smoke count changed")
    seen = set()
    for row in runs:
        need(isinstance(row, dict) and row.get("case") in case_rows
             and row["case"] not in seen, "candidate smoke case set changed")
        seen.add(row["case"])
        case = case_rows[row["case"]]
        expected = contract_rows[row["case"]]["identity"]["expected_output_sha256"]
        need(row.get("format") == "doc" and row.get("input") == case["path"]
             and row.get("source_sha256") == case["sha256"]
             and row.get("expected_output_sha256") == expected
             and row.get("output_sha256") == expected
             and row.get("exit_code") == 0 and row.get("oracle_ok") is True,
             f"candidate smoke run failed: {row.get('case')}")
    need(seen == set(CASES), "candidate smoke case set incomplete")


def qualification(build_receipts: dict[str, dict[str, Any]], retention: dict[str, Any],
                  case_rows: dict[str, dict[str, Any]],
                  contract_rows: dict[str, dict[str, Any]]) -> None:
    source_record = read(PACKET / "candidate-source.json")
    need(isinstance(source_record, dict) and isinstance(source_record.get("files"), dict)
         and bool(source_record["files"]),
         "candidate source custody receipt is missing")
    disposition = read(PACKET / "disposition.json")
    need(isinstance(disposition, dict)
         and disposition.get("production") in ("candidate", "baseline"),
         "production disposition receipt is missing")
    before = read(PACKET / "before-qualification.json")
    need(before.get("status") == "pass" and before.get("inherited") == "0728 final oracle",
         "before qualification failed")
    runs = before.get("runs")
    need(isinstance(runs, list) and len(runs) == 4, "before qualification count changed")
    for row in runs:
        integer(row.get("oracles"), "before oracle count", 1)
        integer(row.get("negative_controls"), "before control count", 1)
    final = read(PACKET / "qualification.json")
    need(isinstance(final, dict) and final.get("status") == "pass",
         "final qualification failed")
    files = final.get("files")
    need(isinstance(files, dict) and bool(files), "qualification files are empty")
    for relative, expected in files.items():
        relative = safe_relative(relative, "qualification file")
        need(digest(expected, f"qualification file {relative}") == sha(PACKET / relative),
             f"qualification file changed: {relative}")
    need(final.get("candidate_builds_sha256") == sha(PACKET / "candidate-builds.json"),
         "qualification candidate build hash changed")
    need(final.get("retention_builds_sha256") == sha(PACKET / "retention-builds.json"),
         "qualification retention build hash changed")
    smoke_files = [relative for relative in files
                   if "smoke" in Path(relative).name.lower() and relative.endswith(".json")]
    need(smoke_files, "qualification omits candidate smoke receipt")
    for relative in smoke_files:
        smoke_receipt(read(PACKET / relative), case_rows, contract_rows)
    manifests = [relative for relative in files
                 if relative.startswith("quality-") and relative.endswith("/manifest.json")]
    need(manifests, "qualification omits quality manifest")
    expected_quality = {relative: digest(value, f"candidate source {relative}")
                        for relative, value in build_receipts["candidate"]["source_sha256"].items()
                        if relative.startswith("crates/") and Path(relative).suffix in {".rs", ".toml"}}
    for relative in manifests:
        manifest = read(PACKET / relative)
        need(isinstance(manifest, dict) and manifest.get("source_sha256") == expected_quality,
             f"quality source custody changed: {relative}")
        runs = manifest.get("runs")
        need(isinstance(runs, list) and len(runs) == 13 and all(row.get("exit_code") == 0 for row in runs),
             f"quality manifest failed: {relative}")
        need([quality_command_key(row.get("command")) for row in runs]
             == expected_quality_commands(), f"quality command matrix changed: {relative}")
    retained = read(PACKET / "retention-analysis.json")
    need(isinstance(retained, dict) and isinstance(retained.get("processes"), list)
         and len(retained["processes"]) == 24, "retention analysis count changed")


def capture_path(raw: Any, root: Path) -> Path:
    need(isinstance(raw, str) and raw and not Path(raw).is_absolute()
         and ".." not in Path(raw).parts, "capture path is unsafe")
    value = (root / raw).resolve()
    base = root.resolve()
    need(base == value or base in value.parents, "capture path escapes packet")
    need(value.is_file() and not value.is_symlink(), f"capture file missing: {raw}")
    return value


def expected(plan_value: dict[str, Any], case_rows: dict[str, dict[str, Any]],
             receipt: dict[str, dict[str, Any]], before: bool) -> list[dict[str, Any]]:
    rows = []
    if before:
        for lane, samples, warmups in (("native", 50, 3), ("allocation", 1, 0)):
            binary_name = f"baseline-ole_format_save_probe" + ("_alloc" if lane == "allocation" else "")
            binary = next(row["path"] for row in receipt["baseline"]["binaries"]
                          if Path(row["path"]).name == binary_name)
            for case in CASES:
                output = f"{lane}-c0-{case}-0-baseline.json"
                rows.append({"lane": lane, "cycle": 0, "case": case, "variant": "baseline",
                             "slot": 0, "output": output, "stderr": output + ".stderr",
                             "samples": samples, "warmups": warmups,
                             "command": ["taskset", "-c", "12", binary, "--case", case,
                                         "--input", case_rows[case]["path"], "--operation", "format",
                                         "--samples", str(samples), "--warmups", str(warmups)]})
        return rows
    for cycle in range(3):
        for case in (CASES if cycle % 2 == 0 else tuple(reversed(CASES))):
            for slot, variant in enumerate(NATIVE_SLOTS):
                binary = next(row["path"] for row in receipt[variant]["binaries"]
                              if Path(row["path"]).name == f"{variant}-ole_format_save_probe")
                output = f"native-c{cycle}-{case}-{slot}-{variant}.json"
                rows.append({"lane": "native", "cycle": cycle, "case": case,
                             "variant": variant, "slot": slot, "output": output,
                             "stderr": output + ".stderr", "samples": 50, "warmups": 3,
                             "command": ["taskset", "-c", "12", binary, "--case", case,
                                         "--input", case_rows[case]["path"], "--operation", "format",
                                         "--samples", "50", "--warmups", "3"]})
    for case in CASES:
        for slot, variant in enumerate(ALLOCATION_SLOTS):
            binary = next(row["path"] for row in receipt[variant]["binaries"]
                          if Path(row["path"]).name == f"{variant}-ole_format_save_probe_alloc")
            output = f"allocation-c0-{case}-{slot}-{variant}.json"
            rows.append({"lane": "allocation", "cycle": 0, "case": case,
                         "variant": variant, "slot": slot, "output": output,
                         "stderr": output + ".stderr", "samples": 1, "warmups": 0,
                         "command": ["taskset", "-c", "12", binary, "--case", case,
                                     "--input", case_rows[case]["path"], "--operation", "format",
                                     "--samples", "1", "--warmups", "0"]})
    return rows


def guard(oracle: Any, label: str) -> None:
    need(isinstance(oracle, dict) and oracle.get("oracle_ok") is True
         and oracle.get("failure_reasons") == [], f"{label}: oracle failed")
    def visit(value: Any, path: str) -> None:
        if isinstance(value, bool):
            need(value, f"{label}: false oracle field {path}")
        elif isinstance(value, dict):
            for key, nested in value.items():
                visit(nested, f"{path}.{key}")
        elif isinstance(value, list):
            for index, nested in enumerate(value):
                visit(nested, f"{path}[{index}]")

    visit(oracle, "oracle")


def stat(values: list[int]) -> dict[str, int | float]:
    need(values, "empty timing samples")
    ordered = sorted(values)
    return {"n": len(values), "p50": statistics.median(ordered),
            "mean": math.fsum(values) / len(values),
            "p95": ordered[math.ceil(.95 * len(ordered)) - 1],
            "p99": ordered[math.ceil(.99 * len(ordered)) - 1],
            "maximum": ordered[-1]}


def report(value: Any, row: dict[str, Any], case: dict[str, Any],
           contract_row: dict[str, Any], ids: dict[str, Any], outputs: dict[str, str]) -> tuple[dict[str, Any], dict[str, int]]:
    need(isinstance(value, dict), "report is not an object")
    label = f"{row['lane']}/c{row['cycle']}/{row['case']}/{row['slot']}/{row['variant']}"
    native = row["lane"] == "native"
    headers = {"schema_version": 1, "case": row["case"], "format": "doc",
               "operation": "format", "scope": "public_format_open_edit_commit",
               "input": case["path"], "policy": "reuse", "policy_applied": False,
               "policy_application_scope": "not_applied_public_format_route",
               "policy_argument_effect": "ignored_public_format_default_route",
               "timing_claim": native, "allocator_instrumented": not native,
               "directory_metadata_fields": META_FIELDS,
               "allocation_ownership_contract": OWNERSHIP,
               "warmups": row["warmups"], "samples_requested": row["samples"],
               "source_sha256": case["sha256"]}
    for key, expected_value in headers.items():
        need(value.get(key) == expected_value, f"{label}: header {key} changed")
    need(isinstance(value.get("policy_contract"), str) and value["policy_contract"],
         f"{label}: policy contract missing")
    keys = ("source_sha256", "expected_output_sha256", "replacements_sha256",
            "source_inventory", "expected_output_inventory", "replacements",
            "changed_length_proof")
    identity = {key: value.get(key) for key in keys}
    need(identity == contract_row["identity"], f"{label}: identity changed")
    need(ids.setdefault(row["case"], identity) == identity, f"{label}: identity drifted")
    oracle = value.get("expected_oracle")
    guard(oracle, f"{label}: expected")
    need(oracle.get("semantic_witness") == contract_row["semantic_witness"],
         f"{label}: semantic witness changed")
    need(oracle.get("directory_metadata_differences") == [], f"{label}: metadata differs")
    proof = value.get("changed_length_proof")
    need(isinstance(proof, dict) and proof.get("logical_stream_length_change_proven") is True
         and proof.get("format_specific_semantic_length_proven") is True
         and proof.get("any_stream_length_changed") is True
         and integer(proof.get("changed_stream_count"), f"{label}: changed streams", 1) >= 1,
         f"{label}: length proof failed")
    controls = value.get("oracle_controls")
    need(isinstance(controls, list) and [item.get("name") for item in controls]
         == contract_row["control_names"], f"{label}: controls changed")
    for control in controls:
        need(isinstance(control, dict) and control.get("status") == "rejected"
             and control.get("rejected") is True and isinstance(control.get("failure_reasons"), list)
             and bool(control["failure_reasons"]), f"{label}: control accepted")
    expected_inventory = value.get("expected_output_inventory")
    need(isinstance(expected_inventory, dict), f"{label}: expected inventory missing")
    timings, allocation, output_hash = [], {}, None
    for index, sample in enumerate(value.get("samples", [])):
        need(isinstance(sample, dict) and sample.get("index") == index, f"{label}: sample index")
        guard(sample.get("oracle"), f"{label}: sample {index}")
        need(sample["oracle"].get("semantic_witness") == contract_row["semantic_witness"],
             f"{label}: sample witness")
        raw = sample["oracle"].get("raw_directory")
        need(isinstance(raw, dict) and raw.get("source_expected_difference_bytes") == 0
             and raw.get("expected_output_difference_bytes") == 0
             and raw.get("source_output_difference_bytes") == 0, f"{label}: raw gate")
        need(sample.get("output_inventory") == expected_inventory, f"{label}: output inventory")
        current = digest(sample.get("output_sha256"), f"{label}: output hash")
        need(current == value["expected_output_sha256"], f"{label}: output identity")
        output_hash = output_hash or current
        need(current == output_hash, f"{label}: output drift")
        if native:
            phase = sample.get("phase_ns")
            need("allocations" not in sample and isinstance(phase, dict)
                 and set(phase) == {"whole_ns"}, f"{label}: timing schema")
            timings.append(integer(phase["whole_ns"], f"{label}: whole_ns", 1))
        else:
            alloc = sample.get("allocations")
            need("phase_ns" not in sample and isinstance(alloc, dict)
                 and set(alloc) == {"whole"} and isinstance(alloc["whole"], dict),
                 f"{label}: allocation schema")
            whole = alloc["whole"]
            need(set(whole) == set(ALLOC_FIELDS), f"{label}: allocation fields")
            current_alloc = {key: integer(whole[key], f"{label}: {key}") for key in ALLOC_FIELDS}
            allocation = allocation or current_alloc
            need(current_alloc == allocation, f"{label}: allocation drift")
    need(output_hash is not None and isinstance(value.get("samples"), list)
         and len(value["samples"]) == row["samples"], f"{label}: sample count")
    outputs.setdefault(row["case"], output_hash)
    need(outputs[row["case"]] == output_hash, f"{label}: case output drift")
    return (stat(timings) if native else {}, allocation)


def capture(expected_rows: list[dict[str, Any]], directory: Path,
            case_rows: dict[str, dict[str, Any]], contract_rows: dict[str, Any]) -> list[dict[str, Any]]:
    manifest = read(directory / "manifest.json")
    need(manifest.get("status") == "complete", f"{directory.name}: capture incomplete")
    runs = manifest.get("runs")
    need(isinstance(runs, list) and len(runs) == len(expected_rows),
         f"{directory.name}: process count")
    seen, ids, outputs, process = set(), {}, {}, []
    for actual, planned in zip(runs, expected_rows):
        need(isinstance(actual, dict), "capture row is not an object")
        for key in ("lane", "cycle", "case", "variant", "slot"):
            need(actual.get(key) == planned[key], f"{directory.name}: order changed at {planned}")
        need(actual.get("exit_code") == 0 and actual.get("command") == planned["command"],
             f"{directory.name}: process command changed")
        output = capture_path(actual.get("output"), directory)
        stderr = capture_path(actual.get("stderr"), directory)
        output_name = output.relative_to(directory.resolve()).as_posix()
        stderr_name = stderr.relative_to(directory.resolve()).as_posix()
        need(output_name == planned["output"] and stderr_name == planned["stderr"],
             f"{directory.name}: output name changed")
        need(output_name not in seen and stderr_name not in seen, f"{directory.name}: duplicate path")
        seen.update((output_name, stderr_name))
        need(actual.get("sha256") == sha(output) and actual.get("stderr_sha256") == sha(stderr),
             f"{directory.name}: raw hash changed")
        timing, allocation = report(read(output), planned, case_rows[planned["case"]],
                                    contract_rows[planned["case"]], ids, outputs)
        process.append({"lane": planned["lane"], "cycle": planned["cycle"],
                        "case": planned["case"], "variant": planned["variant"],
                        "slot": planned["slot"], "output_sha256": outputs[planned["case"]],
                        "timing": timing, "allocations": allocation})
    on_disk = {p.relative_to(directory.resolve()).as_posix() for p in directory.rglob("*")
               if p.is_file() and p.name != "manifest.json"}
    need(on_disk == seen, f"{directory.name}: unmanifested raw files")
    return process


def compare(actual: Any, expected_value: Any, path: str = "analysis") -> None:
    if isinstance(actual, bool) or isinstance(expected_value, bool):
        need(actual is expected_value, f"{path}: boolean differs")
    elif isinstance(actual, dict) or isinstance(expected_value, dict):
        need(isinstance(actual, dict) and isinstance(expected_value, dict)
             and set(actual) == set(expected_value), f"{path}: object shape differs")
        for key in expected_value:
            compare(actual[key], expected_value[key], f"{path}.{key}")
    elif isinstance(actual, list) or isinstance(expected_value, list):
        need(isinstance(actual, list) and isinstance(expected_value, list)
             and len(actual) == len(expected_value), f"{path}: list shape differs")
        for index, (left, right) in enumerate(zip(actual, expected_value)):
            compare(left, right, f"{path}[{index}]")
    elif isinstance(actual, (int, float)) and isinstance(expected_value, (int, float)):
        need(math.isclose(float(actual), float(expected_value), rel_tol=1e-12, abs_tol=1e-9),
             f"{path}: numeric differs")
    else:
        need(actual == expected_value, f"{path}: value differs")


def pair(process: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for case in CASES:
        for cycle in range(3):
            selected = {row["slot"]: row for row in process if row["lane"] == "native"
                        and row["case"] == case and row["cycle"] == cycle}
            need(set(selected) == set(range(6)), f"pair matrix incomplete: {case}/{cycle}")
            for label, first_slot, second_slot in (("AA", 0, 1), ("AB", 2, 3), ("BA", 5, 4)):
                first, second = selected[first_slot]["timing"], selected[second_slot]["timing"]
                differences = {}
                for key in STAT_FIELDS:
                    left, right = float(first[key]), float(second[key])
                    delta = right - left
                    percent = None if left == 0 else delta / abs(left) * 100
                    differences[key] = {"first": left, "second": right, "delta": delta,
                                        "percent": percent,
                                        "flag_gt_5pct": (delta != 0 if percent is None
                                                          else abs(percent) > 5)}
                result.append({"case": case, "cycle": cycle, "pair": label,
                               "first_slot": first_slot, "second_slot": second_slot,
                               "first_variant": selected[first_slot]["variant"],
                               "second_variant": selected[second_slot]["variant"],
                               "differences": differences,
                               "flags_gt_5pct": [key for key, value in differences.items()
                                                  if value["flag_gt_5pct"]]})
    return result


def allocation(process: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result = []
    for case in CASES:
        selected = {row["slot"]: row for row in process if row["lane"] == "allocation"
                    and row["case"] == case}
        need(set(selected) == set(range(4)), f"allocation matrix incomplete: {case}")
        fields = {}
        for key in ALLOC_FIELDS:
            baseline = [selected[0]["allocations"][key], selected[3]["allocations"][key]]
            candidate = [selected[1]["allocations"][key], selected[2]["allocations"][key]]
            left, right = math.fsum(baseline) / 2, math.fsum(candidate) / 2
            delta = right - left
            percent = None if left == 0 else delta / abs(left) * 100
            fields[key] = {"baseline_slots": baseline, "candidate_slots": candidate,
                           "baseline_mean": left, "candidate_mean": right,
                           "first": left, "second": right, "delta": delta,
                           "percent": percent,
                           "flag_gt_5pct": delta != 0 if percent is None else abs(percent) > 5}
        result.append({"case": case, "slots": [0, 1, 2, 3], "fields": fields,
                       "flags_gt_5pct": [key for key, value in fields.items()
                                          if value["flag_gt_5pct"]]})
    return result


def main() -> int:
    try:
        plan_value = plan()
        case_rows = cases()
        contract_rows = contract(case_rows)
        build_receipts = builds()
        retention_receipt = retention_build()
        binary_witness(build_receipts, retention_receipt)
        freeze(build_receipts, retention_receipt)
        ancestry()
        source_guard(build_receipts)
        qualification(build_receipts, retention_receipt, case_rows, contract_rows)
        before_rows = capture(expected(plan_value, case_rows, build_receipts, True), BEFORE,
                              case_rows, contract_rows)
        process = capture(expected(plan_value, case_rows, build_receipts, False), CAPTURES,
                          case_rows, contract_rows)
        need(len(before_rows) == 4 and len(process) == 44, "process totals changed")
        independent = {"packet": "change-0730",
                       "disposition": "bounded validated-render handoff paired measurement; no adoption claim",
                       "matrix": {"before_processes": 4, "native_processes": 36,
                                  "allocation_processes": 8, "cases": list(CASES),
                                  "native_slots": list(NATIVE_SLOTS),
                                  "allocation_slots": list(ALLOCATION_SLOTS)},
                       "statistics": {"quantiles": "nearest-rank p95/p99; midpoint p50",
                                       "fields": list(STAT_FIELDS),
                                       "relative_threshold_percent": 5.0,
                                       "timing_unit": "ns", "allocation_fields": list(ALLOC_FIELDS)},
                       "before_processes": before_rows, "processes": process,
                       "native_pairs": pair(process),
                       "allocation_comparisons": allocation(process)}
        analysis = read(PACKET / "analysis.json")
        compare(analysis, independent)
        record = {"packet": "change-0730", "analysis_match": True,
                  "custody": {"plan_sha256": sha(PACKET / "plan.json"),
                              "analysis_sha256": sha(PACKET / "analysis.json"),
                              "freeze_sha256": sha(PACKET / "freeze.json"),
                              "manifest_sha256": sha(CAPTURES / "manifest.json"),
                              "before_manifest_sha256": sha(BEFORE / "manifest.json"),
                              "processes": len(process), "before_processes": len(before_rows)},
                  "processes": process, "native_pairs": independent["native_pairs"],
                  "allocation_comparisons": independent["allocation_comparisons"]}
        (PACKET / "audit.json").write_text(json.dumps(record, indent=2, sort_keys=True) + "\n",
                                             encoding="utf-8")
        print(json.dumps({"status": "PASS", "processes": len(process), "analysis_match": True},
                         sort_keys=True))
        return 0
    except (AuditError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"FAIL: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
