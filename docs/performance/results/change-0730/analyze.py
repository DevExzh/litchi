#!/usr/bin/env python3
"""Validate and summarize the bounded DOC handoff capture.

The capture is deliberately treated as evidence, rather than as a source of
truth.  Fixture identities, semantic witnesses, command order, probe custody,
and the rejected controls are pinned by the packet contract before timing
numbers are read.  The analyzer does not run Cargo or a native executable.
"""

from __future__ import annotations

import hashlib
import json
import math
import statistics
import subprocess
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
CAPTURES = P / "captures"
BEFORE = P / "before"
HEX = set("0123456789abcdef")
CASES = ("docfloat", "docnohf")
NATIVE_SLOTS = ("baseline", "baseline", "baseline", "candidate", "candidate", "baseline")
ALLOCATION_SLOTS = ("baseline", "candidate", "candidate", "baseline")
META_FIELDS = [
    "entry_type", "name_utf16", "clsid", "bytes", "start_sector", "is_minifat",
    "raw_left_sibling", "raw_right_sibling", "raw_child", "raw_color",
    "raw_state_bits", "raw_creation_time", "raw_modification_time",
]
OWNERSHIP_CONTRACT = (
    "Each allocation region reports boundary-relative allocations. The opened "
    "Editor is retained from open through replace_and_validate; the staged "
    "Editor is retained until finish; the returned Vec is retained across "
    "finish. Format output is also retained outside its region for validation. "
    "retained_bytes is live ownership at the region boundary, not RSS."
)
STAT_FIELDS = ("p50", "mean", "p95", "p99", "maximum")
ALLOCATION_FIELDS = (
    "allocated_bytes", "deallocated_bytes", "allocation_calls",
    "peak_live_bytes", "retained_bytes",
)
ANCESTOR_ORACLE_SHA = "9d1baa0a4978aa9f634545a09d4a594dfb2cfa3785c1895576680045af8b44f9"
LOCAL_ORACLE_SHA = "f870329cc076be509c2066e8ab654b8ce24ab63c3737c110710697a13eb0b4ea"


class Failure(Exception):
    """A malformed, incomplete, or unqualified evidence packet."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise Failure(message)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise Failure(f"invalid JSON {path}: {error}") from error


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing or symlinked file: {path}")
    return sha_bytes(path.read_bytes())


def digest(value: Any, label: str) -> str:
    require(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
            f"{label} is not a lowercase SHA-256 digest")
    return value


def integer(value: Any, label: str, minimum: int = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool) and value >= minimum,
            f"{label} is not an integer >= {minimum}")
    return value


def read_plan() -> dict[str, Any]:
    plan = read_json(P / "plan.json")
    require(isinstance(plan, dict), "plan is not an object")
    expected = {
        "cpu": 12,
        "samples": 50,
        "warmups": 3,
        "cycles": 3,
        "cases": list(CASES),
        "native_order": list(NATIVE_SLOTS),
        "native_processes": 36,
        "allocation_order": list(ALLOCATION_SLOTS),
        "allocation_processes": 8,
        "before_native_processes": 2,
        "before_allocation_processes": 2,
    }
    for key, value in expected.items():
        require(plan.get(key) == value, f"plan field changed: {key}")
    return plan


def read_cases(plan: dict[str, Any]) -> dict[str, dict[str, Any]]:
    rows = read_json(P / "cases.json")
    require(isinstance(rows, list) and [row.get("case") for row in rows] == list(CASES),
            "case order changed")
    result: dict[str, dict[str, Any]] = {}
    expected = {
        "docfloat": ("doc", "test-data/ole/doc/FloatingPictures.doc", 335360,
                     "41cc0cd8f2d7390266f844f5bddbcd29d08ac5338abc33832d9c9a6ed5d0f85d"),
        "docnohf": ("doc", "test-data/ole/doc/NoHeadFoot.doc", 26112,
                    "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48"),
    }
    for row in rows:
        require(isinstance(row, dict), "case row is not an object")
        case = row.get("case")
        require(case in expected and case not in result, f"invalid case identity: {case!r}")
        fmt, path, size, source_hash = expected[case]
        require(row.get("format") == fmt and row.get("path") == path,
                f"{case}: fixture identity changed")
        fixture = ROOT / path
        require(fixture.is_file() and not fixture.is_symlink(), f"missing fixture: {fixture}")
        require(row.get("bytes") == size == fixture.stat().st_size, f"{case}: fixture size changed")
        require(row.get("sha256") == source_hash == sha(fixture), f"{case}: fixture hash changed")
        result[case] = row
    require(set(result) == set(CASES), "case set changed")
    return result


def read_contract(cases: dict[str, dict[str, Any]]) -> dict[str, dict[str, Any]]:
    ancestor_path = ROOT / "docs/performance/results/change-0728/oracle-contract.json"
    require(sha(ancestor_path) == ANCESTOR_ORACLE_SHA,
            "sealed 0728 oracle contract changed")
    require(sha(P / "oracle-contract.json") == LOCAL_ORACLE_SHA,
            "0730 oracle contract copy changed")
    ancestor = read_json(ancestor_path)
    contract = read_json(P / "oracle-contract.json")
    require(isinstance(contract, dict) and set(contract) == set(cases),
            "oracle contract case set changed")
    require(isinstance(ancestor, dict) and set(CASES) <= set(ancestor),
            "sealed ancestor contract is incomplete")
    required = {"directory_metadata_fields", "allocation_ownership_contract",
                "semantic_witness", "control_names", "identity"}
    for case, value in contract.items():
        require(isinstance(value, dict) and set(value) == required,
                f"{case}: oracle contract schema changed")
        require(value["directory_metadata_fields"] == META_FIELDS,
                f"{case}: directory metadata contract changed")
        require(value["allocation_ownership_contract"] == OWNERSHIP_CONTRACT,
                f"{case}: allocation ownership contract changed")
        require(value == ancestor[case], f"{case}: contract differs from sealed 0728 witness")
        require(isinstance(value["semantic_witness"], dict),
                f"{case}: semantic witness is not an object")
        require(isinstance(value["control_names"], list) and value["control_names"],
                f"{case}: controls are empty")
        require(isinstance(value["identity"], dict), f"{case}: identity is not an object")
    return contract


def safe_relative(value: Any, label: str) -> str:
    require(isinstance(value, str) and value and not Path(value).is_absolute()
            and ".." not in Path(value).parts, f"{label} is unsafe: {value!r}")
    return value


def verify_custody_map(mapping: Any, root: Path, label: str) -> dict[str, str]:
    require(isinstance(mapping, dict) and mapping, f"{label} custody map is empty")
    checked: dict[str, str] = {}
    for relative, expected in mapping.items():
        relative = safe_relative(relative, f"{label} path")
        expected = digest(expected, f"{label} {relative}")
        target = root / relative
        archived = [P / "candidate-source" / relative,
                    P / "baseline-source" / relative]
        live_match = target.is_file() and not target.is_symlink() and sha(target) == expected
        archive_match = any(item.is_file() and not item.is_symlink() and sha(item) == expected
                            for item in archived)
        require(live_match or archive_match, f"{label} changed or missing: {relative}")
        checked[relative] = expected
    return checked


def read_build(variant: str) -> dict[str, Any]:
    path = P / f"{variant}-builds.json"
    build = read_json(path)
    require(isinstance(build, dict), f"{variant} build receipt is not an object")
    receipt_name = build.get("receipt")
    safe_relative(receipt_name, f"{variant} build receipt")
    receipt = read_json(P / receipt_name)
    require(receipt.get("exit_code") == 0, f"{variant} build did not pass")
    require(build.get("source_sha256") == receipt.get("source_sha256"),
            f"{variant} source receipt differs from build log")
    require(build.get("probe_sha256") == receipt.get("probe_sha256"),
            f"{variant} probe receipt differs from build log")
    require(build.get("binaries") == receipt.get("binaries"),
            f"{variant} binary receipt differs from build log")
    verify_custody_map(build.get("source_sha256"), ROOT, f"{variant} source")
    verify_custody_map(build.get("probe_sha256"), P, f"{variant} probe")
    binaries = build.get("binaries")
    require(isinstance(binaries, list) and len(binaries) == 2,
            f"{variant} binary count changed")
    expected_names = {f"{variant}-ole_format_save_probe",
                      f"{variant}-ole_format_save_probe_alloc"}
    require({Path(row.get("path", "")).name for row in binaries} == expected_names,
            f"{variant} binary names changed")
    for row in binaries:
        require(isinstance(row, dict), f"{variant} binary row is not an object")
        binary_path = row.get("path")
        require(isinstance(binary_path, str) and Path(binary_path).is_absolute(),
                f"{variant} binary path is not absolute")
        integer(row.get("bytes"), f"{variant} binary bytes", 1)
        digest(row.get("sha256"), f"{variant} binary hash")
    return build


def read_retention_build() -> dict[str, Any]:
    """Validate the separate candidate-only retention probe receipt."""
    build = read_json(P / "retention-builds.json")
    require(isinstance(build, dict), "retention build receipt is not an object")
    require(build.get("exit_code") == 0, "retention build did not pass")
    require(build.get("candidate_builds_sha256") == sha(P / "candidate-builds.json"),
            "retention build is not bound to candidate builds")
    verify_custody_map(build.get("probe_sha256"), P, "retention probe")
    binaries = build.get("binaries")
    require(isinstance(binaries, list) and len(binaries) == 2,
            "retention binary count changed")
    require({Path(row.get("path", "")).name for row in binaries}
            == {"doc_retention_probe", "doc_retention_probe_alloc"},
            "retention binary names changed")
    for row in binaries:
        require(isinstance(row, dict) and isinstance(row.get("path"), str)
                and Path(row["path"]).is_absolute(), "retention binary path changed")
        integer(row.get("bytes"), "retention binary bytes", 1)
        digest(row.get("sha256"), "retention binary hash")
    return build


def cleanup_witnesses() -> list[dict[str, Any]]:
    path = P / "cleanup.json"
    if not path.is_file() or path.is_symlink():
        return []
    cleanup = read_json(path)
    require(isinstance(cleanup, dict), "cleanup receipt is not an object")
    witnesses = cleanup.get("identities")
    require(isinstance(witnesses, list), "cleanup binary identities are not a list")
    return witnesses


def verify_binaries(builds: dict[str, dict[str, Any]], retention: dict[str, Any]) -> list[dict[str, Any]]:
    expected = [row for build in builds.values() for row in build["binaries"]]
    expected.extend(retention["binaries"])
    witnesses = cleanup_witnesses()
    missing = []
    for row in expected:
        path = Path(row["path"])
        if path.is_file():
            require(not path.is_symlink() and path.stat().st_size == row["bytes"]
                    and sha(path) == row["sha256"], f"live binary identity changed: {path}")
        else:
            require(not path.exists() and not path.is_symlink(),
                    f"binary path is neither file nor absent: {path}")
            missing.append(row)
    sorted_rows = lambda rows: sorted(rows, key=lambda row: json.dumps(row, sort_keys=True))
    if missing:
        cleanup = read_json(P / "cleanup.json")
        require(isinstance(cleanup, dict) and cleanup.get("removed") is True,
                "missing binaries have no completed cleanup receipt")
        require(sorted_rows(witnesses) == sorted_rows(expected),
                "cleanup binary identities are not an exact build witness")
    elif witnesses:
        require(sorted_rows(witnesses) == sorted_rows(expected),
                "cleanup binary identities are not an exact build witness")
    return expected


def resolve_binding(raw: str) -> Path:
    path = Path(raw)
    if path.is_absolute():
        return path
    if raw.startswith("docs/") or raw.startswith("crates/") or raw in {"Cargo.toml", "Cargo.lock"}:
        return ROOT / raw
    return P / raw


def read_freeze(builds: dict[str, dict[str, Any]], retention: dict[str, Any]) -> tuple[dict[str, str], dict[str, Any]]:
    frozen = read_json(P / "freeze.json")
    require(isinstance(frozen, dict), "freeze is not an object")
    raw_bindings = frozen.get("bindings", frozen)
    require(isinstance(raw_bindings, dict) and raw_bindings, "freeze bindings are empty")
    bindings: dict[str, str] = {}
    for raw, expected in raw_bindings.items():
        require(isinstance(raw, str), "freeze binding path is not a string")
        expected = digest(expected, f"freeze binding {raw}")
        target = resolve_binding(raw)
        binary = any(raw == row["path"] or target == Path(row["path"])
                     for build in builds.values() for row in build["binaries"])
        binary = binary or any(raw == row["path"] or target == Path(row["path"])
                               for row in retention["binaries"])
        retained_tree = target == ROOT or ROOT in target.parents or target == P or P in target.parents
        require(retained_tree or binary, f"freeze binding is outside the packet/workspace: {raw}")
        if target.is_file():
            require(not target.is_symlink() and sha(target) == expected,
                    f"frozen binding changed: {raw}")
        else:
            require(binary, f"missing non-binary frozen binding: {raw}")
            # The exact binary witness is checked separately after cleanup.
            expected_rows = [row for build in builds.values() for row in build["binaries"]
                             if raw == row["path"] or target == Path(row["path"])]
            expected_rows.extend(row for row in retention["binaries"]
                                 if raw == row["path"] or target == Path(row["path"]))
            require(expected_rows and expected == expected_rows[0]["sha256"],
                    f"missing binary frozen hash changed: {raw}")
        bindings[raw] = expected
    path_keys: dict[Path, list[str]] = {}
    for key in bindings:
        path_keys.setdefault(resolve_binding(key), []).append(key)
    candidate = builds["candidate"]
    # Candidate source and every probe/build packet input must be frozen.  The
    # path may be absolute (the original capture) or packet/root relative.
    for relative, expected in candidate["source_sha256"].items():
        matches = path_keys.get(ROOT / relative, [])
        archived = P / "candidate-source" / relative
        require((matches and bindings[matches[0]] == expected)
                or (archived.is_file() and sha(archived) == expected
                    and any(key.endswith("candidate-source/" + relative) and bindings[key] == expected
                            for key in bindings)),
                f"candidate source is not frozen: {relative}")
    for relative, expected in candidate["probe_sha256"].items():
        matches = path_keys.get(P / relative, [])
        require(matches and bindings[matches[0]] == expected,
                f"candidate probe is not frozen: {relative}")
    if isinstance(frozen.get("binaries"), list):
        expected = [row for build in builds.values() for row in build["binaries"]]
        expected.extend(read_retention_build()["binaries"])
        require(sorted(frozen["binaries"], key=lambda row: json.dumps(row, sort_keys=True))
                == sorted(expected, key=lambda row: json.dumps(row, sort_keys=True)),
                "freeze binary identities changed")
    return bindings, frozen


def verify_ancestry() -> None:
    ancestry = read_json(P / "ancestry.json")
    require(isinstance(ancestry, dict)
            and set(ancestry) == {"packet", "artifact_manifest_sha256",
                                  "oracle_contract_sha256", "baseline_head"},
            "ancestry receipt schema changed")
    require(ancestry["packet"] == "change-0728", "ancestry packet changed")
    require(digest(ancestry["artifact_manifest_sha256"], "ancestor artifact hash")
            == sha(ROOT / "docs/performance/results/change-0728/artifact-manifest.json"),
            "sealed 0728 artifact manifest changed")
    require(digest(ancestry["oracle_contract_sha256"], "ancestor contract hash")
            == ANCESTOR_ORACLE_SHA, "ancestry oracle contract changed")
    head = ancestry["baseline_head"]
    require(isinstance(head, str) and len(head) == 40 and set(head) <= HEX,
            "baseline head is not a commit hash")
    try:
        subprocess.run(["git", "cat-file", "-e", head + "^{commit}"], cwd=ROOT,
                       check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except (OSError, subprocess.CalledProcessError) as error:
        raise Failure("baseline head is not present") from error


def verify_source_guard(builds: dict[str, dict[str, Any]]) -> None:
    record = read_json(P / "candidate-source.json")
    require(isinstance(record, dict) and isinstance(record.get("base_head"), str)
            and isinstance(record.get("files"), dict),
            "candidate source guard receipt is incomplete")
    require(record["base_head"] == read_json(P / "ancestry.json")["baseline_head"],
            "candidate source guard base head differs from ancestry")
    baseline = builds["baseline"]["source_sha256"]
    candidate = builds["candidate"]["source_sha256"]
    changed = {relative for relative in set(baseline) | set(candidate)
               if baseline.get(relative) != candidate.get(relative)}
    files = record["files"]
    require(set(files) == changed, "candidate source guard changed-file set differs")
    for relative in changed:
        safe_relative(relative, "candidate source guard path")
        require(files[relative] == {"before_sha256": baseline.get(relative),
                                    "after_sha256": candidate.get(relative)},
                f"candidate source guard row changed: {relative}")
        before, after = baseline.get(relative), candidate.get(relative)
        if before is not None:
            archive = P / "baseline-source" / relative
            require(archive.is_file() and sha(archive) == before,
                    f"baseline source archive changed: {relative}")
            try:
                original = subprocess.check_output(
                    ["git", "show", f"{record['base_head']}:{relative}"], cwd=ROOT)
            except (OSError, subprocess.CalledProcessError) as error:
                raise Failure(f"baseline source is not recoverable: {relative}") from error
            require(sha_bytes(original) == before, f"baseline head source changed: {relative}")
        if after is not None:
            archive = P / "candidate-source" / relative
            require(archive.is_file() and sha(archive) == after,
                    f"candidate source archive changed: {relative}")
    disposition = read_json(P / "disposition.json")
    require(isinstance(disposition, dict)
            and disposition.get("production") in ("candidate", "baseline"),
            "production disposition is missing")
    selected = candidate if disposition["production"] == "candidate" else baseline
    for relative, expected in selected.items():
        target = ROOT / relative
        require(target.is_file() and not target.is_symlink() and sha(target) == expected,
                f"selected production source changed: {relative}")
    constraints = read_json(P / "constraints.json")
    require(isinstance(constraints, dict), "constraints receipt is not an object")
    for relative, expected in constraints.items():
        safe_relative(relative, "constraint path")
        require(digest(expected, f"constraint {relative}") == sha(ROOT / relative),
                f"constraint source changed: {relative}")


def quality_command_key(command: Any) -> list[str]:
    require(isinstance(command, list) and all(isinstance(item, str) for item in command),
            "quality command is not a string list")
    normalized = []
    for item in command:
        if item.endswith("/probe/Cargo.toml"):
            normalized.append("<probe-manifest>")
        elif item.endswith("/retention-probe/Cargo.toml"):
            normalized.append("<retention-manifest>")
        else:
            normalized.append(item)
    return normalized


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


def verify_smoke_receipt(smoke: Any, cases: dict[str, dict[str, Any]],
                         contract: dict[str, dict[str, Any]],
                         builds: dict[str, dict[str, Any]], retention: dict[str, Any]) -> None:
    require(isinstance(smoke, dict) and smoke.get("status") == "pass"
            and smoke.get("kind") == "candidate-smoke",
            "candidate smoke receipt did not pass")
    require(smoke.get("candidate_builds_sha256") == sha(P / "candidate-builds.json"),
            "candidate smoke build hash changed")
    require(smoke.get("retention_builds_sha256") == sha(P / "retention-builds.json"),
            "candidate smoke retention build hash changed")
    runs = smoke.get("runs")
    require(isinstance(runs, list) and len(runs) == len(CASES),
            "candidate smoke run count changed")
    seen = set()
    for run in runs:
        require(isinstance(run, dict) and run.get("case") in cases
                and run["case"] not in seen, "candidate smoke case set changed")
        seen.add(run["case"])
        case = cases[run["case"]]
        expected = contract[run["case"]]["identity"]["expected_output_sha256"]
        require(run.get("format") == "doc" and run.get("input") == case["path"]
                and run.get("source_sha256") == case["sha256"]
                and run.get("expected_output_sha256") == expected
                and run.get("output_sha256") == expected
                and run.get("exit_code") == 0 and run.get("oracle_ok") is True,
                f"candidate smoke run failed: {run.get('case')}")
    require(seen == set(CASES), "candidate smoke case set incomplete")


def verify_qualification(builds: dict[str, dict[str, Any]], retention: dict[str, Any],
                         cases: dict[str, dict[str, Any]],
                         contract: dict[str, dict[str, Any]]) -> None:
    source_record = read_json(P / "candidate-source.json")
    require(isinstance(source_record, dict) and isinstance(source_record.get("files"), dict)
            and bool(source_record["files"]),
            "candidate source custody receipt is missing or incomplete")
    disposition = read_json(P / "disposition.json")
    require(isinstance(disposition, dict)
            and disposition.get("production") in ("candidate", "baseline"),
            "production disposition receipt is missing or invalid")
    before = read_json(P / "before-qualification.json")
    require(before.get("status") == "pass" and before.get("inherited") == "0728 final oracle",
            "before qualification did not pass the inherited oracle")
    runs = before.get("runs")
    require(isinstance(runs, list) and len(runs) == 4,
            "before qualification process count changed")
    require(all(integer(row.get("oracles"), "before oracle count", 1) >= 1
                and integer(row.get("negative_controls"), "before controls", 1) >= 1
                for row in runs), "before qualification controls are incomplete")
    # Final qualification is a required packet receipt.  Its internal schema
    # intentionally remains owned by the qualification workflow; this gate
    # only accepts a successful, non-empty receipt and binds its hash through
    # the freeze map.
    final = P / "qualification.json"
    require(final.is_file() and not final.is_symlink(), "final qualification receipt is missing")
    receipt = read_json(final)
    require(isinstance(receipt, dict), "final qualification receipt is not an object")
    require(receipt.get("status") == "pass", "final qualification did not pass")
    files = receipt.get("files")
    require(isinstance(files, dict) and bool(files), "final qualification files are empty")
    for relative, expected in files.items():
        relative = safe_relative(relative, "qualification file")
        require(digest(expected, f"qualification file {relative}") == sha(P / relative),
                f"qualification file changed: {relative}")
    require(receipt.get("candidate_builds_sha256") == sha(P / "candidate-builds.json"),
            "qualification candidate build hash changed")
    require(receipt.get("retention_builds_sha256") == sha(P / "retention-builds.json"),
            "qualification retention build hash changed")
    smoke_files = [relative for relative in files
                   if "smoke" in Path(relative).name.lower() and relative.endswith(".json")]
    require(smoke_files, "qualification omits candidate smoke receipt")
    for relative in smoke_files:
        smoke = read_json(P / relative)
        verify_smoke_receipt(smoke, cases, contract, builds, retention)
    quality_manifests = [relative for relative in files
                         if relative.startswith("quality-") and relative.endswith("/manifest.json")]
    require(quality_manifests, "qualification omits quality manifest")
    expected_quality = {relative: digest(value, f"candidate source {relative}")
                        for relative, value in builds["candidate"]["source_sha256"].items()
                        if relative.startswith("crates/")
                        and Path(relative).suffix in {".rs", ".toml"}}
    for relative in quality_manifests:
        manifest = read_json(P / relative)
        require(isinstance(manifest, dict), f"quality manifest is not an object: {relative}")
        runs = manifest.get("runs")
        require(isinstance(runs, list) and len(runs) == 13,
                f"quality command count changed: {relative}")
        require(all(row.get("exit_code") == 0 for row in runs),
                f"quality command failed: {relative}")
        require([quality_command_key(row.get("command")) for row in runs]
                == expected_quality_commands(),
                f"quality command matrix changed: {relative}")
        require(manifest.get("source_sha256") == expected_quality,
                f"quality source custody changed: {relative}")
    retention = P / "retention-analysis.json"
    require(retention.is_file() and not retention.is_symlink(),
            "retention analysis receipt is missing")
    retention_value = read_json(retention)
    require(isinstance(retention_value, dict)
            and isinstance(retention_value.get("processes"), list)
            and len(retention_value["processes"]) == 24,
            "retention analysis process count changed")


def safe_capture_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label} path is invalid")
    relative = Path(raw)
    require(not relative.is_absolute() and ".." not in relative.parts,
            f"{label} escapes capture directory")
    resolved = (CAPTURES / relative).resolve()
    root = CAPTURES.resolve()
    require(root == resolved or root in resolved.parents,
            f"{label} is outside capture directory")
    require(resolved.is_file() and not resolved.is_symlink(),
            f"{label} is missing or symlinked")
    return resolved


def expected_runs(plan: dict[str, Any], cases: dict[str, dict[str, Any]],
                  builds: dict[str, dict[str, Any]], before: bool = False) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    if before:
        cells = [(case, "baseline", 0) for case in CASES]
        lanes = [("native", plan["samples"], plan["warmups"]),
                 ("allocation", 1, 0)]
        for lane, samples, warmups in lanes:
            binary = next(row["path"] for row in builds["baseline"]["binaries"]
                          if Path(row["path"]).name == f"baseline-ole_format_save_probe"
                          + ("_alloc" if lane == "allocation" else ""))
            for case, variant, slot in cells:
                output = f"{lane}-c0-{case}-{slot}-{variant}.json"
                result.append({"lane": lane, "cycle": 0, "case": case,
                               "variant": variant, "slot": slot,
                               "output": output, "stderr": output + ".stderr",
                               "samples": samples, "warmups": warmups,
                               "command": ["taskset", "-c", str(plan["cpu"]), binary,
                                           "--case", case, "--input", cases[case]["path"],
                                           "--operation", "format", "--samples", str(samples),
                                           "--warmups", str(warmups)]})
        return result
    for cycle in range(plan["cycles"]):
        order = CASES if cycle % 2 == 0 else tuple(reversed(CASES))
        for case in order:
            for slot, variant in enumerate(NATIVE_SLOTS):
                binary = next(row["path"] for row in builds[variant]["binaries"]
                              if Path(row["path"]).name == f"{variant}-ole_format_save_probe")
                output = f"native-c{cycle}-{case}-{slot}-{variant}.json"
                result.append({"lane": "native", "cycle": cycle, "case": case,
                               "variant": variant, "slot": slot, "output": output,
                               "stderr": output + ".stderr", "samples": plan["samples"],
                               "warmups": plan["warmups"],
                               "command": ["taskset", "-c", str(plan["cpu"]), binary,
                                           "--case", case, "--input", cases[case]["path"],
                                           "--operation", "format", "--samples",
                                           str(plan["samples"]), "--warmups", str(plan["warmups"]) ]})
    for case in CASES:
        for slot, variant in enumerate(ALLOCATION_SLOTS):
            binary = next(row["path"] for row in builds[variant]["binaries"]
                          if Path(row["path"]).name == f"{variant}-ole_format_save_probe_alloc")
            output = f"allocation-c0-{case}-{slot}-{variant}.json"
            result.append({"lane": "allocation", "cycle": 0, "case": case,
                           "variant": variant, "slot": slot, "output": output,
                           "stderr": output + ".stderr", "samples": 1, "warmups": 0,
                           "command": ["taskset", "-c", str(plan["cpu"]), binary,
                                       "--case", case, "--input", cases[case]["path"],
                                       "--operation", "format", "--samples", "1",
                                       "--warmups", "0"]})
    return result


def oracle_guard(value: Any, label: str) -> None:
    require(isinstance(value, dict), f"{label}: oracle is not an object")
    require(value.get("oracle_ok") is True and value.get("failure_reasons") == [],
            f"{label}: oracle failed")
    def visit(item: Any, path: str) -> None:
        if isinstance(item, bool):
            require(item, f"{label}: oracle boolean {path} is false")
        elif isinstance(item, dict):
            for key, nested in item.items():
                visit(nested, f"{path}.{key}")
        elif isinstance(item, list):
            for index, nested in enumerate(item):
                visit(nested, f"{path}[{index}]")

    visit(value, "oracle")


def stats(values: list[int]) -> dict[str, int | float]:
    require(values, "empty sample vector")
    ordered = sorted(values)
    return {
        "n": len(values),
        "p50": statistics.median(ordered),
        "mean": math.fsum(values) / len(values),
        "p95": ordered[math.ceil(0.95 * len(ordered)) - 1],
        "p99": ordered[math.ceil(0.99 * len(ordered)) - 1],
        "maximum": ordered[-1],
    }


def validate_report(report: Any, row: dict[str, Any], case: dict[str, Any],
                    contract: dict[str, Any], identities: dict[str, dict[str, Any]],
                    output_hashes: dict[str, str]) -> tuple[dict[str, int | float], dict[str, int]]:
    require(isinstance(report, dict), f"{row}: report is not an object")
    label = f"{row['lane']}/c{row['cycle']}/{row['case']}/{row['slot']}/{row['variant']}"
    native = row["lane"] == "native"
    expected_header = {
        "schema_version": 1, "case": row["case"], "format": "doc",
        "operation": "format", "scope": "public_format_open_edit_commit",
        "input": case["path"], "policy": "reuse", "policy_applied": False,
        "policy_application_scope": "not_applied_public_format_route",
        "policy_argument_effect": "ignored_public_format_default_route",
        "timing_claim": native, "allocator_instrumented": not native,
        "directory_metadata_fields": META_FIELDS,
        "allocation_ownership_contract": OWNERSHIP_CONTRACT,
        "warmups": row["warmups"], "samples_requested": row["samples"],
        "source_sha256": case["sha256"],
    }
    for key, expected in expected_header.items():
        require(report.get(key) == expected, f"{label}: header {key} changed")
    require(isinstance(report.get("policy_contract"), str) and report["policy_contract"],
            f"{label}: policy contract missing")
    expected_identity = contract["identity"]
    identity_keys = ("source_sha256", "expected_output_sha256", "replacements_sha256",
                     "source_inventory", "expected_output_inventory", "replacements",
                     "changed_length_proof")
    identity = {key: report.get(key) for key in identity_keys}
    require(identity == expected_identity, f"{label}: frozen oracle identity changed")
    require(identities.setdefault(row["case"], identity) == identity,
            f"{label}: case identity drifted")
    expected_oracle = report.get("expected_oracle")
    oracle_guard(expected_oracle, f"{label}: expected")
    require(expected_oracle.get("semantic_witness") == contract["semantic_witness"],
            f"{label}: expected semantic witness changed")
    require(expected_oracle.get("directory_metadata_differences") == [],
            f"{label}: expected metadata differs")
    proof = report.get("changed_length_proof")
    require(isinstance(proof, dict)
            and proof.get("format_specific_semantic_length_proven") is True
            and proof.get("logical_stream_length_change_proven") is True
            and proof.get("any_stream_length_changed") is True
            and integer(proof.get("changed_stream_count"), f"{label}: changed stream count", 1) >= 1,
            f"{label}: changed-length proof failed")
    controls = report.get("oracle_controls")
    require(isinstance(controls, list)
            and [item.get("name") for item in controls] == contract["control_names"],
            f"{label}: control names changed")
    for control in controls:
        require(isinstance(control, dict) and control.get("status") == "rejected"
                and control.get("rejected") is True
                and isinstance(control.get("failure_reasons"), list)
                and bool(control["failure_reasons"]),
                f"{label}: negative control was not rejected with a reason")
    samples = report.get("samples")
    require(isinstance(samples, list) and len(samples) == row["samples"],
            f"{label}: sample count changed")
    output_expected = report.get("expected_output_inventory")
    require(isinstance(output_expected, dict), f"{label}: expected inventory missing")
    output_hash: str | None = None
    timings: list[int] = []
    allocation: dict[str, int] | None = None
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict) and sample.get("index") == index,
                f"{label}: sample index changed")
        oracle_guard(sample.get("oracle"), f"{label}: sample {index}")
        require(sample["oracle"].get("semantic_witness") == contract["semantic_witness"],
                f"{label}: sample semantic witness changed")
        raw = sample["oracle"].get("raw_directory")
        require(isinstance(raw, dict)
                and raw.get("source_expected_difference_bytes") == 0
                and raw.get("expected_output_difference_bytes") == 0
                and raw.get("source_output_difference_bytes") == 0,
                f"{label}: public raw-directory model gate failed")
        require(sample.get("output_inventory") == output_expected,
                f"{label}: output inventory differs from expected")
        sample_hash = digest(sample.get("output_sha256"), f"{label}: output hash")
        require(sample_hash == report["expected_output_sha256"],
                f"{label}: output differs from expected")
        output_hash = output_hash or sample_hash
        require(sample_hash == output_hash, f"{label}: output hash drifted")
        if native:
            require("allocations" not in sample and isinstance(sample.get("phase_ns"), dict),
                    f"{label}: native sample shape changed")
            require(set(sample["phase_ns"]) == {"whole_ns"}, f"{label}: phase shape changed")
            timings.append(integer(sample["phase_ns"]["whole_ns"], f"{label}: whole_ns", 1))
        else:
            require("phase_ns" not in sample and isinstance(sample.get("allocations"), dict),
                    f"{label}: allocation sample shape changed")
            actual = sample["allocations"]
            require(set(actual) == {"whole"} and isinstance(actual["whole"], dict),
                    f"{label}: allocation shape changed")
            whole = actual["whole"]
            require(set(whole) == set(ALLOCATION_FIELDS), f"{label}: allocation fields changed")
            values = {key: integer(whole[key], f"{label}: allocation {key}")
                      for key in ALLOCATION_FIELDS}
            allocation = allocation or values
            require(values == allocation, f"{label}: allocation repeated values drifted")
    require(output_hash is not None, f"{label}: output hash missing")
    output_hashes.setdefault(row["case"], output_hash)
    require(output_hashes[row["case"]] == output_hash, f"{label}: case output drifted")
    return stats(timings) if native else {}, allocation or {}


def verify_manifest(path: Path, expected: list[dict[str, Any]], cases: dict[str, dict[str, Any]],
                    contract: dict[str, Any], root_captures: bool = True) -> list[dict[str, Any]]:
    manifest = read_json(path)
    require(manifest.get("status") == "complete", f"{path.name}: capture is not complete")
    runs = manifest.get("runs")
    require(isinstance(runs, list) and len(runs) == len(expected),
            f"{path.name}: process count changed")
    seen_files: set[str] = set()
    identities: dict[str, dict[str, Any]] = {}
    output_hashes: dict[str, str] = {}
    rows: list[dict[str, Any]] = []
    old_captures = CAPTURES
    try:
        # The helper always uses CAPTURES.  The before packet is checked by
        # temporarily swapping that scope without weakening path validation.
        globals()["CAPTURES"] = path.parent
        for run, planned in zip(runs, expected):
            require(isinstance(run, dict), "capture run is not an object")
            for key in ("lane", "cycle", "case", "variant", "slot"):
                require(run.get(key) == planned[key], f"capture order changed at {planned}: {key}")
            require(run.get("exit_code") == 0, f"child process failed at {planned}")
            require(run.get("command") == planned["command"], f"command changed at {planned}")
            output = safe_capture_path(run.get("output"), f"{planned}: output")
            stderr_name = run.get("stderr", planned["stderr"])
            stderr = safe_capture_path(stderr_name, f"{planned}: stderr")
            output_rel = output.relative_to(path.parent.resolve()).as_posix()
            stderr_rel = stderr.relative_to(path.parent.resolve()).as_posix()
            require(output_rel not in seen_files and stderr_rel not in seen_files,
                    f"capture raw path reused at {planned}")
            require(run.get("output") == planned["output"]
                    and stderr_name == planned["stderr"],
                    f"capture output name changed at {planned}")
            require(run.get("sha256") == sha(output), f"output hash changed at {planned}")
            require(run.get("stderr_sha256") == sha(stderr), f"stderr hash changed at {planned}")
            report = read_json(output)
            timing, allocation = validate_report(report, planned, cases[planned["case"]],
                                                  contract[planned["case"]], identities,
                                                  output_hashes)
            rows.append({"lane": planned["lane"], "cycle": planned["cycle"],
                         "case": planned["case"], "variant": planned["variant"],
                         "slot": planned["slot"], "output_sha256": output_hashes[planned["case"]],
                         "timing": timing, "allocations": allocation})
            seen_files.update((output_rel, stderr_rel))
        on_disk = {p.relative_to(path.parent.resolve()).as_posix() for p in path.parent.rglob("*")
                   if p.is_file() and p.name != "manifest.json"}
        require(on_disk == seen_files, f"{path.name}: unmanifested capture files")
    finally:
        globals()["CAPTURES"] = old_captures
    return rows


def comparison(first: float, second: float) -> dict[str, Any]:
    delta = second - first
    percent = None if first == 0 else delta / abs(first) * 100.0
    return {"first": first, "second": second, "delta": delta,
            "percent": percent, "flag_gt_5pct": percent is None and delta != 0
            or percent is not None and abs(percent) > 5.0}


def pair_rows(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    pairs = (("AA", 0, 1), ("AB", 2, 3), ("BA", 5, 4))
    for case in CASES:
        for cycle in range(3):
            selected = {(row["slot"]): row for row in rows
                        if row["lane"] == "native" and row["case"] == case
                        and row["cycle"] == cycle}
            require(set(selected) == set(range(6)), f"native pair matrix incomplete: {case}/{cycle}")
            for name, first_slot, second_slot in pairs:
                first = selected[first_slot]["timing"]
                second = selected[second_slot]["timing"]
                differences = {key: comparison(float(first[key]), float(second[key]))
                               for key in STAT_FIELDS}
                output.append({"case": case, "cycle": cycle, "pair": name,
                               "first_slot": first_slot, "second_slot": second_slot,
                               "first_variant": selected[first_slot]["variant"],
                               "second_variant": selected[second_slot]["variant"],
                               "differences": differences,
                               "flags_gt_5pct": [key for key, value in differences.items()
                                                  if value["flag_gt_5pct"]]})
    return output


def allocation_comparisons(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    output: list[dict[str, Any]] = []
    for case in CASES:
        selected = {row["slot"]: row for row in rows
                    if row["lane"] == "allocation" and row["case"] == case}
        require(set(selected) == set(range(4)), f"allocation matrix incomplete: {case}")
        fields: dict[str, Any] = {}
        for key in ALLOCATION_FIELDS:
            baseline = [selected[0]["allocations"][key], selected[3]["allocations"][key]]
            candidate = [selected[1]["allocations"][key], selected[2]["allocations"][key]]
            fields[key] = {"baseline_slots": baseline, "candidate_slots": candidate,
                           "baseline_mean": math.fsum(baseline) / 2,
                           "candidate_mean": math.fsum(candidate) / 2}
            fields[key].update(comparison(fields[key]["baseline_mean"],
                                           fields[key]["candidate_mean"]))
        output.append({"case": case, "slots": [0, 1, 2, 3], "fields": fields,
                       "flags_gt_5pct": [key for key, value in fields.items()
                                          if value["flag_gt_5pct"]]})
    return output


def main() -> None:
    plan = read_plan()
    cases = read_cases(plan)
    contract = read_contract(cases)
    builds = {"baseline": read_build("baseline"), "candidate": read_build("candidate")}
    retention_build = read_retention_build()
    verify_binaries(builds, retention_build)
    read_freeze(builds, retention_build)
    verify_ancestry()
    verify_source_guard(builds)
    verify_qualification(builds, retention_build, cases, contract)
    before_expected = expected_runs(plan, cases, builds, before=True)
    before_rows = verify_manifest(BEFORE / "manifest.json", before_expected, cases, contract)
    expected = expected_runs(plan, cases, builds)
    rows = verify_manifest(CAPTURES / "manifest.json", expected, cases, contract)
    require(len(before_rows) == 4 and len(rows) == 44,
            "capture process totals changed")
    output = {
        "packet": "change-0730",
        "disposition": "bounded validated-render handoff paired measurement; no adoption claim",
        "matrix": {"before_processes": len(before_rows), "native_processes": 36,
                   "allocation_processes": 8, "cases": list(CASES),
                   "native_slots": list(NATIVE_SLOTS), "allocation_slots": list(ALLOCATION_SLOTS)},
        "statistics": {"quantiles": "nearest-rank p95/p99; midpoint p50",
                        "fields": list(STAT_FIELDS), "relative_threshold_percent": 5.0,
                        "timing_unit": "ns", "allocation_fields": list(ALLOCATION_FIELDS)},
        "before_processes": before_rows,
        "processes": rows,
        "native_pairs": pair_rows(rows),
        "allocation_comparisons": allocation_comparisons(rows),
    }
    (P / "analysis.json").write_text(json.dumps(output, indent=2, sort_keys=True) + "\n",
                                       encoding="utf-8")
    print("PASS 4 before + 36 native + 8 allocation processes; exact DOC oracles and custody")


if __name__ == "__main__":
    try:
        main()
    except (Failure, OSError, KeyError, TypeError, ValueError) as error:
        raise SystemExit(f"FAIL: {error}")
