#!/usr/bin/env python3
"""Independent custody and gate audit for the 0726 XLS empty-slot pilot.

This audit is deliberately separate from ``analyze.py``.  It reads the frozen
plan, source/build receipts, capture manifests and raw probe reports, then
recomputes semantic parity, timing, repeat, allocator and budget-fence gates.
It never invokes Cargo, a Rust/native probe, or a profiler.  ``--draft`` is a
pre-capture check; the default command requires terminal captures and (when
owned binaries have been removed) an explicit cleanup witness.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, Iterable


PACKET = Path(__file__).resolve().parent
ROOT = PACKET.parents[3]
PLAN = PACKET / "plan.json"
CAPTURES = PACKET / "captures"
BASELINE_REF = "3ad29e42da"
SOURCE_DELTA = {"crates/litchi-xls/src/workbook/source.rs"}
TOOL_NAMES = ("pilot.py", "analyze.py", "plan.json")
ALLOC_METRICS = {
    "allocation_calls",
    "allocated_bytes",
    "deallocation_calls",
    "deallocated_bytes",
    "peak_live_delta",
    "retained_live_delta",
}
HEX = set("0123456789abcdef")


class AuditError(Exception):
    """A terminal or pre-capture audit failure."""


class Pending(Exception):
    """A draft check completed before terminal evidence existed."""

    def __init__(self, items: Iterable[str]):
        self.items = tuple(dict.fromkeys(items))
        super().__init__("terminal capture evidence is pending")


def sha256_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha256_file(path: Path) -> str:
    if not path.is_file() or path.is_symlink():
        raise AuditError(f"missing or symlinked file: {path}")
    return sha256_bytes(path.read_bytes())


def read_json(path: Path) -> Any:
    if not path.is_file() or path.is_symlink():
        raise AuditError(f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AuditError(f"invalid JSON {path}: {error}") from error


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AuditError(message)


def digest(value: Any) -> str:
    encoded = (json.dumps(value, sort_keys=True, separators=(",", ":")) + "\n").encode()
    return sha256_bytes(encoded)


def require_sha(value: Any, label: str) -> str:
    check(isinstance(value, str) and len(value) == 64 and set(value) <= HEX,
          f"{label} is not a lowercase SHA-256 digest")
    return value


def source_paths(plan: dict[str, Any]) -> list[Path]:
    result: list[Path] = []
    for raw_root in plan["source_roots"]:
        result.extend(path for path in (ROOT / raw_root).rglob("*") if path.is_file())
    return sorted(result)


def relative(path: Path) -> str:
    return path.relative_to(ROOT).as_posix()


def current_source_map(plan: dict[str, Any]) -> dict[str, str]:
    return {relative(path): sha256_file(path) for path in source_paths(plan)}


def git_source_map(plan: dict[str, Any], revision: str) -> dict[str, str | None]:
    result: dict[str, str | None] = {}
    for path in source_paths(plan):
        name = relative(path)
        probe = subprocess.run(
            ["git", "cat-file", "-e", f"{revision}:{name}"],
            cwd=ROOT,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        if probe.returncode:
            result[name] = None
        else:
            result[name] = sha256_bytes(
                subprocess.check_output(["git", "show", f"{revision}:{name}"], cwd=ROOT)
            )
    return result


def probe_map(plan: dict[str, Any]) -> dict[str, str]:
    result: dict[str, str] = {}
    for raw_root in plan["probe_roots"]:
        for path in sorted((ROOT / raw_root).rglob("*")):
            if path.is_file():
                result[relative(path)] = sha256_file(path)
    return result


def corpus_map(plan: dict[str, Any]) -> dict[str, dict[str, int | str]]:
    result: dict[str, dict[str, int | str]] = {}
    cases = list(plan["cases"]) + list(plan["budget_fence"]["cases"])
    for case in cases:
        raw = str(case["path"])
        path = ROOT / raw
        check(path.is_file(), f"missing fixture: {path}")
        if raw not in result:
            result[raw] = {"bytes": path.stat().st_size, "sha256": sha256_file(path)}
    return result


def binary_path(plan: dict[str, Any], phase: str, kind: str) -> Path:
    return Path(plan["binary_root"]) / phase / plan["binaries"][kind]


def tool_map() -> dict[str, str]:
    return {name: sha256_file(PACKET / name) for name in TOOL_NAMES}


def load_plan() -> dict[str, Any]:
    value = read_json(PLAN)
    check(isinstance(value, dict), "plan is not an object")
    check(value.get("schema_version") == 1, "plan schema changed")
    check(value.get("change") == "0726-xls-empty-slot-setup", "plan identity changed")
    check(value.get("baseline_revision") == BASELINE_REF, "baseline revision changed")
    check(value.get("cpu") == 12, "CPU binding changed")
    check(value.get("capture_root") == "captures", "capture root changed")
    check(value.get("case_manifest") == "docs/performance/results/change-0726/cases.json",
          "case manifest binding changed")
    check(value.get("source_roots") == ["crates/litchi-cfb", "crates/litchi-xls"],
          "source roots changed")
    check(value.get("binaries") == {
        "native": "xls-index-retry-probe-0686",
        "repeat": "xls0684-repeat",
        "allocator": "xls0686-alloc",
        "budget": "xls-index-budget-probe-0684",
        "route": "xls-index-probe-0684",
    }, "binary set changed")
    native = value.get("native", {})
    check(native.get("queries") == 8 and native.get("warmups") == 3
          and native.get("samples") == 100, "native sample plan changed")
    check(native.get("legs") == ["aa1", "aa2", "a1", "b1", "b2", "a2"],
          "native leg order changed")
    check(native.get("groups") == 24, "native group plan changed")
    repeat = value.get("repeat", {})
    check(repeat.get("repetitions") == 50000 and repeat.get("samples_per_leg") == 9
          and repeat.get("preparatory_queries") == 2 and repeat.get("groups") == 16,
          "repeat plan changed")
    allocator = value.get("allocator", {})
    check(allocator.get("repeats") == 3 and allocator.get("operations") == ["q1", "q2", "q3", "q8"],
          "allocator operation plan changed")
    check(allocator.get("groups_per_phase") == 96
          and allocator.get("strict_operations") == ["q1", "q2", "q3", "q8"],
          "allocator scope changed")
    check(allocator.get("q2_allowance") == {
        "allocation_calls": 0,
        "allocated_bytes": 0,
        "deallocation_calls": 0,
        "deallocated_bytes": 0,
        "peak_live_delta": 0,
        "retained_live_delta": 0,
    }, "allocator q2 allowance changed")
    check(allocator.get("missing_warm_expected_deltas") == {
        "allocation_calls": -1,
        "allocated_bytes": -16,
        "deallocation_calls": -1,
        "deallocated_bytes": -16,
        "peak_live_delta": -16,
        "retained_live_delta": 0,
    }, "missing warm allocator deltas changed")
    check(len(value.get("cases", [])) == 12 and len(value.get("repeat_cases", [])) == 8,
          "case matrix changed")
    budget_primary = value.get("budget_primary", {})
    check(budget_primary.get("queries") == 8 and budget_primary.get("groups_per_phase") == 12,
          "budget primary plan changed")
    check(value.get("hard_gates") == {
        "timing_percent": 5.0,
        "timing_absolute_ns": 10.0,
        "absolute_exception_metrics": ["q3", "q8", "q3-to-q8-mean"],
        "outcomes_exact": True,
        "allocator_fixed_field": True,
        "budget_fence_metrics_exact": True,
        "zero_budget_and_refusal_parity": True,
    }, "hard gate thresholds changed")
    check(value.get("benefit_gates") == {
        "native": {
            "case": "54016-missing-1048576", "mode": "owned", "metric": "q8",
            "minimum_improvement_percent": 10.0,
            "statistics": ["p50", "mean"], "both_pairs": True,
        },
        "repeat": {
            "case": "54016-missing-1048576", "mode": "owned",
            "minimum_improvement_percent": 10.0,
            "statistics": ["p50", "mean"], "both_pairs": True,
        },
    }, "benefit gates changed")
    fences = value.get("budget_fence", {})
    check(fences.get("queries") == 8 and len(fences.get("cases", [])) == 4,
          "budget fence plan changed")
    return value


def check_source_static(plan: dict[str, Any]) -> dict[str, Any]:
    """Check the one-file empty-slot change and its retained safety fences.

    This is intentionally source-level evidence.  The 0726 candidate moves
    the indexed-slot lookup ahead of path/resolver setup and adds only the
    empty-result return; the fixed cache charge, checkpoint representation,
    scanner, and tests remain baseline code.
    """
    current = current_source_map(plan)
    baseline = git_source_map(plan, plan["baseline_revision"])
    changed = {name for name in set(current) | set(baseline) if current.get(name) != baseline.get(name)}
    check(changed == SOURCE_DELTA, f"source delta changed: {sorted(changed)}")

    candidate_record = read_json(PACKET / "candidate-source.json")
    check(candidate_record == {name: current[name] for name in sorted(SOURCE_DELTA)},
          "candidate-source.json is not the exact changed-source map")
    archived_sources: dict[str, str] = {}
    for name in SOURCE_DELTA:
        archived = PACKET / "candidate-source" / name
        check(archived.is_file() and sha256_file(archived) == current[name],
              f"candidate source archive changed: {name}")
        archived_sources[name] = sha256_file(archived)

    query_path = ROOT / "crates/litchi-xls/src/workbook/query_cache.rs"
    source_path = ROOT / "crates/litchi-xls/src/workbook/source.rs"
    query_text = query_path.read_text(encoding="utf-8")
    source_text = source_path.read_text(encoding="utf-8")
    baseline_query = subprocess.check_output(
        ["git", "show", f"{plan['baseline_revision']}:crates/litchi-xls/src/workbook/query_cache.rs"],
        cwd=ROOT,
        text=True,
    )
    check(query_text == baseline_query,
          "query_cache.rs changed even though the candidate delta is source.rs only")
    check(re.search(r"INDEX_OVERHEAD:\s*u64\s*=\s*224\s*;", query_text) is not None,
          "candidate fixed index charge is not 224 bytes")
    check(re.search(r"INDEX_OVERHEAD:\s*u64\s*=\s*224\s*;", baseline_query) is not None,
          "baseline fixed index charge is not 224 bytes")
    query_code = re.sub(r"/\*.*?\*/", "", query_text, flags=re.S)
    query_code = re.sub(r"//[^\n]*", "", query_code)
    for forbidden in ("CellValue", "SourceBackedError", "String"):
        check(forbidden not in query_code,
              f"query cache retains decoded value/error/string token {forbidden}")

    archived_source = (PACKET / "candidate-source" / "crates/litchi-xls/src/workbook/source.rs").read_text(
        encoding="utf-8"
    )
    check(source_text == archived_source, "live candidate source differs from archived source bytes")
    replay_start = source_text.find("fn replay_indexed_cell(")
    check(replay_start >= 0, "indexed replay function is missing")
    replay_end = source_text.find("\n}\n\n/// Reads and decodes one indexed occurrence.", replay_start)
    check(replay_end > replay_start, "indexed replay function boundary changed")
    replay = source_text[replay_start:replay_end]

    worksheet_match = re.search(r"let sheet = owner\s*\.sheets\s*\.get\(sheet_index\).*?WorksheetNotFound", replay, re.S)
    slots_match = re.search(r"let slots = index\.slots_for\(row, column\);", replay)
    refs_match = re.search(r"let refs = owner\s*\.workbook_path", replay)
    check(worksheet_match is not None, "replay does not retain worksheet lookup/error")
    check(slots_match is not None and refs_match is not None
          and slots_match.start() < refs_match.start(),
          "slots_for is not before path/resolver setup")
    empty_start = replay.find("if slots.is_empty()", slots_match.start() if slots_match else 0)
    check(empty_start > slots_match.start(), "missing-target empty-slot branch is missing")
    empty_branch = replay[empty_start:refs_match.start()]
    check("context.check()" in empty_branch and "owner.ensure_current()" in empty_branch
          and "return Ok(None)" in empty_branch,
          "empty-slot return bypasses execution/version fencing")
    check("workbook_path" not in empty_branch and "stream_cursor" not in empty_branch
          and "SharedStringResolver" not in empty_branch and "Vec" not in empty_branch,
          "empty-slot branch performs retained path/resolver setup")
    loop_match = re.search(r"for slot in slots\s*\{", replay)
    check(loop_match is not None, "indexed replay does not iterate the selected slots")
    check("for slot in index.slots_for" not in replay,
          "indexed replay still performs slot lookup inside the loop")
    tail = replay[loop_match.end():]
    loop_end = tail.rfind("\n    if let Some(context) = execution")
    check(loop_end >= 0, "indexed replay lost final execution fence")
    final_fence = tail[loop_end:]
    check("context.check()" in final_fence and "owner.ensure_current()?" in final_fence
          and "Ok(found)" in final_fence,
          "indexed replay lost final context/version fence")

    # The baseline already owns the trailing-fence integration control.  It is
    # checked by name so a future source-only change cannot silently remove it
    # or replace the public workbook call with a private replay mock.
    cache_test_text = (ROOT / "crates/litchi-xls/tests/xls_query_index_cache.rs").read_text(encoding="utf-8")
    for token in (
        "fn an_indexed_missing_target_still_takes_the_trailing_freshness_fence()",
        "bump_before_observation(2)",
        "assert_eq!(source.read_count(), 0)",
        "worksheet.cell_value(50_000, 7)",
    ):
        check(token in cache_test_text, f"existing cache test lost trailing-fence control: {token}")
    return {
        "source_delta": sorted(changed),
        "baseline_index_overhead": 224,
        "candidate_index_overhead": 224,
        "candidate_source_archive": True,
        "query_cache_unchanged": True,
        "indexed_missing_route": "worksheet lookup; slots_for; fenced Ok(None) before refs",
        "final_replay_fence": True,
        "trailing_fence_test": "xls_query_index_cache::an_indexed_missing_target_still_takes_the_trailing_freshness_fence",
    }


def check_builds(plan: dict[str, Any]) -> dict[str, Any]:
    baseline = git_source_map(plan, plan["baseline_revision"])
    candidate = current_source_map(plan)
    output: dict[str, Any] = {}
    for phase, expected_source in (("baseline", baseline), ("candidate", candidate)):
        path = PACKET / f"{phase}-builds.json"
        rows = read_json(path)
        check(isinstance(rows, list) and len(rows) == len(plan["binaries"]),
              f"{phase}-builds.json row count changed")
        seen: set[str] = set()
        phase_hashes: dict[str, str] = {}
        for row in rows:
            check(isinstance(row, dict) and row.get("exit_code") == 0,
                  f"{phase} build row failed")
            name = row.get("binary")
            check(name in plan["binaries"].values() and name not in seen,
                  f"{phase} build identity is missing/duplicated: {name}")
            seen.add(name)
            # A baseline build predates the untracked candidate-only test, so
            # its census omits paths whose git-baseline value is ``None``.
            # build.py records Rust inputs only.  Freeze/source manifests also
            # include the package metadata and documentation files so that the
            # complete checkout is bound; compare the build receipt at its
            # narrower, explicit Rust census here.
            build_source = {
                name: value for name, value in expected_source.items()
                if value is not None and name.endswith(".rs")
            }
            check(row.get("source_sha256") == build_source,
                  f"{phase} {name} source census differs")
            binary = next(kind for kind, binary_name in plan["binaries"].items() if binary_name == name)
            path_binary = binary_path(plan, phase, binary)
            expected_hash = row.get("binary_sha256")
            require_sha(expected_hash, f"{phase}/{binary} build binary")
            if path_binary.is_file():
                check(not path_binary.is_symlink() and sha256_file(path_binary) == expected_hash,
                      f"{phase}/{binary} build binary hash differs")
            phase_hashes[binary] = expected_hash
        check(seen == set(plan["binaries"].values()), f"{phase} build set is incomplete")
        output[phase] = phase_hashes
    return output


def check_negative_checks() -> dict[str, Any]:
    """Bind the synthetic verifier receipt to the actual analyzer helpers.

    This receipt is a tooling check, not performance evidence.  Requiring its
    analyzer digest to match the current file prevents a later scoring change
    from silently invalidating the recorded positive and negative controls.
    """
    script_path = PACKET / "negative-checks.py"
    receipt_path = PACKET / "negative-checks.json"
    check(script_path.is_file() and not script_path.is_symlink(),
          "negative-checks.py is missing or symlinked")
    script = script_path.read_text(encoding="utf-8")
    for token in ("A.timing_gate(", "A.compare_native(", "A.compare_allocator(",
                  "A.compare_budget_fence(", "change-0690", "TemporaryDirectory"):
        check(token in script, f"negative-checks.py does not exercise {token}")
    receipt = read_json(receipt_path)
    check(receipt.get("scope") == "synthetic verifier checks, not performance evidence",
          "negative-checks receipt scope changed")
    check(receipt.get("analysis_sha256") == sha256_file(PACKET / "analyze.py"),
          "negative-checks receipt is not bound to current analyze.py")
    expected_names = [
        "warm absolute allowance",
        "warm regression rejected",
        "build has no warm allowance",
        "workflow regression rejected",
        "native positive control",
        "actual outcome mutation rejected",
        "actual timing mutation rejected",
        "allocator positive control",
        "strict allocation delta 1 rejected",
        "strict allocation delta -1 rejected",
        "build allowance boundary 0",
        "build allowance boundary 1",
        "build allowance boundary -1",
        "missing warm exact removal accepted",
        "missing warm unremoved allocation rejected",
        "budget fence exact metrics accepted",
        "budget fence changed metrics rejected",
    ]
    checks = receipt.get("checks")
    check(isinstance(checks, list) and len(checks) == len(expected_names),
          "negative-checks receipt count changed")
    check([item.get("name") for item in checks] == expected_names,
          "negative-checks receipt names changed")
    check(all(item.get("pass") is True for item in checks),
          "negative-checks receipt contains a failed control")
    return {
        "script_sha256": sha256_file(script_path),
        "receipt_sha256": sha256_file(receipt_path),
        "analysis_sha256": receipt["analysis_sha256"],
        "checks": len(checks),
    }


def collect_witnesses() -> dict[str, str]:
    result: dict[str, str] = {}
    path = PACKET / "cleanup.json"
    if not path.is_file():
        return result
    value = read_json(path)

    def walk(item: Any) -> None:
        if isinstance(item, dict):
            raw_path = item.get("path")
            expected = item.get("sha256", item.get("binary_sha256"))
            if isinstance(raw_path, str) and isinstance(expected, str) and len(expected) == 64:
                result[str(Path(raw_path).resolve())] = expected
            # Accept the compact historical receipt form
            # {"/owned/binary": "<sha256>"} as well as path/sha fields.
            for key, child in item.items():
                if (isinstance(key, str) and key.startswith("/")
                        and isinstance(child, str) and len(child) == 64
                        and set(child) <= HEX):
                    result[str(Path(key).resolve())] = child
                walk(child)
        elif isinstance(item, list):
            for child in item:
                walk(child)

    walk(value)
    return result


def verify_artifact(path: Path, expected: str, witnesses: dict[str, str], label: str) -> str:
    require_sha(expected, f"{label} hash")
    if path.is_file():
        check(not path.is_symlink() and sha256_file(path) == expected, f"{label} hash differs")
        return "live"
    check(witnesses.get(str(path.resolve())) == expected,
          f"{label} is absent without an exact cleanup witness")
    return "cleanup-witness"


def audit_manifest(
    plan: dict[str, Any], task: str, required_kind: str, current: dict[str, str],
    baseline: dict[str, str | None], probes: dict[str, str], corpus: dict[str, Any],
    frozen: dict[str, Any], witnesses: dict[str, str], draft: bool,
) -> dict[str, Any] | None:
    path = CAPTURES / task / "manifest.json"
    if not path.is_file():
        if draft:
            return None
        raise AuditError(f"missing {task} manifest")
    manifest = read_json(path)
    check(manifest.get("task") == task, f"{task} manifest task changed")
    check(manifest.get("status") == "complete", f"{task} manifest is not complete")
    check(manifest.get("schema_version") == 1 and manifest.get("packet") == "change-0726",
          f"{task} manifest schema/packet changed")
    check(manifest.get("baseline_revision") == plan["baseline_revision"],
          f"{task} baseline revision label changed")
    resolved = subprocess.check_output(
        ["git", "rev-parse", plan["baseline_revision"]], cwd=ROOT, text=True
    ).strip()
    check(manifest.get("baseline_revision_resolved") == resolved,
          f"{task} baseline revision resolution changed")
    check(manifest.get("cpu") == plan["cpu"] and manifest.get("repository_root") == str(ROOT),
          f"{task} execution binding changed")
    check(manifest.get("plan_sha256") == sha256_file(PLAN), f"{task} plan binding differs")
    check(manifest.get("freeze_sha256") == sha256_file(CAPTURES / "freeze.json"),
          f"{task} freeze binding differs")
    check(manifest.get("tools_sha256_start") == frozen.get("tools_sha256_start"),
          f"{task} tool start binding differs")
    check(manifest.get("tools_sha256_end") == manifest.get("tools_sha256_start"),
          f"{task} tooling changed during capture")
    check(manifest.get("source_sha256_start") == current
          and manifest.get("source_sha256_end") == current,
          f"{task} source capture is not the frozen candidate")
    check(manifest.get("baseline_source_sha256") == baseline,
          f"{task} baseline source map differs from git")
    check(manifest.get("probe_sha256") == probes and manifest.get("probe_sha256_end") == probes,
          f"{task} probe map differs")
    check(manifest.get("corpus") == corpus and manifest.get("corpus_end") == corpus,
          f"{task} fixture map differs")
    case_manifest_path = ROOT / plan["case_manifest"]
    case_manifest_sha = sha256_file(case_manifest_path)
    check(manifest.get("case_manifest_sha256") == case_manifest_sha
          and manifest.get("case_manifest_sha256_end") == case_manifest_sha,
          f"{task} case manifest differs")

    expected_config: dict[str, Any] = {
        "native": {
            "groups": plan["native"]["groups"], "cases": len(plan["cases"]),
            "modes": ["owned", "file"], "legs": plan["native"]["legs"],
        },
        "repeat": {
            "groups": plan["repeat"]["groups"], "cases": plan["repeat_cases"],
            "modes": ["owned", "file"], "samples_per_leg": plan["repeat"]["samples_per_leg"],
            "legs": plan["native"]["legs"],
        },
        "allocator": {
            "groups_per_phase": plan["allocator"]["groups_per_phase"], "cases": len(plan["cases"]),
            "modes": ["owned", "file"], "operations": plan["allocator"]["operations"],
            "repeats": plan["allocator"]["repeats"],
        },
        "budget-primary": {
            "cases": len(plan["cases"]), "queries": plan["budget_primary"]["queries"],
            "source_mode": "counted-owned",
        },
        "budget-fence": {
            "cases": len(plan["budget_fence"]["cases"]),
            "budgets": sorted({budget for case in plan["budget_fence"]["cases"] for budget in case["budgets"]}),
            "queries": plan["budget_fence"]["queries"], "source_mode": "counted-owned",
        },
    }
    check(manifest.get("config") == expected_config[task], f"{task} capture config changed")

    binaries = manifest.get("binaries")
    check(isinstance(binaries, dict), f"{task} binary receipt is missing")
    for phase in ("baseline", "candidate"):
        phase_rows = binaries.get(phase)
        check(isinstance(phase_rows, dict), f"{task} {phase} binary receipt is missing")
        for kind, name in plan["binaries"].items():
            row = phase_rows.get(kind)
            check(isinstance(row, dict) and row.get("name") == name,
                  f"{task} {phase}/{kind} binary identity changed")
            expected_path = binary_path(plan, phase, kind)
            check(row.get("path") == str(expected_path),
                  f"{task} {phase}/{kind} binary path changed")
            available = expected_path.is_file()
            check(row.get("available") == available or (not available and row.get("available") is True),
                  f"{task} {phase}/{kind} availability receipt changed")
            # A receipt may say available=true even after terminal cleanup;
            # the recorded hash remains authoritative and must match either
            # the live binary or the exact witness.
            verify_artifact(expected_path, row.get("sha256"), witnesses, f"{task} {phase}/{kind} binary")

    build_receipts = manifest.get("build_manifests")
    check(isinstance(build_receipts, dict), f"{task} build receipts are missing")
    for phase in ("baseline", "candidate"):
        receipt = build_receipts.get(phase)
        check(isinstance(receipt, dict), f"{task} {phase} build receipt is missing")
        build_path = Path(str(receipt.get("path")))
        check(build_path.is_file() and not build_path.is_symlink(), f"{task} {phase} build manifest is missing")
        check(receipt.get("sha256") == sha256_file(build_path), f"{task} {phase} build receipt hash differs")
        check(read_json(build_path) == receipt.get("records"), f"{task} {phase} build receipt content differs")

    commands = manifest.get("commands")
    check(isinstance(commands, list) and commands, f"{task} command receipt is empty")
    check(all(item.get("exit_code") == 0 for item in commands), f"{task} contains failed command")
    raw = manifest.get("raw_sha256")
    check(isinstance(raw, dict) and raw, f"{task} raw hash map is empty")
    for raw_name, expected in raw.items():
        raw_path = Path(raw_name)
        check(not raw_path.is_absolute() and ".." not in raw_path.parts,
              f"{task} raw path is unsafe: {raw_name}")
        path_raw = ROOT / raw_path
        check(path_raw.is_file() and sha256_file(path_raw) == expected,
              f"{task} raw output hash differs: {raw_name}")
    expected_files = {str((CAPTURES / task / name).relative_to(ROOT))
                      for name in path_names(plan, task)}
    observed_files = {name for name in raw if name.endswith((".json", ".tsv"))}
    check(observed_files == expected_files, f"{task} raw output set differs")
    check(baseline_source_for_task(task, required_kind, plan), f"{task} required binary kind changed")
    return manifest


def baseline_source_for_task(task: str, required_kind: str, plan: dict[str, Any]) -> bool:
    return required_kind in plan["binaries"] and task in {"native", "repeat", "allocator", "budget-fence", "budget-primary"}


def path_names(plan: dict[str, Any], task: str) -> list[str]:
    cases = {str(case["case"]): case for case in plan["cases"]}
    legs = list(plan["native"]["legs"])
    if task == "native":
        return [f"{leg}-{case['case']}-{mode}.json"
                for leg in legs for case in plan["cases"] for mode in ("owned", "file")]
    if task == "repeat":
        return [f"{case}-{mode}-{leg}-{sample}.tsv"
                for leg in legs for case in plan["repeat_cases"] for mode in ("owned", "file")
                for sample in range(plan["repeat"]["samples_per_leg"])]
    if task == "allocator":
        return [f"{phase}-{case['case']}-{mode}-{operation}-{repeat}.json"
                for phase in ("baseline", "candidate") for case in plan["cases"]
                for mode in ("owned", "file") for operation in plan["allocator"]["operations"]
                for repeat in range(plan["allocator"]["repeats"])]
    if task == "budget-fence":
        return [f"{phase}-{case['case']}-{budget}.json"
                for phase in ("baseline", "candidate") for case in plan["budget_fence"]["cases"]
                for budget in case["budgets"]]
    if task == "budget-primary":
        return [f"{phase}-{case['case']}.json"
                for phase in ("baseline", "candidate") for case in plan["cases"]]
    raise AuditError(f"unknown capture task {task}")


def number(value: Any, label: str) -> float:
    check(isinstance(value, (int, float)) and not isinstance(value, bool), f"{label} is not numeric")
    result = float(value)
    check(math.isfinite(result) and result >= 0, f"{label} is negative/non-finite")
    return result


def stats(values: list[float]) -> dict[str, float | int]:
    check(values, "empty timing vector")
    ordered = sorted(values)
    return {
        "n": len(values),
        "p50": (ordered[(len(ordered) - 1) // 2] + ordered[len(ordered) // 2]) / 2,
        "mean": sum(values) / len(values),
        "p95": ordered[max(0, math.ceil(.95 * len(ordered)) - 1)],
        "p99": ordered[max(0, math.ceil(.99 * len(ordered)) - 1)],
        "minimum": min(values),
        "maximum": max(values),
    }


def timing_gate(metric: str, before: float, after: float, plan: dict[str, Any]) -> bool:
    change = percent_change(before, after)
    percent_ok = change <= float(plan["hard_gates"]["timing_percent"])
    absolute_ok = metric in set(plan["hard_gates"]["absolute_exception_metrics"]) and after - before <= float(plan["hard_gates"]["timing_absolute_ns"])
    return percent_ok or absolute_ok


def percent_change(before: float, after: float) -> float:
    if before == 0:
        return 0.0 if after == 0 else float("inf")
    return (after / before - 1.0) * 100.0


def outcome_projection(report: dict[str, Any]) -> list[Any]:
    return [query.get("outcome") for record in report["records"] for query in record["queries"]]


def validate_native_report(plan: dict[str, Any], case: dict[str, Any], mode: str,
                           report: dict[str, Any], label: str,
                           continue_on_rejection: bool = False) -> None:
    check(report.get("schema_version") == 1
          and report.get("probe") == "change-0686-xls-index-budget-retry",
          f"{label}: probe schema identity differs")
    check(report.get("mode") == mode, f"{label}: mode differs")
    path = ROOT / case["path"]
    check(report.get("input_sha256") == sha256_file(path), f"{label}: input hash differs")
    check(report.get("input_bytes") == path.stat().st_size, f"{label}: input bytes differ")
    check(Path(report.get("input_path", "")).resolve() == path.resolve(), f"{label}: input path differs")
    for field, expected in (("worksheet", case["sheet"]), ("row", case["row"]),
                            ("column", case["column"]), ("max_query_index_bytes", case["budget"]),
                            ("queries", plan["native"]["queries"]), ("warmups", plan["native"]["warmups"]),
                            ("samples", plan["native"]["samples"])):
        check(report.get(field) == expected, f"{label}: {field} differs")
    check(report.get("fresh_owner_per_sample") is True, f"{label}: owner reuse changed")
    records = report.get("records")
    check(isinstance(records, list) and len(records) == plan["native"]["samples"], f"{label}: sample count differs")
    expected_samples = list(range(plan["native"]["warmups"], plan["native"]["warmups"] + plan["native"]["samples"]))
    check([record.get("sample") for record in records] == expected_samples, f"{label}: sample ordinals differ")
    first: list[Any] | None = None
    for sample in records:
        open_record = sample.get("open", {})
        check(open_record.get("outcome", {}).get("status") == "ok", f"{label}: open failed")
        number(open_record.get("elapsed_ns"), f"{label}: open elapsed")
        queries = sample.get("queries")
        check(isinstance(queries, list) and len(queries) == plan["native"]["queries"], f"{label}: query count differs")
        check([query.get("ordinal") for query in queries] == list(range(plan["native"]["queries"])), f"{label}: query ordinals differ")
        outcomes = [query.get("outcome") for query in queries]
        check(sample.get("all_queries_agree") is True and all(query.get("agrees_with_first") is True for query in queries),
              f"{label}: repeated outcomes disagree")
        check(all(isinstance(item, dict) and item.get("status") in {"value", "missing", "error"} for item in outcomes),
              f"{label}: malformed semantic outcome")
        for index, query in enumerate(queries):
            number(query.get("elapsed_ns"), f"{label}: q{index + 1} elapsed")
        if first is None:
            first = outcomes
        else:
            if outcomes != first and not continue_on_rejection:
                check(False, f"{label}: sample semantic outcome differs")
    check(first is not None, f"{label}: no semantic outcomes")
    if {item["status"] for item in first} != {case["expected_status"]} and not continue_on_rejection:
        check(False, f"{label}: expected status differs")


def native_metrics(plan: dict[str, Any], report: dict[str, Any]) -> dict[str, list[float]]:
    records = report["records"]
    return {
        "open": [float(item["open"]["elapsed_ns"]) for item in records],
        "q1": [float(item["queries"][0]["elapsed_ns"]) for item in records],
        "q2": [float(item["queries"][1]["elapsed_ns"]) for item in records],
        "q3": [float(item["queries"][2]["elapsed_ns"]) for item in records],
        "q8": [float(item["queries"][7]["elapsed_ns"]) for item in records],
        "q3-to-q8-mean": [sum(float(query["elapsed_ns"]) for query in item["queries"][2:]) / 6.0 for item in records],
        "open-plus-eight": [float(item["open"]["elapsed_ns"]) + sum(float(query["elapsed_ns"]) for query in item["queries"]) for item in records],
    }


def audit_native(plan: dict[str, Any], continue_on_rejection: bool = False) -> tuple[list[dict[str, Any]], bool]:
    directory = CAPTURES / "native"
    cases = {case["case"]: case for case in plan["cases"]}
    rows: list[dict[str, Any]] = []
    passed = True
    for case_name, case in cases.items():
        for mode in ("owned", "file"):
            reports: dict[str, dict[str, Any]] = {}
            for leg in plan["native"]["legs"]:
                path = directory / f"{leg}-{case_name}-{mode}.json"
                report = read_json(path)
                validate_native_report(plan, case, mode, report, str(path), continue_on_rejection)
                reports[leg] = report
            first_outcomes = outcome_projection(reports[plan["native"]["legs"][0]])
            equal = all(outcome_projection(report) == first_outcomes for report in reports.values())
            if not equal and not continue_on_rejection:
                check(False, f"native/{case_name}/{mode}: semantic outcomes differ across legs")
            vectors = {leg: native_metrics(plan, report) for leg, report in reports.items()}
            paired: dict[str, dict[str, bool]] = {}
            paired_percent: dict[str, dict[str, dict[str, float]]] = {}
            for metric in plan["native"]["timing_metrics"]:
                pair_gates: dict[str, bool] = {}
                pair_percent: dict[str, dict[str, float]] = {}
                for pair, candidate_leg, baseline_leg in (("b1_a1", "b1", "a1"), ("b2_a2", "b2", "a2")):
                    base = stats(vectors[baseline_leg][metric])
                    candidate = stats(vectors[candidate_leg][metric])
                    pair_gates[pair] = all(timing_gate(metric, float(base[key]), float(candidate[key]), plan) for key in ("p50", "mean"))
                    pair_percent[pair] = {
                        key: percent_change(float(base[key]), float(candidate[key]))
                        for key in ("p50", "mean")
                    }
                paired[metric] = pair_gates
                paired_percent[metric] = pair_percent
                aa_base = stats(vectors["aa1"][metric])
                aa_after = stats(vectors["aa2"][metric])
                # A/A is a retained noise control.  It is reported and
                # checked for finite data, while the frozen hard gate remains
                # the two prospective B/A pairs defined by the plan.
                _ = timing_gate(metric, float(aa_base["p50"]), float(aa_after["p50"]), plan)
            group_pass = equal and all(value for metric in paired.values() for value in metric.values())
            passed &= group_pass
            rows.append({"case": case_name, "mode": mode, "outcome_equal": equal,
                         "paired_timing_pass": paired, "paired_percent": paired_percent,
                         "gate_pass": group_pass})
    check(len(rows) == plan["native"]["groups"], "native group count differs")
    return rows, passed


def parse_repeat(path: Path) -> dict[str, int]:
    fields_raw = path.read_text(encoding="utf-8").strip().split("\t")
    check(all("=" in item for item in fields_raw), f"repeat malformed fields: {path}")
    fields = dict(item.split("=", 1) for item in fields_raw)
    check(set(fields) == {"repeats", "found", "nanos"}, f"repeat fields changed: {path}")
    try:
        result = {key: int(fields[key]) for key in fields}
    except ValueError as error:
        raise AuditError(f"repeat integer parse failed: {path}") from error
    check(all(value >= 0 for value in result.values()), f"repeat negative counter: {path}")
    return result


def audit_repeat(plan: dict[str, Any], continue_on_rejection: bool = False) -> tuple[list[dict[str, Any]], bool]:
    directory = CAPTURES / "repeat"
    cases = {case["case"]: case for case in plan["cases"]}
    rows: list[dict[str, Any]] = []
    passed = True
    for case_name in plan["repeat_cases"]:
        case = cases[case_name]
        for mode in ("owned", "file"):
            values: dict[str, list[dict[str, int]]] = {}
            for leg in plan["native"]["legs"]:
                leg_values = [parse_repeat(directory / f"{case_name}-{mode}-{leg}-{sample}.tsv")
                              for sample in range(plan["repeat"]["samples_per_leg"])]
                check(all(item["repeats"] == plan["repeat"]["repetitions"] for item in leg_values), f"repeat/{case_name}/{mode}/{leg}: repetition count differs")
                expected_found = plan["repeat"]["repetitions"] if case["expected_status"] == "value" else 0
                found_expected = all(item["found"] == expected_found for item in leg_values)
                if not found_expected and not continue_on_rejection:
                    check(False, f"repeat/{case_name}/{mode}/{leg}: found count differs")
                values[leg] = leg_values
            found = {item["found"] for leg in values.values() for item in leg}
            found_equal = len(found) == 1
            if not found_equal and not continue_on_rejection:
                check(False, f"repeat/{case_name}/{mode}: found count differs across legs")
            timing = {leg: stats([item["nanos"] / plan["repeat"]["repetitions"] for item in leg_values])
                      for leg, leg_values in values.items()}
            gates: dict[str, bool] = {}
            pair_percent: dict[str, dict[str, float]] = {}
            for pair, candidate_leg, baseline_leg in (("b1_a1", "b1", "a1"), ("b2_a2", "b2", "a2")):
                # Repeat evidence is its own frozen metric name; the plan's
                # absolute warm-query exception deliberately does not include
                # ``repeat-q8``, keeping this gate at the strict 5% rule.
                gates[pair] = all(timing_gate("repeat-q8", float(timing[baseline_leg][key]), float(timing[candidate_leg][key]), plan)
                                  for key in ("p50", "mean"))
                pair_percent[pair] = {
                    key: percent_change(float(timing[baseline_leg][key]), float(timing[candidate_leg][key]))
                    for key in ("p50", "mean")
                }
            expected_found = plan["repeat"]["repetitions"] if case["expected_status"] == "value" else 0
            found_expected = all(item["found"] == expected_found
                                 for leg_values in values.values() for item in leg_values)
            group_pass = found_expected and found_equal and all(gates.values())
            passed &= group_pass
            rows.append({"case": case_name, "mode": mode, "found_equal": found_equal,
                         "paired_timing_pass": gates, "paired_percent": pair_percent,
                         "gate_pass": group_pass})
    check(len(rows) == plan["repeat"]["groups"], "repeat group count differs")
    return rows, passed


def audit_benefit(rows: list[dict[str, Any]], plan: dict[str, Any], kind: str,
                  continue_on_rejection: bool = False) -> bool:
    criterion = plan["benefit_gates"][kind]
    matches = [row for row in rows
               if row.get("case") == criterion["case"] and row.get("mode") == criterion["mode"]]
    check(len(matches) == 1, f"{kind} benefit row is missing or duplicated")
    row = matches[0]
    if kind == "native":
        source = row.get("paired_percent", {}).get(criterion["metric"], {})
    else:
        source = row.get("paired_percent", {})
    minimum = float(criterion["minimum_improvement_percent"])
    passed = True
    for pair in ("b1_a1", "b2_a2"):
        for statistic in criterion["statistics"]:
            change = float(source.get(pair, {}).get(statistic, float("inf")))
            passed &= -change >= minimum
    if not passed and not continue_on_rejection:
        check(False, f"{kind} 54016-missing-1048576 owned q8 benefit is below {minimum:.1f}%")
    return passed


def audit_allocator(plan: dict[str, Any], continue_on_rejection: bool = False) -> tuple[list[dict[str, Any]], bool]:
    directory = CAPTURES / "allocator"
    cases = {case["case"]: case for case in plan["cases"]}
    rows: list[dict[str, Any]] = []
    passed = True
    strict_operations = set(plan["allocator"]["strict_operations"])
    allowance = plan["allocator"]["q2_allowance"]
    for case_name, case in cases.items():
        for mode in ("owned", "file"):
            for operation in plan["allocator"]["operations"]:
                phase: dict[str, list[dict[str, Any]]] = {}
                for name in ("baseline", "candidate"):
                    records = [read_json(directory / f"{name}-{case_name}-{mode}-{operation}-{repeat}.json")
                               for repeat in range(plan["allocator"]["repeats"])]
                    for item in records:
                        check(item.get("mode") == mode and item.get("operation") == operation,
                              f"allocator/{case_name}/{mode}/{operation}: report identity differs")
                        for field, expected in (("worksheet", case["sheet"]), ("row", case["row"]),
                                                ("column", case["column"]), ("budget", case["budget"])):
                            check(item.get(field) == expected, f"allocator/{case_name}/{mode}/{operation}: {field} differs")
                        for metric in ALLOC_METRICS:
                            check(isinstance(item.get(metric), int) and not isinstance(item.get(metric), bool),
                                  f"allocator/{case_name}/{mode}/{operation}: malformed {metric}")
                    check(all(item == records[0] for item in records[1:]),
                          f"allocator/{case_name}/{mode}/{operation}: repeats differ in {name}")
                    phase[name] = records
                before, after = phase["baseline"][0], phase["candidate"][0]
                repeats_equal = all(item == phase[name][0]
                                    for name in ("baseline", "candidate")
                                    for item in phase[name][1:])
                nonmetrics_before = {key: value for key, value in before.items() if key not in ALLOC_METRICS}
                nonmetrics_after = {key: value for key, value in after.items() if key not in ALLOC_METRICS}
                metadata_equal = nonmetrics_before == nonmetrics_after
                if not repeats_equal and not continue_on_rejection:
                    check(False, f"allocator/{case_name}/{mode}/{operation}: repeats differ")
                if not metadata_equal and not continue_on_rejection:
                    check(False, f"allocator/{case_name}/{mode}/{operation}: metadata differs")
                strict = operation in strict_operations or case["budget"] == 0 or case["expected_status"] == "error"
                deltas = {metric: after[metric] - before[metric] for metric in ALLOC_METRICS}
                limits = {metric: 0 for metric in ALLOC_METRICS} if strict else allowance
                missing_warm = (case["expected_status"] == "missing"
                                and case["budget"] > 0 and operation in ("q3", "q8"))
                expected_deltas = (plan["allocator"]["missing_warm_expected_deltas"]
                                   if missing_warm else None)
                if missing_warm:
                    metric_pass = deltas == expected_deltas
                elif strict:
                    metric_pass = all(delta == 0 for delta in deltas.values())
                else:
                    metric_pass = all(delta <= int(limits[metric]) for metric, delta in deltas.items())
                if not metric_pass and not continue_on_rejection:
                    check(False, f"allocator/{case_name}/{mode}/{operation}: allowance exceeded {deltas}")
                group_pass = repeats_equal and metadata_equal and metric_pass
                passed &= group_pass
                rows.append({"case": case_name, "mode": mode, "operation": operation,
                             "strict": strict and not missing_warm, "deltas": deltas,
                             "allowance": limits, "expected_deltas": expected_deltas,
                             "repeats_equal": repeats_equal,
                             "gate_pass": group_pass})
    expected = len(plan["cases"]) * 2 * len(plan["allocator"]["operations"])
    check(len(rows) == expected, "allocator group count differs")
    return rows, passed


def validate_budget_report(plan: dict[str, Any], case: dict[str, Any], budget: int,
                           queries_expected: int, report: dict[str, Any], label: str,
                           continue_on_rejection: bool = False) -> bool:
    required_metrics = {"read_calls", "read_bytes", "version_calls", "len_calls"}
    check(report.get("schema_version") == 1 and report.get("probe") == "change-0684-xls-index-budget",
          f"{label}: probe schema identity differs")
    path = ROOT / case["path"]
    check(report.get("input_sha256") == sha256_file(path) and report.get("input_bytes") == path.stat().st_size,
          f"{label}: fixture identity differs")
    check(Path(report.get("input_path", "")).resolve() == path.resolve(), f"{label}: input path differs")
    for field, expected in (("worksheet", case["sheet"]), ("row", case["row"]), ("column", case["column"]),
                            ("max_query_index_bytes", budget),
                            ("queries_requested", queries_expected)):
        check(report.get(field) == expected, f"{label}: {field} differs")
    queries = report.get("queries")
    check(isinstance(queries, list) and len(queries) == queries_expected, f"{label}: query count differs")
    check(report.get("all_queries_agree") is True, f"{label}: repeated outcomes disagree")
    first = queries[0].get("outcome")
    check(isinstance(first, dict) and first.get("status") in {"value", "missing", "error"},
          f"{label}: malformed semantic outcome")
    expected_status = first.get("status") == case["expected_status"]
    if not expected_status and not continue_on_rejection:
        check(False, f"{label}: expected status differs")
    open_metrics = report.get("open_metrics")
    check(isinstance(open_metrics, dict) and set(open_metrics) == required_metrics,
          f"{label}: open metric vector shape changed")
    for key, value in open_metrics.items():
        check(isinstance(value, int) and value >= 0, f"{label}: invalid open metric {key}")
    for index, query in enumerate(queries):
        check(query.get("ordinal") == index and query.get("outcome") == first, f"{label}: query semantic parity differs")
        metrics = query.get("metrics")
        check(isinstance(metrics, dict) and set(metrics) == required_metrics,
              f"{label}: source metric vector shape changed")
        for key, value in metrics.items():
            check(isinstance(value, int) and value >= 0, f"{label}: invalid source metric {key}")
    return expected_status


def audit_budget_fence(plan: dict[str, Any], continue_on_rejection: bool = False) -> tuple[list[dict[str, Any]], bool]:
    directory = CAPTURES / "budget-fence"
    rows: list[dict[str, Any]] = []
    passed = True
    for case in plan["budget_fence"]["cases"]:
        for budget in case["budgets"]:
            reports = {phase: read_json(directory / f"{phase}-{case['case']}-{budget}.json")
                       for phase in ("baseline", "candidate")}
            baseline_expected = validate_budget_report(
                plan, case, budget, plan["budget_fence"]["queries"], reports["baseline"],
                f"{case['case']}/{budget}/baseline", continue_on_rejection
            )
            candidate_expected = validate_budget_report(
                plan, case, budget, plan["budget_fence"]["queries"], reports["candidate"],
                f"{case['case']}/{budget}/candidate", continue_on_rejection
            )
            baseline_outcomes = [query["outcome"] for query in reports["baseline"]["queries"]]
            candidate_outcomes = [query["outcome"] for query in reports["candidate"]["queries"]]
            equal = baseline_expected and candidate_expected and baseline_outcomes == candidate_outcomes
            if not equal and not continue_on_rejection:
                check(False, f"budget-fence/{case['case']}/{budget}: semantic outcome differs")
            passed &= equal
            deltas = []
            for before, after in zip(reports["baseline"]["queries"], reports["candidate"]["queries"]):
                before_metrics, after_metrics = before["metrics"], after["metrics"]
                deltas.append({key: int(after_metrics.get(key, 0)) - int(before_metrics.get(key, 0))
                               for key in sorted(set(before_metrics) | set(after_metrics))})
            metrics_equal = all(all(delta == 0 for delta in item.values()) for item in deltas)
            if not metrics_equal and not continue_on_rejection:
                check(False, f"budget-fence/{case['case']}/{budget}: source metrics changed")
            rows.append({"case": case["case"], "budget": budget, "outcomes_equal": equal,
                         "candidate_minus_baseline_metrics": deltas,
                         "metrics_equal": metrics_equal,
                         "route_changed": not metrics_equal,
                         "gate_pass": equal and metrics_equal})
    expected = sum(len(case["budgets"]) for case in plan["budget_fence"]["cases"])
    check(len(rows) == expected, f"budget fence row count differs: {len(rows)} != {expected}")
    return rows, passed and all(row["outcomes_equal"] and row["metrics_equal"] for row in rows)


def audit_budget_primary(plan: dict[str, Any], continue_on_rejection: bool = False) -> tuple[list[dict[str, Any]], bool]:
    directory = CAPTURES / "budget-primary"
    rows: list[dict[str, Any]] = []
    passed = True
    for case in plan["cases"]:
        reports = {
            phase: read_json(directory / f"{phase}-{case['case']}.json")
            for phase in ("baseline", "candidate")
        }
        budget = int(case["budget"])
        queries = int(plan["budget_primary"]["queries"])
        baseline_expected = validate_budget_report(
            plan, case, budget, queries, reports["baseline"],
            f"budget-primary/{case['case']}/baseline", continue_on_rejection
        )
        candidate_expected = validate_budget_report(
            plan, case, budget, queries, reports["candidate"],
            f"budget-primary/{case['case']}/candidate", continue_on_rejection
        )
        before = [query["outcome"] for query in reports["baseline"]["queries"]]
        after = [query["outcome"] for query in reports["candidate"]["queries"]]
        equal = baseline_expected and candidate_expected and before == after
        if not equal and not continue_on_rejection:
            check(False, f"budget-primary/{case['case']}: semantic outcome differs")
        passed &= equal
        rows.append({"case": case["case"], "outcomes_equal": equal})
    check(len(rows) == plan["budget_primary"]["groups_per_phase"], "budget primary group count differs")
    return rows, passed and all(row["outcomes_equal"] for row in rows)


def audit_trace_if_present() -> dict[str, Any]:
    root = PACKET / "trace"
    if not root.exists():
        return {"present": False, "passed": True, "note": "diagnostic trace not captured"}
    findings: list[str] = []
    sequence_reports: dict[str, Any] = {}
    for phase in ("baseline", "candidate"):
        directory = root / phase
        manifest_path = directory / "manifest.json"
        if not manifest_path.is_file():
            findings.append(f"trace/{phase}: manifest missing")
            continue
        manifest = read_json(manifest_path)
        if manifest.get("restored") is not True:
            findings.append(f"trace/{phase}: sources were not restored")
        restored = manifest.get("restored_source_sha256", {})
        for relative_path in (
            "crates/litchi-cfb/src/shared.rs",
            "crates/litchi-xls/src/workbook/query_cache.rs",
            "crates/litchi-xls/src/workbook/source.rs",
        ):
            live = ROOT / relative_path
            if live.is_file() and restored.get(relative_path) != sha256_file(live):
                findings.append(f"trace/{phase}: restored source differs: {relative_path}")
        for output in manifest.get("outputs", []):
            trace = directory / output["trace"]
            semantic = directory / output["semantic"]
            if not trace.is_file() or sha256_file(trace) != output.get("trace_sha256"):
                findings.append(f"trace/{phase}: raw trace hash differs")
            if not semantic.is_file() or sha256_file(semantic) != output.get("semantic_sha256"):
                findings.append(f"trace/{phase}: semantic hash differs")
            if trace.is_file() and re.search(r"(?i)(elapsed|nanos|timing|seconds)", trace.read_text(errors="replace")):
                findings.append(f"trace/{phase}: timing text leaked")
        sequence = manifest.get("sequence_output", {})
        sequence_trace = directory / str(sequence.get("trace", ""))
        sequence_semantic = directory / str(sequence.get("semantic", ""))
        if not sequence_trace.is_file() or not sequence_semantic.is_file():
            findings.append(f"trace/{phase}: sequence evidence missing")
        else:
            report = read_json(sequence_semantic)
            sequence_reports[phase] = report
            labels = [item.get("label") for item in report.get("queries", [])]
            if labels != ["late-build", "late-publish", "first-earlier", "origin-late", "missing"]:
                findings.append(f"trace/{phase}: sequence labels differ: {labels}")
            if any(key.lower().find("timing") >= 0 or key.lower().find("nanos") >= 0
                   for key in report.keys()):
                findings.append(f"trace/{phase}: sequence semantic timing field leaked")
    if set(sequence_reports) == {"baseline", "candidate"} and sequence_reports["baseline"] != sequence_reports["candidate"]:
        findings.append("trace: baseline/candidate sequence semantic or I/O report differs")
    return {"present": True, "passed": not findings, "findings": findings}


def audit_freeze(plan: dict[str, Any], current: dict[str, str], probes: dict[str, str],
                 corpus: dict[str, Any], builds: dict[str, Any]) -> dict[str, Any]:
    path = CAPTURES / "freeze.json"
    check(path.is_file(), "capture freeze is missing")
    frozen = read_json(path)
    check(frozen.get("status") == "frozen", "capture freeze is not final")
    check(frozen.get("task") == "freeze" and frozen.get("plan_sha256") == sha256_file(PLAN),
          "capture freeze identity changed")
    check(frozen.get("tools_sha256_start") == tool_map()
          and frozen.get("tools_sha256_end") == frozen.get("tools_sha256_start"),
          "capture freeze tool binding changed")
    check(frozen.get("source_sha256_start") == current and frozen.get("source_sha256_end") == current,
          "capture freeze source binding changed")
    check(frozen.get("baseline_source_sha256") == git_source_map(plan, plan["baseline_revision"]),
          "capture freeze baseline source differs")
    case_manifest_sha = sha256_file(ROOT / plan["case_manifest"])
    check(frozen.get("probe_sha256") == probes and frozen.get("probe_sha256_end") == probes
          and frozen.get("corpus") == corpus and frozen.get("corpus_end") == corpus
          and frozen.get("case_manifest_sha256") == case_manifest_sha
          and frozen.get("case_manifest_sha256_end") == case_manifest_sha,
          "capture freeze probe/fixture binding changed")
    for phase in ("baseline", "candidate"):
        for kind, expected in builds[phase].items():
            receipt = frozen["binaries"][phase][kind]
            check(receipt.get("sha256") == expected, f"capture freeze binary hash differs: {phase}/{kind}")
    return frozen


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--draft", action="store_true", help="run pre-capture custody/static checks")
    parser.add_argument(
        "--allow-rejected", "--report-rejected", dest="allow_rejected", action="store_true",
        help="emit the complete terminal report after a gate rejection, but still exit 1",
    )
    args = parser.parse_args()
    try:
        plan = load_plan()
        static = check_source_static(plan)
        builds = check_builds(plan)
        negative_checks = check_negative_checks()
        current = current_source_map(plan)
        baseline = git_source_map(plan, plan["baseline_revision"])
        probes = probe_map(plan)
        corpus = corpus_map(plan)
        freeze_path = CAPTURES / "freeze.json"
        if not freeze_path.is_file():
            if not args.draft:
                raise Pending(("captures/freeze.json",))
            print("DRAFT PASS static/source/build/fixture preflight; freeze and captures are pending")
            print(json.dumps({"static": static, "builds": builds,
                              "negative_checks": negative_checks,
                              "pending": ["captures/freeze.json"]}, sort_keys=True))
            return 0
        frozen = audit_freeze(plan, current, probes, corpus, builds)
        witnesses = collect_witnesses()
        manifests: dict[str, Any] = {}
        task_kinds = {
            "native": "native",
            "repeat": "repeat",
            "allocator": "allocator",
            "budget-primary": "budget",
            "budget-fence": "budget",
        }
        for task, kind in task_kinds.items():
            manifest = audit_manifest(plan, task, kind, current, baseline, probes, corpus, frozen, witnesses, args.draft)
            if manifest is None:
                raise Pending((f"captures/{task}/manifest.json",))
            manifests[task] = manifest
        if args.draft:
            print("DRAFT PASS freeze and capture manifests are structurally complete; terminal raw audit pending")
            return 0
        native_rows, native_pass = audit_native(plan)
        repeat_rows, repeat_pass = audit_repeat(plan)
        native_benefit = audit_benefit(native_rows, plan, "native", args.allow_rejected)
        repeat_benefit = audit_benefit(repeat_rows, plan, "repeat", args.allow_rejected)
        allocator_rows, allocator_pass = audit_allocator(plan, args.allow_rejected)
        budget_primary_rows, budget_primary_pass = audit_budget_primary(plan, args.allow_rejected)
        budget_rows, budget_pass = audit_budget_fence(plan, args.allow_rejected)
        trace = audit_trace_if_present()
        if trace["present"] and not trace["passed"] and not args.allow_rejected:
            check(False, f"diagnostic trace failed: {trace.get('findings')}")
        report = {
            "packet": plan["change"],
            "baseline_revision": plan["baseline_revision"],
            "allow_rejected": args.allow_rejected,
            "static": static,
            "negative_checks": negative_checks,
            "manifests": {name: {"status": value.get("status"), "raw_count": len(value.get("raw_sha256", {}))}
                          for name, value in manifests.items()},
            "native_groups": len(native_rows),
            "repeat_groups": len(repeat_rows),
            "benefit_gates": {"native": native_benefit, "repeat": repeat_benefit},
            "allocator_groups": len(allocator_rows),
            "budget_primary_groups": len(budget_primary_rows),
            "budget_fence_rows": len(budget_rows),
            "trace": trace,
            "hard_gates": {
                "bindings": True,
                "native": native_pass,
                "repeat": repeat_pass,
                "native_benefit": native_benefit,
                "repeat_benefit": repeat_benefit,
                "allocator": allocator_pass,
                "budget_primary": budget_primary_pass,
                "budget_fence": budget_pass,
                "trace": not trace["present"] or trace["passed"],
            },
        }
        passed = all(report["hard_gates"].values())
        label = "PASS" if passed else ("REJECTED" if args.allow_rejected else "FAIL")
        print(label, json.dumps(report, sort_keys=True))
        return 0 if passed else 1
    except Pending as pending:
        print("PENDING: " + "; ".join(pending.items), file=sys.stderr)
        return 2
    except (AuditError, OSError, subprocess.CalledProcessError, KeyError, TypeError, ValueError) as error:
        print(f"FAIL: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
