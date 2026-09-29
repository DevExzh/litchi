"""Read-only additive closure for the repaired Callgrind side population.

The admitted ``close.py`` remains the closure for the native/Heaptrack
experiment.  This file repeats its main evidence guards and adds the bounded
reader-v2 replay.  Verification is read-only; explicit clean/seal modes preserve the original ownership and commit guards.
"""

from __future__ import annotations

import re
import shutil
import sys
from pathlib import Path
from typing import Any, Mapping

import analyze as a
import crosscheck as cc
import driver as d
import observer_check as oc
import profile_resume as rp
import qualify as q
import tables


P = d.P
REPORT = "docs/performance/0845-opc-indexed-entry-iteration.md"
INDEXES = [f"docs/performance/{name}.md" for name in
           ["BASELINE", "HOTSPOTS", "REPORT", "CRUD_COVERAGE", "GOAL_AUDIT"]]
PRODUCTION = [
    "crates/soapberry-zip/src/office.rs",
    "crates/litchi-opc/src/source_backed.rs",
]
FAILED_PARENT = "profile-capture"


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def read(path: Path) -> Any:
    return d.read(path)


def descriptor(path: Path) -> dict[str, Any]:
    return d.desc(path)


def bind(item: Mapping[str, Any], label: str) -> dict[str, Any]:
    require(isinstance(item, Mapping), f"{label}: descriptor is missing")
    actual = descriptor(Path(item.get("path", "")))
    require(actual == dict(item), f"{label}: descriptor changed")
    return actual


def bind_reader_repair() -> dict[str, Any]:
    """Bind the additive parser test and the installed Callgrind reference."""

    tests = read(P / "reader-v2-tests.json")
    require(tests.get("status") == "pass", "reader-v2-tests is not passing")
    require(tests.get("qualification_reports") == 64,
            "reader-v2-tests did not retain all qualification reports")
    require(len(tests.get("inherited_mutations_rejected", [])) == 34,
            "reader-v2-tests inherited mutation population changed")
    require(len(tests.get("regressions", [])) == 8,
            "reader-v2-tests grammar regression population changed")
    for key in ("reader", "previous", "test_source", "inherited_test_source"):
        bind(tests[key], f"reader-v2-tests {key}")
    captures = tests.get("captures")
    require(isinstance(captures, list) and len(captures) == 4,
            "reader-v2-tests retained capture set changed")
    for capture in captures:
        for key in ("command", "numbered", "terminal"):
            bind(capture[key], f"reader-v2-tests capture {key}")
        require(capture.get("summary_ir", 0) > 0,
                "reader-v2-tests capture has no positive summary")
    command_path = P / "commands" / "reader-v2-tests" / "receipt.json"
    command = read(command_path)
    require(command.get("argv") == [
        "python3", "-B", str(P.relative_to(d.ROOT) / "reader_v2_tests.py")
    ],
            "reader-v2-tests command argv changed")
    require(command.get("exit_code") == 0, "reader-v2-tests command is not passing")
    bind(command["log"], "reader-v2-tests command log")
    require((command_path.parent / "output.log").read_text(encoding="utf-8") ==
            "Reader v2 PASS: 64 qualifications; 34 inherited mutations; 8 grammar checks; 4 retained profiles\n",
            "reader-v2-tests command output/result metadata changed")
    spec = read(P / "reader-repair-spec.json")
    require(spec.get("status") == "reviewed", "reader repair specification is not reviewed")
    require(spec.get("repair") ==
            "Recognize fe as a current source-file switch and cfl as the cfi callee-file alias; preserve function-identity and cost conservation assertions.",
            "reader repair scope changed")
    bind(spec["installed_reference"], "installed Callgrind reference")
    bind(spec["regression_tests"], "reader repair regression result")
    return tests


def bind_repair_artifacts() -> dict[str, Any]:
    """Replay the freeze, manifest, and v2 analysis without invoking writers."""

    require(rp.FREEZE_PATH.is_file(), "profile-resume-freeze.json is required")
    require(rp.RESULT_PATH.is_file(), "profile-resume.json is required")
    require(rp.ANALYSIS_PATH.is_file(), "profile-analysis-v2.json is required")
    frozen = read(rp.FREEZE_PATH)
    rp.validate_frozen_artifacts(frozen)
    # Profile evidence was captured against the frozen after leg.  Once the
    # decision is closed, the main closure below checks the selected current
    # source; profile replay itself must remain valid if that decision is
    # reject and the checkout is back on the before leg.
    value, admission, bindings, qcache, qualifications, prepared = rp.prepare(
        require_current_after=False,
    )
    failure = prepared["failure"]

    manifest = read(rp.RESULT_PATH)
    require(manifest.get("schema") == "litchi.opc-preservation-profile-resume-0845.v2",
            "profile resume manifest schema changed")
    require(manifest.get("status") == "pass", "profile resume manifest is not passing")
    require(manifest.get("scope") ==
            "diagnostic-only-callgrind; separate from native and Heaptrack populations",
            "profile resume scope changed")
    require(manifest.get("reports") == manifest.get("samples") == 16,
            "profile resume population changed")
    for key, path in {
        "plan": rp.PLAN_PATH,
        "original_runner": P / "profile.py",
        "runner": rp.RUNNER_PATH,
        "reader": rp.READER_PATH,
        "reader_v1": P / "reader.py",
        "admission": P / "admission.json",
        "freeze": rp.FREEZE_PATH,
    }.items():
        require(manifest.get(key) == descriptor(path),
                f"profile resume {key} binding changed")
    require(manifest.get("original_capture_failure") == failure,
            "profile resume failure witness changed")
    require(manifest.get("prefix_count") == 4 and manifest.get("resumed_count") == 12,
            "profile resume prefix accounting changed")
    require(manifest.get("receipts") == [
        descriptor(Path(row["folder"]) / "receipt.json")
        for row in rp.validate_complete(value, admission, bindings, qcache)
    ], "profile resume receipt inventory changed")
    require(manifest.get("rules", {}).get("capture_manifest_written") is False,
            "profile resume synthesized the original capture manifest")
    require(manifest.get("rules", {}).get("native_or_heap_overlap") is False,
            "profile resume permits cross-lane overlap")
    bind(manifest.get("freeze"), "profile resume freeze")
    for item in manifest["receipts"]:
        bind(item, "profile resume receipt")

    validated_rows = rp.validate_complete(value, admission, bindings, qcache)
    require(manifest == rp.resume_result(value, prepared["admission"], bindings,
            qualifications, frozen, prepared["failure"], prepared["repair"], validated_rows),
            "profile resume full reconstruction changed")
    rp.verify_prefix_against_freeze(frozen, value, bindings, qcache, admission,
                                   receipt_for_missing=True)
    analysis = read(rp.ANALYSIS_PATH)
    require(rp.analyze() == analysis, "profile-analysis-v2 full raw replay changed")
    require(analysis.get("schema") ==
            "litchi.opc-preservation-profile-analysis-0845.v2",
            "profile-analysis-v2 schema changed")
    require(analysis.get("status") == "pass" and
            analysis.get("disposition") == "diagnostic-only-no-native-pooling",
            "profile-analysis-v2 disposition changed")
    require(analysis.get("reports") == analysis.get("samples") == 16,
            "profile-analysis-v2 population changed")
    for key, path in {
        "plan": rp.PLAN_PATH,
        "original_runner": P / "profile.py",
        "runner": rp.RUNNER_PATH,
        "reader": rp.READER_PATH,
        "reader_v1": P / "reader.py",
        "admission": P / "admission.json",
        "freeze": rp.FREEZE_PATH,
        "resume": rp.RESULT_PATH,
    }.items():
        require(analysis.get(key) == descriptor(path),
                f"profile-analysis-v2 {key} binding changed")
    require(analysis.get("original_capture_failure") == failure,
            "profile-analysis-v2 failure witness changed")
    require(analysis.get("selected_symbols") == value["selected_symbols"],
            "profile-analysis-v2 selected symbols changed")
    require(analysis.get("limits") == value["limits"],
            "profile-analysis-v2 limits changed")
    rows = analysis.get("rows")
    require(isinstance(rows, list) and len(rows) == 16,
            "profile-analysis-v2 rows changed")
    validated = rp.validate_complete(value, admission, bindings, qcache)
    require([item["row"] for item in rows] == [item["row"] for item in validated],
            "profile-analysis-v2 row order changed")
    for index, item in enumerate(rows):
        require(item.get("folder") == validated[index]["folder"],
                "profile-analysis-v2 folder binding changed")
        bind(item.get("receipt"), f"profile-analysis-v2 receipt {index}")
        bind(item.get("command"), f"profile-analysis-v2 command {index}")
        compact = item.get("compact")
        require(isinstance(compact, Mapping), f"profile-analysis-v2 compact row {index} missing")
        raw = compact.get("raw_callgrind")
        require(isinstance(raw, Mapping), f"profile-analysis-v2 raw row {index} missing")
        bind(raw["numbered"], f"profile-analysis-v2 numbered raw {index}")
        bind(raw["terminal"], f"profile-analysis-v2 terminal raw {index}")
        require(compact.get("summary_ir", 0) > 0,
                f"profile-analysis-v2 row {index} has no positive summary")
        require(compact.get("owner", {}).get("inclusive_ir") == compact.get("summary_ir"),
                f"profile-analysis-v2 row {index} owner conservation changed")
    require(len(analysis.get("comparisons", [])) == 8,
            "profile-analysis-v2 comparison population changed")
    return {
        "freeze": descriptor(rp.FREEZE_PATH),
        "manifest": descriptor(rp.RESULT_PATH),
        "analysis": descriptor(rp.ANALYSIS_PATH),
        "reports": len(rows),
        "samples": len(rows),
        "failure": failure,
    }


def verify_commands() -> list[dict[str, Any]]:
    """Require zero for every command except the exact bound parent failure."""

    failure = rp.validate_failure_witness()
    commands: list[dict[str, Any]] = []
    nonzero: list[str] = []
    for folder in sorted((P / "commands").iterdir()):
        end = read(folder / "receipt.json")
        start = read(folder / "started.json")
        bind(end["log"], f"command log {folder.name}")
        require(end.get("argv") == start.get("argv") and
                end.get("started_unix") == start.get("started_unix") and
                end.get("finished_unix", 0) >= end.get("started_unix", 0),
                f"command receipt changed: {folder.name}")
        code = end.get("exit_code")
        if code != 0:
            nonzero.append(folder.name)
        if folder.name == FAILED_PARENT:
            require(end == read(Path(failure["receipt"]["path"])),
                    "profile-capture failure is not the bound parent receipt")
            require(code == 1, "profile-capture failure exit changed")
        else:
            require(code == 0, f"unexpected command failure: {folder.name}")
        commands.append({"label": folder.name, "exit_code": code})
        rows = re.findall(
            r"test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;",
            (folder / "output.log").read_text(encoding="utf-8"),
        )
        if rows:
            # The original closure's owner-test accounting is retained below.
            pass
    require(nonzero == [FAILED_PARENT], "command failure allowlist changed")
    return commands


def verify() -> dict[str, Any]:
    """Replay the original closure guards plus the additive repair custody."""

    origin = read(P / "origin.json")
    for group in ("normative", "unrelated"):
        require(all(d.sha(d.ROOT / name) == digest
                    for name, digest in origin[group].items()),
                f"{group} source inventory changed")
    for item in read(P / "admission.json")["files"]:
        q.bind(item)

    reader_tests = bind_reader_repair()
    repair = bind_repair_artifacts()

    decision = read(P / "decision.json")
    require(decision.get("status") == "complete-batch" and
            decision.get("decision") in ("adopt", "reject"),
            "decision is not complete")
    leg = "after" if decision["decision"] == "adopt" else "before"
    require(d.source() == read(P / f"source-{leg}.json")["source"],
            "selected source differs from decision")
    require(d.sha(d.ROOT / "Cargo.lock") == d.sha(P / "root-Cargo.lock"),
            "Cargo.lock changed")

    for build_leg in ("before", "after"):
        build = read(P / f"build-{build_leg}.json")
        q.bind(build["freeze"])
        if Path(build["binary"]["path"]).exists():
            q.bind(build["binary"])
        else:
            cleanup = read(P / "cleanup.json")
            require(build["binary"] in cleanup["verified_binaries"] and not d.TARGET.exists(),
                    f"{build_leg} binary has no cleanup witness")
        require(oc.verify(build_leg) == read(P / f"observer-preflight-{build_leg}.json"),
                f"{build_leg} observer preflight changed")

    result = a.analyze()
    require(result == read(P / "analysis.json") and
            result["reports"] == 512 and result["samples"] == 211328,
            "main analysis changed")
    require(cc.verify() == read(P / "crosscheck.json"), "independent crosscheck changed")
    require(tables.render() == (P / "tables.md").read_text(encoding="utf-8"),
            "tables changed")
    require(decision["analysis"] == descriptor(P / "analysis.json"),
            "decision analysis binding changed")
    require(decision["profile_analysis_v2"] == descriptor(P / "profile-analysis-v2.json"),
            "decision profile-analysis-v2 binding changed")
    for field in ("regression_flags", "tail_review", "spread_flags", "allocation_flags"):
        require(decision[field] == len(result[field]), f"decision {field} count changed")
    if decision["decision"] == "adopt":
        require(decision["material_benefit"] and decision["all_flags_reviewed"] and
                decision["unresolved_common_workload_damage"] is False,
                "adopt decision guards changed")

    quality = read(P / "quality-after.json")
    require(quality["status"] == "pass", "quality gate is not passing")
    for item in quality["receipts"]:
        q.bind(item)
        require(read(Path(item["path"]))["exit_code"] == 0,
                "quality receipt is not successful")
    baseline = read(P / "baseline-gate-reuse.json")
    require(baseline["status"] == "pass", "baseline gate reuse is not passing")
    for key in ("production", "prior_quality", "prior_closure", "prior_seal",
                "fresh_replay", "fresh_owner_tests"):
        q.bind(baseline[key])

    for name in ("reader-tests", "heap-reader-tests"):
        tests = read(P / f"{name}.json")
        require(tests["status"] == "pass", f"{name} is not passing")
        for key in ("reader", "test_source", "receipt"):
            q.bind(tests[key])
        require(read(Path(tests["receipt"]["path"]))["exit_code"] == 0,
                f"{name} receipt is not successful")

    commands = verify_commands()
    totals: dict[str, dict[str, int]] = {}
    for folder in sorted((P / "commands").iterdir()):
        rows = re.findall(
            r"test result: (?:ok|FAILED)\. (\d+) passed; (\d+) failed; (\d+) ignored;",
            (folder / "output.log").read_text(encoding="utf-8"),
        )
        if rows:
            totals[folder.name] = {
                key: sum(int(row[index]) for row in rows)
                for index, key in enumerate(("passed", "failed", "ignored"))
            }
    require(totals["before-owner-test"]["failed"] ==
            totals["after-owner-test"]["failed"] == 0,
            "owner tests have failures")
    require(not list(P.rglob("__pycache__")), "Python cache is present")
    return {
        "status": "pass",
        "decision": decision["decision"],
        "reports": 512,
        "samples": 211328,
        "separate_instruction_reports": repair["reports"],
        "separate_instruction_samples": repair["samples"],
        "separate_heap_preflight_reports": 2,
        "separate_heap_preflight_samples": 2,
        "reader_v2_tests": descriptor(P / "reader-v2-tests.json"),
        "profile_resume_freeze": repair["freeze"],
        "profile_resume_manifest": repair["manifest"],
        "profile_analysis_v2": repair["analysis"],
        "controlled_failed_command": FAILED_PARENT,
        "commands": commands,
        "tests": totals,
        "formal_runs_replaced": 0,
        "native_or_heap_pooling": False,
    }


if __name__ == "__main__":
    require(len(sys.argv) == 2, "usage: close_v2.py verify|clean|seal")
    mode = sys.argv[1]
    if mode == "clean":
        d.check()
        value = verify()
        binaries = [read(P / f"build-{leg}.json")["binary"] for leg in ["before", "after"]]
        for item in binaries:
            q.bind(item)
        removed = []
        for root in [d.TARGET, d.SCRATCH]:
            require(read(root / ".owner-0845.json") == dict(packet=str(P), path=str(root)),
                    "owned cleanup marker changed")
            files = [path for path in root.rglob("*") if path.is_file()]
            removed.append(dict(path=str(root), files=len(files),
                                logical_bytes=sum(path.stat().st_size for path in files)))
            shutil.rmtree(root)
            require(not root.exists(), "owned root removal failed")
        d.write(P / "cleanup.json", dict(status="pass", verified_binaries=binaries, removed=removed))
        d.write(P / "closure.json", value)
    elif mode == "seal":
        require(verify() == read(P / "closure.json"), "closure changed")
        require(not d.TARGET.exists() and not d.SCRATCH.exists(), "owned roots still exist")
        names = (PRODUCTION if read(P / "decision.json")["decision"] == "adopt" else []) + INDEXES + [REPORT]
        names += [str(path.relative_to(d.ROOT)) for path in sorted(P.rglob("*"))
                  if path.is_file() and path.name != "seal.json"]
        d.write(P / "seal.json", dict(base=read(P / "origin.json")["base"],
                files={name: d.sha(d.ROOT / name) for name in names}))
    elif mode == "verify":
        value = verify()
        if (P / "closure.json").exists():
            require(value == read(P / "closure.json"), "closure changed")
        if (P / "seal.json").exists():
            seal = read(P / "seal.json")
            require(all(d.sha(d.ROOT / name) == digest for name, digest in seal["files"].items()),
                    "sealed file changed")
            require(not d.TARGET.exists() and not d.SCRATCH.exists(), "owned roots still exist")
            if d.output(["git", "rev-parse", "HEAD"]) != seal["base"]:
                require(d.output(["git", "rev-parse", "HEAD^"]) == seal["base"], "commit parent changed")
                actual = set(d.output(["git", "diff-tree", "--no-commit-id", "--name-only", "-r", "HEAD"]).splitlines())
                require(actual == set(seal["files"]) | {str((P / "seal.json").relative_to(d.ROOT))},
                        "committed path inventory changed")
    else:
        raise AssertionError(mode)
    print(mode, "PASS")
