"""Fail-closed validator for the 0792 evidence packet."""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import Any

import analyze


PACKET = analyze.PACKET
ROOT = analyze.ROOT


def require(condition: bool, message: str) -> None:
    analyze.require(condition, message)


def check_final_seal() -> dict[str, Any]:
    path = PACKET / "seal.json"
    seal = analyze.read_json(path)
    require(isinstance(seal, dict) and isinstance(seal.get("files"), dict)
            and isinstance(seal.get("documents"), dict), "final seal is malformed")
    actual = {
        str(item.relative_to(PACKET)): analyze.sha256(item)
        for item in PACKET.rglob("*")
        if item.is_file() and item.name != "seal.json"
    }
    require(seal["files"] == actual, "final packet seal inventory is stale")
    documents = seal["documents"]
    require(len(documents) == 6, "final document seal cardinality changed")
    for name, digest in documents.items():
        require(analyze.is_sha(digest), f"document seal digest is invalid: {name}")
        candidates = []
        raw = Path(name)
        if raw.is_absolute():
            candidates.append(raw)
        else:
            candidates.extend((ROOT / raw, ROOT / "docs" / raw,
                               ROOT / "docs/performance" / raw))
        matches = [candidate for candidate in candidates
                   if candidate.is_file() and not candidate.is_symlink()]
        require(len(matches) == 1 and analyze.sha256(matches[0]) == digest,
                f"document seal changed: {name}")
    return {"path": analyze.rel(path), "schema": seal.get("schema"),
            "files": len(actual), "documents": len(documents)}


def check_workspace_state() -> dict[str, Any]:
    origin = analyze.origin()
    unrelated = origin.get("unrelated")
    require(isinstance(unrelated, dict), "origin unrelated custody is missing")
    for name, digest in unrelated.items():
        path = ROOT / name
        require(path.is_file() and analyze.sha256(path) == digest,
                f"unrelated workspace file changed: {name}")
    expected = origin.get("worktrees")
    require(isinstance(expected, str) and expected, "origin worktree inventory is missing")
    current = analyze._git_output(["worktree", "list", "--porcelain"], "current worktrees")
    # The main worktree is allowed to advance after this packet is committed;
    # every other worktree remains exact custody.
    def blocks(value: str) -> dict[str, str]:
        output: dict[str, str] = {}
        for block in value.strip().split("\n\n"):
            lines = block.splitlines()
            if not lines or not lines[0].startswith("worktree "):
                continue
            output[lines[0][9:]] = "\n".join(lines)
        return output
    old, new = blocks(expected), blocks(current)
    require(ROOT.as_posix() in new, "main worktree disappeared")
    for path, block in old.items():
        if Path(path).resolve() == ROOT.resolve():
            continue
        require(new.get(path) == block, f"unrelated worktree changed: {path}")
    return {"unrelated_files": len(unrelated), "other_worktrees_checked": max(0, len(old) - 1)}


def validate(*, require_final_seal: bool = False,
             check_workspace: bool = False) -> dict[str, Any]:
    result = analyze.analyze()
    analysis_path = PACKET / "analysis.json"
    require(analysis_path.is_file() and analyze.read_json(analysis_path) == result,
            "analysis.json does not replay byte-for-byte")
    require(result["counts"] == {
        "reports": 255,
        "samples": 5595,
        "native_reports": 180,
        "allocation_reports": 60,
        "qualification_reports": 15,
    }, "aggregate counts changed")
    require(result["native"]["children"] == 180
            and result["allocation"]["children"] == 60
            and result["qualification"]["children"] == 15,
            "lane cardinality changed")
    require(result["quality"]["gates"] == 6
            and result["verification"]["all_raw_report_samples_checked"] is True
            and result["verification"]["semantic_verification_checked"] is True,
            "quality or raw semantic gate changed")
    require(result["probe_tests"]["tests"] == [28, 7]
            and result["probe_tests"]["filtered"] == [0, 26],
            "probe test custody changed")
    require(result["bootstrap"] == {
        "seed": analyze.BOOTSTRAP_SEED,
        "resamples": analyze.BOOTSTRAP_RESAMPLES,
        "confidence": analyze.BOOTSTRAP_CONFIDENCE,
        "statistic": "median",
        "low_rank": analyze.BOOTSTRAP_LOW_RANK,
        "high_rank": analyze.BOOTSTRAP_HIGH_RANK,
    }, "bootstrap contract changed")
    require(result["verification"]["allocation_memory_guards_use_block_medians"] is True
            and result["verification"]["allocation_calls_and_bytes_guards_checked"] is True,
            "allocation guard custody changed")
    for family in (result["native"]["analysis"], result["allocation"]["analysis"]):
        for group in family["paired_by_block_before_after"].values():
            for metric in group["metrics"].values():
                require(metric["bootstrap"]["seed"] == analyze.BOOTSTRAP_SEED
                        and metric["bootstrap"]["resamples"] == analyze.BOOTSTRAP_RESAMPLES
                        and metric["bootstrap"]["low_rank"] == analyze.BOOTSTRAP_LOW_RANK
                        and metric["bootstrap"]["high_rank"] == analyze.BOOTSTRAP_HIGH_RANK
                        and len(metric["by_block"]) == group["blocks"],
                        "paired bootstrap replay changed")
    status = result["disposition"]["status"]
    require(status in {"pending", "retained", "rejected"}, "disposition status changed")
    guards = result["decision_guards"]
    require(guards["latency_guard_passed"] is (not guards["latency_violations"])
            and guards["resource_guard_passed"] is (not guards["resource_violations"])
            and guards["benefit_satisfied"] is bool(guards["eligible_benefits"])
            and guards["adoption_eligible"] is (
                guards["latency_guard_passed"] and guards["resource_guard_passed"]
                and guards["benefit_satisfied"]),
            "decision guard summary changed")
    if status == "rejected":
        require(guards["adoption_eligible"] is False,
                "rejected disposition contradicts frozen adoption guard")
    if require_final_seal:
        require(status in {"retained", "rejected"}, "final disposition is missing")
        require((PACKET / "cleanup.json").is_file(), "final cleanup witness is missing")
        cleanup = analyze.read_json(PACKET / "cleanup.json")
        require(cleanup.get("target_removed") is True
                and cleanup.get("target") == analyze.origin()["target"]
                and len(cleanup.get("removed_binaries", [])) == 4,
                "final cleanup witness is incomplete")
        require(not Path(cleanup["target"]).exists(), "owned target was not removed")
        seal = check_final_seal()
    else:
        seal = None
    workspace = check_workspace_state() if check_workspace else None
    if status == "retained":
        require(result["disposition"]["production_change_retained"] is True,
                "retained disposition flag changed")
        result_guards = analyze.decision_guards(
            result["native"]["analysis"], result["allocation"]["analysis"], result["policy"]
        )
        require(result_guards == guards and guards["adoption_eligible"],
                "retained candidate violates frozen adoption policy")
    return {
        "reports": result["counts"]["reports"],
        "samples": result["counts"]["samples"],
        "native_spread_flags": len(result["native"]["analysis"]["spread_flags_over_5_percent"]),
        "native_regression_flags": len(result["native"]["analysis"]["regression_flags_over_5_percent"]),
        "allocation_spread_flags": len(result["allocation"]["analysis"]["spread_flags_over_5_percent"]),
        "allocation_regression_flags": len(result["allocation"]["analysis"]["regression_flags_over_5_percent"]),
        "disposition": status,
        "seal_checked": seal is not None,
        "workspace_checked": workspace is not None,
    }


if __name__ == "__main__":
    try:
        arguments = set(sys.argv[1:])
        allowed = {"--require-final-seal", "--check-workspace"}
        analyze.require(arguments <= allowed, "unknown validator argument")
        print(json.dumps(validate(require_final_seal="--require-final-seal" in arguments,
                                  check_workspace="--check-workspace" in arguments),
                        indent=2, sort_keys=True))
    except analyze.ReplayError as error:
        print(f"validation failed: {error}", file=sys.stderr)
        raise SystemExit(1)
