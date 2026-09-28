"""Offline terminal validator for the 0822 PPTX edit-profile packet.

The validator delegates all evidence parsing to ``analysis.py`` and adds the
independent-audit and post-cleanup gates used by the final seal.  It never
starts a build, workload, profiler, decoder, or binary inspection command.
"""

from __future__ import annotations

import argparse
from collections import Counter
import json
import subprocess
import sys
from pathlib import Path

import analysis


PACKET = Path(__file__).resolve().parent


def require(condition: bool, message: str) -> None:
    if not condition:
        raise analysis.ReplayError(message)


def validate_cleanup(value: dict, build: dict) -> None:
    require(value.get("schema") == "litchi.performance.0822.cleanup.v1",
            "cleanup schema changed")
    require(value.get("target") == str(analysis.custody.TARGET)
            and value.get("target_removed") is True
            and value.get("scratch") is None
            and value.get("binaries_verified_before_removal") is True,
            "cleanup target/scratch witness changed")
    require(not analysis.custody.TARGET.exists(), "owned target still exists after cleanup")
    removed = value.get("removed_binaries")
    expected = [build["raw"]["binaries"][name]["artifact"] for name in ("ordinary", "fp")]
    require(isinstance(removed, list) and len(removed) == 2
            and set(json.dumps(row, sort_keys=True) for row in removed)
            == set(json.dumps(row, sort_keys=True) for row in expected),
            "cleanup does not account for exactly ordinary/fp binaries")
    require(value.get("source") == build["raw"].get("source")
            and value.get("frozen_inputs") == build["raw"].get("frozen_inputs"),
            "cleanup source/frozen-input witness changed")
    directories = value.get("removed")
    require(isinstance(directories, list) and len(directories) == 1
            and directories[0].get("path") == str(analysis.custody.TARGET),
            "cleanup removed-directory witness changed")
    analysis.finite(value.get("started"), "cleanup start")
    analysis.finite(value.get("ended"), "cleanup end")
    require(value["started"] <= value["ended"], "cleanup timestamps reversed")


def validate_final(value: dict) -> None:
    require(value.get("status") == "accepted", "analysis status is not accepted")
    counts = value["counts"]
    require(value["expected_counts"] == {
        "qualification_reports": 3, "qualification_samples": 9,
        "native_reports": 18, "native_samples": 540,
        "perf_reports": 2, "perf_samples": 4000,
        "total_reports": 23, "total_samples": 4549,
    }, "planned cardinality changed")
    require(counts["qualification_reports"] == 3
            and counts["qualification_samples"] == 9
            and counts["native_reports"] == 18
            and counts["native_samples"] == 540,
            "native/qualification cardinality changed")
    require(value["verification"]["historical_timing_comparison_omitted"] is True
            and value["verification"]["shipping_latency_claim_omitted"] is True
            and value["verification"]["amdahl_fraction_omitted"] is True,
            "analysis claim boundary changed")
    audit = analysis.compare_root_audit(value["native"])
    require(audit is not None, "independent root-audit.json is required for final validation")
    require(value["verification"]["independent_root_audit_present"] is True,
            "analysis did not observe the independent root audit")
    require((PACKET / "results-review.md").is_file()
            and not (PACKET / "results-review.md").is_symlink(),
            "results-review.md is required before cleanup/seal")
    cv = analysis.load_custody()
    quality = analysis.load_quality(cv)
    build = analysis.load_build(cv, quality)
    cleanup = analysis.cleanup_value()
    require(cleanup is not None, "cleanup.json is required for final validation")
    validate_cleanup(cleanup, build)
    perf = value["perf"]
    require(perf["status"] in {"available", "unavailable"}
            and perf.get("profile_fabricated") is False,
            "perf terminal status changed")


def run(final: bool) -> dict:
    # ``analyze(check=True)`` both replays every retained receipt and proves
    # that all derived files are deterministic from those receipts.
    value = analysis.analyze(check=True)
    for script in ('root_audit.py', 'frame_audit.py'):
        subprocess.run([sys.executable, '-B', str(PACKET / script), '--check'], check=True)
    frames = analysis.read_json(PACKET / 'frame-audit.json')
    require(frames['status'] == value['perf']['status'], 'frame audit status differs')
    if frames['status'] == 'available':
        subprocess.run([sys.executable, '-B', str(PACKET / 'stack_diagnostics.py'), '--check'], check=True)
        stacks = analysis.read_json(PACKET / 'stack-diagnostics.json')['repeats']
        require([(r['repeat'], r['samples'], r['empty_stack_samples']) for r in stacks]
                == [(r['repeat'], r['whole_samples'], r['empty_stack_samples']) for r in frames['repeats']],
                'stack diagnostics differ from independent sample census')
        require(len(frames['repeats']) == len(value['perf']['profiles']) == 2,
                'independent frame repeat count differs')
        for independent, profile in zip(frames['repeats'], value['perf']['profiles'], strict=True):
            summary = profile['summary']
            require(independent['repeat'] == profile['repeat'], 'frame repeat identity differs')
            for left, right in (
                ('whole_samples', 'whole_process_samples'),
                ('owner_samples', 'owner_qualified_samples'),
                ('outside_owner_samples', 'unattributed_samples'),
                ('owner_period', 'owner_qualified_period'),
                ('owner_self_leaves', 'qualified_stacks_without_descendant'),
            ):
                require(independent[left] == summary[right], f'frame audit {left} differs')
            leaves = Counter()
            for leaf in independent['leaves']:
                leaves[leaf['symbol']] += leaf['samples']
            require(dict(leaves) == {row['symbol']: row['samples']
                                     for row in summary['ranked_leaf_within_owner']},
                    'independent full leaf partition differs')
    if final:
        validate_final(value)
    return value


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--final", action="store_true",
                        help="apply the post-cleanup and independent-audit seal gates")
    args = parser.parse_args(argv)
    try:
        value = run(args.final)
        print(json.dumps({"status": value["status"], "reports": value["counts"]["reports"],
                          "samples": value["counts"]["samples"],
                          "final": args.final}, sort_keys=True))
    except (analysis.ReplayError, AssertionError, OSError, ValueError, KeyError,
            TypeError, IndexError) as error:
        print(f"0822 validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
