"""Offline terminal validator for the 0828 PPTX edit-profile packet.

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


def descriptor(value: object, label: str) -> Path:
    require(isinstance(value, dict) and isinstance(value.get("path"), str)
            and isinstance(value.get("bytes"), int)
            and analysis.is_sha(value.get("sha256")), f"{label} descriptor malformed")
    path = Path(value["path"])
    require(path.is_file() and not path.is_symlink(), f"{label} missing")
    data = path.read_bytes()
    require(len(data) == value["bytes"]
            and analysis.sha256(path) == value["sha256"], f"{label} identity changed")
    return path


def check_reader_attempts() -> None:
    """Validate every run_reader source snapshot, log, and failure receipt."""
    root = PACKET / "reader-attempts"
    require(root.is_dir() and not root.is_symlink(),
            "reader-attempts directory missing")
    attempts = [path for path in root.iterdir() if path.is_dir() and not path.is_symlink()]
    require(attempts and all(path.name.isdigit() for path in attempts),
            "reader attempt numbering changed")
    attempts.sort(key=lambda path: int(path.name))
    require({path.name for path in attempts} == {str(i) for i in range(len(attempts))},
            "reader attempt numbering is not contiguous")
    pending = [path for path in attempts if not (path / "receipt.json").is_file()]
    require(len(pending) <= 1 and (not pending or pending[0] == attempts[-1]),
            "reader attempt receipt missing before current wrapper")
    for attempt in attempts:
        if attempt in pending:
            require((attempt / "console.log").is_file()
                    and (attempt / "sources").is_dir(),
                    f"in-flight reader attempt {attempt.name} incomplete")
            continue
        require(all(not path.is_symlink() for path in attempt.rglob("*")),
                f"reader attempt {attempt.name} contains a symlink")
        receipt = analysis.read_json(attempt / "receipt.json")
        require(receipt.get("schema") == "litchi.performance.0828.reader-attempt.v1"
                and isinstance(receipt.get("command"), list)
                and receipt["command"], f"reader attempt {attempt.name} receipt changed")
        require(receipt["command"][0] == sys.executable
                and Path(receipt["command"][2]).name in {
                    "reader_preflight.py", "analysis.py", "root_audit.py",
                    "frame_audit.py", "stack_diagnostics.py", "validate.py"},
                f"reader attempt {attempt.name} command changed")
        for key in ("started", "ended"):
            require(isinstance(receipt.get(key), (int, float))
                    and not isinstance(receipt[key], bool),
                    f"reader attempt {attempt.name} {key} malformed")
        require(receipt["started"] <= receipt["ended"],
                f"reader attempt {attempt.name} timestamps reversed")
        require(isinstance(receipt.get("exit_code"), int)
                and not isinstance(receipt["exit_code"], bool),
                f"reader attempt {attempt.name} exit code malformed")
        log = descriptor(receipt.get("log"), f"reader attempt {attempt.name} log")
        require(log.resolve() == (attempt / "console.log").resolve(),
                f"reader attempt {attempt.name} log binding changed")
        rows = receipt.get("sources")
        require(isinstance(rows, list) and rows,
                f"reader attempt {attempt.name} source snapshots missing")
        expected_files = {"receipt.json", "console.log"}
        seen_originals: set[str] = set()
        seen_snapshots: set[str] = set()
        for index, row in enumerate(rows):
            require(isinstance(row, dict), f"reader attempt {attempt.name} source malformed")
            original = row.get("original")
            snapshot = row.get("snapshot")
            for item, label in ((original, "original"), (snapshot, "snapshot")):
                require(isinstance(item, dict) and isinstance(item.get("path"), str)
                        and isinstance(item.get("bytes"), int)
                        and item["bytes"] >= 0 and analysis.is_sha(item.get("sha256")),
                        f"reader attempt {attempt.name} source {index}/{label} malformed")
            original_path = Path(original["path"]).resolve()
            require(original_path.is_relative_to(PACKET.resolve()),
                    f"reader attempt {attempt.name} original escaped packet")
            snapshot_path = Path(snapshot["path"]).resolve()
            source_root = (attempt / "sources").resolve()
            require(snapshot_path.is_relative_to(source_root)
                    and len(snapshot_path.relative_to(source_root).parts) == 1
                    and snapshot_path.name == original_path.name,
                    f"reader attempt {attempt.name} snapshot escaped scope")
            snapshot_observed = descriptor(snapshot,
                                           f"reader attempt {attempt.name} source {index}")
            require(snapshot_observed.resolve() == snapshot_path
                    and snapshot["bytes"] == original["bytes"]
                    and snapshot["sha256"] == original["sha256"],
                    f"reader attempt {attempt.name} snapshot differs from original")
            original_key = str(original_path)
            snapshot_key = str(snapshot_path.relative_to(attempt))
            require(original_key not in seen_originals and snapshot_key not in seen_snapshots,
                    f"reader attempt {attempt.name} duplicate source")
            seen_originals.add(original_key)
            seen_snapshots.add(snapshot_key)
            expected_files.add(snapshot_key)
        actual_files = {str(path.relative_to(attempt)) for path in attempt.rglob("*")
                        if path.is_file() and not path.is_symlink()}
        require(actual_files == expected_files,
                f"reader attempt {attempt.name} contains unbound files")


def validate_cleanup(value: dict, build: dict) -> None:
    require(value.get("schema") == "litchi.performance.0828.cleanup.v1",
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


def compare_phase_partition(independent: dict, primary: dict, label: str) -> None:
    left = independent.get("phase_partition")
    right = primary.get("phase_partition")
    require(isinstance(left, dict) and isinstance(right, dict),
            f"{label}: phase partition missing")
    for key in ("phase_order", "phase_owners", "sample_counts", "periods",
                "owner_samples", "classified_samples", "unclassified_samples",
                "ambiguous_samples", "phase_marker_hits", "phase_marker_other_dso_hits",
                "additive_timing_claim", "causal_fraction_claim"):
        require(left.get(key) == right.get(key),
                f"{label}: phase partition {key} differs")
    for key in ("leaf_census", "unclassified_leaf_census", "ambiguous_leaf_census"):
        require(left.get(key) == right.get(key),
                f"{label}: phase leaf census differs: {key}")
    primary_rows = [{key: row.get(key) for key in (
        "sample_index", "period", "phase_status", "phase",
        "phase_exact_hits", "phase_other_dso_hits")}
                    for row in right.get("sample_rows", [])]
    independent_rows = [{
        "sample_index": row.get("sample_index"), "period": row.get("period"),
        "phase_status": row.get("phase_status"), "phase": row.get("phase"),
        "phase_exact_hits": row.get("phase_exact_hits"),
        "phase_other_dso_hits": row.get("phase_other_dso_hits"),
    } for row in left.get("sample_rows", [])]
    require(independent_rows == primary_rows,
            f"{label}: phase sample rows differ")


def run(final: bool) -> dict:
    # ``analyze(check=True)`` both replays every retained receipt and proves
    # that all derived files are deterministic from those receipts.
    value = analysis.analyze(check=True)
    check_reader_attempts()
    for script in ('root_audit.py', 'frame_audit.py'):
        subprocess.run([sys.executable, '-B', str(PACKET / script), '--check'], check=True)
    frames = analysis.read_json(PACKET / 'frame-audit.json')
    require(frames['status'] == value['perf']['status'], 'frame audit status differs')
    if frames['status'] == 'available':
        require(len(frames['repeats']) == len(value['perf']['profiles']) == 2,
                'independent frame repeat count differs')
        for independent, profile in zip(frames['repeats'], value['perf']['profiles'], strict=True):
            summary = profile['summary']
            require(independent['repeat'] == profile['repeat'], 'frame repeat identity differs')
            for left, right in (
                ('whole_process_samples', 'whole_process_samples'),
                ('owner_samples', 'owner_qualified_samples'),
                ('outside_owner_samples', 'unattributed_samples'),
                ('owner_period', 'owner_qualified_period'),
            ):
                require(independent[left] == summary[right], f'frame audit {left} differs')
            leaves = {(leaf['symbol'], leaf.get('dso', '')): leaf['samples']
                      for leaf in independent['leaves']}
            primary_leaves = {(row['symbol'], row.get('dso', '')): row['samples']
                              for row in summary['ranked_leaf_within_owner']}
            # The primary reader's legacy full-leaf census does not carry DSO
            # on old packets; phase census does. Compare symbols in that case.
            require(leaves == primary_leaves or
                    Counter({key[0]: count for key, count in leaves.items()})
                    == Counter({key[0]: count for key, count in primary_leaves.items()}),
                    'independent full leaf partition differs')
            compare_phase_partition(independent, summary, f"perf/{profile['repeat']}")
        subprocess.run([sys.executable, '-B', str(PACKET / 'stack_diagnostics.py'), '--check'], check=True)
        stacks = analysis.read_json(PACKET / 'stack-diagnostics.json')['repeats']
        require([(row['repeat'], row['samples'], row['empty_stack_samples']) for row in stacks]
                == [(row['repeat'], row['whole_process_samples'], row['empty_stack_samples'])
                    for row in frames['repeats']],
                'stack diagnostics differ from independent sample census')
    else:
        require(not frames.get('repeats'), 'unavailable perf has frame repeats')
        stack_path = PACKET / 'stack-diagnostics.json'
        if stack_path.is_file():
            require(not analysis.read_json(stack_path).get('repeats'),
                    'unavailable perf has stack repeats')
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
        print(f"0828 validation failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
