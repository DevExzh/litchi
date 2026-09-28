"""Final offline validator for the 0827 ordinary-save comparison packet."""

from __future__ import annotations

import json
import hashlib
import sys
from pathlib import Path
from typing import Any

import analysis
import raw_audit
import custody_audit


PACKET = analysis.PACKET
ROOT = analysis.ROOT


def require(condition: bool, message: str) -> None:
    if not condition:
        raise analysis.ReplayError(message)


def descriptor(value: Any, label: str, *, packet_bound: bool = True) -> Path:
    path = analysis.descriptor(value, label, packet_bound=packet_bound)
    require(path is not None, f"{label} missing")
    return path


def check_counts(value: dict[str, Any]) -> None:
    require(value.get("counts") == {
        "qualification_reports": 24, "qualification_samples": 24,
        "native_reports": 144, "native_samples": 4320,
        "observer_reports": 48, "observer_samples": 144,
        "reports": 216, "samples": 4488,
    }, "aggregate counts changed")
    require(len(value.get("native_rows", [])) == 12
            and len(value.get("observer_rows", [])) == 12,
            "summary row cardinality changed")
    expected = {case["case"] for case in analysis.CASES}
    require({row.get("case") for row in value["native_rows"]} == expected
            and {row.get("case") for row in value["observer_rows"]} == expected,
            "summary case coverage changed")


def check_policy(value: dict[str, Any]) -> None:
    decision = value.get("decision_flags")
    require(isinstance(decision, dict)
            and decision.get("adoption_decision") is None
            and isinstance(decision.get("p50_regression_flags"), list)
            and isinstance(decision.get("allocation_increase_flags"), list)
            and isinstance(decision.get("rss_review_flags"), list),
            "ordinary-save packet contains an adoption decision")
    excluded = decision.get("claims_excluded", [])
    for claim in ("adoption", "general benefit", "unsupported general full-save claim",
                  "additive phase medians", "pooled native/observer latency"):
        require(claim in excluded, f"claim exclusion missing: {claim}")


def check_native(value: dict[str, Any]) -> None:
    for row in value["native_rows"]:
        require(row["before"]["blocks"] == 6 and row["after"]["blocks"] == 6,
                f"native blocks changed: {row['case']}")
        pair = row["paired"]
        require(set(pair["metrics"]) == {"p50", "p95", "p99", "mean", "rss_kib"},
                f"native paired metrics changed: {row['case']}")
        for metric, detail in pair["metrics"].items():
            require(len(detail["by_block"]) == 6
                    and isinstance(detail.get("median_ratio"), (int, float)),
                    f"native paired blocks changed: {row['case']}/{metric}")
            boot = detail.get("bootstrap")
            require(isinstance(boot, dict) and boot.get("seed") == analysis.BOOTSTRAP_SEED
                    and boot.get("resamples") == analysis.BOOTSTRAP_RESAMPLES
                    and boot.get("ci_low") is not None and boot.get("ci_high") is not None,
                    f"native CI changed: {row['case']}/{metric}")
        for leg in analysis.LEGS:
            require(set(row[leg]["process_median"]) == {"p50", "p95", "p99", "mean"},
                    f"native summary metrics changed: {row['case']}/{leg}")
            require(set(row[leg]["spread_ratios"]) == {"p50", "p95", "p99", "mean"}
                    and isinstance(row[leg].get("rss_spread_ratio"), (int, float)),
                    f"native spreads changed: {row['case']}/{leg}")


def check_observer(value: dict[str, Any]) -> None:
    for row in value["observer_rows"]:
        for leg in analysis.LEGS:
            allocation = row[leg].get("allocation")
            require(isinstance(allocation, dict)
                    and set(allocation["median_over_blocks"]) == set(analysis.ALLOC_METRICS),
                    f"observer allocation metrics changed: {row['case']}/{leg}")
            processes = allocation.get("processes")
            require(isinstance(processes, list) and len(processes) == 2
                    and [item.get("block") for item in processes] == [0, 1],
                    f"observer allocation blocks changed: {row['case']}/{leg}")
            for process in processes:
                require(isinstance(process.get("values"), dict)
                        and set(process["values"]) == set(analysis.ALLOC_METRICS),
                        f"observer allocation vector names changed: {row['case']}/{leg}")
                for name in analysis.ALLOC_METRICS:
                    values = process["values"][name]
                    require(isinstance(values, list) and len(values) == 3,
                            f"observer allocation values changed: {row['case']}/{leg}/{name}")
                    valid = (all(isinstance(x, (int, float)) for x in values)
                             if name == "net_live" else
                             all(isinstance(x, (int, float)) and x >= 0 for x in values))
                    require(valid,
                            f"observer allocation values changed: {row['case']}/{leg}/{name}")
            diagnostics = row[leg].get("process_diagnostics")
            require(diagnostics is None or isinstance(diagnostics, dict),
                    f"observer process diagnostics malformed: {row['case']}/{leg}")
        require(set(row["paired"]["metrics"]) == set(analysis.OBSERVER_PAIR_METRICS),
                f"observer paired allocation metrics changed: {row['case']}")
        for name in analysis.OBSERVER_PAIR_METRICS:
            pair = row["paired"]["metrics"][name]
            for side in ("before", "after"):
                vectors = pair.get("all_values", {}).get(side)
                width = 1 if name == "rss_kib" else 3
                require(isinstance(vectors, list) and len(vectors) == 2
                        and all(isinstance(vector, list) and len(vector) == width
                                and all(isinstance(x, (int, float)) for x in vector)
                                for vector in vectors),
                        f"observer paired values changed: {row['case']}/{name}/{side}")
        rss = row["paired"]["metrics"]["rss_kib"]
        require(len(rss.get("by_block", [])) == 2
                and isinstance(rss.get("median_ratio"), (int, float))
                and rss["median_ratio"] > 0,
                f"observer paired RSS changed: {row['case']}")


def check_failed_invocations() -> None:
    """Every retained failed command has a unique immutable log."""
    candidates = list(PACKET.glob("quality-*/receipt.json"))
    candidates += list(PACKET.glob("build-*/receipt.json"))
    candidates += [PACKET / name for name in
                   ("reader-failures.json", "capture-failures.json", "failure-receipts.json")]
    seen: set[str] = set()
    for receipt_path in candidates:
        if not receipt_path.is_file() or receipt_path.is_symlink():
            continue
        value = analysis.read_json(receipt_path)
        for index, row in enumerate(value.get("rows", value.get("failures", []))):
            require(isinstance(row, dict), f"failed invocation {receipt_path}/{index} malformed")
            if row.get("exit_code", 1) == 0:
                continue
            log = row.get("log") or row.get("console_log")
            path = descriptor(log, f"failed invocation {receipt_path}/{index} log")
            require(str(path) not in seen, f"failed invocation log reused: {path}")
            seen.add(str(path))


def check_reader_attempts() -> None:
    """Bind every root-owned reader attempt, including failed attempts.

    ``run_reader.py`` snapshots the packet Python sources and the shared
    allocation-schema helper before starting each reader.  The original
    descriptors are historical observations: their live files may have been
    repaired later, so snapshot bytes are checked against the descriptor
    recorded in that attempt rather than against today's source tree.
    """
    root = PACKET / "reader-attempts"
    require(root.is_dir() and not root.is_symlink(),
            "reader attempt directory missing")
    require(all(path.is_dir() and not path.is_symlink() for path in root.iterdir()),
            "reader attempt root contains an unexpected entry")
    all_attempt_dirs = [path for path in root.iterdir()
                        if path.is_dir() and not path.is_symlink()]
    require(all_attempt_dirs and all(path.name.isdigit() for path in all_attempt_dirs)
            and {path.name for path in all_attempt_dirs} ==
            {str(index) for index in range(len(all_attempt_dirs))},
            "reader attempt numbering changed")
    all_attempt_dirs.sort(key=lambda path: int(path.name))
    # The wrapper persists its receipt only after the child exits.  While this
    # validator is the child, allow that one highest numbered in-flight
    # directory; any earlier or multiple receipt-less attempts are evidence
    # loss and remain fatal.  A subsequent wrapped invocation binds this
    # directory through its completed receipt.
    pending = [path for path in all_attempt_dirs if not (path / "receipt.json").is_file()]
    require(len(pending) <= 1 and (not pending or pending[0] == all_attempt_dirs[-1]),
            "reader attempt receipt missing before the current wrapper")
    if pending:
        current = pending[0]
        require((current / "console.log").is_file()
                and not (current / "console.log").is_symlink()
                and (current / "sources").is_dir()
                and not (current / "sources").is_symlink(),
                "in-flight reader attempt is incomplete")
    attempt_dirs = [path for path in all_attempt_dirs if path not in pending]
    for attempt in attempt_dirs:
        require(all(not path.is_symlink() for path in attempt.rglob("*")),
                f"reader attempt {attempt.name} contains a symlink")
        require((attempt / "sources").is_dir()
                and not (attempt / "sources").is_symlink(),
                f"reader attempt {attempt.name} source directory missing")
        receipt_path = attempt / "receipt.json"
        require(receipt_path.is_file() and not receipt_path.is_symlink(),
                f"reader attempt receipt missing: {attempt.name}")
        receipt = analysis.read_json(receipt_path)
        require(receipt.get("schema") == "litchi.performance.0827.reader-attempt.v1"
                and isinstance(receipt.get("command"), list)
                and receipt["command"], f"reader attempt {attempt.name} receipt changed")
        for key in ("started", "ended"):
            require(isinstance(receipt.get(key), (int, float))
                    and not isinstance(receipt[key], bool),
                    f"reader attempt {attempt.name} {key} changed")
        require(receipt["started"] <= receipt["ended"],
                f"reader attempt {attempt.name} timestamps reversed")
        require(isinstance(receipt.get("exit_code"), int)
                and not isinstance(receipt["exit_code"], bool),
                f"reader attempt {attempt.name} exit code changed")
        log = descriptor(receipt.get("log"), f"reader attempt {attempt.name} log")
        require(log.resolve() == (attempt / "console.log").resolve(),
                f"reader attempt {attempt.name} log path changed")
        rows = receipt.get("sources")
        require(isinstance(rows, list) and rows,
                f"reader attempt {attempt.name} source snapshots missing")
        expected_files = {"receipt.json", "console.log"}
        seen_snapshots: set[str] = set()
        seen_originals: set[str] = set()
        for index, row in enumerate(rows):
            require(isinstance(row, dict),
                    f"reader attempt {attempt.name} source {index} malformed")
            original = row.get("original")
            snapshot = row.get("snapshot")
            for value, label in ((original, "original"), (snapshot, "snapshot")):
                require(isinstance(value, dict) and isinstance(value.get("path"), str)
                        and isinstance(value.get("bytes"), int) and value["bytes"] >= 0
                        and analysis.is_sha(value.get("sha256")),
                        f"reader attempt {attempt.name} source {index}/{label} malformed")
            original_path = Path(original["path"]).resolve()
            packet_source = PACKET.resolve()
            helper = (ROOT / "tools/perf_allocation_schema.py").resolve()
            require(original_path.is_relative_to(packet_source)
                    or original_path == helper,
                    f"reader attempt {attempt.name} source {index} original escaped scope")
            snapshot_path = Path(snapshot["path"]).resolve()
            source_root = (attempt / "sources").resolve()
            require(snapshot_path.is_relative_to(source_root)
                    and len(snapshot_path.relative_to(source_root).parts) == 1,
                    f"reader attempt {attempt.name} source {index} snapshot escaped scope")
            require(snapshot_path.name == original_path.name,
                    f"reader attempt {attempt.name} source {index} basename changed")
            snapshot_observed = descriptor(snapshot,
                                           f"reader attempt {attempt.name} source {index}/snapshot")
            require(snapshot_observed.resolve() == snapshot_path
                    and snapshot["bytes"] == original["bytes"]
                    and snapshot["sha256"] == original["sha256"],
                    f"reader attempt {attempt.name} source {index} snapshot differs from original descriptor")
            snapshot_key = str(snapshot_path.relative_to(attempt))
            require(snapshot_key not in seen_snapshots,
                    f"reader attempt {attempt.name} duplicate snapshot")
            require(str(original_path) not in seen_originals,
                    f"reader attempt {attempt.name} duplicate original")
            seen_snapshots.add(snapshot_key)
            seen_originals.add(str(original_path))
            expected_files.add(snapshot_key)
        actual_files = {str(path.relative_to(attempt)) for path in attempt.rglob("*")
                        if path.is_file() and not path.is_symlink()}
        require(actual_files == expected_files,
                f"reader attempt {attempt.name} contains unbound files")


def check_cleanup(value: dict[str, Any]) -> None:
    path = PACKET / "cleanup.json"
    require(path.is_file() and not path.is_symlink(), "cleanup witness missing")
    cleanup = analysis.read_json(path)
    require(cleanup.get("schema") == "litchi.performance.0827.cleanup.v1"
            and cleanup.get("verified") is True
            and cleanup.get("target_absent_after_removal") is True
            and cleanup.get("scratch_absent_after_removal") is True,
            "cleanup status changed")
    require(not analysis.TARGET.exists() and not analysis.SCRATCH.exists(),
            "owned target or scratch remains")
    roots = cleanup.get("roots")
    require(isinstance(roots, list) and len(roots) == 2,
            "cleanup roots changed")
    require({row.get("path") for row in roots} == {str(analysis.TARGET), str(analysis.SCRATCH)},
            "cleanup root paths changed")
    for row in roots:
        require(isinstance(row, dict)
                and isinstance(row.get("files"), int) and row["files"] >= 0
                and isinstance(row.get("logical_bytes"), int) and row["logical_bytes"] >= 0,
                "cleanup root inventory changed")
    binaries = cleanup.get("binaries")
    require(isinstance(binaries, list) and len(binaries) == 6,
            "cleanup binary cardinality changed")
    expected = []
    for leg in analysis.LEGS:
        build = value["build"][leg]
        build_receipt = analysis.read_json(descriptor(build, f"{leg} build receipt"))
        for item in build_receipt.get("binaries", {}).values():
            raw = item.get("artifact", item) if isinstance(item, dict) else item
            expected.append({key: raw[key] for key in ("path", "bytes", "sha256")})
    observed = [{key: row[key] for key in ("path", "bytes", "sha256")} for row in binaries]
    require(sorted(observed, key=lambda x: x["path"]) ==
            sorted(expected, key=lambda x: x["path"]), "cleanup binaries changed")
    marker = cleanup.get("scratch_marker_sha256")
    require(isinstance(marker, str) and len(marker) == 64, "scratch marker witness missing")
    plan = analysis.read_json(PACKET / "plan.json")
    expected_marker = hashlib.sha256(plan["scratch_marker"].encode()).hexdigest()
    require(marker == expected_marker, "scratch marker witness changed")
    finite = cleanup.get("started"), cleanup.get("ended")
    require(all(isinstance(x, (int, float)) for x in finite) and finite[0] <= finite[1],
            "cleanup timestamps changed")


def check_live_after(value: dict[str, Any]) -> None:
    after = value["source"]["after"]
    current = analysis.live_source()
    require(current == after, "live production source differs from frozen after source")
    for name in analysis.SOURCE_ALLOWLIST:
        require(current[name] == after[name], f"live after source differs: {name}")


def independent_readers() -> None:
    raw = raw_audit.derive()
    raw_audit.compare(raw)
    path = PACKET / "raw-audit.json"
    require(path.is_file() and path.read_text(encoding="utf-8") ==
            json.dumps(raw, indent=2, sort_keys=True) + "\n",
            "raw-audit.json does not replay byte-for-byte")


def run(*, final: bool = False) -> dict[str, Any]:
    custody_audit.check()
    value = analysis.analyze(check=True)
    check_counts(value)
    check_policy(value)
    check_native(value)
    check_observer(value)
    check_failed_invocations()
    check_reader_attempts()
    independent_readers()
    if final:
        check_live_after(value)
        check_cleanup(value)
    return {"status": "accepted", "reports": value["counts"]["reports"],
            "samples": value["counts"]["samples"], "final": final}


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    require(args in ([], ["--final"]), "use no arguments or --final")
    try:
        print(json.dumps(run(final=args == ["--final"]), sort_keys=True))
        return 0
    except (analysis.ReplayError, raw_audit.AuditError, AssertionError, OSError,
            ValueError, KeyError, TypeError, IndexError) as error:
        print(f"0827 validation failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
