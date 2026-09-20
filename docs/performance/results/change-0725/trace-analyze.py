#!/usr/bin/env python3
"""Audit 0725 diagnostic route traces independently of timing evidence."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import subprocess
from pathlib import Path
from typing import Any


ROOT = Path(__file__).resolve().parents[4]
PACKET = Path(__file__).resolve().parent
BASELINE_REF = "ee5e0b0650"
TRACE_PATHS = (
    "crates/litchi-cfb/src/shared.rs",
    "crates/litchi-xls/src/workbook/query_cache.rs",
    "crates/litchi-xls/src/workbook/source.rs",
)

QUERY_RE = re.compile(r"^TRACE query sheet=(\d+) row=(\d+) column=(\d+)$")
CALL_RE = re.compile(r"^TRACE call route=([a-z0-9-]+) line=(\d+)$")
CHAIN_RE = re.compile(
    r"^TRACE chain table=([A-Za-z0-9]+) from=(\d+) to=(\d+) links=(\d+)$"
)
TIMING_RE = re.compile(r"(?i)(?:elapsed|nanos|timing|seconds)")


def sha_file(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def check_binary(path: Path, digest: str) -> None:
    if path.exists():
        assert sha_file(path) == digest, path
        return
    cleanup = read_json(PACKET / "cleanup.json")
    assert cleanup["removed"] is True
    matches = [item for item in cleanup["identities"]
               if item["path"] == str(path) and item["sha256"] == digest
               and isinstance(item["bytes"], int) and item["bytes"] > 0]
    assert len(matches) == 1, f"missing exact binary cleanup witness: {path}"


def read_json(path: Path) -> Any:
    return json.loads(path.read_text())


def git_sha(revision: str, relative: str) -> str:
    return hashlib.sha256(
        subprocess.check_output(["git", "show", f"{revision}:{relative}"], cwd=ROOT)
    ).hexdigest()


def baseline_head() -> str:
    return subprocess.check_output(
        ["git", "rev-parse", BASELINE_REF], cwd=ROOT, text=True
    ).strip()


def check_no_timing(value: Any, where: str) -> None:
    if isinstance(value, dict):
        for key, item in value.items():
            assert not TIMING_RE.search(str(key)), f"timing field {where}.{key}"
            check_no_timing(item, f"{where}.{key}")
    elif isinstance(value, list):
        for index, item in enumerate(value):
            check_no_timing(item, f"{where}[{index}]")
    elif isinstance(value, str):
        assert not TIMING_RE.search(value), f"timing text in {where}"


def audit_manifest(phase: str, folder: Path, baseline_head: str) -> dict[str, Any]:
    manifest = read_json(folder / "manifest.json")
    assert manifest["phase"] == phase
    assert manifest["baseline_head"] == baseline_head
    assert manifest["source_paths"] == list(TRACE_PATHS)
    assert manifest["restored"] is True
    for relative, digest in manifest["probe_sha256"].items():
        assert sha_file(ROOT / relative) == digest, relative
    for relative, digest in manifest["sequence_probe_sha256"].items():
        assert sha_file(ROOT / relative) == digest, relative
    for relative in TRACE_PATHS:
        assert manifest["restored_source_sha256"][relative] == manifest[
            "original_source_sha256"
        ][relative]
        assert manifest["instrumented_source_sha256"][relative] != manifest[
            "effective_source_sha256"
        ][relative]
    if phase == "baseline":
        for relative in TRACE_PATHS:
            assert manifest["effective_source_sha256"][relative] == git_sha(
                baseline_head, relative
            )
    else:
        for relative in TRACE_PATHS:
            assert manifest["effective_source_sha256"][relative] == manifest[
                "original_source_sha256"
            ][relative]

    binary = Path(manifest["binary"])
    check_binary(binary, manifest["binary_sha256"])
    assert manifest["outputs"]
    for output in manifest["outputs"]:
        trace = folder / output["trace"]
        semantic = folder / output["semantic"]
        assert sha_file(trace) == output["trace_sha256"]
        assert sha_file(semantic) == output["semantic_sha256"]
        trace_text = trace.read_text()
        assert not TIMING_RE.search(trace_text), f"timing text in {trace}"
        check_no_timing(read_json(semantic), str(semantic))
    sequence_output = manifest["sequence_output"]
    sequence_trace = folder / sequence_output["trace"]
    sequence_semantic = folder / sequence_output["semantic"]
    assert sha_file(sequence_trace) == sequence_output["trace_sha256"]
    assert sha_file(sequence_semantic) == sequence_output["semantic_sha256"]
    assert not TIMING_RE.search(sequence_trace.read_text())
    check_no_timing(read_json(sequence_semantic), str(sequence_semantic))
    sequence_binary = Path(manifest["sequence_binary"])
    check_binary(sequence_binary, manifest["sequence_binary_sha256"])
    for command in manifest["commands"]:
        assert command["exit_code"] == 0, command
    return manifest


def parse_trace(path: Path) -> dict[str, Any]:
    """Parse route markers and attach each chain walk to its preceding call."""

    queries: list[dict[str, Any]] = []
    prequery: dict[str, Any] = {"header": "prequery", "calls": [], "chains": []}
    current = prequery
    pending = "unlabeled"
    for raw in path.read_text().splitlines():
        if TIMING_RE.search(raw):
            raise AssertionError(f"timing text in trace {path}: {raw}")
        query = QUERY_RE.fullmatch(raw)
        if query:
            sheet, row, column = map(int, query.groups())
            current = {
                "header": raw,
                "sheet": sheet,
                "row": row,
                "column": column,
                "calls": [],
                "chains": [],
            }
            queries.append(current)
            pending = "unlabeled"
            continue
        call = CALL_RE.fullmatch(raw)
        if call:
            route, line = call.groups()
            pending = route
            current["calls"].append({"route": route, "line": int(line)})
            continue
        chain = CHAIN_RE.fullmatch(raw)
        if chain:
            table, start, end, links = chain.groups()
            start, end, links = map(int, (start, end, links))
            assert links == max(end - start, 0), raw
            current["chains"].append(
                {
                    "route": pending,
                    "table": table,
                    "from": start,
                    "to": end,
                    "links": links,
                }
            )
            pending = "unlabeled"
            continue
        # Probe errors are useful refusal evidence.  Keep them in the raw
        # trace, but do not mistake ordinary prose for a route record.
    return {"prequery": prequery, "queries": queries}


def route_vectors(parsed: dict[str, Any], route: str) -> list[list[int]]:
    return [
        [item["links"] for item in query["chains"] if item["route"] == route]
        for query in parsed["queries"]
    ]


def owner_ranges(parsed: dict[str, Any]) -> list[tuple[int, int]]:
    """Group trace queries by the fresh owner that emitted them."""

    queries = parsed["queries"]
    if not queries:
        return []
    starts = [0]
    for index in range(1, len(queries)):
        previous_scanned = any(
            call["route"] == "worksheet-scan" for call in queries[index - 1]["calls"]
        )
        current_scanned = any(
            call["route"] == "worksheet-scan" for call in queries[index]["calls"]
        )
        if current_scanned and not previous_scanned:
            starts.append(index)
    return [
        (start, end)
        for start, end in zip(starts, [*starts[1:], len(queries)])
    ]


def route_counts(parsed: dict[str, Any]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for query in [parsed["prequery"], *parsed["queries"]]:
        for item in query["chains"]:
            route = str(item["route"])
            counts[route] = counts.get(route, 0) + 1
    return counts


def route_summary(parsed: dict[str, Any]) -> dict[str, Any]:
    return {
        "query_count": len(parsed["queries"]),
        "prequery_chains": parsed["prequery"]["chains"],
        "queries": [
            {
                "header": query["header"],
                "calls": [call["route"] for call in query["calls"]],
                "chains": query["chains"],
            }
            for query in parsed["queries"]
        ],
        "route_counts": route_counts(parsed),
    }


def compare_case(name: str, before: dict[str, Any], after: dict[str, Any], budget: int) -> dict[str, Any]:
    assert before["prequery"]["chains"] == after["prequery"]["chains"], name
    assert len(before["queries"]) == len(after["queries"]), name
    assert [q["header"] for q in before["queries"]] == [q["header"] for q in after["queries"]], name
    before_owners = owner_ranges(before)
    after_owners = owner_ranges(after)
    assert before_owners == after_owners, (name, before_owners, after_owners)

    # The candidate adds one route during index construction.  All other
    # source-side call routes must remain in the same order per query, after
    # that explicitly named extra route is removed.
    for old, new in zip(before["queries"], after["queries"]):
        old_routes = [call["route"] for call in old["calls"]]
        new_routes = [call["route"] for call in new["calls"]]
        new_routes = [route for route in new_routes if route != "worksheet-checkpoint-build"]
        assert old_routes == new_routes, (name, old_routes, new_routes)

        old_sst = [item for item in old["chains"] if item["route"].startswith("sst-")]
        new_sst = [item for item in new["chains"] if item["route"].startswith("sst-")]
        assert old_sst == new_sst, (name, old_sst, new_sst)

    old_replay = route_vectors(before, "worksheet-replay")
    new_replay = route_vectors(after, "worksheet-replay")
    assert len(old_replay) == len(new_replay)
    for old, new in zip(old_replay, new_replay):
        assert len(old) == len(new), (name, old, new)
        assert all(candidate <= baseline for baseline, candidate in zip(old, new)), (
            name,
            old,
            new,
        )

    old_scan = route_vectors(before, "worksheet-scan")
    new_scan = route_vectors(after, "worksheet-scan")
    assert old_scan == new_scan, (name, old_scan, new_scan)

    old_replay_links = [link for vector in old_replay for link in vector]
    new_replay_links = [link for vector in new_replay for link in vector]
    build_vectors = route_vectors(after, "worksheet-checkpoint-build")
    new_build_links = [link for vector in build_vectors for link in vector]
    old_replay_query_count = sum(bool(vector) for vector in old_replay)
    new_replay_query_count = sum(bool(vector) for vector in new_replay)
    build_query_count = sum(bool(vector) for vector in build_vectors)
    if budget:
        # Whenever the baseline had a replay walk, the candidate must either
        # reuse the exact checkpoint (zero links) or take the old fallback.
        # The bounded target-local experiment is expected to take the former
        # for indexed selected cells.
        if old_replay_links:
            assert new_replay_links == [0] * len(new_replay_links), (
                name,
                old_replay_links,
                new_replay_links,
            )
            assert new_build_links, (name, "candidate omitted checkpoint-build route")
            # A fresh owner builds one checkpoint, while the baseline replays
            # the same target across the remaining queries for that owner.
            # Compare each build query's complete walk with one original
            # target replay walk in the same owner group; do not equate the
            # number of build calls with repeated replay calls.
            old_replay_by_owner = route_vectors(before, "worksheet-replay")
            for start, end in before_owners:
                original_target_walks = [
                    vector for vector in old_replay_by_owner[start:end] if vector
                ]
                candidate_build_walks = [
                    vector for vector in build_vectors[start:end] if vector
                ]
                assert candidate_build_walks, (name, start, end, "missing owner build")
                assert original_target_walks, (
                    name,
                    start,
                    end,
                    "missing owner replay target",
                )
                for build_walk in candidate_build_walks:
                    assert build_walk in original_target_walks, (
                        name,
                        start,
                        end,
                        original_target_walks,
                        build_walk,
                    )
        else:
            assert not new_build_links, (name, new_build_links)
    else:
        assert not old_replay_links and not new_replay_links
        assert not new_build_links

    return {
        "case": name,
        "budget": budget,
        "baseline": route_summary(before),
        "candidate": route_summary(after),
        "worksheet_replay_links": {
            "baseline": old_replay_links,
            "candidate": new_replay_links,
        },
        "worksheet_checkpoint_build_links": new_build_links,
        "worksheet_replay_query_count": {
            "baseline": old_replay_query_count,
            "candidate": new_replay_query_count,
        },
        "worksheet_checkpoint_build_query_count": build_query_count,
        "owner_query_ranges": [list(bounds) for bounds in before_owners],
        "sst_links_equal": True,
    }


def compare_sequence(
    baseline_folder: Path,
    candidate_folder: Path,
) -> dict[str, Any]:
    before_report = read_json(baseline_folder / "sequence.semantic.json")
    after_report = read_json(candidate_folder / "sequence.semantic.json")
    # The source wrapper records every positional read range and len/version
    # call per query.  Exact report equality keeps the metadata-only checkpoint
    # seek from silently changing observable source I/O.
    assert before_report == after_report, "sequence semantic/I/O report changed"

    before = parse_trace(baseline_folder / "sequence.trace")
    after = parse_trace(candidate_folder / "sequence.trace")
    assert len(before["queries"]) == len(after["queries"]) == len(before_report["queries"])
    for parsed, report in zip(before["queries"], before_report["queries"]):
        assert parsed["row"] == report["row"]
        assert parsed["column"] == report["column"]
    for parsed, report in zip(after["queries"], after_report["queries"]):
        assert parsed["row"] == report["row"]
        assert parsed["column"] == report["column"]

    before_replay = route_vectors(before, "worksheet-replay")
    after_replay = route_vectors(after, "worksheet-replay")
    assert len(before_replay) >= 5 and len(after_replay) == len(before_replay)
    labels = [str(query["label"]) for query in before_report["queries"]]
    assert labels == [str(query["label"]) for query in after_report["queries"]]
    earlier_index = labels.index("first-earlier")
    origin_index = labels.index("origin-late")
    # The first late query has no indexed replay yet; optional intermediate
    # and later queries shift the origin-late index, so use semantic labels.
    late_build_index = labels.index("late-build")
    assert before_replay[late_build_index] == after_replay[late_build_index] or not before_replay[late_build_index]
    earlier_before = before_replay[earlier_index]
    earlier_after = after_replay[earlier_index]
    origin_before = before_replay[origin_index]
    origin_after = after_replay[origin_index]
    assert earlier_before and earlier_after and earlier_before == earlier_after
    assert origin_before and origin_after and all(link > 0 for link in origin_before)
    assert origin_after == [0] * len(origin_after)
    build_after = route_vectors(after, "worksheet-checkpoint-build")
    assert any(build_after), "sequence candidate did not trace checkpoint build"
    return {
        "semantic_equal": True,
        "io_ranges_equal": True,
        "earlier_replay_links": {
            "baseline": earlier_before,
            "candidate": earlier_after,
        },
        "origin_late_replay_links": {
            "baseline": origin_before,
            "candidate": origin_after,
        },
        "candidate_checkpoint_build_links": [
            link for vector in build_after for link in vector
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline-head", default=baseline_head())
    parser.add_argument("--cases", type=Path, default=PACKET / "cases.json")
    parser.add_argument("--trace-root", type=Path, default=PACKET / "trace")
    arguments = parser.parse_args()
    baseline_folder = arguments.trace_root / "baseline"
    candidate_folder = arguments.trace_root / "candidate"
    audit_manifest("baseline", baseline_folder, arguments.baseline_head)
    audit_manifest("candidate", candidate_folder, arguments.baseline_head)

    cases = {str(case["case"]): case for case in read_json(arguments.cases)}
    rows: list[dict[str, Any]] = []
    for name, case in cases.items():
        before = parse_trace(baseline_folder / f"{name}.trace")
        after = parse_trace(candidate_folder / f"{name}.trace")
        rows.append(compare_case(name, before, after, int(case["budget"])))

    sequence = compare_sequence(baseline_folder, candidate_folder)

    output = arguments.trace_root / "route-comparison.json"
    output.write_text(
        json.dumps({"cases": rows, "sequence": sequence}, indent=2, sort_keys=True)
        + "\n"
    )
    print(f"PASS diagnostic routes: {len(rows)} cases; worksheet checkpoint and SST routes audited")
    print(output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
