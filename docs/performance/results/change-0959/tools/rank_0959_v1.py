#!/usr/bin/env python3
"""Independent offline semantic-owner ranking for the sealed 0958 traces.

This tool never launches a workload.  It reads only the sealed 0958 input
manifest, decoder trace aggregates, trace-size aggregates, and the sealed
0958 attribution result as a comparison oracle.  Selection is recomputed from
raw stack locations: a row is selected only when it has both the exact
``stage_timed_commit`` and ``timed_iteration`` callsites and no decoder unknown
frame marker.  ``MultiSourceEdit::commit`` is reported as a semantic owner
when it is present, but it is never used as a replacement selection boundary.

Semantic labels overlap.  The first meaningful litchi_xlsx owner is a ranking
key; growth, merge, numeric, duplicate-checking, XML, and copy labels remain
separate diagnostics and are never summed as an additive partition.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Any

INPUT_SCHEMA = "litchi.performance.0958.heaptrack-attribution-input.v1"
TRACE_SCHEMA = "litchi.performance.0949.heaptrack-raw-decode.v1"
REFERENCE_ANALYSIS_SHA = "9fb5ae3b1b484e2fc8294c66d4712d8f818c1aed60a974fbe51bcd03173cf983"
INPUT_MANIFEST_SHA = "0dc7a9a4757d73a05d9fc289581a8ec6add494565f039b94e7addf8d18629495"
EXPECTED = {
    "h00-medium-percent-r1": {"shape": "medium", "workload": "one-percent", "repeat": 1,
                               "calls": 330372, "requested_bytes": 51500040},
    "h01-dense-percent-r1": {"shape": "dense-sparse", "workload": "one-percent", "repeat": 1,
                              "calls": 628254, "requested_bytes": 73812219},
    "h02-medium-one-edit-r1": {"shape": "medium", "workload": "one-edit", "repeat": 1,
                                "calls": 81981, "requested_bytes": 11847945},
    "h03-medium-one-edit-r2": {"shape": "medium", "workload": "one-edit", "repeat": 2,
                                "calls": 81981, "requested_bytes": 11847945},
    "h04-dense-percent-r2": {"shape": "dense-sparse", "workload": "one-percent", "repeat": 2,
                              "calls": 628254, "requested_bytes": 73812219},
    "h05-medium-percent-r2": {"shape": "medium", "workload": "one-percent", "repeat": 2,
                               "calls": 330372, "requested_bytes": 51500040},
}
ORDER = tuple(EXPECTED)
MAX_LINE = 16 * 1024 * 1024
MAX_JSON = 64 * 1024 * 1024
HASH_RE = re.compile(r"::h[0-9a-f]{8,}$")

class RankingError(RuntimeError):
    pass

def require(ok: bool, msg: str) -> None:
    if not ok:
        raise RankingError(msg)

def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()

def desc(path: Path) -> dict[str, Any]:
    path = path.expanduser().resolve()
    require(path.is_absolute() and path.is_file() and not path.is_symlink(), f"bad path: {path}")
    stat = path.stat()
    return {"path": str(path), "bytes": stat.st_size, "mode": stat.st_mode & 0o777, "sha256": sha(path)}

def verify_desc(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str), f"{label} descriptor malformed")
    actual = desc(Path(value["path"]))
    for key in ("bytes", "mode", "sha256"):
        if key in value:
            require(actual[key] == value[key], f"{label}.{key} mismatch")
    return actual

def reject_dupes(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        require(key not in result, f"duplicate JSON key {key!r}")
        result[key] = value
    return result

def read_json(path: Path, label: str) -> Any:
    actual = desc(path)
    require(actual["bytes"] <= MAX_JSON, f"{label} exceeds JSON bound")
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_dupes)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise RankingError(f"cannot parse {label}: {exc}") from exc

def integer(value: Any, label: str, minimum: int = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"{label} is not an integer")
    require(value >= minimum, f"{label} below {minimum}")
    return value

def canonical_function(function: str) -> str:
    return HASH_RE.sub("", function)

def location_function(location: dict[str, Any]) -> str:
    value = location.get("function")
    return value if isinstance(value, str) else ""

def location_file(location: dict[str, Any]) -> str:
    value = location.get("file")
    return value if isinstance(value, str) else ""

def location_line(location: dict[str, Any]) -> int | None:
    value = location.get("line")
    return value if isinstance(value, int) and not isinstance(value, bool) else None

def flatten(row: dict[str, Any]) -> list[dict[str, Any]]:
    stack = row.get("stack")
    require(isinstance(stack, list), "trace stack is not a list")
    result: list[dict[str, Any]] = []
    for frame_index, item in enumerate(stack):
        require(isinstance(item, dict), "trace stack item is not an object")
        frame = item.get("frame")
        require(isinstance(frame, dict), "trace frame is not an object")
        locations = frame.get("locations")
        require(isinstance(locations, list), "trace frame locations are not a list")
        for location_index, location in enumerate(locations):
            require(isinstance(location, dict), "trace location is not an object")
            result.append({"frame_index": frame_index, "location_index": location_index, "location": location})
    return result

def is_rust_or_allocator(function: str, file: str) -> bool:
    lower = (function + "\n" + file).lower()
    return (
        "/rustc/" in file or file.startswith("../") or
        function.startswith("alloc::") or function.startswith("core::") or
        function.startswith("std::") or "globalalloc" in lower or
        "___rdl_" in function or "libc.so" in file or
        "__libc_start" in function or function in {"main", "_start"}
    )

def exact_benchmark_callsite(function: str, name: str) -> bool:
    prefix = "litchi_perf_baseline::xlsx_managed_compression::benchmark::" + name + "::"
    return function.startswith(prefix)

def source_owner(flat: list[dict[str, Any]]) -> dict[str, Any]:
    # The first litchi_xlsx frame above allocator/runtime frames is the
    # nearest concrete semantic owner for this allocation trace.
    candidate = None
    for item in flat:
        loc = item["location"]
        function, file = location_function(loc), location_file(loc)
        if is_rust_or_allocator(function, file):
            continue
        if "litchi_xlsx::" in function or "/crates/litchi-xlsx/" in file:
            candidate = item
            break
    if candidate is None:
        for item in flat:
            loc = item["location"]
            function, file = location_function(loc), location_file(loc)
            if is_rust_or_allocator(function, file):
                continue
            if function:
                candidate = item
                break
    if candidate is None:
        return {"key": "<unknown>", "function": "", "file": "", "line": None}
    loc = candidate["location"]
    function, file, line = location_function(loc), location_file(loc), location_line(loc)
    normalized_file = file
    marker = "/crates/litchi-xlsx/"
    if marker in normalized_file:
        normalized_file = "crates/litchi-xlsx/" + normalized_file.split(marker, 1)[1]
    elif "/tools/perf-baseline/" in normalized_file:
        normalized_file = "tools/perf-baseline/" + normalized_file.split("/tools/perf-baseline/", 1)[1]
    key = f"{canonical_function(function)}|{normalized_file}|{line if line is not None else '?'}"
    return {"key": key, "function": canonical_function(function), "file": normalized_file, "line": line}

def tags(flat: list[dict[str, Any]]) -> list[str]:
    funcs = [location_function(x["location"]) for x in flat]
    files = [location_file(x["location"]) for x in flat]
    blob = "\n".join(funcs + files).lower()
    out: list[str] = []
    if "merge_omitted_cells" in blob or "store::merge" in blob:
        out.append("merge_backing_storage")
    if any(("number" in f.lower() and "clone" in f.lower()) or
           ("cell.rs" in file and location_line(x["location"]) in {389, 390})
           for x, f, file in zip(flat, funcs, files)):
        out.append("numeric_clone")
    if any("number::new" in f.lower() for f in funcs):
        out.append("numeric_construction")
    if "check_for_duplicates" in blob:
        out.append("duplicate_checking")
    if any(token in blob for token in ("quick_xml", "into_owned", "bytesstart", "parse_source", "raw/worksheet/codec")):
        out.append("xml_lexical_or_parse")
    if any(token in blob for token in ("copy_boxed", "to_vec", "vec$lt$t", "clone::") for _ in [0]):
        out.append("copy_or_clone_mechanism")
    if any(token in blob for token in ("grow_amortized", "try_reserve", "finish_grow", "globalalloc", "realloc")):
        out.append("allocator_growth_or_realloc")
    return sorted(set(out))

def parse_trace_sizes(path: Path) -> dict[int, dict[int, int]]:
    result: dict[int, dict[int, int]] = defaultdict(dict)
    with path.open(encoding="utf-8") as stream:
        header = stream.readline().rstrip("\n")
        require(header == "trace_id\tsize_bytes\tcalls\trequested_bytes", f"bad trace-size header: {path}")
        for no, line in enumerate(stream, 2):
            require(len(line.encode()) <= MAX_LINE, f"trace-size line too long: {path}:{no}")
            fields = line.rstrip("\n").split("\t")
            require(len(fields) == 4, f"bad trace-size row: {path}:{no}")
            try:
                trace_id = int(fields[0], 16); size, calls, requested = (int(x) for x in fields[1:])
            except ValueError as exc:
                raise RankingError(f"noninteger trace-size row: {path}:{no}") from exc
            integer(trace_id, "trace id"); integer(size, "size"); integer(calls, "calls", 1); integer(requested, "requested bytes")
            require(requested == size * calls, f"trace-size arithmetic mismatch: {path}:{no}")
            require(size not in result[trace_id], f"duplicate trace-size bucket: {path}:{no}")
            result[trace_id][size] = calls
    return result

def raw_trace_row(line: str, path: Path, no: int) -> dict[str, Any]:
    require(len(line.encode()) <= MAX_LINE, f"trace row too long: {path}:{no}")
    try:
        row = json.loads(line, object_pairs_hook=reject_dupes)
    except (UnicodeError, json.JSONDecodeError) as exc:
        raise RankingError(f"bad trace JSON {path}:{no}: {exc}") from exc
    require(isinstance(row, dict), f"trace row not object: {path}:{no}")
    trace_text = row.get("trace_id")
    require(isinstance(trace_text, str), f"trace id missing: {path}:{no}")
    try: trace_id = int(trace_text, 16)
    except ValueError as exc: raise RankingError(f"bad trace id {trace_text}: {path}:{no}") from exc
    calls = integer(row.get("calls"), f"{path}:{no} calls", 1)
    requested = integer(row.get("requested_bytes"), f"{path}:{no} requested bytes")
    return {"row": row, "trace_id": trace_id, "calls": calls, "requested_bytes": requested, "flat": flatten(row)}

def add_total(dst: dict[str, int], calls: int, requested: int, traces: int = 1) -> None:
    dst["calls"] += calls; dst["requested_bytes"] += requested; dst["trace_count"] += traces

def empty_total() -> dict[str, int]:
    return {"calls": 0, "requested_bytes": 0, "trace_count": 0}

def top_item(value: dict[str, Any]) -> dict[str, Any]:
    return {k: value[k] for k in ("trace_id", "calls", "requested_bytes", "owner", "tags", "meaningful_frames")}

def analyze_run(run: dict[str, Any], reference_run: dict[str, Any] | None) -> dict[str, Any]:
    run_id = run["id"]
    trace_desc = verify_desc(run["trace_aggregates"], f"{run_id}.trace_aggregates")
    size_desc = verify_desc(run["trace_size_aggregates"], f"{run_id}.trace_size_aggregates")
    summary_desc = verify_desc(run["decoder_summary"], f"{run_id}.decoder_summary")
    summary = read_json(Path(summary_desc["path"]), f"{run_id}.decoder_summary")
    require(summary.get("schema") == TRACE_SCHEMA and summary.get("status") == "completed", f"{run_id} decoder summary schema/status")
    size_map = parse_trace_sizes(Path(size_desc["path"]))
    seen: set[int] = set(); raw = empty_total(); selected = empty_total(); unknown = empty_total(); unqualified = empty_total()
    owners: dict[str, dict[str, Any]] = {}
    tag_totals = defaultdict(empty_total); tag_owners: dict[str, dict[str, dict[str, int]]] = defaultdict(lambda: defaultdict(lambda: {"calls": 0, "requested_bytes": 0, "trace_count": 0}))
    size_totals = defaultdict(lambda: {"calls": 0, "requested_bytes": 0})
    selected_rows: list[dict[str, Any]] = []
    all_rows = 0
    with Path(trace_desc["path"]).open(encoding="utf-8") as stream:
        for no, line in enumerate(stream, 1):
            parsed = raw_trace_row(line, Path(trace_desc["path"]), no)
            trace_id, calls, requested, flat = parsed["trace_id"], parsed["calls"], parsed["requested_bytes"], parsed["flat"]
            require(trace_id not in seen, f"{run_id} duplicate trace id {trace_id:x}")
            seen.add(trace_id); all_rows += 1; add_total(raw, calls, requested)
            require(trace_id in size_map, f"{run_id} trace-size missing {trace_id:x}")
            size_calls = sum(size_map[trace_id].values()); size_bytes = sum(size * n for size, n in size_map[trace_id].items())
            require(size_calls == calls and size_bytes == requested, f"{run_id} trace-size mismatch {trace_id:x}")
            functions = [location_function(x["location"]) for x in flat]
            has_stage = any(exact_benchmark_callsite(f, "stage_timed_commit") for f in functions)
            has_timed = any(exact_benchmark_callsite(f, "timed_iteration") for f in functions)
            is_unknown = parsed["row"].get("unknown_any_frame") is True or parsed["row"].get("unknown_leaf_frame") is True
            if is_unknown: add_total(unknown, calls, requested)
            if has_stage and has_timed and not is_unknown:
                add_total(selected, calls, requested)
                for size, count in size_map[trace_id].items():
                    size_totals[size]["calls"] += count; size_totals[size]["requested_bytes"] += size * count
                owner = source_owner(flat); row_tags = tags(flat)
                owner_entry = owners.setdefault(owner["key"], {"owner": owner, "calls": 0, "requested_bytes": 0, "trace_count": 0, "runs": set()})
                owner_entry["calls"] += calls; owner_entry["requested_bytes"] += requested; owner_entry["trace_count"] += 1; owner_entry["runs"].add(run_id)
                meaningful=[]
                for item in flat:
                    loc=item["location"]; f=location_function(loc); file=location_file(loc)
                    if is_rust_or_allocator(f,file): continue
                    if f or file: meaningful.append({"function":canonical_function(f),"file":file,"line":location_line(loc)})
                selected_rows.append({"trace_id":f"{trace_id:x}","calls":calls,"requested_bytes":requested,"owner":owner["key"],"tags":row_tags,"meaningful_frames":meaningful[:8],"run":run_id,"sizes":dict(size_map[trace_id])})
                for tag in row_tags:
                    add_total(tag_totals[tag], calls, requested)
                    entry=tag_owners[tag][owner["key"]]; entry["calls"] += calls; entry["requested_bytes"] += requested; entry["trace_count"] += 1
            else:
                add_total(unqualified, calls, requested)
    require(seen == set(size_map), f"{run_id} trace-size has unreferenced traces")
    require(raw["calls"] == integer(summary.get("allocation_events"), f"{run_id} allocation_events"), f"{run_id} raw calls mismatch decoder")
    require(raw["requested_bytes"] == integer(summary.get("requested_bytes_total"), f"{run_id} requested_bytes_total"), f"{run_id} raw bytes mismatch decoder")
    expected=EXPECTED[run_id]
    require(selected["calls"] == expected["calls"] and selected["requested_bytes"] == expected["requested_bytes"], f"{run_id} exact stage/timed selection mismatch: {selected} != {expected}")
    if reference_run is not None:
        ref=reference_run.get("owner_selection",{}).get("selected_owner",{})
        require(ref.get("calls") == selected["calls"] and ref.get("requested_bytes") == selected["requested_bytes"], f"{run_id} differs from 0958 selected totals")
    owner_list=[]
    for value in owners.values():
        value=dict(value); value["runs"]=sorted(value["runs"]); owner_list.append(value)
    owner_list.sort(key=lambda x:(-x["requested_bytes"],-x["calls"],x["owner"]["key"]))
    tag_owner_out={}
    for tag, owners_for_tag in tag_owners.items():
        rows=[]
        for owner_key,value in owners_for_tag.items(): rows.append({"owner":owner_key,**value})
        rows.sort(key=lambda x:(-x["requested_bytes"],-x["calls"],x["owner"]))
        tag_owner_out[tag]=rows[:20]
    top_traces=sorted(selected_rows,key=lambda x:(-x["requested_bytes"],-x["calls"],x["trace_id"]))[:20]
    top_sizes=[{"size_bytes":size,"calls":value["calls"],"requested_bytes":value["requested_bytes"]} for size,value in sorted(size_totals.items(),key=lambda x:(-x[1]["requested_bytes"],-x[1]["calls"],x[0]))[:30]]
    selected_ids={x["trace_id"] for x in selected_rows}
    all_tag_totals={}
    # Whole-process tag counts are intentionally a separate scan over the same
    # parsed trace rows only when needed; selected tag counts are the primary scope.
    return {"id":run_id,"shape":run["shape"],"workload":run["workload"],"repeat":run["repeat"],"raw_whole_process":raw,"selected_exact_stage_timed":selected,"unknown_any_or_leaf":unknown,"unqualified":unqualified,"trace_rows":all_rows,"top_allocation_sizes":top_sizes,"top_selected_traces":[top_item(x) for x in top_traces],"owners_ranked":owner_list[:50],"tag_totals":dict(sorted(tag_totals.items())),"tag_owner_rankings":tag_owner_out,"selected_tag_scope":"overlapping labels over exact stage/timed selected rows; never additive","conservation":"pass"}

def main() -> int:
    parser=argparse.ArgumentParser()
    parser.add_argument("--manifest",required=True,type=Path)
    parser.add_argument("--reference-analysis",required=True,type=Path)
    parser.add_argument("--output",required=True,type=Path)
    args=parser.parse_args()
    manifest_path=args.manifest.expanduser().resolve()
    actual_manifest=desc(manifest_path)
    require(actual_manifest["sha256"] == INPUT_MANIFEST_SHA, "unexpected sealed input manifest SHA")
    manifest=read_json(manifest_path,"0958 input manifest")
    require(manifest.get("schema")==INPUT_SCHEMA and manifest.get("status")=="sealed", "input manifest schema/status")
    require(tuple(run.get("id") for run in manifest.get("runs",[])) == ORDER, "run order differs from sealed 0958 order")
    reference_path=args.reference_analysis.expanduser().resolve(); reference_desc=desc(reference_path)
    require(reference_desc["sha256"]==REFERENCE_ANALYSIS_SHA,"unexpected 0958 reference analysis SHA")
    reference=read_json(reference_path,"0958 reference analysis")
    ref_runs={r["id"]:r for r in reference.get("runs",[])}
    results=[]
    for run in manifest["runs"]:
        run_id=run["id"]; require(run_id in EXPECTED,"unexpected run id")
        for k in ("trace_aggregates","trace_size_aggregates","decoder_summary"):
            verify_desc(run[k],f"{run_id}.{k}")
        results.append(analyze_run(run,ref_runs.get(run_id)))
    aggregate_raw=empty_total(); aggregate_selected=empty_total(); aggregate_unknown=empty_total(); aggregate_unqualified=empty_total()
    for r in results:
        for dst,src in ((aggregate_raw,r["raw_whole_process"]),(aggregate_selected,r["selected_exact_stage_timed"]),(aggregate_unknown,r["unknown_any_or_leaf"]),(aggregate_unqualified,r["unqualified"])): add_total(dst,src["calls"],src["requested_bytes"],src["trace_count"])
    output={
      "schema":"litchi.performance.0959.semantic-owner-ranking.v1","status":"pass",
      "scope":"Independent raw-trace ranking for exact stage_timed_commit + timed_iteration rows from sealed 0958 profiles; no workload or build launched.",
      "input_manifest":actual_manifest,"reference_analysis":reference_desc,
      "selection":{"exact_callsites":["stage_timed_commit","timed_iteration"],"unknown_policy":"unknown_any_frame/unknown_leaf_frame rows are unqualified","surrogate_policy":"MultiSourceEdit::commit is a reported semantic owner/tag when present but is not a selection surrogate"},
      "aggregate":{"raw_whole_process":aggregate_raw,"selected_exact_stage_timed":aggregate_selected,"unknown_any_or_leaf":aggregate_unknown,"unqualified":aggregate_unqualified},
      "runs":results,
      "interpretation":[
        "owners_ranked is the first meaningful litchi_xlsx frame above allocator/runtime frames; this is a nearest-owner ranking, not an inclusive call-tree sum.",
        "merge_backing_storage, numeric_clone, numeric_construction, duplicate_checking, xml_lexical_or_parse, copy_or_clone_mechanism, and allocator_growth_or_realloc labels overlap and are retained separately.",
        "Requested bytes and calls are raw allocation totals in the selected whole-process stack scope; they do not establish a timed phase total, speedup, or adoption recommendation.",
        "The 0958 analyzer's commit-core surrogate bucket is not used for this independent selection; exact stage/timed selection is cross-checked against its selected totals."
      ]
    }
    require(not args.output.exists() and not args.output.is_symlink(), f"refusing overwrite: {args.output}")
    args.output.parent.mkdir(parents=True,exist_ok=True)
    data=(json.dumps(output,indent=2,sort_keys=True,ensure_ascii=False,allow_nan=False)+"\n").encode()
    with args.output.open("xb") as stream: stream.write(data); stream.flush(); os.fsync(stream.fileno())
    print(json.dumps({"output":str(args.output),"bytes":args.output.stat().st_size,"sha256":sha(args.output),"status":"pass"},sort_keys=True))
    return 0

if __name__ == "__main__":
    try: raise SystemExit(main())
    except RankingError as exc:
        print(json.dumps({"status":"fail","error":str(exc)},sort_keys=True),file=sys.stderr)
        raise SystemExit(1)
