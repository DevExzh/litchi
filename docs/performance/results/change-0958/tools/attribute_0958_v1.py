#!/usr/bin/env python3
"""Offline 0958 heaptrack reconciliation and timed-owner attribution.

The program consumes a sealed manifest and reviewed 0949 decoder outputs.  It
never launches a workload or profiler.  Allocation totals are whole-process;
only decoded rows with resolved ``stage_timed_commit`` and ``timed_iteration``
callsites enter the measured-owner partition. Unknown and unqualified events
remain explicit. Source tags overlap; exclusive leaves are the additive
partition.
"""
from __future__ import annotations
import argparse, hashlib, json, math, os, sys
from collections import Counter, defaultdict
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterator

SCHEMA = "litchi.performance.0958.heaptrack-attribution.v1"
INPUT_SCHEMA = "litchi.performance.0958.heaptrack-attribution-input.v1"
DECODER_SCHEMA = "litchi.performance.0949.heaptrack-raw-decode.v1"
REPORT_SCHEMA = "litchi.xlsx-managed-compression-benchmark.v1"
REPORT_CLAIM = ("none; raw phase timings describe this explicit in-memory harness and do not "
                "claim OS thread counts, speedup, allocation behavior, or physical I/O")
EXPECTED_RUNS = (
    ("h00-medium-percent-r1", "medium", "one-percent", 1),
    ("h01-dense-percent-r1", "dense-sparse", "one-percent", 1),
    ("h02-medium-one-edit-r1", "medium", "one-edit", 1),
    ("h03-medium-one-edit-r2", "medium", "one-edit", 2),
    ("h04-dense-percent-r2", "dense-sparse", "one-percent", 2),
    ("h05-medium-percent-r2", "medium", "one-percent", 2),
)
MAX_JSON = 64 * 1024 * 1024
MAX_LINE = 16 * 1024 * 1024

class AttributionError(RuntimeError):
    pass

def require(ok: bool, message: str) -> None:
    if not ok:
        raise AttributionError(message)

def now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")

def sha(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()

def desc(path: Path) -> dict[str, Any]:
    path = path.expanduser().resolve()
    require(path.is_absolute() and path.is_file() and not path.is_symlink(), f"bad file: {path}")
    st = path.stat()
    return {"path": str(path), "bytes": st.st_size, "mode": st.st_mode & 0o777, "sha256": sha(path)}

def verify_desc(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict) and isinstance(value.get("path"), str), f"{label} descriptor malformed")
    p = Path(value["path"])
    require(p.is_absolute(), f"{label} path is not absolute")
    actual = desc(p)
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

def read_json(path: Path, label: str, limit: int = MAX_JSON) -> Any:
    actual = desc(path)
    require(actual["bytes"] <= limit, f"{label} exceeds {limit} bytes")
    try:
        return json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=reject_dupes)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise AttributionError(f"cannot parse {label}: {exc}") from exc

def write_once(path: Path, value: Any) -> None:
    require(not path.exists() and not path.is_symlink(), f"refusing overwrite: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    data = (json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False, allow_nan=False) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data); stream.flush(); os.fsync(stream.fileno())
    fd = os.open(path.parent, os.O_RDONLY | os.O_DIRECTORY)
    try: os.fsync(fd)
    finally: os.close(fd)

def integer(value: Any, label: str, minimum: int = 0) -> int:
    require(isinstance(value, int) and not isinstance(value, bool), f"{label} is not integer")
    require(value >= minimum, f"{label} below {minimum}")
    return value

def total(calls: int = 0, bytes_: int = 0, traces: int = 0) -> dict[str, int]:
    return {"calls": calls, "requested_bytes": bytes_, "trace_count": traces}

def add(target: dict[str, int], calls: int, bytes_: int, traces: int = 1) -> None:
    target["calls"] += calls; target["requested_bytes"] += bytes_; target["trace_count"] += traces

def same_total(a: dict[str, int], b: dict[str, int], label: str) -> None:
    require(a["calls"] == b["calls"] and a["requested_bytes"] == b["requested_bytes"],
            f"{label} totals differ: {a} != {b}")

def report_contract(run: dict[str, Any], path: Path) -> dict[str, Any]:
    raw = read_json(path, f"report {run['id']}")
    require(raw.get("schema") == REPORT_SCHEMA and raw.get("status") == "completed", f"{run['id']} report status/schema")
    for key, expected in (("shape", run["shape"]), ("workload", run["workload"]),
                          ("workers", 1), ("warmup_iterations", 0), ("measured_samples", 3),
                          ("performance_claim", REPORT_CLAIM)):
        require(raw.get(key) == expected, f"{run['id']} report {key} mismatch")
    case = raw.get("case"); require(isinstance(case, dict), f"{run['id']} report case missing")
    require(case.get("preflight_actual_commit_inverse_verified") is True, f"{run['id']} commit oracle failed")
    samples = case.get("samples"); require(isinstance(samples, list) and len(samples) == 3, f"{run['id']} sample count")
    fields = ("ingress_ns", "edit_commit_ns", "publication_ns", "commit_drop_ns", "total_ns")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"{run['id']} sample {index} malformed")
        require(sample.get("fixture_inverse_restored_exact_source") is True, f"{run['id']} sample {index} inverse failed")
        for field in fields: integer(sample.get(field), f"{run['id']} sample {index} {field}")
    return {"schema": raw["schema"], "status": raw["status"], "shape": raw["shape"],
            "workload": raw["workload"], "workers": raw["workers"], "warmup_iterations": raw["warmup_iterations"],
            "measured_samples": raw["measured_samples"], "source_sha256": case.get("source_sha256"),
            "output_sha256": case.get("output_sha256"), "semantic_sha256": case.get("semantic_sha256"),
            "timing_scope": raw.get("timing_scope"), "phase_scope": raw.get("phase_scope"),
            "performance_claim": raw["performance_claim"],
            "sample_phase_totals": {field: sum(sample[field] for sample in samples) for field in fields}}

def parse_size(path: Path) -> tuple[dict[int, tuple[int, int]], dict[str, int]]:
    rows: dict[int, tuple[int, int]] = {}
    with path.open(encoding="utf-8") as stream:
        require(stream.readline().rstrip("\n") == "size_bytes\tcalls\trequested_bytes", f"bad size header: {path}")
        for no, line in enumerate(stream, 2):
            require(len(line.encode()) <= MAX_LINE, f"size line too long: {path}:{no}")
            fields = line.rstrip("\n").split("\t"); require(len(fields) == 3, f"bad size row: {path}:{no}")
            try: size, calls, bytes_ = (int(x) for x in fields)
            except ValueError as exc: raise AttributionError(f"noninteger size row: {path}:{no}") from exc
            integer(size, "decoder size"); integer(calls, "decoder size calls", 1); integer(bytes_, "decoder size bytes")
            require(bytes_ == size * calls and size not in rows, f"size arithmetic/duplicate: {path}:{no}")
            rows[size] = (calls, bytes_)
    return rows, {"calls": sum(x[0] for x in rows.values()), "requested_bytes": sum(x[1] for x in rows.values())}

def parse_trace_sizes(path: Path) -> dict[int, dict[int, int]]:
    rows: dict[int, dict[int, int]] = defaultdict(dict)
    with path.open(encoding="utf-8") as stream:
        require(stream.readline().rstrip("\n") == "trace_id\tsize_bytes\tcalls\trequested_bytes", f"bad trace-size header: {path}")
        for no, line in enumerate(stream, 2):
            require(len(line.encode()) <= MAX_LINE, f"trace-size line too long: {path}:{no}")
            fields = line.rstrip("\n").split("\t"); require(len(fields) == 4, f"bad trace-size row: {path}:{no}")
            try:
                trace = int(fields[0], 16)
                size, calls, bytes_ = (int(x) for x in fields[1:])
            except ValueError as exc: raise AttributionError(f"noninteger trace-size row: {path}:{no}") from exc
            integer(trace, "trace id"); integer(size, "trace size"); integer(calls, "trace calls", 1); integer(bytes_, "trace bytes")
            require(bytes_ == size * calls and size not in rows[trace], f"trace-size arithmetic/duplicate: {path}:{no}")
            rows[trace][size] = calls
    return rows

def parse_histogram(path: Path) -> tuple[dict[int, int], dict[str, int]]:
    rows: dict[int, int] = {}
    with path.open(encoding="utf-8") as stream:
        for no, line in enumerate(stream, 1):
            require(len(line.encode()) <= MAX_LINE, f"histogram line too long: {path}:{no}")
            fields = line.strip().split()
            if not fields or fields[0].startswith("#"): continue
            require(len(fields) in (2, 3), f"bad histogram row: {path}:{no}")
            try: size, calls = int(fields[0]), int(fields[1])
            except ValueError as exc: raise AttributionError(f"noninteger histogram row: {path}:{no}") from exc
            integer(size, "heaptrack size"); integer(calls, "heaptrack calls"); require(size not in rows, f"duplicate histogram size: {path}:{no}")
            if len(fields) == 3:
                try: bytes_ = int(fields[2])
                except ValueError as exc: raise AttributionError(f"noninteger histogram bytes: {path}:{no}") from exc
                require(bytes_ == size * calls, f"histogram arithmetic: {path}:{no}")
            rows[size] = calls
    return rows, {"calls": sum(rows.values()), "requested_bytes": sum(size * calls for size, calls in rows.items())}

def parse_folded(path: Path) -> dict[str, Any]:
    calls = 0; bytes_: int | None = 0; rows = 0; first = True
    with path.open(encoding="utf-8") as stream:
        for no, line in enumerate(stream, 1):
            require(len(line.encode()) <= MAX_LINE, f"folded line too long: {path}:{no}")
            text = line.rstrip("\n")
            if not text: continue
            fields = text.split("\t")
            if first and fields[:3] == ["stack", "calls", "requested_bytes"]:
                first = False; continue
            first = False
            if "\t" in text:
                require(len(fields) in (2, 3) and fields[0], f"bad folded row: {path}:{no}")
                stack_text, call_text = fields[0], fields[1]
                byte_text = fields[2] if len(fields) == 3 else None
            else:
                try: stack_text, call_text = text.rsplit(None, 1)
                except ValueError as exc: raise AttributionError(f"bad folded row: {path}:{no}") from exc
                byte_text = None
            require(stack_text, f"empty folded stack: {path}:{no}")
            try: row_calls = int(call_text); row_bytes = int(byte_text) if byte_text is not None else None
            except ValueError as exc: raise AttributionError(f"noninteger folded row: {path}:{no}") from exc
            integer(row_calls, "folded calls", 1)
            if row_bytes is None: bytes_ = None
            else:
                integer(row_bytes, "folded bytes")
                require(bytes_ is not None, f"mixed folded byte columns: {path}:{no}")
                bytes_ += row_bytes
            calls += row_calls; rows += 1
    require(rows > 0, f"folded output empty: {path}")
    return {"rows": rows, "calls": calls, "requested_bytes": bytes_, "bytes_present": bytes_ is not None}

def default_rules() -> dict[str, Any]:
    return {
        "number": {"function": ("Number",), "lines": ((388, 397),)},
        "stored": {"function": ("Stored",), "lines": ((646, 676),)},
        "merge": {"function": ("merge_omitted_cells",), "lines": ((893, 959),)},
        "clone": {"function": ("litchi_xlsx..cell..Number$", "litchi_xlsx..cell..Stored$", "litchi_xlsx..cell..Value$", "litchi_xlsx..cell..Cell$"), "lines": ((936, 936),)},
        "copy": {"function": ("copy", "to_owned", "to_vec"), "lines": ()},
        "realloc": {"function": ("realloc", "grow_amortized"), "lines": ()},
    }

def load_rules(manifest: dict[str, Any]) -> dict[str, Any]:
    rules = default_rules()
    raw = manifest.get("source_classification")
    if raw is None:
        return rules
    require(isinstance(raw, dict), "source_classification must be object")
    for name, default in rules.items():
        value = raw.get(name)
        if value is None:
            continue
        require(isinstance(value, dict), f"source rule {name} malformed")
        funcs = value.get("function_substrings", list(default["function"]))
        lines = value.get("line_ranges", [list(x) for x in default["lines"]])
        require(isinstance(funcs, list) and all(isinstance(x, str) and x for x in funcs),
                f"source rule {name} functions malformed")
        clean: list[tuple[int, int]] = []
        require(isinstance(lines, list), f"source rule {name} ranges malformed")
        for pair in lines:
            require(isinstance(pair, list) and len(pair) == 2 and
                    all(isinstance(x, int) and not isinstance(x, bool) for x in pair) and
                    0 <= pair[0] <= pair[1], f"source rule {name} range malformed")
            clean.append((pair[0], pair[1]))
        rules[name] = {"function": tuple(funcs), "lines": tuple(clean)}
    return rules

def fn(location: dict[str, Any]) -> str:
    value = location.get("function")
    return value if isinstance(value, str) else ""

def file_name(location: dict[str, Any]) -> str:
    value = location.get("file")
    return value if isinstance(value, str) else ""

def line_no(location: dict[str, Any]) -> int | None:
    value = location.get("line")
    return value if isinstance(value, int) and not isinstance(value, bool) else None

def flat_stack(stack: Any) -> list[dict[str, Any]]:
    require(isinstance(stack, list), "trace stack is not a list")
    flat: list[dict[str, Any]] = []
    for frame_index, item in enumerate(stack):
        if not isinstance(item, dict):
            continue
        frame = item.get("frame")
        if not isinstance(frame, dict) or not isinstance(frame.get("locations"), list):
            continue
        for location_index, location in enumerate(frame["locations"]):
            if isinstance(location, dict):
                flat.append({"frame_index": frame_index, "location_index": location_index,
                             "location": location})
    return flat

def resolved_owner_index(flat: list[dict[str, Any]], tokens: tuple[str, ...]) -> tuple[int | None, str | None]:
    for index, item in enumerate(flat):
        location = item["location"]
        function = fn(location)
        if (file_name(location) or line_no(location) is not None):
            for token in tokens:
                if token in function:
                    return index, token
    return None, None

def callsite(item: dict[str, Any]) -> str:
    location = item["location"]
    line = line_no(location)
    return f"{fn(location) or '<unknown>'}|{file_name(location) or '<unknown>'}|{line if line is not None else '?'}"

def location_tags(location: dict[str, Any], rules: dict[str, Any]) -> set[str]:
    function, file, line = fn(location), file_name(location), line_no(location)
    result: set[str] = set()
    for name, rule in rules.items():
        function_hit = any(token in function for token in rule["function"])
        line_hit = bool(file.endswith("/cell.rs") and line is not None and
                        any(start <= line <= end for start, end in rule["lines"]))
        if function_hit or line_hit:
            result.add(name)
    return result

def classify(row: dict[str, Any], rules: dict[str, Any]) -> dict[str, Any]:
    calls = integer(row.get("calls"), "trace calls", 1)
    bytes_ = integer(row.get("requested_bytes"), "trace bytes")
    flat = flat_stack(row.get("stack"))
    # stage_timed_commit is often inlined in the uninstrumented native binary.
    # The reviewed MultiSourceEdit::commit symbol is the bounded surrogate for
    # that stage; its use is recorded explicitly and never renamed as the
    # source function.
    stage, stage_token = resolved_owner_index(flat, ("stage_timed_commit", "MultiSourceEdit::commit"))
    timed, timed_token = resolved_owner_index(flat, ("timed_iteration",))
    reasons: list[str] = []
    if stage is None: reasons.append("missing_stage_timed_commit")
    if timed is None: reasons.append("missing_timed_iteration")
    if row.get("unknown_any_frame") is True or row.get("unknown_leaf_frame") is True:
        reasons.append("decoder_unknown_frame")
    selected = not reasons
    tags: set[str] = set()
    by_tag: dict[str, list[dict[str, Any]]] = defaultdict(list)
    exclusive = None
    selected_callsite = None
    if selected:
        # The stack is leaf-to-root.  Only frames at or below the resolved
        # timing boundary are eligible source callsites.
        eligible = flat[:max(stage, timed) + 1]  # type: ignore[arg-type]
        for item in eligible:
            for tag in location_tags(item["location"], rules):
                tags.add(tag); by_tag[tag].append(item)
        for tag in ("clone", "number", "stored", "merge", "copy", "realloc"):
            if by_tag.get(tag):
                selected_callsite = by_tag[tag][0]; exclusive = tag; break
        if selected_callsite is None and eligible:
            selected_callsite = eligible[0]
            exclusive = "selected_owner_other"
    return {
        "calls": calls, "requested_bytes": bytes_, "selected": selected,
        "reasons": reasons, "tags": sorted(tags),
        "exclusive": exclusive, "callsite": callsite(selected_callsite) if selected_callsite else None,
        "stage_index": stage, "timed_index": timed, "stage_resolution": stage_token, "timed_resolution": timed_token,
    }

def process_run(run: dict[str, Any], rules: dict[str, Any]) -> dict[str, Any]:
    run_id = run["id"]
    report_desc = verify_desc(run["report"], f"{run_id}.report")
    report = report_contract(run, Path(report_desc["path"]))
    profile = verify_desc(run["profile"], f"{run_id}.profile")
    summary_desc = verify_desc(run["decoder_summary"], f"{run_id}.decoder_summary")
    size_desc = verify_desc(run["size_aggregates"], f"{run_id}.size_aggregates")
    trace_size_desc = verify_desc(run["trace_size_aggregates"], f"{run_id}.trace_size_aggregates")
    trace_desc = verify_desc(run["trace_aggregates"], f"{run_id}.trace_aggregates")
    histogram_desc = verify_desc(run["heaptrack_histogram"], f"{run_id}.heaptrack_histogram")
    folded_desc = verify_desc(run["heaptrack_folded"], f"{run_id}.heaptrack_folded")
    summary = read_json(Path(summary_desc["path"]), f"{run_id} decoder summary")
    require(summary.get("schema") == DECODER_SCHEMA and summary.get("status") == "completed",
            f"{run_id} decoder summary status/schema")
    require(summary.get("label") in (None, run_id), f"{run_id} decoder label mismatch")
    for field in ("plus_without_metadata", "unresolved_trace_definitions", "a_plus_adjacency_mismatches"):
        require(integer(summary.get(field), f"{run_id} {field}") == 0, f"{run_id} decoder {field} nonzero")
    require(summary.get("string_length_errors", []) == [], f"{run_id} decoder string lengths failed")
    profile_summary = summary.get("profile")
    if isinstance(profile_summary, dict):
        require(profile_summary.get("sha256") in (None, profile["sha256"]), f"{run_id} decoder profile hash differs")
        require(profile_summary.get("bytes") in (None, profile["bytes"]), f"{run_id} decoder profile size differs")
    sizes, size_total = parse_size(Path(size_desc["path"]))
    trace_sizes = parse_trace_sizes(Path(trace_size_desc["path"]))
    histogram, hist_total = parse_histogram(Path(histogram_desc["path"]))
    folded = parse_folded(Path(folded_desc["path"]))
    expected = total(integer(summary.get("allocation_events"), f"{run_id} allocation_events"),
                     integer(summary.get("requested_bytes_total"), f"{run_id} requested_bytes_total"))
    same_total(expected, total(size_total["calls"], size_total["requested_bytes"]), f"{run_id} decoder size")
    same_total(expected, total(hist_total["calls"], hist_total["requested_bytes"]), f"{run_id} histogram")
    require(sizes.keys() == histogram.keys(), f"{run_id} histogram buckets differ")
    for size, (calls, _bytes) in sizes.items():
        require(histogram[size] == calls, f"{run_id} histogram calls differ at size {size}")
    require(folded["calls"] == expected["calls"], f"{run_id} folded calls differ")
    if folded["bytes_present"]:
        require(folded["requested_bytes"] == expected["requested_bytes"], f"{run_id} folded bytes differ")
    folded_total = {"calls": folded["calls"], "requested_bytes": folded["requested_bytes"], "trace_count": 0}
    raw = total(); unknown_any = total(); unknown_leaf = total(); selected = total(); unqualified = total()
    exclusive: dict[str, dict[str, int]] = defaultdict(total)
    tag_totals: dict[str, dict[str, int]] = defaultdict(total)
    intersections: dict[str, dict[str, int]] = defaultdict(total)
    stage_resolution_totals: dict[str, dict[str, int]] = defaultdict(total)
    callsites: dict[str, dict[str, int]] = defaultdict(total)
    reasons: Counter[str] = Counter()
    trace_ids: set[int] = set(); trace_count = 0
    with Path(trace_desc["path"]).open(encoding="utf-8") as stream:
        for line_no_value, line in enumerate(stream, 1):
            require(len(line.encode()) <= MAX_LINE, f"{run_id} trace line too long: {line_no_value}")
            try: row = json.loads(line, object_pairs_hook=reject_dupes)
            except (UnicodeError, json.JSONDecodeError) as exc: raise AttributionError(f"{run_id} trace parse: {exc}") from exc
            require(isinstance(row, dict), f"{run_id} trace row not object")
            text_id = row.get("trace_id"); require(isinstance(text_id, str), f"{run_id} trace id missing")
            try: trace_id = int(text_id, 16)
            except ValueError as exc: raise AttributionError(f"{run_id} invalid trace id {text_id}") from exc
            require(trace_id not in trace_ids, f"{run_id} duplicate trace id {text_id}")
            trace_ids.add(trace_id)
            calls = integer(row.get("calls"), f"{run_id} trace calls", 1)
            bytes_ = integer(row.get("requested_bytes"), f"{run_id} trace bytes")
            by_size = trace_sizes.get(trace_id)
            require(by_size is not None, f"{run_id} trace-size missing {text_id}")
            require(sum(by_size.values()) == calls and sum(size * count for size, count in by_size.items()) == bytes_,
                    f"{run_id} trace-size conservation {text_id}")
            add(raw, calls, bytes_); trace_count += 1
            if row.get("unknown_any_frame") is True: add(unknown_any, calls, bytes_)
            if row.get("unknown_leaf_frame") is True: add(unknown_leaf, calls, bytes_)
            outcome = classify(row, rules)
            if outcome["selected"]:
                add(selected, calls, bytes_); add(exclusive[outcome["exclusive"]], calls, bytes_)
                add(stage_resolution_totals[outcome["stage_resolution"]], calls, bytes_)
                if outcome["callsite"]: add(callsites[outcome["callsite"]], calls, bytes_)
                tags_here = set(outcome["tags"])
                for tag in tags_here: add(tag_totals[tag], calls, bytes_)
                for label, pair in (("number_clone", ("number", "clone")),
                                    ("stored_clone", ("stored", "clone")),
                                    ("merge_clone", ("merge", "clone"))):
                    if set(pair) <= tags_here:
                        add(intersections[label], calls, bytes_)
            else:
                add(unqualified, calls, bytes_)
                for reason in outcome["reasons"]: reasons[reason] += 1
    require(trace_ids == set(trace_sizes), f"{run_id} trace-size has unreferenced trace")
    same_total(expected, raw, f"{run_id} trace")
    require(trace_count == integer(summary.get("unique_trace_ids"), f"{run_id} unique_trace_ids"), f"{run_id} trace count differs")
    same_total(expected, total(selected["calls"] + unqualified["calls"], selected["requested_bytes"] + unqualified["requested_bytes"]), f"{run_id} owner split")
    for actual, field in ((unknown_any, "unknown_any_frame"), (unknown_leaf, "unknown_leaf_frame")):
        part = summary.get(field)
        require(isinstance(part, dict), f"{run_id} {field} summary missing")
        require(actual["calls"] == integer(part.get("calls"), f"{run_id} {field}.calls") and
                actual["requested_bytes"] == integer(part.get("requested_bytes"), f"{run_id} {field}.bytes") and
                actual["trace_count"] == integer(part.get("trace_count"), f"{run_id} {field}.trace_count"),
                f"{run_id} {field} mismatch")
    same_total(selected, total(sum(x["calls"] for x in exclusive.values()), sum(x["requested_bytes"] for x in exclusive.values())), f"{run_id} exclusive partition")
    return {
        "id": run_id, "shape": run["shape"], "workload": run["workload"], "repeat": run["repeat"], "status": "pass",
        "report": {"descriptor": report_desc, "identity": report}, "profile": profile,
        "decoder": {"summary": summary_desc, "size_aggregates": size_desc, "trace_size_aggregates": trace_size_desc, "trace_aggregates": trace_desc,
                     "integrity": {field: summary[field] for field in ("allocation_events", "requested_bytes_total", "plus_without_metadata", "unresolved_trace_definitions", "a_plus_adjacency_mismatches")}},
        "heaptrack_print": {"histogram": histogram_desc, "folded": folded_desc, "folded_rows": folded["rows"], "folded_bytes_present": folded["bytes_present"]},
        "conservation": {"raw_trace": expected, "decoder_size": total(size_total["calls"], size_total["requested_bytes"]),
                          "heaptrack_histogram": total(hist_total["calls"], hist_total["requested_bytes"]),
                          "heaptrack_folded": folded_total, "folded_bytes_comparable": folded["bytes_present"], "status": "pass"},
        "owner_selection": {"status": "admitted" if selected["calls"] else "no_selected_owner", "scope": "whole-process rows requiring both resolved timed callsites and no unknown-frame marker",
                            "selected_owner": selected, "stage_resolution_totals": dict(sorted(stage_resolution_totals.items())), "unqualified": unqualified, "unknown_any_frame": unknown_any, "unknown_leaf_frame": unknown_leaf,
                            "reason_trace_counts": dict(sorted(reasons.items()))},
        "exclusive_leaf_partition": dict(sorted(exclusive.items())),
        "source_callsite_top": [{"callsite": key, **value} for key, value in sorted(callsites.items(), key=lambda x: (-x[1]["requested_bytes"], x[0]))[:25]],
        "nonadditive_source_tags": {key: {**value, "interpretation": "overlap diagnostic; not additive"} for key, value in sorted(tag_totals.items())},
        "source_tag_intersections": {key: {**value, "interpretation": "nonadditive intersection; do not sum with source tags"} for key, value in sorted(intersections.items())},
        "source_tag_scope": {"number_clone": "intersection of Number and clone tags; nonadditive", "stored_clone": "intersection of Stored and clone tags; nonadditive",
                             "merge_total": "merge-tagged selected-owner rows; separate from exclusive buckets", "unknown_unqualified": "never source-tagged"},
    }

def validate_manifest(path: Path) -> tuple[dict[str, Any], list[dict[str, Any]], dict[str, Any]]:
    manifest = read_json(path, "0958 attribution manifest")
    require(isinstance(manifest, dict) and manifest.get("schema") == INPUT_SCHEMA, "input schema mismatch")
    require(manifest.get("status") == "sealed", "input manifest is not sealed")
    runs = manifest.get("runs")
    require(isinstance(runs, list) and len(runs) == len(EXPECTED_RUNS), "input must contain six runs")
    seen: set[str] = set(); artifact_paths: dict[str, str] = {}
    for run, expected in zip(runs, EXPECTED_RUNS):
        require(isinstance(run, dict), "run is not an object")
        name, shape, workload, repeat = expected
        require((run.get("id"), run.get("shape"), run.get("workload"), run.get("repeat")) == expected,
                f"schedule mismatch for {name}")
        require(name not in seen, f"duplicate run id {name}"); seen.add(name)
        require(run.get("workers") == 1 and run.get("warmup") == 0 and run.get("samples") == 3,
                f"run contract mismatch for {name}")
        for field in ("profile", "report", "decoder_summary", "size_aggregates", "trace_size_aggregates", "trace_aggregates", "heaptrack_histogram", "heaptrack_folded"):
            require(field in run, f"{name} missing {field}")
            value = run[field]
            require(isinstance(value, dict) and isinstance(value.get("path"), str), f"{name}.{field} malformed")
            resolved = str(Path(value["path"]).expanduser().resolve())
            if resolved in artifact_paths:
                fail(f"artifact reused by {name}.{field} and {artifact_paths[resolved]}")
            artifact_paths[resolved] = f"{name}.{field}"
    # Source ranges are intentionally optional: without them only function-name
    # evidence is used.  If supplied, all referenced source files are still
    # bound by their ordinary immutable descriptor.
    source_descriptor = manifest.get("source_trace")
    if source_descriptor is not None:
        verify_desc(source_descriptor, "source_trace")
    return manifest, runs, {"rules": load_rules(manifest), "artifact_paths": artifact_paths}

def analyze(manifest_path: Path) -> dict[str, Any]:
    manifest, runs, context = validate_manifest(manifest_path)
    outputs = [process_run(run, context["rules"]) for run in runs]
    raw = total(); selected = total()
    for item in outputs:
        add(raw, item["conservation"]["raw_trace"]["calls"], item["conservation"]["raw_trace"]["requested_bytes"])
        add(selected, item["owner_selection"]["selected_owner"]["calls"], item["owner_selection"]["selected_owner"]["requested_bytes"])
    return {
        "schema": SCHEMA, "status": "pass", "tool_version": "attribute_0958_v1", "created_utc": now(),
        "manifest": desc(manifest_path),
        "scope": {"whole_process": True, "phase_allocation_inference": False,
                  "owner_requirements": ["timed_iteration", "stage_timed_commit_or_reviewed_MultiSourceEdit_commit_owner"],
                  "stage_resolution_scope": {"stage_timed_commit": "full stage callsite when symbolized", "MultiSourceEdit::commit": "commit-core subset only; does not imply complete edit_commit phase"},
                  "unknown_unqualified_retained": True,
                  "exclusive_leaf_partition_additive": True,
                  "source_tags_additive": False},
        "runs": outputs,
        "aggregate": {"raw_whole_process": raw, "selected_timed_owner": selected,
                       "owner_admission": "at least one selected row" if selected["calls"] else "no selected rows"},
    }

def self_test() -> dict[str, Any]:
    rules = default_rules()
    def item(function: str, file: str = "/src/tools/perf-baseline/src/xlsx_managed_compression_benchmark.rs", line: int = 1) -> dict[str, Any]:
        return {"frame": {"locations": [{"function": function, "file": file, "line": line}]}, "trace_id": 1}
    row = {"calls": 2, "requested_bytes": 96,
           "stack": [item("litchi_xlsx::cell::Store::merge_omitted_cells", "/src/crates/litchi-xlsx/src/cell.rs", 936), item("stage_timed_commit"), item("timed_iteration")],
           "unknown_any_frame": False, "unknown_leaf_frame": False}
    admitted = classify(row, rules)
    require(admitted["selected"] and admitted["exclusive"] == "clone" and "merge" in admitted["tags"], "positive owner control failed")
    unknown = dict(row); unknown["unknown_any_frame"] = True
    require(not classify(unknown, rules)["selected"], "unknown control admitted")
    missing = dict(row); missing["stack"] = [item("stage_timed_commit")]
    require(not classify(missing, rules)["selected"], "missing-owner control admitted")
    return {"schema": SCHEMA + ".self-test.v1", "status": "pass", "checks": ["owner-pair", "unknown-retained", "exclusive-clone-priority"]}

def failure(manifest: Path, error: Exception) -> dict[str, Any]:
    return {"schema": SCHEMA, "status": "fail", "tool_version": "attribute_0958_v1", "created_utc": now(),
            "manifest_path": str(manifest.expanduser().resolve()), "error_type": type(error).__name__,
            "error": str(error)[:4000], "interpretation": "No attribution result is admitted."}

def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", type=Path); parser.add_argument("--output", type=Path); parser.add_argument("--self-test", action="store_true")
    args = parser.parse_args(argv)
    if args.self_test:
        print(json.dumps(self_test(), sort_keys=True)); return 0
    if args.manifest is None or args.output is None:
        parser.error("--manifest and --output are required unless --self-test is used")
    output = args.output.expanduser().resolve(); require(not output.exists() and not output.is_symlink(), f"refusing existing output: {output}")
    try:
        result = analyze(args.manifest.expanduser().resolve()); write_once(output, result)
        print(json.dumps({"status": result["status"], "output": str(output), "sha256": sha(output)}, sort_keys=True)); return 0
    except (AttributionError, OSError, ValueError) as error:
        result = failure(args.manifest, error); write_once(output, result)
        print(json.dumps({"status": "fail", "output": str(output), "sha256": sha(output), "error": str(error)[:4000]}, sort_keys=True), file=sys.stderr)
        return 1

if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
