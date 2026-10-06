#!/usr/bin/env python3
"""Read-only Heaptrack stack attribution for retained 0923 profiles.

This tool never starts the profiled binary. It verifies the 0923 independent
memory receipt, decodes the retained interpreted profiles, and compares only
owned_reader against owned_vec within each binary variant. Requested bytes are
assigned once per '+' event to the complete leaf trace that owns the event.
"""
from __future__ import annotations

import argparse
import collections
import hashlib
import json
import os
import pathlib
import re
import subprocess
import sys
from typing import Any, Iterable

EXPECTED_FORMAT = "10500 3"
READ_LIMITED = "litchi_opc::phys_pkg::read_limited"

class ParseError(RuntimeError):
    pass


def require(ok: bool, msg: str) -> None:
    if not ok:
        raise ParseError(msg)


def digest_file(path: pathlib.Path) -> dict[str, Any]:
    require(path.is_file(), f"missing file: {path}")
    h = hashlib.sha256()
    size = 0
    with path.open("rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            size += len(chunk)
            h.update(chunk)
    return {"path": str(path.resolve()), "bytes": size, "sha256": h.hexdigest()}


def recorded_descriptor(value: Any, label: str) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: descriptor is not an object")
    for key in ("path", "bytes", "sha256"):
        require(key in value, f"{label}: descriptor lacks {key}")
    return {"path": str(value["path"]), "bytes": int(value["bytes"]), "sha256": str(value["sha256"])}


def check_recorded_file(path: pathlib.Path, expected: dict[str, Any], label: str) -> dict[str, Any]:
    actual = digest_file(path)
    require(actual["bytes"] == expected["bytes"], f"{label}: byte size mismatch")
    require(actual["sha256"] == expected["sha256"], f"{label}: SHA-256 mismatch")
    return actual


def zstd_lines(path: pathlib.Path) -> Iterable[str]:
    proc = subprocess.Popen(
        ["zstdcat", str(path)], stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        text=True, encoding="utf-8",
    )
    try:
        require(proc.stdout is not None, "zstdcat stdout unavailable")
        assert proc.stdout is not None
        for line in proc.stdout:
            yield line.rstrip("\n")
        err = proc.stderr.read() if proc.stderr is not None else ""
        code = proc.wait()
        require(code == 0, f"zstdcat failed ({code}): {err.strip()}")
    finally:
        if proc.poll() is None:
            proc.kill()
        proc.wait()
        if proc.stdout is not None:
            proc.stdout.close()
        if proc.stderr is not None:
            proc.stderr.close()


def nonnegative_hex(value: str, *, line: int, field: str) -> int:
    try:
        result = int(value, 16)
    except ValueError as exc:
        raise ParseError(f"line {line}: invalid {field} {value!r}") from exc
    require(result >= 0, f"line {line}: negative {field}")
    return result


def parse_profile(path: pathlib.Path, expected: dict[str, Any]) -> dict[str, Any]:
    strings: list[str] = [""]
    instructions: list[tuple[str, ...]] = [()]
    # A trace stores its current instruction first, then its parent (leaf->root).
    traces: list[tuple[str, ...]] = [()]
    allocations: list[tuple[int, int]] = []
    active: list[int] = []
    stack_bytes: collections.Counter[tuple[str, ...]] = collections.Counter()
    stack_calls: collections.Counter[tuple[str, ...]] = collections.Counter()
    command: str | None = None
    version: str | None = None
    version_count = 0
    terminal: set[str] = set()
    data_after_terminal = False
    allocation_events = 0
    deallocation_events = 0
    requested_bytes = 0
    live_bytes = 0
    peak_live = 0

    for line_number, line in enumerate(zstd_lines(path), 1):
        tag, separator, rest = line.partition(" ")
        if terminal == {"strings", "ips"} and tag not in {"", "#"}:
            data_after_terminal = True
            raise ParseError(f"line {line_number}: record follows terminal summaries")
        if tag == "v":
            version_count += 1
            require(version_count == 1, f"line {line_number}: repeated version")
            version = rest
        elif tag == "X":
            require(command is None, f"line {line_number}: repeated command")
            command = rest
        elif tag == "s":
            fields = rest.split(" ", 1)
            require(len(fields) == 2, f"line {line_number}: malformed string")
            length = nonnegative_hex(fields[0], line=line_number, field="string length")
            value = fields[1]
            require(len(value.encode("utf-8")) == length, f"line {line_number}: string length mismatch")
            strings.append(value)
        elif tag == "i":
            fields = rest.split()
            require(len(fields) >= 2, f"line {line_number}: malformed instruction")
            names: list[str] = []
            for index in range(2, len(fields), 3):
                string_index = nonnegative_hex(fields[index], line=line_number, field="instruction string ref")
                require(string_index < len(strings), f"line {line_number}: instruction string out of range")
                names.append(strings[string_index])
            instructions.append(tuple(names))
        elif tag == "t":
            fields = rest.split()
            require(len(fields) == 2, f"line {line_number}: malformed trace")
            instruction_index = nonnegative_hex(fields[0], line=line_number, field="trace instruction ref")
            parent_index = nonnegative_hex(fields[1], line=line_number, field="trace parent ref")
            require(instruction_index < len(instructions), f"line {line_number}: trace instruction out of range")
            require(parent_index < len(traces), f"line {line_number}: trace parent out of range")
            traces.append(instructions[instruction_index] + traces[parent_index])
        elif tag == "a":
            fields = rest.split()
            require(len(fields) == 2, f"line {line_number}: malformed allocation info")
            size = nonnegative_hex(fields[0], line=line_number, field="allocation size")
            trace_index = nonnegative_hex(fields[1], line=line_number, field="allocation trace ref")
            require(trace_index < len(traces), f"line {line_number}: allocation trace out of range")
            allocations.append((size, trace_index))
            active.append(0)
        elif tag == "+":
            index = nonnegative_hex(rest.strip(), line=line_number, field="allocation event ref")
            require(index < len(allocations), f"line {line_number}: allocation event out of range")
            size, trace_index = allocations[index]
            active[index] += 1
            stack = traces[trace_index]
            stack_bytes[stack] += size
            stack_calls[stack] += 1
            allocation_events += 1
            requested_bytes += size
            live_bytes += size
            peak_live = max(peak_live, live_bytes)
        elif tag == "-":
            index = nonnegative_hex(rest.strip(), line=line_number, field="deallocation event ref")
            require(index < len(allocations), f"line {line_number}: deallocation event out of range")
            require(active[index] > 0, f"line {line_number}: deallocation without live allocation")
            size, _ = allocations[index]
            active[index] -= 1
            deallocation_events += 1
            live_bytes -= size
            require(live_bytes >= 0, f"line {line_number}: live heap underflow")
        elif tag == "R":
            nonnegative_hex(rest.strip(), line=line_number, field="RSS sample")
        elif tag == "#":
            comment = rest.strip()
            for name in ("strings", "ips"):
                prefix = name + ":"
                if comment.startswith(prefix):
                    count = comment[len(prefix):].strip()
                    nonnegative_hex(count, line=line_number, field=f"{name} summary count")
                    require(name not in terminal, f"line {line_number}: repeated {name} summary")
                    terminal.add(name)
                    break
        elif tag in {"", "I", "c"}:
            pass
        else:
            raise ParseError(f"line {line_number}: unsupported record {tag!r}")

    require(version == EXPECTED_FORMAT, f"unsupported Heaptrack format {version!r}")
    require(command is not None, "profile has no command")
    require(terminal == {"strings", "ips"}, "profile lacks terminal strings/ips summaries")
    require(not data_after_terminal, "profile has records after terminal summaries")
    require(allocation_events >= deallocation_events, "deallocation count exceeds allocation count")
    expected_metrics = expected.get("metrics", {})
    for key, value in {
        "allocation_event_count": allocation_events,
        "deallocation_event_count": deallocation_events,
        "live_bytes_at_end": live_bytes,
        "peak_live_heap_bytes": peak_live,
        "requested_bytes": requested_bytes,
    }.items():
        if key in expected_metrics:
            require(int(expected_metrics[key]) == value, f"{path}: {key} differs from memory receipt")
    rows = []
    for stack, bytes_value in stack_bytes.items():
        rows.append({
            "bytes": bytes_value,
            "calls": stack_calls[stack],
            "stack": list(stack),
            "contains_read_limited": any(READ_LIMITED in frame for frame in stack),
            "contains_measure_one": any("measure_one" in frame for frame in stack),
        })
    rows.sort(key=lambda row: (-row["bytes"], -row["calls"], row["stack"]))
    read_rows = [row for row in rows if row["contains_read_limited"]]
    return {
        "path": str(path.resolve()),
        "raw": digest_file(path),
        "format": version,
        "command": command,
        "allocation_event_count": allocation_events,
        "deallocation_event_count": deallocation_events,
        "requested_bytes": requested_bytes,
        "peak_live_heap_bytes": peak_live,
        "live_bytes_at_end": live_bytes,
        "trace_count": len(traces) - 1,
        "allocation_info_count": len(allocations),
        "stack_row_count": len(rows),
        "read_limited_bytes": sum(row["bytes"] for row in read_rows),
        "read_limited_calls": sum(row["calls"] for row in read_rows),
        "read_limited_rows": read_rows,
        "top_stacks": rows[:5],
        "all_rows": rows,
    }


def route_variant(binding: dict[str, Any]) -> tuple[str, str]:
    route = str(binding["route"])
    variant = str(binding["variant"])
    require(route in {"generic_read_at", "owned_vec", "owned_reader"}, f"unknown route {route}")
    require(variant in {"baseline", "candidate"}, f"unknown variant {variant}")
    return route, variant


def expected_path_for(path: pathlib.Path, roots: list[pathlib.Path]) -> pathlib.Path:
    if path.is_file():
        return path
    parts = path.parts
    root_name = None
    for candidate in ("allocation-packet-v3", "allocation-packet-v4"):
        if candidate in parts:
            root_name = candidate
            break
    require(root_name is not None, f"cannot map retained raw path: {path}")
    at = parts.index(root_name)
    relative = pathlib.Path(*parts[at + 1:])
    for root in roots:
        if root.name == root_name:
            candidate = root / relative
            if candidate.is_file():
                return candidate
    raise ParseError(f"retained raw path unavailable: {path}")


def source_proof(repo: pathlib.Path, candidate_path: pathlib.Path) -> dict[str, Any]:
    base = candidate_path.parent
    baseline_path = base / "baseline-source-receipt.json"
    candidate = json.loads(candidate_path.read_text(encoding="utf-8"))
    baseline = json.loads(baseline_path.read_text(encoding="utf-8"))
    require(int(candidate["source_file_count"]) == len(candidate["candidate_source_files"]) == 9355,
            "candidate source closure count mismatch")
    require(int(baseline["source_file_count"]) == 9355, "baseline source count mismatch")
    def file_record(path_value: Any, label: str, *, allow_missing: bool = False) -> dict[str, Any]:
        expected = recorded_descriptor(path_value, label)
        path = pathlib.Path(expected["path"])
        if allow_missing and not path.is_file():
            return {"recorded": expected, "present": False}
        actual = check_recorded_file(path, expected, label)
        actual["present"] = True
        return actual
    cand_receipt = digest_file(candidate_path)
    base_receipt = digest_file(baseline_path)
    # Archive/listing/patch/lock/harness files are the portable source custody
    # witnesses. The source_tree values are directory witnesses and retain their
    # recorded aggregate size/SHA without re-walking 9,355 files here.
    candidate_witnesses = {}
    for key in ("source_archive", "full_source_listing", "candidate_patch", "harness_file", "harness_patch", "cargo_lock", "probe_lock"):
        candidate_witnesses[key] = file_record(candidate[key], f"candidate {key}")
    baseline_witnesses = {}
    for key in ("source_archive", "harness_file", "harness_patch", "cargo_lock"):
        # The baseline receipt names base-head.tar, which was removed during
        # the 0923 external cleanup. Keep that custody limitation explicit;
        # the candidate receipt retains the same archive bytes/hash witness.
        baseline_witnesses[key] = file_record(
            baseline[key], f"baseline {key}", allow_missing=(key == "source_archive")
        )
    return {
        "source_revision": candidate["source_revision"],
        "candidate_base_revision": candidate["candidate_base_revision"],
        "source_file_count": int(candidate["source_file_count"]),
        "candidate_source_file_count": len(candidate["candidate_source_files"]),
        "candidate_receipt": cand_receipt,
        "baseline_receipt": base_receipt,
        "candidate_source_tree_recorded": candidate["source_tree"],
        "baseline_source_tree_recorded": baseline["source_tree"],
        "candidate_witnesses": candidate_witnesses,
        "baseline_witnesses": baseline_witnesses,
        "baseline_source_listing_recorded": {
            "bytes": int(baseline["source_listing_bytes"]),
            "sha256": baseline["source_listing_sha256"],
        },
    }


def aggregate_group(items: list[dict[str, Any]]) -> dict[str, Any]:
    bytes_by_stack: collections.Counter[tuple[str, ...]] = collections.Counter()
    calls_by_stack: collections.Counter[tuple[str, ...]] = collections.Counter()
    for item in items:
        for row in item["all_rows"]:
            key = tuple(row["stack"])
            bytes_by_stack[key] += int(row["bytes"])
            calls_by_stack[key] += int(row["calls"])
    return {
        "profiles": len(items),
        "requested_bytes": sum(int(item["requested_bytes"]) for item in items),
        "allocation_event_count": sum(int(item["allocation_event_count"]) for item in items),
        "bytes_by_stack": bytes_by_stack,
        "calls_by_stack": calls_by_stack,
    }


def compare_pair(reader: dict[str, Any], vec: dict[str, Any], label: str) -> dict[str, Any]:
    rb = reader["bytes_by_stack"]
    vb = vec["bytes_by_stack"]
    rc = reader["calls_by_stack"]
    vc = vec["calls_by_stack"]
    keys = set(rb) | set(vb)
    reader_only = sum(rb[key] for key in rb.keys() - vb.keys())
    vec_only = sum(vb[key] for key in vb.keys() - rb.keys())
    primary_rows = [key for key in rb.keys() if key not in vb and any(READ_LIMITED in frame for frame in key)]
    primary_bytes = sum(rb[key] for key in primary_rows)
    primary_calls = sum(rc[key] for key in primary_rows)
    reader_only_other = reader_only - primary_bytes
    common_rows = []
    for key in sorted(rb.keys() & vb.keys()):
        delta = rb[key] - vb[key]
        calls_delta = rc[key] - vc[key]
        if delta or calls_delta:
            common_rows.append({
                "delta_bytes": delta,
                "delta_calls": calls_delta,
                "reader_bytes": rb[key],
                "vec_bytes": vb[key],
                "stack": list(key),
            })
    common_rows.sort(key=lambda row: (-abs(row["delta_bytes"]), -row["delta_bytes"], row["stack"]))
    total_delta = reader["requested_bytes"] - vec["requested_bytes"]
    route_equivalent_net = reader_only_other - vec_only
    accounted = primary_bytes + route_equivalent_net + sum(row["delta_bytes"] for row in common_rows)
    require(accounted == total_delta, f"{label}: stack accounting does not conserve total delta")
    require(reader_only_other == vec_only,
            f"{label}: route-equivalent source setup totals differ ({reader_only_other} vs {vec_only})")
    return {
        "reader_profiles": reader["profiles"],
        "vec_profiles": vec["profiles"],
        "reader_requested_bytes": reader["requested_bytes"],
        "vec_requested_bytes": vec["requested_bytes"],
        "reader_allocation_events": reader["allocation_event_count"],
        "vec_allocation_events": vec["allocation_event_count"],
        "total_delta_bytes": total_delta,
        "primary_read_limited_bytes": primary_bytes,
        "primary_read_limited_calls": primary_calls,
        "primary_stack_count": len(primary_rows),
        "primary_stacks": [list(key) for key in sorted(primary_rows)],
        "reader_route_only_bytes": reader_only,
        "reader_route_only_calls": sum(rc[key] for key in rb.keys() - vb.keys()),
        "reader_route_only_non_read_limited_bytes": reader_only_other,
        "vec_route_only_bytes": vec_only,
        "vec_route_only_calls": sum(vc[key] for key in vb.keys() - rb.keys()),
        "route_equivalent_setup_net_bytes": route_equivalent_net,
        "common_stack_delta_bytes": sum(row["delta_bytes"] for row in common_rows),
        "common_stack_delta_calls": sum(row["delta_calls"] for row in common_rows),
        "common_nonzero_rows": common_rows,
        "accounting_check": {
            "primary_plus_route_equivalent_plus_common": accounted,
            "total_delta": total_delta,
            "pass": accounted == total_delta,
            "parent_stack_rows_added": False,
            "each_plus_event_assigned_once": True,
        },
    }


def main() -> int:
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--repo", type=pathlib.Path, default=pathlib.Path("/home/zhuhe/code/litchi"))
    ap.add_argument("--receipt", type=pathlib.Path)
    ap.add_argument("--raw-root", action="append", type=pathlib.Path, dest="raw_roots")
    ap.add_argument("--source-receipt", type=pathlib.Path)
    ap.add_argument("--output", type=pathlib.Path, required=True)
    ap.add_argument("--stacks-output", type=pathlib.Path)
    args = ap.parse_args()
    repo = args.repo.resolve()
    receipt_path = (args.receipt or repo / "docs/performance/results/change-0923/independent-memory-disposition.json").resolve()
    raw_roots = [p.resolve() for p in (args.raw_roots or [
        pathlib.Path("/tmp/litchi-performance-0923/allocation-packet-v3"),
        pathlib.Path("/tmp/litchi-performance-0923/allocation-packet-v4"),
    ])]
    source_receipt_path = (args.source_receipt or pathlib.Path("/tmp/litchi-performance-0923/packet-v3-final2/custody/candidate-source-receipt.json")).resolve()
    output = args.output.resolve()
    require(not output.exists(), f"refusing to replace output: {output}")
    stacks_output = (args.stacks_output.resolve() if args.stacks_output else None)
    if stacks_output:
        require(not stacks_output.exists(), f"refusing to replace stacks output: {stacks_output}")
    receipt = json.loads(receipt_path.read_text(encoding="utf-8"))
    bindings = receipt.get("all_profile_bindings")
    require(isinstance(bindings, list) and len(bindings) == 25, "memory receipt must bind exactly 25 profiles")
    ids = [str(binding.get("id")) for binding in bindings]
    require(len(set(ids)) == 25, "profile binding IDs are not unique")
    parsed: list[dict[str, Any]] = []
    for binding in bindings:
        raw_record = recorded_descriptor(binding.get("raw"), f"{binding['id']} raw")
        expected_path = pathlib.Path(raw_record["path"])
        actual_path = expected_path_for(expected_path, raw_roots)
        actual_raw = check_recorded_file(actual_path, raw_record, f"{binding['id']} raw")
        parsed_item = parse_profile(actual_path, binding)
        route, variant = route_variant(binding)
        parsed_item.update({
            "id": binding["id"], "route": route, "variant": variant,
            "pair": int(binding["pair"]), "leg": int(binding["leg"]),
            "receipt_raw": raw_record, "raw": actual_raw,
            "report": binding.get("report"), "binary": binding.get("binary"),
        })
        parsed.append(parsed_item)

    # No retained 0923 Heaptrack trace carries the 0919 operation marker. This
    # is a checked absence, so the report cannot accidentally promote a whole
    # process stack to an operation-scoped allocation claim.
    require(
        not any(row["contains_measure_one"] for item in parsed for row in item["all_rows"]),
        "unexpected measure_one marker in a whole-process 0923 profile",
    )

    groups: dict[tuple[str, str], list[dict[str, Any]]] = collections.defaultdict(list)
    for item in parsed:
        groups[(item["route"], item["variant"])].append(item)
    expected_group_counts = {
        ("generic_read_at", "baseline"): 5,
        ("generic_read_at", "candidate"): 4,
        ("owned_vec", "baseline"): 4,
        ("owned_vec", "candidate"): 4,
        ("owned_reader", "baseline"): 4,
        ("owned_reader", "candidate"): 4,
    }
    for key, count in expected_group_counts.items():
        require(len(groups[key]) == count, f"profile group {key} has {len(groups[key])}, expected {count}")
    aggregate_groups = {}
    for key, items in sorted(groups.items()):
        aggregate = aggregate_group(items)
        aggregate_groups[f"{key[0]}:{key[1]}"] = {
            "profiles": aggregate["profiles"],
            "requested_bytes": aggregate["requested_bytes"],
            "allocation_event_count": aggregate["allocation_event_count"],
            "read_limited_bytes": sum(item["read_limited_bytes"] for item in items),
            "read_limited_calls": sum(item["read_limited_calls"] for item in items),
        }
    pairs = {}
    for variant in ("baseline", "candidate"):
        reader_group = aggregate_group(groups[("owned_reader", variant)])
        vec_group = aggregate_group(groups[("owned_vec", variant)])
        pairs[variant] = compare_pair(
            {**reader_group, "profiles": len(groups[("owned_reader", variant)])},
            {**vec_group, "profiles": len(groups[("owned_vec", variant)])},
            variant,
        )

    # Source/build custody is read from the external 0923 receipts. The CLI
    # source-receipt argument is accepted as an override for the candidate path,
    # but all sibling witnesses remain tied to that receipt's directory.
    if source_receipt_path != pathlib.Path("/tmp/litchi-performance-0923/packet-v3-final2/custody/candidate-source-receipt.json").resolve():
        require(source_receipt_path.is_file(), f"source receipt missing: {source_receipt_path}")
    proof = source_proof(repo, source_receipt_path)
    binary_proof = {}
    for variant, value in receipt.get("binary_identities", {}).items():
        recorded = recorded_descriptor(value, f"{variant} binary")
        binary_path = pathlib.Path(recorded["path"])
        if binary_path.is_file():
            binary_proof[variant] = {"present": True, "actual": check_recorded_file(binary_path, recorded, f"{variant} binary")}
        else:
            binary_proof[variant] = {"present": False, "recorded": recorded}
    profile_summaries = []
    for item in parsed:
        profile_summaries.append({
            "id": item["id"], "route": item["route"], "variant": item["variant"],
            "pair": item["pair"], "leg": item["leg"], "raw": item["raw"],
            "allocation_event_count": item["allocation_event_count"],
            "deallocation_event_count": item["deallocation_event_count"],
            "requested_bytes": item["requested_bytes"],
            "peak_live_heap_bytes": item["peak_live_heap_bytes"],
            "live_bytes_at_end": item["live_bytes_at_end"],
            "trace_count": item["trace_count"], "stack_row_count": item["stack_row_count"],
            "read_limited_bytes": item["read_limited_bytes"],
            "read_limited_calls": item["read_limited_calls"],
            "read_limited_stacks": item["read_limited_rows"],
            "top_stacks": item["top_stacks"][:3],
        })
    # Keep the previously executed parser-negative-control receipt visible in
    # the 0924 audit. These controls are parser-integrity evidence only.
    controls_path = repo / "docs/performance/results/change-0923/allocation-parser-controls-v2.json"
    parser_controls = None
    if controls_path.is_file():
        controls = json.loads(controls_path.read_text(encoding="utf-8"))
        require(controls.get("status") == "pass", "retained parser controls are not passing")
        require(all(c.get("exit_code") == 1 and c.get("expected") == "reject" for c in controls.get("controls", [])),
                "retained parser controls include a non-rejecting case")
        parser_controls = {
            "receipt": digest_file(controls_path),
            "status": controls["status"],
            "control_count": len(controls.get("controls", [])),
            "names": [c["name"] for c in controls.get("controls", [])],
            "scope": controls.get("scope"),
        }

    result = {
        "schema": "litchi.performance.0924.heaptrack-stack-attribution.v2",
        "status": "descriptive retained-profile analysis; no recapture or qualification",
        "raw_only": True,
        "profile_count": len(parsed),
        "receipt": {"path": str(receipt_path), "descriptor": digest_file(receipt_path), "expected_bindings": 25},
        "raw_roots": [str(root) for root in raw_roots],
        "source_proof": proof,
        "binary_proof": binary_proof,
        "parser_controls": parser_controls,
        "stack_method": {
            "format": EXPECTED_FORMAT,
            "event_unit": "one '+' allocation event",
            "assignment": "each event contributes its allocation-info size once to the complete leaf->root trace stack",
            "parent_rows_added": False,
            "comparison": "exact full leaf-stack keys; route-only setup rows shown gross on each side and netted once",
            "operation_boundary": "none in Heaptrack; contains_measure_one was absent from all 25 profiles",
            "read_limited_interpretation": "source-ingress stack observed under Package::from_reader; retained whole-process profile cannot prove timing phase by itself",
        },
        "groups": aggregate_groups,
        "owned_reader_vs_owned_vec": pairs,
        "profiles": profile_summaries,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    if stacks_output:
        stacks_output.parent.mkdir(parents=True, exist_ok=True)
        # Keep all exact rows outside the repository for independent review.
        full_rows = [{"id": item["id"], "route": item["route"], "variant": item["variant"],
                      "raw": item["raw"], "rows": item["all_rows"]} for item in parsed]
        stacks_output.write_text(json.dumps(full_rows, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        result["stack_rows_external"] = {
            "descriptor": digest_file(stacks_output),
            "regeneration_argv": [str(pathlib.Path(sys.executable).resolve()), "-B", str(pathlib.Path(__file__).resolve()), *sys.argv[1:]],
            "script": digest_file(pathlib.Path(__file__).resolve()),
        }
    output.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    if stacks_output:
        # The full stack file is deliberately external to the repository.
        pass
    print(json.dumps({
        "output": str(output), "profiles": len(parsed),
        "baseline_gap": pairs["baseline"]["total_delta_bytes"],
        "candidate_gap": pairs["candidate"]["total_delta_bytes"],
        "stack_output": str(stacks_output) if stacks_output else None,
    }, sort_keys=True))
    return 0

if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (OSError, ValueError, ParseError, subprocess.SubprocessError) as exc:
        print(f"0924 stack attribution failed: {exc}", file=sys.stderr)
        raise SystemExit(1)
