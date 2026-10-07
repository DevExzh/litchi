#!/usr/bin/env python3
"""Decode retained heaptrack raw profiles without launching a workload.

The raw-format state machine is adapted from the reviewed 0932/0934 decoder:
implicit string/IP/trace/allocation-info identifiers are resolved in one pass,
then every ``+`` allocation is retained in disjoint trace and size aggregates.
This tool only invokes zstd for decompression and refuses to overwrite output.
"""
from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import subprocess
import time
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

SCHEMA = "litchi.performance.0949.heaptrack-raw-decode.v1"
RUN_SCHEMA = "litchi.performance.0949.heaptrack-raw-decode-run.v1"
LABEL_RE = re.compile(r"^[A-Za-z0-9][A-Za-z0-9._-]{0,127}$")


def now():
    return datetime.now(timezone.utc).isoformat(timespec="microseconds")


def sha(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


def desc(path: Path) -> dict[str, object]:
    require(path.is_file() and not path.is_symlink(), f"profile is not a regular file: {path}")
    return {"path": str(path.absolute()), "bytes": path.stat().st_size, "sha256": sha(path)}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def write_xb(path, value):
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode()
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def hx(value: str) -> int:
    try:
        return int(value, 16)
    except ValueError as error:
        raise RuntimeError(f"invalid hexadecimal field {value!r}") from error


def frame_from_i(fields: list[str], strings: dict[int, str]) -> dict[str, object]:
    require(len(fields) >= 3, f"malformed i record: {fields!r}")
    address = hx(fields[1])
    module_id = hx(fields[2])
    refs = [hx(x) for x in fields[3:]]
    missing = []

    def resolve(ref):
        if ref == 0:
            return None
        if ref not in strings:
            missing.append(ref)
            return f"<missing-string:{ref:x}>"
        return strings[ref]

    locations = []
    if len(refs) == 1:
        locations.append({
            "function": resolve(refs[0]),
            "file": None,
            "line": None,
        })
    else:
        for index in range(0, len(refs) - 2, 3):
            locations.append({
                "function": resolve(refs[index]),
                "file": resolve(refs[index + 1]),
                "line": refs[index + 2],
            })
        if len(refs) % 3:
            locations.append({
                "function": resolve(refs[-1]),
                "file": None,
                "line": None,
            })
    return {
        "address_hex": f"{address:x}",
        "module_id": module_id,
        "module": resolve(module_id),
        "locations": locations,
        "missing_string_refs": sorted(set(missing)),
    }


def decode(profile: Path, out_dir: Path, label: str, zstd_path: Path) -> dict[str, object]:
    strings = {}
    next_string_id = 1
    ip_raw = {}
    next_ip_index = 1
    trace_ip = {}
    parents = {}
    next_trace_id = 1
    allocation_info = {}
    next_alloc_info = 0
    last_a_index = None
    adjacency_mismatches = 0
    trace_stats = defaultdict(lambda: [0, 0])
    trace_size_stats = defaultdict(lambda: defaultdict(int))
    size_stats = defaultdict(lambda: [0, 0])
    command = None
    version = None
    counts = defaultdict(int)
    string_length_errors = []
    total_calls = 0
    total_bytes = 0
    plus_without_metadata = 0
    require(zstd_path.is_file() and not zstd_path.is_symlink(),
            f"zstd decoder is not a regular file: {zstd_path}")
    process = subprocess.Popen(
        [str(zstd_path), "-dc", str(profile)],
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    started = time.monotonic()
    for raw in process.stdout:
        line = raw.decode("utf-8", "replace").rstrip("\n")
        if not line:
            continue
        fields = line.split(" ", 2)
        kind = fields[0]
        counts[kind] += 1
        if kind not in ("a", "+"):
            last_a_index = None
        if kind == "v":
            require(len(fields) >= 2, "malformed v record")
            version = fields[1:]
        elif kind == "X":
            require(len(fields) >= 2, "malformed X record")
            command = fields[1]
        elif kind == "s" and len(fields) >= 3:
            # The first numeric field is the byte length. String IDs are implicit.
            declared_length = hx(fields[1])
            actual_length = len(fields[2].encode("utf-8"))
            if declared_length != actual_length:
                string_length_errors.append((next_string_id, declared_length, actual_length))
            strings[next_string_id] = fields[2]
            next_string_id += 1
        elif kind == "s":
            raise RuntimeError("malformed s record")
        elif kind == "i":
            # IP/frame IDs are implicit in i-record order.
            ip_fields = line.split()
            require(len(ip_fields) >= 3, "malformed i record")
            ip_raw[next_ip_index] = ip_fields
            next_ip_index += 1
        elif kind == "t" and len(fields) >= 3:
            # t's first field selects an IP record; trace IDs are implicit in t order.
            require(len(fields) == 3, "malformed t record")
            ip_index = hx(fields[1])
            parent = hx(fields[2])
            trace_id = next_trace_id
            next_trace_id += 1
            trace_ip[trace_id] = ip_index
            parents[trace_id] = parent
        elif kind == "t":
            raise RuntimeError("malformed t record")
        elif kind == "a" and len(fields) >= 3:
            # Allocation-info IDs are implicit in a-record order.
            require(len(fields) == 3, "malformed a record")
            allocation_info[next_alloc_info] = (hx(fields[1]), hx(fields[2]))
            last_a_index = next_alloc_info
            next_alloc_info += 1
        elif kind == "a":
            raise RuntimeError("malformed a record")
        elif kind == "+" and len(fields) >= 2:
            require(len(fields) == 2, "malformed + record")
            info_index = hx(fields[1])
            if last_a_index is not None:
                if info_index != last_a_index:
                    adjacency_mismatches += 1
                last_a_index = None
            metadata = allocation_info.get(info_index)
            if metadata is None:
                plus_without_metadata += 1
                continue
            size, trace_id = metadata
            total_calls += 1
            total_bytes += size
            trace_stats[trace_id][0] += 1
            trace_stats[trace_id][1] += size
            trace_size_stats[trace_id][size] += 1
            size_stats[size][0] += 1
            size_stats[size][1] += size
        elif kind == "+":
            raise RuntimeError("malformed + record")
        elif kind in ("-", "R", "c"):
            pass
    stderr = process.stderr.read().decode("utf-8", "replace")
    process.wait()
    if process.returncode != 0:
        raise RuntimeError(f"zstd failed for {profile}: {process.returncode}: {stderr}")

    missing_ip_refs = sorted({ip for ip in trace_ip.values() if ip not in ip_raw})
    frames = {
        trace_id: frame_from_i(ip_raw[ip_index], strings)
        for trace_id, ip_index in trace_ip.items()
        if ip_index in ip_raw
    }
    missing_string_refs = sorted({
        ref
        for frame in frames.values()
        for ref in frame.get("missing_string_refs", [])
    })
    unresolved_trace_ids = sorted(trace_id for trace_id in trace_stats if trace_id not in parents)
    if (string_length_errors or missing_ip_refs or missing_string_refs
            or unresolved_trace_ids or plus_without_metadata or adjacency_mismatches):
        raise RuntimeError(json.dumps({
            "string_length_errors": string_length_errors[:10],
            "missing_ip_refs": missing_ip_refs[:10],
            "missing_string_refs": missing_string_refs[:10],
            "unresolved_trace_ids": unresolved_trace_ids[:10],
            "plus_without_metadata": plus_without_metadata,
            "a_plus_adjacency_mismatches": adjacency_mismatches,
        }, sort_keys=True))

    def stack_for(trace_id: int) -> list[dict[str, object]]:
        stack = []
        seen = set()
        current = trace_id
        while current:
            require(current not in seen,
                    f"trace parent cycle for trace {trace_id:x} at {current:x}")
            seen.add(current)
            frame = frames.get(current)
            require(frame is not None, f"missing frame for trace {current:x}")
            stack.append({"trace_id": current, "frame": frame})
            current = parents.get(current, 0)
        return stack

    def unknown_stack(stack):
        for item in stack:
            frame = item["frame"]
            if frame is None or not frame.get("locations"):
                return True
            function = frame["locations"][0].get("function") or ""
            if function.startswith("0x") or "unknown" in function.lower() or function == "??":
                return True
        return False

    require(not out_dir.exists() and not out_dir.is_symlink(),
            f"refusing existing decoder output: {out_dir}")
    out_dir.mkdir(parents=True, exist_ok=False)
    trace_rows = []
    unknown_any_calls = 0
    unknown_any_bytes = 0
    unknown_leaf_calls = 0
    unknown_leaf_bytes = 0
    for trace_id, (calls, bytes_) in trace_stats.items():
        stack = stack_for(trace_id)
        unknown = unknown_stack(stack)
        leaf_unknown = not stack or stack[0]["frame"] is None or not stack[0]["frame"].get("locations")
        if unknown:
            unknown_any_calls += calls
            unknown_any_bytes += bytes_
        if leaf_unknown:
            unknown_leaf_calls += calls
            unknown_leaf_bytes += bytes_
        trace_rows.append({
            "trace_id": f"{trace_id:x}",
            "calls": calls,
            "requested_bytes": bytes_,
            "parent_trace_id": f"{parents.get(trace_id, 0):x}",
            "unknown_any_frame": unknown,
            "unknown_leaf_frame": leaf_unknown,
            "stack": stack,
        })
    trace_rows.sort(key=lambda row: (-row["requested_bytes"], -row["calls"], int(row["trace_id"], 16)))
    trace_path = out_dir / "trace-aggregates.jsonl"
    with trace_path.open("xb") as stream:
        for row in trace_rows:
            stream.write((json.dumps(row, sort_keys=True) + "\n").encode())
        stream.flush()
        os.fsync(stream.fileno())
    size_path = out_dir / "size-aggregates.tsv"
    with size_path.open("xb") as stream:
        stream.write(b"size_bytes\tcalls\trequested_bytes\n")
        for size, (calls, bytes_) in sorted(size_stats.items()):
            stream.write(f"{size}\t{calls}\t{bytes_}\n".encode())
        stream.flush()
        os.fsync(stream.fileno())
    trace_size_path = out_dir / "trace-size-aggregates.tsv"
    with trace_size_path.open("xb") as stream:
        stream.write(b"trace_id\tsize_bytes\tcalls\trequested_bytes\n")
        for trace_id in sorted(trace_size_stats):
            for size, calls in sorted(trace_size_stats[trace_id].items()):
                stream.write(f"{trace_id:x}\t{size}\t{calls}\t{size * calls}\n".encode())
        stream.flush()
        os.fsync(stream.fileno())
    summary = {
        "schema": SCHEMA,
        "status": "completed",
        "label": label,
        "profile": desc(profile),
        "zstd": {"path": str(zstd_path), "version": zstd_version(zstd_path)},
        "header_version": version,
        "debuggee_command": command,
        "raw_marker_counts": dict(sorted(counts.items())),
        "string_length_errors": string_length_errors,
        "string_definitions": len(strings),
        "ip_definitions": len(ip_raw),
        "trace_definitions": len(parents),
        "allocation_info_definitions": len(allocation_info),
        "allocation_events": total_calls,
        "requested_bytes_total": total_bytes,
        "unique_trace_ids": len(trace_stats),
        "unresolved_trace_definitions": len(unresolved_trace_ids),
        "deallocation_events": counts.get("-", 0),
        "allocation_minus_deallocation": total_calls - counts.get("-", 0),
        "plus_without_metadata": plus_without_metadata,
        "a_plus_adjacency_mismatches": adjacency_mismatches,
        "unknown_any_frame": {
            "trace_count": sum(1 for row in trace_rows if row["unknown_any_frame"]),
            "calls": unknown_any_calls,
            "requested_bytes": unknown_any_bytes,
        },
        "unknown_leaf_frame": {
            "trace_count": sum(1 for row in trace_rows if row["unknown_leaf_frame"]),
            "calls": unknown_leaf_calls,
            "requested_bytes": unknown_leaf_bytes,
        },
        "outputs": {
            "trace_aggregates": str(trace_path),
            "size_aggregates": str(size_path),
            "trace_size_aggregates": str(trace_size_path),
        },
        "decoder_scope": "Disjoint by implicit allocation-info index and trace_id: every '+' event looks up its referenced a-record's size and trace. Requested bytes are before frees; this is whole-profile allocation accounting.",
    }
    summary_path = out_dir / "summary.json"
    write_xb(summary_path, summary)
    return summary


def zstd_version(zstd_path: Path) -> str:
    result = subprocess.run([str(zstd_path), "--version"], check=False,
                            capture_output=True, text=True)
    require(result.returncode == 0, f"cannot query zstd version: {result.stderr.strip()}")
    return (result.stdout or result.stderr).strip()


def parse_profile(value: str) -> tuple[str, Path]:
    label, separator, raw_path = value.partition("=")
    require(separator and LABEL_RE.fullmatch(label) is not None and raw_path,
            "--profile must be LABEL=PROFILE with a safe unique label")
    path = Path(raw_path)
    require(path.is_file() and not path.is_symlink(),
            f"profile is not a regular file: {path}")
    return label, path


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--profile", action="append", required=True,
                        help="retained raw profile as LABEL=PATH; repeat per capture")
    parser.add_argument("--output-root", type=Path, required=True,
                        help="new directory for decoder outputs")
    parser.add_argument("--zstd", type=Path, default=Path("/usr/bin/zstd"))
    args = parser.parse_args()

    profiles = [parse_profile(value) for value in args.profile]
    labels = [label for label, _path in profiles]
    require(len(labels) == len(set(labels)), "duplicate profile label")
    output_root = args.output_root.absolute()
    require(not output_root.exists() and not output_root.is_symlink(),
            f"refusing existing decoder output root: {output_root}")
    output_root.mkdir(parents=True, exist_ok=False)
    started = now()
    rows = []
    try:
        for label, profile in profiles:
            rows.append(decode(profile, output_root / label, label, args.zstd.absolute()))
    except Exception as error:
        # Preserve any completed/partial child outputs and leave a terminal
        # failure record so an offline decode cannot be mistaken for success.
        write_xb(output_root / "decoder-terminal.json", {
            "schema": RUN_SCHEMA,
            "status": "failed",
            "started_utc": started,
            "ended_utc": now(),
            "profiles_completed": rows,
            "profiles_requested": labels,
            "output_root": str(output_root),
            "failure": str(error),
            "scope": "offline raw heaptrack decoding only; no workload or profiler launch",
        })
        raise
    write_xb(output_root / "decoder-terminal.json", {
        "schema": RUN_SCHEMA,
        "status": "terminal",
        "started_utc": started,
        "ended_utc": now(),
        "profiles": rows,
        "output_root": str(output_root),
        "scope": "offline raw heaptrack decoding only; no workload or profiler launch",
    })
    print(json.dumps({"output": str(output_root), "profiles": len(rows)}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
