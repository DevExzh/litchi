#!/usr/bin/env python3
"""Offline analysis of the retained 0788 heaptrack packet.

This module deliberately has no profiler, benchmark, or build side effects.
The root capture driver owns creation of the heaptrack traces and the
``heaptrack_print`` text files.  Once those files have been retained this
script verifies their receipts, parses the summaries and peak flamegraphs,
and records a bounded attribution report.  Heaptrack's intercepted
allocation peak is kept separate from RSS and operation timing throughout
the report.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import sys
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
SCHEMA = "litchi.cached-part-heap-analysis.0788.v1"
HEAP_DIR = HERE / "heaptrack"
OUT_JSON = HERE / "heap-analysis.json"
OUT_MD = HERE / "heap-analysis.md"
PLAN_PATH = HERE / "plan.json"

HEX = frozenset("0123456789abcdefABCDEF")
LEG_NAMES = ("before", "after")
SUMMARY_LIMIT = 12
FLAMEGRAPH_LIMIT = 12
FRAME_MARKERS = (
    "litchi",
    "source_backed",
    "read_parts",
    "partcache",
    "cached",
    "serial",
    "thread",
    "worker",
    "alloc",
    "malloc",
    "operator new",
    "std::new",
    "jemalloc",
    "glibc",
)


class EvidenceError(RuntimeError):
    """A retained artifact is absent, stale, or internally contradictory."""


def fail(message: str) -> None:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1 << 20), b""):
                digest.update(chunk)
    except OSError as error:
        fail(f"cannot hash {path}: {error}")
    return digest.hexdigest()


def is_sha(value: Any) -> bool:
    return isinstance(value, str) and len(value) == 64 and all(char in HEX for char in value)


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON evidence: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON evidence {path}: {error}")


def read_text(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing text evidence: {path}")
    try:
        return path.read_text(encoding="utf-8", errors="replace")
    except OSError as error:
        fail(f"cannot read text evidence {path}: {error}")


def rel(path: Path) -> str:
    try:
        return str(path.resolve().relative_to(HERE.resolve()))
    except ValueError:
        return str(path)


def resolve_packet_path(raw: Any, label: str) -> Path:
    require(isinstance(raw, str) and raw, f"{label}.path is missing")
    text = raw.replace("\\", "/")
    marker = "/docs/performance/results/change-0788/"
    if marker in text:
        candidate = HERE / text.split(marker, 1)[1]
    elif text.startswith("docs/performance/results/change-0788/"):
        candidate = HERE / text.split("change-0788/", 1)[1]
    else:
        candidate = Path(raw) if Path(raw).is_absolute() else HERE / raw
    candidate = candidate.resolve(strict=False)
    try:
        candidate.relative_to(HERE.resolve())
    except ValueError:
        fail(f"{label} escapes the 0788 packet: {raw}")
    return candidate


def artifact(value: Any, label: str, *, allow_empty: bool = False) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label} is not an artifact receipt")
    expected_bytes = value.get("bytes", value.get("size"))
    expected_sha = value.get("sha256", value.get("digest"))
    require(type(expected_bytes) is int and expected_bytes >= 0,
            f"{label}.bytes is invalid")
    require(allow_empty or expected_bytes > 0, f"{label}.bytes is empty")
    require(is_sha(expected_sha), f"{label}.sha256 is invalid")
    path = resolve_packet_path(value.get("path"), label)
    require(path.is_file() and not path.is_symlink(), f"missing {label}: {path}")
    require(path.stat().st_size == expected_bytes, f"{label}.bytes changed")
    require(sha256(path) == expected_sha, f"{label}.sha256 changed")
    return {"path": rel(path), "bytes": expected_bytes, "sha256": expected_sha}


def file_identity(path: Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"missing retained file: {path}")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def nested(value: Any, *paths: str) -> Any:
    """Return the first present value among dotted paths."""
    for path in paths:
        cursor = value
        found = True
        for part in path.split("."):
            if not isinstance(cursor, dict) or part not in cursor:
                found = False
                break
            cursor = cursor[part]
        if found:
            return cursor
    return None


def receipt_artifact(row: dict[str, Any], label: str, *names: str) -> dict[str, Any]:
    paths = list(names) + [label]
    for name in paths:
        value = nested(row, name)
        if value is not None:
            return artifact(value, label)
    artifacts = row.get("artifacts")
    if isinstance(artifacts, dict):
        for name in paths:
            if name in artifacts:
                return artifact(artifacts[name], label)
    fail(f"{label} artifact receipt is missing")


def normalise_command(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, list):
        require(all(isinstance(item, str) for item in value), "command has non-string token")
        return value
    require(isinstance(value, str), "command is neither argv nor text")
    # Commands are only inspected for options; shell parsing is unnecessary
    # and would make the report depend on quoting details.
    return value.split()


def command_option(command: list[str], short: str, long: str) -> str | None:
    for index, token in enumerate(command):
        if token in (short, long) and index + 1 < len(command):
            return command[index + 1]
        if token.startswith(long + "="):
            return token.split("=", 1)[1]
    return None


def has_false_merge_option(row: dict[str, Any]) -> bool:
    explicit = nested(row, "print_options.merge_backtraces", "print.merge_backtraces",
                      "heaptrack_print.merge_backtraces", "merge_backtraces")
    if explicit is False or explicit == 0 or explicit == "0" or str(explicit).lower() == "false":
        return True
    for key in ("print_command", "heaptrack_print_command", "commands.print",
                "commands.heaptrack_print"):
        command = normalise_command(nested(row, key))
        for index, token in enumerate(command):
            lowered = token.lower()
            if lowered in ("-m", "--merge-backtraces") and index + 1 < len(command):
                value = command[index + 1].lower()
                if value in ("0", "false", "off", "no"):
                    return True
            if lowered.startswith("--merge-backtraces="):
                if lowered.split("=", 1)[1] in ("0", "false", "off", "no"):
                    return True
    return False


def parse_int(value: str) -> int | None:
    match = re.search(r"[-+]?\d[\d,]*", value)
    return int(match.group(0).replace(",", "")) if match else None


def parse_bytes(value: str) -> int | float | None:
    """Parse heaptrack's decimal/IEC byte spellings without rounding to zero."""
    text = value.strip().replace(",", "")
    match = re.search(r"([-+]?(?:\d+(?:\.\d*)?|\.\d+))\s*([KMGTPE]i?B?|B)?",
                      text, re.IGNORECASE)
    if not match:
        return None
    number = float(match.group(1))
    unit = (match.group(2) or "").lower()
    factors = {
        "": 1,
        "b": 1,
        "k": 1000,
        "kb": 1000,
        "ki": 1024,
        "kib": 1024,
        "m": 1000**2,
        "mb": 1000**2,
        "mi": 1024**2,
        "mib": 1024**2,
        "g": 1000**3,
        "gb": 1000**3,
        "gi": 1024**3,
        "gib": 1024**3,
        "t": 1000**4,
        "tb": 1000**4,
        "ti": 1024**4,
        "tib": 1024**4,
        "p": 1000**5,
        "pb": 1000**5,
        "pi": 1024**5,
        "pib": 1024**5,
        "e": 1000**6,
        "eb": 1000**6,
        "ei": 1024**6,
        "eib": 1024**6,
    }
    if unit not in factors:
        return None
    result = number * factors[unit]
    return int(result) if result.is_integer() else result


def metric_key(label: str) -> str | None:
    lower = label.lower().replace("_", " ").replace("-", " ")
    lower = re.sub(r"\s+", " ", lower).strip()
    if "rss" in lower:
        return None
    if "calls to allocation functions" in lower:
        return "allocation_count"
    if "leaked allocation" in lower and "bytes" not in lower and "memory" not in lower:
        return "leaked_allocations"
    if lower in ("allocations", "allocation count", "number of allocations"):
        return "allocation_count"
    if "temporary memory allocation" in lower or lower in ("temporary allocations", "temporary allocation count"):
        return "temporary_allocations"
    if "temporary allocation" in lower and not any(x in lower for x in ("byte", "memory", "consumption")):
        return "temporary_allocations"
    if "peak heap memory" in lower or "peak memory consumption" in lower:
        return "peak_live_bytes"
    if "peak" in lower and any(x in lower for x in ("heap", "memory", "consumption")):
        return "peak_live_bytes"
    if "total" in lower and lower.startswith("total allocated"):
        return "total_allocated_bytes"
    if "total" in lower and ("allocat" in lower or "heap" in lower) and "byte" in lower:
        return "total_allocated_bytes"
    if "total" in lower and ("leak" in lower or "leaked" in lower):
        return "leaked_bytes"
    if "temporary" in lower and any(x in lower for x in ("byte", "memory consumed", "memory allocated", "size")):
        return "temporary_bytes"
    if "leaked" in lower and any(x in lower for x in ("byte", "memory", "size")):
        return "leaked_bytes"
    if "peak" in lower and "allocation" in lower and "count" in lower:
        return "peak_allocations"
    return None


def parse_summary(text: str) -> dict[str, Any]:
    """Extract summary metrics, retaining source lines for auditability."""
    aliases: dict[str, list[dict[str, Any]]] = {}
    # Labels stop before a colon; heaptrack_print uses both aligned colon and
    # whitespace-delimited variants across releases.
    colon = re.compile(r"^\s*([^:]{2,80}):\s*([^\n]+?)\s*$")
    for line in text.splitlines():
        match = colon.match(line)
        if not match:
            continue
        label, raw = match.groups()
        key = metric_key(label)
        if key is None:
            continue
        if key in ("allocation_count", "temporary_allocations", "leaked_allocations", "peak_allocations"):
            parsed: int | float | None = parse_int(raw)
        else:
            parsed = parse_bytes(raw)
        if parsed is None:
            continue
        aliases.setdefault(key, []).append({"value": parsed, "label": label.strip(), "raw": raw.strip()})

    # Some builds print "N allocations" without a colon in their summary.
    loose_patterns = (
        ("allocation_count", r"^\s*(\d[\d,]*)\s+allocations?\b"),
        ("allocation_count", r"^\s*(\d[\d,]*)\s+calls to allocation functions\b"),
        ("temporary_allocations", r"^\s*(\d[\d,]*)\s+temporary allocations?\b"),
        ("temporary_allocations", r"^\s*(\d[\d,]*)\s+temporary memory allocations?\b"),
        ("leaked_allocations", r"^\s*(\d[\d,]*)\s+leaked allocations?\b"),
    )
    for key, pattern in loose_patterns:
        for match in re.finditer(pattern, text, re.IGNORECASE | re.MULTILINE):
            aliases.setdefault(key, []).append({"value": int(match.group(1).replace(",", "")),
                                                  "label": key, "raw": match.group(0).strip()})

    metrics: dict[str, Any] = {}
    for key, rows in aliases.items():
        # First occurrence is the summary value.  Keep every matching line so
        # a reviewer can see when a version emits two similarly named fields.
        metrics[key] = rows[0]["value"]
    approximate = sorted(key for key, rows in aliases.items()
                         if key.endswith("_bytes")
                         and re.search(r"\d\s*[KMGTPE]i?B?\b", rows[0]["raw"], re.IGNORECASE))
    required = {"allocation_count", "total_allocated_bytes", "peak_live_bytes",
                "leaked_bytes", "temporary_allocations"}
    return {"metrics": metrics, "source_lines": aliases,
            "display": {key: rows[0]["raw"] for key, rows in aliases.items()},
            "approximate": approximate,
            "available": sorted(metrics),
            "missing": sorted(required - set(metrics)),
            "optional_unavailable": sorted({"temporary_bytes"} - set(metrics))}


def parse_ranked_sections(text: str) -> dict[str, list[dict[str, Any]]]:
    """Parse bounded call-stack snippets from heaptrack_print sections.

    Text format has changed slightly between heaptrack versions.  We retain
    only lines carrying an explicit numeric cost and the following indented
    frames, avoiding claims based on decorative percentages or addresses.
    """
    section_names = {
        "peak": ("peak consumption", "peak memory consumers", "top allocators by peak", "print peaks"),
        "allocators": ("most calls to allocation functions", "allocation count", "top allocators", "allocators"),
        "temporary": ("temporary allocations", "temporary"),
        "leaks": ("leaked allocations", "leaks"),
    }
    current: str | None = None
    rows: dict[str, list[dict[str, Any]]] = {key: [] for key in section_names}
    pending: dict[str, Any] | None = None
    heading = re.compile(r"^\s*[A-Z][A-Z0-9 _-]{3,}\s*$")
    calls_record = re.compile(
        r"^\s*(\d[\d,]*)\s+calls to allocation functions"
        r"(?:\s+with\s+([^\s]+)\s+peak consumption)?\s+from\s*$",
        re.IGNORECASE,
    )
    peak_record = re.compile(
        r"^\s*([^\s]+)\s+peak memory consumed over\s+(\d[\d,]*)\s+calls\s+from\s*$",
        re.IGNORECASE,
    )
    cost_line = re.compile(r"^\s*(?:#?\s*)?(\d+(?:\.\d+)?%?)\s+(.*?)(?:\s+)([-+]?\d[\d,.]*(?:\s*[KMGTPE]i?B?)?)\s*$",
                           re.IGNORECASE)
    for line in text.splitlines():
        stripped = line.strip()
        if heading.match(line):
            lower = stripped.lower()
            if pending is not None and current is not None:
                rows[current].append(pending)
            current = None
            for name, aliases in section_names.items():
                if any(alias in lower for alias in aliases):
                    current = name
                    break
            pending = None
            continue
        if current is None:
            continue
        match = calls_record.match(line)
        if match:
            if pending is not None:
                rows[current].append(pending)
            count, peak = match.groups()
            pending = {"cost": parse_int(count), "calls": parse_int(count),
                       "frames": []}
            if peak is not None:
                pending["peak_bytes"] = parse_bytes(peak)
            continue
        match = peak_record.match(line)
        if match:
            if pending is not None:
                rows[current].append(pending)
            peak, calls = match.groups()
            pending = {"cost": parse_bytes(peak), "peak_bytes": parse_bytes(peak),
                       "calls": parse_int(calls), "frames": []}
            continue
        match = cost_line.match(line)
        if match:
            amount = parse_bytes(match.group(3))
            if amount is None:
                continue
            if pending is not None:
                rows[current].append(pending)
            cost_token, location, _ = match.groups()
            pending = {"cost": amount, "cost_token": cost_token,
                       "location": location.strip(), "frames": []}
            continue
        if pending is not None and line[:1].isspace() and stripped and not stripped.startswith(("-", "=")):
            pending["frames"].append(stripped)
    if current is not None and pending is not None:
        rows[current].append(pending)
    for key in rows:
        rows[key] = rows[key][:SUMMARY_LIMIT]
    return rows


def relevant_frames(frames: Iterable[str]) -> list[str]:
    result = []
    for frame in frames:
        lower = frame.lower()
        if any(marker in lower for marker in FRAME_MARKERS):
            result.append(frame)
    return result


def parse_flamegraph(text: str) -> dict[str, Any]:
    rows: list[dict[str, Any]] = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        match = re.match(r"^(.*?)\s+([-+]?(?:\d+(?:\.\d*)?|\.\d+))\s*$", stripped)
        if not match:
            continue
        stack, raw = match.groups()
        try:
            value = float(raw)
        except ValueError:
            continue
        if not math.isfinite(value) or value < 0 or not stack:
            continue
        frames = [frame for frame in stack.split(";") if frame]
        if not frames:
            continue
        rows.append({"value": int(value) if value.is_integer() else value,
                     "frames": frames,
                     "relevant_frames": relevant_frames(frames)})
    rows.sort(key=lambda row: (-float(row["value"]), row["frames"]))
    total = sum(float(row["value"]) for row in rows)
    total_value: int | float = int(total) if total.is_integer() else total
    return {"cost_type": "peak", "entries": rows[:FLAMEGRAPH_LIMIT],
            "entry_count": len(rows), "relevant_entry_count": sum(bool(row["relevant_frames"]) for row in rows),
            "total_cost_bytes": total_value,
            "source_lines": len(text.splitlines())}


def expected_payloads(shape: str) -> tuple[list[str], str, int]:
    require(shape in ("small", "large", "mixed"), f"unknown payload shape {shape}")
    members: list[str] = []
    sequence = hashlib.sha256()
    total = 0
    for index in range(32):
        size = 4096 if shape == "small" or (shape == "mixed" and index == 31) else 262144
        label = f"litchi-0786-member-{index:02}-".encode()
        offset = index * 11 if index * 11 <= 255 else 0
        period = bytes((label[k % len(label)] + (k % 97) * 3 + offset) % 256
                       for k in range(len(label) * 97))
        payload = (period * ((size + len(period) - 1) // len(period)))[:size]
        members.append(hashlib.sha256(payload).hexdigest())
        sequence.update(index.to_bytes(8, "little"))
        sequence.update(payload)
        total += size
    return members, sequence.hexdigest(), total


def validate_execution_report(path: Path, row: dict[str, Any]) -> dict[str, Any]:
    data = read_json(path)
    require(isinstance(data, dict), f"execution report is not an object: {path}")
    samples = data.get("samples")
    require(isinstance(samples, list) and len(samples) == 30,
            f"heaptrack child sample count is not 30: {path}")
    shape = row["shape"]
    members, sequence, logical_bytes = expected_payloads(shape)
    corpus = data.get("corpus")
    require(isinstance(corpus, dict), f"execution corpus is missing: {path}")
    corpus_members = corpus.get("members")
    require(isinstance(corpus_members, list) and len(corpus_members) == 32,
            f"execution corpus members changed: {path}")
    observed_members = []
    for index, member in enumerate(corpus_members):
        require(isinstance(member, dict), f"execution member {index} is malformed: {path}")
        digest = member.get("sha256")
        require(is_sha(digest), f"execution member {index} digest is malformed: {path}")
        observed_members.append(digest)
    require(observed_members == members, f"execution payload bytes changed: {path}")
    for index, sample in enumerate(samples):
        require(isinstance(sample, dict), f"sample {index} is malformed: {path}")
        verification = sample.get("verification")
        require(isinstance(verification, dict)
                and verification.get("ordered") is True
                and verification.get("all_member_sha256_match") is True
                and verification.get("members") == 32
                and verification.get("sequence_sha256") == sequence
                and verification.get("logical_bytes") == logical_bytes,
                f"sample {index} payload verification failed: {path}")
        resources = sample.get("resources")
        require(isinstance(resources, dict), f"sample {index} resources are missing: {path}")
        expected_cpu = 32 if row["state"] == "primed" else 0
        before = resources.get("before_operation")
        after = resources.get("after_operation")
        dropped = resources.get("after_drop")
        require(all(isinstance(value, dict) for value in (before, after, dropped)),
                f"sample {index} resource phases are missing: {path}")
        require(before.get("cpu_tasks") == expected_cpu
                and after.get("cpu_tasks") == expected_cpu + 32
                and dropped.get("cpu_tasks") == expected_cpu + 32
                and dropped.get("workers") == 0
                and dropped.get("io_concurrency") == 0,
                f"sample {index} task accounting changed: {path}")
        limits = resources.get("limits")
        require(isinstance(limits, dict), f"sample {index} resource limits are missing: {path}")
        for phase in (before, after, dropped):
            for key in ("workers", "io_concurrency", "cpu_tasks"):
                value = phase.get(key)
                limit = limits.get(key)
                require(type(value) is int and value >= 0 and type(limit) is int and 0 <= value <= limit,
                        f"sample {index} resource bound failed for {key}: {path}")
        metrics = sample.get("source_metrics")
        if metrics is not None:
            require(isinstance(metrics, dict), f"sample {index} source metrics malformed: {path}")
            if metrics.get("availability") == "unavailable-normal-build":
                # The heaptrack lane intentionally profiles the unchanged
                # normal native binary, which has no source-metrics feature.
                # Preserve that explicit absence rather than treating nulls
                # as measured zeroes.
                for key in ("logical_calls", "requested_bytes", "returned_bytes",
                            "short_reads", "active_reads_after_operation",
                            "max_simultaneous_reads", "request_size_histogram"):
                    require(metrics.get(key) is None,
                            f"sample {index} unavailable source metric became measured: {path}")
            else:
                require(metrics.get("short_reads") == 0
                        and metrics.get("active_reads_after_operation") == 0
                        and metrics.get("requested_bytes") == metrics.get("returned_bytes"),
                        f"sample {index} source metrics changed: {path}")
                histogram = metrics.get("request_size_histogram")
                logical_calls = metrics.get("logical_calls")
                require(isinstance(histogram, list) and type(logical_calls) is int
                        and sum(histogram) == logical_calls,
                        f"sample {index} source histogram changed: {path}")
    return {"path": rel(path), "bytes": path.stat().st_size, "sha256": sha256(path),
            "samples": len(samples), "logical_bytes": logical_bytes, "payload_verified": True}


def load_heap_receipts() -> tuple[Path, list[dict[str, Any]], dict[str, Any]]:
    candidates = [HEAP_DIR / "receipts.json", HERE / "heaptrack-receipts.json"]
    paths = [path for path in candidates if path.is_file() and not path.is_symlink()]
    require(len(paths) == 1, f"expected one heaptrack receipts file, found {len(paths)}")
    path = paths[0]
    value = read_json(path)
    if isinstance(value, list):
        rows = value
        metadata: dict[str, Any] = {}
    elif isinstance(value, dict):
        rows = value.get("receipts")
        metadata = {key: item for key, item in value.items() if key != "receipts"}
    else:
        fail("heaptrack receipts are neither a list nor an object")
    require(isinstance(rows, list), "heaptrack receipts list is missing")
    require(all(isinstance(row, dict) for row in rows), "heaptrack receipt row is malformed")
    return path, rows, metadata


def load_histogram_receipts() -> tuple[Path, list[dict[str, Any]]]:
    path = HERE / "heap-histograms" / "receipts.json"
    require(path.is_file() and not path.is_symlink(), f"missing histogram receipts: {path}")
    value = read_json(path)
    rows = value.get("rows") if isinstance(value, dict) else value
    require(isinstance(rows, list) and all(isinstance(row, dict) for row in rows),
            "histogram receipts are malformed")
    return path, rows


def parse_histogram(path: Path, label: str) -> dict[str, Any]:
    """Parse heaptrack_print -H's exact size/count TSV output."""
    buckets: list[dict[str, int]] = []
    previous_size = -1
    for number, line in enumerate(read_text(path).splitlines(), 1):
        require(line and "\t" in line and line.count("\t") == 1,
                f"{label} line {number} is not two-column TSV")
        raw_size, raw_count = line.split("\t")
        require(re.fullmatch(r"\d+", raw_size) is not None
                and re.fullmatch(r"\d+", raw_count) is not None,
                f"{label} line {number} has non-decimal fields")
        size, count = int(raw_size), int(raw_count)
        require(size >= 0 and count > 0, f"{label} line {number} has invalid values")
        require(size > previous_size, f"{label} sizes are not strictly increasing")
        previous_size = size
        buckets.append({"size": size, "count": count})
    require(buckets, f"{label} is empty")
    allocation_count = sum(row["count"] for row in buckets)
    total_allocated = sum(row["size"] * row["count"] for row in buckets)
    return {"buckets": buckets, "bucket_count": len(buckets),
            "allocation_count": allocation_count,
            "total_allocated_bytes": total_allocated}


def validate_histogram_receipts(heap_rows: list[dict[str, Any]]) -> tuple[dict[tuple[Any, ...], dict[str, Any]], dict[str, Any]]:
    receipt_path, rows = load_histogram_receipts()
    require(len(rows) == len(heap_rows),
            f"histogram receipt count changed: {len(rows)}")
    expected: dict[tuple[Any, ...], dict[str, Any]] = {}
    for index, row in enumerate(rows):
        require(row.get("exit_code") == 0, f"histogram print failed at row {index}")
        case = row.get("case") if isinstance(row.get("case"), dict) else row
        key_without_trace = (case.get("route"), case.get("shape"), case.get("state"),
                             case.get("task_floor"), case.get("workers"), row.get("leg"),
                             row.get("repeat"))
        source = artifact(row.get("source_trace"), f"histogram source trace {index}")
        histogram = artifact(row.get("histogram"), f"histogram TSV {index}")
        log = artifact(row.get("log"), f"histogram log {index}")
        command = normalise_command(row.get("command"))
        require("-H" in command or any(token.startswith("--print-histogram=") for token in command),
                f"histogram command {index} omits -H")
        require("-f" in command or any(token.startswith("--file=") for token in command),
                f"histogram command {index} omits source -f")
        source_path = resolve_packet_path(source["path"], f"histogram source trace {index}")
        command_source = command_option(command, "-f", "--file")
        command_histogram = command_option(command, "-H", "--print-histogram")
        require(command_source is not None and command_histogram is not None,
                f"histogram command {index} options are incomplete")
        require(resolve_packet_path(command_source, f"histogram command source {index}") == source_path,
                f"histogram command {index} source trace path changed")
        histogram_path = resolve_packet_path(histogram["path"], f"histogram TSV {index}")
        require(resolve_packet_path(command_histogram, f"histogram command output {index}") == histogram_path,
                f"histogram command {index} output path changed")
        stats = parse_histogram(histogram_path,
                                f"histogram TSV {index}")
        key = (*key_without_trace, source["sha256"])
        require(key not in expected, f"duplicate histogram receipt {key_without_trace}")
        expected[key] = {"source_trace": source, "histogram": histogram, "log": log,
                         "command": command, **stats}
    expected_heap_keys = {
        (row.get("route"), row.get("shape"), row.get("state"), row.get("task_floor"),
         row.get("workers"), row.get("leg"), row.get("repeat"), row["trace"]["sha256"])
        for row in heap_rows
    }
    require(set(expected) == expected_heap_keys,
            "histogram receipts do not cover the exact heaptrack trace matrix")
    for row in heap_rows:
        key = (row.get("route"), row.get("shape"), row.get("state"), row.get("task_floor"),
               row.get("workers"), row.get("leg"), row.get("repeat"), row["trace"]["sha256"])
        value = expected[key]
        require(value["source_trace"]["bytes"] == row["trace"].get("bytes", row["trace"].get("size"))
                and value["source_trace"]["sha256"] == row["trace"].get("sha256", row["trace"].get("digest")),
                f"histogram source trace identity changed for {key}")
    return expected, file_identity(receipt_path)


def plan_cases() -> list[dict[str, Any]]:
    plan = read_json(PLAN_PATH)
    spec = plan.get("heaptrack") if isinstance(plan, dict) else None
    require(isinstance(spec, dict), "plan.heaptrack is missing")
    cases = spec.get("cases")
    require(isinstance(cases, list) and len(cases) == 4, "heaptrack case plan changed")
    result = []
    for case in cases:
        require(isinstance(case, dict), "heaptrack case is malformed")
        result.append({"route": case.get("route"), "shape": case.get("shape"),
                       "state": case.get("state"), "task_floor": case.get("task_floor"),
                       "workers": case.get("workers")})
    expected = [
        {"route": "parts", "shape": "small", "state": "fresh", "task_floor": 0, "workers": 4},
        {"route": "parts", "shape": "small", "state": "primed", "task_floor": 0, "workers": 4},
        {"route": "parts", "shape": "small", "state": "primed", "task_floor": 65536, "workers": 4},
        {"route": "parts", "shape": "small", "state": "primed", "task_floor": 0, "workers": 32},
    ]
    require(result == expected, "heaptrack case matrix does not match the frozen plan")
    require(spec.get("repeats") == 2 and spec.get("samples") == 30 and spec.get("warmup") == 3,
            "heaptrack repeat/sample protocol changed")
    require(spec.get("print_merge_backtraces") is False,
            "plan does not disable merged peak backtraces")
    return result


def analyse_heaptrack() -> dict[str, Any]:
    cases = plan_cases()
    receipt_path, rows, receipt_metadata = load_heap_receipts()
    histogram_rows, histogram_receipts = validate_histogram_receipts(rows)
    expected_keys = {
        (case["route"], case["shape"], case["state"], case["task_floor"], case["workers"], leg, repeat)
        for case in cases for leg in LEG_NAMES for repeat in range(2)
    }
    require(len(rows) == len(expected_keys), f"heaptrack receipt count changed: {len(rows)}")
    seen: set[tuple[Any, ...]] = set()
    reports: list[dict[str, Any]] = []
    for index, row in enumerate(rows):
        require(isinstance(row, dict), f"heaptrack receipt {index} is malformed")
        key = (row.get("route"), row.get("shape"), row.get("state"), row.get("task_floor"),
               row.get("workers"), row.get("leg"), row.get("repeat"))
        require(key in expected_keys, f"heaptrack receipt {index} has unknown case {key}")
        require(key not in seen, f"duplicate heaptrack receipt {key}")
        seen.add(key)
        require(row.get("exit_code", 0) == 0, f"heaptrack child failed for {key}")
        require(row.get("samples", 30) == 30 and row.get("warmup", 3) == 3,
                f"heaptrack protocol changed for {key}")
        require(has_false_merge_option(row), f"heaptrack_print merge-backtraces was not disabled for {key}")

        trace = receipt_artifact(row, "trace", "trace", "heaptrack_trace", "raw", "artifacts.trace")
        printed = receipt_artifact(row, "print", "print", "summary", "print_output",
                                   "heaptrack_print", "artifacts.print")
        flame = receipt_artifact(row, "flamegraph", "flamegraph", "peak_stacks", "stacks", "artifacts.flamegraph")
        report_value = nested(row, "report", "execution_report", "child_report", "artifacts.report")
        require(report_value is not None, f"execution report receipt is missing for {key}")
        report_path = resolve_packet_path(report_value.get("path"), f"report {key}")
        execution = validate_execution_report(report_path, row)

        print_text = read_text(resolve_packet_path(printed["path"], "print"))
        flame_text = read_text(resolve_packet_path(flame["path"], "flamegraph"))
        summary = parse_summary(print_text)
        ranked = parse_ranked_sections(print_text)
        flame_rows = parse_flamegraph(flame_text)
        histogram_key = (*key, trace["sha256"])
        histogram = histogram_rows[histogram_key]
        require(summary["metrics"].get("allocation_count") == histogram["allocation_count"],
                f"histogram allocation count disagrees with heaptrack summary for {key}")
        summary["metrics"]["total_allocated_bytes"] = histogram["total_allocated_bytes"]
        summary["source_lines"]["total_allocated_bytes"] = [{
            "value": histogram["total_allocated_bytes"],
            "label": "weighted allocation-size histogram",
            "raw": f"{histogram['total_allocated_bytes']}B",
        }]
        summary["display"]["total_allocated_bytes"] = f"{histogram['total_allocated_bytes']}B (histogram)"
        summary["available"] = sorted(summary["metrics"])
        summary["missing"] = [name for name in summary["missing"] if name != "total_allocated_bytes"]
        summary["histogram_provenance"] = "exact weighted sum of retained heaptrack_print -H size/count TSV"
        peak_summary = summary["metrics"].get("peak_live_bytes")
        peak_total = flame_rows["total_cost_bytes"]
        peak_crosscheck = None
        if isinstance(peak_summary, (int, float)) and isinstance(peak_total, (int, float)):
            difference = peak_total - peak_summary
            tolerance = max(1024.0, abs(float(peak_summary)) * 0.005)
            require(abs(difference) <= tolerance,
                    f"peak flamegraph total disagrees with rounded print summary for {key}")
            peak_crosscheck = {"summary_parsed_bytes": peak_summary,
                               "flamegraph_total_bytes": peak_total,
                               "difference_bytes": difference,
                               "within_print_rounding_tolerance": True,
                               "tolerance_bytes": tolerance}
        peak_attribution = {
            "total_stack_cost_bytes": peak_total,
            "top_stack": flame_rows["entries"][0] if flame_rows["entries"] else None,
            "summary_crosscheck": peak_crosscheck,
            "interpretation": "peak-cost flamegraph attribution; not RSS or operation allocation",
        }
        reports.append({
            "route": row["route"], "shape": row["shape"], "state": row["state"],
            "task_floor": row["task_floor"], "workers": row["workers"],
            "leg": row["leg"], "repeat": row["repeat"],
            "trace": trace, "print": printed, "flamegraph": flame,
            "allocation_histogram": histogram,
            "execution_report": execution,
            "print_options": {"merge_backtraces": False, "flamegraph_cost_type": "peak",
                               "summary_limit": SUMMARY_LIMIT, "flamegraph_limit": FLAMEGRAPH_LIMIT},
            "summary": summary, "ranked_call_stacks": ranked, "peak_flamegraph": flame_rows,
            "peak_attribution": peak_attribution,
            "interpretation": "intercepted allocation profile; not RSS, operation timing, or causal proof",
        })
    require(seen == expected_keys, "heaptrack receipt matrix is incomplete")
    reports.sort(key=lambda row: (row["route"], row["shape"], row["state"], row["task_floor"],
                                 row["workers"], row["leg"], row["repeat"]))
    return {"receipts": file_identity(receipt_path), "histogram_receipts": histogram_receipts,
            "receipt_metadata": receipt_metadata, "reports": reports}


def parse_size_output(text: str) -> dict[str, int]:
    result: dict[str, int] = {}
    for line in text.splitlines():
        match = re.match(r"^\s*(\S+)\s+(\d+)\s+(?:0x)?[0-9A-Fa-f]+\s*$", line)
        if match:
            result[match.group(1)] = int(match.group(2))
    return result


def parse_readelf_sections(text: str) -> list[dict[str, Any]]:
    result = []
    # readelf -W -S keeps each section on one line.  The first hexadecimal
    # fields are address, offset, and size; the flags follow entsize.
    pattern = re.compile(
        r"^\s*\[\s*\d+\]\s+(\S+)\s+(\S+)\s+([0-9A-Fa-f]+)\s+"
        r"([0-9A-Fa-f]+)\s+([0-9A-Fa-f]+)\s+([0-9A-Fa-f]+)\s+"
        r"([A-Z]*)\s+\d+\s+\d+\s+\d+\s*$"
    )
    for line in text.splitlines():
        match = pattern.match(line)
        if not match:
            continue
        name, kind, address, offset, size, entsize, flags = match.groups()
        result.append({"name": name, "type": kind, "address": int(address, 16),
                       "offset": int(offset, 16), "size": int(size, 16),
                       "entsize": int(entsize, 16), "flags": flags})
    return result


def classify_section(section: dict[str, Any]) -> str | None:
    name = section["name"]
    flags = section["flags"]
    if not ("A" in flags):
        return None
    if name.startswith(".eh_frame"):
        return "eh_frame"
    if name.startswith((".bss", ".sbss", ".tbss")) or section["type"] == "NOBITS":
        return "bss"
    if "X" in flags or name.startswith((".text", ".init", ".fini", ".plt", ".iplt")):
        return "text"
    if "W" in flags or name.startswith((".data", ".sdata", ".tdata", ".got", ".ctors", ".dtors")):
        return "data"
    return "rodata"


def build_section_diff() -> dict[str, Any]:
    result: dict[str, Any] = {}
    for kind in ("native", "memory"):
        legs: dict[str, Any] = {}
        for leg in LEG_NAMES:
            build_path = HERE / f"build-{leg}" / "build.json"
            build = read_json(build_path)
            sections = build.get("sections", {}).get(kind)
            require(isinstance(sections, dict), f"build-{leg} {kind} section receipts are missing")
            section_art = sections.get("sections")
            require(section_art is not None, f"build-{leg} {kind} readelf section receipt is missing")
            section_receipt = artifact(section_art, f"build-{leg} {kind} readelf sections")
            text = read_text(resolve_packet_path(section_receipt["path"], "readelf sections"))
            parsed = parse_readelf_sections(text)
            require(parsed, f"no ELF sections parsed for build-{leg} {kind}")
            classified = {name: 0 for name in ("text", "rodata", "data", "bss", "eh_frame")}
            rows = []
            for section in parsed:
                category = classify_section(section)
                row = {**section, "category": category}
                rows.append(row)
                if category is not None:
                    classified[category] += section["size"]
            size_art = sections.get("size")
            size_receipt = artifact(size_art, f"build-{leg} {kind} size") if size_art is not None else None
            size_map = parse_size_output(read_text(resolve_packet_path(size_receipt["path"], "size"))) if size_receipt else {}
            binary_value = build.get("binaries", {}).get(kind)
            require(isinstance(binary_value, dict),
                    f"build-{leg} {kind} binary receipt is missing")
            binary = {"name": Path(str(binary_value.get("path", kind))).name,
                      "bytes": binary_value.get("bytes", binary_value.get("size")),
                      "sha256": binary_value.get("sha256", binary_value.get("digest"))}
            require(type(binary["bytes"]) is int and binary["bytes"] > 0
                    and is_sha(binary["sha256"]),
                    f"build-{leg} {kind} binary receipt is malformed")
            legs[leg] = {"binary": binary,
                         "allocated_categories": classified, "allocated_total": sum(classified.values()),
                         "readelf_sections": rows, "size_sections": size_map,
                         "readelf_artifact": section_receipt, "size_artifact": size_receipt}
        before = legs["before"]["allocated_categories"]
        after = legs["after"]["allocated_categories"]
        delta = {key: after[key] - before[key] for key in before}
        result[kind] = {"before": legs["before"], "after": legs["after"],
                        "delta": delta,
                        "allocated_total_delta": sum(delta.values())}
    return result


def make_report() -> dict[str, Any]:
    plan = read_json(PLAN_PATH)
    heap = analyse_heaptrack()
    section_diff = build_section_diff()
    by_case: dict[str, dict[str, int]] = {}
    for row in heap["reports"]:
        key = f"{row['shape']}/{row['state']}/floor-{row['task_floor']}/width-{row['workers']}"
        item = by_case.setdefault(key, {"before": 0, "after": 0})
        item[row["leg"]] += 1
    require(all(item == {"before": 2, "after": 2} for item in by_case.values()),
            "heaptrack report multiplicities changed")
    pair_map: dict[tuple[Any, ...], dict[str, Any]] = {}
    for row in heap["reports"]:
        key = (row["route"], row["shape"], row["state"], row["task_floor"],
               row["workers"], row["repeat"])
        pair_map.setdefault(key, {})[row["leg"]] = row["peak_attribution"]["total_stack_cost_bytes"]
    require(all(set(pair) == {"before", "after"} for pair in pair_map.values()),
            "heaptrack peak pair multiplicities changed")
    peak_pairs = []
    for key, values in sorted(pair_map.items()):
        before, after = values["before"], values["after"]
        peak_pairs.append({"route": key[0], "shape": key[1], "state": key[2],
                           "task_floor": key[3], "workers": key[4], "repeat": key[5],
                           "before_peak_stack_cost_bytes": before,
                           "after_peak_stack_cost_bytes": after,
                           "delta_bytes": after - before})
    return {
        "schema": SCHEMA,
        "purpose": "0788 attribution diagnostic; exact rejected candidate is always restored",
        "plan": {"schema": plan.get("schema"), "heaptrack": plan["heaptrack"]},
        "inputs": {"plan": file_identity(PLAN_PATH), "receipts": heap["receipts"],
                   "histogram_receipts": heap["histogram_receipts"]},
        "settings": {"reports": 16, "samples_per_report": 30, "warmup": 3,
                     "repeats": 2, "merge_backtraces": False,
                     "native_timing_or_rss_pooled": False,
                     "bootstrap": False, "causal_claim": False},
        "reports": heap["reports"],
        "case_multiplicities": by_case,
        "peak_pairs": peak_pairs,
        "elf_section_diff": section_diff,
        "interpretation": {
            "peak": "heaptrack intercepted allocation peak, with merge-backtraces disabled",
            "separation": "This profile is not RSS, smaps, native operation timing, or a causal proof.",
            "scope": "Top stacks are bounded evidence for relevant thread/serial/cache/allocator sites.",
            "missing_summary_fields": "A missing heaptrack_print field remains unavailable rather than being inferred as zero; total allocated bytes come from the retained exact histogram.",
            "rounded_byte_fields": "Human-readable K/M/G byte fields are marked approximate; their raw heaptrack spellings are retained per child.",
        },
    }


def canonical_json(value: Any) -> str:
    return json.dumps(value, indent=2, sort_keys=True, ensure_ascii=False) + "\n"


def markdown(report: dict[str, Any]) -> str:
    lines = [
        "# 0788 heaptrack allocation attribution",
        "",
        "This is a retained heaptrack diagnostic for the rejected cached-Part candidate.",
        "Heaptrack intercepted-allocation peaks are reported separately from RSS and native operation timing.",
        "Merged backtraces were disabled because heaptrack warns that merged peak consumption is inaccurate.",
        "",
        "## Summary",
        "",
        f"- Reports: {len(report['reports'])}; samples per child: {report['settings']['samples_per_report']}; repeats: {report['settings']['repeats']}.",
        "- Allocation fields absent from a heaptrack version remain unavailable; the parser does not substitute zero.",
        "- Human-readable heaptrack byte fields retain their raw rounded spelling below; normalized numbers remain in JSON for comparison.",
        "- The report supplies attribution evidence only and makes no causal claim.",
        "",
        "## Children",
        "",
        "| shape | state | floor | width | leg | repeat | allocations | total allocated | peak live | leaked | temporary | flame stacks |",
        "|---|---|---:|---:|---|---:|---:|---:|---:|---:|---:|---:|",
    ]
    for row in report["reports"]:
        metrics = row["summary"]["metrics"]
        display = row["summary"].get("display", {})
        approximate = set(row["summary"].get("approximate", []))
        def value(name: str) -> str:
            if name not in display:
                return "unavailable"
            return ("~" if name in approximate else "") + str(display[name])
        temporary = (value("temporary_bytes") if "temporary_bytes" in metrics
                     else value("temporary_allocations"))
        lines.append("| {shape} | {state} | {floor} | {width} | {leg} | {repeat} | {alloc} | {total} | {peak} | {leaked} | {temporary} | {flame} |".format(
            shape=row["shape"], state=row["state"], floor=row["task_floor"], width=row["workers"],
            leg=row["leg"], repeat=row["repeat"], alloc=value("allocation_count"),
            total=value("total_allocated_bytes"), peak=value("peak_live_bytes"),
            leaked=value("leaked_bytes"), temporary=temporary,
            flame=row["peak_flamegraph"]["entry_count"]))
    lines.extend(["", "## ELF allocated section deltas", "",
                  "Values come from retained `readelf -W -S` output; only allocated sections are classified.",
                  "", "| binary | text | rodata | data | bss | eh_frame | total |", "|---|---:|---:|---:|---:|---:|---:|"])
    for kind in ("native", "memory"):
        delta = report["elf_section_diff"][kind]["delta"]
        lines.append("| {kind} | {text} | {rodata} | {data} | {bss} | {eh} | {total} |".format(
            kind=kind, text=delta["text"], rodata=delta["rodata"], data=delta["data"],
            bss=delta["bss"], eh=delta["eh_frame"], total=report["elf_section_diff"][kind]["allocated_total_delta"]))
    primed_w4_deltas = [pair["delta_bytes"] for pair in report["peak_pairs"]
                        if pair["state"] == "primed" and pair["task_floor"] == 0
                        and pair["workers"] == 4]
    lines.extend(["", "## Exact peak-cost stack pairs", "",
                  "The totals below sum the unmerged peak-cost flamegraph lines and cross-check against the rounded print summary. They describe intercepted allocation peak attribution.",
                  "", "| shape | state | floor | width | repeat | before stack cost | after stack cost | delta |", "|---|---|---:|---:|---:|---:|---:|---:|"])
    for pair in report["peak_pairs"]:
        lines.append("| {shape} | {state} | {floor} | {width} | {repeat} | {before} | {after} | {delta:+} |".format(
            shape=pair["shape"], state=pair["state"], floor=pair["task_floor"], width=pair["workers"],
            repeat=pair["repeat"], before=pair["before_peak_stack_cost_bytes"],
            after=pair["after_peak_stack_cost_bytes"], delta=pair["delta_bytes"]))
    lines.extend(["", "## Top peak stack evidence", "",
                  "The first relevant frame is retained as a bounded attribution label; it does not establish causation.",
                  "", "| shape | state | floor | width | leg | repeat | top cost | first relevant frame |", "|---|---|---:|---:|---|---:|---:|---|"])
    for row in report["reports"]:
        top = row["peak_attribution"].get("top_stack") or {}
        relevant = top.get("relevant_frames") or top.get("frames") or ["unavailable"]
        frame = str(relevant[0]).replace("|", "\\|")
        lines.append("| {shape} | {state} | {floor} | {width} | {leg} | {repeat} | {cost} | {frame} |".format(
            shape=row["shape"], state=row["state"], floor=row["task_floor"], width=row["workers"],
            leg=row["leg"], repeat=row["repeat"], cost=top.get("value", "unavailable"), frame=frame))
    lines.extend(["", f"The native W4 primed floor-0 pair deltas are {', '.join(f'{value:+} B' for value in primed_w4_deltas)} across repeats. These are allocation-profile deltas and establish no RSS cause.",
                  "The intercepted allocation peak is not a native RSS high-water mark and is not pooled with native timing or RSS.", ""])
    return "\n".join(lines)


def main() -> int:
    parser = argparse.ArgumentParser()
    mode = parser.add_mutually_exclusive_group(required=True)
    mode.add_argument("--write", action="store_true")
    mode.add_argument("--check", action="store_true")
    args = parser.parse_args()
    try:
        report = make_report()
        encoded = canonical_json(report)
        rendered = markdown(report)
        if args.write:
            require(not OUT_JSON.exists() and not OUT_MD.exists(), "analysis outputs already exist")
            OUT_JSON.write_text(encoded, encoding="utf-8")
            OUT_MD.write_text(rendered, encoding="utf-8")
        else:
            require(OUT_JSON.is_file() and not OUT_JSON.is_symlink(), "heap-analysis.json is missing")
            require(OUT_MD.is_file() and not OUT_MD.is_symlink(), "heap-analysis.md is missing")
            require(OUT_JSON.read_text(encoding="utf-8") == encoded, "heap-analysis.json is not reproducible")
            require(OUT_MD.read_text(encoding="utf-8") == rendered, "heap-analysis.md is not reproducible")
        print("heaptrack analysis PASS: 16 reports, 480 samples, section deltas retained")
        return 0
    except EvidenceError as error:
        print(f"heaptrack analysis FAILED: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
