#!/usr/bin/env python3
"""Create a deterministic, read-only summary of the 0888 procfs snapshots.

The monitor keeps the raw status and smaps text in gzip JSONL.  This script
does not launch a child, alter a capture, or infer a cause.  It parses every
raw line, reports mapping-category aggregates in bytes, and keeps sampled
maxima separate from the GNU time process high-water value.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
import os
from pathlib import Path
import re
from typing import Any, Iterable


METRICS = (
    "Rss",
    "Pss",
    "Private_Clean",
    "Private_Dirty",
    "Anonymous",
    "FilePmdMapped",
    "AnonHugePages",
)
STATUS_FIELDS = (
    "VmPeak",
    "VmSize",
    "VmHWM",
    "VmRSS",
    "RssAnon",
    "RssFile",
    "RssShmem",
    "VmData",
    "VmStk",
    "VmExe",
    "VmLib",
    "VmPTE",
)
HEADER = re.compile(
    r"^([0-9a-fA-F]+)-([0-9a-fA-F]+)\s+(\S+)\s+"
    r"([0-9a-fA-F]+)\s+(\S+)\s+(\d+)(?:\s+(.*))?$"
)
SMAPS_METRIC = re.compile(r"^(Rss|Pss|Private_Clean|Private_Dirty|Anonymous|FilePmdMapped|AnonHugePages):\s+(\d+)\s+kB$")


def fail(message: str) -> "NoReturn":
    raise ValueError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def relpath(path: Path, packet: Path) -> str:
    root = packet.resolve()
    while root.parent != root and not (root / ".git").exists():
        root = root.parent
    try:
        return path.resolve().relative_to(root).as_posix()
    except ValueError:
        return str(path.resolve())


def descriptor(path: Path, packet: Path) -> dict[str, Any]:
    path = path.resolve()
    if not path.is_file():
        fail(f"missing regular input: {path}")
    return {
        "path": relpath(path, packet),
        "bytes": path.stat().st_size,
        "sha256": sha256(path),
    }


def load_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def as_bytes(value: str, unit: str | None) -> int:
    if unit == "kB":
        return int(value) * 1024
    if unit in (None, ""):
        return int(value)
    fail(f"unsupported procfs unit {unit!r}")


def parse_status(text: str) -> dict[str, int]:
    parsed: dict[str, int] = {}
    for line in text.splitlines():
        if ":" not in line:
            continue
        name, value = line.split(":", 1)
        name = name.strip()
        if name not in STATUS_FIELDS:
            continue
        parts = value.split()
        if not parts:
            fail(f"empty status value for {name}")
        parsed[name] = as_bytes(parts[0], parts[1] if len(parts) > 1 else None)
    missing = [name for name in STATUS_FIELDS if name not in parsed]
    if missing:
        fail(f"status is missing fields: {missing}")
    return parsed


def without_deleted(path: str) -> str:
    return path.removesuffix(" (deleted)")


def classify(path: str, target: str) -> str:
    """Classify only from the smaps pathname form and the invoked binary path."""
    normalized = without_deleted(path)
    if normalized == target:
        return "executable_file"
    if path == "[heap]":
        return "heap"
    if path.startswith("[stack"):
        return "stack"
    if not path or path.startswith("[anon"):
        return "anonymous"
    if path.startswith("/") or path.startswith("memfd:"):
        return "shared_library"
    return "special"


def empty_bucket() -> dict[str, int]:
    return {"vma_count": 0, **{metric: 0 for metric in METRICS}}


def add_mapping(bucket: dict[str, int], mapping: dict[str, Any]) -> None:
    bucket["vma_count"] += 1
    for metric in METRICS:
        bucket[metric] += mapping[metric]


def aggregate(mappings: list[dict[str, Any]]) -> dict[str, Any]:
    total = empty_bucket()
    by_category: dict[str, dict[str, int]] = {}
    by_category_perms: dict[str, dict[str, dict[str, int]]] = {}
    for mapping in mappings:
        add_mapping(total, mapping)
        category = mapping["category"]
        by_category.setdefault(category, empty_bucket())
        add_mapping(by_category[category], mapping)
        perms = mapping["perms"]
        by_category_perms.setdefault(category, {}).setdefault(perms, empty_bucket())
        add_mapping(by_category_perms[category][perms], mapping)
    return {
        "vma_count": len(mappings),
        "total": total,
        "by_category": by_category,
        "by_category_perms": by_category_perms,
    }


def parse_smaps(text: str, target: str) -> tuple[dict[str, Any], list[dict[str, Any]]]:
    mappings: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for line in text.splitlines():
        match = HEADER.match(line)
        if match:
            current = {
                "start": int(match.group(1), 16),
                "end": int(match.group(2), 16),
                "perms": match.group(3),
                "offset": int(match.group(4), 16),
                "dev": match.group(5),
                "inode": int(match.group(6)),
                "path": (match.group(7) or "").strip(),
            }
            current["category"] = classify(current["path"], target)
            mappings.append(current)
            continue
        if current is None:
            continue
        metric = SMAPS_METRIC.match(line)
        if metric:
            current[metric.group(1)] = int(metric.group(2)) * 1024
    if not mappings:
        fail("smaps contained no mapping headers")
    for mapping in mappings:
        missing = [metric for metric in METRICS if metric not in mapping]
        if missing:
            fail(f"smaps mapping {mapping['path']!r} missing {missing}")
    return aggregate(mappings), mappings


def mapping_inventory(mappings: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    seen: dict[tuple[Any, ...], dict[str, Any]] = {}
    for mapping in mappings:
        key = (
            mapping["category"],
            mapping["path"],
            mapping["perms"],
            mapping["offset"],
            mapping["dev"],
            mapping["inode"],
        )
        seen.setdefault(
            key,
            {
                "category": mapping["category"],
                "path": mapping["path"],
                "perms": mapping["perms"],
                "offset": mapping["offset"],
                "dev": mapping["dev"],
                "inode": mapping["inode"],
            },
        )
    return sorted(seen.values(), key=lambda item: tuple(item[key] for key in ("category", "path", "perms", "offset", "dev", "inode")))


def maxima(aggregates: list[dict[str, Any]]) -> dict[str, Any]:
    if not aggregates:
        fail("cannot calculate maxima of an empty snapshot sequence")

    def bucket_max(buckets: list[dict[str, int]]) -> dict[str, int]:
        return {
            "vma_count": max(bucket["vma_count"] for bucket in buckets),
            **{metric: max(bucket[metric] for bucket in buckets) for metric in METRICS},
        }

    categories = sorted({category for item in aggregates for category in item["by_category"]})
    category_perms = {
        category: sorted({perms for item in aggregates for perms in item["by_category_perms"].get(category, {})})
        for category in categories
    }
    return {
        "vma_count": max(item["vma_count"] for item in aggregates),
        "total": bucket_max([item["total"] for item in aggregates]),
        "by_category": {
            category: bucket_max([item["by_category"].get(category, empty_bucket()) for item in aggregates])
            for category in categories
        },
        "by_category_perms": {
            category: {
                perms: bucket_max([
                    item["by_category_perms"].get(category, {}).get(perms, empty_bucket())
                    for item in aggregates
                ])
                for perms in category_perms[category]
            }
            for category in categories
        },
    }


def stats(values: list[int]) -> dict[str, Any]:
    if not values:
        fail("cannot summarize an empty value list")
    ordered = sorted(values)
    return {
        "count": len(ordered),
        "min": ordered[0],
        "median_lower": ordered[(len(ordered) - 1) // 2],
        "median_upper": ordered[len(ordered) // 2],
        "max": ordered[-1],
    }


def metric_stats(rows: list[dict[str, Any]], path: tuple[str, ...]) -> dict[str, Any]:
    values: list[int] = []
    for row in rows:
        value: Any = row
        missing = False
        for key in path:
            if not isinstance(value, dict):
                fail(f"summary path reached a non-object at {key!r}: {path}")
            if key not in value:
                missing = True
                break
            value = value[key]
        values.append(0 if missing else int(value))
    return stats(values)


def summarize_role(rows: list[dict[str, Any]]) -> dict[str, Any]:
    categories = sorted({category for row in rows for category in row["sampled_max"]["smaps"]["by_category"]})
    category_perms = {
        category: sorted({perms for row in rows for perms in row["sampled_max"]["smaps"]["by_category_perms"].get(category, {})})
        for category in categories
    }
    category_summary = {
        category: {
            metric: metric_stats(rows, ("sampled_max", "smaps", "by_category", category, metric))
            for metric in METRICS
        }
        for category in categories
    }
    category_perm_summary = {
        category: {
            perms: {
                metric: metric_stats(rows, ("sampled_max", "smaps", "by_category_perms", category, perms, metric))
                for metric in METRICS
            }
            for perms in category_perms[category]
        }
        for category in categories
    }
    final_category_summary = {
        category: {
            metric: metric_stats(rows, ("last_snapshot", "smaps", "by_category", category, metric))
            for metric in METRICS
        }
        for category in categories
    }
    return {
        "run_count": len(rows),
        "snapshot_count": sum(row["snapshot_count"] for row in rows),
        "gnu_time_peak_rss_kib": stats([row["gnu_time_peak_rss_kib"] for row in rows]),
        "sampled_status_max_bytes": {
            name: metric_stats(rows, ("sampled_max", "status", name))
            for name in STATUS_FIELDS
        },
        "sampled_smaps_max_bytes": {
            "total": {
                metric: metric_stats(rows, ("sampled_max", "smaps", "total", metric))
                for metric in METRICS
            },
            "by_category": category_summary,
            "by_category_perms": category_perm_summary,
        },
        "last_snapshot_smaps_bytes": {
            "total": {
                metric: metric_stats(rows, ("last_snapshot", "smaps", "total", metric))
                for metric in METRICS
            },
            "by_category": final_category_summary,
        },
    }


def check_receipt_files(folder: Path, receipt: dict[str, Any], packet: Path) -> dict[str, Any]:
    checked: dict[str, Any] = {}
    for name, expected in sorted(receipt.get("files", {}).items()):
        if not isinstance(expected, dict) or not isinstance(expected.get("path"), str):
            fail(f"invalid receipt file descriptor for {folder.name}/{name}")
        path = Path(expected["path"])
        actual = descriptor(path, packet)
        checked[name] = {
            **actual,
            "matches_receipt": all(actual.get(key) == expected.get(key) for key in ("bytes", "sha256")),
        }
        if not checked[name]["matches_receipt"]:
            fail(f"receipt hash mismatch for {path}")
    return checked


def parse_run(folder: Path, packet: Path, plan_sha256: str) -> dict[str, Any]:
    command = load_json(folder / "command.json")
    receipt = load_json(folder / "receipt.json")
    observations = load_json(folder / "observations.json")
    if command.get("exit_code") != 0 or receipt.get("status") != "pass":
        fail(f"run {folder.name} is not a passing child")
    if command.get("plan_sha256") != plan_sha256:
        fail(f"run {folder.name} command is bound to a different plan")
    row = command.get("row")
    if not isinstance(row, dict) or not isinstance(row.get("role"), str):
        fail(f"run {folder.name} has no row role")
    if receipt.get("row") != row:
        fail(f"run {folder.name} receipt row does not match command row")
    argv = command.get("argv")
    if not isinstance(argv, list) or "--case" not in argv:
        fail(f"run {folder.name} has no recognizable argv")
    case_index = argv.index("--case")
    if case_index == 0 or not isinstance(argv[case_index - 1], str):
        fail(f"run {folder.name} has no binary immediately before --case")
    target = os.path.realpath(argv[case_index - 1])
    raw_snapshots: list[dict[str, Any]] = []
    with gzip.open(folder / "snapshots.jsonl.gz", "rt", encoding="utf-8") as handle:
        for line_number, line in enumerate(handle):
            if not line.strip():
                continue
            item = json.loads(line)
            if item.get("index") != len(raw_snapshots):
                fail(f"run {folder.name} snapshot index is not contiguous at line {line_number}")
            if not isinstance(item.get("status"), str) or not isinstance(item.get("smaps"), str):
                fail(f"run {folder.name} snapshot {line_number} lacks raw text")
            raw_snapshots.append(item)
    metadata = observations.get("snapshots")
    if not isinstance(metadata, list) or len(metadata) != len(raw_snapshots):
        fail(f"run {folder.name} observation metadata count does not match raw lines")
    snapshots: list[dict[str, Any]] = []
    inventories: list[dict[str, Any]] = []
    for index, (item, meta) in enumerate(zip(raw_snapshots, metadata, strict=True)):
        for key in ("index", "pid", "exe", "monotonic_ns"):
            if item.get(key) != meta.get(key):
                fail(f"run {folder.name} metadata mismatch at snapshot {index}: {key}")
        smaps, mappings = parse_smaps(item["smaps"], target)
        status = parse_status(item["status"])
        snapshots.append({
            "index": item["index"],
            "monotonic_ns": item["monotonic_ns"],
            "pid": item["pid"],
            "exe": item["exe"],
            "status": status,
            "smaps": smaps,
        })
        inventories.extend(mappings)
    if not snapshots:
        fail(f"run {folder.name} has no snapshots")
    if receipt.get("snapshots") != len(snapshots):
        fail(f"run {folder.name} receipt snapshot count does not match raw lines")
    peak = receipt.get("peak_rss_kib")
    if type(peak) is not int or peak <= 0:
        fail(f"run {folder.name} has invalid GNU time peak")
    sampled_status = {
        key: max(snapshot["status"][key] for snapshot in snapshots)
        for key in STATUS_FIELDS
    }
    smaps_series = [snapshot["smaps"] for snapshot in snapshots]
    return {
        "folder": folder.relative_to(packet).as_posix(),
        "row": row,
        "binary": relpath(Path(target), packet),
        "binary_path_for_classification": target,
        "case": argv[argv.index("--case") + 1] if case_index + 1 < len(argv) else None,
        "gnu_time_peak_rss_kib": peak,
        "snapshot_count": len(snapshots),
        "proc_race_count": len(observations.get("proc_races", [])),
        "pids": sorted({snapshot["pid"] for snapshot in snapshots}),
        "exes": sorted({snapshot["exe"] for snapshot in snapshots}),
        "exe_matches_invoked_binary": all(
            os.path.realpath(snapshot["exe"]) == target for snapshot in snapshots
        ),
        "input_files": {
            **check_receipt_files(folder, receipt, packet),
            "receipt.json": descriptor(folder / "receipt.json", packet),
            "command.json": descriptor(folder / "command.json", packet),
            "observations.json": descriptor(folder / "observations.json", packet),
        },
        "mapping_inventory": mapping_inventory(inventories),
        "snapshots": snapshots,
        "sampled_max": {
            "status": sampled_status,
            "smaps": maxima(smaps_series),
        },
        "first_snapshot": snapshots[0],
        "last_snapshot": snapshots[-1],
    }


def analyse(packet: Path) -> dict[str, Any]:
    packet = packet.resolve()
    plan_path = packet / "plan.json"
    capture_path = packet / "capture.json"
    plan = load_json(plan_path)
    capture = load_json(capture_path)
    if capture.get("status") != "pass":
        fail(f"capture status is not pass: {capture.get('status')!r}")
    plan_digest = sha256(plan_path)
    runs_dir = packet / "runs"
    folders = sorted(path for path in runs_dir.iterdir() if path.is_dir())
    if not folders:
        fail("packet has no run directories")
    rows = [parse_run(folder, packet, plan_digest) for folder in folders]
    expected_children = capture.get("children")
    if expected_children != len(rows):
        fail(f"capture says {expected_children} children but found {len(rows)}")
    plan_rows = {
        item.get("ordinal"): item
        for item in plan.get("rows", [])
        if isinstance(item, dict) and "ordinal" in item
    }
    for row in rows:
        ordinal = row["row"].get("ordinal")
        if ordinal not in plan_rows or plan_rows[ordinal] != row["row"]:
            fail(f"run row is not the frozen plan row for ordinal {ordinal!r}")
    role_rows: dict[str, list[dict[str, Any]]] = {}
    for row in rows:
        role_rows.setdefault(row["row"]["role"], []).append(row)
    return {
        "schema": "litchi.live-mapping-analysis-0888.v1",
        "status": "pass",
        "scope": "Descriptive replay of every retained /proc status and smaps snapshot. No workload launch, timing comparison, causal attribution, or adoption claim.",
        "method": {
            "units": "bytes for procfs kB fields and smaps metrics; KiB for GNU time peak as emitted",
            "smaps_metrics": list(METRICS),
            "status_fields": list(STATUS_FIELDS),
            "classification_rules": {
                "executable_file": "smaps pathname, after removing a trailing ' (deleted)', equals the invoked binary realpath",
                "heap": "pathname is [heap]",
                "stack": "pathname starts with [stack",
                "anonymous": "pathname is empty or starts with [anon",
                "shared_library": "non-executable pathname starts with '/' or 'memfd:'",
                "special": "all remaining bracketed or named forms",
            },
            "aggregation": "Each snapshot sums every parsed VMA; sampled maxima take the maximum independently for each reported field across that child. Role summaries report lower and upper middle values without averaging.",
            "separation": "gnu_time_peak_rss_kib comes only from receipt peak-rss-kib.txt; sampled_max comes only from live status/smaps snapshots.",
        },
        "source": {
            "packet": relpath(packet, packet),
            "plan": descriptor(plan_path, packet),
            "capture": descriptor(capture_path, packet),
            "plan_schema": plan.get("schema"),
            "capture_status": capture.get("status"),
            "capture_children": expected_children,
            "rows_in_plan": plan.get("rows"),
        },
        "rows": rows,
        "role_summary": {
            role: summarize_role(role_rows[role])
            for role in sorted(role_rows)
        },
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    packet = args.packet.resolve()
    output = (args.output or (packet / "analysis.json.gz")).resolve()
    if output.exists():
        raise SystemExit(f"refusing to overwrite existing output: {output}")
    value = analyse(packet)
    output.parent.mkdir(parents=True, exist_ok=True)
    payload = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    if output.name.endswith(".gz"):
        with output.open("xb") as raw:
            with gzip.GzipFile(fileobj=raw, mode="wb", filename="", mtime=0, compresslevel=9) as handle:
                handle.write(payload)
    else:
        with output.open("xb") as handle:
            handle.write(payload)
    print(output)


if __name__ == "__main__":
    main()
