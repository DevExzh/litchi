#!/usr/bin/env python3
"""Reduce raw latency and phase samples from the frozen 0923 packet.

The frozen comparison reducer's paired p50 gates remain authoritative.  This
supplement computes descriptive tails from every raw sample (1,600 samples per
route and version), rather than from the 8 child summaries per cell.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


HERE = Path(__file__).resolve().parent
CAPTURE_PATH = HERE / "comparison-capture-v2.py"
COMPARISON_PATH = HERE / "comparison-v2.py"
ROUTES = ("generic_read_at", "owned_vec", "owned_reader")
VARIANTS = ("baseline", "candidate")
PHASES = ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")
SAMPLES = 200
CHILDREN_PER_ROUTE_VERSION = 8
EXPECTED_RAW_SAMPLES = SAMPLES * CHILDREN_PER_ROUTE_VERSION
ONE_MIB = 1 << 20


class PooledEvidenceError(RuntimeError):
    """A fail-closed raw-sample reduction error."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise PooledEvidenceError(message)


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def nearest_rank(values: list[int], fraction: float) -> float:
    require(values, "cannot calculate a percentile of an empty raw vector")
    ordered = sorted(values)
    index = min(max(int((len(ordered) * fraction) + 0.999999999) - 1, 0), len(ordered) - 1)
    return float(ordered[index])


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file(), f"missing raw report: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": digest.hexdigest()}


def reduce_raw(packet_root: Path, *, replay: bool = False) -> dict[str, Any]:
    capture = load_module(CAPTURE_PATH, "litchi_perf_0923_capture_raw_pool")
    comparison = load_module(COMPARISON_PATH, "litchi_perf_0923_comparison_raw_pool")
    root = packet_root.resolve()
    manifest = capture.load_json(root / "comparison-manifest-v2.json")
    require(manifest.get("status") == "frozen", "raw pooling requires a frozen manifest")
    capture.verify_manifest(manifest, external=not replay)
    rows = manifest["native"]["schedule"]
    require(len(rows) == 48, f"raw pooling schedule changed: {len(rows)}")

    values: dict[str, dict[str, dict[str, Any]]] = {
        route: {
            variant: {
                "elapsed_ns": [],
                "phase_ns": {phase: [] for phase in PHASES},
                "reports": [],
                "children": [],
            }
            for variant in VARIANTS
        }
        for route in ROUTES
    }
    for row in rows:
        receipt, report = capture._read_child(root, row, external=not replay)
        binary = manifest["binaries"][row["variant"]]["binary"]
        validated = comparison.validate_report(
            report,
            row["route"],
            warmups=manifest["native"]["warmups_per_child"],
            samples=manifest["native"]["samples_per_child"],
            label=row["id"],
            binary=binary,
        )
        cell = values[row["route"]][row["variant"]]
        cell["elapsed_ns"].extend(validated["elapsed_ns"])
        for phase in PHASES:
            cell["phase_ns"][phase].extend(validated["phase_ns"][phase])
        cell["reports"].append(descriptor(Path(receipt["report"]["path"])))
        cell["children"].append(row["id"])

    routes: dict[str, Any] = {}
    for route in ROUTES:
        routes[route] = {}
        for variant in VARIANTS:
            cell = values[route][variant]
            require(len(cell["children"]) == CHILDREN_PER_ROUTE_VERSION,
                    f"{route}/{variant}: child count changed")
            require(len(cell["elapsed_ns"]) == EXPECTED_RAW_SAMPLES,
                    f"{route}/{variant}: raw elapsed sample count changed")
            phase_tails = {}
            for phase in PHASES:
                require(len(cell["phase_ns"][phase]) == EXPECTED_RAW_SAMPLES,
                        f"{route}/{variant}/{phase}: raw phase sample count changed")
                phase_tails[phase] = {
                    "p50_ns": nearest_rank(cell["phase_ns"][phase], 0.50),
                    "p95_ns": nearest_rank(cell["phase_ns"][phase], 0.95),
                    "p99_ns": nearest_rank(cell["phase_ns"][phase], 0.99),
                }
            routes[route][variant] = {
                "children": len(cell["children"]),
                "raw_sample_count": len(cell["elapsed_ns"]),
                "elapsed_tails": {
                    "p50_ns": nearest_rank(cell["elapsed_ns"], 0.50),
                    "p95_ns": nearest_rank(cell["elapsed_ns"], 0.95),
                    "p99_ns": nearest_rank(cell["elapsed_ns"], 0.99),
                },
                "phase_raw_sample_count": {phase: len(cell["phase_ns"][phase]) for phase in PHASES},
                "phase_tails": phase_tails,
                "raw_reports": cell["reports"],
                "child_ids": cell["children"],
            }
    return {
        "schema": "litchi.performance.0923.raw-pooled-reduction.v1",
        "created_utc": datetime.now(timezone.utc).isoformat(),
        "manifest": descriptor(root / "comparison-manifest-v2.json"),
        "routes": list(ROUTES),
        "variants": list(VARIANTS),
        "samples_per_child": SAMPLES,
        "children_per_route_version": CHILDREN_PER_ROUTE_VERSION,
        "raw_samples_per_route_version": EXPECTED_RAW_SAMPLES,
        "quantile_method": "nearest-rank over all raw samples; p50/p95/p99",
        "routes_raw": routes,
        "frozen_reducer_note": "paired p50 gates are from comparison-reduction-v2; its pooled_latency_review and phase_summaries are child-summary distributions, not these raw-sample tails",
        "performance_claim": "descriptive timing and phase evidence only; no production adoption",
    }


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet-root", type=Path,
                        default=Path(os.environ.get("LITCHI_PERF_0923_PACKET_ROOT", "/tmp/litchi-performance-0923/packet-v3-final2")))
    parser.add_argument("--output", type=Path)
    parser.add_argument("--replay", action="store_true",
                        help="read packet-local copied reports and receipts")
    args = parser.parse_args(list(argv) if argv is not None else None)
    try:
        value = reduce_raw(args.packet_root, replay=args.replay)
        data = json.dumps(value, indent=2, sort_keys=True) + "\n"
        require(len(data.encode()) < ONE_MIB, "raw pooled result is at least 1 MiB")
        if args.output:
            output = args.output.resolve()
            require(not output.exists(), f"refusing to replace retained result: {output}")
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_text(data)
            print(json.dumps(descriptor(output), sort_keys=True))
        else:
            print(data, end="")
        return 0
    except Exception as error:
        print(f"0923 raw pooled reduction failed: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
