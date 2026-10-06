#!/usr/bin/env python3
"""Exercise the post-capture compact validation receipt against 42 fixtures.

This is an offline tooling control.  It uses the same helper called by the
capture runner, serializes the same child/state shape, and never starts a
benchmark process or modifies a retained packet.
"""

from __future__ import annotations

import argparse
import copy
import importlib.util
import json
import shutil
import tempfile
from pathlib import Path
from typing import Any


HERE = Path(__file__).resolve().parent
MAX_ARTIFACT_BYTES = 1 << 20


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def run(output: Path | None) -> dict[str, Any]:
    capture = load_module(HERE / "comparison-capture.py", "litchi_perf_0925_postcapture_capture")
    comparison = load_module(HERE / "comparison.py", "litchi_perf_0925_postcapture_comparison")
    capture_source = (HERE / "comparison-capture.py").read_text(encoding="utf-8")
    require("row_receipt[\"validation\"] = _compact_validation_receipt(" in capture_source,
            "capture runner does not call the compact validation helper")
    binary = {"path": "/tmp/0925-postcapture-synthetic-binary",
              "sha256": "a" * 64, "bytes": 123_456}
    root = Path(tempfile.mkdtemp(prefix="litchi-performance-0925-postcapture-", dir="/tmp"))
    try:
        packet = root / "packet"
        packet.mkdir()
        rows = comparison.schedule_rows()
        reports = {
            route: comparison.synthetic_report(route, binary=binary)
            for route in comparison.ROUTES
        }
        summaries: dict[str, dict[str, Any]] = {}
        children: list[dict[str, Any]] = []
        for row in rows:
            child = packet / "children" / str(row["id"])
            child.mkdir(parents=True)
            report_path = child / "report.json"
            report_path.write_text(json.dumps(reports[row["route"]], sort_keys=True) + "\n",
                                   encoding="utf-8")
            stdout_path = child / "stdout.log"
            stderr_path = child / "stderr.log"
            time_path = child / "time.log"
            stdout_path.write_bytes(b"")
            stderr_path.write_text("Maximum resident set size (kbytes): 1\n", encoding="utf-8")
            time_path.write_bytes(stderr_path.read_bytes())
            report_descriptor = capture.packet_descriptor(report_path, packet, "postcapture report")
            summary = comparison.validate_report(
                reports[row["route"]], row["route"], warmups=row["warmups"],
                samples=row["samples"], label=str(report_path), binary=binary
            )
            summaries[row["id"]] = summary
            validation = capture._compact_validation_receipt(
                summary, report_descriptor, comparison
            )
            require(set(validation) == {
                "schema", "status", "route", "samples", "warmups",
                "report_bytes", "report_sha256"
            }, f"{row['id']}: compact receipt gained raw fields")
            require(validation["status"] == "passed"
                    and validation["report_sha256"] == report_descriptor["sha256"]
                    and validation["report_bytes"] == report_descriptor["bytes"],
                    f"{row['id']}: compact report binding changed")
            stdout_descriptor = capture.packet_descriptor(stdout_path, packet, "postcapture stdout")
            stderr_descriptor = capture.packet_descriptor(stderr_path, packet, "postcapture stderr")
            time_descriptor = capture.packet_descriptor(time_path, packet, "postcapture time")
            children.append({
                "schema": "litchi.performance.0925.provider-child.v1",
                "schedule": row,
                "argv": capture._child_command(row, {"binary": binary}, report_path),
                "environment": {"LITCHI_DOCX_INGRESS_ROUTE": row["route"]},
                "cwd": "/tmp/0925-postcapture-synthetic-source",
                "started_utc": "synthetic",
                "ended_utc": "synthetic",
                "exit_code": 0,
                "observation_timeout_count": 0,
                "deadline_exceeded": False,
                "stdout": stdout_descriptor,
                "stderr_and_time": stderr_descriptor,
                "time": time_descriptor,
                "rss_bytes": 1024,
                "report": report_descriptor,
                "validation": validation,
                "status": "passed",
            })

        state = {
            "schema": "litchi.performance.0925.provider-capture-state.v1",
            "started_utc": "synthetic",
            "ended_utc": "synthetic",
            "children": children,
            "children_run": len(children),
            "process_deadline_seconds": 900.0,
            "status": "completed",
            "custody_revalidation_after": {"status": "verified"},
            "failure_policy": "invalid cohorts are retained and cannot be recaptured",
        }
        state_bytes = len(json.dumps(state, indent=2, sort_keys=True).encode())
        legacy_children = copy.deepcopy(children)
        for child in legacy_children:
            child["validation"] = summaries[child["schedule"]["id"]]
        legacy_state = {**state, "children": legacy_children}
        legacy_state_bytes = len(json.dumps(legacy_state, indent=2, sort_keys=True).encode())
        require(len(children) == 42 and state_bytes < MAX_ARTIFACT_BYTES,
                f"compact 42-child state exceeds one MiB: {state_bytes}")
        require(all("elapsed_ns" not in child["validation"]
                    and "phase_ns" not in child["validation"]
                    and "provider" not in child["validation"]
                    for child in children), "compact state retained raw vectors")

        result: dict[str, Any] = {
            "schema": "litchi.performance.0925.postcapture-fix-controls.v1",
            "status": "passed",
            "performance_children_run": 0,
            "claim_scope": "future runner custody control only; no repair, capture, timing adoption, or protocol qualification",
            "prior_capture_repaired": False,
            "tooling_bindings": {
                "capture": capture.descriptor(HERE / "comparison-capture.py",
                                                max_bytes=MAX_ARTIFACT_BYTES),
                "comparison": capture.descriptor(HERE / "comparison.py",
                                                   max_bytes=MAX_ARTIFACT_BYTES),
                "control_script": capture.descriptor(Path(__file__),
                                                       max_bytes=MAX_ARTIFACT_BYTES),
            },
            "schedule": {
                "children": len(children), "routes": list(comparison.ROUTES),
                "children_per_route": 6, "samples_per_child": 50,
            },
            "helper": {
                "name": "_compact_validation_receipt",
                "used_by_runner": True,
                "raw_vectors_persisted": "retained report identified by report_sha256",
                "scalar_fields": ["route", "samples", "warmups", "status",
                                  "report_bytes", "report_sha256"],
            },
            "state_size_control": {
                "status": "passed",
                "children": len(children),
                "compact_state_bytes": state_bytes,
                "legacy_full_vector_state_bytes": legacy_state_bytes,
                "writer_limit_bytes": MAX_ARTIFACT_BYTES,
                "compact_under_one_mib": state_bytes < MAX_ARTIFACT_BYTES,
                "legacy_full_vector_state_exceeds_one_mib": legacy_state_bytes >= MAX_ARTIFACT_BYTES,
                "same_helper_as_runner": True,
            },
            "negative_controls": {
                "raw_phase_vectors_absent_from_compact_receipts": "passed",
                "report_hash_present_for_all_children": "passed",
                "all_42_child_ids_unique": "passed",
            },
        }
        data = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode()
        require(len(data) < MAX_ARTIFACT_BYTES, "postcapture control receipt is at least 1 MiB")
        if output is not None:
            output = output.resolve()
            require(not output.exists(), f"refusing to replace retained receipt: {output}")
            output.parent.mkdir(parents=True, exist_ok=True)
            output.write_bytes(data)
        return result
    finally:
        shutil.rmtree(root, ignore_errors=True)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        result = run(args.output)
        print(json.dumps({
            "status": result["status"],
            "children": result["state_size_control"]["children"],
            "compact_state_bytes": result["state_size_control"]["compact_state_bytes"],
            "legacy_full_vector_state_bytes": result["state_size_control"]["legacy_full_vector_state_bytes"],
            "performance_children_run": result["performance_children_run"],
        }, sort_keys=True))
    except Exception as error:
        print(f"0925 postcapture controls failed: {type(error).__name__}: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
