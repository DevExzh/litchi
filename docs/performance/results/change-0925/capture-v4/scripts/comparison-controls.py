#!/usr/bin/env python3
"""Run 0925 validator and compact-reduction preflight controls.

The controls use deterministic in-memory reports shaped like the existing
``litchi-perf-baseline`` envelope.  They never build or launch a benchmark
child.  Every mutation is expected to be rejected by ``comparison.py``.
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


def load_validator() -> Any:
    spec = importlib.util.spec_from_file_location("litchi_perf_0925_comparison", HERE / "comparison.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load comparison.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_capture() -> Any:
    spec = importlib.util.spec_from_file_location("litchi_perf_0925_capture_controls", HERE / "comparison-capture.py")
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load comparison-capture.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_reducer() -> Any:
    spec = importlib.util.spec_from_file_location(
        "litchi_perf_0925_comparison_reducer", HERE / "comparison-reduce.py"
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("cannot load comparison-reduce.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RuntimeError(message)


def mutate(module: Any, report: dict[str, Any], name: str) -> dict[str, Any]:
    value = copy.deepcopy(report)
    result = value["results"][0]
    section = result["source"]["docx_section_layout"]
    if name == "phase_sum":
        section["query_ns"][0] += 1
    elif name == "duplicate_sample_index":
        result["elapsed_ns"]["sample_order"][0] = result["elapsed_ns"]["sample_order"][1]
        section["sample_order"] = result["elapsed_ns"]["sample_order"]
    elif name == "missing_sample_index":
        result["elapsed_ns"]["sample_order"].pop()
    elif name == "wrong_binary_identity":
        value["binary_identity"]["binary_sha256"] = "0" * 64
    elif name == "wrong_output_digest":
        section["output_sha256"][0] = "0" * 64
    elif name == "wrong_route":
        section["ingress_route"] = "range_delayed" if section["ingress_route"] != "range_delayed" else "owned_vec"
    elif name == "provider_histogram":
        require(section.get("provider") is not None, "provider histogram fixture requires provider route")
        section["provider"]["logical"][0]["request_size_counts"][0] += 1
    elif name == "provider_parameter":
        require(section.get("provider") is not None, "provider parameter fixture requires provider route")
        section["provider"]["max_returned_bytes"] = 1
    elif name == "provider_vector_cardinality":
        require(section.get("provider") is not None, "provider vector fixture requires provider route")
        section["provider"]["underlying"].pop()
    elif name == "provider_nominal_formula":
        require(section.get("provider") is not None, "provider delay fixture requires provider route")
        section["provider"]["nominal_service_delay_ns"][0] += 1
    else:
        raise ValueError(f"unknown mutation {name}")
    return value


def expect_rejection(module: Any, report: dict[str, Any], route: str, name: str,
                     binary: dict[str, Any]) -> dict[str, str]:
    try:
        module.validate_report(report, route, label=name, binary=binary)
    except module.EvidenceError as error:
        return {"name": name, "status": "rejected", "reason": str(error)}
    raise RuntimeError(f"negative control unexpectedly passed: {name}")


def copied_packet_replay(module: Any, reports: dict[str, dict[str, Any]],
                         binary: dict[str, Any]) -> dict[str, Any]:
    """Exercise moved-packet replay and report-hash rejection in scratch."""
    capture = load_capture()
    root = Path(tempfile.mkdtemp(prefix="litchi-performance-0925-replay-",
                                 dir="/tmp"))
    try:
        packet = root / "packet"
        packet.mkdir()
        (packet / "scripts").mkdir()
        scripts = {}
        packet_scripts = {}
        for key, name in (
            ("comparison", "comparison.py"),
            ("controls", "comparison-controls.py"),
            ("capture", "comparison-capture.py"),
            ("reduce", "comparison-reduce.py"),
        ):
            source = HERE / name
            target = packet / "scripts" / name
            shutil.copy2(source, target)
            digest = capture.descriptor(source)
            scripts[key] = digest
            packet_scripts[key] = {
                "path": f"scripts/{name}", "bytes": target.stat().st_size,
                "sha256": capture.sha256_path(target),
            }
        rows = module.schedule_rows()
        children = []
        for row in rows:
            child = packet / "children" / str(row["id"])
            child.mkdir(parents=True)
            stdout_path = child / "stdout.log"
            stderr_path = child / "stderr.log"
            time_path = child / "time.log"
            stdout_path.write_bytes(b"")
            stderr_path.write_text("Maximum resident set size (kbytes): 1\n", encoding="utf-8")
            time_path.write_bytes(stderr_path.read_bytes())
            report_path = child / "report.json"
            report_path.write_text(json.dumps(reports[row["route"]], sort_keys=True) + "\n",
                                   encoding="utf-8")
            children.append({
                "schedule": row,
                "status": "passed",
                "argv": capture._child_command(row, {
                    "binary": binary,
                }, report_path),
                "environment": {"LITCHI_DOCX_INGRESS_ROUTE": row["route"]},
                "cwd": "/tmp/0925-synthetic-source",
                "exit_code": 0,
                "deadline_exceeded": False,
                "stdout": capture.packet_descriptor(stdout_path, packet, "synthetic stdout"),
                "stderr_and_time": capture.packet_descriptor(stderr_path, packet,
                                                               "synthetic stderr"),
                "time": capture.packet_descriptor(time_path, packet, "synthetic time"),
                "rss_bytes": 1024,
                "report": capture.packet_descriptor(report_path, packet, "synthetic report"),
            })
        custody_root = packet / "custody"
        custody_root.mkdir()
        custody = {}
        custody_original = {}
        for name in ("source_receipt", "build_receipt", "protocol"):
            path = custody_root / f"{name}.json"
            if name == "source_receipt":
                value = {"base_revision": capture.BASE_REVISION,
                         "source_files": ["synthetic/source.rs"]}
                path.write_text(json.dumps(value, sort_keys=True) + "\n", encoding="utf-8")
            elif name == "build_receipt":
                value = {
                    "base_revision": capture.BASE_REVISION,
                    "profile": "release",
                    "command": "cargo +1.95.0 build --offline -j2 --locked --release",
                    "toolchain": {"rustc": "rustc 1.95.0 synthetic"},
                    "locks": {"synthetic": "bound"},
                    "binary": binary,
                    "source_custody": {
                        "source_receipt": custody_original["source_receipt"],
                        "full_source_listing": custody_original["source_receipt"],
                    },
                }
                path.write_text(json.dumps(value, sort_keys=True) + "\n", encoding="utf-8")
            else:
                shutil.copy2(HERE / "comparison-protocol-v1.json", path)
            local = capture.packet_descriptor(path, packet, f"synthetic {name}")
            custody[name] = local
            custody_original[name] = capture.descriptor(path)
        manifest = {
            "schema": "litchi.performance.0925.provider-capture-manifest.v1",
            "capture_authorized": False,
            "base_revision": capture.BASE_REVISION,
            "binary": binary,
            "source": {
                "source_root": "/tmp/0925-synthetic-source",
                "source_receipt": custody_original["source_receipt"],
                "full_source_listing": custody_original["source_receipt"],
            },
            "build": {"receipt": custody_original["build_receipt"], "binary": binary},
            "protocol": custody_original["protocol"],
            "packet_custody": custody,
            "scripts": scripts,
            "packet_scripts": packet_scripts,
            "process_deadline_seconds": 900.0,
            "schedule": {"rows": rows,
                          "sha256": __import__("hashlib").sha256(capture.canonical(rows)).hexdigest()},
            "selector": {"name": "LITCHI_DOCX_INGRESS_ROUTE", "values": list(module.ROUTES)},
        }
        (packet / "manifest.json").write_text(json.dumps(manifest, indent=2, sort_keys=True) + "\n",
                                                encoding="utf-8")
        (packet / "capture-state.json").write_text(json.dumps({
            "schema": "litchi.performance.0925.provider-capture-state.v1",
            "children": children,
            "children_run": 42,
            "process_deadline_seconds": 900.0,
            "status": "completed",
            "custody_revalidation_after": {"status": "verified"},
        }, sort_keys=True) + "\n",
                                                     encoding="utf-8")
        moved = root / "moved-packet"
        shutil.copytree(packet, moved)
        replay = capture.replay(argparse.Namespace(packet=moved))
        reducer = load_reducer()
        reduced = reducer.reduce_packet(moved)
        reduced_bytes = len(json.dumps(reduced, separators=(",", ":"), sort_keys=True).encode())
        require(reduced_bytes < MAX_ARTIFACT_BYTES,
                f"offline reducer synthetic output is at least 1 MiB: {reduced_bytes}")

        def expect_replay_rejection(name: str, mutate: Any) -> dict[str, str]:
            target = root / name
            shutil.copytree(packet, target)
            mutate(target)
            try:
                capture.replay(argparse.Namespace(packet=target))
            except capture.CaptureError as error:
                return {"status": "rejected", "reason": str(error)}
            raise RuntimeError(f"copied replay accepted {name}")

        negative = {
            "corrupt_report": expect_replay_rejection(
                "corrupt-report",
                lambda target: (target / "children" / str(rows[0]["id"]) / "report.json")
                .write_text(
                    (target / "children" / str(rows[0]["id"]) / "report.json")
                    .read_text(encoding="utf-8") + "\n", encoding="utf-8"
                ),
            ),
            "corrupt_stdout": expect_replay_rejection(
                "corrupt-stdout",
                lambda target: (target / "children" / str(rows[0]["id"]) / "stdout.log")
                .write_bytes(b"tampered"),
            ),
            "incomplete_state": expect_replay_rejection(
                "incomplete-state",
                lambda target: (target / "capture-state.json").write_text(
                    json.dumps({"status": "invalid"}) + "\n", encoding="utf-8"
                ),
            ),
            "missing_child": expect_replay_rejection(
                "missing-child",
                lambda target: (target / "capture-state.json").write_text(
                    json.dumps({
                        **json.loads((target / "capture-state.json").read_text(encoding="utf-8")),
                        "children": json.loads(
                            (target / "capture-state.json").read_text(encoding="utf-8")
                        )["children"][:-1],
                        "children_run": 41,
                    }, sort_keys=True) + "\n", encoding="utf-8"
                ),
            ),
            "escaped_report_path": expect_replay_rejection(
                "escaped-report",
                lambda target: (lambda state: (
                    state["children"][0]["report"].update({"path": "../report.json"}),
                    (target / "capture-state.json").write_text(
                        json.dumps(state, sort_keys=True) + "\n", encoding="utf-8"
                    )
                ))(json.loads((target / "capture-state.json").read_text(encoding="utf-8"))),
            ),
            "tampered_local_custody": expect_replay_rejection(
                "tampered-custody",
                lambda target: (target / "custody" / "protocol.json").write_text(
                    "{}\n", encoding="utf-8"
                ),
            ),
            "tampered_packet_script_descriptor": expect_replay_rejection(
                "tampered-script-descriptor",
                lambda target: (lambda manifest: (
                    manifest["packet_scripts"]["comparison"].update({"sha256": "0" * 64}),
                    (target / "manifest.json").write_text(
                        json.dumps(manifest, sort_keys=True) + "\n", encoding="utf-8"
                    )
                ))(json.loads((target / "manifest.json").read_text(encoding="utf-8"))),
            ),
        }
        return {"moved_packet_reports": replay["reports"],
                "moved_packet_status": replay["status"],
                "offline_reducer": {
                    "status": reduced["status"],
                    "bytes": reduced_bytes,
                    "routes": len(reduced["routes"]),
                    "children": sum(row["children"] for row in reduced["routes"].values()),
                    "rss_fields": all("child_rss_bytes" in row
                                       for row in reduced["routes"].values()),
                    "performance_children_run": reduced["performance_children_run"],
                },
                "negative_controls": negative}
    finally:
        shutil.rmtree(root, ignore_errors=True)


def run(output: Path | None) -> dict[str, Any]:
    module = load_validator()
    capture = load_capture()
    rows = module.schedule_rows()
    binary = {"path": "/tmp/0925-synthetic-binary", "sha256": "a" * 64, "bytes": 123_456}
    positives = []
    reports: dict[str, dict[str, Any]] = {}
    for route in module.ROUTES:
        report = module.synthetic_report(route, binary=binary)
        summary = module.validate_report(report, route, label=f"positive-{route}", binary=binary)
        reports[route] = report
        positives.append({"route": route, "samples": summary["samples"], "status": "accepted"})

    controls: list[dict[str, str]] = []
    for name in ("phase_sum", "duplicate_sample_index", "missing_sample_index",
                 "wrong_binary_identity", "wrong_output_digest", "wrong_route"):
        controls.append(expect_rejection(module, mutate(module, reports["generic_read_at"], name),
                                         "generic_read_at", name, binary))
    for route in module.PROVIDER_ROUTES:
        for name in ("provider_histogram", "provider_parameter", "provider_vector_cardinality",
                     "provider_nominal_formula"):
            controls.append(expect_rejection(module, mutate(module, reports[route], name),
                                             route, f"{route}:{name}", binary))

    replay = copied_packet_replay(module, reports, binary)

    reduction_inputs = {
        route: [
            module.validate_report(reports[route], route, label=f"reduction-{route}-{index}",
                                   binary=binary)
            for index in range(6)
        ]
        for route in module.ROUTES
    }
    reduction = module.reduce_reports(reduction_inputs)
    reduction_bytes = len(json.dumps(reduction, separators=(",", ":"), sort_keys=True).encode())
    require(reduction_bytes < MAX_ARTIFACT_BYTES, "reduction smoke receipt is at least 1 MiB")
    for route in module.ROUTES:
        row = reduction["routes"][route]
        require(all(key in row["child_latency_ns"][0]
                    for key in ("p50", "p95", "p99", "spread")),
                f"{route}: child tail reduction fields missing")
        require((route in module.PROVIDER_ROUTES) == (row["provider_accounting"] is not None),
                f"{route}: provider accounting scope changed")

    planned_bytes, planned_sha = module.planned_reduction_bytes()
    require(planned_bytes < MAX_ARTIFACT_BYTES,
            f"full planned reduction is at least 1 MiB: {planned_bytes}")
    result: dict[str, Any] = {
        "schema": "litchi.performance.0925.comparison-controls.v2",
        "status": "passed",
        "performance_children_run": 0,
        "tooling_bindings": {
            "scripts": capture._script_descriptors(),
            "protocol": capture.descriptor(HERE / "comparison-protocol-v1.json",
                                             max_bytes=MAX_ARTIFACT_BYTES),
        },
        "schedule": {"children": len(rows), "routes": list(module.ROUTES),
                      "children_per_route": 6, "warmups": 5, "samples_per_child": 50,
                      "measured_samples": len(rows) * 50},
        "positive_reports": positives,
        "negative_controls": controls,
        "copied_packet_replay": replay,
        "reduction_contract": {
            "status": "passed",
            "bytes": reduction_bytes,
            "under_one_mib": True,
            "provider_accounting_routes": list(module.PROVIDER_ROUTES),
            "child_tail_fields": ["p50", "p95", "p99", "spread"],
        },
        "planned_reduction": {"bytes": planned_bytes, "sha256": planned_sha,
                              "under_one_mib": True, "children": 42, "samples_per_child": 50},
        "claim_scope": "validator/preflight only; no provider ranking or optimization comparison",
    }
    data = (json.dumps(result, indent=2, sort_keys=True) + "\n").encode()
    require(len(data) < MAX_ARTIFACT_BYTES, "control receipt is at least 1 MiB")
    if output is not None:
        output = output.resolve()
        require(not output.exists(), f"refusing to replace retained receipt: {output}")
        output.parent.mkdir(parents=True, exist_ok=True)
        output.write_bytes(data)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    try:
        print(json.dumps(run(args.output), sort_keys=True))
    except Exception as error:
        print(f"0925 comparison controls failed: {error}")
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
