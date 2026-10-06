#!/usr/bin/env python3
"""Freeze, capture, reduce, copy, and replay the 0923 Heaptrack lane.

The lane is whole-process allocation evidence.  It is deliberately separate
from the native timing runner: a frozen native comparison manifest supplies
the exact binaries and source/build identity, while every raw Heaptrack
profile remains outside the checkout.  ``capture`` cannot start without both
that manifest and an explicit approval environment variable.  A failed batch
leaves a state file and cannot be selectively recaptured.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import shutil
import statistics
import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable


REPO_ROOT = Path(__file__).resolve().parents[4]
HERE = Path(__file__).resolve().parent
PACKET_ROOT = Path(os.environ.get(
    "LITCHI_PERF_0923_ALLOC_PACKET_ROOT",
    "/tmp/litchi-performance-0923/allocation-packet-v2",
)).resolve()
PROTOCOL_PATH = HERE / "allocation-protocol-v2.json"
PARSER_PATH = HERE / "allocation-parse-v2.py"
CONTROLS_PATH = HERE / "allocation-parse-controls-v2.py"
CAPTURE_PATH = Path(__file__).resolve()
DEFAULT_NATIVE_MANIFEST = Path(os.environ.get(
    "LITCHI_PERF_0923_NATIVE_MANIFEST",
    "/tmp/litchi-performance-0923/freeze-manifest-v2.json",
)).resolve()
CASE = "docx_source_backed_existing_section_layout_edit_save"
ROUTES = ("generic_read_at", "owned_vec", "owned_reader")
PAIR_ORDERS = ("baseline,candidate", "candidate,baseline") * 2
CPU = 12
WARMUPS = 0
SAMPLES = 1
CHILDREN = 24
ONE_MIB = 1 << 20
HEAPTRACK = Path("/usr/bin/heaptrack")
HEAPTRACK_PRINT = Path("/usr/bin/heaptrack_print")
ZSTD = Path("/usr/bin/zstd")
ZSTDCAT = Path("/usr/bin/zstdcat")


class AllocationError(RuntimeError):
    """A fail-closed freeze, capture, reduction, or replay error."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AllocationError(message)


def utc_now() -> str:
    return datetime.now(timezone.utc).isoformat()


def canonical(value: Any) -> bytes:
    return json.dumps(value, sort_keys=True, separators=(",", ":")).encode("utf-8")


def sha256_path(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def ensure_external(path: Path) -> Path:
    path = path.resolve()
    try:
        path.relative_to(REPO_ROOT)
    except ValueError:
        return path
    raise AllocationError(f"evidence path must be outside the checkout: {path}")


def descriptor(path: Path, *, limit: int | None = None) -> dict[str, Any]:
    path = path.resolve()
    require(path.is_file(), f"missing evidence file: {path}")
    size = path.stat().st_size
    if limit is not None:
        require(size < limit, f"retained packet file is at least 1 MiB: {path} ({size} bytes)")
    return {"path": str(path), "bytes": size, "sha256": sha256_path(path)}


def write_json(path: Path, value: Any, *, replace: bool = False, limit: int = ONE_MIB) -> dict[str, Any]:
    path = ensure_external(path)
    data = json.dumps(value, indent=2, sort_keys=True).encode("utf-8") + b"\n"
    require(len(data) < limit, f"JSON evidence is at least 1 MiB: {path}")
    if path.exists() and not replace:
        raise AllocationError(f"refusing to replace retained evidence: {path}")
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return descriptor(path, limit=limit)


def load_json(path: Path) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as error:
        raise AllocationError(f"cannot read JSON {path}: {error}") from error
    require(isinstance(value, dict), f"expected JSON object: {path}")
    return value


def embedded_descriptor(value: Any, label: str, *, check_path: bool = True) -> dict[str, Any]:
    require(isinstance(value, dict), f"{label}: descriptor missing")
    require(isinstance(value.get("path"), str), f"{label}: descriptor path missing")
    require(type(value.get("bytes")) is int and value["bytes"] >= 0,
            f"{label}: descriptor byte count invalid")
    digest = value.get("sha256")
    require(isinstance(digest, str) and re.fullmatch(r"[0-9a-f]{64}", digest) is not None,
            f"{label}: descriptor hash invalid")
    if check_path:
        actual = descriptor(Path(value["path"]))
        require(actual["bytes"] == value["bytes"] and actual["sha256"] == value["sha256"],
                f"{label}: bound file changed")
    return value


def schedule() -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for route in ROUTES:
        for pair, order in enumerate(PAIR_ORDERS, start=1):
            for leg, variant in enumerate(order.split(","), start=1):
                rows.append({
                    "index": len(rows) + 1,
                    "route": route,
                    "pair": pair,
                    "order": order,
                    "leg": leg,
                    "variant": variant,
                    "id": f"{route}-pair-{pair:02d}-leg-{leg}-{variant}",
                })
    require(len(rows) == CHILDREN, f"allocation schedule cardinality changed: {len(rows)}")
    require(sum(row["variant"] == "baseline" for row in rows) == 12,
            "allocation baseline cardinality changed")
    require(sum(row["variant"] == "candidate" for row in rows) == 12,
            "allocation candidate cardinality changed")
    return rows


def schedule_descriptor() -> dict[str, Any]:
    rows = schedule()
    return {
        "routes": list(ROUTES),
        "pairs_per_route": 4,
        "pair_orders": list(PAIR_ORDERS),
        "children_per_route": 8,
        "children_per_variant": 12,
        "children_total": CHILDREN,
        "rows": rows,
        "sha256": hashlib.sha256(canonical(rows)).hexdigest(),
    }


def tool_descriptor(path: Path, label: str) -> dict[str, Any]:
    require(path.is_file(), f"{label} is unavailable: {path}")
    value = descriptor(path, limit=ONE_MIB)
    version = subprocess.run([str(path), "--version"], stdout=subprocess.PIPE,
                             stderr=subprocess.STDOUT, text=True, check=False).stdout.strip()
    value["version_output"] = version
    return value


def source_binding(native: dict[str, Any], variant: str) -> dict[str, Any]:
    source = native.get(f"{variant}_source")
    require(isinstance(source, dict), f"native manifest missing {variant} source binding")
    result: dict[str, Any] = {"revision": source.get("revision") or source.get("candidate_base_revision")}
    for key in ("receipt", "archive", "full_source_listing", "harness_patch", "cargo_lock", "probe_lock",
        "production_patch"):
        value = source.get(key)
        if value is not None:
            require(isinstance(value, dict), f"{variant}.source.{key}: descriptor is not an object")
            ensure_external(Path(value.get("path", "")))
            result[key] = embedded_descriptor(value, f"{variant}.source.{key}")
    require(isinstance(result.get("receipt"), dict),
            f"native manifest missing {variant} source receipt")
    return result


def native_binary_binding(native: dict[str, Any], variant: str) -> dict[str, Any]:
    binaries = native.get("binaries")
    require(isinstance(binaries, dict), "native manifest has no binary bindings")
    binding = binaries.get(variant)
    require(isinstance(binding, dict), f"native manifest missing {variant} binary binding")
    binary_claim = binding.get("binary")
    require(isinstance(binary_claim, dict), f"native manifest missing {variant} binary descriptor")
    ensure_external(Path(binary_claim.get("path", "")))
    binary = embedded_descriptor(binary_claim, f"{variant}.binary")
    build = binding.get("build_receipt")
    require(build is not None, f"native manifest missing {variant} build receipt")
    require(isinstance(build, dict), f"native manifest has invalid {variant} build receipt")
    ensure_external(Path(build.get("path", "")))
    build_descriptor = embedded_descriptor(build, f"{variant}.build_receipt")
    record = binding.get("build_record")
    require(isinstance(record, dict), f"native manifest missing {variant} build record")
    source_receipt = record.get("source_receipt")
    require(isinstance(source_receipt, dict), f"{variant} build record is not source-bound")
    require(source_receipt.get("sha256") == native.get(f"{variant}_source", {})
            .get("receipt", {}).get("sha256"),
            f"{variant} build/source receipt binding differs")
    return {
        "kind": variant,
        "binary": binary,
        "build_receipt": build_descriptor,
        "build_record": {
            "argv": record.get("argv") or record.get("regeneration_command"),
            "profile": record.get("profile"),
            "rustc_vv": record.get("rustc_vv"),
            "source_receipt": source_receipt,
            "source_revision": record.get("source_revision"),
            "features": record.get("features"),
        },
    }


def verify_native_manifest(path: Path) -> dict[str, Any]:
    path = ensure_external(path)
    native = load_json(path)
    require(native.get("schema") == "litchi.performance.0923.comparison-manifest.v2",
            "native manifest schema is not comparison-manifest.v2")
    require(native.get("status") == "frozen",
            "allocation capture requires a complete frozen native manifest")
    require(native.get("case") == CASE, "native manifest case differs")
    require(native.get("native", {}).get("cpu") == CPU, "native manifest CPU differs")
    require(native.get("candidate_source") is not None and native.get("baseline_source") is not None,
            "native manifest must contain both source bindings")
    if native.get("allocation_evidence") is not None:
        require(native["allocation_evidence"].get("measured") is False,
                "native manifest already claims allocation evidence")
    binaries = {
        variant: native_binary_binding(native, variant)
        for variant in ("baseline", "candidate")
    }
    sources = {
        variant: source_binding(native, variant)
        for variant in ("baseline", "candidate")
    }
    return {
        "path": str(path),
        "descriptor": descriptor(path, limit=ONE_MIB),
        "case": native["case"],
        "cpu": CPU,
        "binaries": binaries,
        "sources": sources,
        "native_created_utc": native.get("created_utc"),
    }


def build_allocation_manifest(native_path: Path) -> dict[str, Any]:
    native = verify_native_manifest(native_path)
    return {
        "schema": "litchi.performance.0923.heaptrack-allocation.manifest.v2",
        "status": "frozen",
        "created_utc": utc_now(),
        "case": CASE,
        "scope": "whole_process",
        "claims": [
            "allocation_event_count",
            "requested_bytes_sum",
            "peak_live_heap_bytes",
            "live_bytes_at_end",
        ],
        "claim_limits": [
            "Heaptrack overhead is included",
            "profile covers setup, correctness/report serialization, and teardown",
            "no operation-scoped allocation or native latency claim",
            "no production adoption decision",
        ],
        "native_manifest": native["descriptor"],
        "native_identity": {
            "case": native["case"],
            "cpu": native["cpu"],
            "created_utc": native["native_created_utc"],
            "binaries": native["binaries"],
            "sources": native["sources"],
        },
        "protocol": descriptor(PROTOCOL_PATH, limit=ONE_MIB),
        "parser": descriptor(PARSER_PATH, limit=ONE_MIB),
        "controls": descriptor(CONTROLS_PATH, limit=ONE_MIB),
        "capture_runner": descriptor(CAPTURE_PATH, limit=ONE_MIB),
        "tooling": {
            "heaptrack": tool_descriptor(HEAPTRACK, "heaptrack"),
            "heaptrack_print": tool_descriptor(HEAPTRACK_PRINT, "heaptrack_print"),
            "zstd": tool_descriptor(ZSTD, "zstd"),
            "zstdcat": tool_descriptor(ZSTDCAT, "zstdcat"),
        },
        "schedule": schedule_descriptor(),
        "timed_command_template": [
            "/usr/bin/time", "-v", "-o", "$TIME_LOG", "taskset", "-c", str(CPU),
            "env", "LC_ALL=C", "LITCHI_DOCX_INGRESS_ROUTE=$ROUTE", "heaptrack",
            "--record-only", "-o", "$RAW_PROFILE", "$BINARY", "--case", CASE,
            "--samples", "1", "--warmup", "0", "--json", "$REPORT",
        ],
        "analysis_command_template": [
            "/usr/bin/heaptrack_print", "-f", "$RAW_PROFILE", "-p", "0", "-T", "0",
            "-a", "0", "-l", "0", "-H", "$HISTOGRAM",
        ],
    }


def verify_allocation_manifest(manifest: dict[str, Any], *, check_native: bool = True) -> None:
    require(manifest.get("schema") == "litchi.performance.0923.heaptrack-allocation.manifest.v2",
            "allocation manifest schema changed")
    require(manifest.get("status") == "frozen", "allocation manifest is not frozen")
    require(manifest.get("case") == CASE, "allocation manifest case changed")
    schedule_info = manifest.get("schedule")
    require(isinstance(schedule_info, dict) and schedule_info.get("rows") == schedule()[0:]
            and schedule_info.get("sha256") == schedule_descriptor()["sha256"],
            "allocation schedule changed")
    for key in ("native_manifest", "protocol", "parser", "controls", "capture_runner"):
        embedded_descriptor(manifest.get(key), key, check_path=check_native)
    tooling = manifest.get("tooling")
    require(isinstance(tooling, dict), "allocation tooling bindings missing")
    for key in ("heaptrack", "heaptrack_print", "zstd", "zstdcat"):
        embedded_descriptor(tooling.get(key), f"tooling.{key}", check_path=check_native)
    identities = manifest.get("native_identity")
    require(isinstance(identities, dict), "native identity binding missing")
    ensure_external(Path(manifest["native_manifest"]["path"]))
    for variant in ("baseline", "candidate"):
        binding = identities.get("binaries", {}).get(variant)
        require(isinstance(binding, dict), f"native identity missing {variant}")
        binary_claim = binding.get("binary")
        build_claim = binding.get("build_receipt")
        require(isinstance(binary_claim, dict), f"native identity missing {variant} binary descriptor")
        require(isinstance(build_claim, dict), f"native identity missing {variant} build descriptor")
        ensure_external(Path(binary_claim.get("path", "")))
        ensure_external(Path(build_claim.get("path", "")))
        embedded_descriptor(binding.get("binary"), f"native.{variant}.binary", check_path=check_native)
        embedded_descriptor(binding.get("build_receipt"), f"native.{variant}.build_receipt", check_path=check_native)
        source = identities.get("sources", {}).get(variant)
        require(isinstance(source, dict), f"native identity missing {variant} source")
        source_receipt = source.get("receipt")
        require(isinstance(source_receipt, dict), f"native identity missing {variant} source receipt")
        ensure_external(Path(source_receipt.get("path", "")))
        embedded_descriptor(source.get("receipt"), f"native.{variant}.source.receipt", check_path=check_native)
    if check_native:
        native_path = Path(manifest["native_manifest"]["path"])
        native = verify_native_manifest(native_path)
        require(native["descriptor"]["sha256"] == manifest["native_manifest"]["sha256"],
                "native manifest changed after allocation freeze")
        for variant in ("baseline", "candidate"):
            require(native["binaries"][variant]["binary"]["sha256"] ==
                    identities["binaries"][variant]["binary"]["sha256"],
                    f"{variant} binary identity changed")


def child_dir(root: Path, row: dict[str, Any]) -> Path:
    return root / "children" / f"{row['index']:03d}-{row['id']}"


def parse_print_summary(stdout: str) -> dict[str, Any]:
    calls = re.search(r"calls to allocation functions:\s*(\d+)", stdout)
    peak = re.search(r"peak heap memory consumption:\s*([0-9]+(?:\.[0-9]+)?[BKMGTPE])", stdout)
    leaked = re.search(r"total memory leaked:\s*(\d+)B", stdout)
    require(calls and peak and leaked, "heaptrack_print summary is incomplete")
    return {
        "allocation_event_count": int(calls.group(1)),
        "peak_heap_display": peak.group(1),
        "live_bytes_at_end": int(leaked.group(1)),
    }


def format_heaptrack_bytes(value: int) -> str:
    units = ("B", "K", "M", "G", "T", "P", "E")
    amount = float(value)
    unit = 0
    while unit < len(units) - 1 and amount > 1000.0:
        amount /= 1000.0
        unit += 1
    if unit == 0:
        return f"{value}B"
    return f"{amount:.2f}{units[unit]}"


def run_analysis(row: dict[str, Any], directory: Path, raw: Path, report: Path,
                 time_log: Path, stdout_path: Path, stderr_path: Path) -> dict[str, Any]:
    histogram = directory / "heaptrack.histogram.tsv"
    print_stdout = directory / "heaptrack-print.stdout"
    print_stderr = directory / "heaptrack-print.stderr"
    parser_output = directory / "allocation-parser.json"
    print_argv = [str(HEAPTRACK_PRINT), "-f", str(raw), "-p", "0", "-T", "0", "-a", "0",
                  "-l", "0", "-H", str(histogram)]
    with print_stdout.open("wb") as out, print_stderr.open("wb") as err:
        print_result = subprocess.run(print_argv, stdout=out, stderr=err, check=False)
    require(print_result.returncode == 0, f"{row['id']}: heaptrack_print failed ({print_result.returncode})")
    summary = parse_print_summary(print_stdout.read_text(encoding="utf-8", errors="replace"))
    parser_argv = [sys.executable, "-B", str(PARSER_PATH), str(raw), "--output", str(parser_output)]
    parser_result = subprocess.run(parser_argv, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                   text=True, check=False)
    require(parser_result.returncode == 0,
            f"{row['id']}: allocation parser failed ({parser_result.returncode}): {parser_result.stdout[-500:]}")
    parsed = load_json(parser_output)
    require(parsed.get("scope") == "whole_process", f"{row['id']}: parser scope changed")
    for key in ("allocation_event_count", "live_bytes_at_end"):
        require(parsed[key] == summary[key], f"{row['id']}: parser/print {key} differs")
    require(summary["peak_heap_display"] == format_heaptrack_bytes(parsed["peak_live_heap_bytes"]),
            f"{row['id']}: parser/print rounded peak differs")
    return {
        "capture_stdout": descriptor(stdout_path),
        "capture_stderr": descriptor(stderr_path),
        "time_log": descriptor(time_log),
        "report": descriptor(report),
        "raw_profile": descriptor(raw),
        "heaptrack_print_stdout": descriptor(print_stdout),
        "heaptrack_print_stderr": descriptor(print_stderr),
        "histogram": descriptor(histogram),
        "parser_output": descriptor(parser_output),
        "parser_result": parsed,
        "heaptrack_print_summary": summary,
        "analysis_argv": print_argv,
        "parser_argv": parser_argv,
    }


def preflight_capture(manifest: dict[str, Any], root: Path) -> None:
    require(os.environ.get("LITCHI_PERF_0923_ALLOC_CAPTURE_APPROVED") == "1",
            "capture requires LITCHI_PERF_0923_ALLOC_CAPTURE_APPROVED=1")
    verify_allocation_manifest(manifest)
    state = root / "capture-state.json"
    require(not state.exists(), f"capture state already exists; selective recapture is prohibited: {state}")
    children = root / "children"
    if children.exists():
        require(not any(children.iterdir()),
                "capture child directory is not empty; selective recapture is prohibited")
    for row in schedule():
        require(not child_dir(root, row).exists(), f"capture output already exists: {row['id']}")


def action_schedule(args: argparse.Namespace) -> dict[str, Any]:
    value = schedule_descriptor()
    if args.output:
        write_json(args.output, value)
    else:
        print(json.dumps(value, indent=2, sort_keys=True))
    return value


def action_freeze(args: argparse.Namespace) -> dict[str, Any]:
    native_path = (args.manifest or DEFAULT_NATIVE_MANIFEST).resolve()
    value = build_allocation_manifest(native_path)
    output = ensure_external(args.output or PACKET_ROOT / "allocation-manifest-v2.json")
    write_json(output, value)
    print(json.dumps(value, indent=2, sort_keys=True))
    return value


def load_frozen_allocation_manifest(path: Path | None = None, *, check_native: bool = True) -> dict[str, Any]:
    manifest_path = ensure_external(path or PACKET_ROOT / "allocation-manifest-v2.json")
    manifest = load_json(manifest_path)
    verify_allocation_manifest(manifest, check_native=check_native)
    return manifest


def action_capture(args: argparse.Namespace) -> dict[str, Any]:
    root = ensure_external(args.packet_root or PACKET_ROOT)
    manifest = load_frozen_allocation_manifest(args.manifest)
    preflight_capture(manifest, root)
    root.mkdir(parents=True, exist_ok=True)
    (root / "children").mkdir(parents=True, exist_ok=True)
    state_path = root / "capture-state.json"
    state: dict[str, Any] = {
        "schema": "litchi.performance.0923.heaptrack-capture-state.v2",
        "status": "running",
        "started_utc": utc_now(),
        "manifest_sha256": manifest["native_manifest"]["sha256"],
        "schedule_sha256": manifest["schedule"]["sha256"],
        "completed": [],
        "failed": None,
    }
    write_json(state_path, state)
    try:
        for row in schedule():
            directory = child_dir(root, row)
            directory.mkdir(parents=True)
            binary = Path(manifest["native_identity"]["binaries"][row["variant"]]["binary"]["path"])
            # heaptrack appends ``.zst`` to the -o value.  Keep the prefix
            # extensionless and bind the resulting compressed profile below;
            # passing a .zst suffix would silently create .zst.zst.
            raw_prefix = directory / "heaptrack"
            raw = directory / "heaptrack.zst"
            report = directory / "harness-report.json"
            time_log = directory / "process-time.log"
            stdout_path = directory / "capture.stdout"
            stderr_path = directory / "capture.stderr"
            argv = [
                "/usr/bin/time", "-v", "-o", str(time_log), "taskset", "-c", str(CPU), "env",
                "LC_ALL=C", f"LITCHI_DOCX_INGRESS_ROUTE={row['route']}", str(HEAPTRACK),
                "--record-only", "-o", str(raw_prefix), str(binary), "--case", CASE,
                "--samples", str(SAMPLES), "--warmup", str(WARMUPS), "--json", str(report),
            ]
            start = utc_now()
            with stdout_path.open("wb") as out, stderr_path.open("wb") as err:
                result = subprocess.run(argv, stdout=out, stderr=err, check=False)
            finish = utc_now()
            require(result.returncode == 0, f"{row['id']}: capture failed ({result.returncode})")
            profiles = sorted(directory.glob("heaptrack*.zst"))
            require(profiles == [raw],
                    f"{row['id']}: expected exactly {raw.name}, found {[path.name for path in profiles]}")
            require(not raw_prefix.exists(),
                    f"{row['id']}: heaptrack output prefix is unexpectedly a file: {raw_prefix}")
            analysis = run_analysis(row, directory, raw, report, time_log, stdout_path, stderr_path)
            receipt = {
                "schema": "litchi.performance.0923.heaptrack-child-receipt.v2",
                "status": "pass",
                "id": row["id"],
                "index": row["index"],
                "route": row["route"],
                "pair": row["pair"],
                "order": row["order"],
                "leg": row["leg"],
                "variant": row["variant"],
                "case": CASE,
                "scope": "whole_process",
                "claims": ["allocation_event_count", "requested_bytes_sum",
                            "peak_live_heap_bytes", "live_bytes_at_end"],
                "start_utc": start,
                "finish_utc": finish,
                "capture_exit_code": result.returncode,
                "binary": descriptor(binary),
                "raw_profile_prefix": str(raw_prefix),
                "capture_argv": argv,
                **analysis,
            }
            write_json(directory / "receipt.json", receipt)
            state["completed"].append(row["id"])
            write_json(state_path, state, replace=True)
        state["status"] = "complete"
        state["finished_utc"] = utc_now()
        write_json(state_path, state, replace=True)
    except BaseException as error:
        state["status"] = "failed"
        state["failed_utc"] = utc_now()
        state["failed"] = str(error)
        write_json(state_path, state, replace=True)
        raise
    print(json.dumps(state, indent=2, sort_keys=True))
    return state


def load_child_receipts(root: Path, *, verify_artifacts: bool) -> list[dict[str, Any]]:
    state = load_json(root / "capture-state.json")
    require(state.get("status") == "complete", "allocation capture is not complete")
    require(state.get("schedule_sha256") == schedule_descriptor()["sha256"], "capture schedule hash changed")
    receipts: list[dict[str, Any]] = []
    for row in schedule():
        receipt_path = child_dir(root, row) / "receipt.json"
        receipt = load_json(receipt_path)
        require(receipt.get("schema") == "litchi.performance.0923.heaptrack-child-receipt.v2",
                f"{row['id']}: receipt schema changed")
        for key in ("id", "index", "route", "pair", "order", "leg", "variant"):
            require(receipt.get(key) == row[key], f"{row['id']}: receipt schedule identity changed")
        require(receipt.get("status") == "pass", f"{row['id']}: receipt is not successful")
        require(receipt.get("scope") == "whole_process", f"{row['id']}: scope claim changed")
        parser_result = receipt.get("parser_result")
        require(isinstance(parser_result, dict) and parser_result.get("scope") == "whole_process",
                f"{row['id']}: parser result is missing or scope changed")
        if verify_artifacts:
            for key in ("binary", "capture_stdout", "capture_stderr", "time_log", "report",
                        "raw_profile", "heaptrack_print_stdout", "heaptrack_print_stderr",
                        "histogram", "parser_output"):
                embedded_descriptor(receipt.get(key), f"{row['id']}.{key}")
        receipts.append(receipt)
    require(len(receipts) == CHILDREN, "allocation receipt count changed")
    return receipts


def metric(receipt: dict[str, Any], name: str) -> int:
    value = receipt["parser_result"].get(name)
    require(type(value) is int and value >= 0, f"{receipt['id']}: invalid metric {name}")
    return value


def reduce_receipts(receipts: list[dict[str, Any]], *, replay: bool = False) -> dict[str, Any]:
    groups: dict[str, dict[str, list[int]]] = {}
    for receipt in receipts:
        key = f"{receipt['route']}/{receipt['variant']}"
        group = groups.setdefault(key, {name: [] for name in (
            "allocation_event_count", "requested_bytes", "peak_live_heap_bytes", "live_bytes_at_end")})
        group["allocation_event_count"].append(metric(receipt, "allocation_event_count"))
        group["requested_bytes"].append(metric(receipt, "requested_bytes"))
        group["peak_live_heap_bytes"].append(metric(receipt, "peak_live_heap_bytes"))
        group["live_bytes_at_end"].append(metric(receipt, "live_bytes_at_end"))
    summary = {}
    for key, values in sorted(groups.items()):
        summary[key] = {
            "children": len(values["allocation_event_count"]),
            "median_allocation_event_count": int(statistics.median(values["allocation_event_count"])),
            "median_requested_bytes": int(statistics.median(values["requested_bytes"])),
            "median_peak_live_heap_bytes": int(statistics.median(values["peak_live_heap_bytes"])),
            "median_live_bytes_at_end": int(statistics.median(values["live_bytes_at_end"])),
        }
    return {
        "schema": "litchi.performance.0923.heaptrack-allocation.reduction.v2",
        "status": "pass",
        "scope": "whole_process",
        "claims": ["allocation_event_count", "requested_bytes_sum",
                    "peak_live_heap_bytes", "live_bytes_at_end"],
        "children": len(receipts),
        "routes": list(ROUTES),
        "groups": summary,
        "adoption": "not evaluated; allocation evidence is supplemental and whole-process",
        "replay": replay,
    }


def action_reduce(args: argparse.Namespace, *, replay: bool = False) -> dict[str, Any]:
    root = ensure_external(args.packet_root or (args.replay_root if replay else PACKET_ROOT))
    if not replay:
        manifest = load_frozen_allocation_manifest(args.manifest)
        require(root.exists(), f"capture packet is missing: {root}")
        verify_allocation_manifest(manifest)
    receipts = load_child_receipts(root, verify_artifacts=not replay)
    value = reduce_receipts(receipts, replay=replay)
    output = args.output or root / ("allocation-replay-reduction-v2.json" if replay
                                    else "allocation-reduction-v2.json")
    write_json(output, value)
    print(json.dumps(value, indent=2, sort_keys=True))
    return value


def action_copy(args: argparse.Namespace) -> dict[str, Any]:
    source = ensure_external(args.packet_root or PACKET_ROOT)
    destination = ensure_external(args.replay_root)
    require(not destination.exists(), f"refusing to replace replay packet: {destination}")
    manifest = load_frozen_allocation_manifest(args.manifest)
    receipts = load_child_receipts(source, verify_artifacts=True)
    destination.mkdir(parents=True)
    copied: dict[str, Any] = {}

    def copy_file(relative: str, source_path: Path) -> None:
        source_path = source_path.resolve()
        size = source_path.stat().st_size
        if size >= ONE_MIB:
            return
        target = destination / relative
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(source_path, target)
        copied[relative] = {"source": descriptor(source_path), "copy": descriptor(target)}

    manifest_path = source / "allocation-manifest-v2.json"
    copy_file("allocation-manifest-v2.json", manifest_path)
    copy_file("capture-state.json", source / "capture-state.json")
    copy_file("allocation-reduction-v2.json", source / "allocation-reduction-v2.json")
    for relative, path in (("allocation-protocol-v2.json", PROTOCOL_PATH),
                           ("allocation-parse-v2.py", PARSER_PATH),
                           ("allocation-parse-controls-v2.py", CONTROLS_PATH),
                           ("allocation-capture-v2.py", CAPTURE_PATH)):
        copy_file(f"tooling/{relative}", path)
    for row in schedule():
        directory = child_dir(source, row)
        target_prefix = f"children/{row['index']:03d}-{row['id']}"
        copy_file(f"{target_prefix}/receipt.json", directory / "receipt.json")
        receipt = next(value for value in receipts if value["id"] == row["id"])
        for key in ("capture_stdout", "capture_stderr", "time_log", "report",
                    "heaptrack_print_stdout", "heaptrack_print_stderr", "histogram", "parser_output"):
            path = Path(receipt[key]["path"])
            copy_file(f"{target_prefix}/{path.name}", path)
    index = {
        "schema": "litchi.performance.0923.heaptrack-allocation.replay-index.v2",
        "created_utc": utc_now(),
        "source_packet": str(source),
        "files": copied,
        "raw_profiles": "retained only at their external receipt paths; never copied into replay packet",
        "native_manifest_sha256": manifest["native_manifest"]["sha256"],
    }
    write_json(destination / "replay-index.json", index)
    print(json.dumps(index, indent=2, sort_keys=True))
    return index


def action_replay(args: argparse.Namespace) -> dict[str, Any]:
    root = ensure_external(args.replay_root)
    index = load_json(root / "replay-index.json")
    require(index.get("schema") == "litchi.performance.0923.heaptrack-allocation.replay-index.v2",
            "allocation replay index schema changed")
    files = index.get("files")
    require(isinstance(files, dict), "allocation replay file map missing")
    for relative, info in files.items():
        path = (root / relative).resolve()
        try:
            path.relative_to(root.resolve())
        except ValueError as error:
            raise AllocationError(f"unsafe replay path: {relative}") from error
        require(path.is_file(), f"missing replay file: {relative}")
        copy_info = info.get("copy") if isinstance(info, dict) else None
        embedded_descriptor(copy_info, f"replay.{relative}")
        actual = descriptor(path)
        require(actual["bytes"] == copy_info["bytes"] and actual["sha256"] == copy_info["sha256"],
                f"replay file changed: {relative}")
    manifest = load_json(root / "allocation-manifest-v2.json")
    verify_allocation_manifest(manifest)
    for filename, key in (("allocation-protocol-v2.json", "protocol"),
                          ("allocation-parse-v2.py", "parser"),
                          ("allocation-parse-controls-v2.py", "controls"),
                          ("allocation-capture-v2.py", "capture_runner")):
        local = root / "tooling" / filename
        actual = descriptor(local, limit=ONE_MIB)
        claim = manifest[key]
        require(actual["bytes"] == claim["bytes"] and actual["sha256"] == claim["sha256"],
                f"replay tooling differs from frozen {key}")
    receipts = load_child_receipts(root, verify_artifacts=False)
    # Local receipt copies are checked through the index; the raw profile
    # remains external and its hash remains in each receipt for custody.
    value = reduce_receipts(receipts, replay=True)
    output = root / "allocation-replay-reduction-v2.json"
    write_json(output, value)
    print(json.dumps(value, indent=2, sort_keys=True))
    return value


def main(argv: Iterable[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("action", choices=("schedule", "freeze", "capture", "reduce", "copy", "replay"))
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--packet-root", type=Path)
    parser.add_argument("--replay-root", type=Path)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args(argv)
    try:
        if args.action == "schedule":
            action_schedule(args)
        elif args.action == "freeze":
            action_freeze(args)
        elif args.action == "capture":
            action_capture(args)
        elif args.action == "reduce":
            action_reduce(args)
        elif args.action == "copy":
            require(args.replay_root is not None, "copy requires --replay-root")
            action_copy(args)
        elif args.action == "replay":
            require(args.replay_root is not None, "replay requires --replay-root")
            action_replay(args)
        return 0
    except (AllocationError, OSError, ValueError, subprocess.SubprocessError) as error:
        print(f"0923 allocation capture failed: {error}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
