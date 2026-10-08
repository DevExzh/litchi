#!/usr/bin/env python3
"""Fail-closed custody, packet, and diagnostic arithmetic for 0970."""
from __future__ import annotations

import hashlib
import importlib.util
import json
import math
import os
import random
import stat
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Iterable

SCHEMA = "litchi.performance.0970.runner.v2"
PLAN_SCHEMA = "litchi.performance.0970.plan.v2"
# The fresh helper/plan versions are v2, but the root authorization receipt
# intentionally retains the independently reviewed v1 schema contract.
GO_SCHEMA = "litchi.performance.0970.root_go.v1"
STAGE_SCHEMA = "litchi.performance.0970.capture_stage.v2"
MANIFEST_SCHEMA = "litchi.performance.0970.capture_manifest.v2"
REDUCTION_SCHEMA = "litchi.performance.0970.diagnostic_reduction.v2"
HELPER_SEAL_SCHEMA = "litchi.performance.0970.helper_seal.v2"

ROOT = Path("/home/zhuhe/litchi-performance-0970")
CORRECTION_RECEIPT = ROOT / "receipts/correction-0970-v2.json"
DESIGN_SHA256 = "95aa5817f95a038c32dfbd4af7906eb474df3546aa321f0adf72c1572e3b462b"
SOURCE_MANIFEST_SHA256 = "f1818af5c306fb6ea01b59ca8df29710e773386b246071202bd4aa896f4e8e2b"
SOURCE_FILE_MANIFEST_SHA256 = "99e65d43e0feed89133eec60b095130d68058a071a02e8a15b05dea01209bb6c"
SOURCE_ROOT = Path("/home/zhuhe/litchi-performance-0969/source-baseline")
SOURCE_MANIFEST = Path("/home/zhuhe/litchi-performance-0969/manifests/source-manifest-v1.json")
SOURCE_FILE_MANIFEST = Path("/home/zhuhe/litchi-performance-0969/manifests/source-file-manifest-v1.json")
VALIDATOR_PATH = Path("/home/zhuhe/litchi-performance-0969/tools/protocol_common_v4.py")
VALIDATOR_SHA256 = "5ce0f2ac08fc7909df054a4cc2d5137b3279af4cd726acf5bfb4558d97b05c5c"

CASES = ("docnohf", "docfloat")
EDIT_COUNTS = (1, 32)
RETENTION = "default"
PAIRS = 12
WARMUPS = 5
SAMPLES = 80
BOOTSTRAP_SEED = 970001
BOOTSTRAP_RESAMPLES = 10_000
TIMEOUT_SECONDS = 900.0
MAX_STDOUT_BYTES = 8 * 1024 * 1024
EVENTS = ("task-clock", "cycles", "instructions", "context-switches", "cpu-migrations", "page-faults")
EXPECTED_ALLOWED_AFFINITY = list(range(32))
EXPECTED_PIN_CPU = 31
REQUIRED_LOCALE = "C"
INTEGER_TIME_FIELDS = ("voluntary_context_switches", "involuntary_context_switches", "minor_faults", "major_faults", "max_rss_kb")
FLOAT_TIME_FIELDS = ("user_s", "system_s")


class ProtocolError(ValueError):
    """An artifact cannot be admitted to the frozen diagnostic."""


def utc() -> str:
    return datetime.now(timezone.utc).isoformat()


def strict_load_json(path: Path) -> dict[str, Any]:
    def pairs(items: list[tuple[str, Any]]) -> dict[str, Any]:
        result: dict[str, Any] = {}
        for key, value in items:
            if key in result:
                raise ProtocolError(f"{path}: duplicate JSON key {key!r}")
            result[key] = value
        return result

    try:
        value = json.loads(path.read_text(encoding="utf-8"), object_pairs_hook=pairs)
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ProtocolError(f"{path}: invalid JSON: {exc}") from exc
    if not isinstance(value, dict):
        raise ProtocolError(f"{path}: expected a JSON object")
    return value


def write_json_exclusive(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps(value, indent=2, sort_keys=True, allow_nan=False) + "\n"
    try:
        fd = os.open(path, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o644)
    except OSError as exc:
        raise ProtocolError(f"refusing to overwrite artifact {path}: {exc}") from exc
    try:
        with os.fdopen(fd, "w", encoding="utf-8") as handle:
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
    except Exception:
        try:
            path.unlink(missing_ok=True)
        except OSError:
            pass
        raise


def sha256(path: Path) -> str:
    if not path.is_file():
        raise ProtocolError(f"missing file: {path}")
    digest = hashlib.sha256()
    try:
        with path.open("rb") as handle:
            for chunk in iter(lambda: handle.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as exc:
        raise ProtocolError(f"cannot read {path}: {exc}") from exc
    return digest.hexdigest()


def binding(path: Path) -> dict[str, Any]:
    path = path.resolve()
    return {"path": str(path), "sha256": sha256(path), "bytes": path.stat().st_size}


def check_binding(row: Any, label: str, *, expected_path: Path | None = None) -> Path:
    if not isinstance(row, dict):
        raise ProtocolError(f"{label}: binding is missing")
    path_value = row.get("path")
    digest = row.get("sha256")
    if not isinstance(path_value, str) or not path_value or not isinstance(digest, str) or len(digest) != 64:
        raise ProtocolError(f"{label}: path/SHA binding is invalid")
    path = Path(path_value).resolve()
    if expected_path is not None and path != expected_path.resolve():
        raise ProtocolError(f"{label}: path drifted ({path} != {expected_path.resolve()})")
    actual = sha256(path)
    if actual != digest:
        raise ProtocolError(f"{label}: SHA mismatch ({actual} != {digest})")
    if row.get("bytes") is not None and row.get("bytes") != path.stat().st_size:
        raise ProtocolError(f"{label}: byte count changed")
    return path


def finite_number(value: Any, label: str, *, positive: bool = False, nonnegative: bool = False) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ProtocolError(f"{label}: expected numeric value")
    number = float(value)
    if not math.isfinite(number):
        raise ProtocolError(f"{label}: non-finite value")
    if positive and number <= 0:
        raise ProtocolError(f"{label}: expected positive value")
    if nonnegative and number < 0:
        raise ProtocolError(f"{label}: expected non-negative value")
    return number


def cell_id(case: str, edits: int, retention: str = RETENTION) -> str:
    return f"{case}/e{edits}/{retention}"


def cells() -> list[dict[str, Any]]:
    return [
        {"cell_id": cell_id(case, edits), "case": case, "edits": edits, "retention": RETENTION}
        for case in CASES
        for edits in EDIT_COUNTS
    ]


def cell_by_id(identity: str) -> dict[str, Any]:
    for row in cells():
        if row["cell_id"] == identity:
            return dict(row)
    raise ProtocolError(f"unknown cell {identity!r}")


def _load_validator() -> Any:
    if sha256(VALIDATOR_PATH) != VALIDATOR_SHA256:
        raise ProtocolError("sealed 0969 validator changed")
    spec = importlib.util.spec_from_file_location("litchi_protocol_common_0969_v4", VALIDATOR_PATH)
    if spec is None or spec.loader is None:
        raise ProtocolError("cannot load sealed 0969 validator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def validate_probe_report(report: dict[str, Any], *, cell: dict[str, Any], fixture: dict[str, Any]) -> None:
    validator = _load_validator()
    try:
        validator.check_probe_report(report, cell=cell, fixture=fixture, allocator=False)
    except Exception as exc:
        raise ProtocolError(f"0969 report/oracle validator rejected packet: {exc}") from exc


def validate_source_custody(source: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(source, dict):
        raise ProtocolError("source binding is missing")
    root = Path(source.get("root", "")).resolve()
    if root != SOURCE_ROOT.resolve():
        raise ProtocolError("historical source root drifted")
    manifest_path = check_binding(source.get("manifest"), "source manifest", expected_path=SOURCE_MANIFEST)
    file_manifest_path = check_binding(source.get("file_manifest"), "source file manifest", expected_path=SOURCE_FILE_MANIFEST)
    manifest = strict_load_json(manifest_path)
    if manifest.get("schema") != "litchi.performance.0969.source_manifest.v1":
        raise ProtocolError("source manifest schema drifted")
    exports = manifest.get("exports") or {}
    if exports.get("baseline_root") != str(root) or exports.get("baseline_file_count") != 924:
        raise ProtocolError("source manifest baseline export drifted")
    if exports.get("baseline_bytes") != source.get("bytes") or exports.get("baseline_tree_sha256") != source.get("tree_sha256"):
        raise ProtocolError("source tree binding drifted")
    file_manifest = strict_load_json(file_manifest_path)
    if file_manifest.get("schema") != "litchi.performance.0969.source_file_manifest.v1":
        raise ProtocolError("source file manifest schema drifted")
    entries = file_manifest.get("entries")
    if not isinstance(entries, list) or len(entries) != 924:
        raise ProtocolError("source file manifest must contain exactly 924 entries")
    seen: set[str] = set()
    total = 0
    for entry in entries:
        if not isinstance(entry, dict) or entry.get("type") != "blob":
            raise ProtocolError("source manifest contains non-blob entry")
        rel = entry.get("path")
        if not isinstance(rel, str) or not rel or rel in seen:
            raise ProtocolError("source manifest path is missing or duplicated")
        seen.add(rel)
        path = (root / rel).resolve()
        if root not in path.parents and path != root:
            raise ProtocolError(f"source path escapes root: {rel}")
        if path.is_symlink() or not path.is_file():
            raise ProtocolError(f"source file is not a regular file: {rel}")
        if path.stat().st_size != entry.get("bytes"):
            raise ProtocolError(f"source byte count mismatch: {rel}")
        if sha256(path) != entry.get("baseline_sha256"):
            raise ProtocolError(f"source baseline SHA mismatch: {rel}")
        mode = entry.get("mode")
        if not isinstance(mode, str) or len(mode) != 6 or mode[:1] != "1":
            raise ProtocolError(f"source mode is invalid: {rel}")
        expected_mode = int(mode[-4:], 8)
        actual_mode = stat.S_IMODE(path.stat().st_mode)
        # The external source export is read through a 0022-umask workspace:
        # every recorded 100644 file is materialized as 100664.  Preserve the
        # manifest mode as normative while rejecting any other mode drift.
        if actual_mode != expected_mode and (actual_mode | 0o022) != (expected_mode | 0o022):
            raise ProtocolError(f"source mode mismatch: {rel}")
        total += int(entry["bytes"])
    if total != source.get("bytes"):
        raise ProtocolError("source total byte count mismatch")
    return {"entries": len(entries), "bytes": total, "tree_sha256": source.get("tree_sha256")}


def expected_order(pair: int) -> tuple[str, str]:
    if pair < 1 or pair > PAIRS:
        raise ProtocolError(f"pair out of range: {pair}")
    return ("unpinned", "pinned") if pair % 2 == 1 else ("pinned", "unpinned")


def parse_affinity(value: Any, label: str) -> list[int]:
    if not isinstance(value, list) or not value or any(type(item) is not int or item < 0 for item in value):
        raise ProtocolError(f"{label}: expected non-empty CPU list")
    result = sorted(set(value))
    if len(result) != len(value):
        raise ProtocolError(f"{label}: duplicate CPU")
    return result


def validate_affinity_controls(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict) or value.get("schema") != "litchi.performance.0970.affinity_controls.v2":
        raise ProtocolError("affinity controls schema drifted")
    allowed = parse_affinity(value.get("allowed"), "controls.allowed")
    if allowed != EXPECTED_ALLOWED_AFFINITY or value.get("pin_cpu") != EXPECTED_PIN_CPU:
        raise ProtocolError("affinity controls allowed mask/pin CPU drifted")
    controls = value.get("controls")
    if not isinstance(controls, dict) or set(controls) != {"unpinned", "pinned"}:
        raise ProtocolError("affinity control arm set drifted")
    expected = {"unpinned": EXPECTED_ALLOWED_AFFINITY, "pinned": [EXPECTED_PIN_CPU]}
    for condition, target in expected.items():
        row = controls[condition]
        if not isinstance(row, dict) or row.get("condition") != condition:
            raise ProtocolError(f"affinity control {condition} identity drifted")
        for field in ("before", "target", "after", "restored"):
            parse_affinity(row.get(field), f"controls.{condition}.{field}")
        if (
            row["before"] != target
            or row["target"] != target
            or row["after"] != target
            or row["restored"] != row["before"]
        ):
            raise ProtocolError(f"affinity control {condition} mask mismatch")
    if controls["unpinned"]["target"] != EXPECTED_ALLOWED_AFFINITY or controls["pinned"]["target"] != [EXPECTED_PIN_CPU]:
        raise ProtocolError("affinity control labels/masks are swapped")
    return value


def parse_time_sidecar(path: Path) -> dict[str, Any]:
    value = strict_load_json(path)
    required = set(FLOAT_TIME_FIELDS) | set(INTEGER_TIME_FIELDS)
    if set(value) != required:
        raise ProtocolError(f"time sidecar fields drifted: {path}")
    for name in FLOAT_TIME_FIELDS:
        finite_number(value[name], f"time.{name}", nonnegative=True)
    for name in INTEGER_TIME_FIELDS:
        if type(value[name]) is not int or value[name] < 0:
            raise ProtocolError(f"time.{name}: expected non-negative integer")
    return value


def parse_perf_sidecar(path: Path) -> dict[str, Any]:
    try:
        lines: list[str] = []
        for raw_line in path.read_text(encoding="utf-8").splitlines():
            line = raw_line.strip()
            if not line:
                continue
            # GNU perf writes exactly one harmless timestamp comment before
            # the six -x CSV rows.  Admit that wrapper header only; every
            # other non-empty comment remains malformed input.
            if line.startswith("#"):
                if not line.startswith("# started on "):
                    raise ProtocolError(f"perf sidecar unexpected comment: {path}")
                continue
            lines.append(line)
    except (OSError, UnicodeError) as exc:
        raise ProtocolError(f"perf sidecar cannot be read: {path}: {exc}") from exc
    rows: dict[str, dict[str, Any]] = {}
    for line in lines:
        fields = line.split(",")
        if len(fields) < 3:
            raise ProtocolError(f"perf sidecar malformed row: {path}")
        raw_value, unit, event = fields[:3]
        if event in rows or event not in EVENTS:
            raise ProtocolError(f"perf sidecar event set drifted: {event!r}")
        if raw_value.strip() in {"", "<not counted>", "<not supported>"}:
            raise ProtocolError(f"perf sidecar missing counter: {event}")
        try:
            value = finite_number(float(raw_value), f"perf.{event}", nonnegative=True)
        except ValueError as exc:
            raise ProtocolError(f"perf.{event}: invalid value") from exc
        runtime = None
        percent = None
        if len(fields) > 3 and fields[3].strip():
            try:
                runtime = int(fields[3])
            except ValueError as exc:
                raise ProtocolError(f"perf.{event}: invalid runtime") from exc
            if runtime <= 0:
                raise ProtocolError(f"perf.{event}: non-positive runtime")
        if len(fields) > 4 and fields[4].strip():
            try:
                percent = finite_number(float(fields[4]), f"perf.{event}.percent", nonnegative=True)
            except ValueError as exc:
                raise ProtocolError(f"perf.{event}: invalid percentage") from exc
        rows[event] = {
            "value": value,
            "unit": unit,
            "runtime_ns": runtime,
            "percent": percent,
            "multiplexed": percent is not None and percent < 99.999,
        }
    if set(rows) != set(EVENTS):
        raise ProtocolError(f"perf sidecar must contain exactly {EVENTS}")
    return {
        "events": rows,
        "all_events_present": True,
        "multiplexed_events": [event for event, row in rows.items() if row["multiplexed"]],
    }


def parse_launch_sidecar(
    path: Path,
    *,
    helper_sha256: str,
    packet_id: str,
    expected_mask: list[int],
    binary: Path,
    helper_path: Path,
) -> dict[str, Any]:
    value = strict_load_json(path)
    if value.get("schema") != "litchi.performance.0970.launch.v2":
        raise ProtocolError("launch sidecar schema drifted")
    if value.get("packet_id") != packet_id:
        raise ProtocolError("launch sidecar packet identity mismatch")
    if value.get("helper_sha256") != helper_sha256:
        raise ProtocolError("launch sidecar helper hash mismatch")
    if value.get("helper_path") != str(helper_path.resolve()):
        raise ProtocolError("launch sidecar helper path mismatch")
    if value.get("execve") is not True:
        raise ProtocolError("launch sidecar does not prove execve")
    if not Path(path).name.endswith(".launch.json") or Path(path).name[:-len(".launch.json")] != value.get("packet_id"):
        raise ProtocolError("launch sidecar filename/packet identity mismatch")
    if value.get("pid") != value.get("exec_pid") or type(value.get("pid")) is not int or value["pid"] <= 0:
        raise ProtocolError("launch sidecar PID identity is invalid")
    requested = parse_affinity(value.get("requested_affinity"), "launch.requested_affinity")
    actual = parse_affinity(value.get("actual_affinity"), "launch.actual_affinity")
    if requested != expected_mask or actual != expected_mask:
        raise ProtocolError("launch sidecar requested/actual affinity mismatch")
    if value.get("exec_path") != str(binary.resolve()):
        raise ProtocolError("launch sidecar executable identity mismatch")
    argv = value.get("exec_argv")
    if not isinstance(argv, list) or not argv or argv[0] != str(binary.resolve()) or any(not isinstance(item, str) for item in argv):
        raise ProtocolError("launch sidecar exec argv is invalid")
    cgroup = value.get("cgroup")
    if not isinstance(cgroup, list) or not cgroup or any(not isinstance(item, str) for item in cgroup):
        raise ProtocolError("launch sidecar cgroup identity is missing")
    if not isinstance(value.get("created_utc"), str) or not value.get("exec_ready_utc"):
        raise ProtocolError("launch sidecar timestamps are missing")
    return value


def finite_samples(report: dict[str, Any]) -> list[float]:
    elapsed = report.get("elapsed_ns")
    if not isinstance(elapsed, list) or len(elapsed) != SAMPLES:
        raise ProtocolError("report elapsed sample count drifted")
    return [finite_number(value, "report.elapsed_ns", positive=True) for value in elapsed]


def percentile(values: Iterable[float], q: float) -> float:
    ordered = sorted(float(value) for value in values)
    if not ordered:
        raise ProtocolError("percentile of empty values")
    position = (len(ordered) - 1) * q
    lower = math.floor(position)
    upper = math.ceil(position)
    if lower == upper:
        return ordered[lower]
    return ordered[lower] + (ordered[upper] - ordered[lower]) * (position - lower)


def median(values: Iterable[float]) -> float:
    return percentile(list(values), 0.5)


def safe_ratio(numerator: float, denominator: float, reason: str) -> dict[str, Any]:
    if not math.isfinite(numerator) or not math.isfinite(denominator):
        return {"value": None, "reason": "non_finite_operand"}
    if denominator == 0:
        return {"value": None, "reason": reason}
    result = numerator / denominator
    if not math.isfinite(result):
        return {"value": None, "reason": "non_finite_ratio"}
    return {"value": result, "reason": None}


def bootstrap_median(values: list[float], *, seed: int = BOOTSTRAP_SEED, resamples: int = BOOTSTRAP_RESAMPLES) -> dict[str, Any]:
    if not values:
        raise ProtocolError("bootstrap values are empty")
    rng = random.Random(seed)
    medians = [median([values[rng.randrange(len(values))] for _ in values]) for _ in range(resamples)]
    return {
        "median": median(values),
        "lower": percentile(medians, 0.025),
        "upper": percentile(medians, 0.975),
        "seed": seed,
        "resamples": resamples,
    }


def robust_summary(values: list[float]) -> dict[str, Any]:
    med = median(values)
    p25 = percentile(values, 0.25)
    p75 = percentile(values, 0.75)
    p95 = percentile(values, 0.95)
    mad = median([abs(value - med) for value in values])
    return {
        "count": len(values),
        "median": med,
        "p25": p25,
        "p75": p75,
        "p95": p95,
        "iqr": p75 - p25,
        "mad": mad,
        "iqr_over_median": safe_ratio(p75 - p25, med, "zero_median"),
        "mad_over_median": safe_ratio(mad, med, "zero_median"),
        "p95_over_median": safe_ratio(p95, med, "zero_median"),
    }


def child_summary(packet: dict[str, Any]) -> dict[str, Any]:
    values = finite_samples(packet["report"])
    p50 = percentile(values, 0.5)
    p95 = percentile(values, 0.95)
    p99 = percentile(values, 0.99)
    return {
        "packet_id": packet["packet_id"],
        "pair": packet["pair"],
        "condition": packet["condition"],
        "order_position": packet["order_position"],
        "p50_ns": p50,
        "p95_ns": p95,
        "p99_ns": p99,
        "mean_ns": sum(values) / len(values),
        "min_ns": min(values),
        "max_ns": max(values),
        "spread_ns": max(values) - min(values),
        "within_tail": safe_ratio(p95, p50, "zero_p50"),
    }


def paired_bootstrap_ratio(pairs: list[tuple[float, float]], *, numerator: str) -> dict[str, Any]:
    ratios: list[float] = []
    for pinned, unpinned in pairs:
        value, denominator = (pinned, unpinned) if numerator == "pinned_over_unpinned" else (unpinned, pinned)
        if denominator <= 0:
            raise ProtocolError("paired ratio denominator is non-positive")
        ratio = value / denominator
        if not math.isfinite(ratio) or ratio <= 0:
            raise ProtocolError("paired ratio is invalid")
        ratios.append(ratio)
    result = bootstrap_median(ratios)
    result["ratios"] = ratios
    return result


def expected_command(plan: dict[str, Any], packet: dict[str, Any]) -> list[str]:
    context = packet["context"]
    cell = context["cell"]
    expected = context["expected_affinity"]
    paths = packet["sidecars"]
    tools = plan["host"]["tools"]
    return [
        tools["/usr/bin/perf"]["path"], "stat", "-x,", "-o", paths["perf"]["path"],
        "-e", ",".join(EVENTS), "--",
        tools["/usr/bin/time"]["path"], "-f", plan["wrapper"]["time_format"], "-o", paths["time"]["path"],
        tools["/usr/bin/python3.14"]["path"], plan["wrapper"]["launcher"],
        "--metadata", paths["launch"]["path"], "--requested-affinity",
        json.dumps(expected, separators=(",", ":")), "--",
        plan["binary"]["path"], "--case", cell["case"], "--input",
        plan["fixtures"][cell["case"]]["path"], "--edits", str(cell["edits"]),
        "--retention", RETENTION, "--warmup", str(WARMUPS), "--samples", str(SAMPLES),
    ]


def check_packet(packet: dict[str, Any], plan: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(packet, dict) or packet.get("schema") != SCHEMA:
        raise ProtocolError("packet schema mismatch")
    context = packet.get("context")
    if not isinstance(context, dict):
        raise ProtocolError("packet context is missing")
    identity = context.get("cell_id")
    cell = cell_by_id(identity)
    if context.get("cell") != cell:
        raise ProtocolError("packet cell identity mismatch")
    pair = context.get("pair")
    if type(pair) is not int or pair < 1 or pair > PAIRS:
        raise ProtocolError("packet pair is invalid")
    expected_order_value = expected_order(pair)
    if context.get("order") != list(expected_order_value):
        raise ProtocolError("packet pair order drifted")
    condition = context.get("condition")
    if condition not in {"pinned", "unpinned"}:
        raise ProtocolError("packet condition is invalid")
    position = context.get("order_position")
    if position not in {"first", "second"} or expected_order_value[0 if position == "first" else 1] != condition:
        raise ProtocolError("packet order position mismatch")
    packet_id = packet.get("packet_id")
    expected_id = f"diag-{cell['case']}-e{cell['edits']}-default-p{pair:02d}-{condition}"
    if packet_id != expected_id:
        raise ProtocolError("packet identity is not the frozen cell/pair/condition identity")
    terminal = packet.get("terminal")
    if not isinstance(terminal, dict) or terminal.get("returncode") != 0 or terminal.get("timed_out") is True or terminal.get("launch_error"):
        raise ProtocolError("packet terminal status failed")
    if packet.get("validation_error") is not None:
        raise ProtocolError(f"packet validation error: {packet['validation_error']}")
    source = plan["source"]
    if context.get("source_manifest_sha256") != source["manifest"]["sha256"] or context.get("source_file_manifest_sha256") != source["file_manifest"]["sha256"]:
        raise ProtocolError("packet source custody binding mismatch")
    if context.get("binary_sha256") != plan["binary"]["sha256"]:
        raise ProtocolError("packet binary custody binding mismatch")
    if context.get("fixture_sha256") != plan["fixtures"][cell["case"]]["sha256"]:
        raise ProtocolError("packet fixture custody binding mismatch")
    expected_mask = parse_affinity(context.get("expected_affinity"), "packet.expected_affinity")
    observed_mask = parse_affinity(context.get("observed_affinity"), "packet.observed_affinity")
    expected_by_condition = EXPECTED_ALLOWED_AFFINITY if condition == "unpinned" else [EXPECTED_PIN_CPU]
    if expected_mask != expected_by_condition:
        raise ProtocolError("packet condition/mask binding mismatch")
    if observed_mask != expected_mask:
        raise ProtocolError("packet observed affinity mismatch")
    if context.get("lc_all") != REQUIRED_LOCALE or context.get("lang") != REQUIRED_LOCALE or context.get("python_dontwritebytecode") != "1":
        raise ProtocolError("packet locale/bytecode environment binding drifted")
    helper = plan["helpers"]["launcher"]
    if context.get("launcher_sha256") != helper["sha256"]:
        raise ProtocolError("packet launcher helper binding mismatch")
    sidecars = packet.get("sidecars")
    if not isinstance(sidecars, dict) or set(sidecars) != {"launch", "time", "perf"}:
        raise ProtocolError("packet sidecar set drifted")
    output_root = Path(packet.get("output_root", "")).resolve()
    plan_output = Path(plan["outputs"]["capture_root"]).resolve()
    if output_root != plan_output:
        raise ProtocolError("packet output root drifted")
    for name in ("launch", "time", "perf"):
        row = sidecars[name]
        check_binding(row, f"packet {name} sidecar")
        path = Path(row["path"]).resolve()
        if output_root not in path.parents:
            raise ProtocolError(f"packet {name} sidecar escapes output root")
    binary = Path(plan["binary"]["path"]).resolve()
    launch = parse_launch_sidecar(
        Path(sidecars["launch"]["path"]),
        helper_sha256=helper["sha256"],
        packet_id=packet_id,
        expected_mask=expected_mask,
        binary=binary,
        helper_path=Path(plan["helpers"]["launcher"]["path"]),
    )
    time_metrics = parse_time_sidecar(Path(sidecars["time"]["path"]))
    perf_metrics = parse_perf_sidecar(Path(sidecars["perf"]["path"]))
    if packet.get("command") != expected_command(plan, packet):
        raise ProtocolError("packet wrapper argv drifted")
    report = packet.get("report")
    if not isinstance(report, dict):
        raise ProtocolError("packet report is missing")
    validate_probe_report(report, cell=cell, fixture=plan["fixtures"][cell["case"]])
    output = packet.get("stdout")
    error = packet.get("stderr")
    for label, row in (("stdout", output), ("stderr", error)):
        check_binding(row, f"packet {label}")
        path = Path(row["path"]).resolve()
        if output_root not in path.parents:
            raise ProtocolError(f"packet {label} escapes output root")
    if packet.get("launcher") != launch:
        raise ProtocolError("packet launch sidecar copy differs from retained sidecar")
    return {"launch": launch, "time": time_metrics, "perf": perf_metrics}


def validate_manifest(manifest: dict[str, Any], plan_path: Path, plan: dict[str, Any], go_path: Path, go_sha256: str) -> list[Path]:
    if manifest.get("schema") != MANIFEST_SCHEMA or manifest.get("no_retry") is not True:
        raise ProtocolError("capture manifest schema/retry binding drifted")
    plan_row = manifest.get("plan")
    if not isinstance(plan_row, dict) or plan_row.get("path") != str(plan_path.resolve()) or plan_row.get("sha256") != sha256(plan_path):
        raise ProtocolError("capture manifest plan binding mismatch")
    go_row = manifest.get("root_go")
    if not isinstance(go_row, dict) or go_row.get("path") != str(go_path.resolve()) or go_row.get("sha256") != go_sha256:
        raise ProtocolError("capture manifest root GO binding mismatch")
    raw = manifest.get("packets")
    if not isinstance(raw, list) or len(raw) > len(cells()) * PAIRS * 2:
        raise ProtocolError("capture manifest packet list is invalid")
    paths: list[Path] = []
    seen: set[Path] = set()
    for value in raw:
        if not isinstance(value, str):
            raise ProtocolError("capture manifest packet path is not a string")
        path = Path(value).resolve()
        if path in seen:
            raise ProtocolError("capture manifest contains duplicate packet paths")
        seen.add(path)
        paths.append(path)
    return paths
