#!/usr/bin/env python3
"""Prepare and, after authorization, run the 0940 owned-ingress cohort.

The runner has three intentionally separate paths.  ``check-schedule`` is a
pure schedule assertion, ``preflight`` runs only Python report fixtures through
the real child-validation helper, and ``capture`` is inert unless the external
manifest authorizes it and ``LITCHI_PERF_0940_CAPTURE_APPROVED=1`` is present.
It never retries a child.  The helper writes each terminal receipt before this
runner asks the report validator to parse JSON.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
import sys
from pathlib import Path
from typing import Any, Mapping


ROOT = Path(__file__).resolve().parents[4]
HERE = Path(__file__).resolve().parent
COMPARISON = HERE / "comparison.py"
HELPER = ROOT / "tools/perf_child_validation.py"
PLAN_SHA256 = "a653a45f4f15110df6b1ff2fd175a428df677fc9702b02750813df4227fa2417"
PLAN_PATH = Path("/home/zhuhe/litchi-performance-0940/plans/prospective-plan-v3.json")
SOURCE_REVISION_LENGTH = 40
WARMUPS = 5
SAMPLES = 50
CPU = 12
DEADLINE_SECONDS = 900.0
OBSERVATION_SECONDS = 1.0
MAX_REPORT_BYTES = 1 << 20
APPROVAL_ENV = "LITCHI_PERF_0940_CAPTURE_APPROVED"
APPROVAL_VALUE = "1"
ARMS = (
    "normal_probes_off_reference",
    "normal_probes_on_diagnostic",
    "allocator_probes_on_diagnostic",
)
CALIBRATION_ARMS = ("normal_probes_on_diagnostic", "allocator_probes_on_diagnostic")
ARM_PERMUTATIONS = (
    ARMS,
    (ARMS[0], ARMS[2], ARMS[1]),
    (ARMS[1], ARMS[0], ARMS[2]),
    (ARMS[1], ARMS[2], ARMS[0]),
    (ARMS[2], ARMS[0], ARMS[1]),
    (ARMS[2], ARMS[1], ARMS[0]),
)
ROUTES = ("owned_vec", "owned_reader")
CALIBRATION_BOUNDARIES = (0, 4, 8, 12)
CHILD_ENV_KEYS = ("PATH", "LC_ALL", "LANG", "TZ", "RUST_BACKTRACE")
TASKSET = Path("/usr/bin/taskset")


class RunnerError(RuntimeError):
    """A schedule, custody, or child contract failed closed."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise RunnerError(message)


def sha256_path(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    path = path.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256_path(path)}


def write_json_exclusive(path: Path, value: Mapping[str, Any]) -> None:
    data = (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())


def _load_module(path: Path, expected_sha256: str, name: str) -> Any:
    actual = sha256_path(path)
    require(actual == expected_sha256, f"{name} hash changed: {actual} != {expected_sha256}")
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load {name}: {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def schedule_rows() -> list[dict[str, Any]]:
    """Return the immutable 72-workload plus 8-calibration schedule."""

    rows: list[dict[str, Any]] = []
    for block in range(1, 13):
        if block == 1:
            rows.extend(_calibration_rows(0))
        arms = ARM_PERMUTATIONS[(block - 1) % len(ARM_PERMUTATIONS)]
        routes = ROUTES if block % 2 else tuple(reversed(ROUTES))
        for arm in arms:
            for position, route in enumerate(routes, start=1):
                rows.append({
                    "id": f"workload-block-{block:02d}-{arm}-{position:02d}-{route}",
                    "kind": "workload",
                    "block": block,
                    "position": position,
                    "arm": arm,
                    "route": route,
                    "warmups": WARMUPS,
                    "samples": SAMPLES,
                    "cpu": CPU,
                    "deadline_seconds": int(DEADLINE_SECONDS),
                })
        if block in (4, 8, 12):
            rows.extend(_calibration_rows(block))
    require(len(rows) == 80, f"schedule child count changed: {len(rows)}")
    require(sum(row["kind"] == "workload" for row in rows) == 72, "workload count changed")
    require(sum(row["kind"] == "calibration" for row in rows) == 8, "calibration count changed")
    require(len({row["id"] for row in rows}) == len(rows), "schedule child IDs are not unique")
    for arm in ARMS:
        require(sum(row["kind"] == "workload" and row["arm"] == arm for row in rows) == 24,
                f"workload arm count changed: {arm}")
    for arm in CALIBRATION_ARMS:
        require(sum(row["kind"] == "calibration" and row["arm"] == arm for row in rows) == 4,
                f"calibration arm count changed: {arm}")
    require(not any(row["kind"] == "calibration" and row["arm"] == ARMS[0] for row in rows),
            "probes-off calibration unexpectedly added")
    for route in ROUTES:
        for arm in ARMS:
            require(sum(row["kind"] == "workload" and row["route"] == route and row["arm"] == arm for row in rows) == 12,
                    f"route/arm count changed: {route}/{arm}")
    return rows


def _calibration_rows(boundary: int) -> list[dict[str, Any]]:
    rows = []
    for arm in ("normal_probes_on_diagnostic", "allocator_probes_on_diagnostic"):
        rows.append({
            "id": f"calibration-after-{boundary:02d}-{arm}",
            "kind": "calibration",
            "boundary_after_block": boundary,
            "position": 1 if arm.startswith("normal") else 2,
            "arm": arm,
            "route": "route_independent",
            "warmups": WARMUPS,
            "samples": SAMPLES,
            "cpu": CPU,
            "deadline_seconds": int(DEADLINE_SECONDS),
        })
    return rows


def schedule_sha256(rows: list[dict[str, Any]] | None = None) -> str:
    rows = schedule_rows() if rows is None else rows
    data = json.dumps(rows, sort_keys=True, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(data).hexdigest()


def _bound_descriptor(value: Any, label: str) -> tuple[Path, dict[str, Any]]:
    require(isinstance(value, Mapping), f"{label} descriptor missing")
    path = Path(str(value.get("path", ""))).resolve()
    require(path.is_file() and not path.is_symlink(), f"{label} path missing: {path}")
    actual_bytes = path.stat().st_size
    require(value.get("bytes") == actual_bytes, f"{label} size changed")
    actual_sha = sha256_path(path)
    require(value.get("sha256") == actual_sha, f"{label} hash changed")
    return path, {
        "path": str(path),
        "bytes": actual_bytes,
        "sha256": actual_sha,
    }


def _json_descriptor(value: Any, label: str) -> tuple[Path, dict[str, Any], dict[str, Any]]:
    path, normalized = _bound_descriptor(value, label)
    try:
        loaded = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise RunnerError(f"{label} is not readable strict JSON: {error}") from error
    require(isinstance(loaded, Mapping), f"{label} JSON is not an object")
    return path, normalized, dict(loaded)


def _relative_source_path(value: Any, label: str) -> Path:
    require(isinstance(value, str) and value, f"{label} path is missing")
    relative = Path(value)
    require(not relative.is_absolute(), f"{label} path is absolute")
    require(".." not in relative.parts, f"{label} path escapes source root")
    require(relative != Path("."), f"{label} path is empty")
    return relative


def _source_listing(root: Path, listing_path: Path) -> tuple[int, int]:
    """Verify every regular source file against the immutable external listing."""

    rows: list[dict[str, Any]] = []
    seen: set[str] = set()
    with listing_path.open("r", encoding="utf-8") as stream:
        for line_number, line in enumerate(stream, start=1):
            try:
                row = json.loads(line)
            except json.JSONDecodeError as error:
                raise RunnerError(f"source listing line {line_number} is not JSON: {error}") from error
            require(isinstance(row, Mapping), f"source listing line {line_number} is not an object")
            relative = _relative_source_path(row.get("path"), f"source listing line {line_number}")
            key = relative.as_posix()
            require(key not in seen, f"source listing contains duplicate path: {key}")
            seen.add(key)
            require(type(row.get("bytes")) is int and row["bytes"] >= 0,
                    f"source listing line {line_number} has invalid size")
            digest = row.get("sha256")
            require(isinstance(digest, str) and len(digest) == 64
                    and all(char in "0123456789abcdef" for char in digest),
                    f"source listing line {line_number} has invalid digest")
            rows.append({"path": relative, "bytes": row["bytes"], "sha256": digest})

    actual: dict[str, Path] = {}
    for directory, directories, files in os.walk(root, followlinks=False):
        directory_path = Path(directory)
        for name in directories + files:
            path = directory_path / name
            require(not path.is_symlink(), f"source tree contains symlink: {path}")
        for name in files:
            path = directory_path / name
            require(path.is_file(), f"source tree entry is not regular: {path}")
            relative = path.relative_to(root).as_posix()
            require(relative not in actual, f"source tree duplicate path: {relative}")
            actual[relative] = path
    expected = {row["path"].as_posix() for row in rows}
    require(set(actual) == expected,
            f"source tree/listing path set mismatch: actual={len(actual)} listed={len(expected)}")
    for row in rows:
        path = actual[row["path"].as_posix()]
        require(path.stat().st_size == row["bytes"], f"source file size changed: {path}")
        require(sha256_path(path) == row["sha256"], f"source file hash changed: {path}")
    total_bytes = sum(row["bytes"] for row in rows)
    return len(rows), total_bytes


def _require_source_custody(source: Mapping[str, Any]) -> dict[str, Any]:
    source_root = Path(str(source.get("source_root", ""))).resolve()
    require(source_root.is_dir(), f"bound source root missing: {source_root}")
    source_revision = source.get("source_revision")
    require(isinstance(source_revision, str) and len(source_revision) == SOURCE_REVISION_LENGTH
            and all(char in "0123456789abcdef" for char in source_revision),
            "source revision is not a lowercase 40-hex value")
    source_receipt_path, source_receipt_desc, source_receipt = _json_descriptor(
        source.get("source_receipt"), "source receipt"
    )
    listing_path, listing_desc = _bound_descriptor(
        source.get("full_source_listing"), "full source listing"
    )
    require(source_receipt.get("source_root") == str(source_root),
            "source receipt root does not match manifest")
    require(source_receipt.get("source_revision") == source_revision,
            "source receipt revision does not match manifest")
    receipt_listing = source_receipt.get("full_source_listing")
    require(isinstance(receipt_listing, Mapping), "source receipt listing binding missing")
    require(Path(str(receipt_listing.get("path", ""))).resolve() == listing_path,
            "source receipt listing path does not match manifest")
    require(receipt_listing.get("bytes") == listing_desc["bytes"]
            and receipt_listing.get("sha256") == listing_desc["sha256"],
            "source receipt listing digest does not match manifest")
    require(source.get("full_source_listing_sha256", listing_desc["sha256"]) == listing_desc["sha256"],
            "manifest listing digest is not bound")
    listed_count, listed_bytes = _source_listing(source_root, listing_path)
    if source_receipt.get("source_file_count") is not None:
        require(source_receipt["source_file_count"] == listed_count,
                "source receipt file count does not match listing")
    if source_receipt.get("source_bytes") is not None:
        require(source_receipt["source_bytes"] == listed_bytes,
                "source receipt byte total does not match listing")
    return {
        "source_root": source_root,
        "source_revision": source_revision,
        "source_receipt": source_receipt_desc,
        "source_receipt_path": source_receipt_path,
        "full_source_listing": listing_desc,
        "full_source_listing_path": listing_path,
        "source_file_count": listed_count,
        "source_bytes": listed_bytes,
    }


def _require_build_custody(
    entry: Mapping[str, Any], *, lane: str, source: Mapping[str, Any], expected_identity: str
) -> dict[str, Any]:
    require(entry.get("source_root") == str(source["source_root"]),
            f"{lane} build source root is not bound")
    require(entry.get("source_revision") == source["source_revision"],
            f"{lane} build source revision is not bound")
    require(entry.get("source_receipt_sha256") == source["source_receipt"]["sha256"],
            f"{lane} build source receipt is not bound")
    require(entry.get("full_source_listing_sha256") == source["full_source_listing"]["sha256"],
            f"{lane} build listing is not bound")
    start_path, start_desc, start = _json_descriptor(
        entry.get("start"), f"{lane} build start receipt"
    )
    terminal_path, terminal_desc, terminal = _json_descriptor(
        entry.get("terminal"), f"{lane} build terminal receipt"
    )
    binary_receipt_path, binary_receipt_desc, binary_receipt = _json_descriptor(
        entry.get("binary_receipt"), f"{lane} binary receipt"
    )
    require(start.get("source_root") == str(source["source_root"]),
            f"{lane} build start source root is not bound")
    require(terminal.get("source_root") == str(source["source_root"]),
            f"{lane} build terminal source root is not bound")
    require(start.get("cwd") == str(source["source_root"])
            and terminal.get("cwd") == str(source["source_root"]),
            f"{lane} build cwd is not source root")
    require(start.get("argv") == terminal.get("argv"),
            f"{lane} build start/terminal argv changed")
    require(terminal.get("exit_code") == 0, f"{lane} build did not exit successfully")
    require(terminal.get("terminal_written_before_binary_descriptor") is True,
            f"{lane} build terminal ordering is not recorded")
    for record, label in ((start, "start"), (terminal, "terminal")):
        argv = record.get("argv")
        require(isinstance(argv, list) and all(isinstance(value, str) for value in argv),
                f"{lane} build {label} argv missing")
        require("--release" in argv and "--offline" in argv and "--locked" in argv,
                f"{lane} build {label} is not the locked release build")
        require("-j2" in argv, f"{lane} build {label} job count changed")
    argv = terminal["argv"]
    require("--bin" in argv and argv[argv.index("--bin") + 1] == expected_identity,
            f"{lane} build binary identity is not bound")
    if lane == "allocator":
        require("--features" in argv and argv[argv.index("--features") + 1] == "allocator-metrics",
                "allocator build feature is not bound")
    else:
        require("--features" not in argv, "normal build unexpectedly enables allocator metrics")
    policy = terminal.get("environment_policy")
    require(isinstance(policy, str) and "RUSTFLAGS" in policy
            and "CARGO_ENCODED_RUSTFLAGS" in policy and "unset" in policy,
            f"{lane} build environment policy is incomplete")
    for record, label in ((start, "start"), (terminal, "terminal")):
        environment = record.get("environment")
        require(isinstance(environment, Mapping), f"{lane} build {label} environment missing")
        require("RUSTFLAGS" not in environment and "CARGO_ENCODED_RUSTFLAGS" not in environment,
                f"{lane} build {label} retained forbidden RUSTFLAGS")
    binary_path, binary_desc = _bound_descriptor(entry.get("binary"), f"{lane} binary")
    require(binary_desc["bytes"] == binary_receipt.get("bytes")
            and binary_desc["sha256"] == binary_receipt.get("sha256")
            and Path(str(binary_receipt.get("path", ""))).resolve() == binary_path,
            f"{lane} binary receipt does not match binary")
    require(binary_receipt.get("binary_name") in (None, expected_identity),
            f"{lane} binary receipt identity changed")
    expected_reports = entry.get("expected_reports")
    require(isinstance(expected_reports, Mapping), f"{lane} expected report metadata missing")
    expected_keys = ("normal_off", "normal_on") if lane == "normal" else ("allocator_on",)
    for expected_key in expected_keys:
        expected = expected_reports.get(expected_key)
        require(isinstance(expected, Mapping), f"{lane} expected report {expected_key} missing")
        for field in (
            "source_revision", "binary_identity", "diagnostic_arm", "diagnostic_enabled",
            "allocator_identity", "instrumentation_identity", "allocator_counter_revision",
            "timing_scope", "observer_gap_scope",
        ):
            require(field in expected, f"{lane} expected report {expected_key}.{field} missing")
        require(expected.get("source_revision") == source["source_revision"],
                f"{lane} expected report {expected_key} source revision changed")
        require(expected.get("binary_identity") == expected_identity,
                f"{lane} expected report {expected_key} binary identity changed")
        require(expected.get("diagnostic_arm") == expected_key,
                f"{lane} expected report {expected_key} arm changed")
        require(expected.get("diagnostic_enabled") is (expected_key != "normal_off"),
                f"{lane} expected report {expected_key} diagnostic flag changed")
        require(isinstance(expected.get("timing_scope"), str) and expected["timing_scope"],
                f"{lane} expected report {expected_key} timing scope missing")
        require(isinstance(expected.get("observer_gap_scope"), str) and expected["observer_gap_scope"],
                f"{lane} expected report {expected_key} observer scope missing")
        require(type(expected.get("diagnostic_enabled")) is bool,
                f"{lane} expected report {expected_key} diagnostic flag missing")
        expected_allocator = ("CountingSystemAllocator(std::alloc::System)"
                              if expected_key == "allocator_on" else "Rust system allocator")
        expected_instrumentation = ("system_allocator_operation_scoped"
                                     if expected_key == "allocator_on" else "none")
        expected_counter = "serialized_region_peak_v3" if expected_key == "allocator_on" else None
        require(expected.get("allocator_identity") == expected_allocator,
                f"{lane} expected report {expected_key} allocator identity changed")
        require(expected.get("instrumentation_identity") == expected_instrumentation,
                f"{lane} expected report {expected_key} instrumentation identity changed")
        require(expected.get("allocator_counter_revision") == expected_counter,
                f"{lane} expected report {expected_key} counter identity changed")
        require(type(expected.get("expected_allocation_present")) is bool
                and type(expected.get("expected_process_present")) is bool,
                f"{lane} expected report {expected_key} presence flags missing")
        require(expected["expected_allocation_present"] == expected["diagnostic_enabled"]
                and expected["expected_process_present"] == expected["diagnostic_enabled"],
                f"{lane} expected report {expected_key} presence flags inconsistent")
    target_dir = entry.get("target_dir")
    if target_dir is not None:
        require(Path(str(target_dir)).resolve().is_dir(), f"{lane} target directory missing")
    return {
        "lane": lane,
        "start": start_desc,
        "start_path": start_path,
        "terminal": terminal_desc,
        "terminal_path": terminal_path,
        "binary_receipt": binary_receipt_desc,
        "binary_receipt_path": binary_receipt_path,
        "binary": binary_desc,
        "binary_path": binary_path,
        "source_root": source["source_root"],
        "source_revision": source["source_revision"],
        "expected_reports": {key: dict(expected_reports[key]) for key in expected_keys},
    }


def _binary_entry(manifest: Mapping[str, Any], arm: str) -> dict[str, Any]:
    lane = "allocator" if arm == "allocator_probes_on_diagnostic" else "normal"
    build = manifest.get("build")
    require(isinstance(build, Mapping), "manifest build mapping missing")
    value = build.get(lane)
    require(isinstance(value, Mapping), f"manifest {lane} build missing")
    build_entry = value
    if "binary" in value and isinstance(value["binary"], Mapping):
        value = value["binary"]
    value = dict(value)
    path = Path(str(value.get("path", "")))
    require(path.is_file() and not path.is_symlink(), f"bound {lane} binary missing: {path}")
    require(value.get("bytes") == path.stat().st_size, f"bound {lane} binary size changed")
    require(value.get("sha256") == sha256_path(path), f"bound {lane} binary hash changed")
    expected_identity = "litchi-perf-baseline-alloc" if lane == "allocator" else "litchi-perf-baseline"
    require(value.get("binary_identity", value.get("identity", expected_identity)) == expected_identity,
            f"bound {lane} binary identity changed")
    value["binary_identity"] = expected_identity
    selected_key = {
        "normal_probes_off_reference": "normal_off",
        "normal_probes_on_diagnostic": "normal_on",
        "allocator_probes_on_diagnostic": "allocator_on",
    }[arm]
    expected_reports = build_entry.get("expected_reports")
    require(isinstance(expected_reports, Mapping), f"bound {lane} expected reports missing")
    expected = expected_reports.get(selected_key)
    require(isinstance(expected, Mapping), f"bound {lane} expected report missing: {selected_key}")
    value["expected_report"] = dict(expected)
    value["expected_allocation_present"] = expected.get("expected_allocation_present")
    value["expected_process_present"] = expected.get("expected_process_present")
    return value


def validate_manifest(manifest: Mapping[str, Any], runner_path: Path) -> dict[str, Any]:
    require(manifest.get("schema") == "litchi.performance.0940.capture-manifest.v1",
            "manifest schema changed")
    require(manifest.get("plan_sha256") == PLAN_SHA256, "manifest is not bound to plan v3")
    require(manifest.get("schedule_sha256") == schedule_sha256(), "manifest schedule hash changed")
    plan_path, plan_desc, plan = _json_descriptor(manifest.get("plan"), "prospective plan")
    require(plan_path == PLAN_PATH.resolve(), "manifest prospective plan path changed")
    require(plan_desc["sha256"] == PLAN_SHA256, "manifest prospective plan digest changed")
    require(plan.get("schema") == "litchi.performance.0940.prep-plan.v3",
            "manifest prospective plan schema changed")
    source = manifest.get("source")
    require(isinstance(source, Mapping), "manifest source mapping missing")
    source_binding = _require_source_custody(source)
    source_root = source_binding["source_root"]
    source_revision = source_binding["source_revision"]
    require(source_revision == plan["canonical_baseline_custody"]["source_revision"],
            "source revision differs from the prospective canonical baseline")
    scripts = manifest.get("scripts", {})
    require(isinstance(scripts, Mapping), "manifest scripts mapping missing")
    helper = scripts.get("child_validation", manifest.get("validation_helper"))
    comparison = scripts.get("comparison", manifest.get("comparison"))
    runner = scripts.get("runner", manifest.get("runner"))
    taskset = scripts.get("taskset")
    for label, value, fallback in (("child validator", helper, HELPER), ("comparison", comparison, COMPARISON),
                                    ("runner", runner, runner_path), ("taskset", taskset, TASKSET)):
        require(isinstance(value, Mapping), f"manifest {label} descriptor missing")
        path = Path(str(value.get("path", fallback))).resolve()
        require(path == fallback.resolve(), f"manifest {label} path changed")
        _bound_descriptor(value, label)
        if label == "child validator":
            require(path == HELPER.resolve(), "manifest child validator path changed")
        if label == "comparison":
            require(path == COMPARISON.resolve(), "manifest comparison path changed")
        if label == "runner":
            require(path == runner_path.resolve(), "manifest runner path changed")
        if label == "taskset":
            require(path == TASKSET.resolve() and os.access(path, os.X_OK),
                    "manifest taskset executable is not usable")
    environment = manifest.get("child_environment")
    require(isinstance(environment, Mapping), "manifest child environment missing")
    require(set(environment) <= set(CHILD_ENV_KEYS), "manifest child environment has unreviewed keys")
    require("PATH" in environment and "LC_ALL" in environment and "LANG" in environment and "TZ" in environment,
            "manifest child environment is incomplete")
    require(environment.get("TZ") == "UTC", "manifest child timezone is not UTC")
    require(dict(environment) == {
        "PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C",
        "TZ": "UTC", "RUST_BACKTRACE": "0",
    }, "manifest child environment differs from the reviewed real preflight")
    builds = manifest.get("build")
    require(isinstance(builds, Mapping), "manifest build custody missing")
    build_bindings = {
        "normal": _require_build_custody(
            builds.get("normal") if isinstance(builds, Mapping) else {},
            lane="normal", source=source_binding, expected_identity="litchi-perf-baseline",
        ),
        "allocator": _require_build_custody(
            builds.get("allocator") if isinstance(builds, Mapping) else {},
            lane="allocator", source=source_binding, expected_identity="litchi-perf-baseline-alloc",
        ),
    }
    _binary_entry(manifest, "normal_probes_off_reference")
    _binary_entry(manifest, "allocator_probes_on_diagnostic")
    authorization = manifest.get("authorization")
    require(isinstance(authorization, Mapping) and authorization.get("capture_authorized") is True,
            "capture manifest is not authorized")
    return {
        "source_root": source_root,
        "source_revision": source_revision,
        "source_binding": source_binding,
        "build_bindings": build_bindings,
        "environment": {str(key): str(value) for key, value in environment.items()},
        "helper": Path(str(helper["path"])).resolve(),
        "comparison": Path(str(comparison["path"])).resolve(),
    }


def _child_command(row: Mapping[str, Any], binary: Path, report: Path, source_revision: str) -> tuple[str, ...]:
    require(TASKSET.is_file() and os.access(TASKSET, os.X_OK), "taskset executable is unavailable")
    command = [str(TASKSET), "--cpu-list", str(CPU), str(binary), "docx-owned-attribution", "--case", str(row["kind"]),
               "--diagnostic", "on" if row["arm"] != "normal_probes_off_reference" else "off",
               "--warmups", str(WARMUPS), "--samples", str(SAMPLES),
               "--source-revision", source_revision, "--output", str(report)]
    if row["kind"] == "workload":
        command.extend(("--route", str(row["route"])))
    return tuple(command)


def _preflight_child(helper: Any, comparison: Any, scratch: Path, *, case: str, route: str,
                     diagnostic: bool, mutated: bool = False,
                     binary_identity: str = "litchi-perf-baseline") -> dict[str, Any]:
    scratch.mkdir(parents=True, exist_ok=True)
    child = scratch / f"{case}-{route}-{'bad' if mutated else 'good'}"
    child.mkdir()
    report_path = child / "report.json"
    payload = comparison.synthetic_report(
        route=route, case=case, diagnostic=diagnostic, binary_identity=binary_identity
    )
    if mutated and case == "workload":
        payload["output_sha256"][0] = "0" * 64
    payload_path = child / "payload.json"
    payload_path.write_text(json.dumps(payload, sort_keys=True), encoding="utf-8")
    code = "import pathlib,sys; pathlib.Path(sys.argv[2]).write_bytes(pathlib.Path(sys.argv[1]).read_bytes())"
    command = (sys.executable, "-c", code, str(payload_path), str(report_path))
    helper_module = helper
    spec = helper_module.ChildSpec(
        command=command,
        cwd=scratch,
        environment={"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC", "RUST_BACKTRACE": "0"},
        paths=helper_module.ChildPaths(
            report=report_path,
            stdout=child / "stdout.log",
            stderr=child / "stderr.log",
            terminal=child / "terminal.json",
            validation=child / "validation.json",
        ),
        route=route,
        label=f"preflight-{case}-{route}",
        warmups=WARMUPS,
        samples=SAMPLES,
        expected_output_sha256=None if case == "calibration" else comparison.OUTPUT_SHA256,
        binary={
            "identity": binary_identity,
            "source_revision": payload["source_revision"],
            "expected_report": {
                field: payload[field]
                for field in (
                    "source_revision", "binary_identity", "diagnostic_arm", "diagnostic_enabled",
                    "allocator_identity", "instrumentation_identity", "allocator_counter_revision",
                    "timing_scope", "observer_gap_scope",
                )
            },
            "expected_allocation_present": payload.get("allocation") is not None,
            "expected_process_present": payload.get("process") is not None,
        },
        report_contract=(helper_module.REPORT_CONTRACT_DOCX_OWNED_CALIBRATION
                         if case == "calibration"
                         else helper_module.REPORT_CONTRACT_DOCX_OWNED_WORKLOAD),
        deadline_seconds=30.0,
        observation_seconds=0.05,
        max_report_bytes=MAX_REPORT_BYTES,
    )
    return helper_module.run_child_and_validate(spec, comparison)


def run_preflight(helper_path: Path, helper_sha256: str, comparison_path: Path,
                  comparison_sha256: str, output: Path | None, scratch: Path) -> dict[str, Any]:
    helper = _load_module(helper_path, helper_sha256, "child-validation helper")
    comparison = _load_module(comparison_path, comparison_sha256, "0940 comparison validator")
    scratch = scratch.resolve()
    scratch.mkdir(parents=True, exist_ok=False)
    workload = _preflight_child(helper, comparison, scratch, case="workload", route="owned_vec", diagnostic=False)
    calibration = _preflight_child(helper, comparison, scratch, case="calibration", route="route_independent", diagnostic=True)
    allocator = _preflight_child(
        helper, comparison, scratch, case="workload", route="owned_reader", diagnostic=True,
        binary_identity="litchi-perf-baseline-alloc",
    )
    mutated = _preflight_child(helper, comparison, scratch, case="workload", route="owned_vec", diagnostic=False, mutated=True)
    require(workload["status"] == "passed", "positive workload preflight failed")
    require(calibration["status"] == "passed", "positive calibration preflight failed")
    require(allocator["status"] == "passed", "positive allocator preflight failed")
    require(mutated["status"] == "failed", "mutated output preflight was accepted")
    summary = {
        "schema": "litchi.performance.0940.preflight.v1",
        "status": "passed",
        "artifacts_retained": True,
        "artifact_root": str(scratch),
        "helper": {"path": str(helper_path.resolve()), "sha256": helper_sha256},
        "comparison": {"path": str(comparison_path.resolve()), "sha256": comparison_sha256},
        "positive_workload": workload["validation"],
        "positive_calibration": calibration["validation"],
        "positive_allocator": allocator["validation"],
        "negative_mutated_output": mutated["validation"],
        "terminal_before_validation": workload["terminal"]["status"] == "terminal",
    }
    if output is not None:
        write_json_exclusive(output, summary)
    return summary


def capture(manifest_path: Path, packet: Path, runner_path: Path) -> dict[str, Any]:
    require(os.environ.get(APPROVAL_ENV) == APPROVAL_VALUE,
            f"set {APPROVAL_ENV}=1 only for an authorized capture")
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    binding = validate_manifest(manifest, runner_path)
    packet = packet.resolve()
    require(packet != ROOT and ROOT not in packet.parents, "formal packet must be external to checkout")
    packet.mkdir(parents=True, exist_ok=False)
    helper = _load_module(binding["helper"], manifest["scripts"]["child_validation"]["sha256"], "child-validation helper")
    comparison = _load_module(binding["comparison"], manifest["scripts"]["comparison"]["sha256"], "0940 comparison validator")
    rows = schedule_rows()
    results: list[dict[str, Any]] = []
    for index, row in enumerate(rows, start=1):
        child = packet / "children" / row["id"]
        child.mkdir(parents=True, exist_ok=False)
        binary = _binary_entry(manifest, row["arm"])
        report = child / "report.json"
        command = _child_command(row, Path(binary["path"]), report, binding["source_revision"])
        spec = helper.ChildSpec(
            command=command,
            cwd=binding["source_root"],
            environment=binding["environment"],
            paths=helper.ChildPaths(
                report=report,
                stdout=child / "stdout.log",
                stderr=child / "stderr.log",
                terminal=child / "terminal.json",
                validation=child / "validation.json",
            ),
            route=row["route"],
            label=row["id"],
            warmups=WARMUPS,
            samples=SAMPLES,
            expected_output_sha256=None if row["kind"] == "calibration" else comparison.OUTPUT_SHA256,
            binary={**binary, "source_revision": binding["source_revision"]},
            report_contract=(helper.REPORT_CONTRACT_DOCX_OWNED_CALIBRATION
                             if row["kind"] == "calibration"
                             else helper.REPORT_CONTRACT_DOCX_OWNED_WORKLOAD),
            deadline_seconds=DEADLINE_SECONDS,
            observation_seconds=OBSERVATION_SECONDS,
            max_report_bytes=MAX_REPORT_BYTES,
        )
        outcome = helper.run_child_and_validate(spec, comparison)
        results.append({"index": index, "id": row["id"], "kind": row["kind"], "status": outcome["status"]})
        write_json_exclusive(packet / f"capture-state-{index:03d}.json", {
            "schema": "litchi.performance.0940.capture-progress.v1",
            "completed": index,
            "total": len(rows),
            "last": results[-1],
        })
    after_count, after_bytes = _source_listing(
        binding["source_binding"]["source_root"],
        binding["source_binding"]["full_source_listing_path"],
    )
    require(after_count == binding["source_binding"]["source_file_count"]
            and after_bytes == binding["source_binding"]["source_bytes"],
            "source closure changed after capture")
    summary = {
        "schema": "litchi.performance.0940.capture-state.v1",
        "status": "completed" if all(item["status"] == "passed" for item in results) else "completed_with_failures",
        "children": len(results),
        "passed": sum(item["status"] == "passed" for item in results),
        "failed": sum(item["status"] != "passed" for item in results),
        "schedule_sha256": schedule_sha256(rows),
        "source_closure_after": {
            "file_count": after_count,
            "bytes": after_bytes,
            "listing_sha256": binding["source_binding"]["full_source_listing"]["sha256"],
        },
        "results": results,
    }
    write_json_exclusive(packet / "capture-state.json", summary)
    return summary


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check-schedule", action="store_true")
    parser.add_argument("--preflight", action="store_true")
    parser.add_argument("--capture", action="store_true")
    parser.add_argument("--helper", type=Path, default=HELPER)
    parser.add_argument("--helper-sha256")
    parser.add_argument("--comparison", type=Path, default=COMPARISON)
    parser.add_argument("--comparison-sha256")
    parser.add_argument("--preflight-output", type=Path)
    parser.add_argument(
        "--preflight-scratch",
        type=Path,
        default=Path("/home/zhuhe/litchi-performance-0940/preflight-v2"),
    )
    parser.add_argument("--manifest", type=Path)
    parser.add_argument("--packet", type=Path)
    args = parser.parse_args(argv)
    try:
        selected = sum((args.check_schedule, args.preflight, args.capture))
        require(selected == 1, "choose exactly one of --check-schedule, --preflight, or --capture")
        if args.check_schedule:
            result = {"schema": "litchi.performance.0940.schedule.v1", "children": len(schedule_rows()), "schedule_sha256": schedule_sha256()}
        elif args.preflight:
            require(args.helper_sha256 and args.comparison_sha256, "preflight requires helper and comparison SHA-256 values")
            result = run_preflight(
                args.helper.resolve(), args.helper_sha256, args.comparison.resolve(),
                args.comparison_sha256, args.preflight_output, args.preflight_scratch,
            )
        else:
            require(args.manifest is not None and args.packet is not None, "capture requires --manifest and --packet")
            result = capture(args.manifest.resolve(), args.packet.resolve(), Path(__file__).resolve())
        print(json.dumps(result, sort_keys=True))
        return 0
    except Exception as error:
        print(json.dumps({"schema": "litchi.performance.0940.runner-error.v1", "status": "failed", "error": f"{type(error).__name__}: {error}"}, sort_keys=True))
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
