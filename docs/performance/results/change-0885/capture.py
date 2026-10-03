#!/usr/bin/env python3
"""Capture the 0885 release-profile XLSX producer baseline.

The driver deliberately has a small surface.  It launches one fresh process
per child, in the frozen native-then-allocator order, and writes every input
and output needed to audit a child without rerunning it.  It does not compare
timings or make a performance decision; :mod:`verify` performs the descriptive
validation and statistics.
"""

from __future__ import annotations

import argparse
import datetime as _datetime
import hashlib
import json
import os
from pathlib import Path
import stat
import subprocess
from typing import Any


SCHEMA = "litchi.xlsx-producer-full-cell-capture-0885.v1"
COMMAND_SCHEMA = "litchi.xlsx-producer-full-cell-child-0885.v1"
RECEIPT_SCHEMA = "litchi.xlsx-producer-full-cell-receipt-0885.v1"
CAPTURE_SCHEMA = "litchi.xlsx-producer-full-cell-capture-receipt-0885.v1"
CASE = "xlsx_producer_medium_source_full_cell_scan"
PROFILE = "release"
CPU_AFFINITY = "12"
TASKSET = "/usr/bin/taskset"
CHILD_ENV = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC"}
ARCHIVE_SHA256 = "55e0901c093c27d80bcd4e5fef888f00798654cdef8c0d79d130a2ce7cbe0ae8"
ARCHIVE_BYTES = 22_739
ORACLE_SHA256 = "8d29b40718aa9d5b08c746f841195bba17ab8e01608b12e6093daf160c00764b"
CENSUS_SHA256 = "552c35c8b8f958de28792f00c1e9e831ae884471aa28b5b2a7a1479bd06c5207"
COUNTS = {"cell": 4_096, "numeric": 3_686, "text": 410}
NATIVE_CHILDREN = 12
NATIVE_WARMUP = 3
NATIVE_SAMPLES = 30
ALLOCATOR_CHILDREN = 6
ALLOCATOR_WARMUP = 3
ALLOCATOR_SAMPLES = 15
REQUIRED_BINDINGS = (
    "build",
    "source_manifest",
    "source_archive",
    "capture_helper",
    "verify_helper",
    "test_helper",
    "review",
    "environment",
    "taskset",
)


def _utc_now() -> str:
    return _datetime.datetime.now(_datetime.timezone.utc).isoformat()


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, Any]:
    """Return a stable filesystem descriptor for a retained regular file."""

    path = path.resolve()
    info = path.stat()
    if not stat.S_ISREG(info.st_mode):
        raise ValueError(f"expected a regular file: {path}")
    return {
        "path": str(path),
        "bytes": info.st_size,
        "sha256": _sha256(path),
        "mode": stat.S_IMODE(info.st_mode),
        "executable": bool(info.st_mode & stat.S_IXUSR),
    }


def _write_json(path: Path, value: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("x", encoding="utf-8") as handle:
        json.dump(value, handle, indent=2, sort_keys=True)
        handle.write("\n")


def _read_json(path: Path) -> Any:
    with path.open(encoding="utf-8") as handle:
        return json.load(handle)


def _digest_json(value: Any) -> str:
    encoded = json.dumps(value, sort_keys=True, separators=(",", ":")).encode()
    return hashlib.sha256(encoded).hexdigest()


def load_plan(path: Path) -> dict[str, Any]:
    plan = _read_json(path)
    if not isinstance(plan, dict):
        raise ValueError("plan must be a JSON object")
    validate_plan(plan)
    _validate_bindings(path.resolve().parent, plan)
    return plan


def _validate_bindings(packet: Path, plan: dict[str, Any]) -> None:
    """Validate every frozen input before a child can be launched."""

    bindings = plan.get("bindings")
    if not isinstance(bindings, dict):
        raise ValueError("plan must contain frozen bindings")
    missing = [key for key in REQUIRED_BINDINGS if key not in bindings]
    if missing:
        raise ValueError(f"plan bindings missing: {', '.join(missing)}")
    for name in bindings:
        item = bindings[name]
        if not isinstance(item, dict) or not isinstance(item.get("path"), str):
            raise ValueError(f"invalid frozen binding {name}")
        if not isinstance(item.get("bytes"), int) or not isinstance(item.get("sha256"), str):
            raise ValueError(f"frozen binding {name} must include bytes and sha256")
        path = Path(item["path"])
        if not path.is_absolute():
            path = packet / path
        if not path.is_file():
            raise ValueError(f"frozen binding {name} is missing: {path}")
        actual = descriptor(path)
        for field in ("bytes", "sha256", "mode", "executable"):
            if field in item and item[field] != actual[field]:
                raise ValueError(f"frozen binding {name} {field} mismatch")
        if name == "taskset" and path.resolve() != Path(TASKSET).resolve():
            raise ValueError("taskset binding is not /usr/bin/taskset")
    build_path = Path(plan.get("build_json", "build.json"))
    if not build_path.is_absolute():
        build_path = packet / build_path
    bound_build = Path(bindings["build"]["path"])
    if not bound_build.is_absolute():
        bound_build = packet / bound_build
    if build_path.resolve() != bound_build.resolve():
        raise ValueError("plan build_json does not equal its frozen build binding")


def _lane(plan: dict[str, Any], mode: str) -> dict[str, Any]:
    lanes = plan.get("lanes")
    if isinstance(lanes, dict):
        row = lanes.get(mode)
    elif isinstance(lanes, list):
        row = next((candidate for candidate in lanes if candidate.get("name", candidate.get("mode")) == mode), None)
    else:
        row = None
    if not isinstance(row, dict):
        raise ValueError(f"plan has no {mode} lane")
    return row


def validate_plan(plan: dict[str, Any]) -> None:
    if plan.get("schema") != SCHEMA:
        raise ValueError(f"unexpected plan schema: {plan.get('schema')!r}")
    if plan.get("case") != CASE:
        raise ValueError("plan case is not the producer full-cell case")
    if plan.get("profile") != PROFILE:
        raise ValueError("0885 capture requires release profile")
    if str(plan.get("cpu_affinity", plan.get("cpu"))) != CPU_AFFINITY:
        raise ValueError("0885 capture is pinned to CPU 12")
    if plan.get("cache", plan.get("filesystem_cache")) != "warm":
        raise ValueError("0885 capture requires warm filesystem cache")
    if plan.get("fresh_child") is not True:
        raise ValueError("0885 capture requires a fresh child for every row")
    if plan.get("fail_fast") is not True:
        raise ValueError("0885 capture is fail-fast and cannot retry a child")
    if plan.get("schedule") != ["native", "allocator"]:
        raise ValueError("0885 schedule must be native then allocator")
    archive = plan.get("archive", {})
    if not isinstance(archive, dict):
        raise ValueError("plan archive must be an object")
    expected_archive = {
        "sha256": ARCHIVE_SHA256,
        "bytes": ARCHIVE_BYTES,
        "oracle_sha256": ORACLE_SHA256,
        "marker_refusal_census_sha256": CENSUS_SHA256,
        "entry_count": COUNTS["cell"],
        "numeric_count": COUNTS["numeric"],
        "text_count": COUNTS["text"],
    }
    # The key is intentionally named marker_refusal_census_sha256 in the
    # packet contract; accept census_sha256 only while constructing a plan so
    # an older handoff can be diagnosed rather than silently ignored.
    for key, expected in expected_archive.items():
        actual = archive.get(key)
        if key == "marker_refusal_census_sha256" and actual is None:
            actual = archive.get("census_sha256")
        if actual != expected:
            raise ValueError(f"plan archive {key} is {actual!r}, expected {expected!r}")
    for mode, expected_children, expected_warmup, expected_samples in (
        ("native", NATIVE_CHILDREN, NATIVE_WARMUP, NATIVE_SAMPLES),
        ("allocator", ALLOCATOR_CHILDREN, ALLOCATOR_WARMUP, ALLOCATOR_SAMPLES),
    ):
        lane = _lane(plan, mode)
        if lane.get("name", lane.get("mode")) != mode or lane.get("mode", mode) != mode:
            raise ValueError(f"{mode} lane name/mode mismatch")
        if lane.get("children") != expected_children:
            raise ValueError(f"{mode} lane child count mismatch")
        if lane.get("warmup") != expected_warmup:
            raise ValueError(f"{mode} lane warmup mismatch")
        if lane.get("samples") != expected_samples:
            raise ValueError(f"{mode} lane sample count mismatch")
    binaries = plan.get("binaries")
    if not isinstance(binaries, dict) or not all(mode in binaries for mode in ("native", "allocator")):
        raise ValueError("plan must bind native and allocator binaries")
    for mode in ("native", "allocator"):
        binary = _entry(binaries[mode])
        if "bytes" not in binary or "sha256" not in binary or binary.get("executable") is not True:
            raise ValueError(f"{mode} plan binary must bind complete size/hash/executable identity")


def _entry(value: Any) -> dict[str, Any]:
    if isinstance(value, str):
        return {"path": value}
    if isinstance(value, dict) and isinstance(value.get("path"), str):
        return value
    raise ValueError(f"invalid binary descriptor: {value!r}")


def _build_descriptor(packet: Path, plan: dict[str, Any], mode: str) -> dict[str, Any]:
    build_path_value = plan.get("build_json", "build.json")
    build_path = Path(build_path_value)
    if not build_path.is_absolute():
        build_path = packet / build_path
    if not build_path.is_file():
        raise ValueError(f"missing build descriptor: {build_path}")
    build = _read_json(build_path)
    if not isinstance(build, dict):
        raise ValueError("build descriptor must be an object")
    if build.get("status") != "pass":
        raise ValueError("capture requires a passing build descriptor")
    declared = _entry(plan["binaries"][mode])
    build_binaries = build.get("binaries", build)
    built = build_binaries.get(mode) if isinstance(build_binaries, dict) else None
    if built is None:
        raise ValueError(f"build descriptor has no {mode} binary")
    built = _entry(built)
    for key in ("bytes", "sha256", "executable"):
        if key not in built or declared.get(key) != built.get(key):
            raise ValueError(f"{mode} plan/build {key} mismatch")
    expected_features = [] if mode == "native" else ["allocator-metrics"]
    if declared.get("features") != expected_features or built.get("features") != expected_features:
        raise ValueError(f"{mode} binary feature binding mismatch")
    raw_path = declared["path"]
    path = Path(raw_path)
    if not path.is_absolute():
        path = packet / path
    resolved = path.resolve()
    expected_path = (packet / "binaries" / mode).resolve()
    if resolved != expected_path:
        raise ValueError(f"{mode} binary must be the retained packet path {expected_path}")
    if not resolved.is_file() or not os.access(resolved, os.X_OK):
        raise ValueError(f"{mode} binary is not executable: {resolved}")
    actual = descriptor(resolved)
    for key in ("bytes", "sha256", "executable"):
        if declared.get(key) != actual[key]:
            raise ValueError(f"{mode} binary {key} does not match plan")
        if built.get(key) != actual[key]:
            raise ValueError(f"{mode} binary {key} does not match build")
    for source in (declared, built):
        declared_path = Path(source["path"])
        if not declared_path.is_absolute():
            declared_path = packet / declared_path
        if declared_path.resolve() != resolved:
            raise ValueError(f"{mode} plan/build binary path mismatch")
    return {
        "path": str(resolved),
        "identity": actual,
        "plan": declared,
        "build": built,
        "build_json": str(build_path.resolve()),
        "build_sha256": _sha256(build_path),
    }


def _command(mode: str, binary: str, warmup: int, samples: int, report: Path, evidence: Path) -> list[str]:
    return [
        TASKSET,
        "-c",
        CPU_AFFINITY,
        binary,
        "--case",
        CASE,
        "--warmup",
        str(warmup),
        "--samples",
        str(samples),
        "--filesystem-cache",
        "warm",
        "--json",
        str(report),
        "--producer-evidence",
        str(evidence),
    ]


def _artifacts(folder: Path) -> dict[str, dict[str, Any]]:
    result: dict[str, dict[str, Any]] = {}
    for name in ("stdout.txt", "stderr.txt", "report.json", "producer-evidence.json"):
        path = folder / name
        if path.is_file():
            result[name] = descriptor(path)
    return result


def run_lane(
    plan_path: Path,
    output_root: Path,
    mode: str,
    *,
    popen: Any = subprocess.Popen,
) -> dict[str, Any]:
    """Run one already-ordered lane; intended for tests and the main driver."""

    packet = plan_path.resolve().parent
    plan = load_plan(plan_path)
    if mode not in ("native", "allocator"):
        raise ValueError(f"unknown lane: {mode}")
    lane = _lane(plan, mode)
    binary = _build_descriptor(packet, plan, mode)
    output_root.mkdir(parents=True, exist_ok=True)
    lane_root = output_root / mode
    lane_root.mkdir(exist_ok=False)
    records: list[dict[str, Any]] = []
    for index in range(lane["children"]):
        folder = lane_root / f"child-{index:02d}"
        folder.mkdir(exist_ok=False)
        report = folder / "report.json"
        evidence = folder / "producer-evidence.json"
        stdout = folder / "stdout.txt"
        stderr = folder / "stderr.txt"
        argv = _command(mode, binary["path"], lane["warmup"], lane["samples"], report, evidence)
        started = _utc_now()
        command = {
            "schema": COMMAND_SCHEMA,
            "mode": mode,
            "child_index": index,
            "argv": argv,
            "cwd": str(packet),
            "environment": CHILD_ENV.copy(),
            "profile": PROFILE,
            "binary": binary,
            "plan_sha256": _sha256(plan_path),
            "started_utc": started,
        }
        env = CHILD_ENV.copy()
        with stdout.open("xb") as stdout_handle, stderr.open("xb") as stderr_handle:
            process = popen(argv, cwd=packet, env=env, stdout=stdout_handle, stderr=stderr_handle)
            code = process.wait()
        finished = _utc_now()
        command["finished_utc"] = finished
        command["exit_code"] = code
        _write_json(folder / "command.json", command)
        row = {
            "schema": RECEIPT_SCHEMA,
            "mode": mode,
            "child_index": index,
            "argv": argv,
            "cwd": str(packet),
            "profile": PROFILE,
            "plan_sha256": command["plan_sha256"],
            "binary": binary,
            "environment": command["environment"],
            "started_utc": started,
            "finished_utc": finished,
            "exit_code": code,
            "command": descriptor(folder / "command.json"),
            "artifacts": _artifacts(folder),
        }
        _write_json(folder / "receipt.json", row)
        records.append({"path": str((folder / "receipt.json").relative_to(output_root)), "receipt": row})
        if code != 0:
            raise RuntimeError(f"{mode} child {index} exited {code}; capture stops without retry")
        if not report.is_file() or not evidence.is_file():
            raise RuntimeError(f"{mode} child {index} did not produce report and evidence")
        # Import lazily to keep this module directly executable and avoid a
        # module cycle.  A zero exit code is insufficient: malformed semantic
        # evidence stops the frozen capture before the next child can run.
        import verify

        verify._validate_report(_read_json(report), _read_json(evidence), mode, lane)
    result = {
        "schema": CAPTURE_SCHEMA,
        "status": "pass",
        "mode": mode,
        "plan_sha256": _sha256(plan_path),
        "records": records,
        "children": len(records),
        "profile": PROFILE,
        "schedule_position": 0 if mode == "native" else 1,
    }
    _write_json(output_root / f"capture-{mode}.json", result)
    return result


def capture(plan_path: Path, output_root: Path) -> dict[str, Any]:
    plan = load_plan(plan_path)
    if output_root.exists():
        raise FileExistsError(f"capture output already exists: {output_root}")
    output_root.mkdir(parents=True)
    lanes = [run_lane(plan_path, output_root, mode) for mode in plan["schedule"]]
    result = {
        "schema": CAPTURE_SCHEMA,
        "status": "pass",
        "plan": str(plan_path.resolve()),
        "plan_sha256": _sha256(plan_path),
        "schedule": plan["schedule"],
        "lanes": lanes,
        "profile": PROFILE,
        "cpu_affinity": CPU_AFFINITY,
        "no_retries": True,
    }
    _write_json(output_root / "capture.json", result)
    return result


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--plan", type=Path, required=True)
    parser.add_argument("--output-root", type=Path, required=True)
    args = parser.parse_args()
    result = capture(args.plan.resolve(), args.output_root.resolve())
    print(json.dumps(result, indent=2, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
