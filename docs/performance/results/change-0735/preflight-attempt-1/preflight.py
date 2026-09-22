#!/usr/bin/env python3
"""Run the 0735 analyzer and audit on a synthetic two-case matrix.

This is a schema check, not a measurement.  The packet is copied to an
isolated temporary directory, the four retained baseline qualification
reports are used as templates, and their samples are installed under every
baseline/candidate schedule cell.  No Cargo command, native binary, or
profiler is started here.
"""

from __future__ import annotations

import contextlib
import copy
import hashlib
import importlib.util
import io
import json
import shutil
import tempfile
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
NATIVE = (50, 3)
ALLOCATION = (1, 0)


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def write(path: Path, value: Any) -> None:
    path.write_text(json.dumps(value, indent=2) + "\n", encoding="utf-8")


def copy_packet(target: Path) -> None:
    """Copy packet inputs while excluding real captures and Python caches."""
    target.mkdir(parents=True, exist_ok=True)
    excluded = {"__pycache__", "captures"}
    for child in P.iterdir():
        if child.name in excluded or child.name.startswith("preflight-attempt-"):
            continue
        destination = target / child.name
        if child.is_dir():
            shutil.copytree(
                child,
                destination,
                ignore=shutil.ignore_patterns("__pycache__", "captures"),
            )
        elif child.is_file():
            shutil.copy2(child, destination)
    (target / "captures").mkdir()


def replace_prefix(value: Any, original: str, temporary: str) -> Any:
    """Replace only the packet's absolute prefix in nested JSON values."""
    if isinstance(value, str):
        if value == original:
            return temporary
        prefix = original + "/"
        if value.startswith(prefix):
            return temporary + value[len(original):]
        return value
    if isinstance(value, list):
        return [replace_prefix(item, original, temporary) for item in value]
    if isinstance(value, dict):
        return {
            key: replace_prefix(item, original, temporary)
            for key, item in value.items()
        }
    return value


def rebase_json_paths(packet: Path) -> None:
    """Rebase absolute packet paths in copied quality/build receipts.

    The measured workspace fixture and the four retained binary paths stay at
    their original absolute paths.  Only strings beginning with the real
    packet root are moved into the isolated copy.
    """
    original = str(P)
    temporary = str(packet)
    for path in packet.rglob("*.json"):
        if path.name in {"freeze.json", "preflight.json"}:
            # Freeze has path keys and is handled below.  The receipt is made
            # after rebasing.  Build/quality manifests are intentionally
            # included: their command arrays contain absolute probe paths.
            continue
        try:
            value = read(path)
        except (OSError, UnicodeDecodeError, json.JSONDecodeError):
            continue
        rebased = replace_prefix(value, original, temporary)
        if rebased != value:
            write(path, rebased)


def refresh_build_and_qualification(packet: Path) -> None:
    """Refresh hashes changed by quality-manifest path rebasing."""
    for build_path in sorted(packet.glob("*-build.json")):
        build = read(build_path)
        quality_name = build.get("quality")
        if isinstance(quality_name, str):
            quality_path = packet / quality_name
            if quality_path.is_file() and "quality_sha256" in build:
                build["quality_sha256"] = sha(quality_path)
        write(build_path, build)

    qualification_path = packet / "qualification.json"
    if qualification_path.is_file():
        qualification = read(qualification_path)
        files = qualification.get("files")
        if isinstance(files, dict):
            for raw in list(files):
                path = packet / raw
                if path.is_file():
                    files[raw] = sha(path)
        builds = qualification.get("builds")
        if isinstance(builds, dict):
            for variant in list(builds):
                path = packet / f"{variant}-build.json"
                if path.is_file():
                    builds[variant] = sha(path)
        if "build_sha256" in qualification and (packet / "build.json").is_file():
            qualification["build_sha256"] = sha(packet / "build.json")
        write(qualification_path, qualification)

    before_path = packet / "before.json"
    baseline_path = packet / "baseline-build.json"
    if before_path.is_file() and baseline_path.is_file():
        before = read(before_path)
        if "build_sha256" in before:
            before["build_sha256"] = sha(baseline_path)
            write(before_path, before)


def freeze_bindings(packet: Path) -> tuple[dict[str, str], bool]:
    frozen = read(packet / "freeze.json")
    if isinstance(frozen.get("bindings"), dict):
        return frozen["bindings"], True
    return frozen, False


def rebase_freeze(packet: Path) -> None:
    """Move copied packet keys and recompute hashes for copied targets."""
    bindings, wrapped = freeze_bindings(packet)
    original = str(P)
    temporary = str(packet)
    rebased: dict[str, str] = {}
    for raw, expected in bindings.items():
        key = replace_prefix(raw, original, temporary)
        target = Path(key)
        # Workspace fixtures and binary paths are intentionally retained at
        # their original paths.  Copied packet paths get the copied digest.
        rebased[key] = sha(target) if target.is_file() else expected
    frozen = read(packet / "freeze.json")
    if wrapped:
        frozen["bindings"] = rebased
    else:
        frozen = rebased
    write(packet / "freeze.json", frozen)


def rebase_packet(packet: Path) -> None:
    rebase_json_paths(packet)
    refresh_build_and_qualification(packet)
    # Build/qualification edits above are themselves frozen inputs.  Rebuild
    # the copied freeze map only after all dependent hashes are final.
    rebase_freeze(packet)


def cases(packet: Path) -> dict[str, dict[str, Any]]:
    rows = read(packet / "cases.json")
    assert isinstance(rows, list) and len(rows) == 2
    result: dict[str, dict[str, Any]] = {}
    for row in rows:
        assert isinstance(row, dict)
        ident = row.get("id")
        assert ident in {"primary", "secondary"} and ident not in result
        result[ident] = row
    assert set(result) == {"primary", "secondary"}
    return result


def qualification_template(packet: Path, case_id: str, lane: str) -> dict[str, Any]:
    """Load the retained baseline schema for one case and one lane."""
    path = packet / "qualification-raw" / f"baseline-{case_id}-{lane}.json"
    if not path.is_file():
        # The qualification script retains the failed secondary attempts with
        # an ordinal suffix, while the selected successful report has the
        # stable ``baseline-secondary-native`` name.
        if case_id == "secondary" and lane == "native":
            candidates = sorted((packet / "qualification-raw").glob("baseline-secondary-*-native.json"))
            candidates = [p for p in candidates if "receipt" not in p.name]
            assert candidates, "missing selected secondary native qualification"
            path = candidates[-1]
        else:
            raise AssertionError(f"missing qualification schema: {path}")
    report = read(path)
    samples = report.get("samples")
    assert isinstance(samples, list) and samples
    expected_count, expected_warmups = NATIVE if lane == "native" else ALLOCATION
    assert report.get("samples_requested") == expected_count
    assert report.get("warmups") == expected_warmups
    assert report.get("case") == cases(packet)[case_id]["case"]
    assert report.get("allocator_instrumented") is (lane == "allocation")
    return report


def clone_samples(report: dict[str, Any], lane: str) -> dict[str, Any]:
    expected_count, expected_warmups = NATIVE if lane == "native" else ALLOCATION
    synthetic = copy.deepcopy(report)
    source_samples = report["samples"]
    if len(source_samples) == expected_count:
        samples = copy.deepcopy(source_samples)
    else:
        samples = [copy.deepcopy(source_samples[0]) for _ in range(expected_count)]
    for index, sample in enumerate(samples):
        assert isinstance(sample, dict)
        sample["index"] = index
    synthetic["samples_requested"] = expected_count
    synthetic["warmups"] = expected_warmups
    synthetic["samples"] = samples
    return synthetic


def binary_path(packet: Path, variant: str, lane: str) -> str:
    build = read(packet / f"{variant}-build.json")
    binary = build["binaries"][lane]
    assert isinstance(binary.get("path"), str)
    return binary["path"]


def expected_command(
    packet: Path,
    plan: dict[str, Any],
    case: dict[str, Any],
    spec: dict[str, Any],
) -> list[str]:
    lane = spec["lane"]
    samples, warmups = NATIVE if lane == "native" else ALLOCATION
    return [
        "taskset", "-c", str(plan["cpu"]),
        binary_path(packet, spec["variant"], lane),
        "--case", case["case"], "--input", case["path"],
        "--operation", "format", "--samples", str(samples),
        "--warmups", str(warmups),
    ]


def synthetic_captures(packet: Path) -> None:
    plan = read(packet / "plan.json")
    case_rows = cases(packet)
    templates = {
        (lane, ident): qualification_template(packet, ident, lane)
        for lane in ("native", "allocation")
        for ident in case_rows
    }
    captures = packet / "captures"
    runs: list[dict[str, Any]] = []
    schedule = plan.get("schedule")
    assert isinstance(schedule, list) and len(schedule) == 48
    for spec in schedule:
        assert spec["lane"] in {"native", "allocation"}
        assert spec["variant"] in {"baseline", "candidate"}
        ident = spec["case"]
        report = clone_samples(templates[(spec["lane"], ident)], spec["lane"])
        name = (
            f"{spec['lane']}-c{spec['cycle']}-r{spec['repeat']}"
            f"-{ident}-{spec['variant']}.json"
        )
        output = captures / name
        write(output, report)
        stderr = captures / (name + ".stderr")
        stderr.write_bytes(b"")
        runs.append({
            **spec,
            "command": expected_command(packet, plan, case_rows[ident], spec),
            "exit_code": 0,
            "seconds": 0.0,
            "output": name,
            "sha256": sha(output),
            "stderr_sha256": sha(stderr),
        })

    write(captures / "manifest.json", {
        "status": "complete",
        "mode": "synthetic-schema-only",
        "freeze_sha256": sha(packet / "freeze.json"),
        "preflight_sha256": sha(packet / "preflight.json"),
        "runs": runs,
    })
    assert len(runs) == 48
    assert sum(row["lane"] == "native" for row in runs) == 36
    assert sum(row["lane"] == "allocation" for row in runs) == 12


def synthetic_receipt(packet: Path) -> None:
    scripts = {}
    for name in ("preflight.py", "analyze.py", "audit.py"):
        path = packet / name
        assert path.is_file(), f"missing preflight dependency: {name}"
        scripts[name] = sha(path)
    write(packet / "preflight.json", {
        "status": "passed",
        "kind": "synthetic-schema-only",
        "freeze_sha256": sha(packet / "freeze.json"),
        "scripts": scripts,
        "process_count": 48,
        "native_process_count": 36,
        "allocation_process_count": 12,
        "synthetic_capture_count": 48,
    })


def load_module(name: str, path: Path) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def configure(module: Any, packet: Path) -> None:
    for name in ("P", "PACKET", "ROOT"):
        if hasattr(module, name):
            setattr(module, name, packet if name != "ROOT" else ROOT)
    if hasattr(module, "CAPTURES"):
        module.CAPTURES = packet / "captures"


def run_validators(packet: Path) -> None:
    analyzer = load_module("change0735_preflight_analyzer", packet / "analyze.py")
    audit = load_module("change0735_preflight_audit", packet / "audit.py")
    configure(analyzer, packet)
    configure(audit, packet)
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        analyzer.main()
    with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
        audit.main()
    analysis_path = packet / "analysis.json"
    assert analysis_path.is_file(), "analyzer did not write analysis.json"
    analysis = read(analysis_path)
    assert analysis.get("status") == "passed"
    assert analysis.get("process_count", 48) == 48


def main() -> None:
    required = ("freeze.json", "candidate-build.json", "qualification.json",
                "analyze.py", "audit.py")
    for name in required:
        assert (P / name).exists(), f"{name} is required before preflight"
    original_freeze_sha = sha(P / "freeze.json")
    with tempfile.TemporaryDirectory(prefix="litchi-0735-preflight-") as temporary:
        packet = Path(temporary) / "isolated" / "results" / "change-0735"
        copy_packet(packet)
        rebase_packet(packet)
        synthetic_receipt(packet)
        synthetic_captures(packet)
        run_validators(packet)

    receipt = {
        "status": "passed",
        "kind": "synthetic-schema-only",
        "freeze_sha256": original_freeze_sha,
        "scripts": {
            name: sha(P / name) for name in ("preflight.py", "analyze.py", "audit.py")
        },
        "script_sha256": sha(P / "preflight.py"),
        "analyzer_sha256": sha(P / "analyze.py"),
        "audit_sha256": sha(P / "audit.py"),
        "process_count": 48,
        "native_process_count": 36,
        "allocation_process_count": 12,
        "synthetic_capture_count": 48,
    }
    write(P / "preflight.json", receipt)
    print("PASS synthetic 48-process two-case analyzer/audit schema integration; no measurement evidence")


if __name__ == "__main__":
    main()
