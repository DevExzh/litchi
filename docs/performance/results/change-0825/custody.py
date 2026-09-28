"""Fail-closed custody and report checks for the 0825 matched validation.

This module is deliberately execution-free.  The root agent owns Cargo,
artifact export, benchmark children, and every filesystem mutation.  The
drivers call these helpers before and after each child so a temporary install
of the two archived before files cannot silently change the comparison.
"""

from __future__ import annotations

import hashlib
import json
import os
import subprocess
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TOOL = ROOT / "tools/perf-baseline"
PLAN_PATH = P / "plan.json"
BASE = "032b0e89cdeb93b6a741c1712863b6687baeb6e5"
TARGET = Path("/home/zhuhe/code/litchi-target-0825")
SCRATCH = Path("/home/zhuhe/code/litchi-fs-0825")
ALLOWLIST = (
    "crates/litchi-pptx/src/opened/transaction.rs",
    "crates/litchi-pptx/src/opened/xml.rs",
)
REAL_INPUTS = (
    "test-data/ooxml/docx/documentProperties.docx",
    "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx",
    "test-data/ooxml/pptx/shapes.pptx",
)
UNRELATED = {
    "docs/FORMAT_IMPLEMENTATION_REVIEW.md":
        "bffd00f144c4c1bbb3b0805d21352b40e9ae46f7366c03d6e581ca61b1f27ce5",
    "docs/UNIFIED_OPS_API_DESIGN.md":
        "f5672c38393a2a6c52f028b2a501ddad93766ef974dbdb46cdc3c45e1db3ef6d",
    "matrix-analysis.json":
        "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855",
}
DRIVERS = ("plan.json", "custody.py", "freeze.py", "build.py", "capture.py")
SUPPORT = ("prepare.py", "quality.py", "install_source.py")
RAW_ALLOC = (
    "allocation_calls", "deallocation_calls", "reallocation_calls",
    "failed_allocation_calls", "allocated_bytes", "deallocated_bytes",
    "live_bytes_before", "live_bytes_after", "peak_live_bytes_before",
    "peak_live_bytes_after", "region_peak_live_bytes",
)


def sha(path: Path | str) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            digest.update(block)
    return digest.hexdigest()


def read(path: Path | str) -> Any:
    return json.loads(Path(path).read_text(encoding="utf-8"))


def write(path: Path | str, value: Any) -> None:
    Path(path).write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def artifact(path: Path | str) -> dict[str, Any]:
    path = Path(path)
    assert path.is_file() and not path.is_symlink(), f"missing artifact: {path}"
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha(path)}


def resolve_descriptor(value: Any, *, packet_only: bool = False) -> Path:
    assert isinstance(value, dict) and isinstance(value.get("path"), str)
    path = Path(value["path"])
    if not path.is_absolute():
        path = (P if packet_only else ROOT) / path
    path = path.resolve()
    if packet_only:
        assert path == P.resolve() or P.resolve() in path.parents, (
            f"descriptor escapes packet: {path}"
        )
    return path


def verify_descriptor(value: Any, *, packet_only: bool = False) -> Path:
    path = resolve_descriptor(value, packet_only=packet_only)
    observed = artifact(path)
    assert value == observed, f"stale artifact descriptor: {path}"
    return path


def tracked(prefixes: tuple[str, ...]) -> list[str]:
    raw = subprocess.check_output(["git", "ls-files", "-z", "--", *prefixes], cwd=ROOT)
    return [name for name in raw.decode().split("\0") if name]


def tracked_source_names() -> list[str]:
    return tracked(("crates", "Cargo.toml", "clippy.toml", ".cargo/config.toml", "rust-toolchain.toml"))


def tracked_tool_names() -> list[str]:
    names = tracked(("tools/perf-baseline",))
    assert names, "perf-baseline has no tracked files"
    return names


def source() -> dict[str, Any]:
    return {
        "revision": subprocess.check_output(
            ["git", "rev-parse", "HEAD"], cwd=ROOT, text=True
        ).strip(),
        "files": {name: sha(ROOT / name) for name in tracked_source_names()},
    }


def tool_source() -> dict[str, str]:
    return {name: sha(ROOT / name) for name in tracked_tool_names()}


def plan() -> dict[str, Any]:
    value = read(PLAN_PATH)
    assert value["schema"] == "litchi.performance.0825.plan.v1"
    return value


def candidate_descriptors() -> dict[str, dict[str, Any]]:
    value = plan()["candidate"]
    result = {}
    for leg, key in (("before", "before_archives"), ("after", "after_archives")):
        archives = value[key]
        assert set(archives) == set(ALLOWLIST)
        for production in ALLOWLIST:
            descriptor = archives[production]
            path = ROOT / descriptor["path"]
            assert path.is_file() and not path.is_symlink(), path
            actual = artifact(path)
            assert actual["bytes"] == descriptor["bytes"]
            assert actual["sha256"] == descriptor["sha256"]
            result[f"{leg}:{production}"] = actual
    return result


def candidate_files(leg: str) -> dict[str, str]:
    assert leg in {"before", "after"}
    key = "before_archives" if leg == "before" else "after_archives"
    return {
        production: plan()["candidate"][key][production]["sha256"]
        for production in ALLOWLIST
    }


def assert_leg_source(leg: str, frozen: dict[str, Any] | None = None) -> dict[str, Any]:
    assert leg in {"before", "after"}
    current = source()
    assert current["revision"] == BASE, "Git HEAD changed during the matched run"
    if frozen is None:
        expected_files = dict(current["files"])
        expected_files.update(candidate_files(leg))
    else:
        expected_files = dict(frozen["source"]["files"])
        expected_files.update(candidate_files(leg))
    assert current["files"] == expected_files, (
        f"live production source is not the complete frozen {leg} source"
    )
    return current


def root_inputs() -> dict[str, str]:
    manifest = read(P / "root-inputs.json")
    expected = {
        "Cargo.lock": manifest["Cargo.lock"]["sha256"],
        "rustfmt.toml": manifest["rustfmt.toml"]["sha256"],
        "tools/perf-baseline/Cargo.lock": manifest["tool-Cargo.lock"]["sha256"],
    }
    observed = {
        "Cargo.lock": sha(ROOT / "Cargo.lock"),
        "rustfmt.toml": sha(ROOT / "rustfmt.toml"),
        "tools/perf-baseline/Cargo.lock": sha(TOOL / "Cargo.lock"),
    }
    packet = {
        "Cargo.lock": sha(P / "inputs/root-Cargo.lock"),
        "rustfmt.toml": sha(P / "inputs/rustfmt.toml"),
        "tools/perf-baseline/Cargo.lock": sha(P / "inputs/tool-Cargo.lock"),
    }
    assert observed == expected == packet, "root or perf-baseline lock input changed"
    return observed


def architecture() -> dict[str, str]:
    declared = read(P / "architecture-inputs.json")
    assert isinstance(declared, dict) and len(declared) == 35
    observed = {name: sha(ROOT / name) for name in declared}
    assert observed == declared, "normative architecture input changed"
    return observed


def unrelated() -> dict[str, str]:
    observed = {name: sha(ROOT / name) for name in UNRELATED}
    assert observed == UNRELATED, "unrelated workspace file changed"
    return observed


def corpus() -> dict[str, dict[str, Any]]:
    declared = read(P / "corpus-inputs.json")
    assert set(declared) == set(REAL_INPUTS)
    observed = {}
    for name in REAL_INPUTS:
        path = ROOT / name
        assert path.is_file() and not path.is_symlink()
        value = {"bytes": path.stat().st_size, "sha256": sha(path)}
        assert value == declared[name], f"corpus input changed: {name}"
        observed[name] = value
    return observed


def provenance() -> dict[str, Any]:
    value = read(P / "provenance.json")
    assert value["schema"] == "litchi.performance.0825.provenance.v1"
    reference = value["reference"]
    path = ROOT / reference["path"]
    assert path.is_file() and sha(path) == reference["sha256"]
    assert value["corpus"] == corpus()
    return value


def check_no_overrides() -> None:
    for name in (
        "RUSTUP_TOOLCHAIN", "RUSTFLAGS", "RUSTDOCFLAGS",
        "CARGO_ENCODED_RUSTFLAGS", "CARGO_BUILD_RUSTFLAGS",
        "CARGO_TARGET_X86_64_UNKNOWN_LINUX_GNU_RUSTFLAGS", "RUSTC_WRAPPER",
        "RUSTC_WORKSPACE_WRAPPER",
    ):
        assert not os.environ.get(name), f"inherited build override: {name}"


def plan_cases(value: dict[str, Any] | None = None) -> list[dict[str, Any]]:
    cases = (value or plan())["cases"]
    assert isinstance(cases, list) and len(cases) == 12
    required = {"case", "format", "input", "phase"}
    seen = set()
    for case in cases:
        assert set(case) == required
        assert case["format"] in {"docx", "xlsx", "pptx"}
        assert case["phase"] in {"lifecycle", "edit", "atomic_publish", "counting_publish"}
        assert case["case"] == f"{case['format']}_real_file_ordinary_save_{case['phase']}"
        assert case["input"] in REAL_INPUTS
        assert case["case"] not in seen
        seen.add(case["case"])
    assert len(seen) == 12
    return cases


def assert_quality() -> dict[str, Any]:
    value = read(P / "quality.json")
    assert value["schema"] == "litchi.performance.0825.quality.v1"
    assert value["status"] == "pass" and value["gate_count"] == 6
    assert value["driver_sha256"] == sha(P / "quality.py")
    assert value.get("environment", {}).get("CARGO_TARGET_DIR") == str(TARGET / "quality")
    rows = value["rows"]
    assert isinstance(rows, list) and len(rows) == 6
    assert all(row["exit_code"] == 0 for row in rows)
    for row in rows:
        if isinstance(row.get("log"), dict):
            verify_descriptor(row["log"])
    source_value = value["source"]
    source_path = resolve_descriptor(source_value)
    if isinstance(source_value, dict):
        verify_descriptor(source_value)
    quality_source = read(source_path)
    files = quality_source.get("files", quality_source)
    assert isinstance(files, dict)
    expected = {
        **{
            name: candidate_files("after")[name] if name in ALLOWLIST else sha(ROOT / name)
            for name in tracked_source_names()
        },
        **tool_source(),
        **architecture(),
        **unrelated(),
    }
    for name, digest in expected.items():
        assert files.get(name) == digest, f"quality census omitted or changed {name}"
    assert quality_source.get("revision", BASE) == BASE
    return value


def packet_files(names: tuple[str, ...]) -> dict[str, str]:
    result = {}
    for name in names:
        path = P / name
        assert path.is_file() and not path.is_symlink(), path
        result[name] = sha(path)
    return result


def support_files() -> dict[str, str]:
    return packet_files(tuple(name for name in SUPPORT if (P / name).is_file()))


def stable_inputs(frozen: dict[str, Any]) -> None:
    assert plan()["base"] == BASE
    assert root_inputs() == frozen["root_inputs"]
    assert architecture() == frozen["architecture"]
    assert corpus() == frozen["corpus"]
    assert provenance()["corpus"] == frozen["corpus"]
    assert unrelated() == frozen["unrelated"]
    assert tool_source() == frozen["tool"]
    assert candidate_descriptors() == frozen["candidate_archives"]
    assert packet_files(DRIVERS) == frozen["drivers"]
    assert support_files() == frozen["support"]
    for key in (
        "plan", "quality", "host", "toolchain", "root_input_manifest",
        "architecture_manifest", "corpus_manifest", "lock_parity", "origin",
        "provenance",
    ):
        path = resolve_descriptor(frozen[key])
        assert artifact(path) == frozen[key], f"frozen metadata changed: {key}"


def assert_static() -> None:
    value = plan()
    assert value["base"] == BASE and value["cpu"] == 12
    assert value["target"] == str(TARGET) and value["scratch"] == str(SCRATCH)
    assert tuple(value["source_allowlist"]) == ALLOWLIST
    assert value["candidate"]["attribution"].startswith("This packet validates")
    assert value["binaries"]["native"]["cargo_bin"] == "litchi-perf-baseline"
    assert value["binaries"]["artifacts"]["cargo_bin"] == "ordinary_save_artifacts"
    assert value["binaries"]["observer"]["cargo_bin"] == "litchi-perf-baseline-alloc"
    assert value["binaries"]["observer"]["features"] == [
        "allocator-metrics", "ordinary-save-process-metrics"
    ]
    plan_cases(value)
    origin = read(P / "origin.json")
    assert origin["schema"] == "litchi.performance.0825.origin.v1"
    assert origin["base"] == BASE and origin["target"] == str(TARGET)
    assert origin["scratch"] == str(SCRATCH)
    assert tuple(origin["source_allowlist"]) == ALLOWLIST
    assert origin["production_changed_at_freeze"] is False
    assert origin["runtime_harness_changed"] is False and origin["tool_changed"] is False
    assert origin["unrelated"] == UNRELATED
    host = read(P / "host.json")
    assert host["affinity_selected"] == [12]
    assert host["target"] == str(TARGET) and host["scratch"] == str(SCRATCH)
    assert len(read(P / "architecture-inputs.json")) == 35
    candidate_descriptors()
    root_inputs()
    architecture()
    corpus()
    provenance()
    check_no_overrides()


def make_scratch() -> dict[str, Any]:
    marker = SCRATCH / ".litchi-performance-0825-owned"
    expected = plan()["scratch_marker"]
    if SCRATCH.exists():
        assert SCRATCH.is_dir() and not SCRATCH.is_symlink()
        assert marker.is_file() and marker.read_text(encoding="utf-8") == expected
    else:
        SCRATCH.mkdir(parents=True)
        marker.write_text(expected, encoding="utf-8")
    return artifact(marker)


def assert_rss(path: Path) -> None:
    value = path.read_text(encoding="utf-8").strip()
    assert value.isdigit() and int(value) > 0, f"invalid RSS receipt: {path}"


def require_descriptor(value: Any, expected: Path) -> dict[str, Any]:
    path = resolve_descriptor(value, packet_only=True)
    assert path == expected.resolve()
    observed = artifact(path)
    assert value == observed
    return observed


def require_admission(leg: str) -> dict[str, Any]:
    assert leg in {"before", "after"}
    path = P / plan()["paths"]["artifact_admission"][leg]
    assert path.is_file(), f"artifact admission missing: {path}"
    value = read(path)
    assert value["schema"] == "litchi.performance.0825.artifact-admission.v1"
    assert value["accepted"] is True and value["leg"] == leg
    assert value["plan_sha256"] == sha(P / "plan.json")
    complete_path = P / plan()["paths"]["artifact_complete"][leg]
    complete = require_descriptor(value["artifact_complete"], complete_path)
    assert read(complete_path)["cases"] == 6
    output = P / plan()["paths"]["artifact_output"][leg]
    manifest_path = output / "manifest.json"
    require_descriptor(value["manifest"], manifest_path)
    audit_path = resolve_descriptor(value["audit"], packet_only=True)
    require_descriptor(value["audit"], audit_path)
    audit = read(audit_path)
    assert audit.get("ok") is True and not audit.get("errors")
    assert len(audit.get("cases", [])) == 6
    require_descriptor(value["auditor"], P / "artifact_audit.py")
    preservation_path = resolve_descriptor(value["zip_preservation"], packet_only=True)
    require_descriptor(value["zip_preservation"], preservation_path)
    preservation = read(preservation_path)
    assert len(preservation.get("cases", [])) == 6
    selectors = value.get("selectors")
    cases = plan_cases()
    assert isinstance(selectors, list) and len(selectors) == len(cases)
    expected = {case["case"]: case for case in cases}
    seen = set()
    for selector in selectors:
        assert selector["case"] in expected and selector["case"] not in seen
        case = expected[selector["case"]]
        assert selector["format"] == case["format"] and selector["phase"] == case["phase"]
        assert selector["input"] == case["input"]
        assert selector["source_sha256"] == corpus()[case["input"]]["sha256"]
        assert isinstance(selector.get("published_sha256"), str)
        assert selector.get("edit_outcome") == "admitted"
        seen.add(selector["case"])
    assert seen == set(expected)
    return value


def check_alloc(value: Any, label: str) -> None:
    assert isinstance(value, dict) and value.get("status") == "measured", label
    assert value.get("scope") == "operation_global_system_allocator", label
    for field in RAW_ALLOC:
        assert isinstance(value.get(field), int) and value[field] >= 0, f"{label}.{field}"
    assert value["failed_allocation_calls"] == 0
    assert value["live_bytes_after"] == value["live_bytes_before"] + value["allocated_bytes"] - value["deallocated_bytes"]
    assert value["region_peak_live_bytes"] >= value["live_bytes_before"]
    assert value["region_peak_live_bytes"] >= value["live_bytes_after"]


def validate_report(path: Path, case: dict[str, Any], samples: int, warmup: int,
                    binary_name: str, admission: dict[str, Any]) -> dict[str, Any]:
    report = read(path)
    config = report.get("configuration")
    assert isinstance(config, dict)
    assert config.get("samples_per_case") == samples
    assert config.get("warmup_iterations_per_case") == warmup
    results = report.get("results")
    assert isinstance(results, list) and len(results) == 1
    result = results[0]
    assert result.get("case") == case["case"]
    ordinary = result.get("source", {}).get("ordinary_save")
    assert isinstance(ordinary, dict)
    assert ordinary.get("format") == case["format"].upper()
    assert ordinary.get("origin") == "caller-named-real-file"
    phases = {
        "lifecycle": "open+edit+save",
        "edit": "edit",
        "atomic_publish": "save-to-path",
        "counting_publish": "serialize-to-counting-sink",
    }
    assert ordinary.get("phase") == phases[case["phase"]]
    summary = ordinary.get("corpus")
    assert isinstance(summary, dict)
    real_file = summary.get("real_file")
    assert real_file == {
        "path": str(ROOT / case["input"]),
        **corpus()[case["input"]],
    }
    selector = next(row for row in admission["selectors"] if row["case"] == case["case"])
    assert summary.get("source_archive_sha256") == selector["source_sha256"]
    assert summary.get("edit_admitted") is True
    assert summary.get("edit_outcome") == "admitted"
    assert summary.get("published_sha256") == selector["published_sha256"]
    expected_published = [] if case["phase"] == "edit" else [selector["published_sha256"]]
    assert ordinary.get("published_sha256") == expected_published
    if binary_name == "observer":
        assert report.get("tool", {}).get("binary") == "litchi-perf-baseline-alloc"
        assert ordinary.get("process_probe", {}).get("fixed_count") == 32
        check_alloc(result.get("operation_metrics", {}).get("allocation"), case["case"])
    else:
        assert report.get("tool", {}).get("binary") == "litchi-perf-baseline"
        assert ordinary.get("process_probe") is None
    elapsed = report.get("results", [{}])[0].get("elapsed_ns", {})
    assert isinstance(elapsed.get("samples"), list) and len(elapsed["samples"]) == samples
    return report


def changed_files(left: dict[str, Any], right: dict[str, Any]) -> set[str]:
    a, b = left["files"], right["files"]
    return {name for name in a.keys() | b.keys() if a.get(name) != b.get(name)}
