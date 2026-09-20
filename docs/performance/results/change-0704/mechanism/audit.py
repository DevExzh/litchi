#!/usr/bin/env python3
"""Audit the bounded 0704 source-only MCE mechanism packet.

This verifier consumes only this packet's fresh build and trace receipts. It
checks the complete 603-file candidate census, exact codec restoration, the
temporary instrumentation hash, standalone probe bindings, all twelve fresh
process traces, and the expected capture/commit call counts. It makes no
timing, allocation, RSS, or throughput claim.
"""

from __future__ import annotations

import hashlib
import json
import runpy
import subprocess
import tempfile
from pathlib import Path
from typing import Any

P = Path(__file__).resolve().parent
PACKET = P.parent
ROOT = P.parents[4]
CODEC = ROOT / "crates" / "litchi-ooxml-common" / "src" / "mce" / "codec.rs"
PATCH = P / "0704-mce-trace.patch"
SOURCE_CENSUS = PACKET / "source-census-candidate.json"
REAL = ROOT / "test-data" / "libreoffice-core" / "sd" / "qa" / "unit" / "data" / "pptx" / "slide-section-test.pptx"
TARGET = ROOT.parent / "litchi-target-0704"
BIN = ROOT.parent / "litchi-0704-bin"
EXPECTED_CODEC = "a5b5b0aca3ec5a392bc7ae1ea6ca482653bd0a72cb9bdd4c30b8ea9faff87bee"
EXPECTED_NAMES = {
    f"{case}-{workflow}-r{repeat}"
    for repeat in range(2)
    for case in ("real", "generated")
    for workflow in ("noop", "one", "two")
}
EXPECTED_PROBE_FILES = {"probe/Cargo.toml", "probe/Cargo.lock", "probe/src/main.rs"}


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def sha(path: Path) -> str:
    return sha_bytes(path.read_bytes())


def need(path: Path) -> Path:
    if not path.is_file():
        raise AssertionError(f"missing mechanism receipt: {path.relative_to(P)}")
    return path


def map_digest(mapping: dict[str, str]) -> str:
    digest = hashlib.sha256()
    for name, value in sorted(mapping.items()):
        digest.update(name.encode())
        digest.update(b"\0")
        digest.update(bytes.fromhex(value))
    return digest.hexdigest()


def current_source_map() -> dict[str, str]:
    result: dict[str, str] = {}
    for owner in ("litchi-pptx", "litchi-ooxml-common", "litchi-opc"):
        for path in (ROOT / "crates" / owner).rglob("*.rs"):
            result[str(path.relative_to(ROOT))] = sha(path)
    return dict(sorted(result.items()))


def verify_source_census() -> dict[str, str]:
    census = json.loads(need(SOURCE_CENSUS).read_text())
    expected = dict(sorted(census["source_sha256"].items()))
    assert census["phase"] == "candidate"
    assert census["source_file_count"] == 603
    assert len(expected) == 603
    current = current_source_map()
    assert current == expected, "current source map differs from candidate census"
    assert expected["crates/litchi-ooxml-common/src/mce/codec.rs"] == EXPECTED_CODEC
    return expected


def verify_build(source_map: dict[str, str]) -> dict[str, Any]:
    build = json.loads(need(P / "build.json").read_text())
    assert build["schema"] == "litchi-0704-mce-mechanism-build-v1"
    assert build["diagnostic_only"] is True
    assert build["performance_claim"] == "none"
    assert build["status"] == "completed"
    assert build["candidate_source_file_count"] == 603
    assert build["candidate_source_sha256"] == source_map
    assert build["candidate_source_map_sha256"] == map_digest(source_map)
    assert build["candidate_census_sha256"] == sha(SOURCE_CENSUS)
    assert build["original_codec_sha256"] == EXPECTED_CODEC
    assert build["restored_codec_sha256"] == EXPECTED_CODEC
    assert sha(CODEC) == EXPECTED_CODEC
    assert build["patch_sha256"] == sha(PATCH)
    assert build["workspace_lock_sha256"] == sha(ROOT / "Cargo.lock")
    assert build["candidate_source_restored"] is True

    for step in build["steps"]:
        assert step["exit_code"] == 0, step
        assert sha(P / f"{step['name']}.log") == step["log_sha256"]
    step_names = {step["name"] for step in build["steps"]}
    assert {"patch-check", "patch-apply", "build-trace"}.issubset(step_names)

    probe_files = {
        str(path.relative_to(P)): sha(path)
        for path in sorted((P / "probe").rglob("*"))
        if path.is_file()
    }
    assert set(probe_files) == EXPECTED_PROBE_FILES
    assert build["probe_sha256"] == probe_files
    manifest = (P / "probe" / "Cargo.toml").read_text()
    assert 'name = "probe0704mce"' in manifest
    assert 'name = "probe0704mce"' in (P / "probe" / "Cargo.lock").read_text()

    binary = Path(build["binary"])
    assert binary == BIN / "probe0704mce"
    if binary.is_file():
        assert sha(binary) == build["binary_sha256"]
    else:
        cleanup = json.loads((PACKET / "cleanup.json").read_text())
        rows = {row["path"]: row for row in cleanup["binaries"]}
        assert rows[str(binary)]["binary_sha256"] == build["binary_sha256"]
    assert TARGET == ROOT.parent / "litchi-target-0704"

    # Verify the recorded temporary instrumentation independently against the
    # restored production source, without touching the shared worktree.
    with tempfile.TemporaryDirectory(prefix="litchi-0704-mechanism-audit-") as directory:
        base = Path(directory)
        target = base / "crates" / "litchi-ooxml-common" / "src" / "mce" / "codec.rs"
        target.parent.mkdir(parents=True)
        target.write_bytes(CODEC.read_bytes())
        subprocess.run(["git", "apply", str(PATCH)], cwd=base, check=True)
        assert sha(target) == build["instrumented_codec_sha256"]
    assert "LITCHI0703" not in PATCH.read_text()
    assert "LITCHI_0703" not in PATCH.read_text()
    return build


def safe_packet_path(name: str) -> Path:
    path = (P / name).resolve()
    try:
        path.relative_to(P.resolve())
    except ValueError as error:
        raise AssertionError(f"trace path escapes mechanism packet: {name}") from error
    return need(path)


def stdout_fields(path: Path) -> tuple[dict[str, str], dict[str, str]]:
    rows = path.read_text().splitlines()
    header = {line.split("\t", 1)[0]: line.split("\t", 1)[1] for line in rows if "\t" in line and not line.startswith("result\t")}
    result_lines = [line for line in rows if line.startswith("result\t")]
    assert len(result_lines) == 1
    result = dict(part.split("=", 1) for part in result_lines[0].split("\t")[1:])
    assert header["probe"] == "0704-mechanism"
    assert header["iterations"] == "1"
    return header, result


def ordered_trace(path: Path, analyzer: dict[str, Any]) -> tuple[list[dict[str, Any]], list[str]]:
    records, _ = analyzer["parse_file"](path)
    boundaries: list[str] = []
    current_phase = "unset"
    for raw in path.read_text().splitlines():
        if raw.startswith("LITCHI0704_BOUNDARY "):
            fields = dict(token.split("=", 1) for token in raw.split()[1:])
            assert fields["event"] == "begin"
            current_phase = fields["phase"]
            boundaries.append(current_phase)
        elif raw.startswith("LITCHI0704_MCE "):
            row = dict(token.split("=", 1) for token in raw.split()[1:])
            assert row["phase"] == current_phase
    return records, boundaries


def verify_one_trace(row: dict[str, Any], analyzer: dict[str, Any]) -> None:
    name = row["name"]
    case = row["case"]
    workflow = row["workflow"]
    stderr = safe_packet_path(row["stderr"])
    stdout = safe_packet_path(row["stdout"])
    header, result = stdout_fields(stdout)
    assert header["source"] == row["source"]
    assert header["workflow"] == workflow
    assert result["workflow"] == workflow
    assert result["iteration"] == "0"
    expected_changed = "false" if workflow == "noop" else "true"
    assert result["changed"] == expected_changed
    assert result["commit_changed"] == expected_changed

    records, boundaries = ordered_trace(stderr, analyzer)
    assert records
    assert all(record["profile"] == "default-ooxml" for record in records)
    assert all(record["status"] == "ok" for record in records)
    assert [record["call"] for record in records] == list(range(len(records)))
    prefix = f"{workflow}.i0."
    expected_stages = ["open", "capture", "clone", "edit", "commit", "apply", "verify"]
    selected_boundaries = [phase[len(prefix):] for phase in boundaries if phase.startswith(prefix)]
    assert selected_boundaries == expected_stages
    assert boundaries[-1] == "done"
    if workflow == "two":
        assert boundaries[:2] == ["setup.target.first", "setup.target.second"]
    else:
        assert boundaries[0] == "setup.target.first"

    capture_count = sum(record["phase"] == prefix + "capture" for record in records)
    commit_count = sum(record["phase"] == prefix + "commit" for record in records)
    expected_capture = 18 if case == "real" else 19
    expected_commit = {
        "real": {"noop": 0, "one": 6, "two": 7},
        "generated": {"noop": 0, "one": 19, "two": 19},
    }[case][workflow]
    assert capture_count == expected_capture, (name, capture_count)
    assert commit_count == expected_commit, (name, commit_count)

    if case == "generated":
        assert all(record["ownership"] == "borrowed" for record in records)
        assert all(record["raw_sha256"] == record["output_sha256"] for record in records)
        assert all(record["raw_len"] == record["output_len"] for record in records)


def normalized_trace(path: Path, analyzer: dict[str, Any]) -> list[tuple[Any, ...]]:
    records, _ = analyzer["parse_file"](path)
    keys = (
        "call", "phase", "profile", "raw_len", "raw_sha256", "output_len",
        "output_capacity", "ownership", "output_sha256", "status", "options",
    )
    return [tuple(record[key] for key in keys) for record in records]


def verify_runs(build: dict[str, Any]) -> None:
    runs = json.loads(need(P / "runs.json").read_text())
    assert len(runs) == 12
    assert {row["name"] for row in runs} == EXPECTED_NAMES
    assert {(row["case"], row["workflow"], row["repeat"]) for row in runs} == {
        (case, workflow, repeat)
        for case in ("real", "generated")
        for workflow in ("noop", "one", "two")
        for repeat in range(2)
    }
    analyzer = runpy.run_path(str(P / "analyze_trace_0704.py"))
    normalized: dict[tuple[str, str, int], list[tuple[Any, ...]]] = {}
    expected_trace_files: set[str] = set()
    for row in runs:
        assert row["exit_code"] == 0
        assert row["binary_sha256"] == build["binary_sha256"]
        assert row["candidate_source_map_sha256"] == build["candidate_source_map_sha256"]
        stdout = safe_packet_path(row["stdout"])
        stderr = safe_packet_path(row["stderr"])
        expected_trace_files.update({row["stdout"], row["stderr"]})
        assert sha(stdout) == row["stdout_sha256"]
        assert sha(stderr) == row["stderr_sha256"]
        if row["case"] == "real":
            assert row["source_archive_sha256"] == sha(REAL)
        else:
            assert row["source_archive_sha256"] is None
        verify_one_trace(row, analyzer)
        key = (row["case"], row["workflow"], row["repeat"])
        normalized[key] = normalized_trace(stderr, analyzer)
    for case in ("real", "generated"):
        for workflow in ("noop", "one", "two"):
            assert normalized[(case, workflow, 0)] == normalized[(case, workflow, 1)]

    trace_dir = P / "trace-runs"
    assert {
        str(path.relative_to(P))
        for path in trace_dir.iterdir()
        if path.is_file() and path.suffix in {".stdout", ".stderr"}
    } == expected_trace_files
    assert not list(trace_dir.glob("*.phase"))
    report = json.loads(need(P / "trace-summary.json").read_text())
    recomputed = {
        "schema": "litchi-0704-mce-trace-report-v1",
        "diagnostic_only": True,
        "timing_evidence": False,
        "runs": [analyzer["summarize"](safe_packet_path(row["stderr"])) for row in runs],
    }
    assert report == json.loads(json.dumps(recomputed))


def main() -> None:
    source_map = verify_source_census()
    build = verify_build(source_map)
    verify_runs(build)
    for path in P.rglob("*.phase"):
        raise AssertionError(f"ephemeral phase file remains: {path}")
    for script in sorted(P.rglob("*.py")):
        compile(script.read_bytes(), str(script), "exec")
    assert not list(P.rglob("__pycache__"))
    print("PASS: 603-file source binding, codec restoration, build receipt, and 12 traces audit")


if __name__ == "__main__":
    main()
