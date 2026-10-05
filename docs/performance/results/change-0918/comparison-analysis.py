#!/usr/bin/env python3
"""Replay a relocated frozen 0918 packet with a hash-bound fixture alias.

The frozen comparison runner and replay remain unchanged.  Captured reports
retain the original absolute input path, while a copied repository has a
different local fixture path.  This adapter verifies the copied fixture bytes
against the frozen manifest, derives the original validation path solely from
``repository_root / fixture.path``, and supplies that value through the
unchanged 0917 validator's existing ``path_abs`` metadata.  It never opens the
original alias, rewrites a report, changes a metric, or changes a reduction.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
from pathlib import Path
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
DEFAULT_PACKET = HERE
EXPECTED_REPOSITORY_ROOT = "/home/zhuhe/code/litchi"
EXPECTED_FILES = {
    "comparison.py": (63000,
                       "4ac944350e38239e8148c2884d6c2a4f21082f7a2293159dd0e0636575532f88"),
    "comparison-replay.py": (2871,
                             "4c55800b59d3663889c402dd396276434196583ed5c7c88e1cccd7c0d0d1049b"),
    "comparison-protocol.json": (7319,
                                 "0bfd3669f72b0a491f40c489f9d2cd153275e99add8fd65fba3554e5e50b72de"),
    "comparison-manifest.json": (7161,
                                 "3831f99ea1df275970118bbd6452d9f7a44694b77d971d867e33134862b369b0"),
    "comparison-reduced.json": (397754,
                                "237fa7b7e9d3a4773e2b973f8a2baf7b01d323dd8f1f059af484221ca175824f"),
    "comparison-replay-validator.py": (69575,
                                       "fde5219539cf538d51b8bf92ec5527fd9959ac0f39eb2ea5412b8ddf8cd7a894"),
}
FIXTURES = ("docnohf", "docfloat")


class AnalysisError(RuntimeError):
    """Fail-closed relocated replay or custody error."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AnalysisError(message)


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def read_json(path: Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        raise AnalysisError(f"invalid JSON {path}: {error}") from error


def load_module(name: str, path: Path):
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot load module: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def packet_root(packet: Path) -> Path:
    packet = packet.resolve()
    require(packet.name == "change-0918", f"unexpected packet directory: {packet}")
    try:
        root = packet.parents[3]
    except IndexError as error:
        raise AnalysisError(
            "relocated analysis requires <root>/docs/performance/results/change-0918"
        ) from error
    require(packet == root / "docs/performance/results/change-0918",
            "packet is not under the reconstructed repository hierarchy")
    for required in (
        root / "docs/performance/results/change-0917/capture.py",
        root / "docs/performance/results/change-0917/native-protocol.json",
        root / "docs/performance/results/change-0917/probe",
    ):
        require(required.exists(), f"copied 0917 input missing: {required}")
    return root


def check_hashes(packet: Path) -> dict[str, Path]:
    root = packet_root(packet)
    paths = {
        "comparison.py": packet / "comparison.py",
        "comparison-replay.py": packet / "comparison-replay.py",
        "comparison-protocol.json": packet / "comparison-protocol.json",
        "comparison-manifest.json": packet / "comparison-manifest.json",
        "comparison-reduced.json": packet / "comparison-reduced.json",
        "comparison-replay-validator.py": root / "docs/performance/results/change-0917/capture.py",
    }
    for key, path in paths.items():
        expected_bytes, expected_sha = EXPECTED_FILES[key]
        require(path.is_file() and not path.is_symlink(), f"missing frozen file: {path}")
        require(path.stat().st_size == expected_bytes,
                f"frozen {key} byte count changed")
        require(sha256(path) == expected_sha, f"frozen {key} SHA-256 changed")
    return paths


def fixture_aliases(packet: Path, manifest: dict[str, Any]) -> dict[str, str]:
    root = packet_root(packet)
    original_root = Path(str(manifest.get("repository_root")))
    require(original_root.is_absolute() and str(original_root) == EXPECTED_REPOSITORY_ROOT,
            "manifest repository_root is not the frozen source root")
    rows = manifest.get("fixtures")
    require(isinstance(rows, dict) and set(rows) == set(FIXTURES),
            "manifest fixture identities changed")
    aliases: dict[str, str] = {}
    for case in FIXTURES:
        row = rows[case]
        require(isinstance(row, dict), f"manifest fixture row malformed: {case}")
        relative = Path(str(row.get("path")))
        require(not relative.is_absolute() and ".." not in relative.parts,
                f"manifest fixture path is unsafe: {case}")
        local = root / relative
        require(local.is_file() and not local.is_symlink(), f"missing copied fixture: {local}")
        require(local.stat().st_size == row.get("bytes") and sha256(local) == row.get("sha256"),
                f"copied fixture hash changed: {case}")
        # This is validation metadata only.  It is never opened or hashed.
        aliases[case] = str(original_root / relative)
    return aliases


def load_frozen(packet: Path):
    paths = check_hashes(packet)
    manifest = read_json(paths["comparison-manifest.json"])
    require(isinstance(manifest, dict)
            and manifest.get("schema") == "litchi.performance.0918.doc-comparison.manifest.v1",
            "frozen comparison manifest schema changed")
    require(manifest.get("repository_root") == EXPECTED_REPOSITORY_ROOT,
            "frozen manifest repository_root changed")
    runner = load_module("litchi_perf_0918_frozen_runner", paths["comparison.py"])
    return runner, manifest, paths


def report_from_bytes(raw: bytes, label: str) -> dict[str, Any]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise AnalysisError(f"{label}: stdout is not UTF-8: {error}") from error
    for line in reversed([line.strip() for line in text.splitlines() if line.strip()]):
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise AnalysisError(f"{label}: stdout has no JSON report")


def canonical_report_audit(runner: Any, packet: Path, aliases: dict[str, str]) -> dict[str, object]:
    rows: list[str] = []
    input_rows = 0
    protocol = runner.protocol()
    for spec in runner.all_schedule(protocol):
        label = str(spec["label"])
        receipt_path = packet / "comparison-logs" / f"{label}.json"
        receipt = read_json(receipt_path)
        raw = runner.read_artifact_bytes(receipt["stdout"], f"{label}: stdout")
        report = report_from_bytes(raw, label)
        expected_input = aliases[str(spec["case"])]
        require(report.get("input") == expected_input,
                f"{label}: report input is not the authorized frozen alias")
        report_hash = hashlib.sha256(json_bytes(report)).hexdigest()
        summary = receipt.get("report")
        require(isinstance(summary, dict) and summary.get("report_sha256") == report_hash,
                f"{label}: canonical report hash changed")
        if receipt.get("report_canonical_sha256") is not None:
            require(receipt["report_canonical_sha256"] == report_hash,
                    f"{label}: receipt canonical report hash changed")
        rows.append(f"{label}:{report_hash}")
        input_rows += 1
    require(len(rows) == 288, f"canonical report count is {len(rows)}, expected 288")
    aggregate = hashlib.sha256("\n".join(rows).encode("utf-8")).hexdigest()
    return {"reports": len(rows), "authorized_input_rows": input_rows,
            "canonical_sha256": aggregate}


def replay(packet: Path) -> dict[str, object]:
    runner, manifest, paths = load_frozen(packet)
    aliases = fixture_aliases(packet, manifest)
    validator = runner.validator_module(runner.VALIDATOR_PATH)
    validator_modules = {validator, runner.validator_module()}
    original_fixture_rows = validator.fixture_rows

    def fixture_rows_with_aliases(protocol: dict[str, Any]) -> dict[str, dict[str, Any]]:
        rows = original_fixture_rows(protocol)
        result = {case: dict(row) for case, row in rows.items()}
        require(set(result) == set(FIXTURES), "validator fixture set changed")
        for case, alias in aliases.items():
            result[case]["path_abs"] = alias
        return result

    for module in validator_modules:
        module.fixture_rows = fixture_rows_with_aliases
    try:
        result = runner.replay(argparse.Namespace(reduced=None))
        canonical = canonical_report_audit(runner, packet, aliases)
    finally:
        for module in validator_modules:
            module.fixture_rows = original_fixture_rows
    return {
        "status": "verified",
        "packet": str(packet.resolve()),
        "repository_root": str(packet_root(packet)),
        "fixture_hashes_verified": sorted(aliases),
        "original_fixture_aliases_validation_only": True,
        "frozen_hashes": {key: value[1] for key, value in EXPECTED_FILES.items()},
        "canonical_reports": canonical,
        **result,
    }


def fixture_control(path: Path, name: str) -> dict[str, object]:
    packet = path.resolve() / "docs/performance/results/change-0918"
    try:
        runner, manifest, _ = load_frozen(packet)
        fixture_aliases(packet, manifest)
    except Exception as error:
        text = f"{type(error).__name__}: {error}"
        return {"name": name, "status": "rejected", "packet": str(packet),
                "error": text}
    raise AnalysisError(f"fixture corruption control unexpectedly passed: {name}")


def write_report(path: Path, value: object) -> None:
    data = json_bytes(value)
    require(len(data) < (1 << 20), "analysis report is at least 1 MiB")
    if path.exists():
        require(path.is_file() and not path.is_symlink(),
                f"retained analysis report is not a regular file: {path}")
        require(path.read_bytes() == data,
                f"retained analysis report differs; select a new output path: {path}")
        return
    path.write_bytes(data)


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=DEFAULT_PACKET)
    parser.add_argument("--controls-root", type=Path,
                        help="external root containing path-corrupt/ and bytes-corrupt/ copies")
    parser.add_argument("--output", type=Path,
                        help="optional exclusive JSON report path; stdout is the default")
    args = parser.parse_args()
    try:
        positive = replay(args.packet.resolve())
        controls: list[dict[str, object]] = []
        if args.controls_root is not None:
            root = args.controls_root.resolve()
            controls = [fixture_control(root / "path-corrupt", "fixture-path"),
                        fixture_control(root / "bytes-corrupt", "fixture-bytes")]
            require(all(row["status"] == "rejected" for row in controls),
                    "fixture corruption control did not reject")
        report = {
            "schema": "litchi.performance.0918.comparison-analysis.v1",
            "status": "verified",
            "positive_replay": positive,
            "fixture_corruption_controls": controls,
            "commands": {
                "replay": "PYTHONDONTWRITEBYTECODE=1 python3 -B "
                           "docs/performance/results/change-0918/comparison-analysis.py",
                "relocated": "PYTHONDONTWRITEBYTECODE=1 python3 -B "
                             "docs/performance/results/change-0918/comparison-analysis.py "
                             "--packet <copied-packet> --controls-root <control-root>",
            },
            "raw_report_policy": "raw reports and metrics are read-only; canonical hashes are recomputed",
        }
        output = None
        if args.output is not None:
            output = str(args.output.resolve())
            write_report(args.output.resolve(), report)
        print(json.dumps({"status": "verified", "reports": positive["canonical_reports"]["reports"],
                          "adoption": positive["adoption"], "output": output,
                          "fixture_controls": len(controls)}, sort_keys=True))
    except (AnalysisError, OSError, KeyError, TypeError, ValueError) as error:
        print(f"0918 comparison analysis failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
