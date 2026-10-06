#!/usr/bin/env python3
"""Offline replay for a hash-bound 0922 DOC comparison packet.

Replay accepts either the repository packet or a copied external packet when
``LITCHI_PERF_0922_REPO_ROOT`` identifies the actual checkout. Fixture bytes
are verified locally, while the original source-root fixture paths are used
only as validator metadata so retained reports remain auditable after
relocation. No report, metric, or reduction is rewritten.
"""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import sys
from typing import Any


class ReplayError(RuntimeError):
    pass


FIXTURES = ("docnohf", "docfloat")


def require(condition: bool, message: str) -> None:
    if not condition:
        raise ReplayError(message)


def sha256(path: Path) -> str:
    require(path.is_file() and not path.is_symlink(), f"missing regular file: {path}")
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def load_runner(packet: Path):
    script = packet / "comparison.py"
    spec = importlib.util.spec_from_file_location("litchi_perf_0922_comparison", script)
    require(spec is not None and spec.loader is not None, f"cannot load runner: {script}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def require_reconstructed_hierarchy(packet: Path) -> Path:
    packet = packet.resolve()
    require(packet.name == "change-0922",
            f"replay requires a change-0922 packet, got {packet}")
    repository_override = os.environ.get("LITCHI_PERF_0922_REPO_ROOT")
    if repository_override:
        repository = Path(repository_override).resolve()
        require(repository.is_dir(), f"repository root is missing: {repository}")
        require(packet == repository / "docs/performance/results/change-0922"
                or packet.is_dir(),
                "external packet path is missing")
    else:
        try:
            repository = packet.parents[3]
        except IndexError as error:
            raise ReplayError(
                "replay requires <root>/docs/performance/results/change-0922 or "
                "LITCHI_PERF_0922_REPO_ROOT for an external packet"
            ) from error
        require(packet == repository / "docs/performance/results/change-0922",
                "external packet replay requires LITCHI_PERF_0922_REPO_ROOT")
    required = (
        repository / "docs/performance/results/change-0917/capture.py",
        repository / "docs/performance/results/change-0917/probe",
        repository / "test-data/ole/doc/NoHeadFoot.doc",
        repository / "test-data/ole/doc/FloatingPictures.doc",
    )
    missing = [str(path) for path in required if not path.exists()]
    require(not missing, "reconstructed replay inputs are missing: " + ", ".join(missing))
    return repository


def fixture_aliases(packet: Path, manifest: dict[str, Any]) -> dict[str, str]:
    repository = require_reconstructed_hierarchy(packet)
    original_root = Path(str(manifest.get("repository_root")))
    require(original_root.is_absolute(), "manifest repository_root is not absolute")
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
        local = repository / relative
        require(local.is_file() and sha256(local) == row.get("sha256")
                and local.stat().st_size == row.get("bytes"),
                f"copied fixture hash changed: {case}")
        aliases[case] = str(original_root / relative)
    return aliases


def report_from_bytes(raw: bytes, label: str) -> dict[str, Any]:
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as error:
        raise ReplayError(f"{label}: stdout is not UTF-8: {error}") from error
    for line in reversed([line.strip() for line in text.splitlines() if line.strip()]):
        try:
            value = json.loads(line)
        except json.JSONDecodeError:
            continue
        if isinstance(value, dict):
            return value
    raise ReplayError(f"{label}: stdout has no JSON report")


def canonical_report_audit(runner: Any, packet: Path,
                           aliases: dict[str, str]) -> dict[str, object]:
    rows: list[str] = []
    for spec in runner.all_schedule(runner.protocol()):
        label = str(spec["label"])
        receipt = runner.read_json(runner.receipt_path(spec))
        raw = runner.read_artifact_bytes(receipt["stdout"], f"{label}: stdout")
        report = report_from_bytes(raw, label)
        require(report.get("input") == aliases[str(spec["case"])],
                f"{label}: report input is not the authorized frozen alias")
        report_hash = hashlib.sha256(json_bytes(report)).hexdigest()
        summary = receipt.get("report")
        require(isinstance(summary, dict) and summary.get("report_sha256") == report_hash,
                f"{label}: canonical report hash changed")
        require(receipt.get("report_canonical_sha256") == report_hash,
                f"{label}: receipt canonical report hash changed")
        rows.append(f"{label}:{report_hash}")
    require(len(rows) == 288, f"canonical report count is {len(rows)}, expected 288")
    return {"reports": len(rows),
            "canonical_sha256": hashlib.sha256("\n".join(rows).encode()).hexdigest()}


def replay(packet: Path) -> dict[str, object]:
    packet = packet.resolve()
    repository = require_reconstructed_hierarchy(packet)
    # The runner keeps packet outputs and repository inputs separate.  Set both
    # roots before importing a copied external runner so a scratch replay never
    # resolves its validator/probe against the scratch directory's parents.
    os.environ["LITCHI_PERF_0922_PACKET_ROOT"] = str(packet)
    os.environ["LITCHI_PERF_0922_REPO_ROOT"] = str(repository)
    runner = load_runner(packet)
    manifest = runner.load_manifest()
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
        "packet": str(packet),
        "fixture_hashes_verified": sorted(aliases),
        "original_fixture_aliases_validation_only": True,
        "canonical_reports": canonical,
        **result,
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--packet", type=Path, default=Path(__file__).resolve().parent)
    args = parser.parse_args()
    try:
        print(json.dumps(replay(args.packet), sort_keys=True))
    except Exception as error:
        print(f"0922 replay failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
