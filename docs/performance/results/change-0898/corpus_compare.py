#!/usr/bin/env python3
"""Offline, create-only comparison of the 0898 differential reports."""

from __future__ import annotations

import argparse
import filecmp
import gzip
import hashlib
import importlib.util
import json
from pathlib import Path
import shutil
import stat
import sys
from typing import Any


PACKET = Path(__file__).resolve().parent
SCHEMA = "litchi.xml-attribute-bounds.differential.v2"
BASELINE_REVISION = "506aa2edae6666d9c76ae6827c2f5941a8e47896"


def need(ok: bool, message: str) -> None:
    if not ok:
        raise ValueError(message)


def sha(path: Path) -> str:
    value = hashlib.sha256()
    with path.open("rb") as handle:
        for block in iter(lambda: handle.read(1024 * 1024), b""):
            value.update(block)
    return value.hexdigest()


def desc(raw: Path) -> dict[str, Any]:
    raw = Path(raw)
    need(raw.is_file() and not raw.is_symlink(), f"missing regular file: {raw}")
    path = raw.resolve()
    info = path.stat()
    need(stat.S_ISREG(info.st_mode), f"not a regular file: {path}")
    return {"path": str(path), "bytes": info.st_size, "sha256": sha(path),
            "mode": stat.S_IMODE(info.st_mode),
            "executable": bool(info.st_mode & stat.S_IXUSR)}


def read(raw: Path) -> tuple[dict[str, Any], dict[str, Any]]:
    identity = desc(raw)
    path = Path(raw).resolve()
    with path.open("rb") as handle:
        value = json.load(handle)
    need(isinstance(value, dict), f"JSON root is not an object: {path}")
    return value, identity


def oracle_reference(report: dict[str, Any]) -> Any:
    """Run the retained 0889 verifier against the report's frozen source root."""
    roots = report.get("input_inventory", {}).get("roots", [])
    need(isinstance(roots, list) and roots, "report has no input roots")
    test_data = Path(roots[0]).resolve()
    while test_data.name != "test-data" and test_data != test_data.parent:
        test_data = test_data.parent
    need(test_data.name == "test-data", "cannot derive report source root")
    source_root = test_data.parent
    reference_dir = PACKET.parent / "change-0889"
    sys.path.insert(0, str(reference_dir))
    spec = importlib.util.spec_from_file_location(
        "litchi_mce_oracle_reference_0898", reference_dir / "verify_oracles.py")
    need(spec is not None and spec.loader is not None, "cannot load 0889 oracle verifier")
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    module.ROOT = source_root
    module.census.__globals__["ROOT"] = source_root
    return module.check(report, 20000, 20000)


def gzip_copy(source: Path, target: Path) -> dict[str, Any]:
    need(not target.exists() and not target.is_symlink(), f"compressed output exists: {target}")
    target.parent.mkdir(parents=True, exist_ok=True)
    with source.open("rb") as incoming, target.open("xb") as outgoing:
        with gzip.GzipFile(fileobj=outgoing, mode="wb", filename="", mtime=0) as zipped:
            shutil.copyfileobj(incoming, zipped, 1024 * 1024)
    source_id, compressed_id = desc(source), desc(target)
    value, size = hashlib.sha256(), 0
    with gzip.open(target, "rb") as incoming:
        for block in iter(lambda: incoming.read(1024 * 1024), b""):
            size += len(block)
            value.update(block)
    need(size == source_id["bytes"] and value.hexdigest() == source_id["sha256"],
         f"gzip roundtrip mismatch: {target}")
    return {"source": source_id, "compressed": compressed_id, "codec": "gzip",
            "mtime": 0, "filename": "", "roundtrip": True}


def run_receipt(path: Path, label: str) -> dict[str, Any]:
    value, identity = read(path)
    need(value.get("status") == "pass" and value.get("exit_code") == 0,
         f"{label} command did not pass")
    return {"descriptor": identity, "status": value["status"],
            "exit_code": value["exit_code"], "argv": value.get("argv"),
            "log": value.get("log")}


def evidence(path: Path, label: str) -> dict[str, Any]:
    value, identity = read(path)
    need(value.get("status") in {"pass", "approved", "frozen"}, f"{label} receipt status")
    return {"descriptor": identity, "status": value.get("status"),
            "baseline_revision": value.get("baseline_revision"),
            "current_revision": value.get("current_revision"),
            "production_adopted": value.get("production_adopted"),
            "guard_contract": value.get("guard_contract"),
            "source": value.get("source"), "quality": value.get("quality")}


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    for name, default in {
        "baseline": "baseline-differential.json",
        "candidate": "candidate-differential.json",
        "baseline_run": "baseline-differential-run.json",
        "candidate_run": "candidate-differential-run.json",
        "source": "source.json", "candidate_source": "candidate-source.json",
        "baseline_build": "baseline-diagnostic-build-run.json",
        "candidate_build": "candidate-diagnostic-build-run.json",
        "output": "corpus-comparison.json",
        "baseline_gzip": "baseline-differential.json.gz",
        "candidate_gzip": "candidate-differential.json.gz",
    }.items():
        parser.add_argument("--" + name.replace("_", "-"), type=Path, default=PACKET / default)
    args = parser.parse_args()
    output = args.output.resolve()
    need(not output.exists() and not output.is_symlink(), f"comparison output exists: {output}")

    baseline, baseline_id = read(args.baseline)
    candidate, candidate_id = read(args.candidate)
    need(baseline_id["bytes"] == candidate_id["bytes"] and
         baseline_id["sha256"] == candidate_id["sha256"] and
         filecmp.cmp(args.baseline, args.candidate, shallow=False),
         "formal reports are not byte-identical")
    # A byte-identical JSON document compares every oracle, error, output, and
    # report record. Keep this explicit so the receipt records the complete
    # population equality rather than relying only on aggregate counters.
    need(baseline == candidate, "formal oracle/error/output/report records differ")
    baseline_summary = oracle_reference(baseline)
    candidate_summary = oracle_reference(candidate)
    need(baseline_summary == candidate_summary, "formal semantic summaries differ")
    need(baseline_summary["generated"] == 20000 and
         baseline_summary["aliasing"] == 20000 and
         baseline_summary["packages"] == 3 and baseline_summary["members"] == 67,
         "formal population dimensions differ from the frozen corpus")

    compressed = {
        "baseline": gzip_copy(args.baseline, args.baseline_gzip),
        "candidate": gzip_copy(args.candidate, args.candidate_gzip),
    }
    runs = {"baseline": run_receipt(args.baseline_run, "formal baseline"),
            "candidate": run_receipt(args.candidate_run, "candidate")}
    bindings = {name: evidence(path, name) for name, path in {
        "source": args.source, "candidate_source": args.candidate_source,
        "baseline_build": args.baseline_build, "candidate_build": args.candidate_build,
    }.items()}
    source_value, _ = read(args.source)
    need(source_value.get("baseline_revision") == BASELINE_REVISION and
         source_value.get("source_selection_unchanged") is True,
         "source quality baseline is not the unchanged 506aa2 selection")
    result = {
        "schema": "litchi.mce-corpus-comparison-0898.v1", "status": "pass",
        "performance_claim": False, "production_adoption": False,
        "formal_reports": {"baseline": baseline_id, "candidate": candidate_id,
                           "byte_equal": True, "sha256_equal": True,
                           "semantic_equal": True, "records_equal": True,
                           "summary": baseline_summary},
        "compressed_reports": compressed, "command_receipts": runs,
        "source_build_receipts": bindings,
    }
    output.parent.mkdir(parents=True, exist_ok=True)
    with output.open("x", encoding="utf-8") as handle:
        json.dump(result, handle, indent=2, sort_keys=True)
        handle.write("\n")
    print(json.dumps({"status": "pass", "output": str(output),
                      "summary": baseline_summary}, sort_keys=True))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
