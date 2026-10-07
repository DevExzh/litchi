#!/usr/bin/env python3
"""Relocate the sealed 0952 reader without changing its packet or helpers.

The frozen reader's --packet option is compatibility-only.  This adapter loads
that exact reader from a relocated extraction, maps packet-local reads to the
extraction, and preserves historical descriptor paths in the result.  External
0950/0951 dependencies remain at their declared paths.  It performs no
workload or perf operation.
"""
from __future__ import annotations

import argparse
import importlib.util
import json
import os
import sys
from pathlib import Path
from typing import Any, Callable

HIST = Path("/home/zhuhe/litchi-performance-0952").resolve()
MOVED = Path("/home/zhuhe/litchi-performance-0952-replay-v1").resolve()
DEFAULT_READER = MOVED / "tools/analyze_0952_v3.py"


def under(path: Path, root: Path) -> bool:
    try:
        path.relative_to(root)
        return True
    except ValueError:
        return False


def map_path(value: Path) -> Path:
    """Map only packet-local historical paths; retain external dependencies."""
    path = Path(value)
    if not path.is_absolute():
        path = path.absolute()
    # Do not resolve symlinks here: the sealed packet verifier must see them.
    if under(path, HIST):
        return MOVED / path.relative_to(HIST)
    return path


def historical_path(value: Path) -> str:
    path = Path(value).resolve()
    if under(path, MOVED):
        return str(HIST / path.relative_to(MOVED))
    return str(path)


def load_module(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"cannot load {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def patch_semantic(sem: Any, old_sem_desc: Callable[[Path], dict[str, Any]],
                   old_sem_sha: Callable[[Path], str],
                   old_sem_read: Callable[[Path], dict[str, Any]],
                   old_sem_verify_file: Callable[..., None]) -> Any:
    sem.PACKET = MOVED
    sem.PLAN_PATH = MOVED / "plan-0952-v1.json"
    sem.RUNNER_PATH = MOVED / "tools/capture_0952_v1.py"

    def sem_sha(path: Path) -> str:
        return old_sem_sha(map_path(path))

    def sem_desc(path: Path) -> dict[str, Any]:
        actual = map_path(path)
        out = old_sem_desc(actual)
        out["path"] = historical_path(actual)
        return out

    def sem_read(path: Path) -> dict[str, Any]:
        return old_sem_read(map_path(path))

    def sem_verify_file(path: Path, expected_sha: str, expected_bytes: int, label: str) -> None:
        return old_sem_verify_file(map_path(path), expected_sha, expected_bytes, label)

    def sem_verify_desc(value: Any, path: Path, label: str) -> None:
        sem.check(isinstance(value, dict), f"{label} descriptor is missing")
        actual = sem_desc(map_path(path))
        for key in ("path", "bytes", "sha256"):
            sem.check(value.get(key) == actual[key], f"{label} descriptor {key} differs")

    sem.sha256 = sem_sha
    sem.descriptor = sem_desc
    sem.read_json = sem_read
    sem.verify_file = sem_verify_file
    sem.verify_descriptor = sem_verify_desc
    return sem


def configure(reader: Any) -> Any:
    reader.PACKET = MOVED
    reader.PLAN = MOVED / "plan-0952-v1.json"
    reader.RUNNER = MOVED / "tools/capture_0952_v1.py"
    reader.SEMANTIC = MOVED / "tools/semantic_0952_v1.py"
    reader.SAMPLE_EXPORTER = MOVED / "tools/export_samples_0952_v1.py"
    reader.MMAP_EXPORTER = MOVED / "tools/export_mmaps_0952_v2.py"

    old_sha = reader.sha
    old_desc = reader.descriptor
    old_load_json = reader.load_json

    def relocated_sha(path: Path) -> str:
        return old_sha(map_path(path))

    def relocated_desc(path: Path) -> dict[str, Any]:
        actual = map_path(path)
        out = old_desc(actual)
        out["path"] = historical_path(actual)
        return out

    def relocated_load_json(path: Path) -> Any:
        return old_load_json(map_path(path))

    def relocated_verify_desc(value: Any, label: str) -> Path:
        reader.require(isinstance(value, dict), f"{label} descriptor is missing")
        raw = value.get("path")
        reader.require(isinstance(raw, str) and raw and "\x00" not in raw,
                       f"{label} path is invalid")
        actual = map_path(Path(raw))
        observed = relocated_desc(actual)
        for key in ("path", "bytes", "sha256", "mode", "executable"):
            reader.require(observed[key] == value.get(key),
                           f"{label} descriptor {key} differs")
        return actual.resolve()

    reader.sha = relocated_sha
    reader.descriptor = relocated_desc
    reader.load_json = relocated_load_json
    reader.verify_descriptor = relocated_verify_desc

    def relocated_import_semantic() -> Any:
        reader.require(reader.sha(reader.SEMANTIC) == reader.SEMANTIC_SHA,
                       "semantic adapter digest differs")
        sem = load_module(reader.SEMANTIC, "semantic_0952_relocated")
        old_sem_desc = sem.descriptor
        old_sem_sha = sem.sha256
        old_sem_read = sem.read_json
        old_sem_verify_file = sem.verify_file
        return patch_semantic(sem, old_sem_desc, old_sem_sha, old_sem_read,
                              old_sem_verify_file)

    reader.import_semantic = relocated_import_semantic
    return reader


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--index", type=int, required=True)
    parser.add_argument("--sample-export", type=Path, required=True)
    parser.add_argument("--mmap-export", type=Path, required=True)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--reader", type=Path, default=DEFAULT_READER)
    args = parser.parse_args()
    reader = configure(load_module(args.reader.resolve(), "analyze_0952_relocated"))
    output = args.output.resolve()
    try:
        result, code = reader.run(argparse.Namespace(
            index=args.index,
            sample_export=args.sample_export,
            mmap_export=args.mmap_export,
            output=output,
            packet=MOVED,
        ))
        reader.write_exclusive(output, result)
        print(json.dumps({"status": result["status"], "code": code,
                          "output": str(output), "sha256": reader.sha(output)}, sort_keys=True))
        return code
    except (reader.AdmissionError, OSError, ValueError, KeyError, ImportError, reader.struct.error) as exc:
        failure = {"schema": "litchi.performance.0952.native-admission.v3",
                   "status": "error", "error": str(exc),
                   "scope": "No admission; relocated replay failed closed."}
        try:
            reader.write_exclusive(output, failure)
        except Exception:
            pass
        print(f"replay_0952_v1: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
