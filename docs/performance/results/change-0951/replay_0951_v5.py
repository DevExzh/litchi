#!/usr/bin/env python3
"""0951 relocated replay and real custody negative controls.

This is an additive adapter.  The frozen reducer and all sealed packet files
remain byte unchanged.  The v4 adapter is imported only for its already
reviewed path mapping/replay patch; v4's presentation-only negative controls
are intentionally not called.  This version invokes the patched historical
argv validator and an independent, exact manifest verifier for every negative
case.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import importlib.util
import json
import os
import pathlib
import shutil
import stat
import sys
import tempfile
from typing import Any, Callable

ROOT = pathlib.Path(__file__).resolve().parent
V4_PATH = ROOT / "replay_0951_v4.py"
MANIFEST_PATH = ROOT / "manifest-v1.json"
SCHEMA = "litchi.performance.0951.portable-replay-adapter.v5"


class ReplayError(RuntimeError):
    pass


def require(ok: bool, msg: str) -> None:
    if not ok:
        raise ReplayError(msg)


def sha(path: pathlib.Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1 << 20), b""):
            h.update(block)
    return h.hexdigest()


def descriptor(path: pathlib.Path) -> dict[str, Any]:
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    st = path.stat()
    return {
        "path": str(path),
        "bytes": st.st_size,
        "sha256": sha(path),
        "mode": stat.S_IMODE(st.st_mode),
        "executable": bool(st.st_mode & stat.S_IXUSR),
    }


def load_module(path: pathlib.Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    require(spec is not None and spec.loader is not None, f"cannot import {path}")
    module = importlib.util.module_from_spec(spec)
    sys.modules[name] = module
    spec.loader.exec_module(module)
    return module


def load_json(path: pathlib.Path) -> Any:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    return json.loads(path.read_text(encoding="utf-8"))


def safe_manifest_path(raw: Any) -> pathlib.PurePosixPath:
    require(isinstance(raw, str) and raw, "manifest path is empty")
    require("\x00" not in raw and "\\" not in raw, "manifest path has forbidden byte")
    p = pathlib.PurePosixPath(raw)
    require(not p.is_absolute() and raw != ".", "manifest path is absolute")
    require(all(part not in {"", ".", ".."} for part in p.parts), f"unsafe manifest path: {raw}")
    return p


def verify_manifest(root: pathlib.Path, manifest: dict[str, Any]) -> dict[str, Any]:
    """Verify exact regular-file census, bytes, hashes, modes, and no symlinks."""
    require(root.is_dir() and not root.is_symlink(), f"manifest root invalid: {root}")
    entries = manifest.get("entries")
    require(isinstance(entries, list), "manifest entries missing")
    expected: dict[str, dict[str, Any]] = {}
    for entry in entries:
        require(isinstance(entry, dict), "manifest entry is not an object")
        rel = safe_manifest_path(entry.get("path"))
        key = rel.as_posix()
        require(key not in expected, f"duplicate manifest path: {key}")
        expected[key] = entry

    actual_files: set[str] = set()
    symlinks: list[str] = []
    for dirpath, dirnames, filenames in os.walk(root, topdown=True, followlinks=False):
        base = pathlib.Path(dirpath)
        kept_dirs: list[str] = []
        for name in dirnames:
            item = base / name
            if item.is_symlink():
                symlinks.append(item.relative_to(root).as_posix())
            else:
                kept_dirs.append(name)
        dirnames[:] = kept_dirs
        for name in filenames:
            item = base / name
            rel = item.relative_to(root).as_posix()
            if item.is_symlink():
                symlinks.append(rel)
            else:
                actual_files.add(rel)
    require(not symlinks, f"manifest tree contains symlink(s): {sorted(symlinks)}")
    require(actual_files == set(expected), "manifest regular-file set differs")
    for rel, entry in expected.items():
        item = root / pathlib.PurePosixPath(rel)
        require(item.is_file() and not item.is_symlink(), f"manifest file missing: {rel}")
        st = item.stat()
        require(st.st_size == entry.get("bytes"), f"manifest byte count differs: {rel}")
        require(sha(item) == entry.get("sha256"), f"manifest hash differs: {rel}")
        require(stat.S_IMODE(st.st_mode) == entry.get("mode"), f"manifest mode differs: {rel}")
        require(bool(st.st_mode & stat.S_IXUSR) == bool(entry.get("executable")), f"manifest executable bit differs: {rel}")
    return {
        "status": "pass",
        "regular_files": len(actual_files),
        "logical_bytes": sum(int(entry["bytes"]) for entry in expected.values()),
        "symlinks": [],
    }


def make_hardlink_tree(source: pathlib.Path) -> pathlib.Path:
    target = pathlib.Path(tempfile.mkdtemp(prefix="p951-control-", dir=ROOT))
    for dirpath, _dirnames, filenames in os.walk(source, followlinks=False):
        base = pathlib.Path(dirpath)
        relbase = base.relative_to(source)
        outbase = target / relbase
        outbase.mkdir(parents=True, exist_ok=True)
        for name in filenames:
            src = base / name
            require(src.is_file() and not src.is_symlink(), f"control source is not regular: {src}")
            dst = outbase / name
            os.link(src, dst)
            os.chmod(dst, stat.S_IMODE(src.stat().st_mode))
    return target


def expect_rejection(name: str, validator: str, fn: Callable[[], Any]) -> dict[str, Any]:
    try:
        fn()
    except RuntimeError as error:
        return {
            "status": "rejected",
            "validator": validator,
            "exception": type(error).__name__,
            "message": str(error),
        }
    except Exception as error:  # A different exception is not a valid fail-closed result.
        raise ReplayError(f"{name}: validator raised unexpected {type(error).__name__}: {error}") from error
    raise ReplayError(f"{name}: validator accepted the malformed input")


def negative_controls(v4: Any, reducer: Any, moved: pathlib.Path, receipt: pathlib.Path) -> dict[str, Any]:
    manifest = load_json(MANIFEST_PATH)
    manifest_positive = verify_manifest(moved, manifest)
    plan = load_json(moved / "plan-0951-v1.json")
    folder = moved / "capture" / "00-medium"
    start = load_json(folder / "start.json")
    terminal = load_json(folder / "terminal.json")
    command = load_json(folder / "command.json")
    row = next(item for item in plan["schedule"] if item["index"] == 0)
    controls: dict[str, Any] = {
        "manifest_positive": {"status": "accepted", "validator": "verify_manifest", **manifest_positive},
    }

    tampered_start = copy.deepcopy(start)
    tampered_start["argv"][-1] = str(tampered_start["argv"][-1]) + ".tampered"
    controls["tampered_argv"] = expect_rejection(
        "tampered_argv", "reducer.verify_historical_argv", lambda: reducer.verify_historical_argv(plan, row, folder, tampered_start, terminal, command)
    )

    tampered_start = copy.deepcopy(start)
    tampered_terminal = copy.deepcopy(terminal)
    tampered_command = copy.deepcopy(command)
    bad_cwd = str(v4.HIST / "tampered-source")
    tampered_start["cwd"] = bad_cwd
    tampered_terminal["cwd"] = bad_cwd
    tampered_command["cwd"] = bad_cwd
    controls["tampered_cwd"] = expect_rejection(
        "tampered_cwd", "reducer.verify_historical_argv", lambda: reducer.verify_historical_argv(plan, row, folder, tampered_start, tampered_terminal, tampered_command)
    )

    tampered_start = copy.deepcopy(start)
    tampered_terminal = copy.deepcopy(terminal)
    tampered_command = copy.deepcopy(command)
    for doc in (tampered_start, tampered_terminal, tampered_command):
        doc["argv"] = [str(value).replace("report.json", "report-tampered.json") for value in doc["argv"]]
    controls["tampered_artifact_argv"] = expect_rejection(
        "tampered_artifact_argv", "reducer.verify_historical_argv", lambda: reducer.verify_historical_argv(plan, row, folder, tampered_start, tampered_terminal, tampered_command)
    )

    # Each manifest mutation is checked by the same exact-set/hash/mode
    # verifier used for the positive extracted packet.  Hardlinks keep these
    # controls cheap and cannot alter the sealed extraction or source packet.
    temp_roots: list[pathlib.Path] = []
    try:
        missing_root = make_hardlink_tree(moved)
        temp_roots.append(missing_root)
        missing_file = next(pathlib.PurePosixPath(entry["path"]) for entry in manifest["entries"])
        (missing_root / missing_file).unlink()
        controls["missing_artifact"] = expect_rejection(
            "missing_artifact", "verify_manifest", lambda: verify_manifest(missing_root, manifest)
        )

        extra_root = make_hardlink_tree(moved)
        temp_roots.append(extra_root)
        (extra_root / "unexpected-control-artifact.bin").write_bytes(b"control")
        controls["extra_artifact"] = expect_rejection(
            "extra_artifact", "verify_manifest", lambda: verify_manifest(extra_root, manifest)
        )

        symlink_root = make_hardlink_tree(moved)
        temp_roots.append(symlink_root)
        first_rel = pathlib.PurePosixPath(manifest["entries"][0]["path"])
        second_rel = pathlib.PurePosixPath(manifest["entries"][1]["path"])
        (symlink_root / first_rel).unlink()
        os.symlink(second_rel, symlink_root / first_rel)
        controls["symlink_artifact"] = expect_rejection(
            "symlink_artifact", "verify_manifest", lambda: verify_manifest(symlink_root, manifest)
        )

        escaped = copy.deepcopy(manifest)
        escaped["entries"][0]["path"] = "../outside-control-artifact"
        controls["path_escape"] = expect_rejection(
            "path_escape", "verify_manifest", lambda: verify_manifest(moved, escaped)
        )
    finally:
        for temp_root in temp_roots:
            shutil.rmtree(temp_root, ignore_errors=False)

    result = {
        "schema": SCHEMA + ".negative-controls",
        "status": "pass",
        "controls": controls,
        "validator_contract": "Every rejected case invoked the named validator and required ReplayError; mutation-only assertions are absent.",
        "manifest": descriptor(MANIFEST_PATH),
        "moved_packet": str(moved),
        "historical_root": str(v4.HIST),
        "supersedes": {
            "path": str(ROOT / "replay-results" / "replay-negative-controls-v4.json"),
            "sha256": sha(ROOT / "replay-results" / "replay-negative-controls-v4.json"),
            "flaw": "v4 checked mutated values/under() but did not invoke a historical argv or manifest admission validator.",
        },
    }
    receipt.parent.mkdir(parents=True, exist_ok=True)
    receipt.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return result


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("action", choices=("all", "negative-controls"))
    parser.add_argument("--packet", type=pathlib.Path, required=True)
    parser.add_argument("--output", type=pathlib.Path)
    parser.add_argument("--receipt", type=pathlib.Path)
    args = parser.parse_args()
    moved = args.packet.resolve()
    require(moved.is_dir() and not moved.is_symlink(), "moved packet invalid")
    v4 = load_module(V4_PATH, "litchi_0951_adapter_v4_parent")
    reducer_path = moved / v4.REDUCER_REL
    capture_path = moved / "tools/capture_0951_v1.py"
    require(sha(reducer_path) == "c4074267ec9f320671a34fc2ba31461173a68f1d1a3891e1f9b851d974cf6bc2", "frozen reducer hash differs")
    reducer = v4.load_mod(reducer_path, "litchi_0951_v5_frozen_reducer")
    capture = v4.load_mod(capture_path, "litchi_0951_v5_frozen_capture")
    v4.patch(reducer, capture, moved)
    if args.action == "negative-controls":
        receipt = (args.receipt or ROOT / "replay-results" / "replay-negative-controls-v5.json").resolve()
        result = negative_controls(v4, reducer, moved, receipt)
        print(json.dumps(result, sort_keys=True))
        return 0
    require(args.output is not None, "--output required")
    result = reducer.reduce_packet(args.output.resolve())
    print(json.dumps({"status": result["status"], "output": str(args.output.resolve()), "children": len(result["children"])}, sort_keys=True))
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except (ReplayError, v4.ReplayError if "v4" in globals() else RuntimeError, OSError, ImportError, json.JSONDecodeError, ValueError) as error:
        print(f"replay_0951_v5: {error}", file=sys.stderr)
        raise SystemExit(2)
