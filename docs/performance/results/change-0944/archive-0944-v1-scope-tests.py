#!/usr/bin/env python3
"""Synthetic selector and cleanup checks for archive-0944-v1.py.

The tests never invoke archive, verify, or cleanup against the live 0944 root.
They construct a temporary evidence root and exercise closure binding, target
selection, redundant-source exclusion, and post-cleanup preservation.
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import tempfile


SCRIPT = Path(__file__).with_name("archive-0944-v1.py")


def load_script():
    spec = importlib.util.spec_from_file_location("archive_0944_under_test", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def put(root: Path, relative: str, data: bytes = b"x") -> Path:
    path = root / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def bind_module(module, root: Path, work: Path) -> None:
    module.SOURCE_ROOT = root.resolve()
    module.WORK_ROOT = work.resolve()
    module.SCRIPT = SCRIPT.resolve()
    module.ARCHIVE = work / "archive.tar.zst"
    module.MANIFEST = work / "manifest.jsonl"
    module.MEMBERS = work / "members.list"
    module.ARCHIVE_RECEIPT = work / "archive-receipt.json"
    module.VERIFY_RECEIPT = work / "verify-receipt.json"
    module.CLEANUP_RECEIPT = work / "cleanup-receipt.json"
    module.EXTRACT_ROOT = work / "extract"
    module.ARCHIVE_STDOUT = work / "archive.stdout.log"
    module.ARCHIVE_STDERR = work / "archive.stderr.log"
    module.VERIFY_STDOUT = work / "verify.stdout.log"
    module.VERIFY_STDERR = work / "verify.stderr.log"
    module.BINARY_DESCRIPTOR = work / "binary-descriptor.json"
    module.ROOT_NAME = root.name
    module.BOUND_TARGET_FILES = frozenset()


def create_bound_root(module, root: Path, work: Path) -> tuple[Path, Path]:
    root.mkdir(parents=True)
    work.mkdir(parents=True)
    bind_module(module, root, work)

    # The complete export is intentionally redundant but has a valid receipt
    # and hash-bound manifest, which authorizes excluding only top-level source.
    put(root, "source/full-export-marker.txt", b"full")
    put(root, "source/test-data/target-fuzz-seed", b"nested-full")
    full_manifest_value = {
        "schema": "litchi-source-manifest-v1",
        "commit": "commit-synthetic",
        "tree": "tree-synthetic",
        "created_utc": "2026-10-07T00:00:00Z",
        "source_root": "source",
        "file_count": 1,
        "files": [{"path": "full-export-marker.txt", "bytes": 4, "sha256": hashlib.sha256(b"full").hexdigest()}],
    }
    full_manifest = root / "source-manifest.json"
    write_json(full_manifest, full_manifest_value)
    put(root, "source-manifest.sha256", (sha(full_manifest) + "  source-manifest.json\n").encode())
    write_json(root / "export-receipt.json", {
        "schema": "litchi-xlsx-profile-source-export-v1",
        "commit": "commit-synthetic",
        "tree": "tree-synthetic",
        "source_root": "source",
        "manifest": "source-manifest.json",
        "manifest_sha256": sha(full_manifest),
        "file_count": 1,
    })

    # The selected build closure includes a nested target-* seed.  It must
    # survive both selection and cleanup because only root target-native is a
    # cache under the production policy.
    put(root, "source-build/Cargo.toml", b"[workspace]\n")
    put(root, "source-build/test-data/nested/target-seed.bin", b"seed")
    put(root, "source-build/tools/raw-tooling.log", b"tool")
    selected_root = root / "source-build"
    selected_records = []
    for path in sorted(p for p in selected_root.rglob("*") if p.is_file()):
        relative = path.relative_to(selected_root).as_posix()
        selected_records.append({"path": relative, "bytes": path.stat().st_size, "sha256": sha(path)})
    selected_manifest = root / "source-build-manifest.json"
    write_json(selected_manifest, {
        "schema": "litchi-xlsx-profile-source-closure-v1",
        "created_utc": "2026-10-07T00:00:00Z",
        "origin_commit": "commit-synthetic",
        "origin_tree": "tree-synthetic",
        "full_export_manifest": "source-manifest.json",
        "full_export_manifest_sha256": sha(full_manifest),
        "source_root": "source-build",
        "included_prefixes": [".cargo/", "crates/", "test-data/", "tools/perf-baseline/"],
        "included_root_files": ["Cargo.toml"],
        "file_count": len(selected_records),
        "files": selected_records,
    })
    put(root, "source-build-manifest.sha256", (sha(selected_manifest) + "  source-build-manifest.json\n").encode())
    write_json(root / "source-closure-selection.json", {
        "schema": "litchi-xlsx-profile-source-closure-selection-v1",
        "origin_commit": "commit-synthetic",
        "full_export_root": "source",
        "selected_root": "source-build",
    })

    # Evidence categories that must remain in the archive.
    put(root, "build/native.log", b"build")
    put(root, "receipts/capture.json", b"receipt")
    put(root, "raw/capture.bin", b"raw")
    put(root, "tooling/check.py", b"tooling")
    put(root, "tmp/raw.tmp", b"tmp")
    put(root, "python-tools/__pycache__/not-selected.pyc", b"pyc")

    bound = put(root, "target-native/release/litchi-perf-baseline", b"binary")
    put(root, "target-native/release/ignored-fingerprint", b"cache")
    put(root, "target-native/debug/ignored", b"cache")
    put(root, "source/__pycache__/ignored.pyc", b"pyc")
    write_json(work / "binary-descriptor.json", {
        "schema": "litchi.performance.0944.binary-descriptor-v1",
        "source_root": str(root.resolve()),
        "binaries": [{"path": "target-native/release/litchi-perf-baseline", "bytes": bound.stat().st_size, "sha256": sha(bound)}],
        "build_receipt": {"path": "synthetic-build-terminal.json"},
    })
    return bound, selected_root


def test_selector_and_closure(module) -> None:
    with tempfile.TemporaryDirectory(prefix="0944-archive-scope-") as raw:
        root = Path(raw) / "evidence"
        work = Path(raw) / "archive"
        bound, selected_root = create_bound_root(module, root, work)
        closure = module.validate_source_closure()
        assert closure["full_export"]["file_count"] == 1
        assert closure["selected_closure"]["selected_files"] == 3
        binaries = module.expected_binary_descriptors()
        assert binaries["binaries"]["target-native/release/litchi-perf-baseline"]["bytes"] == bound.stat().st_size
        header, records = module.collect_records()
        paths = {item["path"] for item in records}
        assert "source/full-export-marker.txt" not in paths
        assert "source-build/test-data/nested/target-seed.bin" in paths
        assert "target-native/release/litchi-perf-baseline" in paths
        assert "target-native/release/ignored-fingerprint" not in paths
        assert "build/native.log" in paths and "raw/capture.bin" in paths
        assert header["source_closure"]["selected_closure"]["selected_files"] == 3
        assert module.is_redundant_source(Path("source/test-data/target-fuzz-seed"))
        assert not module.may_descend(Path("source"))
        assert module.is_target_path(Path("target-native/release/probe"))
        assert module.selected_file(Path("source-build/test-data/nested/target-seed.bin"))
        assert not module.selected_file(Path("target-native/release/ignored-fingerprint"))
        assert selected_root.joinpath("test-data/nested/target-seed.bin").is_file()


def test_cleanup_scope(module) -> None:
    with tempfile.TemporaryDirectory(prefix="0944-archive-cleanup-") as raw:
        root = Path(raw) / "evidence"
        work = Path(raw) / "archive"
        bound, selected_root = create_bound_root(module, root, work)
        module.expected_binary_descriptors()
        removed_targets = set(module.cleanup_target_native())
        assert "target-native/release/ignored-fingerprint" in removed_targets
        assert "target-native/debug/ignored" in removed_targets
        module.cleanup_pycache()
        assert module.cleanup_redundant_source() == "source"
        preserved = module.assert_post_cleanup()
        assert not (root / "source").exists()
        assert bound.is_file()
        assert selected_root.joinpath("test-data/nested/target-seed.bin").is_file()
        assert selected_root.joinpath("tools/raw-tooling.log").is_file()
        assert not (root / "target-native/debug").exists()
        assert preserved["selected_source"]["selected_files"] == 3


def test_gate_is_required(module) -> None:
    old_flag = module._GO_FLAG
    old_env = __import__("os").environ.pop(module.GO_ENV, None)
    try:
        module._GO_FLAG = False
        try:
            module.archive_mode()
        except module.ArchiveError as error:
            assert "require --go" in str(error)
        else:
            raise AssertionError("archive mode bypassed its GO gate")
    finally:
        module._GO_FLAG = old_flag
        if old_env is not None:
            __import__("os").environ[module.GO_ENV] = old_env


def main() -> None:
    module = load_script()
    test_selector_and_closure(module)
    test_cleanup_scope(module)
    test_gate_is_required(module)
    print("archive-0944-v1 synthetic selector/closure/cleanup tests: PASS")


if __name__ == "__main__":
    main()
