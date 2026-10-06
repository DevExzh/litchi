#!/usr/bin/env python3
"""Synthetic selector/cleanup-scope checks for archive-0942-v2.py.

These tests use a temporary source tree only.  They never invoke --archive,
--verify, or --cleanup against the retained 0942 evidence root.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
import tempfile


SCRIPT = Path(__file__).with_name("archive-0942-v2.py")


def load_script():
    spec = importlib.util.spec_from_file_location("archive_0942_v2_under_test", SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def put(root: Path, relative: str, data: bytes = b"x") -> Path:
    path = root / relative
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(data)
    return path


def test_selectors(module) -> None:
    assert module.is_target_path(Path("target-profile-v1/release/probe"))
    assert not module.is_target_path(Path("source-v1/test-data/iwork/target-fuzz-seed"))
    assert module.may_descend(Path("target-profile-v1"))
    assert module.may_descend(Path("target-profile-v1/release"))
    assert not module.may_descend(Path("target-profile-v1/debug"))
    assert module.may_descend(Path("source-v1/test-data/target-cache"))
    assert module.selected_file(Path("source-v1/test-data/iwork/target-fuzz-seed"))
    assert not module.selected_file(Path("target-profile-v1/release/unbound.bin"))
    assert module.selected_file(Path("target-experimental/cache.bin"))
    module.validate_relative_record("source-v1/test-data/iwork/target-fuzz-seed")
    try:
        module.validate_relative_record("target-profile-v1/release/unbound.bin")
    except module.ArchiveError:
        pass
    else:
        raise AssertionError("unbound known-cache file was accepted")


def test_synthetic_cleanup_scope(module) -> None:
    original_root = module.SOURCE_ROOT
    original_allowlist = module.ALLOWLISTED_TARGET_FILES
    try:
        with tempfile.TemporaryDirectory(prefix="0942-archive-v2-scope-") as raw:
            root = Path(raw) / "source"
            root.mkdir()
            module.SOURCE_ROOT = root.resolve()
            module.ALLOWLISTED_TARGET_FILES = frozenset({
                Path("target-profile-v1/release/probe"),
                Path("target-profile-v2/release/probe"),
            })
            bound_v1 = put(root, "target-profile-v1/release/probe", b"bound-v1")
            bound_v2 = put(root, "target-profile-v2/release/probe", b"bound-v2")
            unbound_v1 = put(root, "target-profile-v1/debug/unbound.bin")
            unbound_v2 = put(root, "target-profile-v2/release/unbound.bin")
            nested_seed = put(root, "source-v1/test-data/iwork/target-fuzz-seed")
            nested_cache = put(root, "source-v2/tools/target-cache/blob")
            other_top_level = put(root, "target-experimental/cache.bin")
            removed = set(module.cleanup_target_caches())
            assert removed == {
                "target-profile-v1/debug/unbound.bin",
                "target-profile-v1/debug",
                "target-profile-v2/release/unbound.bin",
            }
            assert bound_v1.is_file() and bound_v2.is_file()
            assert not unbound_v1.exists() and not unbound_v2.exists()
            assert nested_seed.is_file() and nested_cache.is_file()
            assert other_top_level.is_file()
    finally:
        module.SOURCE_ROOT = original_root
        module.ALLOWLISTED_TARGET_FILES = original_allowlist


def main() -> None:
    module = load_script()
    test_selectors(module)
    test_synthetic_cleanup_scope(module)
    print("archive-0942-v2 synthetic selector/cleanup-scope tests: PASS")


if __name__ == "__main__":
    main()
