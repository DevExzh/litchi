"""Root-owned cleanup for the supplemental amendment target.

Cleanup is allowed only after both native build descriptors and the complete
native lane exist. It records exact binary identities before removing the
packet-owned target so the offline reader can replay evidence after cleanup.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import custody as c


P = c.P


def main() -> None:
    out = P / "cleanup.json"
    if out.exists():
        raise AssertionError(f"cleanup witness already exists: {out}")
    native_complete = P / "native" / "complete.json"
    if not native_complete.is_file():
        raise AssertionError("native lane must be complete before cleanup")
    complete = c.read(native_complete)
    if complete.get("children") != 936 or complete.get("samples") != 28080:
        raise AssertionError("native lane cardinality is not complete")
    removed = []
    for leg in ("before", "after"):
        build = c.read(P / f"build-{leg}" / "build.json")
        descriptor = build["binary"]
        path = Path(descriptor["path"])
        if not path.is_file() or path.is_symlink():
            raise AssertionError(f"missing owned binary: {path}")
        if path.stat().st_size != descriptor["bytes"] or c.sha(path) != descriptor["sha256"]:
            raise AssertionError(f"binary changed before cleanup: {path}")
        if path.parent != c.TARGET:
            raise AssertionError(f"binary escapes owned target: {path}")
        removed.append(descriptor)
    if not c.TARGET.is_dir() or c.TARGET.is_symlink():
        raise AssertionError(f"owned target missing or symlink: {c.TARGET}")
    removed_target_bytes = sum(
        path.stat().st_size for path in c.TARGET.rglob("*") if path.is_file() and not path.is_symlink()
    )
    shutil.rmtree(c.TARGET)
    if c.TARGET.exists():
        raise AssertionError("owned target was not removed")
    c.write(out, {
        "schema": "litchi.performance.0806.amendment-cleanup.v1",
        "target": str(c.TARGET),
        "target_removed": True,
        "removed_target_bytes": removed_target_bytes,
        "removed_binaries": removed,
        "removed_failed_binaries": [],
        "native_complete": c.relative_artifact(native_complete),
    })


if __name__ == "__main__":
    main()
