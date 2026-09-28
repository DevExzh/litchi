"""Root-only cleanup after terminal captures and independent offline review."""
import shutil
import time
import custody as c

assert not (c.P / "cleanup.json").exists()
assert c.TARGET == c.ROOT.parent / "litchi-target-0816"
assert c.TARGET.is_dir() and not c.TARGET.is_symlink()
for name in ("analysis.json", "raw-audit.json", "results-review.md"):
    assert (c.P / name).is_file(), name
build = c.read(c.P / "build.json")
frozen = c.read(build["source"]["path"])
c.unchanged(frozen)
removed = list(build["binaries"].values())
assert len(removed) == 2
for artifact in removed:
    assert c.artifact(artifact["path"]) == artifact
paths = [p for p in c.TARGET.rglob("*") if p.is_file()]
size = sum(p.stat().st_size for p in paths)
started = time.time()
shutil.rmtree(c.TARGET)
assert not c.TARGET.exists()
c.unchanged(frozen)
c.write(c.P / "cleanup.json", {
    "schema": "litchi.performance.0816.cleanup.v1",
    "target": str(c.TARGET), "target_removed": True,
    "removed_files": len(paths), "removed_logical_bytes": size,
    "removed_binaries": removed, "binaries_verified_before_removal": True,
    "source": build["source"], "started": started, "ended": time.time(),
})
print("0816 owned target removed; two binaries and unchanged sources verified")
