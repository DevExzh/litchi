"""Verify and remove only the completed 0825 build and filesystem roots."""
import hashlib
import json
from pathlib import Path
import shutil
import time

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
TARGET = ROOT.parent / "litchi-target-0825"
SCRATCH = ROOT.parent / "litchi-fs-0825"


def read(path):
    return json.loads(path.read_text())


def main():
    assert not (P / "cleanup.json").exists()
    assert read(P / "quality.json")["status"] == "pass"
    for lane in ("qualification-before", "qualification-after", "native", "observer"):
        assert read(P / lane / "complete.json")["status"] == "pass"
    assert (P / "analysis.json").is_file()
    assert (P / "validation-precleanup.log").is_file()
    plan = read(P / "plan.json")
    assert str(TARGET) == plan["target"] and str(SCRATCH) == plan["scratch"]
    marker = SCRATCH / ".litchi-performance-0825-owned"
    assert marker.read_text() == plan["scratch_marker"]
    binaries = []
    for leg in ("before", "after"):
        build = read(P / f"build-{leg}/build.json")
        for descriptor in build["binaries"].values():
            artifact = descriptor["artifact"]
            path = Path(artifact["path"])
            assert path.parent == TARGET and path.is_file() and not path.is_symlink()
            data = path.read_bytes()
            assert len(data) == artifact["bytes"]
            assert hashlib.sha256(data).hexdigest() == artifact["sha256"]
            binaries.append(artifact)
    assert len(binaries) == 6
    for executable in Path("/proc").glob("[0-9]*/exe"):
        try:
            resolved = executable.resolve(strict=True)
        except (FileNotFoundError, PermissionError, ProcessLookupError):
            continue
        assert not resolved.is_relative_to(TARGET), executable
    roots = []
    for directory in (TARGET, SCRATCH):
        assert directory.is_dir() and not directory.is_symlink()
        paths = list(directory.rglob("*"))
        assert not any(path.is_symlink() for path in paths)
        files = [path for path in paths if path.is_file()]
        roots.append({"path": str(directory), "files": len(files),
                      "logical_bytes": sum(path.stat().st_size for path in files)})
    result = {"schema": "litchi.performance.0825.cleanup.v1", "roots": roots,
              "binaries": binaries, "scratch_marker_sha256": hashlib.sha256(marker.read_bytes()).hexdigest(),
              "started": time.time()}
    for directory in (TARGET, SCRATCH):
        shutil.rmtree(directory)
    result.update(ended=time.time(), target_absent_after_removal=not TARGET.exists(),
                  scratch_absent_after_removal=not SCRATCH.exists(), verified=True)
    (P / "cleanup.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print("0825 cleanup PASS:", sum(x["files"] for x in roots), "files;",
          sum(x["logical_bytes"] for x in roots), "logical bytes; both owned roots removed")


if __name__ == "__main__":
    main()
