"""Remove only the verified 0823 target after offline and codegen checks."""
import hashlib
import json
from pathlib import Path
import shutil
import time

P = Path(__file__).resolve().parent
TARGET = Path("/home/zhuhe/code/litchi-target-0823")


def main():
    assert not (P / "cleanup.json").exists()
    assert TARGET.is_dir() and not TARGET.is_symlink()
    for name in ("quality-before.json", "quality-after.json", "quality-probes.json"):
        assert json.loads((P / name).read_text())["status"] == "pass"
    for name in ("native", "allocation", "qualification-before", "qualification-after"):
        assert json.loads((P / name / "complete.json").read_text())["status"] == "pass"
    assert (P / "analysis.json").is_file()
    assert (P / "codegen/result.json").is_file()
    assert (P / "validation-precleanup.log").is_file()
    binaries = []
    for leg in ("before", "after"):
        build = json.loads((P / f"build-{leg}/build.json").read_text())
        for descriptor in build["binaries"].values():
            artifact = descriptor["artifact"]
            path = Path(artifact["path"])
            assert path.parent == TARGET and not path.is_symlink()
            data = path.read_bytes()
            assert len(data) == artifact["bytes"]
            assert hashlib.sha256(data).hexdigest() == artifact["sha256"]
            binaries.append(artifact)
    assert len(binaries) == 8
    # No retained executable may still be mapped by a live process.
    for executable in Path("/proc").glob("[0-9]*/exe"):
        try:
            resolved = executable.resolve(strict=True)
        except (FileNotFoundError, PermissionError, ProcessLookupError):
            continue
        assert not resolved.is_relative_to(TARGET), str(executable)
    paths = list(TARGET.rglob("*"))
    assert not any(x.is_symlink() for x in paths)
    files = [x for x in paths if x.is_file()]
    result = {"schema": "litchi.performance.0823.cleanup.v1", "target": str(TARGET),
              "files": len(files), "logical_bytes": sum(x.stat().st_size for x in files),
              "binaries": binaries, "started": time.time()}
    shutil.rmtree(TARGET)
    result.update({"ended": time.time(), "target_absent_after_removal": not TARGET.exists(), "verified": True})
    (P / "cleanup.json").write_text(json.dumps(result, indent=2, sort_keys=True) + "\n")
    print("0823 cleanup PASS:", result["files"], "files,", result["logical_bytes"], "logical bytes")


if __name__ == "__main__":
    main()
