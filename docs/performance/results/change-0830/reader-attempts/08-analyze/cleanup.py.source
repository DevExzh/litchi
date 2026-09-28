"""Remove only this packet's isolated build target after completed validation."""
import shutil

import driver as d


def main():
    d.check_inputs()
    assert d.read(d.P / "capture.json")["status"] == "pass"
    assert (d.P / "analysis.json.gz").is_file()
    assert d.TARGET == d.ROOT.parent / "litchi-target-0830"
    assert d.TARGET.is_dir() and not d.TARGET.is_symlink()
    assert not (d.P / "cleanup.json").exists()
    binaries = {arm: d.read(d.P / ("binary-" + arm + ".json")) for arm in ("ordinary", "fp")}
    for row in binaries.values():
        assert d.sha(row["path"]) == row["sha256"]
    files = [p for p in d.TARGET.rglob("*") if p.is_file()]
    removed_bytes = sum(p.stat().st_size for p in files)
    shutil.rmtree(d.TARGET)
    assert not d.TARGET.exists()
    d.check_inputs()
    d.write(d.P / "cleanup.json", {"status": "pass", "target": str(d.TARGET),
            "removed_files": len(files), "removed_bytes": removed_bytes, "binaries": binaries})
    print(f"cleanup PASS: {len(files)} files, {removed_bytes} bytes")


if __name__ == "__main__":
    main()
