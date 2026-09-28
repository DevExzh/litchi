"""Seal exactly the owned packet/report/index paths; verify their Git blobs."""
import hashlib
import subprocess
import sys

import driver as d


REPORTS = (
    "docs/performance/0830-xlsx-edit-allocation-profile.md",
    "docs/performance/BASELINE.md",
    "docs/performance/CRUD_COVERAGE.md",
    "docs/performance/GOAL_AUDIT.md",
    "docs/performance/HOTSPOTS.md",
    "docs/performance/REPORT.md",
)


def main():
    mode = sys.argv[1]
    assert mode in ("create", "verify")
    path = d.P / "seal.json"
    assert d.inventory() == d.read(d.P / "inputs.json")
    assert all(d.sha(d.ROOT / n) == h for n, h in d.UNRELATED.items())
    assert not d.TARGET.exists()
    files = {str(p.relative_to(d.ROOT)): d.sha(p) for p in d.P.rglob("*")
             if p.is_file() and p != path}
    files.update({name: d.sha(d.ROOT / name) for name in REPORTS})
    if mode == "create":
        assert d.output(["git", "rev-parse", "HEAD"]) == d.BASE
        d.write(path, {"schema": "litchi.performance.0830.seal.v1", "base": d.BASE,
                       "files": files, "unrelated": d.UNRELATED})
    else:
        sealed = d.read(path)
        assert sealed["files"] == files
        assert sealed["base"] == d.BASE
        assert d.output(["git", "rev-parse", "HEAD^"]) == d.BASE
        all_files = {**files, str(path.relative_to(d.ROOT)): d.sha(path)}
        changed = set(d.output(["git", "diff-tree", "--no-commit-id", "--name-only", "-r", "HEAD"]).splitlines())
        assert changed == set(all_files)
        for name, digest in all_files.items():
            raw = subprocess.check_output(["git", "show", "HEAD:" + name], cwd=d.ROOT)
            assert hashlib.sha256(raw).hexdigest() == digest, name
    print(f"0830 seal {mode} PASS: {len(files)} owned paths; unrelated work preserved")


if __name__ == "__main__":
    main()
