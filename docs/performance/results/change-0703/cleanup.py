#!/usr/bin/env python3
"""Remove exactly the two owned 0703 diagnostic build directories."""
import hashlib
import json
import shutil
from pathlib import Path
P = Path(__file__).resolve().parent
ROOT = P.parents[3]

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

def main():
    build = json.loads((P / "build.json").read_text())
    target = ROOT.parent / "litchi-target-0703"
    binaries = ROOT.parent / "litchi-0703-bin"
    assert Path(build["binary"]) == binaries / "probe0703"
    assert digest(Path(build["binary"])) == build["binary_sha256"]
    lock = ROOT / "Cargo.lock"
    before = digest(lock)
    assert before == build["workspace_lock_sha256"]
    removed = []
    for path in (target, binaries):
        assert path.is_dir() and not path.is_symlink(), path
        shutil.rmtree(path)
        assert not path.exists()
        removed.append(str(path))
    assert digest(lock) == before
    assert not list(P.rglob("*.phase"))
    assert not list(P.rglob("__pycache__"))
    (P / "cleanup.json").write_text(json.dumps(dict(removed_paths=removed,
        workspace_lock_preserved=True,workspace_lock_sha256=before,
        phase_files_remaining=0,python_cache_directories_remaining=0),indent=2)+"\n")
    print("PASS: removed exactly two owned directories; workspace lock preserved")

if __name__ == "__main__":
    main()
