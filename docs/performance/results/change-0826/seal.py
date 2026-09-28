"""Bind the repair and evidence to the exact documentation/tool commit."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys

import validate

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
OWNED = ["tools/perf_allocation_schema.py", "tools/test_perf_allocation_schema.py"] + [
    "docs/performance/" + name for name in (
        "0826-allocation-schema-preflight.md", "BASELINE.md", "CRUD_COVERAGE.md",
        "GOAL_AUDIT.md", "HOTSPOTS.md", "REPORT.md")]


def digest(data):
    return hashlib.sha256(data).hexdigest()


def main():
    assert sys.argv[1:] in (["--write"], ["--check-index"], ["--check-head"])
    validate.check()
    base = json.loads((P / "origin.json").read_text())["base"]
    subprocess.run(["git", "merge-base", "--is-ancestor", base, "HEAD"], cwd=ROOT, check=True)
    files = [p for p in P.rglob("*") if p.is_file() and p.name != "seal.json"]
    files += [ROOT / name for name in OWNED]
    assert not any(p.is_symlink() for p in files)
    hashes = {str(p.relative_to(ROOT)): digest(p.read_bytes()) for p in sorted(files)}
    encoded = json.dumps({"schema": "litchi.performance.0826.seal.v1", "base": base, "files": hashes},
                         indent=2, sort_keys=True) + "\n"
    seal = P / "seal.json"
    if sys.argv[1] == "--write":
        with seal.open("x") as stream:
            stream.write(encoded)
    else:
        assert seal.read_text() == encoded
        expected = hashes | {str(seal.relative_to(ROOT)): digest(encoded.encode())}
        index = sys.argv[1] == "--check-index"
        command = ["git", "diff", "--cached", "--name-only", "-z", base] if index else [
            "git", "diff", "--name-only", "-z", base, "HEAD"]
        actual = {p for p in subprocess.check_output(command, cwd=ROOT).decode().split("\0") if p}
        assert actual == set(expected), (actual - set(expected), set(expected) - actual)
        prefix = ":" if index else "HEAD:"
        for name, wanted in expected.items():
            assert digest(subprocess.check_output(["git", "show", prefix + name], cwd=ROOT)) == wanted, name
    print("0826 seal PASS:", len(hashes) + 1, "owned paths including seal")


if __name__ == "__main__":
    main()
