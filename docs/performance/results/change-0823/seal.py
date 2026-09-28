"""Bind the final 0823 packet and exact staged/committed changes to its base."""
import hashlib
import json
from pathlib import Path
import subprocess
import sys

P = Path(__file__).resolve().parent
ROOT = P.parents[3]
DOCS = ["0823-pptx-scene-reader-workflow.md", "BASELINE.md", "CRUD_COVERAGE.md",
        "GOAL_AUDIT.md", "HOTSPOTS.md", "REPORT.md"]


def digest(data):
    return hashlib.sha256(data).hexdigest()


def main():
    assert len(sys.argv) == 2 and sys.argv[1] in ("--write", "--check-index", "--check-head")
    subprocess.run([sys.executable, "-B", str(P / "validate.py"), "--final"], cwd=ROOT, check=True)
    assert not list(P.rglob("__pycache__"))
    plan = json.loads((P / "plan.json").read_text())
    disposition = json.loads((P / "disposition.json").read_text())
    assert disposition["decision"] in ("adopt", "reject")
    source = plan["source_allowlist"][0]
    selected = "after" if disposition["decision"] == "adopt" else "before"
    assert (ROOT / source).read_bytes() == (P / "candidate" / selected / "reader.rs").read_bytes()
    subprocess.run(["git", "merge-base", "--is-ancestor", plan["base"], "HEAD"], cwd=ROOT, check=True)
    paths = [path for path in P.rglob("*") if path.is_file() and path.name != "seal.json"]
    assert not any(path.is_symlink() for path in paths)
    paths += [ROOT / "docs/performance" / name for name in DOCS]
    if selected == "after":
        paths.append(ROOT / source)
    files = {str(path.relative_to(ROOT)): digest(path.read_bytes()) for path in sorted(paths)}
    encoded = json.dumps({"schema": "litchi.performance.0823.seal.v1", "base": plan["base"],
                          "files": files}, indent=2, sort_keys=True) + "\n"
    if sys.argv[1] == "--write":
        assert not (P / "seal.json").exists()
        (P / "seal.json").write_text(encoded)
    else:
        assert (P / "seal.json").read_text() == encoded
        expected = files | {str((P / "seal.json").relative_to(ROOT)): digest(encoded.encode())}
        index = sys.argv[1] == "--check-index"
        command = ["git", "diff", "--cached", "--name-only", "-z", plan["base"]] if index else [
            "git", "diff", "--name-only", "-z", plan["base"], "HEAD"]
        actual = {name for name in subprocess.check_output(command, cwd=ROOT).decode().split("\0") if name}
        assert actual == set(expected), (actual - set(expected), set(expected) - actual)
        revision = ":" if index else "HEAD:"
        for name, wanted in expected.items():
            assert digest(subprocess.check_output(["git", "show", revision + name], cwd=ROOT)) == wanted, name
    print("0823 seal PASS:", len(files) + 1, "owned paths including seal")


if __name__ == "__main__":
    main()
