#!/usr/bin/env python3
"""Verify the sealed, diagnostic-only 0551 evidence without writing files."""

import hashlib
import importlib.util
import json
from pathlib import Path
import re
import subprocess
import sys

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]


def require(condition, message):
    if not condition:
        raise ValueError(message)


def digest(path):
    require(path.is_file() and not path.is_symlink(), f"not a regular file: {path}")
    return hashlib.sha256(path.read_bytes()).hexdigest()


def read(name):
    return json.loads((HERE / name).read_text())


def check_seal(base):
    manifest = {}
    for line in (base / "SHA256SUMS").read_text().splitlines():
        expected, name = line.split("  ", 1)
        require(name not in manifest, f"duplicate seal entry: {name}")
        require(not Path(name).is_absolute() and ".." not in Path(name).parts,
                "unsafe seal path")
        manifest[name] = expected
    actual = {str(path.relative_to(base)) for path in base.rglob("*")
              if path.is_file() and path.name != "SHA256SUMS"}
    require(actual == set(manifest), f"seal inventory differs: {base}")
    for name, expected in manifest.items():
        require(digest(base / name) == expected, f"seal mismatch: {base / name}")
    return len(manifest)


def main():
    before = {str(path): digest(path) for path in HERE.rglob("*") if path.is_file()}
    require(not list(HERE.rglob("__pycache__")), "owned bytecode cache exists")
    prior_count = check_seal(HERE.parent / "change-0550")
    own_count = check_seal(HERE)
    for name, expected in read("documentation-manifest.json").items():
        require(digest(ROOT / name) == expected, f"documentation differs: {name}")
    spec = importlib.util.spec_from_file_location("analysis_0551", HERE / "analyze.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    replay = json.dumps(module.analyze(), indent=2, sort_keys=True) + "\n"
    require(replay == (HERE / "analysis.json").read_text(), "analysis replay differs")
    for path in HERE.glob("*.py"):
        compile(path.read_text(), str(path), "exec")
    model = (ROOT / "crates/litchi-xlsx/src/raw/worksheet/edit/codec/snapshot/model.rs").read_text()
    layout = model.split("pub(crate) struct Layout {", 1)[1].split("\n}", 1)[0]
    fields = re.findall(r"pub\(crate\) (\w+):", layout)
    documented = re.findall(r"^\| `([^`]+)` \|", (HERE / "design-review.md").read_text(), re.M)
    require(len(fields) == len(documented) == 16 and set(fields) == set(documented),
            "Layout field checklist differs; this checks coverage, not semantic equivalence")
    metadata = read("metadata-checks.json")
    expected_commands = {
        "boundaries": ["python3", "-B", "tools/check_crate_boundaries.py"],
        "claims": ["python3", "-B", "tools/check_perf_claims.py", "--registry",
                   "docs/performance/claim-registry-v1.json", "--repo-root", ".",
                   "--evidence-root", ".", "--mode", "strict"],
    }
    checks = metadata["checks"]
    require(len(checks) == 2 and {row["name"] for row in checks} == set(expected_commands),
            "metadata check matrix differs")
    for row in checks:
        require(row["command"] == expected_commands[row["name"]], "metadata command differs")
        require(row["returncode"] == 0, "metadata check failed")
        for stream in ("stdout", "stderr"):
            require(digest(HERE / f"{row['name']}.{stream}") == row[f"{stream}_sha256"],
                    "metadata output differs")
    revision = read("inputs.json")["source_revision"]
    paths = subprocess.check_output(["git", "diff", "--name-only", revision, "--"],
                                    cwd=ROOT, text=True).splitlines()
    allowed = set(read("documentation-manifest.json"))
    prefix = str(HERE.relative_to(ROOT)) + "/"
    require(all(path in allowed or path.startswith(prefix) for path in paths),
            "changes outside the diagnostic batch")
    after = {str(path): digest(path) for path in HERE.rglob("*") if path.is_file()}
    require(before == after, "verification changed the bundle")
    print(f"PASS: {prior_count} prior sealed files; {own_count} current sealed files; "
          "source/ADR hashes; exact scanner partitions; corpus coverage; "
          "documentation; metadata receipts; read-only replay; no owned temporary cache")


if __name__ == "__main__":
    main()
