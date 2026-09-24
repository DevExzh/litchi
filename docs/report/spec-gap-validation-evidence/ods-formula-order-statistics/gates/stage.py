#!/usr/bin/env python3
"""Stage the exact batch and all selected profile inputs in an isolated checkout."""

import argparse
import ast
import hashlib
import json
from pathlib import Path
import shutil
import subprocess


HERE = Path(__file__).resolve().parent
EVIDENCE = HERE.parent
ROOT = HERE.parents[4]
PREFIX = str(EVIDENCE.relative_to(ROOT))
RUST = [
    "crates/litchi-ods/src/codec/formula/evaluation.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/order.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/rounding.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/order.rs",
    "crates/litchi-ods/src/codec/formula/evaluation/value/matrix.rs",
    *[f"crates/litchi-ods/tests/ods_formula_order_{kind}.rs"
      for kind in ("evaluation", "limits", "native", "oracle")],
]
SELECTED = RUST + [
    f"{PREFIX}/numeric-goldens.json",
    f"{PREFIX}/native/cached-results.json",
    f"{PREFIX}/native/provenance.json",
    "Cargo.lock",
]


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest() if path.is_file() else None


def profile_sources():
    module = ast.parse((EVIDENCE / "performance/run_profile.py").read_text())
    for node in module.body:
        if isinstance(node, ast.Assign) and any(
            isinstance(target, ast.Name) and target.id == "SOURCE_FILES"
            for target in node.targets
        ):
            return ast.literal_eval(node.value)
    raise AssertionError("profile SOURCE_FILES missing")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("checkout", type=Path)
    args = parser.parse_args()
    checkout = args.checkout.resolve()
    assert checkout != ROOT, "must stage into an isolated checkout"
    baseline = json.loads((EVIDENCE / "baseline.json").read_text())["commit"]
    assert subprocess.check_output(
        ["git", "rev-parse", "HEAD"], cwd=checkout, text=True
    ).strip() == baseline
    profile = set(profile_sources())
    assert set(SELECTED) <= profile, "profile omits selected batch inputs"
    assert f"{PREFIX}/numeric_oracle.py" in profile
    sources = {
        relative: HERE / "Cargo.lock" if relative == "Cargo.lock" else ROOT / relative
        for relative in profile
    }
    for relative in SELECTED + [f"{PREFIX}/numeric_oracle.py", f"{PREFIX}/contract.md"]:
        assert sources[relative].is_file(), f"required source missing: {relative}"
    for relative, source in sorted(sources.items()):
        target = checkout / relative
        if source.is_file():
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(source, target)
        else:
            assert not target.exists(), f"unexpected isolated file: {relative}"
    hashes = {relative: digest(source) for relative, source in sorted(sources.items())}
    for relative, expected in hashes.items():
        assert digest(checkout / relative) == expected, relative
    (HERE / "batch-files.json").write_text(json.dumps(RUST, indent=2) + "\n")
    (HERE / "freeze.json").write_text(json.dumps({
        "base_commit": baseline,
        "selected_files": {relative: hashes[relative] for relative in SELECTED},
    }, indent=2) + "\n")
    (HERE / "staged-profile-sources.json").write_text(json.dumps(hashes, indent=2) + "\n")
    print(json.dumps({"staged_sources": len(hashes), "selected_files": len(SELECTED)}))


if __name__ == "__main__":
    main()
