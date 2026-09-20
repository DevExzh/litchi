#!/usr/bin/env python3
"""Regression checks for selected-file workspace closure custody.

The script does not run Cargo or write gate receipts.  It selects a current
ODS Rust module absent from the retained baseline, proves the expected closure
includes it, then proves an observed manifest omitting it is rejected.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
import subprocess
import tempfile


HERE = Path(__file__).resolve().parent
VERIFY = HERE / "verify.py"
ROOT = HERE.parents[4]


def load_verifier():
    spec = importlib.util.spec_from_file_location("lookup_gate_verify_closure", VERIFY)
    if spec is None or spec.loader is None:
        raise RuntimeError("unable to load gate verifier")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def load_runner():
    runner_path = HERE / "run.py"
    spec = importlib.util.spec_from_file_location("lookup_gate_run_closure", runner_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("unable to load gate runner")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def main() -> int:
    verifier = load_verifier()
    runner = load_runner()
    baseline_paths = set(
        subprocess.check_output(
            ["git", "ls-tree", "-r", "--name-only", verifier.BASELINE_COMMIT, "crates/litchi-ods"],
            cwd=ROOT,
            text=True,
        ).splitlines()
    )
    synthetic = "crates/litchi-ods/src/codec/formula/evaluation/lookup/closure_probe.rs"
    if synthetic in baseline_paths:
        raise RuntimeError("synthetic closure probe unexpectedly exists in retained baseline")
    selected = {synthetic: "candidate-hash-placeholder"}
    expected = verifier.expected_workspace_paths(verifier.BASELINE_COMMIT, selected)
    verifier.equal("selected module is in expected closure", synthetic in expected, True)

    omitted = set(expected)
    omitted.remove(synthetic)
    try:
        verifier.verify_workspace_path_set(omitted, expected)
    except RuntimeError as error:
        if "workspace source path set" not in str(error):
            raise AssertionError(f"closure regression failed for an unexpected reason: {error}") from error
        print(f"closure regression passed: rejected omitted {synthetic}")
    else:
        raise AssertionError("closure verifier accepted an observed manifest omitting a selected module")

    with tempfile.TemporaryDirectory() as directory:
        checkout = Path(directory)
        selected = {"crates/litchi-ods/src/codec/formula/evaluation/lookup.rs": "hash"}
        try:
            runner.selected_source_paths(checkout, selected)
        except RuntimeError as error:
            if "absent" not in str(error):
                raise AssertionError(f"selected-source regression failed for an unexpected reason: {error}") from error
            pass
        else:
            raise AssertionError("gate runner accepted an absent selected source")
        source = checkout / next(iter(selected))
        source.parent.mkdir(parents=True)
        source.write_text("// selected module\n", encoding="utf-8")
        observed = runner.selected_source_paths(checkout, selected)
        if observed[next(iter(selected))] != source:
            raise AssertionError("gate runner did not retain the selected source path")
    print("selected-source custody regression passed")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
