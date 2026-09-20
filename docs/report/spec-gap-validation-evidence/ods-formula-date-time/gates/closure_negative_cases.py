#!/usr/bin/env python3
"""Exercise fail-closed date/time gate custody checks without Cargo.

These checks use temporary paths and pure manifest helpers.  They never write a
freeze or a gate receipt and are safe to run while the date/time batch is still
in preparation.
"""

from __future__ import annotations

import importlib.util
from pathlib import Path
import tempfile


HERE = Path(__file__).resolve().parent


def load(name: str):
    path = HERE / name
    spec = importlib.util.spec_from_file_location(f"date_time_gate_{name[:-3]}", path)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"unable to load gate module: {path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def rejects(callable_, *args, **kwargs) -> None:
    try:
        callable_(*args, **kwargs)
    except (RuntimeError, ValueError):
        return
    raise AssertionError(f"expected {callable_.__name__} to reject its input")


def main() -> int:
    stage = load("stage.py")
    runner = load("run.py")
    verifier = load("verify.py")

    rejects(stage.safe_relative, "../outside", "negative path")
    rejects(stage.safe_relative, "/absolute/path", "negative path")
    rejects(runner.safe_relative, "../outside", "negative path")
    rejects(verifier.safe_relative, "../outside", "negative path")

    with tempfile.TemporaryDirectory() as directory:
        checkout = Path(directory) / "checkout"
        checkout.mkdir()
        rejects(runner.safe_repo_path, checkout, "../outside")
        rejects(runner.selected_source_paths, checkout, {"crates/litchi-ods/src/missing.rs": "0" * 64})

    config = stage.baseline()
    selected, categories, warnings = stage.source_closure(config)
    baseline_paths = set(config["selected_baseline_sources"])
    if not baseline_paths <= selected:
        raise AssertionError("source closure dropped a baseline-selected path")
    evaluation_paths = set(stage.evaluation_source_paths())
    if not evaluation_paths <= selected:
        raise AssertionError("source closure dropped a recursive evaluator path")
    shared_lookup = "crates/litchi-ods/src/codec/formula/evaluation/value/lookup.rs"
    if shared_lookup not in selected:
        raise AssertionError("source closure dropped the shared value lookup dispatcher")
    if "coverage-requirements.json" in "\n".join(selected):
        raise AssertionError("mutable coverage receipt bindings entered the frozen source map")
    if "docs/FEATURE_MATRIX.md" not in selected or "crates/litchi-ods/docs/FEATURE_MATRIX.md" not in selected:
        raise AssertionError("feature matrix inputs are outside the frozen source map")
    for category, paths in categories.items():
        if len(paths) != len(set(paths)):
            raise AssertionError(f"source closure category contains duplicates: {category}")
        if not set(paths) <= selected:
            raise AssertionError(f"source closure category escapes selected map: {category}")
    if not isinstance(warnings, list):
        raise AssertionError("source closure warnings are not a list")
    rejects(verifier.validate_map, {"../escape": "0" * 64}, "negative selected map")
    rejects(verifier.validate_map, {"Cargo.lock": "not-a-sha"}, "negative selected map")

    with tempfile.TemporaryDirectory() as directory:
        mirror = Path(directory) / "repo"
        nested = mirror / stage.EVALUATION_ROOT / "new_nested" / "helper.rs"
        nested.parent.mkdir(parents=True)
        nested.write_text("// synthetic nested evaluator helper\n", encoding="utf-8")
        nested_paths = set(stage.evaluation_source_paths(mirror))
        if str(nested.relative_to(mirror)) not in nested_paths:
            raise AssertionError("recursive evaluator closure dropped a newly nested helper")

        evidence = mirror / stage.EVIDENCE.relative_to(stage.ROOT)
        evidence.mkdir(parents=True)
        for relative in stage.IMMUTABLE_EVIDENCE_FILES:
            path = evidence / relative
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text("{}\n", encoding="utf-8")
        (evidence / "completion.md").write_text("future completion receipt\n", encoding="utf-8")
        (evidence / "review-receipt.json").write_text("{}\n", encoding="utf-8")
        (evidence / "oracle-plan.md").write_text("fixed oracle plan\n", encoding="utf-8")
        (evidence / "oracle_verify.py").write_text("# independent oracle\n", encoding="utf-8")
        for receipt in ("oracle-receipt.json", "oracle-execution.json", "oracle-results.json"):
            (evidence / receipt).write_text("{}\n", encoding="utf-8")
        (evidence / "native").mkdir()
        (evidence / "native" / "fixture.ods").write_bytes(b"fixture")
        (evidence / "performance").mkdir()
        (evidence / "performance" / "performance-plan.md").write_text("fixed performance plan\n", encoding="utf-8")
        (evidence / "performance" / "harness" / "src").mkdir(parents=True)
        (evidence / "performance" / "harness" / "Cargo.toml").write_text("[package]\n", encoding="utf-8")
        (evidence / "performance" / "harness" / "Cargo.lock").write_text("# lock\n", encoding="utf-8")
        (evidence / "performance" / "harness" / "src" / "main.rs").write_text("fn main() {}\n", encoding="utf-8")
        (evidence / "performance" / "run_profile.py").write_text("# fixed case matrix\n", encoding="utf-8")
        fixture_paths = set(stage.evidence_input_paths(mirror))
        if str((evidence / "completion.md").relative_to(mirror)) in fixture_paths:
            raise AssertionError("future completion report escaped the evidence allowlist")
        if str((evidence / "review-receipt.json").relative_to(mirror)) in fixture_paths:
            raise AssertionError("future review receipt escaped the evidence allowlist")
        if str((evidence / "oracle_verify.py").relative_to(mirror)) not in fixture_paths:
            raise AssertionError("oracle source escaped the evidence allowlist")
        for receipt in ("oracle-receipt.json", "oracle-execution.json", "oracle-results.json"):
            if str((evidence / receipt).relative_to(mirror)) in fixture_paths:
                raise AssertionError(f"oracle receipt escaped the evidence allowlist: {receipt}")
        if str((evidence / "native" / "fixture.ods").relative_to(mirror)) not in fixture_paths:
            raise AssertionError("native fixture escaped the evidence allowlist")
        harness_output = evidence / "performance" / "harness" / "run.log"
        harness_output.write_text("future harness output\n", encoding="utf-8")
        if str(harness_output.relative_to(mirror)) in set(stage.evidence_input_paths(mirror)):
            raise AssertionError("non-source harness output escaped the evidence allowlist")

    print(
        "date/time closure negative cases passed "
        f"(selected={len(selected)}, categories={len(categories)}, warnings={len(warnings)})"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
