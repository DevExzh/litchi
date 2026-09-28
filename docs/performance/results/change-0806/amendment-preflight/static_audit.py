"""No-execution packet dry run for paths, schemas and frozen cardinalities."""

from __future__ import annotations

import ast
import json
from pathlib import Path

import custody as c


def main() -> None:
    plan = c.read(c.P / "plan.json")
    cases = c.read(c.P / "cases.json")
    handoff = c.read(c.P / "handoff.json")
    assert handoff["status"] == "inputs-frozen-execution-pending"
    assert handoff["native"]["build_legs"] == ["before", "after"]
    assert handoff["exclusions"]["reader_in_prebuild_frozen_inputs"] is True
    assert plan["case_count"] == len(cases) == 39
    assert plan["native"]["blocks"] == 6
    expected_children = 6 * 39 * 2 * 2
    expected_samples = expected_children * 30
    for leg in ("before", "after"):
        assert len(c.archive_manifest(leg)) == 6
        for name in c.SOURCE_FILES:
            assert (c.P / "source" / leg / name).is_file()
    assert len(c.probe_manifest()) == 5
    planned_builds = [str(c.P / f"build-{leg}") for leg in ("before", "after")]
    planned_binaries = [str(c.TARGET / f"{leg}-native") for leg in ("before", "after")]
    planned_quality_sources = [
        str(c.P / "test-src" / leg / "crates" / owner / "src")
        for leg in ("before", "after")
        for owner in plan["quality"]["helper_crates"]
    ]
    assert len(planned_builds) == 2 and len(planned_binaries) == 2
    assert len(planned_quality_sources) == 10
    for driver in ("custody.py", "quality.py", "build.py", "capture.py", "analyze.py", "cleanup.py"):
        ast.parse((c.P / driver).read_text(encoding="utf-8"), filename=driver)
    manifest = (c.P / "probe-src" / "Cargo.toml.template").read_text(encoding="utf-8")
    assert '[workspace]' in manifest and 'quick-xml = "=0.41.0"' in manifest
    assert plan["analysis"]["bootstrap_seed"] == 806082
    assert plan["scope"]["profiles"] is False
    assert plan["scope"]["callgrind"] is False
    print(json.dumps({
        "schema": "litchi.performance.0806.amendment-static-audit.v1",
        "native_children": expected_children,
        "native_samples": expected_samples,
        "planned_build_legs": planned_builds,
        "planned_binary_legs": planned_binaries,
        "planned_quality_source_dirs": planned_quality_sources,
        "source_files_per_leg": 6,
        "probe_files": 5,
        "cargo_or_native_executed": False,
        "reader_frozen_in_build_inputs": False,
    }, sort_keys=True))


if __name__ == "__main__":
    main()
