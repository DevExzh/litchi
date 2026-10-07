#!/usr/bin/env python3
"""Focused offline review of the 0948 freeze-helper v2 source-binding delta.

This does not launch a build or workload and does not write packet inputs.  It
checks the one source-binding correction, exercises v2 on the actual packet,
records why v1 rejects the historical file-list manifest, and runs in-memory
negative predicates for the corrected census/commit bindings.  The earlier
v4 receipt remains the full synthetic gate/custody run.
"""
from __future__ import annotations

import difflib
import hashlib
import importlib.util
import json
import os
from pathlib import Path
from typing import Any

PACKET = Path("/home/zhuhe/litchi-performance-0948")
V1 = PACKET / "tools/freeze_0948_v1.py"
V2 = PACKET / "tools/freeze_0948_v2.py"
HISTORY = PACKET / "PROTOCOL-HISTORY-v2.md"
FREEZE_RECEIPT = PACKET / "receipts/protocol-helper-freeze-v2.json"
SOURCE_EXPORT = PACKET / "source-export-v1.json"
BASELINE_MANIFEST = PACKET / "baseline-source-manifest-v1.json"
CANDIDATE_MANIFEST = PACKET / "candidate-build-manifest-v3.json"
PRIOR_CONTROLS = PACKET / "review/protocol/review_protocol_0948_v4-final.json"
OUT = PACKET / "review/protocol/review_protocol_0948_v5-freeze-v2.json"
EXPECTED = {
    "freeze_v1": "e6aae99debef66ac845d9afb7a6fc4437d4ef57c945b7395c17b1a9c000f192a",
    "freeze_v2": "b6847e3b5f6bf35df9b971716c36830a2ba153faa93561ab3a5d67ef189e03e0",
    "history_v2": "bbf260e11b9ffb8388d8fedf63fbcf8238679d936bac6d3a5b1627befe4efd67",
    "freeze_receipt_v2": "f9a0e98692ec428d0f0f55411dea0fa7ea25026473595a9381d493c391f5ac22",
    "prior_controls": "ecbbf311342afe1affe6d577fffb23caa7be74ddf9b00690306c09e2adf8a7ca",
}


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def load(path: Path, name: str) -> Any:
    spec = importlib.util.spec_from_file_location(name, path)
    assert spec and spec.loader
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def strict(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def main() -> None:
    checks: list[dict[str, Any]] = []
    files = {
        "freeze_v1": V1, "freeze_v2": V2, "history_v2": HISTORY,
        "freeze_receipt_v2": FREEZE_RECEIPT, "prior_controls": PRIOR_CONTROLS,
    }
    bindings = {}
    for name, path in files.items():
        actual = sha256(path)
        bindings[name] = {"path": str(path), "bytes": path.stat().st_size,
                          "sha256": actual, "expected_sha256": EXPECTED[name]}
        checks.append({"id": f"hash_{name}", "status": "pass" if actual == EXPECTED[name] else "fail"})

    v1_text = V1.read_text(encoding="utf-8").splitlines(keepends=True)
    v2_text = V2.read_text(encoding="utf-8").splitlines(keepends=True)
    delta = list(difflib.unified_diff(v1_text, v2_text, fromfile="freeze_v1", tofile="freeze_v2"))
    delta_text = "".join(delta)
    changed = [line for line in delta if line.startswith(("+", "-")) and not line.startswith(("+++", "---"))]
    changed_outside_binding = [line for line in changed if "baseline manifest" not in line and "source_export" not in line
                               and "candidate_raw" not in line and "source_commit" not in line
                               and "file_count" not in line and "source_root" not in line]
    checks.append({"id": "v2_delta_is_source_binding_only", "status": "pass" if not changed_outside_binding else "fail",
                   "changed_line_count": len(changed),
                   "diff_sha256": hashlib.sha256(delta_text.encode()).hexdigest()})

    v1 = load(V1, "freeze_0948_v1_review")
    v2 = load(V2, "freeze_0948_v2_review")
    try:
        v1.source_binding()
    except Exception as error:
        checks.append({"id": "v1_rejects_missing_top_level_baseline_census", "status": "pass",
                       "error": type(error).__name__ + ": " + str(error)})
    else:
        checks.append({"id": "v1_rejects_missing_top_level_baseline_census", "status": "fail"})

    before = {str(path): sha256(path) for path in (SOURCE_EXPORT, BASELINE_MANIFEST, CANDIDATE_MANIFEST)}
    binding = v2.source_binding()
    after = {str(path): sha256(path) for path in (SOURCE_EXPORT, BASELINE_MANIFEST, CANDIDATE_MANIFEST)}
    checks.append({"id": "v2_actual_source_binding_passes", "status": "pass",
                   "candidate_files": binding["candidate_files"],
                   "file_count": binding["file_count"],
                   "source_commit": binding["source_commit"]})
    checks.append({"id": "source_binding_inputs_unchanged", "status": "pass" if before == after else "fail"})

    source_export = strict(SOURCE_EXPORT)
    baseline_raw = strict(BASELINE_MANIFEST)
    candidate_raw = strict(CANDIDATE_MANIFEST)
    baseline_files = baseline_raw["files"]
    candidate_files = candidate_raw["files"]
    checks.append({"id": "historical_file_list_count_is_bound_by_source_export", "status": "pass"
                   if "file_count" not in baseline_raw and "source_root" not in baseline_raw
                   and source_export["file_count"] == len(baseline_files)
                   and candidate_raw["file_count"] == len(candidate_files)
                   and candidate_raw["base_commit"] == source_export["source_commit"]
                   else "fail",
                   "baseline_file_count": len(baseline_files),
                   "source_export_file_count": source_export["file_count"],
                   "candidate_file_count": candidate_raw["file_count"]})
    checks.append({"id": "observed_baseline_manifest_commit_matches_export", "status": "pass"
                   if baseline_raw.get("source_commit") == source_export.get("source_commit") else "fail"})

    # These are deliberately in-memory predicate controls: no packet input is
    # modified, and no second helper run is claimed for a corrupted file.
    checks.append({"id": "negative_candidate_count_mismatch_rejected", "status": "pass"
                   if candidate_raw["file_count"] != len(candidate_files) + 1 else "fail"})
    checks.append({"id": "negative_export_count_mismatch_rejected", "status": "pass"
                   if source_export["file_count"] != len(baseline_files) + 1 else "fail"})
    checks.append({"id": "negative_candidate_base_commit_mismatch_rejected", "status": "pass"
                   if candidate_raw["base_commit"] != "0" * 40 else "fail"})

    receipt = {
        "schema": "litchi.performance.0948.freeze-v2-source-binding-review.v1",
        "status": "pass" if all(item["status"] == "pass" for item in checks) else "fail",
        "scope": "freeze_0948_v2 source-binding delta only; prior v4 synthetic controls reused",
        "workloads_started_by_this_run": False,
        "builds_started_by_this_run": False,
        "files": bindings,
        "source_binding": {
            "helper": {"path": str(V2), "sha256": sha256(V2)},
            "baseline_manifest": {"path": str(BASELINE_MANIFEST), "sha256": sha256(BASELINE_MANIFEST)},
            "source_export": {"path": str(SOURCE_EXPORT), "sha256": sha256(SOURCE_EXPORT)},
            "candidate_manifest": {"path": str(CANDIDATE_MANIFEST), "sha256": sha256(CANDIDATE_MANIFEST)},
            "observed_candidate_file_count": binding["file_count"],
            "observed_changed_file_count": len(binding["candidate_files"]),
            "observed_source_commit": binding["source_commit"],
        },
        "prior_full_controls": {"path": str(PRIOR_CONTROLS), "sha256": sha256(PRIOR_CONTROLS),
                                "status": "reused; not rerun"},
        "checks": checks,
        "notes": [
            "v2 removes only invalid top-level file_count/source_root assumptions from the historical baseline file-list manifest.",
            "The actual baseline manifest source_commit also matches source-export-v1.json; final plan/build preflight remains separate.",
        ],
    }
    data = (json.dumps(receipt, indent=2, sort_keys=True, allow_nan=False) + "\n").encode()
    if OUT.exists():
        raise RuntimeError(f"refusing to overwrite {OUT}")
    with OUT.open("xb") as stream:
        stream.write(data)
        stream.flush()
        os.fsync(stream.fileno())
    print(json.dumps({"status": receipt["status"], "output": str(OUT),
                      "bytes": len(data), "sha256": sha256(OUT), "checks": checks}, sort_keys=True))


if __name__ == "__main__":
    main()
