#!/usr/bin/env python3
"""Compare the 0433 strict harness diagnostics with the retained 0429 log.

The baseline is the exact unmodified 0429 ``harness-strict-final.log`` stream
(SHA-256 ``db722f...d4dee5``).  During generation, the changed source region
is derived from the current worktree versus ``be82ef53d`` for the owned
``tools/perf-baseline/src/lib.rs`` file.  The unified-0 patch and its binding
are retained beside the comparison, so the source may still be dirty while
the comparison is prepared before the final commit.

Prepare the retained baseline at
``checks/baseline-harness-strict.log.gz`` (or the corresponding uncompressed
name), then run the corrected strict command whose retained output is
``checks/harness-strict-final-v2.log`` and invoke this script.  The initial
failed attempt remains at ``checks/harness-strict-final.log`` as development
evidence.  ``--check``
recomputes the comparison from retained logs, patch, and binding only; it does
not require a Git checkout.  The JSON result pins source hashes, worktree
status, and the dynamic current-file line ranges for a later portable reader.
"""

from __future__ import annotations

import argparse
from collections import Counter
import gzip
import hashlib
import json
from pathlib import Path
import re
import subprocess
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 433
BASELINE_REVISION = "be82ef53d"
SOURCE_FILE = "tools/perf-baseline/src/lib.rs"
BASELINE_ORIGIN = "change-0429/checks/harness-strict-final.log.gz"
BASELINE_RAW_SHA256 = "db722f454041840d5e92ceed093c5bfa40936f78aa31d729c520b35a20d4dee5"
BASELINE_LOG = "checks/baseline-harness-strict.log"
CURRENT_LOG = "checks/harness-strict-final-v2.log"
INITIAL_ATTEMPT_LOG = "checks/harness-strict-final.log"
CHANGED_PATCH = "checks/strict-changed-source.patch"
SOURCE_BINDING = "checks/strict-source-binding.json"
DIAGNOSTIC_RE = re.compile(
    r"^error: (.+)\n\s*--> ([^\n]+?):(\d+):\d+",
    re.MULTILINE,
)
HUNK_RE = re.compile(r"^@@ -\d+(?:,\d+)? \+(\d+)(?:,(\d+))? @@", re.MULTILINE)


class Invalid(ValueError):
    pass


def fail(message: str) -> None:
    raise Invalid(message)


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def raw(path_name: str) -> tuple[bytes, str]:
    path = ROOT / path_name
    if path.is_file():
        return path.read_bytes(), path_name
    compressed = Path(str(path) + ".gz")
    if compressed.is_file():
        try:
            value = gzip.decompress(compressed.read_bytes())
        except (OSError, EOFError, gzip.BadGzipFile) as error:
            fail(f"{compressed}: invalid gzip stream: {error}")
        return value, path_name
    fail(f"missing {path_name} and {path_name}.gz")
    raise AssertionError("unreachable")


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def git(*args: str) -> bytes:
    try:
        return subprocess.check_output(["git", *args], cwd=ROOT.parents[3])
    except subprocess.CalledProcessError as error:
        fail(f"git {' '.join(args)} failed with exit {error.returncode}")
    raise AssertionError("unreachable")


def working_source_binding() -> tuple[dict[str, Any], bytes]:
    """Capture source identity and the worktree-vs-base patch.

    This intentionally permits a dirty source file.  The strict comparison is
    a pre-commit gate; the later build/source-manifest verifier binds this
    worktree hash to the final before/after binaries.
    """

    baseline_revision = git("rev-parse", BASELINE_REVISION).decode().strip()
    current_head_revision = git("rev-parse", "HEAD").decode().strip()
    status = git("status", "--porcelain", "--", SOURCE_FILE).decode().splitlines()
    baseline_source = git("show", f"{baseline_revision}:{SOURCE_FILE}")
    current_source = (ROOT.parents[3] / SOURCE_FILE).read_bytes()
    patch = git("diff", "--unified=0", baseline_revision, "--", SOURCE_FILE)
    patch_text = patch.decode("utf-8")
    binding = {
        "baseline_revision_input": BASELINE_REVISION,
        "baseline_revision": baseline_revision,
        "current_head_revision": current_head_revision,
        "current_revision_binding": "working-tree source at comparison time",
        "file": SOURCE_FILE,
        "line_range_convention": "1-based inclusive current-file lines",
        "baseline_sha256": sha(baseline_source),
        "current_worktree_sha256": sha(current_source),
        "tracked_status": status,
        "patch_path": CHANGED_PATCH,
        "patch_format": "git diff --unified=0",
        "patch_sha256": sha(patch),
        "patch_bytes": len(patch),
        "changed_line_ranges": changed_line_ranges(patch_text),
    }
    return binding, patch


def write_source_binding(binding: dict[str, Any], patch: bytes) -> None:
    patch_path = ROOT / CHANGED_PATCH
    binding_path = ROOT / SOURCE_BINDING
    if patch_path.exists() or binding_path.exists():
        fail(f"refusing to overwrite retained source binding artifacts: {CHANGED_PATCH}, {SOURCE_BINDING}")
    patch_path.write_bytes(patch)
    binding_path.write_text(json.dumps(binding, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def hash_text(value: Any, name: str) -> str:
    if not isinstance(value, str) or re.fullmatch(r"[0-9a-f]{64}", value) is None:
        fail(f"{name} must be a lowercase SHA-256 hex string")
    return value


def nonempty_text(value: Any, name: str) -> str:
    if not isinstance(value, str) or not value:
        fail(f"{name} must be non-empty text")
    return value


def load_source_binding() -> dict[str, Any]:
    """Load and validate retained source artifacts without invoking Git."""

    binding_path = ROOT / SOURCE_BINDING
    patch_path = ROOT / CHANGED_PATCH
    if not binding_path.is_file():
        fail(f"missing {SOURCE_BINDING}")
    if not patch_path.is_file():
        fail(f"missing {CHANGED_PATCH}")
    binding = load_json(binding_path)
    if not isinstance(binding, dict):
        fail(f"{SOURCE_BINDING} must contain a JSON object")
    required = (
        "baseline_revision_input",
        "baseline_revision",
        "current_head_revision",
        "current_revision_binding",
        "file",
        "line_range_convention",
        "baseline_sha256",
        "current_worktree_sha256",
        "tracked_status",
        "patch_path",
        "patch_format",
        "patch_sha256",
        "patch_bytes",
        "changed_line_ranges",
    )
    for key in required:
        if key not in binding:
            fail(f"{SOURCE_BINDING} is missing {key}")
    if binding["baseline_revision_input"] != BASELINE_REVISION:
        fail(f"{SOURCE_BINDING}.baseline_revision_input does not match {BASELINE_REVISION}")
    nonempty_text(binding["baseline_revision"], f"{SOURCE_BINDING}.baseline_revision")
    nonempty_text(binding["current_head_revision"], f"{SOURCE_BINDING}.current_head_revision")
    if binding["current_revision_binding"] != "working-tree source at comparison time":
        fail(f"{SOURCE_BINDING}.current_revision_binding has an unsupported value")
    if binding["file"] != SOURCE_FILE:
        fail(f"{SOURCE_BINDING}.file must bind {SOURCE_FILE}")
    if binding["line_range_convention"] != "1-based inclusive current-file lines":
        fail(f"{SOURCE_BINDING}.line_range_convention is unsupported")
    hash_text(binding["baseline_sha256"], f"{SOURCE_BINDING}.baseline_sha256")
    hash_text(binding["current_worktree_sha256"], f"{SOURCE_BINDING}.current_worktree_sha256")
    status = binding["tracked_status"]
    if not isinstance(status, list) or any(not isinstance(item, str) for item in status):
        fail(f"{SOURCE_BINDING}.tracked_status must be a list of strings")
    if binding["patch_path"] != CHANGED_PATCH:
        fail(f"{SOURCE_BINDING}.patch_path must be {CHANGED_PATCH}")
    if binding["patch_format"] != "git diff --unified=0":
        fail(f"{SOURCE_BINDING}.patch_format is unsupported")
    patch = patch_path.read_bytes()
    hash_text(binding["patch_sha256"], f"{SOURCE_BINDING}.patch_sha256")
    if binding["patch_sha256"] != sha(patch):
        fail(f"{SOURCE_BINDING}.patch_sha256 does not match {CHANGED_PATCH}")
    patch_bytes = binding["patch_bytes"]
    if type(patch_bytes) is not int or patch_bytes < 0:
        fail(f"{SOURCE_BINDING}.patch_bytes must be a non-negative integer")
    if patch_bytes != len(patch):
        fail(f"{SOURCE_BINDING}.patch_bytes does not match {CHANGED_PATCH}")
    try:
        patch_text = patch.decode("utf-8")
    except UnicodeDecodeError as error:
        fail(f"{CHANGED_PATCH} is not UTF-8: {error}")
    ranges = binding["changed_line_ranges"]
    if not isinstance(ranges, list) or any(
        not isinstance(item, dict)
        or type(item.get("start")) is not int
        or type(item.get("end")) is not int
        or item["start"] < 1
        or item["end"] < item["start"]
        for item in ranges
    ):
        fail(f"{SOURCE_BINDING}.changed_line_ranges is malformed")
    if ranges != changed_line_ranges(patch_text):
        fail(f"{SOURCE_BINDING}.changed_line_ranges does not match {CHANGED_PATCH}")
    return binding


def changed_line_ranges(diff: str) -> list[dict[str, int]]:
    ranges: list[tuple[int, int]] = []
    for match in HUNK_RE.finditer(diff):
        start = int(match.group(1))
        count = int(match.group(2) or "1")
        if count == 0:
            continue
        ranges.append((start, start + count - 1))
    ranges.sort()
    merged: list[tuple[int, int]] = []
    for start, end in ranges:
        if merged and start <= merged[-1][1] + 1:
            merged[-1] = (merged[-1][0], max(merged[-1][1], end))
        else:
            merged.append((start, end))
    return [{"start": start, "end": end} for start, end in merged]


def diagnostic_rows(data: bytes) -> list[dict[str, Any]]:
    text = data.decode("utf-8", errors="replace")
    return [
        {"message": message, "file": path, "line": int(line)}
        for message, path, line in DIAGNOSTIC_RE.findall(text)
    ]


def findings(rows: list[dict[str, Any]]) -> Counter[tuple[str, str]]:
    return Counter((row["message"], row["file"]) for row in rows)


def in_changed_source(row: dict[str, Any], binding: dict[str, Any]) -> bool:
    path = row["file"].replace("\\", "/")
    if path.startswith("tools/perf-baseline/"):
        path = path.removeprefix("tools/perf-baseline/")
    if path != "src/lib.rs":
        return False
    line = row["line"]
    return any(item["start"] <= line <= item["end"] for item in binding["changed_line_ranges"])


def finding_list(counter: Counter[tuple[str, str]]) -> list[dict[str, Any]]:
    return [
        {"message": message, "file": path, "count": count}
        for (message, path), count in sorted(counter.items())
    ]


def derive(binding: dict[str, Any]) -> dict[str, Any]:
    baseline, baseline_path = raw(BASELINE_LOG)
    current, current_path = raw(CURRENT_LOG)
    initial_attempt, initial_attempt_path = raw(INITIAL_ATTEMPT_LOG)
    if sha(baseline) != BASELINE_RAW_SHA256:
        fail(
            f"retained 0429 baseline digest differs: expected {BASELINE_RAW_SHA256}, "
            f"got {sha(baseline)}"
        )
    baseline_rows = diagnostic_rows(baseline)
    current_rows = diagnostic_rows(current)
    before = findings(baseline_rows)
    after = findings(current_rows)
    if not before:
        fail("baseline diagnostic extraction is empty")
    if before != after:
        fail(
            "strict diagnostic multiset changed: "
            + json.dumps({"added": finding_list(after - before), "removed": finding_list(before - after)}, sort_keys=True)
        )
    changed_rows = [row for row in current_rows if in_changed_source(row, binding)]
    changed_counter = findings(changed_rows)
    if changed_rows:
        fail(f"strict diagnostics landed in changed harness lines: {finding_list(changed_counter)}")
    return {
        "status": "pass",
        "change": CHANGE,
        "baseline_origin": BASELINE_ORIGIN,
        "baseline_path": baseline_path,
        "baseline_original_sha256": sha(baseline),
        "baseline_original_bytes": len(baseline),
        "current_log": current_path,
        "current_original_sha256": sha(current),
        "current_original_bytes": len(current),
        "initial_attempt_log": initial_attempt_path,
        "initial_attempt_original_sha256": sha(initial_attempt),
        "initial_attempt_original_bytes": len(initial_attempt),
        "unique_findings": len(after),
        "same_message_and_source_file_multiset": True,
        "changed_harness_src_lib_findings": 0,
        "changed_line_findings": [],
        "changed_source_patch": CHANGED_PATCH,
        "source_binding_path": SOURCE_BINDING,
        "source_binding": binding,
        "findings": finding_list(after),
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    try:
        target = ROOT / "checks/strict-debt-comparison.json"
        if args.check:
            binding = load_source_binding()
            result = derive(binding)
            if not target.is_file():
                fail("checks/strict-debt-comparison.json is missing")
            if load_json(target) != result:
                fail("retained strict-debt-comparison.json differs from current logs/source")
        else:
            if target.exists():
                fail("refusing to overwrite checks/strict-debt-comparison.json")
            binding, patch = working_source_binding()
            write_source_binding(binding, patch)
            result = derive(binding)
            target.write_text(json.dumps(result, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        print(json.dumps({"status": "pass", "unique_findings": result["unique_findings"],
                          "changed_harness_src_lib_findings": 0}, sort_keys=True))
        return 0
    except (OSError, ValueError, AssertionError) as error:
        print(f"INVALID: {error}")
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
