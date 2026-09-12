#!/usr/bin/env python3
"""Verify sealed InkAction profile receipts and recompute the report.

The verifier is deliberately strict about receipt provenance. A JSON file is
not accepted as a fresh process observation unless its companion timing and
stderr receipts prove a successful, silent `/usr/bin/time -v` invocation.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import subprocess
from pathlib import Path

import profile_pins


SCHEMA = "ink-action-edit-profile-v2"
SOURCE_MANIFEST_FORMAT = "ink-action-edit-build-source-v2"

LANES = (
    "draft_small_8",
    "draft_scaled_128",
    "draft_near_1024",
    "draft_opaque_64",
    "scalar_edit_small_8",
    "scalar_edit_scaled_128",
    "scalar_edit_near_1024",
    "scalar_batch_scaled_128",
    "scalar_batch_near_1024",
    "scalar_coalesce_scaled_128",
    "scalar_coalesce_near_1024",
    "no_op_small_8",
    "no_op_scaled_128",
    "no_op_near_1024",
    "add_small_8",
    "add_scaled_128",
    "add_near_1024",
    "insert_batch_scaled_128",
    "insert_batch_near_1024",
    "remove_small_8",
    "remove_scaled_128",
    "remove_near_1024",
    "remove_batch_scaled_128",
    "remove_batch_near_1024",
    "clear_batch_scaled_128",
    "clear_batch_near_1024",
    "move_small_8",
    "move_scaled_128",
    "move_near_1024",
    "move_batch_scaled_128",
    "move_batch_near_1024",
    "cap_refusal_small_8",
    "cap_refusal_scaled_128",
    "cap_refusal_near_1024",
)
CAP_REFUSALS = {lane for lane in LANES if lane.startswith("cap_refusal_")}
NO_OPS = {lane for lane in LANES if lane.startswith("no_op_")}


def require(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def resolve(root: Path, shown: str) -> Path:
    path = Path(shown)
    return path if path.is_absolute() else root / path


def positive_int(value: object, field: str, path: Path) -> int:
    require(type(value) is int and value > 0, f"{field} is not a positive integer: {path}")
    return int(value)


def nonnegative_int(value: object, field: str, path: Path) -> int:
    require(type(value) is int and value >= 0, f"{field} is negative or malformed: {path}")
    return int(value)


def bool_field(value: object, field: str, path: Path) -> bool:
    require(type(value) is bool, f"{field} is not boolean: {path}")
    return bool(value)


def committed_blob_sha256(root: Path, commit: str, shown: str) -> str:
    """Hash committed blob bytes without consulting the Git index."""

    try:
        committed = subprocess.check_output(
            [
                "git",
                "--no-replace-objects",
                "-C",
                str(root),
                "cat-file",
                "blob",
                f"{commit}:{shown}",
            ]
        )
    except subprocess.CalledProcessError as error:
        raise AssertionError(f"cannot read committed Git input blob: {shown}") from error
    return hashlib.sha256(committed).hexdigest()


def verify_git_snapshot(paths: list[Path], root: Path, commit: str) -> None:
    """Verify local manifest inputs equal their committed Git blobs.

    The comparison reads the working-tree bytes directly.  A Git diff is not
    sufficient because index flags such as ``assume-unchanged`` can hide a
    changed tracked file.
    """

    root = root.resolve()
    relative: list[str] = []
    for path in sorted({path.resolve() for path in paths}, key=str):
        require(path.is_file(), f"retained Git input missing: {path}")
        try:
            relative.append(path.relative_to(root).as_posix())
        except ValueError as error:
            raise AssertionError(
                f"non-Git local input has no retained source snapshot: {path}"
            ) from error
    if not relative:
        return
    current = subprocess.check_output(
        ["git", "-C", str(root), "rev-parse", "HEAD"], text=True
    ).strip()
    require(current == commit, f"Git snapshot commit changed: {commit} -> {current}")

    missing: list[str] = []
    for shown in relative:
        tracked = subprocess.run(
            ["git", "-C", str(root), "ls-files", "--error-unmatch", "--", shown],
            capture_output=True,
            text=True,
            check=False,
        )
        if tracked.returncode != 0:
            missing.append(shown)
    require(not missing, "local Git inputs are not tracked: " + ", ".join(sorted(missing)))

    changed: list[str] = []
    for shown in relative:
        committed_blob = committed_blob_sha256(root, commit, shown)
        working_blob = sha256(root / shown)
        if working_blob != committed_blob:
            changed.append(shown)
    require(
        not changed,
        "local Git inputs differ from committed snapshot: " + ", ".join(sorted(set(changed))),
    )


def verify_source_manifest(path: Path, root: Path) -> tuple[int, str]:
    lines = path.read_text().splitlines()
    require(
        lines and lines[0] == f"format={SOURCE_MANIFEST_FORMAT}",
        "manifest format changed",
    )
    git_commit: str | None = None
    packages: dict[tuple[str, str, str], tuple[str, int, str]] = {}
    files: dict[tuple[str, str, str], list[tuple[str, str]]] = {}
    extras: list[tuple[str, str]] = []
    for line in lines[1:]:
        if line.startswith("git_commit="):
            require(git_commit is None, "duplicate Git commit line")
            git_commit = line.split("=", 1)[1]
            continue
        elif line.startswith("metadata_sha256="):
            continue
        parts = line.split("\t")
        if line.startswith("package="):
            require(len(parts) == 7, f"malformed package line: {line}")
            key = (parts[0][len("package=") :], parts[1], parts[3])
            require(key not in packages, f"duplicate package line: {key}")
            packages[key] = (parts[2], int(parts[5]), parts[6])
        elif line.startswith("file="):
            require(len(parts) == 5, f"malformed file line: {line}")
            key = (parts[0][len("file=") :], parts[1], parts[2])
            files.setdefault(key, []).append((parts[3], parts[4]))
        elif line.startswith("extra=\t"):
            require(len(parts) == 3, f"malformed extra line: {line}")
            extras.append((parts[1], parts[2]))
        else:
            raise AssertionError(f"unknown manifest line: {line}")
    require(git_commit is not None and git_commit, "manifest Git commit missing")
    require(set(packages) == set(files), "package/file manifest sets differ")
    checked = 0
    local_inputs: list[Path] = []
    for key, (source, expected_count, expected_tree) in packages.items():
        entries = files[key]
        require(len(entries) == expected_count, f"package file count changed: {key}")
        tree_payload = "\n".join(f"{shown}\t{digest}" for shown, digest in entries)
        require(
            hashlib.sha256(tree_payload.encode()).hexdigest() == expected_tree,
            f"package tree changed: {key}",
        )
        for shown, expected in entries:
            current = resolve(root, shown)
            require(current.is_file(), f"retained manifest path missing: {shown}")
            require(sha256(current) == expected, f"retained manifest hash changed: {shown}")
            if source == "path":
                local_inputs.append(current)
            checked += 1
    for shown, expected in extras:
        current = resolve(root, shown)
        require(current.is_file(), f"retained extra path missing: {shown}")
        require(sha256(current) == expected, f"retained extra hash changed: {shown}")
        local_inputs.append(current)
        checked += 1
    verify_git_snapshot(local_inputs, root, git_commit)
    require(checked >= 10, f"too few retained manifest inputs checked: {checked}")
    return checked, git_commit


def quantile(values: list[int], percent: int) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * percent / 100) - 1)]


def rss(path: Path) -> int:
    values = [
        line.split(":", 1)[1].strip()
        for line in path.read_text().splitlines()
        if line.lstrip().startswith("Maximum resident set size (kbytes):")
    ]
    require(len(values) == 1, f"RSS must have exactly one reading: {path}")
    value = values[0]
    require(value.isascii() and value.isdigit() and int(value) > 0, f"RSS is malformed: {path}")
    return int(value)


def expected_action_count(lane: str) -> int:
    if lane.endswith("_small_8"):
        return 8
    if lane.endswith("_scaled_128"):
        return 128
    if lane.endswith("_near_1024"):
        return 1024
    if lane.endswith("_opaque_64"):
        return 64
    raise AssertionError(f"unknown action-count suffix: {lane}")


def expected_result_action_count(lane: str) -> int:
    count = expected_action_count(lane)
    if lane.startswith("add_"):
        return count + 1
    if lane.startswith("insert_batch_"):
        return count * 2
    if lane.startswith("remove_") and not lane.startswith("remove_batch_"):
        return count - 1
    if lane.startswith("remove_batch_"):
        return count // 2
    return count


def expected_operation_count(lane: str) -> int:
    count = expected_action_count(lane)
    if lane.startswith(("scalar_batch_", "scalar_coalesce_", "insert_batch_")):
        return count
    if lane.startswith(("remove_batch_", "clear_batch_", "move_batch_")):
        return count // 2
    if lane.startswith("draft_"):
        return count
    if lane.startswith("cap_refusal_"):
        return 1
    return 1


def verify_sample(
    sample: dict[str, object],
    expected_success: bool,
    lane: str,
    input_bytes: int,
    path: Path,
) -> None:
    require(
        bool_field(sample.get("expected_success"), "sample expected_success", path) is expected_success,
        f"sample expected status mismatch: {path}",
    )
    require(
        bool_field(sample.get("actual_success"), "sample actual_success", path) is expected_success,
        f"sample actual status mismatch: {path}",
    )

    for field in (
        "elapsed_ns",
        "requested_alloc_bytes",
        "direct_allocated_bytes",
        "realloc_old_bytes",
        "realloc_new_bytes",
        "deallocated_bytes",
        "alloc_calls",
        "realloc_calls",
        "dealloc_calls",
        "live_before",
        "live_after",
        "peak_live_delta",
        "alloc_failed",
    ):
        nonnegative_int(sample.get(field), field, path)
    require(
        bool_field(sample.get("alloc_balance_ok"), "alloc_balance_ok", path),
        f"allocator balance failed: {path}",
    )
    require(
        not bool_field(sample.get("alloc_invalid"), "alloc_invalid", path),
        f"allocator invalid: {path}",
    )
    require(sample["alloc_failed"] == 0, f"allocation failure failed: {path}")

    direct = int(sample["direct_allocated_bytes"])
    new = int(sample["realloc_new_bytes"])
    old = int(sample["realloc_old_bytes"])
    freed = int(sample["deallocated_bytes"])
    require(
        int(sample["requested_alloc_bytes"]) == direct + new,
        f"allocation equation failed: {path}",
    )
    require(
        int(sample["live_before"]) + direct + new
        == int(sample["live_after"]) + old + freed,
        f"live equation failed: {path}",
    )
    require(
        int(sample["peak_live_delta"]) >= max(0, int(sample["live_after"]) - int(sample["live_before"])),
        f"peak live delta is below retained growth: {path}",
    )

    optional_fields = (
        "semantic_ok",
        "source_exact",
        "source_shared",
        "inverse_ok",
        "opaque_preserved",
        "output_exact",
        "rejection_ok",
        "rejection_source_unchanged",
        "rejection_state_unchanged",
    )
    for field in optional_fields:
        value = sample.get(field)
        require(value is None or type(value) is bool, f"{field} is not boolean or null: {path}")
    rejection_resource = sample.get("rejection_resource")
    rejection_limit = sample.get("rejection_limit")
    require(
        rejection_resource is None or type(rejection_resource) is str,
        f"rejection_resource malformed: {path}",
    )
    require(
        rejection_limit is None or type(rejection_limit) is int,
        f"rejection_limit malformed: {path}",
    )
    rejection_pre_hash = sample.get("rejection_pre_source_hash")
    rejection_post_hash = sample.get("rejection_post_source_hash")
    require(
        rejection_pre_hash is None or (type(rejection_pre_hash) is int and rejection_pre_hash >= 0),
        f"rejection_pre_source_hash malformed: {path}",
    )
    require(
        rejection_post_hash is None or (type(rejection_post_hash) is int and rejection_post_hash >= 0),
        f"rejection_post_source_hash malformed: {path}",
    )

    if lane.startswith("draft_"):
        require(sample["semantic_ok"] is True, f"draft semantic gate failed: {path}")
        require(sample["source_exact"] is None, f"draft source_exact must be N/A: {path}")
        require(sample["source_shared"] is None, f"draft source_shared must be N/A: {path}")
        require(sample["inverse_ok"] is None, f"draft inverse_ok must be N/A: {path}")
        require(sample["opaque_preserved"] is True, f"draft opaque gate failed: {path}")
        require(sample["output_exact"] is True, f"draft output gate failed: {path}")
        require(sample["rejection_ok"] is None, f"draft rejection_ok must be N/A: {path}")
        require(sample["rejection_source_unchanged"] is None, f"draft rejection source check must be N/A: {path}")
        require(sample["rejection_state_unchanged"] is None, f"draft rejection state check must be N/A: {path}")
        require(
            rejection_resource is None and rejection_limit is None
            and rejection_pre_hash is None and rejection_post_hash is None,
            f"draft refusal fields must be N/A: {path}",
        )
    elif expected_success:
        for field in ("semantic_ok", "source_exact", "inverse_ok", "opaque_preserved", "output_exact"):
            require(sample[field] is True, f"{field} gate failed: {path}")
        require(sample["source_shared"] is (lane in NO_OPS), f"source-sharing gate failed: {path}")
        require(sample["rejection_ok"] is None, f"successful rejection_ok must be N/A: {path}")
        require(sample["rejection_source_unchanged"] is None, f"successful rejection source check must be N/A: {path}")
        require(sample["rejection_state_unchanged"] is None, f"successful rejection state check must be N/A: {path}")
        require(
            rejection_resource is None and rejection_limit is None
            and rejection_pre_hash is None and rejection_post_hash is None,
            f"successful refusal fields must be N/A: {path}",
        )
    else:
        for field in ("semantic_ok", "source_exact", "source_shared", "inverse_ok", "output_exact"):
            require(sample[field] is None, f"refusal {field} must be N/A: {path}")
        require(sample["opaque_preserved"] is True, f"refusal opaque gate failed: {path}")
        require(sample["rejection_ok"] is True, f"refusal gate failed: {path}")
        require(sample["rejection_source_unchanged"] is True, f"refusal source gate failed: {path}")
        require(sample["rejection_state_unchanged"] is True, f"refusal state gate failed: {path}")
        require(rejection_resource == "ink action output bytes", f"refusal resource mismatch: {path}")
        require(
            type(rejection_limit) is int and rejection_limit == input_bytes,
            f"refusal limit mismatch: {path}",
        )
        require(
            type(rejection_pre_hash) is int
            and type(rejection_post_hash) is int
            and rejection_pre_hash == rejection_post_hash,
            f"refusal source hash mismatch: {path}",
        )


def verify_process_output(path: Path) -> None:
    stderr = path.with_suffix(".stderr.log")
    require(stderr.is_file(), f"process stderr receipt missing: {path}")
    require(stderr.read_bytes() == b"", f"unexpected process stderr: {path}")
    timing = path.with_suffix(".time.txt")
    require(timing.is_file(), f"process timing receipt missing: {path}")
    lines = timing.read_text().splitlines()
    statuses = [line.strip() for line in lines if line.strip().startswith("Exit status:")]
    require(statuses == ["Exit status: 0"], f"process exit status is not successful: {path}")
    rss(timing)


def verify_lanes(results: Path) -> list[dict[str, object]]:
    rows = []
    for lane in LANES:
        paths = sorted(results.glob(f"{lane}-p*.json"))
        expected_names = [f"{lane}-p{n}.json" for n in range(1, 4)]
        require(
            [path.name for path in paths] == expected_names,
            f"fresh process identities mismatch for {lane}",
        )
        expected_success = lane not in CAP_REFUSALS
        all_samples: list[dict[str, object]] = []
        rss_values = []
        input_bytes: set[int] = set()
        action_counts: set[int] = set()
        result_action_counts: set[int] = set()
        operation_counts: set[int] = set()
        pids: set[int] = set()
        for path in paths:
            verify_process_output(path)
            value = json.loads(path.read_text())
            require(type(value.get("schema")) is str and value["schema"] == SCHEMA, f"schema mismatch in {path}")
            require(type(value.get("lane")) is str and value["lane"] == lane, f"lane mismatch in {path}")
            pids.add(positive_int(value.get("pid"), "pid", path))
            require(type(value.get("warmup")) is int and value["warmup"] == 2, f"warm-up mismatch in {path}")
            require(type(value.get("sample_count")) is int and value["sample_count"] == 20, f"sample count metadata mismatch in {path}")
            require(isinstance(value.get("samples"), list), f"samples are not an array in {path}")
            require(
                value["sample_count"] == len(value["samples"]) == 20,
                f"sample count mismatch in {path}",
            )
            require(
                bool_field(value.get("expected_success"), "expected_success", path)
                is expected_success,
                f"expected status mismatch in {path}",
            )
            require(
                positive_int(value.get("action_count"), "action_count", path)
                == expected_action_count(lane),
                f"action count mismatch in {path}",
            )
            require(
                positive_int(value.get("result_action_count"), "result_action_count", path)
                == expected_result_action_count(lane),
                f"result action count mismatch in {path}",
            )
            require(
                positive_int(value.get("operation_count"), "operation_count", path)
                == expected_operation_count(lane),
                f"operation count mismatch in {path}",
            )
            current_input_bytes = nonnegative_int(value.get("input_bytes"), "input_bytes", path)
            if lane.startswith("draft_"):
                require(current_input_bytes == 0, f"draft input bytes must be zero: {path}")
            else:
                require(current_input_bytes > 0, f"source-backed input bytes must be positive: {path}")
            input_bytes.add(current_input_bytes)
            action_counts.add(int(value["action_count"]))
            result_action_counts.add(int(value["result_action_count"]))
            operation_counts.add(int(value["operation_count"]))
            rss_values.append(rss(path.with_suffix(".time.txt")))
            for sample in value["samples"]:
                require(isinstance(sample, dict), f"sample is not an object: {path}")
                verify_sample(sample, expected_success, lane, current_input_bytes, path)
                all_samples.append(sample)
        require(len(pids) == 3, f"process identities are not distinct for {lane}")
        require(len(input_bytes) == 1, f"input size changed across {lane}")
        require(len(action_counts) == 1, f"action count changed across {lane}")
        require(len(result_action_counts) == 1, f"result action count changed across {lane}")
        require(len(operation_counts) == 1, f"operation count changed across {lane}")
        elapsed = [int(sample["elapsed_ns"]) for sample in all_samples]
        allocated = [int(sample["requested_alloc_bytes"]) for sample in all_samples]
        peak = [int(sample["peak_live_delta"]) for sample in all_samples]
        rows.append(
            {
                "lane": lane,
                "processes": len(paths),
                "samples": len(all_samples),
                "action_count": next(iter(action_counts)),
                "result_action_count": next(iter(result_action_counts)),
                "operation_count": next(iter(operation_counts)),
                "input_bytes": next(iter(input_bytes)),
                "elapsed": (quantile(elapsed, 50), quantile(elapsed, 95), quantile(elapsed, 99)),
                "alloc": (quantile(allocated, 50), quantile(allocated, 95)),
                "peak": (quantile(peak, 50), quantile(peak, 95)),
                "rss": (min(rss_values), max(rss_values)),
            }
        )
    return rows


def verify_binary_receipts(results: Path) -> str:
    before_path = results / "binary.sha256"
    after_path = results / "binary-after.sha256"
    before_lines = before_path.read_text().splitlines()
    after_lines = after_path.read_text().splitlines()
    require(
        len(before_lines) == 1 and before_lines == after_lines,
        "profile executable changed during measurement",
    )
    fields = before_lines[0].split()
    require(len(fields) == 2, "malformed executable digest receipt")
    digest, shown = fields
    require(len(digest) == 64 and all(c in "0123456789abcdef" for c in digest), "malformed executable SHA-256")
    require(shown.startswith("/") and shown.endswith("/ink-action-edit-profile"), "malformed executable path")
    build_lines = (results / "build-provenance.txt").read_text().splitlines()
    binary_lines = [line for line in build_lines if line.startswith("binary=")]
    require(binary_lines == [f"binary={shown}"], "build provenance binary path mismatch")
    require(f"{digest}  {shown}" in build_lines, "build provenance binary digest mismatch")
    live_binary = Path(shown)
    if live_binary.is_file():
        require(sha256(live_binary) == digest, "live profile executable hash mismatch")
    return digest


def verify_provenance(results: Path, root: Path, manifest_commit: str) -> tuple[str, int]:
    provenance = (results / "source-provenance.txt").read_text()
    require(
        f"approved_base_commit={profile_pins.APPROVED_BASE_COMMIT}" in provenance,
        "profile is not pinned to the approved commit",
    )
    profile_arm = next(
        (line.split("=", 1)[1] for line in provenance.splitlines() if line.startswith("profile_arm=")),
        "",
    )
    try:
        source_pin, expected_hashes = profile_pins.profile(profile_arm)
    except ValueError as error:
        raise AssertionError(f"unknown profile arm in provenance: {profile_arm}") from error
    recorded_pin = next(
        (line.split("=", 1)[1] for line in provenance.splitlines() if line.startswith("source_pin=")),
        "",
    )
    require(recorded_pin == source_pin, "profile source pin does not match its arm")
    current_commit = next(
        (line.split("=", 1)[1] for line in provenance.splitlines() if line.startswith("git_head=")),
        "",
    )
    require(
        current_commit
        and current_commit
        == subprocess.check_output(["git", "-C", str(root), "rev-parse", "HEAD"], text=True).strip(),
        "profile git head changed after capture",
    )
    require(current_commit == manifest_commit, "profile Git head differs from source manifest snapshot")
    require(
        subprocess.run(
            ["git", "-C", str(root), "merge-base", "--is-ancestor", source_pin, current_commit],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        ).returncode
        == 0,
        "profile source pin is not an ancestor of the captured head",
    )
    require("git_status_relevant=\n" in provenance, "relevant source tree was dirty during profile")
    source_hashes: dict[str, str] = {}
    for line in provenance.splitlines():
        fields = line.split("  ", 1)
        if len(fields) != 2:
            continue
        digest, shown = fields
        if len(digest) == 64 and shown.startswith("/"):
            current = Path(shown)
            require(current.is_file(), f"source hash path missing: {shown}")
            require(sha256(current) == digest, f"source hash changed: {shown}")
            source_hashes[str(current)] = digest
    for relative, expected in expected_hashes.items():
        current = root / relative
        require(
            source_hashes.get(str(current)) == expected,
            f"approved source hash missing or changed: {relative}",
        )
    evidence_dir = root / "docs/report/spec-gap-validation-evidence/ink-action-edit-performance"
    profile_pins_path = evidence_dir / "profile_pins.py"
    profile_pins_receipt = next(
        (line.split("=", 1)[1] for line in provenance.splitlines() if line.startswith("profile_pins_sha256=")),
        "",
    )
    require(
        profile_pins_receipt == sha256(profile_pins_path),
        "profile pin manifest hash missing or changed",
    )
    for test_name in ("test_verify.py", "test_smoke_target.py", "test_source_snapshot.py"):
        test_path = evidence_dir / test_name
        require(source_hashes.get(str(test_path)) == sha256(test_path), f"test manifest hash missing or changed: {test_name}")
    return provenance, len(source_hashes)


def verify_host_and_commands(results: Path) -> None:
    host = (results / "host.txt").read_text()
    for marker in ("utc=", "rustc", "cargo", "Linux"):
        require(marker in host, f"host provenance missing {marker}")
    commands = (results / "commands.txt").read_text()
    for marker in (
        "profile_arm=",
        "source_pin=",
        "cargo metadata --format-version=1 --locked --offline",
        "cargo build --release --locked --offline",
        "/usr/bin/time -v",
    ):
        require(marker in commands, f"profile command provenance missing {marker}")
    for lane in LANES:
        require(commands.count(f"--lane {lane} ") == 3, f"profile command count mismatch: {lane}")
    build = (results / "build-provenance.txt").read_text()
    for marker in (
        "binary=",
        "profile_arm=",
        "source_pin=",
        "rustc -vV:",
        "cargo=",
        "target=",
        "cargo_incremental=0",
        "flags=none",
    ):
        require(marker in build, f"build provenance missing {marker}")


def verify_report(path: Path, rows: list[dict[str, object]]) -> None:
    parsed: dict[str, dict[str, object]] = {}
    for line in path.read_text().splitlines():
        if not line.startswith("| ") or line.startswith("|---") or line.startswith("| lane"):
            continue
        parts = [part.strip() for part in line.strip().strip("|").split("|")]
        require(len(parts) == 11, f"malformed report row: {line}")
        lane = parts[0]
        require(lane in LANES and lane not in parsed, f"unexpected report lane: {lane}")
        elapsed = tuple(int(value.strip()) for value in parts[7].split("/"))
        alloc = tuple(int(value.strip()) for value in parts[8].split("/"))
        peak = tuple(int(value.strip()) for value in parts[9].split("/"))
        rss_values = tuple(int(value.strip()) for value in parts[10].replace("–", "-").split("-"))
        require(
            len(elapsed) == 3 and len(alloc) == 2 and len(peak) == 2 and len(rss_values) == 2,
            f"report metrics malformed: {line}",
        )
        parsed[lane] = {
            "processes": int(parts[1]),
            "samples": int(parts[2]),
            "action_count": int(parts[3]),
            "result_action_count": int(parts[4]),
            "operation_count": int(parts[5]),
            "input_bytes": int(parts[6]),
            "elapsed": elapsed,
            "alloc": alloc,
            "peak": peak,
            "rss": rss_values,
        }
    require(set(parsed) == set(LANES), "report lane set differs")
    for row in rows:
        actual = parsed[str(row["lane"])]
        for field in (
            "processes",
            "samples",
            "action_count",
            "result_action_count",
            "operation_count",
            "input_bytes",
            "elapsed",
            "alloc",
            "peak",
            "rss",
        ):
            require(actual[field] == row[field], f"report {field} differs: {row['lane']}")


def main() -> None:
    parser = argparse.ArgumentParser()
    here = Path(__file__).resolve().parent
    parser.add_argument("--results", type=Path, default=here / "results")
    parser.add_argument("--report", type=Path, default=here / "report.md")
    args = parser.parse_args()
    root = next(path for path in here.parents if (path / "crates").is_dir())
    results = args.results
    before = results / "source-manifest-before.txt"
    after = results / "source-manifest-after.txt"
    require(before.read_bytes() == after.read_bytes(), "source manifest changed during profile")
    manifest_sha = sha256(before)
    manifest_count, manifest_commit = verify_source_manifest(before, root)
    provenance = (results / "source-provenance.txt").read_text()
    require(f"source_manifest_before_sha256={manifest_sha}" in provenance, "manifest hash missing")
    require(f"source_manifest_after_sha256={manifest_sha}" in provenance, "post-build manifest hash missing")
    _, source_count = verify_provenance(results, root, manifest_commit)
    verify_host_and_commands(results)
    binary_digest = verify_binary_receipts(results)
    rows = verify_lanes(results)
    report = args.report.read_text()
    require("no before/after speedup" in report.lower(), "report omitted no-speedup scope")
    require("no package-wide performance claim" in report.lower(), "report omitted package scope")
    verify_report(args.report, rows)
    result = {
        "passed": True,
        "schema": SCHEMA,
        "lanes": len(LANES),
        "processes_per_lane": 3,
        "samples_per_process": 20,
        "source_manifest_sha256": manifest_sha,
        "manifest_inputs_checked": manifest_count,
        "source_hashes_checked": source_count,
        "binary_sha256": binary_digest,
        "expected_rejections": sorted(CAP_REFUSALS),
        "report_recomputed_from_samples": True,
    }
    (results / "verification.json").write_text(json.dumps(result, indent=2) + "\n")
    print(json.dumps(result, indent=2))


if __name__ == "__main__":
    main()
