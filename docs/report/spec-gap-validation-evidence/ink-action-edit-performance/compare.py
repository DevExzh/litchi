#!/usr/bin/env python3
"""Compare the retained matched InkAction profile receipts.

This is a post-capture analysis tool.  It never starts the profile binary and
does not modify either arm's raw receipts.  Every comparison is scoped to a
lane and to the two retained source arms named below.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import re
import statistics
from pathlib import Path
from typing import Any, Mapping

import verify


HERE = Path(__file__).resolve().parent
DEFAULT_BASELINE = HERE / "results" / "paired-baseline-e5c18ca"
DEFAULT_CANDIDATE = HERE / "results" / "paired-candidate-251d361f"
DEFAULT_REPORT = HERE / "results" / "matched-report.md"
DEFAULT_JSON = HERE / "results" / "matched-report.json"

BASELINE_PIN = "f1cb119361af9ea2227d27050e41915a9a92ae04"
BASELINE_HEAD = "e5c18ca02e120ea1d91e495d4cf58df6d66573be"
CANDIDATE_PIN = "ab94a4d7a02053765bf4c70b4af5022273b821e5"
CANDIDATE_HEAD = "251d361f951a0cfc24fb57e0e0ad0f76044143a0"
GUARD_COMMIT = "03defc9d46a2f8ad2d9e5e0943586dec242e96a6"
PROFILE_PINS_SHA256 = "d6593e82a96423cc54b41f9f945f70219ffc2bd5506de520dbf114ec6afd709d"

# These are the production source entries that define the two matched arms.
# Keep them in the comparator so changing a retained manifest and merely
# updating its self-reported provenance cannot make the pair acceptable.
PINNED_SOURCE_ENTRIES = {
    "baseline": {
        "crates/litchi-drawingml/src/ink/actions.rs": "4e4731d59ab95679f205d424567dcf4d8e02f27c9318ed8ac955a8c163910628",
        "crates/litchi-drawingml/src/ink/actions_edit.rs": "b037eb7fae01b2e62c3c3a1050d9044466ad8d8e1a1402bb13533b2ba8bf9850",
        "crates/litchi-drawingml/src/ink/mod.rs": "a9a55cf0c44b59afa00a7b7c472c0f010eef5c23e62a816d4bb0f27aa6f67ff7",
        "crates/litchi-drawingml/tests/ink_action_edit.rs": "af92fe9d2923ac1197e4f152a5e34350217b5c782fa00365b82c141f69787640",
        "crates/litchi-drawingml/tests/ink_action_id_boundaries.rs": "f07b423989443d68ccba070aef5fed610bc39acb2c8b5a6adf14de8074afbf06",
    },
    "candidate": {
        "crates/litchi-drawingml/src/ink/actions.rs": "4d15cfd25456115bb750622096573118c7f1a652bc499be093db1fd93863a724",
        "crates/litchi-drawingml/src/ink/actions_edit.rs": "8262d2701e497b85c49024640bcaf4d8aaa1b3f856fa2d0956efa2a6f80534bb",
        "crates/litchi-drawingml/src/ink/mod.rs": "a9a55cf0c44b59afa00a7b7c472c0f010eef5c23e62a816d4bb0f27aa6f67ff7",
        "crates/litchi-drawingml/tests/ink_action_edit.rs": "dc565061ccdce79f8eac369ed092f37b608763bd35b91a8079089fb2b8bd2404",
        "crates/litchi-drawingml/tests/ink_action_id_boundaries.rs": "f07b423989443d68ccba070aef5fed610bc39acb2c8b5a6adf14de8074afbf06",
    },
}

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
CONTROL_LANES = tuple(
    lane for lane in LANES if lane.startswith("draft_") or lane.startswith("no_op_")
)
SEMANTIC_FIELDS = (
    "expected_success",
    "actual_success",
    "semantic_ok",
    "source_exact",
    "inverse_ok",
    "opaque_preserved",
    "output_exact",
    "rejection_ok",
    "rejection_source_unchanged",
    "rejection_state_unchanged",
    "rejection_resource",
    "rejection_limit",
    "rejection_pre_source_hash",
    "rejection_post_source_hash",
)
FIXTURE_FIELDS = (
    "action_count",
    "result_action_count",
    "operation_count",
    "input_bytes",
    "expected_success",
)
class CompareError(RuntimeError):
    """Raised when retained receipts cannot form a valid matched pair."""


def require(condition: bool, message: str) -> None:
    if not condition:
        raise CompareError(message)


def read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as error:
        raise CompareError(f"cannot read JSON receipt: {path}") from error


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    try:
        with path.open("rb") as stream:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                digest.update(chunk)
    except OSError as error:
        raise CompareError(f"cannot hash retained receipt: {path}") from error
    return digest.hexdigest()


def read_kv(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        raise CompareError(f"cannot read provenance receipt: {path}") from error
    for line in lines:
        if "=" not in line:
            continue
        key, value = line.split("=", 1)
        require(key, f"empty provenance key: {path}")
        if key in values:
            require(key in {"command", "run"}, f"duplicate provenance key {key}: {path}")
            continue
        values[key] = value
    return values


def positive_int(value: Any, field: str, path: Path) -> int:
    require(type(value) is int and value > 0, f"{field} must be positive: {path}")
    return int(value)


def read_binary_digest(path: Path) -> str:
    try:
        digest = path.read_text(encoding="utf-8").split()[0]
    except (OSError, IndexError) as error:
        raise CompareError(f"missing binary digest receipt: {path}") from error
    require(re.fullmatch(r"[0-9a-f]{64}", digest) is not None, f"malformed binary digest: {path}")
    return digest


def read_rss(path: Path) -> int:
    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        raise CompareError(f"cannot read RSS receipt: {path}") from error
    values = [
        line.split(":", 1)[1].strip()
        for line in lines
        if line.lstrip().startswith("Maximum resident set size (kbytes):")
    ]
    require(len(values) == 1, f"RSS receipt must contain one reading: {path}")
    require(values[0].isascii() and values[0].isdigit() and int(values[0]) > 0, f"invalid RSS: {path}")
    exit_statuses = [
        line.split(":", 1)[1].strip()
        for line in lines
        if line.lstrip().startswith("Exit status:") and ":" in line
    ]
    require(exit_statuses == ["0"], f"timed process must have exactly one successful exit status: {path}")
    return int(values[0])


def validate_source_manifest(path: Path, arm: str, expected_head: str) -> None:
    """Bind manifest identity and production source entries to the arm pins."""

    try:
        lines = path.read_text(encoding="utf-8").splitlines()
    except OSError as error:
        raise CompareError(f"cannot read source manifest: {path}") from error
    require(
        lines and lines[0] == "format=ink-action-edit-build-source-v2",
        f"source manifest format mismatch: {path}",
    )
    commits = [line.split("=", 1)[1] for line in lines if line.startswith("git_commit=")]
    require(commits == [expected_head], f"source manifest commit mismatch: {path}")
    entries: dict[str, list[str]] = {}
    for line in lines:
        if not line.startswith("file="):
            continue
        fields = line.split("\t")
        require(len(fields) == 5, f"malformed source manifest file line: {path}")
        shown = fields[3]
        digest = fields[4]
        require(re.fullmatch(r"[0-9a-f]{64}", digest) is not None, f"malformed source manifest hash: {path}")
        entries.setdefault(shown, []).append(digest)
    expected_entries = PINNED_SOURCE_ENTRIES[arm]
    for shown, expected_digest in expected_entries.items():
        require(
            entries.get(shown) == [expected_digest],
            f"pinned source manifest entry mismatch for {shown}: {path}",
        )


def expected_action_count(lane: str) -> int:
    if lane.endswith("_small_8"):
        return 8
    if lane.endswith("_scaled_128"):
        return 128
    if lane.endswith("_near_1024"):
        return 1024
    if lane.endswith("_opaque_64"):
        return 64
    raise CompareError(f"unknown lane action count: {lane}")


def expected_result_action_count(lane: str) -> int:
    count = expected_action_count(lane)
    if lane.startswith("add_"):
        return count + 1
    if lane.startswith("insert_batch_"):
        return count * 2
    if lane.startswith("remove_small_") or lane.startswith("remove_scaled_"):
        return count - 1
    if lane.startswith("remove_near_"):
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


def fixture_metadata(receipt: Mapping[str, Any], path: Path) -> dict[str, Any]:
    metadata = {}
    for field in FIXTURE_FIELDS:
        require(field in receipt, f"missing fixture field {field}: {path}")
        metadata[field] = receipt[field]
    require(metadata["action_count"] == expected_action_count(str(receipt["lane"])), f"action count mismatch: {path}")
    require(
        metadata["result_action_count"] == expected_result_action_count(str(receipt["lane"])),
        f"result action count mismatch: {path}",
    )
    require(type(metadata["operation_count"]) is int and metadata["operation_count"] > 0, f"invalid operation count: {path}")
    require(type(metadata["input_bytes"]) is int and metadata["input_bytes"] >= 0, f"invalid input bytes: {path}")
    require(type(metadata["expected_success"]) is bool, f"invalid expected success: {path}")
    return metadata


def semantic_signature(sample: Mapping[str, Any], path: Path) -> tuple[Any, ...]:
    for field in SEMANTIC_FIELDS:
        require(field in sample, f"missing semantic field {field}: {path}")
    return tuple(sample[field] for field in SEMANTIC_FIELDS)


def validate_fixture_pair(baseline: Mapping[str, Any], candidate: Mapping[str, Any], lane: str) -> None:
    """Reject a pair when its retained fixture or semantic outcome differs."""

    for field in FIXTURE_FIELDS:
        require(
            baseline[field] == candidate[field],
            f"fixture mismatch for {lane}: {field} {baseline[field]!r} != {candidate[field]!r}",
        )


def load_lane(result_dir: Path, lane: str) -> dict[str, Any]:
    all_samples: list[dict[str, Any]] = []
    fixture: dict[str, Any] | None = None
    statuses: list[tuple[Any, ...]] = []
    rss_values: list[int] = []
    pids: set[int] = set()
    for process in (1, 2, 3):
        path = result_dir / f"{lane}-p{process}.json"
        receipt = read_json(path)
        require(receipt.get("schema") == "ink-action-edit-profile-v2", f"schema mismatch: {path}")
        require(receipt.get("lane") == lane, f"lane mismatch: {path}")
        require(receipt.get("warmup") == 2, f"warm-up mismatch: {path}")
        require(receipt.get("sample_count") == 20, f"sample count mismatch: {path}")
        require(isinstance(receipt.get("samples"), list) and len(receipt["samples"]) == 20, f"sample array mismatch: {path}")
        expected_success = lane not in verify.CAP_REFUSALS
        require(
            type(receipt.get("expected_success")) is bool and receipt["expected_success"] is expected_success,
            f"expected status mismatch: {path}",
        )
        require(receipt.get("action_count") == expected_action_count(lane), f"action count mismatch: {path}")
        require(
            receipt.get("result_action_count") == expected_result_action_count(lane),
            f"result action count mismatch: {path}",
        )
        require(receipt.get("operation_count") == expected_operation_count(lane), f"operation count mismatch: {path}")
        input_bytes = receipt.get("input_bytes")
        require(type(input_bytes) is int and input_bytes >= 0, f"input byte count malformed: {path}")
        if lane.startswith("draft_"):
            require(input_bytes == 0, f"draft input bytes must be zero: {path}")
        else:
            require(input_bytes > 0, f"source-backed input bytes must be positive: {path}")
        pid = positive_int(receipt.get("pid"), "pid", path)
        require(pid not in pids, f"process identity repeated in {result_dir}: {pid}")
        pids.add(pid)
        current_fixture = fixture_metadata(receipt, path)
        if fixture is None:
            fixture = current_fixture
        else:
            validate_fixture_pair(fixture, current_fixture, lane)
        for sample in receipt["samples"]:
            require(isinstance(sample, dict), f"sample is not an object: {path}")
            try:
                verify.verify_sample(sample, expected_success, lane, input_bytes, path)
            except AssertionError as error:
                raise CompareError(str(error)) from error
            statuses.append(semantic_signature(sample, path))
        all_samples.extend(receipt["samples"])
        stderr = result_dir / f"{lane}-p{process}.stderr.log"
        require(stderr.is_file() and stderr.read_bytes() == b"", f"stderr is not empty: {stderr}")
        rss_values.append(read_rss(result_dir / f"{lane}-p{process}.time.txt"))
    require(fixture is not None, f"fixture missing: {lane}")
    return {
        "fixture": fixture,
        "samples": all_samples,
        "semantic_signatures": statuses,
        "rss_kib": rss_values,
        "pids": sorted(pids),
    }


def quantile(values: list[int], percent: int) -> int:
    require(values, "cannot compute a quantile from no samples")
    return sorted(values)[max(0, math.ceil(len(values) * percent / 100) - 1)]


def summarize_lane(loaded: Mapping[str, Any]) -> dict[str, Any]:
    samples = loaded["samples"]
    metrics: dict[str, dict[str, int]] = {}
    for source_field, output_field in (
        ("elapsed_ns", "latency_ns"),
        ("requested_alloc_bytes", "requested_alloc_bytes"),
        ("peak_live_delta", "peak_live_delta_bytes"),
    ):
        values = [int(sample[source_field]) for sample in samples]
        metrics[output_field] = {f"p{percent}": quantile(values, percent) for percent in (50, 95, 99)}
    rss_values = list(loaded["rss_kib"])
    metrics["rss_kib"] = {
        "median": int(statistics.median(rss_values)),
        "min": min(rss_values),
        "max": max(rss_values),
    }
    return metrics


def metric_delta(baseline: Mapping[str, int], candidate: Mapping[str, int]) -> dict[str, dict[str, float | int]]:
    result: dict[str, dict[str, float | int]] = {}
    for key in baseline:
        before = int(baseline[key])
        after = int(candidate[key])
        delta = after - before
        result[key] = {
            "baseline": before,
            "candidate": after,
            "delta": delta,
            "delta_percent": round(100.0 * delta / before, 4) if before else None,
        }
    return result


def read_identity(result_dir: Path, arm: str, expected_pin: str, expected_head: str) -> dict[str, Any]:
    require(result_dir.is_dir(), f"result directory missing: {result_dir}")
    verification = read_json(result_dir / "verification.json")
    require(verification.get("passed") is True, f"arm verifier did not pass: {result_dir}")
    require(verification.get("schema") == "ink-action-edit-profile-v2", f"verification schema mismatch: {result_dir}")
    require(verification.get("lanes") == 34, f"verification lane count mismatch: {result_dir}")
    require(verification.get("processes_per_lane") == 3, f"verification process count mismatch: {result_dir}")
    require(verification.get("samples_per_process") == 20, f"verification sample count mismatch: {result_dir}")
    require(verification.get("report_recomputed_from_samples") is True, f"report was not recomputed: {result_dir}")
    source = read_kv(result_dir / "source-provenance.txt")
    build = read_kv(result_dir / "build-provenance.txt")
    checkout = read_kv(result_dir / "checkout-provenance.txt")
    commands = read_kv(result_dir / "commands.txt")
    for values, name in ((source, "source"), (build, "build"), (checkout, "checkout")):
        require(values.get("profile_arm") == arm, f"{name} arm identity mismatch: {result_dir}")
        require(values.get("source_pin") == expected_pin, f"{name} source pin mismatch: {result_dir}")
    require(source.get("git_head") == expected_head, f"source head mismatch: {result_dir}")
    require(checkout.get("git_head") == expected_head, f"checkout head mismatch: {result_dir}")
    require(checkout.get("expected_control_commit") == expected_head, f"control head mismatch: {result_dir}")
    require(checkout.get("guard_commit") == GUARD_COMMIT, f"guard identity mismatch: {result_dir}")
    require(checkout.get("full_git_status") == "clean", f"checkout was not clean: {result_dir}")
    require(checkout.get("raw_lane_json_count") == "102", f"raw JSON count mismatch: {result_dir}")
    require(checkout.get("raw_timing_count") == "102", f"raw timing count mismatch: {result_dir}")
    require(checkout.get("raw_stderr_count") == "102", f"raw stderr count mismatch: {result_dir}")
    require(checkout.get("target_exists_after_capture") == "false", f"target cleanup mismatch: {result_dir}")
    require(build.get("cargo_incremental") == "0", f"allocator run flags changed: {result_dir}")
    require(build.get("flags") == "none", f"profile flags changed: {result_dir}")
    require(commands.get("warmup") == "2", f"command warm-up changed: {result_dir}")
    require(commands.get("samples_per_process") == "20", f"command sample count changed: {result_dir}")
    require(commands.get("fresh_processes_per_lane") == "3", f"command process count changed: {result_dir}")
    require(source.get("profile_pins_sha256") == PROFILE_PINS_SHA256, f"profile pin manifest changed: {result_dir}")
    require(source.get("source_manifest_before_sha256") == source.get("source_manifest_after_sha256"), f"source manifest changed: {result_dir}")
    before_manifest = result_dir / "source-manifest-before.txt"
    after_manifest = result_dir / "source-manifest-after.txt"
    require(sha256(before_manifest) == source.get("source_manifest_before_sha256"), f"before manifest hash mismatch: {result_dir}")
    require(sha256(after_manifest) == source.get("source_manifest_after_sha256"), f"after manifest hash mismatch: {result_dir}")
    require(sha256(before_manifest) == sha256(after_manifest), f"retained manifest bytes differ: {result_dir}")
    validate_source_manifest(before_manifest, arm, expected_head)
    validate_source_manifest(after_manifest, arm, expected_head)
    require(read_binary_digest(result_dir / "binary.sha256") == read_binary_digest(result_dir / "binary-after.sha256"), f"binary changed: {result_dir}")
    require(verification.get("expected_rejections") == [
        "cap_refusal_near_1024",
        "cap_refusal_scaled_128",
        "cap_refusal_small_8",
    ], f"rejection lane set changed: {result_dir}")
    return {
        "arm": arm,
        "source_pin": expected_pin,
        "git_head": expected_head,
        "guard_commit": GUARD_COMMIT,
        "profile_pins_sha256": source["profile_pins_sha256"],
        "binary_sha256": read_binary_digest(result_dir / "binary.sha256"),
        "source_manifest_sha256": source["source_manifest_before_sha256"],
        "result_dir": result_dir,
    }


def compare_arms(baseline_dir: Path, candidate_dir: Path) -> dict[str, Any]:
    baseline_identity = read_identity(baseline_dir, "baseline", BASELINE_PIN, BASELINE_HEAD)
    candidate_identity = read_identity(candidate_dir, "candidate", CANDIDATE_PIN, CANDIDATE_HEAD)
    require(
        baseline_identity["profile_pins_sha256"] == candidate_identity["profile_pins_sha256"],
        "baseline and candidate profile pin manifests differ",
    )
    rows: list[dict[str, Any]] = []
    for lane in LANES:
        baseline = load_lane(baseline_dir, lane)
        candidate = load_lane(candidate_dir, lane)
        validate_fixture_pair(baseline["fixture"], candidate["fixture"], lane)
        require(
            sorted(baseline["semantic_signatures"]) == sorted(candidate["semantic_signatures"]),
            f"semantic outcome mismatch for {lane}",
        )
        baseline_metrics = summarize_lane(baseline)
        candidate_metrics = summarize_lane(candidate)
        delta = {
            metric: metric_delta(baseline_metrics[metric], candidate_metrics[metric])
            for metric in baseline_metrics
        }
        allocation_exact = [
            sample["requested_alloc_bytes"] for sample in baseline["samples"]
        ] == [sample["requested_alloc_bytes"] for sample in candidate["samples"]]
        peak_exact = [
            sample["peak_live_delta"] for sample in baseline["samples"]
        ] == [sample["peak_live_delta"] for sample in candidate["samples"]]
        rows.append(
            {
                "lane": lane,
                "fixture": baseline["fixture"],
                "baseline": baseline_metrics,
                "candidate": candidate_metrics,
                "delta": delta,
                "allocator_vectors_unchanged": allocation_exact,
                "peak_vectors_unchanged": peak_exact,
                "control_lane": lane in CONTROL_LANES,
            }
        )
    def result_label(path: Path) -> str:
        try:
            return path.relative_to(HERE / "results").as_posix()
        except ValueError:
            return str(path)

    return {
        "schema": "ink-action-edit-matched-compare-v1",
        "comparison": "candidate_minus_baseline",
        "scope": "34 retained InkAction lanes; no new process or timed run",
        "baseline": {
            **{key: value for key, value in baseline_identity.items() if key != "result_dir"},
            "result_dir": result_label(baseline_dir),
        },
        "candidate": {
            **{key: value for key, value in candidate_identity.items() if key != "result_dir"},
            "result_dir": result_label(candidate_dir),
        },
        "common_contract": {
            "toolchain": "Rust/Cargo 1.95.0",
            "cargo_flags": "--release --locked --offline",
            "allocator": "process-local counting allocator",
            "fresh_processes_per_lane": 3,
            "warmups_per_process": 2,
            "samples_per_process": 20,
            "harness_lockfile": "byte-identical per paired-controls-provenance.txt",
        },
        "control_lanes": list(CONTROL_LANES),
        "control_allocator_and_peak_vectors_unchanged": all(
            row["allocator_vectors_unchanged"] and row["peak_vectors_unchanged"]
            for row in rows
            if row["control_lane"]
        ),
        "lanes": rows,
        "interpretation": {
            "expected_local_change": "successful source-backed edits reuse the emitted Arc source during typed readback; the baseline readback copies that emitted source once more",
            "control_observation": "draft and exact no-op lanes do not exercise the changed typed readback path; their allocator and peak-live vectors are unchanged",
            "claim_scope": "per-lane retained-receipt deltas only; no speedup, asymptotic, native-application, or package-wide claim",
        },
    }


def format_delta(metric: Mapping[str, float | int], unit: str = "") -> str:
    baseline = int(metric["baseline"])
    candidate = int(metric["candidate"])
    delta = int(metric["delta"])
    percent = metric["delta_percent"]
    sign = "+" if delta > 0 else ""
    percent_text = "n/a" if percent is None else f"{float(percent):+.2f}%"
    return f"{baseline:,}→{candidate:,} {sign}{delta:,}{unit} ({percent_text})"


def write_report(data: Mapping[str, Any], path: Path) -> None:
    baseline = data["baseline"]
    candidate = data["candidate"]
    lines = [
        "# Matched InkAction profile comparison",
        "",
        "This report compares the retained raw receipts from the clean baseline and candidate arms. It starts no profile process and reports candidate-minus-baseline deltas within each bounded lane.",
        "",
        "## Pair identity",
        "",
        f"- Baseline source pin `{baseline['source_pin']}` at control head `{baseline['git_head']}`: [`{baseline['result_dir']}/report.md`]({baseline['result_dir']}/report.md).",
        f"- Candidate source pin `{candidate['source_pin']}` at control head `{candidate['git_head']}`: [`{candidate['result_dir']}/report.md`]({candidate['result_dir']}/report.md).",
        f"- Both arms use guard `{baseline['guard_commit']}`, profile pin manifest `{baseline['profile_pins_sha256']}`, and the common Rust/Cargo 1.95.0 harness contract.",
        "- The reviewed candidate source change is the three-file `c8d60d5632d79351c325e3e6f1eace2d901b312c` child of the baseline source pin.",
        "",
        "## Reading the table",
        "",
        "Every value is `baseline→candidate delta (relative delta)`. Latency uses all 60 measured samples per lane; requested allocation and peak-live use the same samples; RSS uses the median of the three fresh-process `/usr/bin/time -v` readings. A positive delta means the candidate receipt is larger for that metric.",
        "",
        "Draft and exact no-op lanes are controls for the changed typed readback path. Their requested-allocation and peak-live sample vectors are unchanged between arms; latency and RSS remain listed as observed process measurements.",
        "",
        "The expected local effect in successful source-backed edit lanes is one avoided source copy during typed readback: the candidate reuses the already-emitted `Arc<[u8]>`, while the baseline readback copies those emitted bytes into a second source allocation. The copy size varies with the emitted output. This comparison does not infer a speedup, asymptotic behavior, native-application behavior, or package-wide performance.",
        "",
        "The repeated-scalar-write lanes report their final value and measured cost only; the public API exposes no internal coalescing diagnostic, so this comparison makes no coalescing claim.",
        "",
        "| lane | fixture | latency p50 / p95 / p99 ns | requested alloc p50 / p95 B | peak-live p50 / p95 B | RSS median KiB | note |",
        "|---|---|---:|---:|---:|---:|---|",
    ]
    for row in data["lanes"]:
        fixture = row["fixture"]
        latency = row["delta"]["latency_ns"]
        alloc = row["delta"]["requested_alloc_bytes"]
        peak = row["delta"]["peak_live_delta_bytes"]
        rss = row["delta"]["rss_kib"]["median"]
        if row["control_lane"]:
            note = "control; alloc/peak vectors unchanged"
        elif row["lane"].startswith("cap_refusal_"):
            note = "caller-cap refusal"
        else:
            note = "source-backed edit"
        fixture_text = f"{fixture['action_count']}/{fixture['result_action_count']}/{fixture['operation_count']}"
        lines.append(
            "| {lane} | {fixture} | {latency} | {alloc} | {peak} | {rss} | {note} |".format(
                lane=row["lane"],
                fixture=fixture_text,
                latency=" / ".join(format_delta(latency[f"p{p}"], " ns") for p in (50, 95, 99)),
                alloc=" / ".join(format_delta(alloc[f"p{p}"], " B") for p in (50, 95)),
                peak=" / ".join(format_delta(peak[f"p{p}"], " B") for p in (50, 95)),
                rss=format_delta(rss, " KiB"),
                note=note,
            )
        )
    lines.extend(
        [
            "",
            "Fixture column is `input actions/result actions/queued operations`; input byte sizes remain in the machine-readable comparison. Raw JSON, timing, RSS, allocator, source, and binary receipts remain in the two linked result directories.",
            "",
            "## Control check",
            "",
            "The exact control lanes are `" + "`, `".join(data["control_lanes"]) + "`. Their allocator and peak-live vectors are byte-for-byte unchanged: `" + str(data["control_allocator_and_peak_vectors_unchanged"]).lower() + "`.",
            "",
            "The comparison is intentionally descriptive. It does not combine lanes, extrapolate beyond the 8/128/1,024-action fixtures, or turn these receipts into a general performance claim.",
            "",
        ]
    )
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("\n".join(lines), encoding="utf-8")


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--baseline", type=Path, default=DEFAULT_BASELINE)
    parser.add_argument("--candidate", type=Path, default=DEFAULT_CANDIDATE)
    parser.add_argument("--output", type=Path, default=DEFAULT_REPORT)
    parser.add_argument("--json-output", type=Path, default=DEFAULT_JSON)
    args = parser.parse_args()
    try:
        data = compare_arms(args.baseline.resolve(), args.candidate.resolve())
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(json.dumps(data, indent=2, sort_keys=True) + "\n", encoding="utf-8")
        write_report(data, args.output)
    except CompareError as error:
        parser.error(str(error))
    print(f"wrote {args.output}")
    print(f"wrote {args.json_output}")


if __name__ == "__main__":
    main()
