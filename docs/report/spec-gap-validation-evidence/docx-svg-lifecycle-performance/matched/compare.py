#!/usr/bin/env python3
"""Recompute a matched control/candidate DOCX SVG lifecycle comparison."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
from pathlib import Path


LANES = (
    "native_svg_capture",
    "native_floating_capture",
    "lazy_inventory_1",
    "lazy_inventory_64",
    "single_attach_1",
    "single_attach_16",
    "single_attach_64",
    "single_detach_1",
    "single_detach_16",
    "single_detach_64",
    "batch_attach_1",
    "batch_attach_16",
    "batch_attach_64",
    "batch_detach_1",
    "batch_detach_16",
    "batch_detach_64",
    "shared_svg_cleanup",
    "exact_inverse_single_1",
    "exact_inverse_batch_64",
    "large_unchanged_media_managed_cap",
    "noop_detach_64",
)
REFUSALS = {"batch_attach_64", "exact_inverse_batch_64"}
PHASES = (
    "capture_ns",
    "stage_ns",
    "commit_ns",
    "publish_ns",
    "reopen_ns",
    "inverse_reopen_ns",
    "inverse_ns",
    "payload_ns",
    "validation_ns",
    "readback_ns",
)
EXPECTED_CONTROL = "211806feea6e83279975ff77adeb5c8fa7be01a6"
EXPECTED_CANDIDATE = "ceeecf972e716be850547917fe348c1437186906"


def quantile(values: list[int], percent: int) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * percent / 100) - 1)]


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def rss(path: Path) -> int:
    marker = "Maximum resident set size (kbytes):"
    values = [
        int(line.split(":", 1)[1].strip())
        for line in path.read_text().splitlines()
        if line.lstrip().startswith(marker)
    ]
    if len(values) != 1 or values[0] <= 0:
        raise ValueError(f"invalid RSS receipt: {path}")
    return values[0]


def receipt_samples(root: Path, lane: str) -> tuple[list[dict[str, object]], list[int]]:
    samples: list[dict[str, object]] = []
    rss_values: list[int] = []
    for process in range(1, 4):
        path = root / f"full-{lane}-p{process}.json"
        value = json.loads(path.read_text())
        if value["lane"] != lane or value["sample_count"] != 20 or len(value["samples"]) != 20:
            raise ValueError(f"receipt contract mismatch: {path}")
        if value["expected_success"] is (lane not in REFUSALS):
            pass
        else:
            raise ValueError(f"success expectation mismatch: {path}")
        samples.extend(value["samples"])
        rss_values.append(rss(path.with_suffix(".time.txt")))
        stderr = path.with_suffix(".stderr.log")
        if stderr.read_bytes() != b"":
            raise ValueError(f"unexpected stderr: {stderr}")
    if len(samples) != 60:
        raise ValueError(f"sample count mismatch: {lane}")
    return samples, rss_values


def aggregate(root: Path, lane: str) -> dict[str, object]:
    samples, rss_values = receipt_samples(root, lane)
    elapsed = [int(sample["elapsed_ns"]) for sample in samples]
    allocated = [int(sample["requested_alloc_bytes"]) for sample in samples]
    peak = [int(sample["peak_live_delta"]) for sample in samples]
    phases = {
        phase: quantile([int(sample["phases"][phase]) for sample in samples], 50)
        for phase in PHASES
    }
    dominant = max(phases, key=phases.get)
    errors = [sample["error"] for sample in samples if sample["error"] is not None]
    if lane in REFUSALS:
        if len(errors) != 60 or any(
            error["typed_match"] is not True
            or error["class"] != "topology_part_bound"
            for error in errors
        ):
            raise ValueError(f"typed refusal contract mismatch: {lane}")
    elif errors:
        raise ValueError(f"unexpected error receipt: {lane}")
    return {
        "elapsed": tuple(quantile(elapsed, percent) for percent in (50, 95, 99)),
        "alloc": tuple(quantile(allocated, percent) for percent in (50, 95)),
        "peak": tuple(quantile(peak, percent) for percent in (50, 95)),
        "rss": (min(rss_values), max(rss_values)),
        "phases": phases,
        "dominant": dominant,
        "sample_count": len(samples),
    }


def pct(candidate: int, control: int) -> str:
    if control == 0:
        return "n/a"
    return f"{(candidate / control - 1.0) * 100.0:+.1f}%"


def manifest_field(path: Path, key: str) -> str:
    for line in path.read_text().splitlines():
        if line.startswith(f"{key}="):
            return line.split("=", 1)[1]
    raise ValueError(f"missing {key} in {path}")


def status_field(path: Path, key: str) -> str:
    for line in path.read_text().splitlines():
        if line.startswith(f"{key}="):
            return line.split("=", 1)[1]
    raise ValueError(f"missing {key} in {path}")


def verify_matched_inputs(root: Path, label: str) -> None:
    manifest = root / "matched-source-manifest.txt"
    source_root = root.parent / "source" / label
    counts = {"source": 0, "harness": 0, "fixture": 0}
    for line in manifest.read_text().splitlines():
        prefix = next((prefix for prefix in counts if line.startswith(f"{prefix}=")), None)
        if prefix is None:
            continue
        shown, expected = line[len(prefix) + 1 :].split("\t", 1)
        path = source_root / "harness" / shown if prefix == "harness" else source_root / shown
        if not path.is_file() or sha256(path) != expected:
            raise SystemExit(f"matched source hash mismatch: {label}:{prefix}:{shown}")
        counts[prefix] += 1
    if counts != {"source": 9, "harness": 5, "fixture": 2}:
        raise SystemExit(f"matched source manifest counts changed: {label}: {counts}")
    if (root / "full-metadata-before.json").read_bytes() != (root / "full-metadata-after.json").read_bytes():
        raise SystemExit(f"Cargo metadata changed: {label}")
    if (root / "full-binary.sha256").read_bytes() != (root / "full-binary-after.sha256").read_bytes():
        raise SystemExit(f"profile binary changed: {label}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--matched", type=Path, default=Path(__file__).resolve().parent)
    parser.add_argument("--output", type=Path)
    args = parser.parse_args()
    matched = args.matched.resolve()
    control = matched / "control"
    candidate = matched / "candidate"
    if manifest_field(control / "matched-source-manifest.txt", "production_commit") != EXPECTED_CONTROL:
        raise SystemExit("control production commit drift")
    if manifest_field(candidate / "matched-source-manifest.txt", "production_commit") != EXPECTED_CANDIDATE:
        raise SystemExit("candidate production commit drift")
    if manifest_field(control / "matched-source-manifest.txt", "harness_source_commit") != manifest_field(candidate / "matched-source-manifest.txt", "harness_source_commit"):
        raise SystemExit("harness source commit differs")
    if sha256(control / "matched-source-manifest.txt") == sha256(candidate / "matched-source-manifest.txt"):
        raise SystemExit("control/candidate manifests unexpectedly identical")
    for label, root, expected_commit in (
        ("control", control, EXPECTED_CONTROL),
        ("candidate", candidate, EXPECTED_CANDIDATE),
    ):
        verify_matched_inputs(root, label)
        status = root / "matched-checkout-status.txt"
        if status_field(status, "checkout_head") != expected_commit:
            raise SystemExit(f"checkout HEAD drift: {label}")
        if status_field(status, "checkout_status") != "clean":
            raise SystemExit(f"checkout was not clean: {label}")
        if status_field(status, "shared_dirty_workspace_used_for_build") != "false":
            raise SystemExit(f"shared dirty workspace used: {label}")
        if manifest_field(root / "matched-source-manifest.txt", "checkout_status_receipt_sha256") != sha256(status):
            raise SystemExit(f"checkout status receipt hash mismatch: {label}")
    for label, root in (("control", control), ("candidate", candidate)):
        verification = json.loads((root / "full-verification.json").read_text())
        if verification != {
            "passed": True,
            "mode": "full",
            "lanes": 21,
            "processes": 63,
            "samples": 1260,
            "manifest_inputs_checked": 26,
            "expected_refusals": ["batch_attach_64", "exact_inverse_batch_64"],
            "allocation_peak_rss_separate": True,
            "report_recomputed_from_raw_receipts": True,
            "fail_closed_gates": [
                "source_snapshot",
                "source_manifest",
                "fixture_hash",
                "semantic",
                "opaque",
                "lazy_media",
                "exact_inverse",
                "typed_error",
                "failure_physical_readback",
                "failure_metadata_readback",
                "allocator",
                "rss",
                "process_exit",
            ],
        }:
            raise SystemExit(f"frozen verifier receipt changed: {label}")

    rows: list[dict[str, object]] = []
    for lane in LANES:
        control_row = aggregate(control, lane)
        candidate_row = aggregate(candidate, lane)
        rows.append({"lane": lane, "control": control_row, "candidate": candidate_row})

    output = args.output or matched / "matched-comparison.md"
    lines = [
        "# Matched DOCX source-backed SVG lifecycle profile",
        "",
        "This comparison uses two isolated committed worktrees and the same frozen public-API harness: control `211806fee` and candidate `ceeecf972`. Each lane used three fresh processes, two warmups, and twenty measured samples per process (60 samples per row). The frozen verifier passed independently for both profiles.",
        "",
        "The percentage columns are matched scenario observations computed from p50 values as `(candidate / control - 1) * 100`; they are not a general speedup claim. Allocation traffic, peak live allocation, and process RSS are separate metrics. The temporary Cargo targets and worktrees were removed after verification; per-commit source and harness snapshots remain under `source/`.",
        "",
        f"- control production commit: `{EXPECTED_CONTROL}`",
        f"- candidate production commit: `{EXPECTED_CANDIDATE}`",
        f"- frozen harness anchor: `{manifest_field(control / 'matched-source-manifest.txt', 'harness_source_commit')}`",
        f"- comparison driver SHA-256: `{sha256(Path(__file__))}`",
        f"- control frozen-verifier receipt SHA-256: `{sha256(control / 'full-verification.json')}`",
        f"- candidate frozen-verifier receipt SHA-256: `{sha256(candidate / 'full-verification.json')}`",
        "- allocator: `CountingAllocator` process-local `GlobalAlloc` observer",
        "- RSS: `/usr/bin/time -v` maximum resident set size sidecars",
        "- hard scratch-budget/OOM claim: none",
        "- retained per-profile reports: [`control/full-report.md`](control/full-report.md), [`candidate/full-report.md`](candidate/full-report.md)",
        "",
        "| lane | elapsed p50 ns control/candidate (delta) | requested alloc p50 B control/candidate (delta) | peak live p50 B control/candidate (delta) | RSS KiB control/candidate ranges | dominant p50 phase control → candidate |",
        "|---|---:|---:|---:|---:|---|",
    ]
    for row in rows:
        lane = row["lane"]
        before = row["control"]
        after = row["candidate"]
        lines.append(
            f"| {lane} | {before['elapsed'][0]}/{after['elapsed'][0]} ({pct(after['elapsed'][0], before['elapsed'][0])}) | "
            f"{before['alloc'][0]}/{after['alloc'][0]} ({pct(after['alloc'][0], before['alloc'][0])}) | "
            f"{before['peak'][0]}/{after['peak'][0]} ({pct(after['peak'][0], before['peak'][0])}) | "
            f"{before['rss'][0]}–{before['rss'][1]} / {after['rss'][0]}–{after['rss'][1]} | "
            f"{before['dominant']} ({before['phases'][before['dominant']]}) → {after['dominant']} ({after['phases'][after['dominant']]}) |"
        )
    lines.extend(
        [
            "",
            "## Cost and scaling observations",
            "",
            "The control profile's 16- and 64-owner attach/detach and shared-cleanup rows are dominated by `stage_ns`, and the stage p50 grows sharply with owner count. The candidate removes most of that repeated source-layout/stage work: candidate 16-owner batch attach/detach rows are dominated by publication, and the 64-owner successful detach/shared-cleanup rows show the same shift. Single-owner and native capture rows remain dominated by publication or capture/validation, so their deltas are smaller and should not be generalized beyond this fixture matrix.",
            "",
            "The raw harness receipts intentionally retain `baseline_commit=892441d95db29da4351390716ef5c65b4c7c97de` and `opc_baseline_label=OPC57680dc86`; the outer matched manifests bind those receipts to the control and candidate production commits above.",
            "",
            "Both 64-owner refusal lanes still return the typed `DocxError::Opc(OpcError::SourceBackedOverlayUnavailable { .. })` match in every sample. The receipt class `topology_part_bound` is only a diagnostic label; `typed_match=true` is the gate. Failed operations republish byte-identical physical source bytes and matching metadata, and emit zero output bytes.",
            "",
            "The candidate still builds a full projected story buffer for each staged operation. The profile therefore does not establish a bounded scratch-memory guarantee: the residual projected-story copy remains a material allocation component as same-story owner counts grow, even where repeated source scanning/layout projection is avoided. The `large_unchanged_media_managed_cap` lane keeps its 8-MiB unchanged member under the 2-MiB managed execution cap and passes exact opaque-preservation checks; RSS remains an independent process-level observation.",
            "",
            "Raw receipts, `/usr/bin/time -v` sidecars, per-commit source/harness/fixture snapshots, binary/build metadata, and the two frozen-verifier outputs are retained beside this report. Recompute with `python3 matched/compare.py`.",
        ]
    )
    output.write_text("\n".join(lines) + "\n")
    print(output)


if __name__ == "__main__":
    main()
