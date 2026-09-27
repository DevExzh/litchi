"""Offline replay of the 0791 frame-pointer native stack diagnostic."""

from __future__ import annotations

import argparse
import collections
import gzip
import hashlib
from pathlib import Path
from typing import Any

import analysis_common as c
import native_analysis as native


DIAGNOSTIC_SYMBOLS = (
    c.SCAN,
    c.FINGERPRINT,
    c.RESOLVED,
    c.SHA,
    c.LOAD_INDEX,
    "litchi_pptx::notes::codec::inspect_element",
    "quick_xml::events::attributes::IterState::next",
    "quick_xml::events::attributes::IterState::check_for_duplicates",
    "<litchi_opc::xml_attributes::CheckedAttributes as core::iter::traits::iterator::Iterator>::next",
)


def _gzip_bytes(path: Path, descriptor: dict[str, Any], label: str) -> bytes:
    c.artifact(descriptor, label)
    try:
        with gzip.open(path, "rb") as stream:
            data = stream.read()
    except (OSError, EOFError) as error:
        raise c.EvidenceError(f"{label}: invalid gzip: {error}") from error
    original = descriptor.get("original")
    c.require(isinstance(original, dict), f"{label}: original descriptor is missing")
    c.require(len(data) == original.get("bytes")
              and hashlib.sha256(data).hexdigest() == original.get("sha256"),
              f"{label}: decompressed identity differs")
    return data


def _build(plan: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    _, build = native._build()
    fp_plan = c.read_json(c.PACKET / "perf-fp-plan.json")
    c.require(fp_plan.get("binary") == "profile-fp"
              and fp_plan.get("build_rustflags") == "-C force-frame-pointers=yes"
              and fp_plan.get("call_graph") == "fp"
              and fp_plan.get("event") == "cycles:u"
              and fp_plan.get("cpu") == 12
              and fp_plan.get("frequency") == 499
              and fp_plan.get("repeats") == 2
              and fp_plan.get("samples") == 100
              and fp_plan.get("warmup") == 3,
              "frame-pointer plan changed")
    c.require(fp_plan.get("owner") == c.CHILD,
              "frame-pointer owner changed")
    receipt = c.read_json(c.PACKET / "build-fp" / "receipt.json")
    c.require(receipt.get("exit_code") == 0, "frame-pointer build failed")
    c.require(receipt.get("environment", {}).get("RUSTFLAGS")
              == "-C force-frame-pointers=yes", "frame-pointer RUSTFLAGS changed")
    c.artifact(receipt.get("source"), "frame-pointer source")
    c.artifact(receipt.get("plan"), "frame-pointer plan")
    c.artifact(receipt.get("log"), "frame-pointer build log")
    c.require(receipt.get("source") == build.get("source"),
              "frame-pointer source receipt differs from ordinary build")
    cleanup = c.cleanup_witness()
    fp_binary = receipt.get("binary")
    c.external_artifact(fp_binary, "build profile-fp", cleanup)
    return fp_plan, receipt


def _compressed_index() -> dict[str, dict[str, Any]]:
    values = c.read_json(c.PACKET / "perf-fp" / "compression.json")
    c.require(isinstance(values, list) and len(values) == 4,
              "compressed perf artifact cardinality changed")
    result: dict[str, dict[str, Any]] = {}
    expected = {"0.data", "0.decoded", "1.data", "1.decoded"}
    for row in values:
        c.require(isinstance(row, dict), "compression row is malformed")
        original = row.get("original")
        compressed = row.get("compressed")
        c.require(isinstance(original, dict) and isinstance(compressed, dict),
                  "compression identity is incomplete")
        name = Path(str(original.get("path"))).name
        c.require(name in expected and name not in result,
                  "compression original matrix changed")
        c.artifact(original, f"compressed original {name}", allow_missing=True)
        compressed_path = c.artifact(compressed, f"compressed artifact {name}")
        with gzip.open(compressed_path, "rb") as stream:
            data = stream.read()
        c.require(len(data) == original.get("bytes")
                  and hashlib.sha256(data).hexdigest() == original.get("sha256"),
                  f"compressed artifact {name}: decompressed identity changed")
        result[name] = {"original": original, "compressed": compressed,
                        "data": data}
    c.require(set(result) == expected, "compressed perf artifact names changed")
    return result


def _report_identity(path: Path, fp_plan: dict[str, Any], receipt: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    report = c.read_json(path)
    identity = c.check_report_identity(report, "large", samples=fp_plan["samples"],
                                       warmup=fp_plan["warmup"], binary="profile-fp")
    return report, identity


def _stack_counts(stream: bytes, parser: Any, label: str) -> dict[str, Any]:
    try:
        all_samples = parser.samples(stream)
    except (UnicodeError, ValueError, TypeError, IndexError) as error:
        raise c.EvidenceError(f"{label}: perf script parse failed: {error}") from error
    c.require(all_samples, f"{label}: decoded perf stream is empty")
    c.require(len({sample["timestamp"] for sample in all_samples}) == len(all_samples),
              f"{label}: perf timestamps are not unique")
    c.require(all(sample["period"] > 0 for sample in all_samples),
              f"{label}: non-positive sample period")
    qualified = [sample for sample in all_samples if c.CHILD in sample["symbols"]]
    partitions = {"notes_scan": 0, "fingerprint": 0, "other_capture": 0}
    partition_period = dict.fromkeys(partitions, 0)
    nested_counts = collections.Counter()
    nested_period = collections.Counter()
    leaf_counts = collections.Counter()
    leaf_period = collections.Counter()
    unknown_interior = 0
    for sample in qualified:
        symbols = sample["symbols"]
        c.require(symbols.count(c.CHILD) == 1,
                  f"{label}: capture owner appears more than once")
        owner_index = symbols.index(c.CHILD)
        interior = symbols[:owner_index]
        if c.SCAN in interior:
            category = "notes_scan"
        elif c.FINGERPRINT in interior:
            category = "fingerprint"
        else:
            category = "other_capture"
        partitions[category] += 1
        partition_period[category] += sample["period"]
        for symbol in DIAGNOSTIC_SYMBOLS:
            if symbol in interior:
                nested_counts[symbol] += 1
                nested_period[symbol] += sample["period"]
        if interior:
            leaf_counts[interior[0]] += 1
            leaf_period[interior[0]] += sample["period"]
        if any(token in value.lower() for value in interior
               for token in ("[unknown]", "??", "<unknown>")):
            unknown_interior += 1
    c.require(sum(partitions.values()) == len(qualified),
              f"{label}: flat owner partition does not reconstruct count")
    return {
        "whole_process_samples": len(all_samples),
        "whole_process_period": sum(sample["period"] for sample in all_samples),
        "owner_qualified_samples": len(qualified),
        "owner_qualified_period": sum(sample["period"] for sample in qualified),
        "unqualified_samples": len(all_samples) - len(qualified),
        "exclusive_partition_samples": partitions,
        "exclusive_partition_period": dict(partition_period),
        "nested_inclusive_samples": dict(sorted(nested_counts.items())),
        "nested_inclusive_period": dict(sorted(nested_period.items())),
        "top_sampled_leaf_symbols": [[name, count]
                                      for name, count in leaf_counts.most_common(20)],
        "top_sampled_leaf_period": [[name, value]
                                    for name, value in leaf_period.most_common(20)],
        "diagnostic_symbols": {
            name: {"samples": nested_counts.get(name, 0),
                   "period": nested_period.get(name, 0)}
            for name in DIAGNOSTIC_SYMBOLS
        },
        "qualified_stacks_with_unknown_interior": unknown_interior,
    }


def analyze() -> dict[str, Any]:
    plan = c.read_json(c.PACKET / "plan.json")
    c.require(plan.get("schema") == "litchi.performance.0791.v1", "plan schema changed")
    fp_plan, fp_receipt = _build(plan)
    complete = c.read_json(c.PACKET / "perf-fp" / "complete.json")
    c.require(complete == {
        "processes": 2,
        "plan_sha256": c.sha256(c.PACKET / "perf-fp-plan.json"),
    }, "perf-fp completion receipt changed")
    rows = c.read_json(c.PACKET / "perf-fp" / "receipts.json")
    decodes = c.read_json(c.PACKET / "perf-fp" / "decode-receipts.json")
    frames = c.read_json(c.PACKET / "perf-fp" / "frame-receipts.json")
    c.require(isinstance(rows, list) and len(rows) == 2
              and isinstance(decodes, list) and len(decodes) == 2
              and isinstance(frames, list) and len(frames) == 2,
              "perf-fp receipt cardinality changed")
    packed = _compressed_index()
    parser = c.load_perf_parser()
    captures: list[dict[str, Any]] = []
    for index, (row, decoded, frame) in enumerate(zip(rows, decodes, frames)):
        label = f"perf-fp/{index}"
        c.require(row.get("repeat") == index and row.get("exit_code") == 0,
                  f"{label}: capture receipt changed")
        c.require(row.get("plan_sha256") == c.sha256(c.PACKET / "perf-fp-plan.json"),
                  f"{label}: plan identity changed")
        c.require(row.get("binary") == fp_receipt.get("binary"),
                  f"{label}: binary identity changed")
        raw = row.get("raw")
        report = row.get("report")
        log = row.get("log")
        c.require(isinstance(raw, dict) and isinstance(report, dict) and isinstance(log, dict),
                  f"{label}: capture artifacts are incomplete")
        raw_name = Path(str(raw.get("path"))).name
        c.require(raw_name == f"{index}.data" and raw == packed[raw_name]["original"],
                  f"{label}: raw identity changed")
        c.artifact(raw, f"{label} raw", allow_missing=True)
        report_path = c.artifact(report, f"{label} report")
        c.artifact(log, f"{label} log")
        expected_command = [
            "taskset", "-c", str(fp_plan["cpu"]), "perf", "record",
            "-e", fp_plan["event"], "-F", str(fp_plan["frequency"]),
            "--call-graph", fp_plan["call_graph"], "-o", raw["path"], "--",
            fp_receipt["binary"]["path"], "--mode", "capture", "--shape", "large",
            "--samples", str(fp_plan["samples"]), "--warmup", str(fp_plan["warmup"]),
            "--output", report["path"],
        ]
        c.require(c.normalize_command(row.get("command"))
                  == c.normalize_command(expected_command),
                  f"{label}: perf command changed")
        _, identity = _report_identity(report_path, fp_plan, fp_receipt)
        c.require(decoded.get("repeat") == index and decoded.get("exit_code") == 0,
                  f"{label}: decode receipt changed")
        c.require(decoded.get("raw") == raw
                  and decoded.get("command") == ["perf", "script", "--ns", "-i", raw["path"]],
                  f"{label}: decode input changed")
        decoded_descriptor = decoded.get("decoded")
        c.require(decoded_descriptor == packed[f"{index}.decoded"]["original"],
                  f"{label}: decoded identity changed")
        c.artifact(decoded_descriptor, f"{label} decoded", allow_missing=True)
        c.artifact(decoded.get("log"), f"{label} decode log")
        c.require(frame.get("repeat") == index and frame.get("exit_code") == 0,
                  f"{label}: frame receipt changed")
        c.require(frame.get("raw") == raw
                  and frame.get("command") == ["perf", "script", "--no-inline", "--ns",
                                                "-i", raw["path"]],
                  f"{label}: frame decode input changed")
        frame_descriptor = frame.get("frames")
        compressed_frames = frame.get("compressed")
        c.require(isinstance(frame_descriptor, dict) and isinstance(compressed_frames, dict),
                  f"{label}: no-inline frame identities are incomplete")
        c.artifact(frame_descriptor, f"{label} frames", allow_missing=True)
        frame_path = c.artifact(compressed_frames, f"{label} compressed frames")
        with gzip.open(frame_path, "rb") as stream:
            frame_data = stream.read()
        c.require(len(frame_data) == frame_descriptor.get("bytes")
                  and hashlib.sha256(frame_data).hexdigest() == frame_descriptor.get("sha256"),
                  f"{label}: compressed no-inline frames changed")
        c.artifact(frame.get("log"), f"{label} frame log")
        stack = _stack_counts(frame_data, parser, label)
        captures.append({
            "repeat": index,
            "command": row["command"],
            "report": str(report_path.relative_to(c.PACKET)),
            "report_sha256": c.sha256(report_path),
            "report_identity": {key: identity[key] for key in ("source", "output", "verification")},
            "raw": raw, "decoded": decoded_descriptor,
            "frames": frame_descriptor, "compressed_frames": compressed_frames,
            **stack,
        })
    unknown = sum(item["qualified_stacks_with_unknown_interior"] for item in captures)
    qualified = sum(item["owner_qualified_samples"] for item in captures)
    return {
        "schema": "litchi-0791-native-perf-analysis-v1",
        "packet": "change-0791",
        "plan": {"path": "perf-fp-plan.json", "sha256": c.sha256(c.PACKET / "perf-fp-plan.json"),
                 "owner": fp_plan["owner"], "cpu": fp_plan["cpu"],
                 "event": fp_plan["event"], "frequency": fp_plan["frequency"],
                 "call_graph": fp_plan["call_graph"]},
        "source": c.source_identity(),
        "binary": c.external_artifact(fp_receipt["binary"], "build profile-fp", c.cleanup_witness()),
        "captures": captures,
        "summary": {
            "repeats": len(captures), "whole_process_samples": sum(item["whole_process_samples"] for item in captures),
            "owner_qualified_samples": qualified,
            "owner_qualified_period": sum(item["owner_qualified_period"] for item in captures),
            "qualified_stacks_with_unknown_interior": unknown,
            "exact_owner_counts_available": qualified > 0,
        },
        "phase_fraction_claim_authorized": False,
        "fraction_refusal_reason": "The frozen diagnostic forbids native phase fractions when exact owner qualification is empty or any qualified interior contains unresolved frames; this report retains observed owner, flat-category, nested-symbol, and period counts only.",
        "scope": "Frame-pointer sampled native diagnostics over the large capture probe; no phase fraction, native speedup, or production attribution claim.",
        "claims": [
            "Only samples containing the exact capture owner are included in owner and nested counts.",
            "Flat notes-scan, fingerprint, and other-capture categories are disjoint; nested symbol counts are inclusive diagnostics and are not additive.",
            "Compressed raw perf data, decoded output, and canonical no-inline frame output are identity-checked before parsing.",
        ],
    }


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--check", action="store_true")
    parser.add_argument("--write", action="store_true")
    args = parser.parse_args()
    c.require(args.check ^ args.write, "choose exactly one of --write or --check")
    c.write_or_check(c.PACKET / "perf-analysis.json", analyze(), args.check)
    print("0791 frame-pointer perf analysis PASS", flush=True)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
