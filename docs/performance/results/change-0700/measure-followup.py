#!/usr/bin/env python3
"""Run the isolated 0700 native/refusal follow-up without touching initial data."""

from __future__ import annotations

import hashlib
import json
import math
import statistics
import subprocess
import time
from pathlib import Path
from typing import Any


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BATCH = "0700"
LEGS = (("a0", "baseline"), ("b0", "candidate"), ("b1", "candidate"), ("a3", "baseline"))
SAMPLES = 300
WARMUPS = 10
STAT_FIELDS = ("p50_ns", "mean_ns", "p95_ns", "p99_ns", "min_ns", "max_ns")
NATIVE_SOURCES: dict[str, tuple[str | Path, str]] = {
    "one-real": (ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx", "one"),
    "noop-real": (ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx", "noop"),
    "two-real": (ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx", "two"),
    "one-control": (P / "marker-control.pptx", "one"),
    "one-generated": ("generated:12x8", "one"),
    "one-notes-poi": (ROOT / "test-data/poi/test-data/slideshow/prProps.pptx", "one"),
}
REFUSAL_CASES = {
    "small-valid",
    "generated-12x8-valid",
    "early-name-error",
    "late-root-error",
    "late-root-error-mce",
    "late-missing-relationship",
    "late-missing-relationship-mce",
    "notes-invalid-tail",
    "mixed-conformance",
    "slide-raw-overlimit-16m-to-64m",
}


def sha(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def sha_bytes(value: bytes) -> str:
    return hashlib.sha256(value).hexdigest()


def quantile(values: list[int], fraction: float) -> int:
    return sorted(values)[max(0, math.ceil(len(values) * fraction) - 1)]


def stats(values: list[int]) -> dict[str, float | int]:
    return {
        "samples": len(values),
        "p50_ns": statistics.median(values),
        "mean_ns": statistics.mean(values),
        "p95_ns": quantile(values, 0.95),
        "p99_ns": quantile(values, 0.99),
        "min_ns": min(values),
        "max_ns": max(values),
    }


def parse_native(path: Path) -> tuple[list[str], list[dict[str, int]], dict[str, list[str]]]:
    lines = path.read_text().splitlines()
    header_line = next((line for line in lines if line.startswith("sample\t")), None)
    if header_line is None:
        raise AssertionError(f"{path}: missing native sample header")
    header = header_line.split("\t")
    rows: list[dict[str, int]] = []
    metadata: dict[str, list[str]] = {}
    for line in lines:
        fields = line.split("\t")
        if fields[0].isdigit():
            if len(fields) != len(header):
                raise AssertionError(f"{path}: native row/header width mismatch")
            values = [int(value) for value in fields]
            if any(value < 0 for value in values):
                raise AssertionError(f"{path}: native sample value is negative")
            rows.append(dict(zip(header, values)))
        elif fields[0] not in {"sample", "sample_ns"}:
            metadata[fields[0]] = fields[1:]
    if len(rows) != SAMPLES or [row["sample"] for row in rows] != list(range(SAMPLES)):
        raise AssertionError(f"{path}: native sample census mismatch")
    if metadata.get("samples") != [str(SAMPLES)] or metadata.get("warmups") != [str(WARMUPS)]:
        raise AssertionError(f"{path}: native sample metadata mismatch")
    if metadata.get("probe") != [BATCH]:
        raise AssertionError(f"{path}: native probe identity mismatch")
    return header[1:], rows, metadata


def parse_refusal(
    path: Path,
) -> tuple[dict[str, str], dict[str, list[int]], dict[str, dict[str, list[str]]]]:
    lines = path.read_text().splitlines()
    header: dict[str, str] = {}
    cases: dict[str, list[int]] = {}
    metadata: dict[str, dict[str, list[str]]] = {}
    current: str | None = None
    for line in lines:
        fields = line.split("\t")
        if not fields:
            continue
        key = fields[0]
        if key == "case":
            if len(fields) != 2 or fields[1] in cases:
                raise AssertionError(f"{path}: malformed refusal case")
            current = fields[1]
            cases[current] = []
            metadata[current] = {}
        elif key == "sample_ns":
            continue
        elif key.isdigit():
            if current is None or len(fields) != 2 or int(key) != len(cases[current]):
                raise AssertionError(f"{path}: malformed refusal sample")
            value = int(fields[1])
            if value < 0:
                raise AssertionError(f"{path}: refusal sample value is negative")
            cases[current].append(value)
        elif key == "all_iterations_passed":
            if fields != ["all_iterations_passed", "true"]:
                raise AssertionError(f"{path}: refusal iteration failed")
            header[key] = fields[1]
        elif current is None:
            if len(fields) != 2:
                raise AssertionError(f"{path}: malformed refusal header")
            header[key] = fields[1]
        else:
            metadata[current][key] = fields[1:]
    if header.get("probe") != "0693-refusal":
        raise AssertionError(f"{path}: refusal probe identity mismatch")
    if header.get("samples") != str(SAMPLES) or header.get("warmups") != str(WARMUPS):
        raise AssertionError(f"{path}: refusal sample metadata mismatch")
    if header.get("all_iterations_passed") != "true" or set(cases) != REFUSAL_CASES:
        raise AssertionError(f"{path}: refusal case census mismatch")
    if any(len(values) != SAMPLES for values in cases.values()):
        raise AssertionError(f"{path}: refusal sample count mismatch")
    for case, fields in metadata.items():
        if fields.get("expected_error_debug") is None:
            raise AssertionError(f"{path}/{case}: missing expected error identity")
        if fields.get("observed_error_debug") not in (None, fields["expected_error_debug"]):
            raise AssertionError(f"{path}/{case}: refusal identity changed within run")
    return header, cases, metadata


def semantic_identity(value: Any) -> str:
    canonical = json.dumps(value, ensure_ascii=True, sort_keys=True, separators=(",", ":"))
    return sha_bytes(canonical.encode())


def source_hash(source: str | Path) -> str | None:
    if isinstance(source, str):
        if not source.startswith("generated:"):
            raise AssertionError(f"unsupported non-path source: {source}")
        return None
    return sha(source)


def run_command(kind: str, case: str | None, leg: str, phase: str, command: list[str], output: Path) -> dict[str, Any]:
    error = output.with_suffix(".stderr")
    started = time.monotonic()
    with output.open("w") as stdout, error.open("w") as stderr:
        result = subprocess.run(command, cwd=ROOT, stdout=stdout, stderr=stderr)
    record: dict[str, Any] = {
        "schema": "0700-followup-run-v1",
        "kind": kind,
        "case": case,
        "leg": leg,
        "phase": phase,
        "command": command,
        "exit_code": result.returncode,
        "seconds": time.monotonic() - started,
        "output": str(output.relative_to(P)),
        "output_sha256": sha(output),
        "stderr": str(error.relative_to(P)),
        "stderr_sha256": sha(error),
    }
    if result.returncode != 0:
        raise SystemExit(error.read_text())
    return record


def write_records(records: list[dict[str, Any]]) -> None:
    (P / "followup-runs.json").write_text(json.dumps(records, indent=2) + "\n")


def summarize(records: list[dict[str, Any]]) -> None:
    summary: dict[str, list[dict[str, Any]]] = {"native": [], "refusal": []}
    for record in records:
        path = P / record["output"]
        stderr = P / record["stderr"]
        if sha(path) != record["output_sha256"] or sha(stderr) != record["stderr_sha256"]:
            raise AssertionError(f"follow-up raw hash changed: {path}")
        if record["kind"] == "native":
            columns, rows, metadata = parse_native(path)
            summary["native"].append(
                {
                    "kind": "native",
                    "case": record["case"],
                    "leg": record["leg"],
                    "phase": record["phase"],
                    "workflow": record["workflow"],
                    "binary_sha256": record["binary_sha256"],
                    "source_sha256": record["source_sha256"],
                    "semantic_identity_sha256": record["semantic_identity_sha256"],
                    "metadata": metadata,
                    "stats": {column: stats([row[column] for row in rows]) for column in columns},
                }
            )
        elif record["kind"] == "refusal":
            _header, cases, metadata = parse_refusal(path)
            for case, values in cases.items():
                summary["refusal"].append(
                    {
                        "kind": "refusal",
                        "case": case,
                        "leg": record["leg"],
                        "phase": record["phase"],
                        "binary_sha256": record["binary_sha256"],
                        "semantic_identity_sha256": record["semantic_identity_sha256"][case],
                        "metadata": metadata[case],
                        "stats": {"total_ns": stats(values)},
                    }
                )
        else:
            raise AssertionError(f"unknown follow-up record kind: {record['kind']}")
    for rows in summary.values():
        rows.sort(key=lambda row: (row["case"], row["leg"]))
    if len(summary["native"]) != len(NATIVE_SOURCES) * len(LEGS) or len(summary["refusal"]) != len(REFUSAL_CASES) * len(LEGS):
        raise AssertionError("follow-up summary coverage mismatch")

    comparisons: list[dict[str, Any]] = []
    for kind in ("native", "refusal"):
        rows = summary[kind]
        for case in sorted({row["case"] for row in rows}):
            by_leg = {row["leg"]: row for row in rows if row["case"] == case}
            metrics = sorted({metric for row in by_leg.values() for metric in row["stats"]})
            for metric in metrics:
                for candidate, baseline in (("b0", "a0"), ("b1", "a3"), ("a3", "a0")):
                    candidate_stats = by_leg[candidate]["stats"][metric]
                    baseline_stats = by_leg[baseline]["stats"][metric]
                    comparisons.append(
                        {
                            "kind": kind,
                            "case": case,
                            "phase": metric,
                            "pair": f"{candidate}/{baseline}",
                            "candidate_leg": candidate,
                            "baseline_leg": baseline,
                            "candidate": {name: candidate_stats[name] for name in STAT_FIELDS},
                            "baseline": {name: baseline_stats[name] for name in STAT_FIELDS},
                            "delta_pct": {name: (candidate_stats[name] / baseline_stats[name] - 1) * 100 for name in STAT_FIELDS},
                            "delta_ns": {name: candidate_stats[name] - baseline_stats[name] for name in STAT_FIELDS},
                        }
                    )
    triggers = [
        {
            "kind": row["kind"],
            "case": row["case"],
            "phase": row["phase"],
            "pair": row["pair"],
            "metric": metric,
            "delta_pct": value,
            "delta_ns": row["delta_ns"][metric],
        }
        for row in comparisons
        if row["pair"] in {"b0/a0", "b1/a3"}
        for metric, value in row["delta_pct"].items()
        if value > 5
    ]
    (P / "followup-summary.json").write_text(json.dumps(summary, indent=2) + "\n")
    (P / "followup-comparisons.json").write_text(json.dumps(comparisons, indent=2) + "\n")
    (P / "followup-triggers.json").write_text(json.dumps(triggers, indent=2) + "\n")
    print("follow-up summary", len(summary["native"]), len(summary["refusal"]), "rows;", len(triggers), "triggers")


def main() -> None:
    out = P / "followup"
    native_out = out / "native"
    refusal_out = out / "refusal"
    native_out.mkdir(parents=True, exist_ok=True)
    refusal_out.mkdir(parents=True, exist_ok=True)
    records: list[dict[str, Any]] = []
    semantic: dict[tuple[str, str], str] = {}
    for index, (leg, phase) in enumerate(LEGS):
        binary = ROOT.parent / f"litchi-{BATCH}-bin" / f"{phase}-native"
        cases = list(NATIVE_SOURCES.items())
        if index % 2:
            cases.reverse()
        for case, (source, workflow) in cases:
            source_arg = str(source)
            command = ["taskset", "-c", "12", str(binary), "phases", source_arg, str(SAMPLES), str(WARMUPS), workflow]
            record = run_command("native", case, leg, phase, command, native_out / f"{case}-{leg}.tsv")
            record.update({"workflow": workflow, "binary_sha256": sha(binary), "source_sha256": source_hash(source)})
            columns, rows, metadata = parse_native(P / record["output"])
            for required in ("capture_ns", "clone_ns", "settext_ns", "commit_ns", "apply_ns", "total_ns"):
                if required not in columns:
                    raise AssertionError(f"{case}/{leg}: missing native phase {required}")
            digest = semantic_identity(metadata)
            record["semantic_identity_sha256"] = digest
            key = (case, "native")
            if key in semantic and semantic[key] != digest:
                raise AssertionError(f"{case}: native semantic identity changed across follow-up legs")
            semantic.setdefault(key, digest)
            records.append(record)
            write_records(records)
            print(case, leg, flush=True)

        binary = ROOT.parent / f"litchi-{BATCH}-bin" / f"{phase}-refusal"
        command = ["taskset", "-c", "12", str(binary), "matrix", str(SAMPLES), str(WARMUPS)]
        record = run_command("refusal", None, leg, phase, command, refusal_out / f"{leg}.tsv")
        record["binary_sha256"] = sha(binary)
        _header, cases, metadata = parse_refusal(P / record["output"])
        record["case_count"] = len(cases)
        record["semantic_identity_sha256"] = {case: semantic_identity(fields) for case, fields in metadata.items()}
        for case, digest in record["semantic_identity_sha256"].items():
            key = (case, "refusal")
            if key in semantic and semantic[key] != digest:
                raise AssertionError(f"{case}: refusal semantic identity changed across follow-up legs")
            semantic.setdefault(key, digest)
        records.append(record)
        write_records(records)
        print("refusal", leg, flush=True)
    summarize(records)


if __name__ == "__main__":
    main()
