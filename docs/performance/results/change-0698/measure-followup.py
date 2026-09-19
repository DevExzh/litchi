#!/usr/bin/env python3
"""Run the isolated 300-sample follow-up matrix without touching initial data."""

from __future__ import annotations

import hashlib
import json
import math
import statistics
import subprocess
import time
from pathlib import Path


P = Path(__file__).resolve().parent
ROOT = P.parents[3]
BATCH = "0698"
LEGS = (("a0", "baseline"), ("b0", "candidate"), ("b1", "candidate"), ("a3", "baseline"))
SAMPLES = 300
WARMUPS = 10
STAT_FIELDS = ("p50_ns", "mean_ns", "p95_ns", "p99_ns", "min_ns", "max_ns")

NATIVE_SOURCES = {
    "one-real": (ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx", "one"),
    "noop-real": (ROOT / "test-data/libreoffice-core/sd/qa/unit/data/pptx/slide-section-test.pptx", "noop"),
    "one-control": (P / "marker-control.pptx", "one"),
    "one-notes-poi": (ROOT / "test-data/poi/test-data/slideshow/prProps.pptx", "one"),
}
ORACLE_SOURCE = P / "oracle-declaration-controls" / "mixed.xml"
ORACLE_PROFILE = "opaque"
ORACLE_CASE = "mixed-opaque"


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
    header_line = next(line for line in lines if line.startswith("sample\t"))
    header = header_line.split("\t")
    rows = []
    metadata: dict[str, list[str]] = {}
    for line in lines:
        fields = line.split("\t")
        if fields[0].isdigit():
            if len(fields) != len(header):
                raise AssertionError(f"{path}: native row/header width mismatch")
            rows.append(dict(zip(header, (int(value) for value in fields))))
        elif fields[0] not in {"sample", "sample_ns"}:
            metadata[fields[0]] = fields[1:]
    if len(rows) != SAMPLES or [row["sample"] for row in rows] != list(range(SAMPLES)):
        raise AssertionError(f"{path}: native sample census mismatch")
    if metadata.get("samples") != [str(SAMPLES)] or metadata.get("warmups") != [str(WARMUPS)]:
        raise AssertionError(f"{path}: native sample metadata mismatch")
    if metadata.get("probe") != [BATCH]:
        raise AssertionError(f"{path}: native probe identity mismatch")
    return header[1:], rows, metadata


def parse_refusal(path: Path) -> tuple[dict[str, str], dict[str, list[int]], dict[str, dict[str, list[str]]]]:
    lines = path.read_text().splitlines()
    header: dict[str, str] = {}
    cases: dict[str, list[int]] = {}
    metadata: dict[str, dict[str, list[str]]] = {}
    current = None
    for line in lines:
        fields = line.split("\t")
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
            cases[current].append(int(fields[1]))
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
    if header.get("all_iterations_passed") != "true" or len(cases) != 10:
        raise AssertionError(f"{path}: refusal case census mismatch")
    if any(len(values) != SAMPLES for values in cases.values()):
        raise AssertionError(f"{path}: refusal sample count mismatch")
    for case, fields in metadata.items():
        if fields.get("expected_error_debug") is None:
            raise AssertionError(f"{path}/{case}: missing expected error identity")
        if fields.get("observed_error_debug") not in (None, fields["expected_error_debug"]):
            raise AssertionError(f"{path}/{case}: refusal identity changed within run")
    return header, cases, metadata


def parse_fields(line: str, kind: str) -> dict[str, str]:
    fields = line.split("\t")
    if not fields or fields[0] != kind:
        raise AssertionError(f"malformed {kind} record")
    result: dict[str, str] = {}
    for field in fields[1:]:
        key, separator, value = field.partition("=")
        if not separator or not key or key in result:
            raise AssertionError(f"malformed {kind} field: {field!r}")
        result[key] = value
    return result


def parse_oracle_identity(path: Path) -> tuple[bytes, dict[str, str]]:
    raw = path.read_bytes()
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise AssertionError(f"{path}: oracle identity is not UTF-8: {exc}") from exc
    lines = text.splitlines()
    if len(lines) != 1 or not text.endswith("\n"):
        raise AssertionError(f"{path}: oracle identity must be one newline-terminated line")
    fields = parse_fields(lines[0], "OK")
    if fields.get("probe") != BATCH + "-mce-oracle" or fields.get("profile") != ORACLE_PROFILE:
        raise AssertionError(f"{path}: oracle identity probe/profile mismatch")
    if not fields.get("input_len", "").isdigit() or int(fields["input_len"]) <= 0:
        raise AssertionError(f"{path}: oracle identity input length is invalid")
    return raw, fields


def parse_oracle_timing(path: Path) -> tuple[dict[str, str], list[int]]:
    lines = path.read_text().splitlines()
    if not lines or not lines[0].startswith("TIMING\t"):
        raise AssertionError(f"{path}: missing oracle timing header")
    header = parse_fields(lines[0], "TIMING")
    if header.get("probe") != BATCH + "-mce-oracle" or header.get("profile") != ORACLE_PROFILE:
        raise AssertionError(f"{path}: oracle timing probe/profile mismatch")
    if header.get("warmups") != str(WARMUPS) or header.get("samples") != str(SAMPLES):
        raise AssertionError(f"{path}: oracle timing sample metadata mismatch")
    if not header.get("input_len", "").isdigit() or int(header["input_len"]) <= 0:
        raise AssertionError(f"{path}: oracle timing input length is invalid")
    samples: list[int] = []
    for line in lines[1:]:
        fields = line.split("\t")
        if len(fields) != 3 or fields[0] != "SAMPLE" or not fields[1].startswith("index=") or not fields[2].startswith("elapsed_ns="):
            raise AssertionError(f"{path}: malformed oracle sample")
        index = fields[1][len("index="):]
        elapsed = fields[2][len("elapsed_ns="):]
        if not index.isdigit() or int(index) != len(samples) or not elapsed.isdigit() or int(elapsed) <= 0:
            raise AssertionError(f"{path}: malformed oracle sample index or duration")
        samples.append(int(elapsed))
    if len(samples) != SAMPLES:
        raise AssertionError(f"{path}: oracle sample census mismatch")
    return header, samples


def identity(metadata: dict[str, list[str]]) -> str:
    canonical = json.dumps(metadata, ensure_ascii=True, sort_keys=True, separators=(",", ":"))
    return hashlib.sha256(canonical.encode()).hexdigest()


def run_command(kind: str, case: str | None, leg: str, phase: str, command: list[str], output: Path) -> dict:
    error = output.with_suffix(".stderr")
    started = time.monotonic()
    with output.open("w") as stdout, error.open("w") as stderr:
        result = subprocess.run(command, cwd=ROOT, stdout=stdout, stderr=stderr)
    record = {
        "kind": kind,
        "case": case,
        "leg": leg,
        "phase": phase,
        "command": command,
        "exit_code": result.returncode,
        "seconds": time.monotonic() - started,
        "output": str(output.relative_to(P)),
        "output_sha256": sha(output),
        "stderr_sha256": sha(error),
    }
    if result.returncode != 0:
        raise SystemExit(error.read_text())
    return record


def write_records(records: list[dict]) -> None:
    (P / "followup-runs.json").write_text(json.dumps(records, indent=2) + "\n")


def main() -> None:
    out = P / "followup"
    native_out = out / "native"
    refusal_out = out / "refusal"
    oracle_out = out / "oracle-declaration"
    native_out.mkdir(parents=True, exist_ok=True)
    refusal_out.mkdir(parents=True, exist_ok=True)
    oracle_out.mkdir(parents=True, exist_ok=True)
    records: list[dict] = []
    semantic: dict[tuple[str, str], str] = {}
    oracle_identity: bytes | None = None

    for index, (leg, phase) in enumerate(LEGS):
        binary = ROOT.parent / f"litchi-{BATCH}-bin" / f"{phase}-native"
        cases = list(NATIVE_SOURCES.items())
        if index % 2:
            cases.reverse()
        for case, (source, workflow) in cases:
            command = [
                "taskset",
                "-c",
                "12",
                str(binary),
                "phases",
                str(source),
                str(SAMPLES),
                str(WARMUPS),
                workflow,
            ]
            record = run_command("native", case, leg, phase, command, native_out / f"{case}-{leg}.tsv")
            record["workflow"] = workflow
            record["binary_sha256"] = sha(binary)
            record["source_sha256"] = sha(source)
            columns, rows, metadata = parse_native(P / record["output"])
            for required in ("capture_ns", "clone_ns", "settext_ns", "commit_ns", "apply_ns", "total_ns"):
                if required not in columns:
                    raise AssertionError(f"{case}/{leg}: missing native phase {required}")
            digest = identity(metadata)
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
        header, cases, metadata = parse_refusal(P / record["output"])
        record["case_count"] = len(cases)
        record["semantic_identity_sha256"] = {
            case: identity(fields) for case, fields in metadata.items()
        }
        for case, digest in record["semantic_identity_sha256"].items():
            key = (case, "refusal")
            if key in semantic and semantic[key] != digest:
                raise AssertionError(f"{case}: refusal semantic identity changed across follow-up legs")
            semantic.setdefault(key, digest)
        records.append(record)
        write_records(records)
        print("refusal", leg, flush=True)

        binary = ROOT.parent / f"litchi-{BATCH}-bin" / f"{phase}-oracle"
        identity_command = [str(binary), "run", ORACLE_PROFILE, str(ORACLE_SOURCE)]
        identity_output = oracle_out / f"{ORACLE_CASE}-{leg}.identity"
        identity_error = oracle_out / f"{ORACLE_CASE}-{leg}.identity.stderr"
        with identity_output.open("wb") as stdout, identity_error.open("wb") as stderr:
            identity_result = subprocess.run(
                identity_command,
                cwd=ROOT,
                stdout=stdout,
                stderr=stderr,
            )
        identity_bytes, identity_fields = parse_oracle_identity(identity_output)
        if identity_result.returncode != 0:
            raise SystemExit(identity_error.read_text())
        if identity_error.read_bytes():
            raise AssertionError(f"{ORACLE_CASE}/{leg}: oracle identity emitted stderr")
        if oracle_identity is not None and identity_bytes != oracle_identity:
            raise AssertionError(f"{ORACLE_CASE}: oracle identity changed across follow-up legs")
        oracle_identity = identity_bytes

        command = [
            "taskset",
            "-c",
            "12",
            str(binary),
            "time",
            ORACLE_PROFILE,
            str(ORACLE_SOURCE),
            str(WARMUPS),
            str(SAMPLES),
        ]
        output = oracle_out / f"{ORACLE_CASE}-{leg}.tsv"
        record = run_command("oracle", ORACLE_CASE, leg, phase, command, output)
        record["binary_sha256"] = sha(binary)
        record["source_sha256"] = sha(ORACLE_SOURCE)
        record["profile"] = ORACLE_PROFILE
        record["identity_command"] = identity_command
        record["identity_exit_code"] = identity_result.returncode
        record["identity"] = identity_bytes.decode("utf-8")
        record["identity_stdout"] = str(identity_output.relative_to(P))
        record["identity_stderr"] = str(identity_error.relative_to(P))
        record["identity_stdout_sha256"] = sha(identity_output)
        record["identity_stderr_sha256"] = sha(identity_error)
        record["semantic_identity_sha256"] = sha_bytes(identity_bytes)
        header, samples = parse_oracle_timing(P / record["output"])
        if any(identity_fields.get(key) != header.get(key) for key in ("probe", "profile", "input_len")):
            raise AssertionError(f"{ORACLE_CASE}/{leg}: oracle identity/header mismatch")
        record["sample_count"] = len(samples)
        records.append(record)
        write_records(records)
        print("oracle", leg, flush=True)

    summarize(records)


def summarize(records: list[dict]) -> None:
    summary: dict[str, list[dict]] = {"native": [], "refusal": [], "oracle": []}
    refusal_case_ids = set()
    for record in records:
        path = P / record["output"]
        if sha(path) != record["output_sha256"] or sha(path.with_suffix(".stderr")) != record["stderr_sha256"]:
            raise AssertionError(f"raw follow-up hash changed: {path}")
        if record["kind"] == "native":
            columns, rows, metadata = parse_native(path)
            row = {
                "kind": "native",
                "case": record["case"],
                "leg": record["leg"],
                "phase": record["phase"],
                "workflow": record["workflow"],
                "binary_sha256": record["binary_sha256"],
                "source_sha256": record["source_sha256"],
                "semantic_identity_sha256": record["semantic_identity_sha256"],
                "metadata": metadata,
                "stats": {column: stats([item[column] for item in rows]) for column in columns if column != "sample"},
            }
            summary["native"].append(row)
        elif record["kind"] == "refusal":
            header, cases, metadata = parse_refusal(path)
            for case, values in cases.items():
                refusal_case_ids.add(case)
                summary["refusal"].append({
                    "kind": "refusal",
                    "case": case,
                    "leg": record["leg"],
                    "phase": record["phase"],
                    "binary_sha256": record["binary_sha256"],
                    "semantic_identity_sha256": record["semantic_identity_sha256"][case],
                    "metadata": metadata[case],
                    "stats": {"total_ns": stats(values)},
                })
        elif record["kind"] == "oracle":
            header, samples = parse_oracle_timing(path)
            identity_path = P / record["identity_stdout"]
            identity_bytes, identity_fields = parse_oracle_identity(identity_path)
            if identity_bytes.decode("utf-8") != record["identity"]:
                raise AssertionError(f"oracle identity text changed: {identity_path}")
            if record.get("identity_exit_code") != 0:
                raise AssertionError(f"oracle identity command failed: {identity_path}")
            if sha(identity_path) != record["identity_stdout_sha256"]:
                raise AssertionError(f"oracle identity hash changed: {identity_path}")
            if sha(P / record["identity_stderr"]) != record["identity_stderr_sha256"]:
                raise AssertionError(f"oracle identity stderr hash changed: {identity_path}")
            if any(identity_fields.get(key) != header.get(key) for key in ("probe", "profile", "input_len")):
                raise AssertionError(f"oracle identity/header mismatch: {path}")
            summary["oracle"].append({
                "kind": "oracle",
                "case": record["case"],
                "leg": record["leg"],
                "phase": record["phase"],
                "profile": record["profile"],
                "binary_sha256": record["binary_sha256"],
                "source_sha256": record["source_sha256"],
                "identity": record["identity"],
                "identity_stdout_sha256": record["identity_stdout_sha256"],
                "identity_stderr_sha256": record["identity_stderr_sha256"],
                "semantic_identity_sha256": record["semantic_identity_sha256"],
                "stats": {"total_ns": stats(samples)},
            })
        else:
            raise AssertionError(f"unknown follow-up record kind: {record['kind']}")
    for rows in summary.values():
        rows.sort(key=lambda row: (row["case"], row["leg"]))
    if (
        len(summary["native"]) != len(NATIVE_SOURCES) * len(LEGS)
        or len(summary["refusal"]) != 10 * len(LEGS)
        or len(summary["oracle"]) != len(LEGS)
    ):
        raise AssertionError("follow-up summary coverage mismatch")
    comparisons = []
    for kind, rows in summary.items():
        cases = sorted({row["case"] for row in rows})
        phases = sorted({phase for row in rows for phase in row["stats"]})
        for case in cases:
            for metric in phases:
                by_leg = {row["leg"]: row for row in rows if row["case"] == case}
                for candidate, baseline in (("b0", "a0"), ("b1", "a3"), ("a3", "a0")):
                    if candidate not in by_leg or baseline not in by_leg:
                        raise AssertionError(f"follow-up comparison coverage: {kind}/{case}/{candidate}/{baseline}")
                    candidate_stats = by_leg[candidate]["stats"][metric]
                    baseline_stats = by_leg[baseline]["stats"][metric]
                    delta_pct = {
                        name: (candidate_stats[name] / baseline_stats[name] - 1) * 100
                        for name in STAT_FIELDS
                    }
                    comparisons.append({
                        "kind": kind,
                        "case": case,
                        "phase": metric,
                        "pair": f"{candidate}/{baseline}",
                        "candidate_leg": candidate,
                        "baseline_leg": baseline,
                        "candidate": {name: candidate_stats[name] for name in STAT_FIELDS},
                        "baseline": {name: baseline_stats[name] for name in STAT_FIELDS},
                        "delta_pct": delta_pct,
                        "delta_ns": {
                            name: candidate_stats[name] - baseline_stats[name] for name in STAT_FIELDS
                        },
                    })
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


if __name__ == "__main__":
    main()
