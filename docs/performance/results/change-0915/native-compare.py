#!/usr/bin/env python3
"""Capture and replay the frozen 0915 native DOCX comparison.

The script deliberately keeps source/build custody outside the repository.  A
``freeze`` records immutable binary, protocol, script, validator, and optional
receipt identities.  ``run`` executes only the frozen 8-pair AB/BA schedule;
it never replaces a child directory or committed receipt.  ``reduce`` and
``replay`` validate the reports through the 0914 harness validator and retain
only small summaries under this change directory.

No action in this file adopts a production candidate.  The paired timing gate
is evidence for the caller's review only.
"""

from __future__ import annotations

import argparse
import datetime as dt
import gzip
import hashlib
import importlib.util
import json
import math
import os
from pathlib import Path
import random
import statistics
import subprocess
import sys
from typing import Any


HERE = Path(__file__).resolve().parent
ROOT = HERE.parents[3]
PROTOCOL_PATH = HERE / "native-protocol.json"
SCRIPT_PATH = HERE / "native-compare.py"
VALIDATOR_PATH = HERE.parent / "change-0914" / "capture.py"
MANIFEST_PATH = HERE / "manifest.json"
RECEIPTS_DIR = HERE / "receipts"
REPORTS_DIR = HERE / "reports"
EXTERNAL = Path(
    os.environ.get("LITCHI_PERF_0915_ROOT", "/home/zhuhe/litchi-performance-0915")
).resolve()
COMPARISON = EXTERNAL / "comparison"
DEFAULT_BASELINE = Path(
    os.environ.get(
        "LITCHI_PERF_0915_BASELINE",
        "/home/zhuhe/litchi-performance-0914/target/release/litchi-perf-baseline",
    )
).resolve()


class EvidenceError(RuntimeError):
    pass


def fail(message: str) -> None:
    raise EvidenceError(message)


def require(condition: bool, message: str) -> None:
    if not condition:
        fail(message)


def now() -> str:
    return dt.datetime.now(dt.timezone.utc).isoformat()


def parse_utc(value: object, label: str) -> dt.datetime:
    require(isinstance(value, str), f"{label} timestamp missing")
    try:
        parsed = dt.datetime.fromisoformat(value)
    except ValueError as error:
        fail(f"{label} timestamp malformed: {error}")
    require(parsed.tzinfo is not None, f"{label} timestamp has no timezone")
    return parsed.astimezone(dt.timezone.utc)


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def descriptor(path: Path) -> dict[str, object]:
    original = Path(path)
    require(original.is_file() and not original.is_symlink(), f"missing regular file: {original}")
    path = original.resolve()
    return {"path": str(path), "bytes": path.stat().st_size, "sha256": sha256(path)}


def json_bytes(value: object) -> bytes:
    return (json.dumps(value, indent=2, sort_keys=True) + "\n").encode("utf-8")


def read_json(path: Path) -> object:
    require(path.is_file() and not path.is_symlink(), f"missing JSON: {path}")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as error:
        fail(f"invalid JSON {path}: {error}")


def write_exclusive(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    require(not path.exists(), f"refusing to replace retained file: {path}")
    temporary = path.with_name(path.name + ".tmp")
    require(not temporary.exists(), f"temporary output already exists: {temporary}")
    try:
        with temporary.open("xb") as stream:
            stream.write(data)
        if path.exists():
            fail(f"retained output appeared during write: {path}")
        temporary.replace(path)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise


def write_json(path: Path, value: object) -> None:
    write_exclusive(path, json_bytes(value))


def copy_committed_file(source: Path, target: Path) -> dict[str, object]:
    source_identity = descriptor(source)
    data = source.read_bytes()
    require(len(data) < 1 << 20, f"refusing committed evidence at or above 1 MiB: {source}")
    write_exclusive(target, data)
    committed = descriptor(target)
    require(
        committed["bytes"] == source_identity["bytes"]
        and committed["sha256"] == source_identity["sha256"],
        f"source changed while copying committed evidence: {source}",
    )
    return committed


def repository_relative(path: Path) -> str | None:
    """Return a Git-relative path for a file in this checkout."""
    try:
        return path.resolve().relative_to(ROOT).as_posix()
    except ValueError:
        return None


def recorded_path(identity: dict[str, object], manifest_value: dict[str, Any]) -> Path:
    """Resolve a retained Git file from the repository recorded at freeze."""
    require(isinstance(identity.get("path"), str), "retained descriptor path missing")
    recorded = Path(str(identity["path"]))
    repository = Path(str(manifest_value["repository"])).resolve()
    if recorded.is_absolute():
        try:
            relative = recorded.resolve().relative_to(repository)
        except ValueError:
            return recorded
    else:
        relative = recorded
    candidate = (ROOT / relative).resolve()
    try:
        candidate.relative_to(ROOT)
    except ValueError:
        fail(f"retained descriptor escapes repository: {recorded}")
    return candidate


def verify_recorded(identity: dict[str, object], target: Path, label: str) -> None:
    """Check a retained descriptor by content while allowing checkout relocation."""
    actual = descriptor(target)
    require(
        actual["bytes"] == identity.get("bytes")
        and actual["sha256"] == identity.get("sha256"),
        f"{label} changed since freeze",
    )


def recorded_descriptor(target: Path, manifest_value: dict[str, Any]) -> dict[str, object]:
    """Describe a retained file, preserving the original checkout path label."""
    actual = descriptor(target)
    relative = repository_relative(target)
    if relative is not None:
        actual["path"] = str(Path(str(manifest_value["repository"])) / relative)
    return actual


def recorded_manifest_descriptor(manifest_value: dict[str, Any]) -> dict[str, object]:
    """Describe manifest.json without making its self-hash recursive."""
    return recorded_descriptor(MANIFEST_PATH, manifest_value)


def read_gzip(path: Path) -> bytes:
    try:
        with gzip.open(path, "rb") as stream:
            return stream.read()
    except (OSError, EOFError) as error:
        fail(f"invalid gzip report {path}: {error}")


def write_deterministic_gzip(path: Path, data: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    require(not path.exists(), f"refusing to replace retained gzip: {path}")
    temporary = path.with_name(path.name + ".tmp")
    require(not temporary.exists(), f"temporary gzip already exists: {temporary}")
    try:
        with temporary.open("xb") as output:
            with gzip.GzipFile(
                filename="", mode="wb", fileobj=output, compresslevel=9, mtime=0
            ) as stream:
                stream.write(data)
        if path.exists():
            fail(f"retained gzip appeared during write: {path}")
        temporary.replace(path)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise


def git(*args: str) -> str:
    result = subprocess.run(
        ["git", *args],
        cwd=ROOT,
        stdin=subprocess.DEVNULL,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        check=False,
    )
    if result.returncode:
        fail(f"git {' '.join(args)} failed: {result.stderr.strip()}")
    return result.stdout.strip()


def protocol() -> dict[str, Any]:
    value = read_json(PROTOCOL_PATH)
    require(isinstance(value, dict), "native protocol is not an object")
    require(
        value.get("schema") == "litchi.performance.0915.native-protocol.v1",
        "unexpected native protocol schema",
    )
    require(value.get("case") == "docx_source_backed_existing_section_layout_edit_save", "protocol case changed")
    require(value.get("warmups") == 30 and value.get("samples_per_child") == 200, "protocol samples changed")
    require(value.get("cpu") == 12 and value.get("pairs") == 8, "protocol CPU/pair count changed")
    orders = value.get("pair_orders")
    require(
        orders == [
            "baseline,candidate",
            "candidate,baseline",
            "baseline,candidate",
            "candidate,baseline",
            "baseline,candidate",
            "candidate,baseline",
            "baseline,candidate",
            "candidate,baseline",
        ],
        "protocol pair schedule changed",
    )
    bootstrap = value.get("bootstrap")
    require(
        isinstance(bootstrap, dict)
        and bootstrap.get("draws") == 10000
        and bootstrap.get("seed") == 915,
        "protocol bootstrap changed",
    )
    gate = value.get("timing_gate")
    require(
        isinstance(gate, dict)
        and gate.get("minimum_median_ratio") == 1.05
        and gate.get("lower_95_percent_bound_must_exceed") == 1.0,
        "protocol timing gate changed",
    )
    return value


def validator_module() -> Any:
    require(VALIDATOR_PATH.is_file(), f"missing 0914 validator: {VALIDATOR_PATH}")
    spec = importlib.util.spec_from_file_location("litchi_perf_0914_capture", VALIDATOR_PATH)
    require(spec is not None and spec.loader is not None, "cannot load 0914 validator")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def frozen_protocol_fields(value: dict[str, Any]) -> dict[str, object]:
    return {
        "schema": value["schema"],
        "adoption": value["adoption"],
        "baseline_source_revision": value["baseline_source_revision"],
        "case": value["case"],
        "warmups": value["warmups"],
        "samples_per_child": value["samples_per_child"],
        "cpu": value["cpu"],
        "pairs": value["pairs"],
        "pair_orders": value["pair_orders"],
        "bootstrap": value["bootstrap"],
        "timing_gate": value["timing_gate"],
        "review_flags": value["review_flags"],
    }


def child_specs(manifest: dict[str, Any]) -> list[dict[str, object]]:
    result: list[dict[str, object]] = []
    for pair, order in enumerate(manifest["protocol_fields"]["pair_orders"], start=1):
        variants = order.split(",")
        for leg, variant in enumerate(variants, start=1):
            result.append(
                {
                    "pair": pair,
                    "leg": leg,
                    "variant": variant,
                    "label": f"pair-{pair:02d}-leg-{leg}-{variant}",
                }
            )
    return result


def validate_manifest(manifest: dict[str, Any], *, require_binaries: bool = False) -> dict[str, Any]:
    require(manifest.get("schema") == "litchi.performance.0915.native-compare-freeze.v1", "manifest schema mismatch")
    require(isinstance(manifest.get("repository"), str), "manifest repository missing")
    frozen = protocol()
    require(manifest.get("protocol_fields") == frozen_protocol_fields(frozen), "frozen protocol fields changed")
    created_at = parse_utc(manifest.get("created_utc"), "manifest created")
    frozen_at = parse_utc(frozen.get("frozen_utc"), "protocol frozen")
    require(frozen_at <= created_at, "manifest predates frozen protocol")
    for key, path in (("protocol", PROTOCOL_PATH), ("script", SCRIPT_PATH), ("validator", VALIDATOR_PATH)):
        identity = manifest.get(key)
        require(isinstance(identity, dict), f"{key} descriptor missing")
        verify_recorded(identity, recorded_path(identity, manifest), f"{key} descriptor")
    for key in ("baseline", "candidate"):
        value = manifest.get(key)
        require(isinstance(value, dict) and set(value) == {"path", "bytes", "sha256"}, f"{key} descriptor malformed")
        require(isinstance(value["path"], str) and isinstance(value["bytes"], int) and isinstance(value["sha256"], str), f"{key} descriptor types")
        if require_binaries:
            require(descriptor(Path(value["path"])) == value, f"{key} binary changed since freeze")
    receipts = manifest.get("optional_receipts", [])
    require(isinstance(receipts, list), "optional receipt list malformed")
    for item in receipts:
        require(isinstance(item, dict) and item.get("kind") in {"source", "build"}, "optional receipt kind malformed")
        identity = item.get("descriptor")
        require(isinstance(identity, dict), "optional receipt descriptor missing")
        verify_recorded(identity, recorded_path(identity, manifest), f"optional {item['kind']} receipt")
    return manifest


def manifest() -> dict[str, Any]:
    value = read_json(MANIFEST_PATH)
    require(isinstance(value, dict), "manifest is not an object")
    return validate_manifest(value)


def freeze(args: argparse.Namespace) -> None:
    require(not MANIFEST_PATH.exists(), f"manifest already exists: {MANIFEST_PATH}")
    p = protocol()
    baseline_path = Path(args.baseline).resolve() if args.baseline else DEFAULT_BASELINE
    candidate_path = Path(args.candidate).resolve()
    baseline = descriptor(baseline_path)
    candidate = descriptor(candidate_path)
    require(baseline["bytes"] == p["baseline_binary_bytes"], "baseline binary size differs from frozen protocol")
    require(baseline["sha256"] == p["baseline_binary_sha256"], "baseline binary hash differs from frozen protocol")
    require(baseline["path"] != candidate["path"], "candidate must not reuse baseline path")
    optional: list[dict[str, object]] = []
    for kind, raw in (("source", args.source_receipt), ("build", args.build_receipt)):
        if raw:
            path = Path(raw).resolve()
            optional.append({"kind": kind, "input": descriptor(path), "descriptor": None})
    if COMPARISON.exists():
        require(not any(COMPARISON.iterdir()), f"external comparison directory is not empty: {COMPARISON}")
    COMPARISON.mkdir(parents=True, exist_ok=True)
    for directory in (RECEIPTS_DIR, REPORTS_DIR):
        if directory.exists():
            require(not any(directory.iterdir()), f"committed output directory is not empty: {directory}")
        directory.mkdir(parents=True, exist_ok=True)
    for item in optional:
        source = Path(item["input"]["path"])
        target = RECEIPTS_DIR / f"{item['kind']}-receipt.json"
        item["descriptor"] = copy_committed_file(source, target)
    git_status = git("status", "--porcelain")
    value = {
        "schema": "litchi.performance.0915.native-compare-freeze.v1",
        "created_utc": now(),
        "repository": str(ROOT),
        "git_head_at_freeze": git("rev-parse", "HEAD"),
        "worktree_status_at_freeze": git_status,
        "protocol_fields": frozen_protocol_fields(p),
        "protocol": descriptor(PROTOCOL_PATH),
        "script": descriptor(SCRIPT_PATH),
        "validator": descriptor(VALIDATOR_PATH),
        "baseline": baseline,
        "candidate": candidate,
        "optional_receipts": optional,
        "external_root": str(EXTERNAL),
        "comparison_root": str(COMPARISON),
        "child_labels": [item["label"] for item in child_specs({"protocol_fields": frozen_protocol_fields(p)})],
    }
    write_json(MANIFEST_PATH, value)
    print(json.dumps({"status": "frozen", "manifest": descriptor(MANIFEST_PATH), "baseline": baseline, "candidate": candidate}))


def run_command(argv: list[str], cwd: Path, env: dict[str, str], stdout: Path, stderr: Path) -> int:
    with stdout.open("wb") as out, stderr.open("wb") as err:
        try:
            result = subprocess.run(
                argv,
                cwd=cwd,
                env=env,
                stdin=subprocess.DEVNULL,
                stdout=out,
                stderr=err,
                check=False,
            )
        except OSError:
            return 127
    return result.returncode


def report_summary_path(label: str) -> Path:
    return REPORTS_DIR / f"{label}.json"


def receipt_path(label: str) -> Path:
    return RECEIPTS_DIR / f"{label}.json"


def require_previous_children_finished(
    manifest_value: dict[str, Any], spec: dict[str, object], started: str
) -> None:
    """Require the frozen schedule's earlier children to be complete first."""
    schedule = child_specs(manifest_value)
    index = next(index for index, item in enumerate(schedule) if item["label"] == spec["label"])
    started_at = parse_utc(started, f"{spec['label']} start")
    for previous in schedule[:index]:
        previous_label = str(previous["label"])
        path = receipt_path(previous_label)
        require(path.is_file(), f"cannot start {spec['label']}: prior child {previous_label} is missing")
        value = read_json(path)
        require(isinstance(value, dict), f"prior child {previous_label} receipt malformed")
        require(value.get("status") == "passed", f"cannot start {spec['label']}: prior child {previous_label} failed")
        finished_at = parse_utc(value.get("finished_utc"), f"{previous_label} finish")
        require(finished_at <= started_at, f"schedule overlap: {previous_label} finished after {spec['label']} started")


def run_child(manifest_value: dict[str, Any], spec: dict[str, object]) -> None:
    label = str(spec["label"])
    variant = str(spec["variant"])
    require(variant in {"baseline", "candidate"}, f"unknown child variant: {variant}")
    child = COMPARISON / label
    require(not child.exists(), f"refusing to replace external child directory: {child}")
    require(not receipt_path(label).exists(), f"refusing to replace child receipt: {receipt_path(label)}")
    require(not report_summary_path(label).exists(), f"refusing to replace child summary: {report_summary_path(label)}")
    binary = manifest_value[variant]
    report = child / "report.json"
    compressed = child / "report.json.gz"
    process = child / "process.txt"
    stdout = child / "stdout.txt"
    stderr = child / "stderr.txt"
    argv = [
        "/usr/bin/taskset", "-c", str(manifest_value["protocol_fields"]["cpu"]),
        "/usr/bin/time", "-v", "-o", str(process), str(binary["path"]),
        "--warmup", str(manifest_value["protocol_fields"]["warmups"]),
        "--samples", str(manifest_value["protocol_fields"]["samples_per_child"]),
        "--case", str(manifest_value["protocol_fields"]["case"]),
        "--json", str(report),
    ]
    env = {"PATH": "/usr/bin:/bin", "LC_ALL": "C", "LANG": "C", "TZ": "UTC", "RUST_BACKTRACE": "0"}
    started = now()
    require_previous_children_finished(manifest_value, spec, started)
    require(
        parse_utc(manifest_value["created_utc"], "manifest created")
        <= parse_utc(started, f"{label} start"),
        f"{label} started before manifest freeze",
    )
    child.mkdir(parents=True)
    start = {
        "schema": "litchi.performance.0915.child-start.v1",
        "label": label,
        "pair": spec["pair"],
        "leg": spec["leg"],
        "variant": variant,
        "started_utc": started,
        "argv": argv,
        "cwd": str(ROOT),
        "environment": env,
        "manifest": recorded_manifest_descriptor(manifest_value),
        "protocol": manifest_value["protocol"],
        "script": manifest_value["script"],
        "validator": manifest_value["validator"],
        "binary": binary,
        "cpu": manifest_value["protocol_fields"]["cpu"],
        "host_load_before": list(os.getloadavg()),
    }
    write_json(child / "start.json", start)
    exit_code = run_command(argv, ROOT, env, stdout, stderr)
    finished = now()
    require(parse_utc(started, f"{label} start") <= parse_utc(finished, f"{label} finish"), f"{label} finish precedes start")
    raw_identity: dict[str, object] | None = None
    gzip_identity: dict[str, object] | None = None
    validation: dict[str, object] | None = None
    validation_error: str | None = None
    if report.is_file():
        raw = report.read_bytes()
        raw_identity = {"bytes": len(raw), "sha256": sha256_bytes(raw)}
        write_deterministic_gzip(compressed, raw)
        try:
            value = json.loads(raw.decode("utf-8"))
            validation = validator_module().validate_report(value, label, binary)
        except Exception as error:  # preserve the failed report and receipt
            validation_error = str(error)
        gzip_identity = descriptor(compressed)
    else:
        validation_error = "child produced no report"
    compact = {
        "schema": "litchi.performance.0915.child-summary.v1",
        "label": label,
        "pair": spec["pair"],
        "leg": spec["leg"],
        "variant": variant,
        "report_raw": raw_identity,
        "report_gzip": gzip_identity,
        "validation": validation,
        "validation_error": validation_error,
    }
    committed_report: dict[str, object] | None = None
    if compressed.is_file():
        committed_report_path = REPORTS_DIR / f"{label}.report.json.gz"
        copy_committed_file(compressed, committed_report_path)
        committed_report = recorded_descriptor(committed_report_path, manifest_value)
    committed_process: dict[str, object] | None = None
    committed_process_path = RECEIPTS_DIR / f"{label}.process.txt"
    if process.is_file():
        copy_committed_file(process, committed_process_path)
        committed_process = recorded_descriptor(committed_process_path, manifest_value)
    committed_stdout_path = RECEIPTS_DIR / f"{label}.stdout.txt"
    copy_committed_file(stdout, committed_stdout_path)
    committed_stdout = recorded_descriptor(committed_stdout_path, manifest_value)
    committed_stderr_path = RECEIPTS_DIR / f"{label}.stderr.txt"
    copy_committed_file(stderr, committed_stderr_path)
    committed_stderr = recorded_descriptor(committed_stderr_path, manifest_value)
    process_metrics: dict[str, object] | None = None
    if committed_process is not None:
        try:
            process_metrics = validator_module().parse_time_v(
                committed_process_path
            )
            process_metrics["path"] = str(committed_process["path"])
        except Exception as error:
            validation_error = validation_error or str(error)
    compact["report_gzip"] = committed_report
    compact["validation_error"] = validation_error
    write_json(report_summary_path(label), compact)
    receipt = {
        "schema": "litchi.performance.0915.child-receipt.v1",
        "label": label,
        "pair": spec["pair"],
        "leg": spec["leg"],
        "variant": variant,
        "started_utc": started,
        "finished_utc": finished,
        "argv": argv,
        "cwd": str(ROOT),
        "environment": env,
        "manifest": recorded_manifest_descriptor(manifest_value),
        "protocol": manifest_value["protocol"],
        "script": manifest_value["script"],
        "validator": manifest_value["validator"],
        "binary": binary,
        "cpu": manifest_value["protocol_fields"]["cpu"],
        "host_load_before": start["host_load_before"],
        "host_load_after": list(os.getloadavg()),
        "exit_code": exit_code,
        "process": committed_process,
        "external_process": descriptor(process) if process.is_file() else None,
        "process_metrics": process_metrics,
        "stdout": committed_stdout,
        "stderr": committed_stderr,
        "external_stdout": descriptor(stdout),
        "external_stderr": descriptor(stderr),
        "report_raw": raw_identity,
        "report_gzip": committed_report,
        "external_report_gzip": gzip_identity,
        "summary": recorded_descriptor(report_summary_path(label), manifest_value),
        "status": "passed" if exit_code == 0 and validation is not None and validation_error is None else "failed",
        "validation": validation,
        "validation_error": validation_error,
    }
    write_json(receipt_path(label), receipt)
    if receipt["status"] != "passed":
        fail(f"{label} failed; retained receipt and artifacts under {child}")
    print(json.dumps({"status": "child-complete", "label": label, "report_gzip": gzip_identity, "process": process_metrics}))


def run(args: argparse.Namespace) -> None:
    m = manifest()
    validate_manifest(m, require_binaries=True)
    selected = child_specs(m)
    if args.pair is not None:
        selected = [item for item in selected if item["pair"] == args.pair]
        require(selected, f"unknown pair: {args.pair}")
    if args.label is not None:
        selected = [item for item in selected if item["label"] == args.label]
        require(selected, f"unknown child label: {args.label}")
    errors: list[str] = []
    for spec in selected:
        try:
            run_child(m, spec)
        except EvidenceError as error:
            errors.append(str(error))
            print(f"capture failed: {error}", file=sys.stderr)
    if errors:
        fail("one or more child legs failed; retained all attempted artifacts")


def integer(value: object, label: str, positive: bool = False) -> int:
    require(type(value) is int, f"{label} must be integer")
    require(value > 0 if positive else value >= 0, f"{label} out of range")
    return value


def nearest_rank(values: list[float], fraction: float) -> float:
    ordered = sorted(values)
    index = min(max(math.ceil(fraction * len(ordered)) - 1, 0), len(ordered) - 1)
    return ordered[index]


def bootstrap(ratios: list[float], draws: int, seed: int) -> dict[str, object]:
    require(ratios, "empty paired ratio vector")
    generator = random.Random(seed)
    values = [
        statistics.median(ratios[generator.randrange(len(ratios))] for _ in ratios)
        for _ in range(draws)
    ]
    return {
        "draws": draws,
        "seed": seed,
        "statistic": "median of per-pair baseline_p50/candidate_p50 ratios",
        "ratios": ratios,
        "median_ratio": statistics.median(ratios),
        "confidence_interval_95": {
            "method": "nearest-rank 2.5th/97.5th percentiles",
            "lower": nearest_rank(values, 0.025),
            "upper": nearest_rank(values, 0.975),
        },
    }


def load_child(m: dict[str, Any], spec: dict[str, object]) -> tuple[dict[str, Any], dict[str, Any], dict[str, Any]]:
    label = str(spec["label"])
    receipt_value = read_json(receipt_path(label))
    require(isinstance(receipt_value, dict), f"{label}: receipt malformed")
    receipt = receipt_value
    require(receipt.get("status") == "passed" and receipt.get("exit_code") == 0, f"{label}: child failure retained")
    require(receipt.get("pair") == spec["pair"] and receipt.get("leg") == spec["leg"] and receipt.get("variant") == spec["variant"], f"{label}: schedule binding mismatch")
    started_at = parse_utc(receipt.get("started_utc"), f"{label} start")
    finished_at = parse_utc(receipt.get("finished_utc"), f"{label} finish")
    require(started_at <= finished_at, f"{label}: finish precedes start")
    require(
        parse_utc(m["created_utc"], "manifest created") <= started_at,
        f"{label}: started before manifest freeze",
    )
    require_previous_children_finished(m, spec, str(receipt["started_utc"]))
    manifest_identity = receipt.get("manifest")
    require(isinstance(manifest_identity, dict), f"{label}: manifest binding missing")
    verify_recorded(manifest_identity, MANIFEST_PATH, f"{label}: manifest")
    require(receipt.get("protocol") == m["protocol"] and receipt.get("script") == m["script"] and receipt.get("validator") == m["validator"], f"{label}: tool binding mismatch")
    binary = m[str(spec["variant"])]
    require(receipt.get("binary") == binary, f"{label}: binary binding mismatch")
    for key in ("process", "stdout", "stderr", "summary"):
        identity = receipt.get(key)
        require(isinstance(identity, dict), f"{label}: committed {key} descriptor missing")
        verify_recorded(identity, recorded_path(identity, m), f"{label}: committed {key}")
    compressed = receipt.get("report_gzip")
    require(isinstance(compressed, dict), f"{label}: compressed report descriptor missing")
    compressed_path = recorded_path(compressed, m)
    verify_recorded(compressed, compressed_path, f"{label}: compressed report")
    raw = read_gzip(compressed_path)
    raw_identity = receipt.get("report_raw")
    require(isinstance(raw_identity, dict) and len(raw) == raw_identity.get("bytes") and sha256_bytes(raw) == raw_identity.get("sha256"), f"{label}: raw report identity mismatch")
    report_value = json.loads(raw.decode("utf-8"))
    require(isinstance(report_value, dict), f"{label}: raw report is not an object")
    validator = validator_module()
    summary = validator.validate_report(report_value, label, binary)
    require(receipt.get("validation") == summary, f"{label}: receipt validation summary changed")
    process_identity = receipt["process"]
    process_path = recorded_path(process_identity, m)
    process_metrics = receipt.get("process_metrics")
    require(isinstance(process_metrics, dict), f"{label}: process metrics missing")
    actual_process_metrics = validator_module().parse_time_v(process_path)
    for key, value in actual_process_metrics.items():
        if key != "path":
            require(process_metrics.get(key) == value, f"{label}: recorded process metric {key} changed")
    require(
        set(process_metrics) - {"path"} == set(actual_process_metrics) - {"path"},
        f"{label}: recorded process metric fields changed",
    )
    summary_identity = receipt["summary"]
    compact = read_json(recorded_path(summary_identity, m))
    require(
        isinstance(compact, dict)
        and compact.get("validation") == summary
        and compact.get("report_gzip") == compressed,
        f"{label}: committed summary mismatch",
    )
    require("max_rss_kib" in process_metrics, f"{label}: process metrics missing max RSS")
    return receipt, report_value, summary


def increase_flag(baseline: dict[str, object], candidate: dict[str, object], key: str, limit: float) -> dict[str, object]:
    before = float(baseline[key])
    after = float(candidate[key])
    fraction = after / before - 1.0 if before else float("inf")
    return {"baseline": before, "candidate": after, "increase_fraction": fraction, "flag": fraction > limit}


def compute_reduction(m: dict[str, Any]) -> dict[str, object]:
    specs = child_specs(m)
    details: list[dict[str, object]] = []
    by_pair: dict[int, dict[str, tuple[dict[str, Any], dict[str, Any], dict[str, Any]]]] = {}
    pooled: dict[str, list[int]] = {"baseline": [], "candidate": []}
    pooled_phases: dict[str, dict[str, list[int]]] = {
        variant: {key: [] for key in ("preparation_ns", "open_ns", "query_ns", "edit_ns", "commit_ns", "publication_ns")}
        for variant in ("baseline", "candidate")
    }
    rss: dict[str, list[int]] = {"baseline": [], "candidate": []}
    for spec in specs:
        receipt, report, summary = load_child(m, spec)
        variant = str(spec["variant"])
        result = report["results"][0]
        pooled[variant].extend(int(value) for value in result["elapsed_ns"]["samples"])
        section = result["source"]["docx_section_layout"]
        for key in pooled_phases[variant]:
            pooled_phases[variant][key].extend(int(value) for value in section[key])
        rss[variant].append(int(receipt["process_metrics"]["max_rss_kib"]))
        detail = {
            "label": spec["label"],
            "pair": spec["pair"],
            "leg": spec["leg"],
            "variant": variant,
            "receipt": recorded_descriptor(receipt_path(str(spec["label"])), m),
            "report_gzip": receipt["report_gzip"],
            "report_raw": receipt["report_raw"],
            "process": receipt["process"],
            "process_metrics": receipt["process_metrics"],
            "validation": summary,
        }
        details.append(detail)
        by_pair.setdefault(int(spec["pair"]), {})[variant] = (receipt, report, summary)
    pairs: list[dict[str, object]] = []
    ratios: list[float] = []
    for pair in range(1, int(m["protocol_fields"]["pairs"]) + 1):
        group = by_pair.get(pair, {})
        require(set(group) == {"baseline", "candidate"}, f"pair {pair} is incomplete")
        baseline_summary = group["baseline"][2]
        candidate_summary = group["candidate"][2]
        baseline_p50 = baseline_summary["elapsed_ns"]["p50"]
        candidate_p50 = candidate_summary["elapsed_ns"]["p50"]
        ratio = baseline_p50 / candidate_p50
        ratios.append(ratio)
        pairs.append({
            "pair": pair,
            "baseline_label": next(item["label"] for item in details if item["pair"] == pair and item["variant"] == "baseline"),
            "candidate_label": next(item["label"] for item in details if item["pair"] == pair and item["variant"] == "candidate"),
            "baseline_p50_ns": baseline_p50,
            "candidate_p50_ns": candidate_p50,
            "baseline_over_candidate_p50_ratio": ratio,
        })
    stats = validator_module().stats
    pooled_stats = {variant: stats(values) for variant, values in pooled.items()}
    pooled_phase_stats = {
        variant: {key: stats(values) for key, values in phase_values.items()}
        for variant, phase_values in pooled_phases.items()
    }
    bootstrap_result = bootstrap(
        ratios,
        int(m["protocol_fields"]["bootstrap"]["draws"]),
        int(m["protocol_fields"]["bootstrap"]["seed"]),
    )
    lower = bootstrap_result["confidence_interval_95"]["lower"]
    median_ratio = bootstrap_result["median_ratio"]
    gate = m["protocol_fields"]["timing_gate"]
    timing_eligible = median_ratio >= gate["minimum_median_ratio"] and lower > gate["lower_95_percent_bound_must_exceed"]
    latency_flags = {
        key: increase_flag(pooled_stats["baseline"], pooled_stats["candidate"], key, m["protocol_fields"]["review_flags"]["pooled_latency_increase_fraction"])
        for key in m["protocol_fields"]["review_flags"]["latency_statistics"]
    }
    rss_stats = {variant: stats(values) for variant, values in rss.items()}
    rss_flag = increase_flag(
        {"p50": rss_stats["baseline"]["p50"]},
        {"p50": rss_stats["candidate"]["p50"]},
        "p50",
        m["protocol_fields"]["review_flags"]["whole_process_max_rss_median_increase_fraction"],
    )
    return {
        "schema": "litchi.performance.0915.native-compare-reduction.v1",
        "source_revision": m["protocol_fields"]["baseline_source_revision"],
        "protocol": m["protocol"],
        "script": m["script"],
        "validator": m["validator"],
        "manifest": recorded_manifest_descriptor(m),
        "case": m["protocol_fields"]["case"],
        "baseline": m["baseline"],
        "candidate": m["candidate"],
        "pairs": len(pairs),
        "children": len(details),
        "warmups": m["protocol_fields"]["warmups"],
        "samples_per_child": m["protocol_fields"]["samples_per_child"],
        "cpu": m["protocol_fields"]["cpu"],
        "timing_scope": "harness elapsed_ns: preparation, source-backed open, section query, edit, commit, sequential publication; exact six-phase sum",
        "process_scope": "GNU time -v: setup-inclusive child process including corpus generation, benchmark, verification, JSON serialization, and process teardown",
        "physical_io_metrics": "unavailable; source is in-memory InstrumentedSource and counters are logical harness observations",
        "pooled_elapsed_ns": pooled_stats,
        "pooled_phase_ns": pooled_phase_stats,
        "whole_process_max_rss_kib": rss_stats,
        "pairs_detail": pairs,
        "children_detail": details,
        "paired_bootstrap": bootstrap_result,
        "timing_gate": {
            "minimum_median_ratio": gate["minimum_median_ratio"],
            "lower_95_percent_bound_must_exceed": gate["lower_95_percent_bound_must_exceed"],
            "timing_eligible": timing_eligible,
        },
        "review_flags": {"latency": latency_flags, "whole_process_max_rss": rss_flag},
        "adoption": m["protocol_fields"].get("adoption", "timing eligibility does not adopt the candidate") if isinstance(m["protocol_fields"], dict) else "timing eligibility does not adopt the candidate",
        "performance_claim": "timing eligibility only; no production adoption",
    }


def reduce(args: argparse.Namespace) -> None:
    m = manifest()
    result = compute_reduction(m)
    output = args.output.resolve() if args.output else HERE / "reduced.json"
    write_json(output, result)
    print(json.dumps({"status": "reduced", "output": descriptor(output), "timing_gate": result["timing_gate"]}))


def replay() -> None:
    m = manifest()
    retained = read_json(HERE / "reduced.json")
    require(isinstance(retained, dict), "retained reduction malformed")
    recomputed = compute_reduction(m)
    require(recomputed == retained, "retained reduction differs from deterministic recomputation")
    print(json.dumps({"status": "replayed", "reduced": descriptor(HERE / "reduced.json")}))


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="action", required=True)
    freeze_parser = subparsers.add_parser("freeze")
    freeze_parser.add_argument("--candidate", required=True, type=Path)
    freeze_parser.add_argument("--baseline", type=Path)
    freeze_parser.add_argument("--source-receipt", type=Path)
    freeze_parser.add_argument("--build-receipt", type=Path)
    run_parser = subparsers.add_parser("run")
    run_parser.add_argument("--pair", type=int)
    run_parser.add_argument("--label")
    reduce_parser = subparsers.add_parser("reduce")
    reduce_parser.add_argument("--output", type=Path)
    subparsers.add_parser("replay")
    args = parser.parse_args()
    try:
        if args.action == "freeze":
            freeze(args)
        elif args.action == "run":
            run(args)
        elif args.action == "reduce":
            reduce(args)
        else:
            replay()
    except (EvidenceError, OSError, subprocess.SubprocessError, ValueError) as error:
        print(f"native comparison failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
