#!/usr/bin/env python3
"""Replay the complete 0428 managed-cache evidence bundle.

The replay is intentionally independent of the captured executable. It
checks the frozen 16-process order, report and corpus custody, build/source
identity, compressed-log custody, summary replay, and portable validator
mutation. It makes no timing or optimization claim.
"""

from __future__ import annotations

import argparse
import gzip
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tempfile
from typing import Any


ROOT = Path(__file__).resolve().parent
CHANGE = 428
SAMPLES = 30
WARMUPS = 3
PROCESSES = 16
RETAINED_SAMPLES = 480
FULL_SOURCE_SCOPE = "workspace and standalone tools"


def sha(raw: bytes) -> str:
    return hashlib.sha256(raw).hexdigest()


def _duplicate_key(pairs: list[tuple[str, Any]]) -> dict[str, Any]:
    result: dict[str, Any] = {}
    for key, value in pairs:
        if key in result:
            raise ValueError(f"duplicate JSON key: {key}")
        result[key] = value
    return result


def load_json_bytes(raw: bytes) -> Any:
    def reject_constant(value: str) -> Any:
        raise ValueError(f"non-finite JSON constant: {value}")

    return json.loads(
        raw.decode("utf-8"),
        object_pairs_hook=_duplicate_key,
        parse_constant=reject_constant,
    )


def load_json(path: Path) -> Any:
    return load_json_bytes(path.read_bytes())


def run_quiet(argv: list[str]) -> subprocess.CompletedProcess[bytes]:
    """Run a nested validator without contaminating this replay's JSON output."""
    result = subprocess.run(
        argv,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
    )
    if result.returncode:
        detail = (result.stderr or result.stdout).decode("utf-8", "replace").strip()
        raise AssertionError(detail or f"nested command failed: {argv[0]}")
    return result


def safe_path(root: Path, name: str) -> Path:
    candidate = Path(name)
    if candidate.is_absolute():
        raise AssertionError(f"absolute bundle path: {name}")
    resolved = (root / candidate).resolve()
    if not resolved.is_relative_to(root.resolve()):
        raise AssertionError(f"bundle path escapes root: {name}")
    return resolved


def artifact(root: Path, name: str) -> bytes:
    """Read a retained artifact, accepting only the sealer's .gz fallback."""
    path = safe_path(root, name)
    if path.is_file():
        return path.read_bytes()
    compressed = Path(str(path) + ".gz")
    if compressed.is_file():
        return gzip.decompress(compressed.read_bytes())
    raise AssertionError(f"missing artifact: {name}")


def require_object(value: Any, label: str) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise AssertionError(f"{label} is not an object")
    return value


def require_text(value: Any, label: str) -> str:
    if not isinstance(value, str) or not value:
        raise AssertionError(f"{label} is not nonempty text")
    return value


def require_uint(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise AssertionError(f"{label} is not an unsigned integer")
    return value


def verify_protocol(protocol: dict[str, Any]) -> None:
    assert protocol["change"] == CHANGE
    assert protocol["samples_per_process"] == SAMPLES
    assert protocol["warmups_per_process"] == WARMUPS
    assert protocol["expected_processes"] == PROCESSES
    assert protocol["expected_retained_samples"] == RETAINED_SAMPLES
    assert protocol["cpu"] == 2 and protocol["workers"] == 1

    expected_lanes = {
        ("lifecycle", "plain"),
        ("lifecycle", "media-rich"),
        ("exact-admission", "media-rich"),
        ("one-under", "media-rich"),
        ("pinned-eviction", "media-rich"),
        ("oversized-bypass", "media-rich"),
        ("repeated-publication", "plain"),
        ("repeated-publication", "media-rich"),
    }
    assert len(protocol["lanes"]) == len(expected_lanes)
    lanes = {
        (row.get("scenario"), row.get("corpus"))
        for row in protocol["lanes"]
    }
    assert lanes == expected_lanes
    order = protocol["order"]
    assert len(order) == PROCESSES
    first = [(row["scenario"], row["corpus"]) for row in order[:8]]
    second = [(row["scenario"], row["corpus"]) for row in order[8:]]
    assert all(row["repeat"] == "R1" for row in order[:8])
    assert all(row["repeat"] == "R2" for row in order[8:])
    assert second == list(reversed(first))
    assert set(first) == expected_lanes


def verify_source_manifest(root: Path, manifest: dict[str, Any]) -> None:
    path = require_text(manifest.get("path"), "source manifest path")
    digest = require_text(manifest.get("sha256"), "source manifest digest")
    files = require_uint(manifest.get("files"), "source manifest file count")
    raw = artifact(root, path)
    assert sha(raw) == digest
    parsed = load_json_bytes(raw)
    assert isinstance(parsed, dict) and len(parsed) == files


def verify_build(root: Path, protocol: dict[str, Any], build: dict[str, Any]) -> dict[str, Any]:
    assert build.get("baseline_revision") == protocol.get("baseline_revision")
    build_revision = require_text(build.get("revision"), "build revision")
    binary_digest = require_text(build.get("binary_sha256"), "build binary digest")
    binary_bytes = require_uint(build.get("binary_bytes"), "build binary bytes")
    assert len(binary_digest) == 64

    for field, expected in (
        ("protocol_sha256", "protocol.json"),
        ("planned_checks_sha256", "planned-checks.json"),
        ("machine_sha256", "machine.json"),
        ("strict_comparison_driver_sha256", "compare-strict.py"),
        ("capture_driver_sha256", "capture.py"),
        ("verifier_sha256", "verify-report.py"),
        ("probe_sha256", "probe-report.py"),
        ("summary_driver_sha256", "summarize.py"),
    ):
        assert build[field] == sha((root / expected).read_bytes()), field

    receipt_name = require_text(build.get("build_receipt"), "build receipt")
    receipt_raw = artifact(root, receipt_name)
    assert sha(receipt_raw) == build["build_receipt_sha256"]
    receipt = require_object(load_json_bytes(receipt_raw), "build receipt")
    assert receipt.get("status") == "pass"
    assert receipt.get("revision") == build_revision
    assert receipt.get("source_scope") == FULL_SOURCE_SCOPE
    assert receipt.get("source_before") == receipt.get("source_after")
    assert receipt.get("source_after") == build.get("source_manifest")
    verify_source_manifest(root, require_object(receipt["source_after"], "build source"))
    verify_source_manifest(root, require_object(build["source_manifest"], "bound source"))
    return {"revision": build_revision, "binary_sha256": binary_digest,
            "binary_bytes": binary_bytes}


def verify_planned_checks(root: Path) -> tuple[int, int]:
    planned = require_object(load_json(root / "planned-checks.json"), "planned checks")
    required_pass = planned.get("required_pass")
    required_review = planned.get("required_review")
    assert isinstance(required_pass, list) and all(isinstance(name, str) and name for name in required_pass)
    assert isinstance(required_review, list) and all(isinstance(name, str) and name for name in required_review)
    assert len(set(required_pass)) == len(required_pass)
    assert len(set(required_review)) == len(required_review)
    assert not set(required_pass) & set(required_review)
    assert "harness-strict" in required_review

    for name in required_pass:
        row = require_object(
            load_json(safe_path(root, f"checks/{name}.json")),
            f"required pass {name}",
        )
        assert row.get("status") == "pass", name
    for name in required_review:
        row = require_object(
            load_json(safe_path(root, f"checks/{name}.json")),
            f"required review {name}",
        )
        assert row.get("status") in {"pass", "failed", "review"}, name

    debt = require_object(
        load_json_bytes(artifact(root, "checks/strict-debt-comparison.json")),
        "strict debt comparison",
    )
    assert debt.get("status") == "pass"
    assert debt.get("same_message_and_source_file_multiset") is True
    assert debt.get("cache_retention_module_findings") == 0
    return len(required_pass), len(required_review)


def verify_check_receipts(root: Path) -> int:
    expected_path = root / "expected-checks.json"
    if not expected_path.is_file():
        raise AssertionError("expected-checks.json is missing")
    expected = require_object(load_json(expected_path), "expected checks")
    driver_digest = sha((root / "check.py").read_bytes())
    actual: set[str] = set()
    for name, expected_status in expected.items():
        row_path = safe_path(root, f"checks/{name}.json")
        row = require_object(load_json(row_path), f"check receipt {name}")
        assert row.get("status") == expected_status, name
        assert row.get("change") == CHANGE, name
        assert row.get("driver_sha256") == driver_digest, name
        assert row.get("source_unchanged") is True, name
        assert row.get("source_before") == row.get("source_after"), name
        source = require_object(row.get("source_after"), f"{name} source")
        verify_source_manifest(root, source)
        log = require_object(row.get("log"), f"{name} log")
        raw = artifact(root, require_text(log.get("path"), f"{name} log path"))
        assert sha(raw) == log.get("sha256") and len(raw) == log.get("bytes"), name
        scope = row.get("source_scope")
        assert scope in {"workspace excluding tools", "workspace and standalone tools"}, name
        actual.add(name)
    observed = set()
    for path in (root / "checks").glob("*.json"):
        value = load_json(path)
        if isinstance(value, dict) and "source_before" in value:
            observed.add(path.stem)
    assert observed == actual
    return len(actual)


def verify_argv(row: dict[str, Any], protocol: dict[str, Any], revision: str) -> None:
    argv = row.get("argv")
    assert isinstance(argv, list) and all(isinstance(item, str) for item in argv)
    expected = [
        "taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o",
    ]
    assert argv[: len(expected)] == expected
    assert len(argv) == 21
    name = row["name"]
    assert Path(argv[6]).name == f"{name}-resource.log"
    tail = [
        "cache-retention", "--scenario", row["scenario"], "--corpus", row["corpus"],
        "--samples", str(SAMPLES), "--warmup", str(WARMUPS),
        "--source-revision", revision, "--output",
    ]
    assert argv[8:20] == tail
    assert Path(argv[20]).name == f"{name}.json"
    assert len({flag for flag in argv if flag.startswith("--")}) == 6


def report_identity(report: dict[str, Any], key: str) -> tuple[Any, ...]:
    if key == "corpus":
        names = ("source_archive_sha256", "source_archive_bytes",
                 "destination_archive_sha256", "destination_archive_bytes")
    else:
        names = ("expected_output_sha256", "expected_output_bytes")
    values = tuple(report.get(name) for name in names)
    assert all(value is not None for value in values), f"missing {key} identity"
    return values


def verify_capture_index(root: Path, protocol: dict[str, Any], build: dict[str, Any]) -> tuple[int, int]:
    index_value = load_json(root / "capture-index.json")
    assert isinstance(index_value, list)
    assert len(index_value) == PROCESSES
    expected_order = [
        (row["scenario"], row["corpus"], row["repeat"])
        for row in protocol["order"]
    ]
    identities: dict[str, tuple[Any, ...]] = {}
    outputs: dict[tuple[str, str], tuple[Any, ...]] = {}
    seen_receipts: set[str] = set()
    sample_count = 0

    for receipt_name, expected in zip(index_value, expected_order, strict=True):
        receipt_path = safe_path(root, require_text(receipt_name, "capture receipt path"))
        assert receipt_name not in seen_receipts
        seen_receipts.add(receipt_name)
        row = require_object(load_json(receipt_path), f"capture receipt {receipt_name}")
        scenario, corpus, repeat = expected
        name = f"{scenario}-{corpus}-{repeat.lower()}"
        assert row.get("name") == name
        assert (row.get("scenario"), row.get("corpus"), row.get("repeat")) == expected
        assert row.get("status") == "pass" and row.get("exit_code") == 0
        assert row.get("revision") == build["revision"]
        assert row.get("binary_sha256") == build["binary_sha256"]
        assert row.get("protocol_sha256") == build["protocol_sha256"]
        verify_argv(row, protocol, build["revision"])
        assert row["argv"][7] == build["capture_binary"]

        artifacts = require_object(row.get("artifacts"), f"{name} artifacts")
        expected_artifacts = {
            f"capture/{name}.json",
            f"capture/{name}.log",
            f"capture/{name}-resource.log",
        }
        assert set(artifacts) == expected_artifacts
        for artifact_name, custody in artifacts.items():
            custody = require_object(custody, f"{name} artifact custody")
            raw = artifact(root, artifact_name)
            assert sha(raw) == custody.get("sha256")
            assert len(raw) == custody.get("bytes")

        report_path = root / "capture" / f"{name}.json"
        assert report_path.is_file()
        run_quiet(
            [sys.executable, "-B", str(root / "verify-report.py"), str(report_path)],
        )
        report = require_object(load_json(report_path), f"{name} report")
        assert report.get("source_revision") == build["revision"]
        assert report.get("binary_sha256") == build["binary_sha256"]
        assert report.get("binary_bytes") == build["binary_bytes"]
        assert report.get("current_exe") == build["capture_binary"]
        assert report.get("scenario") == scenario
        assert report.get("corpus") == corpus
        assert report.get("samples") == SAMPLES
        assert report.get("warmup") == WARMUPS
        rows = report.get("samples_raw")
        assert isinstance(rows, list) and len(rows) == SAMPLES
        machine = require_object(load_json(root / "machine.json"), "machine")
        if str(machine.get("platform", "")).startswith("Linux-"):
            for sample in rows:
                for phase in sample["phases"]:
                    assert phase["rss"]["availability"] == "available", "formal Linux RSS unavailable"
        sample_count += len(rows)

        corpus_identity = report_identity(report, "corpus")
        prior = identities.setdefault(corpus, corpus_identity)
        assert corpus_identity == prior, f"{name}: corpus identity differs"
        output_identity = report_identity(report, "output")
        role_key = (scenario, corpus)
        prior_output = outputs.setdefault(role_key, output_identity)
        assert output_identity == prior_output, f"{name}: same-role output differs"
        run_quiet(
            [sys.executable, "-B", str(root / "probe-report.py"), str(report_path)],
        )
    assert sample_count == RETAINED_SAMPLES
    return len(index_value), sample_count


def verify_compression_and_inventory(root: Path) -> int:
    compression = require_object(load_json(root / "compression.json"), "compression")
    for name, row_value in compression.items():
        row = require_object(row_value, f"compression {name}")
        stored = safe_path(root, name).read_bytes()
        raw = gzip.decompress(stored)
        assert sha(stored) == row.get("stored_sha256")
        assert len(stored) == row.get("stored_bytes")
        assert sha(raw) == row.get("original_sha256")
        assert len(raw) == row.get("original_bytes")
        assert row.get("original_path") == name.removesuffix(".gz")

    inventory: dict[str, str] = {}
    for line in (root / "SHA256SUMS").read_text().splitlines():
        digest, name = line.split("  ", 1)
        assert name not in inventory
        safe_path(root, name)
        inventory[name] = digest
        assert sha((root / name).read_bytes()) == digest, name
    expected = {
        str(path.relative_to(root))
        for path in root.rglob("*")
        if path.is_file() and path.name != "SHA256SUMS"
    }
    assert set(inventory) == expected
    return len(inventory)


def verify_cleanup(root: Path, build: dict[str, Any]) -> str:
    """Validate cleanup metadata without requiring a temporary executable.

    A live capture checkout may be replayed before cleanup, in which case the
    copied executable is checked when it is still present.  A published bundle
    normally contains cleanup.json after that executable has been removed; the
    cleanup record is then bound to build.json's copied-binary digest and byte
    count rather than to the removed absolute path.
    """
    cleanup_path = root / "cleanup.json"
    if cleanup_path.is_file():
        cleanup = require_object(load_json(cleanup_path), "cleanup")
        assert cleanup.get("status") == "pass"
        capture_binary = Path(require_text(build.get("capture_binary"), "capture binary"))
        assert cleanup.get("removed_directory") == str(capture_binary.parent)
        assert cleanup.get("removed_binary_sha256") == build["binary_sha256"]
        assert cleanup.get("temporary_directory_absent") is True
        assert cleanup.get("original_build_binary_retained") is True
        assert cleanup.get("root_target_retained") is True
        assert cleanup.get("harness_target_retained") is True
        return "after-cleanup"

    capture_binary = Path(require_text(build.get("capture_binary"), "capture binary"))
    if capture_binary.is_file():
        raw = capture_binary.read_bytes()
        assert sha(raw) == build["binary_sha256"]
        assert len(raw) == build["binary_bytes"]
        return "before-cleanup; copied binary checked"
    # An exported pre-cleanup bundle may not include the host's absolute /tmp
    # path.  The report and build identities still bind it; do not make that
    # portable export depend on a host-local executable.
    return "before-cleanup; copied binary path unavailable in export"


def verify(root: Path) -> dict[str, Any]:
    protocol = require_object(load_json(root / "protocol.json"), "protocol")
    verify_protocol(protocol)
    build = require_object(load_json(root / "build.json"), "build")
    identity = verify_build(root, protocol, build)
    required_pass, required_review = verify_planned_checks(root)
    capture_check = require_object(load_json(root / "checks/release-capture.json"),
                                   "release capture check")
    assert capture_check.get("status") == "pass"
    assert capture_check.get("change") == CHANGE
    assert capture_check.get("revision") == identity["revision"]
    assert capture_check.get("source_scope") == FULL_SOURCE_SCOPE
    assert capture_check.get("source_before") == capture_check.get("source_after")
    assert capture_check.get("source_after") == build.get("source_manifest")
    verify_source_manifest(root, require_object(capture_check["source_after"], "capture source"))

    command_receipts = verify_check_receipts(root)
    run_quiet([sys.executable, "-B", str(root / "compare-strict.py"), "--check"])
    processes, samples = verify_capture_index(root, protocol, build)
    run_quiet(
        [sys.executable, "-B", str(root / "summarize.py"), "--check"],
    )
    run_quiet([sys.executable, "-B", str(root / "resource-audit.py"), "--check"])
    cleanup_state = verify_cleanup(root, build)
    inventory_files = verify_compression_and_inventory(root)
    return {
        "status": "pass",
        "command_receipts": command_receipts,
        "required_pass": required_pass,
        "required_review": required_review,
        "processes": processes,
        "samples": samples,
        "inventory_files": inventory_files,
        "cleanup": cleanup_state,
        "performance_claim": None,
    }


def portable_mutations(root: Path, result: dict[str, Any]) -> dict[str, str]:
    with tempfile.TemporaryDirectory(prefix="litchi-0428-replay-") as temporary:
        exported = Path(temporary) / "bundle"
        shutil.copytree(root, exported)
        replay = load_json_bytes(
            subprocess.check_output([sys.executable, "-B", str(exported / "verify.py")])
        )
        assert replay == result

        # Keep the report syntactically valid and update its receipt custody;
        # only the same-role R1/R2 identity should reject it.
        report = exported / "capture/lifecycle-plain-r2.json"
        receipt_path = exported / "capture/lifecycle-plain-r2-receipt.json"
        original_report = report.read_bytes()
        original_receipt = receipt_path.read_bytes()
        changed = require_object(load_json(report), "portable report")
        digest = changed["expected_output_sha256"]
        changed["expected_output_sha256"] = "0" * 64 if digest != "0" * 64 else "1" * 64
        report.write_text(json.dumps(changed, indent=2) + "\n")
        custody = require_object(load_json(receipt_path), "portable receipt")
        artifacts = require_object(custody["artifacts"], "portable artifacts")
        artifacts["capture/lifecycle-plain-r2.json"] = {
            "sha256": sha(report.read_bytes()), "bytes": report.stat().st_size,
        }
        receipt_path.write_text(json.dumps(custody, indent=2) + "\n")
        rejected = subprocess.run(
            [sys.executable, "-B", str(exported / "verify.py")],
            capture_output=True,
        )
        assert rejected.returncode != 0, "same-role output mutation was accepted"
        rejection = (rejected.stdout + rejected.stderr).decode("utf-8", "replace").upper()
        assert (
            "INVALID" in rejection or "SAME-ROLE OUTPUT DIFFERS" in rejection
        ), "same-role output mutation did not produce a meaningful rejection"
        report.write_bytes(original_report)
        receipt_path.write_bytes(original_receipt)

        validator = exported / "verify-report.py"
        validator.write_text(validator.read_text() + "\n# pinned-validator mutation\n")
        rejected = subprocess.run(
            [sys.executable, "-B", str(exported / "verify.py")],
            capture_output=True,
        )
        assert rejected.returncode != 0, "modified pinned validator was accepted"
    return {
        "portable_export": "pass; copied executable not needed",
        "pinned_validator_mutation": "rejected",
        "valid_digest_repeat_output_mutation": "rejected",
    }


def main() -> None:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--portable", action="store_true")
    args = parser.parse_args()
    result = verify(ROOT)
    if args.portable:
        result.update(portable_mutations(ROOT, result))
    print(json.dumps(result, indent=2, sort_keys=True))


if __name__ == "__main__":
    main()
