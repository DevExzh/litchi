#!/usr/bin/env python3
"""Replay the complete 0429 managed-cache evidence bundle.

The replay is independent of the captured executable. It checks the frozen
32-process order, two supplementary profiles, report and corpus custody,
build/source identity, lossless logs, summary derivation, and portable validator
mutations. API-only duration baselines make no causal optimization claim.
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
CHANGE = 429
SAMPLES = 30
WARMUPS = 3
PROCESSES = 32
RETAINED_SAMPLES = 960
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
    lanes = []
    for command, key, selectors in [("provider-lifecycle", "corpus", ["plain", "media-rich"]),
                                     ("native-image-lifecycle", "fixture", ["poi-slide", "poi-video"])]:
        for selector in selectors:
            for label, provider, cap, delay in [("bytes", "bytes", None, None), ("file", "file", None, None),
                                                ("short-range", "range", 4096, 0), ("delayed-range", "range", 65536, 200)]:
                lanes.append(dict(command=command, selector_key=key, selector=selector,
                                  provider_label=label, provider=provider, max_range=cap, delay_us=delay))
    assert protocol["lanes"] == lanes
    expected = [dict(row, repeat="R1") for row in lanes] + [dict(row, repeat="R2") for row in reversed(lanes)]
    assert protocol["order"] == expected


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
        ("replay_verifier_sha256", "verify.py"),
    ):
        assert build[field] == sha((root / expected).read_bytes()), field

    for name, digest in build["additional_artifact_sha256"].items():
        assert sha(safe_path(root, name).read_bytes()) == digest, name
    assert set(build["additional_artifact_sha256"]) == {"check.py", "profiles.py", "profile-audit.py", "native-oracles.py", "native-image-oracles.json", "native-inputs/original.pptx", "native-inputs/libreoffice.pptx", "native-inputs/poi-slide.pptx", "native-inputs/poi-video.pptx", "shapes-static-oracles.json"}
    run_quiet([sys.executable, "-B", str(root / "native-oracles.py"), "--check"])
    run_quiet([sys.executable, "-B", str(root / "native-oracles.py"), "--check", "--shapes"])

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
    assert "harness-strict-final" in required_review

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
    assert debt.get("new_module_findings") == 0
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
    argv = row["argv"]
    assert argv[:6] == ["taskset", "-c", str(protocol["cpu"]), "/usr/bin/time", "-v", "-o"]
    assert Path(argv[6]).name == row["name"] + "-resource.log"
    expected = [row["command"], "--" + row["selector_key"], row["selector"], "--provider", row["provider"],
                "--samples", str(SAMPLES), "--warmup", str(WARMUPS), "--source-revision", revision, "--output", argv[20]]
    if row["provider"] == "range":
        expected += ["--max-range", str(row["max_range"]), "--delay-us", str(row["delay_us"])]
    assert argv[8:] == expected
    assert Path(argv[20]).name == row["name"] + ".json"


def verify_report_role(report, role):
    assert report["provider"] == role["provider"]
    config = report["provider_config"]
    assert config["provider"] == role["provider"]
    assert config["max_range_bytes"] == role["max_range"]
    assert config["delay_us"] == role["delay_us"]
    if role["command"] == "provider-lifecycle":
        assert report["schema"] == "pptx_provider_lifecycle_v1"
        assert report["corpus"] == role["selector"]
    else:
        assert report["schema"] == "pptx_native_image_lifecycle_v1"
        assert report["fixture"] == role["selector"]


def report_identity(report, role):
    if role["command"] == "provider-lifecycle":
        fields = ["source_archive_sha256", "source_archive_bytes", "destination_archive_sha256",
                  "destination_archive_bytes", "expected_output_sha256", "expected_output_bytes"]
        return tuple(report[key] for key in fields)
    return (report["native_fixture_sha256"], report["native_fixture_bytes"],
            report["payload_oracle"]["payload_sha256"], report["payload_oracle"]["payload_bytes"])


def verify_capture_index(root: Path, protocol: dict[str, Any], build: dict[str, Any]) -> tuple[int, int]:
    index = load_json(root / "capture-index.json")
    assert isinstance(index, list) and len(index) == PROCESSES and len(set(index)) == PROCESSES
    identities = {}
    samples = 0
    for receipt_name, role in zip(index, protocol["order"], strict=True):
        row = load_json(safe_path(root, receipt_name))
        name = role["command"] + "-" + role["selector"] + "-" + role["provider_label"] + "-" + role["repeat"].lower()
        assert row["name"] == name
        assert all(row[key] == value for key, value in role.items())
        assert row["status"] == "pass" and row["exit_code"] == 0
        assert row["revision"] == build["revision"]
        assert row["binary_sha256"] == build["binary_sha256"]
        assert row["protocol_sha256"] == build["protocol_sha256"]
        verify_argv(row, protocol, build["revision"])
        assert row["argv"][7] == build["capture_binary"]
        assert set(row["artifacts"]) == {f"capture/{name}.json", f"capture/{name}.log", f"capture/{name}-resource.log"}
        for path, custody in row["artifacts"].items():
            raw = artifact(root, path)
            assert sha(raw) == custody["sha256"] and len(raw) == custody["bytes"]
        report_path = root / "capture" / (name + ".json")
        run_quiet([sys.executable, "-B", str(root / "verify-report.py"), str(report_path)])
        report = load_json(report_path)
        for key in ["source_revision", "binary_sha256", "binary_bytes", "current_exe"]:
            expected = build[{"source_revision": "revision", "current_exe": "capture_binary"}.get(key, key)]
            assert report[key] == expected
        assert report["samples"] == SAMPLES and report["warmup"] == WARMUPS
        # Report schema/provider/fixture bindings are also independently checked below.
        verify_report_role(report, role)
        rows = report["samples_raw"]
        samples += len(rows)
        for sample in rows:
            for phase in sample["phases"]:
                assert phase["rss"]["availability"] == "available", "formal Linux RSS unavailable"
        identity = report_identity(report, role)
        key = (role["command"], role["selector"])
        assert identities.setdefault(key, identity) == identity, "same-role output or corpus differs"
        run_quiet([sys.executable, "-B", str(root / "probe-report.py"), str(report_path)])
    assert samples == RETAINED_SAMPLES
    return len(index), samples


def verify_profiles(root, protocol, build):
    index = load_json(root / "profile-index.json")
    plan = protocol["supplementary_profiles"]
    assert len(index) == len(plan["cases"]) == 2 and len(set(index)) == 2
    for path, case in zip(index, plan["cases"], strict=True):
        receipt = load_json(safe_path(root, path))
        name = case["corpus"] + "-" + case["provider"]
        assert receipt["name"] == name
        assert receipt["status"] == "pass" and receipt["exit_code"] == 0
        assert receipt["revision"] == build["revision"] and receipt["binary_sha256"] == build["binary_sha256"]
        assert receipt["scope"] == plan["scope"]
        argv = receipt["argv"]
        expected = ["taskset", "-c", str(protocol["cpu"]), "perf", "record", "-e", plan["event"],
                    "-F", str(plan["frequency_hz"]), "--call-graph", plan["call_graph"], "-o", argv[12],
                    "--", build["capture_binary"], case["command"], "--corpus", case["corpus"],
                    "--provider", case["provider"], "--samples", str(plan["samples"]), "--warmup", str(plan["warmups"]),
                    "--source-revision", build["revision"], "--output", argv[27]]
        assert argv == expected
        assert Path(argv[12]).name == name + ".data" and Path(argv[27]).name == name + ".json"
        assert set(receipt["artifacts"]) == {"profiles/" + name + suffix for suffix in [".data", ".json", ".log", "-script.log", "-report.log"]}
        for item, custody in receipt["artifacts"].items():
            raw = artifact(root, item)
            assert sha(raw) == custody["sha256"] and len(raw) == custody["bytes"]
        report_path = root / "profiles" / (name + ".json")
        run_quiet([sys.executable, "-B", str(root / "verify-report.py"), str(report_path)])
        report = load_json(report_path)
        assert report["samples"] == plan["samples"] and report["warmup"] == plan["warmups"]
        assert report["source_revision"] == build["revision"] and report["binary_sha256"] == build["binary_sha256"]
        assert report["binary_bytes"] == build["binary_bytes"] and report["current_exe"] == build["capture_binary"]
        assert report["corpus"] == case["corpus"] and report["provider"] == case["provider"]
    return len(index)


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
    profile_processes = verify_profiles(root, protocol, build)
    run_quiet([sys.executable, "-B", str(root / "profile-audit.py"), "--check"])
    cleanup_state = verify_cleanup(root, build)
    inventory_files = verify_compression_and_inventory(root)
    return {
        "status": "pass",
        "command_receipts": command_receipts,
        "required_pass": required_pass,
        "required_review": required_review,
        "processes": processes,
        "supplementary_profiles": profile_processes,
        "samples": samples,
        "inventory_files": inventory_files,
        "cleanup": cleanup_state,
        "performance_claim": None,
    }


def portable_mutations(root: Path, result: dict[str, Any]) -> dict[str, str]:
    with tempfile.TemporaryDirectory(prefix="litchi-0429-replay-") as temporary:
        exported = Path(temporary) / "bundle"
        shutil.copytree(root, exported)
        replay = load_json_bytes(
            subprocess.check_output([sys.executable, "-B", str(exported / "verify.py")])
        )
        assert replay == result

        # Keep the report syntactically valid and update its receipt custody;
        # only the same-role R1/R2 identity should reject it.
        report = exported / "capture/provider-lifecycle-plain-bytes-r2.json"
        receipt_path = exported / "capture/provider-lifecycle-plain-bytes-r2-receipt.json"
        original_report = report.read_bytes()
        original_receipt = receipt_path.read_bytes()
        changed = require_object(load_json(report), "portable report")
        digest = changed["expected_output_sha256"]
        changed["expected_output_sha256"] = "0" * 64 if digest != "0" * 64 else "1" * 64
        for sample in changed["samples_raw"]:
            sample["output_sha256"] = changed["expected_output_sha256"]
        report.write_text(json.dumps(changed, indent=2) + "\n")
        custody = require_object(load_json(receipt_path), "portable receipt")
        artifacts = require_object(custody["artifacts"], "portable artifacts")
        artifacts["capture/provider-lifecycle-plain-bytes-r2.json"] = {
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
