#!/usr/bin/env python3
"""Mutation tests for the 0478 evidence boundary."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import unittest

import analyze
import verify


def process() -> dict[str, int]:
    return {
        "rchar": 1, "wchar": 2, "read_bytes": 3, "write_bytes": 4,
        "cancelled_write_bytes": 0, "syscr": 5, "syscw": 6,
        "minor_faults": 7, "major_faults": 0, "user_cpu_ticks": 8,
        "system_cpu_ticks": 9, "clock_ticks_per_second": 100,
        "voluntary_context_switches": 0, "nonvoluntary_context_switches": 0,
        "rss_bytes": 10, "peak_rss_bytes": 11,
    }


def allocation() -> dict[str, int | str]:
    return {
        "status": "measured", "scope": "operation_global_system_allocator",
        "allocation_calls": 3, "deallocation_calls": 2,
        "reallocation_calls": 0, "failed_allocation_calls": 0,
        "allocated_bytes": 50, "deallocated_bytes": 50,
        "live_bytes_before": 100, "live_bytes_after": 100,
        "peak_live_bytes_before": 100, "peak_live_bytes_after": 125,
        "region_peak_live_bytes": 125,
    }


def spec(instrumentation: str = "normal", policy: str = "spool", count: int = 8) -> dict[str, object]:
    label = f"r1-{instrumentation}-{count}-{policy}"
    argv = [
        "/usr/bin/time", "-v", "-o", f"/bundle/captures/{label}.resource",
        "/usr/bin/taskset", "-c", "2", f"/tmp/{instrumentation}/pptx_metadata_spool",
        "--mode", policy, "--counts", str(count), "--samples", "30", "--warmups", "3",
        "--repeats", "1", "--max-spool-bytes", str(analyze.MAX_SPOOL_BYTES),
        "--spool-buffer-bytes", str(analyze.SPOOL_BUFFER_BYTES), "--json",
        f"/bundle/captures/{label}.report.json", "--spool-dir", f"/tmp/spools/{label}",
    ]
    return {"label": label, "instrumentation": instrumentation, "count": count,
            "policy": policy, "repeat": 1, "argv": argv}


def report_for(case: dict[str, object], *, allocator_metrics: bool | None = None) -> dict[str, object]:
    instrumentation = str(case["instrumentation"])
    policy = str(case["policy"])
    count = int(case["count"])
    output_hash = "a" * 64
    historical = analyze.HISTORICAL_CORPUS_IDENTITIES[count]
    source_hash = historical["source_archive_sha256"]
    scratch = (
        analyze.expected_spool_scratch_bytes(count)
        if policy == "spool" else None
    )
    corpus = {
        "slide_count": count, "entry_count": historical["entry_count"],
        "input_text_bytes": 100,
        "source_archive_bytes": historical["source_archive_bytes"],
        "source_archive_sha256": source_hash,
        "semantic_sha256": historical["semantic_sha256"],
        "full_text_sha256": historical["full_text_sha256"],
    }
    cases = []
    for mode in analyze.POLICIES:
        cases.append({
            "mode": mode, "slide_count": count, "entry_count": 37 + 2 * count,
            "source_archive_bytes": historical["source_archive_bytes"], "source_archive_sha256": source_hash,
            "output_bytes": 300, "output_sha256": output_hash,
            "scratch_bytes": (
                analyze.expected_spool_scratch_bytes(count)
                if mode == "spool" else None
            ),
            "byte_exact_control_match": True,
            "every_physical_member_verified": True,
            "every_slide_semantic_verified": True,
            "presentation_graph_verified": True,
            "slide_geometry_verified": True,
            "text_digest_verified": True,
        })
    operations = []
    measured = allocation() if instrumentation == "allocator" else None
    for sample in range(analyze.SAMPLES):
        operations.append({
            "mode": policy, "slide_count": count, "entry_count": 37 + 2 * count,
            "repeat": 0, "sample": sample,
            "source_archive_bytes": historical["source_archive_bytes"],
            "source_archive_sha256": source_hash, "input_text_bytes": 100,
            "elapsed_ns": 100 + sample, "output_bytes": 300,
            "output_write_calls": 10, "output_sha256": output_hash,
            "output_matches_oracle": True, "scratch_bytes": scratch,
            "allocation": measured.copy() if measured is not None else None,
            "process": process(),
        })
    return {
        "schema": analyze.REPORT_SCHEMA,
        "instrumentation": "none" if instrumentation == "normal" else "system_allocator_operation_scoped",
        "allocator": "Rust system allocator" if instrumentation == "normal" else "CountingSystemAllocator(std::alloc::System)",
        "timing_scope": "the timed operation constructs the public writer and File, runs finish, flush, and close",
        "oracle_scope": (
            "the control oracle is generated for every selected slide count; "
            "the explicit-spool oracle is byte-compared with control, and both "
            "are passed through the exact 37+2N physical-member, per-slide "
            "text/geometry, and presentation relationship-graph oracle"
        ),
        "cleanup_scope": "spool unlink is outside the timed operation",
        "control_storage_policy": "ordinary in-memory writer",
        "spool_storage_policy": "explicit caller-owned File",
        "samples": analyze.SAMPLES, "warmups": analyze.WARMUPS, "repeats": 1,
        "counts": [count], "modes": [policy],
        "spool_max_bytes": analyze.MAX_SPOOL_BYTES,
        "spool_buffer_bytes": analyze.SPOOL_BUFFER_BYTES,
        "corpora": [corpus], "cases": cases, "operations": operations,
    }


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n", encoding="utf-8")


def source_manifest(root: Path, name: str, entries: dict[str, str]) -> dict[str, object]:
    path = root / "validation-sources" / name
    write_json(path, entries)
    return {
        "path": path.relative_to(root).as_posix(),
        "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        "files": len(entries),
    }


def embedded_fixture(root: Path) -> dict[str, object]:
    """Create a portable 19-input custody fixture for mutation tests."""
    files: dict[str, dict[str, object]] = {}
    for source, destination in sorted(verify.EMBEDDED_INPUT_DESTINATIONS.items()):
        path = root / destination
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes((source + "\n").encode("utf-8"))
        files[source] = {
            "path": destination,
            "bytes": path.stat().st_size,
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
        }
    manifest_path = root / "embedded-inputs.json"
    write_json(manifest_path, {
        "schema": verify.EMBEDDED_INPUT_SCHEMA,
        "files": files,
    })
    return {
        "path": "embedded-inputs.json",
        "bytes": manifest_path.stat().st_size,
        "sha256": hashlib.sha256(manifest_path.read_bytes()).hexdigest(),
    }


def preliminary_fixture(root: Path) -> dict[str, object]:
    """Create a minimal archived-attempt manifest for custody mutations."""
    preliminary = root / "preliminary"
    attempt = preliminary / "attempt-01"
    attempt.mkdir(parents=True)
    (attempt / "protocol.json").write_text("{}\n", encoding="utf-8")
    (attempt / "marker.txt").write_text("retained\n", encoding="utf-8")
    files: dict[str, dict[str, object]] = {}
    for path in sorted(attempt.rglob("*")):
        if path.is_file():
            relative = path.relative_to(preliminary).as_posix()
            data = path.read_bytes()
            files[relative] = {
                "bytes": len(data),
                "sha256": hashlib.sha256(data).hexdigest(),
            }
    protocol_path = attempt / "protocol.json"
    protocol_reference = {
        "path": "attempt-01/protocol.json",
        "bytes": protocol_path.stat().st_size,
        "sha256": hashlib.sha256(protocol_path.read_bytes()).hexdigest(),
    }
    write_json(preliminary / "manifest.json", {
        "schema": verify.PRELIMINARY_SCHEMA,
        "attempt": "formal-before-cleanup-fix",
        "root": "attempt-01",
        "protocol": protocol_reference,
        "files": files,
    })
    manifest_path = preliminary / "manifest.json"
    data = manifest_path.read_bytes()
    return {
        "preliminary": {
            "path": verify.PRELIMINARY_MANIFEST_PATH,
            "bytes": len(data),
            "sha256": hashlib.sha256(data).hexdigest(),
        }
    }


def validation_fixture() -> tuple[Path, dict[str, object], dict[str, dict[str, object]], dict[str, list[str]]]:
    """Build a small complete ledger for source-custody mutation tests."""
    directory = Path(tempfile.mkdtemp(prefix="litchi-0478-evidence-"))
    final = source_manifest(
        directory, "final.json",
        {"src/lib.rs": "f" * 64,
         **{path: "f" * 64 for path in verify.HISTORICAL_SOURCE_EXCLUSIONS}},
    )
    before = source_manifest(directory, "before.json", {"src/lib.rs": "a" * 64})
    changed = source_manifest(directory, "changed.json", {"src/lib.rs": "b" * 64})
    protocol = {
        "schema": analyze.PROTOCOL_SCHEMA,
        "samples": analyze.SAMPLES,
        "warmups": analyze.WARMUPS,
        "cpu": 2,
        "preceding_evidence": analyze.PRECEDING_EVIDENCE,
        "preceding_evidence_sha256": analyze.PRECEDING_EVIDENCE_SHA256,
        "normal_and_allocator_timings_separate": True,
        "memory_gate": analyze.MEMORY_GATE_TEXT,
        "captures": [],
    }
    for capture in analyze.expected_captures():
        item = dict(capture)
        item["argv"] = spec(capture["instrumentation"], capture["policy"], capture["count"])["argv"]
        protocol["captures"].append(item)
    commands = verify.required_validation_argv(protocol)
    validation = directory / "validation"
    validation.mkdir()
    attempts: dict[str, dict[str, object]] = {}
    for label in verify.REQUIRED_FINAL_LABELS:
        source = final
        receipt = {
            "argv": commands[label],
            "exit_code": 0,
            "source_before": source,
            "source_after": source,
            "source_unchanged": True,
        }
        path = validation / f"{label}.json"
        write_json(path, receipt)
        attempts[label] = {
            "path": f"validation/{label}.json",
            "sha256": hashlib.sha256(path.read_bytes()).hexdigest(),
            "exit_code": 0,
            "source_unchanged": True,
            "argv": commands[label],
        }
    development = {
        "argv": ["cargo", "test", "development-attempt"],
        "exit_code": 101,
        "source_before": before,
        "source_after": changed,
        "source_unchanged": False,
    }
    development_path = validation / "development-failure.json"
    write_json(development_path, development)
    attempts["development-failure"] = {
        "path": "validation/development-failure.json",
        "sha256": hashlib.sha256(development_path.read_bytes()).hexdigest(),
        "argv": development["argv"],
        "exit_code": 101,
        "source_unchanged": False,
    }
    write_json(directory / "rust-validation.json", {
        "schema": "pptx-metadata-spool-rust-validation-v1",
        "required": list(verify.REQUIRED_FINAL_LABELS),
        "final_source_sha256": final["sha256"],
        "attempts": attempts,
    })
    binaries = {
        "normal": {"source_manifest_sha256": final["sha256"]},
        "allocator": {"source_manifest_sha256": final["sha256"]},
    }
    return directory, protocol, binaries, attempts


def analysis_fixture(root: Path) -> None:
    protocol: dict[str, object] = {
        "schema": analyze.PROTOCOL_SCHEMA,
        "samples": analyze.SAMPLES,
        "warmups": analyze.WARMUPS,
        "cpu": 2,
        "preceding_evidence": analyze.PRECEDING_EVIDENCE,
        "preceding_evidence_sha256": analyze.PRECEDING_EVIDENCE_SHA256,
        "normal_and_allocator_timings_separate": True,
        "memory_gate": analyze.MEMORY_GATE_TEXT,
        "captures": [],
    }
    captures = root / "captures"
    captures.mkdir(parents=True)
    for capture in analyze.expected_captures():
        item = dict(capture)
        item["argv"] = spec(capture["instrumentation"], capture["policy"], capture["count"])["argv"]
        protocol["captures"].append(item)
        write_json(captures / f"{capture['label']}.report.json", report_for(item))
        (captures / f"{capture['label']}.resource").write_text(
            "Maximum resident set size (kbytes): 10\n", encoding="utf-8"
        )
    write_json(root / "protocol.json", protocol)


class EvidenceTests(unittest.TestCase):
    def test_expected_protocol_has_24_reversed_captures(self) -> None:
        captures = analyze.expected_captures()
        self.assertEqual(len(captures), 24)
        self.assertEqual(captures[0]["label"], "r1-normal-8-control")
        self.assertEqual(captures[-1]["label"], "r2-normal-8-control")
        self.assertEqual(
            [row["label"].replace("r2-", "r1-", 1) for row in captures[12:]],
            [row["label"] for row in reversed(captures[:12])],
        )

    def test_report_requires_byte_exact_physical_and_semantic_oracles(self) -> None:
        current = spec()
        value = report_for(current)
        analyze.validate_report(value, current)
        value["cases"][0]["every_physical_member_verified"] = False
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(value, current)

    def test_spool_extent_matches_fixed_zip_name_derivation(self) -> None:
        self.assertEqual(
            [analyze.expected_spool_scratch_bytes(count) for count in analyze.COUNTS],
            [4074, 40842, 1237692],
        )

    def test_report_rejects_self_consistent_wrong_scratch_extent(self) -> None:
        current = spec(policy="spool")
        value = report_for(current)
        wrong = analyze.expected_spool_scratch_bytes(current["count"]) + 1
        for case in value["cases"]:
            if case["mode"] == "spool":
                case["scratch_bytes"] = wrong
        for operation in value["operations"]:
            operation["scratch_bytes"] = wrong
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(value, current)

    def test_report_rejects_control_scratch_and_failed_allocator(self) -> None:
        control = spec(policy="control")
        value = report_for(control)
        value["operations"][0]["scratch_bytes"] = 1
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(value, control)

        allocator_case = spec(instrumentation="allocator")
        allocator_report = report_for(allocator_case)
        allocator_report["operations"][0]["allocation"]["failed_allocation_calls"] = 1
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(allocator_report, allocator_case)

    def test_allocator_report_must_balance_live_bytes(self) -> None:
        current = spec(instrumentation="allocator")
        value = report_for(current)
        value["operations"][0]["allocation"]["live_bytes_after"] = 101
        with self.assertRaises(analyze.AnalysisError):
            analyze.validate_report(value, current)

    def test_command_cannot_move_report_or_spool_outside_label(self) -> None:
        current = spec()
        argv = list(current["argv"])
        argv[23] = "/bundle/captures/other.report.json"
        with self.assertRaises(verify.VerificationError):
            verify.check_command_argv(argv, current, "capture", Path("/bundle"))

    def test_duplicate_json_keys_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "duplicate.json"
            path.write_text('{"a":1,"a":2}\n', encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.read_json(path, "duplicate")

    def test_embedded_input_resource_mutation_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            reference = embedded_fixture(root)
            self.assertEqual(
                verify.check_embedded_inputs(root, reference, "embedded"),
                reference,
            )
            destination = root / "embedded-resources/notes/notesMaster.xml"
            destination.write_bytes(destination.read_bytes() + b"mutation")
            with self.assertRaises(verify.VerificationError):
                verify.check_embedded_inputs(root, reference, "embedded")

    def test_environment_tmpfs_context_mutation_is_rejected(self) -> None:
        bundle = Path(__file__).resolve().parent
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shutil.copy2(bundle / "environment.json", root / "environment.json")
            protocol_path = bundle / "protocol.json"
            if not protocol_path.is_file():
                protocol_path = next(
                    (bundle / "preliminary").glob("*/protocol.json")
                )
            protocol = verify.read_json(protocol_path, "protocol")
            verify.check_environment(root, protocol)
            value = json.loads((root / "environment.json").read_text(encoding="utf-8"))
            value["filesystem"]["stdout"] = value["filesystem"]["stdout"].replace(
                "tmpfs          tmpfs", "ext4          ext4", 1
            )
            write_json(root / "environment.json", value)
            with self.assertRaises(verify.VerificationError):
                verify.check_environment(root, protocol)

    def test_preliminary_manifest_mutation_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            protocol = preliminary_fixture(root)
            verify.check_preliminary_manifest(root, protocol)
            marker = root / "preliminary/attempt-01/marker.txt"
            marker.write_text("mutated\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.check_preliminary_manifest(root, protocol)

    def test_bundle_paths_reject_traversal_and_symlinks(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "inside.txt").write_text("ok", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.bundle_file(root, "../outside", "path")
            (root / "link").symlink_to(root / "inside.txt")
            with self.assertRaises(verify.VerificationError):
                verify.bundle_file(root, "link", "path")

    def test_seal_rejects_inventory_mutation_after_reseal(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            payload = root / "payload.txt"
            payload.write_text("before\n", encoding="utf-8")
            digest = __import__("hashlib").sha256(payload.read_bytes()).hexdigest()
            for name in ("protocol.json", "summary.json"):
                (root / name).write_text("{}\n", encoding="utf-8")
            entries = {
                "payload.txt": digest,
                "protocol.json": __import__("hashlib").sha256((root / "protocol.json").read_bytes()).hexdigest(),
                "summary.json": __import__("hashlib").sha256((root / "summary.json").read_bytes()).hexdigest(),
            }
            (root / "SHA256SUMS").write_text(
                "".join(f"{value}  {name}\n" for name, value in entries.items()),
                encoding="utf-8",
            )
            verify.check_seal(root, required=True)
            payload.write_text("after\n", encoding="utf-8")
            with self.assertRaises(verify.VerificationError):
                verify.check_seal(root, required=True)

    def test_final_ledger_cannot_drop_a_required_gate(self) -> None:
        root, protocol, binaries, _attempts = validation_fixture()
        try:
            self.assertEqual(
                verify.check_rust_validation(root, binaries, protocol)["required_success"],
                len(verify.REQUIRED_FINAL_LABELS),
            )
            ledger_path = root / "rust-validation.json"
            ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
            ledger["required"] = ledger["required"][1:]
            write_json(ledger_path, ledger)
            with self.assertRaises(verify.VerificationError):
                verify.check_rust_validation(root, binaries, protocol)
        finally:
            # The fixture is outside the repository and is intentionally
            # removed with a narrow, validated path.
            for path in sorted(root.rglob("*"), reverse=True):
                if path.is_file() or path.is_symlink():
                    path.unlink()
                elif path.is_dir():
                    path.rmdir()
            root.rmdir()

    def test_final_ledger_schema_is_authenticated(self) -> None:
        root, protocol, binaries, _attempts = validation_fixture()
        try:
            ledger_path = root / "rust-validation.json"
            ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
            ledger["schema"] = "wrong-schema"
            write_json(ledger_path, ledger)
            with self.assertRaises(verify.VerificationError):
                verify.check_rust_validation(root, binaries, protocol)
        finally:
            for path in sorted(root.rglob("*"), reverse=True):
                if path.is_file() or path.is_symlink():
                    path.unlink()
                elif path.is_dir():
                    path.rmdir()
            root.rmdir()

    def test_final_ledger_attempt_argv_is_authenticated(self) -> None:
        root, protocol, binaries, _attempts = validation_fixture()
        try:
            ledger_path = root / "rust-validation.json"
            ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
            del ledger["attempts"][verify.REQUIRED_FINAL_LABELS[0]]["argv"]
            write_json(ledger_path, ledger)
            with self.assertRaises(verify.VerificationError):
                verify.check_rust_validation(root, binaries, protocol)
        finally:
            for path in sorted(root.rglob("*"), reverse=True):
                if path.is_file() or path.is_symlink():
                    path.unlink()
                elif path.is_dir():
                    path.rmdir()
            root.rmdir()

    def test_changed_development_attempt_is_retained_without_source_exclusion(self) -> None:
        root, protocol, binaries, attempts = validation_fixture()
        try:
            result = verify.check_rust_validation(root, binaries, protocol)
            self.assertEqual(result["retained_nonzero_attempts"], 1)
            self.assertNotIn("source_exclusions", attempts["development-failure"])
        finally:
            for path in sorted(root.rglob("*"), reverse=True):
                if path.is_file() or path.is_symlink():
                    path.unlink()
                elif path.is_dir():
                    path.rmdir()
            root.rmdir()

    def test_required_gate_source_change_is_rejected(self) -> None:
        root, protocol, binaries, _attempts = validation_fixture()
        try:
            changed = source_manifest(root, "late-change.json", {"src/lib.rs": "c" * 64})
            label = verify.REQUIRED_FINAL_LABELS[0]
            path = root / "validation" / f"{label}.json"
            receipt = json.loads(path.read_text(encoding="utf-8"))
            receipt["source_after"] = changed
            receipt["source_unchanged"] = False
            write_json(path, receipt)
            ledger_path = root / "rust-validation.json"
            ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
            ledger["attempts"][label]["sha256"] = hashlib.sha256(path.read_bytes()).hexdigest()
            ledger["attempts"][label]["source_unchanged"] = False
            write_json(ledger_path, ledger)
            with self.assertRaises(verify.VerificationError):
                verify.check_rust_validation(root, binaries, protocol)
        finally:
            for path in sorted(root.rglob("*"), reverse=True):
                if path.is_file() or path.is_symlink():
                    path.unlink()
                elif path.is_dir():
                    path.rmdir()
            root.rmdir()

    def test_source_exclusion_is_rejected_for_every_new_final_gate(self) -> None:
        root, protocol, binaries, _attempts = validation_fixture()
        try:
            label = "shared-streaming-unit"
            ledger_path = root / "rust-validation.json"
            ledger = json.loads(ledger_path.read_text(encoding="utf-8"))
            ledger["attempts"][label]["source_exclusions"] = [
                *verify.HISTORICAL_SOURCE_EXCLUSIONS
            ]
            write_json(ledger_path, ledger)
            with self.assertRaises(verify.VerificationError):
                verify.check_rust_validation(root, binaries, protocol)
        finally:
            for path in sorted(root.rglob("*"), reverse=True):
                if path.is_file() or path.is_symlink():
                    path.unlink()
                elif path.is_dir():
                    path.rmdir()
            root.rmdir()

    def test_anomalous_r2_allocator_peak_is_included_in_memory_gate(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            analysis_fixture(root)
            path = root / "captures" / "r2-allocator-8-spool.report.json"
            report = json.loads(path.read_text(encoding="utf-8"))
            for operation in report["operations"]:
                operation["allocation"]["region_peak_live_bytes"] = 1_000
                operation["allocation"]["peak_live_bytes_after"] = 1_000
            write_json(path, report)
            original = analyze.ROOT
            try:
                analyze.ROOT = root
                summary = analyze.derive()
            finally:
                analyze.ROOT = original
            gate = summary["memory_gate"]
            self.assertIn("r2-allocator-8-spool", gate["lane_maximum_peaks"])
            self.assertGreater(gate["global_maximum"], gate["smallest_count_baseline_minimum"])
            self.assertFalse(gate["within_threshold"])
            self.assertEqual(gate["status"], "requires_investigation")

    def test_low_r2_large_allocator_peak_is_included_in_memory_range(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            analysis_fixture(root)
            path = root / "captures" / "r2-allocator-8192-spool.report.json"
            report = json.loads(path.read_text(encoding="utf-8"))
            for operation in report["operations"]:
                operation["allocation"]["region_peak_live_bytes"] = 100
                operation["allocation"]["peak_live_bytes_after"] = 100
            write_json(path, report)
            original = analyze.ROOT
            try:
                analyze.ROOT = root
                summary = analyze.derive()
            finally:
                analyze.ROOT = original
            gate = summary["memory_gate"]
            self.assertIn("r2-allocator-8192-spool", gate["lane_maximum_peaks"])
            self.assertEqual(gate["global_minimum"], 0)
            self.assertFalse(gate["within_threshold"])
            self.assertEqual(gate["status"], "requires_investigation")


if __name__ == "__main__":
    unittest.main(verbosity=2)
