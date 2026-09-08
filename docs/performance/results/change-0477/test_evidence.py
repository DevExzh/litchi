#!/usr/bin/env python3
"""Portable mutation tests for the 0477 evidence boundary."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import tempfile
import unittest

import verify


ENVIRONMENT = {
    "RUSTUP_TOOLCHAIN": "test-toolchain",
    "CARGO_BUILD_JOBS": "4",
    "CARGO_INCREMENTAL": "0",
    "CARGO_PROFILE_RELEASE_DEBUG": "1",
    "RUSTFLAGS": "-C force-frame-pointers=yes",
    "DEBUGINFOD_URLS": "",
    "LC_ALL": "C",
}


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def write_json(path: Path, value: object) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(value, indent=2, sort_keys=True) + "\n",
                    encoding="utf-8")


def manifest(root: Path, name: str, entries: dict[str, str]) -> dict[str, object]:
    path = root / "validation-sources" / name
    write_json(path, entries)
    return {
        "path": path.relative_to(root).as_posix(),
        "sha256": sha256(path),
        "files": len(entries),
    }


def metadata(path: Path) -> dict[str, object]:
    return {"bytes": path.stat().st_size, "sha256": sha256(path)}


def _build_gate_fixture(root: Path) -> tuple[dict[str, object], dict[str, object], Path, Path]:
    (root / "gate.py").write_text("gate\n", encoding="utf-8")
    (root / "common.py").write_text("common\n", encoding="utf-8")
    source = manifest(root, "source.json", {"src/lib.rs": "a" * 64})
    gate_path = root / "validation" / "build-normal.json"
    wrapper_path = root / "normal-build.json"
    command = [
        "cargo", "build", "--release", "--locked", "--manifest-path",
        "tools/perf-baseline/Cargo.toml", "--bin", "zip_directory_spool",
    ]
    mirrored: dict[str, object] = {
        "argv": command,
        "cwd": "/tmp/litchi-goal-0477-test",
        "environment": ENVIRONMENT,
        "driver_sha256": sha256(root / "gate.py"),
        "common_sha256": sha256(root / "common.py"),
        "source_before": source,
        "source_after": source,
        "started_utc": "2026-01-01T00:00:00+00:00",
        "finished_utc": "2026-01-01T00:00:01+00:00",
        "exit_code": 0,
        "source_unchanged": True,
        "clean_after": True,
    }
    write_json(gate_path, mirrored)
    binary_hash = hashlib.sha256(b"bin").hexdigest()
    copied_binary = {
        "path": "/tmp/litchi-goal-0477-test/normal/zip_directory_spool",
        "bytes": 3,
        "sha256": binary_hash,
    }
    wrapper = dict(
        mirrored,
        binary=copied_binary,
        original_binary=dict(
            copied_binary,
            path="/home/build/target/release/zip_directory_spool",
        ),
        copied_utc="2026-01-01T00:00:02+00:00",
        validation_gate={
            "path": gate_path.relative_to(root).as_posix(),
            "sha256": sha256(gate_path),
        },
    )
    write_json(wrapper_path, wrapper)
    spec: dict[str, object] = {
        **copied_binary,
        "build_gate": wrapper_path.name,
        "build_gate_sha256": sha256(wrapper_path),
        "source_manifest_sha256": source["sha256"],
    }
    protocol = {"environment": ENVIRONMENT}
    return spec, protocol, gate_path, wrapper_path


def _reseal(root: Path) -> None:
    entries = {}
    for name in ("protocol.json", "binaries.json", "summary.json",
                 "capture.report.json"):
        entries[name] = sha256(root / name)
    (root / "SHA256SUMS").write_text(
        "".join(f"{value}  {name}\n" for name, value in sorted(entries.items())),
        encoding="utf-8",
    )


class EvidenceTests(unittest.TestCase):
    def test_resealed_capture_corruption_still_fails_receipt_binding(self) -> None:
        """A new seal cannot rewrite the receipt's original artifact identity."""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            for name in ("protocol.json", "binaries.json", "summary.json"):
                (root / name).write_text("{}\n", encoding="utf-8")
            artifact = root / "capture.report.json"
            artifact.write_text('{"status":"original"}\n', encoding="utf-8")
            retained_metadata = metadata(artifact)
            _reseal(root)
            verify.check_seal(root)

            artifact.write_text('{"status":"changed-after-capture"}\n',
                                encoding="utf-8")
            _reseal(root)
            # The seal is internally consistent after the rewrite, but the
            # retained capture receipt still binds the old artifact bytes.
            verify.check_seal(root)
            with self.assertRaises(verify.VerificationError):
                verify.check_artifact_map(
                    root, root,
                    {"capture.report.json": retained_metadata},
                    {"capture.report.json"},
                    "capture",
                )

    def test_final_build_gate_requires_a_successful_designated_receipt(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            spec, protocol, gate_path, wrapper_path = _build_gate_fixture(root)
            verify.check_gate(root, spec, "normal", protocol)

            receipt = json.loads(gate_path.read_text(encoding="utf-8"))
            receipt["exit_code"] = 101
            write_json(gate_path, receipt)
            wrapper = json.loads(wrapper_path.read_text(encoding="utf-8"))
            wrapper["validation_gate"]["sha256"] = sha256(gate_path)
            write_json(wrapper_path, wrapper)
            spec["build_gate_sha256"] = sha256(wrapper_path)
            with self.assertRaises(verify.VerificationError):
                verify.check_gate(root, spec, "normal", protocol)

    def test_final_build_gate_must_bind_the_declared_source_manifest(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            spec, protocol, gate_path, wrapper_path = _build_gate_fixture(root)
            alternate = manifest(root, "alternate.json", {"src/lib.rs": "b" * 64})
            verify.check_gate(root, spec, "normal", protocol)

            receipt = json.loads(gate_path.read_text(encoding="utf-8"))
            receipt["source_before"] = alternate
            receipt["source_after"] = alternate
            write_json(gate_path, receipt)
            wrapper = json.loads(wrapper_path.read_text(encoding="utf-8"))
            wrapper["source_before"] = alternate
            wrapper["source_after"] = alternate
            wrapper["validation_gate"]["sha256"] = sha256(gate_path)
            write_json(wrapper_path, wrapper)
            spec["build_gate_sha256"] = sha256(wrapper_path)
            with self.assertRaises(verify.VerificationError):
                verify.check_gate(root, spec, "normal", protocol)

    def test_final_build_gate_requires_origin_and_copy_binary_metadata(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            spec, protocol, _gate_path, wrapper_path = _build_gate_fixture(root)
            verify.check_gate(root, spec, "normal", protocol)

            wrapper = json.loads(wrapper_path.read_text(encoding="utf-8"))
            wrapper["original_binary"]["sha256"] = "0" * 64
            write_json(wrapper_path, wrapper)
            spec["build_gate_sha256"] = sha256(wrapper_path)
            with self.assertRaises(verify.VerificationError):
                verify.check_gate(root, spec, "normal", protocol)

    def test_rust_validation_requires_final_source_and_exact_receipt_inventory(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            spec, _protocol, gate_path, wrapper_path = _build_gate_fixture(root)
            validation = root / "validation"
            excluded_source = "tools/perf-baseline/src/zip_directory_spool.rs"
            final_ref = manifest(
                root, "final-source.json",
                {"src/lib.rs": "a" * 64, excluded_source: "b" * 64},
            )
            wrapper = json.loads(wrapper_path.read_text(encoding="utf-8"))
            wrapper["source_before"] = final_ref
            wrapper["source_after"] = final_ref
            write_json(wrapper_path, wrapper)
            final_source = final_ref["sha256"]
            spec["source_manifest_sha256"] = final_source
            old_ref = {
                "path": "validation-sources/source.json",
                "sha256": sha256(root / "validation-sources/source.json"),
                "files": 1,
            }
            alternate = manifest(root, "alternate.json", {"src/lib.rs": "b" * 64})

            def receipt(label: str, argv: list[str], source_before: dict[str, object],
                        source_after: dict[str, object], exit_code: int = 0,
                        unchanged: bool = True) -> Path:
                path = validation / f"{label}.json"
                write_json(path, {
                    "argv": argv,
                    "exit_code": exit_code,
                    "source_before": source_before,
                    "source_after": source_after,
                    "source_unchanged": unchanged,
                })
                return path

            gate_record = {
                "argv": [
                    "cargo", "build", "--release", "--locked", "--manifest-path",
                    "tools/perf-baseline/Cargo.toml", "--bin", "zip_directory_spool",
                ],
                "exit_code": 0,
                "source_before": final_ref,
                "source_after": final_ref,
                "source_unchanged": True,
            }
            write_json(gate_path, gate_record)
            wrapper["validation_gate"]["sha256"] = sha256(gate_path)
            write_json(wrapper_path, wrapper)

            required_labels = list(verify.REQUIRED_FINAL_VALIDATION_LABELS)
            paths: dict[str, Path] = {"build-normal": gate_path}
            argv_by_label = {
                label: ["cargo", "test", label]
                for label in required_labels
                if label not in {"build-normal", "build-allocator"}
            }
            argv_by_label["build-normal"] = gate_record["argv"]
            argv_by_label["build-allocator"] = gate_record["argv"] + [
                "--features", "allocator-metrics"
            ]
            for label in required_labels:
                if label == "build-normal":
                    continue
                before = old_ref if label == "opc-tests" else final_ref
                after = before
                paths[label] = receipt(label, argv_by_label[label], before, after)
            retained_path = receipt(
                "retained-failure", ["cargo", "test", "old"], old_ref,
                alternate, exit_code=101, unchanged=False)
            attempts = {}
            for label in required_labels:
                path = paths[label]
                entry: dict[str, object] = {
                    "path": f"validation/{label}.json",
                    "sha256": sha256(path),
                    "exit_code": 0,
                    "source_unchanged": True,
                    "argv": argv_by_label[label],
                }
                if label == "opc-tests":
                    entry["source_exclusions"] = [excluded_source]
                attempts[label] = entry
            attempts["retained-failure"] = {
                "path": "validation/retained-failure.json",
                "sha256": sha256(retained_path),
                "exit_code": 101,
                "source_unchanged": False,
            }
            ledger = {
                "schema": "zip-directory-spool-rust-validation-v1",
                "required": required_labels,
                "final_source_sha256": final_source,
                "attempts": attempts,
            }
            write_json(root / "rust-validation.json", ledger)
            binaries = {
                "normal": {"build_gate": "normal-build.json",
                           "source_manifest_sha256": final_source},
                "allocator": {"build_gate": "normal-build.json",
                              "source_manifest_sha256": final_source},
            }
            self.assertEqual(
                verify.check_rust_validation(root, binaries),
                {"required_success": len(required_labels),
                 "receipts": len(required_labels) + 1,
                 "retained_nonzero_attempts": 1},
            )
            ledger["required"] = [label for label in required_labels
                                   if label != "pilot-normal"]
            write_json(root / "rust-validation.json", ledger)
            with self.assertRaises(verify.VerificationError):
                verify.check_rust_validation(root, binaries)
            ledger["required"] = required_labels
            ledger["attempts"]["build-normal"]["source_exclusions"] = [excluded_source]
            write_json(root / "rust-validation.json", ledger)
            with self.assertRaises(verify.VerificationError):
                verify.check_rust_validation(root, binaries)
            ledger["attempts"]["build-normal"].pop("source_exclusions")
            ledger["attempts"]["build-normal"]["exit_code"] = 101
            write_json(root / "rust-validation.json", ledger)
            with self.assertRaises(verify.VerificationError):
                verify.check_rust_validation(root, binaries)

    def test_retained_success_may_have_source_change_outside_final_build_gate(self) -> None:
        """Retained development receipts stay historical outside required gates."""
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "gate.py").write_text("gate\n", encoding="utf-8")
            (root / "common.py").write_text("common\n", encoding="utf-8")
            before = manifest(root, "before.json", {"src/lib.rs": "a" * 64})
            after = manifest(root, "after.json", {"src/lib.rs": "b" * 64})
            validation = root / "validation"
            validation.mkdir()
            for label, exit_code, source_after, unchanged in (
                ("successful-development", 0, after, False),
                ("failed-development", 101, before, True),
            ):
                stdout = validation / f"{label}.stdout"
                stderr = validation / f"{label}.stderr"
                stdout.write_text(f"{label} stdout\n", encoding="utf-8")
                stderr.write_text(f"{label} stderr\n", encoding="utf-8")
                started = {
                    "argv": ["gate.py", label],
                    "common_sha256": sha256(root / "common.py"),
                    "cwd": "/tmp/litchi-goal-0477-test",
                    "driver_sha256": sha256(root / "gate.py"),
                    "environment": ENVIRONMENT,
                    "source_before": before,
                    "started_utc": "2026-01-01T00:00:00+00:00",
                }
                finished = dict(
                    started,
                    artifacts={
                        stdout.name: metadata(stdout),
                        stderr.name: metadata(stderr),
                    },
                    exit_code=exit_code,
                    finished_utc="2026-01-01T00:00:01+00:00",
                    source_after=source_after,
                    source_unchanged=unchanged,
                )
                write_json(validation / f"{label}.started.json", started)
                write_json(validation / f"{label}.json", finished)
            self.assertEqual(
                verify.check_validation_receipts(root, {"environment": ENVIRONMENT}),
                {"total": 2, "successful": 1, "failed": 1},
            )


if __name__ == "__main__":
    unittest.main(verbosity=2)
