#!/usr/bin/env python3
"""End-to-end build-free tests for the ODG source-build handoff seam."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
import tempfile
import time
import unittest
from unittest import mock

from tools import odg_joined_qualification_builder as BUILDER
from tools import odg_joined_qualification as RUNTIME
from tools.test_odg_joined_qualification import write_manifest


def _file_identity(path: Path) -> dict[str, object]:
    data = path.read_bytes()
    return {"path": str(path), "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}


class ODGSourceBuildHandoffTests(unittest.TestCase):
    def _reviewed_fixture(self) -> Path:
        # Runtime admission deliberately accepts only this fixed reviewed
        # harness file.  Source endpoints remain separate temporary Git roots.
        return BUILDER.REVIEWED_REPO_ROOT / BUILDER.REVIEWED_FIXTURE_RELATIVE

    def _fixture_and_plan(self, root: Path) -> tuple[Path, dict[str, object], dict[str, object]]:
        manifest = write_manifest(root)
        data = json.loads(manifest.read_text(encoding="utf-8"))
        # Keep the source handoff and runtime validator on the same fixed
        # reviewed harness path; do not replace it with a packet-local copy.
        data["fixture"]["path"] = str(self._reviewed_fixture())
        manifest.write_text(json.dumps(data), encoding="utf-8")
        self.baseline_ref = data["baseline"]["ref"]
        with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref):
            plan = BUILDER.prepare_plan(manifest, allow_built_artifacts=True)
        return manifest, data, plan

    def _source_binder(
        self, root: Path, data: dict[str, object], binder: Path,
    ) -> tuple[Path, Path, Path, Path]:
        def endpoint(source: dict[str, object], profile: str) -> dict[str, object]:
            role = "baseline" if profile == "legacy-apply-joined" else "C2a"
            return {
                "root": source["root"],
                "ref": source["ref"],
                "tree": source["tree"],
                "profile": profile,
                "features": sorted(source["features"]),
                "source_manifest": source["source_manifest"],
                "source_closure": source["source_closure"],
                "lock": source["lock"],
                "toolchain": source["toolchain"],
                "cargo_config": source["cargo_config"],
                "cargo_manifest": source["cargo_manifest"],
                "binary": {
                    "path": str(Path(source["root"]) / "pending" / "release" / "examples" / BUILDER.EXPECTED_EXAMPLE),
                    "bytes": None,
                    "sha256": None,
                    "status": f"pending {role} normal build",
                },
                "allocator_binary": {
                    "path": str(Path(source["root"]) / "pending-allocator" / "release" / "examples" / BUILDER.EXPECTED_EXAMPLE),
                    "bytes": None,
                    "sha256": None,
                    "status": f"pending {role} allocator build",
                },
            }

        fixture = self._reviewed_fixture()
        reviewed = {
            "fixture": {**_file_identity(fixture), "provenance": BUILDER.REVIEWED_FIXTURE_PROVENANCE},
        }
        for name, relative, destination in (
            ("common", BUILDER.REVIEWED_PROBE_RELATIVES["common"], BUILDER.EXPECTED_COMMON_DESTINATION),
            ("allocator", BUILDER.REVIEWED_PROBE_RELATIVES["allocator"], "crates/litchi-odg/examples/odg_joined_qualification_allocator.rs"),
            ("baseline_adapter", BUILDER.REVIEWED_PROBE_RELATIVES["baseline"], BUILDER.EXPECTED_OVERLAY_DESTINATION),
            ("candidate_adapter", BUILDER.REVIEWED_PROBE_RELATIVES["candidate"], BUILDER.EXPECTED_OVERLAY_DESTINATION),
        ):
            path = BUILDER.REVIEWED_REPO_ROOT / relative
            reviewed[name] = {**_file_identity(path), "destination": destination}

        target_root = root / "source-target-root"
        build_output = root / "source-build-output"
        tmpdir = root / "source-tmp"
        tmpdir.mkdir()
        binder.write_text(json.dumps({
            "schema": BUILDER.SOURCE_BINDER_SCHEMA,
            "status": "source-bound; runner manifest is generated only after four binary identities and preparation/build receipts exist",
            "claim_scope": "source/fixture/lock/probe binding only; normal and allocator binaries are pending the authorized C2a build",
            "baseline": endpoint(data["baseline"], "legacy-apply-joined"),
            "candidate": endpoint(data["candidate"], "current-edit-join-commit"),
            "reviewed_inputs": reviewed,
        }), encoding="utf-8")
        return binder, target_root, build_output, tmpdir

    def _fake_stream(self):
        def stream(command: list[str], cwd: Path, log_dir: Path, label: str, target: Path, cpu: int, **kwargs: object) -> dict[str, object]:
            log_dir.mkdir(parents=True, exist_ok=True)
            stdout = log_dir / f"{label}.stdout"
            stderr = log_dir / f"{label}.stderr"
            stdout.write_bytes(b"synthetic bounded build\n")
            stderr.write_bytes(b"")
            artifact = target / "release" / "examples" / BUILDER.EXPECTED_EXAMPLE
            artifact.parent.mkdir(parents=True, exist_ok=True)
            artifact.write_bytes(b"#!/bin/sh\nexit 0\n")
            artifact.chmod(0o755)
            now = time.time_ns()
            return {
                "label": label,
                "command": command,
                "cwd": str(cwd),
                "returncode": 0,
                "wall_ns": 1,
                "target": {"path": str(target), "bytes_before": 0, "bytes_after": BUILDER._directory_bytes(target), "hard_limit_bytes": BUILDER.MEMORY_LIMIT_BYTES},
                "stdout": _file_identity(stdout),
                "stderr": _file_identity(stderr),
                "monitor": {"samples": [{"checked_unix_ns": now, "output_bytes": stdout.stat().st_size + stderr.stat().st_size, "rss_bytes": 0, "load_1m": 0.0, "outside_processes": [], "owned_processes": []}], "memory_limit_bytes": BUILDER.MEMORY_LIMIT_BYTES, "max_output_bytes": BUILDER.MAX_BUILD_OUTPUT, "max_load": kwargs["max_load"], "owned_processes": []},
            }
        return stream

    def _run_prepare_and_build(
        self, root: Path, data: dict[str, object], prep_path: Path,
    ) -> tuple[Path, Path]:
        binder, target_root, build_output, tmpdir = self._source_binder(root, data, root / "source-binder.json")
        BUILDER.prepare_source(
            binder, target_root=target_root, build_output=build_output, tmpdir=tmpdir,
            cpu=0, max_load=1000, max_iterations=1000, timeout_s=2, output=prep_path,
        )
        BUILDER.build_source(prep_path)
        return build_output, prep_path

    def test_public_prepare_build_callback_producer_and_final_validator(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            _, data, _ = self._fixture_and_plan(root)
            prep_path = root / "source-preparation.json"
            with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref), \
                 mock.patch.object(BUILDER, "_quiet_guard", side_effect=lambda limit, phase: {"phase": phase, "load_1m": 0.0, "max_load": limit, "heavy_processes": [], "checked_unix_ns": time.time_ns()}), \
                 mock.patch.object(BUILDER, "_stream_command", side_effect=self._fake_stream()):
                build_output, _ = self._run_prepare_and_build(root, data, prep_path)
                producer = json.loads((build_output / "source-build-receipt.json").read_text(encoding="utf-8"))
                self.assertEqual(
                    [event["label"] for event in producer["events"]],
                    ["baseline-normal", "baseline-allocator", "candidate-normal", "candidate-allocator"],
                )
                self.assertTrue(all(
                    event["monitor"]["max_output_bytes"] == BUILDER.MAX_BUILD_OUTPUT
                    for event in producer["events"]
                ))
                final_manifest = root / "final-input.json"
                BUILDER.seal_source_build(prep_path, output_manifest=final_manifest)
            self.assertTrue(final_manifest.is_file())
            with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref):
                BUILDER.validate_build_receipt(final_manifest)
            parsed = RUNTIME.load_manifest(final_manifest)
            runtime_output = root / "runtime-output"
            runtime_output.mkdir()
            with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref):
                identity = RUNTIME.runtime_identity(parsed, runtime_output, require_build_receipt=True)
            self.assertEqual(identity["fixture"]["path"], str(self._reviewed_fixture()))
            self.assertEqual(
                set(identity["probe_sources"]),
                {"common", "allocator", "baseline", "candidate"},
            )

    def test_seal_rejects_producer_artifact_drift_before_final_admission(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            _, data, _ = self._fixture_and_plan(root)
            prep_path = root / "source-preparation.json"
            with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref), \
                 mock.patch.object(BUILDER, "_quiet_guard", side_effect=lambda limit, phase: {"phase": phase, "load_1m": 0.0, "max_load": limit, "heavy_processes": [], "checked_unix_ns": time.time_ns()}), \
                 mock.patch.object(BUILDER, "_stream_command", side_effect=self._fake_stream()):
                build_output, _ = self._run_prepare_and_build(root, data, prep_path)
                receipt = build_output / "source-build-receipt.json"
                receipt_data = json.loads(receipt.read_text(encoding="utf-8"))
                receipt_data["events"][0]["artifact"]["sha256"] = "0" * 64
                receipt.write_text(json.dumps(receipt_data), encoding="utf-8")
                with self.assertRaises(BUILDER.RunFailure):
                    BUILDER.seal_source_build(prep_path, output_manifest=root / "rejected.json")


if __name__ == "__main__":
    unittest.main()
