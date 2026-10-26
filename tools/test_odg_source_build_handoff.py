#!/usr/bin/env python3
"""End-to-end build-free tests for the ODG source-build handoff seam."""

from __future__ import annotations

import copy
import hashlib
import json
from pathlib import Path
import shutil
import tempfile
import time
import unittest
from unittest import mock

from tools import odg_joined_qualification_builder as BUILDER
from tools.test_odg_joined_qualification import write_manifest


def _file_identity(path: Path) -> dict[str, object]:
    data = path.read_bytes()
    return {"path": str(path), "bytes": len(data), "sha256": hashlib.sha256(data).hexdigest()}


def _source_shape(endpoint: dict[str, object]) -> dict[str, object]:
    return {key: endpoint[key] for key in (
        "ref", "tree", "source_manifest", "source_closure", "lock", "toolchain",
        "cargo_config", "cargo_manifest",
    )}


class ODGSourceBuildHandoffTests(unittest.TestCase):
    def _fixture_and_plan(self, root: Path) -> tuple[Path, dict[str, object], dict[str, object]]:
        manifest = write_manifest(root)
        data = json.loads(manifest.read_text(encoding="utf-8"))
        fixture = root / "external-reviewed-fixture.xml"
        reviewed_fixture = BUILDER.REVIEWED_REPO_ROOT / BUILDER.REVIEWED_FIXTURE_RELATIVE
        fixture.write_bytes(reviewed_fixture.read_bytes())
        data["fixture"]["path"] = str(fixture)
        manifest.write_text(json.dumps(data), encoding="utf-8")
        self.baseline_ref = data["baseline"]["ref"]
        with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref):
            plan = BUILDER.prepare_plan(manifest, allow_built_artifacts=True)
        return manifest, data, plan

    def _preparation(self, root: Path, plan: dict[str, object], binder: Path, output: Path) -> dict[str, object]:
        reviewed = {
            "fixture": {
                **_file_identity(root / "external-reviewed-fixture.xml"),
                "provenance": BUILDER.REVIEWED_FIXTURE_PROVENANCE,
            },
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
        binder.write_text("{}", encoding="utf-8")
        return {
            "schema": BUILDER.SOURCE_PREPARATION_SCHEMA,
            "builder_revision": BUILDER.BUILDER_REVISION,
            "builder_source": BUILDER._identity(Path(BUILDER.__file__).resolve()),
            "status": "source_prepared",
            "source_binder": _file_identity(binder),
            "baseline": plan["baseline"],
            "candidate": plan["candidate"],
            "reviewed_inputs": reviewed,
            "overlays": [
                {"source": reviewed["common"]["path"], "destination": BUILDER.EXPECTED_COMMON_DESTINATION},
                {"source": reviewed["allocator"]["path"], "destination": "crates/litchi-odg/examples/odg_joined_qualification_allocator.rs"},
                {"source": reviewed["baseline_adapter"]["path"], "destination": BUILDER.EXPECTED_OVERLAY_DESTINATION, "endpoint": "baseline"},
                {"source": reviewed["candidate_adapter"]["path"], "destination": BUILDER.EXPECTED_OVERLAY_DESTINATION, "endpoint": "candidate"},
            ],
            "commands": plan["commands"],
            "target_root": str(target_root),
            "build_output": str(build_output),
            "tmpdir": str(tmpdir),
            "cpu": 0,
            "max_load": plan["max_load"],
            "max_iterations": 1000,
            "timeout_s": 2,
            "abba_order": ["a1", "b1", "b2", "a2"],
            "blocks": 2,
            "warmups": 20,
            "samples_per_fresh_child": 1,
            "target_ms": 500,
            "minimum_ms": 150,
            "composition_limits": BUILDER.SOURCE_COMPOSITION_LIMITS,
            "closure_policy": "synthetic test callback; production path revalidates the full closure",
            "claim_scope": "source preparation only; no build, binary, correctness, timing, or acceptance claim",
            "source_preparation": {"path": str(output)},
        }

    def _fake_stream(self, preparation: dict[str, object]):
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

    def test_public_prepare_build_callback_producer_and_final_validator(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest, _, plan = self._fixture_and_plan(root)
            # The source-build producer owns the fresh target directories.
            for endpoint in (plan["baseline"], plan["candidate"]):
                shutil.rmtree(endpoint["target_dir"])
                shutil.rmtree(endpoint["allocator_target_dir"])
            binder = root / "source-binder.json"
            prep_path = root / "source-preparation.json"
            preparation = self._preparation(root, plan, binder, prep_path)
            with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref), \
                 mock.patch.object(BUILDER, "_source_preparation_plan", return_value=copy.deepcopy(preparation)), \
                 mock.patch.object(BUILDER, "_source_closure", return_value={}), \
                 mock.patch.object(BUILDER, "_validate_source_closure"), \
                 mock.patch.object(BUILDER, "_quiet_guard", side_effect=lambda limit, phase: {"phase": phase, "load_1m": 0.0, "max_load": limit, "heavy_processes": [], "checked_unix_ns": time.time_ns()}), \
                 mock.patch.object(BUILDER, "_stream_command", side_effect=self._fake_stream(preparation)):
                BUILDER.prepare_source(
                    binder, target_root=Path(preparation["target_root"]),
                    build_output=Path(preparation["build_output"]), tmpdir=Path(preparation["tmpdir"]),
                    cpu=0, max_load=1000, max_iterations=1000, timeout_s=2, output=prep_path,
                )
                BUILDER.build_source(prep_path)
                final_manifest = root / "final-input.json"
                BUILDER.seal_source_build(prep_path, output_manifest=final_manifest)
            self.assertTrue(final_manifest.is_file())
            with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref):
                BUILDER.validate_build_receipt(final_manifest)

    def test_seal_rejects_producer_artifact_drift_before_final_admission(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            manifest, _, plan = self._fixture_and_plan(root)
            for endpoint in (plan["baseline"], plan["candidate"]):
                shutil.rmtree(endpoint["target_dir"])
                shutil.rmtree(endpoint["allocator_target_dir"])
            binder = root / "source-binder.json"
            prep_path = root / "source-preparation.json"
            preparation = self._preparation(root, plan, binder, prep_path)
            with mock.patch.object(BUILDER, "PRISTINE_BASELINE_REF", self.baseline_ref), \
                 mock.patch.object(BUILDER, "_source_preparation_plan", return_value=copy.deepcopy(preparation)), \
                 mock.patch.object(BUILDER, "_source_closure", return_value={}), \
                 mock.patch.object(BUILDER, "_validate_source_closure"), \
                 mock.patch.object(BUILDER, "_quiet_guard", side_effect=lambda limit, phase: {"phase": phase, "load_1m": 0.0, "max_load": limit, "heavy_processes": [], "checked_unix_ns": time.time_ns()}), \
                 mock.patch.object(BUILDER, "_stream_command", side_effect=self._fake_stream(preparation)):
                BUILDER.prepare_source(
                    binder, target_root=Path(preparation["target_root"]),
                    build_output=Path(preparation["build_output"]), tmpdir=Path(preparation["tmpdir"]),
                    cpu=0, max_load=1000, max_iterations=1000, timeout_s=2, output=prep_path,
                )
                BUILDER.build_source(prep_path)
                receipt = Path(preparation["build_output"]) / "source-build-receipt.json"
                data = json.loads(receipt.read_text(encoding="utf-8"))
                data["events"][0]["artifact"]["sha256"] = "0" * 64
                receipt.write_text(json.dumps(data), encoding="utf-8")
                with self.assertRaises(BUILDER.RunFailure):
                    BUILDER.seal_source_build(prep_path, output_manifest=root / "rejected.json")


if __name__ == "__main__":
    unittest.main()
