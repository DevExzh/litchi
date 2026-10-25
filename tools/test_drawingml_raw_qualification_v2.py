#!/usr/bin/env python3
"""Source-only consumer tests for the version-two DrawingML runner."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import tempfile
import unittest
from unittest import mock

import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

from native_qualification import drawingml_raw_paired_runner_v2 as paired  # noqa: E402
from native_qualification import drawingml_raw_validator_v2 as validator  # noqa: E402
from native_qualification import drawingml_compile_driver_v2 as compile_driver  # noqa: E402
from native_qualification import prepare_drawingml_raw_inputs_v2 as policy  # noqa: E402


class DrawingMlRawQualificationV2Tests(unittest.TestCase):
    def _master_fixture(self, root: Path) -> tuple[dict, dict]:
        manifest = json.loads(
            (ROOT / "tools/native_qualification/manifests/drawingml-raw-v2.json").read_text(
                encoding="utf-8"
            )
        )
        candidate = root / "candidate"
        trusted = root / "trusted"
        output = root / "output"
        target = root / "target"
        for path in (candidate, trusted, output, target):
            path.mkdir()
        fixture = candidate / "tools/native_qualification/fixtures/drawingml_raw_fixtures.json"
        fixture.parent.mkdir(parents=True)
        fixture.write_bytes(
            (ROOT / "tools/native_qualification/fixtures/drawingml_raw_fixtures.json").read_bytes()
        )
        candidate_revision = "a" * 40
        binding = {
            "status": "bound",
            "revision": candidate_revision,
            "manifests": {
                path: {"sha256": "b" * 64, "bytes": 1}
                for path in policy.CANDIDATE_MANIFEST_PATHS
            },
            "vendor_tree": {"sha256": "c" * 64, "files": 3, "bytes": 3},
        }
        manifest.update(
            {
                "status": "ready",
                "candidate_root": str(candidate),
                "candidate_revision": candidate_revision,
                "trusted_root": str(trusted),
                "output_dir": str(output),
                "target_root": str(target),
                "candidate": binding,
            }
        )
        for profile in [*manifest["profiles"], *manifest["route_profiles"], *manifest["owner_profiles"]]:
            baseline = root / profile["id"] / "baseline"
            baseline.mkdir(parents=True)
            receipt = root / profile["id"] / "receipt.json"
            receipt.write_text("{}\n", encoding="utf-8")
            profile["baseline_root"] = str(baseline)
            profile["preparation_receipt"] = str(receipt)
        return manifest, binding

    def test_valid_v2_master_and_profile_preserve_protocol_routes(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-v2-master-") as name:
            data, binding = self._master_fixture(Path(name))
            with mock.patch.object(policy, "CANDIDATE_BINDING", binding):
                master = paired._validate_master(data)
                profile = paired._profile_manifest(
                    master,
                    master["profiles"][0],
                    master["output_dir"],
                )
            self.assertEqual(master["candidate_revision"], binding["revision"])
            self.assertEqual(profile["drawingml_protocol_version"], 2)
            self.assertEqual(profile["shared"]["probe_paths"], list(policy.PROBE_COPY_PATHS))
            self.assertEqual(profile["legs"]["before"]["lock_paths"], list(policy.ROLE_LOCK_PATHS))
            paired.generic._validate_drawingml_raw_workload(profile)

    def test_candidate_identity_cannot_be_overridden_by_manifest(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-v2-identity-") as name:
            data, binding = self._master_fixture(Path(name))
            data["candidate_revision"] = "d" * 40
            with mock.patch.object(policy, "CANDIDATE_BINDING", binding):
                with self.assertRaisesRegex(paired.RunnerError, "candidate revision"):
                    paired._validate_master(data)

    def test_wrong_role_lock_and_owner_route_are_rejected(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-v2-routes-") as name:
            data, binding = self._master_fixture(Path(name))
            with mock.patch.object(policy, "CANDIDATE_BINDING", binding):
                bad_lock = copy.deepcopy(data)
                bad_lock["candidate_native_lock_sha256"] = "d" * 64
                with self.assertRaisesRegex(paired.RunnerError, "role lock"):
                    paired._validate_master(bad_lock)

                bad_route = copy.deepcopy(data)
                bad_route["owner_profiles"][0]["owner_route"] = "drawingml-transform"
                with self.assertRaisesRegex(paired.RunnerError, "owner route"):
                    paired._validate_master(bad_route)

    def test_pending_manifest_fails_before_any_source_path_is_used(self) -> None:
        manifest = json.loads(
            (ROOT / "tools/native_qualification/manifests/drawingml-raw-v2.json").read_text(
                encoding="utf-8"
            )
        )
        with self.assertRaisesRegex(paired.RunnerError, "pending final candidate binding"):
            paired._validate_master(manifest)

    def test_v2_validator_validates_receipt_then_preserves_probe_schema(self) -> None:
        # Reuse the existing unavailable-report fixture generator: this test
        # exercises the complete Python validator path without Cargo/native IO.
        from tools.test_drawingml_raw_qualification import DrawingMlRawProtocolTests

        with tempfile.TemporaryDirectory(prefix="drawingml-v2-validator-") as name:
            output = Path(name)
            DrawingMlRawProtocolTests("test_unavailable_reports_are_preserved_without_acceptance_claim")._write_unavailable_reports(output)
            arguments = type(
                "Arguments",
                (),
                {
                    "output_dir": output,
                    "lane": "drawingml-chart-raw",
                    "endpoint": "raw",
                    "fixture": ROOT / "tools/native_qualification/fixtures/drawingml_raw_fixtures.json",
                    "baseline_revision": "a" * 40,
                    "candidate_revision": "b" * 40,
                    "source_fence": "v2 test fence",
                    "preparation_receipt": output / "receipt.json",
                    "baseline_root": output / "baseline",
                    "candidate_root": output / "candidate",
                    "trusted_root": output / "trusted",
                },
            )()
            receipt = {"baseline_revision": "a" * 40, "candidate_revision": "b" * 40}
            with mock.patch.object(validator.preparation, "validate_preparation_receipt", return_value=receipt):
                summary = validator.validate(arguments)
            self.assertEqual(summary["schema"], validator.SUMMARY_SCHEMA)
            self.assertEqual(summary["protocol_version"], 2)
            self.assertEqual(summary["preparation_receipt"]["schema"], policy.SCHEMA)

    def test_v2_validator_rejects_wrong_receipt_identity(self) -> None:
        arguments = type(
            "Arguments",
            (),
            {
                "preparation_receipt": Path("/tmp/receipt.json"),
                "baseline_root": Path("/tmp/baseline"),
                "candidate_root": Path("/tmp/candidate"),
                "trusted_root": Path("/tmp/trusted"),
                "baseline_revision": "a" * 40,
                "candidate_revision": "b" * 40,
            },
        )()
        with mock.patch.object(
            validator.preparation,
            "validate_preparation_receipt",
            side_effect=ValueError("candidate source identity changed"),
        ):
            with self.assertRaisesRegex(validator.ValidationError, "candidate source identity"):
                validator._validate_preparation(arguments)

    def test_v2_compile_adapter_keeps_candidate_identity_fail_closed(self) -> None:
        with self.assertRaisesRegex(compile_driver.DriverError, "not execution-ready"):
            compile_driver._policy_revision("a" * 40, "candidate")
        # A reviewed historical fence is admissible for the baseline compile
        # path without pretending that it is the final current candidate.
        compile_driver._policy_revision(next(iter(policy.APPROVED_BASELINE_REVISIONS)), "baseline")


if __name__ == "__main__":
    unittest.main()
