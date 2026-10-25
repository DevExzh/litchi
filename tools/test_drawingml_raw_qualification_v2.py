#!/usr/bin/env python3
"""Source-only consumer tests for the version-two DrawingML runner."""

from __future__ import annotations

import copy
import json
from pathlib import Path
import os
import tempfile
import subprocess
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
            with (
                mock.patch.object(policy, "_require_candidate_binding", return_value=binding),
                mock.patch.object(policy, "validate_preparation_receipt", return_value={}),
            ):
                paired.generic._validate_drawingml_raw_workload(profile)
            with self.assertRaisesRegex(paired.generic.RunnerError, "candidate policy is incomplete"):
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

                missing_route = copy.deepcopy(data)
                missing_route["route_profiles"] = []
                with self.assertRaisesRegex(paired.RunnerError, "route_profiles"):
                    paired._validate_master(missing_route)

                missing_owner = copy.deepcopy(data)
                missing_owner["owner_profiles"] = missing_owner["owner_profiles"][:-1]
                with self.assertRaisesRegex(paired.RunnerError, "owner_profiles"):
                    paired._validate_master(missing_owner)

    def test_v2_restores_per_profile_parent_and_feature_contracts(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-v2-profile-contract-") as name:
            root = Path(name)
            data, binding = self._master_fixture(root)
            with mock.patch.object(policy, "CANDIDATE_BINDING", binding):
                missing_parents = copy.deepcopy(data)
                missing_parents["profiles"][0]["approved_baseline_revisions"] = []
                with self.assertRaisesRegex(
                    paired.RunnerError, "reviewed source parents"
                ):
                    paired._validate_master(missing_parents)

                wrong_parent = copy.deepcopy(data)
                wrong_parent["profiles"][0]["approved_baseline_revisions"] = [
                    next(
                        revision
                        for revision in policy.APPROVED_BASELINE_REVISIONS
                        if revision != wrong_parent["profiles"][0]["baseline_revision"]
                    )
                ]
                with self.assertRaisesRegex(paired.RunnerError, "reviewed source parents"):
                    paired._validate_master(wrong_parent)

                missing_routes = copy.deepcopy(data)
                missing_routes["route_profiles"][0]["route_baselines"] = {
                    "theme": missing_routes["route_profiles"][0]["baseline_revision"],
                }
                with self.assertRaisesRegex(paired.RunnerError, "route_baselines"):
                    paired._validate_master(missing_routes)

                wrong_theme_parent = copy.deepcopy(data)
                wrong_theme_parent["route_profiles"][0]["route_baselines"] = {
                    "theme": wrong_theme_parent["profiles"][0]["baseline_revision"],
                    "ink": wrong_theme_parent["route_profiles"][0]["baseline_revision"],
                }
                with self.assertRaisesRegex(paired.RunnerError, "exact theme parent"):
                    paired._validate_master(wrong_theme_parent)

                family_baseline = copy.deepcopy(data)
                family_baseline["route_profiles"][0]["baseline_features"].append(
                    "drawingml-theme-family-current"
                )
                with self.assertRaisesRegex(paired.RunnerError, "legacy branch"):
                    paired._validate_master(family_baseline)

                owner_features = copy.deepcopy(data)
                owner_features["owner_profiles"][0].pop("candidate_features")
                with self.assertRaisesRegex(paired.RunnerError, "all current DrawingML features"):
                    paired._validate_master(owner_features)

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

    def test_v2_compile_requires_preparation_before_legacy_driver(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-v2-compile-gate-") as name:
            root = Path(name)
            baseline = root / "baseline"
            candidate = root / "candidate"
            trusted = root / "trusted"
            for path in (baseline, candidate, trusted):
                path.mkdir()
            arguments = type(
                "Arguments",
                (),
                {
                    "role": "candidate",
                    "preparation_receipt": str(root / "preparation.json"),
                    "baseline_root": str(baseline),
                    "candidate_root": str(candidate),
                    "baseline_revision": next(iter(policy.APPROVED_BASELINE_REVISIONS)),
                    "candidate_revision": "a" * 40,
                    "trusted_root": str(trusted),
                    "source_root": str(candidate),
                    "source_revision": "a" * 40,
                },
            )()
            with (
                mock.patch.object(
                    policy,
                    "validate_preparation_receipt",
                    side_effect=policy.PreparationError("role lock identity differs"),
                ),
                mock.patch.object(compile_driver.legacy, "_compile") as legacy_compile,
            ):
                with self.assertRaisesRegex(compile_driver.DriverError, "preparation receipt validation"):
                    compile_driver._compile(arguments)
            legacy_compile.assert_not_called()

    def test_v2_compile_receipt_consumes_full_policy_projection(self) -> None:
        with tempfile.TemporaryDirectory(prefix="drawingml-v2-compile-receipt-") as name:
            root = Path(name)
            baseline = root / "baseline"
            candidate = root / "candidate"
            trusted = root / "trusted"
            for path in (baseline, candidate, trusted):
                path.mkdir()
            preparation_receipt = root / "preparation.json"
            preparation_receipt.write_text(
                json.dumps({"schema": policy.SCHEMA, "protocol_version": 2}) + "\n",
                encoding="utf-8",
            )
            baseline_revision = next(iter(policy.APPROVED_BASELINE_REVISIONS))
            candidate_revision = "a" * 40
            prepared = {
                "baseline": {"root": str(baseline.resolve()), "head": baseline_revision},
                "candidate": {"root": str(candidate.resolve()), "head": candidate_revision},
            }
            arguments = type(
                "Arguments",
                (),
                {
                    "role": "candidate",
                    "preparation_receipt": str(preparation_receipt),
                    "baseline_root": str(baseline),
                    "candidate_root": str(candidate),
                    "baseline_revision": baseline_revision,
                    "candidate_revision": candidate_revision,
                    "trusted_root": str(trusted),
                    "source_root": str(candidate),
                    "source_revision": candidate_revision,
                    "output": str(root / "compile.json"),
                },
            )()

            def fake_compile(value: object) -> int:
                Path(arguments.output).write_text(
                    json.dumps(
                        {
                            "schema": compile_driver.SCHEMA,
                            "operation": "compile",
                            "role": arguments.role,
                            "source": {
                                "root": str(Path(arguments.source_root).resolve()),
                                "revision": arguments.source_revision,
                            },
                        }
                    )
                    + "\n",
                    encoding="utf-8",
                )
                return 0

            with (
                mock.patch.object(policy, "validate_preparation_receipt", return_value=prepared),
                mock.patch.object(compile_driver, "_policy_revision"),
                mock.patch.object(compile_driver.legacy, "_compile", side_effect=fake_compile),
            ):
                self.assertEqual(compile_driver._compile(arguments), 0)

            with (
                mock.patch.object(policy, "validate_preparation_receipt", return_value=prepared),
            ):
                consumed = compile_driver.validate_compile_receipt(
                    Path(arguments.output),
                    preparation_receipt=preparation_receipt,
                    baseline_root=baseline,
                    candidate_root=candidate,
                    baseline_revision=baseline_revision,
                    candidate_revision=candidate_revision,
                    trusted_root=trusted,
                    role="candidate",
                )
            self.assertEqual(consumed["policy"], compile_driver._receipt_policy("compile", "candidate"))

            tampered = json.loads(Path(arguments.output).read_text(encoding="utf-8"))
            tampered["policy"]["toolchain_paths"] = []
            Path(arguments.output).write_text(json.dumps(tampered) + "\n", encoding="utf-8")
            with mock.patch.object(policy, "validate_preparation_receipt", return_value=prepared):
                with self.assertRaisesRegex(compile_driver.DriverError, "policy projection"):
                    compile_driver.validate_compile_receipt(
                        Path(arguments.output),
                        preparation_receipt=preparation_receipt,
                        baseline_root=baseline,
                        candidate_root=candidate,
                        baseline_revision=baseline_revision,
                        candidate_revision=candidate_revision,
                        trusted_root=trusted,
                        role="candidate",
                    )

    def test_executable_v2_entrypoints_bootstrap_tools_and_pending_policy(self) -> None:
        clean_env = os.environ.copy()
        clean_env.pop("PYTHONPATH", None)
        scripts = (
            ROOT / "tools/native_qualification/drawingml_raw_paired_runner_v2.py",
            ROOT / "tools/native_qualification/drawingml_raw_validator_v2.py",
            ROOT / "tools/native_qualification/drawingml_compile_driver_v2.py",
        )
        for script in scripts:
            result = subprocess.run(
                [sys.executable, str(script), "--help"],
                cwd=ROOT,
                env=clean_env,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
            )
            self.assertEqual(result.returncode, 0, result.stderr)
        pending_manifest = ROOT / "tools/native_qualification/manifests/drawingml-raw-v2.json"
        with tempfile.TemporaryDirectory(prefix="drawingml-v2-cli-") as name:
            env = clean_env.copy()
            for key in (
                "DRAWINGML_CANDIDATE_ROOT",
                "DRAWINGML_CANDIDATE_REVISION",
                "DRAWINGML_TRUSTED_ROOT",
                "DRAWINGML_OUTPUT_DIR",
                "DRAWINGML_TARGET_ROOT",
                "DRAWINGML_CHART_BASELINE_ROOT",
                "DRAWINGML_DIAGRAM_BASELINE_ROOT",
                "DRAWINGML_TEXT_MEDIA_BASELINE_ROOT",
                "DRAWINGML_THEME_INK_BASELINE_ROOT",
                "DRAWINGML_GEOMETRY_BASELINE_ROOT",
                "DRAWINGML_THEME_FAMILY_BASELINE_ROOT",
                "DRAWINGML_OWNER_CHART_BASELINE_ROOT",
                "DRAWINGML_OWNER_THEME_BASELINE_ROOT",
                "DRAWINGML_OWNER_INK_BASELINE_ROOT",
                "DRAWINGML_OWNER_TRANSFORM_BASELINE_ROOT",
            ):
                env[key] = name
            for key in (
                "DRAWINGML_CHART_PREPARATION",
                "DRAWINGML_DIAGRAM_PREPARATION",
                "DRAWINGML_TEXT_MEDIA_PREPARATION",
                "DRAWINGML_THEME_INK_PREPARATION",
                "DRAWINGML_GEOMETRY_PREPARATION",
                "DRAWINGML_THEME_FAMILY_PREPARATION",
                "DRAWINGML_OWNER_CHART_PREPARATION",
                "DRAWINGML_OWNER_THEME_PREPARATION",
                "DRAWINGML_OWNER_INK_PREPARATION",
                "DRAWINGML_OWNER_TRANSFORM_PREPARATION",
            ):
                env[key] = str(Path(name) / f"{key}.json")
            result = subprocess.run(
                [sys.executable, str(scripts[0]), "--manifest", str(pending_manifest), "--phase", "build"],
                cwd=ROOT,
                env=env,
                text=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                check=False,
            )
            self.assertEqual(result.returncode, 2)
            self.assertIn("pending final candidate binding", result.stderr)


if __name__ == "__main__":
    unittest.main()
