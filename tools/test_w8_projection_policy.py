"""Tests for the fixed W8 original-to-final projection policy."""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from tools import w8_projection_policy as policy


class ProjectionPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.before = policy._load_fixed_projection("before")
        cls.after = policy._load_fixed_projection("after")

    def test_reviewed_policy_and_fixed_pair(self) -> None:
        value = policy.load_reviewed_policy()
        self.assertEqual(value, policy._policy_summary())
        proof = policy.validate_projection_pair(self.before, self.after)
        self.assertEqual(proof["registry_package_count"], 518)
        self.assertTrue(proof["registry_identities_checksums_dependencies_equal"])
        self.assertEqual(len(proof["local_edge_differences"]), 6)

    def test_policy_bytes_are_registered_and_canonical(self) -> None:
        path = policy.ROOT / policy.POLICY_RELATIVE
        payload = path.read_bytes()
        self.assertEqual(hashlib.sha256(payload).hexdigest(), policy.REVIEWED_POLICY_SHA256)
        self.assertEqual(payload, policy.canonical(json.loads(payload)))

    def test_registry_checksum_tamper_is_rejected(self) -> None:
        changed = copy.deepcopy(self.after)
        package = next(item for item in changed["lock"]["packages"] if item.get("source", "").startswith("registry+"))
        package["checksum"] = "0" * 64
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.validate_projection_pair(self.before, changed)

    def test_registry_dependency_tamper_is_rejected(self) -> None:
        changed = copy.deepcopy(self.after)
        package = next(item for item in changed["lock"]["packages"] if item.get("source", "").startswith("registry+"))
        package["dependencies"] = list(package["dependencies"]) + ["invented-edge"]
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.validate_projection_pair(self.before, changed)

    def test_local_delta_tamper_is_rejected(self) -> None:
        changed = copy.deepcopy(self.after)
        package = next(item for item in changed["lock"]["packages"] if item["name"] == "litchi-xlsx")
        package["dependencies"] = list(package["dependencies"]) + ["invented-edge"]
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.validate_projection_pair(self.before, changed)

    def test_feature_and_workload_tamper_are_rejected(self) -> None:
        for field, value in (("features", {"package": "litchi-ops", "selected": ["default"], "all_features": True}), ("workload", {"route_count": 108})):
            changed = copy.deepcopy(self.after)
            changed[field] = value
            with self.subTest(field=field), self.assertRaises(policy.ProjectionPolicyError):
                policy.validate_projection_pair(self.before, changed)

    def test_generic_policy_path_cannot_be_replaced(self) -> None:
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.load_reviewed_policy(Path("/tmp/another-policy.json"))

    def test_generic_runner_identity_contract_remains_strict(self) -> None:
        path = Path(__file__).resolve().parents[1] / "docs/unified-ops/results/0362h/run.py"
        spec = importlib.util.spec_from_file_location("w8_run_contract", path)
        self.assertIsNotNone(spec)
        self.assertIsNotNone(spec.loader)
        runner = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(runner)
        before = {
            "dependency_projection_sha256": "a",
            "lock": {"sha256": "b"},
            "probe": {"sha256": "c"},
            "fixture": {"sha256": "d"},
        }
        self.assertEqual(runner.shared_identity(before, copy.deepcopy(before)), {
            "dependency_projection_sha256": "a",
            "lock_sha256": "b",
            "probe_sha256": "c",
            "fixture_sha256": "d",
        })
        changed = copy.deepcopy(before)
        changed["lock"]["sha256"] = "changed"
        with self.assertRaises(RuntimeError):
            runner.shared_identity(before, changed)

    def test_pair_receipt_requires_opaque_full_validation_result(self) -> None:
        with self.assertRaises(TypeError):
            policy._ValidatedEndpoint({}, object())
        with tempfile.TemporaryDirectory() as directory:
            receipt = Path(directory) / "receipt.json"
            receipt.write_bytes(policy.canonical({}))
            with self.assertRaises(policy.ProjectionPolicyError):
                policy.verify_pair_receipt(
                    policy.ROOT / policy.POLICY_RELATIVE, receipt, {}, {}
                )

    def test_validated_endpoint_cannot_replace_or_mutate_nested_fields(self) -> None:
        fields = policy._make_validated_endpoint({
            "manifest": {"path": "/original/endpoint.json"},
            "binary": {"path": "/original/binary"},
        })
        with self.assertRaises(TypeError):
            fields["manifest"] = {"path": "/forged/endpoint.json"}
        with self.assertRaises(TypeError):
            fields["manifest"]["path"] = "/forged/endpoint.json"
        detached = policy.endpoint_public_fields(fields)
        detached["manifest"]["path"] = "/detached/copy.json"
        self.assertEqual(fields["manifest"]["path"], "/original/endpoint.json")

    def test_validator_factory_result_retains_original_manifest_identity(self) -> None:
        class FakeRunner:
            @staticmethod
            def endpoint_manifest(path: Path, binary: Path, label: str) -> dict[str, object]:
                return {"manifest": {"path": str(path)}, "routes": ["docx/open_get"]}

        with mock.patch.object(policy, "_runner_module", return_value=FakeRunner()):
            fields = policy.validate_endpoint_manifest(Path("original.json"), Path("original"), "before")
        with self.assertRaises(TypeError):
            fields["manifest"] = {"path": "forged.json"}
        self.assertEqual(fields["manifest"]["path"], str(Path("original.json").resolve()))

    def test_pair_receipt_binds_metadata_and_route_order(self) -> None:
        expected = {
            "schema": policy.PAIR_SCHEMA,
            "policy": {
                "schema": policy.POLICY_SCHEMA,
                "version": policy.POLICY_VERSION,
                "sha256": policy.REVIEWED_POLICY_SHA256,
            },
            "workload": {"routes": ["docx/open_get"], "route_count": 1},
            "status": "verified_protocol_only",
        }
        before = policy._make_validated_endpoint({})
        after = policy._make_validated_endpoint({})
        with tempfile.TemporaryDirectory() as directory:
            receipt = Path(directory) / "receipt.json"
            receipt.write_bytes(policy.canonical(expected))
            with mock.patch.object(policy, "_receipt_from_validated", return_value=expected):
                self.assertEqual(
                    policy.verify_pair_receipt(
                        policy.ROOT / policy.POLICY_RELATIVE, receipt, before, after
                    ),
                    expected,
                )
                tampered = dict(expected)
                tampered["workload"] = {"routes": ["docx/save"], "route_count": 1}
                receipt.write_bytes(policy.canonical(tampered))
                with self.assertRaises(policy.ProjectionPolicyError):
                    policy.verify_pair_receipt(
                        policy.ROOT / policy.POLICY_RELATIVE, receipt, before, after
                    )

    def test_standalone_pair_uses_complete_endpoint_validator(self) -> None:
        class FakeRunner:
            def __init__(self, message: str) -> None:
                self.message = message

            def endpoint_manifest(self, path: Path, binary: Path, label: str) -> dict[str, object]:
                self.args = (path, binary, label)
                raise RuntimeError(self.message)

        for message in ("source assembly patch artifact identity is stale", "source archive identity is stale"):
            with self.subTest(message=message):
                fake = FakeRunner(message)
                with mock.patch.object(policy, "_runner_module", return_value=fake):
                    with self.assertRaises(policy.ProjectionPolicyError) as raised:
                        policy.validate_endpoint_manifest(Path("before.json"), Path("before"), "before")
                self.assertIn("complete W8 validator", str(raised.exception))

    def test_source_assembly_patch_file_tamper_is_rejected(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)

            def write(name: str, payload: bytes) -> dict[str, object]:
                path = root / name
                path.write_bytes(payload)
                return policy.artifact(path)

            source = write("source.tar", b"source archive")
            alternate_source = write("alternate-source.tar", b"source archive")
            lock = write("Cargo.lock", b"lock")
            probe = write("probe.rs", b"probe")
            fixture = write("fixture.json", b"fixture")
            workspace = write("Cargo.toml", b"workspace")
            package = write("litchi-ops.Cargo.toml", b"package")
            toolchain = write("rust-toolchain.toml", b"toolchain")
            projection = write("projection.json", b"projection")
            binary_path = root / "endpoint"
            binary_path.write_bytes(b"binary")
            os.chmod(binary_path, 0o755)
            binary = policy.artifact(binary_path)
            build = write("build.json", b"build receipt")
            validation = write("validation.json", b"validation receipt")
            patch_path = root / "adapter.patch"
            patch = write("adapter.patch", b"adapter")
            assembly = {
                "schema": "litchi_w8_source_assembly.v1",
                "source_commit": policy.BASELINE_COMMIT,
                "resolved_commit": policy.BASELINE_COMMIT,
                "source_archive": source,
                "overlay": {
                    "kind": "compatibility",
                    "description": "test adapter",
                    "base_commit": policy.BASELINE_COMMIT,
                    "patch": patch,
                },
            }
            assembly_path = root / "source-assembly.json"
            assembly_path.write_bytes(policy.canonical(assembly))
            assembly_artifact = policy.artifact(assembly_path)
            manifest_path = root / "endpoint.json"
            manifest_path.write_bytes(policy.canonical({"schema": "litchi_file_endpoint.v1"}))
            endpoint = {
                "schema": "litchi_file_endpoint.v1",
                "source_commit": policy.BASELINE_COMMIT,
                "dependency_projection_sha256": projection["sha256"],
                "source": source,
                "lock": lock,
                "probe": probe,
                "fixture": fixture,
                "dependency_projection": projection,
                "workspace_manifest": workspace,
                "package_manifest": package,
                "toolchain": toolchain,
                "source_assembly": assembly_artifact,
                "binary": binary,
                "build_receipt": build,
                "validation": validation,
            }
            fields = policy._make_validated_endpoint({
                "manifest": policy.artifact(manifest_path),
                "source_commit": policy.BASELINE_COMMIT,
                "dependency_projection_sha256": projection["sha256"],
                "source": source,
                "binary": binary,
                "lock": lock,
                "probe": probe,
                "fixture": fixture,
                "workspace_manifest": workspace,
                "package_manifest": package,
                "toolchain": toolchain,
                "routes": ["docx/open_get"],
            })
            fixed = copy.deepcopy(policy.FIXED_INPUTS)
            fixed["before"].update({
                "source": source,
                "lock": lock,
                "workspace_manifest": workspace,
                "package_manifest": package,
                "toolchain": toolchain,
                "projection": projection,
                "source_assembly": {
                    "overlay_kind": "compatibility",
                    "overlay_base_commit": policy.BASELINE_COMMIT,
                    "patch_bytes": patch["bytes"],
                    "patch_sha256": patch["sha256"],
                },
            })
            with mock.patch.object(policy, "FIXED_INPUTS", fixed), \
                mock.patch.object(policy, "PROBE_BYTES", probe["bytes"]), \
                mock.patch.object(policy, "PROBE_SHA256", probe["sha256"]), \
                mock.patch.object(policy, "FIXTURE_BYTES", fixture["bytes"]), \
                mock.patch.object(policy, "FIXTURE_SHA256", fixture["sha256"]):
                policy._check_fields_fixed(endpoint, fields, "before")
                original_source_archive = dict(assembly["source_archive"])
                assembly["source_archive"] = dict(alternate_source)
                assembly_path.write_bytes(policy.canonical(assembly))
                endpoint["source_assembly"] = policy.artifact(assembly_path)
                with self.assertRaises(policy.ProjectionPolicyError):
                    policy._check_fields_fixed(endpoint, fields, "before")
                assembly["source_archive"] = original_source_archive
                assembly_path.write_bytes(policy.canonical(assembly))
                endpoint["source_assembly"] = policy.artifact(assembly_path)
                for path, payload in ((binary_path, b"mutated binary"), (root / "build.json", b"mutated build"), (root / "validation.json", b"mutated validation")):
                    path.write_bytes(payload)
                    with self.assertRaises(policy.ProjectionPolicyError):
                        policy._check_fields_fixed(endpoint, fields, "before")
                    if path == binary_path:
                        path.write_bytes(b"binary")
                        os.chmod(path, 0o755)
                    elif path.name == "build.json":
                        path.write_bytes(b"build receipt")
                    else:
                        path.write_bytes(b"validation receipt")
                patch_path.write_bytes(b"tampered adapter")
                with self.assertRaises(policy.ProjectionPolicyError):
                    policy._check_fields_fixed(endpoint, fields, "before")
                patch_path.write_bytes(b"adapter")
                source_path = root / "source.tar"
                source_path.write_bytes(b"tampered source archive")
                with self.assertRaises(policy.ProjectionPolicyError):
                    policy._check_fields_fixed(endpoint, fields, "before")
                source_path.write_bytes(b"source archive")
                source_path.unlink()
                with self.assertRaises(policy.ProjectionPolicyError):
                    policy._check_fields_fixed(endpoint, fields, "before")
                manifest_path.write_bytes(b"tampered endpoint manifest")
                with self.assertRaises(policy.ProjectionPolicyError):
                    policy._check_fields_fixed(endpoint, fields, "before")


if __name__ == "__main__":
    unittest.main()
