"""Tests for the fixed W8 original-to-final projection policy."""

from __future__ import annotations

import copy
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import sys
import tempfile
import types
import unittest
from unittest import mock

from tools import w8_projection_policy as policy
from tools import w8_file_provenance as provenance
from tools import test_w8_file_provenance as provenance_tests


ANALYZER_PATH = Path(__file__).resolve().parents[1] / "docs/unified-ops/results/0362h/analyze.py"
PROBE_PATH = Path(__file__).resolve().parents[1] / "crates/litchi-ops/examples/file_w8_qualified_probe.rs"


def load_analyzer():
    spec = importlib.util.spec_from_file_location("w8_pair_analyzer", ANALYZER_PATH)
    if spec is None or spec.loader is None:
        raise AssertionError(f"cannot load {ANALYZER_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ProjectionPolicyTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.before = policy._load_fixed_projection("before")
        cls.after = policy._load_fixed_projection("after")

    def test_reviewed_policy_and_fixed_pair(self) -> None:
        value = policy.load_reviewed_policy()
        self.assertEqual(value, policy._policy_summary())
        self.assertEqual(value["before_features"], list(policy.BEFORE_FEATURES))
        self.assertEqual(value["features"], list(policy.FEATURES))
        self.assertNotEqual(policy.BEFORE_FEATURES, policy.FEATURES)
        proof = policy.validate_projection_pair(self.before, self.after)
        self.assertEqual(proof["registry_package_count"], 518)
        self.assertTrue(proof["registry_identities_checksums_dependencies_equal"])
        self.assertEqual(len(proof["local_edge_differences"]), 6)

    def test_before_feature_overlay_is_rejected(self) -> None:
        changed = copy.deepcopy(self.before)
        changed["features"]["selected"] = list(policy.FEATURES)
        with self.assertRaises(policy.ProjectionPolicyError):
            policy.validate_projection_pair(changed, self.after)

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
        with self.assertRaises(TypeError):
            fields["manifest"]._items = ()
        with self.assertRaises(TypeError):
            del fields["manifest"]._items
        with self.assertRaises(TypeError):
            fields._fields = {}
        with self.assertRaises(TypeError):
            del fields._fields
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

    def test_pair_runner_accepts_real_validated_endpoints_and_preserves_route_order(self) -> None:
        """Exercise the pair path through the runner, not only its helpers.

        The endpoint validator returns the opaque recursively-frozen capability,
        so this catches the list/tuple boundary between that capability and the
        runner's JSON-shaped route probe result.
        """
        runner_path = Path(__file__).resolve().parents[1] / "docs/unified-ops/results/0362h/run.py"
        spec = importlib.util.spec_from_file_location("w8_run_pair_integration", runner_path)
        self.assertIsNotNone(spec)
        self.assertIsNotNone(spec.loader)
        runner = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(runner)
        reviewed_routes: list[str] = []

        def fake_probe(
            _out: Path,
            _label: str,
            _binary: Path,
            route: str,
            mode: str,
            calls: int | None,
            _blocks: int,
            _cpu: int,
        ) -> dict[str, object]:
            if mode == "calibrate":
                return {
                    "route": route,
                    "common_calls": 1,
                    "minimum_elapsed_ns": runner.MIN_SAMPLE_NS,
                }
            return {"route": route, "calls": calls}

        pair = {
            "policy": {"schema": "test-policy"},
            "projection": {"registry_package_count": 518},
            "workload": {"route_count": 109},
            "build": {"profile": "release"},
        }
        verify_calls: list[tuple[object, object]] = []

        def fake_verify(
            _policy_path: Path,
            _receipt_path: Path,
            before_endpoint: object,
            after_endpoint: object,
        ) -> dict[str, object]:
            self.assertIsInstance(before_endpoint, policy._ValidatedEndpoint)
            self.assertIsInstance(after_endpoint, policy._ValidatedEndpoint)
            self.assertEqual(list(before_endpoint["routes"]), reviewed_routes)
            self.assertEqual(list(after_endpoint["routes"]), reviewed_routes)
            verify_calls.append((before_endpoint, after_endpoint))
            return pair

        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            endpoint_builder = provenance_tests.EndpointManifest()
            before_args, _before_paths, before_values = endpoint_builder.inputs(root / "before")
            after_args, _after_paths, _after_values = endpoint_builder.inputs(root / "after")
            reviewed_routes.extend(before_values["routes"])
            before_binary = before_args.binary
            after_binary = after_args.binary
            before_manifest = root / "before" / "endpoint.json"
            after_manifest = root / "after" / "endpoint.json"
            before_manifest.write_bytes(provenance.canonical(provenance.endpoint_manifest(before_args)))
            after_manifest.write_bytes(provenance.canonical(provenance.endpoint_manifest(after_args)))
            receipt = root / "pair-receipt.json"
            receipt.write_bytes(b"pair receipt")
            output = root / "qualification"
            cpu = min(os.sched_getaffinity(0))
            arguments = [
                str(before_binary),
                str(after_binary),
                str(output),
                "5",
                str(cpu),
                "--before-manifest",
                str(before_manifest),
                "--after-manifest",
                str(after_manifest),
                "--original-to-final-policy",
                str(policy.ROOT / policy.POLICY_RELATIVE),
                "--original-to-final-receipt",
                str(receipt),
            ]
            with (
                mock.patch.object(policy, "verify_pair_receipt", side_effect=fake_verify),
                mock.patch.object(runner, "guard", return_value="guard"),
                mock.patch.object(
                    runner,
                    "routes",
                    side_effect=lambda *_args: (list(reviewed_routes), {"routes": list(reviewed_routes)}),
                ),
                mock.patch.object(runner, "run_probe", side_effect=fake_probe),
                mock.patch.object(sys, "argv", [str(runner_path), *arguments]),
            ):
                self.assertEqual(runner.main(), 0)

            self.assertEqual(len(verify_calls), 1)
            manifest = json.loads((output / "manifest.json").read_text())
            self.assertEqual(manifest["status"], "complete; analyze.py review pending")
            self.assertEqual(manifest["routes"], reviewed_routes)
            self.assertEqual(manifest["provenance"]["before"]["routes"], reviewed_routes)
            self.assertEqual(manifest["provenance"]["after"]["routes"], reviewed_routes)

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


class PairAnalyzerTests(unittest.TestCase):
    def _fixture(self, root: Path) -> tuple[dict, object, dict[str, Path]]:
        analyzer = load_analyzer()
        probe = root / "probe.rs"
        probe.write_bytes(PROBE_PATH.read_bytes())
        fixture = root / "fixture.json"
        fixture.write_bytes(provenance.canonical(provenance.fixture_descriptor(probe)))
        fixture_row = provenance.artifact(fixture)
        probe_row = provenance.artifact(probe)
        routes = provenance.fixture_routes(json.loads(fixture.read_bytes()))
        host = root / "host.json"
        host.write_bytes(provenance.canonical({"blockers": [], "processes": [], "allowed": []}))
        stdout = root / "stdout"
        stderr = root / "stderr"
        stdout.write_bytes(b"probe stdout\n")
        stderr.write_bytes(b"")
        host_row = provenance.artifact(host)
        stdout_row = provenance.artifact(stdout)
        stderr_row = provenance.artifact(stderr)

        binary_rows = []
        endpoint_values = []
        fixed_inputs = {}
        for side, commit in (("before", "a" * 40), ("after", "b" * 40)):
            side_root = root / side
            side_root.mkdir()
            binary = side_root / "probe"
            binary.write_bytes(f"{side} binary\n".encode())
            os.chmod(binary, 0o755)
            source = side_root / "source.tar"
            lock = side_root / "Cargo.lock"
            workspace = side_root / "Cargo.toml"
            package = side_root / "litchi-ops.Cargo.toml"
            toolchain = side_root / "rust-toolchain.toml"
            source.write_bytes(f"{side} source\n".encode())
            lock.write_bytes(f"{side} lock\n".encode())
            workspace.write_bytes(f"{side} workspace\n".encode())
            package.write_bytes(f"{side} package\n".encode())
            toolchain.write_bytes(f"{side} toolchain\n".encode())
            projection = side_root / "projection.json"
            projection.write_bytes(
                provenance.canonical(
                    {"schema": "litchi_w8_dependency_projection.v1", "side": side}
                )
            )
            source_assembly = side_root / "source-assembly.json"
            build = side_root / "build.json"
            validation = side_root / "validation.json"
            source_assembly.write_bytes(provenance.canonical({"side": side, "kind": "none"}))
            build.write_bytes(provenance.canonical({"side": side, "status": "complete"}))
            validation.write_bytes(provenance.canonical({"side": side, "status": "complete"}))
            artifacts = {
                "binary": provenance.artifact(binary),
                "source": provenance.artifact(source),
                "lock": provenance.artifact(lock),
                "probe": probe_row,
                "fixture": fixture_row,
                "dependency_projection": provenance.artifact(projection),
                "workspace_manifest": provenance.artifact(workspace),
                "package_manifest": provenance.artifact(package),
                "toolchain": provenance.artifact(toolchain),
                "source_assembly": provenance.artifact(source_assembly),
                "build_receipt": provenance.artifact(build),
                "validation": provenance.artifact(validation),
            }
            endpoint_path = side_root / "endpoint.json"
            endpoint_value = {
                "schema": "litchi_file_endpoint.v1",
                "source_commit": commit,
                "dependency_projection_sha256": artifacts["dependency_projection"]["sha256"],
                **artifacts,
                "features": {"package": "litchi-ops", "selected": ["all"], "all_features": True},
                "target": "test-target",
                "profile": "release",
            }
            endpoint_path.write_bytes(provenance.canonical(endpoint_value))
            endpoint_value["manifest"] = provenance.artifact(endpoint_path)
            endpoint_value["routes"] = routes
            endpoint_values.append(endpoint_value)
            binary_rows.append(artifacts["binary"])
            fixed_inputs[side] = {
                "source_commit": commit,
                "source": artifacts["source"],
                "workspace_manifest": artifacts["workspace_manifest"],
                "package_manifest": artifacts["package_manifest"],
                "toolchain": artifacts["toolchain"],
                "lock": artifacts["lock"],
                "projection": artifacts["dependency_projection"],
            }

        proof = {
            "registry_package_count": 0,
            "registry_identities_checksums_dependencies_equal": True,
            "local_edge_differences": [],
        }
        policy_identity = {
            "schema": "test.original_to_final_policy.v1",
            "version": "test-v1",
            "sha256": "c" * 64,
        }
        fake_policy = types.SimpleNamespace(
            PAIR_SCHEMA="test.original_to_final_pair.v1",
            REVIEWED_POLICY_SHA256=policy_identity["sha256"],
            FIXED_INPUTS=fixed_inputs,
            PROBE_BYTES=probe_row["bytes"],
            PROBE_SHA256=probe_row["sha256"],
            FIXTURE_BYTES=fixture_row["bytes"],
            FIXTURE_SHA256=fixture_row["sha256"],
            TARGET="test-target",
            PROFILE="release",
            BEFORE_FEATURES=("all",),
            FEATURES=("all",),
            TOOLCHAIN_SHA256=fixed_inputs["before"]["toolchain"]["sha256"],
            load_reviewed_policy=lambda: {
                "schema": policy_identity["schema"],
                "policy_version": policy_identity["version"],
            },
            validate_projection_pair=lambda _before, _after: proof,
        )
        before_endpoint, after_endpoint = endpoint_values
        workload = {
            "probe_sha256": probe_row["sha256"],
            "fixture_sha256": fixture_row["sha256"],
            "route_count": len(routes),
            "routes": routes,
        }
        build = {
            "target": "test-target",
            "profile": "release",
            "features": {"before": ["all"], "after": ["all"]},
            "toolchain_sha256": fake_policy.TOOLCHAIN_SHA256,
        }
        receipt_value = {
            "schema": fake_policy.PAIR_SCHEMA,
            "policy": policy_identity,
            "before": analyzer._pair_endpoint_part(before_endpoint, "before"),
            "after": analyzer._pair_endpoint_part(after_endpoint, "after"),
            "workload": workload,
            "build": build,
            "projection": proof,
            "status": "verified_protocol_only",
        }
        receipt_path = root / "pair-receipt.json"
        receipt_path.write_bytes(provenance.canonical(receipt_value))
        shared = {
            "mode": "original_to_final",
            "policy": policy_identity,
            "receipt": provenance.artifact(receipt_path),
            "projection": proof,
            "workload": workload,
            "build": build,
        }
        return shared, fake_policy, {
            "before": before_endpoint,
            "after": after_endpoint,
            "routes": routes,
            "binary_before": Path(binary_rows[0]["path"]),
            "binary_after": Path(binary_rows[1]["path"]),
            "source_before": Path(before_endpoint["source"]["path"]),
            "projection_before": Path(before_endpoint["dependency_projection"]["path"]),
            "lock_before": Path(before_endpoint["lock"]["path"]),
            "host": host_row,
            "stdout": stdout_row,
            "stderr": stderr_row,
        }

    @staticmethod
    def _phase() -> dict:
        return {"samples": [150_000_000], "calls": 1, "total_ns": 150_000_000, "operation_count": 1}

    def _write_analysis_manifest(
        self,
        root: Path,
        shared: dict,
        endpoint_values: dict,
    ) -> Path:
        changed_suffixes = {
            "set_commit", "set_publication", "batch_commit", "batch_publication",
            "import_commit", "import_publication", "merge_commit", "merge_publication",
        }
        routes = endpoint_values["routes"]
        before_binary = provenance.artifact(endpoint_values["binary_before"])
        after_binary = provenance.artifact(endpoint_values["binary_after"])
        receipt_root = root / "receipts"
        receipt_root.mkdir()
        unique = 0

        def probe_receipt(side: str, route: str, mode: str, calls: int | None) -> dict:
            nonlocal unique
            unique += 1
            binary = before_binary if side == "before" else after_binary
            label = f"synthetic-{unique}"
            path = receipt_root / f"{label}.receipt.json"
            command = [binary["path"], "--mode", mode, "--route", route]
            receipt = {
                "schema": "litchi_file_probe_receipt.v2",
                "label": label,
                "binary": binary,
                "route": route,
                "mode": mode,
                "calls": calls,
                "command": command,
                "stdout": endpoint_values["stdout"],
                "stderr": endpoint_values["stderr"],
                "exit_code": 0,
                "overlap": None,
                "guard_error": None,
                "host_snapshots": [endpoint_values["host"]],
            }
            path.write_bytes(provenance.canonical(receipt))
            return {"_receipt": str(path), "_host_snapshots": [endpoint_values["host"]]}

        def semantic_value(route: str) -> dict:
            suffix = route.rsplit("/", 1)[-1]
            changed = suffix in changed_suffixes
            output = {
                "bytes": 1,
                "sha256": "2" * 64,
                "source_sha256": "1" * 64,
            }
            validation = {
                "semantic_state_valid": True,
                "source_state_preserved": True,
                "semantic_equal": True,
                "source_state_exact": True,
                "leaf_repeated_equal": True,
                "facade_repeated_equal": True,
                "exact_noop": True,
                "leaf_changed": changed,
                "facade_changed": changed,
                "leaf_output_changed": changed,
                "facade_output_changed": changed,
                "source_leaf_output": output,
                "source_facade_output": output,
                "leaf_output": output,
                "facade_output": output,
                "leaf_repeat": output,
                "facade_repeat": output,
                "source_leaf_before_sha256": "3" * 64,
                "source_leaf_after_sha256": "3" * 64,
                "source_facade_before_sha256": "4" * 64,
                "source_facade_after_sha256": "4" * 64,
            }
            if suffix == "open_get":
                validation.update({
                    "read_marker_equal": True,
                    "leaf_read_marker_len": 1,
                    "facade_read_marker_len": 1,
                })
            return validation

        route_lists = {
            "before": probe_receipt("before", "*", "list", None),
            "after": probe_receipt("after", "*", "list", None),
        }
        calibration = []
        for route in routes:
            rows = []
            for side in ("before", "after"):
                rows.append({
                    "common_calls": 1,
                    "minimum_elapsed_ns": 150_000_000,
                    **probe_receipt(side, route, "calibrate", None),
                })
            calibration.append({"route": route, "common_calls": 1, "rows": rows})

        abba = []
        for route in routes:
            blocks = []
            for block_index in range(5):
                legs = []
                for side in ("before", "after", "after", "before"):
                    value = {
                        "route": route,
                        "calls": 1,
                        **probe_receipt(side, route, "measure", 1),
                        "validation": semantic_value(route),
                        "blocks": [{
                            "order": ["leaf", "facade", "facade", "leaf"],
                            "leaf": [self._phase(), self._phase()],
                            "facade": [self._phase(), self._phase()],
                        }],
                    }
                    legs.append({"side": side, "value": value})
                blocks.append({"block": block_index, "legs": legs})
            abba.append({"route": route, "calls": 1, "blocks": blocks})

        manifest = {
            "schema": "litchi_file_qualification.v3",
            "before": before_binary,
            "after": after_binary,
            "routes": routes,
            "route_lists": route_lists,
            "rounds": 5,
            "target_ns": 500_000_000,
            "minimum_sample_ns": 150_000_000,
            "provenance": {
                "before": endpoint_values["before"],
                "after": endpoint_values["after"],
                "shared": shared,
            },
            "calibration": calibration,
            "abba": abba,
        }
        output = root / "analysis-output"
        output.mkdir()
        (output / "manifest.json").write_bytes(provenance.canonical(manifest))
        return output

    def test_pair_analysis_accepts_exact_record_and_rejects_source_projection_lock_tampering(self) -> None:
        analyzer = load_analyzer()
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            shared, fake_policy, values = self._fixture(root)
            output = self._write_analysis_manifest(root, shared, values)
            policy_patch = mock.patch.object(analyzer, "_pair_policy_module", return_value=fake_policy)
            with policy_patch, mock.patch.object(
                sys, "argv", [str(ANALYZER_PATH), str(output)]
            ):
                analyzer.main()
            self.assertTrue((output / "analysis.json").is_file())

            for label, path, payload in (
                ("source", values["source_before"], b"tampered source\n"),
                ("projection", values["projection_before"], b"tampered projection\n"),
                ("lock", values["lock_before"], b"tampered lock\n"),
            ):
                with self.subTest(label=label):
                    original = path.read_bytes()
                    path.write_bytes(payload)
                    try:
                        with policy_patch, mock.patch.object(
                            sys, "argv", [str(ANALYZER_PATH), str(output)]
                        ), self.assertRaises(SystemExit):
                            analyzer.main()
                    finally:
                        path.write_bytes(original)


if __name__ == "__main__":
    unittest.main()
