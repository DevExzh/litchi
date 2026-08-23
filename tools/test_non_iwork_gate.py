from __future__ import annotations

import copy
import re
import subprocess
import tomllib
import unittest
from pathlib import Path
from unittest import mock

from tools import non_iwork_gate as gate


ROOT = Path(__file__).resolve().parents[1]


def metadata_fixture() -> dict[str, object]:
    """Build a topology-sized metadata fixture without running Cargo."""

    facade_manifest = tomllib.loads(
        (ROOT / "crates/litchi/Cargo.toml").read_text(encoding="utf-8")
    )
    feature_values = {
        name: list(values)
        for name, values in facade_manifest["features"].items()
    }
    names = [
        gate.FACADE_PACKAGE,
        *sorted(gate.EXPECTED_EXCLUDED_PACKAGE_NAMES),
        *sorted(gate.EXPECTED_BULK_PACKAGE_NAMES),
    ]
    packages: list[dict[str, object]] = []
    workspace_members: list[str] = []
    for name in names:
        package_id = f"path+file:///fixture/{name}#0.0.1"
        workspace_members.append(package_id)
        raw: dict[str, object] = {
            "name": name,
            "id": package_id,
            "features": feature_values if name == "litchi" else {},
            "dependencies": [],
        }
        if name == "litchi-py":
            raw["dependencies"] = [
                {
                    "name": "litchi",
                    "features": [
                        "doc",
                        "ooxml",
                        "odf",
                        "iwork",
                        "rtf",
                        "formula",
                        "images",
                    ],
                }
            ]
        packages.append(raw)
    return {"packages": packages, "workspace_members": workspace_members}


def bulk_tree_output(plan: gate.WorkspacePlan, suffix: str = "") -> str:
    """Return a complete synthetic workspace tree root listing."""

    return "".join(f"{name} v0.0.1\n" for name in sorted(plan.bulk_packages)) + suffix


class NonIworkGateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.plan = gate.derive_plan(metadata_fixture())

    def test_exact_package_partition_and_feature_closure(self) -> None:
        self.assertEqual(len(self.plan.packages), 64)
        self.assertEqual(len(self.plan.iwork_packages), 17)
        self.assertEqual(self.plan.iwork_packages, gate.EXPECTED_IWORK_PACKAGE_NAMES)
        self.assertEqual(len(self.plan.unsafe_facade_dependents), 1)
        self.assertEqual(self.plan.unsafe_facade_dependents, {"litchi-py"})
        self.assertEqual(len(self.plan.excluded_packages), 18)
        self.assertEqual(
            self.plan.excluded_packages, gate.EXPECTED_EXCLUDED_PACKAGE_NAMES
        )
        self.assertEqual(len(self.plan.bulk_packages), 45)
        self.assertEqual(self.plan.bulk_packages, gate.EXPECTED_BULK_PACKAGE_NAMES)
        self.assertEqual(
            self.plan.unsafe_facade_features,
            {
                "all",
                "all-formats",
                "iwork",
                "keynote",
                "numbers",
                "office",
                "pages",
                "slides",
                "spreadsheets",
                "word",
            },
        )
        self.assertEqual(
            self.plan.safe_facade_features
            | self.plan.unsafe_facade_features,
            set(self.plan.facade_features),
        )

    def test_iwork_feature_aliases_are_not_safe(self) -> None:
        aliases = {
            "word": ["pages"],
            "slides": ["keynote"],
            "spreadsheets": ["numbers"],
            "office": ["iwork"],
        }
        values = {
            "pages": ["dep:litchi-pages"],
            "keynote": ["dep:litchi-keynote"],
            "numbers": ["dep:litchi-numbers"],
            "iwork": ["pages", "keynote", "numbers"],
            **aliases,
        }
        self.assertEqual(
            gate._unsafe_features(
                values,
                frozenset({"litchi-pages", "litchi-keynote", "litchi-numbers"}),
            ),
            frozenset(values),
        )

    def test_tree_parser_and_protobuf_family(self) -> None:
        names = gate.parse_tree_package_names(
            "litchi v0.0.1 (/fixture)\n"
            "litchi-core v0.0.1 (/fixture)\n"
            "prost v0.14.0\n"
            "prost-derive v0.14.0 (proc-macro)\n"
            "prost_types v0.14.0\n"
            "prost_derive v0.14.0 (proc-macro)\n"
            "protobuf_codegen v0.14.0\n"
            "serde v1.0.0 (*)\n"
        )
        self.assertEqual(
            names,
            {
                "litchi",
                "litchi-core",
                "prost",
                "prost-derive",
                "prost_types",
                "prost_derive",
                "protobuf_codegen",
                "serde",
            },
        )
        self.assertEqual(
            gate._protobuf_packages(names),
            {"prost", "prost-derive", "prost_types", "prost_derive", "protobuf_codegen"},
        )

    def test_command_generation_is_argv_safe_and_excludes_all_non_bulk_roots(self) -> None:
        specs = gate.command_specs("cargo", self.plan, "clippy")
        self.assertEqual(len(specs), 2)
        for spec in specs:
            self.assertIsInstance(spec.argv, tuple)
            self.assertNotIn("--locked", spec.argv)
            self.assertIn("--no-deps", spec.argv)
            self.assertNotIn("--tests", spec.argv)
            self.assertLess(spec.argv.index("--no-deps"), spec.argv.index("--"))
            self.assertNotIn("'", " ".join(spec.argv))
        bulk = specs[0].argv
        self.assertIn("--workspace", bulk)
        self.assertIn("--all-features", bulk)
        for package in {"litchi", *self.plan.excluded_packages}:
            self.assertIn(package, bulk)
        facade = specs[1].argv
        self.assertIn("--package", facade)
        self.assertIn("litchi", facade)
        self.assertNotIn("--all-features", facade)
        feature_argument = facade[facade.index("--features") + 1]
        self.assertNotIn("pages", feature_argument)
        self.assertNotIn("keynote", feature_argument)
        self.assertNotIn("numbers", feature_argument)

    def test_verify_rejects_forbidden_bulk_tree(self) -> None:
        result = subprocess.CompletedProcess(
            ["cargo", "tree"],
            0,
            bulk_tree_output(self.plan, "prost v0.14.0\n"),
            "",
        )
        with mock.patch.object(gate.subprocess, "run", return_value=result):
            with self.assertRaisesRegex(gate.GateError, "forbidden packages.*prost"):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_checks_each_safe_facade_feature(self) -> None:
        calls: list[list[str]] = []

        def fake_run(argv: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            calls.append(argv)
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            else:
                output = "litchi v0.0.1\nlitchi-core v0.0.1\n"
            return subprocess.CompletedProcess(argv, 0, output, "")

        with mock.patch.object(gate.subprocess, "run", side_effect=fake_run):
            bulk_count, facade_count = gate.verify_dependency_trees("cargo", self.plan)
        self.assertEqual(bulk_count, 45)
        self.assertEqual(facade_count, len(self.plan.safe_facade_features))
        self.assertEqual(len(calls), 2 + len(self.plan.safe_facade_features))
        self.assertTrue(all("--locked" not in argv for argv in calls))
        combined = calls[-1]
        combined_features = combined[combined.index("--features") + 1]
        self.assertEqual(
            combined_features,
            ",".join(sorted(self.plan.safe_facade_features - {"default"})),
        )

    def test_verify_rejects_protobuf_in_combined_safe_facade_tree(self) -> None:
        def fake_run(argv: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            elif "--features" in argv and "," in argv[argv.index("--features") + 1]:
                output = "litchi v0.0.1\nprost v0.14.0\n"
            else:
                output = "litchi v0.0.1\nlitchi-core v0.0.1\n"
            return subprocess.CompletedProcess(argv, 0, output, "")

        with mock.patch.object(gate.subprocess, "run", side_effect=fake_run):
            with self.assertRaisesRegex(
                gate.GateError, "combined safe facade feature tree.*prost"
            ):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_rejects_protobuf_in_a_safe_facade_tree(self) -> None:
        def fake_run(argv: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            else:
                output = "litchi v0.0.1\nprost v0.14.0\n"
            return subprocess.CompletedProcess(argv, 0, output, "")

        with mock.patch.object(gate.subprocess, "run", side_effect=fake_run):
            with self.assertRaisesRegex(gate.GateError, "safe facade feature.*prost"):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_rejects_empty_or_partial_bulk_tree(self) -> None:
        partial = f"{sorted(self.plan.bulk_packages)[0]} v0.0.1\n"
        for output in ("", partial):
            with self.subTest(output=repr(output)):
                result = subprocess.CompletedProcess(["cargo", "tree"], 0, output, "")
                with mock.patch.object(gate.subprocess, "run", return_value=result):
                    with self.assertRaisesRegex(
                        gate.GateError,
                        "bulk dependency tree is missing required package roots",
                    ):
                        gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_requires_litchi_root_for_each_facade_tree(self) -> None:
        def fake_run(argv: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            else:
                output = "litchi-core v0.0.1\n"
            return subprocess.CompletedProcess(argv, 0, output, "")

        with mock.patch.object(gate.subprocess, "run", side_effect=fake_run):
            with self.assertRaisesRegex(
                gate.GateError,
                "safe facade feature.*missing required package roots.*litchi",
            ):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_requires_litchi_root_for_combined_facade_tree(self) -> None:
        def fake_run(argv: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            elif "--features" in argv and "," in argv[argv.index("--features") + 1]:
                output = "litchi-core v0.0.1\n"
            else:
                output = "litchi v0.0.1\nlitchi-core v0.0.1\n"
            return subprocess.CompletedProcess(argv, 0, output, "")

        with mock.patch.object(gate.subprocess, "run", side_effect=fake_run):
            with self.assertRaisesRegex(
                gate.GateError,
                "combined safe facade feature tree is missing required package roots.*litchi",
            ):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_environment_disables_incremental_artifacts_by_default(self) -> None:
        environment = gate._environment(None)
        self.assertEqual(environment["CARGO_BUILD_JOBS"], "1")
        self.assertEqual(environment["CARGO_INCREMENTAL"], "0")
        self.assertEqual(environment["CARGO_PROFILE_DEV_DEBUG"], "0")
        self.assertEqual(environment["CARGO_PROFILE_TEST_DEBUG"], "0")

    def test_lib_tests_serializes_all_bulk_roots_before_facade(self) -> None:
        specs = gate.command_specs("cargo", self.plan, "lib-tests")
        bulk_tests = [spec for spec in specs if spec.scope.startswith("bulk-test/")]
        bulk_cleans = [spec for spec in specs if spec.scope.startswith("bulk-clean/")]
        self.assertEqual(len(bulk_tests), 45)
        self.assertEqual(len(bulk_cleans), 45)
        self.assertEqual(specs[-2].scope, "facade-safe-features")
        self.assertEqual(specs[-1].scope, "facade-clean")
        self.assertTrue(
            all(
                {"--all-features", "--lib", "--tests"}.issubset(spec.argv)
                and "--workspace" not in spec.argv
                for spec in bulk_tests
            )
        )
        self.assertEqual(
            {spec.argv[spec.argv.index("--package") + 1] for spec in bulk_tests},
            set(self.plan.bulk_packages),
        )
        facade = specs[-2].argv
        self.assertIn("--no-default-features", facade)
        self.assertIn("--lib", facade)
        self.assertIn("--tests", facade)
        self.assertNotIn("--all-features", facade)

    def test_unknown_facade_dependency_feature_fails_closed(self) -> None:
        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        py_package = next(package for package in packages if package["name"] == "litchi-py")
        dependencies = py_package["dependencies"]
        assert isinstance(dependencies, list)
        dependencies[0]["features"].append("future-unknown")
        with self.assertRaisesRegex(gate.GateError, "unknown litchi features"):
            gate.derive_plan(metadata)

    def test_metadata_projection_rejects_missing_member(self) -> None:
        metadata = metadata_fixture()
        metadata["workspace_members"] = list(metadata["workspace_members"])[:-1]
        with self.assertRaisesRegex(gate.GateError, "workspace package inventory mismatch"):
            gate.derive_plan(metadata)

    def test_workspace_inventory_rejects_renamed_or_swapped_package(self) -> None:
        for replacement in ("litchi-iwa-protos-renamed", "litchi-new-unclassified"):
            with self.subTest(replacement=replacement):
                metadata = copy.deepcopy(metadata_fixture())
                packages = metadata["packages"]
                assert isinstance(packages, list)
                package = next(
                    package
                    for package in packages
                    if package["name"] == "litchi-iwa-protos"
                )
                package["name"] = replacement
                with self.assertRaisesRegex(
                    gate.GateError, "workspace package inventory mismatch"
                ):
                    gate.derive_plan(metadata)

    def test_workflow_paths_and_non_iwork_resource_contract(self) -> None:
        workflow = (ROOT / ".github/workflows/rust-ci.yml").read_text(encoding="utf-8")
        paths = set(re.findall(r"^      - '([^']+)'$", workflow, flags=re.MULTILINE))
        self.assertRegex(workflow, re.compile(r"^  push:$", re.MULTILINE))
        self.assertRegex(workflow, re.compile(r"^  pull_request:$", re.MULTILINE))
        self.assertTrue(
            {
                "rust-toolchain.toml",
                "test-data/**",
                "tools/non_iwork_gate.py",
                "tools/test_non_iwork_gate.py",
            }
            <= paths
        )
        self.assertIn("non-iwork-release-gate:", workflow)
        self.assertIn("CARGO_TARGET_DIR: target/non-iwork-gate", workflow)
        self.assertIn("CARGO_INCREMENTAL: 0", workflow)
        self.assertIn("CARGO_PROFILE_DEV_DEBUG: 0", workflow)
        self.assertIn("CARGO_PROFILE_TEST_DEBUG: 0", workflow)


if __name__ == "__main__":
    unittest.main()
