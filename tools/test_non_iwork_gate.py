from __future__ import annotations

import copy
import json
import re
import subprocess
import unittest
from pathlib import Path
from unittest import mock

from tools import non_iwork_gate as gate


ROOT = Path(__file__).resolve().parents[1]


# Keep the fixture's contract independent from the gate's implementation
# constants.  A changed implementation inventory must make these tests fail
# instead of changing both the fixture and the assertion together.
FIXTURE_FACADE_PACKAGE = "litchi"
FIXTURE_IWORK_PACKAGES = frozenset(
    {
        "litchi-iwa",
        "litchi-iwa-archive",
        "litchi-iwa-cache",
        "litchi-iwa-common",
        "litchi-iwa-core",
        "litchi-iwa-detect",
        "litchi-iwa-graph",
        "litchi-iwa-index",
        "litchi-iwa-package",
        "litchi-iwa-protos",
        "litchi-iwa-structured",
        "litchi-iwa-text",
        "litchi-iwa-text-wire",
        "litchi-keynote",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-pages",
    }
)
FIXTURE_EXCLUDED_PACKAGES = frozenset(
    {
        "litchi-iwa",
        "litchi-iwa-archive",
        "litchi-iwa-cache",
        "litchi-iwa-common",
        "litchi-iwa-core",
        "litchi-iwa-detect",
        "litchi-iwa-graph",
        "litchi-iwa-index",
        "litchi-iwa-package",
        "litchi-iwa-protos",
        "litchi-iwa-structured",
        "litchi-iwa-text",
        "litchi-iwa-text-wire",
        "litchi-keynote",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-pages",
        "litchi-py",
    }
)
FIXTURE_BULK_PACKAGES = frozenset(
    {
        "litchi-biff",
        "litchi-cfb",
        "litchi-codepage",
        "litchi-core",
        "litchi-crypto",
        "litchi-doc",
        "litchi-docx",
        "litchi-drawingml",
        "litchi-eval",
        "litchi-fonts",
        "litchi-formula",
        "litchi-imgconv",
        "litchi-markdown",
        "litchi-odb",
        "litchi-odc",
        "litchi-odf",
        "litchi-odf-common",
        "litchi-odf-formula",
        "litchi-odg",
        "litchi-odi",
        "litchi-odm",
        "litchi-odp",
        "litchi-odraw",
        "litchi-ods",
        "litchi-odt",
        "litchi-ograph",
        "litchi-ole-common",
        "litchi-ooxml-common",
        "litchi-opc",
        "litchi-oth",
        "litchi-ppt",
        "litchi-pptx",
        "litchi-rtf",
        "litchi-sheet",
        "litchi-sign",
        "litchi-slide",
        "litchi-spreadsheet-drawing",
        "litchi-vba",
        "litchi-word",
        "litchi-xls",
        "litchi-xlsb",
        "litchi-xlsx",
        "soapberry-zip",
        "xml-minifier",
        "xml-minifier-macros",
    }
)
FIXTURE_WORKSPACE_PACKAGES = frozenset(
    {
        "litchi",
        "litchi-biff",
        "litchi-cfb",
        "litchi-codepage",
        "litchi-core",
        "litchi-crypto",
        "litchi-doc",
        "litchi-docx",
        "litchi-drawingml",
        "litchi-eval",
        "litchi-fonts",
        "litchi-formula",
        "litchi-imgconv",
        "litchi-iwa",
        "litchi-iwa-archive",
        "litchi-iwa-cache",
        "litchi-iwa-common",
        "litchi-iwa-core",
        "litchi-iwa-detect",
        "litchi-iwa-graph",
        "litchi-iwa-index",
        "litchi-iwa-package",
        "litchi-iwa-protos",
        "litchi-iwa-structured",
        "litchi-iwa-text",
        "litchi-iwa-text-wire",
        "litchi-keynote",
        "litchi-markdown",
        "litchi-numbers",
        "litchi-numbers-wire",
        "litchi-odb",
        "litchi-odc",
        "litchi-odf",
        "litchi-odf-common",
        "litchi-odf-formula",
        "litchi-odg",
        "litchi-odi",
        "litchi-odm",
        "litchi-odp",
        "litchi-odraw",
        "litchi-ods",
        "litchi-odt",
        "litchi-ograph",
        "litchi-ole-common",
        "litchi-ooxml-common",
        "litchi-opc",
        "litchi-oth",
        "litchi-pages",
        "litchi-ppt",
        "litchi-pptx",
        "litchi-py",
        "litchi-rtf",
        "litchi-sheet",
        "litchi-sign",
        "litchi-slide",
        "litchi-spreadsheet-drawing",
        "litchi-vba",
        "litchi-word",
        "litchi-xls",
        "litchi-xlsb",
        "litchi-xlsx",
        "soapberry-zip",
        "xml-minifier",
        "xml-minifier-macros",
    }
)
FIXTURE_SAFE_FACADE_FEATURES = frozenset(
    {
        "automatic-fonts",
        "cfb",
        "default",
        "doc",
        "docx",
        "drawingml",
        "encryption",
        "eval",
        "font-discovery",
        "font-subset",
        "fonts",
        "formula",
        "images",
        "legacy",
        "markdown",
        "odf",
        "odf-common",
        "odp",
        "ods",
        "odt",
        "ole",
        "ooxml",
        "ooxml-common",
        "opc",
        "ppt",
        "pptx",
        "rtf",
        "sheet",
        "sign",
        "vba-inspection",
        "web-functions",
        "xls",
        "xlsb",
        "xlsx",
        "yaml",
    }
)
FIXTURE_UNSAFE_FACADE_FEATURES = frozenset(
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
    }
)


def metadata_fixture() -> dict[str, object]:
    """Build a topology-sized, identity-complete metadata fixture."""

    # These values intentionally model only the closure needed by the planner;
    # the exact feature partition is asserted independently below.
    feature_values = {
        name: [] for name in FIXTURE_SAFE_FACADE_FEATURES
    }
    feature_values.update(
        {
            "pages": ["dep:litchi-pages"],
            "keynote": ["dep:litchi-keynote"],
            "numbers": ["dep:litchi-numbers"],
            "iwork": ["pages", "keynote", "numbers"],
            "word": ["pages"],
            "slides": ["keynote"],
            "spreadsheets": ["numbers"],
            "office": ["iwork"],
            "all-formats": ["office"],
            "all": ["all-formats"],
        }
    )
    names = sorted(FIXTURE_WORKSPACE_PACKAGES)
    packages: list[dict[str, object]] = []
    workspace_members: list[str] = []
    for name in names:
        directory = "pyo3-litchi" if name == "litchi-py" else name
        package_root = ROOT / "crates" / directory
        package_id = f"path+file://{package_root}#0.0.1"
        workspace_members.append(package_id)
        if name == "litchi-py":
            target_kind = ["cdylib"]
            crate_types = ["cdylib"]
        elif name == "xml-minifier-macros":
            target_kind = ["proc-macro"]
            crate_types = ["proc-macro"]
        else:
            target_kind = ["lib"]
            crate_types = ["lib"]
        raw: dict[str, object] = {
            "name": name,
            "version": "0.0.1",
            "id": package_id,
            "source": None,
            "manifest_path": str(package_root / "Cargo.toml"),
            "targets": [
                {
                    "name": name.replace("-", "_"),
                    "kind": target_kind,
                    "crate_types": crate_types,
                    "src_path": str(package_root / "src/lib.rs"),
                }
            ],
            "features": feature_values if name == FIXTURE_FACADE_PACKAGE else {},
            "dependencies": [],
        }
        if name == "litchi-py":
            raw["dependencies"] = [
                {
                    "name": "litchi",
                    "source": None,
                    "req": "*",
                    "kind": None,
                    "rename": None,
                    "optional": False,
                    "uses_default_features": False,
                    "target": None,
                    "registry": None,
                    "path": str(ROOT / "crates/litchi"),
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
        if name == FIXTURE_FACADE_PACKAGE:
            raw["dependencies"] = [
                {
                    "name": dependency_name,
                    "source": None,
                    "req": "^0.0.1",
                    "kind": None,
                    "rename": None,
                    "optional": True,
                    "uses_default_features": True,
                    "target": None,
                    "registry": None,
                    "path": str(ROOT / "crates" / dependency_name),
                    "features": [],
                }
                for dependency_name in sorted(FIXTURE_IWORK_PACKAGES)
            ]
        packages.append(raw)
    return {
        "version": 1,
        "packages": packages,
        "workspace_members": workspace_members,
        "workspace_default_members": workspace_members.copy(),
        "workspace_root": str(ROOT),
    }


def bulk_tree_output(plan: gate.WorkspacePlan, suffix: str = "") -> str:
    """Return a complete synthetic workspace tree root listing."""

    return "".join(
        f"{name} v0.0.1 ({plan.packages[name].manifest_path.parent})\n"
        for name in sorted(plan.bulk_packages)
    ) + suffix


def facade_tree_output(plan: gate.WorkspacePlan, extra: str = "") -> str:
    """Return a synthetic facade tree with its canonical local identity."""

    root = plan.packages[FIXTURE_FACADE_PACKAGE].manifest_path.parent
    return f"{FIXTURE_FACADE_PACKAGE} v0.0.1 ({root})\n" + extra


class NonIworkGateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.plan = gate.derive_plan(metadata_fixture())

    def test_exact_package_partition_and_feature_closure(self) -> None:
        self.assertEqual(self.plan.packages.keys(), FIXTURE_WORKSPACE_PACKAGES)
        self.assertEqual(len(self.plan.packages), 64)
        self.assertEqual(len(self.plan.iwork_packages), 17)
        self.assertEqual(self.plan.iwork_packages, FIXTURE_IWORK_PACKAGES)
        self.assertEqual(len(self.plan.unsafe_facade_dependents), 1)
        self.assertEqual(self.plan.unsafe_facade_dependents, {"litchi-py"})
        self.assertEqual(len(self.plan.excluded_packages), 18)
        self.assertEqual(self.plan.excluded_packages, FIXTURE_EXCLUDED_PACKAGES)
        self.assertEqual(len(self.plan.bulk_packages), 45)
        self.assertEqual(self.plan.bulk_packages, FIXTURE_BULK_PACKAGES)
        self.assertEqual(self.plan.unsafe_facade_features, FIXTURE_UNSAFE_FACADE_FEATURES)
        self.assertEqual(self.plan.safe_facade_features, FIXTURE_SAFE_FACADE_FEATURES)
        self.assertEqual(set(self.plan.facade_features), FIXTURE_SAFE_FACADE_FEATURES | FIXTURE_UNSAFE_FACADE_FEATURES)

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
            "litchi v0.0.1 (/fixture/crates/litchi)\n"
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

    def test_cross_platform_cargo_identity_helpers(self) -> None:
        self.assertEqual(
            gate._file_uri_path("path+file:///C:/work/crates/litchi#litchi@0.0.1"),
            ("C:/work/crates/litchi", "litchi@0.0.1"),
        )
        self.assertEqual(
            gate._file_uri_path("path+file://C:/work/crates/litchi#0.0.1"),
            ("C:/work/crates/litchi", "0.0.1"),
        )
        self.assertEqual(
            gate._file_uri_path("path+file://server/share/crates/litchi#0.0.1"),
            ("//server/share/crates/litchi", "0.0.1"),
        )
        windows_suffix = r"(proc-macro) (C:\work\crates\litchi) (*)"
        self.assertEqual(
            gate._tree_local_path_text(windows_suffix), r"C:\work\crates\litchi"
        )
        self.assertEqual(
            gate._tree_local_path_text(r"(C:\work\crates\litchi) (proc-macro)"),
            r"C:\work\crates\litchi",
        )
        self.assertEqual(
            gate._tree_local_path_text(r"(\\server\share\crates\litchi)"),
            r"\\server\share\crates\litchi",
        )
        for source_annotation in (
            "(registry+https://github.com/foo/bar)",
            "(sparse+https://github.com/foo/bar)",
            "(git+https://github.com/foo/bar)",
            "(git+ssh://git@example.com/foo/bar)",
            "(https://github.com/foo/bar)",
        ):
            with self.subTest(source_annotation=source_annotation):
                self.assertIsNone(gate._tree_local_path_text(source_annotation))
        parsed = gate._parse_tree_entries(
            r"litchi v0.0.1 (C:\work\crates\litchi)" + "\n"
        )
        self.assertEqual(len(parsed["litchi"]), 1)
        self.assertEqual(str(next(iter(parsed["litchi"]))), r"C:\work\crates\litchi")

    def test_tree_selection_includes_all_target_dependency_edges(self) -> None:
        command = gate._tree_command("cargo", ["--package", "litchi"])
        self.assertIn("--target", command)
        self.assertEqual(command[command.index("--target") + 1], "all")
        self.assertIn("--edges", command)
        self.assertEqual(
            command[command.index("--edges") + 1], "normal,build,dev"
        )

    def test_metadata_probe_filters_to_host_platform(self) -> None:
        metadata = metadata_fixture()
        rustc_result = subprocess.CompletedProcess(
            ["rustc", "-vV"], 0, "host: x86_64-unknown-linux-gnu\n", ""
        )
        cargo_result = subprocess.CompletedProcess(
            ["cargo", "metadata"], 0, json.dumps(metadata), ""
        )
        calls: list[list[str]] = []

        def fake_run(argv: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            calls.append(argv)
            return rustc_result if argv[0] == "rustc" else cargo_result

        with mock.patch.object(gate.subprocess, "run", side_effect=fake_run):
            gate._cargo_metadata("cargo")
        metadata_call = next(argv for argv in calls if argv[1] == "metadata")
        self.assertIn("--filter-platform", metadata_call)
        self.assertEqual(
            metadata_call[metadata_call.index("--filter-platform") + 1],
            "x86_64-unknown-linux-gnu",
        )

    def test_tree_requires_canonical_workspace_identity(self) -> None:
        fake = {
            "litchi": frozenset({Path("/tmp/name-spoof")}),
        }
        with self.assertRaisesRegex(gate.GateError, "canonical local identity"):
            gate._require_tree_identities(
                fake, self.plan, {FIXTURE_FACADE_PACKAGE}, "facade tree"
            )
        expected = self.plan.packages[FIXTURE_FACADE_PACKAGE].manifest_path.parent
        with self.assertRaisesRegex(gate.GateError, "non-canonical identities"):
            gate._require_tree_identities(
                {"litchi": frozenset({expected, None})},
                self.plan,
                {FIXTURE_FACADE_PACKAGE},
                "facade tree",
            )
        wrong_version = gate._parse_tree_entries(
            f"litchi v9.9.9 ({expected})\n"
        )
        with self.assertRaisesRegex(gate.GateError, "non-canonical versions"):
            gate._require_tree_identities(
                wrong_version,
                self.plan,
                {FIXTURE_FACADE_PACKAGE},
                "facade tree",
            )

    def test_command_generation_is_argv_safe_and_excludes_all_non_bulk_roots(self) -> None:
        specs = gate.command_specs("cargo", self.plan, "clippy")
        self.assertEqual(len(specs), 3)
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
        facade_default = specs[1].argv
        self.assertIn("--package", facade_default)
        self.assertIn("litchi", facade_default)
        self.assertNotIn("--no-default-features", facade_default)
        facade = specs[2].argv
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

    def test_verify_rejects_target_specific_iwork_edge_in_bulk_tree(self) -> None:
        output = bulk_tree_output(
            self.plan,
            f"litchi-pages v0.0.1 ({self.plan.packages['litchi-pages'].manifest_path.parent})\n",
        )
        result = subprocess.CompletedProcess(["cargo", "tree"], 0, output, "")
        with mock.patch.object(gate.subprocess, "run", return_value=result):
            with self.assertRaisesRegex(
                gate.GateError, "forbidden packages.*litchi-pages"
            ):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_checks_each_safe_facade_feature(self) -> None:
        calls: list[list[str]] = []

        def fake_run(argv: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            calls.append(argv)
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            else:
                output = facade_tree_output(self.plan, "litchi-core v0.0.1\n")
            return subprocess.CompletedProcess(argv, 0, output, "")

        with mock.patch.object(gate.subprocess, "run", side_effect=fake_run):
            bulk_count, facade_count = gate.verify_dependency_trees("cargo", self.plan)
        self.assertEqual(bulk_count, 45)
        self.assertEqual(facade_count, len(self.plan.safe_facade_features))
        self.assertEqual(len(calls), 2 + len(self.plan.safe_facade_features))
        self.assertTrue(all("--locked" not in argv for argv in calls))
        self.assertTrue(all("--target" in argv for argv in calls))
        self.assertTrue(all(argv[argv.index("--target") + 1] == "all" for argv in calls))
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
                output = facade_tree_output(self.plan, "prost v0.14.0\n")
            else:
                output = facade_tree_output(self.plan, "litchi-core v0.0.1\n")
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
                output = facade_tree_output(self.plan, "prost v0.14.0\n")
            return subprocess.CompletedProcess(argv, 0, output, "")

        with mock.patch.object(gate.subprocess, "run", side_effect=fake_run):
            with self.assertRaisesRegex(gate.GateError, "safe facade .*prost"):
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
                "safe facade .*missing required package roots.*litchi",
            ):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_requires_litchi_root_for_combined_facade_tree(self) -> None:
        def fake_run(argv: list[str], **kwargs: object) -> subprocess.CompletedProcess[str]:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            elif "--features" in argv and "," in argv[argv.index("--features") + 1]:
                output = "litchi-core v0.0.1\n"
            else:
                output = facade_tree_output(self.plan, "litchi-core v0.0.1\n")
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
        self.assertEqual(specs[-3].scope, "facade-default-feature")
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
        facade_default = specs[-3].argv
        self.assertNotIn("--no-default-features", facade_default)

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

    def test_metadata_rejects_external_or_escaped_workspace_identity(self) -> None:
        cases = {
            "external source": ("source", "registry+https://example.invalid"),
            "escaped manifest": ("manifest_path", "/tmp/escape/Cargo.toml"),
            "mismatched ID": (
                "id",
                f"path+file://{ROOT / 'crates/litchi-core'}#0.0.1",
            ),
        }
        for label, (field, value) in cases.items():
            with self.subTest(label=label):
                metadata = copy.deepcopy(metadata_fixture())
                packages = metadata["packages"]
                assert isinstance(packages, list)
                package = next(
                    package for package in packages if package["name"] == "litchi"
                )
                package[field] = value
                with self.assertRaises(gate.GateError):
                    gate.derive_plan(metadata)

        metadata = metadata_fixture()
        metadata["workspace_root"] = "/tmp/not-this-workspace"
        with self.assertRaisesRegex(gate.GateError, "workspace root"):
            gate.derive_plan(metadata)

    def test_metadata_rejects_duplicate_ids_and_target_escape(self) -> None:
        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        packages.append(copy.deepcopy(packages[0]))
        with self.assertRaisesRegex(gate.GateError, "repeats package ID"):
            gate.derive_plan(metadata)

        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        package = next(package for package in packages if package["name"] == "litchi")
        targets = package["targets"]
        assert isinstance(targets, list)
        targets[0]["src_path"] = "/tmp/escaped-source.rs"
        with self.assertRaisesRegex(gate.GateError, "source escapes"):
            gate.derive_plan(metadata)

        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        package = next(package for package in packages if package["name"] == "litchi-doc")
        targets = package["targets"]
        assert isinstance(targets, list)
        targets[0]["kind"] = ["bin"]
        targets[0]["crate_types"] = ["bin"]
        with self.assertRaisesRegex(gate.GateError, "no local lib/proc-macro target"):
            gate.derive_plan(metadata)

    def test_metadata_rejects_unsynchronized_id_path_and_version(self) -> None:
        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        package = next(package for package in packages if package["name"] == "litchi")
        old_id = package["id"]
        assert isinstance(old_id, str)
        moved_id = "path+file:///tmp/synchronized-litchi#0.0.1"
        package["id"] = moved_id
        metadata["workspace_members"] = [
            moved_id if member == old_id else member
            for member in metadata["workspace_members"]
        ]
        metadata["workspace_default_members"] = [
            moved_id if member == old_id else member
            for member in metadata["workspace_default_members"]
        ]
        with self.assertRaisesRegex(gate.GateError, "ID path"):
            gate.derive_plan(metadata)

        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        package = next(package for package in packages if package["name"] == "litchi")
        old_id = package["id"]
        assert isinstance(old_id, str)
        suffix_id = f"{old_id.rsplit('#', 1)[0]}#litchi@9.9.9"
        package["id"] = suffix_id
        metadata["workspace_members"] = [
            suffix_id if member == old_id else member
            for member in metadata["workspace_members"]
        ]
        metadata["workspace_default_members"] = [
            suffix_id if member == old_id else member
            for member in metadata["workspace_default_members"]
        ]
        with self.assertRaisesRegex(gate.GateError, "ID suffix"):
            gate.derive_plan(metadata)

    def test_metadata_rejects_spoofed_dependency_fields(self) -> None:
        mutations = {
            "missing source": lambda dependency: dependency.pop("source"),
            "source with path": lambda dependency: dependency.update(
                {"source": "registry+https://github.com/rust-lang/crates.io-index"}
            ),
            "escaped path": lambda dependency: dependency.update(
                {"path": "/tmp/escaped-dependency"}
            ),
            "unknown kind": lambda dependency: dependency.update({"kind": "normal"}),
            "nonboolean optional": lambda dependency: dependency.update(
                {"optional": "true"}
            ),
            "invalid rename": lambda dependency: dependency.update(
                {"rename": "not an alias"}
            ),
        }
        for label, mutate in mutations.items():
            with self.subTest(label=label):
                metadata = copy.deepcopy(metadata_fixture())
                packages = metadata["packages"]
                assert isinstance(packages, list)
                package = next(
                    package for package in packages if package["name"] == "litchi-py"
                )
                dependencies = package["dependencies"]
                assert isinstance(dependencies, list)
                dependency = dependencies[0]
                assert isinstance(dependency, dict)
                mutate(dependency)
                with self.assertRaises(gate.GateError):
                    gate.derive_plan(metadata)

    def test_bulk_direct_excluded_dependency_is_rejected_regardless_of_edge_shape(self) -> None:
        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        package = next(package for package in packages if package["name"] == "litchi-doc")
        dependencies = package["dependencies"]
        assert isinstance(dependencies, list)
        dependencies.append(
            {
                "name": "litchi-pages",
                "source": None,
                "req": "^0.0.1",
                "kind": "dev",
                "rename": "pages_alias",
                "optional": True,
                "uses_default_features": True,
                "target": "cfg(windows)",
                "registry": None,
                "path": str(ROOT / "crates/litchi-pages"),
                "features": [],
            }
        )
        with self.assertRaisesRegex(
            gate.GateError, "bulk package 'litchi-doc'.*litchi-pages"
        ):
            gate.derive_plan(metadata)

    def test_renamed_iwork_dependency_alias_is_unsafe(self) -> None:
        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        facade = next(package for package in packages if package["name"] == "litchi")
        dependencies = facade["dependencies"]
        assert isinstance(dependencies, list)
        pages = next(dependency for dependency in dependencies if dependency["name"] == "litchi-pages")
        pages["rename"] = "iwa_pages"
        features = facade["features"]
        assert isinstance(features, dict)
        features["yaml"] = ["iwa_pages?/inner"]
        plan = gate.derive_plan(metadata)
        self.assertIn("yaml", plan.unsafe_facade_features)
        self.assertNotIn("yaml", plan.safe_facade_features)

    def test_unknown_facade_dependency_alias_fails_closed(self) -> None:
        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        facade = next(package for package in packages if package["name"] == "litchi")
        features = facade["features"]
        assert isinstance(features, dict)
        features["yaml"] = ["dep:future_iwork_alias"]
        with self.assertRaisesRegex(gate.GateError, "unknown dependency packages"):
            gate.derive_plan(metadata)

    def test_unknown_weak_facade_dependency_alias_fails_closed(self) -> None:
        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        facade = next(package for package in packages if package["name"] == "litchi")
        features = facade["features"]
        assert isinstance(features, dict)
        features["yaml"] = ["future_iwork_alias?/inner"]
        with self.assertRaisesRegex(gate.GateError, "unknown dependency packages"):
            gate.derive_plan(metadata)

    def test_metadata_rejects_nonempty_default_until_reviewed(self) -> None:
        metadata = copy.deepcopy(metadata_fixture())
        packages = metadata["packages"]
        assert isinstance(packages, list)
        facade = next(package for package in packages if package["name"] == "litchi")
        features = facade["features"]
        assert isinstance(features, dict)
        features["default"] = ["dep:litchi-pages"]
        with self.assertRaisesRegex(gate.GateError, "default feature must remain"):
            gate.derive_plan(metadata)

    def test_dependency_projection_preserves_default_feature_semantics(self) -> None:
        dependency = gate.Dependency("litchi", frozenset(), True, "cfg(target_os = \"x\")")
        self.assertTrue(
            gate._dependency_reaches_unsafe_facade(dependency, frozenset({"default"}))
        )
        self.assertFalse(
            gate._dependency_reaches_unsafe_facade(
                gate.Dependency("litchi", frozenset(), False, None),
                frozenset({"default"}),
            )
        )
        self.assertFalse(self.plan.packages["litchi-py"].dependencies[0].uses_default_features)

    def test_metadata_projection_rejects_missing_member(self) -> None:
        metadata = metadata_fixture()
        metadata["workspace_members"] = list(metadata["workspace_members"])[:-1]
        metadata["workspace_default_members"] = list(
            metadata["workspace_default_members"]
        )[:-1]
        with self.assertRaisesRegex(
            gate.GateError, "package/member identity mismatch|workspace package inventory mismatch"
        ):
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
                    gate.GateError,
                    "manifest directory|workspace package inventory mismatch",
                ):
                    gate.derive_plan(metadata)

    def test_workflow_paths_and_non_iwork_resource_contract(self) -> None:
        workflow = (ROOT / ".github/workflows/rust-ci.yml").read_text(encoding="utf-8")
        self.assertRegex(workflow, re.compile(r"^  push:$", re.MULTILINE))
        self.assertRegex(workflow, re.compile(r"^  pull_request:$", re.MULTILINE))
        required_paths = {
            "rust-toolchain.toml",
            "crates/**",
            "test-data/**",
            "tools/non_iwork_gate.py",
            "tools/test_non_iwork_gate.py",
            ".cargo/**",
            "tools/**",
            ".github/**",
            "deny.toml",
        }
        event_stops = {"push": "pull_request", "pull_request": "permissions"}
        for event, stop in event_stops.items():
            with self.subTest(event=event):
                match = re.search(
                    rf"^  {event}:\n(.*?)(?=^  {stop}:|^permissions:)",
                    workflow,
                    flags=re.MULTILINE | re.DOTALL,
                )
                self.assertIsNotNone(match)
                assert match is not None
                paths = set(
                    re.findall(r"^      - '([^']+)'$", match.group(1), flags=re.MULTILINE)
                )
                self.assertTrue(required_paths <= paths)
        self.assertIn("non-iwork-release-gate:", workflow)
        self.assertIn("CARGO_TARGET_DIR: target/non-iwork-gate", workflow)
        self.assertIn("CARGO_INCREMENTAL: 0", workflow)
        self.assertIn("CARGO_PROFILE_DEV_DEBUG: 0", workflow)
        self.assertIn("CARGO_PROFILE_TEST_DEBUG: 0", workflow)


if __name__ == "__main__":
    unittest.main()
