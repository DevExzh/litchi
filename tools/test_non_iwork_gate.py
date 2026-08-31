from __future__ import annotations

import copy
import json
import re
import subprocess
import tempfile
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


class _FakeCapture:
    def __init__(self, argv: list[str], returncode: int, stdout: str, stderr: str) -> None:
        self.args = argv
        self.returncode = returncode
        self.stdout = stdout
        self.stderr = stderr


def fake_capture_result(
    stdout: str, returncode: int = 0, stderr: str = ""
) -> _FakeCapture:
    return _FakeCapture(["cargo", "tree"], returncode, stdout, stderr)


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
        calls: list[list[str]] = []

        def fake_capture(argv: list[str], **kwargs: object) -> _FakeCapture:
            calls.append(list(argv))
            self.assertEqual(kwargs["cwd"], gate.ROOT)
            if argv[0] == "rustc":
                return _FakeCapture(
                    argv, 0, "host: x86_64-unknown-linux-gnu\n", ""
                )
            self.assertEqual(argv[0], "cargo")
            return _FakeCapture(argv, 0, json.dumps(metadata), "")

        with mock.patch.object(gate, "_run_capped_capture", side_effect=fake_capture):
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
        bulk_specs = specs[:-2]
        self.assertEqual(len(bulk_specs), len(self.plan.bulk_packages))
        self.assertEqual(
            [spec.scope for spec in bulk_specs],
            [f"bulk/{package}" for package in sorted(self.plan.bulk_packages)],
        )
        for spec in specs:
            self.assertIsInstance(spec.argv, tuple)
            self.assertNotIn("--locked", spec.argv)
            self.assertIn("--no-deps", spec.argv)
            self.assertNotIn("--tests", spec.argv)
            self.assertLess(spec.argv.index("--no-deps"), spec.argv.index("--"))
            self.assertNotIn("'", " ".join(spec.argv))
        for spec, package in zip(
            bulk_specs, sorted(self.plan.bulk_packages), strict=True
        ):
            self.assertNotIn("--workspace", spec.argv)
            self.assertIn("--all-features", spec.argv)
            self.assertEqual(
                spec.argv[spec.argv.index("--package") + 1], package
            )
        selected_roots = {
            spec.argv[spec.argv.index("--package") + 1] for spec in bulk_specs
        }
        self.assertTrue(selected_roots.isdisjoint(self.plan.excluded_packages))
        facade_default = specs[-2].argv
        self.assertIn("--package", facade_default)
        self.assertIn("litchi", facade_default)
        self.assertNotIn("--no-default-features", facade_default)
        facade = specs[-1].argv
        self.assertIn("--package", facade)
        self.assertIn("litchi", facade)
        self.assertNotIn("--all-features", facade)
        feature_argument = facade[facade.index("--features") + 1]
        self.assertNotIn("pages", feature_argument)
        self.assertNotIn("keynote", feature_argument)
        self.assertNotIn("numbers", feature_argument)

    def test_verify_rejects_forbidden_bulk_tree(self) -> None:
        result = fake_capture_result(bulk_tree_output(self.plan, "prost v0.14.0\n"))
        with mock.patch.object(gate, "_run_capped_capture", return_value=result):
            with self.assertRaisesRegex(gate.GateError, "forbidden packages.*prost"):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_rejects_target_specific_iwork_edge_in_bulk_tree(self) -> None:
        output = bulk_tree_output(
            self.plan,
            f"litchi-pages v0.0.1 ({self.plan.packages['litchi-pages'].manifest_path.parent})\n",
        )
        result = fake_capture_result(output)
        with mock.patch.object(gate, "_run_capped_capture", return_value=result):
            with self.assertRaisesRegex(
                gate.GateError, "forbidden packages.*litchi-pages"
            ):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_checks_each_safe_facade_feature(self) -> None:
        calls: list[list[str]] = []

        def fake_capture(argv: list[str], **kwargs: object) -> _FakeCapture:
            calls.append(list(argv))
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            else:
                output = facade_tree_output(self.plan, "litchi-core v0.0.1\n")
            return _FakeCapture(argv, 0, output, "")

        with mock.patch.object(gate, "_run_capped_capture", side_effect=fake_capture):
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
        def fake_capture(argv: list[str], **kwargs: object) -> _FakeCapture:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            elif "--features" in argv and "," in argv[argv.index("--features") + 1]:
                output = facade_tree_output(self.plan, "prost v0.14.0\n")
            else:
                output = facade_tree_output(self.plan, "litchi-core v0.0.1\n")
            return _FakeCapture(argv, 0, output, "")

        with mock.patch.object(gate, "_run_capped_capture", side_effect=fake_capture):
            with self.assertRaisesRegex(
                gate.GateError, "combined safe facade feature tree.*prost"
            ):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_rejects_protobuf_in_a_safe_facade_tree(self) -> None:
        def fake_capture(argv: list[str], **kwargs: object) -> _FakeCapture:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            else:
                output = facade_tree_output(self.plan, "prost v0.14.0\n")
            return _FakeCapture(argv, 0, output, "")

        with mock.patch.object(gate, "_run_capped_capture", side_effect=fake_capture):
            with self.assertRaisesRegex(gate.GateError, "safe facade .*prost"):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_rejects_empty_or_partial_bulk_tree(self) -> None:
        partial = f"{sorted(self.plan.bulk_packages)[0]} v0.0.1\n"
        for output in ("", partial):
            with self.subTest(output=repr(output)):
                result = fake_capture_result(output)
                with mock.patch.object(gate, "_run_capped_capture", return_value=result):
                    with self.assertRaisesRegex(
                        gate.GateError,
                        "bulk dependency tree is missing required package roots",
                    ):
                        gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_requires_litchi_root_for_each_facade_tree(self) -> None:
        def fake_capture(argv: list[str], **kwargs: object) -> _FakeCapture:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            else:
                output = "litchi-core v0.0.1\n"
            return _FakeCapture(argv, 0, output, "")

        with mock.patch.object(gate, "_run_capped_capture", side_effect=fake_capture):
            with self.assertRaisesRegex(
                gate.GateError,
                "safe facade .*missing required package roots.*litchi",
            ):
                gate.verify_dependency_trees("cargo", self.plan)

    def test_verify_requires_litchi_root_for_combined_facade_tree(self) -> None:
        def fake_capture(argv: list[str], **kwargs: object) -> _FakeCapture:
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            elif "--features" in argv and "," in argv[argv.index("--features") + 1]:
                output = "litchi-core v0.0.1\n"
            else:
                output = facade_tree_output(self.plan, "litchi-core v0.0.1\n")
            return _FakeCapture(argv, 0, output, "")

        with mock.patch.object(gate, "_run_capped_capture", side_effect=fake_capture):
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

    def test_non_lib_modes_are_per_root_deterministic_and_use_exact_flags(self) -> None:
        mode_specs = {
            "check": ("check", ("--lib", "--tests"), ()),
            "clippy": (
                "clippy",
                ("--lib", "--no-deps", "--", "-D", "warnings"),
                (),
            ),
            "doc": ("doc", ("--no-deps",), (("RUSTDOCFLAGS", "-D warnings"),)),
            "doc-tests": (
                "test",
                ("--doc", "--", "--test-threads=1"),
                (("RUSTDOCFLAGS", "-D warnings"),),
            ),
            "deprecated": ("check", ("--all-targets",), (("RUSTFLAGS", "-D deprecated"),)),
        }
        bulk_packages = sorted(self.plan.bulk_packages)
        bulk_scopes = [f"bulk/{package}" for package in bulk_packages]
        safe_features = sorted(self.plan.safe_facade_features - {"default"})
        combined_features = ",".join(safe_features)

        for mode, (subcommand, common, updates) in mode_specs.items():
            with self.subTest(mode=mode):
                specs = gate.command_specs("cargo", self.plan, mode)
                self.assertEqual(specs, gate.command_specs("cargo", self.plan, mode))
                self.assertEqual(
                    [spec.scope for spec in specs],
                    bulk_scopes + ["facade-default-feature", "facade-safe-features"],
                )
                self.assertTrue(
                    all("--workspace" not in spec.argv for spec in specs)
                )
                for package, spec in zip(bulk_packages, specs):
                    self.assertEqual(
                        spec.argv,
                        ("cargo", subcommand, "--package", package, "--all-features", *common),
                    )
                    self.assertEqual(spec.env, updates)
                    self.assertNotIn("--features", spec.argv)
                    self.assertNotIn("litchi-py", spec.argv)
                    self.assertTrue(FIXTURE_IWORK_PACKAGES.isdisjoint(spec.argv))

                facade_default = specs[-2]
                self.assertEqual(
                    facade_default.argv,
                    ("cargo", subcommand, "--package", "litchi", *common),
                )
                self.assertEqual(facade_default.env, updates)
                self.assertNotIn("--all-features", facade_default.argv)

                facade_safe = specs[-1]
                self.assertEqual(
                    facade_safe.argv,
                    (
                        "cargo",
                        subcommand,
                        "--package",
                        "litchi",
                        "--no-default-features",
                        "--features",
                        combined_features,
                        *common,
                    ),
                )
                self.assertEqual(facade_safe.env, updates)
                selected_features = set(
                    facade_safe.argv[facade_safe.argv.index("--features") + 1].split(",")
                )
                self.assertEqual(selected_features, set(safe_features))
                self.assertTrue(
                    selected_features.isdisjoint(self.plan.unsafe_facade_features)
                )
                self.assertTrue(
                    FIXTURE_IWORK_PACKAGES.isdisjoint(facade_safe.argv)
                )
                self.assertNotIn("litchi-py", facade_safe.argv)

    def test_lib_and_doc_test_serialization_has_single_test_threads_and_cleanup_order(
        self,
    ) -> None:
        lib_specs = gate.command_specs("cargo", self.plan, "lib-tests")
        expected_scopes: list[str] = []
        for package in sorted(self.plan.bulk_packages):
            expected_scopes.extend((f"bulk-test/{package}", f"bulk-clean/{package}"))
        expected_scopes.extend(
            ("facade-default-feature", "facade-safe-features", "facade-clean")
        )
        self.assertEqual([spec.scope for spec in lib_specs], expected_scopes)
        self.assertEqual(
            lib_specs, gate.command_specs("cargo", self.plan, "lib-tests")
        )
        for index, package in enumerate(sorted(self.plan.bulk_packages)):
            test_spec = lib_specs[index * 2]
            clean_spec = lib_specs[index * 2 + 1]
            self.assertEqual(
                test_spec.argv,
                (
                    "cargo",
                    "test",
                    "--package",
                    package,
                    "--all-features",
                    "--lib",
                    "--tests",
                    "--",
                    "--test-threads=1",
                ),
            )
            self.assertEqual(clean_spec.argv, ("cargo", "clean", "--package", package))
            self.assertEqual(test_spec.argv.count("--test-threads=1"), 1)
            self.assertNotIn("--test-threads=1", clean_spec.argv)
            self.assertNotIn("--workspace", test_spec.argv)
            self.assertNotIn("--workspace", clean_spec.argv)

        facade_default, facade_safe, facade_clean = lib_specs[-3:]
        self.assertEqual(facade_default.argv.count("--test-threads=1"), 1)
        self.assertEqual(facade_safe.argv.count("--test-threads=1"), 1)
        self.assertNotIn("--test-threads=1", facade_clean.argv)
        self.assertNotIn("--workspace", facade_default.argv)
        self.assertNotIn("--workspace", facade_safe.argv)
        self.assertNotIn("--workspace", facade_clean.argv)

        doc_specs = gate.command_specs("cargo", self.plan, "doc-tests")
        self.assertTrue(
            all(spec.argv.count("--test-threads=1") == 1 for spec in doc_specs)
        )
        self.assertTrue(
            all("--workspace" not in spec.argv for spec in doc_specs)
        )

    def test_environment_forces_bounded_values_and_honors_explicit_target_dir(self) -> None:
        ambient = {
            "CARGO_TARGET_DIR": "/ambient/target",
            "CARGO_BUILD_JOBS": "96",
            "CARGO_INCREMENTAL": "1",
            "CARGO_PROFILE_DEV_DEBUG": "2",
            "CARGO_PROFILE_TEST_DEBUG": "2",
            "RUSTDOCFLAGS": "-C opt-level=2",
            "RUSTFLAGS": "--cfg ambient",
            "UNRELATED_GATE_VALUE": "preserved",
        }
        with mock.patch.dict(gate.os.environ, ambient, clear=True):
            environment = gate._environment(None)
        self.assertEqual(
            environment["CARGO_TARGET_DIR"], str(ROOT / "target/non-iwork-gate")
        )
        self.assertEqual(environment["CARGO_BUILD_JOBS"], "1")
        self.assertEqual(environment["CARGO_INCREMENTAL"], "0")
        self.assertEqual(environment["CARGO_PROFILE_DEV_DEBUG"], "0")
        self.assertEqual(environment["CARGO_PROFILE_TEST_DEBUG"], "0")
        self.assertEqual(environment["RUSTDOCFLAGS"], ambient["RUSTDOCFLAGS"])
        self.assertEqual(environment["RUSTFLAGS"], ambient["RUSTFLAGS"])
        self.assertEqual(
            environment["UNRELATED_GATE_VALUE"], ambient["UNRELATED_GATE_VALUE"]
        )

        for target_dir, expected in (
            ("target/explicit-gate", ROOT / "target/explicit-gate"),
            ("/tmp/explicit-gate", Path("/tmp/explicit-gate")),
        ):
            with self.subTest(target_dir=target_dir):
                with mock.patch.dict(gate.os.environ, ambient, clear=True):
                    environment = gate._environment(target_dir)
                self.assertEqual(environment["CARGO_TARGET_DIR"], str(expected))
                self.assertEqual(environment["CARGO_BUILD_JOBS"], "1")
                self.assertEqual(environment["CARGO_INCREMENTAL"], "0")

        base = {
            "CARGO_BUILD_JOBS": "1",
            "RUSTDOCFLAGS": "-C opt-level=2",
        }
        updated = gate._updated_environment(
            base,
            (
                ("RUSTDOCFLAGS", "-D warnings"),
                ("RUSTDOCFLAGS", "-D warnings"),
                ("RUSTFLAGS", "-D deprecated"),
            ),
        )
        self.assertEqual(updated["RUSTDOCFLAGS"], "-D warnings")
        self.assertEqual(updated["RUSTFLAGS"], "-D deprecated")
        self.assertEqual(base["RUSTDOCFLAGS"], "-C opt-level=2")

    def test_execution_recorder_records_ordered_success_phases(self) -> None:
        specs = (
            gate.CommandSpec("first", ("cargo", "check", "first")),
            gate.CommandSpec("second", ("cargo", "check", "second")),
            gate.CommandSpec("third", ("cargo", "check", "third")),
        )

        class FakeProcess:
            def __init__(self, pid: int, returncode: int) -> None:
                self.pid = pid
                self.returncode = returncode

            def poll(self) -> int:
                return self.returncode

            def wait(self) -> int:
                return self.returncode

        processes = iter(
            (
                FakeProcess(101, 0),
                FakeProcess(102, 0),
                FakeProcess(103, 0),
            )
        )
        popen_calls: list[tuple[list[str], dict[str, object]]] = []

        def popen_factory(argv: list[str], **kwargs: object) -> FakeProcess:
            popen_calls.append((argv, kwargs))
            return next(processes)

        samples = iter((250, 450, 500))
        sample_pids: list[int] = []

        def rss_sampler(pid: int) -> int:
            sample_pids.append(pid)
            return next(samples)

        ticks = [0]

        def clock() -> int:
            ticks[0] += 1_000
            return ticks[0]

        scans = [0]

        def target_scanner(_: Path | None) -> gate.TargetFootprint:
            scans[0] += 1
            return gate.TargetFootprint("complete", scans[0] * 10, scans[0])

        environment = {
            "CARGO_TARGET_DIR": str(ROOT / "target/non-iwork-gate"),
            "CARGO_BUILD_JOBS": "1",
            "SECRET_NOT_RECORDED": "hidden",
        }
        recorder = gate.ExecutionRecorder(
            "check",
            environment,
            clock=clock,
            target_scanner=target_scanner,
            popen_factory=popen_factory,
            rss_sampler=rss_sampler,
            rss_platform="fixture",
            rss_scope="injected-process-tree",
            sleep_fn=lambda _: None,
        )
        with mock.patch.object(gate, "command_specs", return_value=specs):
            gate.run_mode("cargo", self.plan, "check", environment, recorder=recorder)
        recorder.finish("passed", None)

        self.assertEqual(
            [argv for argv, _ in popen_calls], [list(spec.argv) for spec in specs]
        )
        self.assertEqual(sample_pids, [101, 102, 103])
        self.assertEqual([phase.index for phase in recorder.phases], [1, 2, 3])
        self.assertEqual(
            [phase.scope for phase in recorder.phases], ["first", "second", "third"]
        )
        self.assertEqual(
            [phase.status for phase in recorder.phases], ["passed", "passed", "passed"]
        )
        self.assertEqual(
            [phase.child_rss.status for phase in recorder.phases],
            ["available", "available", "available"],
        )
        self.assertEqual(
            [phase.child_rss.high_water_bytes for phase in recorder.phases],
            [250, 450, 500],
        )
        self.assertEqual(recorder.child_rss_high_water_bytes, 500)
        report = recorder.as_dict()
        self.assertEqual(report["outcome"], "passed")
        self.assertIsNone(report["error"])
        self.assertEqual(report["env_keys"], ["CARGO_BUILD_JOBS", "CARGO_TARGET_DIR"])
        self.assertEqual(report["claim_scope"], "no performance claim")
        self.assertEqual(report["target_dir"], environment["CARGO_TARGET_DIR"])
        self.assertEqual(
            report["cargo_env"],
            {
                "CARGO_BUILD_JOBS": "1",
                "CARGO_TARGET_DIR": environment["CARGO_TARGET_DIR"],
            },
        )
        self.assertEqual(
            set(report),
            {
                "version",
                "mode",
                "claim_scope",
                "outcome",
                "error",
                "elapsed_ns",
                "clock",
                "host",
                "target_dir",
                "env_keys",
                "cargo_env",
                "target_before",
                "target_after",
                "child_rss",
                "cleanup",
                "feature_unification",
                "limitations",
                "phases",
                "target_scan_limits",
            },
        )

    def test_execution_recorder_stops_after_first_failure_and_records_it(self) -> None:
        specs = (
            gate.CommandSpec("passed", ("cargo", "check", "passed")),
            gate.CommandSpec("failed", ("cargo", "check", "failed")),
            gate.CommandSpec("not-run", ("cargo", "check", "not-run")),
        )

        class FakeProcess:
            def __init__(self, pid: int, returncode: int) -> None:
                self.pid = pid
                self.returncode = returncode

            def poll(self) -> int:
                return self.returncode

        processes = iter((FakeProcess(201, 0), FakeProcess(202, 7)))
        popen_calls: list[list[str]] = []

        def popen_factory(argv: list[str], **_: object) -> FakeProcess:
            popen_calls.append(argv)
            return next(processes)

        samples = iter((1_024, 2_048))

        def rss_sampler(_: int) -> int:
            return next(samples)

        ticks = [0]

        def clock() -> int:
            ticks[0] += 100
            return ticks[0]

        recorder = gate.ExecutionRecorder(
            "check",
            {"CARGO_TARGET_DIR": str(ROOT / "target/non-iwork-gate")},
            clock=clock,
            target_scanner=lambda _: gate.TargetFootprint("unavailable", None, None),
            popen_factory=popen_factory,
            rss_sampler=rss_sampler,
            rss_platform="fixture",
            rss_scope="injected-process-tree",
            sleep_fn=lambda _: None,
        )
        failure: gate.GateError | None = None
        with mock.patch.object(gate, "command_specs", return_value=specs):
            try:
                gate.run_mode("cargo", self.plan, "check", recorder.environment, recorder=recorder)
            except gate.GateError as error:
                failure = error
        self.assertIsNotNone(failure)
        assert failure is not None
        self.assertEqual(str(failure), "check/failed failed with exit status 7")
        recorder.finish("failed", str(failure))

        self.assertEqual(
            popen_calls,
            [list(specs[0].argv), list(specs[1].argv)],
        )
        self.assertEqual([phase.index for phase in recorder.phases], [1, 2])
        self.assertEqual(
            [phase.status for phase in recorder.phases], ["passed", "failed"]
        )
        self.assertEqual([phase.returncode for phase in recorder.phases], [0, 7])
        self.assertEqual(recorder.as_dict()["outcome"], "failed")
        self.assertEqual(recorder.as_dict()["error"], str(failure))

    def test_lib_tests_run_one_child_at_a_time_and_record_cleanup_report(self) -> None:
        specs = gate.command_specs("cargo", self.plan, "lib-tests")
        active: set[int] = set()
        processes: list[object] = []
        events: list[tuple[str, tuple[str, ...]]] = []
        max_active = [0]

        class FakeProcess:
            def __init__(self, pid: int, argv: tuple[str, ...]) -> None:
                self.pid = pid
                self.argv = argv
                self.poll_count = 0
                self.completed = False

            def poll(self) -> int | None:
                self.poll_count += 1
                if self.poll_count == 1:
                    return None
                if not self.completed:
                    active.remove(self.pid)
                    events.append(("complete", self.argv))
                    self.completed = True
                return 0

            def wait(self) -> int:
                if not self.completed:
                    active.remove(self.pid)
                    events.append(("complete", self.argv))
                    self.completed = True
                return 0

        def popen_factory(argv: list[str], **_: object) -> FakeProcess:
            self.assertFalse(active, "a second Cargo child started before the first ended")
            pid = 10_000 + len(processes)
            process = FakeProcess(pid, tuple(argv))
            processes.append(process)
            active.add(pid)
            max_active[0] = max(max_active[0], len(active))
            events.append(("start", tuple(argv)))
            return process

        ticks = [0]

        def clock() -> int:
            ticks[0] += 100
            return ticks[0]

        environment = {
            "CARGO_TARGET_DIR": str(ROOT / "target/non-iwork-recorder"),
            "CARGO_BUILD_JOBS": "1",
            "CARGO_INCREMENTAL": "0",
            "CARGO_PROFILE_DEV_DEBUG": "0",
            "CARGO_PROFILE_TEST_DEBUG": "0",
            "SECRET_NOT_RECORDED": "hidden",
        }
        recorder = gate.ExecutionRecorder(
            "lib-tests",
            environment,
            clock=clock,
            target_scanner=lambda _: gate.TargetFootprint(
                "unavailable", None, None, "fixture target"
            ),
            popen_factory=popen_factory,
            rss_sampler=lambda _: 4096,
            rss_platform="fixture",
            rss_scope="injected-process-tree",
            sleep_fn=lambda _: None,
        )
        with mock.patch("builtins.print"):
            gate.run_mode(
                "cargo", self.plan, "lib-tests", environment, recorder=recorder
            )
        recorder.finish("passed", None)

        self.assertEqual(max_active[0], 1)
        self.assertEqual(len(processes), len(specs))
        for index, spec in enumerate(specs):
            self.assertEqual(
                events[index * 2 : index * 2 + 2],
                [("start", spec.argv), ("complete", spec.argv)],
            )

        report = recorder.as_dict()
        expected_env = {
            "CARGO_BUILD_JOBS": "1",
            "CARGO_INCREMENTAL": "0",
            "CARGO_PROFILE_DEV_DEBUG": "0",
            "CARGO_PROFILE_TEST_DEBUG": "0",
            "CARGO_TARGET_DIR": environment["CARGO_TARGET_DIR"],
        }
        self.assertEqual(report["cargo_env"], expected_env)
        self.assertEqual(len(report["phases"]), len(specs))
        self.assertEqual(
            report["phases"][0]["cargo_env"], expected_env
        )
        expected_clean_scopes = [
            spec.scope
            for spec in specs
            if spec.scope.startswith("bulk-clean/") or spec.scope == "facade-clean"
        ]
        self.assertEqual(
            report["cleanup"]["package_clean_scopes"], expected_clean_scopes
        )
        self.assertEqual(
            report["cleanup"]["package_clean_commands"],
            [
                {"scope": spec.scope, "argv": list(spec.argv)}
                for spec in specs
                if spec.scope.startswith("bulk-clean/") or spec.scope == "facade-clean"
            ],
        )

        with tempfile.TemporaryDirectory() as directory:
            report_path = Path(directory) / "recorder.json"
            gate._write_atomic_report(report_path, report)
            encoded = (json.dumps(report, sort_keys=True, indent=2) + "\n").encode()
            self.assertEqual(report_path.read_bytes(), encoded)
            previous = b"previous report\n"
            report_path.write_bytes(previous)
            with mock.patch.object(
                gate.os, "replace", side_effect=OSError("atomic replacement failed")
            ):
                with self.assertRaisesRegex(OSError, "atomic replacement failed"):
                    gate._write_atomic_report(report_path, report)
            self.assertEqual(report_path.read_bytes(), previous)
            self.assertFalse(list(report_path.parent.glob(".recorder.json.*.tmp")))

    def test_main_records_spawn_oserror_with_error_phase(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            record_file = Path(directory) / "spawn-error.json"
            environment = {"CARGO_TARGET_DIR": str(Path(directory) / "target")}
            with (
                mock.patch.object(
                    gate, "_cargo_metadata", return_value=metadata_fixture()
                ),
                mock.patch.object(gate, "_environment", return_value=environment),
                mock.patch.object(
                    gate.subprocess, "Popen", side_effect=OSError("spawn denied")
                ) as popen,
            ):
                status = gate.main(
                    ["check", "--record-file", str(record_file)]
                )

            self.assertEqual(status, 1)
            self.assertEqual(popen.call_count, 1)
            report = json.loads(record_file.read_text(encoding="utf-8"))
            self.assertEqual(report["outcome"], "failed")
            self.assertEqual(report["error"], "spawn denied")
            self.assertEqual(len(report["phases"]), 1)
            phase = report["phases"][0]
            self.assertEqual(phase["status"], "error")
            self.assertIsNone(phase["returncode"])
            self.assertEqual(phase["child_rss"]["status"], "unavailable")
            self.assertEqual(
                phase["child_rss"]["reason"], "child process could not be started"
            )

    def test_main_records_keyboard_interrupt_and_reaps_child(self) -> None:
        class FakeProcess:
            pid = 42_042

            def __init__(self) -> None:
                self.wait_calls = 0

            def poll(self) -> None:
                return None

            def wait(self) -> int:
                self.wait_calls += 1
                return 130

        process = FakeProcess()
        with tempfile.TemporaryDirectory() as directory:
            record_file = Path(directory) / "interrupted.json"
            environment = {"CARGO_TARGET_DIR": str(Path(directory) / "target")}
            with (
                mock.patch.object(
                    gate, "_cargo_metadata", return_value=metadata_fixture()
                ),
                mock.patch.object(gate, "_environment", return_value=environment),
                mock.patch.object(
                    gate.subprocess, "Popen", return_value=process
                ) as popen,
                mock.patch.object(
                    gate, "_linux_process_rss", side_effect=KeyboardInterrupt
                ),
            ):
                status = gate.main(
                    ["check", "--record-file", str(record_file)]
                )

            self.assertEqual(status, 130)
            self.assertEqual(popen.call_count, 1)
            self.assertEqual(process.wait_calls, 1)
            report = json.loads(record_file.read_text(encoding="utf-8"))
            self.assertEqual(report["outcome"], "interrupted")
            self.assertEqual(report["error"], "KeyboardInterrupt")
            self.assertEqual(len(report["phases"]), 1)
            phase = report["phases"][0]
            self.assertEqual(phase["status"], "interrupted")
            self.assertIsNone(phase["returncode"])
            self.assertEqual(
                phase["child_rss"]["reason"],
                "KeyboardInterrupt interrupted RSS sampling",
            )

    def test_bounded_capture_truncates_environment_and_error_values(self) -> None:
        environment_value = "e" * (gate.MAX_RECORDED_ENV_VALUE_LENGTH + 17)
        self.assertEqual(
            gate._bounded_env_values(
                {"RUSTFLAGS": environment_value, "SECRET_NOT_RECORDED": "secret"}
            ),
            ((
                "RUSTFLAGS",
                "e" * gate.MAX_RECORDED_ENV_VALUE_LENGTH + "...",
            ),),
        )

        error_value = "x" * (gate.MAX_REPORT_ERROR_LENGTH + 17)
        recorder = gate.ExecutionRecorder(
            "check",
            {"RUSTFLAGS": environment_value},
            clock=lambda: 1,
            target_scanner=lambda _: gate.TargetFootprint(
                "incomplete", None, None, "fixture scan incomplete"
            ),
            rss_sampler=lambda _: None,
            rss_platform="fixture",
            rss_scope="unavailable",
        )
        recorder.finish("failed", error_value)
        report = recorder.as_dict()
        self.assertEqual(
            report["cargo_env"]["RUSTFLAGS"],
            "e" * gate.MAX_RECORDED_ENV_VALUE_LENGTH + "...",
        )
        self.assertEqual(
            report["error"], "x" * gate.MAX_REPORT_ERROR_LENGTH
        )
        self.assertEqual(report["target_before"]["status"], "incomplete")
        self.assertEqual(report["target_after"]["status"], "incomplete")
        self.assertIsNone(report["target_after"]["regular_file_bytes"])

    def test_required_flags_use_exact_unencoded_spellings(self) -> None:
        for mode in (
            "check",
            "clippy",
            "doc",
            "doc-tests",
            "deprecated",
            "lib-tests",
        ):
            with self.subTest(mode=mode):
                specs = gate.command_specs("cargo", self.plan, mode)
                for spec in specs:
                    self.assertTrue(
                        all("%" not in argument for argument in spec.argv),
                        spec,
                    )
                    self.assertNotIn("--all-features=true", spec.argv)
                    self.assertNotIn("--no-default-features=true", spec.argv)
                    self.assertNotIn("--test-threads%3D1", spec.argv)
                    self.assertNotIn("--test-threads%3d1", spec.argv)
                bulk = [
                    spec
                    for spec in specs
                    if spec.scope.startswith("bulk/")
                    or spec.scope.startswith("bulk-test/")
                ]
                self.assertTrue(all("--all-features" in spec.argv for spec in bulk))
                safe = next(
                    spec for spec in specs if spec.scope == "facade-safe-features"
                )
                self.assertIn("--no-default-features", safe.argv)
                if mode in {"doc-tests", "lib-tests"}:
                    test_specs = [spec for spec in specs if spec.argv[1] == "test"]
                    self.assertTrue(test_specs)
                    self.assertTrue(
                        all(
                            spec.argv.count("--test-threads=1") == 1
                            for spec in test_specs
                        )
                    )

    def test_capped_capture_reaps_process_and_rejects_output_overflow(self) -> None:
        class FakeStream:
            def __init__(self, chunks: list[bytes]) -> None:
                self.chunks = iter(chunks)

            def read(self, _: int) -> bytes:
                return next(self.chunks, b"")

            def close(self) -> None:
                return None

        class FakeProcess:
            def __init__(self) -> None:
                self.stdout = FakeStream([b"12345"])
                self.stderr = FakeStream([b""])
                self.wait_calls = 0

            def poll(self) -> int:
                return 0

            def wait(self, *_args: object, **_kwargs: object) -> int:
                self.wait_calls += 1
                return 0

            def terminate(self) -> None:
                return None

            def kill(self) -> None:
                return None

            def send_signal(self, _: int) -> None:
                return None

        process = FakeProcess()
        with self.assertRaises(gate.CaptureLimitError) as raised:
            gate._run_capped_capture(
                ["cargo", "metadata"],
                cwd=ROOT,
                limit_bytes=4,
                popen_factory=lambda *_args, **_kwargs: process,
                sleep_fn=lambda _: None,
            )
        self.assertIn("stdout", str(raised.exception))
        self.assertGreaterEqual(process.wait_calls, 1)

    def test_cleanup_report_retains_failing_root_and_names_prior_successes(self) -> None:
        specs = gate.command_specs("cargo", self.plan, "lib-tests")
        statuses = iter((0, 0, 17))
        calls: list[list[str]] = []

        class FakeProcess:
            pid = 8_001

            def poll(self) -> int:
                return next(statuses)

        def popen_factory(argv: list[str], **_: object) -> FakeProcess:
            calls.append(argv)
            return FakeProcess()

        recorder = gate.ExecutionRecorder(
            "lib-tests",
            {"CARGO_TARGET_DIR": str(ROOT / "target/non-iwork-cleanup")},
            clock=lambda: 1,
            target_scanner=lambda _: gate.TargetFootprint(
                "unavailable", None, None, "fixture"
            ),
            popen_factory=popen_factory,
            rss_sampler=lambda _: None,
            rss_platform="fixture",
            rss_scope="unavailable",
            sleep_fn=lambda _: None,
        )
        failure: gate.GateError | None = None
        with mock.patch("builtins.print"):
            try:
                gate.run_mode(
                    "cargo", self.plan, "lib-tests", recorder.environment, recorder=recorder
                )
            except gate.GateError as error:
                failure = error
        self.assertIsNotNone(failure)
        assert failure is not None
        recorder.finish("failed", str(failure))

        first_package = sorted(self.plan.bulk_packages)[0]
        self.assertEqual(len(calls), 3)
        self.assertEqual(
            [phase.scope for phase in recorder.phases],
            [
                f"bulk-test/{first_package}",
                f"bulk-clean/{first_package}",
                f"bulk-test/{sorted(self.plan.bulk_packages)[1]}",
            ],
        )
        cleanup = recorder.as_dict()["cleanup"]
        self.assertEqual(
            cleanup["package_clean_scopes"], [f"bulk-clean/{first_package}"]
        )
        self.assertEqual(cleanup["disposition"], "retained_by_policy")
        self.assertEqual(
            cleanup["failure_artifacts"], "failing phase/root retained"
        )

    def test_aggregate_rss_reports_partial_when_samples_are_unreadable(self) -> None:
        spec = gate.CommandSpec("one", ("cargo", "check", "one"))

        class FakeProcess:
            pid = 9_001

            def __init__(self) -> None:
                self.poll_results = iter((None, 0))

            def poll(self) -> int | None:
                return next(self.poll_results)

        samples = iter((8_192, -1))
        recorder = gate.ExecutionRecorder(
            "check",
            {"CARGO_TARGET_DIR": str(ROOT / "target/non-iwork-rss")},
            clock=lambda: 1,
            target_scanner=lambda _: gate.TargetFootprint(
                "unavailable", None, None, "fixture"
            ),
            popen_factory=lambda *_args, **_kwargs: FakeProcess(),
            rss_sampler=lambda _: next(samples),
            rss_platform="fixture",
            rss_scope="injected-process-tree",
            sleep_fn=lambda _: None,
        )
        recorder.run_phase(1, spec, recorder.environment)
        recorder.finish("passed", None)
        report = recorder.as_dict()
        phase_rss = report["phases"][0]["child_rss"]
        self.assertEqual(phase_rss["status"], "partial")
        self.assertIn("unreadable", phase_rss["reason"])
        self.assertEqual(report["child_rss"]["status"], "partial")
        self.assertEqual(
            report["child_rss"]["reason"],
            "one or more phase RSS samples were partial or unavailable",
        )

    def test_aggregate_rss_reports_unavailable_when_pre_poll_samples_are_unreadable(
        self,
    ) -> None:
        spec = gate.CommandSpec("one", ("cargo", "check", "one"))

        class FakeProcess:
            pid = 9_002

            def __init__(self) -> None:
                self.poll_results = iter((None, 0))

            def poll(self) -> int | None:
                return next(self.poll_results)

        samples = iter((-1, "unreadable"))
        recorder = gate.ExecutionRecorder(
            "check",
            {"CARGO_TARGET_DIR": str(ROOT / "target/non-iwork-rss-unavailable")},
            clock=lambda: 1,
            target_scanner=lambda _: gate.TargetFootprint(
                "unavailable", None, None, "fixture"
            ),
            popen_factory=lambda *_args, **_kwargs: FakeProcess(),
            rss_sampler=lambda _: next(samples),
            rss_platform="fixture",
            rss_scope="injected-process-tree",
            sleep_fn=lambda _: None,
        )
        recorder.run_phase(1, spec, recorder.environment)
        recorder.finish("passed", None)
        report = recorder.as_dict()
        self.assertEqual(report["phases"][0]["child_rss"]["status"], "unavailable")
        self.assertEqual(report["child_rss"]["status"], "unavailable")

    def test_target_footprint_refuses_symlinks_and_reports_incomplete_scan(self) -> None:
        with tempfile.TemporaryDirectory() as target_root, tempfile.TemporaryDirectory() as outside_root:
            target = Path(target_root) / "target"
            target.mkdir()
            (target / "inside.bin").write_bytes(b"123")
            outside = Path(outside_root)
            (outside / "outside.bin").write_bytes(b"outside")
            linked_directory = target / "linked-directory"
            linked_target = Path(target_root) / "target-link"
            try:
                linked_directory.symlink_to(outside, target_is_directory=True)
                linked_target.symlink_to(target, target_is_directory=True)
            except (NotImplementedError, OSError) as error:
                self.skipTest(f"symlinks unavailable: {error}")

            footprint = gate._target_footprint(target)
            self.assertEqual(footprint.status, "complete")
            self.assertEqual(footprint.regular_file_bytes, 3)
            self.assertEqual(footprint.regular_file_count, 1)
            self.assertTrue(footprint.reason)
            linked_footprint = gate._target_footprint(linked_target)
            self.assertEqual(linked_footprint.status, "unavailable")
            self.assertIsNone(linked_footprint.regular_file_bytes)
            self.assertIsNone(linked_footprint.regular_file_count)
            self.assertTrue(linked_footprint.reason)

            with mock.patch.object(
                gate.os, "scandir", side_effect=OSError("directory disappeared")
            ):
                incomplete = gate._target_footprint(target)
            self.assertEqual(incomplete.status, "incomplete")
            self.assertIsNone(incomplete.regular_file_bytes)
            self.assertIsNone(incomplete.regular_file_count)
            self.assertTrue(incomplete.reason)

    def test_rss_unavailable_and_normalized_schema(self) -> None:
        with mock.patch.object(gate.sys, "platform", "darwin"):
            sampler, platform, scope = gate._rss_configuration(
                True, None, None, None
            )
        self.assertIsNone(sampler)
        self.assertEqual(platform, "darwin")
        self.assertEqual(scope, "unavailable")

        samples = iter((4096, -1, "invalid", 1024))
        tracker = gate._ChildRssTracker(
            lambda _: next(samples),
            "fixture",
            "injected-process-tree",
        )
        for _ in range(4):
            tracker.sample(1)
        self.assertEqual(tracker.sample_count, 2)
        self.assertEqual(tracker.sample_errors, 2)
        self.assertEqual(
            tracker.result().as_dict(),
            {
                "measurement": gate.RSS_MEASUREMENT,
                "high_water_bytes": 4096,
                "platform": "fixture",
                "scope": "injected-process-tree",
                "status": "partial",
                "reason": "one or more process-tree samples were unreadable",
                "sample_interval_ms": gate.RSS_SAMPLE_INTERVAL_MS,
            },
        )
        self.assertEqual(
            gate.ChildRss(None, "darwin", "unavailable").as_dict(),
            {
                "measurement": gate.RSS_MEASUREMENT,
                "high_water_bytes": None,
                "platform": "darwin",
                "scope": "unavailable",
                "status": "unavailable",
                "reason": "",
                "sample_interval_ms": gate.RSS_SAMPLE_INTERVAL_MS,
            },
        )

        recorder = gate.ExecutionRecorder(
            "check",
            {},
            clock=lambda: 1,
            target_scanner=lambda _: gate.TargetFootprint("unavailable", None, None),
            rss_sampler=lambda _: None,
            rss_platform="fixture",
            rss_scope="unavailable",
        )
        recorder.finish("passed", None)
        self.assertEqual(
            recorder.as_dict()["child_rss"],
            {
                "measurement": gate.RSS_MEASUREMENT,
                "high_water_bytes": None,
                "platform": "fixture",
                "scope": "unavailable",
                "status": "unavailable",
                "reason": "no execution phases were sampled",
                "sample_interval_ms": gate.RSS_SAMPLE_INTERVAL_MS,
            },
        )

    def test_atomic_report_replacement_is_deterministic_and_preserves_on_failure(
        self,
    ) -> None:
        report = {
            "version": 1,
            "mode": "check",
            "outcome": "passed",
            "error": None,
            "features": ["doc", "docx"],
            "packages": ["litchi-core", "litchi-doc"],
        }
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "execution.json"
            gate._write_atomic_report(path, report)
            first = path.read_bytes()
            self.assertEqual(
                first,
                (json.dumps(report, sort_keys=True, indent=2) + "\n").encode(),
            )
            gate._write_atomic_report(path, report)
            self.assertEqual(path.read_bytes(), first)

            replacements: list[tuple[Path, Path]] = []
            real_replace = gate.os.replace

            def recording_replace(source: str | Path, destination: str | Path) -> None:
                replacements.append((Path(source), Path(destination)))
                real_replace(source, destination)

            with mock.patch.object(gate.os, "replace", side_effect=recording_replace):
                gate._write_atomic_report(path, report)
            self.assertEqual(len(replacements), 1)
            temporary, destination = replacements[0]
            self.assertEqual(destination, path)
            self.assertEqual(temporary.parent, path.parent)
            self.assertNotEqual(temporary, path)
            self.assertFalse(list(path.parent.glob(f".{path.name}.*.tmp")))

            previous = b"previous report\n"
            path.write_bytes(previous)
            with mock.patch.object(
                gate.os, "replace", side_effect=OSError("replacement failed")
            ):
                with self.assertRaisesRegex(OSError, "replacement failed"):
                    gate._write_atomic_report(path, report)
            self.assertEqual(path.read_bytes(), previous)
            self.assertFalse(list(path.parent.glob(f".{path.name}.*.tmp")))

    def test_record_file_main_reports_passed_and_failed_outcomes(self) -> None:
        for run_error, expected_status, expected_outcome, expected_error in (
            (None, 0, "passed", None),
            (gate.GateError("synthetic failure"), 1, "failed", "synthetic failure"),
        ):
            with self.subTest(expected_outcome=expected_outcome):
                with tempfile.TemporaryDirectory() as directory:
                    record_file = Path(directory) / "execution.json"
                    environment = {
                        "CARGO_TARGET_DIR": str(Path(directory) / "target")
                    }
                    with (
                        mock.patch.object(
                            gate, "_cargo_metadata", return_value=metadata_fixture()
                        ),
                        mock.patch.object(
                            gate, "_environment", return_value=environment
                        ),
                        mock.patch.object(
                            gate, "run_mode", side_effect=run_error
                        ) as run_mode,
                    ):
                        status = gate.main(
                            ["check", "--record-file", str(record_file)]
                        )
                    self.assertEqual(status, expected_status)
                    payload = json.loads(record_file.read_text(encoding="utf-8"))
                    self.assertEqual(payload["version"], gate.REPORT_VERSION)
                    self.assertEqual(payload["mode"], "check")
                    self.assertEqual(payload["outcome"], expected_outcome)
                    self.assertEqual(payload["error"], expected_error)
                    run_mode.assert_called_once()

    def test_verify_selections_are_sorted_and_contain_no_unsafe_facade_features(
        self,
    ) -> None:
        calls: list[list[str]] = []

        def fake_capture(argv: list[str], **kwargs: object) -> _FakeCapture:
            calls.append(list(argv))
            if "--workspace" in argv:
                output = bulk_tree_output(self.plan)
            else:
                output = facade_tree_output(self.plan, "litchi-core v0.0.1\n")
            return _FakeCapture(argv, 0, output, "")

        with mock.patch.object(gate, "_run_capped_capture", side_effect=fake_capture):
            gate.verify_dependency_trees("cargo", self.plan)

        selections = [argv[2 : argv.index("--target")] for argv in calls]
        expected = [
            ["--workspace", "--all-features", *gate._workspace_exclusions(self.plan)],
            gate._facade_tree_selection(self.plan, "default"),
        ]
        expected.extend(
            gate._facade_tree_selection(self.plan, feature)
            for feature in sorted(self.plan.safe_facade_features - {"default"})
        )
        expected.append(gate._facade_tree_selection(self.plan))
        self.assertEqual(selections, expected)

        for selection in selections[1:]:
            self.assertNotIn("litchi-py", selection)
            self.assertTrue(FIXTURE_IWORK_PACKAGES.isdisjoint(selection))
            if "--features" in selection:
                selected = set(selection[selection.index("--features") + 1].split(","))
                self.assertTrue(selected.isdisjoint(self.plan.unsafe_facade_features))

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
