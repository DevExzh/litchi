from __future__ import annotations

import socket
import subprocess
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from tools import check_example_targets as examples


class ExampleTargetChecks(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    def write_root(
        self,
        members: list[str],
        excludes: list[str] | None = None,
        workspace_package: str = 'version = "0.0.1"\nedition = "2024"',
    ) -> None:
        member_lines = ", ".join(f'"{member}"' for member in members)
        exclude_lines = (
            ""
            if excludes is None
            else f'\nexclude = [{", ".join(f"{item!r}" for item in excludes)}]'
        )
        self.root.joinpath("Cargo.toml").write_text(
            "[workspace]\n"
            f"members = [{member_lines}]{exclude_lines}\n"
            "[workspace.package]\n"
            f"{workspace_package}\n",
            encoding="utf-8",
        )

    def write_package(
        self,
        relative: str,
        *,
        name: str = "package",
        version: str = "version.workspace = true",
        edition: str = "edition.workspace = true",
        package_extra: str = "",
        manifest_extra: str = "",
        files: dict[str, str] | None = None,
    ) -> Path:
        package_root = self.root / relative
        package_root.mkdir(parents=True, exist_ok=True)
        package_root.joinpath("Cargo.toml").write_text(
            "[package]\n"
            f'name = "{name}"\n'
            f"{version}\n"
            f"{edition}\n"
            f"{package_extra}\n"
            f"{manifest_extra}\n",
            encoding="utf-8",
        )
        for file_name, contents in (files or {}).items():
            path = package_root / file_name
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(contents, encoding="utf-8")
        return package_root

    def assert_failure(self, expected: str) -> examples.CheckFailure:
        with self.assertRaises(examples.CheckFailure) as context:
            examples.check_workspace(self.root)
        self.assertIn(expected, str(context.exception))
        return context.exception

    def test_check_uses_no_cargo_command_or_network(self) -> None:
        self.write_root(["package"])
        self.write_package("package", files={"examples/ok.rs": "fn main() {}\n"})

        with (
            patch.object(subprocess, "run", side_effect=AssertionError("Cargo used")),
            patch.object(socket, "socket", side_effect=AssertionError("network used")),
        ):
            targets = examples.check_workspace(self.root)

        self.assertEqual([target.name for target in targets], ["ok"])

    def test_workspace_members_require_string_array_and_existing_globs(self) -> None:
        self.root.joinpath("Cargo.toml").write_text(
            '[workspace]\nmembers = "package"\n', encoding="utf-8"
        )
        self.assert_failure("workspace.members must be an array of strings")

        self.write_root(["missing/*"])
        self.assert_failure("workspace.members[0] matched no paths")

        self.write_root(["package/[foo"])
        self.assert_failure("unmatched opening bracket")

    def test_workspace_excludes_are_validated_and_remove_members(self) -> None:
        self.write_root(["a", "b"], excludes=["b"])
        self.write_package("a", name="a", files={"examples/a.rs": "fn main() {}\n"})
        self.write_package("b", name="b", files={"examples/b.rs": "fn main() {}\n"})

        targets = examples.check_workspace(self.root)

        self.assertEqual(
            [(target.package.name, target.name) for target in targets], [("a", "a")]
        )

        self.write_root(["a"], excludes=["missing"])
        self.assert_failure("workspace.exclude[0] matched no paths")

    def test_package_manifest_requires_membership_and_valid_version(self) -> None:
        self.write_root(["package"])
        self.write_package("package", version="version = 7")
        self.assert_failure("package/Cargo.toml package.version must be a string")

        self.write_package(
            "package",
            version="version.workspace = true",
            package_extra='workspace = "../other"',
        )
        self.root.joinpath("other").mkdir()
        self.assert_failure("package.workspace does not point to the workspace root")

        self.write_package("package", package_extra="autoexamples = \"false\"")
        self.assert_failure("package.autoexamples must be a boolean")

        self.write_root(["package"])
        self.root.joinpath("package/Cargo.toml").write_text(
            "[lib]\npath = \"src/lib.rs\"\n", encoding="utf-8"
        )
        self.assert_failure("package/Cargo.toml [package] must be a TOML table")

    def test_workspace_package_inheritance_requires_valid_version(self) -> None:
        self.write_root(
            ["package"], workspace_package="version = 7\nedition = \"2024\""
        )
        self.write_package("package")
        self.assert_failure("workspace.package.version must be a non-empty string")

        self.write_root(["package"], workspace_package='version = "0.0.1"')
        self.write_package("package")
        self.assert_failure("package/Cargo.toml package.edition inherits")

    def test_explicit_example_fields_and_missing_paths_fail_closed(self) -> None:
        self.write_root(["package"])
        self.write_package(
            "package",
            manifest_extra=(
                "[[example]]\n"
                'name = "bad"\n'
                'required-features = "feature"\n'
            ),
            files={"examples/bad.rs": "fn main() {}\n"},
        )
        self.assert_failure("required-features must be an array of strings")

        self.write_package(
            "package",
            manifest_extra=(
                '[[example]]\nname = "missing"\npath = "examples/missing.rs"\n'
            ),
        )
        self.assert_failure("package/Cargo.toml example[0].path does not exist")

        self.write_package(
            "package",
            manifest_extra='[[example]]\nname = "bad"\npath = 3\n',
        )
        self.assert_failure(
            "package/Cargo.toml example[0].path must be a non-empty string"
        )

    def test_same_package_duplicate_records_and_default_ambiguity_fail(self) -> None:
        self.write_root(["package"])
        self.write_package(
            "package",
            manifest_extra=(
                '[[example]]\nname = "same"\npath = "examples/other.rs"\n'
                '[[example]]\nname = "same"\npath = "examples/same.rs"\n'
            ),
            files={
                "examples/same.rs": "fn main() {}\n",
                "examples/other.rs": "fn main() {}\n",
            },
        )
        failure = self.assert_failure("same-package duplicate example target 'same'")
        self.assertIn("explicit:package/examples/other.rs", str(failure))
        self.assertIn("explicit:package/examples/same.rs", str(failure))

        self.write_package(
            "package",
            manifest_extra='[[example]]\nname = "ambiguous"\n',
            files={
                "examples/ambiguous.rs": "fn main() {}\n",
                "examples/ambiguous/main.rs": "fn main() {}\n",
            },
        )
        self.assert_failure("has ambiguous default source paths")

    def test_explicit_and_auto_same_source_are_one_target(self) -> None:
        self.write_root(["package"])
        self.write_package(
            "package",
            manifest_extra='[[example]]\nname = "same"\n',
            files={"examples/same.rs": "fn main() {}\n"},
        )

        targets = examples.check_workspace(self.root)

        self.assertEqual(len(targets), 1)
        self.assertEqual(targets[0].origin, "explicit")
        self.assertEqual(targets[0].source, self.root / "package/examples/same.rs")

        self.write_package(
            "package",
            manifest_extra=(
                '[[example]]\nname = "same"\npath = "examples/other.rs"\n'
            ),
            files={
                "examples/same.rs": "fn main() {}\n",
                "examples/other.rs": "fn main() {}\n",
            },
        )
        targets = examples.check_workspace(self.root)
        self.assertEqual(targets[0].source, self.root / "package/examples/other.rs")

    def test_auto_examples_include_nested_main_and_detect_collisions(self) -> None:
        self.write_root(["package"])
        self.write_package(
            "package",
            files={
                "examples/top.rs": "fn main() {}\n",
                "examples/nested/main.rs": "fn main() {}\n",
            },
        )
        targets = examples.check_workspace(self.root)
        self.assertEqual([target.name for target in targets], ["nested", "top"])

        self.write_package(
            "package",
            files={
                "examples/collision.rs": "fn main() {}\n",
                "examples/collision/main.rs": "fn main() {}\n",
            },
        )
        self.assert_failure("same-package duplicate example target 'collision'")

    def test_identical_package_names_keep_distinct_manifest_identity(self) -> None:
        self.write_root(["z", "a"])
        self.write_package(
            "z", name="same", files={"examples/shared.rs": "fn main() {}\n"}
        )
        self.write_package(
            "a", name="same", files={"examples/shared.rs": "fn main() {}\n"}
        )

        failure = self.assert_failure("cross-package duplicate example target 'shared'")

        self.assertIn("same@a/Cargo.toml: a/examples/shared.rs", str(failure))
        self.assertIn("same@z/Cargo.toml: z/examples/shared.rs", str(failure))

    def test_path_escape_and_duplicate_member_ids_fail(self) -> None:
        self.write_root(["package"])
        outside = self.root.parent / f"{self.root.name}-outside.rs"
        outside.write_text("fn main() {}\n", encoding="utf-8")
        try:
            self.write_package(
                "package",
                manifest_extra=(
                    '[[example]]\nname = "escape"\npath = "../../'
                    f'{outside.name}"\n'
                ),
            )
            self.assert_failure("example[0].path escapes the workspace")
        finally:
            outside.unlink()

        self.write_root(["package", "package"])
        self.write_package("package", files={"examples/ok.rs": "fn main() {}\n"})
        self.assert_failure("workspace member path is listed more than once")

    def test_diagnostics_are_sorted_by_target_then_package_identity(self) -> None:
        self.write_root(["z", "a"])
        for relative in ("z", "a"):
            self.write_package(
                relative,
                name="same",
                files={
                    "examples/zulu.rs": "fn main() {}\n",
                    "examples/alpha.rs": "fn main() {}\n",
                },
            )

        failure = self.assert_failure("cross-package duplicate example target")
        self.assertEqual(tuple(failure.diagnostics), tuple(sorted(failure.diagnostics)))
        self.assertEqual(
            [diagnostic.split("'")[1] for diagnostic in failure.diagnostics],
            ["alpha", "zulu"],
        )


if __name__ == "__main__":
    unittest.main()
