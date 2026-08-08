from __future__ import annotations

import os
import subprocess
import tempfile
import unittest
from pathlib import Path

from tools import check_numbers_public_dependencies as public_dependencies


class NumbersPublicDependencyGateTests(unittest.TestCase):
    def test_command_marks_only_numbers_wire_private(self) -> None:
        self.assertEqual(
            public_dependencies.command(),
            (
                "cargo",
                "rustc",
                "--locked",
                "--package",
                "litchi-numbers",
                "--lib",
                "--",
                "-Zunstable-options",
                "--extern",
                "priv:litchi_numbers_wire",
                "-Dexported-private-dependencies",
            ),
        )
        self.assertNotIn(
            "priv:litchi_iwa_common",
            public_dependencies.command(),
        )

    def test_environment_enables_the_stable_compiler_modifier(self) -> None:
        source = {"PATH": "/compiler/bin", "RUSTC_BOOTSTRAP": "old"}

        result = public_dependencies.environment(source)

        self.assertEqual(result["RUSTC_BOOTSTRAP"], "1")
        self.assertEqual(result["PATH"], "/compiler/bin")
        self.assertEqual(source["RUSTC_BOOTSTRAP"], "old")

    def test_compiler_allows_semantic_dependency_and_rejects_from_wire(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            semantic = self._compile_dependency(
                root,
                "semantic_common",
                "#[derive(Clone, Copy)] pub struct Point(pub f64, pub f64);\n",
            )
            wire = self._compile_dependency(
                root,
                "numbers_wire",
                "pub struct Error;\n",
            )

            allowed = root / "allowed.rs"
            allowed.write_text(
                "pub fn identity(point: semantic_common::Point) "
                "-> semantic_common::Point { point }\n",
                encoding="utf-8",
            )
            allowed_result = self._compile_host(
                allowed,
                root / "liballowed.rlib",
                semantic,
                wire,
            )

            leaking = root / "leaking.rs"
            leaking.write_text(
                "pub struct LocalError;\n"
                "impl From<numbers_wire::Error> for LocalError {\n"
                "    fn from(_: numbers_wire::Error) -> Self { Self }\n"
                "}\n",
                encoding="utf-8",
            )
            leaking_result = self._compile_host(
                leaking,
                root / "libleaking.rlib",
                semantic,
                wire,
            )

        self.assertEqual(allowed_result.returncode, 0, allowed_result.stderr)
        self.assertNotEqual(leaking_result.returncode, 0)
        self.assertIn("private dependency 'numbers_wire'", leaking_result.stderr)
        self.assertIn("impl From<numbers_wire::Error>", leaking_result.stderr)
        self.assertNotIn("private dependency 'semantic_common'", leaking_result.stderr)

    @staticmethod
    def _compile_dependency(root: Path, name: str, source: str) -> Path:
        source_path = root / f"{name}.rs"
        output_path = root / f"lib{name}.rlib"
        source_path.write_text(source, encoding="utf-8")
        completed = subprocess.run(
            (
                "rustc",
                "--crate-name",
                name,
                "--crate-type",
                "rlib",
                "--edition=2024",
                str(source_path),
                "-o",
                str(output_path),
            ),
            env=public_dependencies.environment(),
            check=False,
            capture_output=True,
            text=True,
        )
        if completed.returncode != 0:
            raise AssertionError(completed.stderr)
        return output_path

    @staticmethod
    def _compile_host(
        source: Path,
        output: Path,
        semantic: Path,
        wire: Path,
    ) -> subprocess.CompletedProcess[str]:
        return subprocess.run(
            (
                "rustc",
                "--crate-name",
                source.stem,
                "--crate-type",
                "rlib",
                "--edition=2024",
                str(source),
                "--extern",
                f"semantic_common={semantic}",
                "--extern",
                f"numbers_wire={wire}",
                "-Zunstable-options",
                "--extern",
                f"priv:numbers_wire={wire}",
                "-Dexported-private-dependencies",
                "-o",
                str(output),
            ),
            env=public_dependencies.environment(os.environ),
            check=False,
            capture_output=True,
            text=True,
        )


if __name__ == "__main__":
    unittest.main()
