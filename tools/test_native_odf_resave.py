#!/usr/bin/env python3
"""Registry-backed tests for the isolated native ODF resave harness."""

from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from tools import native_odf_resave


ROOT = Path(__file__).resolve().parent.parent
FILTER_DIRECTORY = (
    ROOT
    / "3rdparty"
    / "libreoffice-core"
    / "filter"
    / "source"
    / "config"
    / "fragments"
    / "filters"
)


class NativeOdfResaveTest(unittest.TestCase):
    def test_supported_filters_match_libreoffice_primary_registry(self) -> None:
        for filter_spec in native_odf_resave.FILTERS.values():
            registry = FILTER_DIRECTORY / f"{filter_spec.filter_name}.xcu"
            text = registry.read_text(encoding="utf-8")
            self.assertIn(f'<node oor:name="{filter_spec.filter_name}"', text)
            self.assertIn("IMPORT EXPORT", text)
            self.assertIn(
                f"<value>{filter_spec.document_service}</value>",
                text,
            )

    def test_odi_has_no_registered_type_or_filter(self) -> None:
        registry_root = FILTER_DIRECTORY.parent
        registry_text = "".join(
            path.read_text(encoding="utf-8")
            for path in registry_root.rglob("*.xcu")
        )
        self.assertNotIn("application/vnd.oasis.opendocument.image", registry_text)
        self.assertNotIn("<value>odi</value>", registry_text)

    def test_odb_filter_is_import_only(self) -> None:
        registry = FILTER_DIRECTORY / "StarOffice_XML__Base_.xcu"
        text = registry.read_text(encoding="utf-8")
        self.assertIn('<node oor:name="StarOffice XML (Base)"', text)
        self.assertIn("<value>IMPORT OWN", text)
        self.assertNotIn("IMPORT EXPORT", text)

    def test_command_uses_isolated_profile_and_exact_filter(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "changed.odf"
            output = root / "out"
            profile = root / "profile"
            command = native_odf_resave.command_for(
                "/opt/libreoffice/program/soffice", source, output, profile
            )
        self.assertIn("-env:UserInstallation=" + profile.as_uri(), command)
        self.assertIn("odf:math8", command)
        self.assertEqual(command[-1], str(source))


if __name__ == "__main__":
    unittest.main()
