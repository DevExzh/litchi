#!/usr/bin/env python3
"""Registry-backed tests for the isolated native Office resave harness."""

from __future__ import annotations

import hashlib
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

EVIDENCE_SHA256 = {
    "test-data/ole/doc/NoHeadFoot.doc": (
        "45e5df073f34314da6f39d2dad119fb2ef23470878fd2df67f632864cd92ea48"
    ),
    "test-data/office-interop/litchi-changed/noheadfoot-litchi.doc": (
        "152da496f5b376a0d0430bfbc87658a9d1f2f0afc25c592238ec52800347249e"
    ),
    "test-data/office-interop/libreoffice-resaved/noheadfoot-litchi.doc": (
        "02f6b96ed94e027e652df5d9e527ecee825c958db8ddf9bd398ecb7b0870aa35"
    ),
    "test-data/libreoffice-core/sc/qa/extras/testdocuments/tdf78897.xls": (
        "940fb6f143e8d54c545e62599dd7c38e45846db6787397283b7ec93e70eb96ae"
    ),
    "test-data/office-interop/litchi-changed/tdf78897-litchi.xls": (
        "8f881ec4ccdef867154424f296665c37cdd114456c7c9ec5efe85ff460da870d"
    ),
    "test-data/office-interop/libreoffice-resaved/tdf78897-litchi.xls": (
        "eab8ae4797e6499e5c45c01b0ad55502c021103385c6fd14271e91b81e9ec537"
    ),
    "test-data/ooxml/docx/documentProperties.docx": (
        "1cff7a0a94dfce307a70032d21070d26ae34b9fdf742cf70fa66d4a2078ec9d5"
    ),
    "test-data/office-interop/litchi-changed/document-properties-litchi.docx": (
        "e4c76a4cd17a2cc5e66d52fe2109ec62f0a09aa510e97fd291b87073b84697f2"
    ),
    "test-data/office-interop/libreoffice-resaved/document-properties-litchi.docx": (
        "0c10b6489e3b5c02d0470f213611e32fbaa0143746d69d0333bdca821d2c5e47"
    ),
    "test-data/libreoffice-core/sc/qa/unit/data/xlsx/dateAutofilter.xlsx": (
        "d7ab3dbb59388d245ee779bf8547748dc6bac70f3c7216e673e0d97dbbbd6bc4"
    ),
    "test-data/office-interop/litchi-changed/date-autofilter-litchi.xlsx": (
        "0f1c13c528c75b5293b18f4342d5fa95024fcc604af81c1b66c058a06575bac9"
    ),
    "test-data/office-interop/libreoffice-resaved/date-autofilter-litchi.xlsx": (
        "ade103f651e3f7cac423cd6b01d4d5a004d03470667723234105ba39672bc74d"
    ),
    "test-data/ooxml/pptx/shapes.pptx": (
        "19fde9b87e33dd1a95fdbba0cf6abc2278bf03874f4665c7f8b88b6afe4a2571"
    ),
    "test-data/office-interop/litchi-changed/shapes-litchi.pptx": (
        "41cc73dc78e6506900628ec4518a8dee85544f164f5ecb4134c5240c414928df"
    ),
    "test-data/office-interop/libreoffice-resaved/shapes-litchi.pptx": (
        "fcc8acffad88f5091316f67403c099a4c9eaa372e17927edfee29c35fb132034"
    ),
    "test-data/poi/test-data/slideshow/45543.ppt": (
        "218aaac542e5f9b567736407f2631defc65797c6ba2a7818f066e2f93bcfacaf"
    ),
    "test-data/office-interop/litchi-changed/45543-transition-litchi.ppt": (
        "88df2a6cd4bed72ee3f4f0bd224f8246b6388ab860cdf8628ab2057e6a9875b1"
    ),
    "test-data/office-interop/libreoffice-resaved/45543-transition-litchi.ppt": (
        "5ffe19b9e60886a6d2950ccb5ae3fb1bd4ca278f0bafd40b5127575dc93e8d80"
    ),
    "test-data/libreoffice-core/sw/qa/extras/rtfexport/data/relsize.rtf": (
        "1a079582281767c1bf7afa5ef2e63553400cdbc4704aa25d9dbcc34e2c22569d"
    ),
    "test-data/office-interop/litchi-changed/relsize-litchi.rtf": (
        "d0bf70e50972bbf15dc9b0da96b9702d64a92a676cb81529bf28729a2cd91d71"
    ),
    "test-data/office-interop/libreoffice-resaved/relsize-litchi.rtf": (
        "224707aea42c7b38712bc66a76424d3666b52794e2aed010fe64d5adda54f3d9"
    ),
    "test-data/libreoffice-core/dbaccess/qa/unit/data/tdf132924.odb": (
        "ef32cabf31818b2fff52a6fbabb570952e823bfec6237da402a0392546c5d5af"
    ),
    "test-data/office-interop/litchi-changed/tdf132924-litchi.odb": (
        "3af6b848500601f5bb9d3b56e421539880eedfdf5a3ed31146c53b479e55299c"
    ),
    "test-data/office-interop/libreoffice-resaved/tdf132924-litchi.odb": (
        "fbe56e2711dc1876f4b8a0b841e4a530f26e0304095f5a9d000b92fb15d0b607"
    ),
    "test-data/odf/corpus/calc-two-sheets.ods": (
        "67ed3f8831aa078a849badd8f2a15bdee7cf965a628ff4eb73740aa96ba0d4c0"
    ),
    "test-data/odf/native-resave/litchi-changed/calc-two-sheets-litchi.ods": (
        "f942913fe6057266b9233fe539e078213a46f3903ab5ab54bac5e1076a6415b0"
    ),
    "test-data/odf/native-resave/libreoffice-resaved/calc-two-sheets-litchi.ods": (
        "5d60813f36fab58a802da50fd7979b7ea503e38c8721598de093f61f60eb111b"
    ),
    "test-data/odf/native-resave/source/font-styles.odf": (
        "2abee0da450b31c3bc87d007e85fff21714fee742d6dda9e4987415107ffb27f"
    ),
    "test-data/odf/native-resave/litchi-changed/font-styles-litchi.odf": (
        "8b0ce2f1415ec28d579ae0c000fc8a4570dbb3ad9b81f48366f9654db91b6508"
    ),
    "test-data/odf/native-resave/libreoffice-resaved/font-styles-litchi.odf": (
        "4c1ca45f31b5ac919fe82b9b09962d0d57bcffc0c801f7eae3e57882f8c5ea7c"
    ),
    "test-data/odf/corpus/writer-header-footer.odt": (
        "fda7c0be9f1135e7a30b05db6d9ddf96020ba00d87478ca4c74084c8742c5a21"
    ),
    "test-data/odf/native-resave/litchi-changed/writer-header-footer-litchi.odt": (
        "2dd3b2047c89da3352adb7f3b4db027ffdcc77a2e70f4c298645e7815619f952"
    ),
    "test-data/odf/native-resave/libreoffice-resaved/writer-header-footer-litchi.odt": (
        "22e95ca413c468c8ec4de96f23f934ff6a4ee860b5089c512dafb2a0d7f74c32"
    ),
    "test-data/odf/odp/tdf169979.odp": (
        "160908f993c6ba901233695b12d34c4b009142971b36dcd57c0549bf8ee5656b"
    ),
    "test-data/odf/native-resave/litchi-changed/tdf169979-litchi.odp": (
        "2ebbf9efb1a0b26bc60cda62c12874cdd60f0e22027a70e2c1018c44098167fb"
    ),
    "test-data/odf/native-resave/libreoffice-resaved/tdf169979-litchi.odp": (
        "31401a94864ee0a4b68d39287cf65159b977db798946472f4546a393f1d4a4a9"
    ),
    "test-data/odf/native-resave/source/rhbz1870501.odg": (
        "46530e653ca424fd5b985813cdeeceb9f4b99589c45d8bdeb1b1256badad133f"
    ),
    "test-data/odf/native-resave/litchi-changed/rhbz1870501-litchi.odg": (
        "a7282b53e227fa772876e4b697bb70f7aa77e6d2b1384c5a2f6b483153d5de2a"
    ),
    "test-data/odf/native-resave/libreoffice-resaved/rhbz1870501-litchi.odg": (
        "4cf5a94733a11a9a7075284994a191cb87c41173899ee0ebbba592c57232d99c"
    ),
    "test-data/odf/native-resave/source/LICENSE-MPL-2.0.txt": (
        "1f256ecad192880510e84ad60474eab7589218784b9a50bc7ceee34c2b91f1d5"
    ),
}

SUCCESS_LOGS = {
    "test-data/office-interop/logs/doc-resave.log": (
        "test-data/office-interop/libreoffice-resaved/noheadfoot-litchi.doc"
    ),
    "test-data/office-interop/logs/xls-resave.log": (
        "test-data/office-interop/libreoffice-resaved/tdf78897-litchi.xls"
    ),
    "test-data/office-interop/logs/docx-resave.log": (
        "test-data/office-interop/libreoffice-resaved/document-properties-litchi.docx"
    ),
    "test-data/office-interop/logs/xlsx-date-autofilter-resave.log": (
        "test-data/office-interop/libreoffice-resaved/date-autofilter-litchi.xlsx"
    ),
    "test-data/office-interop/logs/pptx-resave.log": (
        "test-data/office-interop/libreoffice-resaved/shapes-litchi.pptx"
    ),
    "test-data/office-interop/logs/ppt-resave.log": (
        "test-data/office-interop/libreoffice-resaved/45543-transition-litchi.ppt"
    ),
    "test-data/office-interop/logs/rtf-resave.log": (
        "test-data/office-interop/libreoffice-resaved/relsize-litchi.rtf"
    ),
    "test-data/office-interop/logs/odb-uno-store.log": (
        "test-data/office-interop/libreoffice-resaved/tdf132924-litchi.odb"
    ),
    "test-data/odf/native-resave/logs/ods-resave.log": (
        "test-data/odf/native-resave/libreoffice-resaved/calc-two-sheets-litchi.ods"
    ),
    "test-data/odf/native-resave/logs/odf-resave.log": (
        "test-data/odf/native-resave/libreoffice-resaved/font-styles-litchi.odf"
    ),
    "test-data/odf/native-resave/logs/odt-resave.log": (
        "test-data/odf/native-resave/libreoffice-resaved/writer-header-footer-litchi.odt"
    ),
    "test-data/odf/native-resave/logs/odp-resave.log": (
        "test-data/odf/native-resave/libreoffice-resaved/tdf169979-litchi.odp"
    ),
    "test-data/odf/native-resave/logs/odg-resave.log": (
        "test-data/odf/native-resave/libreoffice-resaved/rhbz1870501-litchi.odg"
    ),
}


class NativeOdfResaveTest(unittest.TestCase):
    def test_evidence_artifact_checksums(self) -> None:
        for relative_path, expected in EVIDENCE_SHA256.items():
            digest = hashlib.sha256((ROOT / relative_path).read_bytes()).hexdigest()
            self.assertEqual(digest, expected, relative_path)

    def test_success_logs_pin_the_resaved_artifact(self) -> None:
        for relative_log, relative_output in SUCCESS_LOGS.items():
            records = dict(
                line.split("=", 1)
                for line in (ROOT / relative_log).read_text(encoding="utf-8").splitlines()
            )
            self.assertEqual(records["status"], "success", relative_log)
            self.assertEqual(records["output"], relative_output, relative_log)
            self.assertEqual(
                records["output_sha256"], EVIDENCE_SHA256[relative_output], relative_log
            )

    def test_supported_filters_match_libreoffice_primary_registry(self) -> None:
        for filter_spec in native_odf_resave.FILTERS.values():
            registry_name = filter_spec.registry_file or filter_spec.filter_name
            registry = FILTER_DIRECTORY / f"{registry_name}.xcu"
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

    def test_xlsb_filter_is_import_only(self) -> None:
        registry = FILTER_DIRECTORY / "calc_MS_Excel_2007_Binary.xcu"
        text = registry.read_text(encoding="utf-8")
        self.assertIn('<node oor:name="Calc MS Excel 2007 Binary"', text)
        self.assertIn("<value>IMPORT ALIEN", text)
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
