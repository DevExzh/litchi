from __future__ import annotations

import hashlib
import json
import tempfile
import unittest
from pathlib import Path

from tools import officecli_capability_delta as delta


def write(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data), encoding="utf-8")


def report(elements: list[dict], gaps: list[dict] | None = None, get: bool = True) -> dict:
    verb = (
        {"supported": True}
        if get
        else {
            "supported": False,
            "code": "unsupported_capability",
            "followUp": "M1 R4a",
            "reason": "the adapter has not landed",
        }
    )
    return {"format": "xlsx", "verbs": {"get": verb}, "elements": elements, "gaps": gaps or []}


def element(name: str, properties: dict, **extra) -> dict:
    return {"element": name, "properties": properties, **extra}


READ = {"type": "bool", "add": False, "set": False, "get": True}


class SchemaLoading(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.help = Path(self.tempdir.name) / "schemas" / "help"

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    def test_extends_merges_as_officecli_does(self) -> None:
        write(self.help / "_shared" / "a.json", {"note": "a", "properties": {"x": {"type": "string"}, "y": {"type": "bool"}}})
        write(self.help / "_shared" / "b.json", {"properties": {"y": {"type": "number"}}})
        write(
            self.help / "xlsx" / "thing.json",
            {"element": "thing", "extends": ["_shared/a", "_shared/b"], "properties": {"z": {"type": "color"}}},
        )
        merged = delta.load_schema(self.help, self.help / "xlsx" / "thing.json")
        self.assertEqual(list(merged["properties"]), ["z", "y", "x"])
        self.assertEqual(merged["properties"]["y"]["type"], "number")
        self.assertEqual(merged["note"], "a")
        self.assertNotIn("extends", merged)

    def test_xlsb_uses_xlsx_schema_reference(self) -> None:
        write(self.help / "xlsx" / "cell.json", {"element": "cell", "properties": {"value": READ}})
        self.assertEqual(delta.officecli_elements(self.help, "xlsb"), delta.officecli_elements(self.help, "xlsx"))
        self.assertEqual(delta.reference_format("xlsb"), "xlsx")

    def test_xls_uses_xlsx_schema_reference(self) -> None:
        write(self.help / "xlsx" / "cell.json", {"element": "cell", "properties": {"value": READ}})
        self.assertEqual(delta.officecli_elements(self.help, "xls"), delta.officecli_elements(self.help, "xlsx"))
        self.assertEqual(delta.reference_format("xls"), "xlsx")

    def test_path_forms_name_their_element(self) -> None:
        names = delta.path_names(
            {
                "paths": {"positional": ["/SheetName/chart[N]/series[K]", "/Sheet1/A1"]},
                "addressing": {"key": "role", "pathForm": "/SheetName/chart[N]/axis[@role=ROLE]"},
            }
        )
        self.assertEqual(names, ["series", "axis"])

    def test_the_schema_tree_hash_covers_names_and_content(self) -> None:
        write(self.help / "xlsx" / "a.json", {"element": "a"})
        first, count = delta.schema_tree_hash(self.help)
        self.assertEqual(count, 1)
        write(self.help / "xlsx" / "a.json", {"element": "b"})
        self.assertNotEqual(delta.schema_tree_hash(self.help)[0], first)


class Comparison(unittest.TestCase):
    def schemas(self, *elements: delta.OcElement) -> list[delta.OcElement]:
        return list(elements)

    def oc(self, name: str, properties: dict, aliases: list[str] | None = None) -> delta.OcElement:
        return delta.OcElement(name, name, aliases or [], ["sheet"], {}, properties)

    def test_xlsb_uses_spreadsheet_property_aliases_and_raw_policy(self) -> None:
        schemas = self.schemas(self.oc("table", {"showFirstColumn": READ}), self.oc("cell", {"numFmtId": READ}))
        capability = report([element("table", {"firstCol": {"get": True}}), element("cell", {})])
        compared = delta.compare("xlsb", schemas, capability)
        self.assertEqual(compared.name, "xlsb")
        self.assertEqual([row[-2] for row in compared.properties], ["supported", "raw-only"])

    def test_statuses(self) -> None:
        officecli = self.schemas(
            self.oc(
                "cell",
                {
                    "font.bold": {"type": "bool", "aliases": ["bold"], "get": True},
                    "value": {"type": "string", "get": True},
                    "computedValue": {"type": "string", "get": True},
                    "image.fileSize": {"type": "int", "get": True},
                    "font": {"type": "string", "set": True},
                },
            ),
            self.oc("comment", {"text": {"type": "string", "get": True}}, ["note"]),
            self.oc("raw", {}),
            self.oc("gadget", {"size": {"type": "int", "get": True}}),
        )
        litchi = report(
            [
                element(
                    "cell",
                    {"font.bold": READ, "border.left.color": READ},
                    predicates={"keys": ["style", "text", "type", "value"]},
                ),
                element("comment", {}, elementAliases=["note"], gap={"code": "unsupported_capability", "followUp": "L-XLSX-R5", "reason": "r"}),
                element("mergeCell", {}, gap={"code": "unsupported_capability", "followUp": "R4x", "reason": "r"}),
            ],
            gaps=[{"subject": "property", "element": "cell", "key": "computedValue", "code": "unsupported_capability", "followUp": "M4", "reason": "r"}],
        )
        result = delta.compare("xlsx", officecli, litchi)
        elements = {row[0] or row[1]: row[2] for row in result.elements}
        self.assertEqual(
            elements,
            {"cell": "supported", "comment": "gap", "raw": "out of scope", "gadget": "missing", "mergeCell": "litchi-only"},
        )
        properties = {(row[0], row[1]): (row[4], row[5]) for row in result.properties}
        self.assertEqual(properties[("cell", "font.bold")], ("supported", ""))
        self.assertEqual(properties[("cell", "value")], ("supported", "the node's text"))
        self.assertEqual(properties[("cell", "computedValue")], ("gap", "M4"))
        self.assertEqual(properties[("cell", "image.fileSize")], ("missing", ""))
        self.assertEqual(properties[("cell", "font")], ("write-only", "M2"))
        self.assertEqual(properties[("cell", "border.left.color")], ("litchi-only", ""))
        self.assertEqual(properties[("comment", "text")], ("element gap", "L-XLSX-R5"))
        self.assertEqual(properties[("gadget", "size")], ("missing", ""))

    def test_native_selection_metadata_is_a_deliberate_policy_difference(self) -> None:
        schemas = [self.oc("paragraph", {"paraId": READ}), self.oc("sdt", {"id": READ})]
        result = delta.compare("docx", schemas, report([element("paragraph", {}), element("sdt", {})]))
        self.assertTrue(all(row[4] == "raw-only" for row in result.properties))
        result = delta.compare("xlsx", [self.oc("cell", {"numFmtId": READ})], report([element("cell", {})]))
        self.assertEqual(result.properties[0][4], "raw-only")

    def test_contextual_operation_gaps_are_retained_and_rendered(self) -> None:
        gaps = [
            {
                "subject": "operation",
                "verb": "set",
                "element": "paragraph/run",
                "context": "mixed-paragraph-run-fanout",
                "code": "unsupported_capability",
                "followUp": "M2 (homogeneous mutation fanout)",
                "reason": "one mutation admits one target kind",
            }
        ]
        result = delta.compare("docx", [], report([element("paragraph", {})], gaps=gaps))
        self.assertEqual(
            result.operation_gaps,
            [
                (
                    "set",
                    "paragraph/run",
                    "mixed-paragraph-run-fanout",
                    "one mutation admits one target kind",
                    "M2 (homogeneous mutation fanout)",
                )
            ],
        )
        rendered = delta.render([result], [], ("cd" * 32, 0), "unknown")
        self.assertIn("### Contextual operation gaps", rendered)
        self.assertIn(
            "| `set` | `paragraph/run` | `mixed-paragraph-run-fanout` | one mutation admits one target kind | M2 (homogeneous mutation fanout) |",
            rendered,
        )

    def test_optional_reports_add_only_operation_gaps(self) -> None:
        existing = {
            "gaps": [
                {
                    "subject": "operation",
                    "verb": "create",
                    "element": "document",
                    "context": "strict-conformance",
                    "code": "unsupported_capability",
                    "followUp": "Transitional",
                    "reason": "default",
                }
            ]
        }
        optional = {
            "gaps": [
                existing["gaps"][0],
                {
                    "subject": "operation",
                    "verb": "save",
                    "element": "document",
                    "context": "encrypted-without-reencrypt",
                    "code": "document_protected",
                    "followUp": "SaveOptions::reencrypt",
                    "reason": "preserve encryption",
                },
                {
                    "subject": "property",
                    "element": "document",
                    "key": "ignored",
                },
            ]
        }
        merged = delta.merge_optional_operation_gaps(existing, [optional])
        self.assertEqual(len(merged["gaps"]), 2)
        self.assertEqual(merged["gaps"][1]["context"], "encrypted-without-reencrypt")
        self.assertEqual(len(existing["gaps"]), 1)

    def test_generated_table_tracks_current_golden_reports(self) -> None:
        table = (delta.REPO / delta.OUTPUT).read_text()
        for fmt in delta.FORMATS:
            relative = delta.GOLDEN_DIR / f"{fmt}.json"
            digest = hashlib.sha256((delta.REPO / relative).read_bytes()).hexdigest()
            self.assertIn(f"| golden | `{relative}` | `{digest}` |", table)
        relative = delta.GOLDEN_DIR / "docx-encryption.json"
        digest = hashlib.sha256((delta.REPO / relative).read_bytes()).hexdigest()
        self.assertIn(f"| golden | `{relative}` | `{digest}` |", table)

    def test_a_pending_adapter_lists_every_schema_as_pending(self) -> None:
        officecli = self.schemas(self.oc("paragraph", {"align": {"type": "enum", "get": True}}))
        result = delta.compare("docx", officecli, report([], get=False))
        self.assertEqual(result.elements, [("paragraph", "", "pending", "M1 R4a")])
        self.assertEqual(result.properties[0][4], "pending")

    def test_run_subtypes_are_not_reported_as_supported_runs(self) -> None:
        result = delta.compare("docx", [self.oc("fieldChar", {"fieldCharType": READ}, ["r"])],
                               report([element("run", {}, elementAliases=["r"])]))
        self.assertEqual(result.elements[0][2], "missing")

    def test_pending_formats_still_exclude_raw_and_set_only_actions(self) -> None:
        result = delta.compare("xlsx", [self.oc("raw", {"xml": READ}), self.oc("sort", {})], report([], get=False))
        self.assertTrue(all(row[2] == "out of scope" for row in result.elements))
        self.assertEqual(result.properties, [])

    def test_rendering_is_deterministic_and_records_the_inputs(self) -> None:
        officecli = self.schemas(self.oc("cell", {"font.bold": {"type": "bool", "get": True}}))
        result = delta.compare("xlsx", officecli, report([element("cell", {"font.bold": READ})]))
        goldens = [("crates/litchi-ops/tests/golden/capabilities/xlsx.json", "ab" * 32)]
        first = delta.render([result], goldens, ("cd" * 32, 3), "dced0d74")
        self.assertEqual(first, delta.render([result], goldens, ("cd" * 32, 3), "dced0d74"))
        self.assertIn(f"| golden | `{goldens[0][0]}` | `{'ab' * 32}` |", first)
        self.assertIn("| xlsx | 1 | 0 | 0 | 0 | 0 | 0 | 0 |", first)


if __name__ == "__main__":
    unittest.main()
