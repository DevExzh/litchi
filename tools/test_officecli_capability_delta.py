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
    """A registry export (``litchi_ops::schema``): property gaps sit on their
    element as ``propertyGaps``; a format without an adapter is ``pending``."""
    elements = [dict(item) for item in elements]
    rest = []
    for gap in gaps or []:
        if gap.get("subject") != "property":
            rest.append(gap)
            continue
        owner = next(item for item in elements if item["element"] == gap["element"])
        owner.setdefault("propertyGaps", {})[gap["key"]] = {
            name: gap[name] for name in ("code", "followUp", "reason")
        }
    export = {"format": "xlsx", "elements": elements, "gaps": rest}
    if not get:
        export["pending"] = {
            "supported": False,
            "code": "unsupported_capability",
            "followUp": "M1 R4a",
            "reason": "the adapter has not landed",
        }
    return export


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

    def test_legacy_and_rtf_formats_use_their_family_vocabulary(self) -> None:
        self.assertEqual(delta.reference_format("doc"), "docx")
        self.assertEqual(delta.reference_format("rtf"), "docx")
        self.assertEqual(delta.reference_format("ppt"), "pptx")
        self.assertEqual(delta.reference_format("pptx"), "pptx")

    def test_every_registry_snapshot_is_a_table_format(self) -> None:
        listed = delta.formats(delta.REPO)
        for fmt in ("doc", "docx", "ppt", "pptx", "rtf", "xls", "xlsb", "xlsx"):
            self.assertIn(fmt, listed)
        self.assertFalse(any(fmt.endswith(delta.OPTIONAL_SUFFIX) for fmt in listed))
        root = Path(self.tempdir.name)
        for name in ("odt", "docx", "docx-encryption", "ppt"):
            write(root / delta.GOLDEN_DIR / f"{name}.json", {})
        self.assertEqual(delta.formats(root), ["docx", "ppt", "odt"])

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

    def test_a_vendored_tree_reports_its_pinned_commit(self) -> None:
        root = Path(self.tempdir.name)
        self.assertEqual(delta.officecli_commit(root), "unknown")
        (root / delta.UPSTREAM_COMMIT).write_text("0123456789abcdef\n", encoding="utf-8")
        self.assertEqual(delta.officecli_commit(root), "01234567")

    def test_the_vendored_schemas_match_the_committed_table(self) -> None:
        vendored = delta.REPO / delta.VENDORED_OFFICECLI
        tree, count = delta.schema_tree_hash(vendored / "schemas" / "help")
        commit = delta.officecli_commit(vendored)
        table = (delta.REPO / delta.OUTPUT).read_text(encoding="utf-8")
        self.assertIn(
            f"| officecli | `schemas/help` ({count} files, OfficeCLI commit `{commit}`) | `{tree}` |",
            table,
        )


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
        self.assertEqual([row[4] for row in compared.properties], ["supported", "raw-only"])

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
        self.assertEqual(properties[("cell", "font")], ("write-only", ""))
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
        snapshots = sorted((delta.REPO / delta.GOLDEN_DIR).glob("*.json"))
        self.assertGreaterEqual(len(snapshots), 8)
        for path in snapshots:
            relative = path.relative_to(delta.REPO)
            digest = hashlib.sha256(path.read_bytes()).hexdigest()
            self.assertIn(f"| golden | `{relative}` | `{digest}` |", table)
        for fmt in delta.formats(delta.REPO):
            self.assertIn(f"\n## {fmt}\n", table)

    def test_undeclared_element_gaps_are_gaps_not_missing(self) -> None:
        officecli = self.schemas(
            self.oc("comment", {"text": {"type": "string", "set": True, "get": True}}),
            self.oc("gadget", {"size": {"type": "int", "get": True}}),
        )
        gaps = [{"subject": "element", "element": "comment", "code": "unsupported_capability",
                 "followUp": "0248 C11", "reason": "r"}]
        result = delta.compare("doc", officecli, report([element("paragraph", {})], gaps=gaps))
        elements = {row[0]: (row[2], row[3]) for row in result.elements if row[0]}
        self.assertEqual(elements["comment"], ("gap", "0248 C11"))
        self.assertEqual(elements["gadget"], ("missing", ""))
        row = next(row for row in result.properties if row[0] == "comment")
        self.assertEqual((row[4], row[5], row[7]), ("element gap", "0248 C11", "gap"))

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

    def test_write_statuses_compare_add_and_set_flags(self) -> None:
        officecli = self.schemas(
            self.oc(
                "cell",
                {
                    "value": {"type": "string", "add": True, "set": True, "get": True},
                    "font": {"type": "string", "add": True, "set": True, "get": True},
                    "fill": {"type": "color", "set": True, "get": True},
                    "width": {"type": "int", "get": True},
                    "note": {"type": "string", "set": True},
                },
            ),
            self.oc("gadget", {"size": {"type": "int", "set": True}}),
        )
        litchi = report(
            [
                element(
                    "cell",
                    {
                        "value": {"type": "string", "add": True, "set": True, "get": True},
                        "font": {"type": "string", "set": True, "get": True},
                        "width": READ,
                        "note": {"type": "string", "set": True},
                        "hidden": {"type": "bool", "set": True, "get": True},
                    },
                )
            ]
        )
        rows = {(row[0], row[1]): row for row in delta.compare("xlsx", officecli, litchi).properties}
        status = {key: (row[4], row[6], row[7]) for key, row in rows.items()}
        self.assertEqual(status[("cell", "value")], ("supported", "asg", "supported"))
        self.assertEqual(status[("cell", "font")], ("supported", "-sg", "partial"))
        self.assertEqual(status[("cell", "fill")], ("missing", "", "gap"))
        self.assertEqual(status[("cell", "width")], ("supported", "--g", "—"))
        self.assertEqual(status[("cell", "note")], ("write-only", "-s-", "supported"))
        self.assertEqual(status[("cell", "hidden")], ("litchi-only", "-sg", "litchi-only"))
        self.assertEqual(status[("gadget", "size")], ("write-only", "", "gap"))

    def test_a_litchi_write_only_property_is_not_read(self) -> None:
        officecli = self.schemas(self.oc("cell", {"style": {"type": "string", "set": True, "get": True}}))
        litchi = report([element("cell", {"style": {"type": "string", "set": True}})])
        row = delta.compare("xlsx", officecli, litchi).properties[0]
        self.assertEqual((row[4], row[7]), ("write-only", "supported"))

    def test_element_gaps_are_write_gaps(self) -> None:
        officecli = self.schemas(self.oc("comment", {"text": {"type": "string", "set": True, "get": True}}))
        litchi = report([element("comment", {}, gap={"code": "unsupported_capability", "followUp": "F", "reason": "r"})])
        row = delta.compare("xlsx", officecli, litchi).properties[0]
        self.assertEqual((row[4], row[7]), ("element gap", "gap"))

    def test_the_header_describes_write_statuses(self) -> None:
        rendered = delta.render([delta.compare("xlsx", [], report([]))], [], ("cd" * 32, 0), "unknown")
        self.assertNotIn("read-only", rendered)
        self.assertIn("Property writes:", rendered)


if __name__ == "__main__":
    unittest.main()
