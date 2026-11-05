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

    def test_a_mermaid_diagram_is_not_the_slide_it_sits_on(self) -> None:
        result = delta.compare("pptx", [self.oc("diagram", {"mermaid": READ}, ["slide"])],
                               report([element("slide", {"mermaid": READ})]))
        self.assertEqual(result.elements[0][:3], ("diagram", "", "missing"))

    def test_an_element_matches_its_own_name_before_a_path_word(self) -> None:
        # Record 0339: `/slidemaster[N]/shape[K]` names `shape` too, but the
        # element is the slide master.
        result = delta.compare("pptx", [self.oc("slidemaster", {"name": READ}, ["shape"])],
                               report([element("shape", {}), element("slideMaster", {"name": READ})]))
        self.assertEqual(result.elements[0][:3], ("slidemaster", "slideMaster", "supported"))
        gaps = [{"subject": "element", "element": "diagram", "code": "unsupported_path",
                 "followUp": "none", "reason": "r"}]
        result = delta.compare("pptx", [self.oc("diagram", {"mermaid": READ}, ["slide"])],
                               report([element("slide", {})], gaps=gaps))
        self.assertEqual(result.elements[0][:3], ("diagram", "", "gap"))

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


def dispositions(entries: list[dict], tiers: list[dict] | None = None, **policy) -> delta.Dispositions:
    """A dispositions file: by default every xlsx element is P0 and every
    `gadget` P2."""
    if tiers is None:
        tiers = [
            {"tier": "P0", "keys": ["xlsx:*"], "basis": "b"},
            {"tier": "P2", "keys": ["xlsx:gadget"], "basis": "b"},
        ]
    return delta.Dispositions({"policy": policy, "tiers": tiers, "dispositions": entries}, "test")


def entry(id: str, keys: list[str], disposition: str, **extra) -> dict:
    return {"id": id, "keys": keys, "disposition": disposition, "reason": "r", **extra}


class DispositionFile(unittest.TestCase):
    def test_the_committed_file_loads_and_every_key_is_used(self) -> None:
        loaded = delta.Dispositions.load(delta.REPO / delta.DISPOSITIONS)
        _, _, problems = delta.classified(delta.REPO / delta.VENDORED_OFFICECLI, delta.REPO, loaded)
        self.assertEqual(problems, [])

    def test_specificity_prefers_property_then_element_then_format(self) -> None:
        patterns = delta.Patterns(
            [("delta:docx:*", "format"), ("delta:*:raw", "element"), ("delta:*:raw.*", "property"),
             ("delta:docx:raw", "both")],
            "test",
        )
        self.assertEqual(patterns.best("delta:docx:raw")[1], "both")
        self.assertEqual(patterns.best("delta:doc:raw")[1], "element")
        self.assertEqual(patterns.best("delta:doc:raw.xml")[1], "property")
        self.assertEqual(patterns.best("delta:docx:paragraph.align")[1], "format")
        self.assertEqual(patterns.unused(), [])
        # Equal specificity goes to the earlier entry.
        tied = delta.Patterns([("delta:docx:a*", 1), ("delta:docx:*b", 2)], "test")
        self.assertEqual(tied.best("delta:docx:ab")[1], 1)
        self.assertEqual(tied.unused(), ["delta:docx:*b"])
        corpus_keys = delta.Patterns([("corpus:docx|*|*|*|*", "format"), ("corpus:*|*|*|*|bold", "property")], "t")
        self.assertEqual(corpus_keys.best("corpus:docx|set|run|warned|bold")[1], "property")

    def test_a_key_is_listed_once(self) -> None:
        with self.assertRaises(SystemExit):
            dispositions([entry("a", ["delta:x:*"], "deferred"), entry("b", ["delta:x:*"], "wontfix")])

    def test_entries_are_validated(self) -> None:
        with self.assertRaises(SystemExit):
            dispositions([entry("a", ["delta:x:*"], "planned")])  # no work-package or record id
        with self.assertRaises(SystemExit):
            dispositions([{"id": "a", "keys": ["delta:x:*"], "disposition": "wontfix"}])  # no reason
        with self.assertRaises(SystemExit):
            dispositions([entry("a", ["delta:x:*"], "later")])
        with self.assertRaises(SystemExit):
            dispositions([entry("a", ["delta:x:*"], "deferred", ref="M2")])
        with self.assertRaises(SystemExit):
            dispositions([entry("a", ["delta:x:*"], "deferred"), entry("a", ["delta:y:*"], "deferred")])
        dispositions([entry("a", ["delta:x:*"], "planned", ref="WP4"), entry("b", ["delta:y:*"], "planned", ref="0324")])

    def test_milestone_only_follow_ups(self) -> None:
        for text in ("M2 (validate)", "M1-extended (R4x)", "R4x", "M1-Δ1", "M2 XLSX OfficeCLI input policy"):
            self.assertTrue(delta.milestone_only(text), text)
        for text in ("M6 ODT body index (record 0276)", "L-DOCX-W5", "0323-F7 PPT chart series", "design §10", ""):
            self.assertFalse(delta.milestone_only(text), text)

    def test_tiers(self) -> None:
        loaded = dispositions([], tiers=[
            {"tier": "P0", "keys": ["xlsx:cell"], "basis": "b"},
            {"tier": "v1", "keys": ["xlsx:chart"], "basis": "b"},
            {"tier": "P1", "keys": ["xlsx:table"], "basis": "b"},
            {"tier": "P0", "keys": ["xlsx:cell.value", "xlsx:table.style"], "basis": "b"},
        ])
        self.assertEqual(loaded.tier("xlsx", "cell"), "P0")
        self.assertEqual(loaded.tier("xlsx", "cell", "value"), "P0")
        # A P0 or v1 element's other properties are breadth (P1).
        self.assertEqual(loaded.tier("xlsx", "cell", "fill"), "P1")
        self.assertEqual(loaded.tier("xlsx", "chart", "title"), "P1")
        # A property is never in an earlier tier than its element.
        self.assertEqual(loaded.tier("xlsx", "table", "style"), "P1")
        # Outside the v1 boundary a P0 or v1 feature is P1; the vocabulary is the family's.
        self.assertEqual(loaded.tier("xls", "cell"), "P1")
        self.assertEqual(loaded.tier("xls", "cell", "value"), "P1")
        self.assertIsNone(loaded.tier("xlsx", "gadget"))
        wide = dispositions([], tiers=[{"tier": "P0", "keys": ["xlsx:cell"], "basis": "b"}], p0Formats=["xlsx", "xls"])
        self.assertEqual(wide.tier("xls", "cell"), "P0")


class MilestoneFollowUps(unittest.TestCase):
    """Record 0344 (audit WP13): every milestone-only follow-up of the
    registry and capability goldens maps to a current id."""

    def load(self, follow_ups: list[dict]) -> delta.Dispositions:
        return delta.Dispositions({"tiers": [], "dispositions": [], "followUps": follow_ups}, "test")

    def test_a_format_entry_beats_a_general_one_then_the_longer_pattern(self) -> None:
        loaded = self.load([
            {"id": "all", "followUps": ["M2 *"], "ref": "0332", "note": "n"},
            {"id": "xlsx", "followUps": ["M2 *"], "formats": ["xlsx"], "ref": "0338", "note": "n"},
            {"id": "exact", "followUps": ["M2 XLSX input policy"], "formats": ["xlsx"], "ref": "WP16", "note": "n"},
        ])
        self.assertEqual(loaded.follow_up_entry("docx", "M2 tables")["id"], "all")
        self.assertEqual(loaded.follow_up_entry("xlsx", "M2 tables")["id"], "xlsx")
        self.assertEqual(loaded.follow_up_entry("xlsx", "M2 XLSX input policy")["id"], "exact")
        self.assertIsNone(loaded.follow_up_entry("xlsx", "M1-extended"))
        self.assertEqual(loaded.current("xlsx", "M2 XLSX input policy"), "M2 XLSX input policy → WP16")
        # Text that already names a record, and unmapped text, are shown as they are.
        self.assertEqual(loaded.current("xlsx", "record 0327 follow-up"), "record 0327 follow-up")
        self.assertEqual(loaded.current("xlsx", "M1-extended"), "M1-extended")
        self.assertFalse(delta.milestone_only(loaded.current("docx", "M2 tables")))

    def test_entries_are_validated(self) -> None:
        for bad in ({"id": "a", "followUps": ["M2 *"], "ref": "M3", "note": "n"},
                    {"id": "a", "followUps": ["M2 *"], "ref": "", "note": "n"},
                    {"id": "a", "followUps": [], "ref": "0332", "note": "n"},
                    {"id": "a", "followUps": ["M2 *"], "ref": "0332"}):
            with self.assertRaises(SystemExit):
                self.load([bad])
        with self.assertRaises(SystemExit):
            self.load([{"id": "a", "followUps": ["x"], "ref": "0332", "note": "n"}] * 2)
        self.load([{"id": "a", "followUps": ["M2 *"], "ref": "0327/0316", "note": "n"}])

    def test_unmapped_and_stale_entries_fail(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write(root / delta.GOLDEN_DIR / "xlsx.json", {"gaps": [{"followUp": "M2 (validate)"}]})
            write(root / delta.CAPABILITY_DIR / "xlsx-encryption.json",
                  {"elements": [{"gap": {"followUp": "M1-extended (R4x)"}},
                                {"gap": {"followUp": "record 0338 follow-up"}}]})
            found = delta.golden_follow_ups(root)
            self.assertEqual(sorted((fmt, text) for fmt, text, _ in found),
                             [("xlsx", "M1-extended (R4x)"), ("xlsx", "M2 (validate)"),
                              ("xlsx", "record 0338 follow-up")])
            loaded = self.load([
                {"id": "validate", "followUps": ["M2 (validate)"], "ref": "0336", "note": "n"},
                {"id": "unused", "followUps": ["M6 *"], "ref": "0328", "note": "n"},
            ])
            problems = delta.milestone_problems(found, loaded)
            self.assertEqual(len(problems), 2, problems)
            self.assertIn("xlsx: milestone-only follow-up `M1-extended (R4x)`", problems[0])
            self.assertIn("crates/litchi-ops/tests/golden/capabilities/xlsx-encryption.json", problems[0])
            self.assertEqual(problems[1], "stale follow-up map `unused`: it maps no golden follow-up")

    def test_the_committed_goldens_have_no_unmapped_milestone(self) -> None:
        loaded = delta.Dispositions.load(delta.REPO / delta.DISPOSITIONS)
        found = delta.golden_follow_ups(delta.REPO)
        self.assertTrue(any(origin.startswith(delta.CAPABILITY_DIR.as_posix()) for _, _, origin in found))
        self.assertEqual(delta.milestone_problems(found, loaded), [])


class ApplyingDispositions(unittest.TestCase):
    def setUp(self) -> None:
        oc = lambda name, props: delta.OcElement(name, name, [], [], {}, props)
        self.officecli = [
            oc("cell", {"fill": {"type": "color", "set": True, "get": True},
                        "font": {"type": "string", "set": True, "get": True}}),
            oc("gadget", {"size": {"type": "int", "set": True, "get": True},
                          "spin": {"type": "int", "set": True, "get": True}}),
            oc("comment", {"text": {"type": "string", "set": True, "get": True}}),
        ]
        self.litchi = report(
            [element("cell", {}),
             element("comment", {}, gap={"code": "unsupported_capability", "followUp": "M2 (comments)", "reason": "r"})]
        )

    def apply(self, loaded: delta.Dispositions) -> tuple[delta.FormatDelta, list[str]]:
        result = delta.compare("xlsx", self.officecli, self.litchi)
        return result, delta.apply_dispositions(result, loaded)

    def test_rows_resolve_and_missing_becomes_unlisted(self) -> None:
        result, problems = self.apply(dispositions([
            entry("cells", ["delta:xlsx:cell.*"], "planned", ref="WP16"),
            entry("later", ["delta:xlsx:gadget", "delta:xlsx:comment"], "deferred", ref="0332"),
            entry("gadget-spin", ["delta:xlsx:gadget.spin"], "wontfix"),
        ]))
        self.assertEqual(problems, ["xlsx element comment: a P0 row is `deferred` (later); P0 is supported or planned"])
        elements = {row[0]: (row[2], res.cell() if res else "") for row, res in zip(result.elements, result.element_dispositions)}
        self.assertEqual(elements["gadget"], ("unlisted", "deferred-P2 0332 (later)"))
        rows = {(row[0], row[1]): (row[4], res) for row, res in zip(result.properties, result.property_dispositions)}
        self.assertEqual(rows[("cell", "fill")][0], "unlisted")
        self.assertEqual(rows[("cell", "fill")][1].disposition, "planned")
        self.assertEqual(rows[("cell", "fill")][1].tier, "P1")
        # A gap element's properties share its disposition ...
        self.assertEqual(rows[("gadget", "size")][1].entry, "later")
        self.assertEqual(rows[("gadget", "size")][1].via, "element")
        # ... unless an entry names the element and the property.
        self.assertEqual(rows[("gadget", "spin")][1].disposition, "wontfix")
        self.assertEqual(rows[("comment", "text")][1].entry, "later")

    def test_a_format_wide_property_key_does_not_override_a_gap_element(self) -> None:
        result, problems = self.apply(dispositions([
            entry("all", ["delta:xlsx:*"], "planned", ref="WP16"),
            entry("breadth", ["delta:xlsx:*.*"], "deferred", ref="0332"),
        ]))
        self.assertEqual(problems, [])
        rows = {(row[0], row[1]): res for row, res in zip(result.properties, result.property_dispositions)}
        self.assertEqual(rows[("gadget", "size")].entry, "all")
        self.assertEqual(rows[("cell", "fill")].entry, "breadth")
        self.assertEqual(rows[("cell", "fill")].disposition, "deferred-P1")

    def test_untriaged_rows_stay_missing_and_fail(self) -> None:
        result, problems = self.apply(dispositions([entry("cells", ["delta:xlsx:cell.*"], "deferred", ref="0332")]))
        self.assertIn("xlsx element gadget: no disposition", problems)
        self.assertIn("xlsx element comment: no disposition", problems)
        statuses = {row[0]: row[2] for row in result.elements}
        self.assertEqual(statuses["gadget"], "missing")
        self.assertEqual(result.element_dispositions[[row[0] for row in result.elements].index("gadget")].cell(), "**none**")

    def test_tier_mismatches_and_milestone_follow_ups_fail(self) -> None:
        _, problems = self.apply(dispositions([
            entry("cells", ["delta:xlsx:cell.*"], "deferred-P2", ref="0332"),
            entry("gadget", ["delta:xlsx:gadget"], "deferred", ref="0332"),
            entry("comment", ["delta:xlsx:comment"], "planned", ref="WP3"),
        ]))
        self.assertIn("xlsx cell.fill: `deferred-P2` on a P1 row (cells)", problems)
        _, problems = self.apply(dispositions([
            entry("cells", ["delta:xlsx:cell.*"], "deferred"),
            entry("gadget", ["delta:xlsx:gadget"], "deferred"),
            entry("comment", ["delta:xlsx:comment"], "out-of-scope"),
        ], tiers=[{"tier": "P1", "keys": ["xlsx:*"], "basis": "b"}]))
        self.assertEqual(problems, [])
        _, problems = self.apply(dispositions([
            entry("cells", ["delta:xlsx:cell.*"], "deferred"),
            entry("gadget", ["delta:xlsx:gadget"], "deferred"),
            entry("comment", ["delta:xlsx:comment"], "deferred"),
        ], tiers=[{"tier": "P1", "keys": ["xlsx:*"], "basis": "b"}]))
        self.assertEqual(problems, ["xlsx element comment: milestone-only follow-up `M2 (comments)` needs a current id (comment)"])

    def test_a_row_without_a_tier_fails(self) -> None:
        _, problems = self.apply(dispositions([entry("all", ["delta:xlsx:*"], "planned", ref="WP16")],
                                              tiers=[{"tier": "P0", "keys": ["xlsx:cell"], "basis": "b"}]))
        self.assertIn("xlsx element gadget: no tier", problems)

    def test_stale_keys_and_check(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            write(root / delta.GOLDEN_DIR / "xlsx.json", self.litchi)
            for schema in self.officecli:
                write(root / "schemas" / "help" / "xlsx" / f"{schema.name}.json",
                      {"element": schema.name, "properties": schema.properties})
            loaded = dispositions([
                entry("all", ["delta:xlsx:*"], "planned", ref="WP16"),
                entry("stale", ["delta:xlsx:nothing", "delta:docx:*"], "deferred", ref="0332"),
            ], tiers=[{"tier": "P0", "keys": ["xlsx:*"], "basis": "b"}, {"tier": "P1", "keys": ["xlsx:unused"], "basis": "b"}])
            _, _, problems = delta.classified(root, root, loaded)
            # Record 0344: the fixture's milestone-only follow-up has no current id.
            self.assertEqual(problems[0], "xlsx: milestone-only follow-up `M2 (comments)` (1 in "
                             "crates/litchi-ops/tests/golden/registry/xlsx.json) has no current id in `followUps`")
            self.assertEqual(problems[1:], [
                "stale disposition key `delta:xlsx:nothing`: it classifies no row",
                "stale disposition key `delta:docx:*`: it classifies no row",
                "stale tier key `xlsx:unused`: it classifies no row",
            ])
            path = root / "dispositions.json"
            write(path, {"tiers": [], "dispositions": []})
            self.assertEqual(delta.main(["--check", "--dispositions", str(path)]), 1)

    def test_rendering_shows_dispositions_and_property_pages(self) -> None:
        result, _ = self.apply(dispositions([
            entry("cells", ["delta:xlsx:cell.*"], "planned", ref="WP16"),
            entry("later", ["delta:xlsx:gadget", "delta:xlsx:comment"], "planned", ref="WP3"),
        ]))
        loaded = dispositions([entry("cells", ["delta:xlsx:cell.*"], "planned", ref="WP16")])
        main = delta.render([result], [], ("cd" * 32, 0), "unknown", loaded)
        self.assertIn("| gadget |  | unlisted | P2 | planned WP3 (later) |  |", main)
        self.assertIn("| xlsx | 2 | 0 | 0 | 0 | 0 | 0 |", main)  # element gap rows by disposition
        self.assertIn("officecli-capability-delta/xlsx.md", main)
        self.assertIn("| cells | planned | WP16 |", main)
        page = delta.render_properties(result)
        self.assertIn("| cell | `fill` | color | -sg |  | unlisted | gap | P1 | planned WP16 (cells) |  |", page)
        self.assertTrue(all(len(text.encode()) < delta.SIZE_LIMIT for text in (main, page)))

    def test_the_committed_tables_pass_check(self) -> None:
        self.assertEqual(delta.main(["--check"]), 0)


if __name__ == "__main__":
    unittest.main()
