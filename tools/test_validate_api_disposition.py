from __future__ import annotations

import csv
from pathlib import Path
import tempfile
import unittest

from tools import validate_api_disposition as validator


class DispositionValidationTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tmp = tempfile.TemporaryDirectory()
        root = Path(self.tmp.name)
        self.inventory = root / "inventory.tsv"
        self.disposition = root / "disposition.tsv"

    def tearDown(self) -> None:
        self.tmp.cleanup()

    def write_rows(self, inventory_rows, disposition_rows) -> None:
        with self.inventory.open("w", newline="", encoding="utf-8") as stream:
            writer = csv.DictWriter(stream, fieldnames=validator.INVENTORY_FIELDS,
                                    delimiter="\t", lineterminator="\n")
            writer.writeheader()
            writer.writerows(inventory_rows)
        with self.disposition.open("w", newline="", encoding="utf-8") as stream:
            writer = csv.DictWriter(stream, fieldnames=validator.DISPOSITION_FIELDS,
                                    delimiter="\t", lineterminator="\n")
            writer.writeheader()
            writer.writerows(disposition_rows)

    def test_accepts_complete_rows_without_claiming_approval(self) -> None:
        inventory_rows = [
            {"crate": "fixture", "category": "mut_owner", "path": "fixture::Package::set",
             "detail": "&mut self", "target": "Edit"},
            {"crate": "fixture", "category": "mut_draft", "path": "fixture::builder::Edit::set",
             "detail": "&mut self", "target": ""},
            {"crate": "litchi-cfb", "category": "mut_owner", "path": "litchi_cfb::Store::set",
             "detail": "&mut self", "target": "Edit"},
            {"crate": "fixture", "category": "redaction", "path": "fixture::redaction::Edit",
             "detail": "review", "target": ""},
        ]
        disposition_rows = [
            {**inventory_rows[0], "decision": "FIXED", "evidence": "record 0362; source line",
             "owner": "format-owner"},
            {**inventory_rows[1], "decision": "BUILDER_DRAFT", "evidence": "builder module is authoring-only",
             "owner": "format-owner"},
            {**inventory_rows[2], "decision": "SUBSTRATE_TRANSPORT", "evidence": "transport-only substrate",
             "owner": "core-owner"},
            {**inventory_rows[3], "decision": "DEFERRED_REDACTION_22",
             "evidence": "item 22 redaction audit remains deferred", "owner": "coordinator"},
        ]
        self.write_rows(inventory_rows, disposition_rows)
        self.assertEqual(validator.validate(self.inventory, self.disposition), [])

    def test_rejects_scaffold_blanks_and_placeholders(self) -> None:
        row = {"crate": "fixture", "category": "mut_owner", "path": "fixture::Package::set",
               "detail": "&mut self", "target": "Edit"}
        self.write_rows([row], [{**row, "decision": "UNASSIGNED_REQUIRED_REVIEW",
                                "evidence": "", "owner": ""}])
        errors = validator.validate(self.inventory, self.disposition)
        self.assertTrue(any("unsupported decision" in error for error in errors))
        self.assertTrue(any("evidence must be" in error for error in errors))
        self.assertTrue(any("owner must be" in error for error in errors))
        self.assertTrue(any("scaffold decision" in error for error in errors))

    def test_requires_exact_one_to_one_keys_and_rejects_broad_deferral(self) -> None:
        first = {"crate": "fixture", "category": "mut_owner", "path": "fixture::Package::set",
                 "detail": "&mut self", "target": "Edit"}
        second = {"crate": "fixture", "category": "rollback", "path": "fixture::Edit::rollback",
                  "detail": "self", "target": "delete"}
        extra = {"crate": "fixture", "category": "with_limits", "path": "fixture::open_with_limits",
                 "detail": "fn", "target": "OpenOptions"}
        self.write_rows([first, second], [
            {**first, "decision": "deferred", "evidence": "later", "owner": "owner"},
            {**first, "decision": "FIXED", "evidence": "duplicate", "owner": "owner"},
            {**extra, "decision": "FIXED", "evidence": "extra", "owner": "owner"},
        ])
        errors = validator.validate(self.inventory, self.disposition)
        self.assertTrue(any("duplicate row" in error for error in errors))
        self.assertTrue(any("missing inventory row" in error for error in errors))
        self.assertTrue(any("extra row" in error for error in errors))
        self.assertTrue(any("unsupported decision 'deferred'" in error for error in errors))

    def test_redaction_deferral_is_narrow(self) -> None:
        row = {"crate": "fixture", "category": "mut_owner", "path": "fixture::Package::set",
               "detail": "&mut self", "target": "Edit"}
        self.write_rows([row], [{**row, "decision": "DEFERRED_REDACTION_22",
                                "evidence": "item 22 redaction audit", "owner": "owner"}])
        errors = validator.validate(self.inventory, self.disposition)
        self.assertTrue(any("requires a redaction row" in error for error in errors))

    def test_accepts_only_the_reviewed_semantic_rows(self) -> None:
        ast_rows = [
            ("litchi-formula", "mut_owner", "litchi_formula::Formula::set_root"),
            ("litchi-formula", "mut_owner", "litchi_formula::Formula::set_display_style"),
        ]
        request_rows = [
            ("litchi-iwa-protos", "edit_type", path)
            for path in (
                "litchi_iwa_protos::keynote_media_lifecycle_codec::SlideLifecycleEdit",
                "litchi_iwa_protos::keynote_media_lifecycle_codec::BuildLifecycleEdit",
                "litchi_iwa_protos::keynote_media_lifecycle_codec::BuildChunkLifecycleEdit",
                "litchi_iwa_protos::keynote_media_lifecycle_codec::node_cache::SlideNodeBuildCacheEdit",
                "litchi_iwa_protos::numbers_table_cell_storage_codec::TableDataListEntryRefCountEdit",
                "litchi_iwa_protos::numbers_table_cell_storage_codec::HeaderSizeEdit",
                "litchi_iwa_protos::numbers_table_cell_storage_codec::ControlCellSpecTableReferenceEdit",
            )
        ]
        schema_paths = (
            "litchi_iwa_protos::tswp::UndoTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::GenericTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::TextTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::CharIndexTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::ReplaceCharIndexTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::AttributeIndexTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::InsertAttributeTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::InsertNilTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::CharDeltaTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::ParagraphDataTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::ObjectDolcTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::CtDateTransaction",
            "litchi_iwa_protos::tswp::undo_transaction::UnionTransaction",
        )
        schema_rows = [
            ("litchi-iwa-protos", category, path)
            for category in ("edit_type", "transaction_type")
            for path in schema_paths
        ]
        semantic_rows = [
            (row, "AST_VALUE") for row in ast_rows
        ] + [
            (row, "WIRE_REQUEST") for row in request_rows
        ] + [
            (row, "WIRE_SCHEMA") for row in schema_rows
        ]
        inventory_rows = [
            {"crate": crate, "category": category, "path": path,
             "detail": "reviewed semantic row", "target": ""}
            for (crate, category, path), _ in semantic_rows
        ]
        disposition_rows = [
            {**row, "decision": decision,
             "evidence": "record 0362ul2 exact source classification",
             "owner": "coordinator"}
            for row, (_, decision) in zip(inventory_rows, semantic_rows)
        ]
        self.write_rows(inventory_rows, disposition_rows)
        self.assertEqual(validator.validate(self.inventory, self.disposition), [])

    def test_semantic_categories_reject_unknown_or_mismatched_rows(self) -> None:
        cases = [
            (
                "AST_VALUE", "litchi-formula", "mut_owner",
                "litchi_formula::Formula::set_root::nested",
            ),
            (
                "AST_VALUE", "other-crate", "mut_owner",
                "litchi_formula::Formula::set_root",
            ),
            (
                "WIRE_REQUEST", "litchi-iwa-protos", "transaction_type",
                "litchi_iwa_protos::keynote_media_lifecycle_codec::SlideLifecycleEdit",
            ),
            (
                "WIRE_REQUEST", "litchi-iwa-protos", "edit_type",
                "litchi_iwa_protos::keynote_media_lifecycle_codec::LifecycleEdit",
            ),
            (
                "WIRE_SCHEMA", "litchi-iwa-protos", "mut_owner",
                "litchi_iwa_protos::tswp::UndoTransaction",
            ),
            (
                "WIRE_REQUEST", "litchi-iwa-protos", "edit_type",
                "litchi_iwa_protos::numbers_table_cell_storage_codec::CommentStorageTableReferenceEdit",
            ),
            (
                "WIRE_REQUEST", "litchi-iwa-protos", "edit_type",
                "litchi_iwa_protos::numbers_table_cell_storage_codec::ControlCellSpecTableReferenceEditAlias",
            ),
            (
                "WIRE_REQUEST", "litchi-iwa-protos", "mut_owner",
                "litchi_iwa_protos::numbers_table_cell_storage_codec::ControlCellSpecTableReferenceEdit",
            ),
            (
                "WIRE_SCHEMA", "litchi-iwa-protos", "edit_type",
                "litchi_iwa_protos::tswp::UndoTransactionWrapperArchive",
            ),
            (
                "WIRE_SCHEMA", "litchi-iwa-protos", "edit_type",
                "litchi_iwa_protos::keynote_table_physical_sort_codec::HeaderRowEdit",
            ),
        ]
        for decision, crate, category, path in cases:
            with self.subTest(decision=decision, category=category, path=path):
                row = {"crate": crate, "category": category, "path": path,
                       "detail": "reviewed semantic row", "target": ""}
                self.write_rows(
                    [row],
                    [{**row, "decision": decision,
                      "evidence": "record 0362ul2 exact source classification",
                      "owner": "coordinator"}],
                )
                errors = validator.validate(self.inventory, self.disposition)
                self.assertTrue(
                    any(f"{decision} requires an exact reviewed" in error for error in errors),
                    errors,
                )

    def test_accepts_exact_source_role_rows(self) -> None:
        expected_counts = {
            "RAW_READER_METHOD": 38,
            "COMPOSITION_DESCRIPTOR": 4,
            "RAW_ARCHIVE_DESCRIPTOR": 1,
            "RAW_WIRE_DESCRIPTOR": 1,
            "OPC_SOURCE_PLAN_DESCRIPTOR": 4,
            "OPC_XML_SPLICE_DESCRIPTOR": 1,
        }
        self.assertEqual(
            {name: len(validator.SEMANTIC_ROWS[name]) for name in expected_counts},
            expected_counts,
        )
        inventory_rows = []
        disposition_rows = []
        for decision in expected_counts:
            identities = validator.SEMANTIC_ROWS[decision]
            for crate, category, path in sorted(identities):
                row = {
                    "crate": crate,
                    "category": category,
                    "path": path,
                    "detail": "exact source-role review",
                    "target": "",
                }
                inventory_rows.append(row)
                disposition_rows.append({
                    **row,
                    "decision": decision,
                    "evidence": "0362ul7 exact source definition and consumer",
                    "owner": "coordinator",
                })
        self.assertEqual(len(inventory_rows), 49)
        self.write_rows(inventory_rows, disposition_rows)
        self.assertEqual(validator.validate(self.inventory, self.disposition), [])

    def test_rejects_all_removed_raw_reader_aliases(self) -> None:
        inventory_rows = []
        disposition_rows = []
        for crate, category, path in sorted(validator.REMOVED_RAW_READER_ROWS):
            row = {
                "crate": crate,
                "category": category,
                "path": path,
                "detail": "old-source diagnostic alias",
                "target": "",
            }
            inventory_rows.append(row)
            disposition_rows.append({
                **row,
                "decision": "RAW_READER_METHOD",
                "evidence": "0362ul7 old capture only",
                "owner": "coordinator",
            })
        self.write_rows(inventory_rows, disposition_rows)
        errors = validator.validate(self.inventory, self.disposition)
        self.assertEqual(len(validator.REMOVED_RAW_READER_ROWS), 6)
        self.assertEqual(
            sum("rejects removed raw Reader alias" in error for error in errors),
            6,
            errors,
        )

    def test_semantic_rows_reject_changed_identity_and_name_patterns(self) -> None:
        cases = [
            (
                "RAW_READER_METHOD",
                "other-crate",
                "mut_draft",
                "litchi_doc::raw::Reader::document",
            ),
            (
                "COMPOSITION_DESCRIPTOR",
                "litchi-doc",
                "mut_owner",
                "litchi_doc::body_text::SubEdit",
            ),
            (
                "RAW_ARCHIVE_DESCRIPTOR",
                "litchi-iwa-archive",
                "edit_type",
                "litchi_iwa_archive::package::EntryEditExtra",
            ),
            (
                "RAW_WIRE_DESCRIPTOR",
                "litchi-iwa-common",
                "edit_type",
                "litchi_iwa_common::wire::NestedFieldEditAlias",
            ),
            (
                "OPC_SOURCE_PLAN_DESCRIPTOR",
                "litchi-opc",
                "edit_type",
                "litchi_opc::ContentTypeEditUnknown",
            ),
            (
                "OPC_XML_SPLICE_DESCRIPTOR",
                "litchi-opc",
                "mut_draft",
                "litchi_opc::OwnedElementEdit",
            ),
        ]
        for decision, crate, category, path in cases:
            with self.subTest(decision=decision, crate=crate, category=category, path=path):
                row = {
                    "crate": crate,
                    "category": category,
                    "path": path,
                    "detail": "near-match must not infer semantic approval",
                    "target": "",
                }
                self.write_rows(
                    [row],
                    [{
                        **row,
                        "decision": decision,
                        "evidence": "0362ul7 exact source-role review",
                        "owner": "coordinator",
                    }],
                )
                errors = validator.validate(self.inventory, self.disposition)
                self.assertTrue(
                    any(
                        f"{decision} requires an exact reviewed crate/category/path" in error
                        for error in errors
                    ),
                    errors,
                )

    def test_semantic_rows_require_nonblank_proof_and_owner(self) -> None:
        crate, category, path = sorted(validator.SEMANTIC_ROWS["RAW_WIRE_DESCRIPTOR"])[0]
        row = {
            "crate": crate,
            "category": category,
            "path": path,
            "detail": "exact source-role review",
            "target": "",
        }
        self.write_rows(
            [row],
            [{
                **row,
                "decision": "RAW_WIRE_DESCRIPTOR",
                "evidence": "",
                "owner": "",
            }],
        )
        errors = validator.validate(self.inventory, self.disposition)
        self.assertTrue(any("evidence must be nonblank" in error for error in errors))
        self.assertTrue(any("owner must be nonblank" in error for error in errors))

    def test_rejects_short_and_extra_rows_without_traceback(self) -> None:
        self.inventory.write_text(
            "crate\tcategory\tpath\tdetail\ttarget\n"
            "fixture\tmut_owner\tfixture::Package::set\n"
            "fixture\tmut_owner\tfixture::Package::get\tdetail\ttarget\textra\n",
            encoding="utf-8",
        )
        self.disposition.write_text(
            "crate\tcategory\tpath\tdetail\ttarget\tdecision\tevidence\towner\n"
            "fixture\tmut_owner\tfixture::Package::set\t\t\tFIXED\tevidence\towner\n"
            "fixture\tmut_owner\tfixture::Package::get\t\t\tFIXED\tevidence\n",
            encoding="utf-8",
        )
        errors = validator.validate(self.inventory, self.disposition)
        self.assertTrue(any("inventory line 2: expected 5 fields, got 3" in error for error in errors))
        self.assertTrue(any("inventory line 3: expected 5 fields, got 6" in error for error in errors))
        self.assertTrue(any("disposition line 3: expected 8 fields, got 7" in error for error in errors))
        self.assertTrue(any("disposition line 3: owner must be nonblank" in error for error in errors))

    def test_rejects_blank_identity_fields_in_both_inputs(self) -> None:
        self.inventory.write_text(
            "crate\tcategory\tpath\tdetail\ttarget\n"
            "\tmut_owner\tfixture::Package::set\tdetail\ttarget\n"
            "fixture\t\tfixture::Package::get\tdetail\ttarget\n"
            "fixture\tmut_owner\t\tdetail\ttarget\n",
            encoding="utf-8",
        )
        self.disposition.write_text(
            "crate\tcategory\tpath\tdetail\ttarget\tdecision\tevidence\towner\n"
            "\tmut_owner\tfixture::Package::set\t\t\tFIXED\tevidence\towner\n"
            "fixture\t\tfixture::Package::get\t\t\tFIXED\tevidence\towner\n"
            "fixture\tmut_owner\t\t\t\tFIXED\tevidence\towner\n",
            encoding="utf-8",
        )
        errors = validator.validate(self.inventory, self.disposition)
        for label, field in (
            ("inventory", "crate"), ("inventory", "category"),
            ("inventory", "path"), ("disposition", "crate"),
            ("disposition", "category"), ("disposition", "path"),
        ):
            self.assertTrue(
                any(f"{label} line " in error and f"{field} must be nonblank" in error
                    for error in errors),
                (label, field, errors),
            )

    def test_accepts_crlf_and_empty_detail_or_target(self) -> None:
        self.inventory.write_bytes(
            b"crate\tcategory\tpath\tdetail\ttarget\r\n"
            b"fixture\tmut_owner\tfixture::Package::set\t\t\r\n"
        )
        self.disposition.write_bytes(
            b"crate\tcategory\tpath\tdetail\ttarget\tdecision\tevidence\towner\r\n"
            b"fixture\tmut_owner\tfixture::Package::set\t\t\tFIXED\trecord 0362\towner\r\n"
        )
        self.assertEqual(validator.validate(self.inventory, self.disposition), [])

    def test_reports_invalid_utf8_as_a_read_error(self) -> None:
        self.inventory.write_bytes(
            b"crate\tcategory\tpath\tdetail\ttarget\n\xff\n"
        )
        self.disposition.write_text(
            "crate\tcategory\tpath\tdetail\ttarget\tdecision\tevidence\towner\n",
            encoding="utf-8",
        )
        errors = validator.validate(self.inventory, self.disposition)
        self.assertTrue(any(error.startswith("inventory: cannot read:") for error in errors))


if __name__ == "__main__":
    unittest.main()
