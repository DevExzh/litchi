#!/usr/bin/env python3
"""Validate a completed API-debt disposition without approving it.

``final_api_inventory.py`` intentionally emits a blank, human-review scaffold.
This tool is a later audit step: it checks that inventory and disposition rows
are an exact one-to-one relation, that every row has nonblank human evidence,
and that the decision is one of the contract's concrete categories. It never
changes the files and does not infer that a row is semantically resolved.

Accepted decisions are:

* ``FIXED`` — a source repair is evidenced;
* ``BUILDER_DRAFT`` — a proven authoring builder is intentionally mutable;
* ``SUBSTRATE_TRANSPORT`` — only a transport-only CFB/ZIP substrate row;
* ``AST_VALUE`` — only the two reviewed Formula AST value methods;
* ``WIRE_REQUEST`` — only the seven reviewed IWA codec request values;
* ``WIRE_SCHEMA`` — only the reviewed generated IWA protobuf schema values;
* ``RAW_READER_METHOD`` — only the 38 current raw DOC/PPT/XLS Reader methods;
* ``COMPOSITION_DESCRIPTOR`` — only the four exact SubEdit composition values;
* ``RAW_ARCHIVE_DESCRIPTOR`` — only the IWA EntryEdit descriptor;
* ``RAW_WIRE_DESCRIPTOR`` — only the IWA NestedFieldEdit descriptor;
* ``OPC_SOURCE_PLAN_DESCRIPTOR`` — only the four exact OPC source-plan values;
* ``OPC_XML_SPLICE_DESCRIPTOR`` — only the exact OPC XML splice value;
* ``DEFERRED_REDACTION_22`` — only an explicitly identified item-22
  redaction row.

The semantic categories classify inventory rows; they are not waivers and do
not imply that a source capture or final disposition has been approved. The
last category is the sole permitted deferral. Placeholder decisions from the
scaffold and broad ``deferred``/``out-of-scope``/``wontfix`` values are
rejected.
"""
from __future__ import annotations

import argparse
import csv
from pathlib import Path


INVENTORY_FIELDS = ("crate", "category", "path", "detail", "target")
DISPOSITION_FIELDS = INVENTORY_FIELDS + ("decision", "evidence", "owner")
ROW_KEY = INVENTORY_FIELDS
SUBSTRATE_CRATES = frozenset({"litchi-cfb", "soapberry-zip"})
SCAFFOLD_DECISIONS = frozenset({
    "DRAFT_CANDIDATE_REVIEW",
    "SUBSTRATE_SCOPE_REVIEW",
    "UNASSIGNED_REQUIRED_REVIEW",
})

# These are semantic classifications for inventory rows that are not format
# package edits.  Keep the allowlist at full inventory-row granularity: a
# module prefix, crate-wide exception, or old alias must never acquire a
# classification by accident.  The generated protobuf messages occur in the
# inventory as both the edit_type and transaction_type rows, so both categories
# are listed for each exact public path.
SEMANTIC_ROWS = {
    "AST_VALUE": frozenset({
        ("litchi-formula", "mut_owner", "litchi_formula::Formula::set_root"),
        ("litchi-formula", "mut_owner", "litchi_formula::Formula::set_display_style"),
    }),
    "WIRE_REQUEST": frozenset({
        ("litchi-iwa-protos", "edit_type",
         "litchi_iwa_protos::keynote_media_lifecycle_codec::SlideLifecycleEdit"),
        ("litchi-iwa-protos", "edit_type",
         "litchi_iwa_protos::keynote_media_lifecycle_codec::BuildLifecycleEdit"),
        ("litchi-iwa-protos", "edit_type",
         "litchi_iwa_protos::keynote_media_lifecycle_codec::BuildChunkLifecycleEdit"),
        ("litchi-iwa-protos", "edit_type",
         "litchi_iwa_protos::keynote_media_lifecycle_codec::node_cache::SlideNodeBuildCacheEdit"),
        ("litchi-iwa-protos", "edit_type",
         "litchi_iwa_protos::numbers_table_cell_storage_codec::TableDataListEntryRefCountEdit"),
        ("litchi-iwa-protos", "edit_type",
         "litchi_iwa_protos::numbers_table_cell_storage_codec::HeaderSizeEdit"),
        ("litchi-iwa-protos", "edit_type",
         "litchi_iwa_protos::numbers_table_cell_storage_codec::ControlCellSpecTableReferenceEdit"),
    }),
    "WIRE_SCHEMA": frozenset({
        ("litchi-iwa-protos", category, path)
        for category in ("edit_type", "transaction_type")
        for path in (
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
    }),
    "RAW_READER_METHOD": frozenset({
        (crate, "mut_draft", path)
        for crate, paths in {
            "litchi-doc": (
                "litchi_doc::raw::Reader::custom_xml_data_store",
                "litchi_doc::raw::Reader::data_spaces",
                "litchi_doc::raw::Reader::data_spaces_snapshot",
                "litchi_doc::raw::Reader::document",
                "litchi_doc::raw::Reader::document_summary_information",
                "litchi_doc::raw::Reader::document_with_options",
                "litchi_doc::raw::Reader::ole_file",
                "litchi_doc::raw::Reader::save",
                "litchi_doc::raw::Reader::save_with_durability",
                "litchi_doc::raw::Reader::signatures",
                "litchi_doc::raw::Reader::summary_information",
                "litchi_doc::raw::Reader::to_bytes",
                "litchi_doc::raw::Reader::user_defined_hyperlinks",
                "litchi_doc::raw::Reader::user_defined_properties",
                "litchi_doc::raw::Reader::vba",
                "litchi_doc::raw::Reader::vba_signature",
                "litchi_doc::raw::Reader::write_data_spaces_patch",
                "litchi_doc::raw::Reader::write_to",
            ),
            "litchi-ppt": (
                "litchi_ppt::raw::Reader::custom_xml_data_store",
                "litchi_ppt::raw::Reader::document_comparison",
                "litchi_ppt::raw::Reader::document_summary_information",
                "litchi_ppt::raw::Reader::fonts",
                "litchi_ppt::raw::Reader::ole_file",
                "litchi_ppt::raw::Reader::picture_bullets",
                "litchi_ppt::raw::Reader::presentation",
                "litchi_ppt::raw::Reader::presentation_with_options",
                "litchi_ppt::raw::Reader::signatures",
                "litchi_ppt::raw::Reader::summary_information",
                "litchi_ppt::raw::Reader::user_defined_properties",
            ),
            "litchi-xls": (
                "litchi_xls::raw::Reader::custom_xml_data_store",
                "litchi_xls::raw::Reader::document_summary_information",
                "litchi_xls::raw::Reader::revision_log",
                "litchi_xls::raw::Reader::signatures",
                "litchi_xls::raw::Reader::summary_information",
                "litchi_xls::raw::Reader::toolbar",
                "litchi_xls::raw::Reader::user_defined_properties",
                "litchi_xls::raw::Reader::user_names",
                "litchi_xls::raw::Reader::vba",
            ),
        }.items()
        for path in paths
    }),
    "COMPOSITION_DESCRIPTOR": frozenset({
        ("litchi-core", "edit_type", "litchi_core::SubEdit"),
        ("litchi-doc", "edit_type", "litchi_doc::body_text::SubEdit"),
        ("litchi-odb", "edit_type", "litchi_odb::SubEdit"),
        ("litchi-odg", "edit_type", "litchi_odg::edit::SubEdit"),
    }),
    "RAW_ARCHIVE_DESCRIPTOR": frozenset({
        ("litchi-iwa-archive", "edit_type", "litchi_iwa_archive::package::EntryEdit"),
    }),
    "RAW_WIRE_DESCRIPTOR": frozenset({
        ("litchi-iwa-common", "edit_type", "litchi_iwa_common::wire::NestedFieldEdit"),
    }),
    "OPC_SOURCE_PLAN_DESCRIPTOR": frozenset({
        ("litchi-opc", "edit_type", "litchi_opc::ContentTypeEdit"),
        ("litchi-opc", "edit_type", "litchi_opc::ContentTypeFragmentEdit"),
        ("litchi-opc", "edit_type", "litchi_opc::RelationshipEdit"),
        ("litchi-opc", "edit_type", "litchi_opc::RelationshipFragmentEdit"),
    }),
    "OPC_XML_SPLICE_DESCRIPTOR": frozenset({
        ("litchi-opc", "edit_type", "litchi_opc::OwnedElementEdit"),
    }),
}
SEMANTIC_DECISIONS = frozenset(SEMANTIC_ROWS)
ACCEPTED_DECISIONS = frozenset({
    "FIXED",
    "BUILDER_DRAFT",
    "SUBSTRATE_TRANSPORT",
    "DEFERRED_REDACTION_22",
}) | SEMANTIC_DECISIONS


# These six names were present in the retained old-source capture but have no
# exact definition at the reviewed source commit. They must never become raw
# semantic rows through a prefix/suffix or compatibility-name match.
REMOVED_RAW_READER_ROWS = frozenset({
    ("litchi-doc", "mut_draft", "litchi_doc::raw::Reader::signatures_with"),
    ("litchi-doc", "mut_draft", "litchi_doc::raw::Reader::vba_signature_with"),
    ("litchi-doc", "mut_draft", "litchi_doc::raw::Reader::vba_with"),
    ("litchi-ppt", "mut_draft", "litchi_ppt::raw::Reader::signatures_with"),
    ("litchi-xls", "mut_draft", "litchi_xls::raw::Reader::signatures_with"),
    ("litchi-xls", "mut_draft", "litchi_xls::raw::Reader::vba_with"),
})


def _read_rows(path: Path, fields: tuple[str, ...], label: str) -> tuple[list[dict[str, str]], list[str]]:
    errors: list[str] = []
    try:
        with path.open(newline="", encoding="utf-8") as stream:
            reader = csv.reader(stream, delimiter="\t", strict=True)
            try:
                header = next(reader)
            except StopIteration:
                return [], [f"{label}: missing header"]
            if header != list(fields):
                errors.append(
                    f"{label}: expected header {' '.join(fields)}, "
                    f"got {header!r}"
                )
            rows: list[dict[str, str]] = []
            for line, values in enumerate(reader, 2):
                if len(values) != len(fields):
                    errors.append(
                        f"{label} line {line}: expected {len(fields)} fields, "
                        f"got {len(values)}"
                    )
                # Keep malformed rows available for one-to-one diagnostics,
                # while padding short rows so later checks cannot traceback.
                normalized = (values + [""] * len(fields))[:len(fields)]
                row = dict(zip(fields, normalized))
                rows.append(row)
                for field in ("crate", "category", "path"):
                    if not row[field].strip():
                        errors.append(
                            f"{label} line {line}: {field} must be nonblank"
                        )
    except (OSError, UnicodeError, csv.Error) as error:
        return [], [f"{label}: cannot read: {error}"]
    return rows, errors


def _key(row: dict[str, str]) -> tuple[str, ...]:
    return tuple(row.get(field) or "" for field in ROW_KEY)


def _duplicate_errors(rows: list[dict[str, str]], label: str) -> list[str]:
    seen: dict[tuple[str, ...], int] = {}
    errors: list[str] = []
    for number, row in enumerate(rows, 2):
        key = _key(row)
        if key in seen:
            errors.append(
                f"{label}: duplicate row {number} (already at line {seen[key]}): "
                + "\t".join(key)
            )
        else:
            seen[key] = number
    return errors


def _decision_errors(row: dict[str, str], line: int) -> list[str]:
    errors: list[str] = []
    decision = row.get("decision", "").strip()
    evidence = row.get("evidence", "").strip()
    owner = row.get("owner", "").strip()
    label = f"disposition line {line}"
    for field, value in (("decision", decision), ("evidence", evidence), ("owner", owner)):
        if not value:
            errors.append(f"{label}: {field} must be nonblank")
    if decision and decision not in ACCEPTED_DECISIONS:
        errors.append(f"{label}: unsupported decision {decision!r}")
        if decision in SCAFFOLD_DECISIONS:
            errors.append(f"{label}: scaffold decision is still pending human review")
        return errors
    if decision == "BUILDER_DRAFT":
        if row.get("category") != "mut_draft" or "builder" not in row.get("path", "").split("::"):
            errors.append(f"{label}: BUILDER_DRAFT requires a mut_draft builder path")
    elif decision == "SUBSTRATE_TRANSPORT":
        if row.get("crate") not in SUBSTRATE_CRATES:
            errors.append(f"{label}: SUBSTRATE_TRANSPORT is limited to CFB/ZIP substrate crates")
    elif decision == "DEFERRED_REDACTION_22":
        lowered = evidence.lower()
        if row.get("category") != "redaction" or "redaction" not in lowered or "22" not in lowered:
            errors.append(
                f"{label}: DEFERRED_REDACTION_22 requires a redaction row whose evidence names item 22"
            )
    elif decision in SEMANTIC_DECISIONS:
        row_identity = (
            row.get("crate", ""), row.get("category", ""), row.get("path", "")
        )
        if row_identity in REMOVED_RAW_READER_ROWS:
            errors.append(
                f"{label}: {decision} rejects removed raw Reader alias {row_identity!r}"
            )
        elif row_identity not in SEMANTIC_ROWS[decision]:
            errors.append(
                f"{label}: {decision} requires an exact reviewed crate/category/path; "
                f"got {row_identity!r}"
            )
    return errors


def validate(inventory: Path, disposition: Path) -> list[str]:
    """Return audit failures; an empty list means only structural checks passed."""
    inventory_rows, errors = _read_rows(inventory, INVENTORY_FIELDS, "inventory")
    disposition_rows, disposition_errors = _read_rows(
        disposition, DISPOSITION_FIELDS, "disposition")
    errors.extend(disposition_errors)
    errors.extend(_duplicate_errors(inventory_rows, "inventory"))
    errors.extend(_duplicate_errors(disposition_rows, "disposition"))

    inventory_keys = {_key(row) for row in inventory_rows}
    disposition_keys = {_key(row) for row in disposition_rows}
    for key in sorted(inventory_keys - disposition_keys):
        errors.append("disposition: missing inventory row: " + "\t".join(key))
    for key in sorted(disposition_keys - inventory_keys):
        errors.append("disposition: extra row not present in inventory: " + "\t".join(key))
    for line, row in enumerate(disposition_rows, 2):
        errors.extend(_decision_errors(row, line))
    return errors


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("inventory", type=Path)
    parser.add_argument("disposition", type=Path)
    args = parser.parse_args(argv)
    errors = validate(args.inventory, args.disposition)
    if errors:
        print("\n".join(errors))
        return 1
    print(
        f"validated {sum(1 for _ in args.inventory.open(encoding='utf-8')) - 1} "
        "disposition rows; no approval inferred"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
