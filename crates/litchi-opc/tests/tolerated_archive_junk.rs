//! Real-world packages whose archives carry items that are not OPC parts.
//!
//! Each case checks two things: the package opens, and the content that is
//! actually inside it is reachable afterwards. Tolerating junk is only correct
//! if nothing real is lost, and only acceptable if the junk is still reported.

use litchi_opc::{NonPartReason, OpcPackage, Part};
use std::path::PathBuf;

/// Content type of the WordprocessingML main document part.
const WML_DOCUMENT_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
/// Content type of the SpreadsheetML workbook part.
const SML_SHEET_MAIN: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

/// Resolve a fixture path under the repository's `test-data` tree.
fn fixture(relative: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data")
        .join(relative)
}

/// Open a fixture package, failing the test with the fixture name on error.
fn open(relative: &str) -> OpcPackage {
    let path = fixture(relative);
    let bytes = std::fs::read(&path).unwrap_or_else(|error| panic!("read {relative}: {error}"));
    OpcPackage::from_bytes(&bytes).unwrap_or_else(|error| panic!("open {relative}: {error}"))
}

/// Return the part with the given name, or fail the test.
fn part<'package>(package: &'package OpcPackage, partname: &str) -> &'package dyn Part {
    package
        .iter_parts()
        .find(|part| part.partname().as_str() == partname)
        .unwrap_or_else(|| panic!("package is missing {partname}"))
}

#[test]
fn excel_trash_folder_does_not_block_the_workbook() {
    // Excel leaves `[trash]/0000.dat` behind; `[` and `]` are outside RFC 3986
    // pchar, so the item cannot denote a part name (ECMA-376 Part 2 §9.1.1.1).
    let package = open("poi/test-data/spreadsheet/Excel_file_with_trash_item.xlsx");

    let workbook = part(&package, "/xl/workbook.xml");
    assert_eq!(workbook.content_type(), SML_SHEET_MAIN);
    assert!(workbook.blob().starts_with(b"<?xml"));
    assert!(
        package
            .iter_parts()
            .any(|part| part.partname().as_str() == "/xl/worksheets/sheet1.xml")
    );

    let junk = package.non_part_members();
    assert!(
        junk.iter()
            .any(|member| member.name().starts_with("[trash]")
                && member.reason() == NonPartReason::UnmappablePartName),
        "trash items must be reported, got {junk:?}"
    );
}

#[test]
fn ds_store_does_not_block_the_document() {
    // `.DS_Store` has a usable part name but no content type and no incoming
    // relationship, so it is archive junk rather than an untyped part.
    let package = open("libreoffice-core/sw/qa/extras/ooxmlexport/data/tdf124384.docx");

    let document = part(&package, "/word/document.xml");
    assert_eq!(document.content_type(), WML_DOCUMENT_MAIN);
    assert!(document.blob().windows(6).any(|window| window == b"<w:doc"));
    assert!(
        package
            .iter_parts()
            .any(|part| part.partname().as_str() == "/word/styles.xml")
    );

    let junk = package.non_part_members();
    assert!(
        junk.iter().any(|member| member.name() == ".DS_Store"
            && member.reason() == NonPartReason::UntypedAndUnreferenced),
        "editor junk must be reported, got {junk:?}"
    );
}

#[test]
fn dangling_relationship_target_leaves_the_rest_of_the_document_intact() {
    // `word/_rels/document.xml.rels` points at `fontTable.xml`, which the
    // archive does not contain. OPC states no rule that an internal target must
    // resolve, and the relationship is metadata on its source part.
    let package = open("libreoffice-core/sw/qa/extras/ooxmlexport/data/listWithLgl.docx");

    let document = part(&package, "/word/document.xml");
    assert_eq!(document.content_type(), WML_DOCUMENT_MAIN);
    assert!(!part(&package, "/word/numbering.xml").blob().is_empty());
    assert!(
        !package
            .iter_parts()
            .any(|part| part.partname().as_str() == "/word/fontTable.xml"),
        "an absent target must not be fabricated as a part"
    );

    let dangling = document
        .rels()
        .iter()
        .find(|relationship| relationship.target_ref() == "fontTable.xml");
    assert!(
        dangling.is_some(),
        "the unresolved relationship must stay visible on its source part"
    );
}

#[test]
fn content_types_stream_stored_under_a_case_variant_name_still_opens() {
    // ECMA-376 Part 2 §9.1.1.2 compares item names ASCII case-insensitively;
    // this workbook stores the manifest as `[content_types].xml`.
    let package = open("poi/test-data/spreadsheet/49609.xlsx");

    let workbook = part(&package, "/xl/workbook.xml");
    assert_eq!(workbook.content_type(), SML_SHEET_MAIN);
    assert!(
        !part(&package, "/xl/worksheets/sheet1.xml")
            .blob()
            .is_empty()
    );
    // The writer that produced this file also lower-cased some item names, so
    // parts keep the spelling the archive used. Nothing is dropped.
    assert!(!part(&package, "/xl/sharedstrings.xml").blob().is_empty());
    assert!(package.non_part_members().is_empty());
}
