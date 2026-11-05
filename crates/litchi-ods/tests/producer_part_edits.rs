//! Record 0277 leaf gap closures: edits of LibreOffice-produced packages.
//!
//! LibreOffice ends every XML declaration with a newline, indents its
//! manifest, restates each typed cell's `office:value-type` as
//! `calcext:value-type`, writes no padding rows below a small used range, and
//! pads flat documents with comments and a trailing newline. Each of these
//! made an ordinary edit fail before; these tests pin the closures and the
//! byte preservation around them.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "integration-test assertions panic on failure by design"
)]

use std::io::Read;

use litchi_ods::model::names::{Definition, Range, Scope};
use litchi_ods::worksheet::{CellChange, Selector};
use litchi_ods::{Cell, CellValue, Package};

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data")
            .join(name),
    )
    .unwrap()
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = zip::ZipArchive::new(std::io::Cursor::new(bytes)).unwrap();
    let mut out = Vec::new();
    archive
        .by_name(name)
        .unwrap()
        .read_to_end(&mut out)
        .unwrap();
    out
}

fn set_cells(package: &Package, changes: Vec<CellChange>) -> Package {
    let mut edit = package.edit().unwrap();
    edit.worksheets(|worksheets| {
        worksheets.set_cells(Selector::Name("sheet1"), changes)?;
        Ok(())
    })
    .unwrap();
    edit.commit().unwrap().into_package()
}

#[test]
fn new_packages_author_one_empty_sheet() {
    let package = Package::new().unwrap();
    assert_eq!(package.workbook().unwrap().sheet_names(), ["Sheet1"]);
}

#[test]
fn cell_edits_restate_calcext_value_types_and_keep_other_rows() {
    let source = fixture("odf/corpus/calc-two-sheets.ods");
    let package = Package::from_bytes(&source).unwrap();
    let edited = set_cells(
        &package,
        vec![CellChange::new(
            1,
            1,
            Cell::new(CellValue::Number(42.0), "42"),
        )],
    );
    let before = String::from_utf8(member(&source, "content.xml")).unwrap();
    let after = String::from_utf8(member(edited.as_bytes(), "content.xml")).unwrap();
    // The producer prolog newline and the untouched first row stay exact.
    let first_row_end = before.find("</table:table-row>").unwrap();
    assert_eq!(after[..first_row_end], before[..first_row_end]);
    assert!(
        after.contains(r#"office:value-type="float" office:value="42" calcext:value-type="float""#)
    );
    // The cells around the edit keep their restatement too.
    assert!(after.contains(r#"office:value="3" calcext:value-type="float""#));
    let workbook = edited.workbook().unwrap();
    assert_eq!(workbook.sheets()[0].cell(1, 1).unwrap().text, "42");
    // Every member but content.xml is the source's.
    for name in [
        "styles.xml",
        "meta.xml",
        "settings.xml",
        "META-INF/manifest.xml",
    ] {
        assert_eq!(
            member(&source, name),
            member(edited.as_bytes(), name),
            "{name}"
        );
    }
}

#[test]
fn a_cell_below_the_last_row_is_spliced_after_it() {
    let source = fixture("odf/corpus/calc-two-sheets.ods");
    let package = Package::from_bytes(&source).unwrap();
    let edited = set_cells(
        &package,
        vec![CellChange::new(
            6,
            4,
            Cell::new(CellValue::Text("below".to_owned()), "below"),
        )],
    );
    let before = String::from_utf8(member(&source, "content.xml")).unwrap();
    let after = String::from_utf8(member(edited.as_bytes(), "content.xml")).unwrap();
    let table_end = before.find("</table:table>").unwrap();
    assert_eq!(after[..table_end], before[..table_end]);
    let workbook = edited.workbook().unwrap();
    let sheet = &workbook.sheets()[0];
    assert_eq!(sheet.cell(6, 4).unwrap().text, "below");
    assert_eq!(sheet.cell(2, 0).unwrap().text, "a text");
}

#[test]
fn metadata_edits_replace_only_meta_xml() {
    let source = fixture("odf/corpus/calc-two-sheets.ods");
    let package = Package::from_bytes(&source).unwrap();
    let mut metadata = package.workbook().unwrap().metadata().clone();
    metadata.title = Some("Budget".to_owned());
    let mut edit = package.edit().unwrap();
    edit.set_metadata(&metadata).unwrap();
    let edited = edit.commit().unwrap().into_package();
    assert_eq!(
        edited.workbook().unwrap().metadata().title.as_deref(),
        Some("Budget")
    );
    for name in ["content.xml", "styles.xml", "settings.xml"] {
        assert_eq!(
            member(&source, name),
            member(edited.as_bytes(), name),
            "{name}"
        );
    }
    // A field meta.xml's patcher does not own is refused before staging.
    let mut unowned = metadata.clone();
    unowned.last_modified_by = Some("someone".to_owned());
    let mut edit = package.edit().unwrap();
    assert!(edit.set_metadata(&unowned).is_err());
    // An equal projection stages nothing.
    let mut edit = package.edit().unwrap();
    edit.set_metadata(package.workbook().unwrap().metadata())
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert_eq!(commit.package().as_bytes(), source.as_slice());
}

#[test]
fn definition_and_annotation_edits_accept_producer_prologs() {
    let source = fixture("odf/corpus/calc-two-sheets.ods");
    let package = Package::from_bytes(&source).unwrap();
    let mut edit = package.edit().unwrap();
    edit.definitions(|definitions| {
        definitions.add(Definition::Range(
            Range::new("Totals", "$sheet1.$A$1:.$C$2", Scope::Global).unwrap(),
        ))
    })
    .unwrap();
    edit.annotations(|annotations| {
        annotations.add(
            "sheet1",
            1,
            1,
            litchi_ods::annotations::Annotation::new("note"),
        )
    })
    .unwrap();
    let edited = edit.commit().unwrap().into_package();
    let workbook = edited.workbook().unwrap();
    assert_eq!(workbook.definitions().len(), 1);
    assert_eq!(workbook.annotations().unwrap().entries().len(), 1);
}

#[test]
fn flat_documents_admit_xml_misc_around_the_root() {
    let flat = fixture("odf/ods/rich-text-hyperlinks.fods");
    let spreadsheet = litchi_ods::flat::Workbook::from_bytes(&flat).unwrap();
    assert_eq!(spreadsheet.sheets().len(), 1);
    assert!(spreadsheet.text().unwrap().contains("AboldmixtailZ"));
    // Text before or after the root is still refused.
    let mut trailing = flat.clone();
    trailing.extend_from_slice(b"tail");
    assert!(litchi_ods::flat::Workbook::from_bytes(&trailing).is_err());
    let mut leading = b"lead".to_vec();
    leading.extend_from_slice(&flat);
    assert!(litchi_ods::flat::Workbook::from_bytes(&leading).is_err());
}

#[cfg(any(unix, windows))]
#[test]
fn source_backed_spreadsheets_materialize_without_being_consumed() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.ods");
    std::fs::write(&path, fixture("odf/corpus/calc-two-sheets.ods")).unwrap();
    let source = litchi_ods::SourceBackedSpreadsheet::open(&path).unwrap();
    let first = source.materialize().unwrap();
    let second = source.materialize().unwrap();
    assert_eq!(first.as_bytes(), second.as_bytes());
    assert_eq!(source.sheet_count().unwrap(), 2);
}
