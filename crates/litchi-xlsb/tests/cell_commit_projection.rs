#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "projection regressions use compact generated fixtures and immediate diagnostics"
)]

//! Regressions for the package projection retained by a cell-values commit.
//!
//! A package-owned cell commit has already validated the candidate workbook.
//! Turning that package back into an eager [`Workbook`] must therefore expose
//! the exact committed image and retain the source policies that governed the
//! validation.

use std::io::Cursor;

use litchi_core::sheet::WorkbookTrait as _;
use litchi_opc::{OpcPackage, PackURI, SourceLimits};
use litchi_xlsb::builder::{self, Builder};
use litchi_xlsb::cell_values::{CellError, Reference, StyleIndex, Value};
use litchi_xlsb::raw::{Encoder, Records, kind};
use litchi_xlsb::{Leniency, Package, Workbook};

#[path = "support/projection.rs"]
mod projection;
use projection::OwnedWorkbook as _;

const SHEET1: &str = "/xl/worksheets/sheet1.bin";

fn a1() -> Reference {
    Reference::new(0, 0).expect("A1 reference")
}

fn b1() -> Reference {
    Reference::new(0, 1).expect("B1 reference")
}

fn style_zero() -> StyleIndex {
    StyleIndex::new(0).expect("default style")
}

fn numbers_workbook() -> Vec<u8> {
    let mut sheet = builder::Worksheet::new("Numbers");
    sheet.set_cell(0, 0, 1.5);
    sheet.set_cell(0, 1, 2.5);
    let mut writer = Builder::new();
    writer.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    writer
        .build()
        .expect("build package")
        .write_to(&mut output)
        .expect("write numbers fixture");
    output.into_inner()
}

/// Return the numbers fixture with the first cached Xnum replaced by NaN.
fn nan_workbook() -> Vec<u8> {
    let source = numbers_workbook();
    let mut package = OpcPackage::from_reader(Cursor::new(source), SourceLimits::default())
        .expect("read fixture OPC");
    let uri = PackURI::new(SHEET1).expect("worksheet URI");
    let source_sheet = package.get_part(&uri).expect("worksheet part").blob();
    let mut output = Vec::new();
    let mut writer = Encoder::new(&mut output, litchi_xlsb::raw::Limits::DEFAULT);
    let mut replaced = false;
    for item in Records::new(source_sheet, litchi_xlsb::raw::Limits::DEFAULT) {
        let record = item.expect("fixture worksheet record");
        let mut payload = record.payload().to_vec();
        if !replaced && record.kind() == kind::CELL_REAL {
            // The first eight bytes are the column/style fields; the final
            // eight bytes are the cached IEEE-754 Xnum.
            payload[8..16].copy_from_slice(&f64::NAN.to_le_bytes());
            replaced = true;
        }
        writer
            .write_record(record.kind(), &payload)
            .expect("rewrite worksheet record");
    }
    assert!(replaced, "numbers fixture contains a BrtCellReal record");
    package
        .get_part_mut(&uri)
        .expect("worksheet part")
        .set_blob(output);
    let mut bytes = Cursor::new(Vec::new());
    package.write_to(&mut bytes).expect("serialize NaN fixture");
    bytes.into_inner()
}

fn changed_commit(source: &Package) -> litchi_xlsb::cell_values::Commit {
    let mut edit = source
        .edit_cell_values(0)
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet");
    edit.insert(a1(), style_zero(), Value::Number(42.0))
        .expect("insert numeric cell");
    edit.commit().expect("package cell-value commit")
}

fn cell_value(workbook: &Workbook, reference: Reference) -> Value {
    workbook
        .cell_values(0)
        .expect("cell snapshot")
        .cell(reference)
        .expect("cell lookup")
        .expect("stored cell")
        .value()
        .clone()
}

#[test]
fn changed_cell_commit_projects_to_the_exact_patch_image() {
    let source = Package::new().expect("source package");
    let commit = changed_commit(&source);
    let expected = commit.patch().after().to_vec();

    let package = commit.into_package();
    assert_eq!(package.to_bytes().expect("candidate bytes"), expected);

    let workbook = package.owned_workbook().expect("candidate workbook");
    assert_eq!(
        workbook
            .to_package()
            .to_bytes()
            .expect("projected candidate bytes"),
        expected,
        "projecting a validated commit must retain its exact package image"
    );
    assert_eq!(cell_value(&workbook, a1()), Value::Number(42.0));
}

#[test]
fn cloned_commit_and_follow_up_edit_do_not_use_a_stale_projection() {
    let source = Package::new().expect("source package");
    let commit = changed_commit(&source);
    let cloned_commit = commit.clone();
    let expected = commit.patch().after().to_vec();
    let cached = commit.into_package();

    let first = cached
        .clone()
        .owned_workbook()
        .expect("first commit projection");
    let second = cloned_commit
        .into_package()
        .clone()
        .owned_workbook()
        .expect("cloned commit projection");
    assert_eq!(
        first.to_package().to_bytes().expect("first bytes"),
        expected
    );
    assert_eq!(
        second.to_package().to_bytes().expect("second bytes"),
        expected
    );
    assert_eq!(cell_value(&first, a1()), Value::Number(42.0));
    assert_eq!(cell_value(&second, a1()), Value::Number(42.0));

    // The original package still owns the validated projection after both
    // clones were consumed.  A structural edit must invalidate that cached
    // view and publish the renamed workbook rather than replaying the stale
    // Sheet1 projection.
    let mut follow_up = cached.edit().expect("follow-up structural edit");
    follow_up
        .set_sheet_name(0, "Renamed".to_owned())
        .expect("follow-up rename")
        .expect("resolved worksheet selector");
    let follow_up = follow_up.commit().expect("follow-up commit");
    let follow_up_expected = follow_up.patch().after().to_vec();
    let follow_up_workbook = follow_up
        .into_package()
        .owned_workbook()
        .expect("follow-up projection");
    assert_eq!(
        follow_up_workbook
            .to_package()
            .to_bytes()
            .expect("follow-up bytes"),
        follow_up_expected
    );
    assert_eq!(
        follow_up_workbook.worksheet_names(),
        &["Renamed".to_owned()]
    );
}

#[test]
fn no_op_cell_commit_projection_preserves_the_exact_source() {
    let source = Package::new().expect("source package");
    let original = source.to_bytes().expect("source bytes");
    let commit = source
        .edit_cell_values("Sheet1")
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Sheet1 worksheet")
        .commit()
        .expect("no-op commit");

    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().is_noop());
    assert_eq!(commit.patch().before(), original.as_slice());
    assert_eq!(commit.patch().after(), original.as_slice());

    let workbook = commit
        .into_package()
        .owned_workbook()
        .expect("no-op projection");
    assert_eq!(
        workbook.to_package().to_bytes().expect("no-op bytes"),
        original,
        "an exact no-op must not rebuild the source package"
    );
    assert!(
        workbook
            .cell_values(0)
            .expect("cell snapshot")
            .cell(a1())
            .expect("A1 lookup")
            .is_none()
    );
}

#[test]
fn changed_cell_commit_retains_tolerant_cached_number_policy() {
    let workbook = Workbook::new_with_options(
        Cursor::new(nan_workbook()),
        litchi_xlsb::OpenOptions::new().leniency(Leniency::TolerateCachedNumbers),
    )
    .expect("open tolerant workbook");
    let package = workbook.into_package();
    let mut edit = package
        .edit_cell_values(0)
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Numbers worksheet");
    edit.set_number(b1(), 3.5).expect("change B1");
    let commit = edit.commit().expect("tolerant cell commit");

    let published = commit
        .into_package()
        .owned_workbook()
        .expect("tolerant commit projection");
    assert_eq!(published.leniency(), Leniency::TolerateCachedNumbers);
    let snapshot = published.cell_values(0).expect("tolerant cell snapshot");
    assert_eq!(
        snapshot
            .cell(a1())
            .expect("A1 lookup")
            .expect("A1 cell")
            .value(),
        &Value::Error(CellError::Number)
    );
    assert_eq!(
        snapshot
            .cell(b1())
            .expect("B1 lookup")
            .expect("B1 cell")
            .value(),
        &Value::Number(3.5)
    );
    assert_eq!(snapshot.tolerance_report().total(), 1);
}

#[test]
fn changed_cell_commit_retains_read_limits_and_skipped_drawing_policy() {
    let source = numbers_workbook();
    let input = u64::try_from(source.len().saturating_mul(2)).expect("source length fits u64");
    let limits = SourceLimits::builder()
        .max_input_bytes(input)
        .expect("input limit")
        .build()
        .expect("source limits");
    let workbook = Workbook::new_with_options(
        Cursor::new(source),
        litchi_xlsb::OpenOptions::new()
            .source_limits(limits)
            .parse_drawings(false),
    )
    .expect("open drawing-skipped workbook");
    assert!(!workbook.drawings_parsed());

    let package = workbook.into_package();
    assert!(!package.drawings_parsed());
    assert_eq!(package.opc_package().read_limits(), limits);
    let mut edit = package
        .edit_cell_values(0)
        .expect("worksheet selector")
        .resolve()
        .expect("worksheet resolution")
        .expect("Numbers worksheet");
    edit.set_number(b1(), 7.5).expect("change B1");
    let commit = edit.commit().expect("bounded cell commit");
    let published = commit
        .into_package()
        .owned_workbook()
        .expect("bounded commit projection");

    assert!(!published.drawings_parsed());
    assert_eq!(published.opc_package().read_limits(), limits);
    assert_eq!(cell_value(&published, b1()), Value::Number(7.5));
}
