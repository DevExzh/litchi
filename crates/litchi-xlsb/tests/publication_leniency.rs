#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "publication regressions use immediate fixture diagnostics"
)]

//! Changed-publication regressions for the opt-in cached-number policy.
//!
//! A tolerant read turns an invalid cached `Xnum` into the documented typed
//! value, but it must keep that policy while an unrelated owner republishes
//! the workbook.  These tests also retain the original invalid cell record:
//! publication may rebuild the archive, but it must not silently normalize an
//! opaque cached number.

use litchi_core::sheet::WorkbookTrait as _;
use std::io::Cursor;

use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsb::cell_values::{CellError, Reference, StyleIndex, Value};
use litchi_xlsb::cell_watches::Watch;
use litchi_xlsb::raw::{Records, Writer, kind};
use litchi_xlsb::sparkline::{Color, Colors, Group, Groups, Location, Sparkline, SparklineType};
use litchi_xlsb::writer::{Builder, MutableWorksheet};
use litchi_xlsb::{Leniency, Package, Workbook};

const SHEET1: &str = "/xl/worksheets/sheet1.bin";

fn a1() -> Reference {
    Reference::new(0, 0).expect("A1 reference")
}

fn b1() -> Reference {
    Reference::new(0, 1).expect("B1 reference")
}

fn c1() -> Reference {
    Reference::new(0, 2).expect("C1 reference")
}

/// One worksheet with two ordinary `BrtCellReal` values.
fn numbers_workbook() -> Vec<u8> {
    let mut sheet = MutableWorksheet::new("Numbers");
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

/// Replace the first cell's cached `Xnum` with a NaN while leaving the
/// workbook package and every other record untouched.
fn nan_workbook() -> Vec<u8> {
    let source = numbers_workbook();
    let mut package = OpcPackage::from_reader(Cursor::new(source)).expect("read fixture OPC");
    let uri = PackURI::new(SHEET1).expect("worksheet URI");
    let source_sheet = package.get_part(&uri).expect("worksheet part").blob();
    let mut output = Vec::new();
    let mut replaced = false;
    let mut writer = Writer::new(&mut output);
    for item in Records::new(source_sheet) {
        let record = item.expect("fixture worksheet record");
        let mut payload = record.payload().to_vec();
        if !replaced && record.kind() == kind::CELL_REAL {
            // The first eight bytes are the row/style fields; the final
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
    package
        .to_stream(&mut bytes)
        .expect("serialize NaN fixture");
    bytes.into_inner()
}

fn tolerant_workbook() -> Workbook {
    Workbook::new_with_leniency(Cursor::new(nan_workbook()), Leniency::TolerateCachedNumbers)
        .expect("open tolerant workbook")
}

fn save(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook
        .to_package()
        .write_to(&mut output)
        .expect("save workbook");
    output.into_inner()
}

fn worksheet_blob(workbook: &Workbook) -> Vec<u8> {
    let uri = PackURI::new(SHEET1).expect("worksheet URI");
    workbook
        .opc_package()
        .get_part(&uri)
        .expect("worksheet part")
        .blob()
        .to_vec()
}

fn package_edit(workbook: &Workbook, sheet: usize) -> (Package, litchi_xlsb::cell_values::Edit) {
    let package = workbook.to_package();
    let name = workbook.worksheet_names()[sheet].clone();
    let edit = package
        .edit_cell_values(name)
        .expect("cell-value selector")
        .expect("worksheet selector");
    (package, edit)
}

/// Return the complete raw record containing the NaN cache, including its
/// BIFF12 header.  Comparing this value catches an accidental normalization
/// even when the high-level read still reports `#NUM!`.
fn nan_cell_record(sheet: &[u8]) -> Vec<u8> {
    let mut records = Records::new(sheet);
    while let Some(item) = records.next() {
        let record = item.expect("worksheet record");
        let start = record.offset();
        let end = records.offset();
        if record.kind() == kind::CELL_REAL
            && record.payload().len() >= 16
            && record.payload()[8..16] == f64::NAN.to_le_bytes()
        {
            return sheet[start..end].to_vec();
        }
    }
    panic!("worksheet has no NaN BrtCellReal record");
}

fn assert_tolerant_nan(workbook: &Workbook) {
    assert_tolerant_nan_with_b1(workbook, 2.5);
}

fn assert_tolerant_nan_with_b1(workbook: &Workbook, expected_b1: f64) {
    assert_eq!(workbook.leniency(), Leniency::TolerateCachedNumbers);
    let snapshot = workbook.cell_values(0).expect("tolerant cell snapshot");
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
        &Value::Number(expected_b1)
    );
    assert_eq!(snapshot.tolerance_report().total(), 1);
}

fn assert_strict_cell_read_refuses(bytes: &[u8]) {
    let strict = Workbook::new(Cursor::new(bytes.to_vec())).expect("strict lazy open");
    assert!(
        strict.cell_values(0).is_err(),
        "strict cell-value projection must reject the NaN cache"
    );
    let package = Package::from_bytes(bytes.to_vec()).expect("strict package open");
    assert!(
        package
            .into_workbook()
            .expect("strict package projection")
            .cell_values(0)
            .is_err(),
        "strict package projection must reject the NaN cache"
    );
}

fn sparkline_groups() -> Groups {
    let colors = Colors::uniform(Color::rgb(20, 40, 60, 255, 0));
    Groups::new(vec![
        Group::new(
            SparklineType::Line,
            colors,
            vec![Sparkline::new(
                Location::new(4, 5).expect("sparkline location"),
                None,
            )],
        )
        .expect("sparkline group"),
    ])
    .expect("sparkline groups")
}

#[test]
fn changed_cell_values_keep_leniency_and_nan_record_bytes() {
    let mut workbook = tolerant_workbook();
    let original_record = nan_cell_record(&worksheet_blob(&workbook));

    // Keep the raw OPC escape hatch explicit: this commit is owned by the
    // worksheet codec, so it is applied through the raw helper and reopened
    // under the strict workbook projection.
    let mut raw_edit = workbook.cell_values(0).expect("raw cell snapshot").edit();
    raw_edit.set_number(b1(), 3.5).expect("raw change B1");
    let raw_commit = raw_edit.commit().expect("raw cell commit");
    let mut raw_package = workbook.opc_package().clone();
    litchi_xlsb::cell_values::workbook::apply(
        &mut raw_package,
        &PackURI::new(SHEET1).expect("worksheet URI"),
        &raw_commit,
    )
    .expect("publish an explicitly raw commit to raw OPC");
    let raw_workbook = Workbook::from_opc_package(raw_package).expect("raw output workbook");
    assert_eq!(raw_workbook.leniency(), Leniency::Strict);
    assert!(
        raw_workbook.cell_values(0).is_err(),
        "raw OPC does not retain an opt-in projection policy"
    );
    assert_eq!(
        nan_cell_record(&worksheet_blob(&raw_workbook)),
        original_record
    );

    let (_, mut edit) = package_edit(&workbook, 0);
    edit.set_number(b1(), 3.5).expect("change B1");
    let commit = edit.commit().expect("cell commit");
    workbook = commit
        .into_package()
        .into_workbook()
        .expect("tolerant cell publication");
    assert_tolerant_nan_with_b1(&workbook, 3.5);
    assert_eq!(
        nan_cell_record(&worksheet_blob(&workbook)),
        original_record,
        "unrelated B1 publication preserves the raw NaN cell record"
    );

    let output = save(&workbook);
    let reopened =
        Workbook::new_with_leniency(Cursor::new(output.clone()), Leniency::TolerateCachedNumbers)
            .expect("reopen tolerant cell publication");
    assert_tolerant_nan_with_b1(&reopened, 3.5);
    assert_strict_cell_read_refuses(&output);
}

#[test]
fn cell_watch_publication_retains_tolerant_cell_values_atomically() {
    let mut workbook = tolerant_workbook();
    let original_record = nan_cell_record(&worksheet_blob(&workbook));
    let mut edit = workbook.edit_cell_watches(0).expect("watch edit");
    edit.add(Watch::new(3, 5).expect("watch"))
        .expect("add watch");
    let commit = edit.commit().expect("watch commit");
    let snapshot = workbook
        .apply_cell_watches(0, &commit)
        .expect("tolerant watch publication");
    assert_eq!(snapshot.watches(), &[Watch::new(3, 5).expect("watch")]);
    assert_tolerant_nan(&workbook);
    assert_eq!(
        nan_cell_record(&worksheet_blob(&workbook)),
        original_record,
        "unrelated watch publication preserves the raw NaN cell record"
    );

    let output = save(&workbook);
    let reopened =
        Workbook::new_with_leniency(Cursor::new(output.clone()), Leniency::TolerateCachedNumbers)
            .expect("reopen tolerant watch publication");
    assert_tolerant_nan(&reopened);
    assert_eq!(
        reopened.cell_watches(0).expect("watch snapshot").watches(),
        snapshot.watches()
    );
    assert_strict_cell_read_refuses(&output);
}

#[test]
fn sparkline_publication_retains_tolerant_cell_values_atomically() {
    let mut workbook = tolerant_workbook();
    let original_record = nan_cell_record(&worksheet_blob(&workbook));
    let groups = sparkline_groups();
    let mut edit = workbook.edit_sparklines(0).expect("sparkline edit");
    edit.set(groups.clone());
    let commit = edit.commit().expect("sparkline commit");
    let snapshot = workbook
        .apply_sparklines(0, commit)
        .expect("tolerant sparkline publication");
    assert_eq!(snapshot.groups(), Some(&groups));
    assert_tolerant_nan(&workbook);
    assert_eq!(
        nan_cell_record(&worksheet_blob(&workbook)),
        original_record,
        "unrelated sparkline publication preserves the raw NaN cell record"
    );

    let output = save(&workbook);
    let reopened =
        Workbook::new_with_leniency(Cursor::new(output.clone()), Leniency::TolerateCachedNumbers)
            .expect("reopen tolerant sparkline publication");
    assert_tolerant_nan(&reopened);
    assert_eq!(
        reopened.sparklines(0).expect("sparkline snapshot").groups(),
        Some(&groups)
    );
    assert_strict_cell_read_refuses(&output);
}

#[test]
fn structural_transfer_into_tolerant_package_retains_nan_policy_and_bytes() {
    let target_workbook = tolerant_workbook();
    let original_record = nan_cell_record(&worksheet_blob(&target_workbook));
    let donor = Workbook::new(Cursor::new(numbers_workbook())).expect("donor workbook");

    // The legacy mutable-workbook patch entry point must carry the same
    // retained policy.  Its exact inverse must restore the original package
    // image and the original invalid cell record.
    let mut legacy_target = tolerant_workbook();
    let legacy_before = legacy_target
        .to_package()
        .to_bytes()
        .expect("legacy target bytes");
    let mut legacy_edit = legacy_target.to_package().edit().expect("legacy edit");
    legacy_edit
        .transfer_cell(&donor, 0, b1(), 0, c1())
        .expect("legacy transfer");
    let legacy_commit = legacy_edit.commit().expect("legacy commit");
    let legacy_inverse = legacy_commit.patch().inverse();
    legacy_commit
        .patch()
        .apply(&mut legacy_target)
        .expect("legacy tolerant transfer apply");
    assert_tolerant_nan(&legacy_target);
    assert_eq!(
        nan_cell_record(&worksheet_blob(&legacy_target)),
        original_record
    );
    legacy_inverse
        .apply(&mut legacy_target)
        .expect("legacy inverse apply");
    assert_eq!(
        legacy_target
            .to_package()
            .to_bytes()
            .expect("inverse bytes"),
        legacy_before,
        "legacy inverse restores exact package bytes"
    );

    let target = target_workbook.into_package();
    let target_before = target.to_bytes().expect("target bytes");

    // A failed structural plan must not publish anything to the package.
    let mut occupied = target.edit().expect("occupied edit");
    assert!(
        occupied.transfer_cell(&donor, 0, b1(), 0, b1()).is_err(),
        "occupied target cell must refuse transfer"
    );
    assert_eq!(target.to_bytes().expect("unchanged target"), target_before);

    let mut edit = target.edit().expect("transfer edit");
    edit.transfer_cell(&donor, 0, b1(), 0, c1())
        .expect("transfer donor B1 to target C1");
    let commit = edit.commit().expect("transfer commit");
    let applied = target
        .apply(commit.patch())
        .expect("tolerant transfer apply");
    assert_eq!(applied.leniency(), Leniency::TolerateCachedNumbers);

    let published = applied.into_workbook().expect("tolerant transfer workbook");
    assert_tolerant_nan(&published);
    let published_cells = published.cell_values(0).expect("published cells");
    let transferred = published_cells
        .cell(c1())
        .expect("C1 lookup")
        .expect("transferred C1");
    assert_eq!(transferred.value(), &Value::Number(2.5));
    assert_eq!(
        nan_cell_record(&worksheet_blob(&published)),
        original_record,
        "structural transfer preserves the raw NaN cell record"
    );

    let output = save(&published);
    let reopened =
        Workbook::new_with_leniency(Cursor::new(output.clone()), Leniency::TolerateCachedNumbers)
            .expect("reopen tolerant transfer");
    assert_tolerant_nan(&reopened);
    assert_strict_cell_read_refuses(&output);
}

#[test]
fn no_op_publications_keep_source_bytes_and_tolerant_policy() {
    let mut cells = tolerant_workbook();
    let cells_before = save(&cells);
    let (_, edit) = package_edit(&cells, 0);
    let cell_commit = edit.commit().expect("cell no-op commit");
    assert!(cell_commit.patch().is_noop());
    cells = cell_commit
        .into_package()
        .into_workbook()
        .expect("cell no-op publication");
    assert_eq!(save(&cells), cells_before);
    assert_tolerant_nan(&cells);

    let mut watches = tolerant_workbook();
    let watches_before = save(&watches);
    let watch_commit = watches
        .edit_cell_watches(0)
        .expect("watch edit")
        .commit()
        .expect("watch no-op commit");
    watches
        .apply_cell_watches(0, &watch_commit)
        .expect("watch no-op publication");
    assert_eq!(save(&watches), watches_before);
    assert_tolerant_nan(&watches);

    let mut sparklines = tolerant_workbook();
    let sparklines_before = save(&sparklines);
    let sparkline_commit = sparklines
        .edit_sparklines(0)
        .expect("sparkline edit")
        .commit()
        .expect("sparkline no-op commit");
    assert!(sparkline_commit.is_noop());
    sparklines
        .apply_sparklines(0, sparkline_commit)
        .expect("sparkline no-op publication");
    assert_eq!(save(&sparklines), sparklines_before);
    assert_tolerant_nan(&sparklines);

    let package = tolerant_workbook().into_package();
    let package_before = package.to_bytes().expect("package bytes");
    let commit = package
        .edit()
        .expect("structure edit")
        .commit()
        .expect("structure no-op commit");
    assert!(commit.patch().is_noop());
    let applied = package
        .apply(commit.patch())
        .expect("structure no-op apply");
    assert_eq!(applied.leniency(), Leniency::TolerateCachedNumbers);
    assert_eq!(applied.to_bytes().expect("applied bytes"), package_before);
    assert_tolerant_nan(&applied.into_workbook().expect("no-op package workbook"));
}

#[test]
fn failed_tolerant_cell_publication_leaves_source_and_nan_record_untouched() {
    let workbook = tolerant_workbook();
    let before = save(&workbook);
    let original_record = nan_cell_record(&worksheet_blob(&workbook));
    let (_, mut edit) = package_edit(&workbook, 0);
    edit.insert(
        Reference::new(10_000, 100).expect("new cell"),
        StyleIndex::new(0x00ff_ffff).expect("wire style"),
        Value::Number(9.0),
    )
    .expect("stage invalid dependency");
    assert!(
        edit.commit().is_err(),
        "missing style dependency must refuse before publication"
    );
    assert_eq!(save(&workbook), before);
    assert_eq!(nan_cell_record(&worksheet_blob(&workbook)), original_record);
    assert_tolerant_nan(&workbook);
}

#[test]
fn mutable_structural_application_retains_destination_input_budget() {
    let source = Package::new().expect("minimal workbook");
    let bytes = source.to_bytes().expect("source bytes");
    let mut edit = source.edit().expect("unbounded-by-input edit");
    edit.rename_sheet(0, "A significantly longer name".to_owned())
        .unwrap();
    let commit = edit.commit().expect("default-budget commit");
    assert!(commit.patch().after().len() > bytes.len());
    let limits = litchi_opc::SourceLimits::builder()
        .max_input_bytes(u64::try_from(bytes.len()).unwrap())
        .unwrap()
        .build()
        .unwrap();
    let mut destination = Package::from_bytes_with_limits(bytes.clone(), limits)
        .unwrap()
        .into_workbook()
        .unwrap();
    assert!(commit.patch().apply(&mut destination).is_err());
    assert_eq!(destination.to_package().to_bytes().unwrap(), bytes);
    assert_eq!(destination.opc_package().read_limits(), limits);
    let noop = source.edit().unwrap().commit().unwrap();
    noop.patch()
        .apply(&mut destination)
        .expect("bounded exact no-op");
    assert_eq!(destination.to_package().to_bytes().unwrap(), bytes);
    assert_eq!(destination.opc_package().read_limits(), limits);
}
