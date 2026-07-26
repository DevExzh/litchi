//! Grid padding written after the used range of an ODF spreadsheet.
//!
//! Every ODF producer pads a sheet out to its full addressable size with
//! attribute-free `table:table-cell` fillers repeated across the width and a
//! `table:table-row` repeated across the remaining height. Expanding that
//! padding literally costs millions of allocations and used to push ordinary
//! spreadsheets past the parser's expansion safety limits, so the reader now
//! discards a trailing run while still placing interior gaps correctly.

use litchi_odf::{FlatSpreadsheet, MutableSpreadsheet, Spreadsheet};

/// Flat spreadsheet exercising trailing padding, interior gaps, and a short
/// authored tail of empty rows.
const FIXTURE: &str = "../../test-data/odf/ods/grid-padding.fods";

/// Producer-written package whose three sheets each end in a row repeated
/// across the rest of the sheet height.
const PACKAGE_FIXTURE: &str = "../../test-data/odf/ods/grid-padding-package.ods";

fn sheets() -> Vec<litchi_odf::Sheet> {
    let mut flat = FlatSpreadsheet::open(FIXTURE).unwrap();
    flat.spreadsheet_mut().sheets().unwrap()
}

#[test]
fn full_height_and_full_width_padding_is_dropped() {
    let sheets = sheets();
    let padded = &sheets[0];

    // Two authored rows survive; the 1_048_574-row tail does not.
    assert_eq!(padded.rows.len(), 2);
    // Each row keeps only its single authored cell, not the 16_383 fillers.
    assert_eq!(padded.rows[0].cells.len(), 1);
    assert_eq!(padded.rows[0].cells[0].text, "first");
    assert_eq!(padded.rows[1].cells.len(), 1);
    assert_eq!(padded.rows[1].cells[0].text, "42");
}

#[test]
fn interior_gaps_keep_later_content_at_its_true_coordinates() {
    let sheets = sheets();
    let gaps = &sheets[1];

    // One content row, a three-row gap, then a second content row.
    assert_eq!(gaps.rows.len(), 5);
    assert_eq!(gaps.rows[0].cells[0].text, "top");
    for row in &gaps.rows[1..4] {
        assert!(row.cells.iter().all(|cell| cell.text.is_empty()));
    }

    // The two leading fillers are materialised so `bottom` keeps column 2.
    let bottom = &gaps.rows[4];
    assert_eq!(bottom.cells.len(), 3);
    assert_eq!(bottom.cells[2].text, "bottom");
    assert_eq!(bottom.cells[2].coordinates(), (4, 2));
}

#[test]
fn a_short_trailing_run_of_empty_rows_is_authored_spacing_and_survives() {
    let sheets = sheets();
    let tail = &sheets[2];

    assert_eq!(tail.rows.len(), 3);
    assert_eq!(tail.rows[0].cells[0].text, "only");
    assert!(
        tail.rows[1..]
            .iter()
            .all(|row| row.style_name.as_deref() == Some("ro2")),
        "authored row styling must survive the trailing run"
    );
}

#[test]
fn padded_sheets_still_export_only_their_used_range_to_csv() {
    let mut flat = FlatSpreadsheet::open(FIXTURE).unwrap();
    let csv = flat.spreadsheet_mut().to_csv().unwrap();

    assert!(
        csv.contains("first"),
        "used range missing from CSV: {csv:?}"
    );
    assert!(
        csv.lines().count() < 100,
        "CSV expanded the sheet padding: {} lines",
        csv.lines().count()
    );
}

/// Name, row count, and populated-cell count of every sheet.
fn signature(sheets: &[litchi_odf::Sheet]) -> Vec<(String, usize, usize)> {
    sheets
        .iter()
        .map(|sheet| {
            let populated = sheet
                .rows
                .iter()
                .flat_map(|row| row.cells.iter())
                .filter(|cell| !cell.text.is_empty() || cell.formula.is_some())
                .count();
            (sheet.name.clone(), sheet.rows.len(), populated)
        })
        .collect()
}

#[test]
fn a_producer_written_package_opens_without_expanding_its_padding() {
    let mut spreadsheet = Spreadsheet::open(PACKAGE_FIXTURE).unwrap();
    let sheets = spreadsheet.sheets().unwrap();

    assert_eq!(sheets.len(), 3);
    for sheet in &sheets {
        assert!(
            sheet.rows.len() < 4_096,
            "sheet {:?} expanded its padding to {} rows",
            sheet.name,
            sheet.rows.len()
        );
    }
    assert!(
        sheets.iter().any(|sheet| !sheet.rows.is_empty()),
        "the used range must survive"
    );
}

#[test]
fn a_padded_package_survives_the_mutable_save_round_trip() {
    // Dropping a trailing run must not leave the sheet's row structure
    // describing rows that are gone — the writer walks those ranges.
    let mut spreadsheet = Spreadsheet::open(PACKAGE_FIXTURE).unwrap();
    let before = signature(&spreadsheet.sheets().unwrap());

    let mutable = MutableSpreadsheet::from_spreadsheet(spreadsheet).unwrap();
    let saved = mutable.to_bytes().unwrap();

    let mut reopened = Spreadsheet::from_bytes(saved).unwrap();
    assert_eq!(signature(&reopened.sheets().unwrap()), before);
}
