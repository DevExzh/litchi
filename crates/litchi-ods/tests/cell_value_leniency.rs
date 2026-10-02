//! LENIENCY_AUDIT F9.8: one unreadable typed cell value no longer fails the
//! whole spreadsheet. The cell is kept with its displayed text and its value
//! degrades to `CellValue::Unknown` carrying the declared type and the raw
//! lexical, the way LibreOffice's Calc import keeps the cell and continues.
//! Every open path records the repair with its sheet/row/column location; a
//! conforming spreadsheet records nothing, and malformed XML stays fatal.

#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "hand-built fixtures keep tolerance assertions focused"
)]

#[path = "../../litchi-odf-common/tests/support/odf_rezip.rs"]
mod odf_rezip;

use std::sync::Arc;

use litchi_core::OwnedSource;
use litchi_odf_common::{ContentDefect, Defect, DefectLocation, ToleranceReport};
use litchi_ods::worksheet::{CellValue, Sheet};
use litchi_ods::{FlatSnapshot, SourceBackedSpreadsheet, Spreadsheet};

type Recorded = (Defect, u16, Option<DefectLocation>);

fn recorded(report: &ToleranceReport) -> Vec<Recorded> {
    report
        .defects()
        .iter()
        .map(|tolerated| (tolerated.defect, tolerated.context, tolerated.location))
        .collect()
}

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("../../test-data/odf")
            .join(name),
    )
    .unwrap()
}

fn degraded(context: u16, sheet: u32, row: u64, column: u64) -> Recorded {
    (
        Defect::Content(ContentDefect::CellValueUnparseable),
        context,
        Some(DefectLocation::Cell { sheet, row, column }),
    )
}

/// Replace the `nth` (zero-based) occurrence of `from` with `to`.
fn replace_nth(text: &str, from: &str, to: &str, nth: usize) -> String {
    let (start, _) = text
        .match_indices(from)
        .nth(nth)
        .unwrap_or_else(|| panic!("fixture lacks occurrence {nth} of {from}"));
    format!("{}{to}{}", &text[..start], &text[start + from.len()..])
}

/// `calc-two-sheets.ods` with four unreadable typed values: an invalid
/// Boolean and a value-less date on sheet 0, a locale-comma float and a NaN
/// percentage on sheet 1.
fn malformed_spreadsheet() -> Vec<u8> {
    odf_rezip::edit_member(
        &fixture("corpus/calc-two-sheets.ods"),
        "content.xml",
        |xml| {
            let float = |value: &str| {
                format!(
                    r#"office:value-type="float" office:value="{value}" calcext:value-type="float""#
                )
            };
            let xml = replace_nth(
                xml,
                &float("1"),
                r#"office:value-type="boolean" office:boolean-value="yes" calcext:value-type="boolean""#,
                0,
            );
            let xml = replace_nth(
                &xml,
                &float("3"),
                r#"office:value-type="date" calcext:value-type="date""#,
                0,
            );
            let xml = replace_nth(
                &xml,
                &float("2"),
                r#"office:value-type="float" office:value="2,5" calcext:value-type="float""#,
                1,
            );
            replace_nth(
                &xml,
                &float("3"),
                r#"office:value-type="percentage" office:value="NaN" calcext:value-type="percentage""#,
                0,
            )
        },
    )
}

fn expected() -> Vec<Recorded> {
    vec![
        degraded(3, 0, 1, 0),
        degraded(4, 0, 1, 2),
        degraded(0, 1, 2, 2),
        degraded(2, 1, 2, 3),
    ]
}

fn unknown(kind: &str, value: Option<&str>) -> CellValue {
    CellValue::Unknown {
        kind: kind.to_string(),
        value: value.map(str::to_string),
    }
}

fn assert_degraded_cells(sheets: &[Sheet], label: &str) {
    let cell = |sheet: usize, row: usize, column: usize| {
        sheets[sheet]
            .cell(row, column)
            .unwrap_or_else(|| panic!("{label}: cell {sheet}/{row}/{column} exists"))
    };
    assert_eq!(
        cell(0, 1, 0).value,
        unknown("boolean", Some("yes")),
        "{label}"
    );
    assert_eq!(cell(0, 1, 0).text, "1", "{label}: displayed text is kept");
    assert_eq!(cell(0, 1, 1).value, CellValue::Number(2.0), "{label}");
    assert_eq!(cell(0, 1, 2).value, unknown("date", None), "{label}");
    assert_eq!(cell(1, 2, 1).value, CellValue::Number(1.0), "{label}");
    assert_eq!(
        cell(1, 2, 2).value,
        unknown("float", Some("2,5")),
        "{label}"
    );
    assert_eq!(cell(1, 2, 2).text, "2", "{label}: displayed text is kept");
    assert_eq!(
        cell(1, 2, 3).value,
        unknown("percentage", Some("NaN")),
        "{label}"
    );
}

fn assert_every_open_path(bytes: &[u8], defects: &[Recorded], label: &str) {
    let owned = Spreadsheet::from_bytes(bytes.to_vec()).unwrap();
    assert_eq!(
        recorded(owned.tolerance_report()),
        defects,
        "{label}: from_bytes"
    );
    assert_eq!(owned.sheet_names().len(), 2, "{label}");
    if !defects.is_empty() {
        assert_degraded_cells(owned.sheets(), &format!("{label}: from_bytes"));
    }

    let prepared = litchi_odf_common::detect::prepared(bytes.to_vec())
        .unwrap_or_else(|| panic!("{label}: package must prepare"));
    let adopted = Spreadsheet::from_prepared_package(prepared).unwrap();
    assert_eq!(
        recorded(adopted.tolerance_report()),
        defects,
        "{label}: prepared"
    );

    let source =
        SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(bytes.to_vec()))).unwrap();
    assert_eq!(
        recorded(source.tolerance_report()),
        defects,
        "{label}: source-backed"
    );
    if !defects.is_empty() {
        assert_degraded_cells(source.sheets().unwrap(), &format!("{label}: source-backed"));
    }
}

#[test]
fn unreadable_cell_values_degrade_and_are_recorded_on_every_open_path() {
    assert_every_open_path(&malformed_spreadsheet(), &expected(), "malformed");
}

#[test]
fn the_same_spreadsheet_without_the_deviation_repairs_nothing() {
    let conforming = fixture("corpus/calc-two-sheets.ods");
    assert_every_open_path(&conforming, &[], "original");
    let rezipped = odf_rezip::rezip(&conforming, odf_rezip::MimetypeLayout::CONFORMING);
    assert_every_open_path(&rezipped, &[], "re-zipped");
}

#[test]
fn package_and_content_repairs_are_reported_together() {
    // A re-zipped package with a deflated `mimetype` and an unreadable cell:
    // the package repair comes first, then the content repair.
    let bytes = odf_rezip::rezip(
        &malformed_spreadsheet(),
        odf_rezip::MimetypeLayout {
            deflated: true,
            ..odf_rezip::MimetypeLayout::CONFORMING
        },
    );
    let owned = Spreadsheet::from_bytes(bytes).unwrap();
    let mut defects = vec![(
        Defect::Package(litchi_odf_common::PackageDefect::MimetypeDeflated),
        8,
        None,
    )];
    defects.extend(expected());
    assert_eq!(recorded(owned.tolerance_report()), defects);
}

#[test]
fn flat_spreadsheets_degrade_and_record_the_same_way() {
    let conforming = include_str!("fixtures/odfdo-test-flat-lo-compact.fods").to_string();
    let clean = FlatSnapshot::from_bytes(conforming.clone().into_bytes()).unwrap();
    assert!(clean.tolerance_report().is_clean());

    let malformed = replace_nth(
        &conforming,
        r#"office:value-type="float" office:value="123""#,
        r#"office:value-type="float" office:value="one-two-three""#,
        0,
    );
    let flat = FlatSnapshot::from_bytes(malformed.into_bytes()).unwrap();
    let defects = recorded(flat.tolerance_report());
    assert_eq!(defects.len(), 1);
    assert_eq!(
        defects[0].0,
        Defect::Content(ContentDefect::CellValueUnparseable)
    );
    let Some(DefectLocation::Cell { sheet, row, column }) = defects[0].2 else {
        panic!("a degraded cell value is located");
    };
    let row = usize::try_from(row).unwrap();
    let column = usize::try_from(column).unwrap();
    assert_eq!(
        flat.sheets()[sheet as usize]
            .cell(row, column)
            .unwrap()
            .value,
        unknown("float", Some("one-two-three"))
    );
}

#[test]
fn malformed_xml_stays_a_hard_error() {
    // An unclosed cell is broken XML, not a cell value.
    let broken = odf_rezip::edit_member(
        &fixture("corpus/calc-two-sheets.ods"),
        "content.xml",
        |xml| xml.replacen("</table:table-cell>", "", 1),
    );
    assert!(Spreadsheet::from_bytes(broken.clone()).is_err());
    assert!(SourceBackedSpreadsheet::from_read_at(Arc::new(OwnedSource::new(broken))).is_err());
}
