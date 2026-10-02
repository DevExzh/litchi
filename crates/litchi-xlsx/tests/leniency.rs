#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "fixture assertions intentionally panic on failure"
)]

//! Acceptance tests for the litchi-xlsx section of `docs/LENIENCY_AUDIT.md`.
//!
//! Each test patches one member of a conformant workbook so that it carries a
//! single cataloged deviation, then requires that the owned and the
//! source-backed facades both open it by default, that the repair is listed in
//! the workbook's tolerance report with its location, and that the unpatched
//! control reports clean.

use std::io::Cursor;

use litchi_xlsx::{Defect, SourceBackedWorkbook, ToleranceLocation, ToleranceReport, Workbook};
use soapberry_zip::office::{ArchiveReader, StreamingArchiveWriter};

const BASE: &[u8] = include_bytes!("../../../test-data/ooxml/xlsx/DateFormatTests.xlsx");

/// Rebuild [`BASE`] with each named member passed through its rewrite.
fn patched(edits: &[(&str, &dyn Fn(&str) -> String)]) -> Vec<u8> {
    let reader = ArchiveReader::new(BASE).unwrap();
    let names: Vec<String> = reader.file_names().map(str::to_owned).collect();
    let mut writer = StreamingArchiveWriter::new();
    for name in names {
        let bytes = reader.read(&name).unwrap();
        let bytes = match edits.iter().find(|(member, _)| *member == name) {
            Some((_, edit)) => {
                let text = String::from_utf8(bytes).unwrap();
                let rewritten = edit(&text);
                assert_ne!(rewritten, text, "{name}: the patch must change the member");
                rewritten.into_bytes()
            },
            None => bytes,
        };
        writer.write_stored(&name, &bytes).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn replace(from: &'static str, to: &'static str) -> impl Fn(&str) -> String {
    move |text: &str| {
        assert!(text.contains(from), "patch anchor {from:?} not found");
        text.replacen(from, to, 1)
    }
}

/// Open through the owned facade, parse every worksheet, and return the
/// aggregate report.
fn owned_report(bytes: &[u8]) -> ToleranceReport {
    let workbook = Workbook::from_slice(bytes).expect("owned open");
    for sheet in workbook.sheets() {
        sheet.tolerance_report().expect("owned worksheet parse");
    }
    workbook.tolerance_report()
}

/// Open through the source-backed facade, parse every worksheet, and return
/// the aggregate report.
fn source_report(bytes: &[u8]) -> ToleranceReport {
    let workbook =
        SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec())).expect("source open");
    for sheet in workbook.sheets() {
        sheet.tolerance_report().expect("source worksheet parse");
    }
    workbook.tolerance_report()
}

/// Both facades report exactly the same repairs.
fn both_reports(bytes: &[u8]) -> ToleranceReport {
    let owned = owned_report(bytes);
    assert_eq!(owned, source_report(bytes), "facades disagree");
    owned
}

#[test]
fn the_conformant_base_reports_clean_on_every_path() {
    assert!(both_reports(BASE).is_clean());
    let package = litchi_xlsx::Package::from_slice(BASE).unwrap();
    assert!(package.workbook().unwrap().tolerance_report().is_clean());
}

#[test]
fn duplicate_relationship_id_is_folded_into_the_workbook_report() {
    let bytes = patched(&[(
        "_rels/.rels",
        &replace(
            "</Relationships>",
            r#"<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/></Relationships>"#,
        ),
    )]);
    let report = both_reports(&bytes);
    assert_eq!(report.total(), 1, "{report:?}");
    let entry = &report.defects()[0];
    assert_eq!(
        entry.defect,
        Defect::Package(litchi_opc::leniency::Defect::Relationship(
            litchi_opc::leniency::RelationshipDefect::DuplicateId
        ))
    );
    assert_eq!(entry.sheet, None);
    assert_eq!(
        entry.location,
        ToleranceLocation::Part {
            part: "/".into(),
            item: Some("rId1".into()),
            value: Some("docProps/custom.xml".into()),
        }
    );

    let package = litchi_xlsx::Package::from_slice(&bytes).unwrap();
    assert_eq!(package.workbook().unwrap().tolerance_report(), report);
}

#[test]
fn duplicate_content_type_default_is_folded_into_the_workbook_report() {
    let bytes = patched(&[(
        "[Content_Types].xml",
        &replace(
            "<Override ",
            r#"<Default Extension="XML" ContentType="text/xml"/><Override "#,
        ),
    )]);
    let report = both_reports(&bytes);
    assert_eq!(report.total(), 1, "{report:?}");
    assert_eq!(
        report.defects()[0].defect,
        Defect::Package(litchi_opc::leniency::Defect::ContentType(
            litchi_opc::leniency::ContentTypeDefect::DuplicateDefault
        ))
    );
    assert!(matches!(
        &report.defects()[0].location,
        ToleranceLocation::Part { part, .. } if &**part == "/[Content_Types].xml"
    ));
}

#[test]
fn package_framing_stays_fatal() {
    let mut truncated = BASE.to_vec();
    truncated.truncate(BASE.len() / 2);
    assert!(Workbook::from_slice(&truncated).is_err());
    assert!(SourceBackedWorkbook::from_reader(Cursor::new(truncated)).is_err());

    let malformed = patched(&[("_rels/.rels", &replace("</Relationships>", ""))]);
    assert!(Workbook::from_slice(&malformed).is_err());
    assert!(SourceBackedWorkbook::from_reader(Cursor::new(malformed)).is_err());
}

#[test]
fn core_properties_repairs_fold_through_the_public_hand_off() {
    let bytes = patched(&[(
        "docProps/core.xml",
        &replace(
            "</cp:coreProperties>",
            "<dc:title>a</dc:title><dc:title>b</dc:title></cp:coreProperties>",
        ),
    )]);
    let opc = litchi_opc::OpcPackage::from_bytes(&bytes).unwrap();
    let mut metadata = litchi_ooxml_common::leniency::ToleranceReport::default();
    litchi_ooxml_common::properties::read_with_report(&opc, &mut metadata).unwrap();
    assert!(!metadata.is_clean());

    let mut report = Workbook::from_slice(&bytes).unwrap().tolerance_report();
    assert!(report.is_clean());
    report.append_properties(&metadata);
    assert_eq!(report.total(), metadata.total());
    assert!(report.iter().all(|entry| matches!(
        (entry.defect, &entry.location),
        (Defect::Properties(_), ToleranceLocation::Part { .. })
    )));
}

fn worksheet_entries(report: &ToleranceReport) -> Vec<(Defect, Option<u32>, ToleranceLocation)> {
    report
        .iter()
        .map(|entry| (entry.defect, entry.sheet, entry.location.clone()))
        .collect()
}

/// Finding 4.1: out-of-domain row and column geometry clamps (or drops a
/// negative value) and is recorded at the row, column range, or worksheet.
#[test]
fn out_of_domain_geometry_is_clamped_and_recorded() {
    use litchi_xlsx::WorksheetDefect as W;

    let bytes = patched(&[("xl/worksheets/sheet1.xml", &|text: &str| {
        replace(r#"defaultRowHeight="13""#, r#"defaultRowHeight="-13""#)(text)
            .replacen(
                r#"<col min="2" max="2" width="17""#,
                r#"<col min="2" max="2" width="-17""#,
                1,
            )
            .replacen(
                r#"customFormat="1" ht="14""#,
                r#"customFormat="1" ht="500""#,
                1,
            )
    })]);
    let report = both_reports(&bytes);
    let column = litchi_xlsx::ColumnIndex::new(1).unwrap();
    assert_eq!(
        worksheet_entries(&report),
        [
            (
                Defect::Worksheet(W::SheetFormatDefaultOutOfRange),
                Some(0),
                ToleranceLocation::Worksheet
            ),
            (
                Defect::Worksheet(W::ColumnWidthOutOfRange),
                Some(0),
                ToleranceLocation::Columns {
                    first: column,
                    last: column
                }
            ),
            (
                Defect::Worksheet(W::RowHeightOutOfRange),
                Some(0),
                ToleranceLocation::Row(litchi_xlsx::RowIndex::new(0).unwrap())
            ),
        ]
    );

    let workbook = Workbook::from_slice(&bytes).unwrap();
    let sheet = workbook.sheet(0usize).unwrap().unwrap();
    assert_eq!(sheet.row(0).unwrap().height().map(|h| h.get()), Some(409.0));
    assert_eq!(sheet.column(1).unwrap().width(), None);
    assert_eq!(sheet.defaults().unwrap().unwrap().height().get(), 15.0);
}

/// Finding 4.3: malformed number, boolean, and date lexicals are repaired per
/// cell and recorded at the cell; a cold source-backed selected read defers
/// to the repairing parse instead of failing.
#[test]
fn malformed_value_lexicals_are_repaired_per_cell_and_recorded() {
    use litchi_xlsx::{Address, Cell, CellDefect as C, SourceCellView, Value};

    let bytes = patched(&[
        ("xl/worksheets/sheet1.xml", &|text: &str| {
            replace(
                r#"<c r="B3" s="7"/>"#,
                r#"<c r="B3" s="7" t="b"><v>TRUE</v></c>"#,
            )(text)
            .replacen(
                "</row><row r=\"3\"",
                r#"<c r="D2" t="d"><v>2026-13-01</v></c></row><row r="3""#,
                1,
            )
        }),
        (
            "xl/worksheets/sheet2.xml",
            &replace(
                r#"<c r="C2" s="10"><v>17816.607951388887</v>"#,
                r#"<c r="C2" s="10"><v>17,816.6</v>"#,
            ),
        ),
    ]);

    // A cold selected read on the source-backed facade falls back to the
    // repairing parse.
    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone())).unwrap();
    let tests = source.sheet(1usize).unwrap().unwrap();
    match tests.cell("C2").unwrap() {
        SourceCellView::Stored(Cell::Unknown(unknown)) => {
            assert_eq!(unknown.kind(), "n");
            assert_eq!(unknown.value(), Some("17,816.6"));
        },
        other => panic!("expected a demoted cell, got {other:?}"),
    }
    assert_eq!(source.tolerance_report().total(), 1);

    let cell = |a1: &str| ToleranceLocation::Cell(Address::from_a1(a1).unwrap());
    let report = both_reports(&bytes);
    assert_eq!(
        worksheet_entries(&report),
        [
            (Defect::Cell(C::MalformedDate), Some(0), cell("D2")),
            (Defect::Cell(C::MalformedBoolean), Some(0), cell("B3")),
            (Defect::Cell(C::MalformedNumber), Some(1), cell("C2")),
        ]
    );

    let workbook = Workbook::from_slice(&bytes).unwrap();
    let flags = workbook.sheet(0usize).unwrap().unwrap();
    assert!(matches!(
        flags.cell("B3").unwrap(),
        litchi_xlsx::cell::View::Stored(Cell::Value(Value::Bool(true)))
    ));
    assert!(matches!(
        flags.cell("D2").unwrap(),
        litchi_xlsx::cell::View::Stored(Cell::Value(Value::Text(text))) if text.as_str() == "2026-13-01"
    ));
}
