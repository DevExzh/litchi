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

/// Finding 4.5: malformed `sst` count hints are ignored and recorded at the
/// table; a cold source-backed shared-string read falls back to the
/// repairing parse.
#[test]
fn malformed_shared_string_count_hints_are_ignored_and_recorded() {
    use litchi_xlsx::{Cell, SharedStringsDefect, SourceCellView, Value};

    let bytes = patched(&[(
        "xl/sharedStrings.xml",
        &replace(
            r#"count="104" uniqueCount="48""#,
            r#"count="-1" uniqueCount="NaN""#,
        ),
    )]);

    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone())).unwrap();
    let flags = source.sheet(0usize).unwrap().unwrap();
    assert!(matches!(
        flags.cell("A1").unwrap(),
        SourceCellView::Stored(Cell::Value(Value::Text(_)))
    ));

    let report = both_reports(&bytes);
    let hint = (
        Defect::SharedStrings(SharedStringsDefect::CountHint),
        None,
        ToleranceLocation::SharedStringTable,
    );
    assert_eq!(worksheet_entries(&report), [hint.clone(), hint]);
    assert_eq!(source.tolerance_report(), report);
}

/// Finding 4.6: an orphaned shared-formula member keeps its cached value,
/// loses the formula it cannot be given, and is recorded at the cell.
#[test]
fn orphaned_shared_formula_member_keeps_its_cached_value() {
    use litchi_xlsx::{Address, Cell, CellDefect, SourceCellView, Value};

    let bytes = patched(&[(
        "xl/worksheets/sheet1.xml",
        &replace(
            r#"<c r="B3" s="7"/>"#,
            r#"<c r="B3" s="7"><f t="shared" si="7"/><v>12</v></c>"#,
        ),
    )]);

    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone())).unwrap();
    let flags = source.sheet(0usize).unwrap().unwrap();
    assert!(matches!(
        flags.cell("B3").unwrap(),
        SourceCellView::Stored(Cell::Value(Value::Number(number))) if number.as_str() == "12"
    ));

    let report = both_reports(&bytes);
    assert_eq!(
        worksheet_entries(&report),
        [(
            Defect::Cell(CellDefect::BrokenSharedFormulaGroup),
            Some(0),
            ToleranceLocation::Cell(Address::from_a1("B3").unwrap())
        )]
    );
}

/// Finding 4.7: a stale `activeTab` falls back to the first sheet and a
/// defined name scoped to a missing sheet is dropped; both are recorded at
/// open.
#[test]
fn stale_active_tab_and_local_sheet_id_are_repaired_at_open() {
    use litchi_xlsx::CatalogDefect;

    let bytes = patched(&[("xl/workbook.xml", &|text: &str| {
        replace(r#"activeTab="1""#, r#"activeTab="9""#)(text).replacen(
                "</sheets>",
                r#"</sheets><definedNames><definedName name="Stale" localSheetId="5">Flags!$A$1</definedName><definedName name="Kept" localSheetId="1">Tests!$A$1</definedName></definedNames>"#,
                1,
            )
    })]);
    let report = both_reports(&bytes);
    assert_eq!(
        worksheet_entries(&report),
        [
            (
                Defect::Catalog(CatalogDefect::ActiveTabOutOfRange),
                None,
                ToleranceLocation::Workbook
            ),
            (
                Defect::Catalog(CatalogDefect::LocalSheetIdOutOfRange),
                None,
                ToleranceLocation::DefinedName(0)
            ),
        ]
    );

    let workbook = Workbook::from_slice(&bytes).unwrap();
    assert_eq!(workbook.active_sheet().unwrap().name(), "Flags");
    let names = workbook
        .defined_names()
        .iter()
        .map(|name| name.name.as_str())
        .collect::<Vec<_>>();
    assert_eq!(names, ["Kept"]);
    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
    assert_eq!(source.active_sheet().unwrap().name(), "Flags");
}

/// Finding 4.8 (opt-in): a sheet name breaking Excel's authoring rules fails
/// the default open, and opens verbatim with a recorded repair under
/// `Leniency::TolerateSheetNameRules`.
#[test]
fn sheet_name_authoring_rules_are_tolerated_only_when_opted_in() {
    use litchi_xlsx::{CatalogDefect, Leniency};

    const NAME: &str = "Q3: revenue [draft] by region and product line";
    let bytes = patched(&[(
        "xl/workbook.xml",
        &replace(
            r#"<sheet name="Flags""#,
            r#"<sheet name="Q3: revenue [draft] by region and product line""#,
        ),
    )]);

    assert!(Workbook::from_slice(&bytes).is_err());
    assert!(SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone())).is_err());

    let lenient = Leniency::TolerateSheetNameRules;
    let workbook = Workbook::from_slice_with_leniency(&bytes, lenient).unwrap();
    let source =
        SourceBackedWorkbook::from_reader_with_leniency(Cursor::new(bytes.clone()), lenient)
            .unwrap();
    assert_eq!(workbook.leniency(), lenient);
    assert_eq!(workbook.sheet(0usize).unwrap().unwrap().name(), NAME);
    assert_eq!(source.sheet(0usize).unwrap().unwrap().name(), NAME);
    assert!(workbook.sheet(NAME).unwrap().is_some());
    for sheet in workbook.sheets() {
        sheet.tolerance_report().unwrap();
    }
    for sheet in source.sheets() {
        sheet.tolerance_report().unwrap();
    }
    let report = workbook.tolerance_report();
    assert_eq!(report, source.tolerance_report());
    assert_eq!(
        worksheet_entries(&report),
        [(
            Defect::Catalog(CatalogDefect::SheetNameRules),
            None,
            ToleranceLocation::Sheet(0)
        )]
    );

    // The opt-in admits only authoring-rule violations: a conformant
    // workbook reports clean, and an empty name stays fatal.
    assert!(
        Workbook::from_slice_with_leniency(BASE, lenient)
            .unwrap()
            .tolerance_report()
            .is_clean()
    );
    let empty = patched(&[(
        "xl/workbook.xml",
        &replace(r#"<sheet name="Flags""#, r#"<sheet name="""#),
    )]);
    assert!(Workbook::from_slice_with_leniency(&empty, lenient).is_err());
    assert!(SourceBackedWorkbook::from_reader_with_leniency(Cursor::new(empty), lenient).is_err());
}

/// Finding 4.9: a leading `=` is stripped and an empty expression dropped,
/// each recorded at the cell; a cold source-backed read defers to the
/// repairing parse.
#[test]
fn leading_equals_and_empty_formulas_are_repaired_and_recorded() {
    use litchi_xlsx::{Address, Cell, CellDefect as C, SourceCellView, Value};

    let bytes = patched(&[
        (
            "xl/worksheets/sheet1.xml",
            &replace(
                r#"<c r="B3" s="7"/>"#,
                r#"<c r="B3" s="7"><f>=1+1</f><v>2</v></c>"#,
            ),
        ),
        (
            "xl/worksheets/sheet2.xml",
            &replace(
                r#"<c r="C2" s="10"><v>17816.607951388887</v>"#,
                r#"<c r="C2" s="10"><f></f><v>17816.607951388887</v>"#,
            ),
        ),
    ]);

    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone())).unwrap();
    match source.sheet(0usize).unwrap().unwrap().cell("B3").unwrap() {
        SourceCellView::Stored(Cell::Formula(formula)) => assert_eq!(formula.text(), "1+1"),
        other => panic!("expected a formula, got {other:?}"),
    }
    assert!(matches!(
        source.sheet(1usize).unwrap().unwrap().cell("C2").unwrap(),
        SourceCellView::Stored(Cell::Value(Value::Number(_)))
    ));

    let cell = |a1: &str| ToleranceLocation::Cell(Address::from_a1(a1).unwrap());
    assert_eq!(
        worksheet_entries(&both_reports(&bytes)),
        [
            (Defect::Cell(C::FormulaLeadingEquals), Some(0), cell("B3")),
            (Defect::Cell(C::EmptyFormula), Some(1), cell("C2")),
        ]
    );
}

/// Finding 4.10: schema-sequence violations and advisory counts are
/// repaired and recorded; rows out of order sort and a repeated row merges
/// last-wins.
#[test]
fn worksheet_sequencing_and_advisory_metadata_are_repaired_and_recorded() {
    use litchi_xlsx::{Cell, SourceCellView, Value, WorksheetDefect as W};

    let bytes = patched(&[("xl/worksheets/sheet1.xml", &|text: &str| {
        replace(r#"<dimension ref="A1:C4"/>"#, "")(text).replacen(
                "</row></sheetData>",
                r#"</row><row r="2"><c r="D2"><v>5</v></c></row></sheetData><mergeCells count="3"><mergeCell ref="E1:F1"/></mergeCells><dimension ref="A1:D4"/>"#,
                1,
            )
    })]);

    let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.clone())).unwrap();
    assert!(matches!(
        source.sheet(0usize).unwrap().unwrap().cell("D2").unwrap(),
        SourceCellView::Stored(Cell::Value(Value::Number(number))) if number.as_str() == "5"
    ));

    assert_eq!(
        worksheet_entries(&both_reports(&bytes)),
        [
            (
                Defect::Worksheet(W::RowOutOfOrder),
                Some(0),
                ToleranceLocation::Row(litchi_xlsx::RowIndex::new(1).unwrap())
            ),
            (
                Defect::Worksheet(W::CountMismatch),
                Some(0),
                ToleranceLocation::Worksheet
            ),
            (
                Defect::Worksheet(W::MisplacedElement),
                Some(0),
                ToleranceLocation::Worksheet
            ),
        ]
    );

    let workbook = Workbook::from_slice(&bytes).unwrap();
    let flags = workbook.sheet(0usize).unwrap().unwrap();
    // The repeated row 2 merges last-wins: its earlier cells survive.
    assert!(matches!(
        flags.cell("A2").unwrap(),
        litchi_xlsx::cell::View::Stored(Cell::Value(Value::Text(_)))
    ));
    assert_eq!(
        flags.extents().unwrap().declared(),
        Some(litchi_xlsx::Rect::from_a1("A1:D4").unwrap())
    );
}
