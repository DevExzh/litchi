//! XLSX worksheet selection through the `Xlsx` handle's `sheet` method.
//!
//! Selection is XLSX's: names match case-insensitively (`"dAtA"` selects
//! `Data`), positions are zero-based, selecting is catalog-only, and cell
//! reads keep the exact XLSX cell state and the typed XLSX error.
//!
//! M0a step S8c ported these tests from the facade workbook's sheet selector
//! to `litchi::open` / `File::Xlsx(xlsx) => xlsx.sheet(..)`.

#![cfg(feature = "xlsx")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use std::io::{Cursor, Write};

use litchi::error::Source;
use litchi::format::xlsx::{SelectedCellView, SelectedWorksheet, Xlsx};
use litchi::xlsx::{Cell, Error as XlsxError, Value};
use litchi::{Code, File};
use zip::{CompressionMethod, ZipWriter, write::SimpleFileOptions};

const CONTENT_TYPES: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const PACKAGE_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const DOCUMENT_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn workbook_bytes(first_value: &str, second_sheet: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/chartsheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="{DOCUMENT_RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    );
    let workbook = format!(
        r#"<workbook xmlns="{SPREADSHEETML}" xmlns:r="{DOCUMENT_RELATIONSHIPS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Broken" sheetId="2" r:id="rId2"/><sheet name="Chart" sheetId="3" r:id="rId3"/></sheets></workbook>"#
    );
    let workbook_relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="{DOCUMENT_RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{DOCUMENT_RELATIONSHIPS}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{DOCUMENT_RELATIONSHIPS}/chartsheet" Target="chartsheets/sheet3.xml"/></Relationships>"#
    );
    let first_sheet = format!(
        r#"<worksheet xmlns="{SPREADSHEETML}"><sheetData><row r="1"><c r="A1"><v>{first_value}</v></c><c r="B1"/><c r="C1"><f>A1*2</f><v>14</v></c><c r="D1" t="future"><v>opaque</v></c><c r="E1" t="inlineStr"><is><t>anchor</t></is></c></row><row r="3"><c r="F3" t="inlineStr"><is><t>tail</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="E1:F1"/></mergeCells></worksheet>"#
    );
    let chart_sheet = format!(r#"<chartsheet xmlns="{SPREADSHEETML}"/>"#);

    let mut output = Cursor::new(Vec::new());
    let mut writer = ZipWriter::new(&mut output);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let entries: [(&str, &[u8]); 7] = [
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root_relationships.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        (
            "xl/_rels/workbook.xml.rels",
            workbook_relationships.as_bytes(),
        ),
        ("xl/worksheets/sheet1.xml", first_sheet.as_bytes()),
        ("xl/worksheets/sheet2.xml", second_sheet),
        ("xl/chartsheets/sheet3.xml", chart_sheet.as_bytes()),
    ];
    for (name, bytes) in entries {
        writer.start_file(name, options).expect("start XLSX member");
        writer.write_all(bytes).expect("write XLSX member");
    }
    writer.finish().expect("finish XLSX fixture");
    output.into_inner()
}

fn valid_second_sheet() -> &'static [u8] {
    br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>"#
}

fn assert_send_sync<T: Send + Sync>() {}

/// The XLSX handle of an opened workbook.
#[track_caller]
fn xlsx(file: &File) -> &Xlsx {
    match file {
        File::Xlsx(xlsx) => xlsx,
        other => panic!("expected an XLSX workbook, got {other:?}"),
    }
}

/// The XLSX error a selection or cell read reported.
#[track_caller]
fn xlsx_error(error: &litchi::Error) -> &XlsxError {
    match error.leaf() {
        Source::Xlsx(error) => error,
        other => panic!("expected an XLSX error, got {other:?}"),
    }
}

/// `error` reports that the source changed after open.
#[track_caller]
fn assert_source_changed(error: &litchi::Error) {
    assert_eq!(error.code(), Code::IoError, "{error:?}");
    let changed = matches!(
        error.leaf(),
        Source::Core(litchi::common::Error::SourceChanged { .. })
            | Source::Opc(litchi::opc::OpcError::SourceChanged { .. })
            | Source::Xlsx(XlsxError::Package(
                litchi::opc::OpcError::SourceChanged { .. }
            ))
    );
    assert!(changed, "expected a source change, got {error:?}");
}

#[test]
fn public_selector_preserves_exact_owned_cell_semantics() {
    assert_send_sync::<SelectedWorksheet>();

    let file = litchi::open_bytes(workbook_bytes("7", valid_second_sheet()))
        .expect("open source-backed unified XLSX");
    let workbook = xlsx(&file);
    let selected = workbook
        .sheet("dAtA")
        .expect("select by case-folded name")
        .expect("Data sheet exists");
    let by_position = workbook
        .sheet(0usize)
        .expect("select by position")
        .expect("first sheet exists");

    assert_eq!(selected.name(), "Data");
    assert_eq!(selected.position(), 0);
    assert_eq!(by_position.name(), selected.name());
    assert!(format!("{selected:?}").contains("Data"));
    assert!(workbook.sheet("missing").unwrap().is_none());
    assert!(workbook.sheet(3usize).unwrap().is_none());

    let a1 = selected.cell("A1").expect("read A1");
    assert_eq!(a1, selected.cell((0_u32, 0_u32)).unwrap());
    assert!(matches!(
        a1,
        SelectedCellView::Stored(Cell::Value(Value::Number(ref value)))
            if value.as_str() == "7"
    ));
    assert!(matches!(
        selected.cell("B1").unwrap(),
        SelectedCellView::Stored(Cell::Empty)
    ));
    assert!(matches!(
        selected.cell("C1").unwrap(),
        SelectedCellView::Stored(Cell::Formula(_))
    ));
    assert!(matches!(
        selected.cell("D1").unwrap(),
        SelectedCellView::Stored(Cell::Unknown(ref value))
            if value.kind() == "future" && value.value() == Some("opaque")
    ));
    assert!(matches!(
        selected.cell("F1").unwrap(),
        SelectedCellView::Covered(range) if range.a1() == "E1:F1"
    ));
    assert!(matches!(
        selected.cell("A2").unwrap(),
        SelectedCellView::Missing
    ));

    let cells = selected.cells("A1:F3").expect("read sparse range");
    assert_eq!(
        cells
            .iter()
            .map(|entry| entry.address.a1())
            .collect::<Vec<_>>(),
        ["A1", "B1", "C1", "D1", "E1", "F3"]
    );
    assert_eq!(
        selected.cells((0_u32, 0_u32, 1_u32, 5_u32)).unwrap(),
        cells[..5]
    );

    // A coordinate or range the caller passed is the caller's error, not a
    // corrupt file (M0a S9); the typed leaf is unchanged.
    let coordinate_error = selected.cell("A0").unwrap_err();
    assert!(matches!(
        xlsx_error(&coordinate_error),
        XlsxError::Coordinate(_)
    ));
    assert_eq!(coordinate_error.code(), Code::InvalidArgument);
    let range_error = selected.cells((1_u32, 1_u32, 1_u32, 2_u32)).unwrap_err();
    assert!(matches!(xlsx_error(&range_error), XlsxError::Range(_)));
    assert_eq!(range_error.code(), Code::InvalidArgument);
    let outside = selected.cell((1_048_576_u32, 0_u32)).unwrap_err();
    assert!(matches!(xlsx_error(&outside), XlsxError::Coordinate(_)));
    assert_eq!(outside.code(), Code::OutOfRange);
}

#[test]
fn selected_access_defers_unselected_malformed_worksheet() {
    let malformed = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">"#;
    let file = litchi::open_bytes(workbook_bytes("7", malformed))
        .expect("catalog does not read malformed worksheet payload");
    let workbook = xlsx(&file);

    let data = workbook.sheet("Data").unwrap().unwrap();
    assert!(matches!(
        data.cell("A1").unwrap(),
        SelectedCellView::Stored(Cell::Value(Value::Number(_)))
    ));
    let broken = workbook
        .sheet("Broken")
        .expect("selecting a handle remains catalog-only")
        .expect("Broken catalog entry exists");
    assert!(broken.cell("A1").is_err());
}

#[test]
fn non_grid_sheet_returns_typed_not_worksheet_error() {
    let file = litchi::open_bytes(workbook_bytes("7", valid_second_sheet())).unwrap();
    let chart = xlsx(&file).sheet("Chart").unwrap().unwrap();
    let error = chart.cell("A1").unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedType, "{error:?}");
    assert!(matches!(
        xlsx_error(&error),
        XlsxError::NotWorksheet { sheet } if sheet == "Chart"
    ));
}

/// The facade workbook's selector refused a non-XLSX workbook at run time
/// (`Unsupported("... only for XLSX")`). The selector is now a method of
/// the XLSX handle, so the refusal is static: an XLSB workbook opens as
/// `File::Xlsb`, whose handle has no `sheet` method, and its generic
/// workbook view still reads.
#[cfg(feature = "xlsb")]
#[test]
fn non_xlsx_workbook_returns_explicit_unsupported_error() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/Simple.xlsb"
    ))
    .expect("read XLSB fixture");
    let file = litchi::open_bytes(bytes).expect("open unified XLSB workbook");
    assert!(matches!(file, File::Xlsb(_)), "{file:?}");
    assert!(
        !file
            .workbook()
            .unwrap()
            .worksheet_names()
            .unwrap()
            .is_empty()
    );
}

#[cfg(any(unix, windows))]
#[test]
fn selected_handle_preserves_source_change_error() {
    let temporary = tempfile::NamedTempFile::new().expect("temporary XLSX");
    std::fs::write(temporary.path(), workbook_bytes("7", valid_second_sheet()))
        .expect("write initial XLSX");
    let file = litchi::open(temporary.path()).expect("open positional XLSX");
    let workbook = xlsx(&file);
    let selected = workbook.sheet("Data").unwrap().unwrap();

    std::fs::write(
        temporary.path(),
        workbook_bytes("12345678901234567890", valid_second_sheet()),
    )
    .expect("replace XLSX source");

    assert_source_changed(&selected.cell("A1").unwrap_err());
    assert_source_changed(&workbook.sheet("Data").unwrap_err());
}
