#![cfg(feature = "xlsx")]

use std::io::{Cursor, Write};

use litchi::sheet::{SelectedCellView, SelectedWorksheet, Workbook};
use litchi::xlsx::{Cell, Error as XlsxError, Value};
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

#[test]
fn public_selector_preserves_exact_owned_cell_semantics() {
    assert_send_sync::<SelectedWorksheet>();

    let workbook = Workbook::from_bytes(workbook_bytes("7", valid_second_sheet()))
        .expect("open source-backed unified XLSX");
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

    let coordinate_error = selected.cell("A0").unwrap_err();
    assert!(matches!(
        coordinate_error.downcast_ref::<XlsxError>(),
        Some(XlsxError::Coordinate(_))
    ));
    let range_error = selected.cells((1_u32, 1_u32, 1_u32, 2_u32)).unwrap_err();
    assert!(matches!(
        range_error.downcast_ref::<XlsxError>(),
        Some(XlsxError::Range(_))
    ));
}

#[test]
fn selected_access_defers_unselected_malformed_worksheet() {
    let malformed = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">"#;
    let workbook = Workbook::from_bytes(workbook_bytes("7", malformed))
        .expect("catalog does not read malformed worksheet payload");

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
    let workbook = Workbook::from_bytes(workbook_bytes("7", valid_second_sheet())).unwrap();
    let chart = workbook.sheet("Chart").unwrap().unwrap();
    let error = chart.cell("A1").unwrap_err();
    assert!(matches!(
        error.downcast_ref::<XlsxError>(),
        Some(XlsxError::NotWorksheet { sheet }) if sheet == "Chart"
    ));
}

#[cfg(feature = "xlsb")]
#[test]
fn non_xlsx_workbook_returns_explicit_unsupported_error() {
    let bytes = std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/Simple.xlsb"
    ))
    .expect("read XLSB fixture");
    let workbook = Workbook::from_bytes(bytes).expect("open unified XLSB workbook");
    let error = workbook.sheet(0usize).unwrap_err();
    assert!(matches!(
        error.downcast_ref::<litchi::Error>(),
        Some(litchi::Error::Unsupported(message)) if message.contains("only for XLSX")
    ));
}

#[cfg(any(unix, windows))]
#[test]
fn selected_handle_preserves_source_change_error() {
    let file = tempfile::NamedTempFile::new().expect("temporary XLSX");
    std::fs::write(file.path(), workbook_bytes("7", valid_second_sheet()))
        .expect("write initial XLSX");
    let workbook = Workbook::open(file.path()).expect("open positional XLSX");
    let selected = workbook.sheet("Data").unwrap().unwrap();

    std::fs::write(
        file.path(),
        workbook_bytes("12345678901234567890", valid_second_sheet()),
    )
    .expect("replace XLSX source");

    let error = selected.cell("A1").unwrap_err();
    assert!(matches!(
        error.downcast_ref::<litchi::Error>(),
        Some(litchi::Error::SourceChanged { .. })
    ));
    let selection_error = workbook.sheet("Data").unwrap_err();
    assert!(matches!(
        selection_error.downcast_ref::<litchi::Error>(),
        Some(litchi::Error::SourceChanged { .. })
    ));
}
