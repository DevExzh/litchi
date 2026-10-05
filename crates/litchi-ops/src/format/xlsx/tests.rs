//! `Xlsx` handle tests: every constructor route on real fixtures, parity
//! with the facade's `sheet::Workbook` output, the source-freshness guard,
//! the generic sheet traits, the XLSX selector semantics carried over from
//! `crates/litchi/tests/unified_sheet_selectors.rs`, and editor construction.

use std::io::{Cursor, Write};
use std::path::PathBuf;

use litchi_core::sheet::{CellValue, WorkbookTrait};
use litchi_opc::SourceBackedPackage;
use litchi_xlsx::{Cell, Value};

use super::{SelectedCellView, Xlsx, XlsxEditor, XlsxOwner};
use crate::error::{Code, Error, Source};
use crate::options::{OpenOptions, SourceMode};

fn fixture_path(name: &str) -> PathBuf {
    PathBuf::from(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/"
    ))
    .join(name)
}

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(fixture_path(name)).expect("read fixture")
}

/// The path route, through the unified opener.
#[cfg(any(unix, windows))]
fn open_path(path: &std::path::Path, options: &OpenOptions) -> Xlsx {
    match crate::open_with(path, options).expect("path open") {
        crate::File::Xlsx(xlsx) => xlsx,
        other => panic!("{} opened as {other:?}", path.display()),
    }
}

/// The byte route, through the unified opener.
fn open_bytes(bytes: Vec<u8>, options: &OpenOptions) -> Xlsx {
    match crate::open_bytes_with(bytes, options).expect("bytes open") {
        crate::File::Xlsx(xlsx) => xlsx,
        other => panic!("bytes opened as {other:?}"),
    }
}

/// The eager fallback route: smart detection's parsed OPC package.
fn open_eager(bytes: Vec<u8>, options: &OpenOptions) -> Xlsx {
    let Some(crate::detect::DetectedFormat::Xlsx(package)) =
        crate::detect::detect_format_smart(bytes)
    else {
        panic!("smart detection did not find XLSX");
    };
    Xlsx::from_opc_package(package, options).expect("eager open")
}

/// The prepared route: `detect_and_prepare`, then the unified opener.
#[cfg(any(unix, windows))]
fn open_prepared(path: &std::path::Path, options: &OpenOptions) -> Xlsx {
    let detection = crate::detect::detect_and_prepare(path).expect("prepared OOXML source");
    match crate::open_prepared_with(detection, options).expect("prepared open") {
        crate::File::Xlsx(xlsx) => xlsx,
        other => panic!("{} opened as {other:?}", path.display()),
    }
}

fn is_source(handle: &Xlsx) -> bool {
    matches!(handle.0.owner, XlsxOwner::Source(_))
}

fn fnv(text: &str) -> u64 {
    let mut hash: u64 = 0xcbf2_9ce4_8422_2325;
    for byte in text.bytes() {
        hash ^= u64::from(byte);
        hash = hash.wrapping_mul(0x0100_0000_01b3);
    }
    hash
}

/// `(fixture, text length, FNV-1a of text, sheet names, author)`, as the
/// facade's `sheet::Workbook::{open, from_bytes}` report them at the base
/// of this step (captured with a throwaway facade example; both facade
/// routes agree).
const FACADE: &[(&str, usize, u64, &[&str], &str)] = &[
    (
        "xlsx/DateFormatTests.xlsx",
        2417,
        0x6fb2_3c8a_898e_d3d1,
        &["Flags", "Tests"],
        "Ken Arnold",
    ),
    (
        "xlsx/InlineStrings.xlsx",
        169,
        0x5763_71dd_1909_be93,
        &["Sheet1", "Sheet2", "Sheet3"],
        "Nick Burch",
    ),
    (
        "xlsx/ExcelTables.xlsx",
        118,
        0x00c2_0376_5ce8_4a05,
        &["ExcelTable"],
        "",
    ),
];

fn assert_facade_parity(handle: &Xlsx, expected: &(&str, usize, u64, &[&str], &str)) {
    let (name, length, hash, names, author) = *expected;
    let text = handle.text().expect("text");
    assert_eq!(text.len(), length, "{name}: text length");
    assert_eq!(fnv(&text), hash, "{name}: text digest");
    assert_eq!(handle.worksheet_names().unwrap(), names, "{name}: names");
    assert_eq!(handle.worksheet_count().unwrap(), names.len());
    let metadata = handle
        .metadata()
        .unwrap()
        .expect("XLSX metadata is always reported");
    assert_eq!(metadata.author.as_deref(), Some(author), "{name}: author");
    assert_eq!(metadata.title, None);
    assert!(handle.tolerance_report().unwrap().is_clean(), "{name}");
}

#[cfg(any(unix, windows))]
#[test]
fn every_route_matches_the_facade_on_real_fixtures() {
    for expected in FACADE {
        let path = fixture_path(expected.0);
        let bytes = fixture(expected.0);

        let by_path = open_path(&path, &OpenOptions::new());
        assert!(is_source(&by_path));
        assert_facade_parity(&by_path, expected);

        let by_bytes = open_bytes(bytes.clone(), &OpenOptions::new());
        assert!(is_source(&by_bytes));
        assert_facade_parity(&by_bytes, expected);

        let by_prepared = open_prepared(&path, &OpenOptions::new());
        assert!(is_source(&by_prepared));
        assert_facade_parity(&by_prepared, expected);

        let eager = open_eager(bytes.clone(), &OpenOptions::new());
        assert!(!is_source(&eager));
        assert_facade_parity(&eager, expected);

        let forced = open_bytes(bytes, &OpenOptions::new().source_mode(SourceMode::Eager));
        assert!(!is_source(&forced));
        assert_facade_parity(&forced, expected);
        let forced = open_path(&path, &OpenOptions::new().source_mode(SourceMode::Eager));
        assert!(!is_source(&forced));
        assert_facade_parity(&forced, expected);
    }
}

#[test]
fn metadata_matches_the_leaf_core_properties_read() {
    let bytes = fixture("xlsx/DateFormatTests.xlsx");
    let package = SourceBackedPackage::from_vec(bytes.clone()).unwrap();
    let (leaf, _) =
        crate::detect::workbook_metadata::read_source_backed_metadata(&package).unwrap();
    let handle = open_bytes(bytes, &OpenOptions::new());
    assert_eq!(
        format!("{:?}", handle.metadata().unwrap()),
        format!("{:?}", Some(leaf))
    );
}

// --- Synthetic two-sheet fixture from the facade adapter's tests ---------

const SPREADSHEETML: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_RELATIONSHIPS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn zip(entries: &[(&str, &[u8])]) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    let mut writer = zip::ZipWriter::new(&mut output);
    let options =
        zip::write::SimpleFileOptions::default().compression_method(zip::CompressionMethod::Stored);
    for (name, bytes) in entries {
        writer.start_file(*name, options).expect("start member");
        writer.write_all(bytes).expect("write member");
    }
    writer.finish().expect("finish archive");
    output.into_inner()
}

/// `First` (A1 = 7, B2 = "second") and `Second` (A1 = 9), 1904 dates,
/// `Second` active.
fn two_sheets() -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#;
    let root = format!(
        r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="{RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    );
    let workbook = format!(
        r#"<workbook xmlns="{SPREADSHEETML}" xmlns:r="{RELATIONSHIPS}"><workbookPr date1904="1"/><bookViews><workbookView activeTab="1"/></bookViews><sheets><sheet name="First" sheetId="1" r:id="rId1"/><sheet name="Second" sheetId="2" r:id="rId2"/></sheets></workbook>"#
    );
    let relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="{RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{RELATIONSHIPS}/worksheet" Target="worksheets/sheet2.xml"/></Relationships>"#
    );
    let first = format!(
        r#"<worksheet xmlns="{SPREADSHEETML}"><dimension ref="A1:B2"/><sheetData><row r="1"><c r="A1"><v>7</v></c></row><row r="2"><c r="B2" t="inlineStr"><is><t>second</t></is></c></row></sheetData></worksheet>"#
    );
    let second = format!(
        r#"<worksheet xmlns="{SPREADSHEETML}"><dimension ref="A1"/><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>"#
    );
    zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", relationships.as_bytes()),
        ("xl/worksheets/sheet1.xml", first.as_bytes()),
        ("xl/worksheets/sheet2.xml", second.as_bytes()),
    ])
}

fn assert_workbook_trait_surface(workbook: &dyn WorkbookTrait) {
    assert_eq!(workbook.worksheet_names(), ["First", "Second"]);
    assert_eq!(workbook.worksheet_count(), 2);
    assert_eq!(workbook.active_sheet_index(), 1);
    assert!(workbook.is_1904_date_system());

    let active = workbook.active_worksheet().expect("active worksheet");
    assert_eq!(active.name(), "Second");
    let by_name = workbook.worksheet_by_name("First").expect("by name");
    let by_index = workbook.worksheet_by_index(0).expect("by index");
    assert_eq!(by_name.name(), by_index.name());
    let missing = workbook
        .worksheet_by_index(2)
        .err()
        .expect("index 2 is out of range");
    assert_eq!(
        missing.downcast_ref::<Error>().map(Error::code),
        Some(Code::OutOfRange)
    );
    let missing = workbook
        .worksheet_by_name("Third")
        .err()
        .expect("no such sheet");
    assert_eq!(
        missing.downcast_ref::<Error>().map(Error::code),
        Some(Code::NotFound)
    );

    let mut worksheets = workbook.worksheets();
    assert_eq!(worksheets.next().unwrap().unwrap().name(), "First");
    assert_eq!(worksheets.next().unwrap().unwrap().name(), "Second");
    assert!(worksheets.next().is_none());

    assert_eq!(by_name.row_count(), 2);
    assert_eq!(by_name.column_count(), 2);
    assert_eq!(by_name.dimensions(), Some((0, 0, 1, 1)));
    assert_eq!(
        by_name.cell_value(1, 1).unwrap().as_ref(),
        &CellValue::Int(7)
    );
    assert_eq!(
        by_name.cell_by_coordinate("B2").unwrap().value(),
        &CellValue::String("second".to_owned())
    );
    let cell = by_name.cell(1, 1).expect("cell");
    assert_eq!(cell.coordinate(), "A1");
    assert_eq!((cell.row(), cell.column()), (1, 1));
    assert!(!cell.is_empty());
    assert!(!cell.is_formula());

    let mut cells = by_name.cells();
    assert_eq!(cells.next().unwrap().unwrap().coordinate(), "A1");
    assert_eq!(cells.next().unwrap().unwrap().coordinate(), "B2");
    assert!(cells.next().is_none());
    assert_eq!(
        by_name.row(1).unwrap().as_ref(),
        &[CellValue::Empty, CellValue::String("second".to_owned())]
    );
    let mut rows = by_name.rows();
    assert_eq!(rows.next().unwrap().unwrap().len(), 2);
    assert_eq!(rows.next().unwrap().unwrap().len(), 2);
    assert!(rows.next().is_none());
}

#[test]
fn workbook_trait_surface_matches_across_owner_states() {
    let bytes = two_sheets();
    let source = open_bytes(bytes.clone(), &OpenOptions::new());
    let eager = open_eager(bytes, &OpenOptions::new());
    assert_workbook_trait_surface(&source);
    assert_workbook_trait_surface(&eager);
    // The facade's text projection: tab-separated rows, sheet after sheet.
    assert_eq!(source.text().unwrap(), "7\t\n\tsecond\n9\n");
    assert_eq!(eager.text().unwrap(), source.text().unwrap());
    assert!(format!("{source:?}").contains("source_backed: true"));
}

#[cfg(any(unix, windows))]
#[test]
fn path_route_matches_the_trait_surface() {
    let file = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(file.path(), two_sheets()).unwrap();
    assert_workbook_trait_surface(&open_path(file.path(), &OpenOptions::new()));
}

// --- Selector semantics (unified_sheet_selectors.rs) ---------------------

fn selector_workbook(first_value: &str, second_sheet: &[u8]) -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/chartsheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml"/></Types>"#;
    let root = format!(
        r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="{RELATIONSHIPS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
    );
    let workbook = format!(
        r#"<workbook xmlns="{SPREADSHEETML}" xmlns:r="{RELATIONSHIPS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Broken" sheetId="2" r:id="rId2"/><sheet name="Chart" sheetId="3" r:id="rId3"/></sheets></workbook>"#
    );
    let relationships = format!(
        r#"<Relationships xmlns="{PACKAGE_RELATIONSHIPS}"><Relationship Id="rId1" Type="{RELATIONSHIPS}/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="{RELATIONSHIPS}/worksheet" Target="worksheets/sheet2.xml"/><Relationship Id="rId3" Type="{RELATIONSHIPS}/chartsheet" Target="chartsheets/sheet3.xml"/></Relationships>"#
    );
    let first = format!(
        r#"<worksheet xmlns="{SPREADSHEETML}"><sheetData><row r="1"><c r="A1"><v>{first_value}</v></c><c r="B1"/><c r="C1"><f>A1*2</f><v>14</v></c><c r="D1" t="future"><v>opaque</v></c><c r="E1" t="inlineStr"><is><t>anchor</t></is></c></row><row r="3"><c r="F3" t="inlineStr"><is><t>tail</t></is></c></row></sheetData><mergeCells count="1"><mergeCell ref="E1:F1"/></mergeCells></worksheet>"#
    );
    let chart = format!(r#"<chartsheet xmlns="{SPREADSHEETML}"/>"#);
    zip(&[
        ("[Content_Types].xml", content_types.as_bytes()),
        ("_rels/.rels", root.as_bytes()),
        ("xl/workbook.xml", workbook.as_bytes()),
        ("xl/_rels/workbook.xml.rels", relationships.as_bytes()),
        ("xl/worksheets/sheet1.xml", first.as_bytes()),
        ("xl/worksheets/sheet2.xml", second_sheet),
        ("xl/chartsheets/sheet3.xml", chart.as_bytes()),
    ])
}

const VALID_SECOND: &[u8] = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>9</v></c></row></sheetData></worksheet>"#;

fn xlsx_leaf(error: &Error) -> &litchi_xlsx::Error {
    match error.leaf() {
        Source::Xlsx(error) => error,
        other => panic!("expected an XLSX leaf error, got {other:?}"),
    }
}

#[test]
fn selector_matches_names_case_insensitively_on_both_owner_states() {
    let bytes = selector_workbook("7", VALID_SECOND);
    for handle in [
        open_bytes(bytes.clone(), &OpenOptions::new()),
        open_eager(bytes.clone(), &OpenOptions::new()),
    ] {
        let selected = handle
            .sheet("dAtA")
            .expect("select by case-folded name")
            .expect("Data sheet exists");
        let by_position = handle.sheet(0_usize).unwrap().expect("first sheet");
        assert_eq!(selected.name(), "Data");
        assert_eq!(selected.position(), 0);
        assert_eq!(by_position.name(), selected.name());
        assert!(format!("{selected:?}").contains("Data"));
        assert!(handle.sheet("missing").unwrap().is_none());
        assert!(handle.sheet(3_usize).unwrap().is_none());
        // The generic trait keeps the same case-insensitive match.
        assert_eq!(
            WorkbookTrait::worksheet_by_name(&handle, "DATA")
                .unwrap()
                .name(),
            "Data"
        );

        let a1 = selected.cell("A1").expect("read A1");
        assert_eq!(a1, selected.cell((0_u32, 0_u32)).unwrap());
        assert!(matches!(
            a1,
            SelectedCellView::Stored(Cell::Value(Value::Number(ref value))) if value.as_str() == "7"
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

        let cells = selected.cells("A1:F3").expect("sparse range");
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

        // A coordinate or range the caller passed is the caller's error
        // (M0a S9), with the XLSX coordinate error as the typed leaf: a
        // reference or bounds that name no cells are an invalid argument,
        // indices or bounds outside the grid are out of range.
        let coordinate = selected.cell("A0").unwrap_err();
        assert!(matches!(
            xlsx_leaf(&coordinate),
            litchi_xlsx::Error::Coordinate(_)
        ));
        assert_eq!(coordinate.code(), Code::InvalidArgument);
        assert_eq!(coordinate.format(), Some(litchi_core::FileFormat::Xlsx));
        let range = selected.cells((1_u32, 1_u32, 1_u32, 2_u32)).unwrap_err();
        assert!(matches!(xlsx_leaf(&range), litchi_xlsx::Error::Range(_)));
        assert_eq!(range.code(), Code::InvalidArgument);
        assert_eq!(
            selected.cells("A1:Z").unwrap_err().code(),
            Code::InvalidArgument
        );
        let outside = selected.cell((1_048_576_u32, 0_u32)).unwrap_err();
        assert!(matches!(
            xlsx_leaf(&outside),
            litchi_xlsx::Error::Coordinate(_)
        ));
        assert_eq!(outside.code(), Code::OutOfRange);
        let outside = selected
            .cells((0_u32, 0_u32, 1_048_577_u32, 1_u32))
            .unwrap_err();
        assert!(matches!(xlsx_leaf(&outside), litchi_xlsx::Error::Range(_)));
        assert_eq!(outside.code(), Code::OutOfRange);
    }
}

/// The generic sheet traits classify a caller's coordinate the same way,
/// while a coordinate read from a malformed worksheet stays a corrupt file.
#[test]
fn sheet_trait_coordinates_are_argument_errors_and_payload_coordinates_corruption() {
    let handle = open_bytes(selector_workbook("7", VALID_SECOND), &OpenOptions::new());
    let sheet = WorkbookTrait::worksheet_by_index(&handle, 0).unwrap();
    let code = |error: Box<dyn std::error::Error + Send + Sync>| {
        error
            .downcast_ref::<Error>()
            .expect("a unified error")
            .code()
    };
    assert_eq!(
        code(sheet.cell_by_coordinate("A0").map(drop).unwrap_err()),
        Code::InvalidArgument
    );
    assert_eq!(
        code(sheet.cell(1_048_577, 1).map(drop).unwrap_err()),
        Code::OutOfRange
    );
    assert_eq!(
        code(sheet.cell_value(1, 16_385).map(drop).unwrap_err()),
        Code::OutOfRange
    );
    assert_eq!(
        code(sheet.row(1_048_576).map(drop).unwrap_err()),
        Code::OutOfRange
    );
    assert_eq!(
        code(sheet.row(usize::MAX).map(drop).unwrap_err()),
        Code::OutOfRange
    );

    // A cell reference the worksheet payload itself names out of the grid.
    let malformed = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A0"><v>1</v></c></row></sheetData></worksheet>"#;
    let handle = open_bytes(selector_workbook("7", malformed), &OpenOptions::new());
    let broken = handle.sheet("Broken").unwrap().unwrap();
    let error = broken.cell("A1").unwrap_err();
    assert_eq!(error.code(), Code::CorruptFile, "{error:?}");
}

#[test]
fn selected_access_defers_an_unselected_malformed_worksheet() {
    let malformed = br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1">"#;
    let handle = open_bytes(selector_workbook("7", malformed), &OpenOptions::new());
    let data = handle.sheet("Data").unwrap().unwrap();
    assert!(matches!(
        data.cell("A1").unwrap(),
        SelectedCellView::Stored(Cell::Value(Value::Number(_)))
    ));
    let broken = handle
        .sheet("Broken")
        .expect("selecting stays catalog-only")
        .expect("Broken exists");
    assert_eq!(broken.cell("A1").unwrap_err().code(), Code::CorruptFile);
    // Text reads every sheet, so it reports the malformed payload.
    assert_eq!(handle.text().unwrap_err().code(), Code::CorruptFile);
}

#[test]
fn a_non_grid_sheet_reports_a_typed_not_worksheet_error() {
    let handle = open_bytes(selector_workbook("7", VALID_SECOND), &OpenOptions::new());
    let chart = handle.sheet("Chart").unwrap().unwrap();
    let error = chart.cell("A1").unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedType);
    assert!(matches!(
        xlsx_leaf(&error),
        litchi_xlsx::Error::NotWorksheet { sheet } if sheet == "Chart"
    ));
    // The generic trait models a non-grid sheet as empty, as the facade did.
    let generic = WorkbookTrait::worksheet_by_name(&handle, "Chart").unwrap();
    assert_eq!(generic.dimensions(), None);
    assert!(generic.rows().next().is_none());
}

// --- Source freshness ----------------------------------------------------

#[cfg(any(unix, windows))]
#[test]
fn a_changed_source_is_a_typed_io_error_on_every_read() {
    let file = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(file.path(), selector_workbook("7", VALID_SECOND)).unwrap();
    let handle = open_path(file.path(), &OpenOptions::new());
    let selected = handle.sheet("Data").unwrap().unwrap();
    assert_eq!(handle.text().unwrap(), handle.clone().text().unwrap());

    std::fs::write(
        file.path(),
        selector_workbook("12345678901234567890", VALID_SECOND),
    )
    .unwrap();

    let changed = |error: Error| {
        assert_eq!(error.code(), Code::IoError, "{error}");
        assert_eq!(error.format(), Some(litchi_core::FileFormat::Xlsx));
    };
    changed(selected.cell("A1").unwrap_err());
    changed(handle.sheet("Data").unwrap_err());
    changed(handle.text().unwrap_err());
    changed(handle.metadata().unwrap_err());
    changed(handle.tolerance_report().unwrap_err());
    changed(handle.worksheet_names().unwrap_err());
    changed(handle.worksheet_count().unwrap_err());
    let generic = WorkbookTrait::worksheet_by_index(&handle, 0)
        .unwrap()
        .cell_value(1, 1)
        .unwrap_err();
    changed(*generic.downcast::<Error>().expect("typed trait error"));
    changed(
        handle
            .into_editor()
            .err()
            .expect("editor refuses a changed source"),
    );
}

// --- Editor construction (L-XLSX-1) --------------------------------------

#[cfg(any(unix, windows))]
#[test]
fn sole_owner_editor_reuses_the_package_indexed_at_open() {
    let handle = open_path(
        &fixture_path("xlsx/DateFormatTests.xlsx"),
        &OpenOptions::new(),
    );
    let XlsxOwner::Source(workbook) = &handle.0.owner else {
        panic!("path route is source-backed");
    };
    let lineage = workbook.source_backed_package().source_lineage();
    let Ok(XlsxEditor::Source(package)) = handle.into_editor() else {
        panic!("a sole owner hands over its package");
    };
    assert!(
        package.source_lineage() == lineage,
        "the editor adopts the package indexed at open, not a re-index"
    );
    let editor = litchi_xlsx::tab_state::SourceBackedEditor::from_source_backed_package(*package)
        .expect("a per-feature source editor adopts the package");
    editor.snapshot().expect("tab-state snapshot");
}

#[test]
fn shared_handles_materialize_an_eager_editor_and_stay_readable() {
    let bytes = fixture("xlsx/DateFormatTests.xlsx");
    let handle = open_bytes(bytes, &OpenOptions::new());
    let clone = handle.clone();
    let Ok(XlsxEditor::Eager(workbook)) = handle.into_editor() else {
        panic!("a shared handle materializes the owning snapshot");
    };
    workbook.edit().expect("eager edit transaction");
    assert_eq!(clone.worksheet_names().unwrap(), ["Flags", "Tests"]);

    // A selected worksheet also shares the source catalog.
    let selected = clone.sheet(0_usize).unwrap().unwrap();
    let Ok(XlsxEditor::Eager(_)) = clone.into_editor() else {
        panic!("a selected worksheet keeps the catalog shared");
    };
    assert_eq!(selected.name(), "Flags");
}

#[test]
fn eager_owner_editor_is_the_snapshot_and_retention_can_be_refused() {
    let bytes = fixture("xlsx/InlineStrings.xlsx");
    let Ok(XlsxEditor::Eager(workbook)) =
        open_eager(bytes.clone(), &OpenOptions::new()).into_editor()
    else {
        panic!("an eager handle hands over its snapshot");
    };
    workbook.edit().expect("eager edit transaction");

    for handle in [
        open_bytes(bytes.clone(), &OpenOptions::new().retain_source(false)),
        open_eager(bytes.clone(), &OpenOptions::new().retain_source(false)),
    ] {
        assert_eq!(handle.text().unwrap().len(), 169, "reads still work");
        let error = handle
            .into_editor()
            .err()
            .expect("editing was not retained");
        assert_eq!(error.code(), Code::UnsupportedCapability);
        assert!(error.suggestion().is_some());
    }
}

#[test]
fn clones_share_one_state() {
    let handle = open_bytes(two_sheets(), &OpenOptions::new());
    let clone = handle.clone();
    assert!(std::sync::Arc::ptr_eq(&handle.0, &clone.0));
    assert_eq!(handle.format(), litchi_core::FileFormat::Xlsx);
    std::thread::spawn(move || clone.text().unwrap())
        .join()
        .unwrap();
}
