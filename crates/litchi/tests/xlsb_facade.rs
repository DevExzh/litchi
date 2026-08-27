#![cfg(feature = "xlsb")]

use std::io::{Cursor, Write};

use litchi::opc::{OpcPackage, PackURI};
use litchi::sheet::{Workbook, WorkbookTrait};
use litchi::xlsb;
use litchi_core::sheet::{CellValue, Worksheet as WorksheetTrait};

fn fixture_bytes() -> Vec<u8> {
    std::fs::read(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsb/Simple.xlsb"
    ))
    .unwrap()
}

fn fixture_file(bytes: &[u8]) -> tempfile::NamedTempFile {
    let file = tempfile::NamedTempFile::new().unwrap();
    std::fs::write(file.path(), bytes).unwrap();
    file
}

fn mixed_chart_tabs_bytes() -> Vec<u8> {
    mixed_chart_tabs_bytes_with_sparkline(false)
}

fn mixed_chart_tabs_with_sparkline_bytes() -> Vec<u8> {
    mixed_chart_tabs_bytes_with_sparkline(true)
}

fn mixed_chart_tabs_bytes_with_sparkline(with_sparkline: bool) -> Vec<u8> {
    use xlsb::chart::{Anchor, Chart};
    use xlsb::sparkline::{Color as SparkColor, Colors, Group, Groups, Location, Sparkline};
    use xlsb::writer::{MutableChartSheet, MutableWorksheet, WorkbookWriter};

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .unwrap();
    let mut data = MutableWorksheet::new("Data");
    data.set_cell(0, 0, "DATA");
    let mut tail = MutableWorksheet::new("Tail");
    tail.set_cell(0, 0, "TAIL");
    if with_sparkline {
        let groups = Groups::new(vec![
            Group::new(
                xlsb::sparkline::SparklineType::Stacked,
                Colors::uniform(SparkColor::rgb(20, 40, 60, 255, 0)),
                vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
            )
            .unwrap(),
        ])
        .unwrap();
        tail.set_sparkline_groups(groups).unwrap();
    }

    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(data);
    writer
        .add_chart_sheet(MutableChartSheet::new("Sales Chart", chart))
        .unwrap();
    writer.add_worksheet(tail);

    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).unwrap();
    output.into_inner()
}

fn malformed_second_sheet() -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(fixture_bytes())).unwrap();
    let sheet = PackURI::new("/xl/worksheets/sheet2.bin").unwrap();
    package
        .get_part_mut(&sheet)
        .unwrap()
        .set_blob(vec![0xff; 8]);
    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn assert_source_changed<T>(result: litchi::sheet::Result<T>) {
    let error = match result {
        Ok(_) => panic!("stale source operation unexpectedly succeeded"),
        Err(error) => error,
    };
    assert!(matches!(
        error.downcast_ref::<litchi_core::Error>(),
        Some(litchi_core::Error::SourceChanged { .. })
    ));
}

fn append_eager_cell_text(output: &mut String, value: &CellValue) {
    match value {
        CellValue::Empty => {},
        CellValue::Bool(value) => output.push_str(if *value { "TRUE" } else { "FALSE" }),
        CellValue::Int(value) => output.push_str(&value.to_string()),
        CellValue::Float(value) | CellValue::DateTime(value) => output.push_str(&value.to_string()),
        CellValue::String(value) | CellValue::Error(value) => output.push_str(value),
        CellValue::Formula {
            formula,
            cached_value,
            ..
        } => match cached_value.as_deref() {
            Some(value) if !matches!(value, CellValue::Empty) => {
                append_eager_cell_text(output, value)
            },
            _ => {
                output.push('=');
                output.push_str(formula);
            },
        },
    }
}

fn eager_text(workbook: &xlsb::Workbook) -> String {
    let mut output = String::new();
    for index in 0..WorkbookTrait::worksheet_count(workbook) {
        let worksheet = WorkbookTrait::worksheet_by_index(workbook, index).unwrap();
        let mut rows = WorksheetTrait::rows(worksheet.as_ref());
        while let Some(row) = rows.next() {
            let row = row.unwrap();
            for (column, cell) in row.iter().enumerate() {
                if column != 0 {
                    output.push('\t');
                }
                append_eager_cell_text(&mut output, cell);
            }
            output.push('\n');
        }
    }
    output
}

#[test]
fn path_facade_defers_malformed_unselected_worksheet() {
    let malformed = malformed_second_sheet();
    let file = fixture_file(&malformed);
    let workbook = Workbook::open(file.path()).expect("lazy XLSB facade open");

    let names = workbook
        .worksheet_names()
        .expect("catalog names should not read worksheet payloads");
    assert_eq!(workbook.worksheet_count().unwrap(), names.len());
    assert!(names.len() >= 2);

    // The first worksheet remains a valid readable source; the malformed
    // second stream is only observed when the facade walks into it.
    let valid = Workbook::open(fixture_file(&fixture_bytes()).path())
        .expect("valid XLSB facade open")
        .text()
        .expect("valid first worksheet should be readable");
    assert!(!valid.is_empty());
    assert!(workbook.text().is_err());
}

#[test]
fn owned_bytes_facade_defers_malformed_unselected_worksheet() {
    let valid = Workbook::from_bytes(fixture_bytes()).expect("valid XLSB bytes open");
    assert!(
        !valid
            .text()
            .expect("valid first worksheet should be readable")
            .is_empty()
    );

    let malformed = Workbook::from_bytes(malformed_second_sheet())
        .expect("owned bytes catalog should defer worksheet parsing");
    let names = malformed
        .worksheet_names()
        .expect("owned bytes catalog names");
    assert_eq!(malformed.worksheet_count().unwrap(), names.len());
    assert!(names.len() >= 2);
    assert!(malformed.text().is_err());
}

#[test]
fn path_facade_reports_typed_source_change_for_catalog_and_reads() {
    let file = fixture_file(&fixture_bytes());
    let workbook = Workbook::open(file.path()).expect("lazy XLSB facade open");
    let _ = workbook
        .text()
        .expect("materialize the retained worksheet source");

    let mut changed = std::fs::OpenOptions::new()
        .append(true)
        .open(file.path())
        .unwrap();
    changed.write_all(b"source mutation").unwrap();
    changed.flush().unwrap();

    assert_source_changed(workbook.worksheet_names());
    assert_source_changed(workbook.worksheet_count());
    assert_source_changed(workbook.text());
}

#[test]
fn lazy_facade_catalog_matches_eager_xlsb_tabs_and_date_system() {
    let bytes = fixture_bytes();
    let file = fixture_file(&bytes);
    let lazy = Workbook::open(file.path()).expect("lazy XLSB facade open");
    let eager = xlsb::Workbook::new(Cursor::new(bytes)).expect("eager XLSB open");
    let eager_dyn = litchi::sheet::open_xlsb_workbook_dyn(file.path())
        .expect("eager XLSB facade entrypoint open");

    assert_eq!(lazy.worksheet_names().unwrap(), eager.worksheet_names());
    assert_eq!(lazy.worksheet_count().unwrap(), eager.worksheet_count());
    assert_eq!(eager_dyn.is_1904_date_system(), eager.is_1904_date_system());
}

#[test]
fn source_backed_facade_text_matches_eager_workbook() {
    let bytes = fixture_bytes();
    let file = fixture_file(&bytes);
    let source_backed = Workbook::open(file.path()).expect("lazy XLSB facade open");
    let eager = xlsb::Workbook::new(Cursor::new(bytes)).expect("eager XLSB open");

    assert_eq!(source_backed.text().unwrap(), eager_text(&eager));
}

#[test]
fn source_backed_facade_skips_chart_tabs_when_selecting_worksheets() {
    let bytes = mixed_chart_tabs_bytes();
    let file = fixture_file(&bytes);
    let workbook =
        litchi::sheet::open_xlsb_workbook_dyn(file.path()).expect("lazy XLSB facade open");

    assert_eq!(WorkbookTrait::worksheet_count(workbook.as_ref()), 2);
    assert_eq!(
        WorkbookTrait::worksheet_names(workbook.as_ref()),
        &["Data".to_string(), "Tail".to_string()]
    );

    let first = WorkbookTrait::worksheet_by_index(workbook.as_ref(), 0).expect("first worksheet");
    assert_eq!(first.name(), "Data");
    assert_eq!(
        first
            .cell_by_coordinate("A1")
            .expect("first worksheet cell")
            .value(),
        &CellValue::String("DATA".to_string())
    );

    let tail = WorkbookTrait::worksheet_by_index(workbook.as_ref(), 1).expect("tail worksheet");
    assert_eq!(tail.name(), "Tail");
    assert_eq!(
        tail.cell_by_coordinate("A1")
            .expect("tail worksheet cell")
            .value(),
        &CellValue::String("TAIL".to_string())
    );
    assert!(WorkbookTrait::worksheet_by_index(workbook.as_ref(), 2).is_err());

    let named_tail =
        WorkbookTrait::worksheet_by_name(workbook.as_ref(), "Tail").expect("named tail worksheet");
    assert_eq!(named_tail.name(), "Tail");
    assert!(WorkbookTrait::worksheet_by_name(workbook.as_ref(), "Sales Chart").is_err());

    let mut worksheets = WorkbookTrait::worksheets(workbook.as_ref());
    assert_eq!(worksheets.next().unwrap().unwrap().name(), "Data");
    assert_eq!(worksheets.next().unwrap().unwrap().name(), "Tail");
    assert!(worksheets.next().is_none());
}

#[test]
fn source_backed_facade_text_skips_chart_tabs() {
    let file = fixture_file(&mixed_chart_tabs_bytes());
    let workbook = Workbook::open(file.path()).expect("lazy XLSB facade open");

    assert_eq!(
        workbook.text().expect("source-backed XLSB text"),
        "DATA\nTAIL\n"
    );
}

#[test]
fn source_backed_facade_eager_fallback_keeps_tail_after_chart_tab() {
    let bytes = mixed_chart_tabs_with_sparkline_bytes();
    let file = fixture_file(&bytes);
    let workbook =
        litchi::sheet::open_xlsb_workbook_dyn(file.path()).expect("source-backed XLSB facade open");

    let tail =
        WorkbookTrait::worksheet_by_index(workbook.as_ref(), 1).expect("logical second worksheet");
    assert_eq!(tail.name(), "Tail");
    assert_eq!(
        tail.cell_by_coordinate("A1")
            .expect("tail worksheet cell")
            .value(),
        &CellValue::String("TAIL".to_string())
    );

    let named_tail =
        WorkbookTrait::worksheet_by_name(workbook.as_ref(), "Tail").expect("named tail worksheet");
    assert_eq!(named_tail.name(), "Tail");
    assert_eq!(
        named_tail
            .cell_by_coordinate("A1")
            .expect("named tail worksheet cell")
            .value(),
        &CellValue::String("TAIL".to_string())
    );

    let text = Workbook::open(file.path())
        .expect("lazy XLSB facade open")
        .text()
        .expect("eager fallback XLSB text");
    assert_eq!(text, "DATA\nTAIL\n");
}

#[test]
fn eager_xlsb_byte_entrypoint_rejects_malformed_sheet_while_lazy_facade_opens() {
    let malformed = malformed_second_sheet();
    let file = fixture_file(&malformed);
    let lazy = Workbook::open(file.path()).expect("lazy facade should open catalog");
    assert!(lazy.worksheet_names().is_ok());

    assert!(xlsb::Workbook::new(Cursor::new(malformed.clone())).is_err());
    assert!(litchi::sheet::open_xlsb_workbook_from_bytes(&malformed).is_err());
}
