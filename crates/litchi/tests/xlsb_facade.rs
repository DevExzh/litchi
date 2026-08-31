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

fn all_worksheet_bytes() -> Vec<u8> {
    use xlsb::writer::{MutableWorksheet, WorkbookWriter};

    let mut first = MutableWorksheet::new("First");
    first.set_cell(0, 0, "FIRST");
    let mut tail = MutableWorksheet::new("Tail");
    tail.set_cell(0, 0, "TAIL");

    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(first);
    writer.add_worksheet(tail);

    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).unwrap();
    output.into_inner()
}

fn chart_first_tabs_bytes() -> Vec<u8> {
    use xlsb::chart::{Anchor, Chart};
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

    let mut writer = WorkbookWriter::new();
    writer
        .add_chart_sheet(MutableChartSheet::new("Sales Chart", chart))
        .unwrap();
    writer.add_worksheet(data);

    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).unwrap();
    output.into_inner()
}

fn rewrite_active_book_view(bytes: &[u8], active_catalog_position: u32) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let original = package.get_part(&workbook_uri).unwrap().blob().to_vec();
    let mut rewritten = Vec::new();
    let mut book_view_count = 0;
    {
        let mut writer = xlsb::raw::Writer::new(&mut rewritten);
        for record in xlsb::raw::Records::new(&original) {
            let record = record.unwrap();
            let mut payload = record.payload().to_vec();
            if record.kind() == xlsb::raw::kind::BOOK_VIEW {
                assert_eq!(payload.len(), 29, "unexpected BrtBookView payload length");
                payload[24..28].copy_from_slice(&active_catalog_position.to_le_bytes());
                book_view_count += 1;
            }
            writer.write_record(record.kind(), &payload).unwrap();
        }
    }
    assert_eq!(
        book_view_count, 1,
        "generated workbook must have one BrtBookView"
    );
    package
        .get_part_mut(&workbook_uri)
        .unwrap()
        .set_blob(rewritten);

    let mut output = Vec::new();
    package.to_stream(&mut output).unwrap();
    output
}

fn active_book_view_catalog_position(bytes: &[u8]) -> u32 {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let workbook = package.get_part(&workbook_uri).unwrap();
    let mut active_catalog_position = None;
    for record in xlsb::raw::Records::new(workbook.blob()) {
        let record = record.unwrap();
        if record.kind() != xlsb::raw::kind::BOOK_VIEW {
            continue;
        }
        assert_eq!(
            record.payload().len(),
            29,
            "unexpected BrtBookView payload length"
        );
        assert!(
            active_catalog_position
                .replace(u32::from_le_bytes(
                    record.payload()[24..28].try_into().unwrap()
                ))
                .is_none(),
            "generated workbook must have one BrtBookView"
        );
    }
    active_catalog_position.expect("generated workbook must have one BrtBookView")
}

fn save_eager_workbook(workbook: &xlsb::Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
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

fn assert_unsupported_feature<T>(result: litchi::sheet::Result<T>, expected: &str) {
    let error = match result {
        Ok(_) => panic!("unsupported XLSB operation unexpectedly succeeded"),
        Err(error) => error,
    };
    assert!(matches!(
        error.downcast_ref::<xlsb::package::error::Error>(),
        Some(xlsb::package::error::Error::UnsupportedFeature(message))
            if message.contains(expected)
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
fn dynamic_owned_bytes_facade_defers_malformed_unselected_worksheet() {
    let malformed = malformed_second_sheet();
    let workbook = litchi::sheet::open_xlsb_workbook_from_bytes_dyn(&malformed)
        .expect("dynamic XLSB bytes facade should open the catalog");
    let names = WorkbookTrait::worksheet_names(workbook.as_ref());
    assert_eq!(
        WorkbookTrait::worksheet_count(workbook.as_ref()),
        names.len()
    );
    assert!(names.len() >= 2);

    let first = WorkbookTrait::worksheet_by_index(workbook.as_ref(), 0)
        .expect("the valid first worksheet should remain selectable");
    assert!(first.cell_by_coordinate("A1").is_ok());

    let second = WorkbookTrait::worksheet_by_index(workbook.as_ref(), 1)
        .expect("the malformed worksheet catalog entry should remain selectable");
    assert!(second.cell_by_coordinate("A1").is_err());
}

#[test]
fn dynamic_owned_bytes_facade_retains_input_after_caller_mutation() {
    let mut bytes = mixed_chart_tabs_bytes();
    let workbook = litchi::sheet::open_xlsb_workbook_from_bytes_dyn(&bytes)
        .expect("dynamic XLSB bytes facade open");
    bytes.fill(0);

    let data = WorkbookTrait::worksheet_by_name(workbook.as_ref(), "Data")
        .expect("owned source should outlive the caller's input slice");
    assert_eq!(
        data.cell_by_coordinate("A1").unwrap().value(),
        &CellValue::String("DATA".to_string())
    );
}

#[test]
fn dynamic_owned_bytes_facade_enforces_exact_input_limit() {
    let bytes = all_worksheet_bytes();
    let input_bytes = u64::try_from(bytes.len()).unwrap();
    let exact_limits = xlsb::ReadLimits::builder()
        .max_input_bytes(input_bytes)
        .unwrap()
        .build()
        .unwrap();
    let workbook =
        litchi::sheet::open_xlsb_workbook_from_bytes_dyn_with_limits(&bytes, exact_limits)
            .expect("an input exactly at the configured limit should be accepted");
    assert_eq!(WorkbookTrait::worksheet_count(workbook.as_ref()), 2);

    let below_limit = xlsb::ReadLimits::builder()
        .max_input_bytes(input_bytes - 1)
        .unwrap()
        .build()
        .unwrap();
    let error = litchi::sheet::open_xlsb_workbook_from_bytes_dyn_with_limits(&bytes, below_limit)
        .expect_err("an input over the configured limit must be rejected before copying");
    assert!(error.downcast_ref::<litchi_core::Error>().is_some());
}

#[test]
fn dynamic_owned_bytes_facade_enforces_exact_part_limit() {
    let bytes = all_worksheet_bytes();
    let package = OpcPackage::from_bytes(&bytes).unwrap();
    let largest_part_bytes = package
        .iter_parts()
        .map(|part| part.blob().len())
        .max()
        .expect("generated package must contain ordinary parts");
    let exact_limits = xlsb::ReadLimits::builder()
        .max_part_bytes(u64::try_from(largest_part_bytes).unwrap())
        .unwrap()
        .build()
        .unwrap();
    litchi::sheet::open_xlsb_workbook_from_bytes_dyn_with_limits(&bytes, exact_limits)
        .expect("parts exactly at the configured limit should be accepted");

    let below_limit = xlsb::ReadLimits::builder()
        .max_part_bytes(u64::try_from(largest_part_bytes - 1).unwrap())
        .unwrap()
        .build()
        .unwrap();
    let error = litchi::sheet::open_xlsb_workbook_from_bytes_dyn_with_limits(&bytes, below_limit)
        .expect_err("a part over the configured limit must be rejected");
    assert!(error.downcast_ref::<litchi_core::Error>().is_some());
}

#[test]
fn dynamic_xlsb_byte_entrypoint_rejects_xlsx() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsx/SimpleNormal.xlsx"
    ));
    assert!(litchi::sheet::open_xlsb_workbook_from_bytes_dyn(bytes).is_err());
}

#[test]
fn dynamic_xlsb_byte_entrypoint_rejects_arbitrary_bytes() {
    assert!(litchi::sheet::open_xlsb_workbook_from_bytes_dyn(b"not an XLSB package").is_err());
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
fn source_backed_facade_strictly_refuses_sparkline_materialization() {
    let bytes = mixed_chart_tabs_with_sparkline_bytes();
    let file = fixture_file(&bytes);
    let workbook =
        litchi::sheet::open_xlsb_workbook_dyn(file.path()).expect("source-backed XLSB facade open");

    let tail =
        WorkbookTrait::worksheet_by_index(workbook.as_ref(), 1).expect("logical second worksheet");
    assert_eq!(tail.name(), "Tail");
    assert_unsupported_feature(tail.cell_by_coordinate("A1"), "sparkline groups");

    let named_tail =
        WorkbookTrait::worksheet_by_name(workbook.as_ref(), "Tail").expect("named tail worksheet");
    assert_eq!(named_tail.name(), "Tail");
    assert_unsupported_feature(named_tail.cell_by_coordinate("A1"), "sparkline groups");

    assert_unsupported_feature(
        Workbook::open(file.path())
            .expect("lazy XLSB facade open")
            .text(),
        "sparkline groups",
    );

    let bytes_dynamic = litchi::sheet::open_xlsb_workbook_from_bytes_dyn(&bytes)
        .expect("source-backed XLSB bytes facade open");
    let bytes_tail = WorkbookTrait::worksheet_by_name(bytes_dynamic.as_ref(), "Tail")
        .expect("bytes-backed tail worksheet");
    assert_unsupported_feature(bytes_tail.cell_by_coordinate("A1"), "sparkline groups");

    let unified =
        Workbook::from_bytes(bytes.clone()).expect("unified source-backed XLSB facade open");
    assert_unsupported_feature(unified.text(), "sparkline groups");

    let eager = litchi::sheet::open_xlsb_workbook_from_bytes(&bytes)
        .expect("explicit eager XLSB API should support sparkline-bearing sheets");
    assert_eq!(eager_text(&eager), "DATA\nTAIL\n");
    assert_eq!(
        WorkbookTrait::worksheet_by_name(&eager, "Tail")
            .expect("eager tail worksheet")
            .cell_by_coordinate("A1")
            .expect("eager tail cell")
            .value(),
        &CellValue::String("TAIL".to_string())
    );
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

#[test]
fn eager_and_dynamic_facade_select_active_tail_in_all_worksheet_catalog() {
    let bytes = rewrite_active_book_view(&all_worksheet_bytes(), 1);

    let eager = xlsb::Workbook::new(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&eager), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(&eager).unwrap().name(),
        "Tail"
    );

    let dynamic = litchi::sheet::open_xlsb_workbook_from_bytes_dyn(&bytes).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(dynamic.as_ref()), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(dynamic.as_ref())
            .unwrap()
            .name(),
        "Tail"
    );
}

#[test]
fn eager_and_dynamic_facade_map_active_tail_catalog_position_to_worksheet_ordinal() {
    let bytes = rewrite_active_book_view(&mixed_chart_tabs_bytes(), 2);

    let eager = xlsb::Workbook::new(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&eager), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(&eager).unwrap().name(),
        "Tail"
    );

    let dynamic = litchi::sheet::open_xlsb_workbook_from_bytes_dyn(&bytes).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(dynamic.as_ref()), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(dynamic.as_ref())
            .unwrap()
            .name(),
        "Tail"
    );

    let reopened = xlsb::Workbook::new(Cursor::new(save_eager_workbook(&eager))).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&reopened), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(&reopened).unwrap().name(),
        "Tail"
    );
}

#[test]
fn eager_and_dynamic_facade_reject_active_chart_as_active_worksheet() {
    let bytes = rewrite_active_book_view(&mixed_chart_tabs_bytes(), 1);

    let eager = xlsb::Workbook::new(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&eager), 0);
    let eager_error = WorkbookTrait::active_worksheet(&eager)
        .err()
        .expect("a chart tab cannot be returned as an active worksheet");
    assert!(
        matches!(
            eager_error
                .downcast_ref::<Box<xlsb::package::error::Error>>()
                .map(Box::as_ref),
            Some(xlsb::package::error::Error::UnsupportedFeature(message))
                if message == "XLSB active sheet is not a worksheet"
        ),
        "active chart failure must retain the typed XLSB package error"
    );

    let dynamic = litchi::sheet::open_xlsb_workbook_from_bytes_dyn(&bytes).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(dynamic.as_ref()), 0);
    let dynamic_error = WorkbookTrait::active_worksheet(dynamic.as_ref())
        .err()
        .expect("a chart tab cannot be returned as an active worksheet");
    assert!(
        matches!(
            dynamic_error
                .downcast_ref::<xlsb::package::error::Error>(),
            Some(xlsb::package::error::Error::UnsupportedFeature(message))
                if message == "XLSB active sheet is not a worksheet"
        ),
        "active chart failure must retain the typed XLSB package error"
    );
}

#[test]
fn eager_xlsb_rejects_out_of_range_active_catalog_position() {
    for active_catalog_position in [3, u32::MAX] {
        let bytes = rewrite_active_book_view(&mixed_chart_tabs_bytes(), active_catalog_position);
        let eager_error = xlsb::Workbook::new(Cursor::new(bytes.clone()))
            .expect_err("out-of-range active catalog position must be rejected");
        assert!(
            matches!(eager_error, xlsb::package::error::Error::InvalidFormula(message) if message.contains("BrtBookView sheet position")),
            "out-of-range active catalog position must retain a typed XLSB error"
        );

        let dynamic_error = litchi::sheet::open_xlsb_workbook_from_bytes_dyn(&bytes)
            .expect_err("dynamic facade must reject out-of-range active catalog position");
        assert!(
            dynamic_error.downcast_ref::<litchi_core::Error>().is_some(),
            "dynamic facade must retain a typed core error"
        );

        let file = fixture_file(&bytes);
        let path_error = litchi::sheet::open_xlsb_workbook_dyn(file.path())
            .expect_err("path facade must reject out-of-range active catalog position");
        assert!(
            matches!(
                path_error.downcast_ref::<xlsb::package::error::Error>(),
                Some(xlsb::package::error::Error::InvalidFormula(message))
                    if message.contains("BrtBookView sheet position")
            ),
            "path facade must retain the typed XLSB error"
        );
    }
}

#[test]
fn xlsb_writer_default_active_tab_is_deterministic() {
    let bytes = mixed_chart_tabs_bytes();
    let eager = xlsb::Workbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&eager), 0);
    assert_eq!(
        WorkbookTrait::active_worksheet(&eager).unwrap().name(),
        "Data"
    );

    let reopened = xlsb::Workbook::new(Cursor::new(save_eager_workbook(&eager))).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&reopened), 0);
    assert_eq!(
        WorkbookTrait::active_worksheet(&reopened).unwrap().name(),
        "Data"
    );
}

#[test]
fn path_facade_selects_active_tail_after_chart_tab_and_observes_source_changes() {
    let bytes = rewrite_active_book_view(&mixed_chart_tabs_bytes(), 2);
    let file = fixture_file(&bytes);
    let workbook = litchi::sheet::open_xlsb_workbook_dyn(file.path()).unwrap();

    assert_eq!(WorkbookTrait::active_sheet_index(workbook.as_ref()), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(workbook.as_ref())
            .unwrap()
            .name(),
        "Tail"
    );

    let mut changed = std::fs::OpenOptions::new()
        .append(true)
        .open(file.path())
        .unwrap();
    changed.write_all(b"source mutation").unwrap();
    changed.flush().unwrap();

    assert_source_changed(WorkbookTrait::active_worksheet(workbook.as_ref()));
}

#[test]
fn path_facade_reports_typed_error_for_active_chart_and_source_change_takes_precedence() {
    let bytes = rewrite_active_book_view(&mixed_chart_tabs_bytes(), 1);
    let file = fixture_file(&bytes);
    let workbook = litchi::sheet::open_xlsb_workbook_dyn(file.path()).unwrap();

    assert_eq!(WorkbookTrait::active_sheet_index(workbook.as_ref()), 0);
    let error = WorkbookTrait::active_worksheet(workbook.as_ref())
        .err()
        .expect("a chart tab cannot be returned as an active worksheet");
    assert!(
        matches!(
            error.downcast_ref::<xlsb::package::error::Error>(),
            Some(xlsb::package::error::Error::UnsupportedFeature(message))
                if message == "XLSB active sheet is not a worksheet"
        ),
        "path facade active-chart failure must retain the typed XLSB error"
    );

    let mut changed = std::fs::OpenOptions::new()
        .append(true)
        .open(file.path())
        .unwrap();
    changed.write_all(b"source mutation").unwrap();
    changed.flush().unwrap();

    assert_source_changed(WorkbookTrait::active_worksheet(workbook.as_ref()));
}

#[test]
fn chart_first_writer_defaults_active_catalog_to_first_worksheet() {
    let bytes = chart_first_tabs_bytes();
    assert_eq!(active_book_view_catalog_position(&bytes), 1);

    let eager = xlsb::Workbook::new(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&eager), 0);
    assert_eq!(
        WorkbookTrait::active_worksheet(&eager).unwrap().name(),
        "Data"
    );

    let dynamic = litchi::sheet::open_xlsb_workbook_from_bytes_dyn(&bytes).unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(dynamic.as_ref()), 0);
    assert_eq!(
        WorkbookTrait::active_worksheet(dynamic.as_ref())
            .unwrap()
            .name(),
        "Data"
    );
}
