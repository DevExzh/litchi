//! XLSB through the unified API: `litchi::open` / `litchi::open_bytes` and
//! the workbook view's static `WorkbookTrait`, against the eager leaf
//! `litchi::xlsb::Workbook` as the oracle.
//!
//! M0a step S8c ported these tests from the facade's workbook type and its
//! XLSB-specific workbook openers: the facade's source-backed path and
//! byte routes are `litchi::open` and `litchi::open_bytes`, its
//! `Box<dyn WorkbookTrait>` is the `WorkbookView` of `File::workbook`, its
//! eager byte entry is `SourceMode::Eager` (and the leaf it wrapped), and its
//! `*_with_limits` forms are `OpenOptions::ooxml_limits`. Errors that crossed
//! the trait boxed `litchi_core::Error` or the XLSB package error; they now
//! box a `litchi::Error` whose typed source is that XLSB or OPC error.

#![cfg(feature = "xlsb")]
#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    clippy::panic,
    reason = "test assertions panic on failure by design"
)]

use std::io::{Cursor, Write};

use litchi::error::Source;
use litchi::opc::{OpcError, OpcPackage, PackURI, SourceLimitKind};
use litchi::sheet::WorkbookTrait;
use litchi::xlsb;
use litchi::{Code, File, FileFormat, OpenOptions, SourceMode};
use litchi_core::sheet::Cell as _;
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
    use xlsb::builder::chart::{Anchor, Chart};
    use xlsb::builder::{self, Builder};
    use xlsb::sparkline::{Color as SparkColor, Colors, Group, Groups, Location, Sparkline};

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .unwrap();
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "DATA");
    let mut tail = builder::Worksheet::new("Tail");
    tail.set_cell(0, 0, "TAIL");
    if with_sparkline {
        let limits = xlsb::sparkline::Limits::default();
        let groups = Groups::new(
            vec![
                Group::new(
                    xlsb::sparkline::SparklineType::Stacked,
                    Colors::uniform(SparkColor::rgb(20, 40, 60, 255, 0)),
                    vec![Sparkline::new(Location::new(0, 0).unwrap(), None)],
                    limits,
                )
                .unwrap(),
            ],
            limits,
        )
        .unwrap();
        tail.set_sparkline_groups(groups).unwrap();
    }

    let mut writer = Builder::new();
    writer.add_worksheet(data);
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .unwrap();
    writer.add_worksheet(tail);

    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn all_worksheet_bytes() -> Vec<u8> {
    use xlsb::builder::{self, Builder};

    let mut first = builder::Worksheet::new("First");
    first.set_cell(0, 0, "FIRST");
    let mut tail = builder::Worksheet::new("Tail");
    tail.set_cell(0, 0, "TAIL");

    let mut writer = Builder::new();
    writer.add_worksheet(first);
    writer.add_worksheet(tail);

    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn chart_first_tabs_bytes() -> Vec<u8> {
    use xlsb::builder::chart::{Anchor, Chart};
    use xlsb::builder::{self, Builder};

    let chart = Chart::bar_chart_with_cache(
        "Sales",
        "Data!$A$1:$A$2",
        &["North", "South"],
        "Data!$B$1:$B$2",
        &[42.0, 55.0],
        Anchor::new(0, 0, 10, 20),
    )
    .unwrap();
    let mut data = builder::Worksheet::new("Data");
    data.set_cell(0, 0, "DATA");

    let mut writer = Builder::new();
    writer
        .add_chart_sheet(builder::ChartSheet::new("Sales Chart", chart))
        .unwrap();
    writer.add_worksheet(data);

    let mut output = Cursor::new(Vec::new());
    writer.build().unwrap().write_to(&mut output).unwrap();
    output.into_inner()
}

fn rewrite_active_book_view(bytes: &[u8], active_catalog_position: u32) -> Vec<u8> {
    let mut package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let original = package.get_part(&workbook_uri).unwrap().blob().to_vec();
    let mut rewritten = Vec::new();
    let mut book_view_count = 0;
    {
        let mut writer = xlsb::raw::Encoder::new(&mut rewritten, xlsb::raw::Limits::default());
        for record in xlsb::raw::Records::new(&original, xlsb::raw::Limits::default()) {
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
    package.write_to(&mut output).unwrap();
    output
}

fn active_book_view_catalog_position(bytes: &[u8]) -> u32 {
    let package = OpcPackage::from_bytes(bytes, litchi_opc::SourceLimits::default()).unwrap();
    let workbook_uri = PackURI::new("/xl/workbook.bin").unwrap();
    let workbook = package.get_part(&workbook_uri).unwrap();
    let mut active_catalog_position = None;
    for record in xlsb::raw::Records::new(workbook.blob(), xlsb::raw::Limits::default()) {
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
    let mut package = OpcPackage::from_reader(
        Cursor::new(fixture_bytes()),
        litchi_opc::SourceLimits::default(),
    )
    .unwrap();
    let sheet = PackURI::new("/xl/worksheets/sheet2.bin").unwrap();
    package
        .get_part_mut(&sheet)
        .unwrap()
        .set_blob(vec![0xff; 8]);
    let mut output = Vec::new();
    package.write_to(&mut output).unwrap();
    output
}

/// Open the workbook at `path`: the facade's lazy, source-backed path route.
fn open(path: &std::path::Path) -> File {
    let file = litchi::open(path).expect("lazy XLSB open");
    assert_eq!(file.format(), Some(FileFormat::Xlsb));
    file
}

/// Open owned workbook bytes: the facade's source-backed byte route.
fn open_bytes(bytes: impl Into<Vec<u8>>) -> File {
    let file = litchi::open_bytes(bytes).expect("XLSB bytes open");
    assert_eq!(file.format(), Some(FileFormat::Xlsb));
    file
}

/// Open workbook bytes into the eager owner: the facade's eager byte entry.
fn open_eager(bytes: impl Into<Vec<u8>>) -> litchi::Result<File> {
    litchi::open_bytes_with(bytes, &OpenOptions::new().source_mode(SourceMode::Eager))
}

/// `error` reports that the source changed after open.
#[track_caller]
fn assert_source_changed(error: &litchi::Error) {
    assert_eq!(error.code(), Code::IoError, "{error:?}");
    let changed = matches!(
        error.leaf(),
        Source::Core(litchi_core::Error::SourceChanged { .. })
            | Source::Opc(OpcError::SourceChanged { .. })
            | Source::Xlsb(xlsb::package::error::Error::Opc(
                OpcError::SourceChanged { .. }
            ))
    );
    assert!(changed, "expected a source change, got {error:?}");
}

/// The trait result failed with a source change (the trait errors are
/// typed `litchi::Error` values).
#[track_caller]
fn assert_boxed_source_changed<T>(result: Result<T, litchi::Error>) {
    let error = match result {
        Ok(_) => panic!("stale source operation unexpectedly succeeded"),
        Err(error) => error,
    };
    assert_source_changed(&error);
}

/// `error` is the XLSB package's `UnsupportedFeature` refusal, naming
/// `expected`.
#[track_caller]
fn assert_unsupported_feature(error: &litchi::Error, expected: &str) {
    assert_eq!(error.code(), Code::UnsupportedCapability, "{error:?}");
    assert!(
        matches!(
            error.leaf(),
            Source::Xlsb(xlsb::package::error::Error::UnsupportedFeature(message))
                if message.contains(expected)
        ),
        "expected an unsupported {expected:?} feature, got {error:?}"
    );
}

/// The trait result failed with the XLSB `UnsupportedFeature` refusal.
#[track_caller]
fn assert_boxed_unsupported_feature<T>(result: Result<T, litchi::Error>, expected: &str) {
    let error = match result {
        Ok(_) => panic!("unsupported XLSB operation unexpectedly succeeded"),
        Err(error) => error,
    };
    assert_unsupported_feature(&error, expected);
}

/// `error` is the OPC read-limit refusal of `resource`: `observed` against
/// a ceiling of `maximum`, or any observed value when `observed` is `None`.
#[track_caller]
fn assert_read_limit(
    error: &litchi::Error,
    resource: SourceLimitKind,
    observed: Option<u64>,
    maximum: u64,
) {
    assert_eq!(error.code(), Code::LimitExceeded, "{error:?}");
    assert!(
        matches!(
            error.leaf(),
            Source::Opc(OpcError::Limit { kind: actual_resource, observed: actual, maximum: limit })
                if *actual_resource == resource
                    && observed.is_none_or(|observed| *actual == observed)
                    && *limit == maximum
        ),
        "expected the {resource:?} limit {maximum} to refuse, got {error:?}"
    );
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
        for row in WorksheetTrait::rows(worksheet.as_ref()) {
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

fn mutate(path: &std::path::Path) {
    let mut changed = std::fs::OpenOptions::new().append(true).open(path).unwrap();
    changed.write_all(b"source mutation").unwrap();
    changed.flush().unwrap();
}

#[test]
fn path_facade_defers_malformed_unselected_worksheet() {
    let malformed = malformed_second_sheet();
    let file = fixture_file(&malformed);
    let workbook = open(file.path());

    let names = workbook
        .workbook()
        .unwrap()
        .worksheet_names()
        .expect("catalog names should not read worksheet payloads");
    assert_eq!(
        workbook.workbook().unwrap().worksheet_count().unwrap(),
        names.len()
    );
    assert!(names.len() >= 2);

    // The first worksheet remains a valid readable source; the malformed
    // second stream is only observed when the read walks into it.
    let valid = open(fixture_file(&fixture_bytes()).path())
        .text()
        .expect("valid first worksheet should be readable");
    assert!(!valid.is_empty());
    assert!(workbook.text().is_err());
}

#[test]
fn owned_bytes_facade_defers_malformed_unselected_worksheet() {
    let valid = open_bytes(fixture_bytes());
    assert!(
        !valid
            .text()
            .expect("valid first worksheet should be readable")
            .is_empty()
    );

    let malformed = open_bytes(malformed_second_sheet());
    let view = malformed.workbook().unwrap();
    let names = view.worksheet_names().expect("owned bytes catalog names");
    assert_eq!(view.worksheet_count().unwrap(), names.len());
    assert!(names.len() >= 2);
    assert!(malformed.text().is_err());
}

#[test]
fn dynamic_owned_bytes_facade_defers_malformed_unselected_worksheet() {
    let malformed = malformed_second_sheet();
    let file = open_bytes(malformed);
    let workbook = file
        .workbook()
        .expect("the workbook view should open the catalog");
    let names = WorkbookTrait::worksheet_names(&workbook);
    assert_eq!(WorkbookTrait::worksheet_count(&workbook), names.len());
    assert!(names.len() >= 2);

    let first = WorkbookTrait::worksheet_by_index(&workbook, 0)
        .expect("the valid first worksheet should remain selectable");
    assert!(first.cell_by_coordinate("A1").is_ok());

    let second = WorkbookTrait::worksheet_by_index(&workbook, 1)
        .expect("the malformed worksheet catalog entry should remain selectable");
    assert!(second.cell_by_coordinate("A1").is_err());
}

#[test]
fn dynamic_owned_bytes_facade_retains_input_after_caller_mutation() {
    let mut bytes = mixed_chart_tabs_bytes();
    // A borrowed slice is copied into the handle once.
    let file = open_bytes(bytes.as_slice());
    bytes.fill(0);

    let workbook = file.workbook().unwrap();
    let data = WorkbookTrait::worksheet_by_name(&workbook, "Data")
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
    let exact_limits = xlsb::SourceLimits::builder()
        .max_input_bytes(input_bytes)
        .unwrap()
        .build()
        .unwrap();
    let file = litchi::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().ooxml_limits(exact_limits),
    )
    .expect("an input exactly at the configured limit should be accepted");
    assert_eq!(WorkbookTrait::worksheet_count(&file.workbook().unwrap()), 2);

    let below_limit = xlsb::SourceLimits::builder()
        .max_input_bytes(input_bytes - 1)
        .unwrap()
        .build()
        .unwrap();
    let error = litchi::open_bytes_with(bytes, &OpenOptions::new().ooxml_limits(below_limit))
        .expect_err("an input over the configured limit must be rejected before copying");
    assert_read_limit(
        &error,
        SourceLimitKind::InputBytes,
        Some(input_bytes),
        input_bytes - 1,
    );
}

#[test]
fn dynamic_owned_bytes_facade_enforces_exact_part_limit() {
    let bytes = all_worksheet_bytes();
    let package = OpcPackage::from_bytes(&bytes, litchi_opc::SourceLimits::default()).unwrap();
    let largest_part_bytes = package
        .try_iter_parts()
        .map(|part| part.expect("generated package parts decode").blob().len())
        .max()
        .expect("generated package must contain ordinary parts");
    let largest_part_bytes = u64::try_from(largest_part_bytes).unwrap();
    let exact_limits = xlsb::SourceLimits::builder()
        .max_part_bytes(largest_part_bytes)
        .unwrap()
        .build()
        .unwrap();
    litchi::open_bytes_with(
        bytes.clone(),
        &OpenOptions::new().ooxml_limits(exact_limits),
    )
    .expect("parts exactly at the configured limit should be accepted");

    let below_limit = xlsb::SourceLimits::builder()
        .max_part_bytes(largest_part_bytes - 1)
        .unwrap()
        .build()
        .unwrap();
    let error = litchi::open_bytes_with(bytes, &OpenOptions::new().ooxml_limits(below_limit))
        .expect_err("a part over the configured limit must be rejected");
    assert_read_limit(
        &error,
        SourceLimitKind::PartBytes,
        None,
        largest_part_bytes - 1,
    );
}

/// The facade's XLSB byte entry refused an XLSX package. Detection now
/// routes it: an XLSX never reaches the XLSB owner, and the XLSB leaf entry
/// still refuses it.
#[test]
fn dynamic_xlsb_byte_entrypoint_rejects_xlsx() {
    let bytes = include_bytes!(concat!(
        env!("CARGO_MANIFEST_DIR"),
        "/../../test-data/ooxml/xlsx/SimpleNormal.xlsx"
    ));
    #[cfg(feature = "xlsx")]
    assert!(matches!(
        litchi::open_bytes(bytes.as_slice()).unwrap(),
        File::Xlsx(_)
    ));
    #[cfg(not(feature = "xlsx"))]
    assert_eq!(
        litchi::open_bytes(bytes.as_slice()).unwrap_err().detected(),
        Some(FileFormat::Xlsx)
    );
    assert!(xlsb::Package::from_bytes(bytes).is_err());
}

#[test]
fn dynamic_xlsb_byte_entrypoint_rejects_arbitrary_bytes() {
    let error = litchi::open_bytes(b"not an XLSB package".as_slice()).unwrap_err();
    assert_eq!(error.code(), Code::UnsupportedFormat, "{error:?}");
    assert_eq!(error.detected(), None, "{error:?}");
}

#[test]
fn path_facade_reports_typed_source_change_for_catalog_and_reads() {
    let file = fixture_file(&fixture_bytes());
    let workbook = open(file.path());
    let _ = workbook
        .text()
        .expect("materialize the retained worksheet source");

    mutate(file.path());

    let view = workbook.workbook().unwrap();
    assert_source_changed(&view.worksheet_names().unwrap_err());
    assert_source_changed(&view.worksheet_count().unwrap_err());
    assert_source_changed(&workbook.text().unwrap_err());
}

#[test]
fn lazy_facade_catalog_matches_eager_xlsb_tabs_and_date_system() {
    let bytes = fixture_bytes();
    let file = fixture_file(&bytes);
    let lazy = open(file.path());
    let eager = xlsb::Workbook::new(Cursor::new(bytes.clone())).expect("eager XLSB open");
    let eager_facade = open_eager(bytes).expect("eager XLSB facade open");

    let view = lazy.workbook().unwrap();
    assert_eq!(view.worksheet_names().unwrap(), eager.worksheet_names());
    assert_eq!(view.worksheet_count().unwrap(), eager.worksheet_count());
    assert_eq!(
        WorkbookTrait::is_1904_date_system(&view),
        eager.is_1904_date_system()
    );
    assert_eq!(
        WorkbookTrait::is_1904_date_system(&eager_facade.workbook().unwrap()),
        eager.is_1904_date_system()
    );
}

#[test]
fn source_backed_facade_text_matches_eager_workbook() {
    let bytes = fixture_bytes();
    let file = fixture_file(&bytes);
    let source_backed = open(file.path());
    let eager = xlsb::Workbook::new(Cursor::new(bytes)).expect("eager XLSB open");

    assert_eq!(source_backed.text().unwrap(), eager_text(&eager));
}

#[test]
fn source_backed_facade_skips_chart_tabs_when_selecting_worksheets() {
    let bytes = mixed_chart_tabs_bytes();
    let file = fixture_file(&bytes);
    let opened = open(file.path());
    let workbook = opened.workbook().expect("lazy XLSB workbook view");

    assert_eq!(WorkbookTrait::worksheet_count(&workbook), 2);
    assert_eq!(
        WorkbookTrait::worksheet_names(&workbook),
        &["Data".to_string(), "Tail".to_string()]
    );

    let first = WorkbookTrait::worksheet_by_index(&workbook, 0).expect("first worksheet");
    assert_eq!(first.name(), "Data");
    assert_eq!(
        first
            .cell_by_coordinate("A1")
            .expect("first worksheet cell")
            .value(),
        &CellValue::String("DATA".to_string())
    );

    let tail = WorkbookTrait::worksheet_by_index(&workbook, 1).expect("tail worksheet");
    assert_eq!(tail.name(), "Tail");
    assert_eq!(
        tail.cell_by_coordinate("A1")
            .expect("tail worksheet cell")
            .value(),
        &CellValue::String("TAIL".to_string())
    );
    assert!(WorkbookTrait::worksheet_by_index(&workbook, 2).is_err());

    let named_tail =
        WorkbookTrait::worksheet_by_name(&workbook, "Tail").expect("named tail worksheet");
    assert_eq!(named_tail.name(), "Tail");
    assert!(WorkbookTrait::worksheet_by_name(&workbook, "Sales Chart").is_err());

    let mut worksheets = WorkbookTrait::worksheets(&workbook);
    assert_eq!(worksheets.next().unwrap().unwrap().name(), "Data");
    assert_eq!(worksheets.next().unwrap().unwrap().name(), "Tail");
    assert!(worksheets.next().is_none());
}

#[test]
fn source_backed_facade_text_skips_chart_tabs() {
    let file = fixture_file(&mixed_chart_tabs_bytes());
    let workbook = open(file.path());

    assert_eq!(
        workbook.text().expect("source-backed XLSB text"),
        "DATA\nTAIL\n"
    );
}

#[test]
fn source_backed_facade_strictly_refuses_sparkline_materialization() {
    let bytes = mixed_chart_tabs_with_sparkline_bytes();
    let file = fixture_file(&bytes);
    let opened = open(file.path());
    let workbook = opened.workbook().expect("source-backed XLSB workbook view");

    let tail = WorkbookTrait::worksheet_by_index(&workbook, 1).expect("logical second worksheet");
    assert_eq!(tail.name(), "Tail");
    assert_boxed_unsupported_feature(tail.cell_by_coordinate("A1"), "sparkline groups");

    let named_tail =
        WorkbookTrait::worksheet_by_name(&workbook, "Tail").expect("named tail worksheet");
    assert_eq!(named_tail.name(), "Tail");
    assert_boxed_unsupported_feature(named_tail.cell_by_coordinate("A1"), "sparkline groups");

    assert_unsupported_feature(&open(file.path()).text().unwrap_err(), "sparkline groups");

    let bytes_file = open_bytes(bytes.clone());
    let bytes_workbook = bytes_file.workbook().unwrap();
    let bytes_tail = WorkbookTrait::worksheet_by_name(&bytes_workbook, "Tail")
        .expect("bytes-backed tail worksheet");
    assert_boxed_unsupported_feature(bytes_tail.cell_by_coordinate("A1"), "sparkline groups");

    assert_unsupported_feature(&bytes_file.text().unwrap_err(), "sparkline groups");

    // The explicit eager APIs support sparkline-bearing sheets: the eager
    // owner of the unified opener, and the leaf entry the facade wrapped.
    let eager_facade =
        open_eager(bytes.clone()).expect("the eager owner should support sparkline-bearing sheets");
    assert_eq!(eager_facade.text().unwrap(), "DATA\nTAIL\n");
    assert_eq!(
        WorkbookTrait::worksheet_by_name(&eager_facade.workbook().unwrap(), "Tail")
            .expect("eager facade tail worksheet")
            .cell_by_coordinate("A1")
            .expect("eager facade tail cell")
            .value(),
        &CellValue::String("TAIL".to_string())
    );
    let package = xlsb::Package::from_bytes(&bytes)
        .expect("explicit eager XLSB API should support sparkline-bearing sheets");
    let eager = package.workbook().expect("eager projection");
    assert_eq!(eager_text(eager), "DATA\nTAIL\n");
    assert_eq!(
        WorkbookTrait::worksheet_by_name(eager, "Tail")
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
    let lazy = open(file.path());
    assert!(lazy.workbook().unwrap().worksheet_names().is_ok());

    assert!(xlsb::Workbook::new(Cursor::new(malformed.clone())).is_err());
    assert!(open_eager(malformed).is_err());
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

    let file = open_bytes(bytes);
    let dynamic = file.workbook().unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&dynamic), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(&dynamic).unwrap().name(),
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

    let file = open_bytes(bytes);
    let dynamic = file.workbook().unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&dynamic), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(&dynamic).unwrap().name(),
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
        .expect_err("a chart tab cannot be returned as an active worksheet");
    assert!(
        matches!(
            &eager_error,
            xlsb::package::error::Error::UnsupportedFeature(message)
                if *message == "XLSB active sheet is not a worksheet"
        ),
        "active chart failure must retain the typed XLSB package error"
    );

    let file = open_bytes(bytes);
    let dynamic = file.workbook().unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&dynamic), 0);
    let dynamic_error = WorkbookTrait::active_worksheet(&dynamic)
        .expect_err("a chart tab cannot be returned as an active worksheet");
    assert!(
        matches!(
            dynamic_error.leaf(),
            Source::Xlsb(xlsb::package::error::Error::UnsupportedFeature(message))
                if *message == "XLSB active sheet is not a worksheet"
        ),
        "active chart failure must retain the typed XLSB package error: {dynamic_error:?}"
    );
}

#[test]
fn eager_xlsb_rejects_out_of_range_active_catalog_position() {
    #[track_caller]
    fn assert_out_of_range(error: &litchi::Error, route: &str) {
        assert!(
            matches!(
                error.leaf(),
                Source::Xlsb(xlsb::package::error::Error::InvalidFormula(message))
                    if message.contains("BrtBookView sheet position")
            ),
            "{route} must retain the typed XLSB error: {error:?}"
        );
    }

    for active_catalog_position in [3, u32::MAX] {
        let bytes = rewrite_active_book_view(&mixed_chart_tabs_bytes(), active_catalog_position);
        let eager_error = xlsb::Workbook::new(Cursor::new(bytes.clone()))
            .expect_err("out-of-range active catalog position must be rejected");
        assert!(
            matches!(eager_error, xlsb::package::error::Error::InvalidFormula(message) if message.contains("BrtBookView sheet position")),
            "out-of-range active catalog position must retain a typed XLSB error"
        );

        let dynamic_error = litchi::open_bytes(bytes.clone())
            .expect_err("the byte route must reject out-of-range active catalog position");
        assert_out_of_range(&dynamic_error, "the byte route");

        let file = fixture_file(&bytes);
        let path_error = litchi::open(file.path())
            .expect_err("the path route must reject out-of-range active catalog position");
        assert_out_of_range(&path_error, "the path route");
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
    let opened = open(file.path());
    let workbook = opened.workbook().unwrap();

    assert_eq!(WorkbookTrait::active_sheet_index(&workbook), 1);
    assert_eq!(
        WorkbookTrait::active_worksheet(&workbook).unwrap().name(),
        "Tail"
    );

    mutate(file.path());

    assert_boxed_source_changed(WorkbookTrait::active_worksheet(&workbook));
}

#[test]
fn path_facade_reports_typed_error_for_active_chart_and_source_change_takes_precedence() {
    let bytes = rewrite_active_book_view(&mixed_chart_tabs_bytes(), 1);
    let file = fixture_file(&bytes);
    let opened = open(file.path());
    let workbook = opened.workbook().unwrap();

    assert_eq!(WorkbookTrait::active_sheet_index(&workbook), 0);
    let error = WorkbookTrait::active_worksheet(&workbook)
        .expect_err("a chart tab cannot be returned as an active worksheet");
    assert!(
        matches!(
            error.leaf(),
            Source::Xlsb(xlsb::package::error::Error::UnsupportedFeature(message))
                if *message == "XLSB active sheet is not a worksheet"
        ),
        "path facade active-chart failure must retain the typed XLSB error: {error:?}"
    );

    mutate(file.path());

    assert_boxed_source_changed(WorkbookTrait::active_worksheet(&workbook));
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

    let file = open_bytes(bytes);
    let dynamic = file.workbook().unwrap();
    assert_eq!(WorkbookTrait::active_sheet_index(&dynamic), 0);
    assert_eq!(
        WorkbookTrait::active_worksheet(&dynamic).unwrap().name(),
        "Data"
    );
}
