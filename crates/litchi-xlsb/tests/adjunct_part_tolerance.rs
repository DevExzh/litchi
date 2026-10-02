#![allow(
    clippy::unwrap_used,
    reason = "test assertions use panic-on-failure extraction by design"
)]

//! Audit finding F3.11 in adjunct parts: trailing payload bytes in table,
//! PivotCache, connections, and comments parts, and reserved `Cell` flag bits
//! on the binary-index point lookup, are repaired and recorded. Each fixture
//! injects one deviation into a conformant workbook; the unmodified workbook
//! is the control and repairs nothing.

use std::io::Cursor;

use litchi_core::sheet::traits::WorkbookTrait;
use litchi_opc::{OpcPackage, PackURI};
use litchi_xlsb::package::connections::{
    CommandType, Connection, Connections, DbProperties, Properties, SourceType,
};
use litchi_xlsb::package::table::{Column, Range as TableRange, Table, Type as TableType};
use litchi_xlsb::raw::{Kind, Records, Writer, kind};
use litchi_xlsb::writer::{MutableWorksheet, WorkbookWriter};
use litchi_xlsb::{
    Defect, LayoutDefect, Package, ReservedFieldDefect, SourceBackedWorkbook, ToleranceReport,
    Workbook,
};

type Stream = Vec<(Kind, Vec<u8>)>;

fn test_data(name: &str) -> Vec<u8> {
    std::fs::read(format!(
        "{}/../../test-data/ooxml/xlsb/{name}",
        env!("CARGO_MANIFEST_DIR")
    ))
    .unwrap()
}

fn save(workbook: &Workbook) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    workbook.save(&mut output).unwrap();
    output.into_inner()
}

/// The first part whose name contains `fragment`.
fn part_name(package_bytes: &[u8], fragment: &str) -> String {
    let package = OpcPackage::from_reader(Cursor::new(package_bytes)).unwrap();
    package
        .iter_parts()
        .map(|part| part.partname().to_string())
        .find(|name| name.contains(fragment))
        .unwrap()
}

fn rewrite(package_bytes: &[u8], part: &str, transform: impl FnOnce(&mut Stream)) -> Vec<u8> {
    let mut package = OpcPackage::from_reader(Cursor::new(package_bytes)).unwrap();
    let uri = PackURI::new(part).unwrap();
    let mut stream: Stream = Records::new(package.get_part(&uri).unwrap().blob())
        .map(|record| {
            let record = record.unwrap();
            (record.kind(), record.payload().to_vec())
        })
        .collect();
    transform(&mut stream);
    let mut blob = Vec::new();
    let mut writer = Writer::new(&mut blob);
    for (record_type, payload) in &stream {
        writer.write_record(*record_type, payload).unwrap();
    }
    package.get_part_mut(&uri).unwrap().set_blob(blob);
    let mut output = Cursor::new(Vec::new());
    package.to_stream(&mut output).unwrap();
    output.into_inner()
}

fn first_payload(stream: &mut Stream, record_type: Kind) -> &mut Vec<u8> {
    &mut stream
        .iter_mut()
        .find(|(kind, _)| *kind == record_type)
        .unwrap()
        .1
}

fn assert_single(
    report: &ToleranceReport,
    defect: impl Into<Defect>,
    record_type: Kind,
    observed: u32,
) {
    assert_eq!(report.total(), 1, "{:?}", report.defects());
    let entry = report.defects()[0];
    assert_eq!(entry.defect(), defect.into());
    assert_eq!(entry.record_type(), record_type.get());
    assert_eq!(entry.observed(), observed);
}

/// A workbook with one structured table on its first sheet.
fn table_workbook() -> Vec<u8> {
    let mut data = MutableWorksheet::new("Data");
    data.set_cell(0, 0, "Region");
    data.set_cell(0, 1, "Amount");
    data.set_cell(1, 0, "North");
    data.set_cell(1, 1, 42.5);
    data.add_table(Table {
        id: 3,
        name: Some("SalesTable".to_string()),
        display_name: Some("SalesTable".to_string()),
        range: TableRange {
            first_row: 0,
            last_row: 1,
            first_column: 0,
            last_column: 1,
        },
        table_type: TableType::Range,
        header_row_count: 1,
        columns: vec![
            Column {
                id: 1,
                name: Some("Region".to_string()),
                ..Column::default()
            },
            Column {
                id: 2,
                name: Some("Amount".to_string()),
                ..Column::default()
            },
        ],
        ..Table::default()
    })
    .unwrap();
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(data);
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).unwrap();
    output.into_inner()
}

/// A workbook carrying one inert ODBC connection.
fn connections_workbook() -> Vec<u8> {
    let mut sheet = MutableWorksheet::new("Sheet1");
    sheet.set_cell(0, 0, "value");
    let mut writer = WorkbookWriter::new();
    writer.add_worksheet(sheet);
    let mut output = Cursor::new(Vec::new());
    writer.save(&mut output).unwrap();
    let mut workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    workbook
        .set_connections(Connections {
            connections: vec![Connection {
                connection_id: 42,
                source_type: SourceType::Odbc,
                name: "Sales".to_string(),
                properties: Properties::Database(DbProperties {
                    command_type: CommandType::Sql,
                    connection_string: "Driver={Generated};Server=example.invalid".to_string(),
                    command: Some("SELECT 1".to_string()),
                    server_command: None,
                }),
                ..Connection::default()
            }],
        })
        .unwrap();
    save(&workbook)
}

#[test]
fn conformant_adjunct_parts_repair_nothing() {
    for bytes in [
        table_workbook(),
        connections_workbook(),
        test_data("pivot-cache-ac-prefixed.xlsb"),
        test_data("comments.xlsb"),
    ] {
        let workbook = Workbook::new(Cursor::new(bytes.clone())).unwrap();
        assert!(workbook.tolerance_report().is_clean());
        for index in 0..workbook.worksheet_count() {
            assert!(
                workbook
                    .worksheet(index)
                    .unwrap()
                    .tolerance_report()
                    .is_clean()
            );
        }
        assert!(
            Package::from_bytes(bytes)
                .unwrap()
                .tolerance_report()
                .is_clean()
        );
    }
}

#[test]
fn trailing_table_record_bytes_are_ignored_and_recorded_once() {
    let source = table_workbook();
    let table = part_name(&source, "/tables/");
    let bytes = rewrite(&source, &table, |stream| {
        first_payload(stream, kind::BEGIN_LIST).extend_from_slice(&[0; 3]);
    });

    let workbook = Workbook::new(Cursor::new(bytes.clone())).unwrap();
    // The typed table parser and the formula-context pass read the same
    // record; the repair is still recorded once.
    assert_single(
        workbook.tolerance_report(),
        LayoutDefect::TrailingBytes,
        kind::BEGIN_LIST,
        3,
    );
    assert_eq!(workbook.structured_tables().len(), 1);
    assert_eq!(
        workbook.structured_tables()[0].1.display_name.as_deref(),
        Some("SalesTable")
    );
    assert_eq!(
        Package::from_bytes(bytes.clone())
            .unwrap()
            .tolerance_report(),
        workbook.tolerance_report()
    );

    // The source-backed catalog parses table parts lazily for formula
    // resolution, once.
    let source_backed = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
    for _ in 0..2 {
        source_backed
            .worksheet_by_index(0)
            .unwrap()
            .unwrap()
            .materialize()
            .unwrap();
    }
    assert_single(
        &source_backed.tolerance_report(),
        LayoutDefect::TrailingBytes,
        kind::BEGIN_LIST,
        3,
    );
}

#[test]
fn trailing_table_column_bytes_are_ignored_and_recorded() {
    let source = table_workbook();
    let table = part_name(&source, "/tables/");
    let bytes = rewrite(&source, &table, |stream| {
        first_payload(stream, kind::BEGIN_LIST_COL).push(0xEE);
    });
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_single(
        workbook.tolerance_report(),
        LayoutDefect::TrailingBytes,
        kind::BEGIN_LIST_COL,
        1,
    );
    let table = &workbook.structured_tables()[0].1;
    assert_eq!(table.columns[0].name.as_deref(), Some("Region"));
}

#[test]
fn trailing_pivot_cache_definition_bytes_are_ignored_and_recorded() {
    let source = test_data("pivot-cache-ac-prefixed.xlsb");
    let bytes = rewrite(
        &source,
        "/xl/pivotCache/pivotCacheDefinition1.bin",
        |stream| {
            first_payload(stream, kind::BEGIN_PIVOT_CACHE_DEF).extend_from_slice(&[0; 2]);
        },
    );
    let workbook = Workbook::new(Cursor::new(bytes.clone())).unwrap();
    assert_single(
        workbook.tolerance_report(),
        LayoutDefect::TrailingBytes,
        kind::BEGIN_PIVOT_CACHE_DEF,
        2,
    );
    let control = Workbook::new(Cursor::new(source)).unwrap();
    assert_eq!(
        format!("{:?}", workbook.pivot_cache_definitions()),
        format!("{:?}", control.pivot_cache_definitions())
    );
    assert_eq!(
        Package::from_bytes(bytes).unwrap().tolerance_report(),
        workbook.tolerance_report()
    );
}

#[test]
fn trailing_connection_bytes_are_ignored_and_recorded() {
    let source = connections_workbook();
    let bytes = rewrite(&source, "/xl/connections.bin", |stream| {
        first_payload(stream, kind::BEGIN_EXT_CONNECTION).extend_from_slice(&[0; 4]);
    });
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert_single(
        workbook.tolerance_report(),
        LayoutDefect::TrailingBytes,
        kind::BEGIN_EXT_CONNECTION,
        4,
    );
    let connection = &workbook.connections().unwrap().connections[0];
    assert_eq!(connection.name, "Sales");
    assert_eq!(connection.connection_id, 42);
}

/// Open a comments fixture eagerly and source-backed; return the report of
/// the worksheet that owns the comments.
fn comments_report(bytes: &[u8]) -> ToleranceReport {
    let workbook = Workbook::new(Cursor::new(bytes.to_vec())).unwrap();
    assert!(workbook.tolerance_report().is_clean());
    let control = Workbook::new(Cursor::new(test_data("comments.xlsb"))).unwrap();
    let mut merged = ToleranceReport::default();
    for index in 0..workbook.worksheet_count() {
        let worksheet = workbook.worksheet(index).unwrap();
        assert_eq!(
            worksheet.comments().len(),
            control.worksheet(index).unwrap().comments().len()
        );
        if !worksheet.tolerance_report().is_clean() {
            assert!(merged.is_clean(), "only one sheet owns comments");
            merged = worksheet.tolerance_report().clone();

            let source = SourceBackedWorkbook::from_reader(Cursor::new(bytes.to_vec())).unwrap();
            let materialized = source
                .worksheet_by_index(index)
                .unwrap()
                .unwrap()
                .materialize()
                .unwrap();
            assert_eq!(materialized.tolerance_report(), &merged);
        }
    }
    merged
}

#[test]
fn trailing_comment_author_bytes_are_ignored_and_recorded() {
    let source = test_data("comments.xlsb");
    let comments = part_name(&source, "comments");
    let bytes = rewrite(&source, &comments, |stream| {
        first_payload(stream, kind::COMMENT_AUTHOR).extend_from_slice(&[0; 2]);
    });
    let report = comments_report(&bytes);
    assert_single(
        &report,
        LayoutDefect::TrailingBytes,
        kind::COMMENT_AUTHOR,
        2,
    );
}

#[test]
fn trailing_comment_text_bytes_are_ignored_and_recorded() {
    let source = test_data("comments.xlsb");
    let comments = part_name(&source, "comments");
    let bytes = rewrite(&source, &comments, |stream| {
        first_payload(stream, kind::COMMENT_TEXT).push(0);
    });
    let report = comments_report(&bytes);
    assert_single(&report, LayoutDefect::TrailingBytes, kind::COMMENT_TEXT, 1);
}

#[test]
fn indexed_lookups_mask_reserved_cell_flags_and_record_each_repair() {
    let source = test_data("Simple.xlsb");
    let mut cell = None;
    let bytes = rewrite(&source, "/xl/worksheets/sheet1.bin", |stream| {
        let mut row = 0;
        for (record_type, payload) in stream.iter_mut() {
            if *record_type == kind::ROW_HDR {
                row = u32::from_le_bytes(payload[0..4].try_into().unwrap());
            } else if *record_type == kind::CELL_ISST {
                payload[7] |= 0x80;
                cell = Some((row, u32::from_le_bytes(payload[0..4].try_into().unwrap())));
                break;
            }
        }
    });
    let (row, column) = cell.unwrap();

    let control = SourceBackedWorkbook::from_reader(Cursor::new(source)).unwrap();
    let expected = control
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .indexed_values()
        .unwrap()
        .unwrap()
        .cached_value(row, column)
        .unwrap();
    assert!(expected.is_some());
    assert!(control.tolerance_report().is_clean());

    let workbook = SourceBackedWorkbook::from_reader(Cursor::new(bytes)).unwrap();
    let indexed = workbook
        .worksheet_by_index(0)
        .unwrap()
        .unwrap()
        .indexed_values()
        .unwrap()
        .unwrap();
    assert_eq!(indexed.cached_value(row, column).unwrap(), expected);
    assert_single(
        &workbook.tolerance_report(),
        ReservedFieldDefect::CellFlags,
        kind::CELL_ISST,
        0x80,
    );
    assert_eq!(workbook.tolerance_report().defects()[0].location(), row);
}
