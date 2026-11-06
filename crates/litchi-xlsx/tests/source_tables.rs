#![allow(clippy::unwrap_used, reason = "focused fixture assertions")]

//! Worksheet tables on the source-backed owner, case-insensitive lookup on
//! both owners, and the typed header, data and totals ranges of a table.

use std::io::Cursor;
use std::path::{Path, PathBuf};

use litchi_opc::{OpcError, SourceLimitKind};
use litchi_xlsx::workbook::worksheet::TablePart;
use litchi_xlsx::{Error, Rect, SourceBackedWorkbook, SourceLimits, WorksheetKind};

fn fixture(name: &str) -> PathBuf {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ooxml/xlsx")
        .join(name)
}

fn describe(tables: &[TablePart]) -> Vec<String> {
    tables
        .iter()
        .map(|part| {
            format!(
                "{}|{}|{:?}",
                part.relationship_id(),
                part.part_name(),
                part.table()
            )
        })
        .collect()
}

fn range(reference: &str) -> Rect {
    Rect::from_a1(reference).unwrap()
}

fn assert_event_limit(error: Error) {
    assert!(matches!(
        error,
        Error::Package(OpcError::Limit {
            kind: SourceLimitKind::Events,
            maximum: 40,
            observed,
        }) if observed > 40
    ));
}

#[test]
fn source_tables_equal_the_eager_owner_tables_across_the_corpus() {
    let roots = [
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/ooxml/xlsx"),
        Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test-data/poi/test-data/spreadsheet"),
    ];
    let mut tables = 0usize;
    for root in roots {
        let Ok(entries) = std::fs::read_dir(root) else {
            continue;
        };
        for entry in entries.flatten() {
            let path = entry.path();
            if path.extension().and_then(|value| value.to_str()) != Some("xlsx") {
                continue;
            }
            let (Ok(eager), Ok(source)) = (
                litchi_xlsx::Package::open(&path),
                SourceBackedWorkbook::open(&path),
            ) else {
                continue;
            };
            for (eager_sheet, source_sheet) in
                eager.workbook_ref().sheets().zip(source.sheets().unwrap())
            {
                match (eager_sheet.tables(), source_sheet.tables()) {
                    (Ok(left), Ok(right)) => {
                        assert_eq!(describe(&left), describe(&right), "{}", path.display());
                        tables += right.len();
                        for part in &right {
                            let table = part.table();
                            // Both owners resolve every table by either name,
                            // in any case.
                            for name in [&table.display_name, &table.name] {
                                for spelling in [name.to_lowercase(), name.to_uppercase()] {
                                    let found = source_sheet.table(&spelling).unwrap().unwrap();
                                    let eager_found =
                                        eager_sheet.table(&spelling).unwrap().unwrap();
                                    assert_eq!(
                                        describe(std::slice::from_ref(&found)),
                                        describe(&[eager_found]),
                                        "{} {spelling}",
                                        path.display()
                                    );
                                    assert!(
                                        crate_equivalent(&found.table().display_name, &spelling)
                                            || crate_equivalent(&found.table().name, &spelling)
                                    );
                                }
                            }
                            // Every table's bands tile its range.
                            let whole = table.range().unwrap();
                            let rows = [
                                table.header_range().unwrap(),
                                table.data_range().unwrap(),
                                table.totals_range().unwrap(),
                            ]
                            .into_iter()
                            .flatten()
                            .map(Rect::rows)
                            .sum::<u32>();
                            assert_eq!(rows, whole.rows(), "{} {}", path.display(), table.name);
                        }
                    },
                    (Err(_), Err(_)) => {},
                    (left, right) => panic!(
                        "{}: owners disagree, eager {:?}, source {:?}",
                        path.display(),
                        left.map(|tables| tables.len()),
                        right.map(|tables| tables.len())
                    ),
                }
            }
        }
    }
    assert!(tables >= 30, "only {tables} tables compared");
}

fn crate_equivalent(left: &str, right: &str) -> bool {
    left.to_lowercase() == right.to_lowercase()
}

#[test]
fn table_ranges_split_header_data_and_totals_rows() {
    let workbook = SourceBackedWorkbook::open(fixture("table-sample.xlsx")).unwrap();
    let (sheet, part) = workbook.table("tabelle1").unwrap().unwrap();
    assert_eq!(sheet.position(), 0);
    let table = part.table();
    assert_eq!(table.range().unwrap(), range("C4:G9"));
    assert_eq!(table.header_range().unwrap(), Some(range("C4:G4")));
    assert_eq!(table.data_range().unwrap(), Some(range("C5:G8")));
    assert_eq!(table.totals_range().unwrap(), Some(range("C9:G9")));
    assert_eq!(
        table.column("field 2").unwrap().map(|column| column.id),
        Some(2),
        "column names compare without regard to case"
    );
    assert!(table.column("Field 6").unwrap().is_none());

    let workbook = litchi_xlsx::Package::open(fixture("ExcelTables.xlsx")).unwrap();
    let (_, part) = workbook.workbook_ref().table("TABLENAME").unwrap().unwrap();
    let table = part.table();
    assert_eq!(table.header_range().unwrap(), Some(range("G1:I1")));
    assert_eq!(table.data_range().unwrap(), Some(range("G2:I4")));
    assert_eq!(table.totals_range().unwrap(), None);
    assert!(
        workbook
            .workbook_ref()
            .table("NoSuchTable")
            .unwrap()
            .is_none()
    );
}

#[test]
fn workbook_lookup_searches_every_worksheet_in_order() {
    let path = fixture("ConditionalFormattingSamples.xlsx");
    let eager = litchi_xlsx::Package::open(&path).unwrap();
    let source = SourceBackedWorkbook::open(&path).unwrap();
    let mut seen = 0usize;
    for sheet in source.sheets().unwrap() {
        if sheet.kind() != WorksheetKind::Worksheet {
            continue;
        }
        for part in sheet.tables().unwrap() {
            let name = part.table().display_name.to_ascii_uppercase();
            let (found_sheet, found) = source.table(&name).unwrap().unwrap();
            assert_eq!(found_sheet.position(), sheet.position());
            assert_eq!(found.part_name(), part.part_name());
            let (eager_sheet, eager_found) = eager.workbook_ref().table(&name).unwrap().unwrap();
            assert_eq!(eager_sheet.position(), sheet.position());
            assert_eq!(eager_found.part_name(), part.part_name());
            seen += 1;
        }
    }
    assert_eq!(seen, 9);
    assert!(source.table("missing").unwrap().is_none());
    // These saved sort extents predate the current table bounds. Both readers
    // retain them as metadata; reading a table does not execute its sort.
    for (name, table_range, sort_range, key_range) in [
        ("Table1", "A3:E25", "A3:E24", "A2:A24"),
        ("Table3", "A2:F21", "A2:G22", "A1:A22"),
    ] {
        for table in [
            source.table(name).unwrap().unwrap().1,
            eager.workbook_ref().table(name).unwrap().unwrap().1,
        ] {
            assert_eq!(table.table().ref_range, table_range);
            let sort = table.table().sort_state.as_ref().unwrap();
            assert_eq!(sort.ref_range, sort_range);
            assert_eq!(sort.conditions.len(), 1);
            assert_eq!(sort.conditions[0].ref_range, key_range);
        }
    }
}

#[test]
fn source_tables_read_only_table_parts() {
    let workbook = SourceBackedWorkbook::open(fixture("ExcelTables.xlsx")).unwrap();
    let sheet = workbook.sheet(0_usize).unwrap().unwrap();
    let before = workbook.cache_diagnostics();
    let tables = sheet.tables().unwrap();
    assert_eq!(tables.len(), 1);
    let after = workbook.cache_diagnostics();
    assert_eq!(
        after.cold_loads - before.cold_loads,
        1,
        "only the table part is read, not the worksheet"
    );
    assert_eq!(tables[0].part_name(), "/xl/tables/table1.xml");
    assert_eq!(tables[0].relationship_id(), "rId2");
}

#[test]
fn non_cell_sheets_own_no_tables() {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet/chart_sheet.xlsx");
    let source = SourceBackedWorkbook::open(&path).unwrap();
    let eager = litchi_xlsx::Package::open(&path).unwrap();
    let mut checked = false;
    for (source_sheet, eager_sheet) in source.sheets().unwrap().zip(eager.workbook_ref().sheets()) {
        if source_sheet.kind() == WorksheetKind::Worksheet {
            continue;
        }
        checked = true;
        assert!(matches!(
            source_sheet.tables(),
            Err(Error::NotWorksheet { .. })
        ));
        assert!(matches!(
            eager_sheet.tables(),
            Err(Error::NotWorksheet { .. })
        ));
    }
    assert!(checked, "the fixture has a chart sheet");
    assert!(source.table("anything").unwrap().is_none());
    assert!(eager.workbook_ref().table("anything").unwrap().is_none());
}

#[test]
fn caller_event_limits_reach_eager_and_source_table_reads() {
    // The selected table part has more than 40 physical XML events, while
    // this package's content-types, workbook and relationship metadata stay
    // below that ceiling. Opening therefore succeeds and the deferred table
    // read is the operation that must report the caller's limit.
    let bytes = std::fs::read(fixture("CustomXMLMappings-complex-type.xlsx")).unwrap();
    let limits = SourceLimits::builder()
        .max_events(40)
        .unwrap()
        .build()
        .unwrap();

    let eager = litchi_xlsx::Package::from_bytes_with_options(
        &bytes,
        litchi_xlsx::OpenOptions::new().source_limits(limits),
    )
    .unwrap();
    let eager_sheet = eager.workbook_ref().sheet(0_usize).unwrap().unwrap();
    assert_event_limit(eager_sheet.tables().unwrap_err());

    let source = SourceBackedWorkbook::from_reader_with_options(
        Cursor::new(bytes),
        litchi_xlsx::OpenOptions::new().source_limits(limits),
    )
    .unwrap();
    let source_sheet = source.sheet(0_usize).unwrap().unwrap();
    assert_event_limit(source_sheet.tables().unwrap_err());
}
