//! Integration tests for the inert BIFF8 QUERYTABLE sequence reader against
//! real-world POI fixtures. Connection strings, URLs, and file paths are
//! asserted verbatim; they are never opened, resolved, or contacted.

use std::fs::File;
use std::path::PathBuf;

use litchi_xls::{QuerySource, QueryTable, TextCodePage, TextDelimiter, TextFieldFormat, Workbook};

fn poi_fixture(name: &str) -> PathBuf {
    PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet")
        .join(name)
}

fn open(name: &str) -> Workbook<File> {
    Workbook::new(File::open(poi_fixture(name)).unwrap()).unwrap()
}

/// The single query table expected anywhere in the workbook.
fn only_query_table(workbook: &Workbook<File>) -> QueryTable {
    let tables: Vec<QueryTable> = (0..workbook.sheets().len())
        .filter_map(|index| workbook.xls_worksheet(index).ok())
        .flat_map(|worksheet| worksheet.query_tables().to_vec())
        .collect();
    assert_eq!(
        tables.len(),
        1,
        "expected exactly one query table in the workbook"
    );
    tables.into_iter().next().unwrap()
}

fn assert_text_fixture(name: &str, table_name: &str, file: &str) {
    let table = only_query_table(&open(name));
    assert_eq!(table.name, table_name);
    assert_eq!(table.source, QuerySource::Text);
    assert_eq!(table.enable_refresh, Some(false));
    assert_eq!(table.qsi_future, 3);
    assert!(table.titles);
    assert!(table.async_refresh);
    assert!(table.save_data);
    assert!(table.overwrite);
    assert!(!table.disable_refresh);
    assert!(!table.disable_edit);
    assert!(!table.auto_refresh);
    assert_eq!(table.auto_format_index, 18);
    assert!(table.auto_format_pattern);
    assert!(!table.auto_format_number);
    assert_eq!(table.command_text, None);
    assert_eq!(table.connection_string, None);
    assert!(table.parameters.is_empty());
    assert!(table.ole_db_connections.is_empty());
    assert!(table.future_bytes.is_empty());
    assert!(table.sort_data_bytes.is_empty());

    let text_query = table.text_query.as_ref().expect("text query present");
    assert!(text_query.delimited);
    assert_eq!(text_query.codepage, TextCodePage::WindowsAnsi);
    assert!(!text_query.prompt_for_file);
    assert_eq!(text_query.row_start_at, 1);
    assert!(text_query.tab);
    assert!(text_query.comma);
    assert!(!text_query.space);
    assert!(!text_query.semicolon);
    assert_eq!(text_query.custom_delimiter, None);
    assert!(!text_query.consecutive);
    assert_eq!(text_query.text_delimiter, TextDelimiter::QuotationMark);
    assert_eq!(text_query.decimal_separator, '.');
    assert_eq!(text_query.thousands_separator, ',');
    assert_eq!(text_query.fields.len(), 1);
    assert_eq!(text_query.fields[0].format, TextFieldFormat::General);
    assert_eq!(text_query.fields[0].start, 0);
    assert_eq!(text_query.file, file);
}

#[test]
fn text_query_45365() {
    assert_text_fixture(
        "45365.xls",
        "Jac-Jackson-MSC_1",
        "D:\\Jac-Jackson-MSC_1.csv",
    );
}

#[test]
fn text_query_45365_2() {
    assert_text_fixture(
        "45365-2.xls",
        "Bos-Walpole-MSC_1",
        "D:\\Bos-Walpole-MSC_1.csv",
    );
}

#[test]
fn text_query_mr_extra_lines() {
    assert_text_fixture("MRExtraLines.xls", "SPFDMATABS0", "D:\\SPFDMATABS0.csv");
}

#[test]
fn web_query_57456_is_blocked_by_an_unrelated_sst_defect() {
    // 57456.xls declares an SST whose total count underflows its unique
    // count; the deliberately strict workbook parser rejects the file before
    // any worksheet is reached. Its query table is therefore covered by a
    // record-level test inside the crate (query_table::tests). This test
    // documents the pre-existing boundary so a future SST-leniency change
    // can promote the fixture to a full read test.
    let result = Workbook::new(File::open(poi_fixture("57456.xls")).unwrap());
    assert!(result.is_err());
}

#[test]
fn pivot_table_fixture_yields_no_query_tables() {
    // QsiSXTag records with fSx=1 belong to PivotTable views and must never
    // surface as query tables.
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join(
        "../../test-data/libreoffice-core/sc/qa/unit/data/xls/pivottable_number_grouping.xls",
    );
    let workbook = Workbook::new(File::open(path).unwrap()).unwrap();
    for index in 0..workbook.sheets().len() {
        let worksheet = workbook.xls_worksheet(index).unwrap();
        assert!(
            worksheet.query_tables().is_empty(),
            "sheet {index} unexpectedly has query tables"
        );
    }
}

#[test]
fn ordinary_workbook_has_no_query_tables() {
    let workbook = open("Simple.xls");
    for index in 0..workbook.sheets().len() {
        assert!(
            workbook
                .xls_worksheet(index)
                .unwrap()
                .query_tables()
                .is_empty()
        );
    }
}
