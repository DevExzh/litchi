use super::{codec::*, model::*, validation::*};

/// Encode an `XLUnicodeString` (compressed, 16-bit character count).
fn xl_string(text: &str) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&(text.len() as u16).to_le_bytes());
    out.push(0); // compressed (one byte per character)
    out.extend_from_slice(text.as_bytes());
    out
}

fn qsi(name: &str, flags: u16) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&QSI_ATR_PAT.to_le_bytes());
    out.extend_from_slice(&18u16.to_le_bytes()); // itblAutoFmt
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved
    out.extend_from_slice(&xl_string(name));
    out.extend_from_slice(&0u16.to_le_bytes()); // unused4
    out
}

#[allow(clippy::too_many_arguments)]
fn db_query(
    flags: u16,
    cparams: i16,
    cst_query: i16,
    cst_web_post: i16,
    cst_sql_sav: i16,
    cst_odbc_conn: i16,
) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&cparams.to_le_bytes());
    out.extend_from_slice(&cst_query.to_le_bytes());
    out.extend_from_slice(&cst_web_post.to_le_bytes());
    out.extend_from_slice(&cst_sql_sav.to_le_bytes());
    out.extend_from_slice(&cst_odbc_conn.to_le_bytes());
    out
}

fn qsi_sx_tag(name: &str, f_sx: u16, flags: u16, future: u32) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&QSI_SX_TAG_RECORD_TYPE.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&f_sx.to_le_bytes());
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&future.to_le_bytes());
    out.push(0); // verSxLastUpdated
    out.push(0); // verSxUpdatableMin
    out.push(16); // obCchName
    out.push(0); // reserved2
    out.extend_from_slice(&xl_string(name));
    out.extend_from_slice(&0u16.to_le_bytes()); // unused
    out
}

#[allow(clippy::too_many_arguments)]
fn db_query_ext(
    dbt: u16,
    flags: u16,
    grbit_dbt: u16,
    trailing: u16,
    coledb: u16,
    future_bytes: &[u8],
    refresh_interval: u16,
    html_fmt: u16,
    parameter_flags: &[u16],
) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&DB_QUERY_EXT_RECORD_TYPE.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&dbt.to_le_bytes());
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&grbit_dbt.to_le_bytes());
    out.extend_from_slice(&trailing.to_le_bytes());
    out.push(3); // bVerDbqueryEdit
    out.push(2); // bVerDbqueryRefreshed
    out.push(1); // bVerDbqueryRefreshableMin
    out.push(0); // reserved4
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved5
    out.extend_from_slice(&coledb.to_le_bytes());
    out.extend_from_slice(&(future_bytes.len() as u16).to_le_bytes());
    out.extend_from_slice(&refresh_interval.to_le_bytes());
    out.extend_from_slice(&html_fmt.to_le_bytes());
    out.extend_from_slice(&(parameter_flags.len() as u16).to_le_bytes());
    for flag in parameter_flags {
        out.extend_from_slice(&flag.to_le_bytes());
    }
    out.extend_from_slice(future_bytes);
    out
}

fn ext_string(text: &str) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&EXT_STRING_RECORD_TYPE.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&xl_string(text));
    out
}

fn txt_qry(
    flags: u16,
    delimiters: u8,
    custom: Option<char>,
    row_start: i32,
    fields: &[(u32, i32)],
    file: &str,
) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&TXT_QRY_RECORD_TYPE.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // reserved
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // unused1
    out.extend_from_slice(&row_start.to_le_bytes());
    out.push(delimiters);
    out.extend_from_slice(&custom.map_or(0u16, |c| c as u16).to_le_bytes());
    out.push(0); // unused2
    out.extend_from_slice(&(fields.len() as i32).to_le_bytes());
    out.push(b'.'); // chDecimal
    out.push(b','); // chThousSep
    for (format, start) in fields {
        out.extend_from_slice(&format.to_le_bytes());
        out.extend_from_slice(&start.to_le_bytes());
    }
    out.extend_from_slice(&xl_string(file));
    out
}

fn ole_db_conn(flags: u16, cst: u16) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&OLE_DB_CONN_RECORD_TYPE.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // grbitFrt
    out.extend_from_slice(&flags.to_le_bytes());
    out.extend_from_slice(&cst.to_le_bytes());
    out.extend_from_slice(&0u32.to_le_bytes()); // reserved2
    out
}

fn param_qry(sql_type: u16, pbt: u16, grbit: u16, prompt: Option<&str>) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(&sql_type.to_le_bytes());
    out.extend_from_slice(&pbt.to_le_bytes());
    out.extend_from_slice(&grbit.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes()); // fVal
    if let Some(prompt) = prompt {
        out.extend_from_slice(&xl_string(prompt));
        out.push(0); // unused byte
    }
    out
}

fn sort_data(conditions: u32) -> Vec<u8> {
    let mut out = vec![0xAA; 34];
    out[30..34].copy_from_slice(&conditions.to_le_bytes());
    out
}

/// A record type that never belongs to a QUERYTABLE sequence.
const OTHER_RECORD: u16 = 0x000A;

#[test]
fn text_query_sequence() {
    let mut collector = QueryTableCollector::new();
    assert!(collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0x2209)));
    assert!(collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0)
    ));
    assert!(collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Table1", 0, 0x0001, 3)));
    assert!(collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0006, 0, 0, 0, 0, &[1, 2, 3], 15, 0, &[]),
    ));
    let txt = txt_qry(
        0x0001 | 0x0002 | (1 << 2), // fFile, fDelimited, iCpid = Windows (ANSI)
        TXT_DELIM_COMMA | TXT_DELIM_CUSTOM,
        Some('|'),
        2,
        &[(0, 0), (1, 5)],
        "D:\\data\\table1.csv",
    );
    assert!(collector.feed_record(TXT_QRY_RECORD_TYPE, &txt));
    assert!(!collector.feed_record(OTHER_RECORD, &[]));

    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    assert_eq!(table.name, "Table1");
    assert!(table.titles);
    assert!(table.async_refresh);
    assert!(table.save_data);
    assert!(!table.disable_edit);
    assert!(table.overwrite);
    assert!(!table.shrink);
    assert_eq!(table.auto_format_index, 18);
    assert!(table.auto_format_pattern);
    assert_eq!(table.enable_refresh, Some(true));
    assert_eq!(table.qsi_future, 3);
    assert_eq!(table.source, QuerySource::Text);
    assert_eq!(table.refresh_interval, 15);
    assert_eq!(table.edited_version, 3);
    assert_eq!(table.future_bytes, vec![1, 2, 3]);
    let text_query = table.text_query.as_ref().expect("text query present");
    assert!(text_query.delimited);
    assert_eq!(text_query.codepage, TextCodePage::WindowsAnsi);
    assert_eq!(text_query.row_start_at, 2);
    assert!(text_query.comma);
    assert!(!text_query.tab);
    assert_eq!(text_query.custom_delimiter, Some('|'));
    assert_eq!(text_query.text_delimiter, TextDelimiter::QuotationMark);
    assert_eq!(text_query.decimal_separator, '.');
    assert_eq!(text_query.thousands_separator, ',');
    assert_eq!(
        text_query.fields,
        vec![
            TextField {
                format: TextFieldFormat::General,
                start: 0
            },
            TextField {
                format: TextFieldFormat::Text,
                start: 5
            },
        ]
    );
    assert_eq!(text_query.file, "D:\\data\\table1.csv");
}

#[test]
fn odbc_query_with_parameters() {
    let flags = 0x0001 | DBQUERY_SQL | DBQUERY_ODBC_CONN | DBQUERY_SAVE_PWD;
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Odbc1", 0));
    collector.feed_record(DB_OR_PARAM_QRY_RECORD_TYPE, &db_query(flags, 1, 2, 0, 0, 1));
    collector.feed_record(
        SX_STRING_RECORD_TYPE,
        &xl_string("SELECT * FROM t WHERE id = "),
    );
    collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("?"));
    collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("DSN=warehouse;UID=scott"));
    collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("id"));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &param_qry(4, 0, 0, Some("Enter id")),
    );
    collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Odbc1", 0, 0, 0));
    collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0001, DBEXT_MAINTAIN, 0, 0, 0, &[], 0, 0, &[0x0004]),
    );
    collector.feed_record(OTHER_RECORD, &[]);

    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    assert_eq!(table.source, QuerySource::Odbc);
    assert!(table.maintain_connection);
    assert!(table.save_password);
    assert_eq!(
        table.command_text.as_deref(),
        Some("SELECT * FROM t WHERE id = ?")
    );
    assert_eq!(
        table.connection_string.as_deref(),
        Some("DSN=warehouse;UID=scott")
    );
    assert_eq!(table.parameter_flags, vec![0x0004]);
    assert_eq!(table.parameters.len(), 1);
    let parameter = &table.parameters[0];
    assert_eq!(parameter.name, "id");
    assert_eq!(parameter.parameter_type, QueryParameterType::Prompt);
    assert_eq!(parameter.sql_type, 4);
    assert_eq!(parameter.prompt.as_deref(), Some("Enter id"));
}

#[test]
fn web_query_with_post_data() {
    let flags = 0x0004 | DBQUERY_WEB | DBQUERY_TABLES_ONLY_HTML;
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Web1", 0));
    collector.feed_record(DB_OR_PARAM_QRY_RECORD_TYPE, &db_query(flags, 0, 2, 1, 0, 0));
    collector.feed_record(
        SX_STRING_RECORD_TYPE,
        &xl_string("https://example.invalid/report?a="),
    );
    collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("1"));
    collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("q=1&r=2"));
    collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Web1", 0, 0, 0));
    collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0004, 0, 0x0023, DBEXT_TABLE_NAMES, 0, &[], 0, 0x0002, &[]),
    );
    collector.feed_record(EXT_STRING_RECORD_TYPE, &ext_string("Table1,Table2"));
    collector.feed_record(OTHER_RECORD, &[]);

    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    assert_eq!(table.source, QuerySource::Web);
    assert!(table.tables_only_html);
    assert_eq!(
        table.command_text.as_deref(),
        Some("https://example.invalid/report?a=1")
    );
    assert_eq!(table.web_post.as_deref(), Some("q=1&r=2"));
    assert_eq!(table.connection_flags, 0x0023);
    assert_eq!(table.html_formatting, HtmlFormatting::RichText);
    assert_eq!(table.table_names.as_deref(), Some("Table1,Table2"));
}

#[test]
fn ole_db_query_with_chunked_connection() {
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Ole1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0005 | DBQUERY_SQL, 0, 1, 0, 0, 0),
    );
    collector.feed_record(SX_STRING_RECORD_TYPE, &xl_string("SELECT 1"));
    collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Ole1", 0, 0, 0));
    collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0005, 0, 0, 0, 1, &[], 0, 0, &[]),
    );
    collector.feed_record(OLE_DB_CONN_RECORD_TYPE, &ole_db_conn(0x0001, 2));
    collector.feed_record(EXT_STRING_RECORD_TYPE, &ext_string("Provider=SQLOLEDB;"));
    collector.feed_record(EXT_STRING_RECORD_TYPE, &ext_string("Data Source=srv"));
    collector.feed_record(OTHER_RECORD, &[]);

    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    assert_eq!(table.source, QuerySource::OleDb);
    assert_eq!(table.ole_db_connections.len(), 1);
    let connection = &table.ole_db_connections[0];
    assert!(connection.password_stripped);
    assert!(!connection.local);
    assert_eq!(
        connection.connection_string,
        "Provider=SQLOLEDB;Data Source=srv"
    );
}

#[test]
fn multiple_query_tables_per_sheet() {
    let mut collector = QueryTableCollector::new();
    for name in ["First", "Second"] {
        collector.feed_record(QSI_RECORD_TYPE, &qsi(name, 0));
        collector.feed_record(
            DB_OR_PARAM_QRY_RECORD_TYPE,
            &db_query(0x0006, 0, 0, 0, 0, 0),
        );
        collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag(name, 0, 0, 0));
        collector.feed_record(
            DB_QUERY_EXT_RECORD_TYPE,
            &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
        );
    }
    let tables = collector.finish();
    assert_eq!(tables.len(), 2);
    assert_eq!(tables[0].name, "First");
    assert_eq!(tables[1].name, "Second");
}

#[test]
fn pivot_tag_is_handed_back() {
    // fSx=1 while a sequence is open: the tag belongs to a PivotTable
    // view and must not be consumed; the pending table is completed
    // without a bound tag.
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    assert!(!collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("PivotTable1", 1, 0, 0)));
    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].enable_refresh, None);
    assert_eq!(tables[0].qsi_future, 0);
}

#[test]
fn pivot_tag_without_query_table_is_ignored() {
    let mut collector = QueryTableCollector::new();
    assert!(!collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("PivotTable1", 1, 0, 0)));
    assert!(!collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0005, 0, 0, 0, 0, &[], 0, 0, &[]),
    ));
    assert!(collector.finish().is_empty());
}

#[test]
fn mismatched_tag_name_is_ignored() {
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    assert!(collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Other", 0, 0x0001, 7)));
    collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
    );
    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].enable_refresh, None);
    assert_eq!(tables[0].source, QuerySource::Text);
}

#[test]
fn sort_data_inside_sequence_is_consumed() {
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Table1", 0, 0, 0));
    collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
    );
    let base = sort_data(1);
    assert!(collector.feed_record(SORT_DATA_RECORD_TYPE, &base));
    assert!(collector.feed_record(CONTINUE_FRT12_RECORD_TYPE, &[0xBB; 12]));
    assert!(!collector.feed_record(OTHER_RECORD, &[]));

    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    let mut expected = base.clone();
    expected.extend_from_slice(&[0xBB; 12]);
    assert_eq!(tables[0].sort_data_bytes, expected);
}

#[test]
fn sort_data_with_missing_continues_still_completes() {
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    let base = sort_data(3);
    assert!(collector.feed_record(SORT_DATA_RECORD_TYPE, &base));
    // No ContinueFrt12 records arrive; the sequence just ends.
    assert!(!collector.feed_record(OTHER_RECORD, &[]));
    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].sort_data_bytes, base);
}

#[test]
fn sxaddl_qsi_class_is_consumed() {
    let sxaddl = |sxd: u8| {
        let mut out = Vec::new();
        out.extend_from_slice(&SXADDL_RECORD_TYPE.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.push(SXC_QSI_CLASS);
        out.push(sxd);
        out.extend_from_slice(&[0; 6]);
        out
    };
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Table1", 0, 0, 0));
    collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
    );
    assert!(collector.feed_record(SXADDL_RECORD_TYPE, &sxaddl(0x00))); // SXDId
    assert!(collector.feed_record(SXADDL_RECORD_TYPE, &sxaddl(SXD_END)));
    assert!(!collector.feed_record(OTHER_RECORD, &[]));
    assert_eq!(collector.finish().len(), 1);
}

#[test]
fn foreign_sxaddl_class_ends_sequence() {
    let other_class = {
        let mut out = Vec::new();
        out.extend_from_slice(&SXADDL_RECORD_TYPE.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.push(0x01); // SxcView class
        out.push(0x00);
        out
    };
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    assert!(!collector.feed_record(SXADDL_RECORD_TYPE, &other_class));
    assert_eq!(collector.finish().len(), 1);
}

#[test]
fn malformed_records_never_panic() {
    // Truncated Qsi: consumed, no sequence started.
    let mut collector = QueryTableCollector::new();
    assert!(collector.feed_record(QSI_RECORD_TYPE, &[0x01, 0x02]));
    assert!(collector.finish().is_empty());

    // Truncated DbQuery drops the in-progress sequence.
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    assert!(collector.feed_record(DB_OR_PARAM_QRY_RECORD_TYPE, &[0x06, 0x00]));
    assert!(collector.finish().is_empty());

    // Truncated DBQueryExt drops the in-progress sequence.
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    assert!(collector.feed_record(DB_QUERY_EXT_RECORD_TYPE, &[0x03, 0x08, 0x00]));
    assert!(collector.finish().is_empty());

    // Malformed SXString drops the in-progress sequence.
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0001 | DBQUERY_SQL, 0, 1, 0, 0, 0),
    );
    assert!(collector.feed_record(SX_STRING_RECORD_TYPE, &[0xFF, 0xFF, 0x01]));
    assert!(collector.finish().is_empty());

    // Malformed TxtQry is ignored; the table still completes.
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    collector.feed_record(QSI_SX_TAG_RECORD_TYPE, &qsi_sx_tag("Table1", 0, 0, 0));
    collector.feed_record(
        DB_QUERY_EXT_RECORD_TYPE,
        &db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]),
    );
    assert!(collector.feed_record(TXT_QRY_RECORD_TYPE, &[0x05, 0x08]));
    collector.feed_record(OTHER_RECORD, &[]);
    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    assert!(tables[0].text_query.is_none());

    // Declared counts exceeding the payload are clamped/dropped.
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0006, 0, 0, 0, 0, 0),
    );
    let mut ext = db_query_ext(0x0006, 0, 0, 0, 0, &[], 0, 0, &[]);
    ext[20] = 0xFF; // cstFuture claims 255 bytes that are absent
    ext[21] = 0x00;
    assert!(collector.feed_record(DB_QUERY_EXT_RECORD_TYPE, &ext));
    collector.feed_record(OTHER_RECORD, &[]);
    assert_eq!(collector.finish().len(), 1);
}

/// Feed every record of a fixture's Workbook stream through the
/// collector, mirroring the worksheet walker.
fn collect_from_fixture(path: &str) -> Vec<QueryTable> {
    let bytes = std::fs::read(path).expect("fixture readable");
    let mut ole =
        litchi_cfb::OleFile::open(std::io::Cursor::new(bytes)).expect("fixture is a CFB container");
    let stream = ole
        .open_stream(&["Workbook"])
        .or_else(|_| ole.open_stream(&["Book"]))
        .expect("workbook stream present");
    let mut collector = QueryTableCollector::new();
    let mut offset = 0;
    while offset + 4 <= stream.len() {
        let record_type = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]) as usize;
        let body_end = (offset + 4 + length).min(stream.len());
        collector.feed_record(record_type, &stream[offset + 4..body_end]);
        offset = body_end;
    }
    collector.finish()
}

/// 57456.xls declares an SST whose total count underflows its unique
/// count, which the (deliberately strict) workbook parser rejects before
/// any worksheet is reached; walk the fixture's records directly so the
/// real bytes still drive the collector.
#[test]
fn web_query_57456_fixture() {
    let tables = collect_from_fixture("../../test-data/poi/test-data/spreadsheet/57456.xls");
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    assert_eq!(table.name, "ExternalData_1");
    assert_eq!(table.source, QuerySource::Web);
    assert!(table.titles);
    assert!(table.async_refresh);
    assert!(table.auto_refresh);
    assert!(table.shrink);
    assert_eq!(table.enable_refresh, Some(false));
    assert_eq!(table.qsi_future, 3);
    assert_eq!(table.html_formatting, HtmlFormatting::None);
    assert_eq!(table.connection_flags, 0x0023);
    assert!(table.text_query.is_none());
    let url = table.command_text.as_deref().expect("web query URL");
    assert!(url.starts_with("http://bugstop.lenexa.ibm.com:8080/disp_bugs.php?flow=&comp=totals"));
    assert!(url.ends_with("txmgt:wlm&x=z"));
    assert_eq!(url.len(), 867);
    assert!(table.connection_string.is_none());
    assert!(table.web_post.is_none());
}

/// The text-query fixtures parse through the public workbook API; this
/// cross-checks the raw record walk against the same bytes.
#[test]
fn text_query_45365_fixture() {
    let tables = collect_from_fixture("../../test-data/poi/test-data/spreadsheet/45365.xls");
    assert_eq!(tables.len(), 1);
    let table = &tables[0];
    assert_eq!(table.name, "Jac-Jackson-MSC_1");
    assert_eq!(table.source, QuerySource::Text);
    let text_query = table.text_query.as_ref().expect("text query present");
    assert_eq!(text_query.file, "D:\\Jac-Jackson-MSC_1.csv");
    assert_eq!(text_query.row_start_at, 1);
    assert!(text_query.tab);
    assert!(text_query.comma);
}

#[test]
fn negative_chunk_counts_are_clamped() {
    let mut collector = QueryTableCollector::new();
    collector.feed_record(QSI_RECORD_TYPE, &qsi("Table1", 0));
    collector.feed_record(
        DB_OR_PARAM_QRY_RECORD_TYPE,
        &db_query(0x0001 | DBQUERY_SQL | DBQUERY_ODBC_CONN, -1, -2, -3, -4, -5),
    );
    collector.feed_record(OTHER_RECORD, &[]);
    let tables = collector.finish();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].command_text, None);
    assert_eq!(tables[0].connection_string, None);
}
