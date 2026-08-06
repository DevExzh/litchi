//! Focused binary and semantic regression tests for mail-merge parts.

use super::codec::parse_odso_properties;
use super::validation::*;
use super::*;
use crate::parts::fib::FileInformationBlock;

fn utf16(text: &str) -> Vec<u8> {
    text.encode_utf16().flat_map(u16::to_le_bytes).collect()
}

fn pmfs_bytes(kind: u8, flags: u8, tk_field: i16, tk_rec: i16, fnpi: u16) -> [u8; PMFS_LEN] {
    let mut bytes = [0u8; PMFS_LEN];
    bytes[0] = kind;
    bytes[1] = flags;
    bytes[2..4].copy_from_slice(&tk_field.to_le_bytes());
    bytes[4..6].copy_from_slice(&tk_rec.to_le_bytes());
    bytes[6..8].copy_from_slice(&fnpi.to_le_bytes());
    bytes
}

struct PmsBuilder {
    wpms: u16,
    ipmf_mf: u8,
    ipmf_fetch: u8,
    irec_cur: u32,
    rfs: u32,
    sql: Option<String>,
    sttbf: Option<Vec<Vec<u8>>>,
    wpmsdt: Option<u32>,
}

impl PmsBuilder {
    fn new() -> Self {
        PmsBuilder {
            wpms: 0x0409,
            ipmf_mf: 0,
            ipmf_fetch: 0,
            irec_cur: IREC_NIL,
            rfs: 0,
            sql: None,
            sttbf: None,
            wpmsdt: None,
        }
    }

    fn build(&self) -> Vec<u8> {
        let mut data = Vec::new();
        data.extend_from_slice(&self.wpms.to_le_bytes());
        data.push(self.ipmf_mf);
        data.push(self.ipmf_fetch);
        data.extend_from_slice(&self.irec_cur.to_le_bytes());
        data.extend_from_slice(&pmfs_bytes(0x00, 0, 0, 0, 0xFFF3));
        data.extend_from_slice(&pmfs_bytes(0xFF, 0, 0, 0, 0xFFF3));
        data.extend_from_slice(&self.rfs.to_le_bytes());
        match &self.sql {
            Some(sql) => {
                let mut encoded = utf16(sql);
                encoded.extend_from_slice(&[0, 0]);
                data.extend_from_slice(&(encoded.len() as u16).to_le_bytes());
                data.extend_from_slice(&encoded);
            },
            None => data.extend_from_slice(&0u16.to_le_bytes()),
        }
        if let Some(strings) = &self.sttbf {
            data.extend_from_slice(&STTB_F_EXTEND.to_le_bytes());
            data.extend_from_slice(&(strings.len() as u16).to_le_bytes());
            data.extend_from_slice(&0u16.to_le_bytes());
            for string in strings {
                data.extend_from_slice(&((string.len() / 2) as u16).to_le_bytes());
                data.extend_from_slice(string);
            }
        }
        if let Some(doc_type) = self.wpmsdt {
            data.extend_from_slice(&doc_type.to_le_bytes());
        }
        data
    }
}

#[test]
fn parses_full_pms() {
    let mut builder = PmsBuilder::new();
    builder.ipmf_mf = 1;
    builder.ipmf_fetch = 0;
    builder.irec_cur = 1;
    builder.sql = Some("SELECT * FROM [myTable] WHERE x".to_string());
    builder.wpmsdt = Some(0x01);
    let pms = Pms::parse(&builder.build()).unwrap();
    assert_eq!(pms.state.merge_type, MailMergeType::Letters);
    assert!(pms.state.main_document);
    assert!(!pms.state.data_source);
    assert!(pms.state.suppress_blank_lines);
    assert_eq!(pms.state.destination, MailMergeDestination::None);
    assert_eq!(pms.header_source_index, 1);
    assert_eq!(pms.fetch_source_index, 0);
    assert_eq!(pms.current_record, Some(1));
    assert_eq!(pms.sources[0].source_kind, MergeDataSourceKind::DataFile);
    assert_eq!(pms.sources[1].source_kind, MergeDataSourceKind::None);
    assert!(pms.sources[0].file_name.is_mail_merge_source());
    assert_eq!(pms.sources[0].file_name.identifier(), 0x0FFF);
    assert_eq!(
        pms.sql_query.as_deref(),
        Some("SELECT * FROM [myTable] WHERE x")
    );
    assert!(pms.strings.is_none());
    assert_eq!(pms.document_type, Some(MailMergeDocumentType::Letters));
}

#[test]
fn parses_minimal_pms() {
    let pms = Pms::parse(&PmsBuilder::new().build()).unwrap();
    assert_eq!(pms.current_record, None);
    assert_eq!(pms.sql_query, None);
    assert_eq!(pms.document_type, None);
}

#[test]
fn parses_pms_with_sttbf_rfs() {
    let mut builder = PmsBuilder::new();
    builder.rfs = 0x0001_0000 | RFS_SHOW_DATA; // hsttbRfs nonzero
    builder.sttbf = Some(vec![
        utf16("DSN=mailmerge;"),
        utf16(""),
        utf16("Your order"),
        utf16("Email"),
        utf16("ignored"),
    ]);
    builder.wpmsdt = Some(0x10);
    let pms = Pms::parse(&builder.build()).unwrap();
    assert!(pms.filter.show_data);
    let strings = pms.strings.unwrap();
    assert_eq!(strings.strings().len(), 5);
    assert_eq!(strings.connection_string(), "DSN=mailmerge;");
    assert_eq!(strings.header_connection_string(), "");
    assert_eq!(strings.email_subject(), "Your order");
    assert_eq!(strings.address_column(), "Email");
    assert_eq!(pms.document_type, Some(MailMergeDocumentType::Email));
}

#[test]
fn parses_pmfs_flags_and_tokens() {
    let pmfs = Pmfs::parse(&pmfs_bytes(0x00, 0x0F, 0x06, 0x02, 0xFFF3)).unwrap();
    assert!(pmfs.link_to_file);
    assert!(pmfs.link_to_connection);
    assert!(pmfs.no_prompt_query_tools);
    assert!(pmfs.uses_query);
    assert_eq!(pmfs.field_separator(), Some(MergeFileToken::Tab));
    assert_eq!(pmfs.record_separator(), Some(MergeFileToken::Enter));
    assert_eq!(
        MergeFileToken::from_raw(0x48),
        Some(MergeFileToken::TableRow)
    );
    assert_eq!(MergeFileToken::from_raw(0x30), None);
}

#[test]
fn parses_rfs_flags() {
    // byte0 = 0x82: grfChkErr=1, fMailAsHtml; hsttbRfs = 0.
    let rfs = Rfs::parse(0x0082).unwrap();
    assert!(!rfs.show_data);
    assert_eq!(rfs.error_checking, MergeErrorCheck::PauseAndReport);
    assert!(rfs.mail_as_html);
    assert!(!rfs.mail_as_text);
    assert!(!rfs.has_string_table);
    assert!(Rfs::parse(0x0006).is_err()); // grfChkErr = 3
}

#[test]
fn rejects_malformed_pms() {
    let good = PmsBuilder::new().build();
    // Truncated header.
    assert!(Pms::parse(&good[..PMS_HEADER_LEN - 1]).is_err());
    // ipmfMF out of range.
    let mut bad = good.clone();
    bad[2] = 2;
    assert!(Pms::parse(&bad).is_err());
    // ipmfFetch out of range.
    let mut bad = good.clone();
    bad[3] = 2;
    assert!(Pms::parse(&bad).is_err());
    // iRecCur out of range.
    let mut builder = PmsBuilder::new();
    builder.irec_cur = IREC_MAX + 1;
    assert!(Pms::parse(&builder.build()).is_err());
    // Undefined wpmsType.
    let mut builder = PmsBuilder::new();
    builder.wpms = 0x0003 << WPMS_TYPE_SHIFT;
    assert!(Pms::parse(&builder.build()).is_err());
    // Undefined wpmsDest.
    let mut builder = PmsBuilder::new();
    builder.wpms = 0x0003 << WPMS_DEST_SHIFT;
    assert!(Pms::parse(&builder.build()).is_err());
    // Undefined data source kind.
    let mut bad = good.clone();
    bad[8] = 0x06;
    assert!(Pms::parse(&bad).is_err());
    // Odd SQL length.
    let mut bad = PmsBuilder::new().build();
    bad[28] = 3;
    bad.extend_from_slice(&[0, 0, 0]);
    assert!(Pms::parse(&bad).is_err());
    // SQL length too small (null terminator only).
    let mut bad = PmsBuilder::new().build();
    bad[28] = 2;
    bad.extend_from_slice(&[0, 0]);
    assert!(Pms::parse(&bad).is_err());
    // SQL length too large.
    let mut builder = PmsBuilder::new();
    builder.sql = Some("x".repeat(300));
    assert!(Pms::parse(&builder.build()).is_err());
    // SQL missing its null terminator.
    let mut bad = PmsBuilder::new().build();
    bad[28] = 4;
    bad.extend_from_slice(&utf16("xy"));
    assert!(Pms::parse(&bad).is_err());
    // Declared string table missing.
    let mut builder = PmsBuilder::new();
    builder.rfs = 0x0001_0000;
    assert!(Pms::parse(&builder.build()).is_err());
    // Partial trailing Wpmsdt.
    let mut bad = PmsBuilder::new().build();
    bad.extend_from_slice(&[0, 0]);
    assert!(Pms::parse(&bad).is_err());
    // Undefined Wpmsdt document type.
    let mut builder = PmsBuilder::new();
    builder.wpmsdt = Some(0x03);
    assert!(Pms::parse(&builder.build()).is_err());
}

#[test]
fn rejects_malformed_sttbf_rfs() {
    let with_table = |strings: Vec<Vec<u8>>, f_extend: u16, cb_extra: u16| {
        let mut builder = PmsBuilder::new();
        builder.rfs = 0x0001_0000;
        builder.sttbf = Some(strings);
        let mut data = builder.build();
        if f_extend != STTB_F_EXTEND {
            let at = PMS_HEADER_LEN;
            data[at..at + 2].copy_from_slice(&f_extend.to_le_bytes());
        }
        if cb_extra != 0 {
            let at = PMS_HEADER_LEN + 4;
            data[at..at + 2].copy_from_slice(&cb_extra.to_le_bytes());
        }
        data
    };
    let strings = || vec![utf16("a"), utf16(""), utf16(""), utf16(""), utf16("")];
    // Bad fExtend.
    assert!(Pms::parse(&with_table(strings(), 0x0000, 0)).is_err());
    // Bad cbExtra.
    assert!(Pms::parse(&with_table(strings(), STTB_F_EXTEND, 2)).is_err());
    // Too few strings (cData = 3): truncated table.
    assert!(Pms::parse(&with_table(strings()[..3].to_vec(), STTB_F_EXTEND, 0)).is_err());
    // String exceeding 255 characters.
    let mut oversized = strings();
    oversized[0] = utf16(&"x".repeat(256));
    assert!(Pms::parse(&with_table(oversized, STTB_F_EXTEND, 0)).is_err());
}

fn odso_item(id: u16, value: &[u8]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&id.to_le_bytes());
    if value.len() >= ODSO_LARGE as usize {
        data.extend_from_slice(&ODSO_LARGE.to_le_bytes());
        data.extend_from_slice(&(value.len() as u32).to_le_bytes());
    } else {
        data.extend_from_slice(&(value.len() as u16).to_le_bytes());
    }
    data.extend_from_slice(value);
    data
}

fn recipient_info_bytes(recipients: &[Vec<(u16, Vec<u8>)>]) -> Vec<u8> {
    let mut list = Vec::new();
    for items in recipients {
        for (id, value) in items {
            list.extend_from_slice(&id.to_le_bytes());
            list.extend_from_slice(&(value.len() as u16).to_le_bytes());
            list.extend_from_slice(value);
        }
        list.extend_from_slice(&[0, 0, 0, 0]);
    }
    let mut data = Vec::new();
    data.extend_from_slice(&COUNT_MARKER.to_le_bytes());
    data.extend_from_slice(&CB_COUNT.to_le_bytes());
    data.extend_from_slice(&(recipients.len() as u32).to_le_bytes());
    data.extend_from_slice(&LIST_SIZE_MARKER.to_le_bytes());
    data.extend_from_slice(&(list.len() as u16).to_le_bytes());
    data.extend_from_slice(&list);
    data
}

fn field_map_info_bytes(mappings: &[Vec<(u16, Vec<u8>)>]) -> Vec<u8> {
    let mut list = Vec::new();
    for items in mappings {
        for (id, value) in items {
            list.extend_from_slice(&id.to_le_bytes());
            list.extend_from_slice(&(value.len() as u16).to_le_bytes());
            list.extend_from_slice(value);
        }
        list.extend_from_slice(&[0, 0, 0, 0]);
    }
    let mut data = Vec::new();
    data.extend_from_slice(&COUNT_MARKER.to_le_bytes());
    data.extend_from_slice(&CB_COUNT.to_le_bytes());
    data.extend_from_slice(&FIELD_MAP_COUNT.to_le_bytes());
    data.extend_from_slice(&LIST_SIZE_MARKER.to_le_bytes());
    data.extend_from_slice(&(list.len() as u16).to_le_bytes());
    data.extend_from_slice(&list);
    data
}

#[test]
fn parses_odso_scalar_properties() {
    let mut bag = Vec::new();
    bag.extend_from_slice(&odso_item(
        ODSO_ID_CONNECTION_STRING,
        &utf16("Provider=SQLOLEDB;Data Source=srv;"),
    ));
    bag.extend_from_slice(&odso_item(ODSO_ID_DATA_TABLE, &utf16("Customers")));
    bag.extend_from_slice(&odso_item(
        ODSO_ID_DATA_SOURCE_FILE,
        &utf16("C:\\data\\customers.mdb"),
    ));
    bag.extend_from_slice(&odso_item(ODSO_ID_CONNECTION_TYPE, &5u32.to_le_bytes()));
    bag.extend_from_slice(&odso_item(ODSO_ID_COLUMN_DELIMITER, &0x2Cu16.to_le_bytes()));
    bag.extend_from_slice(&odso_item(ODSO_ID_FIRST_ROW_IS_HEADER, &1u32.to_le_bytes()));
    bag.extend_from_slice(&odso_item(ODSO_ID_WIZARD_STEP, &3u16.to_le_bytes()));
    let properties = parse_odso_properties(&bag).unwrap();
    assert_eq!(
        properties,
        vec![
            OdsoProperty::ConnectionString("Provider=SQLOLEDB;Data Source=srv;".into()),
            OdsoProperty::DataTable("Customers".into()),
            OdsoProperty::DataSourceFile("C:\\data\\customers.mdb".into()),
            OdsoProperty::ConnectionType(5),
            OdsoProperty::ColumnDelimiter(0x2C),
            OdsoProperty::FirstRowIsHeader(true),
            OdsoProperty::WizardStep(3),
        ]
    );
}

#[test]
fn parses_odso_large_and_unknown_properties() {
    let large_value = utf16(&"x".repeat(0x1_0000));
    let mut bag = odso_item(ODSO_ID_CONNECTION_STRING, &large_value);
    bag.extend_from_slice(&odso_item(0x0099, &[1, 2, 3]));
    let properties = parse_odso_properties(&bag).unwrap();
    assert_eq!(properties.len(), 2);
    assert_eq!(
        properties[1],
        OdsoProperty::Unknown {
            id: 0x0099,
            data: vec![1, 2, 3],
        }
    );
}

#[test]
fn rejects_malformed_odso_bags() {
    // Partial property header.
    assert!(parse_odso_properties(&[0, 0, 4]).is_err());
    // Value overruns the bag.
    let mut bag = odso_item(ODSO_ID_DATA_TABLE, &utf16("abc"));
    bag.truncate(bag.len() - 2);
    assert!(parse_odso_properties(&bag).is_err());
    // Large property with an overrunning size.
    let mut bag = Vec::new();
    bag.extend_from_slice(&0u16.to_le_bytes());
    bag.extend_from_slice(&ODSO_LARGE.to_le_bytes());
    bag.extend_from_slice(&100u32.to_le_bytes());
    assert!(parse_odso_properties(&bag).is_err());
    // Odd-length Unicode string.
    assert!(parse_odso_properties(&odso_item(ODSO_ID_DATA_TABLE, b"a")).is_err());
    // Wrong scalar sizes.
    assert!(parse_odso_properties(&odso_item(ODSO_ID_CONNECTION_TYPE, &[0; 3])).is_err());
    assert!(parse_odso_properties(&odso_item(ODSO_ID_COLUMN_DELIMITER, &[0; 4])).is_err());
    assert!(parse_odso_properties(&odso_item(ODSO_ID_WIZARD_STEP, &[0; 4])).is_err());
    // Out-of-range scalar values.
    assert!(
        parse_odso_properties(&odso_item(ODSO_ID_FIRST_ROW_IS_HEADER, &2u32.to_le_bytes()))
            .is_err()
    );
    assert!(parse_odso_properties(&odso_item(ODSO_ID_WIZARD_STEP, &7u16.to_le_bytes())).is_err());
    assert!(parse_odso_properties(&odso_item(ODSO_ID_WIZARD_STEP, &0u16.to_le_bytes())).is_err());
}

fn filter_item_bytes(column: u32, comparison: u32, condition: u32, value: &str) -> Vec<u8> {
    let mut body = Vec::new();
    body.extend_from_slice(&column.to_le_bytes());
    body.extend_from_slice(&comparison.to_le_bytes());
    body.extend_from_slice(&condition.to_le_bytes());
    body.extend_from_slice(&utf16(value));
    body.extend_from_slice(&[0, 0]);
    let mut item = Vec::new();
    item.extend_from_slice(&((body.len() + 4) as u32).to_le_bytes());
    item.extend_from_slice(&body);
    item
}

#[test]
fn parses_odso_recipient_filters() {
    let mut value = filter_item_bytes(2, 3, 0, "smith");
    value.extend_from_slice(&filter_item_bytes(0, 7, 1, ""));
    let bag = odso_item(ODSO_ID_RECIPIENT_FILTERS, &value);
    let properties = parse_odso_properties(&bag).unwrap();
    assert_eq!(
        properties,
        vec![OdsoProperty::RecipientFilters(vec![
            FilterDataItem {
                column: 2,
                comparison: FilterComparison::GreaterThan,
                condition: FilterCondition::And,
                value: "smith".into(),
            },
            FilterDataItem {
                column: 0,
                comparison: FilterComparison::NotEmpty,
                condition: FilterCondition::Or,
                value: String::new(),
            },
        ])]
    );
}

#[test]
fn rejects_malformed_filters() {
    let wrap = |value: &[u8]| parse_odso_properties(&odso_item(ODSO_ID_RECIPIENT_FILTERS, value));
    // cbItem smaller than the fixed prefix.
    assert!(wrap(&4u32.to_le_bytes()).is_err());
    // cbItem overrunning the value.
    let mut bad = filter_item_bytes(0, 0, 0, "a");
    bad[0] = 0xFF;
    assert!(wrap(&bad).is_err());
    // Column index out of range.
    assert!(wrap(&filter_item_bytes(255, 0, 0, "a")).is_err());
    // Undefined comparison operator.
    assert!(wrap(&filter_item_bytes(0, 8, 0, "a")).is_err());
    // Undefined condition.
    assert!(wrap(&filter_item_bytes(0, 0, 2, "a")).is_err());
    // Missing null terminator: trim the terminator and adjust cbItem.
    let mut bad = filter_item_bytes(0, 0, 0, "a");
    bad.truncate(bad.len() - 2);
    let size = (bad.len() as u32).to_le_bytes();
    bad[..4].copy_from_slice(&size);
    assert!(wrap(&bad).is_err());
    // Comparison string exceeding 212 characters.
    assert!(wrap(&filter_item_bytes(0, 0, 0, &"x".repeat(213))).is_err());
}

#[test]
fn parses_odso_sort_order() {
    let mut value = Vec::new();
    value.extend_from_slice(&1u32.to_le_bytes());
    value.extend_from_slice(&0u32.to_le_bytes());
    value.extend_from_slice(&2u32.to_le_bytes());
    value.extend_from_slice(&1u32.to_le_bytes());
    let bag = odso_item(ODSO_ID_SORT_ORDER, &value);
    let properties = parse_odso_properties(&bag).unwrap();
    assert_eq!(
        properties,
        vec![OdsoProperty::SortOrder(vec![
            SortColumnAndDirection {
                column: 1,
                direction: SortDirection::Ascending,
            },
            SortColumnAndDirection {
                column: 2,
                direction: SortDirection::Descending,
            },
        ])]
    );
    // Partial item.
    assert!(parse_odso_properties(&odso_item(ODSO_ID_SORT_ORDER, &[0; 4])).is_err());
    // More than three keys.
    let mut too_many = Vec::new();
    for _ in 0..4 {
        too_many.extend_from_slice(&0u32.to_le_bytes());
        too_many.extend_from_slice(&0u32.to_le_bytes());
    }
    assert!(parse_odso_properties(&odso_item(ODSO_ID_SORT_ORDER, &too_many)).is_err());
    // Column out of range.
    let mut bad_column = Vec::new();
    bad_column.extend_from_slice(&255u32.to_le_bytes());
    bad_column.extend_from_slice(&0u32.to_le_bytes());
    assert!(parse_odso_properties(&odso_item(ODSO_ID_SORT_ORDER, &bad_column)).is_err());
    // Undefined direction.
    let mut bad_direction = Vec::new();
    bad_direction.extend_from_slice(&0u32.to_le_bytes());
    bad_direction.extend_from_slice(&2u32.to_le_bytes());
    assert!(parse_odso_properties(&odso_item(ODSO_ID_SORT_ORDER, &bad_direction)).is_err());
}

#[test]
fn parses_odso_recipient_info() {
    let first = vec![
        (RECIPIENT_INCLUDED, 0u32.to_le_bytes().to_vec()),
        (RECIPIENT_UNIQUE_COLUMN, 4u32.to_le_bytes().to_vec()),
        (RECIPIENT_UNIQUE_VALUE, utf16("key-1")),
    ];
    let second = vec![(RECIPIENT_HASH, 0xDEAD_BEEFu32.to_le_bytes().to_vec())];
    let bag = odso_item(ODSO_ID_RECIPIENTS, &recipient_info_bytes(&[first, second]));
    let properties = parse_odso_properties(&bag).unwrap();
    let [OdsoProperty::Recipients(info)] = properties.as_slice() else {
        panic!("expected recipient info");
    };
    assert_eq!(info.recipients.len(), 2);
    assert!(!info.recipients[0].included);
    assert_eq!(info.recipients[0].unique_column, Some(4));
    assert_eq!(info.recipients[0].unique_value.as_deref(), Some("key-1"));
    // Inclusion defaults to true when no status item is stored.
    assert!(info.recipients[1].included);
    assert_eq!(info.recipients[1].record_hash, Some(0xDEAD_BEEF));
}

#[test]
fn rejects_malformed_recipient_info() {
    let wrap = |value: &[u8]| parse_odso_properties(&odso_item(ODSO_ID_RECIPIENTS, value));
    // Wrong count marker.
    let mut bad = recipient_info_bytes(&[]);
    bad[1] = 1;
    assert!(wrap(&bad).is_err());
    // Wrong cbCount.
    let mut bad = recipient_info_bytes(&[]);
    bad[3] = 8;
    assert!(wrap(&bad).is_err());
    // Wrong list size marker.
    let mut bad = recipient_info_bytes(&[]);
    bad[9] = 2;
    assert!(wrap(&bad).is_err());
    // List size overrun.
    let mut bad = recipient_info_bytes(&[]);
    bad[10] = 4;
    assert!(wrap(&bad).is_err());
    // Undefined item id.
    let bad = recipient_info_bytes(&[vec![(0x0009, 0u32.to_le_bytes().to_vec())]]);
    assert!(wrap(&bad).is_err());
    // Terminator carrying data.
    let bad = recipient_info_bytes(&[vec![(ITEM_TERMINATOR, vec![0, 0, 0, 0])]]);
    // The terminator is emitted after the items, so inject a bad one.
    let mut with_bad_terminator = bad.clone();
    let at = 12; // start of the list
    with_bad_terminator[at + 2] = 4;
    assert!(wrap(&with_bad_terminator).is_err());
    // Inclusion value other than 0/1.
    let bad = recipient_info_bytes(&[vec![(RECIPIENT_INCLUDED, 2u32.to_le_bytes().to_vec())]]);
    assert!(wrap(&bad).is_err());
    // Missing terminator: declared list ends mid-recipient.
    let mut good =
        recipient_info_bytes(&[vec![(RECIPIENT_UNIQUE_COLUMN, 4u32.to_le_bytes().to_vec())]]);
    let total = good.len();
    good[10] = (total - 12 - 4) as u8; // drop the terminator from the size
    good.truncate(total - 4);
    assert!(wrap(&good).is_err());
}

#[test]
fn parses_odso_field_map_info() {
    let mut mappings: Vec<Vec<(u16, Vec<u8>)>> = vec![Vec::new(); FIELD_MAP_COUNT as usize];
    mappings[2] = vec![
        (
            FIELD_MAP_MAPPED,
            FIELD_MAP_MAPPED_VALUE.to_le_bytes().to_vec(),
        ),
        (FIELD_MAP_COLUMN_NAME, utf16("GivenName")),
        (FIELD_MAP_FIELD_NAME, utf16("First Name")),
        (FIELD_MAP_COLUMN_INDEX, 3u32.to_le_bytes().to_vec()),
    ];
    mappings[19] = vec![(
        FIELD_MAP_COLUMN_INDEX,
        FIELD_MAP_COLUMN_NIL.to_le_bytes().to_vec(),
    )];
    let bag = odso_item(ODSO_ID_FIELD_MAP, &field_map_info_bytes(&mappings));
    let properties = parse_odso_properties(&bag).unwrap();
    let [OdsoProperty::FieldMap(info)] = properties.as_slice() else {
        panic!("expected field map info");
    };
    assert_eq!(info.mappings.len(), FIELD_MAP_COUNT as usize);
    assert_eq!(info.mappings[2].column_index, Some(3));
    assert_eq!(info.mappings[2].column_name.as_deref(), Some("GivenName"));
    // 0xFFFFFFFF means "not mapped".
    assert_eq!(info.mappings[19].column_index, None);
    assert_eq!(FieldMapInfo::STANDARD_ADDRESS_FIELDS[2], "First Name");
    assert_eq!(FieldMapInfo::STANDARD_ADDRESS_FIELDS[29], "Department");
}

#[test]
fn rejects_malformed_field_map_info() {
    let empty = || vec![Vec::new(); FIELD_MAP_COUNT as usize];
    let wrap = |value: &[u8]| parse_odso_properties(&odso_item(ODSO_ID_FIELD_MAP, value));
    // Wrong count marker.
    let mut bad = field_map_info_bytes(&empty());
    bad[1] = 1;
    assert!(wrap(&bad).is_err());
    // Wrong field count.
    let mut bad = field_map_info_bytes(&empty());
    bad[4] = 29;
    assert!(wrap(&bad).is_err());
    // Wrong list size marker.
    let mut bad = field_map_info_bytes(&empty());
    bad[9] = 2;
    assert!(wrap(&bad).is_err());
    // Mapped flag other than 1.
    let mut mappings = empty();
    mappings[0] = vec![(FIELD_MAP_MAPPED, 2u32.to_le_bytes().to_vec())];
    assert!(wrap(&field_map_info_bytes(&mappings)).is_err());
    // Undefined item id.
    let mut mappings = empty();
    mappings[0] = vec![(0x0005, 0u32.to_le_bytes().to_vec())];
    assert!(wrap(&field_map_info_bytes(&mappings)).is_err());
    // Missing terminator on the last mapping.
    let mut good = field_map_info_bytes(&empty());
    let total = good.len();
    good[10] = (total - 12 - 4) as u8;
    good.truncate(total - 4);
    assert!(wrap(&good).is_err());
}

#[test]
fn parses_pms_new_through_the_fib() {
    const FIB_POINTERS: usize = 145;

    fn set_fib_pointer(fib: &mut [u8], index: usize, offset: u32, length: u32) {
        let declared = u16::from_le_bytes([fib[152], fib[153]]);
        let count = declared.max(u16::try_from(index + 1).unwrap());
        fib[152..154].copy_from_slice(&count.to_le_bytes());
        let start = 154 + index * 8;
        fib[start..start + 4].copy_from_slice(&offset.to_le_bytes());
        fib[start + 4..start + 8].copy_from_slice(&length.to_le_bytes());
    }

    let mut fib_data = vec![0; 154 + FIB_POINTERS * 8];
    fib_data[0..2].copy_from_slice(&0xA5ECu16.to_le_bytes());
    fib_data[2..4].copy_from_slice(&0x0101u16.to_le_bytes());
    fib_data[152..154].copy_from_slice(&(FIB_POINTERS as u16).to_le_bytes());

    let mut builder = PmsBuilder::new();
    builder.irec_cur = 7;
    let pms_new = builder.build();
    set_fib_pointer(&mut fib_data, FC_PMS_NEW, 0, pms_new.len() as u32);
    let fib = FileInformationBlock::parse(&fib_data).unwrap();

    let mail_merge = DocumentMailMerge::parse(&fib, &pms_new)
        .unwrap()
        .expect("merge state present");
    assert!(mail_merge.state().is_none());
    assert_eq!(
        mail_merge.new_state().and_then(|pms| pms.current_record),
        Some(7)
    );

    // A malformed PmsNew is reported, not ignored.
    let mut fib_data = fib.raw_data().to_vec();
    set_fib_pointer(&mut fib_data, FC_PMS_NEW, 0, 1);
    let fib = FileInformationBlock::parse(&fib_data).unwrap();
    assert!(DocumentMailMerge::parse(&fib, &pms_new).is_err());
}
