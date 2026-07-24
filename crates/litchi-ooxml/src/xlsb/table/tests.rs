//! Synthetic Brt-record stream tests for the Table (ListObject) parser.

use super::model::*;
use super::parse::{parse_table_part, parse_table_part_rel_ids};
use crate::xlsb::error::{XlsbError, XlsbResult};
use crate::xlsb::records::record_types as rt;
use crate::xlsb::writer::RecordWriter;

fn wide_string(value: &str) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&(value.encode_utf16().count() as u32).to_le_bytes());
    for unit in value.encode_utf16() {
        data.extend_from_slice(&unit.to_le_bytes());
    }
    data
}

fn nullable_wide_string(value: Option<&str>) -> Vec<u8> {
    match value {
        Some(value) => wide_string(value),
        None => u32::MAX.to_le_bytes().to_vec(),
    }
}

fn stream(records: &[(u16, Vec<u8>)]) -> Vec<u8> {
    let mut data = Vec::new();
    let mut writer = RecordWriter::new(&mut data);
    for (record_type, payload) in records {
        writer.write_record(*record_type, payload).unwrap();
    }
    data
}

fn parse(records: &[(u16, Vec<u8>)]) -> XlsbResult<XlsbTable> {
    parse_table_part(&stream(records))
}

/// `BrtBeginList` payload: two-column table over A1:B10 with a header row.
fn list_payload() -> Vec<u8> {
    list_payload_with_display_name("Sales Table")
}

/// `BrtBeginList` payload with a chosen display name.
fn list_payload_with_display_name(display_name: &str) -> Vec<u8> {
    let mut data = Vec::new();
    for value in [0u32, 9, 0, 1] {
        data.extend_from_slice(&value.to_le_bytes()); // rfxList
    }
    data.extend_from_slice(&0u32.to_le_bytes()); // lt = LTRANGE
    data.extend_from_slice(&7u32.to_le_bytes()); // idList
    data.extend_from_slice(&1u32.to_le_bytes()); // crwHeader
    data.extend_from_slice(&0u32.to_le_bytes()); // crwTotals
    data.extend_from_slice(&0b0001_1001u32.to_le_bytes()); // fShownTotalRow | fInsertRowInsCells | fPublished
    data.extend_from_slice(&3u32.to_le_bytes()); // nDxfHeader
    for _ in 0..5 {
        data.extend_from_slice(&u32::MAX.to_le_bytes()); // remaining DXFIds
    }
    data.extend_from_slice(&0u32.to_le_bytes()); // dwConnID
    data.extend_from_slice(&nullable_wide_string(Some("SalesTable"))); // stName
    data.extend_from_slice(&nullable_wide_string(Some(display_name))); // stDisplayName
    data.extend_from_slice(&nullable_wide_string(Some(" quarterly numbers "))); // stComment
    data.extend_from_slice(&nullable_wide_string(None)); // stStyleHeader
    data.extend_from_slice(&nullable_wide_string(None)); // stStyleData
    data.extend_from_slice(&nullable_wide_string(None)); // stStyleAgg
    data
}

/// `BrtBeginListCol` payload with the given totals-row function and strings.
fn column_payload(id: u32, ilta: u32, caption: &str, total: Option<&str>) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&id.to_le_bytes()); // idField
    data.extend_from_slice(&ilta.to_le_bytes()); // ilta
    data.extend_from_slice(&u32::MAX.to_le_bytes()); // nDxfHdr
    data.extend_from_slice(&u32::MAX.to_le_bytes()); // nDxfInsertRow
    data.extend_from_slice(&5u32.to_le_bytes()); // nDxfAgg
    data.extend_from_slice(&0u32.to_le_bytes()); // idqsif
    data.extend_from_slice(&nullable_wide_string(None)); // stName
    data.extend_from_slice(&nullable_wide_string(Some(caption))); // stCaption
    data.extend_from_slice(&nullable_wide_string(total)); // stTotal
    data.extend_from_slice(&nullable_wide_string(None)); // stStyleHeader
    data.extend_from_slice(&nullable_wide_string(None)); // stStyleInsertRow
    data.extend_from_slice(&nullable_wide_string(None)); // stStyleAgg
    data
}

fn minimal_table_records() -> Vec<(u16, Vec<u8>)> {
    vec![
        (rt::BEGIN_LIST, list_payload()),
        (rt::BEGIN_LIST_COLS, 2u32.to_le_bytes().to_vec()),
        (rt::BEGIN_LIST_COL, column_payload(1, 0, "Item", None)),
        (rt::END_LIST_COL, Vec::new()),
        (rt::BEGIN_LIST_COL, column_payload(2, 6, "Price", None)),
        (rt::END_LIST_COL, Vec::new()),
        (rt::END_LIST_COLS, Vec::new()),
        (rt::END_LIST, Vec::new()),
    ]
}

#[test]
fn parses_minimal_table_with_columns() {
    let table = parse(&minimal_table_records()).unwrap();
    assert_eq!(table.id, 7);
    assert_eq!(table.name.as_deref(), Some("SalesTable"));
    // The display name contains a space, which Excel rejects when creating a
    // table; the parser preserves it verbatim rather than sanitizing it.
    assert_eq!(table.display_name.as_deref(), Some("Sales Table"));
    assert_eq!(table.comment.as_deref(), Some(" quarterly numbers "));
    assert_eq!(
        table.range,
        XlsbTableRange {
            first_row: 0,
            last_row: 9,
            first_column: 0,
            last_column: 1,
        }
    );
    assert_eq!(table.table_type, XlsbTableType::Range);
    assert_eq!(table.header_row_count, 1);
    assert_eq!(table.totals_row_count, 0);
    assert!(table.totals_row_shown);
    assert!(!table.single_cell);
    assert!(!table.insert_row_visible);
    assert!(table.insert_row_inserted_cells);
    assert!(table.published);
    assert_eq!(table.header_dxf_id, Some(3));
    assert_eq!(table.data_dxf_id, None);
    assert_eq!(table.connection_id, None);
    assert_eq!(table.style_info, None);
    assert_eq!(table.alternate_text, None);

    assert_eq!(table.columns.len(), 2);
    let item = &table.columns[0];
    assert_eq!(item.id, 1);
    assert_eq!(item.totals_row_function, XlsbTableTotalsRowFunction::None);
    assert_eq!(item.caption.as_deref(), Some("Item"));
    assert_eq!(item.totals_dxf_id, Some(5));
    assert_eq!(item.query_table_field_id, 0);
    let price = &table.columns[1];
    assert_eq!(price.totals_row_function, XlsbTableTotalsRowFunction::Sum);
}

#[test]
fn parses_totals_row_functions_labels_and_formulas() {
    let mut formula = vec![0b0000_0010]; // fArray
    formula.extend_from_slice(&3u32.to_le_bytes()); // cce
    formula.extend_from_slice(&[0x1E, 0x6D, 0x00]); // rgce (inert tokens)
    formula.extend_from_slice(&0u32.to_le_bytes()); // cb

    let table = parse(&[
        (rt::BEGIN_LIST, list_payload()),
        (rt::BEGIN_LIST_COLS, 2u32.to_le_bytes().to_vec()),
        (
            rt::BEGIN_LIST_COL,
            column_payload(1, 2, "Item", Some("3 items")),
        ),
        (rt::LIST_CC_FMLA, {
            let mut data = vec![0x00]; // not an array formula
            data.extend_from_slice(&1u32.to_le_bytes());
            data.extend_from_slice(&[0x1E]);
            data.extend_from_slice(&0u32.to_le_bytes());
            data
        }),
        (rt::END_LIST_COL, Vec::new()),
        (rt::BEGIN_LIST_COL, column_payload(2, 9, "Price", None)),
        (rt::LIST_TR_FMLA, formula),
        (rt::END_LIST_COL, Vec::new()),
        (rt::END_LIST_COLS, Vec::new()),
        (rt::END_LIST, Vec::new()),
    ])
    .unwrap();

    let item = &table.columns[0];
    assert_eq!(item.totals_row_function, XlsbTableTotalsRowFunction::Count);
    assert_eq!(item.totals_row_label.as_deref(), Some("3 items"));
    let calculated = item.calculated_column_formula.as_ref().unwrap();
    assert!(!calculated.array);
    assert_eq!(calculated.tokens, [0x1E]);
    assert!(calculated.extra.is_empty());

    let price = &table.columns[1];
    assert_eq!(
        price.totals_row_function,
        XlsbTableTotalsRowFunction::Custom
    );
    assert_eq!(price.totals_row_label, None);
    let totals = price.totals_row_formula.as_ref().unwrap();
    assert!(totals.array);
    assert_eq!(totals.tokens, [0x1E, 0x6D, 0x00]);
}

#[test]
fn parses_style_info_flags_and_alternate_text() {
    let mut style = 0b0000_1101u16.to_le_bytes().to_vec(); // first column | row stripes | column stripes
    style.extend_from_slice(&nullable_wide_string(Some("TableStyleMedium9")));

    let mut alt = vec![0xFF, 0xFF, 0x00, 0x00]; // FRTBlank header
    alt.extend_from_slice(&nullable_wide_string(Some("Sales by region")));
    alt.extend_from_slice(&nullable_wide_string(Some("Pivot-like summary")));

    let table = parse(&[
        (rt::BEGIN_LIST, list_payload()),
        (rt::BEGIN_LIST_COLS, 2u32.to_le_bytes().to_vec()),
        (rt::BEGIN_LIST_COL, column_payload(1, 0, "Item", None)),
        (rt::END_LIST_COL, Vec::new()),
        (rt::BEGIN_LIST_COL, column_payload(2, 0, "Price", None)),
        (rt::END_LIST_COL, Vec::new()),
        (rt::END_LIST_COLS, Vec::new()),
        (rt::TABLE_STYLE_CLIENT, style),
        (rt::LIST14, alt),
        (rt::END_LIST, Vec::new()),
    ])
    .unwrap();

    let style = table.style_info.unwrap();
    assert_eq!(style.name.as_deref(), Some("TableStyleMedium9"));
    assert!(style.show_first_column);
    assert!(!style.show_last_column);
    assert!(style.show_row_stripes);
    assert!(style.show_column_stripes);
    assert_eq!(table.alternate_text.as_deref(), Some("Sales by region"));
    assert_eq!(
        table.alternate_text_summary.as_deref(),
        Some("Pivot-like summary")
    );
}

#[test]
fn skips_unknown_records_and_balanced_collections() {
    let mut records = minimal_table_records();
    // Unknown standalone record before the column collection.
    records.insert(1, (0x0FFF, vec![1, 2, 3]));
    // Unknown begin/end pair wrapping noise, between the columns and the end.
    records.insert(5, (0x0ABC, vec![0]));
    records.insert(6, (0x0DEF, vec![9; 8]));
    records.insert(7, (0x0ABD, Vec::new()));
    // FRT wrapper around an unknown record.
    records.push((rt::FRT_BEGIN, vec![0; 4]));
    records.push((0x0EEE, vec![0; 2]));
    records.push((rt::FRT_END, Vec::new()));
    // Records after BrtEndList are ignored entirely.
    records.push((rt::BEGIN_LIST, Vec::new()));
    let table = parse(&records).unwrap();
    assert_eq!(table.id, 7);
    assert_eq!(table.columns.len(), 2);
}

#[test]
fn skips_xml_column_properties_inside_columns() {
    let table = parse(&[
        (rt::BEGIN_LIST, list_payload()),
        (rt::BEGIN_LIST_COLS, 1u32.to_le_bytes().to_vec()),
        (rt::BEGIN_LIST_COL, column_payload(1, 0, "Item", None)),
        (rt::BEGIN_LIST_XML_CPR, vec![0; 12]),
        (rt::END_LIST_XML_CPR, Vec::new()),
        (rt::END_LIST_COL, Vec::new()),
        (rt::END_LIST_COLS, Vec::new()),
        (rt::END_LIST, Vec::new()),
    ])
    .unwrap();
    assert_eq!(table.columns.len(), 1);
    assert_eq!(table.columns[0].caption.as_deref(), Some("Item"));
}

#[test]
fn rejects_truncated_and_malformed_streams() {
    // Empty stream.
    assert!(matches!(
        parse_table_part(&[]),
        Err(XlsbError::UnexpectedEndOfStream(_))
    ));
    // Wrong first record.
    assert!(matches!(
        parse(&[(rt::END_LIST, Vec::new())]),
        Err(XlsbError::UnexpectedRecord { .. })
    ));
    // Truncated BrtBeginList payload.
    assert!(matches!(
        parse(&[(rt::BEGIN_LIST, vec![0; 8])]),
        Err(XlsbError::InvalidLength { .. })
    ));
    // Missing BrtEndList.
    assert!(matches!(
        parse(&[(rt::BEGIN_LIST, list_payload())]),
        Err(XlsbError::UnexpectedEndOfStream(_))
    ));
    // Unterminated column collection.
    assert!(matches!(
        parse(&[
            (rt::BEGIN_LIST, list_payload()),
            (rt::BEGIN_LIST_COLS, 0u32.to_le_bytes().to_vec()),
        ]),
        Err(XlsbError::UnexpectedEndOfStream(_))
    ));
    // Trailing bytes in an understood payload.
    assert!(matches!(
        parse(&[
            (rt::BEGIN_LIST, {
                let mut data = list_payload();
                data.push(0);
                data
            }),
            (rt::END_LIST, Vec::new()),
        ]),
        Err(XlsbError::Unrecognized { .. })
    ));
    // Invalid table type enumeration value.
    let mut bad_type = list_payload();
    bad_type[16] = 1; // lt = 0x00000001 is not a ListType
    assert!(matches!(
        parse(&[(rt::BEGIN_LIST, bad_type), (rt::END_LIST, Vec::new())]),
        Err(XlsbError::Unrecognized { .. })
    ));
    // Invalid totals-row function enumeration value.
    assert!(matches!(
        parse(&[
            (rt::BEGIN_LIST, list_payload()),
            (rt::BEGIN_LIST_COLS, 1u32.to_le_bytes().to_vec()),
            (rt::BEGIN_LIST_COL, column_payload(1, 0x0A, "Item", None)),
            (rt::END_LIST_COL, Vec::new()),
            (rt::END_LIST_COLS, Vec::new()),
            (rt::END_LIST, Vec::new()),
        ]),
        Err(XlsbError::Unrecognized { .. })
    ));
    // Non-Boolean crwHeader.
    let mut bad_flag = list_payload();
    bad_flag[24] = 2;
    assert!(matches!(
        parse(&[(rt::BEGIN_LIST, bad_flag), (rt::END_LIST, Vec::new())]),
        Err(XlsbError::Unrecognized { .. })
    ));
    // Declared column count disagrees with the record collection.
    assert!(matches!(
        parse(&[
            (rt::BEGIN_LIST, list_payload()),
            (rt::BEGIN_LIST_COLS, 3u32.to_le_bytes().to_vec()),
            (rt::BEGIN_LIST_COL, column_payload(1, 0, "Item", None)),
            (rt::END_LIST_COL, Vec::new()),
            (rt::END_LIST_COLS, Vec::new()),
            (rt::END_LIST, Vec::new()),
        ]),
        Err(XlsbError::Unrecognized { .. })
    ));
}

#[test]
fn tolerates_out_of_order_collections() {
    // The style record precedes the column collection.
    let mut style = 0b0000_0100u16.to_le_bytes().to_vec(); // row stripes only
    style.extend_from_slice(&nullable_wide_string(None));
    let mut records = minimal_table_records();
    records.insert(1, (rt::TABLE_STYLE_CLIENT, style));
    let table = parse(&records).unwrap();
    assert_eq!(table.columns.len(), 2);
    assert!(table.style_info.unwrap().show_row_stripes);
}

#[test]
fn extracts_table_part_relationship_ids_from_worksheet_stream() {
    let rel_ids = parse_table_part_rel_ids(&stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (0x0FFF, vec![0; 3]), // unknown record before the collection
        (rt::BEGIN_LIST_PARTS, 2u32.to_le_bytes().to_vec()),
        (rt::LIST_PART, wide_string("rIdTable1")),
        (0x0ABC, vec![0]), // unknown record inside the collection
        (rt::LIST_PART, wide_string("rIdTable2")),
        (rt::END_LIST_PARTS, Vec::new()),
        (rt::END_SHEET, Vec::new()),
    ]))
    .unwrap();
    assert_eq!(rel_ids, ["rIdTable1".to_string(), "rIdTable2".to_string()]);

    // Declared count disagreement is rejected.
    assert!(matches!(
        parse_table_part_rel_ids(&stream(&[
            (rt::BEGIN_LIST_PARTS, 2u32.to_le_bytes().to_vec()),
            (rt::LIST_PART, wide_string("rIdTable1")),
            (rt::END_LIST_PARTS, Vec::new()),
        ])),
        Err(XlsbError::Unrecognized { .. })
    ));
    // Unterminated collection is rejected.
    assert!(matches!(
        parse_table_part_rel_ids(&stream(&[(
            rt::BEGIN_LIST_PARTS,
            0u32.to_le_bytes().to_vec()
        )])),
        Err(XlsbError::UnexpectedEndOfStream(_))
    ));
}

/// Build a synthetic package with one worksheet that references one table
/// part through its `BrtBeginListParts` collection, and verify the workbook
/// accessors.
#[test]
fn resolves_tables_through_workbook_relationships() {
    use crate::xlsb::XlsbWorkbook;
    use litchi_opc::constants::relationship_type;
    use litchi_opc::part::Part;
    use litchi_opc::{BlobPart, OpcPackage, PackURI};

    // workbook.bin declares one worksheet.
    let mut bundle_sheet = 0u32.to_le_bytes().to_vec(); // hsState = visible
    bundle_sheet.extend_from_slice(&1u32.to_le_bytes()); // iTabID
    bundle_sheet.extend_from_slice(&wide_string("rIdSheet1"));
    bundle_sheet.extend_from_slice(&wide_string("Sheet1"));
    let workbook_data = stream(&[(rt::BUNDLE_SH, bundle_sheet)]);
    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.bin").unwrap(),
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
        workbook_data,
    );
    workbook_part.rels_mut().add_relationship(
        relationship_type::WORKSHEET.to_string(),
        "worksheets/sheet1.bin".to_string(),
        "rIdSheet1".to_string(),
        false,
    );

    // The worksheet references one table part.
    let sheet_data = stream(&[
        (rt::BEGIN_SHEET, Vec::new()),
        (rt::BEGIN_LIST_PARTS, 1u32.to_le_bytes().to_vec()),
        (rt::LIST_PART, wide_string("rIdTable1")),
        (rt::END_LIST_PARTS, Vec::new()),
        (rt::END_SHEET, Vec::new()),
    ]);
    let mut sheet_part = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.bin").unwrap(),
        "application/vnd.ms-excel.worksheet".to_string(),
        sheet_data,
    );
    sheet_part.rels_mut().add_relationship(
        relationship_type::TABLE.to_string(),
        "../tables/table1.bin".to_string(),
        "rIdTable1".to_string(),
        false,
    );

    // Use a display name without spaces: the workbook-level formula parser
    // rejects names that violate XLNameWideString grammar, while the typed
    // table parser preserves them verbatim (covered by the parser tests).
    let mut table_records = minimal_table_records();
    table_records[0] = (rt::BEGIN_LIST, list_payload_with_display_name("SalesTable"));
    let table_part = BlobPart::new(
        PackURI::new("/xl/tables/table1.bin").unwrap(),
        "application/vnd.ms-excel.table".to_string(),
        stream(&table_records),
    );

    let mut package = OpcPackage::new();
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(sheet_part));
    package.add_part(Box::new(table_part));
    let workbook = XlsbWorkbook::from_opc_package(package).unwrap();

    let tables = workbook.structured_tables();
    assert_eq!(tables.len(), 1);
    assert_eq!(tables[0].0, 0);
    assert_eq!(tables[0].1.id, 7);
    assert_eq!(tables[0].1.columns.len(), 2);

    let on_sheet = workbook.tables_on_sheet(0);
    assert_eq!(on_sheet.len(), 1);
    assert_eq!(on_sheet[0].display_name.as_deref(), Some("SalesTable"));
    assert!(workbook.tables_on_sheet(1).is_empty());

    // A sheet that declares a broken table part surfaces the error, matching
    // the eager failure handling of PivotCache definitions.
    let mut sheet_part = BlobPart::new(
        PackURI::new("/xl/worksheets/sheet1.bin").unwrap(),
        "application/vnd.ms-excel.worksheet".to_string(),
        stream(&[
            (rt::BEGIN_LIST_PARTS, 1u32.to_le_bytes().to_vec()),
            (rt::LIST_PART, wide_string("rIdTable1")),
            (rt::END_LIST_PARTS, Vec::new()),
        ]),
    );
    sheet_part.rels_mut().add_relationship(
        relationship_type::TABLE.to_string(),
        "../tables/table1.bin".to_string(),
        "rIdTable1".to_string(),
        false,
    );
    let broken_table = BlobPart::new(
        PackURI::new("/xl/tables/table1.bin").unwrap(),
        "application/vnd.ms-excel.table".to_string(),
        stream(&[(rt::BEGIN_LIST, vec![0; 8])]),
    );
    let mut bundle_sheet = 0u32.to_le_bytes().to_vec();
    bundle_sheet.extend_from_slice(&1u32.to_le_bytes());
    bundle_sheet.extend_from_slice(&wide_string("rIdSheet1"));
    bundle_sheet.extend_from_slice(&wide_string("Sheet1"));
    let mut workbook_part = BlobPart::new(
        PackURI::new("/xl/workbook.bin").unwrap(),
        "application/vnd.ms-excel.sheet.binary.macroEnabled.main".to_string(),
        stream(&[(rt::BUNDLE_SH, bundle_sheet)]),
    );
    workbook_part.rels_mut().add_relationship(
        relationship_type::WORKSHEET.to_string(),
        "worksheets/sheet1.bin".to_string(),
        "rIdSheet1".to_string(),
        false,
    );
    let mut package = OpcPackage::new();
    package.add_part(Box::new(workbook_part));
    package.add_part(Box::new(sheet_part));
    package.add_part(Box::new(broken_table));
    assert!(XlsbWorkbook::from_opc_package(package).is_err());
}
