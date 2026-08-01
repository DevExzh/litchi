use std::fs::File;
use std::io::Cursor;
use std::path::PathBuf;

use litchi_core::sheet::WorkbookTrait;
use litchi_xls::writer::{XlsCustomTableStyles, XlsWriter};
use litchi_xls::{
    XlsListColumnId, XlsListObject, XlsListObjectColumn, XlsListObjectFeatureVersion,
    XlsListObjectId, XlsListObjectRange, XlsListObjectStyleOptions, XlsListTotalAggregation,
    XlsSortCondition, XlsSortData, XlsSortOn, XlsSortParent, XlsSortRange, XlsTableStyle,
    XlsWorkbook,
};

fn workbook_records(bytes: &[u8]) -> Vec<(u16, usize)> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole.open_stream(&["Workbook"]).unwrap();
    let mut records = Vec::new();
    let mut offset = 0usize;
    while offset + 4 <= stream.len() {
        let record_type = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        records.push((record_type, length));
        offset += 4 + length;
    }
    records
}

fn column(id: u32, name: &str) -> XlsListObjectColumn {
    XlsListObjectColumn::try_new(XlsListColumnId::try_new(id).unwrap(), name).unwrap()
}

fn table(id: u32, name: &str, first_row: u16, last_row: u16, style: &str) -> XlsListObject {
    XlsListObject::try_new(
        XlsListObjectId::try_new(id).unwrap(),
        name,
        XlsListObjectRange::try_new(first_row, last_row, 0, 1).unwrap(),
        vec![column(1, "Region"), column(2, "Sales")],
        XlsListObjectStyleOptions::try_new(style).unwrap(),
    )
    .unwrap()
}

#[test]
fn writer_custom_list_object_round_trips_through_reader() {
    let custom_style = XlsTableStyle::try_new("WriterTable", true, false, Vec::new()).unwrap();
    let styles = XlsCustomTableStyles::try_from_styles(
        Vec::new(),
        "WriterTable",
        "PivotStyleLight16",
        vec![custom_style],
    )
    .unwrap();
    let mut writer = XlsWriter::new();
    writer.set_custom_table_styles(styles).unwrap();
    let sheet = writer.add_worksheet("Data").unwrap();
    writer
        .add_list_object(sheet, table(7, "SalesTable", 0, 3, "WriterTable"))
        .unwrap();
    writer.write_string(sheet, 1, 0, "East").unwrap();
    writer.write_number(sheet, 1, 1, 42.0).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let table = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(table.id().value(), 7);
    assert_eq!(table.name(), "SalesTable");
    assert_eq!(table.columns()[1].name(), "Sales");
    assert_eq!(table.style().unwrap().name(), "WriterTable");
    assert!(table.style().unwrap().shows_row_stripes());
}

#[test]
fn reads_excel_producer_feature11_and_list12_tables() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet/ConditionalFormattingSamples.xls");
    let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
    let mut found = None;
    let mut observed = Vec::new();
    for index in 0..workbook.worksheet_count() {
        let worksheet = workbook.xls_worksheet(index).unwrap();
        observed.extend(
            worksheet
                .list_objects()
                .iter()
                .map(|table| table.name().to_string()),
        );
        if let Some(table) = worksheet
            .list_objects()
            .iter()
            .find(|table| table.name() == "Table6")
        {
            found = Some(table.clone());
            break;
        }
    }
    let table = found.unwrap_or_else(|| panic!("producer Table6; observed {observed:?}"));
    assert_eq!(
        table.range(),
        XlsListObjectRange::try_new(2, 11, 0, 1).unwrap()
    );
    assert_eq!(table.style().unwrap().name(), "TableStyleMedium9");
    assert_eq!(table.columns()[0].name(), "Region");
}

#[test]
fn writer_rejects_table_collisions_ranges_and_style_references_atomically() {
    assert!(
        XlsListObject::try_new(
            XlsListObjectId::try_new(1).unwrap(),
            "Bad",
            XlsListObjectRange::try_new(0, 2, 0, 1).unwrap(),
            vec![column(1, "A"), column(1, "B")],
            XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
        )
        .is_err()
    );
    let mut writer = XlsWriter::new();
    let first = writer.add_worksheet("One").unwrap();
    let second = writer.add_worksheet("Two").unwrap();
    writer
        .add_list_object(first, table(1, "First", 0, 3, "TableStyleMedium2"))
        .unwrap();
    assert!(
        writer
            .add_list_object(first, table(2, "Overlap", 2, 5, "TableStyleMedium2"))
            .is_err()
    );
    assert!(
        writer
            .add_list_object(second, table(1, "Other", 0, 3, "TableStyleMedium2"))
            .is_err()
    );
    assert!(
        writer
            .add_list_object(second, table(3, "First", 0, 3, "TableStyleMedium2"))
            .is_err()
    );
    assert!(
        writer
            .add_list_object(second, table(3, "Unknown", 0, 3, "NotConfigured"))
            .is_err()
    );
    assert!(writer.write_to(&mut Cursor::new(Vec::new())).is_ok());
}

#[test]
fn long_feature11_uses_strict_continue_frt11_chain_and_round_trips() {
    let columns = (0..256)
        .map(|index| column(index + 1, &format!("Column_{index:03}_{}", "x".repeat(220))))
        .collect::<Vec<_>>();
    let table = XlsListObject::try_new(
        XlsListObjectId::try_new(19).unwrap(),
        "WideTable",
        XlsListObjectRange::try_new(0, 2, 0, 255).unwrap(),
        columns,
        XlsListObjectStyleOptions::try_new("TableStyleMedium4").unwrap(),
    )
    .unwrap();
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Wide").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    let records = workbook_records(&bytes);
    let feature = records.iter().position(|v| v.0 == 0x0872).unwrap();
    assert_eq!(records[feature].1, 8_224);
    assert!(
        records[feature + 1..]
            .iter()
            .take_while(|v| v.0 == 0x0875)
            .count()
            > 1
    );
    let workbook = XlsWorkbook::new(Cursor::new(bytes)).unwrap();
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().list_objects()[0]
            .columns()
            .len(),
        256
    );
}

#[test]
fn headerless_internal_feature12_and_table_sort_continuations_round_trip_in_order() {
    let value = table(23, "Headerless", 0, 3, "TableStyleLight8")
        .with_header_row(false)
        .unwrap();
    let mut sort = XlsSortData::new(
        XlsSortRange::new(0, 3, 0, 1).unwrap(),
        XlsSortParent::Table { id: 23 },
    );
    sort.add_condition(XlsSortCondition::new(
        XlsSortRange::new(0, 3, 0, 0).unwrap(),
        false,
        XlsSortOn::Values {
            custom_list: Some("East,West".to_string()),
        },
    ));
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Extended").unwrap();
    writer.add_list_object(sheet, value).unwrap();
    writer.set_sort_data(sheet, sort.clone()).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    let types = workbook_records(&bytes)
        .into_iter()
        .map(|v| v.0)
        .collect::<Vec<_>>();
    let feature = types.iter().position(|v| *v == 0x0878).unwrap();
    let sort_position = types.iter().position(|v| *v == 0x0895).unwrap();
    let last_list12 = types
        .iter()
        .enumerate()
        .filter(|(_, v)| **v == 0x0877)
        .map(|v| v.0)
        .max()
        .unwrap();
    assert!(feature < last_list12 && last_list12 < sort_position);
    assert_eq!(types[sort_position + 1], 0x087f);
    let workbook = XlsWorkbook::new(Cursor::new(bytes)).unwrap();
    let worksheet = workbook.xls_worksheet(0).unwrap();
    assert_eq!(
        worksheet.list_objects()[0].feature_version(),
        XlsListObjectFeatureVersion::Feature12
    );
    assert!(!worksheet.list_objects()[0].has_header_row());
    assert_eq!(worksheet.sort_data(), Some(&sort));
}

#[test]
fn producer_table_sort_continue_frt12_chain_remains_attached() {
    let fixture = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/poi/test-data/spreadsheet/ConditionalFormattingSamples.xls");
    let workbook = XlsWorkbook::new(File::open(fixture).unwrap()).unwrap();
    let mut found = false;
    for index in 0..workbook.worksheet_count() {
        let worksheet = workbook.xls_worksheet(index).unwrap();
        if let Some(sort) = worksheet.sort_data()
            && let XlsSortParent::Table { id } = sort.parent()
        {
            assert!(
                worksheet
                    .list_objects()
                    .iter()
                    .any(|table| table.id().value() == id)
            );
            assert!(!sort.conditions().is_empty());
            found = true;
        }
    }
    assert!(
        found,
        "expected producer table SortData/ContinueFrt12 chain"
    );
}

#[test]
fn feature12_total_formula_string_and_aggregation_round_trip_inertly() {
    let columns = vec![
        column(1, "Region")
            .with_total_string("Grand total")
            .unwrap(),
        column(2, "Sales")
            .with_total_formula_tokens(vec![0x1e, 1, 0])
            .unwrap(),
    ];
    let value = XlsListObject::try_new(
        XlsListObjectId::try_new(31).unwrap(),
        "TotalsTable",
        XlsListObjectRange::try_new(0, 3, 0, 1).unwrap(),
        columns,
        XlsListObjectStyleOptions::try_new("TableStyleMedium6").unwrap(),
    )
    .unwrap()
    .with_totals_row(true)
    .unwrap();
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Totals").unwrap();
    writer.add_list_object(sheet, value).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let parsed = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(
        parsed.feature_version(),
        XlsListObjectFeatureVersion::Feature12
    );
    assert_eq!(parsed.columns()[0].total_string(), Some("Grand total"));
    assert_eq!(
        parsed.columns()[1].total_aggregation(),
        XlsListTotalAggregation::Custom
    );
    assert_eq!(
        parsed.columns()[1].total_formula_tokens(),
        Some(&[0x1e, 1, 0][..])
    );
}
