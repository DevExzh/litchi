use std::{fs::File, io::Cursor, path::PathBuf};

use litchi_xls::{
    XlsAutoFilter12Criterion, XlsAutoFilter12Icon, XlsAutoFilter12IconSet, XlsAutoFilter12Operator,
    XlsAutoFilter12Value, XlsListColumnId, XlsListObject, XlsListObjectColumn, XlsListObjectId,
    XlsListObjectRange, XlsListObjectStyleOptions, XlsTableAutoFilter12, XlsWorkbook, XlsWriter,
};

#[test]
fn table_autofilter12_scalar_criteria_writer_reader_round_trip_inertly() {
    let filter = XlsTableAutoFilter12::try_new(
        1,
        vec![
            XlsAutoFilter12Criterion::try_new(
                XlsAutoFilter12Operator::GreaterThanOrEqual,
                XlsAutoFilter12Value::Number(12.5),
            )
            .unwrap(),
            XlsAutoFilter12Criterion::try_new(
                XlsAutoFilter12Operator::Equal,
                XlsAutoFilter12Value::String("POI*".to_string()),
            )
            .unwrap(),
            XlsAutoFilter12Criterion::try_new(
                XlsAutoFilter12Operator::NotEqual,
                XlsAutoFilter12Value::Boolean(false),
            )
            .unwrap(),
            XlsAutoFilter12Criterion::try_new(
                XlsAutoFilter12Operator::Equal,
                XlsAutoFilter12Value::NonBlanks,
            )
            .unwrap(),
        ],
    )
    .unwrap()
    .with_hidden_arrow(true);
    let table = XlsListObject::try_new(
        XlsListObjectId::try_new(23).unwrap(),
        "ProducerData",
        XlsListObjectRange::try_new(0, 8, 0, 1).unwrap(),
        vec![
            XlsListObjectColumn::try_new(XlsListColumnId::try_new(1).unwrap(), "Name").unwrap(),
            XlsListObjectColumn::try_new(XlsListColumnId::try_new(2).unwrap(), "Value").unwrap(),
        ],
        XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_autofilter12_criteria(filter.clone())
    .unwrap();

    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Filters").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let parsed = &workbook.xls_worksheet(0).unwrap().list_objects()[0];
    assert_eq!(parsed.autofilter12_criteria(), Some(&filter));
    assert!(parsed.opaque_future_records().is_empty());
}

#[test]
fn apache_poi_icon_filter_table_round_trips_through_writer() {
    let path = PathBuf::from(env!("CARGO_MANIFEST_DIR"))
        .join("../../test-data/ole/xls/ConditionalFormattingSamples.xls");
    let producer = XlsWorkbook::new(File::open(path).unwrap()).unwrap();
    let table = (0..producer.sheets().len())
        .filter_map(|sheet| producer.xls_worksheet(sheet).ok())
        .flat_map(|sheet| sheet.list_objects())
        .find(|table| table.range().column_count() >= 2 && table.opaque_future_records().is_empty())
        .cloned()
        .expect("Apache POI table suitable for icon-filter producer metadata");
    let expected =
        XlsAutoFilter12Icon::try_new(XlsAutoFilter12IconSet::ThreeTrafficLights1, 1).unwrap();
    let table = XlsListObject::try_new(
        table.id(),
        table.name(),
        table.range(),
        table.columns().to_vec(),
        XlsListObjectStyleOptions::try_new("TableStyleMedium2").unwrap(),
    )
    .unwrap()
    .with_autofilter12_criteria(XlsTableAutoFilter12::try_new_icon(1, expected).unwrap())
    .unwrap();

    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("POI icon filter").unwrap();
    writer.add_list_object(sheet, table).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().list_objects()[0]
            .autofilter12_criteria()
            .unwrap()
            .icon_filter(),
        Some(expected),
    );
}
