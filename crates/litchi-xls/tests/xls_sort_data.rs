use std::io::Cursor;

use litchi_xls::{
    CONTINUE_FRT12_RECORD_TYPE, SORT_DATA_RECORD_TYPE, XlsDifferentialFormatIndex,
    XlsSortCondition, XlsSortData, XlsSortIcon, XlsSortIconSet, XlsSortMethod, XlsSortOn,
    XlsSortOrientation, XlsSortParent, XlsSortRange, XlsWorkbook, XlsWriter,
};

const LEGACY_SORT_RECORD_TYPE: u16 = 0x0090;

fn workbook_record_types(bytes: &[u8]) -> Vec<u16> {
    let mut ole = litchi_cfb::OleFile::open(Cursor::new(bytes)).unwrap();
    let stream = ole.open_stream(&["Workbook"]).unwrap();
    let mut record_types = Vec::new();
    let mut offset = 0usize;
    while offset + 4 <= stream.len() {
        let record_type = u16::from_le_bytes([stream[offset], stream[offset + 1]]);
        let length = usize::from(u16::from_le_bytes([stream[offset + 2], stream[offset + 3]]));
        offset += 4 + length;
        record_types.push(record_type);
    }
    record_types
}

fn extended_sort() -> XlsSortData {
    let range = XlsSortRange::new(1, 20, 0, 4).unwrap();
    let mut sort = XlsSortData::new(range, XlsSortParent::Sheet);
    sort.set_orientation(XlsSortOrientation::Columns);
    sort.set_case_sensitive(true);
    sort.set_method(XlsSortMethod::Alternate);
    sort.add_condition(XlsSortCondition::new(
        XlsSortRange::new(1, 20, 0, 0).unwrap(),
        true,
        XlsSortOn::Values {
            custom_list: Some("高,中,低".to_string()),
        },
    ));
    sort.add_condition(XlsSortCondition::new(
        XlsSortRange::new(1, 20, 1, 1).unwrap(),
        false,
        XlsSortOn::CellColor {
            differential_format: XlsDifferentialFormatIndex::new(9),
        },
    ));
    sort.add_condition(XlsSortCondition::new(
        XlsSortRange::new(1, 20, 2, 2).unwrap(),
        false,
        XlsSortOn::Icon {
            set: XlsSortIconSet::FiveRating,
            icon: XlsSortIcon::Fourth,
        },
    ));
    sort
}

#[test]
fn writer_stream_round_trip_preserves_extended_sort_and_record_order() {
    let expected = extended_sort();
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Sorted").unwrap();
    writer.set_sort(sheet, true, true, &[(0, true)]).unwrap();
    writer.set_sort_data(sheet, expected.clone()).unwrap();

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let bytes = output.into_inner();
    let record_types = workbook_record_types(&bytes);
    let sort_position = record_types
        .iter()
        .position(|record_type| *record_type == LEGACY_SORT_RECORD_TYPE)
        .unwrap();
    assert_eq!(record_types[sort_position + 1], SORT_DATA_RECORD_TYPE);
    assert_eq!(
        &record_types[sort_position + 2..sort_position + 5],
        &[CONTINUE_FRT12_RECORD_TYPE; 3]
    );

    let workbook = XlsWorkbook::new(Cursor::new(bytes)).unwrap();
    let parsed = workbook.xls_worksheet(0).unwrap().sort_data().unwrap();
    assert_eq!(parsed, &expected);
}

#[test]
fn writer_and_parser_support_zero_condition_sort_data() {
    let expected = XlsSortData::new(
        XlsSortRange::new(0, 0, 0, 0).unwrap(),
        XlsSortParent::AutoFilter,
    );
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Empty sort").unwrap();
    writer.set_sort_data(sheet, expected.clone()).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(
        workbook.xls_worksheet(0).unwrap().sort_data(),
        Some(&expected)
    );
}
