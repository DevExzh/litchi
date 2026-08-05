use std::io::Cursor;

use litchi_xls::writer::sort::{
    Axis, CONTINUE_FRT12_RECORD_TYPE, Config, Dxf, Icon, IconSet, Key, Method, On, Parent, Range,
    SORT_DATA_RECORD_TYPE,
};
use litchi_xls::{Workbook, Writer};

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

fn extended_sort() -> Config {
    let range = Range::new(1..=20, 0..=4).unwrap();
    let mut sort = Config::new(range, Parent::Sheet);
    sort.put_axis(Axis::Cols).unwrap();
    sort.set_case(true);
    sort.set_method(Method::Alternate);
    sort.add(
        Key::row(
            Range::new(1..=1, 0..=4).unwrap(),
            true,
            On::Values {
                custom_list: Some("高,中,低".to_string()),
            },
        )
        .unwrap(),
    )
    .unwrap();
    sort.add(
        Key::row(
            Range::new(2..=2, 0..=4).unwrap(),
            false,
            On::CellColor {
                differential_format: Dxf::new(9),
            },
        )
        .unwrap(),
    )
    .unwrap();
    sort.add(
        Key::row(
            Range::new(3..=3, 0..=4).unwrap(),
            false,
            On::Icon {
                set: IconSet::FiveRating,
                icon: Icon::Fourth,
            },
        )
        .unwrap(),
    )
    .unwrap();
    sort
}

#[test]
fn writer_stream_round_trip_preserves_extended_sort_and_record_order() {
    let expected = extended_sort();
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Sorted").unwrap();
    writer.set_sort(sheet, true, true, &[(0, true)]).unwrap();
    assert_eq!(writer.put_sort(sheet, expected.clone()).unwrap(), None);

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

    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    let parsed = workbook.xls_worksheet(0).unwrap().sort().unwrap();
    assert_eq!(parsed, &expected);
}

#[test]
fn writer_and_parser_support_zero_condition_sort_data() {
    let expected = Config::new(Range::new(0..=0, 0..=0).unwrap(), Parent::AutoFilter);
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Empty sort").unwrap();
    assert_eq!(writer.put_sort(sheet, expected.clone()).unwrap(), None);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(workbook.xls_worksheet(0).unwrap().sort(), Some(&expected));
}

#[test]
fn writer_sort_crud_is_move_first_idempotent_and_failure_atomic() {
    let first = Config::new(Range::new(0..=5, 0..=1).unwrap(), Parent::Sheet);
    let mut second = Config::new(Range::new(2..=9, 1..=2).unwrap(), Parent::Sheet);
    second
        .add(
            Key::col(
                Range::new(2..=9, 1..=1).unwrap(),
                false,
                On::Values { custom_list: None },
            )
            .unwrap(),
        )
        .unwrap();

    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("CRUD").unwrap();
    assert_eq!(writer.put_sort(sheet, first.clone()).unwrap(), None);
    assert_eq!(
        writer.put_sort(sheet, second.clone()).unwrap(),
        Some(first.clone())
    );

    let invalid_sheet = sheet + 1;
    assert!(writer.put_sort(invalid_sheet, first.clone()).is_err());
    assert!(writer.remove_sort(invalid_sheet).is_err());
    let unknown_table = Config::new(Range::new(0..=0, 0..=0).unwrap(), Parent::Table { id: 99 });
    assert!(writer.put_sort(sheet, unknown_table).is_err());
    assert_eq!(writer.remove_sort(sheet).unwrap(), Some(second));
    assert_eq!(writer.remove_sort(sheet).unwrap(), None);

    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert_eq!(workbook.xls_worksheet(0).unwrap().sort(), None);
}
