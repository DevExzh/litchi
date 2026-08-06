use std::io::Cursor;

use litchi_ole_common::toolbar::{Flags, Header, Restrictions, Type, WString};
use litchi_xls::{Toolbar, Workbook, Wrapper, Writer};

fn toolbar() -> Toolbar<'static> {
    Toolbar::new(
        Header::new(
            0,
            Restrictions::new(Type::Basic).unwrap(),
            1,
            Flags::from_raw(0),
            WString::new("XCB").unwrap(),
        )
        .unwrap(),
        0x0000_0FFF,
    )
    .unwrap()
}

#[test]
fn writer_and_workbook_facades_round_trip_xcb_stream() {
    let mut writer = Writer::new();
    writer.add_worksheet("Sheet1").unwrap();
    writer
        .set_toolbar(Wrapper::new(vec![toolbar()]).unwrap())
        .unwrap();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();

    let mut workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    let decoded = workbook.toolbar().unwrap().unwrap();
    assert_eq!(decoded.toolbars().len(), 1);
    assert_eq!(decoded.toolbars()[0].header().name().text(), "XCB");
}

#[test]
fn writer_can_clear_xcb_stream() {
    let mut writer = Writer::new();
    writer.add_worksheet("Sheet1").unwrap();
    writer
        .set_toolbar(Wrapper::new(vec![toolbar()]).unwrap())
        .unwrap();
    writer.clear_toolbar();

    let mut bytes = Cursor::new(Vec::new());
    writer.write_to(&mut bytes).unwrap();
    let mut workbook = Workbook::new(Cursor::new(bytes.into_inner())).unwrap();
    assert!(workbook.toolbar().unwrap().is_none());
}
