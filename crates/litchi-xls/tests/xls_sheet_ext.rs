//! Round-trip tests for the BIFF8 SheetExt record (sheet tab color).

use litchi_xls::Workbook;
use litchi_xls::writer::Writer;
use std::io::Cursor;

fn written_workbook(configure: impl FnOnce(&mut Writer)) -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Tab").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    configure(&mut writer);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn tab_color_round_trips_through_writer_and_reader() {
    let bytes = written_workbook(|writer| {
        writer.set_worksheet_tab_color(0, Some(0x0A)).unwrap();
    });
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    let sheet_ext = workbook
        .xls_worksheet(0)
        .unwrap()
        .sheet_ext()
        .expect("SheetExt record present");
    assert_eq!(sheet_ext.tab_color(), Some(0x0A));
    assert!(sheet_ext.optional().is_none());
}

#[test]
fn workbook_without_tab_color_has_no_sheet_ext() {
    let bytes = written_workbook(|_| {});
    let workbook = Workbook::new(Cursor::new(bytes)).unwrap();
    assert!(workbook.xls_worksheet(0).unwrap().sheet_ext().is_none());
}

#[test]
fn tab_color_validation_rejects_out_of_palette_indices() {
    let mut writer = Writer::new();
    writer.add_worksheet("Tab").unwrap();
    assert!(writer.set_worksheet_tab_color(0, Some(0x05)).is_err());
    assert!(writer.set_worksheet_tab_color(0, Some(0x40)).is_err());
    assert!(writer.set_worksheet_tab_color(0, Some(0x7F)).is_err());
    assert!(writer.set_worksheet_tab_color(9, Some(0x0A)).is_err());
    // Boundary indices and clearing are accepted.
    writer.set_worksheet_tab_color(0, Some(0x08)).unwrap();
    writer.set_worksheet_tab_color(0, Some(0x3F)).unwrap();
    writer.set_worksheet_tab_color(0, None).unwrap();
}
