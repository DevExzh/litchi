//! Round-trip tests for the BIFF8 BookExt record (workbook extension flags).

use litchi_xls::writer::XlsWriter;
use litchi_xls::{
    XlsBookExt, XlsBookExtConditional11, XlsBookExtConditional12, XlsFactoidDisplay, XlsWorkbook,
};
use std::io::Cursor;

fn written_workbook(book_ext: Option<XlsBookExt>) -> Vec<u8> {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Flags").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_book_ext(book_ext);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn book_ext_round_trips_through_writer_and_reader() {
    let book_ext = XlsBookExt {
        dont_auto_recover: true,
        filter_privacy: true,
        embed_factoids: true,
        factoid_display: XlsFactoidDisplay::ButtonOnly,
        saved_during_recovery: true,
        conditional11: Some(XlsBookExtConditional11 {
            bugged_user_about_solution: false,
            show_ink_annotation: true,
        }),
        conditional12: Some(XlsBookExtConditional12 {
            published_book_items: true,
        }),
        ..XlsBookExt::default()
    };
    let workbook = XlsWorkbook::new(Cursor::new(written_workbook(Some(book_ext.clone())))).unwrap();
    assert_eq!(workbook.book_ext(), Some(&book_ext));
}

#[test]
fn workbook_without_book_ext_has_none() {
    let workbook = XlsWorkbook::new(Cursor::new(written_workbook(None))).unwrap();
    assert!(workbook.book_ext().is_none());
}
