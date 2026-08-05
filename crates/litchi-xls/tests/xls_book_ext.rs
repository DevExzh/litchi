//! Round-trip tests for the BIFF8 BookExt record (workbook extension flags).

use litchi_xls::writer::Writer;
use litchi_xls::{BookExt, BookExtConditional11, BookExtConditional12, FactoidDisplay, Workbook};
use std::io::Cursor;

fn written_workbook(book_ext: Option<BookExt>) -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Flags").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_book_ext(book_ext);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn book_ext_round_trips_through_writer_and_reader() {
    let book_ext = BookExt {
        dont_auto_recover: true,
        filter_privacy: true,
        embed_factoids: true,
        factoid_display: FactoidDisplay::ButtonOnly,
        saved_during_recovery: true,
        conditional11: Some(BookExtConditional11 {
            bugged_user_about_solution: false,
            show_ink_annotation: true,
        }),
        conditional12: Some(BookExtConditional12 {
            published_book_items: true,
        }),
        ..BookExt::default()
    };
    let workbook = Workbook::new(Cursor::new(written_workbook(Some(book_ext.clone())))).unwrap();
    assert_eq!(workbook.book_ext(), Some(&book_ext));
}

#[test]
fn workbook_without_book_ext_has_none() {
    let workbook = Workbook::new(Cursor::new(written_workbook(None))).unwrap();
    assert!(workbook.book_ext().is_none());
}
