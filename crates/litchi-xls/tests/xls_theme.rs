//! Round-trip tests for the BIFF8 Theme record.

use litchi_xls::writer::Writer;
use litchi_xls::{Theme, Workbook};
use std::io::Cursor;

fn written_workbook(theme: Option<Theme>) -> Vec<u8> {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Theme").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_theme(theme);
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn default_theme_round_trips() {
    let workbook =
        Workbook::new(Cursor::new(written_workbook(Some(Theme::default_theme())))).unwrap();
    let theme = workbook.theme().expect("Theme record present");
    assert!(!theme.is_custom());
    assert_eq!(theme.version(), 124_226);
    assert_eq!(theme.contents(), None);
}

#[test]
fn large_custom_theme_round_trips_across_continuations() {
    // Exceeds one BIFF8 record, forcing ContinueFrt12 chunking.
    let contents: Vec<u8> = (0..40_000u32).flat_map(u32::to_le_bytes).collect();
    let workbook = Workbook::new(Cursor::new(written_workbook(Some(
        Theme::custom(contents.clone()).unwrap(),
    ))))
    .unwrap();
    let theme = workbook.theme().expect("Theme record present");
    assert!(theme.is_custom());
    assert_eq!(theme.contents(), Some(contents.as_slice()));
}

#[test]
fn workbook_without_theme_has_none() {
    let workbook = Workbook::new(Cursor::new(written_workbook(None))).unwrap();
    assert!(workbook.theme().is_none());
}
