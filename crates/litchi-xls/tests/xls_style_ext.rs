//! Round-trip tests for the BIFF8 `StyleExt` record (cell-style extensions).

use litchi_xls::writer::Writer;
use litchi_xls::{StyleCategory, StyleExt, Workbook};
use litchi_xls::{XfProperties, XfProperty};
use std::io::Cursor;

fn extensions() -> Vec<StyleExt> {
    let mut heading = StyleExt::try_new(
        true,
        StyleCategory::TitleAndHeading,
        "Heading 1".to_string(),
        XfProperties::try_new(vec![XfProperty::FontItalic(true)]).unwrap(),
    )
    .unwrap();
    heading.set_hidden(true);
    let custom = StyleExt::try_new(
        false,
        StyleCategory::Custom,
        "My Style".to_string(),
        XfProperties::default(),
    )
    .unwrap();
    vec![heading, custom]
}

#[test]
fn style_extensions_round_trip_through_writer_and_reader() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Styles").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_style_extensions(extensions());
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let parsed = workbook.style_extensions();
    assert_eq!(parsed.len(), 2);
    assert_eq!(parsed[0].name(), "Heading 1");
    assert!(parsed[0].built_in());
    assert!(parsed[0].hidden());
    assert_eq!(parsed[0].category(), StyleCategory::TitleAndHeading);
    assert_eq!(
        parsed[0].properties().properties(),
        &[XfProperty::FontItalic(true)]
    );
    assert_eq!(parsed[1].name(), "My Style");
    assert!(!parsed[1].built_in());
    assert_eq!(parsed[1].built_in_data(), None);
}

#[test]
fn workbook_without_style_extensions_has_none() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Styles").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert!(workbook.style_extensions().is_empty());
}
