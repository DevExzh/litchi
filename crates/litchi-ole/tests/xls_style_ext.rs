//! Round-trip tests for the BIFF8 StyleExt record (cell-style extensions).

use litchi_ole::xls::writer::XlsWriter;
use litchi_ole::xls::{XlsStyleCategory, XlsStyleExt, XlsWorkbook};
use litchi_ole::xls::{XlsXfProperties, XlsXfProperty};
use std::io::Cursor;

fn extensions() -> Vec<XlsStyleExt> {
    let mut heading = XlsStyleExt::try_new(
        true,
        XlsStyleCategory::TitleAndHeading,
        "Heading 1".to_string(),
        XlsXfProperties::try_new(vec![XlsXfProperty::FontItalic(true)]).unwrap(),
    )
    .unwrap();
    heading.set_hidden(true);
    let custom = XlsStyleExt::try_new(
        false,
        XlsStyleCategory::Custom,
        "My Style".to_string(),
        XlsXfProperties::default(),
    )
    .unwrap();
    vec![heading, custom]
}

#[test]
fn style_extensions_round_trip_through_writer_and_reader() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Styles").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_style_extensions(extensions());
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    let parsed = workbook.style_extensions();
    assert_eq!(parsed.len(), 2);
    assert_eq!(parsed[0].name(), "Heading 1");
    assert!(parsed[0].built_in());
    assert!(parsed[0].hidden());
    assert_eq!(parsed[0].category(), XlsStyleCategory::TitleAndHeading);
    assert_eq!(
        parsed[0].properties().properties(),
        &[XlsXfProperty::FontItalic(true)]
    );
    assert_eq!(parsed[1].name(), "My Style");
    assert!(!parsed[1].built_in());
    assert_eq!(parsed[1].built_in_data(), None);
}

#[test]
fn workbook_without_style_extensions_has_none() {
    let mut writer = XlsWriter::new();
    let sheet = writer.add_worksheet("Styles").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = XlsWorkbook::new(Cursor::new(output.into_inner())).unwrap();
    assert!(workbook.style_extensions().is_empty());
}
