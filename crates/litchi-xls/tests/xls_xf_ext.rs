//! Round-trip tests for the BIFF8 XFExt record (formatting extensions).

use litchi_xls::writer::Writer;
use litchi_xls::{ExtProp, FullColorExt, FullColorType, Workbook, XfExt};
use std::io::Cursor;

fn extensions() -> Vec<XfExt> {
    let fill =
        ExtProp::FillForegroundColor(FullColorExt::try_new(FullColorType::Theme, -3, 4).unwrap());
    let text =
        ExtProp::TextColor(FullColorExt::try_new(FullColorType::Rgb, 0, 0x00FF_0000).unwrap());
    let indent = ExtProp::Indent(4);
    vec![
        XfExt::try_new(15, vec![fill, text]).unwrap(),
        XfExt::try_new(20, vec![indent]).unwrap(),
    ]
}

#[test]
fn xf_extensions_round_trip_through_writer_and_reader() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Ext").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_xf_extensions(extensions());
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();

    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    let formatting = workbook.formatting();
    assert_eq!(formatting.xf_extensions(), extensions().as_slice());
}

#[test]
fn xf_extension_index_is_validated_against_the_xf_table() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Ext").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    writer.set_xf_extensions(vec![XfExt::try_new(4000, Vec::new()).unwrap()]);
    let mut output = Cursor::new(Vec::new());
    assert!(writer.write_to(&mut output).is_err());
}

#[test]
fn workbook_without_xf_extensions_has_none() {
    let mut writer = Writer::new();
    let sheet = writer.add_worksheet("Ext").unwrap();
    writer.write_string(sheet, 0, 0, "content").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    let workbook = Workbook::new(Cursor::new(output.into_inner())).unwrap();
    assert!(workbook.formatting().xf_extensions().is_empty());
}
