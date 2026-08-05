//! Focused invariants for the layered Property Set codec.

use super::super::model::{Standard, Value};
use super::Editor;
use super::binary::parse_typed_property;
use litchi_cfb::consts::VT_I4;
use litchi_cfb::{OleFile, OleWriter};
use std::io::Cursor;

#[test]
fn typed_property_codec_preserves_scalar_values() {
    let bytes = [VT_I4 as u8, 0, 0, 0, 42, 0, 0, 0];
    assert_eq!(
        parse_typed_property(&bytes, 1252, 0).expect("valid typed property"),
        Value::I4(42)
    );
}

#[test]
fn editor_preserves_nested_streams_with_matching_leaf_names() {
    let property_set = super::super::model::Stream::new(super::super::model::Section::new(
        super::super::model::SUMMARY_INFORMATION_FMTID,
    ))
    .to_bytes()
    .expect("empty property set is serializable");

    let mut writer = OleWriter::new();
    writer
        .create_stream(&["\u{0005}SummaryInformation"], &property_set)
        .expect("root property set stream");
    writer
        .create_stream(&["Nested", "\u{0005}SummaryInformation"], b"nested payload")
        .expect("nested stream");
    let mut source = Cursor::new(Vec::new());
    writer.write_to(&mut source).expect("CFB source");

    let mut editor = Editor::new(source.into_inner()).expect("editor source");
    editor
        .update(Standard::SummaryInformation, |section| {
            section.add(2, Value::Lpstr("Title".into()))
        })
        .expect("summary edit");
    let mut ole =
        OleFile::open(Cursor::new(editor.finish().expect("edited CFB"))).expect("edited CFB opens");
    assert_eq!(
        ole.open_stream(&["Nested", "\u{0005}SummaryInformation"])
            .expect("nested stream survives"),
        b"nested payload"
    );
}
