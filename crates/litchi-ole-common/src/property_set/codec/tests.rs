//! Focused invariants for the layered Property Set codec.

use super::super::binding::Binding;
use super::super::model::{Guid, Section, Stream, Value};
use super::binary::parse_typed_property;
use super::{Edit, PropertySetReader};
use litchi_cfb::consts::{VT_BLOB, VT_I4};
use litchi_cfb::{OleError, OleFile, OleFileLimits, OleWriter};
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
fn outer_typed_property_padding_is_ignored_while_blob_data_is_preserved() {
    let bytes = [
        VT_BLOB as u8,
        0,
        0xfe,
        0xca,
        4,
        0,
        0,
        0,
        0xde,
        0xad,
        0xbe,
        0xef,
    ];
    assert_eq!(
        parse_typed_property(&bytes, 1252, 0).expect("nonzero outer padding is ignored"),
        Value::Blob(vec![0xde, 0xad, 0xbe, 0xef])
    );
}

#[test]
fn editor_preserves_nested_streams_with_matching_leaf_names() {
    let property_set = Stream::new(Section::new(super::super::model::SUMMARY_INFORMATION_FMTID))
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

    let mut editor =
        Edit::new(source.into_inner(), OleFileLimits::default()).expect("editor source");
    editor
        .update(Binding::SummaryInformation, |section| {
            Ok(section.add(2, Value::Lpstr("Title".into()))?)
        })
        .expect("summary edit");
    let mut ole = OleFile::open(
        Cursor::new(editor.finish().expect("edited CFB")),
        OleFileLimits::default(),
    )
    .expect("edited CFB opens");
    assert_eq!(
        ole.open_stream(&["Nested", "\u{0005}SummaryInformation"])
            .expect("nested stream survives"),
        b"nested payload"
    );
}

#[test]
fn editor_requires_and_enforces_the_callers_cfb_limits() {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["Workbook"], b"source")
        .expect("source stream");
    let mut source = Cursor::new(Vec::new());
    writer.write_to(&mut source).expect("CFB source");
    let bytes = source.into_inner();
    let limits = OleFileLimits::new(u64::try_from(bytes.len() - 1).expect("fixture length"))
        .expect("finite ingress limit");
    assert!(matches!(
        Edit::new(bytes, limits),
        Err(crate::Error::Cfb(OleError::LimitExceeded {
            resource: "input bytes",
            ..
        }))
    ));
}

#[test]
fn reader_opens_a_guid_derived_binding_without_a_format_specific_path() {
    let binding = Binding::custom(Guid::from_bytes([
        0x10, 0x32, 0x54, 0x76, 0x98, 0xBA, 0xDC, 0xFE, 0x01, 0x23, 0x45, 0x67, 0x89, 0xAB, 0xCD,
        0xEF,
    ]));
    let bytes = Stream::new(Section::new(binding.format_identifier()))
        .to_bytes()
        .expect("custom property set should serialize");
    let name = binding.name();

    let mut writer = OleWriter::new();
    writer
        .create_stream(&[name.as_str()], &bytes)
        .expect("standard binding stream should be writable");
    let mut cfb = Cursor::new(Vec::new());
    writer.write_to(&mut cfb).expect("CFB should serialize");

    let mut ole = OleFile::open(Cursor::new(cfb.into_inner()), OleFileLimits::default())
        .expect("CFB should open");
    let parsed = PropertySetReader::property_set(&mut ole, binding)
        .expect("binding reader should resolve the standard name");
    assert_eq!(
        parsed.sections[0].format_identifier,
        binding.format_identifier()
    );
}

#[test]
fn a_user_defined_edit_creates_a_typed_document_summary_section() {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Workbook"], b"host").expect("host");
    let mut source = Cursor::new(Vec::new());
    writer.write_to(&mut source).expect("CFB source");

    let mut editor =
        Edit::new(source.into_inner(), OleFileLimits::default()).expect("editor source");
    editor
        .update(Binding::UserDefinedProperties, |section| {
            Ok(section.add_named(2, "Department".into(), Value::Lpstr("Research".into()))?)
        })
        .expect("user-defined edit");
    let mut ole = OleFile::open(
        Cursor::new(editor.finish().expect("edited CFB")),
        OleFileLimits::default(),
    )
    .expect("edited CFB opens");
    let stream = PropertySetReader::property_set(&mut ole, Binding::DocumentSummaryInformation)
        .expect("document summary stream");
    assert_eq!(stream.sections.len(), 2);
    assert_eq!(
        stream.sections[0].page(),
        Some(super::super::model::CodePage::WINDOWS_1252)
    );
    // The standard section now satisfies the typed owner.
    super::super::document_summary::Snapshot::from_stream(&stream)
        .expect("typed document summary section");
    assert!(
        stream
            .section(super::super::model::USER_DEFINED_PROPERTIES_FMTID)
            .expect("user-defined section")
            .find_named("Department")
            .is_some()
    );
}
