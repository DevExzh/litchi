use std::io::Cursor;
use std::path::PathBuf;

use litchi_ole::{
    OleFile, OlePropertySetEditor, OleWriter, PropertySetGuid, PropertyValue,
    StandardPropertySet,
};

fn ole_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (name, data) in streams { writer.create_stream(&[*name], data).unwrap(); }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn generated_property_sets_round_trip_all_office_value_families() {
    let payload = b"untouched payload".to_vec();
    let original = ole_with_streams(&[("Payload", &payload)]);
    let mut editor = OlePropertySetEditor::new(original).unwrap();
    editor.update(StandardPropertySet::SummaryInformation, |section| {
        section.set_codepage(1252)?;
        section.add(2, PropertyValue::Lpstr("Title".into()))?;
        section.add(3, PropertyValue::I4(-42))?;
        section.add(4, PropertyValue::UI8(u64::MAX - 1))?;
        section.add(5, PropertyValue::Bool(true))?;
        section.add(6, PropertyValue::Filetime(116_444_736_000_000_000))?;
        section.add(7, PropertyValue::Clsid(PropertySetGuid::from_bytes([7; 16])))?;
        section.add(8, PropertyValue::Blob(vec![1, 2, 3]))?;
        section.add(9, PropertyValue::Clipboard { format: 13, data: vec![4, 5, 6] })?;
        section.add(10, PropertyValue::Vector(vec![PropertyValue::I4(1), PropertyValue::Lpwstr("two".into())]))?;
        section.add(11, PropertyValue::Unknown { variant_type: 0x7777, data: vec![9, 8, 7, 6] })?;
        Ok(())
    }).unwrap();
    editor.update(StandardPropertySet::UserDefinedProperties, |section| {
        section.set_codepage(1200)?;
        section.add_named(2, "担当者".into(), PropertyValue::Lpwstr("山田".into()))?;
        section.add_named(3, "Enabled".into(), PropertyValue::Bool(false))?;
        Ok(())
    }).unwrap();
    let bytes = editor.finish().unwrap();

    let mut ole = OleFile::open(Cursor::new(bytes.clone())).unwrap();
    assert_eq!(ole.open_stream(&["Payload"]).unwrap(), payload);
    let summary = ole.property_set_stream(&["\u{0005}SummaryInformation"]).unwrap();
    assert_eq!(summary.sections[0].property(2), Some(&PropertyValue::Lpstr("Title".into())));
    assert_eq!(summary.sections[0].property(10), Some(&PropertyValue::Vector(vec![PropertyValue::I4(1), PropertyValue::Lpwstr("two".into())])));
    assert_eq!(summary.sections[0].property(11), Some(&PropertyValue::Unknown { variant_type: 0x7777, data: vec![9, 8, 7, 6] }));
    let document = ole.property_set_stream(&["\u{0005}DocumentSummaryInformation"]).unwrap();
    let custom = document.sections.iter().find(|section| section.property_name(2) == Some("担当者")).unwrap();
    assert_eq!(custom.find_named("担当者").unwrap().1, &PropertyValue::Lpwstr("山田".into()));

    let mut editor = OlePropertySetEditor::new(bytes).unwrap();
    editor.update(StandardPropertySet::SummaryInformation, |section| { section.update(2, PropertyValue::Lpstr("Changed".into()))?; Ok(()) }).unwrap();
    let mut ole = OleFile::open(Cursor::new(editor.finish().unwrap())).unwrap();
    let summary = ole.property_set_stream(&["\u{0005}SummaryInformation"]).unwrap();
    assert_eq!(summary.sections[0].property(11), Some(&PropertyValue::Unknown { variant_type: 0x7777, data: vec![9, 8, 7, 6] }));
}

#[test]
fn mutations_are_atomic_ordered_and_noops_are_byte_exact() {
    let original = ole_with_streams(&[("Payload", b"same")]);
    assert_eq!(OlePropertySetEditor::new(original.clone()).unwrap().finish().unwrap(), original);

    let mut editor = OlePropertySetEditor::new(original.clone()).unwrap();
    let error = editor.update(StandardPropertySet::UserDefinedProperties, |section| {
        section.set_codepage(1200)?;
        section.add_named(2, "Name".into(), PropertyValue::I4(1))?;
        section.add_named(3, "name".into(), PropertyValue::I4(2))
    });
    assert!(error.is_err());
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = OlePropertySetEditor::new(original).unwrap();
    editor.update(StandardPropertySet::SummaryInformation, |section| {
        section.set_codepage(1252)?; section.add(4, PropertyValue::I4(4))?; section.add(2, PropertyValue::I4(2))?;
        section.reorder(&[1, 2, 4])?; Ok(())
    }).unwrap();
    let mut section = editor.property_set(StandardPropertySet::SummaryInformation).unwrap().unwrap();
    assert_eq!(section.property_ids().collect::<Vec<_>>(), vec![1, 2, 4]);
    assert_eq!(section.remove_named("missing"), None);
}

#[test]
fn signed_and_encrypted_containers_are_rejected_without_execution() {
    for name in ["EncryptionInfo", "\u{0005}DigitalSignature"] {
        assert!(OlePropertySetEditor::new(ole_with_streams(&[(name, b"opaque")])).is_err());
    }
}

#[test]
fn poi_and_libreoffice_fixtures_preserve_exact_noop_bytes() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    for relative in [
        "3rdparty/poi/test-data/hpsf/TestMickey.doc",
        "3rdparty/poi/test-data/hpsf/TestUnicode.xls",
        "test-data/ole/doc/documentProperties.doc",
    ] {
        let bytes = std::fs::read(root.join(relative)).unwrap();
        assert_eq!(OlePropertySetEditor::new(bytes.clone()).unwrap().finish().unwrap(), bytes);
    }
}

#[test]
fn doc_xls_and_ppt_facades_expose_standard_property_sets() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    let mut doc = litchi_ole::doc::Package::open(root.join("test-data/ole/doc/documentProperties.doc")).unwrap();
    assert!(doc.summary_information().unwrap().is_some());
    let mut xls = litchi_ole::xls::XlsWorkbook::new(std::fs::File::open(root.join("test-data/ole/xls/Simple.xls")).unwrap()).unwrap();
    let _ = xls.summary_information().unwrap();
    let mut ppt = litchi_ole::ppt::Package::open(root.join("3rdparty/poi/test-data/slideshow/text-margins.ppt")).unwrap();
    let _ = ppt.document_summary_information().unwrap();
}
