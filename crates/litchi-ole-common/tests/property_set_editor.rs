#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    clippy::shadow_reuse,
    clippy::shadow_unrelated,
    clippy::cast_possible_truncation,
    reason = "integration tests use concise assertions and checked fixture-sized literals"
)]

use std::io::Cursor;
use std::path::PathBuf;

use litchi_cfb::{OleFile, OleFileLimits, OleWriter};
use litchi_ole_common::property_set::{
    CodePage, Edit, Guid, PropertySetReader, Standard, Value, Vector,
};

fn ole_with_streams(streams: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    for (name, data) in streams {
        writer.create_stream(&[*name], data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn ole_with_path(path: &[&str]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_stream(path, b"opaque").unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn ole_with_storage(path: &[&str]) -> Vec<u8> {
    let mut writer = OleWriter::new();
    writer.create_storage(path).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}
#[test]
fn generated_property_sets_round_trip_all_office_value_families() {
    let payload = b"untouched payload".to_vec();
    let original = ole_with_streams(&[("Payload", &payload)]);
    let mut editor = Edit::new(original, OleFileLimits::default()).unwrap();
    editor
        .update(Standard::SummaryInformation, |section| {
            section.set_page(CodePage::WINDOWS_1252);
            section.add(2, Value::Lpstr("Title".into()))?;
            section.add(3, Value::I4(-42))?;
            section.add(4, Value::UI8(u64::MAX - 1))?;
            section.add(5, Value::Bool(true))?;
            section.add(6, Value::Filetime(116_444_736_000_000_000))?;
            section.add(7, Value::Clsid(Guid::from_bytes([7; 16])))?;
            section.add(8, Value::Blob(vec![1, 2, 3]))?;
            section.add(
                9,
                Value::Clipboard {
                    format: 13,
                    data: vec![4, 5, 6],
                },
            )?;
            section.add(
                10,
                Value::Vector(
                    Vector::variant(vec![Value::I4(1), Value::Lpwstr("two".into())])
                        .expect("variant vector should validate"),
                ),
            )?;
            section.add(
                11,
                Value::Unknown {
                    variant_type: 0x7777,
                    data: vec![9, 8, 7, 6],
                },
            )?;
            Ok(())
        })
        .unwrap();
    editor
        .update(Standard::UserDefinedProperties, |section| {
            section.set_page(CodePage::Utf16Le);
            section.add_named(2, "担当者".into(), Value::Lpwstr("山田".into()))?;
            section.add_named(3, "Enabled".into(), Value::Bool(false))?;
            Ok(())
        })
        .unwrap();
    let bytes = editor.finish().unwrap();

    let mut ole = OleFile::open(Cursor::new(bytes.clone()), OleFileLimits::default()).unwrap();
    assert_eq!(ole.open_stream(&["Payload"]).unwrap(), payload);
    let summary = ole
        .property_set_stream(&["\u{0005}SummaryInformation"])
        .unwrap();
    assert_eq!(
        summary.sections[0].property(2),
        Some(&Value::Lpstr("Title".into()))
    );
    assert_eq!(
        summary.sections[0].property(10),
        Some(&Value::Vector(
            Vector::variant(vec![Value::I4(1), Value::Lpwstr("two".into())])
                .expect("variant vector should validate"),
        ))
    );
    assert_eq!(
        summary.sections[0].property(11),
        Some(&Value::Unknown {
            variant_type: 0x7777,
            data: vec![9, 8, 7, 6]
        })
    );
    let document = ole
        .property_set_stream(&["\u{0005}DocumentSummaryInformation"])
        .unwrap();
    let custom = document
        .sections
        .iter()
        .find(|section| section.property_name(2) == Some("担当者"))
        .unwrap();
    assert_eq!(
        custom.find_named("担当者").unwrap().1,
        &Value::Lpwstr("山田".into())
    );

    let mut editor = Edit::new(bytes, OleFileLimits::default()).unwrap();
    editor
        .update(Standard::SummaryInformation, |section| {
            section.update(2, Value::Lpstr("Changed".into()))?;
            Ok(())
        })
        .unwrap();
    let mut ole = OleFile::open(
        Cursor::new(editor.finish().unwrap()),
        OleFileLimits::default(),
    )
    .unwrap();
    let summary = ole
        .property_set_stream(&["\u{0005}SummaryInformation"])
        .unwrap();
    assert_eq!(
        summary.sections[0].property(11),
        Some(&Value::Unknown {
            variant_type: 0x7777,
            data: vec![9, 8, 7, 6]
        })
    );
}

#[test]
fn mutations_are_atomic_ordered_and_noops_are_byte_exact() {
    let original = ole_with_streams(&[("Payload", b"same")]);
    assert_eq!(
        Edit::new(original.clone(), OleFileLimits::default())
            .unwrap()
            .finish()
            .unwrap(),
        original
    );

    let mut editor = Edit::new(original.clone(), OleFileLimits::default()).unwrap();
    let error = editor.update(Standard::UserDefinedProperties, |section| {
        section.set_page(CodePage::Utf16Le);
        section.add_named(2, "Name".into(), Value::I4(1))?;
        Ok(section.add_named(3, "name".into(), Value::I4(2))?)
    });
    assert!(error.is_err());
    assert_eq!(editor.finish().unwrap(), original);

    let mut editor = Edit::new(original, OleFileLimits::default()).unwrap();
    editor
        .update(Standard::SummaryInformation, |section| {
            section.set_page(CodePage::WINDOWS_1252);
            section.add(4, Value::I4(4))?;
            section.add(2, Value::I4(2))?;
            section.reorder(&[1, 2, 4])?;
            Ok(())
        })
        .unwrap();
    let mut section = editor
        .property_set(Standard::SummaryInformation)
        .unwrap()
        .unwrap();
    assert_eq!(section.property_ids().collect::<Vec<_>>(), vec![1, 2, 4]);
    assert_eq!(section.remove_named("missing"), None);
}

#[test]
fn protected_containers_are_rejected_at_root_and_nested_case_insensitively() {
    for marker in [
        "_xmlsignatures",
        "_SIGNATURES",
        "dIGITALsIGNATURE",
        "\u{0005}dIGITALsIGNATURE",
        "\u{0006}DATAsPACES",
        "\u{0009}drmcontent",
        "\u{0009}drmvIEWERcONTENT",
        "ENCRYPTEDpACKAGE",
        "eNCRYPTIONiNFO",
    ] {
        assert!(
            Edit::new(ole_with_path(&[marker]), OleFileLimits::default()).is_err(),
            "{marker:?}"
        );
        assert!(
            Edit::new(
                ole_with_path(&["Nested", marker, "Payload"]),
                OleFileLimits::default()
            )
            .is_err(),
            "nested {marker:?}"
        );
    }
}

#[test]
fn drm_viewer_content_storages_are_rejected_at_root_and_nested_case_insensitively() {
    for path in [
        &["\u{0009}DRMViewerContent"][..],
        &["Nested", "\u{0009}drmvIEWERcONTENT"][..],
    ] {
        assert!(
            Edit::new(ole_with_storage(path), OleFileLimits::default()).is_err(),
            "{path:?}"
        );
    }
}

#[test]
fn poi_and_libreoffice_fixtures_preserve_exact_noop_bytes() {
    let root = PathBuf::from(env!("CARGO_MANIFEST_DIR")).join("../..");
    for relative in [
        "test-data/poi/test-data/hpsf/TestMickey.doc",
        "test-data/poi/test-data/hpsf/TestUnicode.xls",
        "test-data/ole/doc/documentProperties.doc",
    ] {
        let bytes = std::fs::read(root.join(relative)).unwrap();
        assert_eq!(
            Edit::new(bytes.clone(), OleFileLimits::default())
                .unwrap()
                .finish()
                .unwrap(),
            bytes
        );
    }
}
