#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

//! Independent vectors for the embedded-font structures in MS-PPT 2.9.8-2.9.12
//! and 2.11.5.
//!
//! These fixtures are assembled from literal record identifiers and field
//! layouts. They do not use the production writer, system font discovery, or
//! a rendering engine.

#[cfg(feature = "encryption")]
use litchi_ppt::FontPackageOptions;
use litchi_ppt::{
    Font, FontCollection, FontEmbeddingFlags, FontScope, FontSnapshot, Record, RecordLimits,
    RecordType,
};
use std::io::Cursor;

const RT_FONT_COLLECTION: u16 = 0x07d5;
const RT_FONT_COLLECTION_10: u16 = 0x07d6;
const RT_FONT_ENTITY_ATOM: u16 = 0x0fb7;
const RT_FONT_EMBED_DATA_BLOB: u16 = 0x0fb8;

fn record(version: u16, instance: u16, record_type: u16, payload: &[u8]) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(8 + payload.len());
    bytes.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    bytes.extend_from_slice(&record_type.to_le_bytes());
    bytes.extend_from_slice(&u32::try_from(payload.len()).unwrap().to_le_bytes());
    bytes.extend_from_slice(payload);
    bytes
}

fn font_entity(instance: u16, name: &[u16]) -> Vec<u8> {
    assert!(name.len() <= 32);
    let mut payload = [0_u8; 68];
    for (index, unit) in name.iter().copied().enumerate() {
        payload[index * 2..index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
    }
    payload[64] = 0;
    payload[65] = 1;
    payload[66] = 4;
    payload[67] = 0x22;
    record(0, instance, RT_FONT_ENTITY_ATOM, &payload)
}

fn terminated_name(value: &str) -> Vec<u16> {
    let mut units = value.encode_utf16().collect::<Vec<_>>();
    units.push(0);
    units
}

fn collection(record_type: u16, payload: &[u8]) -> Vec<u8> {
    record(0x0f, 0, record_type, payload)
}

fn parsed_collection(bytes: &[u8]) -> FontCollection {
    let (record, consumed) = Record::parse(bytes, 0).expect("valid outer PPT record");
    assert_eq!(consumed, bytes.len());
    FontCollection::parse(&record).expect("valid MS-PPT font collection")
}

#[test]
fn base_collection_uses_ordinal_refs_and_accepts_all_four_facets() {
    let full_width_name = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdef";
    assert_eq!(full_width_name.encode_utf16().count(), 32);

    let eot = valid_eot();
    let mut payload = font_entity(7, &full_width_name.encode_utf16().collect::<Vec<_>>());
    for style in 0..=3 {
        payload.extend_from_slice(&record(0, style, RT_FONT_EMBED_DATA_BLOB, &eot));
    }
    // recInstance is deliberately duplicated. FontIndexRef addresses collection
    // order, not FontEntityAtom.recInstance.
    payload.extend_from_slice(&font_entity(7, &terminated_name("Second")));

    let fonts = parsed_collection(&collection(RT_FONT_COLLECTION, &payload));
    assert!(!fonts.international);
    assert_eq!(fonts.fonts.len(), 2);
    assert_eq!(fonts.fonts[0].name, full_width_name);
    assert_eq!(fonts.fonts[0].embedded_fonts.len(), 4);
    assert_eq!(
        fonts.fonts[0]
            .embedded_fonts
            .iter()
            .map(|facet| facet.style)
            .collect::<Vec<_>>(),
        [0, 1, 2, 3]
    );
    assert!(std::ptr::eq(
        fonts.get(0).unwrap(),
        &raw const fonts.fonts[0]
    ));
    assert!(std::ptr::eq(
        fonts.get(1).unwrap(),
        &raw const fonts.fonts[1]
    ));
    assert!(fonts.get(2).is_none());

    #[cfg(feature = "fonts")]
    for facet in &fonts.fonts[0].embedded_fonts {
        let view = litchi_fonts::embedding::powerpoint::View::parse(&facet.data)
            .expect("fixture facet is canonical uncompressed EOT 1.0");
        assert_eq!(view.family_name().decode().unwrap(), "Fixture");
    }
}

#[test]
fn international_collection_and_pp10_flags_keep_ignored_bits_inert() {
    let payload = font_entity(128, &terminated_name("International"));
    let fonts = parsed_collection(&collection(RT_FONT_COLLECTION_10, &payload));
    assert!(fonts.international);

    let raw = 0x8000_0003_u32;
    let flags = FontEmbeddingFlags::parse(&Record {
        record_type: RecordType::FontEmbedFlags10Atom,
        record_type_raw: 0x32c8,
        version: 0,
        instance: 0,
        data_length: 4,
        data: raw.to_le_bytes().to_vec(),
        children: Vec::new(),
    })
    .expect("undefined PP10 bits are ignored, not rejected");
    assert_eq!(flags.raw, raw);
    assert!(flags.subset);
    assert!(flags.subset_option_confirmed);
}

#[test]
fn malformed_font_names_and_entity_lengths_are_rejected() {
    let malformed_utf16 = font_entity(0, &[0xd800, 0]);
    let parsed = Record::parse(&collection(RT_FONT_COLLECTION, &malformed_utf16), 0)
        .unwrap()
        .0;
    assert!(FontCollection::parse(&parsed).is_err());

    let short_entity = record(0, 0, RT_FONT_ENTITY_ATOM, &[0; 67]);
    let record = Record::parse(&collection(RT_FONT_COLLECTION, &short_entity), 0)
        .unwrap()
        .0;
    assert!(FontCollection::parse(&record).is_err());
}

#[test]
fn record_limits_reject_oversize_font_records_before_materialization() {
    let bytes = collection(RT_FONT_COLLECTION, &[0; 128]);
    let limits = RecordLimits {
        max_input_bytes: bytes.len(),
        max_record_bytes: 64,
        max_record_payload_bytes: 56,
        max_copied_payload_bytes: 56,
        ..RecordLimits::default()
    };
    assert!(Record::parse_with_limits(&bytes, 0, limits).is_err());
}

#[cfg(feature = "fonts")]
#[test]
fn eot_parser_rejects_malformed_utf16_sizes_and_explicit_limit_overflow() {
    use litchi_fonts::embedding::powerpoint::{Limits, View};

    let eot = valid_eot();
    assert!(View::parse(&eot).is_ok());

    let mut wrong_size = eot.clone();
    wrong_size[0..4].copy_from_slice(&0_u32.to_le_bytes());
    assert!(View::parse(&wrong_size).is_err());

    let mut malformed_name = eot.clone();
    // The first name begins after the 82-byte fixed header and its u16 size.
    malformed_name[84..86].copy_from_slice(&0xd800_u16.to_le_bytes());
    assert!(View::parse(&malformed_name).is_err());

    let limits = Limits {
        max_input_bytes: eot.len() - 1,
        ..Limits::default()
    };
    assert!(View::parse_with(&eot, limits).is_err());
}

#[test]
fn signed_source_allows_exact_no_op_but_refuses_changed_font_commit() {
    let source = with_stream(
        unsigned_ppt(),
        &["_xmlsignatures", "origin.sigs"],
        b"fixture",
    );
    let snapshot =
        FontSnapshot::from_bytes(source.clone()).expect("signed PPT remains inspectable");

    let commit = snapshot.edit().unwrap().commit().expect("exact no-op");
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source);

    let mut transaction = snapshot.edit().unwrap();
    transaction
        .append_font(FontScope::Base, Font::new("Changed"))
        .unwrap();
    assert!(transaction.commit().is_err());
}

#[test]
fn signed_inconsistent_save_with_fonts_flag_is_preserved_on_exact_no_op() {
    let inconsistent = with_inconsistent_save_with_fonts(unsigned_ppt());
    let source = with_stream(inconsistent, &["_xmlsignatures", "origin.sigs"], b"fixture");
    let snapshot =
        FontSnapshot::from_bytes(source.clone()).expect("signed PPT remains inspectable");
    assert!(snapshot.fonts().base.as_ref().is_some_and(|collection| {
        collection
            .fonts
            .iter()
            .all(|font| font.embedded_fonts.is_empty())
    }));

    let commit = snapshot.edit().unwrap().commit().expect("exact no-op");
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source);
}

#[test]
#[cfg(feature = "encryption")]
fn encrypted_source_allows_exact_no_op_but_refuses_changed_font_commit() {
    let password = "font fixture password";
    let source = encrypted_ppt(password);
    let snapshot = FontSnapshot::from_bytes_with_options(
        source.clone(),
        FontPackageOptions {
            password: Some(password),
            ..FontPackageOptions::default()
        },
    )
    .expect("password-authorized font inspection");

    let commit = snapshot.edit().unwrap().commit().expect("exact no-op");
    assert!(commit.patch().is_empty());
    assert_eq!(commit.snapshot().bytes(), source);

    let mut transaction = snapshot.edit().unwrap();
    transaction
        .append_font(FontScope::Base, Font::new("Changed"))
        .unwrap();
    assert!(transaction.commit().is_err());
}

#[test]
fn empty_root_and_nested_protection_storages_refuse_changed_publication() {
    for path in [
        &["_SIGNATURES"][..],
        &["Nested", "\u{6}DataSpaces"][..],
        &["Nested", "_XmlSignatures"][..],
        &["\u{9}DRMViewerContent"][..],
        &["Nested", "\u{9}drmvIEWERcONTENT"][..],
    ] {
        let source = with_storage(unsigned_ppt(), path);
        let snapshot = FontSnapshot::from_bytes(source.clone()).expect("protected PPT inspection");
        let no_op = snapshot
            .edit()
            .unwrap()
            .commit()
            .expect("exact protected no-op");
        assert_eq!(no_op.snapshot().bytes(), source);

        let mut changed = snapshot.edit().unwrap();
        changed
            .append_font(FontScope::Base, Font::new("Changed"))
            .unwrap();
        assert!(changed.commit().is_err());
    }
}

fn unsigned_ppt() -> Vec<u8> {
    let mut writer = litchi_ppt::writer::Writer::new();
    writer.add_slide().unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[cfg(feature = "encryption")]
fn encrypted_ppt(password: &str) -> Vec<u8> {
    use litchi_ppt::writer::EncryptionProfile;

    let mut writer = litchi_ppt::writer::Writer::new();
    writer.add_slide().unwrap();
    writer
        .set_password(password, EncryptionProfile::CryptoApiRc4 { key_bits: 128 })
        .unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn with_stream(source: Vec<u8>, path: &[&str], payload: &[u8]) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleWriter};

    let mut source_file = OleFile::open(Cursor::new(source)).unwrap();
    let streams = source_file
        .list_streams()
        .into_iter()
        .map(|stream_path| {
            let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = source_file.open_stream(&refs).unwrap();
            (stream_path, data)
        })
        .collect::<Vec<_>>();
    let mut writer = OleWriter::new();
    for (stream_path, data) in streams {
        let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer.create_stream(path, payload).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn with_storage(source: Vec<u8>, path: &[&str]) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleWriter};

    let mut source_file = OleFile::open(Cursor::new(source)).unwrap();
    let streams = source_file
        .list_streams()
        .into_iter()
        .map(|stream_path| {
            let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
            let data = source_file.open_stream(&refs).unwrap();
            (stream_path, data)
        })
        .collect::<Vec<_>>();
    let mut writer = OleWriter::new();
    for (stream_path, data) in streams {
        let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
        writer.create_stream(&refs, &data).unwrap();
    }
    writer.create_storage(path).unwrap();
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn with_inconsistent_save_with_fonts(source: Vec<u8>) -> Vec<u8> {
    use litchi_cfb::{OleFile, OleWriter};

    let snapshot = FontSnapshot::from_bytes(source.clone()).unwrap();
    let owner = snapshot.document_bytes();
    let mut offset = 8usize;
    let flag_offset = loop {
        let header = owner
            .get(offset..offset + 8)
            .expect("direct document child");
        let kind = u16::from_le_bytes([header[2], header[3]]);
        let length = u32::from_le_bytes(header[4..8].try_into().unwrap()) as usize;
        if kind == RecordType::DocumentAtom.as_u16() {
            assert_eq!(length, 40);
            break offset + 8 + 36;
        }
        offset += 8 + length;
    };

    let mut source_file = OleFile::open(Cursor::new(source)).unwrap();
    let streams = source_file
        .list_streams()
        .into_iter()
        .map(|stream_path| {
            let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
            let mut data = source_file.open_stream(&refs).unwrap();
            if stream_path
                .last()
                .is_some_and(|name| name == "PowerPoint Document")
            {
                let start = data
                    .windows(owner.len())
                    .rposition(|candidate| candidate == owner)
                    .expect("live document owner in stream");
                data[start + flag_offset] = 1;
            }
            (stream_path, data)
        })
        .collect::<Vec<_>>();
    let mut writer = OleWriter::new();
    for (stream_path, data) in streams {
        let refs = stream_path.iter().map(String::as_str).collect::<Vec<_>>();
        writer.create_stream(&refs, &data).unwrap();
    }
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

fn valid_eot() -> Vec<u8> {
    let font = minimal_sfnt();
    let names = ["Fixture", "Regular", "Version 1", "Fixture Regular"];
    let names_bytes = names
        .iter()
        .enumerate()
        .map(|(index, name)| 2 + name.encode_utf16().count() * 2 + usize::from(index != 3))
        .sum::<usize>()
        + 3;
    let size = 82 + names_bytes + font.len();
    let mut eot = vec![0_u8; 82];
    set_le_u32(&mut eot, 0, u32::try_from(size).unwrap());
    set_le_u32(&mut eot, 4, u32::try_from(font.len()).unwrap());
    set_le_u32(&mut eot, 8, 0x0001_0000);
    eot[16..26].copy_from_slice(&[2, 11, 6, 4, 2, 2, 2, 2, 2, 4]);
    set_le_u32(&mut eot, 28, 400);
    set_le_u16(&mut eot, 32, 0);
    set_le_u16(&mut eot, 34, 0x504c);
    for (index, name) in names.iter().enumerate() {
        let units = name.encode_utf16().collect::<Vec<_>>();
        eot.extend_from_slice(&u16::try_from(units.len() * 2).unwrap().to_le_bytes());
        for unit in units {
            eot.extend_from_slice(&unit.to_le_bytes());
        }
        if index != 3 {
            eot.extend_from_slice(&0_u16.to_le_bytes());
        }
    }
    eot.extend_from_slice(&font);
    assert_eq!(eot.len(), size);
    eot
}

fn minimal_sfnt() -> Vec<u8> {
    const TABLE_OFFSET: usize = 28;
    const OS2_LEN: usize = 96;
    let mut font = vec![0_u8; TABLE_OFFSET + OS2_LEN];
    set_be_u32(&mut font, 0, 0x0001_0000);
    set_be_u16(&mut font, 4, 1);
    font[12..16].copy_from_slice(b"OS/2");
    set_be_u32(&mut font, 20, u32::try_from(TABLE_OFFSET).unwrap());
    set_be_u32(&mut font, 24, u32::try_from(OS2_LEN).unwrap());
    set_be_u16(&mut font, TABLE_OFFSET, 2);
    set_be_u16(&mut font, TABLE_OFFSET + 4, 400);
    set_be_u16(&mut font, TABLE_OFFSET + 6, 5);
    set_be_u16(&mut font, TABLE_OFFSET + 8, 0);
    font[TABLE_OFFSET + 32..TABLE_OFFSET + 42].copy_from_slice(&[2, 11, 6, 4, 2, 2, 2, 2, 2, 4]);
    font
}

fn set_le_u16(bytes: &mut [u8], offset: usize, value: u16) {
    bytes[offset..offset + 2].copy_from_slice(&value.to_le_bytes());
}

fn set_le_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_le_bytes());
}

fn set_be_u16(bytes: &mut [u8], offset: usize, value: u16) {
    bytes[offset..offset + 2].copy_from_slice(&value.to_be_bytes());
}

fn set_be_u32(bytes: &mut [u8], offset: usize, value: u32) {
    bytes[offset..offset + 4].copy_from_slice(&value.to_be_bytes());
}
