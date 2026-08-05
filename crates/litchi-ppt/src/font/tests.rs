use super::{PowerPointFontCollection, PowerPointFontCollections};
use crate::{PptRecord, PptRecordType};

fn record_bytes(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    let mut data = Vec::new();
    data.extend_from_slice(&((instance << 4) | version).to_le_bytes());
    data.extend_from_slice(&kind.to_le_bytes());
    data.extend_from_slice(&(payload.len() as u32).to_le_bytes());
    data.extend_from_slice(payload);
    data
}

fn collection(kind: PptRecordType, payload: Vec<u8>) -> PptRecord {
    PptRecord {
        record_type: kind,
        record_type_raw: kind.as_u16(),
        version: 0x0f,
        instance: 0,
        data_length: payload.len() as u32,
        data: payload,
        children: Vec::new(),
    }
}

fn prog_tags_record(version: u8, blob_payload: &[u8]) -> PptRecord {
    let tag_name: Vec<u8> = format!("___PPT{version}")
        .encode_utf16()
        .flat_map(u16::to_le_bytes)
        .collect();
    let name = record_bytes(0, 0, 4026, &tag_name);
    let blob = record_bytes(0, 0, 0x138b, blob_payload);
    let mut tag_payload = name;
    tag_payload.extend_from_slice(&blob);
    let tag = record_bytes(0x0f, 0, 0x138a, &tag_payload);
    PptRecord {
        record_type: PptRecordType::ProgTags,
        record_type_raw: 0x1388,
        version: 0x0f,
        instance: 0,
        data_length: tag.len() as u32,
        data: tag,
        children: Vec::new(),
    }
}

#[test]
fn parses_font_collections_and_embedded_facets() {
    let mut entity = vec![0u8; 68];
    for (index, unit) in "Noto Sans CJK"
        .encode_utf16()
        .chain(std::iter::once(0))
        .enumerate()
    {
        entity[index * 2..index * 2 + 2].copy_from_slice(&unit.to_le_bytes());
    }
    entity[64] = 0x80;
    entity[65] = 0x01;
    entity[66] = 0x0c;
    entity[67] = 0x22;
    let mut payload = record_bytes(0, 7, 4023, &entity);
    payload.extend_from_slice(&record_bytes(0, 0, 4024, b"plain-font"));
    payload.extend_from_slice(&record_bytes(0, 3, 4024, b"bold-italic-font"));

    let fonts =
        PowerPointFontCollection::parse(&collection(PptRecordType::FontCollection10, payload))
            .unwrap();

    assert!(fonts.international);
    let font = fonts.get(7).unwrap();
    assert_eq!(font.name, "Noto Sans CJK");
    assert_eq!(font.charset, 0x80);
    assert!(font.embedded_subset);
    assert!(font.truetype);
    assert!(font.no_substitution);
    assert_eq!(font.pitch_and_family, 0x22);
    assert_eq!(font.embedded_fonts.len(), 2);
    assert_eq!(font.embedded_fonts[1].style, 3);
}

#[test]
fn rejects_malformed_font_collections() {
    let mut unterminated = vec![b'A'; 68];
    unterminated[66] = 0;
    let data = record_bytes(0, 0, 4023, &unterminated);
    assert!(
        PowerPointFontCollection::parse(&collection(PptRecordType::FontCollection, data,)).is_err()
    );

    let embedded_first = record_bytes(0, 0, 4024, b"font");
    assert!(
        PowerPointFontCollection::parse(
            &collection(PptRecordType::FontCollection, embedded_first,)
        )
        .is_err()
    );
}

#[test]
fn resolves_base_and_international_font_collections() {
    let mut entity = vec![0u8; 68];
    entity[..4].copy_from_slice(&[b'A', 0, 0, 0]);
    entity[66] = 4;
    let base = collection(
        PptRecordType::FontCollection,
        record_bytes(0, 0, 4023, &entity),
    );
    let international_bytes = record_bytes(0, 9, 4023, &entity);
    let international = record_bytes(0x0f, 0, 2006, &international_bytes);
    let embedding_flags = record_bytes(0, 0, 0x32c8, &0xffff_ffffu32.to_le_bytes());
    let mut extension = international;
    extension.extend_from_slice(&embedding_flags);
    let root = PptRecord {
        record_type: PptRecordType::Document,
        record_type_raw: 1000,
        version: 0x0f,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children: vec![base, prog_tags_record(10, &extension)],
    };

    let fonts = PowerPointFontCollections::parse(&root).unwrap();
    assert_eq!(fonts.get_base(0).unwrap().name, "A");
    assert_eq!(fonts.get_international(9).unwrap().name, "A");
    let flags = fonts.embedding_flags.unwrap();
    assert_eq!(flags.raw, 0xffff_ffff);
    assert!(flags.subset);
    assert!(flags.subset_option_confirmed);
}

#[test]
fn rejects_malformed_or_duplicate_font_embedding_flags() {
    let malformed_header = record_bytes(1, 0, 0x32c8, &0u32.to_le_bytes());
    let root = PptRecord {
        record_type: PptRecordType::Document,
        record_type_raw: 1000,
        version: 0x0f,
        instance: 0,
        data_length: 0,
        data: Vec::new(),
        children: vec![prog_tags_record(10, &malformed_header)],
    };
    assert!(PowerPointFontCollections::parse(&root).is_err());

    let malformed_size = record_bytes(0, 0, 0x32c8, &[0, 0, 0]);
    let mut duplicate = record_bytes(0, 0, 0x32c8, &0u32.to_le_bytes());
    duplicate.extend_from_slice(&record_bytes(0, 0, 0x32c8, &1u32.to_le_bytes()));
    for payload in [&malformed_size[..], &duplicate[..]] {
        let root = PptRecord {
            record_type: PptRecordType::Document,
            record_type_raw: 1000,
            version: 0x0f,
            instance: 0,
            data_length: 0,
            data: Vec::new(),
            children: vec![prog_tags_record(10, payload)],
        };
        assert!(PowerPointFontCollections::parse(&root).is_err());
    }
}
