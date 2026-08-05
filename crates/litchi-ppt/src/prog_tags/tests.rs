use super::codec::encode_record;
use super::*;
use crate::consts::RecordType;
use crate::records::Record;

fn record(version: u16, instance: u16, kind: u16, payload: &[u8]) -> Vec<u8> {
    encode_record(version, instance, kind, payload).unwrap()
}

fn string_tag(name: &str, value: Option<&str>) -> Vec<u8> {
    let units = |text: &str| {
        text.encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>()
    };
    let mut data = record(0, 0, RecordType::CString.as_u16(), &units(name));
    if let Some(value) = value {
        data.extend_from_slice(&record(0, 1, RecordType::CString.as_u16(), &units(value)));
    }
    record(0x0f, 0, RecordType::ProgStringTag.as_u16(), &data)
}

fn binary_tag(name: &str, payload: &[u8]) -> Vec<u8> {
    let name_data: Vec<u8> = name.encode_utf16().flat_map(u16::to_le_bytes).collect();
    let mut data = record(0, 0, RecordType::CString.as_u16(), &name_data);
    data.extend_from_slice(&record(0, 0, RecordType::BinaryTagData.as_u16(), payload));
    record(0x0f, 0, RecordType::ProgBinaryTag.as_u16(), &data)
}

fn complete_record(payload: &[u8]) -> (Vec<u8>, Record) {
    let bytes = record(0x0f, 3, RecordType::ProgTags.as_u16(), payload);
    let (parsed, consumed) = Record::parse_strict(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    (bytes, parsed)
}

fn versioned_payload() -> Vec<u8> {
    // Two arbitrary atom records forming a valid strict record sequence.
    let mut payload = record(0, 0, RecordType::TextHeaderAtom.as_u16(), &[0; 4]);
    payload.extend_from_slice(&record(
        0,
        0,
        RecordType::StyleTextProp9Atom.as_u16(),
        &[0; 12],
    ));
    payload
}

#[test]
fn parses_document_scope_variants_and_round_trips_exactly() {
    let limits = ProgTagLimits::default();
    let mut payload = string_tag("author", Some("Ada"));
    payload.extend_from_slice(&binary_tag("___PPT9", &versioned_payload()));
    payload.extend_from_slice(&binary_tag("___PPT10", &versioned_payload()));
    payload.extend_from_slice(&binary_tag("___PPT11", &versioned_payload()));
    payload.extend_from_slice(&binary_tag("___PPT12", &versioned_payload()));
    payload.extend_from_slice(&binary_tag("vendor", &[1, 2, 3]));
    let (bytes, record) = complete_record(&payload);

    let parsed = ProgTags::parse(&record, ProgTagScope::Document, limits).unwrap();

    assert_eq!(parsed.scope, ProgTagScope::Document);
    assert_eq!(parsed.instance, 3);
    assert_eq!(parsed.tags.len(), 6);
    for version in [
        ProgBinaryTagVersion::PowerPoint9,
        ProgBinaryTagVersion::PowerPoint10,
        ProgBinaryTagVersion::PowerPoint11,
        ProgBinaryTagVersion::PowerPoint12,
    ] {
        let tag = parsed.binary_tag(version).unwrap();
        assert_eq!(tag.records().unwrap().len(), 2);
    }
    let unknown = parsed.binary_tag(ProgBinaryTagVersion::Unknown).unwrap();
    assert_eq!(unknown.name, "vendor");
    assert_eq!(unknown.payload, [1, 2, 3]);
    assert!(unknown.records().is_err());
    assert_eq!(
        parsed.tags.iter().find_map(|tag| match tag {
            ProgTag::String(tag) => tag.value.as_deref(),
            _ => None,
        }),
        Some("Ada")
    );
    assert_eq!(parsed.to_bytes(limits).unwrap(), bytes);
}

#[test]
fn slide_scope_treats_ppt11_as_unknown() {
    let limits = ProgTagLimits::default();
    let mut payload = binary_tag("___PPT9", &versioned_payload());
    payload.extend_from_slice(&binary_tag("___PPT11", &versioned_payload()));
    let (bytes, record) = complete_record(&payload);

    let parsed = ProgTags::parse(&record, ProgTagScope::Slide, limits).unwrap();

    assert_eq!(
        parsed
            .binary_tag(ProgBinaryTagVersion::PowerPoint9)
            .unwrap()
            .name,
        "___PPT9"
    );
    assert!(
        parsed
            .binary_tag(ProgBinaryTagVersion::PowerPoint11)
            .is_none()
    );
    let unknown = parsed.binary_tag(ProgBinaryTagVersion::Unknown).unwrap();
    assert_eq!(unknown.name, "___PPT11");
    assert_eq!(parsed.to_bytes(limits).unwrap(), bytes);
}

#[test]
fn enforces_duplicate_pair_and_versioned_payload_rules() {
    let limits = ProgTagLimits::default();
    let duplicate = [
        binary_tag("___PPT9", &versioned_payload()),
        binary_tag("___PPT9", &versioned_payload()),
    ]
    .concat();
    assert!(ProgTags::parse_payload(&duplicate, 0, ProgTagScope::Document, limits).is_err());

    // A versioned tag whose blob is not a strict record sequence is invalid.
    let invalid_blob = binary_tag("___PPT10", &[1, 2, 3]);
    assert!(ProgTags::parse_payload(&invalid_blob, 0, ProgTagScope::Document, limits).is_err());

    let mut missing_blob = record(
        0,
        0,
        RecordType::CString.as_u16(),
        &"vendor"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect::<Vec<_>>(),
    );
    missing_blob = record(0x0f, 0, RecordType::ProgBinaryTag.as_u16(), &missing_blob);
    assert!(ProgTags::parse_payload(&missing_blob, 0, ProgTagScope::Document, limits).is_err());

    let disallowed = record(0, 0, RecordType::CString.as_u16(), &[65, 0]);
    assert!(ProgTags::parse_payload(&disallowed, 0, ProgTagScope::Document, limits).is_err());
}

#[test]
fn rejects_malformed_strings_headers_truncation_and_every_limit() {
    let defaults = ProgTagLimits::default();
    let valid = binary_tag("vendor", &[1, 2, 3, 4]);

    let mut truncated = valid.clone();
    truncated.pop();
    assert!(ProgTags::parse_payload(&truncated, 0, ProgTagScope::Document, defaults).is_err());

    let invalid_utf16 = record(0, 0, RecordType::CString.as_u16(), &[0x00, 0xd8]);
    let invalid_utf16 = record(0x0f, 0, RecordType::ProgStringTag.as_u16(), &invalid_utf16);
    assert!(ProgTags::parse_payload(&invalid_utf16, 0, ProgTagScope::Document, defaults).is_err());

    let control_name = string_tag("bad\nname", None);
    assert!(ProgTags::parse_payload(&control_name, 0, ProgTagScope::Document, defaults).is_err());

    let cases = [
        ProgTagLimits {
            max_container_bytes: valid.len() - 1,
            ..defaults
        },
        ProgTagLimits {
            max_tags: 0,
            ..defaults
        },
        ProgTagLimits {
            max_tag_bytes: 0,
            ..defaults
        },
        ProgTagLimits {
            max_string_code_units: 1,
            ..defaults
        },
        ProgTagLimits {
            max_binary_payload_bytes: 3,
            ..defaults
        },
    ];
    for limits in cases {
        assert!(ProgTags::parse_payload(&valid, 0, ProgTagScope::Document, limits).is_err());
    }

    let known = binary_tag("___PPT12", &versioned_payload());
    assert!(
        ProgTags::parse_payload(
            &known,
            0,
            ProgTagScope::Document,
            ProgTagLimits {
                max_binary_records: 1,
                ..defaults
            }
        )
        .is_err()
    );
}

fn parsed_container(version: u16, kind: RecordType, payload: &[u8]) -> Record {
    let bytes = record(version, 0, kind.as_u16(), payload);
    let (parsed, consumed) = Record::parse_strict(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    parsed
}

#[test]
fn parse_document_locates_prog_tags_inside_doc_info_list() {
    let limits = ProgTagLimits::default();
    let tags_payload = string_tag("author", None);
    let prog_tags = record(0x0f, 0, RecordType::ProgTags.as_u16(), &tags_payload);
    let doc_info_list = record(0x0f, 0, RecordType::DocInfoList.as_u16(), &prog_tags);
    let document = parsed_container(0x0f, RecordType::Document, &doc_info_list);

    let parsed = ProgTags::parse_document(&document, limits)
        .unwrap()
        .unwrap();
    assert_eq!(parsed.scope, ProgTagScope::Document);
    assert_eq!(parsed.tags.len(), 1);

    // A document without a DocInfoListContainer has no tags.
    let bare = parsed_container(0x0f, RecordType::Document, &[]);
    assert!(ProgTags::parse_document(&bare, limits).unwrap().is_none());

    // A DocInfoListContainer without ProgTags has no tags.
    let empty_list = record(0x0f, 0, RecordType::DocInfoList.as_u16(), &[]);
    let no_tags = parsed_container(0x0f, RecordType::Document, &empty_list);
    assert!(
        ProgTags::parse_document(&no_tags, limits)
            .unwrap()
            .is_none()
    );

    // Duplicate DocInfoListContainer or ProgTags children are rejected.
    let duplicate_list = parsed_container(
        0x0f,
        RecordType::Document,
        &[doc_info_list.clone(), doc_info_list.clone()].concat(),
    );
    assert!(ProgTags::parse_document(&duplicate_list, limits).is_err());
    let duplicate_tags = record(
        0x0f,
        0,
        RecordType::DocInfoList.as_u16(),
        &[prog_tags.clone(), prog_tags].concat(),
    );
    let duplicate_tags = parsed_container(0x0f, RecordType::Document, &duplicate_tags);
    assert!(ProgTags::parse_document(&duplicate_tags, limits).is_err());

    // A non-Document record cannot provide document tags.
    let slide = parsed_container(0x0f, RecordType::Slide, &[]);
    assert!(ProgTags::parse_document(&slide, limits).is_err());
}

#[test]
fn parse_slide_locates_direct_prog_tags_child() {
    let limits = ProgTagLimits::default();
    let prog_tags = record(
        0x0f,
        0,
        RecordType::ProgTags.as_u16(),
        &binary_tag("___PPT9", &versioned_payload()),
    );
    let slide = parsed_container(0x0f, RecordType::Slide, &prog_tags);

    let parsed = ProgTags::parse_slide(&slide, limits).unwrap().unwrap();
    assert_eq!(parsed.scope, ProgTagScope::Slide);
    assert!(
        parsed
            .binary_tag(ProgBinaryTagVersion::PowerPoint9)
            .is_some()
    );

    let bare = parsed_container(0x0f, RecordType::Slide, &[]);
    assert!(ProgTags::parse_slide(&bare, limits).unwrap().is_none());

    let duplicate = parsed_container(
        0x0f,
        RecordType::Slide,
        &[prog_tags.clone(), prog_tags].concat(),
    );
    assert!(ProgTags::parse_slide(&duplicate, limits).is_err());
}
