#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

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
    if let Some(text) = value {
        data.extend_from_slice(&record(0, 1, RecordType::CString.as_u16(), &units(text)));
    }
    record(0x0f, 0, RecordType::ProgStringTag.as_u16(), &data)
}

fn binary_tag(name: &str, style_type: Option<RecordType>, payload: &[u8]) -> Vec<u8> {
    let name_data: Vec<u8> = name.encode_utf16().flat_map(u16::to_le_bytes).collect();
    let mut data = record(0, 0, RecordType::CString.as_u16(), &name_data);
    let blob_data = style_type.map_or_else(
        || payload.to_vec(),
        |kind| record(0, 0, kind.as_u16(), payload),
    );
    data.extend_from_slice(&record(
        0,
        0,
        RecordType::BinaryTagData.as_u16(),
        &blob_data,
    ));
    record(0x0f, 0, RecordType::ProgBinaryTag.as_u16(), &data)
}

fn complete_record(payload: &[u8]) -> (Vec<u8>, Record) {
    let bytes = record(0x0f, 7, RecordType::ProgTags.as_u16(), payload);
    let (parsed, consumed) = Record::parse_strict(&bytes, 0).unwrap();
    assert_eq!(consumed, bytes.len());
    (bytes, parsed)
}

#[test]
fn parses_all_defined_variants_and_round_trips_exactly() {
    let mut payload = string_tag("author", Some("Ada"));
    payload.extend_from_slice(&binary_tag(
        "___PPT9",
        Some(RecordType::StyleTextProp9Atom),
        &[0; 12],
    ));
    payload.extend_from_slice(&binary_tag(
        "___PPT10",
        Some(RecordType::StyleTextProp10Atom),
        &[0; 4],
    ));
    payload.extend_from_slice(&binary_tag(
        "___PPT11",
        Some(RecordType::StyleTextProp11Atom),
        &[0; 4],
    ));
    payload.extend_from_slice(&binary_tag("vendor", None, &[1, 2, 3, 4, 5]));
    let (bytes, record) = complete_record(&payload);

    let parsed =
        ShapeProgrammableTags::parse(&record, ShapeProgrammableTagLimits::default()).unwrap();

    assert_eq!(parsed.instance, 7);
    assert_eq!(parsed.tags.len(), 5);
    assert_eq!(parsed.powerpoint9().unwrap().runs.len(), 1);
    assert_eq!(parsed.powerpoint10().unwrap().runs.len(), 1);
    assert_eq!(parsed.powerpoint11().unwrap().runs.len(), 1);
    assert_eq!(
        parsed.tags.iter().find_map(|tag| match tag {
            ShapeProgrammableTag::String(string_tag_data) => string_tag_data.value.as_deref(),
            ShapeProgrammableTag::Binary(_) => None,
        }),
        Some("Ada")
    );
    assert_eq!(
        parsed
            .to_bytes(ShapeProgrammableTagLimits::default())
            .unwrap(),
        bytes
    );
}

#[test]
fn enforces_container_pair_and_version_ownership() {
    let limits = ShapeProgrammableTagLimits::default();
    let duplicate = [
        binary_tag("___PPT9", Some(RecordType::StyleTextProp9Atom), &[0; 12]),
        binary_tag("___PPT9", Some(RecordType::StyleTextProp9Atom), &[0; 12]),
    ]
    .concat();
    assert!(ShapeProgrammableTags::parse_payload(&duplicate, 0, limits).is_err());

    let wrong_style = binary_tag("___PPT10", Some(RecordType::StyleTextProp9Atom), &[0; 12]);
    assert!(ShapeProgrammableTags::parse_payload(&wrong_style, 0, limits).is_err());

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
    assert!(ShapeProgrammableTags::parse_payload(&missing_blob, 0, limits).is_err());

    let disallowed = record(0, 0, RecordType::CString.as_u16(), &[65, 0]);
    assert!(ShapeProgrammableTags::parse_payload(&disallowed, 0, limits).is_err());
}

#[test]
fn rejects_malformed_strings_headers_truncation_and_every_limit() {
    let defaults = ShapeProgrammableTagLimits::default();
    let valid = binary_tag("vendor", None, &[1, 2, 3, 4]);

    let mut truncated = valid.clone();
    truncated.pop();
    assert!(ShapeProgrammableTags::parse_payload(&truncated, 0, defaults).is_err());

    let invalid_utf16_cstring = record(0, 0, RecordType::CString.as_u16(), &[0x00, 0xd8]);
    let invalid_utf16 = record(
        0x0f,
        0,
        RecordType::ProgStringTag.as_u16(),
        &invalid_utf16_cstring,
    );
    assert!(ShapeProgrammableTags::parse_payload(&invalid_utf16, 0, defaults).is_err());

    let control_name = string_tag("bad\nname", None);
    assert!(ShapeProgrammableTags::parse_payload(&control_name, 0, defaults).is_err());

    let cases = [
        ShapeProgrammableTagLimits {
            max_container_bytes: valid.len() - 1,
            ..defaults
        },
        ShapeProgrammableTagLimits {
            max_tags: 0,
            ..defaults
        },
        ShapeProgrammableTagLimits {
            max_tag_bytes: 0,
            ..defaults
        },
        ShapeProgrammableTagLimits {
            max_string_code_units: 1,
            ..defaults
        },
        ShapeProgrammableTagLimits {
            max_unknown_binary_bytes: 3,
            ..defaults
        },
    ];
    for limits in cases {
        assert!(ShapeProgrammableTags::parse_payload(&valid, 0, limits).is_err());
    }

    let known = binary_tag("___PPT10", Some(RecordType::StyleTextProp10Atom), &[0; 4]);
    assert!(
        ShapeProgrammableTags::parse_payload(
            &known,
            0,
            ShapeProgrammableTagLimits {
                max_style_payload_bytes: 3,
                ..defaults
            },
        )
        .is_err()
    );
    assert!(
        ShapeProgrammableTags::parse_payload(
            &known,
            0,
            ShapeProgrammableTagLimits {
                max_style_runs: 0,
                ..defaults
            },
        )
        .is_err()
    );
}
