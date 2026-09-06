#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "The fixtures deliberately use small, checked wire values."
)]

use litchi_iwa_core::{
    Archive, ArchiveLimits, ArchiveObject, Error, FieldInfo, FieldPath, FieldType, LimitKind,
    RawMessage,
};

const ROOT_UNKNOWN: &[u8] = &[
    0xa0, 0x06, 0x81, 0x00, // unknown varint with a non-canonical value
    0xa9, 0x06, 0x91, 0x92, 0x93, 0x94, 0x95, 0x96, 0x97, 0x98, // unknown fixed64
    0xb2, 0x06, 0x05, b'R', b'O', b'O', b'T', b'?', // unknown bytes
];

const MESSAGE_ZERO_UNKNOWN: &[u8] = &[
    0xa0, 0x06, 0x83, 0x00, // unknown varint with a non-canonical value
    0xb2, 0x06, 0x05, b'M', b'Z', b'E', b'R', b'O', // unknown bytes
];

const MESSAGE_ONE_UNKNOWN: &[u8] = &[
    0xa0, 0x06, 0x85, 0x00, // unknown varint with a non-canonical value
    0xb2, 0x06, 0x05, b'M', b'O', b'N', b'E', b'!', // unknown bytes
];

#[test]
fn clone_remap_scratch_obeys_header_memory_limit_for_a_small_source() -> Result<(), Error> {
    let source = ArchiveObject::new(
        1,
        vec![RawMessage {
            type_: 7,
            data: vec![1],
        }],
    )?;
    let before = archive_object_bytes(&source)?;
    let limits = ArchiveLimits::default().with_header_memory_bytes(4096)?;
    source.validate_with_limits(limits)?;
    let remap = (100_u64..2100)
        .map(|identifier| (identifier, identifier + 10_000))
        .collect::<Vec<_>>();
    let error = source
        .clone_with_identity_remap_with_limits(2, &remap, &source.messages, limits)
        .expect_err("the remap scratch exceeds the budget although the source fits");
    assert!(matches!(
        error,
        Error::Limit {
            kind: LimitKind::HeaderMemoryBytes,
            ..
        }
    ));
    assert_eq!(archive_object_bytes(&source)?, before);
    Ok(())
}

#[test]
fn clone_remaps_aggregate_and_field_objects_preserves_data_and_raw_header_spans()
-> Result<(), Error> {
    let (source, source_bytes) = source_object_with_adversarial_header()?;
    let replacement_messages = vec![
        RawMessage {
            type_: 101,
            data: vec![0xa5; 130],
        },
        RawMessage {
            type_: 202,
            data: vec![0xb6; 3],
        },
    ];
    let remap = [
        (41, 401),
        (42, 402),
        (44, 404),
        (45, 405),
        (46, 406),
        (47, 407),
    ];

    let cloned = source.clone_with_identity_remap(401, &remap, &replacement_messages)?;

    // The source is borrowed by the clone operation and remains byte-exact.
    assert_eq!(archive_object_bytes(&source)?, source_bytes);
    assert_eq!(cloned.archive_info.identifier, Some(401));
    assert_eq!(cloned.messages, replacement_messages);
    assert_eq!(cloned.messages.len(), source.messages.len());
    assert_eq!(
        cloned
            .archive_info
            .message_infos
            .iter()
            .map(|info| info.type_)
            .collect::<Vec<_>>(),
        [101, 202]
    );

    let first = &cloned.archive_info.message_infos[0];
    assert_eq!(first.object_references, [401, 402, 43]);
    assert_eq!(first.data_references, [9001]);
    assert_eq!(first.field_infos[0].object_references, [404, 405]);
    assert_eq!(first.field_infos[0].data_references, [901]);
    assert_eq!(first.field_infos[1].object_references, [406, 43]);
    assert_eq!(first.field_infos[1].data_references, [902]);

    let second = &cloned.archive_info.message_infos[1];
    assert_eq!(second.object_references, [407, 48]);
    assert_eq!(second.data_references, [9002]);
    assert_eq!(second.field_infos[0].object_references, [401, 49]);
    assert_eq!(second.field_infos[0].data_references, [903]);

    // Reparse the clone so the retained raw header is exercised by the public
    // serializer. Known type/length spans changed, but unrelated unknown and
    // non-canonical spans remain present in their original byte form.
    let cloned_bytes = archive_object_bytes(&cloned)?;
    let cloned_header = archive_header(&cloned_bytes);
    for marker in [ROOT_UNKNOWN, MESSAGE_ZERO_UNKNOWN, MESSAGE_ONE_UNKNOWN] {
        assert!(
            cloned_header
                .windows(marker.len())
                .any(|window| window == marker),
            "retained raw header marker missing: {marker:?}"
        );
    }
    let reparsed = Archive::parse(&cloned_bytes)?;
    assert_eq!(reparsed.objects.len(), 1);
    assert_eq!(reparsed.objects[0].messages, replacement_messages);
    assert_eq!(reparsed.objects[0].archive_info.identifier, Some(401));
    assert_eq!(
        reparsed.objects[0].archive_info.message_infos[0].object_references,
        [401, 402, 43]
    );
    assert_eq!(
        reparsed.objects[0].archive_info.message_infos[0].field_infos[0].object_references,
        [404, 405]
    );
    assert_eq!(
        reparsed.objects[0].archive_info.message_infos[0].data_references,
        [9001]
    );
    Ok(())
}

#[test]
fn clone_identity_self_map_is_byte_consistent_and_does_not_mutate_source() -> Result<(), Error> {
    let (source, source_bytes) = source_object_with_adversarial_header()?;
    let replacement_messages = source.messages.clone();
    let cloned = source.clone_with_identity_remap(
        41,
        &[(41, 41), (42, 42), (44, 44), (45, 45), (46, 46), (47, 47)],
        &replacement_messages,
    )?;

    assert_eq!(archive_object_bytes(&source)?, source_bytes);
    assert_eq!(archive_object_bytes(&cloned)?, source_bytes);
    assert_eq!(cloned.archive_info, source.archive_info);
    assert_eq!(cloned.messages, source.messages);
    Ok(())
}

#[test]
fn clone_rejects_invalid_remaps_and_replacement_arity_atomically() -> Result<(), Error> {
    let (source, source_bytes) = source_object_with_adversarial_header()?;
    let replacement_messages = source.messages.clone();
    let valid_remap = [
        (41, 401),
        (42, 402),
        (44, 404),
        (45, 405),
        (46, 406),
        (47, 407),
    ];

    let invalid_remaps: &[&[(u64, u64)]] = &[
        &[(41, 401), (41, 402)], // duplicate source
        &[(41, 0)],              // zero target
        &[(0, 401)],             // zero source
        &[(41, 401), (42, 401)], // target collision
        &[(41, 402)],            // source identity does not match new_identifier
        &[(42, 401)],            // destination collides with implicit clone identity
    ];
    for remap in invalid_remaps {
        assert!(
            source
                .clone_with_identity_remap(401, remap, &replacement_messages)
                .is_err()
        );
        assert_eq!(archive_object_bytes(&source)?, source_bytes);
    }

    let too_few_messages = vec![replacement_messages[0].clone()];
    assert!(
        source
            .clone_with_identity_remap(401, &valid_remap, &too_few_messages)
            .is_err()
    );
    assert_eq!(archive_object_bytes(&source)?, source_bytes);

    let mut wrong_type = replacement_messages.clone();
    wrong_type[1].type_ = 203;
    assert!(
        source
            .clone_with_identity_remap(401, &valid_remap, &wrong_type)
            .is_err()
    );
    assert_eq!(archive_object_bytes(&source)?, source_bytes);
    Ok(())
}

#[test]
fn clone_with_finite_limits_fails_before_mutating_the_source() -> Result<(), Error> {
    let (source, source_bytes) = source_object_with_adversarial_header()?;
    let replacement_messages = vec![
        RawMessage {
            type_: 101,
            data: vec![0xa5; 130],
        },
        RawMessage {
            type_: 202,
            data: vec![0xb6; 3],
        },
    ];
    let remap = [
        (41, 401),
        (42, 402),
        (44, 404),
        (45, 405),
        (46, 406),
        (47, 407),
    ];

    let limits = ArchiveLimits::default().with_message_bytes(3)?;
    let error = source
        .clone_with_identity_remap_with_limits(401, &remap, &replacement_messages, limits)
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::MessageBytes,
            ..
        })
    ));
    assert_eq!(archive_object_bytes(&source)?, source_bytes);

    let limits = ArchiveLimits::default().with_metadata_items(1)?;
    let error = source
        .clone_with_identity_remap_with_limits(401, &remap, &replacement_messages, limits)
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::MetadataItems,
            ..
        })
    ));
    assert_eq!(archive_object_bytes(&source)?, source_bytes);

    let limits = ArchiveLimits::default().with_header_memory_bytes(1)?;
    let error = source
        .clone_with_identity_remap_with_limits(401, &remap, &replacement_messages, limits)
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::HeaderMemoryBytes,
            ..
        })
    ));
    assert_eq!(archive_object_bytes(&source)?, source_bytes);

    // The retained source header fits exactly. The larger identity and the
    // 3-to-130 length transition require additional encoded header bytes, so
    // the clone must reject the rewrite before publishing it.
    let limits = ArchiveLimits::default().with_header_bytes(archive_header(&source_bytes).len())?;
    let wide_identifier = u64::from(u32::MAX) + 1;
    let wide_remap = [
        (41, wide_identifier),
        (42, 402),
        (44, 404),
        (45, 405),
        (46, 406),
        (47, 407),
    ];
    let error = source
        .clone_with_identity_remap_with_limits(
            wide_identifier,
            &wide_remap,
            &replacement_messages,
            limits,
        )
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::HeaderBytes,
            ..
        })
    ));
    assert_eq!(archive_object_bytes(&source)?, source_bytes);

    // The source object fits this ceiling, while the replacement grows the
    // aggregate object beyond it. Each replacement message remains below the
    // default per-message ceiling, so this specifically exercises the
    // aggregate ObjectBytes check and its atomic failure path.
    let grown_messages = vec![
        RawMessage {
            type_: 101,
            data: vec![0xa5; 200],
        },
        RawMessage {
            type_: 202,
            data: vec![0xb6; 200],
        },
    ];
    let limits = ArchiveLimits::default().with_object_bytes(source_bytes.len())?;
    let error = source
        .clone_with_identity_remap_with_limits(401, &remap, &grown_messages, limits)
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::ObjectBytes,
            ..
        })
    ));
    assert_eq!(archive_object_bytes(&source)?, source_bytes);
    Ok(())
}

#[test]
fn clone_refuses_merge_and_diff_metadata_without_mutating_source() -> Result<(), Error> {
    let replacement_messages = vec![RawMessage {
        type_: 101,
        data: vec![0xa5; 3],
    }];
    let remap = [(41, 401)];

    for source in [
        source_object_with_merge_metadata(true, false, false)?,
        source_object_with_merge_metadata(false, true, false)?,
        source_object_with_merge_metadata(false, false, true)?,
    ] {
        let (source, source_bytes) = source;
        assert!(
            source
                .clone_with_identity_remap(401, &remap, &replacement_messages)
                .is_err()
        );
        assert_eq!(archive_object_bytes(&source)?, source_bytes);
    }
    Ok(())
}

fn source_object_with_adversarial_header() -> Result<(ArchiveObject, Vec<u8>), Error> {
    let mut object = ArchiveObject::new(
        41,
        vec![
            RawMessage {
                type_: 101,
                data: vec![0x11; 3],
            },
            RawMessage {
                type_: 202,
                data: vec![0x22; 130],
            },
        ],
    )?;
    let first = &mut object.archive_info.message_infos[0];
    first.object_references = vec![41, 42, 43];
    first.data_references = vec![9001];
    first.field_infos = vec![
        FieldInfo {
            path: FieldPath::new(vec![7, 1]),
            r#type: Some(FieldType::ObjectReference),
            object_references: vec![44, 45],
            data_references: vec![901],
            ..FieldInfo::default()
        },
        FieldInfo {
            path: FieldPath::new(vec![7, 2]),
            object_references: vec![46, 43],
            data_references: vec![902],
            ..FieldInfo::default()
        },
    ];
    let second = &mut object.archive_info.message_infos[1];
    second.object_references = vec![47, 48];
    second.data_references = vec![9002];
    second.field_infos = vec![FieldInfo {
        path: FieldPath::new(vec![8]),
        object_references: vec![41, 49],
        data_references: vec![903],
        ..FieldInfo::default()
    }];
    object.validate()?;

    let archive = Archive {
        objects: vec![object],
    };
    let canonical = archive.to_bytes()?;
    let mut message_zero_suffix = vec![
        0x08, 0xe5, 0x00, // duplicate type = 101, overlong value
        0x18, 0x83, 0x00, // duplicate length = 3, overlong value
    ];
    message_zero_suffix.extend_from_slice(MESSAGE_ZERO_UNKNOWN);
    let with_message_zero = append_to_nested_message_info(&canonical, 0, &message_zero_suffix);
    let mut message_one_suffix = vec![
        0x08, 0xca, 0x01, // duplicate type = 202, overlong value
        0x18, 0x82, 0x01, // duplicate length = 130, overlong value
    ];
    message_one_suffix.extend_from_slice(MESSAGE_ONE_UNKNOWN);
    let with_message_zero =
        append_to_nested_message_info(&with_message_zero, 1, &message_one_suffix);
    let modified = append_to_archive_header(&with_message_zero, ROOT_UNKNOWN);
    let parsed = Archive::parse(&modified)?;
    assert_eq!(parsed.objects.len(), 1);
    assert_eq!(archive_object_bytes(&parsed.objects[0])?, modified);
    Ok((
        parsed
            .objects
            .into_iter()
            .next()
            .expect("one fixture object"),
        modified,
    ))
}

fn source_object_with_merge_metadata(
    should_merge: bool,
    diff_metadata: bool,
    base_metadata: bool,
) -> Result<(ArchiveObject, Vec<u8>), Error> {
    let mut object = ArchiveObject::new(
        41,
        vec![RawMessage {
            type_: 101,
            data: vec![0x11; 3],
        }],
    )?;
    object.archive_info.should_merge = should_merge.then_some(true);
    if diff_metadata {
        object.archive_info.message_infos[0].diff_merge_version = vec![1, 5];
    }
    if base_metadata {
        object.archive_info.message_infos[0].base_message_index = Some(0);
    }
    object.validate()?;
    let bytes = archive_object_bytes(&object)?;
    Ok((object, bytes))
}

fn archive_object_bytes(object: &ArchiveObject) -> Result<Vec<u8>, Error> {
    Archive {
        objects: vec![object.clone()],
    }
    .to_bytes()
}

fn archive_header(bytes: &[u8]) -> &[u8] {
    let (length, prefix) = read_varint(bytes, 0);
    let length = usize::try_from(length).expect("fixture header length fits usize");
    &bytes[prefix..prefix + length]
}

fn append_to_archive_header(bytes: &[u8], suffix: &[u8]) -> Vec<u8> {
    let (length, prefix) = read_varint(bytes, 0);
    let length = usize::try_from(length).expect("fixture header length fits usize");
    let header = &bytes[prefix..prefix + length];
    let mut new_header = Vec::with_capacity(header.len() + suffix.len());
    new_header.extend_from_slice(header);
    new_header.extend_from_slice(suffix);
    rebuild_archive(new_header, &bytes[prefix + length..])
}

fn append_to_nested_message_info(bytes: &[u8], occurrence: usize, suffix: &[u8]) -> Vec<u8> {
    let (length, prefix) = read_varint(bytes, 0);
    let length = usize::try_from(length).expect("fixture header length fits usize");
    let header = &bytes[prefix..prefix + length];
    let mut cursor = 0usize;
    let mut seen = 0usize;
    let mut rewritten = Vec::with_capacity(header.len() + suffix.len());
    while cursor < header.len() {
        let field_start = cursor;
        let (key, key_end) = read_varint(header, cursor);
        cursor = key_end;
        let field_number = key >> 3;
        let wire_type = key & 7;
        let (payload_start, field_end) = wire_payload_bounds(header, cursor, wire_type);
        if field_number == 2 && wire_type == 2 && seen == occurrence {
            let (nested_length, _) = read_varint(header, cursor);
            let nested_length = usize::try_from(nested_length).expect("nested length fits usize");
            let nested = &header[payload_start..field_end];
            assert_eq!(nested_length, nested.len());
            let new_length = nested
                .len()
                .checked_add(suffix.len())
                .expect("fixture length does not overflow");
            rewritten.extend_from_slice(&header[field_start..cursor]);
            push_varint(
                &mut rewritten,
                u64::try_from(new_length).expect("length fits u64"),
            );
            rewritten.extend_from_slice(nested);
            rewritten.extend_from_slice(suffix);
            seen += 1;
        } else {
            rewritten.extend_from_slice(&header[field_start..field_end]);
            if field_number == 2 && wire_type == 2 {
                seen += 1;
            }
        }
        cursor = field_end;
    }
    assert!(seen > occurrence, "expected nested MessageInfo occurrence");
    rebuild_archive(rewritten, &bytes[prefix + length..])
}

fn rebuild_archive(header: Vec<u8>, payload: &[u8]) -> Vec<u8> {
    let mut output = Vec::new();
    push_varint(
        &mut output,
        u64::try_from(header.len()).expect("fixture header length fits u64"),
    );
    output.extend_from_slice(&header);
    output.extend_from_slice(payload);
    output
}

fn wire_payload_bounds(bytes: &[u8], payload_start: usize, wire_type: u64) -> (usize, usize) {
    match wire_type {
        0 => {
            let (_, end) = read_varint(bytes, payload_start);
            (payload_start, end)
        },
        1 => {
            let end = payload_start.checked_add(8).expect("fixture range fits");
            assert!(end <= bytes.len());
            (payload_start, end)
        },
        2 => {
            let (length, length_end) = read_varint(bytes, payload_start);
            let length = usize::try_from(length).expect("fixture length fits usize");
            let end = length_end.checked_add(length).expect("fixture range fits");
            assert!(end <= bytes.len());
            (length_end, end)
        },
        5 => {
            let end = payload_start.checked_add(4).expect("fixture range fits");
            assert!(end <= bytes.len());
            (payload_start, end)
        },
        other => panic!("unsupported fixture wire type {other}"),
    }
}

fn read_varint(bytes: &[u8], mut offset: usize) -> (u64, usize) {
    let start = offset;
    let mut value = 0u64;
    let mut shift = 0u32;
    loop {
        let byte = *bytes.get(offset).expect("complete fixture varint");
        offset += 1;
        value |= u64::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return (value, offset);
        }
        shift += 7;
        assert!(shift < 64, "fixture varint overflow at {start}");
    }
}

fn push_varint(bytes: &mut Vec<u8>, mut value: u64) {
    while value >= 0x80 {
        bytes.push((value as u8 & 0x7f) | 0x80);
        value >>= 7;
    }
    bytes.push(value as u8);
}
