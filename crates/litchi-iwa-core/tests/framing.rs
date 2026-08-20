#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    reason = "Test fixtures intentionally generate byte patterns from bounded counters."
)]
#![allow(
    clippy::shadow_unrelated,
    reason = "Short-lived error assertions keep each malformed-input case independent."
)]

use litchi_iwa_core::{
    Archive, ArchiveLimits, ArchiveObject, Error, FieldInfo, FieldPath, FieldType, HeaderKind,
    HeaderOperation, KnownFieldRule, LimitKind, RawMessage, SnappyLimits, SnappyStream,
    UnknownFieldRule,
};
use prost::Message;
use std::io::{self, Cursor, Read};

struct FailingReader;

impl Read for FailingReader {
    fn read(&mut self, _buffer: &mut [u8]) -> io::Result<usize> {
        Err(io::Error::new(
            io::ErrorKind::UnexpectedEof,
            "test reader failure",
        ))
    }
}

#[test]
fn round_trip_spans_multiple_frames() -> Result<(), Error> {
    let input: Vec<u8> = (0..(SnappyStream::WRITE_CHUNK_SIZE * 2 + 17))
        .map(|value| value as u8)
        .collect();

    let encoded = SnappyStream::compress(&input)?;
    let decoded = SnappyStream::decompress(&encoded)?;

    assert_eq!(decoded.as_bytes(), input.as_slice());
    assert_eq!(decoded.into_bytes(), input);
    Ok(())
}

#[test]
fn empty_stream_round_trips() -> Result<(), Error> {
    let encoded = SnappyStream::compress(&[])?;
    assert!(encoded.is_empty());
    assert!(SnappyStream::decompress(&encoded)?.as_bytes().is_empty());
    Ok(())
}

#[test]
fn truncated_header_and_payload_are_rejected() -> Result<(), Error> {
    let error = SnappyStream::decompress(&[0, 1, 2]).err();
    assert!(matches!(
        error,
        Some(Error::InvalidArchive {
            reason: "truncated Snappy frame header",
            ..
        })
    ));

    let encoded = SnappyStream::compress(b"truncated")?;
    let error = SnappyStream::decompress(&encoded[..encoded.len() - 1]).err();
    assert!(matches!(
        error,
        Some(Error::InvalidArchive {
            reason: "truncated Snappy frame payload",
            ..
        })
    ));
    Ok(())
}

#[test]
fn configured_limits_are_enforced() -> Result<(), Error> {
    let encoded = SnappyStream::compress(b"bounded")?;
    let limits = SnappyLimits::new(1, 7)?;
    let error = SnappyStream::decompress_with_limits(&encoded, limits).err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::SnappyChunkBytes,
            ..
        })
    ));

    let error = SnappyLimits::new(8, 7).err();
    assert!(matches!(error, Some(Error::InvalidLimits { .. })));
    let error = SnappyLimits::default().with_input_limits(8, 11, 1).err();
    assert!(matches!(error, Some(Error::InvalidLimits { .. })));
    Ok(())
}

#[test]
fn archive_limits_are_copyable_and_checked() -> Result<(), Error> {
    let limits = ArchiveLimits::default().with_objects(3)?;
    let copy = limits;
    assert_eq!(limits.max_objects(), 3);
    assert_eq!(copy.max_objects(), 3);

    let error = ArchiveLimits::default()
        .with_archive_bytes(ArchiveLimits::MAX_ARCHIVE_BYTES + 1)
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::ArchiveBytes,
            ..
        })
    ));
    Ok(())
}

#[test]
fn malformed_large_length_is_rejected_without_arithmetic_overflow() {
    let error = SnappyStream::decompress(&[0, 0xff, 0xff, 0xff]).err();
    assert!(matches!(
        error,
        Some(Error::InvalidArchive {
            reason: "truncated Snappy frame payload",
            ..
        })
    ));
}

#[test]
fn archive_round_trip_preserves_order_and_payloads() -> Result<(), Error> {
    let archive = Archive {
        objects: vec![
            ArchiveObject::new(
                17,
                vec![
                    RawMessage {
                        type_: 100,
                        data: b"first".to_vec(),
                    },
                    RawMessage {
                        type_: 200,
                        data: (0..300).map(|value| value as u8).collect(),
                    },
                ],
            )?,
            ArchiveObject::new(
                23,
                vec![RawMessage {
                    type_: 300,
                    data: b"last".to_vec(),
                }],
            )?,
        ],
    };

    let encoded = archive.to_bytes()?;
    let parsed = Archive::parse(&encoded)?;
    assert_eq!(parsed.to_bytes()?, encoded);
    assert_eq!(parsed.objects[0].archive_info.identifier, Some(17));
    assert_eq!(parsed.objects[1].primary_message_type(), Some(300));
    assert_eq!(parsed.objects[0].messages[1].data.len(), 300);
    Ok(())
}

#[test]
fn selected_archive_parse_retains_only_requested_objects_but_validates_all_frames()
-> Result<(), Error> {
    let archive = Archive {
        objects: vec![
            ArchiveObject::new(
                9,
                vec![RawMessage {
                    type_: 100,
                    data: vec![0x11; 1024],
                }],
            )?,
            ArchiveObject::new(
                1,
                vec![RawMessage {
                    type_: 200,
                    data: b"selected".to_vec(),
                }],
            )?,
        ],
    };
    let encoded = archive.to_bytes()?;

    let selected = Archive::parse_objects_with_identifier(&encoded, 1)?;
    assert_eq!(selected.objects.len(), 1);
    assert_eq!(selected.objects[0].archive_info.identifier, Some(1));
    assert_eq!(selected.objects[0].messages[0].data, b"selected");

    let discarded = Archive {
        objects: vec![archive.objects[0].clone()],
    }
    .to_bytes()?;
    let selected_only = Archive {
        objects: vec![archive.objects[1].clone()],
    }
    .to_bytes()?;
    let duplicate_discarded = [
        discarded.as_slice(),
        discarded.as_slice(),
        selected_only.as_slice(),
    ]
    .concat();
    assert!(matches!(
        Archive::parse_objects_with_identifier(&duplicate_discarded, 1),
        Err(Error::InvalidArchive {
            reason: "duplicate object identifier",
            ..
        })
    ));

    let mut malformed_discarded = discarded;
    malformed_discarded.extend_from_slice(&[0x01, 0xff]);
    assert!(matches!(
        Archive::parse_objects_with_identifier(&malformed_discarded, 1),
        Err(Error::HeaderCodec {
            header: HeaderKind::ArchiveInfo,
            operation: HeaderOperation::Decode,
            ..
        })
    ));

    let mut truncated = encoded;
    truncated.pop();
    assert!(matches!(
        Archive::parse_objects_with_identifier(&truncated, 1),
        Err(Error::InvalidArchive {
            reason: "truncated message payload",
            ..
        })
    ));

    let limited = ArchiveLimits::default().with_objects(1)?;
    assert!(matches!(
        Archive::parse_objects_with_identifier_with_limits(&archive.to_bytes()?, 1, limited),
        Err(Error::Limit {
            kind: LimitKind::Objects,
            observed: 2,
            maximum: 1,
        })
    ));
    Ok(())
}

#[test]
fn archive_rejects_malformed_lengths_and_truncated_payloads() -> Result<(), Error> {
    let error = Archive::parse(&[0x80]).err();
    assert!(matches!(
        error,
        Some(Error::InvalidArchive {
            reason: "truncated archive length varint",
            ..
        })
    ));

    let error = Archive::parse(&[0xff; 10]).err();
    assert!(matches!(
        error,
        Some(Error::InvalidArchive {
            reason: "archive length varint overflow",
            ..
        })
    ));

    let archive = Archive {
        objects: vec![ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 9,
                data: b"payload".to_vec(),
            }],
        )?],
    };
    let mut encoded = archive.to_bytes()?;
    encoded.pop();
    let error = Archive::parse(&encoded).err();
    assert!(matches!(
        error,
        Some(Error::InvalidArchive {
            reason: "truncated message payload",
            ..
        })
    ));
    Ok(())
}

#[test]
fn archive_limits_bound_objects_messages_payloads_and_metadata() -> Result<(), Error> {
    let archive = Archive {
        objects: vec![ArchiveObject::new(
            1,
            vec![RawMessage {
                type_: 9,
                data: b"payload".to_vec(),
            }],
        )?],
    };
    let encoded = archive.to_bytes()?;

    let error = ArchiveLimits::default().with_objects(0).err();
    assert!(matches!(error, Some(Error::InvalidLimits { .. })));

    let error = Archive::parse_with_limits(
        &encoded,
        ArchiveLimits::default().with_messages_per_object(1)?,
    )
    .err();
    assert!(error.is_none());

    let error =
        Archive::parse_with_limits(&encoded, ArchiveLimits::default().with_message_bytes(6)?).err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::MessageBytes,
            ..
        })
    ));

    let error =
        Archive::parse_with_limits(&encoded, ArchiveLimits::default().with_metadata_items(3)?)
            .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::MetadataItems,
            ..
        })
    ));

    let error = Archive::parse_with_limits(
        &encoded,
        ArchiveLimits::default().with_object_bytes(encoded.len() - 1)?,
    )
    .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::ObjectBytes,
            ..
        })
    ));
    Ok(())
}

#[test]
fn archive_crud_rejects_duplicates_and_preserves_order() -> Result<(), Error> {
    let first = ArchiveObject::new(
        11,
        vec![RawMessage {
            type_: 1,
            data: b"first".to_vec(),
        }],
    )?;
    let second = ArchiveObject::new(
        22,
        vec![RawMessage {
            type_: 2,
            data: b"second".to_vec(),
        }],
    )?;
    let duplicate = ArchiveObject::new(
        11,
        vec![RawMessage {
            type_: 3,
            data: b"duplicate".to_vec(),
        }],
    )?;
    let mut archive = Archive::new();
    archive.insert_object(first)?;
    archive.insert_object(second)?;
    assert_eq!(
        archive
            .object(11)
            .and_then(ArchiveObject::primary_message_type),
        Some(1)
    );
    assert_eq!(archive.objects.len(), 2);

    let error = archive.insert_object(duplicate.clone()).err();
    assert!(matches!(
        error,
        Some(Error::InvalidArchive {
            reason: "duplicate object identifier",
            ..
        })
    ));
    assert_eq!(archive.objects.len(), 2);

    let previous = archive.upsert_object(duplicate)?;
    assert_eq!(
        previous.and_then(|object| object.primary_message_type()),
        Some(1)
    );
    assert_eq!(
        archive
            .object(11)
            .and_then(ArchiveObject::primary_message_type),
        Some(3)
    );

    let removed = archive.remove_object(22);
    assert_eq!(
        removed.and_then(|object| object.primary_message_type()),
        Some(2)
    );
    assert!(archive.object(22).is_none());
    assert_eq!(archive.objects.len(), 1);
    Ok(())
}

#[test]
fn archive_validation_rejects_duplicate_structural_ids() -> Result<(), Error> {
    let object = ArchiveObject::new(
        7,
        vec![RawMessage {
            type_: 1,
            data: vec![1],
        }],
    )?;
    let archive = Archive {
        objects: vec![object.clone(), object],
    };
    assert!(matches!(
        archive.validate(),
        Err(Error::InvalidArchive {
            reason: "duplicate object identifier",
            ..
        })
    ));
    Ok(())
}

#[test]
fn message_mutations_are_atomic_and_update_metadata() -> Result<(), Error> {
    let mut object = ArchiveObject::new(
        9,
        vec![RawMessage {
            type_: 10,
            data: b"old".to_vec(),
        }],
    )?;
    let before_info = object.archive_info.clone();
    let limits = ArchiveLimits::default().with_message_bytes(3)?;
    let error = object
        .replace_message_with_limits(
            0,
            RawMessage {
                type_: 20,
                data: b"newer".to_vec(),
            },
            limits,
        )
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::MessageBytes,
            ..
        })
    ));
    assert_eq!(object.messages[0].type_, 10);
    assert_eq!(object.messages[0].data, b"old");
    assert_eq!(object.archive_info, before_info);

    let old = object.replace_message(
        0,
        RawMessage {
            type_: 20,
            data: b"new".to_vec(),
        },
    )?;
    assert_eq!(old.data, b"old");
    assert_eq!(object.archive_info.message_infos[0].type_, 20);
    assert_eq!(object.archive_info.message_infos[0].length, 3);

    let count_limited = ArchiveLimits::default().with_messages_per_object(1)?;
    let before_push = object.clone();
    let error = object
        .push_message_with_limits(
            RawMessage {
                type_: 30,
                data: vec![4],
            },
            count_limited,
        )
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::MessagesPerObject,
            ..
        })
    ));
    assert_eq!(object.messages, before_push.messages);
    assert_eq!(object.archive_info, before_push.archive_info);

    object.push_message(RawMessage {
        type_: 30,
        data: vec![4],
    })?;
    assert_eq!(object.messages.len(), 2);
    assert_eq!(
        object.remove_message(0).map(|message| message.type_),
        Some(20)
    );
    assert_eq!(object.messages.len(), 1);
    assert!(object.remove_message(9).is_none());
    Ok(())
}

#[test]
fn upsert_failure_rolls_back_without_replacing_existing_object() -> Result<(), Error> {
    let mut archive = Archive::new();
    archive.insert_object(ArchiveObject::new(
        1,
        vec![RawMessage {
            type_: 1,
            data: vec![1],
        }],
    )?)?;
    let replacement = ArchiveObject::new(
        1,
        vec![RawMessage {
            type_: 2,
            data: vec![2, 3, 4],
        }],
    )?;
    let limits = ArchiveLimits::default().with_message_bytes(1)?;
    let error = archive.upsert_object_with_limits(replacement, limits).err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::MessageBytes,
            ..
        })
    ));
    assert_eq!(
        archive
            .object(1)
            .and_then(ArchiveObject::primary_message_type),
        Some(1)
    );
    assert_eq!(archive.objects.len(), 1);
    Ok(())
}

#[test]
fn reader_headers_are_bounded_and_truncation_is_structured() -> Result<(), Error> {
    let archive = Archive {
        objects: vec![ArchiveObject::new(
            31,
            vec![RawMessage {
                type_: 41,
                data: b"payload".to_vec(),
            }],
        )?],
    };
    let encoded = archive.to_bytes()?;
    let (header_length, prefix_length) = decode_test_varint(&encoded);
    let mut reader = Cursor::new(&encoded[prefix_length..prefix_length + header_length]);
    let info = litchi_iwa_core::ArchiveInfo::parse(&mut reader)?;
    assert_eq!(info.identifier, Some(31));

    let mut truncated = Cursor::new(vec![0x80]);
    assert!(matches!(
        litchi_iwa_core::ArchiveInfo::parse(&mut truncated),
        Err(Error::HeaderCodec {
            header: HeaderKind::ArchiveInfo,
            operation: HeaderOperation::Decode,
            ..
        })
    ));

    let mut oversized = Cursor::new(vec![0, 0]);
    let limits = ArchiveLimits::default().with_header_bytes(1)?;
    assert!(matches!(
        litchi_iwa_core::ArchiveInfo::parse_with_limits(&mut oversized, limits),
        Err(Error::Limit {
            kind: LimitKind::HeaderBytes,
            ..
        })
    ));
    Ok(())
}

#[test]
fn message_info_reader_and_io_errors_are_reported() -> Result<(), Error> {
    // required type=61, packed version=[1, 0, 5], required length=2
    let encoded = [0x08, 0x3d, 0x12, 0x03, 0x01, 0x00, 0x05, 0x18, 0x02];
    let mut reader = Cursor::new(encoded);
    let info = litchi_iwa_core::MessageInfo::parse(&mut reader)?;
    assert_eq!(info.type_, 61);
    assert_eq!(info.length, 2);

    let mut failing = FailingReader;
    assert!(matches!(
        litchi_iwa_core::MessageInfo::parse(&mut failing),
        Err(Error::Io(_))
    ));
    Ok(())
}

#[test]
fn neutral_field_metadata_preserves_presence_and_unknown_enum_values() -> Result<(), Error> {
    let absent = FieldInfo::new(vec![1]);
    assert_eq!(absent.effective_type(), FieldType::Value);
    assert_eq!(
        absent.effective_unknown_field_rule(),
        UnknownFieldRule::IgnoreAndPreserveUntilModified
    );
    assert_eq!(absent.effective_known_field_rule(), KnownFieldRule::None);

    let explicit_defaults = FieldInfo {
        path: FieldPath::new(vec![2]),
        r#type: Some(FieldType::Value),
        unknown_field_rule: Some(UnknownFieldRule::IgnoreAndPreserveUntilModified),
        known_field_rule: Some(KnownFieldRule::None),
        ..Default::default()
    };
    let known = FieldInfo {
        path: FieldPath::new(vec![3]),
        r#type: Some(FieldType::Message),
        unknown_field_rule: Some(UnknownFieldRule::NotSupported),
        known_field_rule: Some(KnownFieldRule::PreserveNewerValue),
        ..Default::default()
    };
    let unrecognized = FieldInfo {
        path: FieldPath::new(vec![4]),
        r#type: Some(FieldType::Unrecognized(99)),
        unknown_field_rule: Some(UnknownFieldRule::Unrecognized(-2)),
        known_field_rule: Some(KnownFieldRule::Unrecognized(77)),
        ..Default::default()
    };
    let expected = vec![absent, explicit_defaults, known, unrecognized];

    let mut object = ArchiveObject::new(
        91,
        vec![RawMessage {
            type_: 92,
            data: Vec::new(),
        }],
    )?;
    object.archive_info.message_infos[0].field_infos = expected.clone();
    let archive = Archive {
        objects: vec![object],
    };
    let encoded = archive.to_bytes()?;
    let parsed = Archive::parse(&encoded)?;
    assert_eq!(
        parsed.objects[0].archive_info.message_infos[0].field_infos,
        expected
    );
    assert_eq!(parsed.to_bytes()?, encoded);
    Ok(())
}

#[test]
fn missing_required_field_info_path_is_rejected_before_publication() {
    // required type=1, required length=0, then an empty FieldInfo without its
    // required path.
    let missing_path = [0x08, 0x01, 0x18, 0x00, 0x22, 0x00];
    assert!(matches!(
        litchi_iwa_core::MessageInfo::decode(&missing_path),
        Err(Error::HeaderCodec {
            header: HeaderKind::MessageInfo,
            operation: HeaderOperation::Decode,
            ..
        })
    ));
}

#[test]
fn unknown_and_noncanonical_header_bytes_round_trip_byte_for_byte() -> Result<(), Error> {
    let archive = Archive {
        objects: vec![ArchiveObject::new(
            71,
            vec![RawMessage {
                type_: 81,
                data: vec![0, 1, 2, 255],
            }],
        )?],
    };
    let encoded = archive.to_bytes()?;
    let (header_length, prefix_length) = decode_test_varint(&encoded);
    let duplicate_identifier = [0x08, 71];
    let unknown_field = [0xa2, 0x06, 0x01, 0xff];
    let noncanonical_header_suffix = [duplicate_identifier.as_slice(), &unknown_field].concat();
    let new_header_length = header_length + noncanonical_header_suffix.len();
    let mut modified = Vec::new();
    modified.push(new_header_length as u8);
    modified.extend_from_slice(&encoded[prefix_length..prefix_length + header_length]);
    modified.extend_from_slice(&noncanonical_header_suffix);
    modified.extend_from_slice(&encoded[prefix_length + header_length..]);

    let parsed = Archive::parse(&modified)?;
    assert_eq!(parsed.objects[0].messages[0].data, vec![0, 1, 2, 255]);
    assert_eq!(parsed.to_bytes()?, modified);
    let selected = Archive::parse_objects_with_identifier(&modified, 71)?;
    assert_eq!(selected.to_bytes()?, modified);
    Ok(())
}

#[test]
fn nested_duplicates_and_unknown_wire_payloads_preserve_the_raw_header() -> Result<(), Error> {
    let mut message_info = vec![
        0x08, 0x01, 0x08, 0x07, // duplicate required type; last wins
        0x18, 0x01, 0x18, 0x03, // duplicate required length; last wins
        0x90, 0x00, 0x81, 0x00, // noncanonical unpacked version = 1
    ];
    // Unknown varint, fixed64, length-delimited, and fixed32 records.
    message_info.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]);
    message_info.extend_from_slice(&[0xa9, 0x06, 1, 2, 3, 4, 5, 6, 7, 8]);
    message_info.extend_from_slice(&[0xb2, 0x06, 0x81, 0x00, 0xff]);
    message_info.extend_from_slice(&[0xc5, 0x06, 9, 10, 11, 12]);

    let mut header = vec![0x08, 0x01, 0x88, 0x00, 0xaa, 0x00];
    header.extend_from_slice(&[0x12, one_byte_length(message_info.len())]);
    header.extend_from_slice(&message_info);
    header.extend_from_slice(&[0xcd, 0x06, 13, 14, 15, 16]);

    let prost = match litchi_iwa_protos::tsp::ArchiveInfo::decode(header.as_slice()) {
        Ok(prost) => prost,
        Err(error) => panic!("adversarial header must remain valid protobuf: {error}"),
    };
    assert_eq!(prost.identifier, Some(42));
    assert_eq!(prost.message_infos[0].r#type, 7);
    assert_eq!(prost.message_infos[0].length, 3);
    assert_eq!(prost.message_infos[0].version, [1]);

    let mut source = vec![one_byte_length(header.len())];
    source.extend_from_slice(&header);
    source.extend_from_slice(&[0xde, 0xad, 0xbe]);
    let archive = Archive::parse(&source)?;
    assert_eq!(archive.objects[0].archive_info.identifier, Some(42));
    assert_eq!(archive.objects[0].messages[0].type_, 7);
    assert_eq!(archive.objects[0].messages[0].data, [0xde, 0xad, 0xbe]);
    assert_eq!(archive.to_bytes()?, source);
    Ok(())
}

#[test]
fn hostile_native_header_bytes_remain_the_no_op_authority() -> Result<(), Error> {
    let header = [
        0x08, 0xaa, 0x00, 0x90, 0x03, 0x09, 0x12, 0x8f, 0x00, 0x08, 0x87, 0x00, 0x98, 0x06, 0x96,
        0x01, 0x18, 0x8b, 0x00, 0xa2, 0x06, 0x81, 0x00, 0xff, 0x9a, 0x03, 0x82, 0x00, 0xde, 0xad,
        0x18, 0x81, 0x00,
    ];
    let mut source = vec![header.len() as u8];
    source.extend_from_slice(&header);
    source.extend_from_slice(&[0x5a; 11]);

    let archive = Archive::parse(&source)?;
    assert_eq!(archive.objects[0].archive_info.identifier, Some(42));
    assert_eq!(archive.objects[0].messages[0].type_, 7);
    assert_eq!(archive.to_bytes()?, source);
    Ok(())
}

#[test]
fn preserving_replacement_retains_adversarial_header_bytes_and_restores_exactly()
-> Result<(), Error> {
    let mut message_info = vec![
        0x08, 0x01, // duplicate type = 1
        0x88, 0x00, 0x87, 0x00, // effective type = 7, overlong key and value
        0x12, 0x03, 0x01, 0x00, 0x05, // conventional version
        0x18, 0x01, // duplicate length = 1
        0x98, 0x00, 0x83, 0x00, // effective length = 3, overlong key and value
    ];
    message_info.extend_from_slice(&[0xa0, 0x06, 0x81, 0x00]);
    message_info.extend_from_slice(&[0xa9, 0x06, 1, 2, 3, 4, 5, 6, 7, 8]);
    message_info.extend_from_slice(&[0xb2, 0x06, 0x81, 0x00, 0xff]);
    message_info.extend_from_slice(&[0xc5, 0x06, 9, 10, 11, 12]);

    let mut header = vec![
        0x08,
        0x01, // duplicate identifier = 1
        0x88,
        0x00,
        0xaa,
        0x00, // effective identifier = 42, overlong
        0xd1,
        0x06,
        1,
        2,
        3,
        4,
        5,
        6,
        7,
        8, // unknown fixed64
        0x92,
        0x00, // overlong MessageInfo field key
        0x80 | one_byte_length(message_info.len()),
        0x00, // overlong MessageInfo length prefix
    ];
    header.extend_from_slice(&message_info);
    header.extend_from_slice(&[0xda, 0x06, 0x81, 0x00, 0xee]);
    header.extend_from_slice(&[0x18, 0x81, 0x00]); // overlong should_merge = true

    let mut source = vec![one_byte_length(header.len())];
    source.extend_from_slice(&header);
    source.extend_from_slice(&[0xde, 0xad, 0xbe]);
    let mut archive = Archive::parse(&source)?;

    let old = archive.objects[0].replace_message_preserving_header(
        0,
        RawMessage {
            type_: 9,
            data: vec![0x5a; 130],
        },
    )?;
    assert_eq!(old.type_, 7);
    assert_eq!(old.data, [0xde, 0xad, 0xbe]);

    let changed = archive.to_bytes()?;
    let (changed_header_length, changed_prefix_length) = decode_test_varint(&changed);
    let changed_header =
        &changed[changed_prefix_length..changed_prefix_length + changed_header_length];
    let mut expected_header = header.clone();
    assert!(replace_unique_bytes(
        &mut expected_header,
        &[0x87, 0x00],
        &[0x89, 0x00]
    ));
    assert!(replace_unique_bytes(
        &mut expected_header,
        &[0x83, 0x00],
        &[0x82, 0x01]
    ));
    assert_eq!(changed_header, expected_header);

    let reparsed = Archive::parse(&changed)?;
    assert_eq!(reparsed.objects[0].messages[0].type_, 9);
    assert_eq!(reparsed.objects[0].messages[0].data, vec![0x5a; 130]);

    let replaced = archive.objects[0].replace_message_preserving_header(0, old)?;
    assert_eq!(replaced.type_, 9);
    assert_eq!(replaced.data, vec![0x5a; 130]);
    assert_eq!(archive.to_bytes()?, source);

    let object_before_failure = archive.objects[0].clone();
    let object_limit = header.len() + 3;
    let rollback_limits = ArchiveLimits::default()
        .with_object_bytes(object_limit)?
        .with_header_bytes(header.len())?;
    let error = archive.objects[0]
        .replace_message_preserving_header_with_limits(
            0,
            RawMessage {
                type_: 9,
                data: vec![0xaa, 0xbb, 0xcc],
            },
            rollback_limits,
        )
        .err();
    assert!(matches!(
        error,
        Some(Error::Limit {
            kind: LimitKind::ObjectBytes,
            observed,
            maximum,
        }) if observed == object_limit + 1 && maximum == object_limit
    ));
    assert_eq!(archive.objects[0], object_before_failure);
    assert_eq!(archive.to_bytes()?, source);
    Ok(())
}

#[test]
fn preserving_replacement_grows_scalar_and_enclosing_length_widths() -> Result<(), Error> {
    let mut message_info = vec![
        0x08, 0x01, // type = 1
        0x18, 0x01, // length = 1
        0xa2, 0x06, 0x78, // unknown field 100, 120-byte payload
    ];
    message_info.extend(std::iter::repeat_n(0x5a, 120));
    assert_eq!(message_info.len(), 127);

    let mut header = vec![
        0x08, 0x01, // identifier = 1
        0x12, 0x7f, // MessageInfo with a one-byte length prefix
    ];
    header.extend_from_slice(&message_info);
    let mut source = vec![0x83, 0x01]; // 131-byte ArchiveInfo
    source.extend_from_slice(&header);
    source.push(0xcc);
    let mut archive = Archive::parse(&source)?;

    archive.objects[0].replace_message_preserving_header(
        0,
        RawMessage {
            type_: 1,
            data: vec![0xdd; 130],
        },
    )?;
    let changed = archive.to_bytes()?;
    let (changed_header_length, changed_prefix_length) = decode_test_varint(&changed);
    assert_eq!(changed_header_length, 133);
    let changed_header =
        &changed[changed_prefix_length..changed_prefix_length + changed_header_length];

    let mut expected_message_info = vec![
        0x08, 0x01, // unchanged type
        0x18, 0x82, 0x01, // length = 130, now a two-byte scalar
        0xa2, 0x06, 0x78,
    ];
    expected_message_info.extend(std::iter::repeat_n(0x5a, 120));
    assert_eq!(expected_message_info.len(), 128);
    let mut expected_header = vec![
        0x08, 0x01, 0x12, 0x80, 0x01, // MessageInfo length = 128, now two bytes
    ];
    expected_header.extend_from_slice(&expected_message_info);
    assert_eq!(changed_header, expected_header);
    assert_eq!(
        &changed[changed_prefix_length + changed_header_length..],
        vec![0xdd; 130]
    );
    Ok(())
}

fn replace_unique_bytes(bytes: &mut [u8], from: &[u8], to: &[u8]) -> bool {
    assert_eq!(from.len(), to.len());
    let mut matches = bytes
        .windows(from.len())
        .enumerate()
        .filter_map(|(index, candidate)| (candidate == from).then_some(index));
    let Some(index) = matches.next() else {
        return false;
    };
    if matches.next().is_some() {
        return false;
    }
    bytes[index..index + to.len()].copy_from_slice(to);
    true
}

#[test]
fn buffa_preflight_rejects_malformed_deferred_children() {
    // Each case is a complete MessageInfo with one malformed deferred child.
    // A root-only lazy decode would otherwise publish the parent and postpone
    // the error until that particular child route was accessed.
    let malformed_messages: &[(&str, &[u8])] = &[
        (
            "FieldInfo body",
            &[0x08, 0x01, 0x18, 0x00, 0x22, 0x01, 0x80],
        ),
        (
            "FieldInfo.path",
            &[0x08, 0x01, 0x18, 0x00, 0x22, 0x03, 0x0a, 0x01, 0x80],
        ),
        (
            "diff_field_path",
            &[0x08, 0x01, 0x18, 0x00, 0x4a, 0x01, 0x80],
        ),
        (
            "fields_to_remove",
            &[0x08, 0x01, 0x18, 0x00, 0x52, 0x01, 0x80],
        ),
    ];

    for (context, malformed_message) in malformed_messages {
        assert!(
            matches!(
                litchi_iwa_core::MessageInfo::decode(malformed_message),
                Err(Error::HeaderCodec {
                    header: HeaderKind::MessageInfo,
                    operation: HeaderOperation::Decode,
                    ..
                })
            ),
            "malformed {context} escaped direct MessageInfo preflight"
        );

        let mut malformed_archive = vec![0x08, 0x01, 0x12];
        malformed_archive.push(one_byte_length(malformed_message.len()));
        malformed_archive.extend_from_slice(malformed_message);
        assert!(
            matches!(
                litchi_iwa_core::ArchiveInfo::decode(&malformed_archive),
                Err(Error::HeaderCodec {
                    header: HeaderKind::ArchiveInfo,
                    operation: HeaderOperation::Decode,
                    ..
                })
            ),
            "malformed {context} escaped ArchiveInfo preflight"
        );
    }
}

#[test]
fn buffa_preflight_enforces_field_nesting_and_memory_budgets() -> Result<(), Error> {
    let nested = litchi_iwa_protos::tsp::ArchiveInfo {
        identifier: Some(1),
        message_infos: vec![litchi_iwa_protos::tsp::MessageInfo {
            r#type: 7,
            length: 0,
            field_infos: vec![litchi_iwa_protos::tsp::FieldInfo {
                path: litchi_iwa_protos::tsp::FieldPath { path: vec![1, 2] },
                ..Default::default()
            }],
            ..Default::default()
        }],
        should_merge: None,
    }
    .encode_to_vec();
    assert!(litchi_iwa_core::ArchiveInfo::decode(&nested).is_ok());

    let nesting_limits = ArchiveLimits::default().with_header_nesting(2)?;
    assert!(matches!(
        litchi_iwa_core::ArchiveInfo::decode_with_limits(&nested, nesting_limits),
        Err(Error::Limit {
            kind: LimitKind::HeaderNesting,
            ..
        })
    ));

    let fields = [0x08, 0x01, 0x20, 0x00, 0x20, 0x00];
    let field_limits = ArchiveLimits::default().with_header_fields(2)?;
    assert!(matches!(
        litchi_iwa_core::ArchiveInfo::decode_with_limits(&fields, field_limits),
        Err(Error::Limit {
            kind: LimitKind::HeaderFields,
            ..
        })
    ));

    let one_empty_message = [0x08, 0x01, 0x12, 0x00];
    let memory_limits = ArchiveLimits::default().with_header_memory_bytes(1)?;
    assert!(matches!(
        litchi_iwa_core::ArchiveInfo::decode_with_limits(&one_empty_message, memory_limits),
        Err(Error::Limit {
            kind: LimitKind::HeaderMemoryBytes,
            ..
        })
    ));
    Ok(())
}

#[test]
fn buffa_preflight_counts_packed_metadata_before_decode() -> Result<(), Error> {
    // required type=1, packed version=[1, 2], required length=0
    let message_info = [0x08, 0x01, 0x12, 0x02, 0x01, 0x02, 0x18, 0x00];
    let limits = ArchiveLimits::default().with_metadata_items(1)?;
    assert!(matches!(
        litchi_iwa_core::MessageInfo::decode_with_limits(&message_info, limits),
        Err(Error::Limit {
            kind: LimitKind::MetadataItems,
            observed: 2,
            maximum: 1,
        })
    ));
    Ok(())
}

#[test]
fn buffa_preflight_charges_unrecognized_closed_enum_records() -> Result<(), Error> {
    let mut field_info = vec![0x0a, 0x00]; // present, empty required FieldPath
    for _ in 0..96 {
        // FieldInfo.type=99. Buffa retains each unrecognized closed-enum
        // occurrence in its unknown-field collection before the compatibility
        // projection selects the final semantic value.
        field_info.extend_from_slice(&[0x10, 0x63]);
    }
    let mut message_info = vec![0x08, 0x01, 0x18, 0x00, 0x22];
    message_info.extend_from_slice(&litchi_iwa_common::encode_varint(
        u64::try_from(field_info.len()).unwrap_or(u64::MAX),
    ));
    message_info.extend_from_slice(&field_info);

    let limits = ArchiveLimits::default().with_header_memory_bytes(4_096)?;
    assert!(matches!(
        litchi_iwa_core::MessageInfo::decode_with_limits(&message_info, limits),
        Err(Error::Limit {
            kind: LimitKind::HeaderMemoryBytes,
            observed,
            maximum: 4_096,
        }) if observed > 4_096
    ));
    Ok(())
}

#[test]
fn buffa_preflight_byte_and_field_limits_are_inclusive() -> Result<(), Error> {
    let minimal_message = [0x08, 0x01, 0x18, 0x00];

    let exact_bytes = ArchiveLimits::default().with_header_bytes(minimal_message.len())?;
    assert!(
        litchi_iwa_core::MessageInfo::decode_with_limits(&minimal_message, exact_bytes).is_ok()
    );
    let below_bytes = ArchiveLimits::default().with_header_bytes(minimal_message.len() - 1)?;
    assert!(matches!(
        litchi_iwa_core::MessageInfo::decode_with_limits(&minimal_message, below_bytes),
        Err(Error::Limit {
            kind: LimitKind::HeaderBytes,
            observed: 4,
            maximum: 3,
        })
    ));

    let exact_fields = ArchiveLimits::default().with_header_fields(2)?;
    assert!(
        litchi_iwa_core::MessageInfo::decode_with_limits(&minimal_message, exact_fields).is_ok()
    );
    let below_fields = ArchiveLimits::default().with_header_fields(1)?;
    assert!(matches!(
        litchi_iwa_core::MessageInfo::decode_with_limits(&minimal_message, below_fields),
        Err(Error::Limit {
            kind: LimitKind::HeaderFields,
            observed: 2,
            maximum: 1,
        })
    ));
    Ok(())
}

#[test]
fn buffa_preflight_nesting_and_memory_limits_are_inclusive() -> Result<(), Error> {
    let nested = litchi_iwa_protos::tsp::ArchiveInfo {
        identifier: Some(1),
        message_infos: vec![litchi_iwa_protos::tsp::MessageInfo {
            r#type: 7,
            length: 0,
            field_infos: vec![litchi_iwa_protos::tsp::FieldInfo {
                path: litchi_iwa_protos::tsp::FieldPath { path: vec![1] },
                ..Default::default()
            }],
            ..Default::default()
        }],
        should_merge: None,
    }
    .encode_to_vec();
    let exact_nesting = ArchiveLimits::default().with_header_nesting(3)?;
    assert!(litchi_iwa_core::ArchiveInfo::decode_with_limits(&nested, exact_nesting).is_ok());
    let below_nesting = ArchiveLimits::default().with_header_nesting(2)?;
    assert!(matches!(
        litchi_iwa_core::ArchiveInfo::decode_with_limits(&nested, below_nesting),
        Err(Error::Limit {
            kind: LimitKind::HeaderNesting,
            observed: 3,
            maximum: 2,
        })
    ));

    let one_message = [0x08, 0x01, 0x12, 0x04, 0x08, 0x01, 0x18, 0x00];
    let observed_memory = match litchi_iwa_core::ArchiveInfo::decode_with_limits(
        &one_message,
        ArchiveLimits::default().with_header_memory_bytes(1)?,
    ) {
        Err(Error::Limit {
            kind: LimitKind::HeaderMemoryBytes,
            observed,
            maximum: 1,
        }) => observed,
        other => panic!("expected exact header-memory observation, got {other:?}"),
    };
    assert!(observed_memory > 1);
    let exact_memory = ArchiveLimits::default().with_header_memory_bytes(observed_memory)?;
    assert!(litchi_iwa_core::ArchiveInfo::decode_with_limits(&one_message, exact_memory).is_ok());
    let below_memory = ArchiveLimits::default().with_header_memory_bytes(observed_memory - 1)?;
    assert!(matches!(
        litchi_iwa_core::ArchiveInfo::decode_with_limits(&one_message, below_memory),
        Err(Error::Limit {
            kind: LimitKind::HeaderMemoryBytes,
            observed,
            maximum,
        }) if observed == observed_memory && maximum == observed_memory - 1
    ));
    Ok(())
}

#[test]
fn buffa_preflight_metadata_and_message_count_limits_are_inclusive() -> Result<(), Error> {
    let packed_metadata = [0x08, 0x01, 0x12, 0x02, 0x01, 0x02, 0x18, 0x00];
    let exact_metadata = ArchiveLimits::default().with_metadata_items(2)?;
    assert!(
        litchi_iwa_core::MessageInfo::decode_with_limits(&packed_metadata, exact_metadata).is_ok()
    );
    let below_metadata = ArchiveLimits::default().with_metadata_items(1)?;
    assert!(matches!(
        litchi_iwa_core::MessageInfo::decode_with_limits(&packed_metadata, below_metadata),
        Err(Error::Limit {
            kind: LimitKind::MetadataItems,
            observed: 2,
            maximum: 1,
        })
    ));

    let two_messages = [
        0x08, 0x01, 0x12, 0x04, 0x08, 0x01, 0x18, 0x00, 0x12, 0x04, 0x08, 0x02, 0x18, 0x00,
    ];
    let exact_messages = ArchiveLimits::default().with_messages_per_object(2)?;
    assert!(
        litchi_iwa_core::ArchiveInfo::decode_with_limits(&two_messages, exact_messages).is_ok()
    );
    let below_messages = ArchiveLimits::default().with_messages_per_object(1)?;
    assert!(matches!(
        litchi_iwa_core::ArchiveInfo::decode_with_limits(&two_messages, below_messages),
        Err(Error::Limit {
            kind: LimitKind::MessagesPerObject,
            observed: 2,
            maximum: 1,
        })
    ));
    Ok(())
}

#[test]
fn parse_does_not_reserve_object_slots_from_payload_size() -> Result<(), Error> {
    let archive = Archive {
        objects: vec![ArchiveObject::new(
            17,
            vec![RawMessage {
                type_: 300,
                data: vec![0x5a; 64 * 1024],
            }],
        )?],
    };
    let encoded = archive.to_bytes()?;
    let parsed = Archive::parse(&encoded)?;

    assert_eq!(parsed.objects.len(), 1);
    assert!(
        parsed.objects.capacity() <= 4,
        "object capacity grew from payload bytes: {}",
        parsed.objects.capacity()
    );
    Ok(())
}

fn decode_test_varint(data: &[u8]) -> (usize, usize) {
    let mut value = 0usize;
    let mut shift = 0usize;
    for (index, byte) in data.iter().copied().enumerate() {
        value |= usize::from(byte & 0x7f) << shift;
        if byte & 0x80 == 0 {
            return (value, index + 1);
        }
        shift += 7;
    }
    (0, 0)
}

fn one_byte_length(length: usize) -> u8 {
    match u8::try_from(length) {
        Ok(byte_length) => byte_length,
        Err(error) => panic!("test fixture length must fit one byte: {error}"),
    }
}
