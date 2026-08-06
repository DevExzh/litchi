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
    Archive, ArchiveLimits, ArchiveObject, Error, LimitKind, RawMessage, SnappyLimits, SnappyStream,
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
        Err(Error::Protobuf(_))
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
    let source = ArchiveObject::new(
        51,
        vec![RawMessage {
            type_: 61,
            data: vec![1, 2],
        }],
    )?;
    let proto = litchi_iwa_protos::tsp::MessageInfo::from(&source.archive_info.message_infos[0]);
    let encoded = proto.encode_to_vec();
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
