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
