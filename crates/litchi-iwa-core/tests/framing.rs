use litchi_iwa_core::{ArchiveLimits, Error, LimitKind, SnappyLimits, SnappyStream};

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
