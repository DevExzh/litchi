#![allow(
    clippy::unwrap_used,
    clippy::expect_used,
    reason = "test assertions panic on failure by design"
)]

use super::{Error, OpenOptions, Reader, SourceLimits};
use litchi_cfb::{OleFile, OleWriter};
use std::io::Cursor;
use std::path::Path;

fn serialize_ole(writer: &mut OleWriter) -> Vec<u8> {
    let mut output = Cursor::new(Vec::new());
    writer.write_to(&mut output).unwrap();
    output.into_inner()
}

#[test]
fn test_open_package() {
    let base = Path::new(env!("CARGO_MANIFEST_DIR"))
        .parent()
        .unwrap()
        .parent()
        .unwrap();
    let result = Reader::open(
        base.join("test-data")
            .join("ole")
            .join("ppt")
            .join("empty.ppt"),
    );
    assert!(result.is_ok());
}

#[test]
fn rejects_powerpoint_document_storage_before_package_publication() {
    let mut writer = OleWriter::new();
    writer.create_storage(&["PowerPoint Document"]).unwrap();
    let bytes = serialize_ole(&mut writer);

    let from_reader = Reader::from_reader(Cursor::new(bytes.clone()));
    assert!(matches!(
        from_reader,
        Err(Error::InvalidFormat(message)) if message.contains("is not a stream")
    ));

    let ole = OleFile::open(Cursor::new(bytes), litchi_cfb::OleFileLimits::default()).unwrap();
    let from_ole_file = Reader::from_ole_file(ole);
    assert!(matches!(
        from_ole_file,
        Err(Error::InvalidFormat(message)) if message.contains("is not a stream")
    ));
}

#[test]
fn rejects_oversized_clear_or_encrypted_pictures_before_presentation_parsing() {
    for token in [0xe391_c05fu32, 0xf3d1_c4dfu32] {
        let mut current_user = vec![0u8; 32];
        current_user[2..4].copy_from_slice(&0x0ff6u16.to_le_bytes());
        current_user[4..8].copy_from_slice(&24u32.to_le_bytes());
        current_user[8..12].copy_from_slice(&20u32.to_le_bytes());
        current_user[12..16].copy_from_slice(&token.to_le_bytes());

        let mut writer = OleWriter::new();
        writer.create_stream(&["PowerPoint Document"], &[]).unwrap();
        writer
            .create_stream(&["Current User"], &current_user)
            .unwrap();
        writer.create_stream(&["Pictures"], &[0u8; 65]).unwrap();
        let bytes = serialize_ole(&mut writer);
        let mut package = Reader::from_reader(Cursor::new(bytes)).unwrap();
        let error = package
            .presentation_with_options(OpenOptions::new().source_limits(SourceLimits {
                max_input_bytes: 64,
                ..SourceLimits::default()
            }))
            .err()
            .unwrap();
        assert!(matches!(
            error,
            Error::Limit {
                kind: crate::SourceLimitKind::EntryBytes,
                ..
            }
        ));
    }
}

#[test]
fn rejects_aggregate_stream_bytes_at_the_exact_boundary_plus_one() {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &[0u8; 8])
        .unwrap();
    writer.create_stream(&["Current User"], &[0u8; 8]).unwrap();
    writer.create_stream(&["Pictures"], &[0u8; 8]).unwrap();
    let bytes = serialize_ole(&mut writer);
    let mut package = Reader::from_reader(Cursor::new(bytes)).unwrap();
    let error = package
        .presentation_with_options(OpenOptions::new().source_limits(SourceLimits {
            max_input_bytes: 8,
            max_aggregate_input_bytes: 23,
            ..SourceLimits::default()
        }))
        .err()
        .unwrap();
    assert!(matches!(
        error,
        Error::Limit {
            kind: crate::SourceLimitKind::AggregateBytes,
            ..
        }
    ));
}

#[test]
fn package_byte_limit_accepts_exact_size_and_rejects_one_less() {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &[0u8; 8])
        .unwrap();
    let bytes = serialize_ole(&mut writer);
    let exact = SourceLimits {
        max_package_bytes: bytes.len(),
        ..SourceLimits::default()
    };
    assert!(
        Reader::from_reader_with_options(
            Cursor::new(bytes.clone()),
            OpenOptions::new().source_limits(exact)
        )
        .is_ok()
    );
    let error = Reader::from_reader_with_options(
        Cursor::new(bytes.clone()),
        OpenOptions::new().source_limits(SourceLimits {
            max_package_bytes: bytes.len() - 1,
            ..SourceLimits::default()
        }),
    )
    .err()
    .unwrap();
    assert!(matches!(
        error,
        Error::Limit {
            kind: crate::SourceLimitKind::InputBytes,
            ..
        }
    ));
}

#[test]
#[ignore = "requires a test file"]
fn test_invalid_file() {
    // Create a non-PPT file
    std::fs::write("test_invalid.tmp", b"Not a PPT file").unwrap();
    let result = Reader::open("test_invalid.tmp");
    assert!(result.is_err());
    std::fs::remove_file("test_invalid.tmp").ok();
}

#[test]
fn unspecified_open_options_preserve_profiles_above_defaults() {
    use super::{OpenOptions, SourceBackedPackage};
    use litchi_core::OwnedSource;
    use std::sync::Arc;

    let bytes = include_bytes!("../../../../test-data/ole/ppt/empty.ppt").to_vec();
    let selected = SourceLimits::builder()
        .max_input_bytes(SourceLimits::MAX_INPUT_BYTES)
        .unwrap()
        .max_aggregate_payload_bytes(SourceLimits::MAX_AGGREGATE_PAYLOAD_BYTES)
        .unwrap()
        .build()
        .unwrap();
    let mut eager = Reader::from_reader_with_options(
        Cursor::new(bytes.clone()),
        OpenOptions::new().source_limits(selected),
    )
    .unwrap();
    assert_eq!(eager.presentation().unwrap().source_limits(), selected);
    assert_eq!(
        eager
            .presentation_with_options(OpenOptions::new())
            .unwrap()
            .source_limits(),
        selected
    );
    assert_eq!(
        eager
            .presentation_with_options(OpenOptions::new().source_limits(SourceLimits::default()))
            .unwrap()
            .source_limits(),
        SourceLimits::default()
    );

    let source = SourceBackedPackage::from_read_at_with_options(
        Arc::new(OwnedSource::new(bytes)),
        OpenOptions::new().source_limits(selected),
    )
    .unwrap();
    assert_eq!(source.presentation().unwrap().source_limits(), selected);
    assert_eq!(
        source
            .presentation_with_options(OpenOptions::new().source_limits(SourceLimits::default()))
            .unwrap()
            .source_limits(),
        SourceLimits::default()
    );
}

#[test]
fn core_limit_conversion_preserves_dimension_and_values() {
    use super::SourceLimitKind;
    use litchi_core::Resource;
    for (kind, resource) in [
        (SourceLimitKind::Events, Resource::Objects),
        (SourceLimitKind::Depth, Resource::Depth),
        (SourceLimitKind::EntryBytes, Resource::InputBytes),
    ] {
        let error: litchi_core::Error = Error::Limit {
            kind,
            observed: 5,
            maximum: 4,
        }
        .into();
        assert!(matches!(error, litchi_core::Error::ResourceLimit(limit)
            if limit.resource == resource && limit.observed == 5 && limit.limit == 4));
    }
}
