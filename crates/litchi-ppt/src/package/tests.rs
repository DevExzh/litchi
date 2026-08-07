use super::{Error, Package, RecordLimits};
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
    let result = Package::open(
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

    let from_reader = Package::from_reader(Cursor::new(bytes.clone()));
    assert!(matches!(
        from_reader,
        Err(Error::InvalidFormat(message)) if message.contains("is not a stream")
    ));

    let ole = OleFile::open(Cursor::new(bytes)).unwrap();
    let from_ole_file = Package::from_ole_file(ole);
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
        let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
        let error = package
            .presentation_with_limits(RecordLimits {
                max_input_bytes: 64,
                ..RecordLimits::default()
            })
            .err()
            .unwrap();
        assert!(matches!(error, Error::ResourceLimit(message) if message.contains("Pictures")));
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
    let mut package = Package::from_reader(Cursor::new(bytes)).unwrap();
    let error = package
        .presentation_with_limits(RecordLimits {
            max_input_bytes: 8,
            max_aggregate_input_bytes: 23,
            ..RecordLimits::default()
        })
        .err()
        .unwrap();
    assert!(matches!(error, Error::ResourceLimit(message) if message.contains("aggregate")));
}

#[test]
fn package_byte_limit_accepts_exact_size_and_rejects_one_less() {
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["PowerPoint Document"], &[0u8; 8])
        .unwrap();
    let bytes = serialize_ole(&mut writer);
    let exact = RecordLimits {
        max_package_bytes: bytes.len(),
        ..RecordLimits::default()
    };
    assert!(Package::from_reader_with_limits(Cursor::new(bytes.clone()), exact).is_ok());
    let error = Package::from_reader_with_limits(
        Cursor::new(bytes.clone()),
        RecordLimits {
            max_package_bytes: bytes.len() - 1,
            ..RecordLimits::default()
        },
    )
    .err()
    .unwrap();
    assert!(matches!(error, Error::ResourceLimit(message) if message.contains("package size")));
}

#[test]
#[ignore] // Requires test file
fn test_invalid_file() {
    // Create a non-PPT file
    std::fs::write("test_invalid.tmp", b"Not a PPT file").unwrap();
    let result = Package::open("test_invalid.tmp");
    assert!(result.is_err());
    std::fs::remove_file("test_invalid.tmp").ok();
}
