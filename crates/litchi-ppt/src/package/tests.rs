use super::{Error, Package};
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
#[ignore] // Requires test file
fn test_invalid_file() {
    // Create a non-PPT file
    std::fs::write("test_invalid.tmp", b"Not a PPT file").unwrap();
    let result = Package::open("test_invalid.tmp");
    assert!(result.is_err());
    std::fs::remove_file("test_invalid.tmp").ok();
}
