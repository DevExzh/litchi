//! Integration tests for OLE writer
//!
//! These tests verify that the OLE writer can create valid OLE2 files
//! that can be read back by the OLE reader.

use super::super::file::OleFile;
use super::core::OleWriter;
use std::io::Cursor;

#[test]
fn test_write_simple_ole_file() {
    // Create a simple OLE file with one stream
    let mut writer = OleWriter::new();
    writer
        .create_stream(&["TestStream"], b"Hello, World!")
        .unwrap();

    // Write to memory buffer
    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    // Read it back
    let data = buffer.into_inner();
    assert!(data.len() >= 1536); // Minimum OLE file size

    // Verify magic bytes
    assert_eq!(&data[0..8], b"\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1");

    // Try to open with reader
    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    // Verify we can read the stream back
    let stream_data = ole.open_stream(&["TestStream"]).unwrap();
    assert_eq!(stream_data, b"Hello, World!");
}

#[test]
fn test_write_multiple_streams() {
    let mut writer = OleWriter::new();

    // Add multiple streams of different sizes
    writer.create_stream(&["Small1"], b"Small").unwrap();
    writer.create_stream(&["Small2"], b"Data").unwrap();
    writer
        .create_stream(&["Large1"], &vec![0xAAu8; 5000])
        .unwrap();
    writer
        .create_stream(&["Large2"], &vec![0xBBu8; 10000])
        .unwrap();

    // Write to memory buffer
    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    // Read it back
    let data = buffer.into_inner();
    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    // Verify all streams
    let small1 = ole.open_stream(&["Small1"]).unwrap();
    assert_eq!(small1, b"Small");

    let small2 = ole.open_stream(&["Small2"]).unwrap();
    assert_eq!(small2, b"Data");

    let large1 = ole.open_stream(&["Large1"]).unwrap();
    assert_eq!(large1.len(), 5000);
    assert!(large1.iter().all(|&b| b == 0xAA));

    let large2 = ole.open_stream(&["Large2"]).unwrap();
    assert_eq!(large2.len(), 10000);
    assert!(large2.iter().all(|&b| b == 0xBB));
}

#[test]
fn test_write_empty_stream() {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Empty"], b"").unwrap();

    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    let data = buffer.into_inner();
    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    let empty = ole.open_stream(&["Empty"]).unwrap();
    assert_eq!(empty.len(), 0);
}

#[test]
fn test_write_with_minifat() {
    // Create streams smaller than 4096 bytes (should use MiniFAT)
    let mut writer = OleWriter::new();

    for i in 0..10 {
        let name = format!("Stream{}", i);
        let data = vec![i as u8; 100 + i * 50]; // Different sizes < 4096
        writer.create_stream(&[&name], &data).unwrap();
    }

    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    let data = buffer.into_inner();
    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    // Verify all streams
    for i in 0..10 {
        let name = format!("Stream{}", i);
        let stream_data = ole.open_stream(&[&name]).unwrap();
        assert_eq!(stream_data.len(), 100 + i * 50);
        assert!(stream_data.iter().all(|&b| b == i as u8));
    }
}

#[test]
fn test_write_large_stream() {
    // Test with a large stream (> 64KB to test multiple sectors)
    let mut writer = OleWriter::new();
    let large_data = vec![0x42u8; 100_000];
    writer.create_stream(&["LargeStream"], &large_data).unwrap();

    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    let data = buffer.into_inner();
    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    let read_data = ole.open_stream(&["LargeStream"]).unwrap();
    assert_eq!(read_data.len(), 100_000);
    assert!(read_data.iter().all(|&b| b == 0x42));
}

#[test]
fn test_write_update_delete() {
    let mut writer = OleWriter::new();

    // Create initial stream
    writer.create_stream(&["Test"], b"Initial").unwrap();

    // Update it
    writer.update_stream(&["Test"], b"Updated").unwrap();

    // Create another stream
    writer.create_stream(&["Test2"], b"Data").unwrap();

    // Delete the first stream
    writer.delete_stream(&["Test"]).unwrap();

    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    let data = buffer.into_inner();
    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    // Test should not exist
    assert!(ole.open_stream(&["Test"]).is_err());

    // Test2 should exist
    let test2_data = ole.open_stream(&["Test2"]).unwrap();
    assert_eq!(test2_data, b"Data");
}

#[test]
fn test_write_sector_size_4096() {
    let mut writer = OleWriter::with_sector_size(4096);
    writer.create_stream(&["Test"], b"Hello, 4096!").unwrap();

    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    let data = buffer.into_inner();

    // Verify sector size in header (offset 0x1E, should be 12 for 4096 bytes)
    let sector_shift = u16::from_le_bytes([data[0x1E], data[0x1F]]);
    assert_eq!(sector_shift, 12);

    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    let stream_data = ole.open_stream(&["Test"]).unwrap();
    assert_eq!(stream_data, b"Hello, 4096!");
}

#[test]
fn test_list_streams_after_write() {
    let mut writer = OleWriter::new();
    writer.create_stream(&["Stream1"], b"Data1").unwrap();
    writer.create_stream(&["Stream2"], b"Data2").unwrap();
    writer.create_stream(&["Stream3"], b"Data3").unwrap();

    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    let data = buffer.into_inner();
    let cursor = Cursor::new(data);
    let ole = OleFile::open(cursor).unwrap();

    let streams = ole.list_streams();
    assert_eq!(streams.len(), 3);

    // Check that all streams are present (order may vary)
    let stream_names: Vec<&str> = streams.iter().map(|s| s[0].as_str()).collect();
    assert!(stream_names.contains(&"Stream1"));
    assert!(stream_names.contains(&"Stream2"));
    assert!(stream_names.contains(&"Stream3"));
}

#[test]
fn test_write_mixed_sizes() {
    // Test with a mix of small (MiniFAT) and large (FAT) streams
    let mut writer = OleWriter::new();

    // Small streams (< 4096 bytes, should use MiniFAT)
    writer.create_stream(&["Tiny"], b"tiny").unwrap();
    writer
        .create_stream(&["Small"], &vec![0x11u8; 1000])
        .unwrap();
    writer
        .create_stream(&["Medium"], &vec![0x22u8; 3000])
        .unwrap();

    // Large streams (>= 4096 bytes, should use FAT)
    writer
        .create_stream(&["Large"], &vec![0x33u8; 5000])
        .unwrap();
    writer
        .create_stream(&["Huge"], &vec![0x44u8; 20000])
        .unwrap();

    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    let data = buffer.into_inner();
    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    // Verify all streams
    assert_eq!(ole.open_stream(&["Tiny"]).unwrap(), b"tiny");

    let small = ole.open_stream(&["Small"]).unwrap();
    assert_eq!(small.len(), 1000);
    assert!(small.iter().all(|&b| b == 0x11));

    let medium = ole.open_stream(&["Medium"]).unwrap();
    assert_eq!(medium.len(), 3000);
    assert!(medium.iter().all(|&b| b == 0x22));

    let large = ole.open_stream(&["Large"]).unwrap();
    assert_eq!(large.len(), 5000);
    assert!(large.iter().all(|&b| b == 0x33));

    let huge = ole.open_stream(&["Huge"]).unwrap();
    assert_eq!(huge.len(), 20000);
    assert!(huge.iter().all(|&b| b == 0x44));
}

#[test]
#[should_panic(expected = "Sector size must be 512 or 4096")]
fn test_invalid_sector_size() {
    let _ = OleWriter::with_sector_size(1024);
}

#[test]
fn test_write_to_file() {
    use std::env;
    use std::fs;

    let mut writer = OleWriter::new();
    writer
        .create_stream(&["TestFile"], b"File content")
        .unwrap();

    // Create a temporary file path
    let temp_dir = env::temp_dir();
    let temp_path = temp_dir.join("litchi_test_ole_file.ole");

    // Write to file
    writer.save(&temp_path).unwrap();

    // Read back from file
    let file = fs::File::open(&temp_path).unwrap();
    let mut ole = OleFile::open(file).unwrap();

    let data = ole.open_stream(&["TestFile"]).unwrap();
    assert_eq!(data, b"File content");

    // Clean up
    let _ = fs::remove_file(&temp_path);
}

#[test]
fn test_boundary_conditions() {
    let mut writer = OleWriter::new();

    // Test exactly at MiniFAT cutoff boundary (4095 and 4096 bytes)
    writer
        .create_stream(&["JustUnder"], &vec![0xAAu8; 4095])
        .unwrap();
    writer
        .create_stream(&["Exactly"], &vec![0xBBu8; 4096])
        .unwrap();
    writer
        .create_stream(&["JustOver"], &vec![0xCCu8; 4097])
        .unwrap();

    let mut buffer = Cursor::new(Vec::new());
    writer.write_to(&mut buffer).unwrap();

    let data = buffer.into_inner();
    let cursor = Cursor::new(data);
    let mut ole = OleFile::open(cursor).unwrap();

    let under = ole.open_stream(&["JustUnder"]).unwrap();
    assert_eq!(under.len(), 4095);

    let exactly = ole.open_stream(&["Exactly"]).unwrap();
    assert_eq!(exactly.len(), 4096);

    let over = ole.open_stream(&["JustOver"]).unwrap();
    assert_eq!(over.len(), 4097);
}
