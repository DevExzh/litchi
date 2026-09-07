#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "The source-member fixtures are fixed and assertion-driven."
)]

use litchi_core::{Error, OwnedSource, ReadAt, SourceVersion};
use litchi_odf_common::core::{
    PackageWriter, Profile, SourceBackedPackage, SourceMemberReaderError,
};
use soapberry_zip::office::StreamingArchiveWriter;
use std::io;
use std::sync::{
    Arc,
    atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering},
};

const MIME: &[u8] = b"application/vnd.oasis.opendocument.text";
const MANIFEST: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3">
  <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
  <manifest:file-entry manifest:full-path="mimetype" manifest:media-type="text/plain"/>
  <manifest:file-entry manifest:full-path="Pictures/target.bin" manifest:media-type="application/octet-stream"/>
</manifest:manifest>"#;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct CallbackFailure;

impl std::fmt::Display for CallbackFailure {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("callback stopped")
    }
}

impl std::error::Error for CallbackFailure {}

#[derive(Debug)]
struct CountingSource {
    bytes: Arc<Vec<u8>>,
    bytes_read: AtomicUsize,
    revision: AtomicU64,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: Arc::new(bytes),
            bytes_read: AtomicUsize::new(0),
            revision: AtomicU64::new(0),
        })
    }

    fn bytes_read(&self) -> usize {
        self.bytes_read.load(Ordering::Relaxed)
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("fixture length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "fixture offset overflow"))?;
        let Some(input) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let amount = input.len().min(output.len());
        output[..amount].copy_from_slice(&input[..amount]);
        self.bytes_read.fetch_add(amount, Ordering::Relaxed);
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x534f_5552_4345_0001,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Arc<Vec<u8>>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: Arc::new(bytes),
            revision: AtomicU64::new(0),
        })
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::Relaxed);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("fixture length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "fixture offset overflow"))?;
        let Some(input) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let amount = input.len().min(output.len());
        output[..amount].copy_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x534f_5552_4345_0002,
            self.revision.load(Ordering::Relaxed),
        ))
    }
}

#[derive(Debug)]
struct ZeroAfterArmSource {
    bytes: Arc<Vec<u8>>,
    payload_start: usize,
    armed: AtomicBool,
}

impl ZeroAfterArmSource {
    fn new(bytes: Vec<u8>, payload_start: usize) -> Arc<Self> {
        Arc::new(Self {
            bytes: Arc::new(bytes),
            payload_start,
            armed: AtomicBool::new(false),
        })
    }

    fn arm(&self) {
        self.armed.store(true, Ordering::Relaxed);
    }
}

impl ReadAt for ZeroAfterArmSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("fixture length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "fixture offset overflow"))?;
        if self.armed.load(Ordering::Relaxed) && start >= self.payload_start {
            return Ok(0);
        }
        let Some(input) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let amount = input
            .len()
            .min(output.len())
            .min(if start >= self.payload_start {
                1
            } else {
                usize::MAX
            });
        output[..amount].copy_from_slice(&input[..amount]);
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x534f_5552_4345_0003, 0))
    }
}

fn package_bytes(deflated: bool, payload: &[u8]) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer.write_stored("mimetype", MIME).unwrap();
    if deflated {
        writer
            .write_deflated_sized("Pictures/target.bin", payload)
            .unwrap();
    } else {
        writer.write_stored("Pictures/target.bin", payload).unwrap();
    }
    writer
        .write_deflated_sized("META-INF/manifest.xml", MANIFEST)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn open(bytes: Vec<u8>) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap()
}

fn open_counted(bytes: Vec<u8>) -> (Arc<CountingSource>, SourceBackedPackage) {
    let source = CountingSource::new(bytes);
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    (source, package)
}

fn open_versioned(bytes: Vec<u8>) -> (Arc<VersionedSource>, SourceBackedPackage) {
    let source = VersionedSource::new(bytes);
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    (source, package)
}

fn encrypted_package_bytes() -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer
        .set_mimetype(std::str::from_utf8(MIME).unwrap())
        .unwrap();
    writer
        .set_encryption("test-password", Profile::compatible())
        .unwrap();
    writer
        .add_file("content.xml", b"<office:document-content/>")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn patch_crc(mut bytes: Vec<u8>, member: &[u8]) -> Vec<u8> {
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let local = cursor + relative;
        let name_len = usize::from(u16::from_le_bytes(
            bytes[local + 26..local + 28].try_into().unwrap(),
        ));
        let name_start = local + 30;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[local + 14..local + 18].copy_from_slice(&0_u32.to_le_bytes());
            break;
        }
        cursor = name_start + name_len;
    }

    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let central = cursor + relative;
        let name_len = usize::from(u16::from_le_bytes(
            bytes[central + 28..central + 30].try_into().unwrap(),
        ));
        let name_start = central + 46;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[central + 16..central + 20].copy_from_slice(&0_u32.to_le_bytes());
            return bytes;
        }
        cursor = name_start + name_len;
    }
    panic!("target member CRC was not found");
}

fn patch_compression_method(mut bytes: Vec<u8>, member: &[u8], method: u16) -> Vec<u8> {
    let mut local_found = false;
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let local = cursor + relative;
        let name_len = usize::from(u16::from_le_bytes(
            bytes[local + 26..local + 28].try_into().unwrap(),
        ));
        let name_start = local + 30;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[local + 8..local + 10].copy_from_slice(&method.to_le_bytes());
            local_found = true;
            break;
        }
        cursor = name_start + name_len;
    }

    let mut central_found = false;
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let central = cursor + relative;
        let name_len = usize::from(u16::from_le_bytes(
            bytes[central + 28..central + 30].try_into().unwrap(),
        ));
        let name_start = central + 46;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[central + 10..central + 12].copy_from_slice(&method.to_le_bytes());
            central_found = true;
            break;
        }
        cursor = name_start + name_len;
    }
    assert!(local_found);
    assert!(central_found);
    bytes
}

fn patch_uncompressed_size(mut bytes: Vec<u8>, member: &[u8], size: u32) -> Vec<u8> {
    let mut local_found = false;
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let local = cursor + relative;
        let name_len = usize::from(u16::from_le_bytes(
            bytes[local + 26..local + 28].try_into().unwrap(),
        ));
        let name_start = local + 30;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[local + 22..local + 26].copy_from_slice(&size.to_le_bytes());
            local_found = true;
            break;
        }
        cursor = name_start + name_len;
    }

    let mut central_found = false;
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x01\x02")
    {
        let central = cursor + relative;
        let name_len = usize::from(u16::from_le_bytes(
            bytes[central + 28..central + 30].try_into().unwrap(),
        ));
        let name_start = central + 46;
        if &bytes[name_start..name_start + name_len] == member {
            bytes[central + 24..central + 28].copy_from_slice(&size.to_le_bytes());
            central_found = true;
            break;
        }
        cursor = name_start + name_len;
    }
    assert!(local_found);
    assert!(central_found);
    bytes
}

fn local_payload_offset(bytes: &[u8], member: &[u8]) -> usize {
    let mut cursor = 0;
    while let Some(relative) = bytes[cursor..]
        .windows(4)
        .position(|window| window == b"PK\x03\x04")
    {
        let local = cursor + relative;
        let name_len = usize::from(u16::from_le_bytes(
            bytes[local + 26..local + 28].try_into().unwrap(),
        ));
        let extra_len = usize::from(u16::from_le_bytes(
            bytes[local + 28..local + 30].try_into().unwrap(),
        ));
        let name_start = local + 30;
        if &bytes[name_start..name_start + name_len] == member {
            return name_start + name_len + extra_len;
        }
        cursor = name_start + name_len + extra_len;
    }
    panic!("target member local payload was not found");
}

#[test]
fn verified_reader_streams_store_and_deflate_full_or_early_without_materializing() {
    let payload = vec![b'x'; 64 * 1024 + 37];
    for deflated in [false, true] {
        let package = open(package_bytes(deflated, &payload));
        let full = package
            .with_verified_member_reader("Pictures/target.bin", |reader| {
                let mut output = Vec::new();
                reader.read_to_end(&mut output)?;
                Ok::<_, io::Error>(output)
            })
            .unwrap();
        assert_eq!(full, payload);

        let early = package
            .with_verified_member_reader("/Pictures/target.bin", |reader| {
                let mut prefix = [0_u8; 11];
                reader.read_exact(&mut prefix)?;
                Ok::<_, io::Error>(prefix.to_vec())
            })
            .unwrap();
        assert_eq!(early, payload[..11]);
    }
}

#[test]
fn callback_error_and_bad_crc_keep_archive_error_primary() {
    let payload = vec![b'c'; 32 * 1024 + 5];
    let bytes = patch_crc(package_bytes(true, &payload), b"Pictures/target.bin");
    let package = open(bytes);
    let error = package
        .with_verified_member_reader("Pictures/target.bin", |_reader| {
            Err::<(), _>(CallbackFailure)
        })
        .unwrap_err();

    assert!(matches!(error, SourceMemberReaderError::Core { .. }));
    assert!(matches!(error.callback(), Some(&CallbackFailure)));
    assert!(
        matches!(error.core(), Some(Error::InvalidFormat(reason)) if reason.contains("checksum"))
    );
}

#[test]
fn unsupported_compression_refuses_before_callback() {
    let payload = vec![b'u'; 1024];
    let bytes = patch_compression_method(package_bytes(true, &payload), b"Pictures/target.bin", 12);
    let package = open(bytes);
    let called = Arc::new(AtomicBool::new(false));
    let callback_called = Arc::clone(&called);
    let error = package
        .with_verified_member_reader("Pictures/target.bin", move |_| {
            callback_called.store(true, Ordering::Relaxed);
            Ok::<_, CallbackFailure>(())
        })
        .unwrap_err();

    assert!(!called.load(Ordering::Relaxed));
    assert!(
        matches!(error.core(), Some(Error::InvalidFormat(reason)) if reason.contains("compression"))
    );
}

#[test]
fn normal_reader_drains_bad_size_but_abortable_reader_stops_before_it() {
    let payload = vec![b'z'; 4096];
    let bytes = patch_uncompressed_size(
        package_bytes(true, &payload),
        b"Pictures/target.bin",
        u32::try_from(payload.len() + 1).unwrap(),
    );

    let normal = open(bytes.clone());
    let error = normal
        .with_verified_member_reader("Pictures/target.bin", |_reader| {
            Err::<(), _>(CallbackFailure)
        })
        .unwrap_err();
    assert!(matches!(error.core(), Some(Error::InvalidFormat(reason)) if reason.contains("size")));
    assert!(matches!(error.callback(), Some(&CallbackFailure)));

    let abortable = open(bytes);
    let error = abortable
        .with_verified_member_reader_abortable("Pictures/target.bin", |_reader| {
            Err::<(), _>(CallbackFailure)
        })
        .unwrap_err();
    assert!(matches!(error, SourceMemberReaderError::Callback(_)));
}

#[test]
fn abortable_callback_failure_does_not_drain_the_member() {
    let payload = vec![b'a'; 256 * 1024 + 1];
    let (source, package) = open_counted(package_bytes(false, &payload));
    let before = source.bytes_read();
    let error = package
        .with_verified_member_reader_abortable("Pictures/target.bin", |reader| {
            let mut prefix = [0_u8; 1];
            reader
                .read_exact(&mut prefix)
                .expect("fixture member prefix must be readable");
            Err::<(), _>(CallbackFailure)
        })
        .unwrap_err();
    let abort_read = source.bytes_read() - before;
    assert!(matches!(error, SourceMemberReaderError::Callback(_)));

    let before = source.bytes_read();
    let error = package
        .with_verified_member_reader("Pictures/target.bin", |reader| {
            let mut prefix = [0_u8; 1];
            reader
                .read_exact(&mut prefix)
                .expect("fixture member prefix must be readable");
            Err::<(), _>(CallbackFailure)
        })
        .unwrap_err();
    let normal_read = source.bytes_read() - before;
    assert!(matches!(error, SourceMemberReaderError::Callback(_)));
    assert!(normal_read > abort_read);
}

#[test]
fn early_source_zero_is_observed_only_by_the_draining_reader() {
    let payload = vec![b'e'; 4096];
    let bytes = package_bytes(false, &payload);
    let payload_start = local_payload_offset(&bytes, b"Pictures/target.bin");
    let source = ZeroAfterArmSource::new(bytes.clone(), payload_start);
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let error = package
        .with_verified_member_reader("Pictures/target.bin", |reader| {
            let mut prefix = [0_u8; 1];
            reader
                .read_exact(&mut prefix)
                .expect("fixture member prefix must be readable");
            source.arm();
            Err::<(), _>(CallbackFailure)
        })
        .unwrap_err();
    assert!(matches!(error, SourceMemberReaderError::Core { .. }));
    assert!(matches!(error.callback(), Some(&CallbackFailure)));
    assert!(matches!(error.core(), Some(Error::InvalidFormat(reason)) if reason.contains("size")));

    let source = ZeroAfterArmSource::new(bytes, payload_start);
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let error = package
        .with_verified_member_reader_abortable("Pictures/target.bin", |reader| {
            let mut prefix = [0_u8; 1];
            reader
                .read_exact(&mut prefix)
                .expect("fixture member prefix must be readable");
            source.arm();
            Err::<(), _>(CallbackFailure)
        })
        .unwrap_err();
    assert!(matches!(error, SourceMemberReaderError::Callback(_)));
}

#[test]
fn final_source_change_is_primary_and_retains_abortable_callback_error() {
    let payload = vec![b's'; 8 * 1024];
    let (source, package) = open_versioned(package_bytes(false, &payload));
    let error = package
        .with_verified_member_reader_abortable("Pictures/target.bin", |reader| {
            let mut prefix = [0_u8; 1];
            reader
                .read_exact(&mut prefix)
                .expect("fixture member prefix must be readable");
            source.bump();
            Err::<(), _>(CallbackFailure)
        })
        .unwrap_err();

    assert!(matches!(error.core(), Some(Error::SourceChanged { .. })));
    assert!(matches!(error.callback(), Some(&CallbackFailure)));
}

#[test]
fn encrypted_target_refuses_before_callback() {
    let source = Arc::new(OwnedSource::new(encrypted_package_bytes()));
    let package = SourceBackedPackage::from_read_at_with_password(source, "test-password").unwrap();
    let called = Arc::new(AtomicBool::new(false));
    let callback_called = Arc::clone(&called);
    let error = package
        .with_verified_member_reader("content.xml", move |_| {
            callback_called.store(true, Ordering::Relaxed);
            Ok::<_, CallbackFailure>(())
        })
        .unwrap_err();

    assert!(!called.load(Ordering::Relaxed));
    assert!(
        matches!(error.core(), Some(Error::InvalidFormat(reason)) if reason.contains("Encrypted"))
    );
}

#[test]
fn repeated_reads_traverse_source_again_without_a_payload_cache() {
    let payload = vec![b'r'; 48 * 1024];
    let (source, package) = open_counted(package_bytes(true, &payload));
    let before = source.bytes_read();
    package
        .with_verified_member_reader("Pictures/target.bin", |reader| {
            let mut output = Vec::new();
            reader.read_to_end(&mut output)?;
            Ok::<_, io::Error>(output.len())
        })
        .unwrap();
    let after_first = source.bytes_read();
    package
        .with_verified_member_reader("Pictures/target.bin", |reader| {
            let mut output = Vec::new();
            reader.read_to_end(&mut output)?;
            Ok::<_, io::Error>(output.len())
        })
        .unwrap();
    let after_second = source.bytes_read();

    assert!(after_first > before);
    assert!(after_second > after_first);
}
