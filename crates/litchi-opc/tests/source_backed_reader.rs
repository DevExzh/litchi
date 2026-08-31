#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "focused sequential-reader assertions intentionally fail on fixture errors"
)]

//! Sequential-reader ingress tests for the lazy source-backed OPC package.

use std::io::{self, Read};

use litchi_opc::{OpcError, PackURI, ReadLimits, ReadResource, SourceBackedPackage};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const DOCUMENT_URI: &str = "/word/document.xml";
const UNUSED_MEMBER: &str = "custom/unused.bin";
const UNUSED_URI: &str = "/custom/unused.bin";
const DOCUMENT_PAYLOAD: &[u8] = b"selected reader payload";
const UNUSED_PAYLOAD: &[u8] = b"ordinary payload remains deferred";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

fn archive_bytes() -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="{DOCUMENT_URI}" ContentType="{DOCUMENT_CONTENT_TYPE}"/></Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="{DOCUMENT_MEMBER}"/></Relationships>"#
    );

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    writer
        .write_deflated_sized(DOCUMENT_MEMBER, DOCUMENT_PAYLOAD)
        .unwrap();
    writer
        .write_deflated_sized(UNUSED_MEMBER, UNUSED_PAYLOAD)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

struct SinglePassReader {
    bytes: Vec<u8>,
    offset: usize,
    bytes_read: usize,
    eof_returned: bool,
}

impl SinglePassReader {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            offset: 0,
            bytes_read: 0,
            eof_returned: false,
        }
    }
}

impl Read for SinglePassReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.eof_returned {
            return Err(io::Error::other("sequential input was read after EOF"));
        }
        if self.offset == self.bytes.len() {
            self.eof_returned = true;
            return Ok(0);
        }

        let count = output.len().min(self.bytes.len() - self.offset);
        output[..count].copy_from_slice(&self.bytes[self.offset..self.offset + count]);
        self.offset += count;
        self.bytes_read += count;
        Ok(count)
    }
}

struct InterruptOnceReader {
    inner: SinglePassReader,
    interrupted: bool,
}

impl InterruptOnceReader {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            inner: SinglePassReader::new(bytes),
            interrupted: false,
        }
    }
}

impl Read for InterruptOnceReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(io::Error::from(io::ErrorKind::Interrupted));
        }
        self.inner.read(output)
    }
}

struct InvalidCountReader;

impl Read for InvalidCountReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        Ok(output
            .len()
            .checked_add(1)
            .expect("test buffer length must fit"))
    }
}

#[test]
fn reader_ingress_consumes_once_and_keeps_ordinary_payloads_cold() {
    let source = archive_bytes();
    let input_length = source.len();
    let limits = ReadLimits::builder()
        .max_input_bytes(input_length as u64)
        .unwrap()
        .build()
        .unwrap();
    let mut reader = SinglePassReader::new(source);

    let package = SourceBackedPackage::from_reader_with_limits(&mut reader, limits).unwrap();

    assert_eq!(reader.bytes_read, input_length);
    assert_eq!(reader.offset, input_length);
    assert!(reader.eof_returned);
    assert_eq!(package.iter_parts().count(), 2);
    assert!(package.part(&pack(UNUSED_URI)).is_ok());
    assert_eq!(package.cache_diagnostics().cold_loads, 0);
    assert_eq!(package.cache_diagnostics().successful_loads, 0);

    let selected = package.part(&pack(DOCUMENT_URI)).unwrap().data().unwrap();
    assert_eq!(selected.as_bytes(), DOCUMENT_PAYLOAD);
    assert_eq!(package.cache_diagnostics().cold_loads, 1);
    assert_eq!(package.cache_diagnostics().successful_loads, 1);
}

#[test]
fn reader_ingress_rejects_input_above_max_input_bytes() {
    let source = archive_bytes();
    let configured_max = source.len() - 1;
    let limits = ReadLimits::builder()
        .max_input_bytes(configured_max as u64)
        .unwrap()
        .build()
        .unwrap();
    let mut reader = SinglePassReader::new(source);

    let error = SourceBackedPackage::from_reader_with_limits(&mut reader, limits)
        .err()
        .expect("input larger than max_input_bytes must be rejected");

    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::InputBytes,
            actual,
            maximum,
        }
        if actual == configured_max as u64 + 1 && maximum == configured_max as u64
    ));
    assert_eq!(reader.bytes_read, configured_max + 1);
    assert_eq!(reader.offset, configured_max + 1);
}

#[test]
fn reader_ingress_retries_one_interrupted_read() {
    let source = archive_bytes();
    let input_length = source.len();
    let limits = ReadLimits::builder()
        .max_input_bytes(input_length as u64)
        .unwrap()
        .build()
        .unwrap();
    let mut reader = InterruptOnceReader::new(source);

    let package = SourceBackedPackage::from_reader_with_limits(&mut reader, limits)
        .expect("a transient interrupted read must be retried");

    assert!(reader.interrupted);
    assert_eq!(reader.inner.bytes_read, input_length);
    assert_eq!(reader.inner.offset, input_length);
    assert_eq!(package.iter_parts().count(), 2);
}

#[test]
fn reader_ingress_rejects_invalid_read_count_without_panicking() {
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        SourceBackedPackage::from_reader_with_limits(InvalidCountReader, ReadLimits::default())
    }));
    let result = result.expect("an invalid Read count must not panic public ingress");
    let error = match result {
        Ok(_) => panic!("an invalid Read count must be rejected"),
        Err(error) => error,
    };

    assert!(matches!(
        error,
        OpcError::IoError(error) if error.kind() == io::ErrorKind::InvalidData
    ));
}
