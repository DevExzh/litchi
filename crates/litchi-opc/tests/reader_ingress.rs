#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "the focused ingress tests use fixture construction assertions"
)]

//! Cross-check the sequential reader boundary shared by the eager, physical,
//! and source-backed OPC constructors.

use std::io::{self, Read};

use litchi_opc::SourceCacheLimits;
use litchi_opc::phys_pkg::OwnedPhysPkgReader;
use litchi_opc::{
    OpcError, OpcPackage, PackURI, SourceBackedPackage, SourceLimitKind, SourceLimits,
};
use soapberry_zip::office::StreamingArchiveWriter;
use tempfile::NamedTempFile;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const DOCUMENT_URI: &str = "/word/document.xml";
const DOCUMENT_PAYLOAD: &[u8] = b"reader ingress parity payload";

fn limits(maximum: usize) -> SourceLimits {
    SourceLimits::builder()
        .max_input_bytes(maximum as u64)
        .expect("positive test input limit")
        .build()
        .expect("consistent test input limits")
}

fn archive_bytes() -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="bin" ContentType="application/octet-stream"/><Override PartName="{DOCUMENT_URI}" ContentType="{DOCUMENT_CONTENT_TYPE}"/></Types>"#
    );
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    );

    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("write content types");
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .expect("write package relationships");
    writer
        .write_stored("word/document.xml", DOCUMENT_PAYLOAD)
        .expect("write document");
    writer
        .write_stored("custom/unused.bin", b"deferred payload")
        .expect("write unused payload");
    writer.finish_to_bytes().expect("finish archive")
}

struct ChunkedReader {
    bytes: Vec<u8>,
    offset: usize,
    chunk_size: usize,
    interruptions: usize,
}

impl ChunkedReader {
    fn new(bytes: Vec<u8>, chunk_size: usize, interruptions: usize) -> Self {
        assert!(chunk_size > 0);
        Self {
            bytes,
            offset: 0,
            chunk_size,
            interruptions,
        }
    }
}

impl Read for ChunkedReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.interruptions != 0 {
            self.interruptions -= 1;
            return Err(io::Error::from(io::ErrorKind::Interrupted));
        }
        if self.offset == self.bytes.len() {
            return Ok(0);
        }
        let count = output
            .len()
            .min(self.chunk_size)
            .min(self.bytes.len() - self.offset);
        output[..count].copy_from_slice(&self.bytes[self.offset..self.offset + count]);
        self.offset += count;
        Ok(count)
    }
}

struct BrokenReader {
    kind: io::ErrorKind,
}

impl Read for BrokenReader {
    fn read(&mut self, _output: &mut [u8]) -> io::Result<usize> {
        Err(io::Error::from(self.kind))
    }
}

struct ErrorAfterLimit {
    bytes: Vec<u8>,
    offset: usize,
    fail_at: usize,
}

impl Read for ErrorAfterLimit {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.offset >= self.fail_at {
            return Err(io::Error::from(io::ErrorKind::BrokenPipe));
        }
        let count = output
            .len()
            .min(self.fail_at - self.offset)
            .min(self.bytes.len() - self.offset);
        output[..count].copy_from_slice(&self.bytes[self.offset..self.offset + count]);
        self.offset += count;
        Ok(count)
    }
}

fn input_limit_error(error: OpcError, maximum: usize, actual: usize) {
    assert!(matches!(
        error,
        OpcError::Limit {
            kind: SourceLimitKind::InputBytes,
            observed,
            maximum: configured,
        } if observed == actual as u64 && configured == maximum as u64
    ));
}

fn take_error<T>(result: litchi_opc::Result<T>, message: &str) -> OpcError {
    match result {
        Ok(_) => panic!("{message}"),
        Err(error) => error,
    }
}

fn document_from_physical(source: Vec<u8>, policy: SourceLimits) -> Vec<u8> {
    let reader = OwnedPhysPkgReader::from_reader(ChunkedReader::new(source, 3, 2), policy)
        .expect("physical reader ingress");
    reader
        .blob_for(&PackURI::new(DOCUMENT_URI).expect("document URI"))
        .expect("physical document payload")
}

fn document_from_package(source: Vec<u8>, policy: SourceLimits) -> Vec<u8> {
    let package = OpcPackage::from_reader(ChunkedReader::new(source, 3, 2), policy)
        .expect("eager package reader ingress");
    package
        .get_part(&PackURI::new(DOCUMENT_URI).expect("document URI"))
        .expect("eager document part")
        .blob()
        .to_vec()
}

fn document_from_source_backed(source: Vec<u8>, policy: SourceLimits) -> Vec<u8> {
    let package = SourceBackedPackage::from_reader(
        ChunkedReader::new(source, 3, 2),
        policy,
        SourceCacheLimits::default(),
        None,
    )
    .expect("source-backed reader ingress");
    package
        .part(&PackURI::new(DOCUMENT_URI).expect("document URI"))
        .expect("source-backed document part")
        .data()
        .expect("source-backed document payload")
        .as_bytes()
        .to_vec()
}

#[test]
fn reader_routes_preserve_payloads_with_short_and_interrupted_reads() {
    let source = archive_bytes();
    let policy = limits(source.len());

    let physical = document_from_physical(source.clone(), policy);
    let eager = document_from_package(source.clone(), policy);
    let source_backed = document_from_source_backed(source, policy);

    assert_eq!(physical, DOCUMENT_PAYLOAD);
    assert_eq!(eager, physical);
    assert_eq!(source_backed, physical);
}

#[test]
fn reader_routes_report_input_limit_before_zip_validation() {
    let source = b"not a zip archive".to_vec();
    let maximum = source.len() - 1;
    let policy = limits(maximum);

    input_limit_error(
        take_error(
            OwnedPhysPkgReader::from_reader(ChunkedReader::new(source.clone(), 1, 0), policy),
            "physical reader must reject oversized input",
        ),
        maximum,
        maximum + 1,
    );
    input_limit_error(
        take_error(
            OpcPackage::from_reader(ChunkedReader::new(source.clone(), 1, 0), policy),
            "eager package must reject oversized input",
        ),
        maximum,
        maximum + 1,
    );
    input_limit_error(
        take_error(
            SourceBackedPackage::from_reader(
                ChunkedReader::new(source, 1, 0),
                policy,
                SourceCacheLimits::default(),
                None,
            ),
            "source-backed package must reject oversized input",
        ),
        maximum,
        maximum + 1,
    );
}

#[test]
fn reader_routes_preserve_underlying_errors_before_zip_validation() {
    let policy = limits(17);

    for error in [
        take_error(
            OwnedPhysPkgReader::from_reader(
                BrokenReader {
                    kind: io::ErrorKind::BrokenPipe,
                },
                policy,
            ),
            "physical reader must preserve input errors",
        ),
        take_error(
            OpcPackage::from_reader(
                BrokenReader {
                    kind: io::ErrorKind::BrokenPipe,
                },
                policy,
            ),
            "eager package must preserve input errors",
        ),
        take_error(
            SourceBackedPackage::from_reader(
                BrokenReader {
                    kind: io::ErrorKind::BrokenPipe,
                },
                policy,
                SourceCacheLimits::default(),
                None,
            ),
            "source-backed package must preserve input errors",
        ),
    ] {
        assert!(matches!(
            error,
            OpcError::IoError(error) if error.kind() == io::ErrorKind::BrokenPipe
        ));
    }
}

#[test]
fn reader_routes_report_probe_errors_after_exact_limit() {
    let source = archive_bytes();
    let policy = limits(source.len());

    for error in [
        take_error(
            OwnedPhysPkgReader::from_reader(
                ErrorAfterLimit {
                    bytes: source.clone(),
                    offset: 0,
                    fail_at: source.len(),
                },
                policy,
            ),
            "physical reader must report the probe error",
        ),
        take_error(
            OpcPackage::from_reader(
                ErrorAfterLimit {
                    bytes: source.clone(),
                    offset: 0,
                    fail_at: source.len(),
                },
                policy,
            ),
            "eager package must report the probe error",
        ),
        take_error(
            SourceBackedPackage::from_reader(
                ErrorAfterLimit {
                    bytes: source.clone(),
                    offset: 0,
                    fail_at: source.len(),
                },
                policy,
                SourceCacheLimits::default(),
                None,
            ),
            "source-backed package must report the probe error",
        ),
    ] {
        assert!(matches!(
            error,
            OpcError::IoError(error) if error.kind() == io::ErrorKind::BrokenPipe
        ));
    }
}

#[test]
fn physical_path_ingress_uses_file_size_for_input_limit_precheck() {
    let source = archive_bytes();
    let file = NamedTempFile::new().expect("create package path");
    std::fs::write(file.path(), &source).expect("write package path");
    let document = PackURI::new(DOCUMENT_URI).expect("document URI");

    let exact = OwnedPhysPkgReader::open(file.path(), limits(source.len()))
        .expect("exact path input limit");
    assert_eq!(
        exact.blob_for(&document).expect("path document"),
        DOCUMENT_PAYLOAD
    );

    let maximum = source.len() - 1;
    let error = take_error(
        OwnedPhysPkgReader::open(file.path(), limits(maximum)),
        "path metadata must reject an oversized input before opening ZIP",
    );
    input_limit_error(error, maximum, source.len());
}

#[test]
fn input_limit_builder_rejects_zero_and_unaddressable_values() {
    assert!(matches!(
        SourceLimits::builder().max_input_bytes(0),
        Err(OpcError::InvalidLimit {
            kind: SourceLimitKind::InputBytes,
            value: 0,
            ..
        })
    ));
    assert!(matches!(
        SourceLimits::builder().max_input_bytes(u64::MAX),
        Err(OpcError::InvalidLimit {
            kind: SourceLimitKind::InputBytes,
            value: u64::MAX,
            ..
        })
    ));
}
