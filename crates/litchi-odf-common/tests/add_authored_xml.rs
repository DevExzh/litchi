#![allow(
    clippy::unwrap_used,
    reason = "These are fixed-package boundary tests; an unexpected result is a test failure."
)]

//! Typed authored-XML publication, preflight, and transport failure contracts.

use litchi_core::Error;
use litchi_odf_common::core::{PackageWriter, PackageWriterError, PackageWriterLimits, Profile};
use litchi_odf_common::signature::{DocumentSigner, SignatureAlgorithm};
use soapberry_zip::office::ArchiveReader;
use std::cell::RefCell;
use std::error::Error as StdError;
use std::io::{self, Write};
use std::rc::Rc;

const MIME: &str = "application/vnd.oasis.opendocument.text";
const XML_MEDIA_TYPE: &str = "text/xml";
const AUTHORED_XML: &[u8] = b"<root><child>payload</child></root>";

const RSA_KEY: &[u8] = include_bytes!("fixtures/signatures/rsa-key.pk8");
const RSA_CERT: &[u8] = include_bytes!("fixtures/signatures/rsa-cert.der");

#[derive(Debug, Clone)]
struct SharedSink(Rc<RefCell<Vec<u8>>>);

impl Write for SharedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.0.borrow_mut().extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn shared_writer(bytes: &Rc<RefCell<Vec<u8>>>) -> PackageWriter<SharedSink> {
    PackageWriter::with_writer(SharedSink(Rc::clone(bytes)))
}

fn signer() -> DocumentSigner {
    DocumentSigner::from_pkcs8_der(
        SignatureAlgorithm::RsaSha256,
        RSA_KEY,
        vec![RSA_CERT.to_vec()],
        "2026-07-19T12:00:00Z",
    )
    .unwrap()
}

fn authored_package_with(content: &[u8]) -> Vec<u8> {
    // Keep the expected archive on the same sequential/typed path as the
    // bounded controls below.  In particular, do not compare `new()` plus the
    // compatibility `set_mimetype` path to a caller-owned sink publication.
    let mut sink = Vec::new();
    let mut writer = PackageWriter::with_writer(&mut sink);
    writer.set_mimetype_streaming(MIME).unwrap();
    writer
        .add_authored_xml("content.xml", content, XML_MEDIA_TYPE)
        .unwrap();
    let _ = writer.finish_to_writer().unwrap();
    sink
}

fn authored_package() -> Vec<u8> {
    authored_package_with(AUTHORED_XML)
}

fn assert_core_invalid_format(error: PackageWriterError) {
    assert!(matches!(
        error,
        PackageWriterError::Core(Error::InvalidFormat(_))
    ));
}

#[test]
fn authored_xml_preserves_valid_comments_byte_for_byte() {
    let content = b"<root><!--keep this authored comment--><child>text</child></root>";
    let output = authored_package_with(content);

    let archive = ArchiveReader::new(&output).unwrap();
    assert_eq!(archive.read("content.xml").unwrap(), content);
}

#[test]
fn malformed_authored_xml_is_rejected_before_member_output_and_is_retryable() {
    let bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = shared_writer(&bytes);
    writer.set_mimetype_streaming(MIME).unwrap();
    let before = bytes.borrow().len();

    let error = writer
        .add_authored_xml("content.xml", b"<root><child></root>", XML_MEDIA_TYPE)
        .unwrap_err();
    assert_core_invalid_format(error);
    assert_eq!(bytes.borrow().len(), before);

    // A validation refusal is preflight-only; it must not poison the writer.
    writer
        .add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)
        .unwrap();
    let _ = writer.finish_to_writer().unwrap();
}

#[test]
fn authored_xml_refuses_non_xml_reserved_and_duplicate_members_atomically() {
    for (path, media_type) in [
        ("Pictures/payload.bin", "application/octet-stream"),
        ("mimetype", XML_MEDIA_TYPE),
        ("META-INF/manifest.xml", XML_MEDIA_TYPE),
    ] {
        let bytes = Rc::new(RefCell::new(Vec::new()));
        let mut writer = shared_writer(&bytes);
        writer.set_mimetype_streaming(MIME).unwrap();
        let before = bytes.borrow().len();
        let error = writer
            .add_authored_xml(path, AUTHORED_XML, media_type)
            .unwrap_err();
        assert_core_invalid_format(error);
        assert_eq!(
            bytes.borrow().len(),
            before,
            "refusal wrote bytes for {path:?}"
        );
    }

    let bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = shared_writer(&bytes);
    writer.set_mimetype_streaming(MIME).unwrap();
    writer
        .add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)
        .unwrap();
    let before_duplicate = bytes.borrow().len();
    let duplicate = writer
        .add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)
        .unwrap_err();
    assert_core_invalid_format(duplicate);
    assert_eq!(bytes.borrow().len(), before_duplicate);
    let _ = writer.finish_to_writer().unwrap();
}

#[test]
fn authored_xml_refuses_encryption_and_signing_before_member_output() {
    let encrypted_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut encrypted = shared_writer(&encrypted_bytes);
    encrypted.set_mimetype_streaming(MIME).unwrap();
    encrypted
        .set_encryption("secret", Profile::compatible())
        .unwrap();
    let encrypted_before = encrypted_bytes.borrow().len();
    let encrypted_error = encrypted
        .add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)
        .unwrap_err();
    assert_core_invalid_format(encrypted_error);
    assert_eq!(encrypted_bytes.borrow().len(), encrypted_before);

    let signed_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut signed = shared_writer(&signed_bytes);
    signed.set_mimetype_streaming(MIME).unwrap();
    signed.set_document_signer(signer()).unwrap();
    let signed_before = signed_bytes.borrow().len();
    let signed_error = signed
        .add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)
        .unwrap_err();
    assert_core_invalid_format(signed_error);
    assert_eq!(signed_bytes.borrow().len(), signed_before);
}

#[derive(Debug)]
struct AuthoredIoMarker;

impl std::fmt::Display for AuthoredIoMarker {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("authored XML sink marker")
    }
}

impl StdError for AuthoredIoMarker {}

#[derive(Debug)]
struct MarkerSink {
    bytes: Rc<RefCell<Vec<u8>>>,
    remaining: usize,
}

impl Write for MarkerSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, AuthoredIoMarker));
        }
        let count = bytes.len().min(self.remaining);
        self.bytes.borrow_mut().extend_from_slice(&bytes[..count]);
        self.remaining -= count;
        if count < bytes.len() {
            // A short write must report the accepted prefix as `Ok(count)`;
            // returning an error here would make the caller unable to account
            // for bytes that this sink already accepted.  The next write is
            // the failing operation and carries the marker.
            return Ok(count);
        }
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn assert_authored_io_marker(error: &PackageWriterError) {
    let mut current: &(dyn StdError + 'static) = error;
    loop {
        if let Some(io_error) = current.downcast_ref::<io::Error>()
            && io_error
                .get_ref()
                .and_then(|source| source.downcast_ref::<AuthoredIoMarker>())
                .is_some()
        {
            return;
        }
        let Some(next) = current.source() else {
            break;
        };
        current = next;
    }
    panic!("authored XML sink marker was lost: {error:?}");
}

#[test]
fn authored_xml_io_failure_retains_marker_and_poisoned_finalize() {
    let bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(MarkerSink {
        bytes: Rc::clone(&bytes),
        // Leave enough room for the mimetype and the beginning of the authored
        // member, then fail during its local header/data publication.
        remaining: 128,
    });
    writer.set_mimetype_streaming(MIME).unwrap();
    let before_member = bytes.borrow().len();
    let error = writer
        .add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)
        .unwrap_err();
    assert!(bytes.borrow().len() > before_member);
    assert!(error.written().unwrap() > before_member as u64);
    assert_eq!(error.written(), Some(bytes.borrow().len() as u64));
    assert_authored_io_marker(&error);

    // A failed member poisons ZIP publication; finalization must fail rather
    // than silently append a manifest to an incomplete archive.
    let before_finalize = bytes.borrow().len();
    let finalize_error = writer.finish_to_writer().unwrap_err();
    assert_eq!(finalize_error.written(), Some(before_finalize as u64));
    assert_eq!(bytes.borrow().len(), before_finalize);
    assert!(finalize_error.to_string().contains("poisoned"));
}

fn output_limited_writer(sink: &mut Vec<u8>, maximum_output: u64) -> PackageWriter<&mut Vec<u8>> {
    let limits = PackageWriterLimits::new(3, 32, maximum_output.min(64 * 1024)).with_byte_limits(
        64 * 1024,
        64 * 1024,
        maximum_output,
    );
    PackageWriter::with_writer_and_limits(sink, limits)
}

#[test]
fn authored_xml_output_limit_is_inclusive_and_one_byte_under_is_typed() {
    let expected = authored_package();
    let expected_size = expected.len() as u64;

    let mut exact_sink = Vec::new();
    let mut exact = output_limited_writer(&mut exact_sink, expected_size);
    exact.set_mimetype_streaming(MIME).unwrap();
    exact
        .add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)
        .unwrap();
    let _ = exact.finish_to_writer().unwrap();
    assert_eq!(exact_sink, expected);

    let mut under_sink = Vec::new();
    let mut under = output_limited_writer(&mut under_sink, expected_size - 1);
    under.set_mimetype_streaming(MIME).unwrap();
    under
        .add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)
        .unwrap();
    let error = under.finish_to_writer().unwrap_err();
    match error {
        PackageWriterError::LimitExceeded { written, limit, .. } => {
            assert_eq!(written, under_sink.len() as u64);
            assert_eq!(
                limit.resource(),
                litchi_odf_common::core::PackageWriterLimitResource::OutputBytes
            );
            assert_eq!(limit.maximum(), expected_size - 1);
            assert!(limit.actual() > limit.maximum());
        },
        other => panic!("expected typed output limit, got {other:?}"),
    }
    assert!(under_sink.len() < expected.len());
}

fn metadata_limits(max_metadata_bytes: u64) -> PackageWriterLimits {
    let defaults = PackageWriterLimits::default();
    PackageWriterLimits::new(
        defaults.max_entries,
        defaults.max_member_name_bytes,
        max_metadata_bytes,
    )
    .with_byte_limits(
        defaults.max_entry_size,
        defaults.max_total_size,
        defaults.max_output_bytes,
    )
}

fn metadata_limited_package(max_metadata_bytes: u64) -> Result<Vec<u8>, PackageWriterError> {
    let limits = metadata_limits(max_metadata_bytes);
    let mut sink = Vec::new();
    let mut writer = PackageWriter::with_writer_and_limits(&mut sink, limits);
    writer.set_mimetype_streaming(MIME)?;
    writer.add_authored_xml("content.xml", AUTHORED_XML, XML_MEDIA_TYPE)?;
    let _ = writer.finish_to_writer()?;
    Ok(sink)
}

fn minimum_successful_metadata_limit() -> u64 {
    let maximum = PackageWriterLimits::default().max_metadata_bytes;
    assert!(metadata_limited_package(maximum).is_ok());
    let mut low = 0_u64;
    let mut high = maximum;
    while low < high {
        let midpoint = low + (high - low) / 2;
        if metadata_limited_package(midpoint).is_ok() {
            high = midpoint;
        } else {
            low = midpoint + 1;
        }
    }
    assert!(metadata_limited_package(low).is_ok());
    low
}

#[test]
fn authored_xml_metadata_limit_is_inclusive_and_one_byte_under_is_typed() {
    let expected = authored_package();
    let metadata_bytes = minimum_successful_metadata_limit();
    assert!(metadata_bytes > 0);

    let exact = metadata_limited_package(metadata_bytes).unwrap();
    assert_eq!(exact, expected);

    let error = metadata_limited_package(metadata_bytes - 1).unwrap_err();
    match error {
        PackageWriterError::LimitExceeded { limit, written, .. } => {
            assert_eq!(
                limit.resource(),
                litchi_odf_common::core::PackageWriterLimitResource::MetadataBytes
            );
            assert_eq!(limit.maximum(), metadata_bytes - 1);
            assert!(limit.actual() > limit.maximum());
            assert!(written > 0);
        },
        other => panic!("expected typed metadata limit, got {other:?}"),
    }
}
