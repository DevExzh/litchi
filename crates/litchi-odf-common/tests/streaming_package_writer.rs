#![allow(
    clippy::unwrap_used,
    reason = "Streaming writer tests use fixed inputs and direct assertions."
)]

use litchi_core::Error;
use litchi_odf_common::core::{
    OwnedPackage, PackageCompression, PackageWriter, PackageWriterError, PackageWriterLimits,
    Structure,
};
use litchi_odf_common::signature::{DocumentSigner, SignatureAlgorithm};
use sha2::{Digest, Sha256};
use soapberry_zip::office::ArchiveReader;
use std::cell::RefCell;
use std::io::{self, Cursor, Read, Write};
use std::rc::Rc;

const MIME: &str = "application/vnd.oasis.opendocument.text";
const RSA_KEY: &[u8] = include_bytes!("fixtures/signatures/rsa-key.pk8");
const RSA_CERT: &[u8] = include_bytes!("fixtures/signatures/rsa-cert.der");

fn payload() -> Vec<u8> {
    (0..32_768).map(|index| (index % 251) as u8).collect()
}

fn digest(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

fn memory_artifact(bytes: &[u8]) -> Vec<u8> {
    let mut writer = PackageWriter::new();
    writer.set_mimetype(MIME).unwrap();
    writer
        .add_file_reader_with_media_type_and_compression(
            "Pictures/stream.bin",
            Cursor::new(bytes),
            "application/octet-stream",
            PackageCompression::Stored,
        )
        .unwrap();
    writer.finish().unwrap()
}

#[derive(Debug)]
struct ShortSink {
    bytes: Rc<RefCell<Vec<u8>>>,
    maximum_write: usize,
}

impl Write for ShortSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let count = bytes.len().min(self.maximum_write);
        if count == 0 {
            return Err(io::Error::new(
                io::ErrorKind::WriteZero,
                "short sink limit is zero",
            ));
        }
        self.bytes.borrow_mut().extend_from_slice(&bytes[..count]);
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct FailingSink {
    bytes: Rc<RefCell<Vec<u8>>>,
    remaining: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        if self.remaining == 0 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "test sink failed",
            ));
        }
        let count = bytes.len().min(self.remaining);
        self.bytes.borrow_mut().extend_from_slice(&bytes[..count]);
        self.remaining -= count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct FailingReader {
    bytes: Vec<u8>,
    offset: usize,
}

impl Read for FailingReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.offset == self.bytes.len() {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                "test reader failed",
            ));
        }
        let count = output
            .len()
            .min(self.bytes.len().saturating_sub(self.offset));
        output[..count].copy_from_slice(&self.bytes[self.offset..self.offset + count]);
        self.offset += count;
        Ok(count)
    }
}

#[derive(Debug)]
struct NestedPublicationSource;

impl std::fmt::Display for NestedPublicationSource {
    fn fmt(&self, formatter: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        formatter.write_str("nested publication failure")
    }
}

impl std::error::Error for NestedPublicationSource {}

#[derive(Debug)]
struct NestedReader {
    emitted: bool,
}

impl Read for NestedReader {
    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.emitted {
            return Err(io::Error::new(
                io::ErrorKind::InvalidData,
                NestedPublicationSource,
            ));
        }
        output[0] = b'x';
        self.emitted = true;
        Ok(1)
    }
}

#[derive(Debug)]
struct NestedFailingSink {
    bytes: Rc<RefCell<Vec<u8>>>,
    remaining: usize,
}

impl Write for NestedFailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if bytes.is_empty() {
            return Ok(0);
        }
        if self.remaining == 0 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                NestedPublicationSource,
            ));
        }
        let count = bytes.len().min(self.remaining);
        self.bytes.borrow_mut().extend_from_slice(&bytes[..count]);
        self.remaining -= count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
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

#[test]
fn short_non_seek_sink_round_trips_exact_stream_and_is_deterministic() {
    let bytes = payload();
    let expected_digest = digest(&bytes);
    let expected = memory_artifact(&bytes);

    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let sink = ShortSink {
        bytes: Rc::clone(&sink_bytes),
        maximum_write: 7,
    };
    let mut writer = PackageWriter::with_writer(sink);
    writer.set_mimetype(MIME).unwrap();
    writer
        .add_file_reader_with_media_type_and_compression(
            "Pictures/stream.bin",
            Cursor::new(bytes.clone()),
            "application/octet-stream",
            PackageCompression::Stored,
        )
        .unwrap();
    writer.finish_to_writer().unwrap();

    let output = sink_bytes.borrow();
    let archive = ArchiveReader::new(&output).unwrap();
    assert!(archive.is_stored("mimetype").unwrap());
    assert_eq!(
        digest(&archive.read("Pictures/stream.bin").unwrap()),
        expected_digest
    );
    assert_eq!(*output, expected);
}

#[test]
fn exact_output_limit_succeeds_and_one_byte_under_reports_limit_progress() {
    let bytes = payload();
    let expected = memory_artifact(&bytes);
    let mut exact_limits = PackageWriterLimits::default().with_byte_limits(
        64 * 1024,
        64 * 1024,
        expected.len() as u64,
    );
    exact_limits.max_metadata_bytes = expected.len() as u64;

    let mut exact_sink = Vec::new();
    let mut exact = PackageWriter::with_writer_and_limits(&mut exact_sink, exact_limits);
    exact.set_mimetype(MIME).unwrap();
    exact
        .add_file_reader_with_media_type_and_compression(
            "Pictures/stream.bin",
            Cursor::new(bytes.clone()),
            "application/octet-stream",
            PackageCompression::Stored,
        )
        .unwrap();
    exact.finish_to_writer().unwrap();
    assert_eq!(exact_sink, expected);

    let mut under_limits = PackageWriterLimits::default().with_byte_limits(
        64 * 1024,
        64 * 1024,
        (expected.len() - 1) as u64,
    );
    under_limits.max_metadata_bytes = (expected.len() - 1) as u64;
    let mut under_sink = Vec::new();
    let mut under = PackageWriter::with_writer_and_limits(&mut under_sink, under_limits);
    under.set_mimetype(MIME).unwrap();
    under
        .add_file_reader_with_media_type_and_compression(
            "Pictures/stream.bin",
            Cursor::new(bytes),
            "application/octet-stream",
            PackageCompression::Stored,
        )
        .unwrap();
    let error = under.finish_to_writer().unwrap_err();
    match error {
        PackageWriterError::LimitExceeded { written, limit, .. } => {
            assert_eq!(written, under_sink.len() as u64);
            assert!(written < expected.len() as u64);
            assert!(limit.actual() > limit.maximum());
            assert_eq!(limit.maximum(), (expected.len() - 1) as u64);
        },
        other => panic!("expected typed output limit, got {other:?}"),
    }
    assert!(under_sink.len() < expected.len());
}

#[test]
fn streamed_entry_limit_is_typed_after_the_member_header_starts() {
    let bytes = payload();
    let mut limits = PackageWriterLimits::default().with_byte_limits(
        (bytes.len() - 1) as u64,
        64 * 1024,
        1024 * 1024,
    );
    limits.max_metadata_bytes = 1024 * 1024;
    let mut sink = Vec::new();
    let mut writer = PackageWriter::with_writer_and_limits(&mut sink, limits);
    writer.set_mimetype(MIME).unwrap();
    let error = writer
        .add_file_reader_with_media_type_and_compression(
            "Pictures/stream.bin",
            Cursor::new(bytes.clone()),
            "application/octet-stream",
            PackageCompression::Stored,
        )
        .unwrap_err();
    match error {
        PackageWriterError::LimitExceeded { written, limit, .. } => {
            assert!(written > 0);
            assert_eq!(
                limit.resource(),
                litchi_odf_common::core::PackageWriterLimitResource::EntrySize
            );
            assert_eq!(limit.actual(), bytes.len() as u64);
            assert_eq!(limit.maximum(), (bytes.len() - 1) as u64);
        },
        other => panic!("expected typed entry limit, got {other:?}"),
    }
}

#[test]
fn reader_failure_reports_incomplete_output_and_poisoned_writer() {
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&sink_bytes),
        maximum_write: usize::MAX,
    });
    writer.set_mimetype(MIME).unwrap();
    let error = writer
        .add_file_reader_with_media_type(
            "Pictures/failing.bin",
            FailingReader {
                bytes: payload(),
                offset: 0,
            },
            "application/octet-stream",
        )
        .unwrap_err();
    match error {
        PackageWriterError::IncompleteOutput { written, source } => {
            assert_eq!(written, sink_bytes.borrow().len() as u64);
            assert!(written > 0);
            assert!(matches!(source.as_ref(), PackageWriterError::Archive(_)));
        },
        other => panic!("expected incomplete output, got {other:?}"),
    }
    assert!(
        writer
            .add_file_reader_with_media_type(
                "Pictures/after-failure.bin",
                Cursor::new(b"later"),
                "application/octet-stream",
            )
            .is_err()
    );
}

#[test]
fn reader_failure_preserves_nested_source_chain() {
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&sink_bytes),
        maximum_write: usize::MAX,
    });
    writer.set_mimetype_streaming(MIME).unwrap();
    let error = writer
        .add_file_reader_with_media_type(
            "Pictures/nested-reader.bin",
            NestedReader { emitted: false },
            "application/octet-stream",
        )
        .unwrap_err();
    let source = match error {
        PackageWriterError::IncompleteOutput { source, .. } => source,
        other => panic!("expected incomplete output, got {other:?}"),
    };
    let archive = match source.as_ref() {
        PackageWriterError::Archive(error) => error,
        other => panic!("expected archive source, got {other:?}"),
    };
    let io_error = match archive.kind() {
        soapberry_zip::ErrorKind::IO(error) | soapberry_zip::ErrorKind::Io(error) => error,
        other => panic!("expected retained I/O source, got {other:?}"),
    };
    assert!(
        io_error
            .get_ref()
            .and_then(|nested| nested.downcast_ref::<NestedPublicationSource>())
            .is_some(),
        "nested reader source was lost"
    );
    let source = std::error::Error::source(source.as_ref()).expect("archive source");
    assert!(source.downcast_ref::<io::Error>().is_some());
    assert!(!sink_bytes.borrow().is_empty());
}

#[test]
fn finalization_failure_preserves_accepted_output_progress() {
    let bytes = payload();
    let expected = memory_artifact(&bytes);
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let sink = FailingSink {
        bytes: Rc::clone(&sink_bytes),
        remaining: expected.len() - 1,
    };
    let mut writer = PackageWriter::with_writer(sink);
    writer.set_mimetype(MIME).unwrap();
    writer
        .add_file_reader_with_media_type_and_compression(
            "Pictures/stream.bin",
            Cursor::new(bytes),
            "application/octet-stream",
            PackageCompression::Stored,
        )
        .unwrap();
    let error = writer.finish_to_writer().unwrap_err();
    match error {
        PackageWriterError::IncompleteOutput { written, source } => {
            assert_eq!(written, (expected.len() - 1) as u64);
            assert_eq!(sink_bytes.borrow().len(), expected.len() - 1);
            assert!(matches!(
                source.as_ref(),
                PackageWriterError::ArchiveFailure(_)
            ));
        },
        other => panic!("expected incomplete finalization, got {other:?}"),
    }
}

#[test]
fn finalization_sink_failure_preserves_nested_source_chain() {
    let bytes = payload();
    let expected = {
        let mut probe = PackageWriter::new();
        probe.set_mimetype(MIME).unwrap();
        probe
            .add_file_reader_with_media_type_and_compression(
                "Pictures/nested-sink.bin",
                Cursor::new(bytes.clone()),
                "application/octet-stream",
                PackageCompression::Stored,
            )
            .unwrap();
        probe.finish().unwrap()
    };
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let sink = NestedFailingSink {
        bytes: Rc::clone(&sink_bytes),
        remaining: expected.len() - 1,
    };
    let mut writer = PackageWriter::with_writer(sink);
    writer.set_mimetype_streaming(MIME).unwrap();
    writer
        .add_file_reader_with_media_type_and_compression(
            "Pictures/nested-sink.bin",
            Cursor::new(bytes),
            "application/octet-stream",
            PackageCompression::Stored,
        )
        .unwrap();
    let error = writer.finish_to_writer().unwrap_err();
    let source = match error {
        PackageWriterError::IncompleteOutput { source, .. } => source,
        other => panic!("expected incomplete output, got {other:?}"),
    };
    let failure = match source.as_ref() {
        PackageWriterError::ArchiveFailure(failure) => failure,
        other => panic!("expected archive failure source, got {other:?}"),
    };
    let io_error = match failure.error().kind() {
        soapberry_zip::ErrorKind::IO(error) | soapberry_zip::ErrorKind::Io(error) => error,
        other => panic!("expected retained I/O source, got {other:?}"),
    };
    assert!(
        io_error
            .get_ref()
            .and_then(|nested| nested.downcast_ref::<NestedPublicationSource>())
            .is_some(),
        "nested sink source was lost"
    );
    let source = std::error::Error::source(source.as_ref()).expect("archive failure source");
    assert!(source.downcast_ref::<io::Error>().is_some());
    assert!(!sink_bytes.borrow().is_empty());
}

#[test]
fn duplicate_and_xml_refusals_do_not_emit_new_member_headers() {
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&sink_bytes),
        maximum_write: usize::MAX,
    });
    writer.set_mimetype(MIME).unwrap();
    writer
        .add_file_reader_with_media_type(
            "Pictures/one.bin",
            Cursor::new(b"one"),
            "application/octet-stream",
        )
        .unwrap();
    let before_duplicate = sink_bytes.borrow().len();
    let duplicate = writer
        .add_file_reader_with_media_type(
            "Pictures/one.bin",
            Cursor::new(b"duplicate"),
            "application/octet-stream",
        )
        .unwrap_err();
    assert!(matches!(duplicate, PackageWriterError::Core(_)));
    assert_eq!(sink_bytes.borrow().len(), before_duplicate);

    let before_xml = sink_bytes.borrow().len();
    let xml = writer
        .add_file_reader_with_media_type("content.xml", Cursor::new(b"<root/>"), "text/xml")
        .unwrap_err();
    match xml {
        PackageWriterError::Core(Error::InvalidFormat(message)) => {
            assert!(message.contains("rejects XML member"));
        },
        other => panic!("expected XML refusal, got {other:?}"),
    }
    assert_eq!(sink_bytes.borrow().len(), before_xml);

    writer
        .add_file_reader_with_media_type(
            "Pictures/two.bin",
            Cursor::new(b"two"),
            "application/octet-stream",
        )
        .unwrap();
    writer.finish_to_writer().unwrap();
}

#[test]
fn encryption_and_signing_are_refused_before_streamed_member_output() {
    let encrypted_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut encrypted = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&encrypted_bytes),
        maximum_write: usize::MAX,
    });
    encrypted.set_mimetype(MIME).unwrap();
    encrypted
        .set_encryption("secret", litchi_odf_common::core::Profile::compatible())
        .unwrap();
    let encrypted_before = encrypted_bytes.borrow().len();
    let encrypted_error = encrypted
        .add_file_reader_with_media_type(
            "Pictures/encrypted.bin",
            Cursor::new(b"payload"),
            "application/octet-stream",
        )
        .unwrap_err();
    assert!(matches!(encrypted_error, PackageWriterError::Core(_)));
    assert_eq!(encrypted_bytes.borrow().len(), encrypted_before);

    let signed_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut signed = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&signed_bytes),
        maximum_write: usize::MAX,
    });
    signed.set_mimetype(MIME).unwrap();
    signed.set_document_signer(signer()).unwrap();
    let signed_before = signed_bytes.borrow().len();
    let signed_error = signed
        .add_file_reader_with_media_type(
            "Pictures/signed.bin",
            Cursor::new(b"payload"),
            "application/octet-stream",
        )
        .unwrap_err();
    assert!(matches!(signed_error, PackageWriterError::Core(_)));
    assert_eq!(signed_bytes.borrow().len(), signed_before);

    let finalization_error = signed.finish_to_writer().unwrap_err();
    match finalization_error {
        PackageWriterError::IncompleteOutput { written, source } => {
            assert_eq!(written, signed_before as u64);
            assert!(matches!(source.as_ref(), PackageWriterError::Core(_)));
        },
        PackageWriterError::Core(Error::InvalidFormat(message)) => {
            assert!(message.contains("does not support document signing"));
        },
        other => panic!("expected signing refusal, got {other:?}"),
    }
    assert_eq!(signed_bytes.borrow().len(), signed_before);
}

#[test]
fn canonical_paths_reopen_through_owned_indexed_and_family_packages() {
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&sink_bytes),
        maximum_write: usize::MAX,
    });
    writer.set_mimetype_streaming(MIME).unwrap();
    let content = Structure::default_content_xml("office:text");
    writer.add_file("content.xml", content.as_bytes()).unwrap();

    for path in [
        "",
        "Pictures/",
        "META-INF/",
        "META-INF",
        "META-INF/manifest.xml",
        "META-INF/documentsignatures.xml",
        "Pictures/../alias.bin",
        "./Pictures/alias.bin",
        "/Pictures/alias.bin",
        "Pictures//alias.bin",
        "Pictures\\alias.bin",
        "Pictures/query?cache=1.bin",
        "Pictures/hash#fragment.bin",
        "Pictures/percent%2Ebin",
    ] {
        let before = sink_bytes.borrow().len();
        let error = writer
            .add_file_reader_with_media_type(
                path,
                Cursor::new(b"hostile"),
                "application/octet-stream",
            )
            .unwrap_err();
        assert!(matches!(error, PackageWriterError::Core(_)));
        assert_eq!(sink_bytes.borrow().len(), before, "path: {path:?}");
    }

    let before_generic_alias = sink_bytes.borrow().len();
    assert!(
        writer
            .add_file_with_media_type(
                "Pictures/generic%2Ebin",
                b"hostile",
                "application/octet-stream",
            )
            .is_err()
    );
    assert_eq!(sink_bytes.borrow().len(), before_generic_alias);

    let before_generic_admin = sink_bytes.borrow().len();
    assert!(writer.add_file("META-INF", b"hostile").is_err());
    assert_eq!(sink_bytes.borrow().len(), before_generic_admin);
    assert!(
        writer
            .add_file_with_media_type("META-INF", b"hostile", "application/octet-stream")
            .is_err()
    );
    assert_eq!(sink_bytes.borrow().len(), before_generic_admin);

    writer
        .add_file_reader_with_media_type(
            "Pictures/canonical.bin",
            Cursor::new(b"canonical"),
            "application/octet-stream",
        )
        .unwrap();
    writer.finish_to_writer().unwrap();
    let output = sink_bytes.borrow().clone();

    let owned = OwnedPackage::from_bytes(output.clone()).unwrap();
    let indexed = owned.package().unwrap();
    assert_eq!(
        indexed.get_file("Pictures/canonical.bin").unwrap(),
        b"canonical"
    );
    assert_eq!(owned.mimetype().unwrap(), MIME);

    let prepared = litchi_odf_common::detect::prepared(output.clone()).unwrap();
    assert_eq!(
        prepared
            .package()
            .get_file("Pictures/canonical.bin")
            .unwrap(),
        b"canonical"
    );

    let family =
        litchi_odf_common::core::Package::from_bytes(output, MIME, "<office:text", "ODT").unwrap();
    assert_eq!(
        family.package().get_file("Pictures/canonical.bin").unwrap(),
        b"canonical"
    );
}

#[test]
fn manifest_only_paths_use_strict_odf_validation_and_reopen() {
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&sink_bytes),
        maximum_write: usize::MAX,
    });
    writer.set_mimetype_streaming(MIME).unwrap();
    let content = Structure::default_content_xml("office:text");
    writer.add_file("content.xml", content.as_bytes()).unwrap();

    for path in [
        "",
        "Object/../escape/",
        "Object/%2E%2E/escape/",
        "Object/query?cache=1/",
        "Object/hash#fragment/",
        "Object/percent%2Ebin/",
        "./Object/alias/",
        "/Object/absolute/",
        "Object//empty-segment/",
        "META-INF/",
        "META-INF",
        "manifest.xml",
        "META-INF/manifest.xml",
        "META-INF/documentsignatures.xml",
        "mimetype",
    ] {
        let before = sink_bytes.borrow().len();
        assert!(
            writer.add_manifest_entry(path, "").is_err(),
            "path: {path:?}"
        );
        assert_eq!(sink_bytes.borrow().len(), before, "path: {path:?}");
    }

    writer
        .add_manifest_directory("Object..1/", "")
        .expect("double-dot characters inside a segment are not traversal");
    writer.finish_to_writer().unwrap();
    let output = sink_bytes.borrow().clone();

    let owned = OwnedPackage::from_bytes(output.clone()).unwrap();
    let indexed = owned.package().unwrap();
    assert_eq!(indexed.manifest().get_media_type("Object..1/"), Some(""));
    let prepared = litchi_odf_common::detect::prepared(output.clone()).unwrap();
    assert_eq!(
        prepared
            .package()
            .package()
            .unwrap()
            .manifest()
            .get_media_type("Object..1/"),
        Some("")
    );
    let family =
        litchi_odf_common::core::Package::from_bytes(output, MIME, "<office:text", "ODT").unwrap();
    assert_eq!(
        family
            .package()
            .package()
            .unwrap()
            .manifest()
            .get_media_type("Object..1/"),
        Some("")
    );
}

#[test]
fn metadata_is_validated_and_bounded_before_member_output() {
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&sink_bytes),
        maximum_write: usize::MAX,
    });
    let invalid_mime = writer.set_mimetype_streaming("not-a-mime").unwrap_err();
    assert!(matches!(invalid_mime, PackageWriterError::Core(_)));
    assert!(sink_bytes.borrow().is_empty());
    writer.set_mimetype_streaming(MIME).unwrap();

    for mimetype in [
        " application/vnd.oasis.opendocument.text",
        "application/vnd.oasis.opendocument.text ",
        "application/vnd.oasis.opendocument. text",
        "application/vnd.oasis.opendocument.text\t",
        "application/vnd.oasis.opendocument.text/extra",
        "application/vnd.oasis.opendocument.text; charset=utf-8",
    ] {
        let candidate_bytes = Rc::new(RefCell::new(Vec::new()));
        let mut candidate = PackageWriter::with_writer(ShortSink {
            bytes: Rc::clone(&candidate_bytes),
            maximum_write: usize::MAX,
        });
        let error = candidate.set_mimetype_streaming(mimetype).unwrap_err();
        assert!(
            matches!(error, PackageWriterError::Core(_)),
            "MIME: {mimetype:?}"
        );
        assert!(candidate_bytes.borrow().is_empty(), "MIME: {mimetype:?}");
    }

    let oversized_media_type = format!("application/{}", "m".repeat(2_000));
    for media_type in [
        "application/octet-stream\0",
        " text/plain",
        "text/plain ",
        "text/ plain",
        "text/plain/extra",
        oversized_media_type.as_str(),
    ] {
        let before = sink_bytes.borrow().len();
        let error = writer
            .add_file_reader_with_media_type(
                "Pictures/metadata.bin",
                Cursor::new(b"metadata"),
                media_type,
            )
            .unwrap_err();
        assert!(matches!(error, PackageWriterError::Core(_)));
        assert_eq!(sink_bytes.borrow().len(), before);
    }

    let parameterized_media_type = "text/plain; charset=utf-8";
    writer
        .add_file_reader_with_media_type(
            "Pictures/metadata-parameterized.bin",
            Cursor::new(b"metadata"),
            parameterized_media_type,
        )
        .unwrap();
    writer.finish_to_writer().unwrap();
    let output = sink_bytes.borrow().clone();
    let owned = OwnedPackage::from_bytes(output).unwrap();
    assert_eq!(
        owned
            .package()
            .unwrap()
            .manifest()
            .get_media_type("Pictures/metadata-parameterized.bin"),
        Some(parameterized_media_type)
    );

    let metadata_limits = PackageWriterLimits {
        max_metadata_bytes: 1_024,
        ..PackageWriterLimits::default()
    };
    let mut bounded_sink = Vec::new();
    let mut bounded = PackageWriter::with_writer_and_limits(&mut bounded_sink, metadata_limits);
    bounded.set_mimetype_streaming(MIME).unwrap();
    let long_path = format!("Pictures/{}.bin", "p".repeat(700));
    let error = bounded
        .add_file_reader_with_media_type(
            &long_path,
            Cursor::new(b"metadata"),
            "application/octet-stream",
        )
        .unwrap_err();
    match error {
        PackageWriterError::LimitExceeded { limit, written, .. } => {
            assert_eq!(
                limit.resource(),
                litchi_odf_common::core::PackageWriterLimitResource::MetadataBytes
            );
            assert_eq!(written, bounded_sink.len() as u64);
        },
        other => panic!("expected bounded manifest metadata, got {other:?}"),
    }
}

#[test]
fn manifest_collisions_and_finish_entry_limit_are_atomic_with_progress() {
    let collision_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut collision_writer = PackageWriter::with_writer(ShortSink {
        bytes: Rc::clone(&collision_bytes),
        maximum_write: usize::MAX,
    });
    collision_writer.set_mimetype_streaming(MIME).unwrap();
    collision_writer
        .add_manifest_entry("Pictures/collision.bin", "application/octet-stream")
        .unwrap();
    let before_collision = collision_bytes.borrow().len();
    let collision = collision_writer
        .add_file_reader_with_media_type(
            "Pictures/collision.bin",
            Cursor::new(b"collision"),
            "application/octet-stream",
        )
        .unwrap_err();
    assert!(matches!(collision, PackageWriterError::Core(_)));
    assert_eq!(collision_bytes.borrow().len(), before_collision);
    collision_writer.finish_to_writer().unwrap();

    let limits = PackageWriterLimits {
        max_entries: 2,
        ..PackageWriterLimits::default()
    };
    let mut sink = Vec::new();
    let mut limited = PackageWriter::with_writer_and_limits(&mut sink, limits);
    limited.set_mimetype_streaming(MIME).unwrap();
    limited
        .add_file_reader_with_media_type(
            "Pictures/ok.bin",
            Cursor::new(b"ok"),
            "application/octet-stream",
        )
        .unwrap();
    let error = limited.finish_to_writer().unwrap_err();
    match error {
        PackageWriterError::LimitExceeded { limit, written, .. } => {
            assert_eq!(
                limit.resource(),
                litchi_odf_common::core::PackageWriterLimitResource::FileCount
            );
            assert_eq!(written, sink.len() as u64);
            assert!(written > 0);
        },
        other => panic!("expected finish entry-count limit, got {other:?}"),
    }
}

#[test]
fn typed_mimetype_and_reader_failures_preserve_archive_sources() {
    let sink_bytes = Rc::new(RefCell::new(Vec::new()));
    let mut writer = PackageWriter::with_writer(FailingSink {
        bytes: Rc::clone(&sink_bytes),
        remaining: 10,
    });
    let error = writer.set_mimetype_streaming(MIME).unwrap_err();
    match error {
        PackageWriterError::IncompleteOutput { written, source } => {
            assert_eq!(written, sink_bytes.borrow().len() as u64);
            assert!(matches!(source.as_ref(), PackageWriterError::Archive(_)));
            assert!(std::error::Error::source(source.as_ref()).is_some());
        },
        other => panic!("expected typed MIME sink failure, got {other:?}"),
    }
}
