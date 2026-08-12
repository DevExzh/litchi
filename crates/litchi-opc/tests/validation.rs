#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use std::io;
use std::sync::Arc;
use std::sync::atomic::{AtomicU64, AtomicUsize, Ordering};

use litchi_core::{CheckStatus, OwnedSource, ReadAt, SourceVersion, ValidationLimits};
use litchi_opc::{OpcError, ReadLimits, validate_read_at, validate_read_at_with_limits};

const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
const ROOT_RELS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

fn package_bytes(content_types: &[u8], relationships: &[u8], payload: &[u8]) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types)
        .unwrap();
    writer.write_stored("_rels/.rels", relationships).unwrap();
    writer.write_stored("word/document.xml", payload).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn status<'a>(report: &'a litchi_core::ValidateReport, id: &str) -> &'a CheckStatus {
    report
        .checks()
        .iter()
        .find(|check| check.id().as_str() == id)
        .expect("declared validation capability")
        .status()
}

fn phase_vector(report: &litchi_core::ValidateReport) -> Vec<String> {
    [
        "opc.package.ingress",
        "opc.package.catalog",
        "opc.package.reachable_relationship_graph",
        "opc.package.signature_presence",
    ]
    .into_iter()
    .map(|capability| match status(report, capability) {
        CheckStatus::Complete => "complete".to_owned(),
        CheckStatus::NotApplicable => "not_applicable".to_owned(),
        CheckStatus::Blocked { .. } => "blocked".to_owned(),
        CheckStatus::StoppedBy { check } => format!("stopped_by:{}", check.as_str()),
        _ => "unknown".to_owned(),
    })
    .collect()
}

#[test]
fn valid_catalog_is_complete_deterministic_and_does_not_read_payload() {
    let payload = vec![0xa5; 2 * 1024 * 1024];
    let bytes = package_bytes(CONTENT_TYPES, ROOT_RELS, &payload);
    let source = Arc::new(CountingSource::new(bytes.clone()));

    let first = validate_read_at(source.clone()).unwrap();
    let second = validate_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();

    assert!(first.is_complete());
    assert!(!first.has_errors());
    assert_eq!(first, second);
    assert!(matches!(
        status(&first, "opc.package.signature_presence"),
        CheckStatus::NotApplicable
    ));
    assert!(source.bytes_read() < payload.len() / 8);
    assert!(source.read_calls() < 128);
}

#[test]
fn malformed_content_types_and_relationships_are_conclusive_issues() {
    for (bytes, expected_code, expected_check, expected_phases) in [
        (
            package_bytes(b"<Types>", ROOT_RELS, b"document"),
            "opc.content_types.invalid",
            "opc.package.catalog",
            ["complete", "complete", "blocked", "blocked"],
        ),
        (
            package_bytes(CONTENT_TYPES, b"<Relationships>", b"document"),
            "opc.relationships.invalid",
            "opc.package.reachable_relationship_graph",
            ["complete", "blocked", "complete", "blocked"],
        ),
    ] {
        let report = validate_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
        assert!(!report.is_complete());
        assert!(report.has_errors());
        assert!(matches!(
            status(&report, "opc.package.ingress"),
            CheckStatus::Complete
        ));
        assert_eq!(report.issues().len(), 1);
        assert_eq!(report.issues()[0].code(), expected_code);
        assert_eq!(report.issues()[0].check().as_str(), expected_check);
        assert_eq!(phase_vector(&report), expected_phases);
    }
}

#[test]
fn duplicate_physical_entries_are_reported_as_structural_rejection() {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", CONTENT_TYPES)
        .unwrap();
    writer.write_stored("_rels/.rels", ROOT_RELS).unwrap();
    writer.write_stored("word/document.xml", b"one").unwrap();
    writer.write_stored("word/document.xml", b"two").unwrap();
    let bytes = writer.finish_to_bytes().unwrap();

    let report = validate_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    assert!(report.has_errors());
    assert!(matches!(
        status(&report, "opc.package.ingress"),
        CheckStatus::Complete
    ));
    assert_eq!(
        phase_vector(&report),
        ["complete", "blocked", "blocked", "blocked",]
    );
}

#[test]
fn zip_and_relationship_ceilings_are_blocked_without_issues() {
    let bytes = package_bytes(CONTENT_TYPES, ROOT_RELS, b"document");
    let input_limits = ReadLimits::builder()
        .max_input_bytes((bytes.len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    let input = validate_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes.clone())),
        input_limits,
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(input.issues().is_empty());
    assert!(matches!(
        status(&input, "opc.package.ingress"),
        CheckStatus::Blocked { .. }
    ));
    assert_eq!(
        phase_vector(&input),
        [
            "blocked",
            "stopped_by:opc.package.ingress",
            "stopped_by:opc.package.ingress",
            "blocked",
        ]
    );

    let archive_limits = ReadLimits::builder()
        .max_archive_members(1)
        .unwrap()
        .max_relationship_parts(1)
        .unwrap()
        .max_parts(1)
        .unwrap()
        .build()
        .unwrap();
    let archive = validate_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes.clone())),
        archive_limits,
        ValidationLimits::default(),
    )
    .unwrap();
    assert_eq!(
        phase_vector(&archive),
        [
            "blocked",
            "stopped_by:opc.package.ingress",
            "stopped_by:opc.package.ingress",
            "blocked",
        ]
    );

    let declared_size_limits = ReadLimits::builder()
        .max_archive_total_bytes(32)
        .unwrap()
        .build()
        .unwrap();
    let declared_size = validate_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes.clone())),
        declared_size_limits,
        ValidationLimits::default(),
    )
    .unwrap();
    assert_eq!(
        phase_vector(&declared_size),
        [
            "blocked",
            "stopped_by:opc.package.ingress",
            "stopped_by:opc.package.ingress",
            "blocked",
        ]
    );

    let catalog_limits = ReadLimits::builder()
        .max_content_types_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    let catalog = validate_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes.clone())),
        catalog_limits,
        ValidationLimits::default(),
    )
    .unwrap();
    assert_eq!(
        phase_vector(&catalog),
        [
            "complete",
            "blocked",
            "stopped_by:opc.package.catalog",
            "blocked",
        ]
    );

    let relationship_limits = ReadLimits::builder()
        .max_relationship_xml_bytes(16)
        .unwrap()
        .build()
        .unwrap();
    let relationships = validate_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        relationship_limits,
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(relationships.issues().is_empty());
    assert!(matches!(
        status(&relationships, "opc.package.reachable_relationship_graph"),
        CheckStatus::Blocked { .. }
    ));
    assert_eq!(
        phase_vector(&relationships),
        ["complete", "blocked", "blocked", "blocked",]
    );
}

#[test]
fn report_limits_are_fallible_errors_not_unbounded_retention() {
    let bytes = package_bytes(CONTENT_TYPES, ROOT_RELS, b"document");
    let tiny = ValidationLimits::new(1, 1, 1, 1, 8, 8, 8, 8, 8, 32);
    assert!(matches!(
        validate_read_at_with_limits(
            Arc::new(OwnedSource::new(bytes)),
            ReadLimits::default(),
            tiny,
        ),
        Err(OpcError::ValidationReport(_))
    ));
}

#[test]
fn signature_presence_is_information_not_signature_validity() {
    let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/_xmlsignatures/origin.sigs" ContentType="application/vnd.openxmlformats-package.digital-signature-origin"/></Types>"#;
    let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="_xmlsignatures/origin.sigs"/></Relationships>"#;
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types)
        .unwrap();
    writer.write_stored("_rels/.rels", relationships).unwrap();
    writer
        .write_stored("word/document.xml", b"document")
        .unwrap();
    writer
        .write_stored("_xmlsignatures/origin.sigs", b"not a signature")
        .unwrap();
    let report = validate_read_at(Arc::new(OwnedSource::new(
        writer.finish_to_bytes().unwrap(),
    )))
    .unwrap();

    assert!(report.is_complete());
    assert!(!report.has_errors());
    assert_eq!(report.issues().len(), 1);
    assert_eq!(
        report.issues()[0].code(),
        "opc.signature.infrastructure_present"
    );
    assert!(
        report.issues()[0]
            .message()
            .contains("validity was not checked")
    );
}

#[test]
fn macro_payload_is_inert_and_not_claimed_as_a_validation_capability() {
    let content_types = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>"#;
    let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="word/vbaProject.bin"/></Relationships>"#;
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types)
        .unwrap();
    writer.write_stored("_rels/.rels", relationships).unwrap();
    writer
        .write_stored("word/document.xml", b"document")
        .unwrap();
    writer
        .write_stored("word/vbaProject.bin", b"inert")
        .unwrap();
    let report = validate_read_at(Arc::new(OwnedSource::new(
        writer.finish_to_bytes().unwrap(),
    )))
    .unwrap();

    assert!(report.is_complete());
    assert!(
        report
            .checks()
            .iter()
            .all(|check| !check.id().as_str().contains("macro"))
    );
}

#[test]
fn io_and_source_change_failures_remain_errors() {
    let bytes = package_bytes(CONTENT_TYPES, ROOT_RELS, b"document");
    let io_error = validate_read_at(Arc::new(FailingSource::new(bytes.clone()))).unwrap_err();
    assert!(matches!(
        io_error,
        OpcError::IoError(error) if error.kind() == io::ErrorKind::BrokenPipe
    ));

    let stale = validate_read_at(Arc::new(MutatingSource::new(bytes))).unwrap_err();
    assert!(matches!(stale, OpcError::SourceChanged { .. }));

    let blocked_bytes = package_bytes(CONTENT_TYPES, ROOT_RELS, b"document");
    let blocked_limits = ReadLimits::builder()
        .max_input_bytes((blocked_bytes.len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    let blocked_stale = validate_read_at_with_limits(
        Arc::new(LengthMutatingSource::new(blocked_bytes)),
        blocked_limits,
        ValidationLimits::default(),
    )
    .unwrap_err();
    assert!(matches!(blocked_stale, OpcError::SourceChanged { .. }));

    let structural_stale =
        validate_read_at(Arc::new(MutatingSource::new(b"not a zip".to_vec()))).unwrap_err();
    assert!(matches!(structural_stale, OpcError::SourceChanged { .. }));
}

struct CountingSource {
    bytes: Vec<u8>,
    reads: AtomicUsize,
    read_bytes: AtomicUsize,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            reads: AtomicUsize::new(0),
            read_bytes: AtomicUsize::new(0),
        }
    }

    fn read_calls(&self) -> usize {
        self.reads.load(Ordering::SeqCst)
    }

    fn bytes_read(&self) -> usize {
        self.read_bytes.load(Ordering::SeqCst)
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let source = self.bytes.get(start..).unwrap_or_default();
        let count = source.len().min(output.len());
        output[..count].copy_from_slice(&source[..count]);
        self.read_bytes.fetch_add(count, Ordering::SeqCst);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(700, 0))
    }
}

struct FailingSource {
    bytes: Vec<u8>,
}

impl FailingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self { bytes }
    }
}

impl ReadAt for FailingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, _offset: u64, _output: &mut [u8]) -> io::Result<usize> {
        Err(io::Error::new(
            io::ErrorKind::BrokenPipe,
            "injected read failure",
        ))
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(701, 0))
    }
}

struct MutatingSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
}

struct LengthMutatingSource {
    bytes: Vec<u8>,
    lengths: AtomicUsize,
}

impl LengthMutatingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            lengths: AtomicUsize::new(0),
        }
    }
}

impl ReadAt for LengthMutatingSource {
    fn len(&self) -> io::Result<u64> {
        self.lengths.fetch_add(1, Ordering::SeqCst);
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let source = self.bytes.get(start..).unwrap_or_default();
        let count = source.len().min(output.len());
        output[..count].copy_from_slice(&source[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            703,
            u64::from(self.lengths.load(Ordering::SeqCst) > 0),
        ))
    }
}

impl MutatingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
        }
    }
}

impl ReadAt for MutatingSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let source = self.bytes.get(start..).unwrap_or_default();
        let count = source.len().min(output.len());
        output[..count].copy_from_slice(&source[..count]);
        self.revision.store(1, Ordering::SeqCst);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            702,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}
