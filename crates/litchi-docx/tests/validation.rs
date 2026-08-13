#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "validation assertions intentionally panic on fixture failure"
)]

use std::{
    io,
    sync::{
        Arc,
        atomic::{AtomicUsize, Ordering},
    },
};

use litchi_core::{
    CheckStatus, EvidenceValue, OwnedSource, ReadAt, SourceVersion, ValidationLimits,
};
use litchi_docx::{
    DocxValidationError, DocxValidationLimits, source_backed, validate_read_at,
    validate_read_at_with_limits,
};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{BlobPart, OpcError, OpcPackage, PackURI, PackageWriter, Part};
use soapberry_zip::office::StreamingArchiveWriter;

const WORD_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn fixture(main_xml: &str, content_type: &str, external: bool) -> Vec<u8> {
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        content_type.to_owned(),
        main_xml.as_bytes().to_vec(),
    );
    if external {
        main.relate_to_ext("https://example.invalid/not-fetched", rt::HYPERLINK);
    }
    package.try_add_part(Box::new(main)).unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn malformed_fixture(main_xml: &str) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{}"/></Types>"#,
        ct::WML_DOCUMENT_MAIN
    );
    let relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="{}" Target="word/document.xml"/></Relationships>"#,
        rt::OFFICE_DOCUMENT
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .unwrap();
    writer
        .write_stored("word/document.xml", main_xml.as_bytes())
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn relationship_fixture() -> Vec<u8> {
    let main_xml =
        format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p/></w:body></w:document>"#);
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MAIN.to_owned(),
        main_xml.into_bytes(),
    );
    main.relate_to("missing.xml", rt::HEADER);
    main.relate_to("wrong.bin", rt::HEADER);
    // A vendor URI ending in `/image` is not the OOXML image relationship.
    // It must remain unknown rather than inheriting suffix-based semantics.
    main.relate_to("wrong.bin", "https://vendor.invalid/image");
    main.relate_to("header1.xml", rt::HEADER);
    package.try_add_part(Box::new(main)).unwrap();
    package
        .try_add_part(Box::new(BlobPart::new(
            PackURI::new("/word/wrong.bin").unwrap(),
            ct::PNG.to_owned(),
            vec![0x89, b'P', b'N', b'G'],
        )))
        .unwrap();
    let mut header = BlobPart::new(
        PackURI::new("/word/header1.xml").unwrap(),
        ct::WML_HEADER.to_owned(),
        b"<w:hdr/>".to_vec(),
    );
    header.relate_to("missing-image.png", rt::IMAGE);
    package.try_add_part(Box::new(header)).unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn macro_and_activex_fixture() -> Vec<u8> {
    let main_xml =
        format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p/></w:body></w:document>"#);
    let mut package = OpcPackage::new();
    let mut main = BlobPart::new(
        PackURI::new("/word/document.xml").unwrap(),
        ct::WML_DOCUMENT_MACRO_MAIN.to_owned(),
        main_xml.into_bytes(),
    );
    main.relate_to(
        "missing-control.bin",
        "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary",
    );
    package.try_add_part(Box::new(main)).unwrap();
    package.relate_to("word/document.xml", rt::OFFICE_DOCUMENT);
    PackageWriter::to_bytes(&package).unwrap()
}

fn valid_fixture() -> Vec<u8> {
    fixture(
        &format!(
            r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p><w:r><w:t>alpha</w:t></w:r></w:p><w:tbl/><w:sectPr/></w:body></w:document>"#
        ),
        ct::WML_DOCUMENT_MAIN,
        false,
    )
}

fn status<'a>(report: &'a litchi_core::ValidateReport, id: &str) -> &'a CheckStatus {
    report
        .checks()
        .iter()
        .find(|check| check.id().as_str() == id)
        .expect("declared DOCX validation capability")
        .status()
}

fn count_evidence(report: &litchi_core::ValidateReport, code: &str, key: &str) -> Option<u64> {
    report
        .issues()
        .iter()
        .find(|issue| issue.code() == code)
        .and_then(|issue| {
            issue
                .evidence()
                .iter()
                .find(|evidence| evidence.key() == key)
        })
        .and_then(|evidence| match evidence.value() {
            EvidenceValue::Count(value) => Some(value),
            _ => None,
        })
}

fn has_issue(report: &litchi_core::ValidateReport, code: &str) -> bool {
    report.issues().iter().any(|issue| issue.code() == code)
}

#[derive(Clone)]
struct TestSource {
    bytes: Arc<Vec<u8>>,
    fail_all: bool,
    fail_range: Option<(u64, u64)>,
    payload_reads: Option<Arc<AtomicUsize>>,
}

impl TestSource {
    fn failing(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Arc::new(bytes),
            fail_all: true,
            fail_range: None,
            payload_reads: None,
        }
    }

    fn counted_payload(bytes: Vec<u8>, payload: &[u8]) -> Self {
        let start = bytes
            .windows(payload.len())
            .position(|window| window == payload)
            .expect("stored payload is present in the archive");
        let start = u64::try_from(start).unwrap();
        let length = u64::try_from(payload.len()).unwrap();
        Self {
            bytes: Arc::new(bytes),
            fail_all: false,
            fail_range: Some((start, start.saturating_add(length))),
            payload_reads: Some(Arc::new(AtomicUsize::new(0))),
        }
    }

    fn payload_read_count(&self) -> usize {
        self.payload_reads
            .as_ref()
            .map_or(0, |count| count.load(Ordering::Acquire))
    }
}

impl ReadAt for TestSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("test source length does not fit u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if output.is_empty() {
            return Ok(0);
        }
        let end = offset.saturating_add(u64::try_from(output.len()).unwrap());
        let overlaps_payload = self
            .fail_range
            .is_some_and(|(start, finish)| offset < finish && end > start);
        if overlaps_payload {
            if let Some(count) = &self.payload_reads {
                count.fetch_add(1, Ordering::AcqRel);
            }
        }
        if self.fail_all || overlaps_payload {
            return Err(io::Error::other("test positional source failure"));
        }
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "test offset overflow"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let length = output.len().min(self.bytes.len().saturating_sub(start));
        output[..length].copy_from_slice(&self.bytes[start..start + length]);
        Ok(length)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0xD0C0_0098, 0))
    }
}

#[test]
fn valid_main_document_is_complete_without_retaining_content() {
    let bytes = valid_fixture();
    let original = bytes.clone();
    let report = validate_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();

    assert!(report.is_complete());
    assert!(!report.has_errors());
    assert_eq!(
        status(&report, "docx.package.ingress"),
        &CheckStatus::Complete
    );
    assert_eq!(
        status(&report, "docx.main_document.markup_compatibility"),
        &CheckStatus::NotApplicable
    );
    assert_eq!(
        status(&report, "docx.main_document.semantics"),
        &CheckStatus::Complete
    );
    assert_eq!(
        status(&report, "docx.security.macro_presence"),
        &CheckStatus::NotApplicable
    );
    assert_eq!(
        status(&report, "docx.relationships.external_target_presence"),
        &CheckStatus::NotApplicable
    );
    assert!(!original.is_empty());
}

#[test]
fn malformed_and_unsupported_main_semantics_are_content_free_findings() {
    let malformed = malformed_fixture(&format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p>"#));
    let malformed_report = validate_read_at(Arc::new(OwnedSource::new(malformed))).unwrap();
    assert!(malformed_report.has_errors());
    assert_eq!(
        status(&malformed_report, "docx.main_document.markup_compatibility"),
        &CheckStatus::NotApplicable
    );
    assert_eq!(
        status(&malformed_report, "docx.main_document.semantics"),
        &CheckStatus::Complete
    );
    assert!(
        malformed_report
            .issues()
            .iter()
            .any(|issue| issue.code() == "docx.main_document.malformed_xml")
    );

    let unsupported = fixture(
        r#"<x:document xmlns:x="urn:unsupported"><x:body/></x:document>"#,
        ct::WML_DOCUMENT_MAIN,
        false,
    );
    let unsupported_report = validate_read_at(Arc::new(OwnedSource::new(unsupported))).unwrap();
    assert!(unsupported_report.has_errors());
    assert!(
        unsupported_report
            .issues()
            .iter()
            .any(|issue| issue.code() == "docx.main_document.unsupported_root")
    );
}

#[test]
fn security_presence_is_inert_and_external_targets_are_not_fetched() {
    let bytes = fixture(
        &format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p/></w:body></w:document>"#),
        ct::WML_DOCUMENT_MACRO_MAIN,
        true,
    );
    let report = validate_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();

    assert_eq!(
        status(&report, "docx.security.macro_presence"),
        &CheckStatus::Complete
    );
    assert_eq!(
        status(&report, "docx.relationships.external_target_presence"),
        &CheckStatus::Complete
    );
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "docx.macro.storage_present")
    );
    assert!(
        report
            .issues()
            .iter()
            .any(|issue| issue.code() == "docx.external_target.present")
    );
}

#[test]
fn semantic_limits_block_without_materializing_an_unbounded_payload() {
    let bytes = valid_fixture();
    let report = validate_read_at_with_limits(
        Arc::new(OwnedSource::new(bytes)),
        litchi_opc::ReadLimits::default(),
        DocxValidationLimits::default().with_max_main_document_bytes(1),
        ValidationLimits::default(),
    )
    .unwrap();

    assert!(!report.is_complete());
    assert!(matches!(
        status(&report, "docx.main_document.markup_compatibility"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "docx.main_document.semantics"),
        CheckStatus::StoppedBy { check } if check.as_str() == "docx.main_document.markup_compatibility"
    ));
}

#[test]
fn package_source_backed_open_remains_available_for_follow_up_queries() {
    let bytes = valid_fixture();
    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(bytes))).unwrap();
    assert_eq!(package.document().unwrap().paragraph_count().unwrap(), 1);
}

#[test]
fn source_io_is_not_reclassified_as_a_structural_zip_finding() {
    let error = validate_read_at(Arc::new(TestSource::failing(valid_fixture()))).unwrap_err();
    assert!(matches!(
        error,
        DocxValidationError::Ingress(OpcError::IoError(error))
            if error.kind() == io::ErrorKind::Other
    ));
}

#[test]
fn main_part_io_is_reported_as_ingress_failure_after_catalog_open() {
    let main_xml =
        format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p/></w:body></w:document>"#);
    let bytes = malformed_fixture(&main_xml);
    let source = TestSource::counted_payload(bytes, main_xml.as_bytes());
    let error = validate_read_at(Arc::new(source)).unwrap_err();
    assert!(matches!(
        error,
        DocxValidationError::Ingress(OpcError::IoError(error))
            if error.kind() == io::ErrorKind::Other
    ));
}

#[test]
fn declared_main_size_is_rejected_before_payload_read() {
    let main_xml =
        format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p/></w:body></w:document>"#);
    let bytes = malformed_fixture(&main_xml);
    let source = TestSource::counted_payload(bytes, main_xml.as_bytes());
    let report = validate_read_at_with_limits(
        Arc::new(source.clone()),
        litchi_opc::ReadLimits::default(),
        DocxValidationLimits::default()
            .with_max_main_document_bytes(u64::try_from(main_xml.len() - 1).unwrap()),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "docx.main_document.markup_compatibility"),
        CheckStatus::Blocked { .. }
    ));
    assert_eq!(source.payload_read_count(), 0);
}

#[test]
fn visible_xml_rejects_special_events_invalid_references_and_extra_roots() {
    let cases = [
        format!(r#"<!DOCTYPE w:document><w:document xmlns:w="{WORD_NS}"><w:body/></w:document>"#),
        format!(r#"<?not-supported?><w:document xmlns:w="{WORD_NS}"><w:body/></w:document>"#),
        format!(r#"<?xml version="1.2"?><w:document xmlns:w="{WORD_NS}"><w:body/></w:document>"#),
        format!(
            r#"<?xml version="1.0" standalone="maybe"?><w:document xmlns:w="{WORD_NS}"><w:body/></w:document>"#
        ),
        format!(
            r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p>&unknown;</w:p></w:body></w:document>"#
        ),
        format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p>&#0;</w:p></w:body></w:document>"#),
        format!(r#"<w:document xmlns:w="{WORD_NS}" bad="&unknown;"><w:body/></w:document>"#),
        format!(r#"<w:document xmlns:w="{WORD_NS}" bad="&#0;"><w:body/></w:document>"#),
        format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body/></w:document>outside"#),
        format!(
            r#"<w:document xmlns:w="{WORD_NS}"><w:body/></w:document><w:document xmlns:w="{WORD_NS}"/>"#
        ),
        format!(r#"<w:document xmlns:w="{WORD_NS}"><x:p/></w:document>"#),
        format!(r#"<w:document xmlns:w="{WORD_NS}" x:flag="1"><w:body/></w:document>"#),
        format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body/><w:body/></w:document>"#),
        format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p/></w:body><w:body/></w:document>"#),
    ];
    for xml in cases {
        let report = validate_read_at(Arc::new(OwnedSource::new(malformed_fixture(&xml)))).unwrap();
        assert!(
            has_issue(&report, "docx.main_document.malformed_xml"),
            "{xml}"
        );
        assert_eq!(
            status(&report, "docx.main_document.semantics"),
            &CheckStatus::Complete
        );
    }
}

#[test]
fn xml_depth_is_bounded_before_namespace_reader_and_public_limit_is_safe() {
    assert_eq!(
        DocxValidationLimits::default()
            .with_max_xml_depth(usize::MAX)
            .max_xml_depth(),
        u16::MAX as usize - 1
    );
    let xml = format!(r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p/></w:body></w:document>"#);
    let report = validate_read_at_with_limits(
        Arc::new(OwnedSource::new(fixture(
            &xml,
            ct::WML_DOCUMENT_MAIN,
            false,
        ))),
        litchi_opc::ReadLimits::default(),
        DocxValidationLimits::default().with_max_xml_depth(1),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "docx.main_document.markup_compatibility"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "docx.main_document.semantics"),
        CheckStatus::StoppedBy { check }
            if check.as_str() == "docx.main_document.markup_compatibility"
    ));
}

#[test]
fn predefined_and_numeric_references_remain_valid() {
    let xml = format!(
        r#"<w:document xmlns:w="{WORD_NS}"><w:body><w:p>&amp;&#x41;&#65;</w:p></w:body></w:document>"#
    );
    let report = validate_read_at(Arc::new(OwnedSource::new(fixture(
        &xml,
        ct::WML_DOCUMENT_MAIN,
        false,
    ))))
    .unwrap();
    assert!(!has_issue(&report, "docx.main_document.malformed_xml"));
    assert_eq!(
        status(&report, "docx.main_document.semantics"),
        &CheckStatus::Complete
    );
}

#[test]
fn mce_preprocessing_obeys_the_xml_event_ceiling() {
    let xml = format!(
        r#"<w:document xmlns:w="{WORD_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><w:body><w:p/><w:p/><w:p/></w:body></w:document>"#
    );
    let report = validate_read_at_with_limits(
        Arc::new(OwnedSource::new(fixture(
            &xml,
            ct::WML_DOCUMENT_MAIN,
            false,
        ))),
        litchi_opc::ReadLimits::default(),
        DocxValidationLimits::default().with_max_xml_events(2),
        ValidationLimits::default(),
    )
    .unwrap();
    assert!(matches!(
        status(&report, "docx.main_document.markup_compatibility"),
        CheckStatus::Blocked { .. }
    ));
    assert!(matches!(
        status(&report, "docx.main_document.semantics"),
        CheckStatus::StoppedBy { check } if check.as_str() == "docx.main_document.markup_compatibility"
    ));
}

#[test]
fn main_relationship_closure_reports_missing_and_incompatible_targets_at_main_uri() {
    let report = validate_read_at(Arc::new(OwnedSource::new(relationship_fixture()))).unwrap();
    let missing = report
        .issues()
        .iter()
        .find(|issue| issue.code() == "docx.main_document.relationship_target_missing")
        .expect("missing main target finding");
    assert_eq!(missing.locations()[0].path(), Some("/word/document.xml"));
    assert_eq!(
        count_evidence(
            &report,
            "docx.main_document.relationship_target_missing",
            "missing_internal_targets"
        ),
        Some(2)
    );
    assert_eq!(
        count_evidence(
            &report,
            "docx.main_document.relationship_target_content_type",
            "content_type_mismatches"
        ),
        Some(1)
    );
    assert_eq!(
        status(&report, "docx.main_document.relationship_closure"),
        &CheckStatus::Complete
    );
}

#[test]
fn macro_main_is_counted_once_and_dangling_activex_is_visible() {
    let report = validate_read_at(Arc::new(OwnedSource::new(macro_and_activex_fixture()))).unwrap();
    assert_eq!(
        count_evidence(&report, "docx.macro.storage_present", "observations"),
        Some(1)
    );
    assert_eq!(
        count_evidence(&report, "docx.embedded_content.present", "observations"),
        Some(1)
    );
}
