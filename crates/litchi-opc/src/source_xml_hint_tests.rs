//! Internal proof-field coverage for the source XML hint fast path.
//!
//! The public integration tests exercise the opaque API. These tests live
//! under `source_backed` so they can independently alter one proof field at a
//! time and show that every field in the admission predicate is authoritative.

#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used,
    reason = "focused internal proof tests deliberately panic on fixture errors"
)]

use super::*;
use crate::content_type::ContentType;
use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits, SourceVersion};
use soapberry_zip::office::StreamingArchiveWriter;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Arc;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const DOCUMENT_URI: &str = "/word/document.xml";
const SOURCE: &[u8] = b"<document><current/></document>";
const OTHER: &[u8] = b"<document><changed/></document>";

fn document_uri() -> PackURI {
    PackURI::new(DOCUMENT_URI).expect("fixture URI must be valid")
}

fn archive_bytes(document: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
    );
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="{DOCUMENT_MEMBER}"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .expect("relationships fixture must be writable");
    writer
        .write_stored(DOCUMENT_MEMBER, document)
        .expect("document fixture must be writable");
    writer
        .finish_to_bytes()
        .expect("archive fixture must finish")
}

fn open_package(document: &[u8]) -> SourceBackedPackage {
    SourceBackedPackage::from_vec(archive_bytes(document)).expect("fixture package must open")
}

fn source_parts(package: &SourceBackedPackage) -> (usize, SourceXmlPart) {
    let view = package.part(&document_uri()).expect("document must exist");
    let index = view.index;
    let source = package
        .source_xml_part(index)
        .expect("document source must be valid");
    (index, source)
}

fn current_from_hint(
    package: &SourceBackedPackage,
    index: usize,
    hint: &SourceXmlPart,
) -> SourceXmlPart {
    package
        .source_xml_part_with_hint(index, Some(hint))
        .expect("mismatched proof field must fall back to current source")
}

#[test]
fn each_identity_and_policy_field_is_required_for_a_hint_hit() {
    let package = open_package(SOURCE);
    let (index, original) = source_parts(&package);
    let expected_version = package.source.version();
    let expected_lineage = package.source.lineage.clone();
    let expected_limits = package.limits;

    let mut cases = Vec::new();

    let mut lineage = original.clone();
    lineage.source_lineage = SourceLineage(Arc::new(()));
    cases.push(("lineage", lineage));

    let mut version = original.clone();
    version.source_version = SourceVersion::new(0x4849_4e54, 99);
    cases.push(("version", version));

    let mut partname = original.clone();
    partname.source_partname = Arc::new(PackURI::new("/custom/other.xml").unwrap());
    cases.push(("Part URI", partname));

    let mut content_type = original.clone();
    content_type.source_content_type = Arc::new(ContentType::new("application/other+xml").unwrap());
    cases.push(("content type", content_type));

    let mut limits = original.clone();
    limits.limits = ReadLimits::builder()
        .max_xml_events(expected_limits.max_xml_events() - 1)
        .unwrap()
        .build()
        .unwrap();
    cases.push(("limits", limits));

    for (field, candidate) in cases {
        let current = current_from_hint(&package, index, &candidate);
        assert_eq!(
            current.bytes(),
            SOURCE,
            "mismatched {field} must return current bytes"
        );
        assert_eq!(current.source_version, expected_version);
        assert_eq!(current.source_lineage, expected_lineage);
        assert_eq!(current.limits, expected_limits);
        assert_eq!(current.partname(), &document_uri());
        assert_eq!(current.content_type(), "application/xml");
    }
}

#[test]
fn original_bytes_are_checked_after_all_metadata_matches() {
    let package = open_package(SOURCE);
    let (index, original) = source_parts(&package);
    let other_package = open_package(OTHER);
    let other_view = other_package.part(&document_uri()).unwrap();
    let wrong_original = other_view.data().unwrap();

    let mut candidate = original;
    // Keep the candidate in the original state, but replace both opaque
    // payload handles with another valid XML allocation. This reaches the
    // exact-byte comparison after every metadata predicate has matched.
    candidate.original = wrong_original.clone();
    candidate.payload = wrong_original.shared_bytes();
    assert_eq!(candidate.bytes(), OTHER);

    let current = current_from_hint(&package, index, &candidate);
    assert_eq!(current.bytes(), SOURCE);
    assert_eq!(current.partname(), &document_uri());
    assert_eq!(current.content_type(), "application/xml");
}

#[test]
fn equal_bytes_in_a_distinct_allocation_still_take_the_hint_hit() {
    let package = open_package(SOURCE);
    let other_package = open_package(SOURCE);
    let (index, original) = source_parts(&package);
    let other_data = other_package.part(&document_uri()).unwrap().data().unwrap();
    assert!(!original.original.shares_allocation_with(&other_data));

    let mut candidate = original.clone();
    candidate.original = other_data.clone();
    candidate.payload = other_data.shared_bytes();
    let current = current_from_hint(&package, index, &candidate);

    assert_eq!(current.bytes(), SOURCE);
    assert!(Arc::ptr_eq(&current.payload, &candidate.payload));
    assert!(Arc::ptr_eq(
        &current.source_partname,
        &candidate.source_partname
    ));
    assert!(Arc::ptr_eq(
        &current.source_content_type,
        &candidate.source_content_type
    ));
}

#[test]
fn eligible_hint_comparison_consumes_its_bounded_work_charge() {
    let archive = archive_bytes(SOURCE);
    let source_bytes = SOURCE.len() as u64;
    let budget = Budget::root(
        "source-xml-hint-work-limit",
        Limits::new(
            64 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            1 << 40,
        ),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let context = ExecutionContext::new(
        budget.clone(),
        cancellation,
        ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(64 * 1024 * 1024).unwrap(),
            0,
        )
        .unwrap(),
    );
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let part = package.part(&document_uri()).unwrap();
    let hint = part.source_xml().unwrap();
    let remaining = budget
        .limit(Resource::Work)
        .saturating_sub(budget.used(Resource::Work));
    assert!(remaining > source_bytes);
    budget
        .consume(
            Resource::Work,
            remaining.saturating_sub(source_bytes).saturating_add(1),
        )
        .expect("the budget must leave exactly less than one comparison charge");
    let error = part
        .source_xml_with_hint(&hint)
        .expect_err("the second bounded byte charge must exceed the work budget");
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Work
    ));
}
