//! Focused tests for reusing an already validated source XML proof at the
//! OPC topology-publication boundary.
//!
//! The tests intentionally live beside `SourceXmlPart` so they can exercise
//! the private proof fields and call `check_for_publication` directly. Public
//! topology tests continue to cover archive/security/output behavior; this
//! module isolates the XML-proof decision and its accounting contract.

#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used,
    reason = "focused publication-proof tests deliberately panic on fixture errors"
)]

use super::*;
use crate::source_backed::SourceBackedPackage;
use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, Resource, SourceVersion,
};
use soapberry_zip::office::StreamingArchiveWriter;
use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, RwLock,
    atomic::{AtomicU64, Ordering},
};

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const DOCUMENT_URI: &str = "/word/document.xml";
const CONTENT_TYPE: &str = "application/xml";

const SOURCE_XML: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<document>
  <before/>
  <nested>
    <leaf/>
  </nested>
</document>
"#;

fn document_uri() -> PackURI {
    PackURI::new(DOCUMENT_URI).expect("fixture URI must be valid")
}

fn archive_bytes(document: &[u8]) -> Vec<u8> {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="{CONTENT_TYPE}"/></Types>"#
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

fn managed_context(
    scope: &'static str,
    memory: u64,
    work: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        scope,
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker limit is nonzero"),
        NonZeroUsize::new(1).expect("one operation limit is nonzero"),
        NonZeroU64::new(memory).expect("memory limit is nonzero"),
        0,
    )
    .expect("execution limits must be valid");
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn managed_package(
    document: &[u8],
    limits: ReadLimits,
    scope: &'static str,
) -> (SourceBackedPackage, Budget, CancellationSource) {
    let (budget, cancellation_source, context) = managed_context(scope, 128 * 1024 * 1024, 1 << 40);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive_bytes(document))),
        limits,
        context,
    )
    .expect("managed fixture package must open");
    (package, budget, cancellation_source)
}

fn capture(package: &SourceBackedPackage) -> SourceXmlPart {
    package
        .part(&document_uri())
        .expect("document must exist")
        .source_xml()
        .expect("source XML must be valid")
}

fn make_derived(source: SourceXmlPart) -> SourceXmlPart {
    let insertion = source
        .bytes()
        .windows(b"</document>".len())
        .position(|window| window == b"</document>")
        .expect("fixture close tag must exist");
    let proof = source
        .checked_range(insertion..insertion, &[])
        .expect("insertion proof must be valid");
    let mut publication = source.into_publication().expect("publication must start");
    publication
        .replace(
            proof,
            AuthoredXmlFragment::markup(b"<added/>".to_vec())
                .expect("added XML fragment must be compact-authorized"),
        )
        .expect("source insertion must be accepted");
    publication.finish().expect("derived source must be valid")
}

fn assert_work_charge(budget: &Budget, before: u64, bytes: &[u8]) {
    assert_eq!(
        budget.used(Resource::Work) - before,
        bytes.len() as u64,
        "one publication proof check must consume exactly one payload-sized Work charge"
    );
}

#[test]
fn equal_original_and_derived_proofs_cover_replacement_and_addition() {
    let (package, budget, _cancellation) =
        managed_package(SOURCE_XML, ReadLimits::default(), "proof-equal");
    let original = capture(&package);
    let source_partname = original.source_partname.clone();
    let source_content_type = original.source_content_type.clone();
    let destination_original = original.original.as_bytes();

    // Existing replacement: destination identity and original bytes are both
    // checked before the publication proof is reused.
    let before = budget.used(Resource::Work);
    original
        .check_for_replacement(
            &original.source,
            &source_partname,
            source_content_type.as_str(),
            destination_original,
            original.limits,
        )
        .expect("an original source proof must authorize equal-limit replacement");
    assert_work_charge(&budget, before, original.bytes());

    // A finished source splice has independently validated its assembled
    // payload. Equal destination limits may reuse that proof too.
    let derived = make_derived(original.clone());
    assert_ne!(derived.bytes(), original.bytes());
    let before = budget.used(Resource::Work);
    derived
        .check_for_replacement(
            &derived.source,
            &source_partname,
            source_content_type.as_str(),
            destination_original,
            derived.limits,
        )
        .expect("a validated derived proof must authorize equal-limit replacement");
    assert_work_charge(&budget, before, derived.bytes());

    // Source additions have no destination-original comparison, but the
    // proof and destination policy still meet at check_for_publication.
    let before = budget.used(Resource::Work);
    original
        .check_for_publication(source_content_type.as_str(), original.limits)
        .expect("an original source proof must authorize equal-limit addition");
    assert_work_charge(&budget, before, original.bytes());

    drop(derived);
    drop(original);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn equal_limit_proof_hit_needs_no_parser_memory_but_keeps_work_bounded() {
    let (package, budget, _cancellation) =
        managed_package(SOURCE_XML, ReadLimits::default(), "proof-memory");
    let source = capture(&package);
    let payload_len = source.bytes().len() as u64;
    let before = budget.used(Resource::Work);

    // Saturate every remaining managed-memory byte. The equal-policy branch
    // is expected to charge Work and finish without allocating the temporary
    // quick-xml validation workspace.
    let remaining = budget
        .limit(Resource::Memory)
        .saturating_sub(budget.used(Resource::Memory));
    assert!(
        remaining > 0,
        "capture must leave memory for the saturation guard"
    );
    let memory_guard = budget
        .reserve(Resource::Memory, remaining)
        .expect("the saturation reservation must fit exactly");

    source
        .check_for_publication(source.content_type(), source.limits)
        .expect("equal limits must reuse the retained XML proof at the memory ceiling");
    assert_eq!(
        budget.used(Resource::Work) - before,
        payload_len,
        "proof reuse retains one bounded payload-sized Work charge"
    );
    drop(memory_guard);
    drop(source);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn different_limits_still_require_the_validator_memory_reservation() {
    let (package, budget, _cancellation) =
        managed_package(SOURCE_XML, ReadLimits::default(), "proof-memory-miss");
    let source = capture(&package);
    let changed = ReadLimits::builder()
        .max_parts(source.limits.max_parts() - 1)
        .expect("changed non-XML limit must be valid")
        .build()
        .expect("changed limits must build");
    assert_ne!(changed, source.limits);
    let remaining = budget
        .limit(Resource::Memory)
        .saturating_sub(budget.used(Resource::Memory));
    assert!(remaining > 0);
    let memory_guard = budget
        .reserve(Resource::Memory, remaining)
        .expect("the saturation reservation must fit exactly");
    let before = budget.used(Resource::Work);
    let error = source
        .check_for_publication(source.content_type(), changed)
        .expect_err("a policy miss must retain the full validator path");
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Memory
    ));
    assert_eq!(
        budget.used(Resource::Work) - before,
        source.bytes().len() as u64,
        "the full validator retains one payload-sized Work charge"
    );
    drop(memory_guard);
    drop(source);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn lower_xml_limits_take_the_full_validator_path() {
    let (package, budget, _cancellation) =
        managed_package(SOURCE_XML, ReadLimits::default(), "proof-limits");
    let source = capture(&package);
    let lower = ReadLimits::builder()
        .max_xml_depth(1)
        .expect("lower depth limit must be valid")
        .build()
        .expect("lower limits must build");
    assert_ne!(lower, source.limits);
    let before = budget.used(Resource::Work);
    let error = source
        .check_for_publication(source.content_type(), lower)
        .expect_err("a nested source must fail the lower destination depth limit");
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::XmlDepth,
            ..
        }
    ));
    assert_eq!(
        budget.used(Resource::Work) - before,
        source.bytes().len() as u64,
        "a lower-limit miss retains the full validator's one payload-sized Work charge"
    );
    drop(source);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn proof_work_refusal_is_typed_and_does_not_partially_charge() {
    let (package, budget, _cancellation) =
        managed_package(SOURCE_XML, ReadLimits::default(), "proof-work-limit");
    let source = capture(&package);
    let payload_len = source.bytes().len() as u64;
    let remaining = budget
        .limit(Resource::Work)
        .saturating_sub(budget.used(Resource::Work));
    assert!(remaining > payload_len);
    budget
        .consume(
            Resource::Work,
            remaining.saturating_sub(payload_len).saturating_add(1),
        )
        .expect("the budget must leave less than one proof Work charge");
    let before = budget.used(Resource::Work);
    let error = source
        .check_for_publication(source.content_type(), source.limits)
        .expect_err("proof Work charge must respect the cumulative budget");
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Work
    ));
    assert_eq!(budget.used(Resource::Work), before);
    drop(source);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn content_type_and_lineage_guards_precede_proof_reuse() {
    let (package, budget, _cancellation) =
        managed_package(SOURCE_XML, ReadLimits::default(), "proof-guards");
    let source = capture(&package);
    let before = budget.used(Resource::Work);
    let content_error = source
        .check_for_publication("application/other+xml", source.limits)
        .expect_err("a different content type must remain refused");
    assert!(matches!(
        content_error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
    assert_eq!(budget.used(Resource::Work), before);

    let mut foreign = source.clone();
    foreign.source_lineage = SourceBackedPackage::from_vec(archive_bytes(SOURCE_XML))
        .expect("foreign package must open")
        .source_lineage();
    let lineage_error = foreign
        .check_for_publication(foreign.content_type(), foreign.limits)
        .expect_err("an inconsistent source lineage must remain refused");
    assert!(matches!(
        lineage_error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));

    drop(foreign);
    drop(source);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[derive(Debug)]
struct MutableSource {
    bytes: RwLock<Vec<u8>>,
    revision: AtomicU64,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: RwLock::new(bytes),
            revision: AtomicU64::new(0),
        }
    }

    fn replace_bytes(&self, bytes: Vec<u8>) {
        *self.bytes.write().expect("mutable source lock must work") = bytes;
        self.revision.fetch_add(1, Ordering::Release);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self
            .bytes
            .read()
            .expect("mutable source lock must work")
            .len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let bytes = self.bytes.read().expect("mutable source lock must work");
        let start = usize::try_from(offset).map_err(|_| io::Error::other("offset overflow"))?;
        if start >= bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - start);
        output[..count].copy_from_slice(&bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5052_4f4f_465f_5354,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[test]
fn source_revision_invalidates_a_previously_valid_publication_proof() {
    let source = Arc::new(MutableSource::new(archive_bytes(SOURCE_XML)));
    let package = SourceBackedPackage::from_read_at(source.clone())
        .expect("mutable fixture package must open");
    let proof = capture(&package);
    source.replace_bytes(archive_bytes(SOURCE_XML));
    let error = proof
        .check_for_publication(proof.content_type(), proof.limits)
        .expect_err("source revision must invalidate the retained proof");
    assert!(matches!(error, OpcError::SourceChanged { .. }));
}

#[test]
fn destination_identity_is_checked_for_replacements() {
    let first =
        SourceBackedPackage::from_vec(archive_bytes(SOURCE_XML)).expect("first package must open");
    let second =
        SourceBackedPackage::from_vec(archive_bytes(SOURCE_XML)).expect("second package must open");
    let first_proof = capture(&first);
    let second_proof = capture(&second);
    let error = first_proof
        .check_for_replacement(
            &second_proof.source,
            &second_proof.source_partname,
            second_proof.content_type(),
            second_proof.original.as_bytes(),
            second_proof.limits,
        )
        .expect_err("a foreign replacement destination must be refused");
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
}
