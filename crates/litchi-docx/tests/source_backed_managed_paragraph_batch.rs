#![cfg(any(unix, windows))]

//! Managed source-backed coverage for the direct-body paragraph batch edit.
//!
//! These tests keep the source proof, the retained owners, and every finite
//! execution dimension visible. The scalar route is used only as an oracle
//! for the same semantic final state; the batch route remains the API under
//! test.

use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::atomic::{AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, Position, ReadAt, Resource, SourceVersion,
};
use litchi_docx::document::{ParagraphTextReplacement, Refusal, TransactionError};
use litchi_docx::{Error, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::{OpcError, OpcPackage, PackURI};
use soapberry_zip::office::StreamingArchiveWriter;

const MAIN: &str = "word/document.xml";
const MEDIA: &str = "word/media/image1.png";
const OPAQUE: &str = "word/opaque.bin";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const FINITE_MEMORY: u64 = 16 * 1024 * 1024;
const FINITE_INPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OBJECTS: u64 = 1_000_000;
const FINITE_DEPTH: u64 = 1024;
const FINITE_WORK: u64 = 1 << 30;

fn document_xml() -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p w:rsidR="1"><w:pPr><w:keepNext/></w:pPr><w:r><w:rPr><w:b/></w:rPr><w:t>first</w:t><w:custom w:value="opaque"/></w:r><w:r><w:t> tail</w:t></w:r></w:p><w:p><w:r><w:t>second</w:t></w:r></w:p><w:p><w:r><w:rPr><w:i/></w:rPr><w:t>third</w:t></w:r></w:p><w:p><w:r><w:t>fourth</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn archive_fixture(document: &[u8], signed: bool) -> Vec<u8> {
    let media: Vec<u8> = (0_u8..=u8::MAX)
        .cycle()
        .take(4096)
        .map(|byte| byte.wrapping_mul(37))
        .collect();
    let opaque: Vec<u8> = (0_u8..=u8::MAX)
        .cycle()
        .take(8192)
        .map(|byte| byte.wrapping_mul(13))
        .collect();
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="sigs" ContentType="{signature_origin}"/><Override PartName="/{MAIN}" ContentType="{document_content_type}"/></Types>"#,
        signature_origin = ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        document_content_type = ct::WML_DOCUMENT_MAIN,
    );
    let signature_relationship = if signed {
        format!(
            r#"<Relationship Id="rIdSignature" Type="{}" Target="_xmlsignatures/origin.sigs"/>"#,
            rt::DIGITAL_SIGNATURE_ORIGIN,
        )
    } else {
        String::new()
    };
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rIdDocument" Type="{}" Target="{MAIN}"/>{signature_relationship}</Relationships>"#,
        rt::OFFICE_DOCUMENT,
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .unwrap();
    writer.write_stored(MEDIA, &media).unwrap();
    writer.write_stored(OPAQUE, &opaque).unwrap();
    writer.write_stored(MAIN, document).unwrap();
    if signed {
        writer
            .write_stored("_xmlsignatures/origin.sigs", b"<origin/>")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn source_bytes() -> Vec<u8> {
    archive_fixture(&document_xml(), false)
}

fn signed_source_bytes() -> Vec<u8> {
    archive_fixture(&document_xml(), true)
}

fn replacements() -> Vec<ParagraphTextReplacement> {
    vec![
        ParagraphTextReplacement::new(Position::new(0), "first changed"),
        ParagraphTextReplacement::new(Position::new(2), "third changed"),
    ]
}

fn original_replacements() -> Vec<ParagraphTextReplacement> {
    vec![
        ParagraphTextReplacement::new(Position::new(0), "first tail"),
        ParagraphTextReplacement::new(Position::new(2), "third"),
    ]
}

fn managed_context(memory: u64, work: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-managed-paragraph-batch-test",
        Limits::new(
            memory,
            FINITE_INPUT_BYTES,
            FINITE_OUTPUT_BYTES,
            FINITE_OBJECTS,
            FINITE_DEPTH,
            work,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::MIN,
        NonZeroUsize::MIN,
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, limits),
    )
}

fn managed_with_work(
    bytes: Vec<u8>,
    memory: u64,
    work: u64,
) -> (Budget, CancellationSource, source_backed::Package) {
    let (budget, cancellation_source, context) = managed_context(memory, work);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    (budget, cancellation_source, package)
}

fn managed(bytes: Vec<u8>) -> (Budget, CancellationSource, source_backed::Package) {
    managed_with_work(bytes, FINITE_MEMORY, FINITE_WORK)
}

fn part_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .get_part(&PackURI::new(format!("/{name}")).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn is_cancelled(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(OpcError::Cancelled))
    )
}

fn is_resource_limit(error: &TransactionError, resource: Resource) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(OpcError::Execution(
            ExecutionError::ResourceLimit(limit),
        ))) if limit.resource == resource
    )
}

#[derive(Debug)]
struct MutableSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            revision: AtomicU64::new(0),
        }
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }
}

impl ReadAt for MutableSource {
    fn len(&self) -> io::Result<u64> {
        let bytes = self
            .bytes
            .lock()
            .map_err(|_| io::Error::other("mutable source mutex poisoned"))?;
        u64::try_from(bytes.len()).map_err(|_| io::Error::other("source length overflows u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let offset = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflows usize"))?;
        let bytes = self
            .bytes
            .lock()
            .map_err(|_| io::Error::other("mutable source mutex poisoned"))?;
        if offset >= bytes.len() || output.is_empty() {
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - offset);
        output[..count].copy_from_slice(&bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4d_50_42_41,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[test]
fn managed_batch_matches_repeated_scalar_xml_and_preserves_unknown_members() {
    let source = source_bytes();
    let replacements = replacements();

    let (batch_budget, _batch_cancel, batch_package) = managed(source.clone());
    let batch_snapshot = batch_package.document_snapshot().unwrap();
    let mut batch_edit = batch_snapshot.edit();
    batch_edit
        .replace_body_paragraph_texts(&replacements)
        .unwrap();
    let batch_commit = batch_edit.commit().unwrap();
    let batch_xml = batch_commit.snapshot().xml_bytes().to_vec();
    for replacement in &replacements {
        assert_eq!(
            batch_commit
                .snapshot()
                .paragraph(replacement.position())
                .unwrap()
                .text()
                .unwrap(),
            replacement.text()
        );
    }
    let mut batch_output = Vec::new();
    batch_package
        .publish_document_commit_to_stream(&mut batch_output, &batch_commit)
        .unwrap();
    drop(batch_commit);
    drop(batch_snapshot);
    assert_eq!(batch_budget.used(Resource::Memory), 0);
    assert_eq!(batch_budget.used(Resource::Objects), 0);

    let (scalar_budget, _scalar_cancel, scalar_package) = managed(source.clone());
    let scalar_snapshot = scalar_package.document_snapshot().unwrap();
    let mut scalar_edit = scalar_snapshot.edit();
    for replacement in &replacements {
        scalar_edit
            .replace_paragraph_text(replacement.position(), replacement.text())
            .unwrap();
    }
    let scalar_commit = scalar_edit.commit().unwrap();
    assert_eq!(batch_xml, scalar_commit.snapshot().xml_bytes());
    let mut scalar_output = Vec::new();
    scalar_package
        .publish_document_commit_to_stream(&mut scalar_output, &scalar_commit)
        .unwrap();
    assert_eq!(batch_output, scalar_output);
    assert_eq!(part_bytes(&batch_output, MEDIA), part_bytes(&source, MEDIA));
    assert_eq!(
        part_bytes(&batch_output, OPAQUE),
        part_bytes(&source, OPAQUE)
    );

    let changed = String::from_utf8(part_bytes(&batch_output, MAIN)).unwrap();
    assert!(changed.contains(r#"<w:rPr><w:b/></w:rPr>"#));
    assert!(changed.contains(r#"<w:custom w:value="opaque"/>"#));
    assert!(changed.contains(r#"<w:rPr><w:i/></w:rPr>"#));

    drop(scalar_commit);
    drop(scalar_snapshot);
    assert_eq!(scalar_budget.used(Resource::Memory), 0);
    assert_eq!(scalar_budget.used(Resource::Objects), 0);
}

#[test]
fn managed_batch_forward_and_inverse_are_source_checked() {
    let source = source_bytes();
    let (budget, _cancel, package) = managed(source.clone());
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.replace_body_paragraph_texts(&replacements()).unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    let publication = package
        .publish_document_commit_with_inverse_to_stream(&mut output, &commit)
        .unwrap();
    drop(commit);
    drop(snapshot);
    assert!(budget.used(Resource::Memory) > 0);

    let (inverse_budget, _inverse_cancel, reopened) = managed(output.clone());
    let mut restored = Vec::new();
    let restored_snapshot = reopened
        .publish_document_inverse_to_stream(&mut restored, &publication)
        .unwrap();
    assert_eq!(restored, source);
    assert_eq!(restored_snapshot.xml_bytes(), part_bytes(&source, MAIN));
    drop(restored_snapshot);
    assert_eq!(inverse_budget.used(Resource::Memory), 0);
    drop(publication);
    assert_eq!(budget.used(Resource::Memory), 0);

    let mutable = Arc::new(MutableSource::new(source.clone()));
    let (mutable_budget, _mutable_cancel, context) = managed_context(FINITE_MEMORY, FINITE_WORK);
    let package = source_backed::Package::from_read_at_with_execution_context(
        mutable.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.replace_body_paragraph_texts(&replacements()).unwrap();
    let commit = edit.commit().unwrap();
    mutable.bump_revision();
    let mut refused_output = Vec::new();
    let result = package.publish_document_commit_to_stream(&mut refused_output, &commit);
    assert!(matches!(
        result,
        Err(TransactionError::Document(Error::Opc(
            OpcError::SourceChanged { .. }
        )))
    ));
    assert!(refused_output.is_empty());
    drop(commit);
    drop(snapshot);
    assert_eq!(mutable_budget.used(Resource::Memory), 0);
}

#[test]
fn managed_batch_noop_and_revert_restore_the_exact_source_owner() {
    let source = source_bytes();
    let original = original_replacements();

    let (budget, _cancel, package) = managed(source.clone());
    let snapshot = package.document_snapshot().unwrap();
    let source_ptr = snapshot.xml_bytes().as_ptr();
    let mut edit = snapshot.edit();
    edit.replace_body_paragraph_texts(&original).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().operations().is_empty());
    assert_eq!(commit.snapshot().xml_bytes().as_ptr(), source_ptr);
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
    drop(commit);
    drop(snapshot);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancel, package) = managed(source.clone());
    let snapshot = package.document_snapshot().unwrap();
    let source_ptr = snapshot.xml_bytes().as_ptr();
    let mut edit = snapshot.edit();
    edit.replace_body_paragraph_texts(&replacements()).unwrap();
    assert_ne!(edit.projected().xml_bytes().as_ptr(), source_ptr);
    edit.replace_body_paragraph_texts(&original).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().operations().is_empty());
    assert_eq!(commit.snapshot().xml_bytes().as_ptr(), source_ptr);
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
    drop(commit);
    drop(snapshot);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_batch_composes_with_scalar_edits_in_either_order() {
    let source = source_bytes();
    let batch = replacements();

    let (first_budget, _cancel, first_package) = managed(source.clone());
    let first_snapshot = first_package.document_snapshot().unwrap();
    let mut first_edit = first_snapshot.edit();
    first_edit
        .replace_paragraph_text(Position::new(1), "second scalar")
        .unwrap();
    first_edit.replace_body_paragraph_texts(&batch).unwrap();
    let first_commit = first_edit.commit().unwrap();
    let first_xml = first_commit.snapshot().xml_bytes().to_vec();
    let mut first_output = Vec::new();
    first_package
        .publish_document_commit_to_stream(&mut first_output, &first_commit)
        .unwrap();

    let (second_budget, _cancel, second_package) = managed(source);
    let second_snapshot = second_package.document_snapshot().unwrap();
    let mut second_edit = second_snapshot.edit();
    second_edit.replace_body_paragraph_texts(&batch).unwrap();
    second_edit
        .replace_paragraph_text(Position::new(1), "second scalar")
        .unwrap();
    let second_commit = second_edit.commit().unwrap();
    let mut second_output = Vec::new();
    second_package
        .publish_document_commit_to_stream(&mut second_output, &second_commit)
        .unwrap();

    assert_eq!(first_xml, second_commit.snapshot().xml_bytes());
    assert_eq!(first_output, second_output);
    assert_eq!(
        first_commit
            .snapshot()
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "second scalar"
    );
    drop(first_commit);
    drop(first_snapshot);
    drop(second_commit);
    drop(second_snapshot);
    assert_eq!(first_budget.used(Resource::Memory), 0);
    assert_eq!(second_budget.used(Resource::Memory), 0);
}

#[test]
fn managed_batch_rejects_empty_duplicate_and_unsorted_selectors_atomically() {
    let (budget, _cancel, package) = managed(source_bytes());
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    let before = edit.projected().xml_bytes().to_vec();
    for invalid in [
        Vec::new(),
        vec![
            ParagraphTextReplacement::new(Position::new(0), "one"),
            ParagraphTextReplacement::new(Position::new(0), "duplicate"),
        ],
        vec![
            ParagraphTextReplacement::new(Position::new(2), "third first"),
            ParagraphTextReplacement::new(Position::new(0), "first second"),
        ],
    ] {
        let result = edit.replace_body_paragraph_texts(&invalid);
        assert!(matches!(
            result,
            Err(TransactionError::Refused {
                reason: Refusal::AmbiguousCompositeSelector,
                ..
            })
        ));
        assert_eq!(edit.projected().xml_bytes(), before.as_slice());
    }
    drop(edit);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_batch_late_leaf_failures_leave_the_staged_projection_unchanged() {
    let (budget, _cancel, package) = managed(source_bytes());
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.replace_paragraph_text(Position::new(0), "first scalar")
        .unwrap();
    let staged = edit.projected().xml_bytes().to_vec();
    let work_before = budget.used(Resource::Work);
    let input_before = budget.used(Resource::InputBytes);

    let invalid_text = vec![
        ParagraphTextReplacement::new(Position::new(1), "second valid"),
        ParagraphTextReplacement::new(Position::new(2), "late\nbreak"),
    ];
    let result = edit.replace_body_paragraph_texts(&invalid_text);
    assert!(matches!(
        result,
        Err(TransactionError::Refused {
            reason: Refusal::StructuralText,
            ..
        })
    ));
    assert_eq!(edit.projected().xml_bytes(), staged.as_slice());
    assert!(budget.used(Resource::Work) >= work_before);
    assert!(budget.used(Resource::InputBytes) >= input_before);

    let invalid_position = vec![
        ParagraphTextReplacement::new(Position::new(1), "second valid"),
        ParagraphTextReplacement::new(Position::new(99), "out of bounds"),
    ];
    let result = edit.replace_body_paragraph_texts(&invalid_position);
    assert!(matches!(
        result,
        Err(TransactionError::OutOfBounds { position: 99, .. })
    ));
    assert_eq!(edit.projected().xml_bytes(), staged.as_slice());
    assert!(budget.used(Resource::Work) >= work_before);
    assert!(budget.used(Resource::InputBytes) >= input_before);

    let commit = edit.commit().unwrap();
    assert_eq!(commit.diagnostics().operations(), 1);
    drop(commit);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_batch_cancellation_work_and_memory_failures_are_typed_and_release_owners() {
    let (budget, cancellation, package) = managed(source_bytes());
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    let before = edit.projected().xml_bytes().to_vec();
    let work_before = budget.used(Resource::Work);
    let input_before = budget.used(Resource::InputBytes);
    cancellation.cancel();
    let result = edit.replace_body_paragraph_texts(&replacements());
    assert!(result.as_ref().is_err_and(is_cancelled));
    assert_eq!(edit.projected().xml_bytes(), before.as_slice());
    assert!(budget.used(Resource::Work) >= work_before);
    assert!(budget.used(Resource::InputBytes) >= input_before);
    drop(edit);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);

    let (probe_budget, _probe_cancel, probe_package) = managed(source_bytes());
    let probe_snapshot = probe_package.document_snapshot().unwrap();
    let baseline_work = probe_budget.used(Resource::Work);
    assert!(baseline_work > 0);
    drop(probe_snapshot);
    drop(probe_package);
    assert_eq!(probe_budget.used(Resource::Memory), 0);

    let (budget, _cancel, package) =
        managed_with_work(source_bytes(), FINITE_MEMORY, baseline_work + 1);
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    let before = edit.projected().xml_bytes().to_vec();
    let work_before = budget.used(Resource::Work);
    let input_before = budget.used(Resource::InputBytes);
    let result = edit.replace_body_paragraph_texts(&replacements());
    assert!(
        result
            .as_ref()
            .is_err_and(|error| is_resource_limit(error, Resource::Work))
    );
    assert_eq!(edit.projected().xml_bytes(), before.as_slice());
    assert!(budget.used(Resource::Work) >= work_before);
    assert!(budget.used(Resource::InputBytes) >= input_before);
    drop(edit);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);

    let (budget, _cancel, package) = managed(source_bytes());
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    let before = edit.projected().xml_bytes().to_vec();
    let remaining = budget
        .limit(Resource::Memory)
        .saturating_sub(budget.used(Resource::Memory));
    assert!(remaining > 0);
    let hold = budget.reserve(Resource::Memory, remaining).unwrap();
    let work_before = budget.used(Resource::Work);
    let input_before = budget.used(Resource::InputBytes);
    let result = edit.replace_body_paragraph_texts(&replacements());
    assert!(
        result
            .as_ref()
            .is_err_and(|error| is_resource_limit(error, Resource::Memory))
    );
    assert_eq!(edit.projected().xml_bytes(), before.as_slice());
    assert!(budget.used(Resource::Work) >= work_before);
    assert!(budget.used(Resource::InputBytes) >= input_before);
    drop(hold);
    drop(edit);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_signed_batch_noop_is_exact_and_change_is_refused() {
    let source = signed_source_bytes();
    let original = original_replacements();

    let (budget, _cancel, package) = managed(source.clone());
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.replace_body_paragraph_texts(&original).unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
    drop(commit);
    drop(snapshot);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancel, package) = managed(source);
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    let result = edit.replace_body_paragraph_texts(&replacements());
    assert!(matches!(
        result,
        Err(TransactionError::Document(Error::UnsafeEdit { .. }))
    ));
    drop(edit);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}
