#![cfg(any(unix, windows))]

//! Ordinary managed DOCX document transactions.
//!
//! These cases keep the execution context and memory budget visible while an
//! opened source is captured, edited, committed, and published. They are
//! deliberately separate from the selective managed-read coverage: the
//! ordinary document snapshot/edit path must retain its source owner and every
//! candidate owner through the commit and publication boundary.

use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};
use std::sync::{Arc, Mutex};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, Position, ReadAt, Resource, SourceVersion,
};
use litchi_docx::document::{Commit, TransactionError};
use litchi_docx::paragraph::{Collapsed, Inline, Symbols};
use litchi_docx::{Error, ReadLimits, source_backed};
use litchi_opc::constants::{content_type as ct, relationship_type as rt};
use litchi_opc::rel::Relationships;
use litchi_opc::{OpcError, OpcPackage, PackURI, SourceCacheLimits};
use soapberry_zip::office::StreamingArchiveWriter;

const MAIN: &str = "word/document.xml";
const MEDIA: &str = "word/media/image1.png";
const OPAQUE: &str = "word/opaque.bin";
const SETTINGS: &str = "word/settings.xml";
const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W12: &str = "http://schemas.microsoft.com/office/word/2012/wordml";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const SYMEX: &str = "http://schemas.microsoft.com/office/word/2015/wordml/symex";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";
const STRICT_W: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const MCE: &str = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const STRICT_SETTINGS_RELATIONSHIP: &str =
    "http://purl.oclc.org/ooxml/officeDocument/relationships/settings";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const FINITE_INPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OUTPUT_BYTES: u64 = 64 * 1024 * 1024;
const FINITE_OBJECTS: u64 = 1_000_000;
const FINITE_DEPTH: u64 = 1024;
const FINITE_WORK: u64 = 1 << 30;

fn document_xml(first: &str, second: &str) -> Vec<u8> {
    format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r><w:t>{first}</w:t><w:custom w:value="opaque"/></w:r><w:r><w:t> tail</w:t></w:r></w:p><w:p><w:r><w:t>{second}</w:t></w:r></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn empty_text_slot_document(slot_count: usize) -> Vec<u8> {
    let mut slots = String::with_capacity(slot_count.saturating_mul(6));
    for _ in 0..slot_count {
        slots.push_str("<w:t/>");
    }
    format!(
        r#"<w:document xmlns:w="{W}"><w:body><w:p><w:r>{slots}</w:r></w:p></w:body></w:document>"#
    )
    .into_bytes()
}

fn fixture(first: &str, second: &str) -> Vec<u8> {
    fixture_with_document(document_xml(first, second), 37, 13)
}

fn byte_order_marked_document(first: &str, second: &str) -> Vec<u8> {
    let mut document = b"\xEF\xBB\xBF".to_vec();
    document.extend_from_slice(&document_xml(first, second));
    document
}

fn managed_collapsed_hyperlink_fixture() -> Vec<u8> {
    fixture_with_document(
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:w12="{W12}" xmlns:r="{R}"><w:body><w:p><w:pPr><w12:collapsed w12:val="on"/></w:pPr><w:hyperlink r:id="rLink"><w:r><w:t>linked</w:t></w:r></w:hyperlink><w:r><w:t>direct</w:t></w:r></w:p></w:body></w:document>"#
        )
        .into_bytes(),
        37,
        13,
    )
}

fn managed_symex_namespace_fixture() -> Vec<u8> {
    fixture_with_document(
        format!(
            r#"<w:document xmlns:w="{W}" xmlns:sx="{SYMEX}"><w:body><w:p><w:r><sx:symEx sx:font="Segoe UI Symbol" sx:char="00000041"/><w:t>inherited</w:t></w:r></w:p><w:p xmlns:sx="urn:foreign-shadow"><w:r xmlns:sx="urn:foreign-run"><sx:symEx sx:char="00000042"/><w:t>shadowed</w:t></w:r></w:p><w:p><w:r><sx:symEx sx:char="00000044"/><w:t>restored</w:t></w:r></w:p><w:p><w:r xmlns:foreign="urn:foreign-binding"><foreign:symEx foreign:char="00000043"/><w:t>foreign</w:t></w:r></w:p></w:body></w:document>"#
        )
        .into_bytes(),
        37,
        13,
    )
}

fn hyperlink_relationships() -> Relationships {
    let mut relationships = Relationships::new("/word/document.xml".to_owned());
    relationships.add_relationship(
        rt::HYPERLINK.to_owned(),
        "https://managed.invalid/".to_owned(),
        "rLink".to_owned(),
        true,
    );
    relationships
}

fn multi_paragraph_document(values: &[&str]) -> Vec<u8> {
    let mut body = String::new();
    for value in values {
        body.push_str(&format!(r#"<w:p><w:r><w:t>{value}</w:t></w:r></w:p>"#));
    }
    format!(r#"<w:document xmlns:w="{W}"><w:body>{body}</w:body></w:document>"#).into_bytes()
}

fn fixture_signed(first: &str, second: &str) -> Vec<u8> {
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
    archive_fixture(&document_xml(first, second), &media, &opaque, None, true)
}

fn fixture_with_document(
    document: Vec<u8>,
    media_multiplier: u8,
    opaque_multiplier: u8,
) -> Vec<u8> {
    let media: Vec<u8> = (0_u8..=u8::MAX)
        .cycle()
        .take(4096)
        .map(|byte| byte.wrapping_mul(media_multiplier))
        .collect();
    let opaque: Vec<u8> = (0_u8..=u8::MAX)
        .cycle()
        .take(8192)
        .map(|byte| byte.wrapping_mul(opaque_multiplier))
        .collect();
    archive_fixture(&document, &media, &opaque, None, false)
}

fn settings_xml(namespace: &str, body: &str) -> Vec<u8> {
    format!(r#"<w:settings xmlns:w="{namespace}">{body}</w:settings>"#).into_bytes()
}

fn standard_ignorable_settings() -> Vec<u8> {
    format!(
        r#"<w:settings xmlns:w="{W}" xmlns:mc="{MCE}" xmlns:w14="{W14}" mc:Ignorable="w14"><w:zoom w:val="bestFit"/><w14:docId w14:val="7F00AA10"/></w:settings>"#
    )
    .into_bytes()
}

fn fixture_with_settings(
    document: Vec<u8>,
    settings: Vec<u8>,
    strict_settings_relationship: bool,
) -> Vec<u8> {
    let settings_relationship = if strict_settings_relationship {
        STRICT_SETTINGS_RELATIONSHIP
    } else {
        rt::SETTINGS
    };
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
    archive_fixture(
        &document,
        &media,
        &opaque,
        Some((&settings, settings_relationship)),
        false,
    )
}

fn archive_fixture(
    document: &[u8],
    media: &[u8],
    opaque: &[u8],
    settings: Option<(&[u8], &str)>,
    signed: bool,
) -> Vec<u8> {
    let settings_override = if settings.is_some() {
        format!(
            r#"<Override PartName="/{SETTINGS}" ContentType="{}"/>"#,
            ct::WML_SETTINGS
        )
    } else {
        String::new()
    };
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="bin" ContentType="application/octet-stream"/><Default Extension="sigs" ContentType="{signature_origin}"/><Override PartName="/{MAIN}" ContentType="{document_content_type}"/>{settings_override}</Types>"#,
        signature_origin = ct::OPC_DIGITAL_SIGNATURE_ORIGIN,
        document_content_type = ct::WML_DOCUMENT_MAIN,
    );

    let mut relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rIdDocument" Type="{office_document}" Target="{MAIN}"/>"#,
        office_document = rt::OFFICE_DOCUMENT,
    );
    if signed {
        relationships.push_str(&format!(
            r#"<Relationship Id="rIdSignature" Type="{signature_relationship}" Target="_xmlsignatures/origin.sigs"/>"#,
            signature_relationship = rt::DIGITAL_SIGNATURE_ORIGIN,
        ));
    }
    relationships.push_str("</Relationships>");

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .unwrap();
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .unwrap();
    if let Some((_, settings_relationship)) = settings {
        let main_relationships = format!(
            r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rIdSettings" Type="{settings_relationship}" Target="settings.xml"/></Relationships>"#
        );
        writer
            .write_stored(
                "word/_rels/document.xml.rels",
                main_relationships.as_bytes(),
            )
            .unwrap();
    }
    writer.write_stored(MEDIA, media).unwrap();
    writer.write_stored(OPAQUE, opaque).unwrap();
    writer.write_stored(MAIN, document).unwrap();
    if let Some((settings, _)) = settings {
        writer.write_stored(SETTINGS, settings).unwrap();
    }
    if signed {
        writer
            .write_stored("_xmlsignatures/origin.sigs", b"<origin/>")
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn part_bytes(bytes: &[u8], name: &str) -> Vec<u8> {
    let package = OpcPackage::from_bytes(bytes).unwrap();
    package
        .get_part(&PackURI::new(format!("/{name}")).unwrap())
        .unwrap()
        .blob()
        .to_vec()
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    managed_context_with_output(memory, FINITE_OUTPUT_BYTES)
}

fn managed_context_with_output(
    memory: u64,
    output_bytes: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "docx-managed-document-edit-test",
        Limits::new(
            memory,
            FINITE_INPUT_BYTES,
            output_bytes,
            FINITE_OBJECTS,
            FINITE_DEPTH,
            FINITE_WORK,
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

fn managed(bytes: Vec<u8>, memory: u64) -> (Budget, CancellationSource, source_backed::Package) {
    let (budget, cancellation_source, context) = managed_context(memory);
    let package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    (budget, cancellation_source, package)
}

fn publish_managed_edit_with_inverse(
    source: Vec<u8>,
    replacement: &str,
) -> (Vec<u8>, source_backed::DocumentPublication, Budget, Commit) {
    let (budget, _cancellation_source, package) = managed(source, 1 << 20);
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.replace_paragraph_text(Position::new(0), replacement)
        .unwrap();
    let commit = edit.commit().unwrap();
    drop(snapshot);
    let mut output = Vec::new();
    let publication = package
        .publish_document_commit_with_inverse_to_stream(&mut output, &commit)
        .unwrap();
    (output, publication, budget, commit)
}

fn is_source_changed(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(OpcError::SourceChanged { .. }))
    )
}

fn is_cancelled(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(OpcError::Cancelled))
    )
}

fn is_budget_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(OpcError::Execution(
            ExecutionError::ResourceLimit(limit),
        ))) if limit.resource == Resource::Memory
    )
}

fn is_memory_error(error: &Error) -> bool {
    matches!(
        error,
        Error::Opc(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::Memory
    )
}

fn is_output_budget_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(OpcError::Execution(
            ExecutionError::ResourceLimit(limit),
        ))) if limit.resource == Resource::OutputBytes
    )
}

fn is_artifact_mismatch(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(OpcError::SourceArtifactMismatch {
            artifact: "current",
            ..
        }))
    )
}

fn is_protection_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::UnsafeEdit {
            operation: "changed document publication",
            reason: "document or write protection, or tracked revisions, is enabled",
            ..
        })
    )
}

fn is_mixed_dialect_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::InvalidRelationship(reason))
            if reason.contains("mixed OOXML conformance families")
    )
}

fn is_mce_policy_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::UnsafeEdit {
            operation: "changed document publication",
            reason: "markup-compatibility projection or unknown source policy is not admitted",
            ..
        })
    )
}

fn is_settings_policy_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::UnsafeEdit {
            operation: "changed document publication",
            ..
        }) | TransactionError::Document(Error::InvalidFormat(_))
    )
}

fn is_literal_namespace_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(
            OpcError::SourceBackedOverlayUnavailable { reason },
        )) if reason == "source XML requires literal namespace bindings"
    )
}

fn is_dtd_policy_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(
            OpcError::SourceBackedOverlayUnavailable { reason },
        )) if reason == "source XML DTDs are not permitted"
    )
}

fn is_mce_document_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::UnsafeEdit {
            operation: "source-backed document read",
            reason: "markup-compatibility preprocessing would require an unbudgeted owned payload",
            ..
        })
    )
}

fn is_dtd_document_refusal(error: &TransactionError) -> bool {
    matches!(
        error,
        TransactionError::Document(Error::Opc(
            OpcError::SourceBackedOverlayUnavailable { reason },
        )) if reason == "source XML DTDs are not permitted"
    )
}

#[derive(Debug)]
struct MutableSource {
    bytes: Mutex<Vec<u8>>,
    revision: AtomicU64,
    revision_on_read: AtomicBool,
    cancellation_on_read: Mutex<Option<CancellationSource>>,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes: Mutex::new(bytes),
            revision: AtomicU64::new(0),
            revision_on_read: AtomicBool::new(false),
            cancellation_on_read: Mutex::new(None),
        }
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }

    fn arm_revision_on_read(&self) {
        self.revision_on_read.store(true, Ordering::Release);
    }

    fn arm_cancellation_on_read(&self, cancellation: CancellationSource) {
        *self.cancellation_on_read.lock().unwrap() = Some(cancellation);
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
        let count = {
            let bytes = self
                .bytes
                .lock()
                .map_err(|_| io::Error::other("mutable source mutex poisoned"))?;
            if offset >= bytes.len() || output.is_empty() {
                0
            } else {
                let count = output.len().min(bytes.len() - offset);
                output[..count].copy_from_slice(&bytes[offset..offset + count]);
                count
            }
        };
        if count > 0 {
            if self.revision_on_read.swap(false, Ordering::AcqRel) {
                self.bump_revision();
            }
            let cancellation = self
                .cancellation_on_read
                .lock()
                .map_err(|_| io::Error::other("mutable source cancellation mutex poisoned"))?
                .take();
            if let Some(cancellation) = cancellation {
                cancellation.cancel();
            }
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4d_44_4f_43,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[derive(Debug)]
struct PartialWriter {
    bytes: Vec<u8>,
    limit: usize,
}

impl Write for PartialWriter {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.bytes.len() >= self.limit {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "intentional managed publication failure",
            ));
        }
        let accepted = (self.limit - self.bytes.len()).min(bytes.len());
        self.bytes.extend_from_slice(&bytes[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn managed_document_snapshot_and_edit_retain_source_and_candidate_reservations() {
    let source = fixture("managed", "second");
    let (budget, _cancellation_source, package) = managed(source, 1 << 20);
    assert_eq!(budget.used(Resource::Memory), 0);

    let snapshot = package.document_snapshot().unwrap();
    let source_reserved = budget.used(Resource::Memory);
    assert!(source_reserved > 0);
    let paragraph = snapshot.paragraph(Position::new(0)).unwrap();
    let retained_runs = paragraph.runs().unwrap();
    assert_eq!(retained_runs.len(), 2);
    assert_eq!(retained_runs[0].text().unwrap(), "managed");
    let retained_contents = retained_runs[0].contents().unwrap();
    assert!(retained_contents.iter().any(|content| {
        matches!(content, litchi_docx::RunContent::Unknown(opaque) if !opaque.xml_bytes().is_empty())
    }));
    let cloned_snapshot = snapshot.clone();
    assert_eq!(budget.used(Resource::Memory), source_reserved);

    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "candidate")
        .unwrap();
    let candidate_reserved = budget.used(Resource::Memory);
    assert!(candidate_reserved >= source_reserved);
    let commit = edit.commit().unwrap();
    assert!(budget.used(Resource::Memory) >= candidate_reserved);

    drop(commit);
    drop(cloned_snapshot);
    drop(snapshot);
    drop(package);
    assert!(budget.used(Resource::Memory) > 0);
    drop(retained_contents);
    assert!(budget.used(Resource::Memory) > 0);
    drop(retained_runs);
    assert!(budget.used(Resource::Memory) > 0);
    drop(paragraph);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_document_bom_offsets_are_preserved_through_snapshot_and_edit() {
    let (budget, _cancellation_source, package) = managed(
        fixture_with_document(byte_order_marked_document("before", "kept"), 37, 13),
        1 << 20,
    );
    let snapshot = package.document_snapshot().unwrap();
    assert_eq!(
        snapshot
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "before tail"
    );
    assert_eq!(
        snapshot
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "kept"
    );
    assert!(snapshot.xml_bytes().starts_with(b"\xEF\xBB\xBF"));

    let mut edit = snapshot.edit();
    edit.replace_paragraph_text(Position::new(0), "after")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "after"
    );
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "kept"
    );
    assert!(commit.snapshot().xml_bytes().starts_with(b"\xEF\xBB\xBF"));
    drop(commit);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_paragraph_views_refuse_parser_growth_at_memory_boundary() {
    let (budget, _cancellation_source, package) = managed(fixture("managed", "second"), 1 << 20);
    let snapshot = package.document_snapshot().unwrap();
    let paragraph = snapshot.paragraph(Position::new(0)).unwrap();
    let source_reserved = budget.used(Resource::Memory);
    let remaining = budget
        .limit(Resource::Memory)
        .saturating_sub(source_reserved);
    assert!(remaining > 0);
    let hold = budget.reserve(Resource::Memory, remaining).unwrap();
    assert_eq!(
        budget.used(Resource::Memory),
        budget.limit(Resource::Memory)
    );

    let text = paragraph.text();
    assert!(text.as_ref().is_err_and(is_memory_error));
    let runs = paragraph.runs();
    assert!(runs.as_ref().is_err_and(is_memory_error));
    let inlines = paragraph.inlines();
    assert!(inlines.as_ref().is_err_and(is_memory_error));

    drop(hold);
    assert_eq!(budget.used(Resource::Memory), source_reserved);
    let text = paragraph.text().unwrap();
    assert_eq!(text, "managed tail");
    let runs = paragraph.runs().unwrap();
    assert_eq!(runs.len(), 2);
    let inlines = paragraph.inlines().unwrap();
    assert_eq!(inlines.len(), 2);

    drop(snapshot);
    drop(package);
    assert!(budget.used(Resource::Memory) > 0);
    drop(inlines);
    assert!(budget.used(Resource::Memory) > 0);
    drop(runs);
    assert!(budget.used(Resource::Memory) > 0);
    drop(paragraph);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_paragraph_and_nested_run_mutations_refuse_without_state_change() {
    let (budget, _cancellation_source, package) =
        managed(managed_collapsed_hyperlink_fixture(), 1 << 20);
    let snapshot = package.document_snapshot().unwrap();
    let mut paragraph = snapshot.paragraph(Position::new(0)).unwrap();

    let before_collapsed = paragraph.collapsed().unwrap();
    assert_eq!(before_collapsed, Some(Collapsed::Enabled));
    for value in [None, Some(Collapsed::Disabled)] {
        let result = paragraph.set_collapsed(value);
        assert!(matches!(
            result,
            Err(Error::UnsafeEdit {
                operation: "paragraph.set_collapsed",
                ..
            })
        ));
        assert_eq!(paragraph.collapsed().unwrap(), before_collapsed);
    }

    let relationships = hyperlink_relationships();
    let inlines = paragraph
        .inlines_with_relationships(&relationships)
        .unwrap();
    let mut nested_run = match &inlines[0] {
        Inline::Hyperlink(hyperlink) => hyperlink.runs()[0].clone(),
        _ => panic!("expected a relationship-resolved hyperlink"),
    };
    let before_symbols = nested_run.symbols().unwrap();
    assert!(before_symbols.is_empty());
    let result = nested_run.set_symbols(Symbols::default());
    assert!(matches!(
        result,
        Err(Error::UnsafeEdit {
            operation: "run.set_symbols",
            ..
        })
    ));
    assert_eq!(nested_run.symbols().unwrap(), before_symbols);

    drop(nested_run);
    drop(inlines);
    drop(paragraph);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_collapsed_and_symbols_reads_refuse_when_memory_is_held() {
    let (budget, _cancellation_source, package) =
        managed(managed_collapsed_hyperlink_fixture(), 1 << 20);
    let snapshot = package.document_snapshot().unwrap();
    let paragraph = snapshot.paragraph(Position::new(0)).unwrap();
    let relationships = hyperlink_relationships();
    let inlines = paragraph
        .inlines_with_relationships(&relationships)
        .unwrap();
    let nested_run = match &inlines[0] {
        Inline::Hyperlink(hyperlink) => hyperlink.runs()[0].clone(),
        _ => panic!("expected a relationship-resolved hyperlink"),
    };

    let source_reserved = budget.used(Resource::Memory);
    let remaining = budget
        .limit(Resource::Memory)
        .saturating_sub(source_reserved);
    assert!(remaining > 0);
    let hold = budget.reserve(Resource::Memory, remaining).unwrap();
    assert_eq!(
        budget.used(Resource::Memory),
        budget.limit(Resource::Memory)
    );

    let collapsed = paragraph.collapsed();
    assert!(collapsed.as_ref().is_err_and(is_memory_error));
    let symbols = nested_run.symbols();
    assert!(symbols.as_ref().is_err_and(is_memory_error));

    drop(hold);
    assert_eq!(paragraph.collapsed().unwrap(), Some(Collapsed::Enabled));
    assert!(nested_run.symbols().unwrap().is_empty());

    drop(nested_run);
    drop(inlines);
    drop(paragraph);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_hyperlink_nested_run_retains_source_after_package_drop_and_refuses_mutation() {
    let (budget, _cancellation_source, package) =
        managed(managed_collapsed_hyperlink_fixture(), 1 << 20);
    let snapshot = package.document_snapshot().unwrap();
    let paragraph = snapshot.paragraph(Position::new(0)).unwrap();
    let relationships = hyperlink_relationships();
    let inlines = paragraph
        .inlines_with_relationships(&relationships)
        .unwrap();
    let mut nested_run = match &inlines[0] {
        Inline::Hyperlink(hyperlink) => hyperlink.runs()[0].clone(),
        _ => panic!("expected a relationship-resolved hyperlink"),
    };
    drop(inlines);
    drop(paragraph);
    drop(snapshot);
    drop(package);
    assert!(budget.used(Resource::Memory) > 0);

    let before_symbols = nested_run.symbols().unwrap();
    let result = nested_run.set_symbols(Symbols::default());
    assert!(matches!(
        result,
        Err(Error::UnsafeEdit {
            operation: "run.set_symbols",
            ..
        })
    ));
    assert_eq!(nested_run.symbols().unwrap(), before_symbols);

    drop(nested_run);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_retained_runs_resolve_inherited_symex_and_ignore_foreign_bindings() {
    let (budget, _cancellation_source, package) =
        managed(managed_symex_namespace_fixture(), 1 << 20);
    let snapshot = package.document_snapshot().unwrap();
    let inherited_paragraph = snapshot.paragraph(Position::new(0)).unwrap();
    let shadowed_paragraph = snapshot.paragraph(Position::new(1)).unwrap();
    let restored_paragraph = snapshot.paragraph(Position::new(2)).unwrap();
    let foreign_paragraph = snapshot.paragraph(Position::new(3)).unwrap();
    let inherited_run = inherited_paragraph
        .runs()
        .unwrap()
        .into_iter()
        .next()
        .expect("inherited symEx paragraph must retain one run");
    let shadowed_run = shadowed_paragraph
        .runs()
        .unwrap()
        .into_iter()
        .next()
        .expect("shadowed symEx paragraph must retain one run");
    let restored_run = restored_paragraph
        .runs()
        .unwrap()
        .into_iter()
        .next()
        .expect("sibling after shadowed symEx paragraph must retain one run");
    let foreign_run = foreign_paragraph
        .runs()
        .unwrap()
        .into_iter()
        .next()
        .expect("foreign symEx paragraph must retain one run");

    drop(inherited_paragraph);
    drop(shadowed_paragraph);
    drop(restored_paragraph);
    drop(foreign_paragraph);
    drop(snapshot);
    drop(package);
    assert!(budget.used(Resource::Memory) > 0);

    let inherited_symbols = inherited_run.symbols().unwrap();
    assert_eq!(inherited_symbols.len(), 1);
    assert_eq!(
        inherited_symbols.first().unwrap().font(),
        Some("Segoe UI Symbol")
    );
    assert_eq!(inherited_symbols.first().unwrap().char_code(), Some(0x41));
    assert!(shadowed_run.symbols().unwrap().is_empty());
    let restored_symbols = restored_run.symbols().unwrap();
    assert_eq!(restored_symbols.len(), 1);
    assert_eq!(restored_symbols.first().unwrap().char_code(), Some(0x44));
    assert!(foreign_run.symbols().unwrap().is_empty());

    drop(inherited_run);
    drop(shadowed_run);
    drop(restored_run);
    drop(foreign_run);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_edit_try_clone_refuses_memory_boundary_without_mutating_projection() {
    let (budget, _cancellation_source, package) = managed(fixture("managed", "second"), 1 << 20);
    let snapshot = package.document_snapshot().unwrap();
    let mut edit = snapshot.edit();
    edit.replace_paragraph_text(Position::new(0), "candidate")
        .unwrap();
    let expected_xml = edit.projected().xml_bytes().to_vec();
    let source_reserved = budget.used(Resource::Memory);
    let remaining = budget
        .limit(Resource::Memory)
        .saturating_sub(source_reserved);
    assert!(remaining > 0);
    let hold = budget.reserve(Resource::Memory, remaining).unwrap();

    let result = edit.try_clone();
    assert!(result.as_ref().is_err_and(is_budget_refusal));
    assert_eq!(edit.projected().xml_bytes(), expected_xml.as_slice());

    drop(hold);
    let cloned = edit.try_clone().unwrap();
    assert_eq!(cloned.projected().xml_bytes(), expected_xml.as_slice());
    let commit = cloned.commit().unwrap();
    assert_eq!(commit.snapshot().xml_bytes(), expected_xml.as_slice());

    drop(commit);
    drop(edit);
    drop(snapshot);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_signed_noop_and_equal_setter_preserve_exact_source() {
    let source = fixture_signed("before", "second");

    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let noop = package.edit_document().unwrap().commit().unwrap();
    assert!(!noop.diagnostics().changed());
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &noop)
        .unwrap();
    assert_eq!(output, source);
    drop(noop);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "before tail")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    let result = edit.replace_paragraph_text(Position::new(0), "changed");
    assert!(matches!(
        result,
        Err(TransactionError::Document(Error::UnsafeEdit { .. }))
    ));
    drop(edit);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);

    let package = source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.publish_document_commit_to_stream(&mut output, &commit),
        Err(TransactionError::Document(Error::Opc(
            OpcError::SignedSourceRequiresExplicitPolicy
        )))
    ));
    assert!(output.is_empty());
}

#[test]
fn managed_disjoint_updates_publish_in_arbitrary_order_and_inverse_exact() {
    let source = fixture("before", "second");
    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(1), "second changed")
        .unwrap();
    edit.replace_paragraph_text(Position::new(0), "first changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "first changed"
    );
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "second changed"
    );
    let restored = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(restored.xml_bytes(), commit.patch().source().xml_bytes());
    drop(restored);

    let mut output = Vec::new();
    let publication = package
        .publish_document_commit_with_inverse_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(part_bytes(&output, MEDIA), part_bytes(&source, MEDIA));
    assert_eq!(part_bytes(&output, OPAQUE), part_bytes(&source, OPAQUE));
    let (inverse_budget, _cancellation_source, reopened) = managed(output, 1 << 20);
    let snapshot = reopened.document_snapshot().unwrap();
    assert_eq!(
        snapshot
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "first changed"
    );
    assert_eq!(
        snapshot
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "second changed"
    );
    let mut restored = Vec::new();
    reopened
        .publish_document_inverse_to_stream(&mut restored, &publication)
        .unwrap();
    assert_eq!(restored, source);
    drop(snapshot);
    assert_eq!(inverse_budget.used(Resource::Memory), 0);
    drop(publication);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn unmanaged_foreign_package_with_equal_main_xml_refuses_stale_commit_before_output() {
    let main_xml = document_xml("before", "second");
    let first = fixture_with_document(main_xml.clone(), 37, 13);
    let second = fixture_with_document(main_xml, 41, 29);
    assert_eq!(part_bytes(&first, MAIN), part_bytes(&second, MAIN));
    assert_ne!(part_bytes(&first, MEDIA), part_bytes(&second, MEDIA));
    assert_ne!(part_bytes(&first, OPAQUE), part_bytes(&second, OPAQUE));

    let first_package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(first))).unwrap();
    let mut edit = first_package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "first package change")
        .unwrap();
    let commit = edit.commit().unwrap();
    drop(first_package);

    let second_package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(second))).unwrap();
    let mut output = Vec::new();
    let result = second_package.publish_document_commit_to_stream(&mut output, &commit);
    assert!(matches!(result, Err(TransactionError::StaleSource)));
    assert!(output.is_empty());
}

#[test]
fn managed_four_disjoint_updates_release_scan_depth_and_restore_exact_artifact() {
    let source = fixture_with_document(
        multi_paragraph_document(&["zero", "one", "two", "three"]),
        37,
        13,
    );
    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    for (position, replacement) in [
        (3_usize, "three changed"),
        (1_usize, "one changed"),
        (0_usize, "zero changed"),
        (2_usize, "two changed"),
    ] {
        edit.replace_paragraph_text(Position::new(position), replacement)
            .unwrap();
    }
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().operations().len(), 4);
    for (position, expected) in [
        (0_usize, "zero changed"),
        (1_usize, "one changed"),
        (2_usize, "two changed"),
        (3_usize, "three changed"),
    ] {
        assert_eq!(
            commit
                .snapshot()
                .paragraph(Position::new(position))
                .unwrap()
                .text()
                .unwrap(),
            expected
        );
    }

    let mut changed = Vec::new();
    let publication = package
        .publish_document_commit_with_inverse_to_stream(&mut changed, &commit)
        .unwrap();
    let (inverse_budget, _cancellation_source, reopened) = managed(changed, 1 << 20);
    let mut restored = Vec::new();
    reopened
        .publish_document_inverse_to_stream(&mut restored, &publication)
        .unwrap();
    assert_eq!(restored, source);
    assert_eq!(inverse_budget.used(Resource::Memory), 0);
    drop(publication);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_repeated_position_keeps_final_value_and_collapses_equal_setter() {
    let (budget, _cancellation_source, package) = managed(fixture("before", "second"), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "first changed")
        .unwrap();
    edit.replace_paragraph_text(Position::new(0), "first final")
        .unwrap();
    edit.replace_paragraph_text(Position::new(0), "first final")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().operations().len(), 1);
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "first final"
    );
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "second"
    );
    drop(commit);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_reverting_one_operation_preserves_the_other_and_inverse_artifact() {
    let source = fixture("before", "second");
    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "first changed")
        .unwrap();
    edit.replace_paragraph_text(Position::new(1), "second changed")
        .unwrap();
    edit.replace_paragraph_text(Position::new(0), "before tail")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(commit.patch().operations().len(), 1);
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "before tail"
    );
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "second changed"
    );
    let restored_xml = commit.patch().inverse().apply(commit.snapshot()).unwrap();
    assert_eq!(
        restored_xml.xml_bytes(),
        commit.patch().source().xml_bytes()
    );
    drop(restored_xml);

    let mut changed = Vec::new();
    let publication = package
        .publish_document_commit_with_inverse_to_stream(&mut changed, &commit)
        .unwrap();
    assert_ne!(changed, source);
    let (inverse_budget, _cancellation_source, reopened) = managed(changed, 1 << 20);
    let mut restored = Vec::new();
    reopened
        .publish_document_inverse_to_stream(&mut restored, &publication)
        .unwrap();
    assert_eq!(restored, source);
    assert_eq!(inverse_budget.used(Resource::Memory), 0);
    drop(publication);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_revert_to_original_publishes_exact_noop_bytes() {
    let source = fixture("before", "second");
    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "temporary")
        .unwrap();
    edit.replace_paragraph_text(Position::new(0), "before tail")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(!commit.diagnostics().changed());
    assert!(commit.patch().operations().is_empty());
    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_failed_second_edit_budget_leaves_first_projection_intact() {
    let (budget, _cancellation_source, package) = managed(fixture("before", "second"), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "first changed")
        .unwrap();
    let first_projected = edit.projected().xml_bytes().to_vec();
    let remaining = budget
        .limit(Resource::Memory)
        .saturating_sub(budget.used(Resource::Memory));
    assert!(remaining > 0);
    let hold = budget.reserve(Resource::Memory, remaining).unwrap();
    let result = edit.replace_paragraph_text(Position::new(1), "second changed");
    assert!(result.as_ref().is_err_and(is_budget_refusal));
    assert_eq!(edit.projected().xml_bytes(), first_projected.as_slice());
    drop(hold);

    let commit = edit.commit().unwrap();
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "first changed"
    );
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(1))
            .unwrap()
            .text()
            .unwrap(),
        "second"
    );
    drop(commit);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_failed_second_edit_cancellation_leaves_first_projection_intact() {
    let (budget, cancellation_source, package) = managed(fixture("before", "second"), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "first changed")
        .unwrap();
    let first_projected = edit.projected().xml_bytes().to_vec();
    cancellation_source.cancel();
    let result = edit.replace_paragraph_text(Position::new(1), "second changed");
    assert!(result.as_ref().is_err_and(is_cancelled));
    assert_eq!(edit.projected().xml_bytes(), first_projected.as_slice());
    drop(edit);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_noop_inverse_and_changed_publication_preserve_untouched_media() {
    let source = fixture("before", "second");
    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let noop = package.edit_document().unwrap().commit().unwrap();
    assert!(!noop.diagnostics().changed());
    let mut noop_output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut noop_output, &noop)
        .unwrap();
    assert_eq!(noop_output, source);
    drop(noop);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    let source_xml = edit.source().xml_bytes().to_vec();
    edit.replace_paragraph_text(Position::new(0), "after")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    let inverse = commit.patch().inverse();
    let restored = inverse.apply(commit.snapshot()).unwrap();
    assert_eq!(restored.xml_bytes(), source_xml.as_slice());
    drop(restored);
    drop(inverse);

    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_ne!(output, source);
    assert_eq!(part_bytes(&output, MEDIA), part_bytes(&source, MEDIA));
    assert_eq!(part_bytes(&output, OPAQUE), part_bytes(&source, OPAQUE));
    let reopened =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(output))).unwrap();
    assert_eq!(
        reopened
            .document_snapshot()
            .unwrap()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "after"
    );
    drop(reopened);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_protected_settings_noop_preserves_relationship_and_payload_exactly() {
    let settings = settings_xml(
        W,
        r#"<w:documentProtection w:edit="readOnly" w:enforcement="1"/>"#,
    );
    let source = fixture_with_settings(document_xml("before", "second"), settings.clone(), false);
    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let commit = package.edit_document().unwrap().commit().unwrap();
    assert!(!commit.diagnostics().changed());

    let mut output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(output, source);
    assert_eq!(part_bytes(&output, SETTINGS), settings);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_changed_publication_enforces_settings_protection_flags() {
    let cases = [
        (
            "documentProtection enabled",
            r#"<w:documentProtection w:edit="readOnly" w:enforcement="1"/>"#,
            true,
        ),
        (
            "documentProtection disabled",
            r#"<w:documentProtection w:edit="readOnly" w:enforcement="0"/>"#,
            false,
        ),
        (
            "writeProtection enabled",
            r#"<w:writeProtection w:val="1"/>"#,
            true,
        ),
        (
            "writeProtection disabled",
            r#"<w:writeProtection w:val="0"/>"#,
            false,
        ),
        (
            "trackRevisions enabled",
            r#"<w:trackRevisions w:val="1"/>"#,
            true,
        ),
        (
            "trackRevisions disabled",
            r#"<w:trackRevisions w:val="0"/>"#,
            false,
        ),
    ];
    for (label, body, refused) in cases {
        let source = fixture_with_settings(
            document_xml("before", "second"),
            settings_xml(W, body),
            false,
        );
        let (budget, _cancellation_source, package) = managed(source, 1 << 20);
        let mut edit = package.edit_document().unwrap();
        edit.replace_paragraph_text(Position::new(0), "changed")
            .unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        let result = package.publish_document_commit_to_stream(&mut output, &commit);
        if refused {
            assert!(
                result.as_ref().is_err_and(is_protection_refusal),
                "{label} must refuse changed publication: {:?}",
                result.as_ref().err()
            );
            assert!(output.is_empty(), "{label} must refuse before output");
        } else {
            result.expect("explicitly disabled policy must admit changed publication");
            assert!(!output.is_empty(), "{label} must publish changed output");
        }
        drop(commit);
        assert_eq!(budget.used(Resource::Memory), 0, "{label} leaked budget");
    }
}

#[test]
fn changed_publication_preserves_standard_ignorable_settings_for_managed_and_unmanaged_sources() {
    let settings = standard_ignorable_settings();
    let source = fixture_with_settings(document_xml("before", "second"), settings.clone(), false);

    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "managed changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut managed_output = Vec::new();
    package
        .publish_document_commit_to_stream(&mut managed_output, &commit)
        .unwrap();
    assert_eq!(part_bytes(&managed_output, SETTINGS), settings);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);

    let unmanaged_package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = unmanaged_package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "unmanaged changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut unmanaged_output = Vec::new();
    unmanaged_package
        .publish_document_commit_to_stream(&mut unmanaged_output, &commit)
        .unwrap();
    assert_eq!(part_bytes(&unmanaged_output, SETTINGS), settings);
}

#[test]
fn protected_track_revisions_with_ignorable_extension_remains_refused() {
    let settings = format!(
        r#"<w:settings xmlns:w="{W}" xmlns:mc="{MCE}" xmlns:w14="{W14}" mc:Ignorable="w14"><w:trackRevisions w:val="1"/><w14:docId w14:val="7F00AA10"/></w:settings>"#
    )
    .into_bytes();
    let source = fixture_with_settings(document_xml("before", "second"), settings, false);

    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "managed changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut managed_output = Vec::new();
    let result = package.publish_document_commit_to_stream(&mut managed_output, &commit);
    assert!(result.as_ref().is_err_and(is_protection_refusal));
    assert!(managed_output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);

    let unmanaged_package =
        source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
    let mut edit = unmanaged_package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "unmanaged changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut unmanaged_output = Vec::new();
    let result =
        unmanaged_package.publish_document_commit_to_stream(&mut unmanaged_output, &commit);
    assert!(result.as_ref().is_err_and(is_protection_refusal));
    assert!(unmanaged_output.is_empty());
}

#[test]
fn changed_publication_rejects_unresolved_ignorable_and_active_mce_controls() {
    let cases: [(&str, Vec<u8>, fn(&TransactionError) -> bool); 5] = [
        (
            "unresolved Ignorable prefix",
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:mc="{MCE}" mc:Ignorable="u"><w:docVars/></w:settings>"#
            )
                .into_bytes(),
            is_settings_policy_refusal,
        ),
        (
            "ProcessContent",
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:mc="{MCE}" xmlns:u="urn:future" mc:ProcessContent="u:opaque"><w:docVars/></w:settings>"#
            )
                .into_bytes(),
            is_mce_policy_refusal,
        ),
        (
            "AlternateContent",
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:mc="{MCE}"><mc:AlternateContent><mc:Fallback><w:docVars/></mc:Fallback></mc:AlternateContent></w:settings>"#
            )
                .into_bytes(),
            is_mce_policy_refusal,
        ),
        (
            "entity-escaped MCE namespace with ProcessContent",
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/200&#54;" xmlns:u="urn:future" mc:ProcessContent="u:opaque"><w:docVars/></w:settings>"#
            )
                .into_bytes(),
            is_literal_namespace_refusal,
        ),
        (
            "entity-escaped Word namespace protection alias",
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/200&#54;/main"><x:documentProtection x:enforcement="1"/></w:settings>"#
            )
                .into_bytes(),
            is_literal_namespace_refusal,
        ),
    ];
    for (label, settings, refusal) in cases {
        let source = fixture_with_settings(document_xml("before", "second"), settings, false);
        let (budget, _cancellation_source, package) = managed(source, 1 << 20);
        let mut edit = package.edit_document().unwrap();
        edit.replace_paragraph_text(Position::new(0), "changed")
            .unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        let result = package.publish_document_commit_to_stream(&mut output, &commit);
        assert!(
            result.as_ref().is_err_and(refusal),
            "{label} must refuse changed publication: {:?}",
            result.as_ref().err()
        );
        assert!(output.is_empty(), "{label} must refuse before output");
        drop(commit);
        assert_eq!(budget.used(Resource::Memory), 0, "{label} leaked budget");
    }
}

#[test]
fn managed_changed_publication_rejects_mixed_settings_dialect_before_output() {
    let strict_settings = settings_xml(STRICT_W, r#"<w:docVars/>"#);
    let source = fixture_with_settings(document_xml("before", "second"), strict_settings, true);
    let (budget, _cancellation_source, package) = managed(source, 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut output = Vec::new();
    let result = package.publish_document_commit_to_stream(&mut output, &commit);
    assert!(result.as_ref().is_err_and(is_mixed_dialect_refusal));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_changed_publication_rejects_settings_mce_and_dtd_before_output() {
    let cases: [(&str, Vec<u8>, fn(&TransactionError) -> bool); 2] = [
        (
            "settings MCE",
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:mc="{MCE}"><mc:AlternateContent><mc:Choice Requires="u"><w:docVars/></mc:Choice><mc:Fallback><w:docVars/></mc:Fallback></mc:AlternateContent></w:settings>"#
            )
            .into_bytes(),
            is_mce_policy_refusal,
        ),
        (
            "settings DTD",
            format!(
                r#"<!DOCTYPE w:settings [<!ELEMENT settings ANY>]><w:settings xmlns:w="{W}"/>"#
            )
            .into_bytes(),
            is_dtd_policy_refusal,
        ),
    ];
    for (label, settings, refusal) in cases {
        let source = fixture_with_settings(document_xml("before", "second"), settings, false);
        let (budget, _cancellation_source, package) = managed(source, 1 << 20);
        let mut edit = package.edit_document().unwrap();
        edit.replace_paragraph_text(Position::new(0), "changed")
            .unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        let result = package.publish_document_commit_to_stream(&mut output, &commit);
        assert!(
            result.as_ref().is_err_and(refusal),
            "{label} must refuse changed publication: {:?}",
            result.as_ref().err()
        );
        assert!(output.is_empty(), "{label} must refuse before output");
        drop(commit);
        assert_eq!(budget.used(Resource::Memory), 0, "{label} leaked budget");
    }

    let escaped_cases = [
        (
            "unmanaged entity-escaped MCE namespace",
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/200&#54;" xmlns:u="urn:future" mc:ProcessContent="u:opaque"><w:docVars/></w:settings>"#
            )
            .into_bytes(),
        ),
        (
            "unmanaged entity-escaped Word namespace",
            format!(
                r#"<w:settings xmlns:w="{W}" xmlns:x="http://schemas.openxmlformats.org/wordprocessingml/200&#54;/main"><x:documentProtection x:enforcement="1"/></w:settings>"#
            )
            .into_bytes(),
        ),
    ];
    for (label, settings) in escaped_cases {
        let source = fixture_with_settings(document_xml("before", "second"), settings, false);
        let package =
            source_backed::Package::from_read_at(Arc::new(OwnedSource::new(source))).unwrap();
        let mut edit = package.edit_document().unwrap();
        edit.replace_paragraph_text(Position::new(0), "changed")
            .unwrap();
        let commit = edit.commit().unwrap();
        let mut output = Vec::new();
        let result = package.publish_document_commit_to_stream(&mut output, &commit);
        assert!(
            result.as_ref().is_err_and(is_literal_namespace_refusal),
            "{label} must refuse changed publication: {:?}",
            result.as_ref().err()
        );
        assert!(output.is_empty(), "{label} must refuse before output");
    }
}

#[test]
fn managed_source_document_rejects_mce_and_dtd_before_edit() {
    let cases: [(&str, Vec<u8>, fn(&TransactionError) -> bool); 2] = [
        (
            "document MCE",
            format!(
                r#"<w:document xmlns:w="{W}" xmlns:mc="{MCE}"><w:body><mc:AlternateContent/></w:body></w:document>"#
            )
            .into_bytes(),
            is_mce_document_refusal,
        ),
        (
            "document DTD",
            format!(
                r#"<!DOCTYPE w:document [<!ELEMENT document ANY>]><w:document xmlns:w="{W}"><w:body/></w:document>"#
            )
            .into_bytes(),
            is_dtd_document_refusal,
        ),
    ];
    for (label, document, refusal) in cases {
        let (budget, _cancellation_source, package) =
            managed(fixture_with_document(document, 37, 13), 1 << 20);
        let result = package.document_snapshot();
        assert!(
            result.as_ref().is_err_and(refusal),
            "{label} must refuse source capture: {:?}",
            result.as_ref().err()
        );
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0, "{label} leaked budget");
    }
}

#[test]
fn managed_reopened_inverse_restores_exact_artifact_and_rejects_foreign_member() {
    let source = fixture("before", "second");
    let (changed, publication, budget, commit) =
        publish_managed_edit_with_inverse(source.clone(), "after");
    assert_eq!(
        publication
            .snapshot()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "after"
    );

    let changed_document = part_bytes(&changed, MAIN);
    let foreign = fixture_with_document(changed_document, 37, 29);
    assert_eq!(part_bytes(&foreign, MAIN), part_bytes(&changed, MAIN));
    assert_ne!(foreign, changed);
    let (foreign_budget, _cancellation_source, foreign_package) = managed(foreign, 1 << 20);
    let mut foreign_output = Vec::new();
    let result =
        foreign_package.publish_document_inverse_to_stream(&mut foreign_output, &publication);
    assert!(result.as_ref().is_err_and(is_artifact_mismatch));
    assert!(foreign_output.is_empty());
    assert_eq!(foreign_budget.used(Resource::Memory), 0);

    let (inverse_budget, _cancellation_source, reopened) = managed(changed, 1 << 20);
    let mut restored = Vec::new();
    reopened
        .publish_document_inverse_to_stream(&mut restored, &publication)
        .unwrap();
    assert_eq!(restored, source);
    assert_eq!(inverse_budget.used(Resource::Memory), 0);

    drop(publication);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_reopened_inverse_rejects_source_mutation_before_output() {
    let source = fixture("before", "second");
    let (changed, publication, budget, commit) = publish_managed_edit_with_inverse(source, "after");
    let mutable = Arc::new(MutableSource::new(changed));
    let (inverse_budget, _cancellation_source, context) = managed_context(1 << 20);
    let package = source_backed::Package::from_read_at_with_execution_context(
        mutable.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    mutable.bump_revision();

    let mut output = Vec::new();
    let result = package.publish_document_inverse_to_stream(&mut output, &publication);
    assert!(result.as_ref().is_err_and(is_source_changed));
    assert!(output.is_empty());
    assert_eq!(inverse_budget.used(Resource::Memory), 0);
    drop(publication);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_reopened_inverse_cancellation_and_output_budget_are_before_output() {
    let source = fixture("before", "second");
    let (changed, publication, budget, commit) = publish_managed_edit_with_inverse(source, "after");
    let (inverse_budget, cancellation_source, package) = managed(changed.clone(), 1 << 20);
    cancellation_source.cancel();
    let mut cancelled_output = Vec::new();
    let result = package.publish_document_inverse_to_stream(&mut cancelled_output, &publication);
    assert!(result.as_ref().is_err_and(is_cancelled));
    assert!(cancelled_output.is_empty());
    assert_eq!(inverse_budget.used(Resource::Memory), 0);

    let (limited_budget, _cancellation_source, context) = managed_context_with_output(1 << 20, 1);
    let limited_package = source_backed::Package::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(changed)),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut limited_output = Vec::new();
    let result =
        limited_package.publish_document_inverse_to_stream(&mut limited_output, &publication);
    assert!(result.as_ref().is_err_and(is_output_budget_refusal));
    assert!(limited_output.is_empty());
    assert_eq!(limited_budget.used(Resource::Memory), 0);

    drop(publication);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_reopened_inverse_reports_partial_sink_failure() {
    let source = fixture("before", "second");
    let (changed, publication, budget, commit) = publish_managed_edit_with_inverse(source, "after");
    let (inverse_budget, _cancellation_source, package) = managed(changed, 1 << 20);
    let mut sink = PartialWriter {
        bytes: Vec::new(),
        limit: 1,
    };
    let result = package.publish_document_inverse_to_stream(&mut sink, &publication);
    assert!(matches!(
        result,
        Err(TransactionError::Document(Error::Opc(
            OpcError::IncompleteOutput { written, .. }
        ))) if written > 0
    ));
    assert!(!sink.bytes.is_empty());
    assert_eq!(inverse_budget.used(Resource::Memory), 0);
    drop(publication);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_file_source_pins_the_open_source_across_path_replacement() {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("source.docx");
    let replacement_path = directory.path().join("replacement.docx");
    let original = fixture("original", "second");
    let replacement = fixture("replacement", "second");
    std::fs::write(&path, &original).unwrap();
    std::fs::write(&replacement_path, &replacement).unwrap();

    let (budget, _cancellation_source, context) = managed_context(1 << 20);
    let package = source_backed::Package::from_path_with_execution_context(
        &path,
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let version = package.source_version().unwrap();
    std::fs::remove_file(&path).unwrap();
    std::fs::rename(&replacement_path, &path).unwrap();

    assert_eq!(package.source_version().unwrap(), version);
    assert_eq!(
        package
            .document_snapshot()
            .unwrap()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "original tail"
    );
    assert_ne!(std::fs::read(&path).unwrap(), original);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_source_change_refuses_publication_before_sink_output() {
    let source_bytes = fixture("before", "second");
    let source = Arc::new(MutableSource::new(source_bytes));
    let (budget, _cancellation_source, context) = managed_context(1 << 20);
    let package = source_backed::Package::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    source.bump_revision();

    let mut output = Vec::new();
    let result = package.publish_document_commit_to_stream(&mut output, &commit);
    assert!(result.as_ref().is_err_and(is_source_changed));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_refuses_edit_and_publication_without_output() {
    let source = fixture("before", "second");
    let (budget, cancellation_source, package) = managed(source.clone(), 1 << 20);
    cancellation_source.cancel();
    assert!(is_cancelled(&package.edit_document().unwrap_err()));
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, package) = managed(source, 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    cancellation_source.cancel();
    let mut output = Vec::new();
    let result = package.publish_document_commit_to_stream(&mut output, &commit);
    assert!(result.as_ref().is_err_and(is_cancelled));
    assert!(output.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_after_capture_refuses_empty_commit_and_releases_memory() {
    let (budget, cancellation_source, package) = managed(fixture("before", "second"), 1 << 20);
    let edit = package.edit_document().unwrap();
    assert!(budget.used(Resource::Memory) > 0);
    cancellation_source.cancel();

    let result = edit.commit();
    assert!(result.as_ref().is_err_and(is_cancelled));
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_empty_text_slots_refuse_low_scan_memory_before_projection_change() {
    let source = fixture_with_document(empty_text_slot_document(4096), 37, 13);
    let (probe_budget, _cancellation_source, probe_package) = managed(source.clone(), 1 << 24);
    let probe_edit = probe_package.edit_document().unwrap();
    let baseline_memory = probe_budget.used(Resource::Memory);
    assert!(baseline_memory > 0);
    drop(probe_edit);
    drop(probe_package);
    assert_eq!(probe_budget.used(Resource::Memory), 0);

    let low_memory_limit = baseline_memory.saturating_add(2 * 1024 * 1024);
    let (budget, _cancellation_source, package) = managed(source, low_memory_limit);
    let mut edit = package.edit_document().unwrap();
    let before = edit.projected().xml_bytes().to_vec();
    let result = edit.replace_paragraph_text(Position::new(0), "x");
    assert!(result.as_ref().is_err_and(is_budget_refusal));
    assert_eq!(edit.projected().xml_bytes(), before.as_slice());
    drop(edit);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_candidate_budget_failure_happens_before_publication_output() {
    let source = fixture("before", "second");
    let (budget, _cancellation_source, package) = managed(source, 1 << 20);
    let mut edit = package.edit_document().unwrap();
    let baseline_memory = budget.used(Resource::Memory);
    assert!(baseline_memory > 0);
    let remaining = budget
        .limit(Resource::Memory)
        .saturating_sub(baseline_memory);
    assert!(remaining > 0);
    let hold = budget.reserve(Resource::Memory, remaining).unwrap();
    let result = edit.replace_paragraph_text(Position::new(0), "candidate");
    let output = Vec::<u8>::new();
    assert!(result.as_ref().is_err_and(is_budget_refusal));
    assert!(output.is_empty());
    drop(hold);
    drop(edit);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_publication_reports_a_partial_sink_failure() {
    let source = fixture("before", "second");
    let (budget, _cancellation_source, package) = managed(source, 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    let mut sink = PartialWriter {
        bytes: Vec::new(),
        limit: 1,
    };
    let result = package.publish_document_commit_to_stream(&mut sink, &commit);
    assert!(matches!(
        result,
        Err(TransactionError::Document(Error::Opc(
            OpcError::IncompleteOutput { written, .. }
        ))) if written > 0
    ));
    assert!(!sink.bytes.is_empty());
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_changed_publication_preserves_readback_and_releases_snapshot_budget() {
    let source = fixture("before", "second");
    let original_main = part_bytes(&source, MAIN);
    let original_media = part_bytes(&source, MEDIA);
    let original_opaque = part_bytes(&source, OPAQUE);
    let (budget, _cancellation_source, package) = managed(source.clone(), 1 << 20);
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    assert_eq!(
        commit.patch().source().xml_bytes(),
        original_main.as_slice()
    );
    assert_eq!(
        commit
            .snapshot()
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "changed"
    );
    assert!(budget.used(Resource::Memory) > 0);

    let mut output = Vec::new();
    let published = package
        .publish_document_commit_to_stream(&mut output, &commit)
        .unwrap();
    assert_eq!(
        published
            .paragraph(Position::new(0))
            .unwrap()
            .text()
            .unwrap(),
        "changed"
    );
    assert_eq!(part_bytes(&output, MAIN), published.xml_bytes());
    assert_eq!(part_bytes(&output, MEDIA), original_media);
    assert_eq!(part_bytes(&output, OPAQUE), original_opaque);
    assert!(budget.used(Resource::Memory) > 0);

    drop(published);
    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_foreign_equal_main_xml_refuses_commit_before_output() {
    let main_xml = document_xml("before", "second");
    let first = fixture_with_document(main_xml.clone(), 37, 13);
    let second = fixture_with_document(main_xml, 41, 29);
    assert_eq!(part_bytes(&first, MAIN), part_bytes(&second, MAIN));
    assert_ne!(part_bytes(&first, MEDIA), part_bytes(&second, MEDIA));
    assert_ne!(part_bytes(&first, OPAQUE), part_bytes(&second, OPAQUE));

    let (first_budget, _first_cancellation, first_package) = managed(first, 1 << 20);
    let mut edit = first_package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "first package change")
        .unwrap();
    let commit = edit.commit().unwrap();
    assert!(commit.diagnostics().changed());
    drop(first_package);

    let (second_budget, _second_cancellation, second_package) = managed(second, 1 << 20);
    let mut output = Vec::new();
    let result = second_package.publish_document_commit_to_stream(&mut output, &commit);
    assert!(matches!(result, Err(TransactionError::StaleSource)));
    assert!(output.is_empty());

    drop(commit);
    assert_eq!(first_budget.used(Resource::Memory), 0);
    assert_eq!(second_budget.used(Resource::Memory), 0);
}

#[test]
fn managed_source_change_during_current_part_read_refuses_before_output() {
    let source_bytes = fixture("before", "second");
    let source = Arc::new(MutableSource::new(source_bytes));
    let (budget, _cancellation_source, context) = managed_context(1 << 20);
    let cache_limits = SourceCacheLimits::new(1, 1).unwrap();
    let package =
        source_backed::Package::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source.clone(),
            ReadLimits::default(),
            cache_limits,
            context,
        )
        .unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    source.arm_revision_on_read();

    let mut output = Vec::new();
    let result = package.publish_document_commit_to_stream(&mut output, &commit);
    assert!(result.as_ref().is_err_and(is_source_changed));
    assert!(output.is_empty());

    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_during_current_part_read_refuses_before_output() {
    let source_bytes = fixture("before", "second");
    let source = Arc::new(MutableSource::new(source_bytes));
    let (budget, cancellation_source, context) = managed_context(1 << 20);
    let cache_limits = SourceCacheLimits::new(1, 1).unwrap();
    let package =
        source_backed::Package::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source.clone(),
            ReadLimits::default(),
            cache_limits,
            context,
        )
        .unwrap();
    let mut edit = package.edit_document().unwrap();
    edit.replace_paragraph_text(Position::new(0), "changed")
        .unwrap();
    let commit = edit.commit().unwrap();
    source.arm_cancellation_on_read(cancellation_source);

    let mut output = Vec::new();
    let result = package.publish_document_commit_to_stream(&mut output, &commit);
    assert!(result.as_ref().is_err_and(is_cancelled));
    assert!(output.is_empty());

    drop(commit);
    assert_eq!(budget.used(Resource::Memory), 0);
}
