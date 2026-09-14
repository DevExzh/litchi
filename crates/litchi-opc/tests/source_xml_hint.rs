#![allow(
    clippy::cast_possible_truncation,
    clippy::cast_sign_loss,
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used,
    reason = "focused source-proof tests deliberately panic on fixture errors"
)]

//! Focused coverage for `PartView::source_xml_with_hint`.
//!
//! The hint is an optimization aid, not an authority transfer.  These tests
//! keep the public boundary honest: a matching original proof can reuse its
//! exact byte allocation, while foreign, derived, stale, or differently
//! bounded proofs follow the existing source-validation and guard paths.

use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, RwLock,
    atomic::{AtomicBool, AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::{
    AuthoredXmlFragment, OpcError, PackURI, ReadLimits, SourceBackedPackage, SourceCacheLimits,
};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SIGNATURE_ORIGIN_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const DOCUMENT_URI: &str = "/word/document.xml";
const OTHER_MEMBER: &str = "custom/other.bin";
const OTHER_URI: &str = "/custom/other.bin";
const SIGNATURE_MEMBER: &str = "signature/origin.xml";

const FORMATTED_DOCUMENT: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?>
<document>
  <before attr="one"/>
  <!--keep-->
  <nested>
    <leaf/>
  </nested>
</document>
"#;
const OTHER_DOCUMENT: &[u8] = b"<document><changed/></document>";
const MALFORMED_DOCUMENT: &[u8] = b"<document><unclosed></document>";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).expect("fixture URI must be valid")
}

fn archive_bytes(document: &[u8]) -> Vec<u8> {
    archive_with_options(document, "application/xml", None, false)
}

fn archive_with_other(document: &[u8], other_content_type: &str) -> Vec<u8> {
    archive_with_options(document, "application/xml", Some(other_content_type), false)
}

fn archive_with_options(
    document: &[u8],
    document_content_type: &str,
    other_content_type: Option<&str>,
    signed: bool,
) -> Vec<u8> {
    let other_override = other_content_type.map_or_else(String::new, |content_type| {
        format!(r#"<Override PartName="/{OTHER_MEMBER}" ContentType="{content_type}"/>"#)
    });
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/{DOCUMENT_MEMBER}" ContentType="{document_content_type}"/>{other_override}</Types>"#
    );
    let signature_relationship = if signed {
        format!(
            r#"<Relationship Id="rSig" Type="{SIGNATURE_ORIGIN_REL}" Target="{SIGNATURE_MEMBER}"/>"#
        )
    } else {
        String::new()
    };
    let other_relationship = other_content_type.map_or_else(String::new, |_| {
        format!(
            r#"<Relationship Id="rOther" Type="urn:litchi:test/other" Target="{OTHER_MEMBER}"/>"#
        )
    });
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rDoc" Type="{OFFICE_DOCUMENT_REL}" Target="{DOCUMENT_MEMBER}"/>{other_relationship}{signature_relationship}</Relationships>"#
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
    if other_content_type.is_some() {
        writer
            .write_stored(OTHER_MEMBER, OTHER_DOCUMENT)
            .expect("other XML fixture must be writable");
    }
    if signed {
        writer
            .write_stored(SIGNATURE_MEMBER, b"signature-origin")
            .expect("signature fixture must be writable");
    }
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
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

fn managed_package(
    source: Arc<dyn ReadAt>,
    limits: ReadLimits,
    scope: &'static str,
) -> (SourceBackedPackage, Budget, CancellationSource) {
    managed_package_with_cache_limits(source, limits, SourceCacheLimits::default(), scope)
}

fn managed_package_with_cache_limits(
    source: Arc<dyn ReadAt>,
    limits: ReadLimits,
    cache_limits: SourceCacheLimits,
    scope: &'static str,
) -> (SourceBackedPackage, Budget, CancellationSource) {
    let (budget, cancellation_source, context) = managed_context(scope, 128 * 1024 * 1024, 1 << 40);
    let package =
        SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
            source,
            limits,
            cache_limits,
            context,
        )
        .expect("managed fixture package must open");
    (package, budget, cancellation_source)
}

fn owned_package(
    archive: Vec<u8>,
    limits: ReadLimits,
    scope: &'static str,
) -> (SourceBackedPackage, Budget, CancellationSource) {
    managed_package(Arc::new(OwnedSource::new(archive)), limits, scope)
}

/// Separate package instances can intentionally report one equal caller
/// version.  `SourceLineage` must still keep their proofs isolated.
struct FixedVersionSource {
    source: OwnedSource,
    version: SourceVersion,
}

impl FixedVersionSource {
    fn new(bytes: Vec<u8>, version: SourceVersion) -> Self {
        Self {
            source: OwnedSource::new(bytes),
            version,
        }
    }
}

impl ReadAt for FixedVersionSource {
    fn len(&self) -> io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.source.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(self.version)
    }
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
            0x4f50_435f_4849_4e54,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

struct RevisionDuringRead {
    source: OwnedSource,
    revision: AtomicU64,
    armed: AtomicBool,
}

impl RevisionDuringRead {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            source: OwnedSource::new(bytes),
            revision: AtomicU64::new(0),
            armed: AtomicBool::new(false),
        }
    }

    fn arm(&self) {
        self.armed.store(true, Ordering::Release);
    }
}

impl ReadAt for RevisionDuringRead {
    fn len(&self) -> io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let read = self.source.read_at(offset, output)?;
        if self.armed.swap(false, Ordering::AcqRel) {
            self.revision.fetch_add(1, Ordering::Release);
        }
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4f50_435f_5245_4144,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

struct CancellationDuringRead {
    source: OwnedSource,
    cancellation: CancellationSource,
    armed: AtomicBool,
}

impl CancellationDuringRead {
    fn new(bytes: Vec<u8>, cancellation: CancellationSource) -> Self {
        Self {
            source: OwnedSource::new(bytes),
            cancellation,
            armed: AtomicBool::new(false),
        }
    }

    fn arm(&self) {
        self.armed.store(true, Ordering::Release);
    }
}

impl ReadAt for CancellationDuringRead {
    fn len(&self) -> io::Result<u64> {
        self.source.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let read = self.source.read_at(offset, output)?;
        if self.armed.swap(false, Ordering::AcqRel) {
            self.cancellation.cancel();
        }
        Ok(read)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.source.version()
    }
}

fn mark_entry_encrypted(mut bytes: Vec<u8>, wanted: &str) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .expect("fixture must contain ZIP32 EOCD");
    let count = usize::from(u16::from_le_bytes(
        bytes[eocd + 10..eocd + 12]
            .try_into()
            .expect("EOCD count must fit"),
    ));
    let mut central = usize::try_from(u32::from_le_bytes(
        bytes[eocd + 16..eocd + 20]
            .try_into()
            .expect("EOCD offset must fit"),
    ))
    .expect("central offset must fit");
    for _ in 0..count {
        assert_eq!(&bytes[central..central + 4], &0x0201_4b50_u32.to_le_bytes());
        let name_len = usize::from(u16::from_le_bytes(
            bytes[central + 28..central + 30]
                .try_into()
                .expect("central name length must fit"),
        ));
        let extra_len = usize::from(u16::from_le_bytes(
            bytes[central + 30..central + 32]
                .try_into()
                .expect("central extra length must fit"),
        ));
        let comment_len = usize::from(u16::from_le_bytes(
            bytes[central + 32..central + 34]
                .try_into()
                .expect("central comment length must fit"),
        ));
        let name_start = central + 46;
        let name = &bytes[name_start..name_start + name_len];
        let local = usize::try_from(u32::from_le_bytes(
            bytes[central + 42..central + 46]
                .try_into()
                .expect("local offset must fit"),
        ))
        .expect("local offset must fit");
        if name == wanted.as_bytes() {
            let central_flags = u16::from_le_bytes(
                bytes[central + 8..central + 10]
                    .try_into()
                    .expect("central flags must fit"),
            ) | 1;
            bytes[central + 8..central + 10].copy_from_slice(&central_flags.to_le_bytes());
            let local_flags = u16::from_le_bytes(
                bytes[local + 6..local + 8]
                    .try_into()
                    .expect("local flags must fit"),
            ) | 1;
            bytes[local + 6..local + 8].copy_from_slice(&local_flags.to_le_bytes());
            return bytes;
        }
        central += 46 + name_len + extra_len + comment_len;
    }
    panic!("missing ZIP member {wanted}");
}

#[test]
fn matching_original_hint_reuses_exact_bytes_without_extra_budget_work() {
    let archive = archive_bytes(FORMATTED_DOCUMENT);

    let (cold_package, cold_budget, _cold_cancellation) =
        owned_package(archive.clone(), ReadLimits::default(), "opc-hint-cold");
    let cold_part = cold_package.part(&pack(DOCUMENT_URI)).unwrap();
    // Warm the bounded payload cache so the comparison isolates the source
    // XML validation work that the hint is intended to avoid.
    let cold_data = cold_part.data().unwrap();
    drop(cold_data);
    let cold_before = cold_budget.used(Resource::Work);
    let cold = cold_part.source_xml().unwrap();
    let cold_work = cold_budget.used(Resource::Work) - cold_before;
    assert_eq!(cold.bytes(), FORMATTED_DOCUMENT);

    let (package, budget, _cancellation) =
        owned_package(archive, ReadLimits::default(), "opc-hint-hit");
    let part = package.part(&pack(DOCUMENT_URI)).unwrap();
    let hint = part.source_xml().unwrap();
    let memory_with_hint = budget.used(Resource::Memory);
    let work_before = budget.used(Resource::Work);
    let reused = part
        .source_xml_with_hint(&hint)
        .expect("matching original hint must be accepted");
    let hit_work = budget.used(Resource::Work) - work_before;

    assert_eq!(reused.bytes(), FORMATTED_DOCUMENT);
    assert_eq!(reused.partname(), hint.partname());
    assert_eq!(reused.content_type(), hint.content_type());
    assert_eq!(
        reused.bytes().as_ptr(),
        hint.bytes().as_ptr(),
        "a hit must retain the original immutable byte allocation"
    );
    assert_eq!(
        budget.used(Resource::Memory),
        memory_with_hint,
        "cloning a hit must retain metadata and payload reservations"
    );
    assert_eq!(
        hit_work, cold_work,
        "the hint's conservative byte comparison charge should equal the cold source-byte work"
    );

    drop(reused);
    drop(hint);
    drop(package);
    drop(cold);
    drop(cold_package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
    assert_eq!(cold_budget.used(Resource::Memory), 0);
    assert_eq!(cold_budget.used(Resource::Objects), 0);
}

#[test]
fn equal_version_foreign_lineage_falls_back_to_current_original() {
    let version = SourceVersion::new(0x0046_4f52_4549_474e, 7);
    let (first, first_budget, _first_cancellation) = managed_package(
        Arc::new(FixedVersionSource::new(
            archive_bytes(FORMATTED_DOCUMENT),
            version,
        )),
        ReadLimits::default(),
        "opc-hint-foreign-first",
    );
    let (second, second_budget, _second_cancellation) = managed_package(
        Arc::new(FixedVersionSource::new(
            archive_bytes(FORMATTED_DOCUMENT),
            version,
        )),
        ReadLimits::default(),
        "opc-hint-foreign-second",
    );
    assert_eq!(
        first.source_version().unwrap(),
        second.source_version().unwrap()
    );
    assert_ne!(first.source_lineage(), second.source_lineage());

    let hint = first
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml()
        .unwrap();
    let target_part = second.part(&pack(DOCUMENT_URI)).unwrap();
    let before = second_budget.used(Resource::Work);
    let current = target_part
        .source_xml_with_hint(&hint)
        .expect("foreign hint must fall back to the current source Part");
    let work = second_budget.used(Resource::Work) - before;

    assert_eq!(current.bytes(), FORMATTED_DOCUMENT);
    assert_ne!(
        current.bytes().as_ptr(),
        hint.bytes().as_ptr(),
        "foreign lineage must not return the hint allocation"
    );
    assert!(
        work > 0,
        "foreign fallback must perform bounded current work"
    );

    drop(current);
    drop(hint);
    drop(first);
    drop(second);
    assert_eq!(first_budget.used(Resource::Memory), 0);
    assert_eq!(second_budget.used(Resource::Memory), 0);
}

#[test]
fn different_part_and_content_type_hints_return_the_target_original() {
    let archive = archive_with_other(FORMATTED_DOCUMENT, "application/vnd.example.other+xml");
    let (package, budget, _cancellation) =
        owned_package(archive, ReadLimits::default(), "opc-hint-part");
    let document = package.part(&pack(DOCUMENT_URI)).unwrap();
    let other = package.part(&pack(OTHER_URI)).unwrap();
    let hint = document.source_xml().unwrap();
    let target = other
        .source_xml_with_hint(&hint)
        .expect("different Part must use the target source bytes");

    assert_eq!(target.bytes(), OTHER_DOCUMENT);
    assert_eq!(target.partname(), &pack(OTHER_URI));
    assert_eq!(target.content_type(), "application/vnd.example.other+xml");
    assert_ne!(target.bytes(), hint.bytes());
    drop(target);
    drop(hint);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn different_content_type_hint_falls_back_even_with_equal_source_version() {
    let version = SourceVersion::new(0x4354_5950_4543_484b, 11);
    let first_archive = archive_with_options(
        FORMATTED_DOCUMENT,
        "application/vnd.example.first+xml",
        None,
        false,
    );
    let second_archive = archive_with_options(
        FORMATTED_DOCUMENT,
        "application/vnd.example.second+xml",
        None,
        false,
    );
    let (first, first_budget, _first_cancellation) = managed_package(
        Arc::new(FixedVersionSource::new(first_archive, version)),
        ReadLimits::default(),
        "opc-hint-content-type-first",
    );
    let (second, second_budget, _second_cancellation) = managed_package(
        Arc::new(FixedVersionSource::new(second_archive, version)),
        ReadLimits::default(),
        "opc-hint-content-type-second",
    );
    let hint = first
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml()
        .unwrap();
    let target = second
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml_with_hint(&hint)
        .expect("different content type must use a fresh current proof");

    assert_eq!(target.bytes(), FORMATTED_DOCUMENT);
    assert_eq!(target.content_type(), "application/vnd.example.second+xml");
    assert_ne!(target.bytes().as_ptr(), hint.bytes().as_ptr());

    drop(target);
    drop(hint);
    drop(first);
    drop(second);
    assert_eq!(first_budget.used(Resource::Memory), 0);
    assert_eq!(second_budget.used(Resource::Memory), 0);
}

#[test]
fn different_limits_hint_falls_back_to_validation_under_current_limits() {
    let changed_limits = ReadLimits::builder()
        .max_xml_events(999_999)
        .expect("changed event limit must be valid")
        .build()
        .expect("changed limits must build");
    assert_ne!(changed_limits, ReadLimits::default());
    let version = SourceVersion::new(0x4c49_4d49_5453_484e, 3);
    let archive = archive_bytes(FORMATTED_DOCUMENT);
    let (first, first_budget, _first_cancellation) = managed_package(
        Arc::new(FixedVersionSource::new(archive.clone(), version)),
        ReadLimits::default(),
        "opc-hint-limits-first",
    );
    let (second, second_budget, _second_cancellation) = managed_package(
        Arc::new(FixedVersionSource::new(archive, version)),
        changed_limits,
        "opc-hint-limits-second",
    );
    let hint = first
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml()
        .unwrap();
    let target = second
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml_with_hint(&hint)
        .expect("different limits must use a current proof");

    assert_eq!(target.bytes(), FORMATTED_DOCUMENT);
    assert_ne!(target.bytes().as_ptr(), hint.bytes().as_ptr());

    drop(target);
    drop(hint);
    drop(first);
    drop(second);
    assert_eq!(first_budget.used(Resource::Memory), 0);
    assert_eq!(second_budget.used(Resource::Memory), 0);
}

#[test]
fn derived_hint_returns_current_original_and_keeps_derived_bytes_live() {
    let (package, budget, _cancellation) = owned_package(
        archive_bytes(FORMATTED_DOCUMENT),
        ReadLimits::default(),
        "opc-hint-derived",
    );
    let part = package.part(&pack(DOCUMENT_URI)).unwrap();
    let source = part.source_xml().unwrap();
    let insertion = source
        .bytes()
        .windows(b"</document>".len())
        .position(|window| window == b"</document>")
        .expect("fixture close tag must exist");
    let proof = source
        .checked_range(insertion..insertion, &[])
        .expect("insertion proof must be valid");
    let mut publication = source.into_publication().unwrap();
    publication
        .replace(
            proof,
            AuthoredXmlFragment::markup(b"<derived/>".to_vec()).unwrap(),
        )
        .unwrap();
    let derived = publication.finish().unwrap();
    assert_ne!(derived.bytes(), FORMATTED_DOCUMENT);
    let memory_with_derived = budget.used(Resource::Memory);

    let current = part
        .source_xml_with_hint(&derived)
        .expect("derived hint must fall back to the current original source");
    assert_eq!(current.bytes(), FORMATTED_DOCUMENT);
    assert_ne!(current.bytes(), derived.bytes());
    assert!(
        budget.used(Resource::Memory) >= memory_with_derived,
        "the derived token's original and edit reservations must remain live"
    );

    drop(current);
    assert_eq!(budget.used(Resource::Memory), memory_with_derived);
    drop(derived);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn malformed_current_source_is_revalidated_after_foreign_hint_mismatch() {
    let hint_package = SourceBackedPackage::from_vec(archive_bytes(FORMATTED_DOCUMENT)).unwrap();
    let hint = hint_package
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml()
        .unwrap();
    let current = SourceBackedPackage::from_vec(archive_bytes(MALFORMED_DOCUMENT)).unwrap();
    let error = current
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml_with_hint(&hint)
        .expect_err("a mismatched malformed Part must take the full validator path");
    assert!(
        matches!(error, OpcError::XmlError(_)),
        "unexpected malformed-source error: {error:?}"
    );
}

#[test]
fn source_revision_change_before_hint_read_is_refused() {
    let source = Arc::new(MutableSource::new(archive_bytes(FORMATTED_DOCUMENT)));
    let (package, budget, _cancellation) = managed_package(
        source.clone(),
        ReadLimits::default(),
        "opc-hint-revision-before",
    );
    let part = package.part(&pack(DOCUMENT_URI)).unwrap();
    let hint = part.source_xml().unwrap();
    source.replace_bytes(archive_bytes(FORMATTED_DOCUMENT));
    let error = part
        .source_xml_with_hint(&hint)
        .expect_err("a revised source must fail before hint reuse");
    assert!(matches!(error, OpcError::SourceChanged { .. }));
    drop(hint);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn source_revision_change_during_hint_read_is_refused_after_read() {
    let source = Arc::new(RevisionDuringRead::new(archive_bytes(FORMATTED_DOCUMENT)));
    let cache_limits = SourceCacheLimits::new(1, 1).expect("tiny cache limits must be valid");
    let (package, budget, _cancellation) = managed_package_with_cache_limits(
        source.clone(),
        ReadLimits::default(),
        cache_limits,
        "opc-hint-revision-during",
    );
    let part = package.part(&pack(DOCUMENT_URI)).unwrap();
    let hint = part.source_xml().unwrap();
    source.arm();
    let error = part
        .source_xml_with_hint(&hint)
        .expect_err("a source revision during the current read must be refused");
    assert!(matches!(error, OpcError::SourceChanged { .. }));
    drop(hint);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cancellation_before_and_during_hint_read_remains_typed_and_budgeted() {
    let (package, budget, cancellation) = owned_package(
        archive_bytes(FORMATTED_DOCUMENT),
        ReadLimits::default(),
        "opc-hint-cancel-before",
    );
    let part = package.part(&pack(DOCUMENT_URI)).unwrap();
    let hint = part.source_xml().unwrap();
    cancellation.cancel();
    let error = part
        .source_xml_with_hint(&hint)
        .expect_err("pre-cancelled hint read must stop");
    assert!(matches!(error, OpcError::Cancelled));
    drop(hint);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (during_budget, during_cancellation, context) =
        managed_context("opc-hint-cancel-during", 128 * 1024 * 1024, 1 << 40);
    let source = Arc::new(CancellationDuringRead::new(
        archive_bytes(FORMATTED_DOCUMENT),
        during_cancellation.clone(),
    ));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .expect("managed cancellation fixture must open");
    let foreign = SourceBackedPackage::from_vec(archive_bytes(FORMATTED_DOCUMENT)).unwrap();
    let hint = foreign
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml()
        .unwrap();
    let part = package.part(&pack(DOCUMENT_URI)).unwrap();
    source.arm();
    let error = part
        .source_xml_with_hint(&hint)
        .expect_err("cancellation observed during current read must stop");
    assert!(matches!(error, OpcError::Cancelled));
    drop(hint);
    drop(foreign);
    drop(package);
    assert_eq!(during_budget.used(Resource::Memory), 0);
}

#[test]
fn signed_and_encrypted_packages_refuse_hints_before_payload_transfer() {
    let hint_package = SourceBackedPackage::from_vec(archive_bytes(FORMATTED_DOCUMENT)).unwrap();
    let hint = hint_package
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml()
        .unwrap();

    let signed = SourceBackedPackage::from_vec(archive_with_options(
        FORMATTED_DOCUMENT,
        "application/xml",
        None,
        true,
    ))
    .expect("signed fixture must open");
    let signed_error = signed
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml_with_hint(&hint)
        .expect_err("signature infrastructure must remain policy-gated");
    assert!(matches!(
        signed_error,
        OpcError::SignedSourceRequiresExplicitPolicy
    ));

    let encrypted_archive =
        mark_entry_encrypted(archive_bytes(FORMATTED_DOCUMENT), DOCUMENT_MEMBER);
    let encrypted = SourceBackedPackage::from_vec(encrypted_archive)
        .expect("encrypted metadata fixture must open");
    let encrypted_error = encrypted
        .part(&pack(DOCUMENT_URI))
        .unwrap()
        .source_xml_with_hint(&hint)
        .expect_err("encrypted entries must remain transfer-refused");
    assert!(matches!(
        encrypted_error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
}
