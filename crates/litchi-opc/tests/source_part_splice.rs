#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "focused source-splice assertions intentionally panic on fixture errors"
)]

//! Contract coverage for bounded decoded-member source splicing.
//!
//! These tests deliberately construct both Store and Deflate members, retain
//! one untyped physical member, and keep all source assertions on the public
//! OPC boundary.  The splice API must never turn the selected Part into a
//! cache entry merely to publish it.

use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, RwLock,
    atomic::{AtomicBool, AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_opc::{
    OpcError, OpcOperationAccounting, PackURI, SourceBackedPackage, SourcePartSpliceLimits,
    SourcePartSpliceProof, SpliceResource,
};
use sha2::{Digest as _, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SIGNATURE_ORIGIN_REL: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const DOCUMENT_URI: &str = "/word/document.xml";
const BINARY_MEMBER: &str = "word/document.bin";
const BINARY_URI: &str = "/word/document.bin";
const SCRATCH_MEMBER: &str = "scratch.bin";
const SIGNATURE_MEMBER: &str = "signature/origin.xml";

const VALID_SOURCE: &[u8] = br#"<?xml version="1.0" encoding="UTF-8"?><root><before/></root>"#;
const INSERTION_OFFSET: usize = VALID_SOURCE
    .len()
    .checked_sub(b"</root>".len())
    .expect("fixture close tag must fit");
const FRAGMENT: &[u8] = b"<inserted/>";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).expect("fixture URI must be canonical")
}

fn digest(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

fn candidate_bytes(source: &[u8], offset: usize, fragment: &[u8]) -> Vec<u8> {
    [&source[..offset], fragment, &source[offset..]].concat()
}

fn content_types(_signed: bool) -> Vec<u8> {
    content_types_for(false)
}

fn content_types_for(binary_target: bool) -> Vec<u8> {
    let target_override = if binary_target {
        format!(r#"<Override PartName="{BINARY_URI}" ContentType="application/octet-stream"/>"#)
    } else {
        String::new()
    };
    format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{target_override}</Types>"#
    )
    .into_bytes()
}

fn package_relationships(signed: bool) -> Vec<u8> {
    package_relationships_for(DOCUMENT_MEMBER, signed)
}

fn package_relationships_for(document_member: &str, signed: bool) -> Vec<u8> {
    let signature_relationship = if signed {
        format!(
            r#"<Relationship Id="rSig" Type="{SIGNATURE_ORIGIN_REL}" Target="{SIGNATURE_MEMBER}"/>"#
        )
    } else {
        String::new()
    };
    format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rDoc" Type="{OFFICE_DOCUMENT_REL}" Target="{document_member}"/>{signature_relationship}</Relationships>"#
    )
    .into_bytes()
}

fn archive_bytes(source: &[u8], deflated: bool, signed: bool) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", &content_types(signed))
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", &package_relationships(signed))
        .expect("package relationships fixture must be writable");
    // No content type is declared for this member, so it remains an opaque
    // non-Part and can be checked for exact physical preservation.
    writer
        .write_stored(SCRATCH_MEMBER, b"opaque physical bytes\0\xff")
        .expect("opaque fixture must be writable");
    if deflated {
        writer
            .write_deflated_sized(DOCUMENT_MEMBER, source)
            .expect("Deflate fixture must be writable");
    } else {
        writer
            .write_stored(DOCUMENT_MEMBER, source)
            .expect("Store fixture must be writable");
    }
    if signed {
        writer
            .write_stored(SIGNATURE_MEMBER, b"signature-origin")
            .expect("signature fixture must be writable");
    }
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

fn binary_archive_bytes(source: &[u8], deflated: bool) -> Vec<u8> {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", &content_types_for(true))
        .expect("binary content types fixture must be writable");
    writer
        .write_stored(
            "_rels/.rels",
            &package_relationships_for(BINARY_MEMBER, false),
        )
        .expect("package relationships fixture must be writable");
    writer
        .write_stored(SCRATCH_MEMBER, b"opaque physical bytes\0\xff")
        .expect("opaque fixture must be writable");
    if deflated {
        writer
            .write_deflated_sized(BINARY_MEMBER, source)
            .expect("binary Deflate fixture must be writable");
    } else {
        writer
            .write_stored(BINARY_MEMBER, source)
            .expect("binary Store fixture must be writable");
    }
    writer
        .finish_to_bytes()
        .expect("binary fixture archive must finish")
}

fn open(source: &[u8], deflated: bool, signed: bool) -> (SourceBackedPackage, Vec<u8>) {
    let archive = archive_bytes(source, deflated, signed);
    let package =
        SourceBackedPackage::from_vec(archive.clone()).expect("fixture package must open");
    (package, archive)
}

fn proof(
    package: &SourceBackedPackage,
    source: &[u8],
    offset: usize,
    fragment: &[u8],
) -> SourcePartSpliceProof {
    proof_for(package, source, offset, fragment)
}

fn proof_for(
    package: &SourceBackedPackage,
    source: &[u8],
    offset: usize,
    fragment: &[u8],
) -> SourcePartSpliceProof {
    let candidate = candidate_bytes(source, offset, fragment);
    SourcePartSpliceProof {
        source_version: package
            .source_version()
            .expect("source version must be available"),
        source_len: source.len() as u64,
        source_sha256: digest(source),
        insertion_offset: offset as u64,
        fragment_len: fragment.len() as u64,
        fragment_sha256: digest(fragment),
        candidate_len: candidate.len() as u64,
        candidate_sha256: digest(&candidate),
    }
}

fn prepare<'a>(
    package: &'a SourceBackedPackage,
    source: &[u8],
    offset: usize,
    fragment: &[u8],
    limits: SourcePartSpliceLimits,
) -> litchi_opc::SourcePartSplicePlan<'a> {
    package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(package, source, offset, fragment),
            Arc::new(fragment.to_vec()),
            limits,
        )
        .expect("source splice proof must prepare")
}

fn target_bytes(archive: &[u8]) -> Vec<u8> {
    target_bytes_for(archive, DOCUMENT_URI)
}

fn target_bytes_for(archive: &[u8], uri: &str) -> Vec<u8> {
    let package = SourceBackedPackage::from_vec(archive.to_vec()).expect("published archive opens");
    package
        .part(&pack(uri))
        .expect("published target exists")
        .data()
        .expect("published target reads")
        .as_bytes()
        .to_vec()
}

fn u16_at(bytes: &[u8], offset: usize) -> usize {
    usize::from(u16::from_le_bytes(
        bytes[offset..offset + 2]
            .try_into()
            .expect("ZIP16 field must fit"),
    ))
}

fn u32_at(bytes: &[u8], offset: usize) -> usize {
    usize::try_from(u32::from_le_bytes(
        bytes[offset..offset + 4]
            .try_into()
            .expect("ZIP32 field must fit"),
    ))
    .expect("fixture offset must fit usize")
}

/// Return a complete local record, including compressed payload and descriptor.
/// This independent parser is restricted to the generated non-ZIP64, comment-free
/// fixtures. Sizes come from the central directory: descriptor-mode local sizes
/// can be zero and byte signatures can occur inside compressed payloads.
fn local_record(bytes: &[u8], wanted: &str) -> Vec<u8> {
    let eocd = bytes.len().checked_sub(22).expect("fixture EOCD must fit");
    assert_eq!(&bytes[eocd..eocd + 4], b"PK\x05\x06");
    assert_eq!(u16_at(bytes, eocd + 20), 0);
    let mut central = u32_at(bytes, eocd + 16);
    for _ in 0..u16_at(bytes, eocd + 10) {
        assert_eq!(&bytes[central..central + 4], b"PK\x01\x02");
        let name_len = u16_at(bytes, central + 28);
        let extra_len = u16_at(bytes, central + 30);
        let comment_len = u16_at(bytes, central + 32);
        let name = &bytes[central + 46..central + 46 + name_len];
        if name == wanted.as_bytes() {
            let start = u32_at(bytes, central + 42);
            let compressed = u32_at(bytes, central + 20);
            assert_ne!(compressed, u32::MAX as usize, "fixture must not use ZIP64");
            assert_eq!(&bytes[start..start + 4], b"PK\x03\x04");
            let payload = start + 30 + u16_at(bytes, start + 26) + u16_at(bytes, start + 28);
            let mut end = payload + compressed;
            if u16_at(bytes, central + 8) & 8 != 0 {
                end += if &bytes[end..end + 4] == b"PK\x07\x08" {
                    16
                } else {
                    12
                };
            }
            return bytes[start..end].to_vec();
        }
        central += 46 + name_len + extra_len + comment_len;
    }
    panic!("fixture member {wanted} is missing")
}

/// Preserve every central-directory byte except the relocated local offset.
/// This is independent of production ZIP metadata parsing and covers names,
/// extra fields, comments, flags, versions, timestamps, CRC, sizes and attrs.
fn central_record(bytes: &[u8], wanted: &str) -> Vec<u8> {
    let eocd = bytes.len().checked_sub(22).expect("fixture EOCD must fit");
    assert_eq!(&bytes[eocd..eocd + 4], b"PK\x05\x06");
    assert_eq!(u16_at(bytes, eocd + 20), 0);
    let mut central = u32_at(bytes, eocd + 16);
    for _ in 0..u16_at(bytes, eocd + 10) {
        assert_eq!(&bytes[central..central + 4], b"PK\x01\x02");
        let name_len = u16_at(bytes, central + 28);
        let end =
            central + 46 + name_len + u16_at(bytes, central + 30) + u16_at(bytes, central + 32);
        if &bytes[central + 46..central + 46 + name_len] == wanted.as_bytes() {
            assert_ne!(u32_at(bytes, central + 42), u32::MAX as usize);
            let mut record = bytes[central..end].to_vec();
            record[42..46].fill(0);
            return record;
        }
        central = end;
    }
    panic!("fixture member {wanted} is missing")
}

#[test]
fn central_record_oracle_detects_metadata_changes_outside_local_record() {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_deflated(DOCUMENT_MEMBER, VALID_SOURCE)
        .unwrap();
    let archive = writer.finish_to_bytes().unwrap();
    let mut changed = archive.clone();
    let central = u32_at(&archive, archive.len() - 22 + 16);
    changed[central + 38] ^= 1; // External attributes occur only in the directory.
    assert_eq!(
        local_record(&archive, DOCUMENT_MEMBER),
        local_record(&changed, DOCUMENT_MEMBER)
    );
    assert_ne!(
        central_record(&archive, DOCUMENT_MEMBER),
        central_record(&changed, DOCUMENT_MEMBER)
    );
    changed[central + 38] ^= 1;
    changed[central + 42] ^= 1; // Relocation alone is excluded from this comparison.
    assert_eq!(
        central_record(&archive, DOCUMENT_MEMBER),
        central_record(&changed, DOCUMENT_MEMBER)
    );
}

#[test]
fn raw_record_oracle_includes_descriptor_mode_compressed_payload() {
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_deflated(DOCUMENT_MEMBER, VALID_SOURCE)
        .expect("descriptor-mode fixture must write");
    let archive = writer.finish_to_bytes().expect("fixture must finish");
    let record = local_record(&archive, DOCUMENT_MEMBER);
    let payload = 30 + u16_at(&record, 26) + u16_at(&record, 28);
    assert_eq!(u32_at(&record, 18), 0, "fixture uses deferred local sizes");
    assert_ne!(u16_at(&record, 6) & 8, 0);
    assert_eq!(&record[record.len() - 16..record.len() - 12], b"PK\x07\x08");
    assert!(
        record.len() > payload + 16,
        "oracle must include compressed bytes"
    );
}

fn managed_context_with_limits(
    memory: u64,
    work: u64,
    output: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let (budget, context) = managed_context_for_token(memory, work, output, cancellation);
    (budget, cancellation_source, context)
}

fn managed_context_for_token(
    memory: u64,
    work: u64,
    output: u64,
    cancellation: litchi_core::CancellationToken,
) -> (Budget, ExecutionContext) {
    let budget = Budget::root(
        "source-part-splice-test",
        Limits::new(memory, u64::MAX, output, u64::MAX, u64::MAX, work),
    );
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker limit is nonzero"),
        NonZeroUsize::new(1).expect("operation limit is nonzero"),
        NonZeroU64::new(memory).expect("memory limit is nonzero"),
        0,
    )
    .expect("execution limits must be valid");
    (
        budget.clone(),
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn managed_context_with_budget() -> (Budget, CancellationSource, ExecutionContext) {
    managed_context_with_limits(64 * 1024 * 1024, u64::MAX, u64::MAX)
}

fn managed_context() -> (CancellationSource, ExecutionContext) {
    let (_, cancellation_source, context) = managed_context_with_budget();
    (cancellation_source, context)
}

fn xml_workspace_requirement(limits: SourcePartSpliceLimits) -> u64 {
    u64::try_from(
        limits
            .xml_audit_limits
            .streaming_memory_upper_bound()
            .expect("the default XML audit profile has a finite workspace bound"),
    )
    .expect("XML workspace must fit u64")
    .checked_add(64 * 1024)
    .expect("XML workspace fixture bound must fit u64")
}

#[derive(Debug)]
struct PrefixFailSink {
    bytes: Vec<u8>,
    remaining: usize,
    fail_on_flush: bool,
}

impl PrefixFailSink {
    fn after(bytes: usize) -> Self {
        Self {
            bytes: Vec::new(),
            remaining: bytes,
            fail_on_flush: false,
        }
    }

    fn flush_fails() -> Self {
        Self {
            bytes: Vec::new(),
            remaining: usize::MAX,
            fail_on_flush: true,
        }
    }
}

impl Write for PrefixFailSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "test sink stopped",
            ));
        }
        let accepted = bytes.len().min(self.remaining);
        self.bytes.extend_from_slice(&bytes[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.fail_on_flush {
            Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "test flush stopped",
            ))
        } else {
            Ok(())
        }
    }
}

struct MutatingSink {
    bytes: Vec<u8>,
    source: Arc<MutableSource>,
    replacement: Vec<u8>,
    mutated: bool,
}

impl MutatingSink {
    fn new(source: Arc<MutableSource>, replacement: Vec<u8>) -> Self {
        Self {
            bytes: Vec::new(),
            source,
            replacement,
            mutated: false,
        }
    }
}

impl Write for MutatingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.mutated {
            self.source.replace_bytes(self.replacement.clone());
            self.mutated = true;
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct MutatingFlushSink {
    bytes: Vec<u8>,
    source: Arc<MutableSource>,
    replacement: Vec<u8>,
    mutated: bool,
}

impl MutatingFlushSink {
    fn new(source: Arc<MutableSource>, replacement: Vec<u8>) -> Self {
        Self {
            bytes: Vec::new(),
            source,
            replacement,
            mutated: false,
        }
    }
}

impl Write for MutatingFlushSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        if !self.mutated {
            self.source.replace_bytes(self.replacement.clone());
            self.mutated = true;
        }
        Ok(())
    }
}

struct CancellingSink {
    bytes: Vec<u8>,
    cancellation: CancellationSource,
    on_flush: bool,
    cancelled: bool,
}

impl CancellingSink {
    fn on_write(cancellation: CancellationSource) -> Self {
        Self {
            bytes: Vec::new(),
            cancellation,
            on_flush: false,
            cancelled: false,
        }
    }

    fn on_flush(cancellation: CancellationSource) -> Self {
        Self {
            bytes: Vec::new(),
            cancellation,
            on_flush: true,
            cancelled: false,
        }
    }
}

impl Write for CancellingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        if !self.on_flush && !self.cancelled {
            self.cancellation.cancel();
            self.cancelled = true;
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.on_flush && !self.cancelled {
            self.cancellation.cancel();
            self.cancelled = true;
        }
        Ok(())
    }
}

fn contains_source_change(error: &OpcError) -> bool {
    match error {
        OpcError::SourceChanged { .. } => true,
        OpcError::IncompleteOutput { source, .. } => contains_source_change(source),
        _ => false,
    }
}

fn contains_cancellation(error: &OpcError) -> bool {
    match error {
        OpcError::Cancelled => true,
        OpcError::IncompleteOutput { source, .. } => contains_cancellation(source),
        _ => false,
    }
}

fn assert_sink_failure_count(error: &OpcError, accepted: usize) {
    match error {
        OpcError::IncompleteOutput { written, .. } => assert_eq!(
            usize::try_from(*written).expect("accepted output count must fit usize"),
            accepted
        ),
        OpcError::IoError(_) => {},
        other => panic!("unexpected sink failure: {other:?}"),
    }
}

#[test]
fn store_and_deflate_splices_preserve_opaque_members_without_caching_target() {
    let expected = candidate_bytes(VALID_SOURCE, INSERTION_OFFSET, FRAGMENT);
    for deflated in [false, true] {
        let (package, source_archive) = open(VALID_SOURCE, deflated, false);
        assert!(
            package
                .non_part_members()
                .iter()
                .any(|member| member.name() == SCRATCH_MEMBER)
        );
        let before = package.cache_diagnostics();
        assert_eq!(before.cold_loads, 0, "fixture setup must not load target");

        let plan = prepare(
            &package,
            VALID_SOURCE,
            INSERTION_OFFSET,
            FRAGMENT,
            SourcePartSpliceLimits::default(),
        );
        let mut output = Vec::new();
        let publication = plan
            .write_to_stream(&mut output)
            .expect("bounded decoded splice must publish");

        assert_eq!(target_bytes(&output), expected);
        assert!(!publication.is_noop());
        assert_eq!(publication.candidate_artifact_len(), output.len() as u64);
        assert_eq!(package.cache_diagnostics().cold_loads, 0);
        assert_eq!(
            local_record(&source_archive, SCRATCH_MEMBER),
            local_record(&output, SCRATCH_MEMBER),
            "untyped physical member must remain byte exact"
        );
        assert_eq!(
            central_record(&source_archive, SCRATCH_MEMBER),
            central_record(&output, SCRATCH_MEMBER),
            "untyped physical member directory metadata must survive relocation"
        );
    }
}

#[test]
fn accounting_reports_two_fresh_decoded_reader_passes() {
    for deflated in [false, true] {
        let (package, _) = open(VALID_SOURCE, deflated, false);
        let plan = prepare(
            &package,
            VALID_SOURCE,
            INSERTION_OFFSET,
            FRAGMENT,
            SourcePartSpliceLimits::default(),
        );
        let mut output = Vec::new();
        let mut accounting = OpcOperationAccounting::default();
        plan.write_to_stream_with_accounting(&mut output, &mut accounting)
            .expect("accounted decoded splice must publish");

        let expected_decoded_reads = (VALID_SOURCE.len() as u64) * 2;
        if deflated {
            assert_eq!(
                accounting.deflate_bytes_produced(),
                expected_decoded_reads,
                "both replay callback passes must decode the complete Deflate member"
            );
        } else {
            assert_eq!(
                accounting.stored_payload_bytes_read(),
                expected_decoded_reads,
                "both replay callback passes must read the complete Store member"
            );
        }
    }
}

#[test]
fn binary_part_splice_skips_xml_audit_but_keeps_decoded_proofs() {
    const BINARY_SOURCE: &[u8] = b"opaque\0source\xff";
    const BINARY_FRAGMENT: &[u8] = b"\x00fragment\xfe";
    let offset = 6usize;
    for deflated in [false, true] {
        let archive = binary_archive_bytes(BINARY_SOURCE, deflated);
        let package = SourceBackedPackage::from_vec(archive).expect("binary fixture must open");
        let candidate = candidate_bytes(BINARY_SOURCE, offset, BINARY_FRAGMENT);
        let proof = SourcePartSpliceProof {
            source_version: package.source_version().expect("source version must exist"),
            source_len: BINARY_SOURCE.len() as u64,
            source_sha256: digest(BINARY_SOURCE),
            insertion_offset: offset as u64,
            fragment_len: BINARY_FRAGMENT.len() as u64,
            fragment_sha256: digest(BINARY_FRAGMENT),
            candidate_len: candidate.len() as u64,
            candidate_sha256: digest(&candidate),
        };
        let plan = package
            .prepare_source_part_splice(
                &pack(BINARY_URI),
                proof,
                Arc::new(BINARY_FRAGMENT.to_vec()),
                SourcePartSpliceLimits::default(),
            )
            .expect("binary Part must not be forced through XML audit");
        let mut output = Vec::new();
        plan.write_to_stream(&mut output)
            .expect("binary decoded splice must publish");
        assert_eq!(target_bytes_for(&output, BINARY_URI), candidate);
    }
}

#[test]
fn exact_noop_copies_malformed_and_signed_source_byte_for_byte() {
    let malformed = b"<not-xml";
    for signed in [false, true] {
        let (package, source_archive) = open(malformed, false, signed);
        let proof = proof(&package, malformed, 0, &[]);
        let plan = package
            .prepare_source_part_splice(
                &pack(DOCUMENT_URI),
                proof,
                Arc::new(Vec::new()),
                SourcePartSpliceLimits::default(),
            )
            .expect("exact no-op must retain source authority");
        assert!(plan.is_noop());
        let mut output = Vec::new();
        let publication = plan
            .write_to_stream(&mut output)
            .expect("exact no-op must bypass XML and signature policy");
        assert!(publication.is_noop());
        assert_eq!(output, source_archive);
        assert_eq!(publication.candidate_artifact_len(), output.len() as u64);
    }
}

#[test]
fn exact_noop_authenticates_source_and_candidate_hashes_and_archive_limit() {
    let malformed = b"<not-xml";
    let (package, source_archive) = open(malformed, false, false);
    let baseline = proof(&package, malformed, 0, &[]);

    let mut wrong_source_hash = baseline;
    wrong_source_hash.source_sha256[0] ^= 1;
    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            wrong_source_hash,
            Arc::new(Vec::new()),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("a no-op must authenticate its source hash");
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));

    let mut wrong_candidate_hash = baseline;
    wrong_candidate_hash.candidate_sha256[0] ^= 1;
    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            wrong_candidate_hash,
            Arc::new(Vec::new()),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("a no-op must authenticate its candidate hash");
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));

    let limits = SourcePartSpliceLimits::new(
        malformed.len() as u64,
        1,
        malformed.len() as u64,
        source_archive.len() as u64,
        source_archive.len() as u64 - 1,
        source_archive.len() as u64 + 1,
    )
    .expect("finite no-op limits must be valid");
    let plan = package
        .prepare_source_part_splice(&pack(DOCUMENT_URI), baseline, Arc::new(Vec::new()), limits)
        .expect("valid no-op proof must prepare");
    let mut output = Vec::new();
    let error = plan
        .write_to_stream(&mut output)
        .expect_err("physical archive limit must apply to an exact no-op");
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: litchi_opc::ReadResource::ArchiveTotalBytes,
            ..
        }
    ));
    assert!(output.is_empty());
}

#[test]
fn signed_non_noop_refuses_before_output() {
    let (package, _) = open(VALID_SOURCE, false, true);
    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("signed source requires an explicit edit policy");
    assert!(matches!(
        error,
        OpcError::SignedSourceRequiresExplicitPolicy
    ));
}

#[test]
fn malformed_source_or_candidate_is_rejected_before_output() {
    let malformed_source = b"<root>";
    let (package, _) = open(malformed_source, false, false);
    let fragment = b"<inserted/>";
    let source_proof = proof(&package, malformed_source, malformed_source.len(), fragment);
    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            source_proof,
            Arc::new(fragment.to_vec()),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("malformed source must fail the independent XML audit");
    assert!(matches!(error, OpcError::XmlPublication { .. }));

    let (package, _) = open(VALID_SOURCE, false, false);
    let malformed_fragment = b"<inserted";
    let candidate_proof = proof(&package, VALID_SOURCE, INSERTION_OFFSET, malformed_fragment);
    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            candidate_proof,
            Arc::new(malformed_fragment.to_vec()),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("malformed candidate must fail before publication");
    assert!(matches!(error, OpcError::XmlPublication { .. }));
}

#[test]
fn mismatched_proof_fields_are_refused_before_publication() {
    let (package, _) = open(VALID_SOURCE, false, false);
    let fragment = FRAGMENT;
    let baseline = proof(&package, VALID_SOURCE, INSERTION_OFFSET, fragment);
    let mut cases = Vec::new();

    let mut source_length = baseline;
    source_length.source_len += 1;
    cases.push(source_length);

    let mut source_hash = baseline;
    source_hash.source_sha256[0] ^= 1;
    cases.push(source_hash);

    let mut fragment_length = baseline;
    fragment_length.fragment_len += 1;
    cases.push(fragment_length);

    let mut candidate_length = baseline;
    candidate_length.candidate_len += 1;
    cases.push(candidate_length);

    let mut offset = baseline;
    offset.insertion_offset = offset.source_len + 1;
    cases.push(offset);

    let mut fragment_hash = baseline;
    fragment_hash.fragment_sha256[0] ^= 1;
    cases.push(fragment_hash);

    let mut candidate_hash = baseline;
    candidate_hash.candidate_sha256[0] ^= 1;
    cases.push(candidate_hash);

    for invalid in cases {
        let result = package.prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            invalid,
            Arc::new(fragment.to_vec()),
            SourcePartSpliceLimits::default(),
        );
        assert!(result.is_err(), "invalid proof was accepted: {invalid:?}");
    }
}

#[test]
fn source_version_mutation_is_refused_before_output() {
    let source_archive = archive_bytes(VALID_SOURCE, false, false);
    let source = Arc::new(MutableSource::new(source_archive));
    let package = SourceBackedPackage::from_read_at(source.clone())
        .expect("mutable fixture package must open");
    let expected_version = package.source_version().expect("source version must exist");
    source.replace_bytes(archive_bytes(VALID_SOURCE, false, false));
    assert_ne!(source.version().unwrap(), expected_version);

    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            SourcePartSpliceProof {
                source_version: expected_version,
                source_len: VALID_SOURCE.len() as u64,
                source_sha256: digest(VALID_SOURCE),
                insertion_offset: INSERTION_OFFSET as u64,
                fragment_len: FRAGMENT.len() as u64,
                fragment_sha256: digest(FRAGMENT),
                candidate_len: (VALID_SOURCE.len() + FRAGMENT.len()) as u64,
                candidate_sha256: digest(&candidate_bytes(
                    VALID_SOURCE,
                    INSERTION_OFFSET,
                    FRAGMENT,
                )),
            },
            Arc::new(FRAGMENT.to_vec()),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("source mutation must invalidate the captured version");
    assert!(matches!(error, OpcError::SourceChanged { .. }));
}

#[test]
fn source_mutation_after_the_first_sink_write_wins_over_publication() {
    let source_archive = archive_bytes(VALID_SOURCE, false, false);
    let source = Arc::new(MutableSource::new(source_archive.clone()));
    let package = SourceBackedPackage::from_read_at(source.clone())
        .expect("mutable fixture package must open");
    let plan = prepare(
        &package,
        VALID_SOURCE,
        INSERTION_OFFSET,
        FRAGMENT,
        SourcePartSpliceLimits::default(),
    );
    let mut sink = MutatingSink::new(Arc::clone(&source), source_archive);
    let error = plan
        .write_to_stream(&mut sink)
        .expect_err("source mutation during output must invalidate publication");
    assert!(
        contains_source_change(&error),
        "unexpected error: {error:?}"
    );
    assert!(
        !sink.bytes.is_empty(),
        "the sink should report accepted progress"
    );
}

#[test]
fn large_binary_fragment_refuses_scoped_work_before_preparation_succeeds() {
    const FRAGMENT_BYTES: usize = 2 * 1024 * 1024;
    const SOURCE_WORK_ALLOWANCE: u64 = 64 * 1024;
    const WORK_LIMIT: u64 = 1024 * 1024 * 1024;
    const BINARY_SOURCE: &[u8] = b"binary source";

    let fragment = vec![0xa5; FRAGMENT_BYTES];
    let archive = binary_archive_bytes(BINARY_SOURCE, false);
    let (budget, _cancellation, context) =
        managed_context_with_limits(64 * 1024 * 1024, WORK_LIMIT, u64::MAX);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed binary fixture must open");
    let baseline = budget.used(Resource::Work);
    let available = WORK_LIMIT
        .checked_sub(baseline)
        .expect("fixture setup must stay within the work budget");
    assert!(available > SOURCE_WORK_ALLOWANCE);
    budget
        .consume(Resource::Work, available - SOURCE_WORK_ALLOWANCE)
        .expect("fixture should leave the source allowance available");

    let error = package
        .prepare_source_part_splice(
            &pack(BINARY_URI),
            proof_for(&package, BINARY_SOURCE, BINARY_SOURCE.len(), &fragment),
            Arc::new(fragment),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("large fragment hashing must refuse the exhausted Work budget");
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Work
    ));
    assert!(budget.used(Resource::Work) <= WORK_LIMIT);
}

#[test]
fn large_binary_fragment_observes_managed_cancellation_between_work_chunks() {
    const FRAGMENT_BYTES: usize = 8 * 1024 * 1024;
    const BINARY_SOURCE: &[u8] = b"binary source";

    let fragment = vec![0x3c; FRAGMENT_BYTES];
    let archive = binary_archive_bytes(BINARY_SOURCE, false);
    let (cancellation, cancellation_token) = CancellationSource::pair();
    let (budget, context) =
        managed_context_for_token(64 * 1024 * 1024, u64::MAX, u64::MAX, cancellation_token);
    let source = Arc::new(WorkCancellingSource::new(
        archive,
        budget.clone(),
        cancellation,
    ));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source.clone(),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed binary fixture must open");
    let proof = proof_for(&package, BINARY_SOURCE, BINARY_SOURCE.len(), &fragment);
    let baseline = budget.used(Resource::Work);
    source.arm(baseline);
    let result = package.prepare_source_part_splice(
        &pack(BINARY_URI),
        proof,
        Arc::new(fragment),
        SourcePartSpliceLimits::default(),
    );
    assert!(matches!(result, Err(OpcError::Cancelled)));
    let work_delta = budget.used(Resource::Work).saturating_sub(baseline);
    assert!(work_delta > 0);
    assert!(work_delta < FRAGMENT_BYTES as u64);
}

#[test]
fn candidate_limit_is_refused_without_output() {
    let (package, _) = open(VALID_SOURCE, false, false);
    let candidate_len = (VALID_SOURCE.len() + FRAGMENT.len()) as u64;
    let limits = SourcePartSpliceLimits::new(
        VALID_SOURCE.len() as u64,
        FRAGMENT.len() as u64,
        candidate_len - 1,
        1024 * 1024,
        1024 * 1024,
        1024 * 1024,
    )
    .expect("finite test limits must be valid");
    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            limits,
        )
        .expect_err("candidate limit must reject the plan");
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            actual,
            maximum,
            ..
        } if actual == candidate_len && maximum == candidate_len - 1
    ));
}

#[test]
fn cancellation_is_observed_before_plan_and_write() {
    let archive = archive_bytes(VALID_SOURCE, false, false);
    let (cancellation, context) = managed_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed source fixture must open");
    cancellation.cancel();
    let result = package.prepare_source_part_splice(
        &pack(DOCUMENT_URI),
        proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
        Arc::new(FRAGMENT.to_vec()),
        SourcePartSpliceLimits::default(),
    );
    assert!(matches!(result, Err(OpcError::Cancelled)));

    let archive = archive_bytes(VALID_SOURCE, false, false);
    let (cancellation, context) = managed_context();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed source fixture must open");
    let plan = prepare(
        &package,
        VALID_SOURCE,
        INSERTION_OFFSET,
        FRAGMENT,
        SourcePartSpliceLimits::default(),
    );
    cancellation.cancel();
    let mut output = Vec::new();
    let error = plan
        .write_to_stream(&mut output)
        .expect_err("cancellation must stop publication");
    assert!(matches!(error, OpcError::Cancelled));
    assert!(output.is_empty());
}

#[test]
fn partial_and_flush_sink_failures_preserve_typed_progress() {
    let (package, _) = open(VALID_SOURCE, false, false);
    let plan = prepare(
        &package,
        VALID_SOURCE,
        INSERTION_OFFSET,
        FRAGMENT,
        SourcePartSpliceLimits::default(),
    );
    let mut partial = PrefixFailSink::after(17);
    let error = plan
        .write_to_stream(&mut partial)
        .expect_err("partial sink must fail publication");
    assert_eq!(partial.bytes.len(), 17);
    assert_sink_failure_count(&error, 17);

    let (package, _) = open(VALID_SOURCE, false, false);
    let plan = prepare(
        &package,
        VALID_SOURCE,
        INSERTION_OFFSET,
        FRAGMENT,
        SourcePartSpliceLimits::default(),
    );
    let mut expected_output = Vec::new();
    prepare(
        &package,
        VALID_SOURCE,
        INSERTION_OFFSET,
        FRAGMENT,
        SourcePartSpliceLimits::default(),
    )
    .write_to_stream(&mut expected_output)
    .expect("reference publication must succeed");
    let mut flush_failure = PrefixFailSink::flush_fails();
    let error = plan
        .write_to_stream(&mut flush_failure)
        .expect_err("flush failure must fail publication");
    assert_eq!(flush_failure.bytes.len(), expected_output.len());
    assert_sink_failure_count(&error, expected_output.len());
}

#[test]
fn publication_inverse_requires_the_authenticated_candidate_artifact() {
    let (package, source_archive) = open(VALID_SOURCE, true, false);
    let plan = prepare(
        &package,
        VALID_SOURCE,
        INSERTION_OFFSET,
        FRAGMENT,
        SourcePartSpliceLimits::default(),
    );
    let mut candidate_archive = Vec::new();
    let publication = plan
        .write_to_stream(&mut candidate_archive)
        .expect("candidate publication must succeed");
    let candidate = SourceBackedPackage::from_vec(candidate_archive.clone())
        .expect("candidate package must reopen");
    let mut inverse = Vec::new();
    publication
        .write_inverse_to_stream(&candidate, &mut inverse)
        .expect("authenticated candidate must authorize exact inverse");
    assert_eq!(inverse, source_archive);

    let (other_package, _) = open(VALID_SOURCE, true, false);
    let other_plan = prepare(
        &other_package,
        VALID_SOURCE,
        INSERTION_OFFSET,
        b"<different/>",
        SourcePartSpliceLimits::default(),
    );
    let mut other_archive = Vec::new();
    other_plan
        .write_to_stream(&mut other_archive)
        .expect("second candidate publication must succeed");
    let other_candidate =
        SourceBackedPackage::from_vec(other_archive).expect("second candidate package must reopen");
    let mut rejected = Vec::new();
    let error = publication
        .write_inverse_to_stream(&other_candidate, &mut rejected)
        .expect_err("a different candidate must not authorize inverse");
    assert!(matches!(
        error,
        OpcError::SourceArtifactMismatch {
            artifact: "current",
            field: "length",
        }
    ));
    assert!(rejected.is_empty());
}

fn published_candidate() -> (litchi_opc::SourcePartSplicePublication, Vec<u8>, Vec<u8>) {
    let (package, source_archive) = open(VALID_SOURCE, true, false);
    let plan = prepare(
        &package,
        VALID_SOURCE,
        INSERTION_OFFSET,
        FRAGMENT,
        SourcePartSpliceLimits::default(),
    );
    let mut candidate_archive = Vec::new();
    let publication = plan
        .write_to_stream(&mut candidate_archive)
        .expect("candidate publication must succeed");
    (publication, candidate_archive, source_archive)
}

#[test]
fn publication_inverse_rechecks_candidate_freshness_during_write_and_flush() {
    let (publication, candidate_archive, _) = published_candidate();
    let candidate_source = Arc::new(MutableSource::new(candidate_archive.clone()));
    let candidate = SourceBackedPackage::from_read_at(candidate_source.clone())
        .expect("mutable candidate package must open");
    let mut sink = MutatingSink::new(Arc::clone(&candidate_source), candidate_archive.clone());
    let error = publication
        .write_inverse_to_stream(&candidate, &mut sink)
        .expect_err("candidate mutation during inverse writes must invalidate publication");
    assert!(
        contains_source_change(&error),
        "unexpected error: {error:?}"
    );
    assert!(
        !sink.bytes.is_empty(),
        "inverse sink must report accepted bytes"
    );

    let (publication, candidate_archive, _) = published_candidate();
    let candidate_source = Arc::new(MutableSource::new(candidate_archive.clone()));
    let candidate = SourceBackedPackage::from_read_at(candidate_source.clone())
        .expect("mutable candidate package must open");
    let mut sink = MutatingFlushSink::new(Arc::clone(&candidate_source), candidate_archive);
    let error = publication
        .write_inverse_to_stream(&candidate, &mut sink)
        .expect_err("candidate mutation during inverse flush must invalidate publication");
    assert!(
        contains_source_change(&error),
        "unexpected error: {error:?}"
    );
    assert!(
        !sink.bytes.is_empty(),
        "inverse flush must report accepted bytes"
    );
}

#[test]
fn publication_inverse_observes_candidate_cancellation_during_write_and_flush() {
    let (publication, candidate_archive, _) = published_candidate();
    let (cancellation, context) = managed_context();
    let candidate = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(candidate_archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed candidate package must open");
    let mut sink = CancellingSink::on_write(cancellation);
    let error = publication
        .write_inverse_to_stream(&candidate, &mut sink)
        .expect_err("candidate cancellation during inverse writes must stop publication");
    assert!(contains_cancellation(&error), "unexpected error: {error:?}");
    assert!(
        !sink.bytes.is_empty(),
        "inverse sink must report accepted bytes"
    );

    let (publication, candidate_archive, _) = published_candidate();
    let (cancellation, context) = managed_context();
    let candidate = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(candidate_archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed candidate package must open");
    let mut sink = CancellingSink::on_flush(cancellation);
    let error = publication
        .write_inverse_to_stream(&candidate, &mut sink)
        .expect_err("candidate cancellation during inverse flush must stop publication");
    assert!(contains_cancellation(&error), "unexpected error: {error:?}");
    assert!(
        !sink.bytes.is_empty(),
        "inverse flush must report accepted bytes"
    );
}

#[test]
fn publication_inverse_uses_current_candidate_output_budget() {
    let (publication, candidate_archive, _) = published_candidate();
    let (budget, _cancellation, context) =
        managed_context_with_limits(64 * 1024 * 1024, u64::MAX, 0);
    let candidate = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(candidate_archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed candidate package must open");
    let mut output = Vec::new();
    let error = publication
        .write_inverse_to_stream(&candidate, &mut output)
        .expect_err("inverse output must use the current candidate budget");
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::OutputBytes
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::OutputBytes), 0);
}

#[test]
fn publication_inverse_shared_context_charges_retained_archive_once() {
    let source_archive = archive_bytes(VALID_SOURCE, true, false);
    let (budget, _cancellation, context) = managed_context_with_budget();
    let original = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(source_archive.clone())),
        litchi_opc::ReadLimits::default(),
        context.clone(),
    )
    .expect("managed original package must open");
    let plan = original
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&original, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            SourcePartSpliceLimits::default(),
        )
        .expect("managed source splice must prepare");
    let mut candidate_archive = Vec::new();
    let publication = plan
        .write_to_stream(&mut candidate_archive)
        .expect("managed candidate publication must succeed");
    let before_inverse = budget.used(Resource::OutputBytes);

    let candidate = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(candidate_archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed candidate package must open");
    let mut inverse = Vec::new();
    publication
        .write_inverse_to_stream(&candidate, &mut inverse)
        .expect("shared-context inverse must publish");

    assert_eq!(inverse, source_archive);
    assert_eq!(
        budget.used(Resource::OutputBytes) - before_inverse,
        source_archive.len() as u64,
        "inverse must charge the retained original archive exactly once"
    );
}

#[test]
fn splice_memory_quotas_are_typed_and_release_after_success_or_failure() {
    let archive = archive_bytes(VALID_SOURCE, false, false);
    let (budget, _cancellation, context) = managed_context_with_budget();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive)),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed source fixture must open");
    let baseline = budget.used(Resource::Memory);
    let defaults = SourcePartSpliceLimits::default();
    let workspace = xml_workspace_requirement(defaults);

    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            defaults.with_max_xml_workspace_bytes(workspace - 1),
        )
        .expect_err("one byte below XML workspace must be refused");
    assert!(matches!(
        error,
        OpcError::SourcePartSpliceLimit {
            resource: SpliceResource::XmlWorkspaceBytes,
            actual,
            maximum,
        } if actual == workspace && maximum == workspace - 1
    ));
    assert_eq!(budget.used(Resource::Memory), baseline);

    let plan = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            defaults.with_max_xml_workspace_bytes(workspace),
        )
        .expect("the exact XML workspace boundary must prepare");
    assert!(budget.used(Resource::Memory) > baseline);
    drop(plan);
    assert_eq!(budget.used(Resource::Memory), baseline);

    let error = SourcePartSpliceLimits::default().with_max_xml_workspace_bytes(0);
    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            error,
        )
        .expect_err("zero XML workspace must be rejected as an invalid quota");
    assert!(matches!(
        error,
        OpcError::InvalidSourcePartSpliceLimit {
            resource: SpliceResource::XmlWorkspaceBytes,
            value: 0,
        }
    ));
    assert_eq!(budget.used(Resource::Memory), baseline);

    let plan = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            defaults.with_max_preservation_memory_bytes(1),
        )
        .expect("preservation quota is checked during writing");
    let mut output = Vec::new();
    let error = plan
        .write_to_stream(&mut output)
        .expect_err("preservation quota must be typed");
    assert!(matches!(
        error,
        OpcError::SourcePartSpliceLimit {
            resource: SpliceResource::PreservationMemoryBytes,
            ..
        }
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::Memory), baseline);

    let plan = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            defaults.with_max_replay_memory_bytes(1),
        )
        .expect("replay quota is checked during writing");
    let mut output = Vec::new();
    let error = plan
        .write_to_stream(&mut output)
        .expect_err("replay quota must be typed");
    assert!(matches!(
        error,
        OpcError::SourcePartSpliceLimit {
            resource: SpliceResource::ReplayMemoryBytes,
            ..
        }
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::Memory), baseline);
}

#[test]
fn authored_fragment_transfers_one_reservation_and_publishes_reversibly() {
    for deflated in [false, true] {
        let source_archive = archive_bytes(VALID_SOURCE, deflated, false);
        let (budget, _cancellation, context) = managed_context_with_budget();
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            Arc::new(OwnedSource::new(source_archive.clone())),
            litchi_opc::ReadLimits::default(),
            context,
        )
        .expect("managed source must open");
        let baseline = budget.used(Resource::Memory);
        let length = FRAGMENT.len() as u64;
        let mut fragment = package
            .allocate_source_part_splice_fragment(length, length)
            .expect("exact fragment limit must allocate");
        assert_eq!(budget.used(Resource::Memory), baseline + length);
        assert_eq!(fragment.as_slice(), vec![0; FRAGMENT.len()]);
        fragment.as_mut_slice().copy_from_slice(FRAGMENT);
        let limits = SourcePartSpliceLimits::default();
        let plan = package
            .prepare_source_part_splice_with_fragment(
                &pack(DOCUMENT_URI),
                proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
                fragment,
                limits,
            )
            .expect("authored fragment must transfer into the plan");
        assert_eq!(
            budget.used(Resource::Memory),
            baseline + length + xml_workspace_requirement(limits),
            "preparation must retain exactly one fragment charge"
        );
        let mut output = Vec::new();
        let publication = plan
            .write_to_stream(&mut output)
            .expect("splice must publish");
        assert_eq!(budget.used(Resource::Memory), baseline);
        let candidate = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(output)))
            .expect("candidate must reopen");
        let mut inverse = Vec::new();
        publication
            .write_inverse_to_stream(&candidate, &mut inverse)
            .expect("authored splice must invert");
        assert_eq!(inverse, source_archive);

        let fragment = package
            .allocate_source_part_splice_fragment(length, length)
            .expect("second fragment must allocate");
        drop(fragment);
        assert_eq!(budget.used(Resource::Memory), baseline);
    }
}

#[test]
fn authored_fragment_cannot_transfer_a_foreign_package_budget() {
    let source_archive = archive_bytes(VALID_SOURCE, false, false);
    let (budget, _cancellation, context) = managed_context_with_budget();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(source_archive.clone())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed source must open");
    let foreign = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(source_archive)))
        .expect("identical foreign package must open");
    let baseline = budget.used(Resource::Memory);
    let mut fragment = package
        .allocate_source_part_splice_fragment(11, 11)
        .expect("fragment must allocate");
    fragment.as_mut_slice().copy_from_slice(FRAGMENT);
    assert_eq!(budget.used(Resource::Memory), baseline + 11);
    let error = foreign
        .prepare_source_part_splice_with_fragment(
            &pack(DOCUMENT_URI),
            proof(&foreign, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            fragment,
            SourcePartSpliceLimits::default(),
        )
        .expect_err("even identical bytes cannot substitute a budget owner");
    assert!(error.to_string().contains("different source package"));
    assert_eq!(budget.used(Resource::Memory), baseline);
}

#[test]
fn authored_fragment_refusals_release_memory_and_preserve_typed_limits() {
    let archive = archive_bytes(VALID_SOURCE, false, false);
    let (budget, cancellation, context) = managed_context_with_budget();
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(archive)),
        litchi_opc::ReadLimits::default(),
        context.clone(),
    )
    .expect("managed source must open");
    let baseline = budget.used(Resource::Memory);
    let remaining = 64 * 1024 * 1024 - baseline;
    let occupied = context
        .reserve(Resource::Memory, remaining - 10)
        .expect("test must reserve all but ten remaining bytes");
    let error = package
        .allocate_source_part_splice_fragment(11, 11)
        .expect_err("one byte above remaining budget must refuse");
    assert!(
        matches!(error, OpcError::Execution(ExecutionError::ResourceLimit(ref limit))
        if limit.resource == Resource::Memory)
    );
    assert_eq!(budget.used(Resource::Memory), 64 * 1024 * 1024 - 10);
    drop(occupied);
    assert_eq!(budget.used(Resource::Memory), baseline);

    let fragment = package
        .allocate_source_part_splice_fragment(11, 11)
        .expect("fragment must allocate");
    let error = package
        .prepare_source_part_splice_with_fragment(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            fragment,
            SourcePartSpliceLimits::default(),
        )
        .expect_err("uninitialized authored bytes must fail proof validation");
    assert!(error.to_string().contains("fragment hash"));
    assert_eq!(budget.used(Resource::Memory), baseline);

    cancellation.cancel();
    let error = package
        .allocate_source_part_splice_fragment(11, 11)
        .expect_err("cancelled allocation must refuse");
    assert!(matches!(error, OpcError::Cancelled));
    assert_eq!(budget.used(Resource::Memory), baseline);
}

#[test]
fn authored_fragment_zero_fill_checks_cancellation_between_work_chunks() {
    const LENGTH: u64 = 1024 * 1024;
    let (budget, cancellation, context) = managed_context_with_budget();
    let source = Arc::new(WorkCancellingSource::new(
        archive_bytes(VALID_SOURCE, false, false),
        budget.clone(),
        cancellation,
    ));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source.clone(),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed source must open");
    let memory = budget.used(Resource::Memory);
    let work = budget.used(Resource::Work);
    source.arm(work);
    let error = package
        .allocate_source_part_splice_fragment(LENGTH, LENGTH)
        .expect_err("cancellation during initialization must stop allocation");
    assert!(matches!(error, OpcError::Cancelled));
    let consumed = budget.used(Resource::Work) - work;
    assert!(consumed > 0 && consumed < LENGTH);
    assert_eq!(budget.used(Resource::Memory), memory);
}

#[derive(Debug)]
struct WorkCancellingSource {
    inner: OwnedSource,
    budget: Budget,
    cancellation: CancellationSource,
    baseline: AtomicU64,
    armed: AtomicBool,
}

#[derive(Debug)]
struct InterruptedSource {
    inner: OwnedSource,
    armed: AtomicBool,
    cancellation: Option<CancellationSource>,
}

impl ReadAt for InterruptedSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.armed.swap(false, Ordering::AcqRel) {
            if let Some(cancellation) = &self.cancellation {
                cancellation.cancel();
            }
            return Err(io::ErrorKind::Interrupted.into());
        }
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.inner.version()
    }
}

#[test]
fn interrupted_splice_source_retries_without_double_charging_input() {
    let mut observations = Vec::new();
    for interrupt in [false, true] {
        let (budget, _cancellation, context) = managed_context_with_budget();
        let source = Arc::new(InterruptedSource {
            inner: OwnedSource::new(archive_bytes(VALID_SOURCE, true, false)),
            armed: AtomicBool::new(false),
            cancellation: None,
        });
        let package = SourceBackedPackage::from_read_at_with_execution_context(
            source.clone(),
            litchi_opc::ReadLimits::default(),
            context,
        )
        .expect("managed fixture must open");
        let input = budget.used(Resource::InputBytes);
        let work = budget.used(Resource::Work);
        let memory = budget.used(Resource::Memory);
        source.armed.store(interrupt, Ordering::Release);
        let plan = prepare(
            &package,
            VALID_SOURCE,
            INSERTION_OFFSET,
            FRAGMENT,
            SourcePartSpliceLimits::default(),
        );
        drop(plan);
        assert_eq!(budget.used(Resource::Memory), memory);
        observations.push((
            budget.used(Resource::InputBytes) - input,
            budget.used(Resource::Work) - work,
        ));
    }
    assert_eq!(
        observations[0].0, observations[1].0,
        "interruption accepts no bytes and must not consume input budget"
    );
    assert_eq!(
        observations[0].1 + 1,
        observations[1].1,
        "each interrupted retry consumes one bounded work unit"
    );
}

#[test]
fn interrupted_splice_source_observes_cancellation_before_retry() {
    let (budget, cancellation, context) = managed_context_with_budget();
    let source = Arc::new(InterruptedSource {
        inner: OwnedSource::new(archive_bytes(VALID_SOURCE, false, false)),
        armed: AtomicBool::new(false),
        cancellation: Some(cancellation),
    });
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source.clone(),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed fixture must open");
    let input = budget.used(Resource::InputBytes);
    let memory = budget.used(Resource::Memory);
    source.armed.store(true, Ordering::Release);
    let error = package
        .prepare_source_part_splice(
            &pack(DOCUMENT_URI),
            proof(&package, VALID_SOURCE, INSERTION_OFFSET, FRAGMENT),
            Arc::new(FRAGMENT.to_vec()),
            SourcePartSpliceLimits::default(),
        )
        .expect_err("cancelled interrupted reader must stop");
    assert!(matches!(error, OpcError::Cancelled));
    assert_eq!(budget.used(Resource::InputBytes), input);
    assert_eq!(budget.used(Resource::Memory), memory);
}

impl WorkCancellingSource {
    fn new(bytes: Vec<u8>, budget: Budget, cancellation: CancellationSource) -> Self {
        Self {
            inner: OwnedSource::new(bytes),
            budget,
            cancellation,
            baseline: AtomicU64::new(0),
            armed: AtomicBool::new(false),
        }
    }

    fn arm(&self, baseline: u64) {
        self.baseline.store(baseline, Ordering::Release);
        self.armed.store(true, Ordering::Release);
    }
}

impl ReadAt for WorkCancellingSource {
    fn len(&self) -> io::Result<u64> {
        self.inner.len()
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.inner.read_at(offset, output)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        if self.armed.load(Ordering::Acquire)
            && self.budget.used(Resource::Work) > self.baseline.load(Ordering::Acquire)
        {
            self.cancellation.cancel();
        }
        self.inner.version()
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
            0x5350_4c49_4345,
            self.revision.load(Ordering::Acquire),
        ))
    }
}
