#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "these integration tests fail explicitly when their small OPC fixture is invalid"
)]

//! Public contract coverage for the opt-in bounded source-read window.
//!
//! The fixture deliberately contains one untyped physical member.  The
//! read-ahead policy is exercised only through the source-backed package API;
//! the archive writer is used here solely to make a representative OPC ZIP.

use std::collections::HashMap;
use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, Mutex, Weak,
    atomic::{AtomicBool, AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_opc::{
    OpcError, PackURI, ReadLimits, SourceArtifactFingerprint, SourceArtifactRestoreProof,
    SourceBackedPackage, SourceCacheLimits, SourcePartSpliceLimits, SourcePartSpliceProof,
    SourceReadPolicy, SourceTopologyPlan,
};
use sha2::{Digest as _, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;
use soapberry_zip::{PreservationIndex, ZipArchive};

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const HYPERLINK_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const DOCUMENT_URI: &str = "/word/document.xml";
const SECOND_MEMBER: &str = "word/styles.xml";
const SECOND_URI: &str = "/word/styles.xml";
const DOCUMENT_RELS_MEMBER: &str = "word/_rels/document.xml.rels";
const EXTERNAL_RELATIONSHIP_ID: &str = "rIdExternal";
const UNKNOWN_MEMBER: &str = "custom/opaque.bin";
const UNKNOWN_PAYLOAD: &[u8] = b"opaque vendor bytes remain byte exact\0\xff";
const DOCUMENT_PAYLOAD: &[u8] =
    b"<w:document xmlns:w=\"urn:example:word\"><w:p>read ahead</w:p></w:document>";
const REPLACEMENT_PAYLOAD: &[u8] =
    b"<w:document xmlns:w=\"urn:example:word\"><w:p>published</w:p></w:document>";
const SECOND_PAYLOAD: &[u8] = b"<w:styles xmlns:w=\"urn:example:word\"><w:style/></w:styles>";
const DOCUMENT_RELS_PAYLOAD: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/vendor" TargetMode="External"/></Relationships>"#;
const SPLICE_FRAGMENT: &[u8] = b"<w:bookmark/>";

fn uri(value: &str) -> PackURI {
    PackURI::new(value).expect("fixture URI must be valid")
}

fn digest(bytes: &[u8]) -> [u8; 32] {
    Sha256::digest(bytes).into()
}

fn splice_proof(package: &SourceBackedPackage, fragment: &[u8]) -> SourcePartSpliceProof {
    let insertion_offset = DOCUMENT_PAYLOAD.len() - b"</w:document>".len();
    let candidate = [
        &DOCUMENT_PAYLOAD[..insertion_offset],
        fragment,
        &DOCUMENT_PAYLOAD[insertion_offset..],
    ]
    .concat();
    SourcePartSpliceProof {
        source_version: package
            .source_version()
            .expect("fixture source version must be available"),
        source_len: DOCUMENT_PAYLOAD.len() as u64,
        source_sha256: digest(DOCUMENT_PAYLOAD),
        insertion_offset: insertion_offset as u64,
        fragment_len: fragment.len() as u64,
        fragment_sha256: digest(fragment),
        candidate_len: candidate.len() as u64,
        candidate_sha256: digest(&candidate),
    }
}

fn splice_candidate(fragment: &[u8]) -> Vec<u8> {
    let insertion_offset = DOCUMENT_PAYLOAD.len() - b"</w:document>".len();
    [
        &DOCUMENT_PAYLOAD[..insertion_offset],
        fragment,
        &DOCUMENT_PAYLOAD[insertion_offset..],
    ]
    .concat()
}

fn archive_bytes() -> Vec<u8> {
    archive_bytes_with_typed_unknown(false)
}

fn typed_archive_bytes() -> Vec<u8> {
    archive_bytes_with_typed_unknown(true)
}

fn archive_bytes_with_typed_unknown(typed_unknown: bool) -> Vec<u8> {
    let unknown_default = if typed_unknown {
        r#"<Default Extension="bin" ContentType="application/octet-stream"/>"#
    } else {
        ""
    };
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{unknown_default}</Types>"#
    );
    let root_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="{DOCUMENT_MEMBER}"/></Relationships>"#
    );

    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .expect("package relationships fixture must be writable");
    writer
        .write_deflated_sized(DOCUMENT_MEMBER, DOCUMENT_PAYLOAD)
        .expect("document fixture must be writable");
    writer
        .write_stored(DOCUMENT_RELS_MEMBER, DOCUMENT_RELS_PAYLOAD)
        .expect("document relationship fixture must be writable");
    writer
        .write_deflated_sized(SECOND_MEMBER, SECOND_PAYLOAD)
        .expect("second part fixture must be writable");
    // There is intentionally no `bin` Default in the manifest.  This is an
    // untyped and unreferenced physical member, so it must remain a reported
    // non-Part and survive every preservation path.
    writer
        .write_stored(UNKNOWN_MEMBER, UNKNOWN_PAYLOAD)
        .expect("unknown fixture member must be writable");
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

fn cache_limits() -> SourceCacheLimits {
    SourceCacheLimits::new(1024 * 1024, 1).expect("cache limits must be positive")
}

fn managed_context(memory: u64, input_bytes: u64) -> (Budget, ExecutionContext) {
    let budget = Budget::root(
        "opc-source-read-ahead-integration-test",
        Limits::new(memory, input_bytes, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker limit is nonzero"),
        NonZeroUsize::new(1).expect("in-flight task limit is nonzero"),
        NonZeroU64::new(memory.max(1)).expect("in-flight byte limit is nonzero"),
        0,
    )
    .expect("execution limits must be valid");
    (
        budget.clone(),
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn read_ahead_package(source: Arc<dyn ReadAt>, policy: SourceReadPolicy) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy(
        source,
        ReadLimits::default(),
        cache_limits(),
        policy,
    )
    .expect("source-backed read-ahead fixture must open")
}

fn managed_read_ahead_package(
    source: Arc<dyn ReadAt>,
    policy: SourceReadPolicy,
    context: ExecutionContext,
) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_source_read_policy_and_execution_context(
        source,
        ReadLimits::default(),
        cache_limits(),
        policy,
        context,
    )
    .expect("managed source-backed read-ahead fixture must open")
}

fn unknown_member(bytes: &[u8]) -> Vec<u8> {
    PhysPkgReader::new(bytes)
        .expect("published archive must be readable")
        .read_member(UNKNOWN_MEMBER)
        .expect("published archive must retain the unknown member")
}

#[derive(Debug, PartialEq, Eq)]
struct RawRecord {
    local: Vec<u8>,
    central: Vec<u8>,
}

fn raw_records(data: &[u8]) -> HashMap<Vec<u8>, RawRecord> {
    let archive = ZipArchive::from_slice(data)
        .expect("published archive must be valid ZIP")
        .into_zip_archive();
    let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut scratch)
        .expect("published archive must have preservable physical records");

    index
        .entries()
        .iter()
        .map(|entry| {
            let local = entry.local_span();
            let central = entry.central_record();
            (
                entry.raw_name_bytes().to_vec(),
                RawRecord {
                    local: data[local.start as usize..local.end as usize].to_vec(),
                    central: data[central.start as usize..central.end as usize].to_vec(),
                },
            )
        })
        .collect()
}

fn central_without_local_offset(record: &[u8]) -> Vec<u8> {
    let mut record = record.to_vec();
    // Replacing an earlier member may move this member's local header. The
    // offset is regenerated while every other central-directory byte remains
    // part of the physical unknown-member contract.
    record[42..46].fill(0);
    record
}

fn unknown_raw_record(data: &[u8]) -> RawRecord {
    // Keep the public physical-reader check beside the byte-level comparison:
    // the payload must remain readable through the same API callers use.
    assert_eq!(unknown_member(data), UNKNOWN_PAYLOAD);
    raw_records(data)
        .remove(UNKNOWN_MEMBER.as_bytes())
        .expect("archive must contain the unknown physical member")
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ReadCall {
    offset: u64,
    length: usize,
}

#[derive(Debug)]
struct LoggedSource {
    bytes: Arc<Vec<u8>>,
    calls: Mutex<Vec<ReadCall>>,
}

impl LoggedSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: Arc::new(bytes),
            calls: Mutex::new(Vec::new()),
        })
    }

    fn calls(&self) -> Vec<ReadCall> {
        self.calls
            .lock()
            .expect("source call log must not be poisoned")
            .clone()
    }
}

impl ReadAt for LoggedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "fixture is too large"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.calls
            .lock()
            .expect("source call log must not be poisoned")
            .push(ReadCall {
                offset,
                length: output.len(),
            });
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset is too large"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x5245_4144_4c4f_4747, 0))
    }
}

#[derive(Debug, Clone, Copy)]
struct ReentrantObservation {
    typed_refusal: bool,
    version_calls_before: u64,
    version_calls_after: u64,
}

#[derive(Debug)]
struct ReentrantSource {
    bytes: Arc<Vec<u8>>,
    owner: Mutex<Option<Weak<SourceBackedPackage>>>,
    observation: Mutex<Option<ReentrantObservation>>,
    armed: AtomicBool,
    version_calls: AtomicU64,
}

impl ReentrantSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: Arc::new(bytes),
            owner: Mutex::new(None),
            observation: Mutex::new(None),
            armed: AtomicBool::new(false),
            version_calls: AtomicU64::new(0),
        })
    }

    fn set_owner(&self, owner: Weak<SourceBackedPackage>) {
        *self
            .owner
            .lock()
            .expect("reentrant owner lock must not be poisoned") = Some(owner);
    }

    fn arm(&self) {
        self.armed.store(true, Ordering::Release);
    }

    fn observation(&self) -> Option<ReentrantObservation> {
        *self
            .observation
            .lock()
            .expect("reentrant observation lock must not be poisoned")
    }
}

impl ReadAt for ReentrantSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "fixture is too large"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.armed.swap(false, Ordering::AcqRel) {
            let owner = self
                .owner
                .lock()
                .expect("reentrant owner lock must not be poisoned")
                .clone()
                .and_then(|owner| owner.upgrade());
            if let Some(owner) = owner {
                let version_calls_before = self.version_calls.load(Ordering::Acquire);
                let result = owner.to_opc_package();
                let version_calls_after = self.version_calls.load(Ordering::Acquire);
                let observation = ReentrantObservation {
                    typed_refusal: matches!(
                        result,
                        Err(OpcError::SourceBackedOverlayUnavailable { .. })
                    ),
                    version_calls_before,
                    version_calls_after,
                };
                *self
                    .observation
                    .lock()
                    .expect("reentrant observation lock must not be poisoned") = Some(observation);
            }
        }

        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset is too large"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        self.version_calls.fetch_add(1, Ordering::AcqRel);
        Ok(SourceVersion::new(0x5245_454e_5452_414e, 0))
    }
}

struct FirstWriteMemorySink {
    budget: Budget,
    bytes: Vec<u8>,
    first_write_memory: Option<u64>,
}

impl FirstWriteMemorySink {
    fn new(budget: Budget) -> Self {
        Self {
            budget,
            bytes: Vec::new(),
            first_write_memory: None,
        }
    }
}

impl Write for FirstWriteMemorySink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !bytes.is_empty() && self.first_write_memory.is_none() {
            self.first_write_memory = Some(self.budget.used(Resource::Memory));
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_read_policy_is_bounded_and_exact_is_the_default_control() {
    assert!(SourceReadPolicy::forward_start(0).is_err());
    assert!(SourceReadPolicy::forward_start(65_537).is_err());
    assert!(SourceReadPolicy::forward_start(1).is_ok());
    assert!(SourceReadPolicy::forward_start(65_536).is_ok());

    let source = archive_bytes();
    let exact = SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(source.clone())))
        .expect("exact control package must open");
    assert!(
        exact
            .source_read_diagnostics()
            .expect("exact diagnostics must be readable")
            .is_none(),
        "existing constructors must remain exact and uninstrumented"
    );

    let default_policy = read_ahead_package(
        Arc::new(OwnedSource::new(source)),
        SourceReadPolicy::default(),
    );
    assert!(
        default_policy
            .source_read_diagnostics()
            .expect("default-policy diagnostics must be readable")
            .is_none(),
        "Default must select the exact policy"
    );

    let (_budget, managed_context) = managed_context(8 * 1024 * 1024, u64::MAX);
    let managed_exact =
        SourceBackedPackage::from_read_at_with_limits_and_cache_limits_and_execution_context(
            Arc::new(OwnedSource::new(archive_bytes())),
            ReadLimits::default(),
            cache_limits(),
            managed_context,
        )
        .expect("managed exact control package must open");
    assert!(
        managed_exact
            .source_read_diagnostics()
            .expect("managed exact diagnostics must be readable")
            .is_none(),
        "the existing managed constructor must retain exact reads"
    );
}

#[test]
fn package_publication_callback_refuses_before_recursive_source_version_fence() {
    let source = ReentrantSource::new(archive_bytes());
    let source_adapter: Arc<dyn ReadAt> = source.clone();
    let package = Arc::new(read_ahead_package(
        source_adapter,
        SourceReadPolicy::forward_start(1).expect("one-byte window policy must be valid"),
    ));
    source.set_owner(Arc::downgrade(&package));
    source.arm();

    let document = package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("semantic read must survive the reentrant publication refusal");
    assert_eq!(document.as_bytes(), DOCUMENT_PAYLOAD);
    drop(document);

    let observation = source
        .observation()
        .expect("the armed provider callback must reenter publication");
    assert!(
        observation.typed_refusal,
        "package publication must return the typed overlay-unavailable refusal"
    );
    assert_eq!(
        observation.version_calls_before, observation.version_calls_after,
        "the package publication guard must run before its source-version fence"
    );
}

#[test]
fn source_xml_capture_permanently_disables_opt_in_reads() {
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(archive_bytes())),
        SourceReadPolicy::forward_start(1).expect("one-byte window policy must be valid"),
    );
    let before = package
        .source_read_diagnostics()
        .expect("diagnostics must be readable before source XML capture")
        .expect("read-ahead diagnostics must be enabled");
    let source_xml = package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .source_xml()
        .expect("source XML capture must succeed");
    assert_eq!(source_xml.partname().as_str(), DOCUMENT_URI);
    let after = package
        .source_read_diagnostics()
        .expect("diagnostics must be readable after source XML capture")
        .expect("source XML capture must retain diagnostics");
    assert!(!after.enabled);
    assert_eq!(after.retained_window_bytes, 0);
    assert_eq!(after.requests, before.requests);
    assert_eq!(after.hits, before.hits);
    assert_eq!(after.misses, before.misses);
}

#[test]
fn exact_restore_noop_permanently_releases_the_opt_in_window() {
    let source = archive_bytes();
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(source.clone())),
        SourceReadPolicy::forward_start(1).expect("one-byte window policy must be valid"),
    );
    let original = package.source_artifact();
    let proof = SourceArtifactRestoreProof {
        current_len: original.len(),
        current_sha256: original
            .fingerprint()
            .expect("current source artifact must fingerprint"),
        original_len: original.len(),
        original_sha256: original
            .fingerprint()
            .expect("original source artifact must fingerprint"),
    };
    let before = package
        .source_read_diagnostics()
        .expect("diagnostics must be readable before exact restore")
        .expect("read-ahead diagnostics must be enabled");
    let mut output = Vec::new();
    package
        .restore_source_artifact_to_stream(&original, proof, original.len(), &mut output)
        .expect("exact restore must publish the original artifact");
    assert_eq!(output, source);
    let after = package
        .source_read_diagnostics()
        .expect("diagnostics must be readable after exact restore")
        .expect("exact restore must retain diagnostics");
    assert!(!after.enabled);
    assert_eq!(after.retained_window_bytes, 0);
    assert_eq!(after.requests, before.requests);
    assert_eq!(after.hits, before.hits);
    assert_eq!(after.misses, before.misses);
}

#[test]
fn managed_read_ahead_accounts_physical_input_and_retained_window_memory() {
    let source = archive_bytes();
    let (budget, context) = managed_context(8 * 1024 * 1024, u64::MAX);
    let package = managed_read_ahead_package(
        Arc::new(OwnedSource::new(source)),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
        context,
    );

    let before_input = budget.used(Resource::InputBytes);
    let before_memory = budget.used(Resource::Memory);
    let initial = package
        .source_read_diagnostics()
        .expect("read-ahead diagnostics must be available")
        .expect("enabled read-ahead must expose diagnostics");
    assert!(initial.enabled);
    assert_eq!(initial.configured_window_bytes, 4096);

    let data = package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("document payload must decode");
    assert_eq!(data.as_bytes(), DOCUMENT_PAYLOAD);
    drop(data);

    let after = package
        .source_read_diagnostics()
        .expect("read-ahead diagnostics must remain available")
        .expect("enabled read-ahead must expose diagnostics");
    assert!(after.requests > 0);
    assert_eq!(after.requests, after.hits + after.misses);
    assert!(after.fills <= after.misses);
    assert!(after.requested_bytes > 0);
    assert!(after.retained_window_bytes > 0);
    assert!(after.retained_window_bytes <= after.configured_window_bytes);

    let physical_input = budget.used(Resource::InputBytes);
    let retained_memory = budget.used(Resource::Memory);
    assert!(
        after.requests > initial.requests,
        "the semantic read must still be counted even when the archive is already cached"
    );
    assert!(
        after.hits > initial.hits,
        "the small fixture's semantic read should be served by the retained window"
    );
    assert!(physical_input >= before_input);
    assert!(retained_memory > before_memory);
    assert_eq!(
        physical_input, after.returned_bytes,
        "managed InputBytes must equal accepted physical read-ahead bytes before direct artifact reads"
    );
    assert!(
        retained_memory
            >= u64::try_from(after.retained_window_bytes).expect("window size fits u64"),
        "managed Memory must include the retained bounded read window"
    );
}

#[test]
fn exact_source_save_preserves_unknown_member_after_known_payload_read() {
    let source = archive_bytes();
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(source.clone())),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    assert_eq!(
        package.non_part_members().len(),
        1,
        "fixture must classify the opaque member as non-Part"
    );
    assert_eq!(package.non_part_members()[0].name(), UNKNOWN_MEMBER);
    assert_eq!(
        package
            .part(&uri(DOCUMENT_URI))
            .expect("document part must be present")
            .data()
            .expect("known payload must decode")
            .as_bytes(),
        DOCUMENT_PAYLOAD
    );

    let mut saved = Vec::new();
    package
        .source_artifact()
        .write_to_stream(&mut saved)
        .expect("exact source save must succeed");
    assert_eq!(saved, source, "exact source save must retain all ZIP bytes");
    assert_eq!(unknown_member(&saved), UNKNOWN_PAYLOAD);
}

#[test]
fn publication_disables_read_ahead_before_output_and_keeps_future_reads_exact() {
    let source = typed_archive_bytes();
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(source.clone())),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    let document = package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("document payload must decode");
    assert_eq!(document.as_bytes(), DOCUMENT_PAYLOAD);
    drop(document);
    let before_publication = package
        .source_read_diagnostics()
        .expect("diagnostics must be readable before publication")
        .expect("read-ahead diagnostics must be enabled");
    assert!(before_publication.enabled);
    assert!(before_publication.requests > 0);

    let materialized = package
        .to_opc_package()
        .expect("publication to the owning OPC package must succeed");
    let after_publication = package
        .source_read_diagnostics()
        .expect("diagnostics must remain readable after publication")
        .expect("a previously enabled policy retains diagnostics");
    assert!(!after_publication.enabled);
    assert_eq!(after_publication.retained_window_bytes, 0);
    assert_eq!(after_publication.requests, before_publication.requests);
    assert_eq!(after_publication.hits, before_publication.hits);
    assert_eq!(after_publication.misses, before_publication.misses);

    // The cache is one-entry, so materialization leaves the second Part warm
    // and evicts the document.  This semantic read must traverse the exact
    // source path after the publication switch.  Read-ahead counters must
    // stay frozen.
    let document_again = package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("future semantic read must remain valid");
    assert_eq!(document_again.as_bytes(), DOCUMENT_PAYLOAD);
    drop(document_again);
    let after_future_read = package
        .source_read_diagnostics()
        .expect("diagnostics must remain readable")
        .expect("read-ahead diagnostics must remain attached");
    assert!(!after_future_read.enabled);
    assert_eq!(after_future_read.retained_window_bytes, 0);
    assert_eq!(after_future_read.requests, after_publication.requests);
    assert_eq!(
        after_future_read.returned_bytes,
        after_publication.returned_bytes
    );

    let mut saved = Vec::new();
    materialized
        .to_stream(&mut saved)
        .expect("materialized package save must succeed");
    assert_eq!(unknown_member(&saved), UNKNOWN_PAYLOAD);
    let reopened = SourceBackedPackage::from_vec(saved).expect("saved package must reopen");
    assert_eq!(
        reopened
            .part(&uri(DOCUMENT_URI))
            .expect("published document must exist")
            .data()
            .expect("published document must decode")
            .as_bytes(),
        DOCUMENT_PAYLOAD
    );
}

fn assert_document_replacement_and_unknown_preservation(source: &[u8], output: &[u8]) {
    assert_eq!(
        PhysPkgReader::new(output)
            .expect("published archive must be readable")
            .read_member(DOCUMENT_MEMBER)
            .expect("published document must be readable"),
        REPLACEMENT_PAYLOAD
    );
    assert_eq!(unknown_member(output), UNKNOWN_PAYLOAD);
    let source_unknown = unknown_raw_record(source);
    let output_unknown = unknown_raw_record(output);
    assert_eq!(output_unknown.local, source_unknown.local);
    assert_eq!(
        central_without_local_offset(&output_unknown.central),
        central_without_local_offset(&source_unknown.central)
    );
}

fn assert_external_relationship_was_removed(output: &[u8]) {
    let reopened = SourceBackedPackage::from_vec(output.to_vec())
        .expect("published relationship-removal archive must reopen");
    assert_eq!(
        reopened
            .part(&uri(DOCUMENT_URI))
            .expect("published document must exist")
            .rels()
            .len(),
        0,
        "the selected external relationship must be absent after publication"
    );
    assert!(
        PhysPkgReader::new(output)
            .expect("published archive must be readable")
            .read_member(DOCUMENT_RELS_MEMBER)
            .expect("published document relationships must be readable")
            .windows(EXTERNAL_RELATIONSHIP_ID.len())
            .all(|window| window != EXTERNAL_RELATIONSHIP_ID.as_bytes()),
        "the regenerated relationships member must not retain the removed id"
    );
}

#[test]
fn one_part_external_relationship_removal_disables_opt_in_reads_before_publication() {
    let source = archive_bytes();
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(source.clone())),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    let mut output = Vec::new();
    package
        .write_part_overlay_with_external_relationship_removals_to_stream(
            &mut output,
            &uri(DOCUMENT_URI),
            REPLACEMENT_PAYLOAD.to_vec(),
            vec![EXTERNAL_RELATIONSHIP_ID.to_owned()],
        )
        .expect("one-part relationship removal must publish");

    assert_document_replacement_and_unknown_preservation(&source, &output);
    assert_external_relationship_was_removed(&output);
}

#[test]
fn batch_external_relationship_removal_disables_opt_in_reads_before_publication() {
    let source = archive_bytes();
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(source.clone())),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    let mut output = Vec::new();
    package
        .write_part_overlays_with_external_relationship_removals_to_stream(
            &mut output,
            vec![(
                uri(DOCUMENT_URI),
                REPLACEMENT_PAYLOAD.to_vec(),
                vec![EXTERNAL_RELATIONSHIP_ID.to_owned()],
            )],
        )
        .expect("batch relationship removal must publish");

    assert_document_replacement_and_unknown_preservation(&source, &output);
    assert_external_relationship_was_removed(&output);
}

#[test]
fn batch_replacement_and_deletion_preserve_unknown_physical_records() {
    let source = archive_bytes();
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(source.clone())),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    let mut output = Vec::new();
    package
        .write_part_overlays_with_deletions_to_stream(
            &mut output,
            vec![(uri(DOCUMENT_URI), REPLACEMENT_PAYLOAD.to_vec())],
            vec![uri(SECOND_URI)],
        )
        .expect("batch replacement and deletion must publish");

    assert_document_replacement_and_unknown_preservation(&source, &output);
    let physical = PhysPkgReader::new(&output).expect("published archive must be readable");
    assert!(
        !physical
            .member_names()
            .expect("published member names must be available")
            .iter()
            .any(|name| name == SECOND_MEMBER),
        "the selected Part must be physically deleted"
    );
}

#[test]
fn non_empty_topology_publication_disables_opt_in_reads_and_adds_a_typed_part() {
    let source = typed_archive_bytes();
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(source.clone())),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    let added = uri("/custom/topology.bin");
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        added.clone(),
        "application/octet-stream",
        b"topology payload".to_vec(),
    )
    .expect("typed topology addition must be accepted");
    plan.try_add_external_relationship(
        uri(DOCUMENT_URI),
        "rIdTopology",
        HYPERLINK_REL,
        "https://example.invalid/topology",
    )
    .expect("topology relationship addition must be accepted");

    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, plan)
        .expect("non-empty topology publication must succeed");
    assert_eq!(unknown_member(&output), UNKNOWN_PAYLOAD);
    assert_eq!(
        PhysPkgReader::new(&output)
            .expect("published archive must be readable")
            .read_member(added.membername())
            .expect("added topology member must be readable"),
        b"topology payload"
    );
    let reopened = SourceBackedPackage::from_vec(output).expect("topology output must reopen");
    assert_eq!(
        reopened
            .part(&added)
            .expect("added topology Part must exist")
            .data()
            .expect("added topology Part must decode")
            .as_bytes(),
        b"topology payload"
    );
    assert_eq!(
        reopened
            .part(&uri(DOCUMENT_URI))
            .expect("document Part must exist")
            .rels()
            .get("rIdTopology")
            .expect("topology relationship must exist")
            .target_mode(),
        litchi_opc::rel::TargetMode::External
    );
}

#[test]
fn precompressed_topology_addition_disables_opt_in_reads_before_physical_transfer() {
    let source = typed_archive_bytes();
    let package = read_ahead_package(
        Arc::new(OwnedSource::new(source.clone())),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    let source_part = package
        .part(&uri(SECOND_URI))
        .expect("leaf styles Part must be present");
    let authorized = source_part
        .authorize_precompressed(Arc::new(SECOND_PAYLOAD.to_vec()))
        .expect("the leaf styles Part must authorize its verified compressed payload");
    let added = uri("/custom/precompressed.xml");
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_precompressed_part(added.clone(), "application/xml", authorized)
        .expect("authorized precompressed addition must be accepted");

    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, plan)
        .expect("precompressed topology publication must succeed");
    assert_eq!(unknown_member(&output), UNKNOWN_PAYLOAD);
    assert_eq!(
        PhysPkgReader::new(&output)
            .expect("published archive must be readable")
            .read_member(added.membername())
            .expect("precompressed topology member must be readable"),
        SECOND_PAYLOAD
    );
    let reopened = SourceBackedPackage::from_vec(output).expect("topology output must reopen");
    assert_eq!(
        reopened
            .part(&added)
            .expect("precompressed topology Part must exist")
            .data()
            .expect("precompressed topology Part must decode")
            .as_bytes(),
        SECOND_PAYLOAD
    );
}

fn exact_splice_artifact() -> (Vec<u8>, Vec<ReadCall>) {
    let source = LoggedSource::new(archive_bytes());
    let source_adapter: Arc<dyn ReadAt> = source.clone();
    let package = read_ahead_package(source_adapter, SourceReadPolicy::exact());
    let proof = splice_proof(&package, SPLICE_FRAGMENT);
    let plan = package
        .prepare_source_part_splice(
            &uri(DOCUMENT_URI),
            proof,
            Arc::new(SPLICE_FRAGMENT.to_vec()),
            SourcePartSpliceLimits::default(),
        )
        .expect("exact splice control must prepare");
    let before_publication = source.calls().len();
    let mut output = Vec::new();
    plan.write_to_stream(&mut output)
        .expect("exact splice control must publish");
    let calls = source
        .calls()
        .into_iter()
        .skip(before_publication)
        .collect();
    (output, calls)
}

#[test]
fn expected_artifact_splice_preview_releases_opt_in_reads_and_runs_exact_twice() {
    let (expected, one_pass_calls) = exact_splice_artifact();
    assert!(!one_pass_calls.is_empty());
    let source = LoggedSource::new(archive_bytes());
    let source_adapter: Arc<dyn ReadAt> = source.clone();
    let package = read_ahead_package(
        source_adapter,
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    let document = package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("semantic warm read must decode");
    assert_eq!(document.as_bytes(), DOCUMENT_PAYLOAD);
    drop(document);
    let proof = splice_proof(&package, SPLICE_FRAGMENT);
    let plan = package
        .prepare_source_part_splice(
            &uri(DOCUMENT_URI),
            proof,
            Arc::new(SPLICE_FRAGMENT.to_vec()),
            SourcePartSpliceLimits::default(),
        )
        .expect("forward-policy splice must prepare");
    let before_publication = source.calls().len();
    let expected_len = u64::try_from(expected.len()).expect("fixture output length fits u64");
    let expected_hash = SourceArtifactFingerprint::from_sha256(digest(&expected));
    let mut output = Vec::new();
    plan.write_to_stream_with_expected_artifact(&mut output, expected_len, expected_hash)
        .expect("expected-artifact splice publication must succeed");
    assert_eq!(output, expected);

    let publication_calls: Vec<_> = source
        .calls()
        .into_iter()
        .skip(before_publication)
        .collect();
    let expected_calls: Vec<_> = one_pass_calls
        .iter()
        .chain(one_pass_calls.iter())
        .copied()
        .collect();
    assert_eq!(
        publication_calls, expected_calls,
        "expected-artifact preview and emission must use two exact physical traversals"
    );
    assert_eq!(
        SourceBackedPackage::from_vec(output.clone())
            .expect("published splice archive must reopen")
            .part(&uri(DOCUMENT_URI))
            .expect("published splice document must exist")
            .data()
            .expect("published splice document must decode")
            .as_bytes(),
        splice_candidate(SPLICE_FRAGMENT)
    );
    assert_eq!(unknown_member(&output), UNKNOWN_PAYLOAD);
}

fn exact_publication_calls() -> Vec<ReadCall> {
    let source = LoggedSource::new(archive_bytes());
    let source_adapter: Arc<dyn ReadAt> = source.clone();
    let package = read_ahead_package(source_adapter, SourceReadPolicy::exact());
    package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("known payload must decode");
    let before_publication = source.calls().len();

    let mut output = Vec::new();
    package
        .write_part_overlay_to_stream(
            &mut output,
            &uri(DOCUMENT_URI),
            REPLACEMENT_PAYLOAD.to_vec(),
        )
        .expect("exact publication control must succeed");
    source
        .calls()
        .into_iter()
        .skip(before_publication)
        .collect()
}

#[test]
fn managed_changed_overlay_releases_window_before_first_write_and_preserves_unknown_records() {
    let source_bytes = archive_bytes();
    let source = LoggedSource::new(source_bytes.clone());
    let source_adapter: Arc<dyn ReadAt> = source.clone();
    let (budget, context) = managed_context(8 * 1024 * 1024, u64::MAX);
    let package = managed_read_ahead_package(
        source_adapter,
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
        context,
    );
    let document = package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("known payload must decode");
    assert_eq!(document.as_bytes(), DOCUMENT_PAYLOAD);
    drop(document);

    let before_publication = source.calls().len();
    let before_diagnostics = package
        .source_read_diagnostics()
        .expect("diagnostics must be readable before publication")
        .expect("read-ahead diagnostics must be enabled");
    let before_memory = budget.used(Resource::Memory);
    assert!(before_diagnostics.retained_window_bytes > 0);

    let mut sink = FirstWriteMemorySink::new(budget.clone());
    package
        .write_part_overlay_to_stream(&mut sink, &uri(DOCUMENT_URI), REPLACEMENT_PAYLOAD.to_vec())
        .expect("managed changed overlay must publish");
    let first_write_memory = sink
        .first_write_memory
        .expect("publication must perform a non-empty first write");
    let retained_window = u64::try_from(before_diagnostics.retained_window_bytes)
        .expect("retained window size fits u64");
    assert!(
        first_write_memory < before_memory,
        "the first output write must observe the read-ahead window already released"
    );
    assert_eq!(
        before_memory - first_write_memory,
        retained_window,
        "the only managed memory released before output is the read-ahead window"
    );

    let publication_calls: Vec<_> = source
        .calls()
        .into_iter()
        .skip(before_publication)
        .collect();
    assert_eq!(
        publication_calls,
        exact_publication_calls(),
        "publication must use the exact physical source ranges after disabling read-ahead"
    );

    let output = sink.bytes;
    assert_eq!(unknown_member(&output), UNKNOWN_PAYLOAD);
    let source_unknown = unknown_raw_record(&source_bytes);
    let output_unknown = unknown_raw_record(&output);
    assert_eq!(output_unknown.local, source_unknown.local);
    assert_eq!(
        central_without_local_offset(&output_unknown.central),
        central_without_local_offset(&source_unknown.central)
    );
    assert_eq!(
        PhysPkgReader::new(&output)
            .expect("published archive must be readable")
            .read_member(DOCUMENT_MEMBER)
            .expect("published document must be readable"),
        REPLACEMENT_PAYLOAD
    );
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "managed publication must release package memory after the consumed package drops"
    );
}

#[test]
fn retained_source_artifact_does_not_pin_the_read_ahead_window() {
    let (budget, context) = managed_context(8 * 1024 * 1024, u64::MAX);
    let package = managed_read_ahead_package(
        Arc::new(OwnedSource::new(archive_bytes())),
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
        context,
    );
    let artifact = package.source_artifact();
    let data = package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("document payload must decode");
    drop(data);
    let diagnostics = package
        .source_read_diagnostics()
        .expect("diagnostics must be readable")
        .expect("read-ahead diagnostics must be enabled");
    assert!(diagnostics.retained_window_bytes > 0);
    assert!(budget.used(Resource::Memory) > 0);

    drop(package);
    assert_eq!(
        budget.used(Resource::Memory),
        0,
        "the exact source artifact must not retain the read-ahead allocation"
    );
    let mut saved = Vec::new();
    artifact
        .write_to_stream(&mut saved)
        .expect("retained source artifact must remain usable");
    assert_eq!(saved, archive_bytes());
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Arc<Vec<u8>>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: Arc::new(bytes),
            revision: AtomicU64::new(0),
        })
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "fixture is too large"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset is too large"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x5245_4144_4148_4541,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[test]
fn source_version_change_is_a_typed_publication_error_before_output() {
    let source = VersionedSource::new(archive_bytes());
    let source_adapter: Arc<dyn ReadAt> = source.clone();
    let package = read_ahead_package(
        source_adapter,
        SourceReadPolicy::forward_start(4096).expect("window policy must be valid"),
    );
    package
        .part(&uri(DOCUMENT_URI))
        .expect("document part must be present")
        .data()
        .expect("known payload must decode before source mutation");
    source.bump_revision();

    let mut output = Vec::new();
    let error = package
        .write_part_overlay_to_stream(
            &mut output,
            &uri(DOCUMENT_URI),
            REPLACEMENT_PAYLOAD.to_vec(),
        )
        .expect_err("a changed source must reject publication");
    assert!(matches!(error, OpcError::SourceChanged { .. }));
    assert!(
        output.is_empty(),
        "source-version refusal must happen before publication output"
    );
}
