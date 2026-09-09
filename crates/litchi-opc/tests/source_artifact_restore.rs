#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used,
    reason = "focused artifact-restore assertions intentionally panic on fixture errors"
)]

//! Contract coverage for the public exact-artifact restoration boundary.
//!
//! The original and current packages in these tests are opened through
//! independent positional sources.  A successful restore therefore proves
//! that the operation uses the supplied artifact and authenticated byte
//! identities, rather than a process-local package or source-version token.

use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, RwLock,
    atomic::{AtomicBool, AtomicU8, AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    OwnedSource, ReadAt, Resource, SourceVersion,
};
use litchi_opc::{
    OpcError, SourceArtifact, SourceArtifactFingerprint, SourceArtifactRestoreProof,
    SourceBackedPackage, SpliceResource,
};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const DOCUMENT_MEMBER: &str = "word/document.xml";
const LARGE_DOCUMENT_BYTES: usize = 192 * 1024;

fn document_bytes(extra: &[u8]) -> Vec<u8> {
    let body = vec![b'x'; LARGE_DOCUMENT_BYTES];
    let mut document = Vec::with_capacity(body.len() + extra.len() + 13);
    document.extend_from_slice(b"<root>");
    document.extend_from_slice(&body);
    document.extend_from_slice(extra);
    document.extend_from_slice(b"</root>");
    document
}

fn archive_bytes() -> Vec<u8> {
    archive_bytes_with_extra(&[])
}

fn archive_bytes_with_extra(extra: &[u8]) -> Vec<u8> {
    let document = document_bytes(extra);
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#
    );
    let relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rDoc" Type="{OFFICE_DOCUMENT_REL}" Target="{DOCUMENT_MEMBER}"/></Relationships>"#
    );
    let mut writer = StreamingArchiveWriter::new();
    writer
        .write_stored("[Content_Types].xml", content_types.as_bytes())
        .expect("content types fixture must be writable");
    writer
        .write_stored("_rels/.rels", relationships.as_bytes())
        .expect("package relationships fixture must be writable");
    writer
        .write_stored("opaque.bin", b"opaque bytes\0\xff")
        .expect("opaque fixture must be writable");
    writer
        .write_stored(DOCUMENT_MEMBER, &document)
        .expect("document fixture must be writable");
    writer
        .finish_to_bytes()
        .expect("fixture archive must finish")
}

fn open_unmanaged(bytes: Vec<u8>) -> SourceBackedPackage {
    SourceBackedPackage::from_read_at(Arc::new(OwnedSource::new(bytes)))
        .expect("fixture package must open")
}

fn proof_for(current: &SourceArtifact, original: &SourceArtifact) -> SourceArtifactRestoreProof {
    SourceArtifactRestoreProof {
        current_len: current.len(),
        current_sha256: current.fingerprint().expect("current artifact must hash"),
        original_len: original.len(),
        original_sha256: original.fingerprint().expect("original artifact must hash"),
    }
}

fn changed_fingerprint(fingerprint: SourceArtifactFingerprint) -> SourceArtifactFingerprint {
    let mut digest = fingerprint.into_sha256();
    digest[0] ^= 1;
    SourceArtifactFingerprint::from_sha256(digest)
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

#[derive(Debug)]
struct PrefixFailSink {
    bytes: Vec<u8>,
    remaining: usize,
}

impl PrefixFailSink {
    fn after(remaining: usize) -> Self {
        Self {
            bytes: Vec::new(),
            remaining,
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
        Ok(())
    }
}

struct CancellingPrefixSink {
    bytes: Vec<u8>,
    remaining: usize,
    cancellation: CancellationSource,
    cancelled: bool,
}

impl CancellingPrefixSink {
    fn after(remaining: usize, cancellation: CancellationSource) -> Self {
        Self {
            bytes: Vec::new(),
            remaining,
            cancellation,
            cancelled: false,
        }
    }
}

impl Write for CancellingPrefixSink {
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
        if !self.cancelled {
            self.cancellation.cancel();
            self.cancelled = true;
        }
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct MutableSource {
    bytes: RwLock<Vec<u8>>,
    revision: AtomicU64,
    mutate_after_read: AtomicBool,
}

impl MutableSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: RwLock::new(bytes),
            revision: AtomicU64::new(0),
            mutate_after_read: AtomicBool::new(false),
        })
    }

    fn arm_after_read(&self) {
        self.mutate_after_read.store(true, Ordering::Release);
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }

    fn mutate_tail_without_revision(&self) {
        let mut bytes = self.bytes.write().expect("mutable source lock must work");
        if let Some(last) = bytes.last_mut() {
            *last ^= 1;
        }
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
        drop(bytes);
        if self.mutate_after_read.swap(false, Ordering::AcqRel) {
            self.bump_revision();
        }
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x4152_5446_4143_5452,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

struct MutatingSink {
    bytes: Vec<u8>,
    current: Arc<MutableSource>,
    original: Arc<MutableSource>,
    on_flush: bool,
    mutated: bool,
}

impl MutatingSink {
    fn on_write(current: Arc<MutableSource>, original: Arc<MutableSource>) -> Self {
        Self {
            bytes: Vec::new(),
            current,
            original,
            on_flush: false,
            mutated: false,
        }
    }

    fn on_flush(current: Arc<MutableSource>, original: Arc<MutableSource>) -> Self {
        Self {
            bytes: Vec::new(),
            current,
            original,
            on_flush: true,
            mutated: false,
        }
    }

    fn mutate(&mut self) {
        if !self.mutated {
            self.current.bump_revision();
            self.original.bump_revision();
            self.mutated = true;
        }
    }
}

impl Write for MutatingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        if !self.on_flush {
            self.mutate();
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        if self.on_flush {
            self.mutate();
        }
        Ok(())
    }
}

struct SilentMutatingSink {
    bytes: Vec<u8>,
    original: Arc<MutableSource>,
    mutated: bool,
}

impl Write for SilentMutatingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        if !self.mutated {
            self.original.mutate_tail_without_revision();
            self.mutated = true;
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

const FAULT_NONE: u8 = 0;
const FAULT_TRUNCATED: u8 = 1;
const FAULT_OVERREPORTED: u8 = 2;

#[derive(Debug)]
struct FaultSource {
    bytes: Arc<Vec<u8>>,
    fault: AtomicU8,
}

impl FaultSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: Arc::new(bytes),
            fault: AtomicU8::new(FAULT_NONE),
        })
    }

    fn arm_truncated(&self) {
        self.fault.store(FAULT_TRUNCATED, Ordering::Release);
    }

    fn arm_overreported(&self) {
        self.fault.store(FAULT_OVERREPORTED, Ordering::Release);
    }
}

impl ReadAt for FaultSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len()).map_err(|error| io::Error::other(error.to_string()))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        match self.fault.load(Ordering::Acquire) {
            FAULT_TRUNCATED => return Ok(0),
            FAULT_OVERREPORTED => {
                return Ok(output
                    .len()
                    .checked_add(1)
                    .expect("test read request must leave one representable count"));
            },
            FAULT_NONE => {},
            _ => return Err(io::Error::other("unknown test source fault")),
        }
        let start = usize::try_from(offset).map_err(|_| io::Error::other("offset overflow"))?;
        if start >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - start);
        output[..count].copy_from_slice(&self.bytes[start..start + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x4152_5446_4155_4c54, 0))
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

fn managed_context(memory: u64, output: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let budget = Budget::root(
        "source-artifact-restore-test",
        Limits::new(memory, u64::MAX, output, u64::MAX, u64::MAX, u64::MAX),
    );
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("worker limit is nonzero"),
        NonZeroUsize::new(1).expect("operation limit is nonzero"),
        NonZeroU64::new(memory).expect("memory limit is nonzero"),
        0,
    )
    .expect("execution limits must be valid");
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

#[test]
fn restore_succeeds_with_independently_reopened_artifacts() {
    let original_bytes = archive_bytes();
    let current_bytes = archive_bytes_with_extra(b"current-source");
    assert_ne!(current_bytes, original_bytes);
    let current = open_unmanaged(current_bytes);
    let original = open_unmanaged(original_bytes.clone());
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    assert_ne!(current_artifact.len(), original_artifact.len());
    let proof = proof_for(&current_artifact, &original_artifact);

    let mut output = Vec::new();
    current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut output,
        )
        .expect("independent artifacts must restore exactly");
    assert_eq!(output, original_bytes);
}

#[test]
fn restore_rejects_wrong_hashes_and_lengths_before_output() {
    let bytes = archive_bytes();
    let current = open_unmanaged(bytes.clone());
    let original = open_unmanaged(bytes);
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let baseline = proof_for(&current_artifact, &original_artifact);
    let mut cases = Vec::new();

    let mut current_hash = baseline;
    current_hash.current_sha256 = changed_fingerprint(current_hash.current_sha256);
    cases.push(current_hash);

    let mut original_hash = baseline;
    original_hash.original_sha256 = changed_fingerprint(original_hash.original_sha256);
    cases.push(original_hash);

    let mut current_length = baseline;
    current_length.current_len += 1;
    cases.push(current_length);

    let mut original_length = baseline;
    original_length.original_len += 1;
    cases.push(original_length);

    for invalid in cases {
        let mut output = Vec::new();
        let error = current
            .restore_source_artifact_to_stream(
                &original_artifact,
                invalid,
                original_artifact.len(),
                &mut output,
            )
            .expect_err("invalid artifact proof must refuse");
        assert!(matches!(error, OpcError::SourceArtifactMismatch { .. }));
        assert!(output.is_empty(), "proof failure must precede output");
    }
}

#[test]
fn restore_rejects_finite_output_cap_before_output() {
    let bytes = archive_bytes();
    let current = open_unmanaged(bytes.clone());
    let original = open_unmanaged(bytes);
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let mut output = Vec::new();
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len() - 1,
            &mut output,
        )
        .expect_err("a one-byte-short output cap must refuse");
    assert!(matches!(
        error,
        OpcError::SourcePartSpliceLimit {
            resource: SpliceResource::OutputBytes,
            actual,
            maximum,
        } if actual == original_artifact.len() && maximum + 1 == actual
    ));
    assert!(output.is_empty());
}

#[test]
fn restore_reports_exact_prefix_for_short_sink() {
    const PREFIX: usize = 19;
    let bytes = archive_bytes();
    let current = open_unmanaged(bytes.clone());
    let original = open_unmanaged(bytes);
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let mut sink = PrefixFailSink::after(PREFIX);
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut sink,
        )
        .expect_err("a short sink must fail with typed progress");
    match error {
        OpcError::IncompleteOutput { written, .. } => {
            assert_eq!(written, PREFIX as u64);
        },
        other => panic!("unexpected short-sink error: {other:?}"),
    }
    assert_eq!(sink.bytes.len(), PREFIX);
}

#[test]
fn restore_detects_both_artifact_versions_during_authentication() {
    for mutate_current in [true, false] {
        let bytes = archive_bytes();
        let current_source = MutableSource::new(bytes.clone());
        let original_source = MutableSource::new(bytes);
        let current = SourceBackedPackage::from_read_at(current_source.clone())
            .expect("mutable current package must open");
        let original = SourceBackedPackage::from_read_at(original_source.clone())
            .expect("mutable original package must open");
        let current_artifact = current.source_artifact();
        let original_artifact = original.source_artifact();
        let proof = proof_for(&current_artifact, &original_artifact);
        if mutate_current {
            current_source.arm_after_read();
        } else {
            original_source.arm_after_read();
        }

        let mut output = Vec::new();
        let error = current
            .restore_source_artifact_to_stream(
                &original_artifact,
                proof,
                original_artifact.len(),
                &mut output,
            )
            .expect_err("source changes during authentication must refuse");
        assert!(
            contains_source_change(&error),
            "unexpected error: {error:?}"
        );
        assert!(output.is_empty());
    }
}

#[test]
fn restore_detects_both_artifact_versions_from_write_and_flush_callbacks() {
    for on_flush in [false, true] {
        let bytes = archive_bytes();
        let current_source = MutableSource::new(bytes.clone());
        let original_source = MutableSource::new(bytes);
        let current = SourceBackedPackage::from_read_at(current_source.clone())
            .expect("mutable current package must open");
        let original = SourceBackedPackage::from_read_at(original_source.clone())
            .expect("mutable original package must open");
        let current_artifact = current.source_artifact();
        let original_artifact = original.source_artifact();
        let proof = proof_for(&current_artifact, &original_artifact);
        let mut sink = if on_flush {
            MutatingSink::on_flush(current_source, original_source)
        } else {
            MutatingSink::on_write(current_source, original_source)
        };
        let error = current
            .restore_source_artifact_to_stream(
                &original_artifact,
                proof,
                original_artifact.len(),
                &mut sink,
            )
            .expect_err("sink-time source changes must refuse publication");
        assert!(
            contains_source_change(&error),
            "unexpected error: {error:?}"
        );
        assert!(!sink.bytes.is_empty());
    }
}

#[test]
fn restore_rehashes_original_after_output_started_when_version_is_stable() {
    let bytes = archive_bytes();
    let current = open_unmanaged(bytes.clone());
    let original_source = MutableSource::new(bytes);
    let original = SourceBackedPackage::from_read_at(original_source.clone())
        .expect("mutable original package must open");
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let mut sink = SilentMutatingSink {
        bytes: Vec::new(),
        original: original_source,
        mutated: false,
    };
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut sink,
        )
        .expect_err("a stable-version byte mutation must invalidate restore");
    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert_eq!(written, sink.bytes.len() as u64);
            assert!(matches!(
                *source,
                OpcError::SourceArtifactMismatch {
                    artifact: "original",
                    ..
                }
            ));
        },
        other => panic!("unexpected stable-version mutation error: {other:?}"),
    }
    assert!(!sink.bytes.is_empty());
}

#[test]
fn restore_rejects_truncated_and_overreported_providers_before_output() {
    for fault_current in [true, false] {
        for overreported in [false, true] {
            let bytes = archive_bytes();
            let current_source = FaultSource::new(bytes.clone());
            let original_source = FaultSource::new(bytes);
            let current = SourceBackedPackage::from_read_at(current_source.clone())
                .expect("faulting current package must open");
            let original = SourceBackedPackage::from_read_at(original_source.clone())
                .expect("faulting original package must open");
            let current_artifact = current.source_artifact();
            let original_artifact = original.source_artifact();
            let proof = proof_for(&current_artifact, &original_artifact);
            let fault_source = if fault_current {
                &current_source
            } else {
                &original_source
            };
            if overreported {
                fault_source.arm_overreported();
            } else {
                fault_source.arm_truncated();
            }

            let mut output = Vec::new();
            let error = current
                .restore_source_artifact_to_stream(
                    &original_artifact,
                    proof,
                    original_artifact.len(),
                    &mut output,
                )
                .expect_err("a malformed provider must refuse before output");
            assert!(
                matches!(error, OpcError::IoError(_)),
                "unexpected error: {error:?}"
            );
            assert!(output.is_empty(), "provider failure must precede output");
        }
    }
}

#[test]
fn restore_uses_current_context_for_all_io_work_and_output() {
    let bytes = archive_bytes();
    let (current_budget, _current_cancel, current_context) =
        managed_context(64 * 1024 * 1024, u64::MAX);
    let (original_budget, _original_cancel, original_context) =
        managed_context(64 * 1024 * 1024, u64::MAX);
    let current = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        current_context,
    )
    .expect("managed current package must open");
    let original = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        original_context,
    )
    .expect("managed original package must open");
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let current_input = current_budget.used(Resource::InputBytes);
    let current_work = current_budget.used(Resource::Work);
    let current_output = current_budget.used(Resource::OutputBytes);
    let original_input = original_budget.used(Resource::InputBytes);
    let original_work = original_budget.used(Resource::Work);
    let original_output = original_budget.used(Resource::OutputBytes);
    let mut output = Vec::new();

    current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut output,
        )
        .expect("restore must use the current execution context");
    assert_eq!(output, bytes);
    assert!(current_budget.used(Resource::InputBytes) > current_input);
    assert!(current_budget.used(Resource::Work) > current_work);
    assert_eq!(
        current_budget.used(Resource::OutputBytes) - current_output,
        bytes.len() as u64
    );
    assert_eq!(original_budget.used(Resource::InputBytes), original_input);
    assert_eq!(original_budget.used(Resource::Work), original_work);
    assert_eq!(original_budget.used(Resource::OutputBytes), original_output);
}

#[test]
fn restore_observes_original_context_cancellation_after_an_exact_short_prefix() {
    const PREFIX: usize = 23;
    let bytes = archive_bytes();
    let (current_budget, _current_cancel, current_context) =
        managed_context(64 * 1024 * 1024, u64::MAX);
    let (original_budget, original_cancel, original_context) =
        managed_context(64 * 1024 * 1024, u64::MAX);
    let current = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        current_context,
    )
    .expect("managed current package must open");
    let original = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes)),
        litchi_opc::ReadLimits::default(),
        original_context,
    )
    .expect("managed original package must open");
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let current_memory = current_budget.used(Resource::Memory);
    let original_memory = original_budget.used(Resource::Memory);
    let mut sink = CancellingPrefixSink::after(PREFIX, original_cancel);
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut sink,
        )
        .expect_err("original-context cancellation must stop short output");
    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert_eq!(written, PREFIX as u64);
            assert!(contains_cancellation(&source));
        },
        other => panic!("unexpected original-context cancellation: {other:?}"),
    }
    assert_eq!(sink.bytes.len(), PREFIX);
    assert_eq!(current_budget.used(Resource::Memory), current_memory);
    assert_eq!(original_budget.used(Resource::Memory), original_memory);
}

#[test]
fn managed_restore_releases_workspace_on_memory_hash_and_sink_failures() {
    const WORKSPACE: u64 = 64 * 1024;
    const PREFIX: usize = 29;
    let bytes = archive_bytes();
    let (budget, _cancellation, context) = managed_context(64 * 1024 * 1024, u64::MAX);
    let current = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        context.clone(),
    )
    .expect("managed current package must open");
    let original = open_unmanaged(bytes);
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let baseline_memory = budget.used(Resource::Memory);
    let available = budget
        .limit(Resource::Memory)
        .checked_sub(baseline_memory)
        .expect("managed package must fit inside its memory budget");
    assert!(available >= WORKSPACE);
    let occupied = context
        .reserve(
            Resource::Memory,
            available
                .checked_sub(WORKSPACE)
                .and_then(|remaining| remaining.checked_add(1))
                .expect("test reservation must fit the memory budget"),
        )
        .expect("test must leave less than one restore window available");
    let mut output = Vec::new();
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut output,
        )
        .expect_err("restore workspace must honor the memory limit");
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Memory
    ));
    assert!(output.is_empty());
    drop(occupied);
    assert_eq!(budget.used(Resource::Memory), baseline_memory);

    let mut wrong_hash = proof;
    wrong_hash.original_sha256 = changed_fingerprint(wrong_hash.original_sha256);
    let hash_baseline = budget.used(Resource::Memory);
    let mut output = Vec::new();
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            wrong_hash,
            original_artifact.len(),
            &mut output,
        )
        .expect_err("hash authentication must fail after releasing its workspace");
    assert!(matches!(
        error,
        OpcError::SourceArtifactMismatch {
            artifact: "original",
            field: "fingerprint",
        }
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::Memory), hash_baseline);

    let sink_baseline = budget.used(Resource::Memory);
    let mut sink = PrefixFailSink::after(PREFIX);
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut sink,
        )
        .expect_err("sink failure must release its restore workspace");
    assert!(
        matches!(error, OpcError::IncompleteOutput { written, .. } if written == PREFIX as u64)
    );
    assert_eq!(sink.bytes.len(), PREFIX);
    assert_eq!(budget.used(Resource::Memory), sink_baseline);
}

#[test]
fn managed_restore_charges_shared_output_once_and_releases_memory_lease() {
    let bytes = archive_bytes();
    let (budget, _cancellation, context) = managed_context(64 * 1024 * 1024, u64::MAX);
    let current = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        context.clone(),
    )
    .expect("managed current package must open");
    let original = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed original package must open");
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let baseline_memory = budget.used(Resource::Memory);
    let baseline_output = budget.used(Resource::OutputBytes);
    let mut output = Vec::new();
    current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut output,
        )
        .expect("managed exact restore must publish");
    assert_eq!(output, bytes);
    assert_eq!(
        budget.used(Resource::OutputBytes) - baseline_output,
        bytes.len() as u64,
        "shared restore must charge accepted output once"
    );
    assert_eq!(
        budget.used(Resource::Memory),
        baseline_memory,
        "the bounded restore workspace must be released"
    );
}

#[test]
fn restore_observes_cancellation_before_output_and_from_sink_callbacks() {
    let bytes = archive_bytes();
    let (budget, cancellation, context) = managed_context(64 * 1024 * 1024, u64::MAX);
    let current = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed package must open");
    let original = open_unmanaged(bytes);
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let baseline_memory = budget.used(Resource::Memory);

    cancellation.cancel();
    let mut output = Vec::new();
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut output,
        )
        .expect_err("pre-cancelled restore must refuse");
    assert!(matches!(error, OpcError::Cancelled));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::Memory), baseline_memory);

    let bytes = archive_bytes();
    let (budget, cancellation, context) = managed_context(64 * 1024 * 1024, u64::MAX);
    let current = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed package must open");
    let original = open_unmanaged(bytes);
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let baseline_memory = budget.used(Resource::Memory);
    let mut sink = CancellingSink::on_write(cancellation.clone());
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut sink,
        )
        .expect_err("sink cancellation must stop restore");
    assert!(contains_cancellation(&error), "unexpected error: {error:?}");
    assert!(!sink.bytes.is_empty());
    assert_eq!(budget.used(Resource::Memory), baseline_memory);

    let bytes = archive_bytes();
    let (budget, cancellation, context) = managed_context(64 * 1024 * 1024, u64::MAX);
    let current = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(OwnedSource::new(bytes.clone())),
        litchi_opc::ReadLimits::default(),
        context,
    )
    .expect("managed package must open");
    let original = open_unmanaged(bytes);
    let current_artifact = current.source_artifact();
    let original_artifact = original.source_artifact();
    let proof = proof_for(&current_artifact, &original_artifact);
    let baseline_memory = budget.used(Resource::Memory);
    let mut sink = CancellingSink::on_flush(cancellation);
    let error = current
        .restore_source_artifact_to_stream(
            &original_artifact,
            proof,
            original_artifact.len(),
            &mut sink,
        )
        .expect_err("flush cancellation must stop restore");
    assert!(contains_cancellation(&error), "unexpected error: {error:?}");
    assert_eq!(sink.bytes.len(), original_artifact.len() as usize);
    assert_eq!(budget.used(Resource::Memory), baseline_memory);
}
