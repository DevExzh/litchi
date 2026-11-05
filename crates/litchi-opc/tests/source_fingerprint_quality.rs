#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used,
    reason = "focused fingerprint controls use small generated fixtures"
)]

//! Focused controls for the source-artifact fingerprint boundary.
//!
//! These controls intentionally open through the public source-backed API.  The
//! managed case proves that a repeated fingerprint still performs the ordinary
//! budgeted read path; the generic cases prove that stable version tokens do not
//! authorize byte reuse and that malformed/short providers retain read-order
//! semantics.

use std::io;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, RwLock,
    atomic::{AtomicU64, AtomicUsize, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, OwnedSource, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::{OpcError, SourceBackedPackage, SourceLimits};
use litchi_opc::{SourceCacheLimits, SourceReadPolicy};
use sha2::{Digest as _, Sha256};
use soapberry_zip::office::StreamingArchiveWriter;

const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
const ROOT_RELS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rDoc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;

fn archive_bytes(payload: &[u8]) -> Vec<u8> {
    let mut writer =
        StreamingArchiveWriter::new(soapberry_zip::office::StreamingArchiveLimits::default());
    writer
        .write_stored("[Content_Types].xml", CONTENT_TYPES)
        .unwrap();
    writer.write_stored("_rels/.rels", ROOT_RELS).unwrap();
    writer.write_stored("word/document.xml", payload).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn managed_context(
    memory: u64,
    input_bytes: u64,
    work: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "opc-fingerprint-quality",
        Limits::new(memory, input_bytes, u64::MAX, u64::MAX, u64::MAX, work),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory.max(1)).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    (budget, cancellation_source, context)
}

#[derive(Debug)]
struct StableMutableSource {
    bytes: RwLock<Vec<u8>>,
    reads: AtomicUsize,
}

impl StableMutableSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: RwLock::new(bytes),
            reads: AtomicUsize::new(0),
        })
    }

    fn mutate_without_version(&self) {
        let mut bytes = self.bytes.write().unwrap();
        let last = bytes.last_mut().unwrap();
        *last ^= 1;
    }
}

impl ReadAt for StableMutableSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.read().unwrap().len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let bytes = self.bytes.read().unwrap();
        let offset = usize::try_from(offset).map_err(|_| io::Error::other("offset overflow"))?;
        if offset >= bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(bytes.len() - offset);
        output[..count].copy_from_slice(&bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x5354_4142_4c45, 0))
    }
}

#[derive(Debug)]
struct ShortReadSource {
    bytes: Arc<Vec<u8>>,
    max_read: usize,
    overreport: AtomicU64,
}

impl ShortReadSource {
    fn new(bytes: Vec<u8>, max_read: usize) -> Arc<Self> {
        Arc::new(Self {
            bytes: Arc::new(bytes),
            max_read: max_read.max(1),
            overreport: AtomicU64::new(0),
        })
    }

    fn arm_overreport(&self) {
        self.overreport.store(1, Ordering::Release);
    }
}

impl ReadAt for ShortReadSource {
    fn len(&self) -> io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.overreport.load(Ordering::Acquire) != 0 {
            return Ok(output.len().saturating_add(1));
        }
        let offset = usize::try_from(offset).map_err(|_| io::Error::other("offset overflow"))?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output
            .len()
            .min(self.max_read)
            .min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(0x0053_484f_5254, 0))
    }
}

#[test]
fn managed_fingerprint_repeated_calls_charge_work_and_input_and_release_memory() {
    let bytes = archive_bytes(b"<document>managed</document>");
    let (budget, cancellation, context) = managed_context(64 * 1024 * 1024, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(OwnedSource::new(bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let baseline_memory = budget.used(Resource::Memory);
    let before_input = budget.used(Resource::InputBytes);
    let before_work = budget.used(Resource::Work);

    let first = artifact.fingerprint().unwrap();
    let after_first_input = budget.used(Resource::InputBytes);
    let after_first_work = budget.used(Resource::Work);
    assert!(after_first_input > before_input);
    assert!(after_first_work > before_work);
    assert_eq!(budget.used(Resource::Memory), baseline_memory);

    assert_eq!(artifact.fingerprint().unwrap(), first);
    assert!(budget.used(Resource::InputBytes) > after_first_input);
    assert!(budget.used(Resource::Work) > after_first_work);
    assert_eq!(budget.used(Resource::Memory), baseline_memory);

    cancellation.cancel();
    assert!(matches!(artifact.fingerprint(), Err(OpcError::Cancelled)));
    assert_eq!(budget.used(Resource::Memory), baseline_memory);
}

#[test]
fn generic_stable_version_mutation_is_rehashed_on_every_fingerprint() {
    let bytes = archive_bytes(b"<document>stable</document>");
    let source = StableMutableSource::new(bytes);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let artifact = package.source_artifact();
    let first = artifact.fingerprint().unwrap();
    let reads_after_first = source.reads.load(Ordering::Acquire);

    source.mutate_without_version();
    let second = artifact.fingerprint().unwrap();

    assert_ne!(
        second, first,
        "stable version must not authorize stale bytes"
    );
    assert!(source.reads.load(Ordering::Acquire) > reads_after_first);
}

#[test]
fn generic_short_reads_are_complete_and_overreported_reads_fail_before_digest() {
    let bytes = archive_bytes(b"<document>short</document>");
    let expected = Sha256::digest(&bytes);
    let source = ShortReadSource::new(bytes.clone(), 7);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let artifact = package.source_artifact();
    assert_eq!(artifact.fingerprint().unwrap().as_sha256(), &expected[..]);

    source.arm_overreport();
    let error = artifact
        .fingerprint()
        .expect_err("an overreported read must fail before finalizing a digest");
    assert!(matches!(
        error,
        OpcError::IoError(error) if error.kind() == io::ErrorKind::InvalidData
    ));
}
