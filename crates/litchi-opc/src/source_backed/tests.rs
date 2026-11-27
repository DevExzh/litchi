#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "test assertions panic on failure by design"
)]

use super::*;
use crate::{SourceCacheLimits, SourceReadPolicy};
use std::collections::HashMap;
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::Barrier;
use std::sync::atomic::{AtomicBool, AtomicU64, AtomicUsize, Ordering};
use std::time::Duration;

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits, Resource,
};

struct CountingSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
    reads: AtomicUsize,
    versions: AtomicUsize,
    read_bytes: AtomicU64,
    read_ranges: Mutex<Vec<(u64, usize)>>,
    max_read: usize,
}

impl CountingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
            reads: AtomicUsize::new(0),
            versions: AtomicUsize::new(0),
            read_bytes: AtomicU64::new(0),
            read_ranges: Mutex::new(Vec::new()),
            max_read: usize::MAX,
        }
    }

    fn chunked(bytes: Vec<u8>, max_read: usize) -> Self {
        let mut source = Self::new(bytes);
        source.max_read = max_read.max(1);
        source
    }

    fn changed(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn read_ranges(&self) -> Vec<(u64, usize)> {
        self.read_ranges
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .clone()
    }
}

impl ReadAt for CountingSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        self.reads.fetch_add(1, Ordering::SeqCst);
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        let count = count.min(self.max_read);
        self.read_ranges
            .lock()
            .unwrap_or_else(|poisoned| poisoned.into_inner())
            .push((offset as u64, count));
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        self.read_bytes.fetch_add(count as u64, Ordering::SeqCst);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        self.versions.fetch_add(1, Ordering::SeqCst);
        Ok(SourceVersion::new(42, self.revision.load(Ordering::SeqCst)))
    }
}

struct CancelOnHitVersionSource {
    bytes: Vec<u8>,
    cancellation_source: CancellationSource,
    skip_versions: AtomicUsize,
    armed: AtomicBool,
}

impl CancelOnHitVersionSource {
    fn new(bytes: Vec<u8>, cancellation_source: CancellationSource) -> Self {
        Self {
            bytes,
            cancellation_source,
            skip_versions: AtomicUsize::new(0),
            armed: AtomicBool::new(false),
        }
    }

    fn arm_after_versions(&self, skip_versions: usize) {
        self.skip_versions.store(skip_versions, Ordering::SeqCst);
        self.armed.store(true, Ordering::SeqCst);
    }
}

impl ReadAt for CancelOnHitVersionSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        if self.armed.load(Ordering::SeqCst)
            && self
                .skip_versions
                .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |remaining| {
                    remaining.checked_sub(1)
                })
                .is_err()
        {
            self.armed.store(false, Ordering::SeqCst);
            self.cancellation_source.cancel();
        }
        Ok(SourceVersion::new(43, 0))
    }
}

struct ChangeOnHitVersionSource {
    bytes: Vec<u8>,
    revision: AtomicU64,
    skip_versions: AtomicUsize,
    armed: AtomicBool,
}

impl ChangeOnHitVersionSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            revision: AtomicU64::new(0),
            skip_versions: AtomicUsize::new(0),
            armed: AtomicBool::new(false),
        }
    }

    fn arm_after_versions(&self, skip_versions: usize) {
        self.skip_versions.store(skip_versions, Ordering::SeqCst);
        self.armed.store(true, Ordering::SeqCst);
    }
}

impl ReadAt for ChangeOnHitVersionSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        if self.armed.load(Ordering::SeqCst)
            && self
                .skip_versions
                .fetch_update(Ordering::SeqCst, Ordering::SeqCst, |remaining| {
                    remaining.checked_sub(1)
                })
                .is_err()
        {
            self.armed.store(false, Ordering::SeqCst);
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
        Ok(SourceVersion::new(
            102,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

struct OverReportingSource {
    bytes: Vec<u8>,
    overreport: AtomicBool,
}

impl OverReportingSource {
    fn new(bytes: Vec<u8>) -> Self {
        Self {
            bytes,
            overreport: AtomicBool::new(false),
        }
    }
}

impl ReadAt for OverReportingSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        if self.overreport.load(Ordering::SeqCst) {
            return Ok(output.len().saturating_add(1));
        }
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(93, 0))
    }
}

struct CatalogReadFailureSource {
    bytes: Vec<u8>,
    fail_offset: usize,
}

impl CatalogReadFailureSource {
    fn new(bytes: Vec<u8>, fail_offset: usize) -> Self {
        Self { bytes, fail_offset }
    }
}

impl ReadAt for CatalogReadFailureSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        let end = offset.saturating_add(output.len());
        if offset <= self.fail_offset && self.fail_offset < end {
            return Err(std::io::Error::new(
                std::io::ErrorKind::BrokenPipe,
                "catalog-stage read failed",
            ));
        }
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(94, 0))
    }
}

struct OverReportingSink {
    calls: usize,
    accepted: usize,
}

impl Write for OverReportingSink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.calls = self.calls.saturating_add(1);
        Ok(bytes.len().saturating_add(1))
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

/// The furthest before a member's payload that a read *of that member* can
/// begin: the fixed local header plus the widest local variable region
/// measured in this repository's corpus.
const MEMBER_RECORD_PREFIX_BYTES: usize = 640;

/// Whether one positional read fetches the payload of the member whose
/// first payload byte is at `payload`.
///
/// The test sources below fire — sleep, cancel, bump their version, park on
/// a barrier — when the member they watch is fetched, and used to recognise
/// that read by its start offset, because a member's payload arrived in a
/// read that *began* at it. Today a member's first read is one bounded read
/// of its whole local record, which begins at the member's local header, so
/// the test is now that the read delivers the payload's first byte *and*
/// began inside the member's own local record. Both halves matter:
/// delivering the byte distinguishes the member read from the short
/// local-record probes issued at the same offset, and beginning inside the
/// member's own record keeps a bulk copy that spans the member from
/// counting as a fetch of it, which is how these triggers behaved before.
///
/// The non-empty requirement excludes the zero-length terminating read a
/// range reader issues at the end of a member.
fn read_fetches_payload(offset: usize, requested: usize, payload: usize) -> bool {
    requested > 0
        && offset <= payload
        && payload - offset < requested
        && payload - offset <= MEMBER_RECORD_PREFIX_BYTES
}

struct SlowPayloadSource {
    bytes: Vec<u8>,
    payload_offset: usize,
    payload_reads: AtomicUsize,
}

impl SlowPayloadSource {
    /// Whether one positional read fetches this member's payload.
    ///
    /// See [`read_fetches_payload`].
    fn fetches_payload(&self, offset: usize, requested: usize) -> bool {
        read_fetches_payload(offset, requested, self.payload_offset)
    }
    fn new(bytes: Vec<u8>, payload_offset: usize) -> Self {
        Self {
            bytes,
            payload_offset,
            payload_reads: AtomicUsize::new(0),
        }
    }
}

impl ReadAt for SlowPayloadSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if self.fetches_payload(offset, output.len()) {
            self.payload_reads.fetch_add(1, Ordering::SeqCst);
            // Keep the cold load in flight long enough for the peer to
            // enter the same part concurrently.
            std::thread::sleep(Duration::from_millis(100));
        }
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(77, 0))
    }
}

#[cfg(feature = "performance-diagnostics")]
struct GatedPayloadSource {
    bytes: Vec<u8>,
    payload_offset: usize,
    payload_gate_entered: AtomicBool,
    payload_gate_released: AtomicBool,
    payload_gate_armed: AtomicBool,
    payload_reads: AtomicUsize,
}

#[cfg(feature = "performance-diagnostics")]
impl GatedPayloadSource {
    /// Whether one positional read fetches this member's payload.
    ///
    /// See [`read_fetches_payload`].
    fn fetches_payload(&self, offset: usize, requested: usize) -> bool {
        read_fetches_payload(offset, requested, self.payload_offset)
    }
    fn new(bytes: Vec<u8>, payload_offset: usize) -> Self {
        Self {
            bytes,
            payload_offset,
            payload_gate_entered: AtomicBool::new(false),
            payload_gate_released: AtomicBool::new(false),
            payload_gate_armed: AtomicBool::new(true),
            payload_reads: AtomicUsize::new(0),
        }
    }

    fn wait_until_payload_read(&self, timeout: Duration) -> bool {
        let deadline = std::time::Instant::now() + timeout;
        while !self.payload_gate_entered.load(Ordering::Acquire)
            && std::time::Instant::now() < deadline
        {
            std::thread::yield_now();
        }
        self.payload_gate_entered.load(Ordering::Acquire)
    }

    fn release_payload_read(&self) {
        self.payload_gate_released.store(true, Ordering::Release);
    }
}

#[cfg(feature = "performance-diagnostics")]
impl ReadAt for GatedPayloadSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if self.fetches_payload(offset, output.len()) {
            self.payload_reads.fetch_add(1, Ordering::SeqCst);
            if self.payload_gate_armed.swap(false, Ordering::AcqRel) {
                self.payload_gate_entered.store(true, Ordering::Release);
                while !self.payload_gate_released.load(Ordering::Acquire) {
                    std::thread::yield_now();
                }
            }
        }
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(78, 0))
    }
}

struct GatedCorruptSource {
    bytes: Vec<u8>,
    payload_offset: usize,
    payload_gate_entered: Arc<Barrier>,
    payload_gate_release: Arc<Barrier>,
    payload_gate_armed: AtomicBool,
    payload_reads: AtomicUsize,
}

impl GatedCorruptSource {
    /// Whether one positional read fetches this member's payload.
    ///
    /// See [`read_fetches_payload`].
    fn fetches_payload(&self, offset: usize, requested: usize) -> bool {
        read_fetches_payload(offset, requested, self.payload_offset)
    }
    fn new(bytes: Vec<u8>, payload_offset: usize) -> Self {
        Self {
            bytes,
            payload_offset,
            payload_gate_entered: Arc::new(Barrier::new(2)),
            payload_gate_release: Arc::new(Barrier::new(2)),
            payload_gate_armed: AtomicBool::new(true),
            payload_reads: AtomicUsize::new(0),
        }
    }

    fn wait_until_payload_read(&self) {
        self.payload_gate_entered.wait();
    }

    fn release_payload_read(&self) {
        self.payload_gate_release.wait();
    }
}

impl ReadAt for GatedCorruptSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if self.fetches_payload(offset, output.len()) {
            self.payload_reads.fetch_add(1, Ordering::SeqCst);
            if self.payload_gate_armed.swap(false, Ordering::SeqCst) {
                self.payload_gate_entered.wait();
                self.payload_gate_release.wait();
            }
        }
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(101, 0))
    }
}

struct CancelDuringPayloadSource {
    bytes: Vec<u8>,
    payload_offset: usize,
    cancellation_source: CancellationSource,
    armed: AtomicBool,
}

struct CancelDuringOpenSource {
    bytes: Vec<u8>,
    cancellation_source: CancellationSource,
    reads: AtomicUsize,
    cancel_after: usize,
}

impl CancelDuringOpenSource {
    fn new(bytes: Vec<u8>, cancellation_source: CancellationSource) -> Self {
        Self {
            bytes,
            cancellation_source,
            reads: AtomicUsize::new(0),
            cancel_after: 1,
        }
    }
}

impl ReadAt for CancelDuringOpenSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        if self.reads.fetch_add(1, Ordering::SeqCst) + 1 >= self.cancel_after {
            self.cancellation_source.cancel();
        }
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(95, 0))
    }
}

impl CancelDuringPayloadSource {
    /// Whether one positional read fetches this member's payload.
    ///
    /// See [`read_fetches_payload`].
    fn fetches_payload(&self, offset: usize, requested: usize) -> bool {
        read_fetches_payload(offset, requested, self.payload_offset)
    }
    fn new(bytes: Vec<u8>, payload_offset: usize, cancellation_source: CancellationSource) -> Self {
        Self {
            bytes,
            payload_offset,
            cancellation_source,
            armed: AtomicBool::new(true),
        }
    }
}

impl ReadAt for CancelDuringPayloadSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        if self.fetches_payload(offset, output.len()) && self.armed.swap(false, Ordering::SeqCst) {
            // The bytes have been decompressed into the loader's private
            // allocation, but the publication checks must reject them.
            self.cancellation_source.cancel();
        }
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(79, 0))
    }
}

struct ChangeDuringPayloadSource {
    bytes: Vec<u8>,
    payload_offset: usize,
    revision: AtomicU64,
    armed: AtomicBool,
}

impl ChangeDuringPayloadSource {
    /// Whether one positional read fetches this member's payload.
    ///
    /// See [`read_fetches_payload`].
    fn fetches_payload(&self, offset: usize, requested: usize) -> bool {
        read_fetches_payload(offset, requested, self.payload_offset)
    }
    fn new(bytes: Vec<u8>, payload_offset: usize) -> Self {
        Self {
            bytes,
            payload_offset,
            revision: AtomicU64::new(0),
            armed: AtomicBool::new(false),
        }
    }
}

impl ReadAt for ChangeDuringPayloadSource {
    fn len(&self) -> std::io::Result<u64> {
        Ok(self.bytes.len() as u64)
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> std::io::Result<usize> {
        let offset = usize::try_from(offset).map_err(|_| {
            std::io::Error::new(std::io::ErrorKind::InvalidInput, "offset too large")
        })?;
        if offset >= self.bytes.len() {
            return Ok(0);
        }
        let count = output.len().min(self.bytes.len() - offset);
        output[..count].copy_from_slice(&self.bytes[offset..offset + count]);
        if self.fetches_payload(offset, output.len()) && self.armed.swap(false, Ordering::SeqCst) {
            self.revision.fetch_add(1, Ordering::SeqCst);
        }
        Ok(count)
    }

    fn version(&self) -> std::io::Result<SourceVersion> {
        Ok(SourceVersion::new(88, self.revision.load(Ordering::SeqCst)))
    }
}

fn managed_context_with_cancellation(
    memory: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let (budget, cancellation_source, context) =
        managed_context_with_resources(memory, u64::MAX, u64::MAX, u64::MAX);
    (budget, cancellation_source, context)
}

fn managed_context_with_all_resources(
    memory: u64,
    input_bytes: u64,
    output_bytes: u64,
    objects: u64,
    work: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "opc-source-cache-test",
        Limits::new(memory, input_bytes, output_bytes, objects, u64::MAX, work),
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

fn managed_context_with_resources(
    memory: u64,
    input_bytes: u64,
    objects: u64,
    work: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    managed_context_with_all_resources(memory, input_bytes, u64::MAX, objects, work)
}

fn managed_context_with_output(
    output_bytes: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let memory = SOURCE_PUBLICATION_CHUNK_BYTES as u64 + 4096;
    managed_context_with_all_resources(memory, u64::MAX, output_bytes, u64::MAX, u64::MAX)
}

fn managed_context(memory: u64) -> (Budget, ExecutionContext) {
    let (budget, cancellation_source, context) = managed_context_with_cancellation(memory);
    drop(cancellation_source);
    (budget, context)
}

fn archive_bytes(root_relationships: &[u8], document: &[u8], include_junk: bool) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships)
        .unwrap();
    writer.write_stored("word/document.xml", document).unwrap();
    writer
        .write_stored("custom/orphan.xml", b"<orphan/>")
        .unwrap();
    if include_junk {
        writer.write_stored("scratch.bin", b"not a part").unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn zip32_entry_count_boundary_source() -> Vec<u8> {
    const PART_COUNT: usize = (u16::MAX as usize - 1) - 2;
    const CONTENT_TYPES: &[u8] = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="bin" ContentType="application/vnd.ms-excel.sheet.binary.macroEnabled.main"/><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;

    let package_relationships = Relationships::new(PACKAGE_URI.to_owned()).to_xml();
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
        .write_stored("[Content_Types].xml", CONTENT_TYPES)
        .expect("write boundary content types");
    writer
        .write_stored("_rels/.rels", package_relationships.as_bytes())
        .expect("write boundary package relationships");
    for index in 0..PART_COUNT {
        let name = format!("custom/opaque-{index:05}.bin");
        let payload = [u8::try_from(index % 251).expect("boundary payload byte")];
        writer
            .write_stored(&name, &payload)
            .expect("write boundary opaque Part");
    }
    writer
        .finish_to_bytes()
        .expect("finish boundary source archive")
}

fn read_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(bytes[offset..offset + 2].try_into().unwrap())
}

fn read_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn promote_to_zip64(mut bytes: Vec<u8>) -> Vec<u8> {
    let archive = soapberry_zip::ZipArchive::from_slice(&bytes).expect("parse ZIP");
    let eocd = archive.eocd_offset() as usize;
    let entries = u64::from(read_u16(&bytes, eocd + 10));
    let central_size = u64::from(read_u32(&bytes, eocd + 12));
    let central_offset = u64::from(read_u32(&bytes, eocd + 16));
    let mut records = Vec::new();
    records.extend_from_slice(&0x0606_4b50_u32.to_le_bytes());
    records.extend_from_slice(&44_u64.to_le_bytes());
    records.extend_from_slice(&45_u16.to_le_bytes());
    records.extend_from_slice(&45_u16.to_le_bytes());
    records.extend_from_slice(&0_u32.to_le_bytes());
    records.extend_from_slice(&0_u32.to_le_bytes());
    records.extend_from_slice(&entries.to_le_bytes());
    records.extend_from_slice(&entries.to_le_bytes());
    records.extend_from_slice(&central_size.to_le_bytes());
    records.extend_from_slice(&central_offset.to_le_bytes());
    records.extend_from_slice(&0x0706_4b50_u32.to_le_bytes());
    records.extend_from_slice(&0_u32.to_le_bytes());
    records.extend_from_slice(&(eocd as u64).to_le_bytes());
    records.extend_from_slice(&1_u32.to_le_bytes());
    bytes.splice(eocd..eocd, records);
    let ordinary_eocd = eocd + 76;
    bytes[ordinary_eocd + 8..ordinary_eocd + 12].fill(0xff);
    bytes[ordinary_eocd + 12..ordinary_eocd + 20].fill(0xff);
    bytes
}

fn add_zip64_tail_extension(mut bytes: Vec<u8>, extension: &[u8]) -> Vec<u8> {
    let archive = soapberry_zip::ZipArchive::from_slice(&bytes).expect("parse ZIP64");
    assert!(archive.is_zip64());
    let head = archive.head_eocd_offset() as usize;
    let old_size = u64::from_le_bytes(bytes[head + 4..head + 12].try_into().unwrap());
    bytes.splice(head + 56..head + 56, extension.iter().copied());
    bytes[head + 4..head + 12].copy_from_slice(
        &old_size
            .checked_add(extension.len() as u64)
            .unwrap()
            .to_le_bytes(),
    );
    bytes
}

fn zip64_tail_extensible_data(data: &[u8]) -> Vec<u8> {
    let archive = soapberry_zip::ZipArchive::from_slice(data).expect("parse ZIP64");
    let head = archive.head_eocd_offset() as usize;
    let locator = archive.eocd_offset() as usize - 20;
    data[head + 56..locator].to_vec()
}

fn mixed_compression_archive() -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships())
        .unwrap();
    writer
        .write_deflated(
            "word/document.xml",
            b"<document>deflated source payload</document>",
        )
        .unwrap();
    writer
        .write_deflated(
            "custom/second.xml",
            b"<second>deflated source payload</second>",
        )
        .unwrap();
    writer
        .write_stored_stream(
            "custom/orphan.xml",
            std::io::Cursor::new(b"<orphan>stored source payload</orphan>".to_vec()),
        )
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[test]
fn validation_archive_index_preserves_raw_source_io_errors() {
    let source = Arc::new(OverReportingSource::new(archive_bytes(
        root_relationships(),
        b"validation ingress",
        false,
    )));
    source.overreport.store(true, Ordering::SeqCst);

    let error =
        match SourceBackedPackage::from_read_at_for_validation(source, SourceLimits::default()) {
            Ok(_) => panic!("over-reported source reads must fail archive indexing"),
            Err(error) => error,
        };
    assert_eq!(error.phase, ValidationCatalogPhase::Ingress);
    assert!(matches!(
        error.error,
        OpcError::IoError(error) if error.kind() == std::io::ErrorKind::InvalidData
    ));
}

#[test]
fn validation_catalog_preserves_raw_source_io_errors() {
    let bytes = archive_bytes(root_relationships(), b"validation catalog", false);
    let marker = b"<Types xmlns=\"";
    let fail_offset = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("the content-types marker must be present in the fixture");
    let source = Arc::new(CatalogReadFailureSource::new(bytes, fail_offset));

    let error =
        match SourceBackedPackage::from_read_at_for_validation(source, SourceLimits::default()) {
            Ok(_) => panic!("the catalog-stage source read must fail"),
            Err(error) => error,
        };
    assert_eq!(error.phase, ValidationCatalogPhase::Catalog);
    assert!(matches!(
        error.error,
        OpcError::IoError(error) if error.kind() == std::io::ErrorKind::BrokenPipe
    ));
}

#[test]
fn normal_open_catalog_preserves_raw_source_io_errors() {
    let bytes = archive_bytes(root_relationships(), b"normal catalog", false);
    let marker = b"<Types xmlns=\"";
    let fail_offset = bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("the content-types marker must be present in the fixture");
    let source = Arc::new(CatalogReadFailureSource::new(bytes, fail_offset));

    let error = match SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    ) {
        Ok(_) => panic!("the catalog-stage source read must fail"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        OpcError::IoError(error) if error.kind() == std::io::ErrorKind::BrokenPipe
    ));
}

#[test]
fn owned_bytes_open_keeps_ordinary_payloads_deferred() {
    let package = SourceBackedPackage::from_vec(
        archive_bytes(root_relationships(), b"<document/>", true),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();

    let opened = package.cache_diagnostics();
    assert_eq!(opened.cold_loads, 0);
    assert_eq!(opened.retained_entries, 0);

    let document = PackURI::new("/word/document.xml").unwrap();
    let data = package.part(&document).unwrap().data().unwrap();
    assert_eq!(data.as_bytes(), b"<document/>");

    let loaded = package.cache_diagnostics();
    assert_eq!(loaded.cold_loads, 1);
    assert_eq!(loaded.successful_loads, 1);
    assert_eq!(loaded.retained_entries, 1);
}

#[test]
fn part_view_declared_size_is_metadata_only() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<document/>",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let before = package.cache_diagnostics();
    let document = PackURI::new("/word/document.xml").unwrap();
    let view = package.part(&document).unwrap();

    assert_eq!(
        view.declared_uncompressed_size().unwrap(),
        b"<document/>".len() as u64
    );
    let after = package.cache_diagnostics();
    assert_eq!(after.cold_loads, before.cold_loads);
    assert_eq!(after.retained_entries, before.retained_entries);
}

#[test]
fn part_view_declared_size_checks_source_and_cancellation() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<document/>",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let document = PackURI::new("/word/document.xml").unwrap();
    let view = package.part(&document).unwrap();
    source.changed();
    assert!(matches!(
        view.declared_uncompressed_size(),
        Err(OpcError::SourceChanged { .. })
    ));

    let (_budget, cancellation_source, context) = managed_context_with_cancellation(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<document/>",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let view = package.part(&document).unwrap();
    cancellation_source.cancel();
    assert!(matches!(
        view.declared_uncompressed_size(),
        Err(OpcError::Cancelled)
    ));
}

#[test]
fn topology_add_part_inserts_lossless_content_type_override() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        PackURI::new("/custom/new.bin").unwrap(),
        "application/octet-stream",
        b"new payload".to_vec(),
    )
    .unwrap();
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();

    let archive = soapberry_zip::office::ArchiveReader::new(
        &output,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
    let content_types = archive.read("[Content_Types].xml").unwrap();
    let source_content_types = archive_bytes(root_relationships(), b"<before/>", false);
    assert!(content_types
            .windows(b"<Override PartName=\"/custom/new.bin\" ContentType=\"application/octet-stream\"/>".len())
            .any(|window| window
                == b"<Override PartName=\"/custom/new.bin\" ContentType=\"application/octet-stream\"/>"));
    assert_ne!(output, source_content_types);
    assert_eq!(archive.read("custom/new.bin").unwrap(), b"new payload");
    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(
        reopened
            .try_iter_parts()
            .find_map(|part| {
                let part = part.unwrap();
                (part.partname().as_str() == "/custom/new.bin").then_some(part)
            })
            .unwrap()
            .blob(),
        b"new payload"
    );
    assert_eq!(archive.read("custom/orphan.xml").unwrap(), b"<orphan/>");
}

#[test]
fn prepared_topology_exposes_complete_effective_graph_and_publishes_same_overlays() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let package = SourceBackedPackage::from_vec(
        source_bytes.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let document = PackURI::new("/word/document.xml").unwrap();
    let added = PackURI::new("/custom/new.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(document.clone(), b"<after/>".to_vec())
        .unwrap();
    plan.try_add_part(added.clone(), "application/xml", b"<new/>".to_vec())
        .unwrap();
    plan.try_add_internal_relationship(
        document.clone(),
        "rId2",
        "urn:test:prepared",
        added.clone(),
    )
    .unwrap();

    let prepared = package.prepare_topology(plan).unwrap();
    prepared
        .with_candidate(|candidate| {
            assert_eq!(candidate.parts().len(), 3);
            let document_view = candidate.part(&document)?;
            assert_eq!(document_view.data()?.as_bytes(), b"<after/>");
            assert_eq!(
                document_view.relationships().get("rId2").unwrap().reltype(),
                "urn:test:prepared"
            );
            assert_eq!(
                document_view
                    .relationships()
                    .get("rId2")
                    .unwrap()
                    .target_partname()?,
                added
            );
            assert_eq!(candidate.part(&added)?.data()?.as_bytes(), b"<new/>");
            assert_eq!(candidate.content_type(&added)?.as_str(), "application/xml");
            assert!(candidate.has_physical_member("custom/new.xml")?);
            assert!(candidate.has_physical_member("word/document.xml")?);
            Ok(())
        })
        .unwrap();

    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    let reopened = SourceBackedPackage::from_vec(
        output,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    assert_eq!(
        reopened.part(&document).unwrap().data().unwrap().as_bytes(),
        b"<after/>"
    );
    assert_eq!(
        reopened.part(&added).unwrap().data().unwrap().as_bytes(),
        b"<new/>"
    );
    assert_eq!(
        reopened
            .part(&document)
            .unwrap()
            .rels()
            .get("rId2")
            .unwrap()
            .target_partname()
            .unwrap(),
        added
    );
}

#[test]
fn prepared_topology_stale_fence_rejects_readback_and_publication_before_output() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(
        PackURI::new("/word/document.xml").unwrap(),
        b"<after/>".to_vec(),
    )
    .unwrap();
    let prepared = package.prepare_topology(plan).unwrap();
    source.changed();
    assert!(matches!(
        prepared.with_candidate(|_| Ok(())),
        Err(OpcError::SourceChanged { .. })
    ));
    let mut output = Vec::new();
    assert!(matches!(
        prepared.publish_to_stream(&mut output),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn prepared_topology_removal_updates_effective_membership_and_content_types() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let package = SourceBackedPackage::from_vec(
        source_bytes,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let removed = PackURI::new("/custom/orphan.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_part(removed.clone()).unwrap();
    let prepared = package.prepare_topology(plan).unwrap();
    prepared
        .with_candidate(|candidate| {
            assert!(candidate.part(&removed).is_err());
            assert!(!candidate.has_physical_member("custom/orphan.xml")?);
            assert_eq!(candidate.parts().len(), 1);
            assert_eq!(
                candidate
                    .content_type(&PackURI::new("/word/document.xml").unwrap())?
                    .as_str(),
                "application/xml"
            );
            Ok(())
        })
        .unwrap();
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    let reopened = SourceBackedPackage::from_vec(
        output,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    assert!(reopened.part(&removed).is_err());
}

#[test]
fn prepared_topology_empty_plan_keeps_exact_source_fast_path_without_candidate() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let package = SourceBackedPackage::from_vec(
        source_bytes.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let prepared = package.prepare_topology(SourceTopologyPlan::new()).unwrap();
    assert!(matches!(
        prepared.with_candidate(|_| Ok(())),
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
    let mut output = Vec::new();
    prepared.publish_to_stream(&mut output).unwrap();
    assert_eq!(output, source_bytes);
}

#[test]
fn topology_add_part_shared_reuses_payload_and_matches_add_validation() {
    let partname = PackURI::new("/custom/shared.bin").unwrap();
    let payload = Arc::new(b"shared payload".to_vec());
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part_shared(
        partname.clone(),
        "application/octet-stream",
        Arc::clone(&payload),
    )
    .unwrap();
    let TopologyPartPayload::Decoded(retained) = &plan.additions[0].payload else {
        panic!("shared decoded addition must retain its decoded payload");
    };
    assert!(Arc::ptr_eq(&payload, retained));

    assert!(matches!(
        plan.try_add_part_shared(
            partname,
            "application/octet-stream",
            Arc::new(b"duplicate".to_vec()),
        ),
        Err(OpcError::DuplicatePartName(_))
    ));

    let mut root_plan = SourceTopologyPlan::new();
    assert!(matches!(
        root_plan.try_add_part_shared(
            PackURI::new("/").unwrap(),
            "application/octet-stream",
            Arc::new(Vec::new()),
        ),
        Err(OpcError::InvalidPackUri(_))
    ));

    let constructed = Arc::new(AtomicBool::new(false));
    let constructed_marker = Arc::clone(&constructed);
    let mut lazy_plan = SourceTopologyPlan::new();
    assert!(matches!(
        lazy_plan.try_add_part_with_payload(
            PackURI::new("/").unwrap(),
            "application/octet-stream",
            || {
                constructed_marker.store(true, Ordering::SeqCst);
                Arc::new(Vec::new())
            },
        ),
        Err(OpcError::InvalidPackUri(_))
    ));
    assert!(!constructed.load(Ordering::SeqCst));

    let mut content_type_plan = SourceTopologyPlan::new();
    assert!(matches!(
        content_type_plan.try_add_part_shared(
            PackURI::new("/custom/invalid.bin").unwrap(),
            "invalid content type",
            Arc::new(Vec::new()),
        ),
        Err(OpcError::InvalidContentType { .. })
    ));

    let mut bounded_plan = SourceTopologyPlan::new();
    for index in 0..MAX_SOURCE_TOPOLOGY_PARTS {
        bounded_plan
            .try_add_part_shared(
                PackURI::new(format!("/custom/bounded{index}.bin")).unwrap(),
                "application/octet-stream",
                Arc::new(Vec::new()),
            )
            .unwrap();
    }
    assert!(matches!(
        bounded_plan.try_add_part_shared(
            PackURI::new("/custom/over-bound.bin").unwrap(),
            "application/octet-stream",
            Arc::new(Vec::new()),
        ),
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));

    let mut retained_payload = Arc::new(b"copy-on-write".to_vec());
    let mut retained_plan = SourceTopologyPlan::new();
    retained_plan
        .try_add_part_shared(
            PackURI::new("/custom/retained.bin").unwrap(),
            "application/octet-stream",
            Arc::clone(&retained_payload),
        )
        .unwrap();
    assert_eq!(Arc::strong_count(&retained_payload), 2);
    drop(retained_plan);
    assert_eq!(Arc::strong_count(&retained_payload), 1);
    Arc::make_mut(&mut retained_payload)[0] = b'C';
    assert_eq!(retained_payload.as_slice(), b"Copy-on-write");
}

#[test]
fn topology_replace_part_shared_preserves_arc_identity_and_cow_isolation() {
    let partname = PackURI::new("/word/document.xml").unwrap();
    let mut payload = Arc::new(b"<shared/>".to_vec());
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part_shared(partname, Arc::clone(&payload))
        .unwrap();

    let TopologyReplacement {
        replacement,
        source_xml,
        ..
    } = &plan.replacements[0];
    assert!(Arc::ptr_eq(&payload, replacement));
    assert!(source_xml.is_none());
    assert_eq!(Arc::strong_count(&payload), 2);

    Arc::make_mut(&mut payload)[0] = b'X';
    assert_eq!(payload.as_slice(), b"Xshared/>");
    assert_eq!(replacement.as_slice(), b"<shared/>");

    drop(plan);
    assert_eq!(Arc::strong_count(&payload), 1);
}

#[test]
fn topology_replace_part_shared_publishes_the_same_bytes_as_owned_replacement() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let partname = PackURI::new("/word/document.xml").unwrap();

    let mut owned_plan = SourceTopologyPlan::new();
    owned_plan
        .try_replace_part(partname.clone(), b"<after/>".to_vec())
        .unwrap();
    let mut owned_output = Vec::new();
    SourceBackedPackage::from_vec(
        source_bytes.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut owned_output, owned_plan)
    .unwrap();

    let mut shared_plan = SourceTopologyPlan::new();
    shared_plan
        .try_replace_part_shared(partname.clone(), Arc::new(b"<after/>".to_vec()))
        .unwrap();
    let mut shared_output = Vec::new();
    SourceBackedPackage::from_vec(
        source_bytes,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut shared_output, shared_plan)
    .unwrap();

    assert_eq!(shared_output, owned_output);
    assert_eq!(
        soapberry_zip::office::ArchiveReader::new(
            &shared_output,
            soapberry_zip::office::ArchiveLimits::default()
        )
        .unwrap()
        .read(partname.membername())
        .unwrap(),
        b"<after/>"
    );
}

#[test]
fn topology_replace_part_shared_with_spans_publishes_multiwindow_xml() {
    let original = "<r><a>0</a><b>0</b><c>unchanged</c></r>";
    let replacement = "<r><a>1</a><b>12345</b><c>unchanged</c></r>";
    let original_a = original.find("<a>0</a>").unwrap();
    let replacement_a = replacement.find("<a>1</a>").unwrap();
    let original_b = original.find("<b>0</b>").unwrap();
    let replacement_b = replacement.find("<b>12345</b>").unwrap();
    let spans: Arc<[ReplacementSpan]> = Arc::from(vec![
        ReplacementSpan {
            original: original_a..original_a + "<a>0</a>".len(),
            replacement: replacement_a..replacement_a + "<a>1</a>".len(),
        },
        ReplacementSpan {
            original: original_b..original_b + "<b>0</b>".len(),
            replacement: replacement_b..replacement_b + "<b>12345</b>".len(),
        },
    ]);
    let proof = audit_source_part_replacement_xml(
        "/word/document.xml",
        original.as_bytes(),
        replacement.as_bytes(),
        Some(spans.as_ref()),
    );
    assert!(matches!(
        proof,
        Ok(xml_minifier::audit::ReplacementProof::Windows { count: 2, .. })
    ));

    let partname = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part_shared_with_spans(
        partname.clone(),
        Arc::new(replacement.as_bytes().to_vec()),
        Arc::clone(&spans),
    )
    .unwrap();
    assert_eq!(
        plan.replacements[0].replacement_spans.as_deref(),
        Some(spans.as_ref())
    );

    let mut output = Vec::new();
    SourceBackedPackage::from_vec(
        archive_bytes(root_relationships(), original.as_bytes(), false),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, plan)
    .unwrap();
    assert_eq!(
        soapberry_zip::office::ArchiveReader::new(
            &output,
            soapberry_zip::office::ArchiveLimits::default()
        )
        .unwrap()
        .read(partname.membername())
        .unwrap(),
        replacement.as_bytes()
    );
}

#[test]
fn topology_replace_part_shared_with_spans_cannot_bypass_malformed_replacement() {
    let original = "<r><a>0</a><b>0</b><c>0</c></r>";
    let replacement = "<r><a>1</broken><b>2</b><c>0</c></r>";
    let original_a = original.find("<a>0</a>").unwrap();
    let replacement_a = replacement.find("<a>1</broken>").unwrap();
    let original_b = original.find("<b>0</b>").unwrap();
    let spans: Arc<[ReplacementSpan]> = Arc::from(vec![
        ReplacementSpan {
            original: original_a..original_a + "<a>0</a>".len(),
            replacement: replacement_a..replacement_a + "<a>1</broken>".len(),
        },
        // This out-of-bounds hint forces the auditor's complete fallback;
        // it cannot turn the malformed first window into a proof.
        ReplacementSpan {
            original: original_b..original_b + "<b>0</b>".len(),
            replacement: replacement.len()..replacement.len() + 1,
        },
    ]);
    assert!(matches!(
        audit_source_part_replacement_xml(
            "/word/document.xml",
            original.as_bytes(),
            replacement.as_bytes(),
            Some(spans.as_ref()),
        ),
        Err(OpcError::XmlPublication { .. })
    ));

    let partname = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part_shared_with_spans(
        partname,
        Arc::new(replacement.as_bytes().to_vec()),
        spans,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        SourceBackedPackage::from_vec(
            archive_bytes(root_relationships(), original.as_bytes(), false,),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            None
        )
        .unwrap()
        .write_topology_to_stream(&mut output, plan),
        Err(OpcError::XmlPublication { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn a_remembered_original_proves_insertions_from_its_boundaries() {
    // The remembered original keeps its indexed proof, and an unproven
    // replacement gets its complete audit's verdict and error from a scan
    // of the inserted run alone.
    let partname = PackURI::new("/word/document.xml").unwrap();
    let mut document = String::from("<w:document xmlns:w=\"urn:w\"><w:body>");
    for index in 0..40 {
        document.push_str(&format!(
            "<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Paragraph {index}</w:t></w:r></w:p>"
        ));
    }
    document.push_str("</w:body></w:document>");
    let package = SourceBackedPackage::from_vec(
        archive_bytes(root_relationships(), document.as_bytes(), false),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let target = package.part_index(&partname).unwrap();
    let original = package.read_part(target).unwrap();
    let version = package.source.version();
    let at = document
        .find("<w:p><w:r><w:rPr><w:i/></w:rPr><w:t>Paragraph 20<")
        .unwrap();
    let insert = |fragment: &str| {
        Arc::new(format!("{}{fragment}{}", &document[..at], &document[at..]).into_bytes())
    };
    let inserted = insert("<w:p><w:r><w:t>Inserted</w:t></w:r></w:p>");
    // The first pair audits the original once and keeps its boundaries.
    let proof = package
        .audit_physical_replacement_xml(target, &original, &inserted, None, None, None)
        .unwrap()
        .unwrap();
    assert!(proof.covers(&inserted, xml_minifier::audit::Limits::default()));
    let base = package
        .remembered_source_xml_audit(target, version)
        .unwrap()
        .expect("an unmanaged original keeps its indexed proof");
    assert!(format!("{base:?}").contains("boundaries: Some("));
    assert_eq!(base.bytes().as_slice(), document.as_bytes());
    for replacement in [
        inserted,
        insert("<w:tbl/>"),
        insert("<w:p>"),
        insert("</w:body><w:body>"),
        insert("<x:p/>"),
        insert("<w:p><w:t>&bogus;</w:t></w:p>"),
    ] {
        let actual = package.audit_physical_replacement_xml(
            target,
            &original,
            &replacement,
            None,
            None,
            None,
        );
        match xml_minifier::audit::verify_source(
            &replacement,
            xml_minifier::audit::Limits::default(),
        ) {
            Ok(_) => {
                let proof = actual.unwrap().unwrap();
                assert!(proof.covers(&replacement, xml_minifier::audit::Limits::default()));
            },
            Err(expected) => match actual {
                Err(OpcError::XmlPublication { source, .. }) => assert_eq!(source, expected),
                other => panic!("{other:?} instead of {expected:?}"),
            },
        }
    }
    // At most four remembered originals keep their proofs; every verdict
    // stays remembered.
    for other in 1_000..1_006 {
        package.remember_source_xml_audit(other, version, Some(base.clone()));
    }
    let audited = package.cache.audited_source_xml.lock().unwrap();
    assert_eq!(audited.len(), 7);
    assert_eq!(
        audited
            .iter()
            .filter(|audited| audited.proof.is_some())
            .count(),
        RETAINED_SOURCE_XML_PROOFS
    );
    assert!(audited[0].proof.is_none());
}

#[test]
fn physical_replacement_audit_remembers_only_passing_originals() {
    // The remembered original verdict and the plan's proof.
    let partname = PackURI::new("/word/document.xml").unwrap();
    let package = SourceBackedPackage::from_vec(
        archive_bytes(root_relationships(), b"<before/>", false),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let target = package.part_index(&partname).unwrap();
    let original = package.read_part(target).unwrap();
    let version = package.source.version();
    // The original passes here and the replacement fails: nothing is
    // remembered, so the next pair is audited completely again.
    assert!(
        package
            .audit_physical_replacement_xml(
                target,
                &original,
                &Arc::new(b"<a>".to_vec()),
                None,
                None,
                None,
            )
            .is_err()
    );
    assert!(!package.original_source_xml_audited(target, version));
    let replacement = Arc::new(b"<after/>".to_vec());
    let proof = package
        .audit_physical_replacement_xml(target, &original, &replacement, None, None, None)
        .unwrap()
        .expect("a complete pair audit proves its replacement");
    assert!(proof.covers(&replacement, xml_minifier::audit::Limits::default()));
    assert!(package.original_source_xml_audited(target, version));
    assert!(package.clone().original_source_xml_audited(target, version));
    // A remembered original still audits an unproven replacement.
    assert!(matches!(
        package.audit_physical_replacement_xml(
            target,
            &original,
            &Arc::new(b"<a>".to_vec()),
            None,
            None,
            None,
        ),
        Err(OpcError::XmlPublication { .. })
    ));
    // The plan entry keeps the proof and publishes its allocation.
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part_verified(partname.clone(), proof.clone())
        .unwrap();
    assert!(plan.replacements[0].proof.is_some());
    assert!(Arc::ptr_eq(
        &plan.replacements[0].replacement,
        proof.bytes()
    ));
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    assert_eq!(
        soapberry_zip::office::ArchiveReader::new(
            &output,
            soapberry_zip::office::ArchiveLimits::default()
        )
        .unwrap()
        .read(partname.membername())
        .unwrap(),
        b"<after/>"
    );
}

#[test]
fn topology_replace_part_shared_with_spans_drops_overbound_hints() {
    let mut plan = SourceTopologyPlan::new();
    let spans: Arc<[ReplacementSpan]> = Arc::from(
        (0..=MAX_SOURCE_REPLACEMENT_SPANS)
            .map(|index| ReplacementSpan {
                original: index..index + 1,
                replacement: index..index + 1,
            })
            .collect::<Vec<_>>(),
    );
    let partname = PackURI::new("/word/document.xml").unwrap();
    plan.try_replace_part_shared_with_spans(
        partname.clone(),
        Arc::new(b"<replacement/>".to_vec()),
        spans,
    )
    .unwrap();
    assert!(plan.replacements[0].replacement_spans.is_none());

    let mut output = Vec::new();
    SourceBackedPackage::from_vec(
        archive_bytes(root_relationships(), b"<before/>", false),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, plan)
    .unwrap();
    assert_eq!(
        soapberry_zip::office::ArchiveReader::new(
            &output,
            soapberry_zip::office::ArchiveLimits::default()
        )
        .unwrap()
        .read(partname.membername())
        .unwrap(),
        b"<replacement/>"
    );
}

#[test]
fn topology_replace_part_shared_preserves_refusal_order_and_plan_on_failure() {
    let root = PackURI::new("/").unwrap();
    let mut root_plan = SourceTopologyPlan::new();
    let root_payload = Arc::new(b"root payload".to_vec());
    assert!(matches!(
        root_plan.try_replace_part_shared(root, Arc::clone(&root_payload)),
        Err(OpcError::InvalidPackUri(_))
    ));
    assert!(root_plan.replacements.is_empty());
    assert_eq!(Arc::strong_count(&root_payload), 1);

    let mut lazy_root_plan = SourceTopologyPlan::new();
    let constructed = Arc::new(AtomicBool::new(false));
    let constructed_marker = Arc::clone(&constructed);
    assert!(matches!(
        lazy_root_plan.try_replace_part_with_payload(PackURI::new("/").unwrap(), || {
            constructed_marker.store(true, Ordering::SeqCst);
            Arc::new(Vec::new())
        },),
        Err(OpcError::InvalidPackUri(_))
    ));
    assert!(!constructed.load(Ordering::SeqCst));

    let duplicate_name = PackURI::new("/word/document.xml").unwrap();
    let mut duplicate_plan = SourceTopologyPlan::new();
    duplicate_plan
        .try_replace_part_shared(duplicate_name.clone(), Arc::new(b"<first/>".to_vec()))
        .unwrap();
    let duplicate_payload = Arc::new(b"<second/>".to_vec());
    let duplicate_equivalent = PackURI::new("/WORD/DOCUMENT.XML").unwrap();
    assert!(matches!(
        duplicate_plan
            .try_replace_part_shared(duplicate_equivalent, Arc::clone(&duplicate_payload),),
        Err(OpcError::DuplicatePartName(_))
    ));
    assert_eq!(duplicate_plan.replacements.len(), 1);
    assert_eq!(
        duplicate_plan.replacements[0].replacement.as_slice(),
        b"<first/>"
    );
    assert_eq!(Arc::strong_count(&duplicate_payload), 1);

    let cross_duplicate = PackURI::new("/custom/cross-duplicate.bin").unwrap();
    let mut addition_plan = SourceTopologyPlan::new();
    addition_plan
        .try_add_part(
            cross_duplicate.clone(),
            "application/octet-stream",
            b"existing".to_vec(),
        )
        .unwrap();
    assert!(matches!(
        addition_plan.try_replace_part_shared(cross_duplicate, Arc::new(b"replacement".to_vec()),),
        Err(OpcError::DuplicatePartName(_))
    ));
    assert_eq!(addition_plan.additions.len(), 1);
    assert!(addition_plan.replacements.is_empty());

    let removed_name = PackURI::new("/custom/removed.bin").unwrap();
    let mut removal_plan = SourceTopologyPlan::new();
    removal_plan.try_remove_part(removed_name.clone()).unwrap();
    assert!(matches!(
        removal_plan.try_replace_part_shared(removed_name, Arc::new(b"replacement".to_vec()),),
        Err(OpcError::DuplicatePartName(_))
    ));
    assert_eq!(removal_plan.removals.len(), 1);
    assert!(removal_plan.replacements.is_empty());

    let mut bounded_plan = SourceTopologyPlan::new();
    for index in 0..MAX_SOURCE_TOPOLOGY_PARTS {
        bounded_plan
            .try_replace_part_shared(
                PackURI::new(format!("/custom/bounded{index}.bin")).unwrap(),
                Arc::new(Vec::new()),
            )
            .unwrap();
    }
    let over_bound_payload = Arc::new(b"over bound".to_vec());
    assert!(matches!(
        bounded_plan.try_replace_part_shared(
            PackURI::new("/custom/over-bound.bin").unwrap(),
            Arc::clone(&over_bound_payload),
        ),
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
    assert_eq!(bounded_plan.replacements.len(), MAX_SOURCE_TOPOLOGY_PARTS);
    assert_eq!(Arc::strong_count(&over_bound_payload), 1);

    let mut full_root_plan = SourceTopologyPlan::new();
    for index in 0..MAX_SOURCE_TOPOLOGY_PARTS {
        full_root_plan
            .try_add_part(
                PackURI::new(format!("/custom/root-bound{index}.bin")).unwrap(),
                "application/octet-stream",
                Vec::new(),
            )
            .unwrap();
    }
    assert!(matches!(
        full_root_plan.try_replace_part_shared(PackURI::new("/").unwrap(), Arc::new(Vec::new()),),
        Err(OpcError::InvalidPackUri(_))
    ));
    assert_eq!(full_root_plan.additions.len(), MAX_SOURCE_TOPOLOGY_PARTS);
    assert!(full_root_plan.replacements.is_empty());
}

#[test]
fn topology_add_at_zip32_entry_count_boundary_promotes_source_backed_source() {
    const SOURCE_MEMBER_COUNT: usize = u16::MAX as usize - 1;

    let source_bytes = zip32_entry_count_boundary_source();
    let source_archive = soapberry_zip::ZipArchive::from_slice(&source_bytes).unwrap();
    assert!(!source_archive.is_zip64());
    assert_eq!(source_archive.entries_hint(), u64::from(u16::MAX - 1));
    assert_eq!(source_archive.entries().count(), SOURCE_MEMBER_COUNT);
    let source_raw = raw_records(&source_bytes);
    let source = Arc::new(CountingSource::new(source_bytes.clone()));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_eq!(package.physical_member_names().len(), SOURCE_MEMBER_COUNT);
    assert_eq!(package.iter_parts().count(), SOURCE_MEMBER_COUNT - 2);

    let added = PackURI::new("/custom/new.bin").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        added.clone(),
        content_type::XLSB_BIN,
        b"new payload".to_vec(),
    )
    .unwrap();
    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, plan)
        .expect("publish boundary addition");
    assert_eq!(source.bytes.as_slice(), source_bytes.as_slice());

    let output_archive = soapberry_zip::ZipArchive::from_slice(&output).unwrap();
    assert!(output_archive.is_zip64());
    assert_eq!(output_archive.entries_hint(), u64::from(u16::MAX));
    assert_eq!(output_archive.entries().count(), usize::from(u16::MAX));

    let output_raw = raw_records(&output);
    assert_eq!(output_raw.len(), source_raw.len() + 1);
    assert!(output_raw.contains_key(added.membername().as_bytes()));
    for (name, source_record) in &source_raw {
        let output_record = output_raw.get(name).expect("source member retained");
        assert_eq!(output_record.local, source_record.local, "{name:?}");
        assert_eq!(
            central_without_local_offset(&output_record.central),
            central_without_local_offset(&source_record.central),
            "{name:?}"
        );
    }

    let archive_reader = soapberry_zip::office::ArchiveReader::new(
        &output,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
    assert_eq!(
        archive_reader.read(added.membername()).unwrap(),
        b"new payload"
    );
    let reopened = SourceBackedPackage::from_vec(
        output,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .expect("reopen boundary output");
    assert_eq!(reopened.iter_parts().count(), SOURCE_MEMBER_COUNT - 1);
    assert_eq!(
        reopened.part(&added).unwrap().data().unwrap().as_bytes(),
        b"new payload"
    );
}

#[test]
fn topology_changed_replacement_reads_source_payload_once() {
    const BEFORE: &[u8] = b"<before/>";
    let source_bytes = archive_bytes(root_relationships(), BEFORE, false);
    let payload_start = source_bytes
        .windows(BEFORE.len())
        .position(|window| window == BEFORE)
        .expect("the source payload must be present") as u64;
    let payload_end = payload_start + BEFORE.len() as u64;
    let source = Arc::new(CountingSource::new(source_bytes));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::new(1, 1).unwrap(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let ranges_before = source.read_ranges().len();

    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_part(
        PackURI::new("/word/document.xml").unwrap(),
        b"<after/>".to_vec(),
    )
    .unwrap();
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();

    let payload_reads = source.read_ranges()[ranges_before..]
        .iter()
        .filter(|&&(offset, length)| {
            let end = offset.saturating_add(length as u64);
            offset < payload_end && end > payload_start
        })
        .count();
    assert_eq!(payload_reads, 1);
    assert_eq!(
        soapberry_zip::office::ArchiveReader::new(
            &output,
            soapberry_zip::office::ArchiveLimits::default()
        )
        .unwrap()
        .read("word/document.xml")
        .unwrap(),
        b"<after/>"
    );
}

#[test]
fn topology_source_xml_addition_rechecks_source_version_after_initial_validation() {
    const SOURCE_XML: &[u8] = b"<source><nested/></source>";
    let source = Arc::new(ChangeOnHitVersionSource::new(archive_bytes(
        root_relationships(),
        SOURCE_XML,
        false,
    )));
    let source_package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let document = PackURI::new("/word/document.xml").unwrap();
    let source_xml = source_package
        .part(&document)
        .unwrap()
        .source_xml()
        .unwrap();
    // A source XML addition receives two source checks in its initial
    // publication proof. The next check is the transfer-boundary fence.
    source.arm_after_versions(2);

    let destination = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<destination/>",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_source_xml_part(PackURI::new("/custom/copied.xml").unwrap(), source_xml)
        .unwrap();
    let mut output = Vec::new();
    let error = destination
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(error, OpcError::SourceChanged { .. }));
    assert!(output.is_empty());
    drop(source_package);
}

#[test]
fn topology_source_xml_addition_rechecks_cancellation_after_initial_validation() {
    const SOURCE_XML: &[u8] = b"<source><nested/></source>";
    let (budget, cancellation_source, context) = managed_context_with_cancellation(64 * 1024);
    let source = Arc::new(CancelOnHitVersionSource::new(
        archive_bytes(root_relationships(), SOURCE_XML, false),
        cancellation_source,
    ));
    let source_package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let document = PackURI::new("/word/document.xml").unwrap();
    let source_xml = source_package
        .part(&document)
        .unwrap()
        .source_xml()
        .unwrap();
    // The source's cancellation is injected by the first version check
    // after the two checks that complete the initial XML proof.
    source.arm_after_versions(2);

    let destination = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<destination/>",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_source_xml_part(PackURI::new("/custom/copied.xml").unwrap(), source_xml)
        .unwrap();
    let mut output = Vec::new();
    let error = destination
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(error, OpcError::Cancelled));
    assert!(output.is_empty());
    drop(source_package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn topology_source_xml_addition_charges_one_xml_validation_pass() {
    const SOURCE_XML: &[u8] = b"<source><nested/></source>";
    let (budget, _cancellation_source, context) = managed_context_with_cancellation(64 * 1024);
    let source_package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            SOURCE_XML,
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let document = PackURI::new("/word/document.xml").unwrap();
    let source_xml = source_package
        .part(&document)
        .unwrap()
        .source_xml()
        .unwrap();
    let work_before = budget.used(Resource::Work);

    let destination = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<destination/>",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_source_xml_part(PackURI::new("/custom/copied.xml").unwrap(), source_xml)
        .unwrap();
    let mut output = Vec::new();
    destination
        .write_topology_to_stream(&mut output, plan)
        .unwrap();
    assert_eq!(
        budget.used(Resource::Work) - work_before,
        SOURCE_XML.len() as u64
    );

    drop(source_package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn topology_zip64_add_remove_preserves_untouched_records_and_tail() {
    let extension = [0x11, 0x22, 0x33, 0x44];
    let source_bytes = add_zip64_tail_extension(
        promote_to_zip64(archive_bytes(root_relationships(), b"<before/>", false)),
        &extension,
    );
    let source_raw = raw_records(&source_bytes);
    let source_archive = soapberry_zip::ZipArchive::from_slice(&source_bytes).unwrap();
    assert!(source_archive.is_zip64());
    assert_eq!(zip64_tail_extensible_data(&source_bytes), extension);

    let document = PackURI::new("/word/document.xml").unwrap();
    let orphan = PackURI::new("/custom/orphan.xml").unwrap();
    let added = PackURI::new("/custom/new.bin").unwrap();
    let publish = || {
        let source = Arc::new(CountingSource::new(source_bytes.clone()));
        let package = SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_remove_part(orphan.clone()).unwrap();
        plan.try_add_part(
            added.clone(),
            "application/octet-stream",
            b"new payload".to_vec(),
        )
        .unwrap();
        let mut output = Vec::new();
        package.write_topology_to_stream(&mut output, plan).unwrap();
        assert_eq!(source.bytes.as_slice(), source_bytes.as_slice());
        output
    };

    let output = publish();
    assert_eq!(output, publish());
    let output_archive = soapberry_zip::ZipArchive::from_slice(&output).unwrap();
    assert!(output_archive.is_zip64());
    assert_eq!(zip64_tail_extensible_data(&output), extension);

    let output_raw = raw_records(&output);
    assert_eq!(output_raw.len(), source_raw.len());
    assert!(output_raw.contains_key(b"custom/new.bin".as_slice()));
    assert!(!output_raw.contains_key(b"custom/orphan.xml".as_slice()));
    for (name, source_record) in &source_raw {
        if name.as_slice() == b"custom/orphan.xml" || name.as_slice() == b"[Content_Types].xml" {
            continue;
        }
        let output_record = output_raw.get(name).expect("untouched member retained");
        assert_eq!(output_record.local, source_record.local, "{name:?}");
        assert_eq!(
            central_without_local_offset(&output_record.central),
            central_without_local_offset(&source_record.central),
            "{name:?}"
        );
    }

    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default())
        .expect("reopen ZIP64 topology output");
    assert_eq!(reopened.get_part(&document).unwrap().blob(), b"<before/>");
    assert_eq!(reopened.get_part(&added).unwrap().blob(), b"new payload");
    assert!(matches!(
        reopened.get_part(&orphan),
        Err(OpcError::PartNotFound(_))
    ));
}

#[test]
fn content_type_override_insertion_uses_the_parsed_root_close_span() {
    let source = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><!-- text that looks like </Types> --><?keep?></Types>"#;
    let partname = PackURI::new("/custom/new.bin").unwrap();
    let additions = [(partname, "application/octet-stream".to_string())];
    let plan = content_types_plan::ContentTypesPlan::plan(
        source,
        &additions,
        &[],
        SourceLimits::default(),
        None,
        None,
    )
    .unwrap();
    let (output, _memory_reservation) = plan
        .materialize(SourceLimits::default(), None, None)
        .unwrap();
    assert_eq!(
            output,
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><!-- text that looks like </Types> --><?keep?><Override PartName="/custom/new.bin" ContentType="application/octet-stream"/></Types>"#
        );
    ContentTypeMap::from_xml(&output, SourceLimits::default()).unwrap();
}

#[test]
fn topology_adds_package_relationship_to_new_part() {
    let root = br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer.write_stored("_rels/.rels", root).unwrap();
    writer
        .write_stored("word/document.xml", b"<before/>")
        .unwrap();
    let source = writer.finish_to_bytes().unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    let target = PackURI::new("/custom/new.bin").unwrap();
    plan.try_add_part(
        target.clone(),
        "application/octet-stream",
        b"new payload".to_vec(),
    )
    .unwrap();
    plan.try_add_internal_relationship(
        PackURI::new("/").unwrap(),
        "rId2",
        "urn:litchi:test",
        target,
    )
    .unwrap();
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_eq!(
        reopened.rels().get("rId2").unwrap().target_ref(),
        "custom/new.bin"
    );
    assert_eq!(
        reopened
            .part(&PackURI::new("/custom/new.bin").unwrap())
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"new payload"
    );
}

fn large_archive_bytes(root_relationships: &[u8], document: &[u8], junk_bytes: usize) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships)
        .unwrap();
    writer.write_stored("word/document.xml", document).unwrap();
    writer
        .write_stored("custom/orphan.xml", b"<orphan/>")
        .unwrap();
    let junk = vec![0xA5; junk_bytes];
    writer.write_stored("scratch.bin", &junk).unwrap();
    writer.finish_to_bytes().unwrap()
}

fn archive_with_part_names(root_target: &str, part_names: &[&str]) -> Vec<u8> {
    let root_relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{root_target}"/></Relationships>"#
    );
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    for (index, part_name) in part_names.iter().enumerate() {
        writer
            .write_stored(
                part_name,
                if index == 0 {
                    &b"document"[..]
                } else {
                    &b"other"[..]
                },
            )
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn archive_with_ordered_xml_parts(parts: &[(&str, &[u8])]) -> Vec<u8> {
    assert!(!parts.is_empty());
    let root_relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{}"/></Relationships>"#,
        parts[0].0
    );
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    for (name, payload) in parts {
        writer.write_stored(name, payload).unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn archive_with_descending_xml_parts(count: usize) -> Vec<u8> {
    assert!(count > 0);
    let first_name = format!("part-{index:04}.xml", index = count - 1);
    let root_relationships = format!(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="{first_name}"/></Relationships>"#
    );
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships.as_bytes())
        .unwrap();
    for index in (0..count).rev() {
        let name = format!("part-{index:04}.xml");
        writer.write_stored(&name, b"payload").unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn archive_with_document_relationships(document_relationships: &[u8]) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships())
        .unwrap();
    writer
        .write_stored("word/document.xml", b"<before/>")
        .unwrap();
    writer
        .write_stored("word/_rels/document.xml.rels", document_relationships)
        .unwrap();
    writer
        .write_stored("custom/orphan.xml", b"<orphan/>")
        .unwrap();
    writer.write_stored("scratch.bin", b"untouched").unwrap();
    writer.finish_to_bytes().unwrap()
}

fn document_relationships() -> &'static [u8] {
    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://remove.invalid/" TargetMode="External"/><Relationship Id="rInternal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../custom/orphan.xml"/></Relationships>"#
}

fn topology_relationship_archive(document_relationships: &[u8]) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships())
        .unwrap();
    writer
        .write_stored("word/document.xml", b"<before/>")
        .unwrap();
    writer
        .write_stored("word/target.xml", b"<target/>")
        .unwrap();
    writer
        .write_stored("word/_rels/document.xml.rels", document_relationships)
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

fn canonical_document_relationships(entries: &[(&str, &str, &str, TargetMode)]) -> Vec<u8> {
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut relationships = Relationships::for_source(&owner);
    for (r_id, reltype, target, mode) in entries {
        relationships
            .try_add_relationship(
                (*reltype).to_string(),
                (*target).to_string(),
                (*r_id).to_string(),
                *mode,
            )
            .unwrap();
    }
    relationships.try_to_xml_bytes().unwrap()
}

#[test]
fn canonical_relationship_length_matches_fallible_serializer() {
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut relationships = Relationships::for_source(&owner);
    relationships
        .try_add_relationship(
            "urn:test:a&b".to_string(),
            "https://example.invalid/?q=a&b=2".to_string(),
            "rId2".to_string(),
            TargetMode::External,
        )
        .unwrap();
    relationships
        .try_add_relationship(
            "urn:internal".to_string(),
            "target.xml".to_string(),
            "rId1".to_string(),
            TargetMode::Internal,
        )
        .unwrap();

    let serialized = relationships.try_to_xml_bytes().unwrap();
    assert_eq!(
        canonical_relationship_xml_len(&relationships).unwrap(),
        serialized.len()
    );
}

#[test]
fn topology_canonical_relationship_publication_honors_memory_admission() {
    const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    let source = topology_relationship_archive(&canonical_document_relationships(&[]));
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    for index in 0..64 {
        plan.try_add_external_relationship(
            owner.clone(),
            format!("rBudget{index}"),
            HYPERLINK,
            format!("https://example.invalid/{}", "x".repeat(2048)),
        )
        .unwrap();
    }
    let (budget, _cancellation_source, context) = managed_context_with_cancellation(64 * 1024);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Memory
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn topology_canonical_relationship_publication_honors_tight_limit_before_output() {
    const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    let source = topology_relationship_archive(&canonical_document_relationships(&[]));
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut expected = Relationships::for_source(&owner);
    expected
        .try_add_relationship(
            HYPERLINK.to_string(),
            "https://example.invalid/".to_string(),
            "rLimit".to_string(),
            TargetMode::External,
        )
        .unwrap();
    let limit = canonical_relationship_xml_len(&expected)
        .unwrap()
        .checked_sub(1)
        .expect("canonical relationship XML is nonempty");
    let limits = SourceLimits::builder()
        .max_relationship_xml_bytes(limit)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_external_relationship(owner, "rLimit", HYPERLINK, "https://example.invalid/")
        .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::Limit {
            kind: SourceLimitKind::RelationshipXmlBytes,
            ..
        }
    ));
    assert!(output.is_empty());
}

fn noncanonical_document_relationships() -> &'static [u8] {
    br#"<?xml version="1.0" encoding="UTF-8"?>
<!--preserve-before-->
<r:Relationships producer="opaque" xmlns="urn:litchi:unused" xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
  <?preserve instruction?>
  <r:Relationship TargetMode='External'
      Target='https://before.invalid/?a=1&amp;b=2#fragment'
      Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
      Id='rExisting'></r:Relationship>
  <!--preserve-between-->
  <r:Relationship Id='rInternal' Target="./target.xml" Type="urn:litchi:test"><?remaining?><!----></r:Relationship>
  <?preserve-after?>
</r:Relationships>
<!--preserve-tail-->"#
}

fn noncanonical_relationship_append_plan() -> SourceTopologyPlan {
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_external_relationship(
        PackURI::new("/word/document.xml").unwrap(),
        "rAdded",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        "https://added.invalid/?q=two#fragment",
    )
    .unwrap();
    plan
}

fn noncanonical_relationship_mixed_plan() -> SourceTopologyPlan {
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_external_relationship(owner.clone(), "rExisting")
        .unwrap();
    plan.try_add_external_relationship(
        owner,
        "rAdded",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        "https://added.invalid/?q=two#fragment",
    )
    .unwrap();
    plan
}

#[test]
fn topology_appends_prefixed_noncanonical_relationships_losslessly_and_reopens() {
    let relationships = noncanonical_document_relationships();
    let source = topology_relationship_archive(relationships);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, noncanonical_relationship_append_plan())
        .unwrap();

    let close = b"</r:Relationships>";
    let close_start = relationships
        .windows(close.len())
        .position(|window| window == close)
        .unwrap();
    let generated = br#"<r:Relationship Id="rAdded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://added.invalid/?q=two#fragment" TargetMode="External"/>"#;
    let mut expected_relationships = Vec::new();
    expected_relationships.extend_from_slice(&relationships[..close_start]);
    expected_relationships.extend_from_slice(generated);
    expected_relationships.extend_from_slice(&relationships[close_start..]);

    let archive = soapberry_zip::office::ArchiveReader::new(
        &output,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
    assert_eq!(
        archive.read("word/_rels/document.xml.rels").unwrap(),
        expected_relationships
    );

    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let relationships = reopened.part(&owner).unwrap().rels();
    assert_eq!(
        relationships.get("rExisting").unwrap().target_ref(),
        "https://before.invalid/?a=1&b=2#fragment"
    );
    assert_eq!(
        relationships.get("rAdded").unwrap().target_ref(),
        "https://added.invalid/?q=two#fragment"
    );
    assert_eq!(
        relationships.get("rAdded").unwrap().target_mode(),
        TargetMode::External
    );
    let reopened_owned = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(
        reopened_owned
            .get_part(&owner)
            .unwrap()
            .rels()
            .get("rAdded")
            .unwrap()
            .target_ref(),
        "https://added.invalid/?q=two#fragment"
    );
}

#[test]
fn topology_appends_default_namespace_noncanonical_relationships_losslessly() {
    let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <!--preserve-->
  <Relationship Target="https://before.invalid/" Type="urn:litchi:test" Id="rExisting" TargetMode="External" />
</Relationships>"#;
    let source = topology_relationship_archive(relationships);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, noncanonical_relationship_append_plan())
        .unwrap();

    let close = b"</Relationships>";
    let close_start = relationships
        .windows(close.len())
        .position(|window| window == close)
        .unwrap();
    let generated = br#"<Relationship Id="rAdded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://added.invalid/?q=two#fragment" TargetMode="External"/>"#;
    let mut expected_relationships = Vec::new();
    expected_relationships.extend_from_slice(&relationships[..close_start]);
    expected_relationships.extend_from_slice(generated);
    expected_relationships.extend_from_slice(&relationships[close_start..]);

    let archive = soapberry_zip::office::ArchiveReader::new(
        &output,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
    assert_eq!(
        archive.read("word/_rels/document.xml.rels").unwrap(),
        expected_relationships
    );
    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_eq!(
        reopened
            .part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .rels()
            .get("rAdded")
            .unwrap()
            .target_mode(),
        TargetMode::External
    );
}

#[test]
fn topology_mixes_noncanonical_relationship_removal_and_append_losslessly_with_bom() {
    let mut relationships = Vec::with_capacity(3 + noncanonical_document_relationships().len());
    relationships.extend_from_slice(b"\xEF\xBB\xBF");
    relationships.extend_from_slice(noncanonical_document_relationships());
    let source = topology_relationship_archive(&relationships);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut output = Vec::new();
    package
        .write_topology_to_stream(&mut output, noncanonical_relationship_mixed_plan())
        .unwrap();

    let removed_start = relationships
        .windows(b"<r:Relationship TargetMode".len())
        .position(|window| window == b"<r:Relationship TargetMode")
        .unwrap();
    let removed_end = relationships
        .windows(b"</r:Relationship>".len())
        .position(|window| window == b"</r:Relationship>")
        .unwrap()
        + b"</r:Relationship>".len();
    let close = b"</r:Relationships>";
    let close_start = relationships
        .windows(close.len())
        .position(|window| window == close)
        .unwrap();
    let mut compacted = Vec::with_capacity(relationships.len() - (removed_end - removed_start));
    compacted.extend_from_slice(&relationships[..removed_start]);
    compacted.extend_from_slice(&relationships[removed_end..]);
    let insertion = close_start - (removed_end - removed_start);
    let generated = br#"<r:Relationship Id="rAdded" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://added.invalid/?q=two#fragment" TargetMode="External"/>"#;
    let mut expected_relationships = Vec::with_capacity(compacted.len() + generated.len());
    expected_relationships.extend_from_slice(&compacted[..insertion]);
    expected_relationships.extend_from_slice(generated);
    expected_relationships.extend_from_slice(&compacted[insertion..]);

    let archive = soapberry_zip::office::ArchiveReader::new(
        &output,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
    assert_eq!(
        archive.read("word/_rels/document.xml.rels").unwrap(),
        expected_relationships
    );

    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let relationships = reopened.part(&owner).unwrap().rels();
    assert!(relationships.get("rExisting").is_none());
    assert_eq!(
        relationships.get("rInternal").unwrap().target_ref(),
        "./target.xml"
    );
    assert_eq!(
        relationships.get("rAdded").unwrap().target_ref(),
        "https://added.invalid/?q=two#fragment"
    );
    assert_eq!(
        relationships.get("rAdded").unwrap().target_mode(),
        TargetMode::External
    );
}

#[test]
fn topology_mixed_noncanonical_relationships_honor_output_byte_limit_before_output() {
    let relationships = noncanonical_document_relationships();
    let source = topology_relationship_archive(relationships);
    let relationship_limit = relationships.len().max(root_relationships().len());
    let limits = SourceLimits::builder()
        .max_relationship_xml_bytes(relationship_limit)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_external_relationship(owner.clone(), "rExisting")
        .unwrap();
    plan.try_add_external_relationship(
        owner,
        "rAdded",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        format!("https://added.invalid/?q={}", "x".repeat(1024)),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::Limit {
            kind: SourceLimitKind::RelationshipXmlBytes,
            ..
        }
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_mixed_noncanonical_relationships_honor_memory_admission_before_output() {
    let relationships = noncanonical_document_relationships();
    let source = topology_relationship_archive(relationships);
    let (_budget, _cancellation_source, context) =
        managed_context_with_cancellation((relationships.len() as u64).saturating_mul(2) - 1);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, noncanonical_relationship_mixed_plan())
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Memory
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_noncanonical_relationship_replace_refuses_opaque_children_before_output() {
    let relationships = std::str::from_utf8(noncanonical_document_relationships())
        .unwrap()
        .replace(
            "Id='rExisting'></r:Relationship>",
            "Id='rExisting'><!--opaque--></r:Relationship>",
        );
    let source = topology_relationship_archive(relationships.as_bytes());
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_external_relationship(
        owner,
        "rExisting",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        "https://replacement.invalid/",
    )
    .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_noncanonical_relationship_replace_preserves_unselected_source_bytes() {
    for self_closing in [false, true] {
        let mut xml = String::from("\u{feff}");
        xml.push_str(std::str::from_utf8(noncanonical_document_relationships()).unwrap());
        if self_closing {
            xml = xml.replace("Id='rExisting'></r:Relationship>", "Id='rExisting'/>");
        }
        let start = xml.find("<r:Relationship TargetMode").unwrap();
        let end = if self_closing {
            start + xml[start..].find("/>").unwrap() + 2
        } else {
            start + xml[start..].find("</r:Relationship>").unwrap() + "</r:Relationship>".len()
        };
        let mut expected = xml.clone();
        expected.replace_range(start..end, "");
        let insertion = expected.find("</r:Relationships>").unwrap();
        expected.insert_str(insertion, r#"<r:Relationship Id="rExisting" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://replacement.invalid/?a=1&amp;b=2" TargetMode="External"/>"#);
        let package = SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(topology_relationship_archive(
                xml.as_bytes(),
            ))),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let owner = PackURI::new("/word/document.xml").unwrap();
        let mut plan = SourceTopologyPlan::new();
        plan.try_replace_external_relationship(
            owner.clone(),
            "rExisting",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            "https://replacement.invalid/?a=1&b=2",
        )
        .unwrap();
        let mut output = Vec::new();
        package.write_topology_to_stream(&mut output, plan).unwrap();
        let archive = soapberry_zip::office::ArchiveReader::new(
            &output,
            soapberry_zip::office::ArchiveLimits::default(),
        )
        .unwrap();
        assert_eq!(
            archive.read("word/_rels/document.xml.rels").unwrap(),
            expected.as_bytes()
        );
        let reopened = SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(output)),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let part = reopened.part(&owner).unwrap();
        assert_eq!(part.rels().len(), 2);
        assert_eq!(
            part.rels().get("rExisting").unwrap().target_ref(),
            "https://replacement.invalid/?a=1&b=2"
        );
        assert_eq!(
            part.rels().get("rInternal").unwrap().target_ref(),
            "./target.xml"
        );
    }
}

#[test]
fn topology_noncanonical_relationship_replace_honors_byte_limit_before_output() {
    let relationships = noncanonical_document_relationships();
    let limits = SourceLimits::builder()
        .max_relationship_xml_bytes(relationships.len().max(root_relationships().len()))
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(topology_relationship_archive(
            relationships,
        ))),
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_external_relationship(
        PackURI::new("/word/document.xml").unwrap(),
        "rExisting",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        format!("https://replacement.invalid/?q={}", "x".repeat(1024)),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::Limit {
            kind: SourceLimitKind::RelationshipXmlBytes,
            ..
        }
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_noncanonical_relationship_remove_preserves_source_and_reopens() {
    let relationships = noncanonical_document_relationships();
    let source = topology_relationship_archive(relationships);
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_external_relationship(owner.clone(), "rExisting")
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();

    let removed_start = relationships
        .windows(b"<r:Relationship TargetMode".len())
        .position(|window| window == b"<r:Relationship TargetMode")
        .unwrap();
    let removed_end = relationships
        .windows(b"</r:Relationship>".len())
        .position(|window| window == b"</r:Relationship>")
        .unwrap()
        + b"</r:Relationship>".len();
    let mut expected_relationships = Vec::new();
    expected_relationships.extend_from_slice(&relationships[..removed_start]);
    expected_relationships.extend_from_slice(&relationships[removed_end..]);

    let archive = soapberry_zip::office::ArchiveReader::new(
        &output,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
    assert_eq!(
        archive.read("word/_rels/document.xml.rels").unwrap(),
        expected_relationships
    );
    assert!(
        expected_relationships
            .windows(b"producer=\"opaque\"".len())
            .any(|window| window == b"producer=\"opaque\"")
    );
    assert!(
        expected_relationships
            .windows(b"preserve-before".len())
            .any(|window| window == b"preserve-before")
    );
    assert!(
        expected_relationships
            .windows(b"preserve-after".len())
            .any(|window| window == b"preserve-after")
    );

    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let relationships = reopened.part(&owner).unwrap().rels();
    assert!(relationships.get("rExisting").is_none());
    assert_eq!(
        relationships.get("rInternal").unwrap().target_ref(),
        "./target.xml"
    );
}

#[test]
fn topology_noncanonical_relationship_remove_last_keeps_empty_member_and_opaque_context() {
    let relationships = br#"<?xml version="1.0" encoding="UTF-8"?>
<!--before-->
<r:Relationships producer="opaque" xmlns="urn:litchi:unused" xmlns:r="http://schemas.openxmlformats.org/package/2006/relationships">
  <?inside?>
  <r:Relationship Id="rOnly" Type="urn:litchi:test" Target="https://remove.invalid/" TargetMode="External"></r:Relationship>
  <!--after-->
</r:Relationships>
<!--tail-->"#;
    let source = topology_relationship_archive(relationships);
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_remove_external_relationship(owner.clone(), "rOnly")
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();

    let removed_start = relationships
        .windows(b"<r:Relationship Id=\"rOnly\"".len())
        .position(|window| window == b"<r:Relationship Id=\"rOnly\"")
        .unwrap();
    let removed_end = relationships
        .windows(b"</r:Relationship>".len())
        .position(|window| window == b"</r:Relationship>")
        .unwrap()
        + b"</r:Relationship>".len();
    let mut expected_relationships = Vec::new();
    expected_relationships.extend_from_slice(&relationships[..removed_start]);
    expected_relationships.extend_from_slice(&relationships[removed_end..]);

    let archive = soapberry_zip::office::ArchiveReader::new(
        &output,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
    assert_eq!(
        archive.read("word/_rels/document.xml.rels").unwrap(),
        expected_relationships
    );
    assert!(
        expected_relationships
            .windows(b"producer=\"opaque\"".len())
            .any(|window| window == b"producer=\"opaque\"")
    );

    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert!(reopened.part(&owner).unwrap().rels().is_empty());
}

#[test]
fn topology_noncanonical_relationship_inverse_add_retains_source_context() {
    let relationships = noncanonical_document_relationships();
    let source = topology_relationship_archive(relationships);
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut remove = SourceTopologyPlan::new();
    remove
        .try_remove_external_relationship(owner.clone(), "rExisting")
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut removed_output = Vec::new();
    package
        .write_topology_to_stream(&mut removed_output, remove)
        .unwrap();

    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(removed_output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut add = SourceTopologyPlan::new();
    add.try_add_external_relationship(
        owner.clone(),
        "rExisting",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        "https://before.invalid/?a=1&b=2#fragment",
    )
    .unwrap();
    let mut inverse_output = Vec::new();
    reopened
        .write_topology_to_stream(&mut inverse_output, add)
        .unwrap();

    let archive = soapberry_zip::office::ArchiveReader::new(
        &inverse_output,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
    let inverse_relationships = archive.read("word/_rels/document.xml.rels").unwrap();
    assert!(
        inverse_relationships
            .windows(b"producer=\"opaque\"".len())
            .any(|window| window == b"producer=\"opaque\"")
    );
    assert!(
        inverse_relationships
            .windows(b"preserve-before".len())
            .any(|window| window == b"preserve-before")
    );
    assert!(
        inverse_relationships
            .windows(b"preserve-after".len())
            .any(|window| window == b"preserve-after")
    );

    let inverse = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(inverse_output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let relationship = inverse
        .part(&owner)
        .unwrap()
        .rels()
        .get("rExisting")
        .unwrap();
    assert_eq!(relationship.target_mode(), TargetMode::External);
    assert_eq!(
        relationship.target_ref(),
        "https://before.invalid/?a=1&b=2#fragment"
    );
}

#[test]
fn topology_noncanonical_relationship_unknown_markup_refuses_before_output() {
    let relationships = br#"<Relationships foo="unknown" xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExisting" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://before.invalid/" TargetMode="External">opaque</Relationship></Relationships>"#;
    let source = topology_relationship_archive(relationships);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, noncanonical_relationship_append_plan())
        .unwrap_err();
    assert!(matches!(error, OpcError::InvalidRelationshipsManifest(_)));
    assert!(output.is_empty());
}

#[test]
fn topology_noncanonical_relationship_parser_rejects_dtd_namespace_and_root_forms() {
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut expected = Relationships::for_source(&owner);
    expected
        .try_add_relationship(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
                .to_string(),
            "https://before.invalid/".to_string(),
            "rExisting".to_string(),
            TargetMode::External,
        )
        .unwrap();
    let unsafe_sources: &[&[u8]] = &[
            br#"<!DOCTYPE Relationships><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExisting" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://before.invalid/" TargetMode="External"/></Relationships>"#,
            br#"<Relationships xmlns="urn:litchi:wrong"><Relationship Id="rExisting" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://before.invalid/" TargetMode="External"/></Relationships>"#,
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"/>"#,
            br#"<Relationships 1bad="opaque" xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExisting" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://before.invalid/" TargetMode="External"/></Relationships>"#,
            br#"<Relationships p:producer="opaque" xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExisting" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://before.invalid/" TargetMode="External"/></Relationships>"#,
            br#"<Relationships xmlns:p="urn:litchi:alias" xmlns:q="urn:litchi:alias" p:producer="one" q:producer="two" xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rExisting" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://before.invalid/" TargetMode="External"/></Relationships>"#,
        ];
    for source in unsafe_sources {
        assert!(
            parse_noncanonical_relationship_source(
                source,
                "word/_rels/document.xml.rels",
                &expected,
                SourceLimits::default(),
                || Ok(()),
            )
            .is_err()
        );
    }
    let canonical = String::from_utf8(expected.try_to_xml_bytes().unwrap()).unwrap();
    assert!(
        parse_noncanonical_relationship_source(
            canonical.as_bytes(),
            "word/_rels/document.xml.rels",
            &expected,
            SourceLimits::default(),
            || Ok(()),
        )
        .is_ok()
    );
    for (before, after) in [
        ("encoding=\"UTF-8\"", "encoding=\"UTF-16\""),
        ("version=\"1.0\"", "version=\"1.&#48;\""),
    ] {
        let invalid = canonical.replace(before, after);
        assert!(
            parse_noncanonical_relationship_source(
                invalid.as_bytes(),
                "word/_rels/document.xml.rels",
                &expected,
                SourceLimits::default(),
                || Ok(()),
            )
            .is_err()
        );
    }
    let shallow = SourceLimits::builder()
        .max_depth(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        parse_noncanonical_relationship_source(
            canonical.as_bytes(),
            "word/_rels/document.xml.rels",
            &expected,
            shallow,
            || Ok(()),
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::Depth,
            ..
        })
    ));
}

#[test]
fn topology_noncanonical_relationship_append_honors_limits_before_output() {
    let relationships = noncanonical_document_relationships();
    let source = topology_relationship_archive(relationships);
    let relationship_limit = relationships.len().max(root_relationships().len());
    let limits = SourceLimits::builder()
        .max_relationship_xml_bytes(relationship_limit)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, noncanonical_relationship_append_plan())
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::Limit {
            kind: SourceLimitKind::RelationshipXmlBytes,
            ..
        }
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_noncanonical_relationship_append_honors_memory_admission() {
    let relationships = noncanonical_document_relationships();
    let source = topology_relationship_archive(relationships);
    let (_budget, _cancellation_source, context) =
        managed_context_with_cancellation((relationships.len() as u64).saturating_mul(2) - 1);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, noncanonical_relationship_append_plan())
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Memory
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_noncanonical_relationship_append_honors_cancellation() {
    let source = topology_relationship_archive(noncanonical_document_relationships());
    let (_budget, cancellation_source, context) = managed_context_with_cancellation(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    cancellation_source.cancel();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, noncanonical_relationship_append_plan())
        .unwrap_err();
    assert!(matches!(error, OpcError::Cancelled));
    assert!(output.is_empty());
}

#[test]
fn source_relationship_handle_absence_is_fenced_without_budget_charge() {
    let source = archive_bytes(root_relationships(), b"<document/>", false);
    let (budget, _cancellation_source, context) =
        managed_context_with_all_resources(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let owner = PackURI::new("/word/document.xml").unwrap();
    let memory_before = budget.used(Resource::Memory);
    let objects_before = budget.used(Resource::Objects);
    let work_before = budget.used(Resource::Work);

    assert!(
        package
            .relationships_data_for_with_limit(&owner, 64 * 1024)
            .unwrap()
            .is_none()
    );
    assert_eq!(budget.used(Resource::Memory), memory_before);
    assert_eq!(budget.used(Resource::Objects), objects_before);
    assert_eq!(budget.used(Resource::Work), work_before);
}

#[test]
fn source_relationship_handle_charges_work_and_releases_memory_on_drop() {
    let relationships = canonical_document_relationships(&[]);
    let source = topology_relationship_archive(&relationships);
    let (budget, _cancellation_source, context) =
        managed_context_with_all_resources(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let owner = PackURI::new("/word/document.xml").unwrap();
    let memory_before = budget.used(Resource::Memory);
    let objects_before = budget.used(Resource::Objects);
    let work_before = budget.used(Resource::Work);
    let data = package
        .relationships_data_for_with_limit(&owner, relationships.len())
        .unwrap()
        .expect("the physical relationship member is present");

    assert_eq!(data.as_bytes(), relationships.as_slice());
    assert_eq!(
        budget.used(Resource::Memory),
        memory_before + relationships.len() as u64
    );
    assert_eq!(budget.used(Resource::Objects), objects_before + 1);
    assert_eq!(
        budget.used(Resource::Work),
        work_before + relationships.len() as u64
    );
    drop(data);
    assert_eq!(budget.used(Resource::Memory), memory_before);
    assert_eq!(budget.used(Resource::Objects), objects_before);
    // Work is cumulative by design and therefore remains charged after
    // the source handle releases its retained decoded bytes.
    assert_eq!(
        budget.used(Resource::Work),
        work_before + relationships.len() as u64
    );
}

#[test]
fn source_relationship_handle_checks_cancellation_before_read() {
    let relationships = canonical_document_relationships(&[]);
    let source = topology_relationship_archive(&relationships);
    let (budget, cancellation_source, context) =
        managed_context_with_all_resources(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let owner = PackURI::new("/word/document.xml").unwrap();
    let memory_before = budget.used(Resource::Memory);
    let work_before = budget.used(Resource::Work);
    cancellation_source.cancel();

    assert!(matches!(
        package.relationships_data_for_with_limit(&owner, relationships.len()),
        Err(OpcError::Cancelled)
    ));
    assert_eq!(budget.used(Resource::Memory), memory_before);
    assert_eq!(budget.used(Resource::Work), work_before);
}

#[test]
fn source_content_types_handle_charges_work_and_releases_memory_on_drop() {
    let source = archive_bytes(root_relationships(), b"<document/>", false);
    let (budget, _cancellation_source, context) =
        managed_context_with_all_resources(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let entry = package
        .archive
        .entry_id(&package.catalog.content_types_member)
        .expect("the content-types member is present");
    let expected = package.archive.read_entry(entry).unwrap();
    let memory_before = budget.used(Resource::Memory);
    let objects_before = budget.used(Resource::Objects);
    let work_before = budget.used(Resource::Work);

    let data = package
        .content_types_data_with_limit(expected.len())
        .unwrap();
    assert_eq!(data.as_bytes(), expected.as_slice());
    assert_eq!(
        budget.used(Resource::Memory),
        memory_before + expected.len() as u64
    );
    assert_eq!(budget.used(Resource::Objects), objects_before + 1);
    assert_eq!(
        budget.used(Resource::Work),
        work_before + expected.len() as u64
    );
    drop(data);
    assert_eq!(budget.used(Resource::Memory), memory_before);
    assert_eq!(budget.used(Resource::Objects), objects_before);
    // Work is cumulative by design and therefore remains charged after
    // the source handle releases its retained decoded bytes.
    assert_eq!(
        budget.used(Resource::Work),
        work_before + expected.len() as u64
    );
}

#[test]
fn source_content_types_handle_checks_caller_limit_before_budget_charge() {
    let source = archive_bytes(root_relationships(), b"<document/>", false);
    let (budget, _cancellation_source, context) =
        managed_context_with_all_resources(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let entry = package
        .archive
        .entry_id(&package.catalog.content_types_member)
        .expect("the content-types member is present");
    let declared = package
        .archive
        .metadata_for(entry)
        .unwrap()
        .uncompressed_size() as usize;
    let maximum = declared - 1;
    let memory_before = budget.used(Resource::Memory);
    let objects_before = budget.used(Resource::Objects);
    let work_before = budget.used(Resource::Work);

    let error = package.content_types_data_with_limit(maximum).unwrap_err();
    assert!(matches!(
        error,
        OpcError::Limit {
            kind: SourceLimitKind::ContentTypesBytes,
            observed: actual,
            maximum: observed_maximum,
        } if actual == declared as u64 && observed_maximum == maximum as u64
    ));
    assert_eq!(budget.used(Resource::Memory), memory_before);
    assert_eq!(budget.used(Resource::Objects), objects_before);
    assert_eq!(budget.used(Resource::Work), work_before);
}

#[test]
fn source_content_types_handle_checks_cancellation_before_read() {
    let source = archive_bytes(root_relationships(), b"<document/>", false);
    let (budget, cancellation_source, context) =
        managed_context_with_all_resources(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let memory_before = budget.used(Resource::Memory);
    let objects_before = budget.used(Resource::Objects);
    let work_before = budget.used(Resource::Work);
    cancellation_source.cancel();

    assert!(matches!(
        package.content_types_data_with_limit(usize::MAX),
        Err(OpcError::Cancelled)
    ));
    assert_eq!(budget.used(Resource::Memory), memory_before);
    assert_eq!(budget.used(Resource::Objects), objects_before);
    assert_eq!(budget.used(Resource::Work), work_before);
}

#[test]
fn source_content_types_handle_releases_reservations_when_source_changes_during_read() {
    let source_bytes = archive_bytes(root_relationships(), b"<document/>", false);
    let marker = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#;
    let payload_offset = source_bytes
        .windows(marker.len())
        .position(|window| window == marker)
        .expect("content-types payload is present");
    let source = Arc::new(ChangeDuringPayloadSource::new(source_bytes, payload_offset));
    let (budget, _cancellation_source, context) =
        managed_context_with_all_resources(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let memory_before = budget.used(Resource::Memory);
    let objects_before = budget.used(Resource::Objects);
    let work_before = budget.used(Resource::Work);
    source.armed.store(true, Ordering::SeqCst);

    assert!(matches!(
        package.content_types_data_with_limit(usize::MAX),
        Err(OpcError::SourceChanged { .. })
    ));
    assert_eq!(budget.used(Resource::Memory), memory_before);
    assert_eq!(budget.used(Resource::Objects), objects_before);
    // Work is charged for the attempted decode even though publication of
    // the returned handle is rejected by the source-version fence.
    assert!(budget.used(Resource::Work) > work_before);
}

#[test]
fn source_signature_query_is_metadata_only_and_conservative() {
    let ordinary = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<document/>",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert!(!ordinary.has_signature_infrastructure());

    let (budget, _cancellation_source, context) =
        managed_context_with_all_resources(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let signed = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(signed_archive(b"<document/>"))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let memory_before = budget.used(Resource::Memory);
    let objects_before = budget.used(Resource::Objects);
    let work_before = budget.used(Resource::Work);
    assert!(signed.has_signature_infrastructure());
    assert_eq!(budget.used(Resource::Memory), memory_before);
    assert_eq!(budget.used(Resource::Objects), objects_before);
    assert_eq!(budget.used(Resource::Work), work_before);
}

#[test]
fn topology_adds_external_relationship_with_exact_uri_reference() {
    let relationships = canonical_document_relationships(&[]);
    let source = topology_relationship_archive(&relationships);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_external_relationship(
        owner.clone(),
        "rLink",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        "https://example.invalid/path?q=one#bookmark",
    )
    .unwrap();

    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = reopened.part(&owner).unwrap();
    let relationship = part.rels().get("rLink").unwrap();
    assert_eq!(relationship.target_mode(), TargetMode::External);
    assert_eq!(
        relationship.target_ref(),
        "https://example.invalid/path?q=one#bookmark"
    );
}

#[test]
fn topology_replaces_and_removes_external_relationships() {
    const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    let relationships = canonical_document_relationships(&[
        (
            "rReplace",
            HYPERLINK,
            "https://before.invalid/",
            TargetMode::External,
        ),
        (
            "rRemove",
            HYPERLINK,
            "https://remove.invalid/",
            TargetMode::External,
        ),
    ]);
    let source = topology_relationship_archive(&relationships);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_external_relationship(
        owner.clone(),
        "rReplace",
        HYPERLINK,
        "https://after.invalid/?q=two#fragment",
    )
    .unwrap();
    plan.try_remove_external_relationship(owner.clone(), "rRemove")
        .unwrap();

    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = reopened.part(&owner).unwrap();
    let relationships = part.rels();
    assert_eq!(
        relationships.get("rReplace").unwrap().target_ref(),
        "https://after.invalid/?q=two#fragment"
    );
    assert!(relationships.get("rRemove").is_none());
}

#[test]
fn topology_exact_external_replacement_preserves_noncanonical_source() {
    const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    let relationships = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship TargetMode="External" Target="https://same.invalid/?q=1#same" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Id="rSame"/></Relationships>"#;
    let source = topology_relationship_archive(relationships);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_external_relationship(
        PackURI::new("/word/document.xml").unwrap(),
        "rSame",
        HYPERLINK,
        "https://same.invalid/?q=1#same",
    )
    .unwrap();

    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    assert_eq!(output, source);
}

#[test]
fn topology_signed_noop_copies_and_mutation_refuses_before_output() {
    const OFFICE_DOCUMENT: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    let source = signed_archive(b"<signed/>");
    let owner = PackURI::new("/").unwrap();
    let target = PackURI::new("/word/document.xml").unwrap();
    let mut noop = SourceTopologyPlan::new();
    noop.try_replace_internal_relationship(owner.clone(), "rId1", OFFICE_DOCUMENT, target.clone())
        .unwrap();
    let mut output = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, noop)
    .unwrap();
    assert_eq!(output, source);

    let mut mutation = SourceTopologyPlan::new();
    mutation
        .try_replace_internal_relationship(
            owner,
            "rId1",
            OFFICE_DOCUMENT,
            PackURI::new("/signature/origin.xml").unwrap(),
        )
        .unwrap();
    output.clear();
    let error = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, mutation)
    .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SignedSourceRequiresExplicitPolicy
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_signed_part_noop_copies_and_mutation_refuses_before_output() {
    let source = signed_archive(b"<signed/>");
    let target = PackURI::new("/word/document.xml").unwrap();

    let mut noop = SourceTopologyPlan::new();
    noop.try_replace_part(target.clone(), b"<signed/>".to_vec())
        .unwrap();
    let mut output = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, noop)
    .unwrap();
    assert_eq!(output, source);

    let mut mutation = SourceTopologyPlan::new();
    mutation
        .try_replace_part(target, b"<changed/>".to_vec())
        .unwrap();
    output.clear();
    let error = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_topology_to_stream(&mut output, mutation)
    .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SignedSourceRequiresExplicitPolicy
    ));
    assert!(output.is_empty());
}

#[test]
fn topology_external_wrappers_refuse_internal_relationships_before_output() {
    const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    let relationships = canonical_document_relationships(&[(
        "rInternal",
        HYPERLINK,
        "document.xml",
        TargetMode::Internal,
    )]);
    let source = topology_relationship_archive(&relationships);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_external_relationship(
        PackURI::new("/word/document.xml").unwrap(),
        "rInternal",
        HYPERLINK,
        "https://example.invalid/",
    )
    .unwrap();

    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(error, OpcError::InvalidRelationship(_)));
    assert!(output.is_empty());
}

#[test]
fn topology_internal_replace_requires_and_preserves_internal_mode() {
    const TEST_RELATIONSHIP: &str = "urn:litchi:test";
    let relationships = canonical_document_relationships(&[
        (
            "rInternal",
            TEST_RELATIONSHIP,
            "document.xml",
            TargetMode::Internal,
        ),
        (
            "rExternal",
            TEST_RELATIONSHIP,
            "https://example.invalid/",
            TargetMode::External,
        ),
    ]);
    let source = topology_relationship_archive(&relationships);
    let owner = PackURI::new("/word/document.xml").unwrap();
    let mut plan = SourceTopologyPlan::new();
    plan.try_replace_internal_relationship(
        owner.clone(),
        "rInternal",
        TEST_RELATIONSHIP,
        PackURI::new("/word/target.xml").unwrap(),
    )
    .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan).unwrap();
    let reopened = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(output)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = reopened.part(&owner).unwrap();
    let relationship = part.rels().get("rInternal").unwrap();
    assert_eq!(relationship.target_mode(), TargetMode::Internal);
    assert_eq!(
        relationship.target_partname().unwrap(),
        PackURI::new("/word/target.xml").unwrap()
    );

    let mut mismatch = SourceTopologyPlan::new();
    mismatch
        .try_replace_internal_relationship(
            owner,
            "rExternal",
            TEST_RELATIONSHIP,
            PackURI::new("/word/target.xml").unwrap(),
        )
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut mismatch_output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut mismatch_output, mismatch)
        .unwrap_err();
    assert!(matches!(error, OpcError::InvalidRelationship(_)));
    assert!(mismatch_output.is_empty());
}

#[test]
fn topology_relationship_publication_honors_configured_reopen_limits() {
    const HYPERLINK: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
    let relationships = canonical_document_relationships(&[]);
    let source = topology_relationship_archive(&relationships);
    let owner = PackURI::new("/word/document.xml").unwrap();

    let target_limits = SourceLimits::builder()
        .max_relationship_target_bytes(24)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source.clone())),
        target_limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut target_plan = SourceTopologyPlan::new();
    target_plan
        .try_add_external_relationship(
            owner.clone(),
            "rTargetLimit",
            HYPERLINK,
            "https://example.invalid/too-long",
        )
        .unwrap();
    let mut target_output = Vec::new();
    assert!(matches!(
        package.write_topology_to_stream(&mut target_output, target_plan),
        Err(OpcError::Limit {
            kind: SourceLimitKind::RelationshipTargetBytes,
            ..
        })
    ));
    assert!(target_output.is_empty());

    let attribute_limits = SourceLimits::builder()
        .max_xml_attribute_bytes(128)
        .unwrap()
        .max_relationship_target_bytes(128)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source.clone())),
        attribute_limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut attribute_plan = SourceTopologyPlan::new();
    attribute_plan
        .try_add_external_relationship(
            owner.clone(),
            "rAttributeLimit",
            HYPERLINK,
            format!("https://example.invalid/?{}", "&".repeat(30)),
        )
        .unwrap();
    let mut attribute_output = Vec::new();
    assert!(matches!(
        package.write_topology_to_stream(&mut attribute_output, attribute_plan),
        Err(OpcError::Limit {
            kind: SourceLimitKind::XmlAttributeBytes,
            ..
        })
    ));
    assert!(attribute_output.is_empty());

    let name_limits = SourceLimits::builder()
        .max_archive_member_name_bytes(32)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source.clone())),
        name_limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut name_plan = SourceTopologyPlan::new();
    name_plan
        .try_add_part(
            PackURI::new(format!("/custom/{}.xml", "x".repeat(40))).unwrap(),
            "application/xml",
            b"<new/>".to_vec(),
        )
        .unwrap();
    let mut name_output = Vec::new();
    assert!(matches!(
        package.write_topology_to_stream(&mut name_output, name_plan),
        Err(OpcError::Limit {
            kind: SourceLimitKind::ArchiveMemberNameBytes,
            ..
        })
    ));
    assert!(name_output.is_empty());

    let default_limits = SourceLimits::default();
    let source_events = relationship_xml_event_count(root_relationships(), default_limits)
        .unwrap()
        .checked_add(relationship_xml_event_count(&relationships, default_limits).unwrap())
        .unwrap();
    let event_limits = SourceLimits::builder()
        .max_events(5)
        .unwrap()
        .max_total_relationship_xml_events(source_events as usize)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        event_limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut event_plan = SourceTopologyPlan::new();
    event_plan
        .try_add_external_relationship(owner, "rEventLimit", HYPERLINK, "https://example.invalid/")
        .unwrap();
    let mut event_output = Vec::new();
    assert!(matches!(
        package.write_topology_to_stream(&mut event_output, event_plan),
        Err(OpcError::Limit {
            kind: SourceLimitKind::TotalRelationshipXmlEvents,
            ..
        })
    ));
    assert!(event_output.is_empty());
}

fn relationship_batch(prefix: &str, count: usize) -> (Vec<String>, Vec<u8>) {
    let mut ids = Vec::new();
    let mut xml = String::from(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for index in 0..count {
        let id = format!("r{prefix}{index}");
        ids.push(id.clone());
        xml.push_str(&format!(
                r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://{prefix}{index}.invalid/" TargetMode="External"/>"#
            ));
    }
    xml.push_str("</Relationships>");
    (ids, xml.into_bytes())
}

fn archive_with_relationship_batches(
    first: usize,
    second: usize,
) -> (Vec<u8>, Vec<String>, Vec<String>) {
    let (first_ids, first_rels) = relationship_batch("first", first);
    let (second_ids, second_rels) = relationship_batch("second", second);
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships())
        .unwrap();
    writer
        .write_stored("word/document.xml", b"<before/>")
        .unwrap();
    writer
        .write_stored("word/_rels/document.xml.rels", &first_rels)
        .unwrap();
    writer
        .write_stored("custom/orphan.xml", b"<before/>")
        .unwrap();
    writer
        .write_stored("custom/_rels/orphan.xml.rels", &second_rels)
        .unwrap();
    (writer.finish_to_bytes().unwrap(), first_ids, second_ids)
}

fn source_archive_total(package: &SourceBackedPackage) -> u64 {
    package
        .archive
        .file_names()
        .map(|name| package.archive.metadata(name).unwrap().uncompressed_size())
        .sum()
}

#[test]
fn source_artifact_paths_reject_overreported_read_counts_without_output() {
    let source = Arc::new(OverReportingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        false,
    )));
    let read_at: Arc<dyn ReadAt> = source.clone();
    let package = SourceBackedPackage::from_read_at(
        read_at,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let artifact = package.source_artifact();
    source.overreport.store(true, Ordering::SeqCst);

    assert!(matches!(
        artifact.fingerprint(),
        Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
    ));

    let mut output = Vec::new();
    assert!(matches!(
        artifact.write_to(&mut output),
        Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
    ));
    assert!(output.is_empty());
}

#[test]
fn source_artifact_copy_rejects_overreporting_sink_without_false_progress() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"<before/>",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut sink = OverReportingSink {
        calls: 0,
        accepted: 0,
    };

    assert!(matches!(
        artifact.write_to(&mut sink),
        Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
    ));
    assert_eq!(sink.calls, 1);
    assert_eq!(sink.accepted, 0);
}

fn root_relationships() -> &'static [u8] {
    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#
}

fn signed_root_relationships() -> &'static [u8] {
    br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" Target="signature/origin.xml"/></Relationships>"#
}

fn signed_archive(document: &[u8]) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", signed_root_relationships())
        .unwrap();
    writer.write_stored("word/document.xml", document).unwrap();
    writer
        .write_stored("signature/origin.xml", b"<origin/>")
        .unwrap();
    writer.finish_to_bytes().unwrap()
}

#[derive(Debug)]
struct RawRecord {
    local: Vec<u8>,
    central: Vec<u8>,
}

fn raw_records(data: &[u8]) -> HashMap<Vec<u8>, RawRecord> {
    let archive = soapberry_zip::ZipArchive::from_slice(data)
        .unwrap()
        .into_zip_archive();
    let mut scratch = vec![0; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = soapberry_zip::PreservationIndex::new(
        &archive,
        &mut scratch,
        soapberry_zip::office::ArchiveLimits::default(),
    )
    .unwrap();
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

fn central_without_local_offset(bytes: &[u8]) -> Vec<u8> {
    let mut bytes = bytes.to_vec();
    bytes[42..46].fill(0);
    bytes
}

struct MutatingSink {
    source: Arc<CountingSource>,
    bytes: Vec<u8>,
    change_after: usize,
    changed: bool,
}

impl Write for MutatingSink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        if !self.changed && self.bytes.len() >= self.change_after {
            self.source.changed();
            self.changed = true;
        }
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

struct BoundedFailingSink {
    accepted: usize,
    limit: usize,
    largest_write: usize,
}

struct CancelAfterWriteSink {
    cancellation_source: CancellationSource,
    bytes: Vec<u8>,
    cancelled: bool,
}

struct ShortInterruptedSink {
    bytes: Vec<u8>,
    maximum_write: usize,
    interrupted: bool,
}

struct FlushFailingSink {
    bytes: Vec<u8>,
}

impl Write for CancelAfterWriteSink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        if !self.cancelled {
            self.cancelled = true;
            self.cancellation_source.cancel();
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

impl Write for BoundedFailingSink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.largest_write = self.largest_write.max(bytes.len());
        let remaining = self.limit.saturating_sub(self.accepted);
        if remaining == 0 {
            return Err(std::io::Error::other("injected sink failure"));
        }
        let written = remaining.min(bytes.len());
        self.accepted += written;
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

impl Write for ShortInterruptedSink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        if !self.interrupted {
            self.interrupted = true;
            return Err(std::io::Error::new(
                std::io::ErrorKind::Interrupted,
                "injected retryable interruption",
            ));
        }
        if bytes.is_empty() {
            return Ok(0);
        }
        let written = self.maximum_write.min(bytes.len());
        self.bytes.extend_from_slice(&bytes[..written]);
        Ok(written)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}

impl Write for FlushFailingSink {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.bytes.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Err(std::io::Error::other("injected flush failure"))
    }
}

#[test]
fn mandatory_xml_is_opened_but_ordinary_payload_corruption_is_deferred() {
    let malformed = archive_bytes(b"<Relationships", b"document", false);
    let malformed_source = Arc::new(CountingSource::new(malformed));
    assert!(matches!(
        SourceBackedPackage::from_read_at(
            malformed_source,
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None
        ),
        Err(OpcError::QuickXmlError(_))
    ));

    const DOCUMENT: &[u8] = b"source-backed deferred corruption";
    let mut corrupt = archive_bytes(root_relationships(), DOCUMENT, false);
    let position = corrupt
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    corrupt[position] ^= 0xff;
    let source = Arc::new(CountingSource::new(corrupt));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let main = package.main_document_part().unwrap();
    assert!(matches!(main.data(), Err(OpcError::ZipError(_))));
}

#[test]
fn source_backed_catalog_open_does_not_read_ordinary_payloads() {
    const DOCUMENT: &[u8] = b"source-backed cold catalog payload sentinel";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_start = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .expect("stored payload must be present in the archive");
    let payload_end = payload_start + DOCUMENT.len();
    let source = Arc::new(CountingSource::new(bytes));

    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_eq!(package.cache_diagnostics().cold_loads, 0);
    assert_eq!(package.cache_diagnostics().retained_entries, 0);
    assert_eq!(package.iter_parts().count(), 2);
    assert!(source.read_bytes.load(Ordering::SeqCst) > 0);
    assert!(source.read_ranges().into_iter().all(|(offset, length)| {
        let start = usize::try_from(offset).expect("test source offset fits usize");
        let end = start + length;
        end <= payload_start || start >= payload_end
    }));
}

#[test]
fn source_backed_catalog_matches_eager_part_admission() {
    let bytes = archive_bytes(root_relationships(), b"catalog parity", true);
    let eager = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    let source = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    let mut eager_parts: Vec<_> = eager
        .iter_parts()
        .map(|part| {
            let mut relationships: Vec<_> = part
                .rels()
                .iter()
                .map(|relationship| {
                    (
                        relationship.r_id().to_owned(),
                        relationship.reltype().to_owned(),
                        relationship.target_ref().to_owned(),
                        relationship.target_mode(),
                    )
                })
                .collect();
            relationships.sort_by(|left, right| left.0.cmp(&right.0));
            (
                part.partname().to_string(),
                part.content_type().to_owned(),
                relationships,
            )
        })
        .collect();
    let mut deferred_parts: Vec<_> = source
        .iter_parts()
        .map(|part| {
            let mut relationships: Vec<_> = part
                .rels()
                .iter()
                .map(|relationship| {
                    (
                        relationship.r_id().to_owned(),
                        relationship.reltype().to_owned(),
                        relationship.target_ref().to_owned(),
                        relationship.target_mode(),
                    )
                })
                .collect();
            relationships.sort_by(|left, right| left.0.cmp(&right.0));
            (
                part.partname().to_string(),
                part.content_type().to_owned(),
                relationships,
            )
        })
        .collect();
    eager_parts.sort_by(|left, right| left.0.cmp(&right.0));
    deferred_parts.sort_by(|left, right| left.0.cmp(&right.0));

    assert_eq!(eager_parts, deferred_parts);
    assert_eq!(eager.non_part_members(), source.non_part_members());
    assert_eq!(eager.rels().len(), source.rels().len());
    assert_eq!(
        eager.main_document_part().unwrap().partname(),
        source.main_document_part().unwrap().partname()
    );
}

#[test]
fn source_backed_catalog_matches_eager_relationship_typing_and_non_parts() {
    let bytes = archive_with_document_relationships(document_relationships());
    let eager = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();
    let source = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    assert_eq!(eager.non_part_members(), source.non_part_members());
    assert_eq!(source.non_part_members().len(), 1);
    assert_eq!(source.non_part_members()[0].name(), "scratch.bin");
    assert!(
        source
            .physical_member_names()
            .any(|name| name == "word/_rels/document.xml.rels")
    );
    assert!(
        source
            .iter_parts()
            .all(|part| part.partname().as_str() != "/word/_rels/document.xml.rels")
    );
    assert_eq!(
        source
            .part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .rels()
            .len(),
        eager
            .get_part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .rels()
            .len()
    );
}

#[test]
fn source_backed_catalog_enforces_part_limits_during_ordered_admission() {
    let bytes = archive_with_ordered_xml_parts(&[
        ("word/document.xml", b"document"),
        ("custom/orphan.xml", b"orphan"),
    ]);

    let per_part = SourceLimits::builder()
        .max_part_bytes(b"document".len() as u64 - 1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(bytes.clone())),
            per_part,
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::PartBytes,
            ..
        })
    ));

    let aggregate = SourceLimits::builder()
        .max_total_part_bytes((b"document".len() + b"orphan".len() - 1) as u64)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(bytes.clone())),
            aggregate,
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::TotalPartBytes,
            ..
        })
    ));

    let part_count = SourceLimits::builder()
        .max_parts(1)
        .unwrap()
        .build()
        .unwrap();
    assert!(matches!(
        SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(bytes)),
            part_count,
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::Parts,
            ..
        })
    ));
}

#[test]
fn source_backed_part_limit_precedes_later_derived_or_equivalent_conflicts() {
    let oversized = b"oversized first part";
    let cases = [
        (
            [
                ("word/oversized.xml", &oversized[..]),
                ("word/container.xml", &b"container"[..]),
                ("word/container.xml/child.xml", &b"child"[..]),
            ],
            true,
        ),
        (
            [
                ("word/oversized.xml", &oversized[..]),
                ("word/equivalent.xml", &b"first"[..]),
                ("WORD/EQUIVALENT.XML", &b"second"[..]),
            ],
            false,
        ),
    ];
    let limits = SourceLimits::builder()
        .max_part_bytes(oversized.len() as u64 - 1)
        .unwrap()
        .build()
        .unwrap();

    for (parts, derived) in cases {
        let bytes = archive_with_ordered_xml_parts(&parts);
        assert!(matches!(
            SourceBackedPackage::from_read_at(
                Arc::new(CountingSource::new(bytes.clone())),
                limits,
                SourceCacheLimits::default(),
                SourceReadPolicy::exact(),
                None
            ),
            Err(OpcError::Limit {
                kind: SourceLimitKind::PartBytes,
                ..
            })
        ));
        let eager_error = match OpcPackage::from_bytes(&bytes, limits) {
            Ok(_) => panic!("eager package unexpectedly accepted a name conflict"),
            Err(error) => error,
        };
        if derived {
            assert!(matches!(eager_error, OpcError::DerivedPartNames { .. }));
        } else {
            assert!(matches!(eager_error, OpcError::EquivalentPartNames { .. }));
        }
    }
}

#[test]
fn source_backed_part_limit_precedes_later_max_parts_error() {
    let oversized = b"oversized first part";
    let bytes = archive_with_ordered_xml_parts(&[
        ("word/oversized.xml", oversized),
        ("word/later.xml", b"later"),
    ]);
    let limits = SourceLimits::builder()
        .max_parts(1)
        .unwrap()
        .max_part_bytes(oversized.len() as u64 - 1)
        .unwrap()
        .build()
        .unwrap();

    assert!(matches!(
        SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(bytes.clone())),
            limits,
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::PartBytes,
            ..
        })
    ));
    assert!(matches!(
        OpcPackage::from_bytes(&bytes, limits),
        Err(OpcError::Limit {
            kind: SourceLimitKind::Parts,
            ..
        })
    ));
}

#[test]
fn part_lookup_prefers_exact_names_and_matches_ascii_case_insensitive_targets() {
    let exact_name = PackURI::new("/word/document.xml").unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_with_part_names(
            "word/document.xml",
            &["word/document.xml"],
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    assert_eq!(package.part(&exact_name).unwrap().partname(), &exact_name);
    let case_variant = PackURI::new("/WORD/DOCUMENT.XML").unwrap();
    assert_eq!(package.part(&case_variant).unwrap().partname(), &exact_name);

    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_with_part_names(
            "WORD/DOCUMENT.XML",
            &["word/document.xml"],
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_eq!(
        package.main_document_part().unwrap().partname(),
        &exact_name
    );
}

#[test]
fn casefold_lookup_index_handles_unsorted_parts_without_changing_iteration_order() {
    let parts: [(&str, &[u8]); 5] = [
        ("zeta.xml", b"zeta"),
        ("ALPHA.xml", b"alpha"),
        ("middle.XML", b"middle"),
        ("beta.xml", b"beta"),
        ("Omega.xml", b"omega"),
    ];
    let bytes = archive_with_ordered_xml_parts(&parts);

    fn assert_lookup_behavior(package: &SourceBackedPackage) {
        assert!(package.catalog.casefold_order.is_none());
        let source_order = [
            "/zeta.xml",
            "/ALPHA.xml",
            "/middle.XML",
            "/beta.xml",
            "/Omega.xml",
        ];
        let iterated: Vec<_> = package
            .iter_parts()
            .map(|part| part.partname().as_str())
            .collect();
        assert_eq!(iterated, source_order);

        for (canonical, alias) in [
            ("/ALPHA.xml", "/alpha.XML"),
            ("/middle.XML", "/MIDDLE.xml"),
            ("/zeta.xml", "/ZETA.XML"),
        ] {
            let canonical = PackURI::new(canonical).unwrap();
            assert_eq!(package.part(&canonical).unwrap().partname(), &canonical);
            let alias = PackURI::new(alias).unwrap();
            assert_eq!(package.part(&alias).unwrap().partname(), &canonical);
        }

        let missing = PackURI::new("/delta.xml").unwrap();
        assert!(matches!(
            package.part(&missing),
            Err(OpcError::PartNotFound(name)) if name == "/delta.xml"
        ));
    }

    let normal = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_lookup_behavior(&normal);

    let validation = match SourceBackedPackage::from_read_at_for_validation(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
    ) {
        Ok(package) => package,
        Err(error) => panic!("validation open failed: {}", error.error),
    };
    assert_lookup_behavior(&validation);
}

fn assert_large_catalog_lookup_behavior(
    package: &SourceBackedPackage,
    expects_casefold_index: bool,
) {
    assert_eq!(
        package.catalog.casefold_order.is_some(),
        expects_casefold_index
    );
    let iterated: Vec<_> = package
        .iter_parts()
        .map(|part| part.partname().as_str())
        .collect();
    assert_eq!(iterated.len(), SOURCE_CASEFOLD_INDEX_MIN_PARTS);
    assert_eq!(iterated[0], "/part-2047.xml");
    assert_eq!(iterated[1024], "/part-1023.xml");
    assert_eq!(iterated[2047], "/part-0000.xml");

    for (canonical, alias) in [
        ("/part-0000.xml", "/PART-0000.XML"),
        ("/part-1024.xml", "/PART-1024.XML"),
        ("/part-2047.xml", "/PART-2047.XML"),
    ] {
        let canonical = PackURI::new(canonical).unwrap();
        assert_eq!(package.part(&canonical).unwrap().partname(), &canonical);
        let alias = PackURI::new(alias).unwrap();
        assert_eq!(package.part(&alias).unwrap().partname(), &canonical);
    }

    for missing in [
        PackURI::new("/part-2048.xml").unwrap(),
        PackURI::new("/part-1024.bin").unwrap(),
        PackURI::new("/part-0000.bin").unwrap(),
    ] {
        assert!(matches!(
            package.part(&missing),
            Err(OpcError::PartNotFound(name)) if name == missing.as_str()
        ));
    }
}

#[test]
fn casefold_lookup_index_is_retained_at_measured_large_catalog_threshold() {
    let bytes = archive_with_descending_xml_parts(SOURCE_CASEFOLD_INDEX_MIN_PARTS);

    let normal = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_large_catalog_lookup_behavior(&normal, true);

    let validation = match SourceBackedPackage::from_read_at_for_validation(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
    ) {
        Ok(package) => package,
        Err(error) => panic!("validation open failed: {}", error.error),
    };
    assert_large_catalog_lookup_behavior(&validation, true);
}

#[test]
fn managed_open_omits_casefold_lookup_index_and_keeps_linear_fallback() {
    let bytes = archive_with_descending_xml_parts(SOURCE_CASEFOLD_INDEX_MIN_PARTS);
    let (budget, _cancellation_source, context) = managed_context_with_cancellation(1);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();

    assert!(package.catalog.casefold_order.is_none());
    assert_eq!(budget.used(Resource::Memory), 0);

    assert_large_catalog_lookup_behavior(&package, false);
}

#[test]
fn source_catalog_still_rejects_case_equivalent_part_names() {
    let bytes = archive_with_part_names(
        "word/document.xml",
        &["word/document.xml", "WORD/DOCUMENT.XML"],
    );
    let Err(error) = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    ) else {
        panic!("case-equivalent parts must remain ambiguous");
    };
    assert!(matches!(error, OpcError::EquivalentPartNames { .. }));

    let validation_error = match SourceBackedPackage::from_read_at_for_validation(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
    ) {
        Ok(_) => panic!("case-equivalent parts must remain ambiguous"),
        Err(error) => error,
    };
    assert!(matches!(
        validation_error.error,
        OpcError::EquivalentPartNames { .. }
    ));
}

#[test]
fn cache_hits_pin_payloads_and_failures_are_not_retained() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"cached payload",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    let first = part.data().unwrap();
    let after_first = source.reads.load(Ordering::SeqCst);
    let second = part.data().unwrap();
    assert_eq!(source.reads.load(Ordering::SeqCst), after_first);
    assert!(first.shares_allocation_with(&second));
    assert_eq!(
        package.cache_diagnostics(),
        SourceCacheDiagnostics {
            hits: 1,
            cold_loads: 1,
            successful_loads: 1,
            retained_entries: 1,
            retained_bytes: b"cached payload".len(),
            ..SourceCacheDiagnostics::default()
        }
    );

    const DOCUMENT: &[u8] = b"never cache a failed read";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let position = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[position] ^= 0xff;
    let corrupt_source = Arc::new(CountingSource::new(bytes));
    let corrupt_package = SourceBackedPackage::from_read_at(
        corrupt_source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let corrupt_part = corrupt_package.main_document_part().unwrap();
    assert!(corrupt_part.data().is_err());
    let after_failure = corrupt_source.reads.load(Ordering::SeqCst);
    assert!(corrupt_part.data().is_err());
    assert!(corrupt_source.reads.load(Ordering::SeqCst) > after_failure);
    let diagnostics = corrupt_package.cache_diagnostics();
    assert_eq!(diagnostics.cold_loads, 2);
    assert_eq!(diagnostics.failed_loads, 2);
    assert_eq!(diagnostics.retained_entries, 0);
}

#[test]
fn managed_cache_reserves_declared_payload_and_releases_on_package_drop() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"managed payload",
        false,
    )));
    let (budget, context) = managed_context(1024);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    assert_eq!(budget.used(Resource::Memory), 0);

    let data = package.main_document_part().unwrap().data().unwrap();
    assert_eq!(data.as_bytes(), b"managed payload");
    assert_eq!(
        budget.used(Resource::Memory),
        b"managed payload".len() as u64
    );
    let diagnostics = package.cache_diagnostics();
    assert!(diagnostics.budget_managed);
    assert_eq!(diagnostics.budget_reservation_failures, 0);
    assert_eq!(
        diagnostics.budget_cache_reserved_bytes,
        b"managed payload".len() as u64
    );
    assert_eq!(
        diagnostics.budget_memory_used,
        b"managed payload".len() as u64
    );

    drop(data);
    // The clean cache entry owns the reservation until the package drops.
    assert_eq!(
        budget.used(Resource::Memory),
        b"managed payload".len() as u64
    );
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_input_bytes_are_exact_for_chunked_reads_and_cache_hits_are_free() {
    const DOCUMENT: &[u8] = b"short-read physical input accounting";
    let source = Arc::new(CountingSource::chunked(
        archive_bytes(root_relationships(), DOCUMENT, false),
        3,
    ));
    let (budget, _cancellation_source, context) =
        managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    assert_eq!(
        budget.used(Resource::InputBytes),
        source.read_bytes.load(Ordering::SeqCst)
    );
    let first = package.main_document_part().unwrap().data().unwrap();
    let after_cold = source.read_bytes.load(Ordering::SeqCst);
    assert_eq!(budget.used(Resource::InputBytes), after_cold);
    let second = package.main_document_part().unwrap().data().unwrap();
    assert!(first.shares_allocation_with(&second));
    assert_eq!(source.read_bytes.load(Ordering::SeqCst), after_cold);
    assert_eq!(
        package.cache_diagnostics().budget_input_bytes_used,
        after_cold
    );
    drop(second);
    drop(first);
    drop(package);
    // InputBytes and Work are cumulative; only retained resources release
    // when the package and its handles are dropped.
    assert_eq!(budget.used(Resource::InputBytes), after_cold);
}

#[test]
fn managed_input_budget_refusal_counts_only_the_terminal_read_reservation() {
    const DOCUMENT: &[u8] = b"input reservation refusal accounting";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let (budget, _cancellation_source, context) =
        managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context.clone()),
    )
    .unwrap();
    let input_before = budget.used(Resource::InputBytes);
    context
        .consume(Resource::InputBytes, u64::MAX - input_before)
        .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    let error = package.main_document_part().unwrap().data().unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::InputBytes
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(budget.used(Resource::InputBytes), u64::MAX);
    assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
    assert_eq!(package.cache_diagnostics().budget_reservation_failures, 1);
    assert_eq!(package.cache_diagnostics().retained_entries, 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    drop(package);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_work_one_under_refuses_payload_before_physical_io() {
    const DOCUMENT: &[u8] = b"work preflight one under";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let (budget, _cancellation_source, context) =
        managed_context_with_resources(4096, u64::MAX, u64::MAX, (DOCUMENT.len() - 1) as u64);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    let error = package.main_document_part().unwrap().data().unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Work
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(budget.used(Resource::Work), 0);
    assert_eq!(package.cache_diagnostics().retained_entries, 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_object_one_under_refuses_payload_before_physical_io() {
    const DOCUMENT: &[u8] = b"object preflight one under";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    // archive_bytes(false) contains four non-directory members and one
    // package-level catalog owner is retained by SourceBackedPackage.
    let (budget, _cancellation_source, context) =
        managed_context_with_resources(4096, u64::MAX, 5, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    let error = package.main_document_part().unwrap().data().unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Objects
    ));
    // The catalog reservation is retained by the package; the one-under
    // payload-object preflight happens before any ordinary payload read.
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(budget.used(Resource::Objects), 5);
    drop(package);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_failed_cold_load_consumes_work_and_input_but_releases_retained_objects() {
    const DOCUMENT: &[u8] = b"managed failed cold-load accounting";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let position = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[position] ^= 0xff;
    let source = Arc::new(CountingSource::new(bytes));
    let (budget, _cancellation_source, context) =
        managed_context_with_resources(4096, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.read_bytes.load(Ordering::SeqCst);
    assert!(matches!(
        package.main_document_part().unwrap().data(),
        Err(OpcError::ZipError(_))
    ));
    let reads_after = source.read_bytes.load(Ordering::SeqCst);
    assert!(reads_after > reads_before);
    assert_eq!(budget.used(Resource::InputBytes), reads_after);
    assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(package.cache_diagnostics().retained_entries, 0);
    assert_eq!(package.cache_diagnostics().budget_cache_reserved_objects, 0);
    drop(package);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_cancellation_between_cache_admission_and_source_read_is_typed() {
    const DOCUMENT: &[u8] = b"cancel before managed source read";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let (budget, cancellation_source, context) = managed_context_with_cancellation(4096);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    let index = part.index;
    let entry_id = package.catalog.parts[index].entry_id;
    let declared = package
        .archive
        .metadata_for(entry_id)
        .unwrap()
        .uncompressed_size();
    let reads_before = source.reads.load(Ordering::SeqCst);
    let flight = match package.cache.enter(entry_id, declared).unwrap() {
        CacheAccess::Loader(flight) => flight,
        CacheAccess::Hit(_) | CacheAccess::Waiter(_) | CacheAccess::Bypass(_) => {
            panic!("fresh managed Part must become the loader")
        },
    };
    // The loader has charged its bounded cold-load resources, but has not
    // entered SourceReader yet. This is the exact race boundary where a
    // cancellation must survive ZIP's std::io::Error conversion.
    cancellation_source.cancel();
    let mut observer = NoopDiagnosticObserver;
    let error = package
        .load_part_with_accounting(
            index,
            entry_id,
            Some(declared),
            Some(flight),
            None,
            None,
            None,
            &mut observer,
            None,
        )
        .unwrap_err();
    assert!(matches!(error, OpcError::Cancelled));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    assert_eq!(package.cache_diagnostics().retained_entries, 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(package.cache_diagnostics().budget_cache_reserved_objects, 0);
    drop(package);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_cancellation_before_preservation_source_read_is_typed() {
    const DOCUMENT: &[u8] = b"<before/>";
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let source = Arc::new(CancelOnHitVersionSource::new(
        archive_bytes(root_relationships(), DOCUMENT, false),
        cancellation_source,
    ));
    let budget = Budget::root(
        "opc-source-cache-preservation-cancel-test",
        Limits::new(4096, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(4096).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let target = package.main_document_part().unwrap().partname().clone();
    // One observation for the part lookup, then two more inside the
    // cold load, before the pre-read check this test targets.
    source.arm_after_versions(3);
    let mut output = Vec::new();
    let error = package
        .write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec())
        .unwrap_err();
    assert!(matches!(error, OpcError::Cancelled));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_constructor_honors_pre_cancellation_without_source_reads() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"cancelled before open",
        false,
    )));
    let reads_before = source.reads.load(Ordering::SeqCst);
    let budget = Budget::root(
        "opc-source-cache-cancel-test",
        Limits::new(1024, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    cancellation_source.cancel();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(1024).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget, cancellation, execution_limits);

    assert!(matches!(
        SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            Some(context)
        ),
        Err(OpcError::Cancelled)
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
}

#[test]
fn managed_cache_hit_honors_cancellation_without_releasing_cached_budget() {
    const DOCUMENT: &[u8] = b"managed cancellation hit";
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let source = Arc::new(CancelOnHitVersionSource::new(
        archive_bytes(root_relationships(), DOCUMENT, false),
        cancellation_source,
    ));
    let budget = Budget::root(
        "opc-source-cache-hit-cancel-test",
        Limits::new(
            DOCUMENT.len() as u64,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, execution_limits);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let first = package.main_document_part().unwrap().data().unwrap();
    assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
    // The part lookup takes one observation and the warm read takes
    // one more, which is the hit's post-entry check this test targets.
    source.arm_after_versions(2);

    assert!(matches!(
        package.main_document_part().unwrap().data(),
        Err(OpcError::Cancelled)
    ));
    assert_eq!(package.cache_diagnostics().hits, 1);
    // Cancellation rejects the handle request, but never steals the
    // clean entry's reservation from the package-owned cache.
    assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
    drop(first);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

fn scheduled_payload(seed: usize, bytes: usize) -> Vec<u8> {
    let mut data = Vec::with_capacity(bytes + 32);
    data.extend_from_slice(b"<p>");
    let mut state = (seed as u64).wrapping_mul(0x9e37_79b9_7f4a_7c15).max(1);
    while data.len() < bytes {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        data.extend_from_slice(format!("<c r=\"{}\"/>", state % 8192).as_bytes());
    }
    data.extend_from_slice(b"</p>");
    data
}

/// A package whose ordinary Parts are deflated, so a publication that
/// replaces them regenerates deflate members rather than stored ones.
fn deflated_package_bytes(parts: usize, bytes: usize) -> Vec<u8> {
    let mut writer = soapberry_zip::office::StreamingArchiveWriter::new(
        soapberry_zip::office::StreamingArchiveLimits::default(),
    );
    writer
            .write_stored(
                "[Content_Types].xml",
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#,
            )
            .unwrap();
    writer
        .write_stored("_rels/.rels", root_relationships())
        .unwrap();
    writer
        .write_deflated("word/document.xml", &scheduled_payload(1, bytes))
        .unwrap();
    for index in 0..parts {
        writer
            .write_deflated(
                &format!("custom/part-{index:04}.xml"),
                &scheduled_payload(index + 2, bytes),
            )
            .unwrap();
    }
    writer.finish_to_bytes().unwrap()
}

fn scheduled_context(workers: usize) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "opc-scheduled-publication-test",
        Limits::new(
            256 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(workers).unwrap(),
        NonZeroUsize::new(workers.max(64)).unwrap(),
        NonZeroU64::new(64 * 1024 * 1024).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, limits);
    (budget, cancellation_source, context)
}

fn ordered_read_context(
    workers: usize,
    io_concurrency: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "opc-ordered-read-test",
        Limits::new(
            256 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        )
        .with_execution_io(workers as u64, io_concurrency, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(workers).unwrap(),
        NonZeroUsize::new(workers.max(4)).unwrap(),
        NonZeroU64::new(64 * 1024 * 1024).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, limits);
    (budget, cancellation_source, context)
}

fn scheduled_replacements(parts: usize, bytes: usize) -> Vec<(PackURI, Arc<Vec<u8>>)> {
    (0..parts)
        .map(|index| {
            (
                PackURI::new(format!("/custom/part-{index:04}.xml")).unwrap(),
                Arc::new(scheduled_payload(index + 100, bytes)),
            )
        })
        .collect()
}

/// A sink that records the worker permits the publication held when the
/// first output byte was accepted — every member is compressed before
/// that byte, so this is the width the wave actually ran at.
struct PermitProbeSink {
    inner: Vec<u8>,
    budget: Budget,
    observed: Arc<AtomicU64>,
    seen: bool,
}

impl Write for PermitProbeSink {
    fn write(&mut self, data: &[u8]) -> std::io::Result<usize> {
        if !self.seen {
            self.seen = true;
            self.observed
                .store(self.budget.used(Resource::Workers), Ordering::Relaxed);
        }
        self.inner.write(data)
    }

    fn flush(&mut self) -> std::io::Result<()> {
        self.inner.flush()
    }
}

fn publish_overlays_observing_permits(
    source: &Arc<dyn ReadAt>,
    budget: &Budget,
    context: ExecutionContext,
    replacements: Vec<(PackURI, Arc<Vec<u8>>)>,
) -> (Vec<u8>, u64) {
    let package = SourceBackedPackage::from_read_at(
        Arc::clone(source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let observed = Arc::new(AtomicU64::new(0));
    let sink = PermitProbeSink {
        inner: Vec::new(),
        budget: budget.clone(),
        observed: Arc::clone(&observed),
        seen: false,
    };
    let mut sink = sink;
    package
        .write_part_overlays_shared_to_stream(&mut sink, replacements)
        .unwrap();
    (sink.inner, observed.load(Ordering::Relaxed))
}

fn publish_overlays(
    source: &Arc<dyn ReadAt>,
    context: Option<ExecutionContext>,
    replacements: Vec<(PackURI, Arc<Vec<u8>>)>,
) -> Vec<u8> {
    let package = match context {
        Some(context) => SourceBackedPackage::from_read_at(
            Arc::clone(source),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            Some(context),
        )
        .unwrap(),
        None => SourceBackedPackage::from_read_at(
            Arc::clone(source),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap(),
    };
    let mut output = Vec::new();
    package
        .write_part_overlays_shared_to_stream(&mut output, replacements)
        .unwrap();
    output
}

#[test]
fn a_scheduled_publication_is_byte_identical_at_every_wave_width() {
    let source: Arc<dyn ReadAt> =
        Arc::new(CountingSource::new(deflated_package_bytes(15, 6 * 1024)));
    let replacements = scheduled_replacements(15, 6 * 1024);
    let sequential = publish_overlays(&source, None, replacements.clone());

    for workers in [1_usize, 2, 4, 8] {
        let (budget, _cancellation, context) = scheduled_context(workers);
        let (scheduled, permits) =
            publish_overlays_observing_permits(&source, &budget, context, replacements.clone());
        assert_eq!(scheduled, sequential, "width {workers} changed the bytes");
        // A width of one takes the sequential path and holds no permits;
        // every wider wave held exactly the permits its width asked for.
        assert_eq!(
            permits,
            if workers > 1 { workers as u64 } else { 0 },
            "width {workers} ran at the wrong width",
        );
        // One CPU task per regenerated deflate member, at every width.
        assert_eq!(budget.used(Resource::CpuTasks), 15, "width {workers}");
        // Worker permits and the memory their compressor states retain are
        // released with the package that held them.
        assert_eq!(budget.used(Resource::Workers), 0, "width {workers}");
        assert_eq!(budget.used(Resource::Memory), 0, "width {workers}");
    }
}

#[derive(Debug, Default)]
struct CountingCallerWorkers {
    waves: AtomicU64,
    tasks: AtomicU64,
}

impl litchi_core::ScopedWorkers for CountingCallerWorkers {
    fn run_all<'task>(&self, tasks: &mut [&mut (dyn FnMut() + Send + 'task)]) {
        self.waves.fetch_add(1, Ordering::Relaxed);
        self.tasks.fetch_add(tasks.len() as u64, Ordering::Relaxed);
        for task in tasks.iter_mut() {
            task();
        }
    }
}

#[test]
fn a_caller_facility_runs_every_task_and_keeps_the_bytes() {
    let source: Arc<dyn ReadAt> =
        Arc::new(CountingSource::new(deflated_package_bytes(15, 6 * 1024)));
    let replacements = scheduled_replacements(15, 6 * 1024);
    let sequential = publish_overlays(&source, None, replacements.clone());

    let (budget, _cancellation, context) = scheduled_context(4);
    let workers: Arc<CountingCallerWorkers> = Arc::new(CountingCallerWorkers::default());
    let facility: Arc<dyn litchi_core::ScopedWorkers> = workers.clone();
    let context = context.with_scoped_workers(facility);
    let scheduled = publish_overlays(&source, Some(context), replacements);

    assert_eq!(scheduled, sequential);
    assert_eq!(workers.waves.load(Ordering::Relaxed), 1);
    assert_eq!(workers.tasks.load(Ordering::Relaxed), 15);
    assert_eq!(budget.used(Resource::CpuTasks), 15);
    assert_eq!(budget.used(Resource::Workers), 0);
}

#[test]
fn ordered_source_reads_use_caller_facility_and_preserve_order() {
    let source: Arc<dyn ReadAt> = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"ordered document",
        false,
    )));
    let (budget, _cancellation, context) = ordered_read_context(2, 2);
    let workers: Arc<CountingCallerWorkers> = Arc::new(CountingCallerWorkers::default());
    let context = context.with_scoped_workers(workers.clone());
    let package = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let document = PackURI::new("/word/document.xml").unwrap();
    let orphan = PackURI::new("/custom/orphan.xml").unwrap();
    let batch = package.read_parts_ordered(&[orphan, document]).unwrap();

    assert_eq!(batch.len(), 2);
    assert_eq!(batch.get(0).unwrap().as_bytes(), b"<orphan/>");
    assert_eq!(batch.get(1).unwrap().as_bytes(), b"ordered document");
    assert!(workers.tasks.load(Ordering::Relaxed) >= 2);
    assert_eq!(budget.used(Resource::Workers), 0);
    assert_eq!(budget.used(Resource::IoConcurrency), 0);
}

#[test]
fn ordered_source_reads_refuse_zero_io_before_payload_reads() {
    let source_probe = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"ordered document",
        false,
    )));
    let source: Arc<dyn ReadAt> = source_probe.clone();
    let (budget, _cancellation, context) = ordered_read_context(2, 0);
    let package = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let baseline = source_probe.reads.load(Ordering::SeqCst);
    let document = PackURI::new("/word/document.xml").unwrap();
    let orphan = PackURI::new("/custom/orphan.xml").unwrap();
    assert!(matches!(
        package.read_parts_ordered(&[document, orphan]),
        Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::IoConcurrency
    ));
    assert_eq!(source_probe.reads.load(Ordering::SeqCst), baseline);
    assert_eq!(budget.used(Resource::Workers), 0);
    assert_eq!(budget.used(Resource::IoConcurrency), 0);
    assert_eq!(budget.used(Resource::CpuTasks), 0);
}

#[test]
fn ordered_serial_read_failure_releases_worker_and_io_admissions() {
    const DOCUMENT: &[u8] = b"ordered failure";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let position = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[position] ^= 0xff;
    let source: Arc<dyn ReadAt> = Arc::new(CountingSource::new(bytes));
    let (budget, _cancellation, context) = ordered_read_context(1, 1);
    let package = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let document = PackURI::new("/word/document.xml").unwrap();

    assert!(package.read_parts_ordered(&[document]).is_err());
    assert_eq!(budget.used(Resource::Workers), 0);
    assert_eq!(budget.used(Resource::IoConcurrency), 0);
}

#[test]
fn a_serial_scheduled_publication_reserves_no_worker_permits() {
    let source: Arc<dyn ReadAt> =
        Arc::new(CountingSource::new(deflated_package_bytes(15, 6 * 1024)));
    let budget = Budget::root(
        "opc-scheduled-serial-test",
        Limits::new(
            256 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        )
        .with_execution(0, u64::MAX),
    );
    let (_source_handle, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(8).unwrap(),
        NonZeroUsize::new(64).unwrap(),
        NonZeroU64::new(64 * 1024 * 1024).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, limits);
    // A budget that grants no worker permits narrows the wave to one and
    // publishes the sequential bytes rather than refusing.
    let replacements = scheduled_replacements(15, 6 * 1024);
    let sequential = publish_overlays(&source, None, replacements.clone());
    let narrowed = publish_overlays(&source, Some(context), replacements);
    assert_eq!(narrowed, sequential);
    assert_eq!(budget.used(Resource::Workers), 0);
}

#[test]
fn an_exhausted_cpu_task_budget_refuses_before_any_output_byte() {
    let source: Arc<dyn ReadAt> =
        Arc::new(CountingSource::new(deflated_package_bytes(15, 6 * 1024)));
    let budget = Budget::root(
        "opc-scheduled-task-budget-test",
        Limits::new(
            256 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        )
        .with_execution(u64::MAX, 4),
    );
    let (_source_handle, cancellation) = CancellationSource::pair();
    let limits = ExecutionLimits::new(
        NonZeroUsize::new(4).unwrap(),
        NonZeroUsize::new(64).unwrap(),
        NonZeroU64::new(64 * 1024 * 1024).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(budget.clone(), cancellation, limits);
    let package = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut output = b"untouched".to_vec();
    let error = package
        .write_part_overlays_shared_to_stream(&mut output, scheduled_replacements(15, 6 * 1024))
        .expect_err("fifteen tasks do not fit a four-task budget");
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(ref limit))
            if limit.resource == Resource::CpuTasks
    ));
    assert_eq!(output, b"untouched");
}

#[test]
fn a_cancelled_scheduled_publication_publishes_nothing() {
    let source: Arc<dyn ReadAt> =
        Arc::new(CountingSource::new(deflated_package_bytes(15, 6 * 1024)));
    let (_budget, cancellation_source, context) = scheduled_context(4);
    let package = SourceBackedPackage::from_read_at(
        Arc::clone(&source),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    cancellation_source.cancel();
    let mut output = b"untouched".to_vec();
    let error = package
        .write_part_overlays_shared_to_stream(&mut output, scheduled_replacements(15, 6 * 1024))
        .expect_err("a cancelled publication refuses");
    assert!(matches!(error, OpcError::Cancelled));
    assert_eq!(output, b"untouched");
}

#[test]
fn managed_cache_eviction_releases_unpinned_reservation_before_next_read() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"document",
        false,
    )));
    let (budget, context) = managed_context(64);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::new(9, 2).unwrap(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let first_name = package.catalog.parts[0].partname.clone();
    let second_name = package.catalog.parts[1].partname.clone();
    let first = package.part(&first_name).unwrap().data().unwrap();
    assert_eq!(budget.used(Resource::Memory), b"document".len() as u64);
    drop(first);

    let second = package.part(&second_name).unwrap().data().unwrap();
    assert_eq!(second.as_bytes(), b"<orphan/>");
    assert_eq!(budget.used(Resource::Memory), b"<orphan/>".len() as u64);
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.evictions, 1);
    assert_eq!(diagnostics.retained_bytes, b"<orphan/>".len());
    assert_eq!(
        diagnostics.budget_cache_reserved_bytes,
        b"<orphan/>".len() as u64
    );
    drop(second);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cache_does_not_evict_externally_pinned_entry_and_bypasses_retention() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"document",
        false,
    )));
    let (budget, context) = managed_context(64);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::new(b"document".len(), 1).unwrap(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let first_name = package.catalog.parts[0].partname.clone();
    let second_name = package.catalog.parts[1].partname.clone();
    let first = package.part(&first_name).unwrap().data().unwrap();
    let first_error = first.into_arc().expect_err("managed Arc escape must fail");
    assert_eq!(budget.used(Resource::Memory), b"document".len() as u64);
    assert_eq!(
        package.cache_diagnostics().retained_bytes,
        b"document".len()
    );
    let second = package.part(&second_name).unwrap().data().unwrap();

    assert!(matches!(first_error, OpcError::ManagedPartDataArcEscape));
    assert_eq!(first.as_bytes(), b"document");
    assert_eq!(second.as_bytes(), b"<orphan/>");
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.retained_entries, 1);
    assert_eq!(diagnostics.retained_bytes, b"document".len());
    assert_eq!(diagnostics.bypasses, 1);
    assert_eq!(
        budget.used(Resource::Memory),
        (b"document".len() + b"<orphan/>".len()) as u64
    );
    assert_eq!(budget.used(Resource::Objects), 7);
    assert_eq!(diagnostics.budget_cache_reserved_objects, 1);
    assert!(
        package
            .cache
            .state
            .lock()
            .unwrap()
            .entries
            .contains_key(&package.catalog.parts[0].entry_id)
    );

    drop(second);
    drop(first);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn managed_budget_rejects_before_payload_io_and_reports_content_free_failure() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"payload too large",
        false,
    )));
    let (budget, context) = managed_context((b"payload too large".len() - 1) as u64);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    let error = package.main_document_part().unwrap().data().unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::Memory
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(budget.used(Resource::Memory), 0);
    let diagnostics = package.cache_diagnostics();
    assert!(diagnostics.budget_managed);
    assert_eq!(diagnostics.budget_reservation_failures, 2);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.retained_bytes, 0);
}

#[test]
fn managed_cache_respects_hierarchical_parent_memory_limit() {
    const DOCUMENT: &[u8] = b"hierarchical budget payload";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let root = Budget::root(
        "opc-source-cache-root",
        Limits::new(
            (DOCUMENT.len() - 1) as u64,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let child = root.child(
        "opc-source-cache-child",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    drop(cancellation_source);
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(child, cancellation, execution_limits);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);

    assert!(matches!(
        package.main_document_part().unwrap().data(),
        Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::Memory
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(root.used(Resource::Memory), 0);
    assert_eq!(package.cache_diagnostics().budget_reservation_failures, 2);
}

#[test]
fn managed_sibling_caches_compete_for_parent_memory_before_payload_io() {
    const DOCUMENT: &[u8] = b"sibling parent budget payload";
    let root = Budget::root(
        "opc-source-cache-sibling-root",
        Limits::new(
            DOCUMENT.len() as u64,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let first_budget = root.child(
        "opc-source-cache-sibling-first",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let second_budget = root.child(
        "opc-source-cache-sibling-second",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(DOCUMENT.len() as u64).unwrap(),
        0,
    )
    .unwrap();
    let (first_cancellation_source, first_cancellation) = CancellationSource::pair();
    let (second_cancellation_source, second_cancellation) = CancellationSource::pair();
    let first_context = ExecutionContext::new(first_budget, first_cancellation, execution_limits);
    let second_context =
        ExecutionContext::new(second_budget, second_cancellation, execution_limits);
    drop(first_cancellation_source);
    drop(second_cancellation_source);
    let first_source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let second_source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let first_package = SourceBackedPackage::from_read_at(
        first_source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(first_context),
    )
    .unwrap();
    let second_package = SourceBackedPackage::from_read_at(
        second_source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(second_context),
    )
    .unwrap();
    let reads_before = second_source.reads.load(Ordering::SeqCst);
    let first = first_package.main_document_part().unwrap().data().unwrap();
    assert_eq!(root.used(Resource::Memory), DOCUMENT.len() as u64);
    drop(first);

    assert!(matches!(
        second_package.main_document_part().unwrap().data(),
        Err(OpcError::Execution(ExecutionError::ResourceLimit(limit)))
            if limit.resource == Resource::Memory
    ));
    assert_eq!(second_source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(root.used(Resource::Memory), DOCUMENT.len() as u64);
    assert_eq!(
        second_package
            .cache_diagnostics()
            .budget_reservation_failures,
        2
    );
    drop(second_package);
    drop(first_package);
    assert_eq!(root.used(Resource::Memory), 0);
}

#[test]
fn managed_same_part_waiters_share_one_reservation_and_flight() {
    const DOCUMENT: &[u8] = b"managed single-flight payload";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
    let (budget, context) = managed_context(DOCUMENT.len() as u64);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let start = Arc::new(Barrier::new(3));
    let (first, second) = std::thread::scope(|scope| {
        let package = &package;
        let first_start = Arc::clone(&start);
        let first_task = scope.spawn(move || {
            first_start.wait();
            package.main_document_part().unwrap().data().unwrap()
        });
        let second_start = Arc::clone(&start);
        let second_task = scope.spawn(move || {
            second_start.wait();
            package.main_document_part().unwrap().data().unwrap()
        });
        start.wait();
        std::thread::sleep(Duration::from_millis(10));
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.in_flight_loads, 1);
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            DOCUMENT.len() as u64
        );
        assert_eq!(diagnostics.budget_cache_reserved_objects, 2);
        assert_eq!(budget.used(Resource::Objects), 7);
        (first_task.join().unwrap(), second_task.join().unwrap())
    });
    assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
    assert!(first.shares_allocation_with(&second));
    assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.cold_loads, 1);
    assert_eq!(diagnostics.waiter_joins, 1);
    assert_eq!(diagnostics.successful_loads, 1);
    assert_eq!(
        diagnostics.budget_cache_reserved_bytes,
        DOCUMENT.len() as u64
    );
    assert_eq!(diagnostics.budget_cache_reserved_objects, 1);
    drop(first);
    drop(second);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_waiter_cancellation_does_not_block_or_publish_loader_payload() {
    const DOCUMENT: &[u8] = b"managed waiter cancellation payload";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
    let (budget, cancellation_source, context) =
        managed_context_with_cancellation(DOCUMENT.len() as u64);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let start = Arc::new(Barrier::new(3));
    let (first, second) = std::thread::scope(|scope| {
        let package = &package;
        let first_start = Arc::clone(&start);
        let first_task = scope.spawn(move || {
            first_start.wait();
            package.main_document_part().unwrap().data()
        });
        let second_start = Arc::clone(&start);
        let second_task = scope.spawn(move || {
            second_start.wait();
            package.main_document_part().unwrap().data()
        });
        start.wait();
        std::thread::sleep(Duration::from_millis(10));
        assert_eq!(package.cache_diagnostics().in_flight_loads, 1);
        cancellation_source.cancel();
        (first_task.join().unwrap(), second_task.join().unwrap())
    });

    assert!(matches!(first, Err(OpcError::Cancelled)));
    assert!(matches!(second, Err(OpcError::Cancelled)));
    assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.failed_loads, 1);
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
    assert!(budget.used(Resource::InputBytes) > 0);
    assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_source_change_drops_reservation_and_does_not_retain_payload() {
    const DOCUMENT: &[u8] = b"source changes during managed read";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let source = Arc::new(ChangeDuringPayloadSource::new(bytes, payload_offset));
    let (budget, context) = managed_context(1024);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    source.armed.store(true, Ordering::SeqCst);

    assert!(matches!(
        package.main_document_part().unwrap().data(),
        Err(OpcError::SourceChanged { .. })
    ));
    assert_eq!(budget.used(Resource::Memory), 0);
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.failed_loads, 1);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.budget_cache_reserved_bytes, 0);
    assert_eq!(budget.used(Resource::Work), DOCUMENT.len() as u64);
    assert!(budget.used(Resource::InputBytes) > 0);
    assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
}

#[test]
fn source_change_takes_precedence_over_archive_failure_and_releases_flight() {
    const DOCUMENT: &[u8] = b"source changes while a corrupt payload is read";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[payload_offset] ^= 0xff;
    let source = Arc::new(ChangeDuringPayloadSource::new(bytes, payload_offset));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    source.armed.store(true, Ordering::SeqCst);

    assert!(matches!(
        package.main_document_part().unwrap().data(),
        Err(OpcError::SourceChanged { .. })
    ));
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.failed_loads, 1);
}

#[test]
fn cancellation_takes_precedence_over_accounted_archive_failure_and_releases_flight() {
    const DOCUMENT: &[u8] = b"cancellation races with a corrupt accounted payload";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[payload_offset] ^= 0xff;
    let (budget, cancellation_source, context) =
        managed_context_with_cancellation(DOCUMENT.len() as u64);
    let source = Arc::new(CancelOnHitVersionSource::new(bytes, cancellation_source));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    // One observation for the part lookup, then two more inside the
    // cold load, before the pre-read check this test targets.
    source.arm_after_versions(3);
    let mut accounting = OpcOperationAccounting::default();

    assert!(matches!(
        package
            .main_document_part()
            .unwrap()
            .data_with_accounting(&mut accounting),
        Err(OpcError::Cancelled)
    ));
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.failed_loads, 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_cancellation_after_decompression_prevents_publication() {
    const DOCUMENT: &[u8] = b"managed prepublication cancellation payload";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let (budget, cancellation_source, context) =
        managed_context_with_cancellation(DOCUMENT.len() as u64);
    let source = Arc::new(CancelDuringPayloadSource::new(
        bytes,
        payload_offset,
        cancellation_source,
    ));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();

    assert!(matches!(
        package.main_document_part().unwrap().data(),
        Err(OpcError::Cancelled)
    ));
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.failed_loads, 1);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn managed_open_checks_cancellation_after_indexing_before_catalog_publication() {
    let (budget, cancellation_source, context) = managed_context_with_cancellation(u64::MAX);
    let source = Arc::new(CancelDuringOpenSource::new(
        archive_bytes(root_relationships(), b"open cancellation", false),
        cancellation_source,
    ));
    assert!(matches!(
        SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            Some(context)
        ),
        Err(OpcError::Cancelled)
    ));
    assert!(source.reads.load(Ordering::SeqCst) > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn concurrent_cold_reads_share_one_archive_load_and_one_arc() {
    const DOCUMENT: &[u8] = b"single-flight source-backed payload";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let source = Arc::new(SlowPayloadSource::new(bytes, payload_offset));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let start = Arc::new(Barrier::new(3));
    let (first, second) = std::thread::scope(|scope| {
        let package = &package;
        let first_start = Arc::clone(&start);
        let first_task = scope.spawn(move || {
            first_start.wait();
            package
                .main_document_part()
                .unwrap()
                .data()
                .unwrap()
                .into_arc()
                .unwrap()
        });
        let second_start = Arc::clone(&start);
        let second_task = scope.spawn(move || {
            second_start.wait();
            package
                .main_document_part()
                .unwrap()
                .data()
                .unwrap()
                .into_arc()
                .unwrap()
        });
        start.wait();
        std::thread::sleep(Duration::from_millis(10));
        assert_eq!(package.cache_diagnostics().in_flight_loads, 1);
        (first_task.join().unwrap(), second_task.join().unwrap())
    });
    assert_eq!(source.payload_reads.load(Ordering::SeqCst), 1);
    assert!(Arc::ptr_eq(&first, &second));
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.cold_loads, 1);
    assert_eq!(diagnostics.waiter_joins, 1);
    assert_eq!(diagnostics.successful_loads, 1);
    assert_eq!(diagnostics.in_flight_loads, 0);
}

#[test]
fn provisional_publication_source_change_rolls_back_everything() {
    const DOCUMENT: &[u8] = b"source changed after provisional publication";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let (budget, context) = managed_context(4096);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let hook = Arc::new(TestPublicationHook::new());
    let _hook_guard = package
        .cache
        .install_test_publication_hook(Arc::clone(&hook));

    let result = std::thread::scope(|scope| {
        let package = &package;
        let task = scope.spawn(move || package.main_document_part().unwrap().data());
        hook.wait_until_entered();
        {
            let state = package.cache.state.lock().unwrap();
            assert_eq!(state.entries.len(), 1);
            assert_eq!(state.pending.len(), 1);
            assert_eq!(state.flights.len(), 1);
        }
        source.changed();
        hook.release();
        task.join().unwrap()
    });

    assert!(matches!(result, Err(OpcError::SourceChanged { .. })));
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.retained_bytes, 0);
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert_eq!(diagnostics.failed_loads, 1);
    assert_eq!(diagnostics.budget_cache_reserved_bytes, 0);
    assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    let state = package.cache.state.lock().unwrap();
    assert!(state.entries.is_empty());
    assert!(state.pending.is_empty());
    assert!(state.flights.is_empty());
}

#[test]
fn provisional_publication_waiter_joins_and_stable_hit_reuses_allocation() {
    const DOCUMENT: &[u8] = b"waiter joins provisional publication";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let hook = Arc::new(TestPublicationHook::new());
    let _hook_guard = package
        .cache
        .install_test_publication_hook(Arc::clone(&hook));

    let (first, second) = std::thread::scope(|scope| {
        let package = &package;
        let first_task = scope.spawn(move || package.main_document_part().unwrap().data());
        hook.wait_until_entered();
        let second_task = scope.spawn(move || package.main_document_part().unwrap().data());
        hook.wait_until_waiter_joined();
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.cold_loads, 1);
        assert_eq!(diagnostics.waiter_joins, 1);
        assert_eq!(diagnostics.hits, 0);
        assert_eq!(diagnostics.in_flight_loads, 1);
        hook.release();
        (
            first_task.join().unwrap().unwrap(),
            second_task.join().unwrap().unwrap(),
        )
    });

    assert_eq!(first.as_bytes(), DOCUMENT);
    assert_eq!(second.as_bytes(), DOCUMENT);
    assert!(first.shares_allocation_with(&second));
    let hit = package.main_document_part().unwrap().data().unwrap();
    assert!(first.shares_allocation_with(&hit));
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.cold_loads, 1);
    assert_eq!(diagnostics.waiter_joins, 1);
    assert_eq!(diagnostics.hits, 1);
    assert_eq!(diagnostics.successful_loads, 1);
    assert_eq!(diagnostics.retained_entries, 1);
    assert_eq!(diagnostics.in_flight_loads, 0);
}

#[test]
fn combined_capture_and_ordinary_read_share_the_exact_loader_payload() {
    const DOCUMENT: &[u8] = b"shared combined-capture loader payload";
    for first_combined in [false, true] {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let (budget, context) = managed_context(65536);
        let package = SourceBackedPackage::from_read_at(
            source,
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            Some(context),
        )
        .unwrap();
        let hook = Arc::new(TestPublicationHook::new());
        let hook_guard = package
            .cache
            .install_test_publication_hook(Arc::clone(&hook));
        let work_before = budget.used(Resource::Work);
        let read = |combined: bool| {
            let view = package.main_document_part().unwrap();
            if combined {
                view.data_and_authorize_precompressed()
                    .map(|(data, token)| (data, Some(token)))
            } else {
                view.data().map(|data| (data, None))
            }
        };
        let (first, second) = std::thread::scope(|scope| {
            let first_task = scope.spawn(|| read(first_combined));
            hook.wait_until_entered();
            let second_task = scope.spawn(|| read(!first_combined));
            hook.wait_until_waiter_joined();
            let diagnostics = package.cache_diagnostics();
            assert_eq!(diagnostics.cold_loads, 1);
            assert_eq!(diagnostics.waiter_joins, 1);
            assert_eq!(diagnostics.in_flight_loads, 1);
            hook.release();
            (
                first_task.join().unwrap().unwrap(),
                second_task.join().unwrap().unwrap(),
            )
        });
        assert_eq!(first.0.as_bytes(), DOCUMENT);
        assert!(first.0.shares_allocation_with(&second.0));
        let token = first.1.as_ref().or(second.1.as_ref()).unwrap();
        assert!(Arc::ptr_eq(&token.expected_decoded, &first.0.payload.bytes));
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.cold_loads, 1);
        assert_eq!(diagnostics.successful_loads, 1);
        assert_eq!(diagnostics.in_flight_loads, 0);
        let passes = if first_combined { 2 } else { 3 };
        assert_eq!(
            budget.used(Resource::Work) - work_before,
            DOCUMENT.len() as u64 * passes
        );
        drop(first);
        drop(second);
        drop(hook_guard);
        drop(package);
        assert_eq!(budget.used(Resource::Memory), 0);
        assert_eq!(budget.used(Resource::Objects), 0);
    }
}

#[test]
fn combined_capture_provisional_failure_releases_cache_and_capture_reservations() {
    const DOCUMENT: &[u8] = b"combined capture rolls back before token publication";
    for change_source in [false, true] {
        let source = Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            DOCUMENT,
            false,
        )));
        let (budget, cancellation, context) = managed_context_with_cancellation(65536);
        let package = SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            Some(context),
        )
        .unwrap();
        let hook = Arc::new(TestPublicationHook::new());
        let _hook_guard = package
            .cache
            .install_test_publication_hook(Arc::clone(&hook));
        let result = std::thread::scope(|scope| {
            let task = scope.spawn(|| {
                package
                    .main_document_part()
                    .unwrap()
                    .data_and_authorize_precompressed()
            });
            hook.wait_until_entered();
            assert_eq!(package.cache_diagnostics().retained_entries, 1);
            assert!(budget.used(Resource::Memory) > DOCUMENT.len() as u64);
            if change_source {
                source.changed();
            } else {
                cancellation.cancel();
            }
            hook.release();
            task.join().unwrap()
        });
        if change_source {
            assert!(matches!(result, Err(OpcError::SourceChanged { .. })));
        } else {
            assert!(matches!(result, Err(OpcError::Cancelled)));
        }
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.retained_entries, 0);
        assert_eq!(diagnostics.in_flight_loads, 0);
        assert_eq!(diagnostics.failed_loads, 1);
        assert_eq!(diagnostics.budget_cache_reserved_bytes, 0);
        assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

#[test]
fn provisional_publication_cancellation_rolls_back_everything() {
    const DOCUMENT: &[u8] = b"cancelled after provisional publication";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let (budget, cancellation_source, context) = managed_context_with_cancellation(4096);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let hook = Arc::new(TestPublicationHook::new());
    let _hook_guard = package
        .cache
        .install_test_publication_hook(Arc::clone(&hook));

    let result = std::thread::scope(|scope| {
        let package = &package;
        let task = scope.spawn(move || package.main_document_part().unwrap().data());
        hook.wait_until_entered();
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.retained_entries, 1);
        assert_eq!(diagnostics.in_flight_loads, 1);
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            DOCUMENT.len() as u64
        );
        assert_eq!(diagnostics.budget_cache_reserved_objects, 2);
        assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
        {
            let state = package.cache.state.lock().unwrap();
            assert_eq!(state.entries.len(), 1);
            assert_eq!(state.pending.len(), 1);
            assert_eq!(state.flights.len(), 1);
        }
        cancellation_source.cancel();
        hook.release();
        task.join().unwrap()
    });

    assert!(matches!(result, Err(OpcError::Cancelled)));
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.retained_bytes, 0);
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert_eq!(diagnostics.failed_loads, 1);
    assert_eq!(diagnostics.budget_cache_reserved_bytes, 0);
    assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    let state = package.cache.state.lock().unwrap();
    assert!(state.entries.is_empty());
    assert!(state.pending.is_empty());
    assert!(state.flights.is_empty());
}

#[test]
fn managed_failed_loader_releases_reservations_before_waiter_retry() {
    const DOCUMENT: &[u8] = b"stable corrupt payload for a managed retry";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[payload_offset] ^= 0xff;
    let source = Arc::new(GatedCorruptSource::new(bytes, payload_offset));
    let (budget, context) = managed_context(DOCUMENT.len() as u64);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let hook = Arc::new(TestPublicationHook::new());
    let _hook_guard = package
        .cache
        .install_test_publication_hook(Arc::clone(&hook));

    let (loader_result, waiter_result) = std::thread::scope(|scope| {
        let package = &package;
        let loader = scope.spawn(move || package.main_document_part().unwrap().data());
        source.wait_until_payload_read();
        let waiter = scope.spawn(move || package.main_document_part().unwrap().data());
        hook.wait_until_waiter_joined();
        let diagnostics = package.cache_diagnostics();
        assert_eq!(diagnostics.in_flight_loads, 1);
        assert_eq!(diagnostics.waiter_joins, 1);
        assert_eq!(
            diagnostics.budget_cache_reserved_bytes,
            DOCUMENT.len() as u64
        );
        assert_eq!(diagnostics.budget_cache_reserved_objects, 2);
        source.release_payload_read();
        (loader.join().unwrap(), waiter.join().unwrap())
    });

    assert!(matches!(loader_result, Err(OpcError::ZipError(_))));
    assert!(matches!(waiter_result, Err(OpcError::ZipError(_))));
    assert!(source.payload_reads.load(Ordering::SeqCst) >= 2);
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.failed_loads, 2);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.retained_bytes, 0);
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert_eq!(diagnostics.budget_cache_reserved_bytes, 0);
    assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
    assert_eq!(budget.used(Resource::Memory), 0);
    let state = package.cache.state.lock().unwrap();
    assert!(state.entries.is_empty());
    assert!(state.pending.is_empty());
    assert!(state.flights.is_empty());
}

#[test]
fn managed_allocation_admission_bypass_returns_correct_bytes_without_retention() {
    const DOCUMENT: &[u8] = b"forced allocation-admission bypass";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        false,
    )));
    let (budget, context) = managed_context(4096);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let _bypass_guard = package.cache.force_test_cache_admission_bypass();
    let data = package.main_document_part().unwrap().data().unwrap();
    assert_eq!(data.as_bytes(), DOCUMENT);
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.cold_loads, 1);
    assert_eq!(diagnostics.successful_loads, 1);
    assert_eq!(diagnostics.bypasses, 1);
    assert_eq!(diagnostics.allocation_bypasses, 1);
    assert_eq!(diagnostics.oversized_bypasses, 0);
    assert_eq!(diagnostics.retained_entries, 0);
    assert_eq!(diagnostics.retained_bytes, 0);
    assert_eq!(diagnostics.in_flight_loads, 0);
    assert_eq!(diagnostics.budget_cache_reserved_bytes, 0);
    assert_eq!(diagnostics.budget_cache_reserved_objects, 0);
    assert_eq!(diagnostics.budget_reservation_failures, 0);
    assert_eq!(budget.used(Resource::Memory), DOCUMENT.len() as u64);
    drop(data);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn cache_cleanup_is_exact_identity_scoped_and_late_success_cannot_resurrect() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"document",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let first_id = package.catalog.parts[0].entry_id;
    let second_id = package.catalog.parts[1].entry_id;
    let cache = PartCache::new(SourceCacheLimits::new(32, 3).unwrap());
    let old_flight = Arc::new(LoadFlight::new(None, None, None));
    let new_flight = Arc::new(LoadFlight::new(None, None, None));
    let old_payload = CachedPayload {
        bytes: Arc::new(vec![1, 2]),
        reservation: None,
        object_reservation: None,
    };
    let new_payload = CachedPayload {
        bytes: Arc::new(vec![3, 4]),
        reservation: None,
        object_reservation: None,
    };
    assert!(matches!(
        cache.insert_for_test(first_id, new_payload.clone()),
        CacheRetention::Retained
    ));
    {
        let mut state = cache.state.lock().unwrap();
        state.flights.insert(first_id, Arc::clone(&new_flight));
        state.pending.insert(
            first_id,
            PendingPublication {
                flight: Arc::clone(&new_flight),
                payload: Arc::clone(&new_payload.bytes),
            },
        );
    }

    cache.rollback_publication(
        first_id,
        &old_flight,
        old_payload.clone(),
        CachePublication::Pending,
    );
    {
        let state = cache.state.lock().unwrap();
        assert!(
            state
                .flights
                .get(&first_id)
                .is_some_and(|flight| Arc::ptr_eq(flight, &new_flight))
        );
        assert!(state.pending.contains_key(&first_id));
        assert!(state.entries.contains_key(&first_id));
    }
    assert!(
            cache
                .complete_uncached_success(
                    first_id,
                    &old_flight,
                    old_payload,
                    CacheRetention::Retained,
                )
                .is_none()
        );
    assert!(
        cache
            .commit_pending(first_id, &new_flight, new_payload.clone())
            .is_some()
    );
    assert_eq!(cache.state.lock().unwrap().entries.len(), 1);

    let failed_flight = match cache.enter(second_id, 1).unwrap() {
        CacheAccess::Loader(flight) => flight,
        _ => panic!("expected a fresh loader flight"),
    };
    cache.complete_failure(second_id, &failed_flight);
    assert!(failed_flight.wait(None).unwrap().is_none());
    assert!(
        cache
            .complete_uncached_success(
                second_id,
                &failed_flight,
                CachedPayload {
                    bytes: Arc::new(vec![5]),
                    reservation: None,
                    object_reservation: None,
                },
                CacheRetention::Retained,
            )
            .is_none()
    );
    let state = cache.state.lock().unwrap();
    assert!(!state.entries.contains_key(&second_id));
    assert!(!state.pending.contains_key(&second_id));
    assert!(!state.flights.contains_key(&second_id));
}

#[test]
fn cache_ready_hint_rejects_flights_and_pending_without_touching_state() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"cache hint state",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let first_id = package.catalog.parts[0].entry_id;
    let second_id = package.catalog.parts[1].entry_id;
    let cache = PartCache::new(SourceCacheLimits::new(32, 3).unwrap());
    assert!(!cache.all_entries_ready([first_id].into_iter()));
    assert!(matches!(
        cache.insert_for_test(
            first_id,
            CachedPayload {
                bytes: Arc::new(vec![1, 2]),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Retained
    ));
    let clock_before = cache.state.lock().unwrap().clock;
    let hits_before = cache.diagnostics().hits;
    assert!(cache.all_entries_ready([first_id].into_iter()));
    assert_eq!(cache.state.lock().unwrap().clock, clock_before);
    assert_eq!(cache.diagnostics().hits, hits_before);

    assert!(matches!(
        cache.insert_for_test(
            second_id,
            CachedPayload {
                bytes: Arc::new(vec![4, 5]),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Retained
    ));
    let flight = Arc::new(LoadFlight::new(None, None, None));
    {
        let mut state = cache.state.lock().unwrap();
        state.flights.insert(second_id, Arc::clone(&flight));
    }
    assert!(cache.state.lock().unwrap().entries.contains_key(&second_id));
    assert!(!cache.all_entries_ready([second_id].into_iter()));
    {
        let mut state = cache.state.lock().unwrap();
        state.flights.remove(&second_id);
    }
    assert!(cache.all_entries_ready([second_id].into_iter()));
    {
        let mut state = cache.state.lock().unwrap();
        state.pending.insert(
            second_id,
            PendingPublication {
                flight,
                payload: Arc::new(vec![3]),
            },
        );
    }
    assert!(!cache.all_entries_ready([second_id].into_iter()));
    cache.state.lock().unwrap().pending.remove(&second_id);
    assert!(cache.all_entries_ready([second_id].into_iter()));

    let evicting = PartCache::new(SourceCacheLimits::new(4, 3).unwrap());
    assert!(matches!(
        evicting.insert_for_test(
            first_id,
            CachedPayload {
                bytes: Arc::new(vec![6, 7]),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Retained
    ));
    assert!(evicting.all_entries_ready([first_id].into_iter()));
    assert!(matches!(
        evicting.insert_for_test(
            second_id,
            CachedPayload {
                bytes: Arc::new(vec![8, 9, 10]),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Retained
    ));
    assert!(!evicting.all_entries_ready([first_id].into_iter()));
    assert!(evicting.all_entries_ready([second_id].into_iter()));
}

#[test]
fn cache_evicts_by_byte_weight_and_entry_count_and_rejects_oversized_values() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"document",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let first_id = package.catalog.parts[0].entry_id;
    let second_id = package.catalog.parts[1].entry_id;
    let cache = PartCache::new(SourceCacheLimits::new(3, 3).unwrap());
    let first = Arc::new(vec![1, 2]);
    assert!(matches!(
        cache.insert_for_test(
            first_id,
            CachedPayload {
                bytes: Arc::clone(&first),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Retained
    ));
    assert!(Arc::ptr_eq(
        &cache.state.lock().unwrap().entries[&first_id].payload.bytes,
        &first
    ));
    drop(first);
    assert!(matches!(
        cache.insert_for_test(
            second_id,
            CachedPayload {
                bytes: Arc::new(vec![3, 4]),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Retained
    ));
    assert!(!cache.state.lock().unwrap().entries.contains_key(&first_id));
    assert!(cache.state.lock().unwrap().entries.contains_key(&second_id));

    let entry_limited = PartCache::new(SourceCacheLimits::new(10, 1).unwrap());
    assert!(matches!(
        entry_limited.insert_for_test(
            first_id,
            CachedPayload {
                bytes: Arc::new(vec![1, 2]),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Retained
    ));
    assert!(matches!(
        entry_limited.insert_for_test(
            second_id,
            CachedPayload {
                bytes: Arc::new(vec![3, 4]),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Retained
    ));
    assert!(
        !entry_limited
            .state
            .lock()
            .unwrap()
            .entries
            .contains_key(&first_id)
    );
    assert!(
        entry_limited
            .state
            .lock()
            .unwrap()
            .entries
            .contains_key(&second_id)
    );

    assert!(matches!(
        cache.insert_for_test(
            first_id,
            CachedPayload {
                bytes: Arc::new(vec![0, 0, 0, 0]),
                reservation: None,
                object_reservation: None,
            },
        ),
        CacheRetention::Oversized
    ));
    assert!(!cache.state.lock().unwrap().entries.contains_key(&first_id));
    assert_eq!(cache.diagnostics().evictions, 1);
    assert_eq!(cache.diagnostics().oversized_bypasses, 0);
}

#[test]
fn cache_limits_reject_zero_bounds() {
    assert_eq!(
        SourceCacheLimits::new(0, 1),
        Err(SourceCacheLimitError::ZeroMaximumBytes)
    );
    assert_eq!(
        SourceCacheLimits::new(1, 0),
        Err(SourceCacheLimitError::ZeroMaximumEntries)
    );
}

#[test]
fn checked_counter_delta_rejects_a_regression() {
    let before = SourceCacheDiagnostics {
        hits: 4,
        ..SourceCacheDiagnostics::default()
    };
    let after = SourceCacheDiagnostics {
        hits: 3,
        ..SourceCacheDiagnostics::default()
    };

    assert_eq!(
        SourceCacheDiagnostics::checked_counter_delta(before, after),
        Err(SourceCacheDiagnosticsError::CounterMovedBackwards { counter: "hits" })
    );
}

#[test]
fn checked_counter_delta_preserves_event_counts_without_gauge_subtraction() {
    let before = SourceCacheDiagnostics {
        hits: 2,
        retained_entries: 1,
        retained_bytes: 10,
        ..SourceCacheDiagnostics::default()
    };
    let after = SourceCacheDiagnostics {
        hits: 5,
        retained_entries: 0,
        retained_bytes: 0,
        ..SourceCacheDiagnostics::default()
    };

    assert_eq!(
        SourceCacheDiagnostics::checked_counter_delta(before, after)
            .expect("valid counter interval"),
        SourceCacheCounterDelta {
            hits: 3,
            ..SourceCacheCounterDelta::default()
        }
    );
}

#[test]
fn checked_counter_overflow_fails_closed_snapshot() {
    let cache = PartCache::new(SourceCacheLimits::new(3, 3).unwrap());
    cache.counters.hits.value.store(u64::MAX, Ordering::Relaxed);
    cache.counters.hits.increment();

    assert_eq!(
        cache.try_diagnostics(),
        Err(SourceCacheDiagnosticsError::CounterOverflow)
    );
    assert_eq!(cache.counters.hits.load(), u64::MAX);
}

#[test]
fn try_cache_diagnostics_matches_compatibility_snapshot_when_healthy() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"healthy diagnostics",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    assert_eq!(
        package
            .try_cache_diagnostics()
            .expect("healthy diagnostic snapshot"),
        package.cache_diagnostics()
    );
}

#[test]
fn managed_input_reservation_counter_overflow_fails_closed_through_wrapper() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"managed input diagnostics",
        false,
    )));
    let (_budget, context) = managed_context(4096);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let counter = package
        .cache
        .input_reservation_failures
        .as_ref()
        .expect("managed input reservation counter");
    counter.value.store(u64::MAX, Ordering::Relaxed);
    counter.increment();

    assert_eq!(
        package.try_cache_diagnostics(),
        Err(SourceCacheDiagnosticsError::CounterOverflow)
    );
}

#[test]
fn managed_output_reservation_counter_overflow_fails_closed_through_wrapper() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"managed output diagnostics",
        false,
    )));
    let (_budget, context) = managed_context(4096);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let counter = package
        .cache
        .output_reservation_failures
        .as_ref()
        .expect("managed output reservation counter");
    counter.value.store(u64::MAX, Ordering::Relaxed);
    counter.increment();

    assert_eq!(
        package.try_cache_diagnostics(),
        Err(SourceCacheDiagnosticsError::CounterOverflow)
    );
}

#[test]
fn aggregate_reservation_counter_overflow_fails_closed_through_wrapper() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"managed aggregate diagnostics",
        false,
    )));
    let (_budget, context) = managed_context(4096);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    package
        .cache
        .counters
        .budget_reservation_failures
        .value
        .store(u64::MAX, Ordering::Relaxed);
    package
        .cache
        .input_reservation_failures
        .as_ref()
        .expect("managed input reservation counter")
        .value
        .store(1, Ordering::Relaxed);

    assert_eq!(
        package.try_cache_diagnostics(),
        Err(SourceCacheDiagnosticsError::CounterOverflow)
    );
}

#[test]
fn retained_gauge_aggregation_overflow_fails_closed_through_wrapper() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_with_part_names(
            "word/document.xml",
            &["word/document.xml", "custom/other.xml"],
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_eq!(package.catalog.parts.len(), 2);

    let bytes_budget_a = Budget::root(
        "diagnostic-gauge-bytes-a",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let bytes_budget_b = Budget::root(
        "diagnostic-gauge-bytes-b",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let objects_budget_a = Budget::root(
        "diagnostic-gauge-objects-a",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let objects_budget_b = Budget::root(
        "diagnostic-gauge-objects-b",
        Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let bytes_a = Arc::new(
        bytes_budget_a
            .reserve(Resource::Memory, u64::MAX)
            .expect("maximum diagnostic byte reservation"),
    );
    let bytes_b = Arc::new(
        bytes_budget_b
            .reserve(Resource::Memory, u64::MAX)
            .expect("maximum diagnostic byte reservation"),
    );
    let objects_a = Arc::new(
        objects_budget_a
            .reserve(Resource::Objects, u64::MAX)
            .expect("maximum diagnostic object reservation"),
    );
    let objects_b = Arc::new(
        objects_budget_b
            .reserve(Resource::Objects, u64::MAX)
            .expect("maximum diagnostic object reservation"),
    );
    let mut state = package.cache.state.lock().expect("unpoisoned cache state");
    for (index, (entry_id, bytes, objects)) in [
        (package.catalog.parts[0].entry_id, bytes_a, objects_a),
        (package.catalog.parts[1].entry_id, bytes_b, objects_b),
    ]
    .into_iter()
    .enumerate()
    {
        state.entries.insert(
            entry_id,
            CacheEntry {
                payload: CachedPayload {
                    bytes: Arc::new(Vec::new()),
                    reservation: Some(bytes),
                    object_reservation: Some(objects),
                },
                last_used: u64::try_from(index).expect("bounded test index"),
            },
        );
    }
    drop(state);

    assert_eq!(
        package.try_cache_diagnostics(),
        Err(SourceCacheDiagnosticsError::CounterOverflow)
    );
}

#[test]
fn poisoned_cache_state_fails_closed_snapshot() {
    let cache = Arc::new(PartCache::new(SourceCacheLimits::new(3, 3).unwrap()));
    let poisoner = Arc::clone(&cache);
    let join = std::thread::spawn(move || {
        let _state = poisoner.state.lock().expect("unpoisoned cache state");
        panic!("test cache-state poison");
    });
    assert!(join.join().is_err());

    assert_eq!(
        cache.try_diagnostics(),
        Err(SourceCacheDiagnosticsError::StatePoisoned)
    );
}

#[test]
fn poisoned_cache_state_fails_closed_through_wrapper() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"poisoned diagnostics",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let _state = package.cache.state.lock().expect("unpoisoned cache state");
        panic!("test cache-state poison through package wrapper");
    }));
    assert!(result.is_err());

    assert_eq!(
        package.try_cache_diagnostics(),
        Err(SourceCacheDiagnosticsError::StatePoisoned)
    );
}

#[test]
fn catalog_entry_id_resolves_to_the_same_payload_as_its_member_name() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"entry identity",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    let catalog = package
        .catalog
        .parts
        .iter()
        .find(|catalog| catalog.partname == *part.partname())
        .unwrap();
    assert_eq!(
        package.archive.read_entry(catalog.entry_id).unwrap(),
        package.archive.read(catalog.partname.membername()).unwrap()
    );
}

#[test]
fn source_changes_reject_metadata_cache_and_conversion_access() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"stable payload",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    part.data().unwrap();
    source.changed();
    assert!(matches!(
        package.part(part.partname()),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(matches!(part.data(), Err(OpcError::SourceChanged { .. })));
    assert!(matches!(
        package.into_opc_package(),
        Err(OpcError::SourceChanged { .. })
    ));
}

#[test]
fn catalog_reports_non_parts_and_conversion_matches_loaded_parts() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"document",
        true,
    )));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_eq!(package.iter_parts().count(), 2);
    assert_eq!(package.non_part_members().len(), 1);
    assert_eq!(package.non_part_members()[0].name(), "scratch.bin");
    assert_eq!(
        package
            .main_document_part()
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"document"
    );
    let owned = package.into_opc_package().unwrap();
    assert_eq!(owned.part_count(), 2);
    assert_eq!(owned.non_part_members().len(), 1);
    assert_eq!(owned.main_document_part().unwrap().blob(), b"document");
}

#[test]
fn consuming_materialization_retains_cached_payload_allocation() {
    const DOCUMENT: &[u8] = b"consuming materialization payload";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        true,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let cached_payload = package
        .main_document_part()
        .unwrap()
        .data()
        .unwrap()
        .into_arc()
        .unwrap();

    let owned = package.into_opc_package().unwrap();
    let materialized_payload = owned.main_document_part().unwrap().blob_arc();
    assert!(Arc::ptr_eq(&cached_payload, &materialized_payload));
    assert_eq!(materialized_payload.as_slice(), DOCUMENT);

    // The owning package must retain the adopted allocation after every
    // source-backed handle is gone.
    drop(materialized_payload);
    drop(cached_payload);
    drop(source);
    assert_eq!(owned.main_document_part().unwrap().blob(), DOCUMENT);
}

#[test]
fn borrowed_materialization_retains_cached_payload_allocation_after_source_drop() {
    const DOCUMENT: &[u8] = b"borrowed materialization payload";
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        DOCUMENT,
        true,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let cached_payload = package
        .main_document_part()
        .unwrap()
        .data()
        .unwrap()
        .into_arc()
        .unwrap();

    let mut owned = package.to_opc_package().unwrap();
    let materialized_payload = owned.main_document_part().unwrap().blob_arc();
    assert!(Arc::ptr_eq(&cached_payload, &materialized_payload));
    assert_eq!(materialized_payload.as_slice(), DOCUMENT);

    // Mutating the owning package must detach its part without changing
    // the source-backed cache's shared payload.
    let target = PackURI::new("/word/document.xml").unwrap();
    owned
        .get_part_mut(&target)
        .unwrap()
        .set_blob(b"edited owning payload".to_vec());
    assert_eq!(
        owned.get_part(&target).unwrap().blob(),
        b"edited owning payload"
    );
    assert_eq!(
        package
            .main_document_part()
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        DOCUMENT
    );

    // Drop the borrowed source-backed package and all other handles. The
    // materialized package owns the shared payload independently.
    drop(materialized_payload);
    drop(cached_payload);
    drop(package);
    drop(source);
    assert_eq!(
        owned.main_document_part().unwrap().blob(),
        b"edited owning payload"
    );
}

#[test]
fn borrowed_materialization_preserves_source_and_matches_eager_graph() {
    let bytes = archive_bytes(root_relationships(), b"document", true);
    let source = Arc::new(CountingSource::new(bytes.clone()));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let owned = package.to_opc_package().unwrap();
    let eager = OpcPackage::from_bytes(&bytes, SourceLimits::default()).unwrap();

    assert_eq!(owned.part_count(), eager.part_count());
    assert_eq!(owned.non_part_members(), eager.non_part_members());
    assert_eq!(owned.rels().iter().count(), eager.rels().iter().count());
    for partname in [
        PackURI::new("/word/document.xml").unwrap(),
        PackURI::new("/custom/orphan.xml").unwrap(),
    ] {
        assert_eq!(
            owned.get_part(&partname).unwrap().blob(),
            eager.get_part(&partname).unwrap().blob()
        );
    }

    // The borrowed conversion must leave the source-backed catalog usable;
    // this second read is a cache hit and therefore does not touch the
    // positional source again.
    let reads_after_materialization = source.reads.load(Ordering::SeqCst);
    assert_eq!(
        package
            .main_document_part()
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        b"document"
    );
    assert_eq!(
        source.reads.load(Ordering::SeqCst),
        reads_after_materialization
    );
}

#[test]
fn preserving_materialization_retains_positional_source_for_later_edits() {
    let bytes = archive_bytes(root_relationships(), b"<document/>", true);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    let owned = package.to_opc_package_preserving_source().unwrap();
    assert_eq!(
        owned.exact_source_shared().as_deref().map(Vec::as_slice),
        Some(bytes.as_slice())
    );

    let document = PackURI::new("/word/document.xml").unwrap();
    let mut changed = owned.clone();
    changed
        .try_replace_owned_xml_part_bytes(
            &document,
            b"<document/>",
            Arc::new(b"<changed/>".to_vec()),
        )
        .unwrap();
    let mut output = Vec::new();
    changed.write_to(&mut output).unwrap();
    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(reopened.get_part(&document).unwrap().blob(), b"<changed/>");
}

#[test]
fn preserving_materialization_defers_part_payloads_until_read() {
    let bytes = archive_bytes(root_relationships(), b"<document/>", true);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    // No ordinary payload is inflated by the conversion.
    let owned = package.to_opc_package_preserving_source().unwrap();
    assert_eq!(owned.deferred_decode_counters(), Some((0, 0)));

    let document = PackURI::new("/word/document.xml").unwrap();
    assert_eq!(owned.get_part(&document).unwrap().blob(), b"<document/>");
    assert_eq!(owned.deferred_decode_counters(), Some((1, 11)));

    // An untouched package still publishes its exact source.
    let mut output = Vec::new();
    owned.write_to(&mut output).unwrap();
    assert_eq!(output, bytes);
}

#[test]
fn consuming_materialization_with_accounting_reports_mixed_cold_payload_work() {
    const DEFLATED_DOCUMENT: &[u8] = b"<document>deflated source payload</document>";
    const DEFLATED_SECOND: &[u8] = b"<second>deflated source payload</second>";
    const STORED_ORPHAN: &[u8] = b"<orphan>stored source payload</orphan>";
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(mixed_compression_archive())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut accounting = OpcOperationAccounting::default();

    let owned = package
        .into_opc_package_with_accounting(&mut accounting)
        .unwrap();

    assert_eq!(
        owned
            .get_part(&PackURI::new("/word/document.xml").unwrap())
            .unwrap()
            .blob(),
        DEFLATED_DOCUMENT
    );
    assert_eq!(
        owned
            .get_part(&PackURI::new("/custom/second.xml").unwrap())
            .unwrap()
            .blob(),
        DEFLATED_SECOND
    );
    assert_eq!(
        owned
            .get_part(&PackURI::new("/custom/orphan.xml").unwrap())
            .unwrap()
            .blob(),
        STORED_ORPHAN
    );
    assert_eq!(
        accounting.stored_payload_bytes_read(),
        STORED_ORPHAN.len() as u64
    );
    assert_eq!(
        accounting.stored_payload_bytes_accepted(),
        STORED_ORPHAN.len() as u64
    );
    assert!(accounting.compressed_deflate_payload_bytes_read() > 0);
    assert_eq!(
        accounting.deflate_bytes_produced(),
        (DEFLATED_DOCUMENT.len() + DEFLATED_SECOND.len()) as u64
    );
    assert_eq!(
        accounting.deflate_bytes_accepted(),
        accounting.deflate_bytes_produced()
    );
    assert_eq!(accounting.output_bytes_accepted(), 0);
    assert_eq!(accounting.generated_deflate_payload_bytes_emitted(), 0);
    assert_eq!(accounting.stored_payload_bytes_emitted(), 0);
    assert_eq!(accounting.precompressed_payload_bytes_emitted(), 0);
    assert_eq!(accounting.raw_unchanged_source_bytes_accepted(), 0);
}

#[test]
fn consuming_materialization_with_accounting_reports_owned_provenance_reads() {
    const CONTENT_TYPES: &[u8] = br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/></Types>"#;
    const STORED_ORPHAN: &[u8] = b"<orphan>stored source payload</orphan>";

    let package = SourceBackedPackage::from_vec(
        mixed_compression_archive(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    let mut accounting = OpcOperationAccounting::default();

    package
        .into_opc_package_with_accounting(&mut accounting)
        .unwrap();

    // The owning conversion reads the physical manifest and package
    // relationships once to authorize targeted preservation. Those are
    // ZIP payload reads just like the stored ordinary Part and therefore
    // belong in the caller's report.
    assert_eq!(
        accounting.stored_payload_bytes_read(),
        (CONTENT_TYPES.len() + root_relationships().len() + STORED_ORPHAN.len()) as u64
    );
    assert_eq!(
        accounting.stored_payload_bytes_accepted(),
        accounting.stored_payload_bytes_read()
    );
}

#[test]
fn borrowed_materialization_with_accounting_does_not_charge_preloaded_cache_hits() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(mixed_compression_archive())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    for part in package.iter_parts() {
        part.data().unwrap();
    }

    let mut accounting = OpcOperationAccounting::default();
    let owned = package
        .to_opc_package_with_accounting(&mut accounting)
        .unwrap();

    assert_eq!(accounting, OpcOperationAccounting::default());
    assert_eq!(owned.part_count(), package.iter_parts().count());
}

#[test]
fn consuming_materialization_with_accounting_refuses_managed_package_before_reads() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"managed consuming document",
        true,
    )));
    let (budget, context) = managed_context(4096);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    let read_bytes_before = source.read_bytes.load(Ordering::SeqCst);
    let mut accounting = OpcOperationAccounting::default();

    assert!(matches!(
        package.into_opc_package_with_accounting(&mut accounting),
        Err(OpcError::ManagedPackageMaterialization)
    ));
    assert_eq!(accounting, OpcOperationAccounting::default());
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(source.read_bytes.load(Ordering::SeqCst), read_bytes_before);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn consuming_materialization_with_accounting_keeps_malformed_read_progress() {
    const DOCUMENT: &[u8] = b"payload whose CRC is invalid for accounting";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[payload_offset] ^= 0xff;
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut accounting = OpcOperationAccounting::default();

    assert!(matches!(
        package.into_opc_package_with_accounting(&mut accounting),
        Err(OpcError::ZipError(_))
    ));
    assert!(accounting.stored_payload_bytes_read() > 0);
}

#[test]
fn consuming_materialization_with_accounting_keeps_source_change_read_progress() {
    const DOCUMENT: &[u8] = b"document changes during accounted materialization";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, true);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let source = Arc::new(ChangeDuringPayloadSource::new(bytes, payload_offset));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    source.armed.store(true, Ordering::SeqCst);
    let mut accounting = OpcOperationAccounting::default();

    assert!(matches!(
        package.into_opc_package_with_accounting(&mut accounting),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(accounting.stored_payload_bytes_read() > 0);
}

#[test]
fn borrowed_materialization_rejects_managed_package_before_payload_reads() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"managed document",
        true,
    )));
    let (budget, context) = managed_context(4096);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    let read_bytes_before = source.read_bytes.load(Ordering::SeqCst);

    assert!(matches!(
        package.to_opc_package(),
        Err(OpcError::ManagedPackageMaterialization)
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(source.read_bytes.load(Ordering::SeqCst), read_bytes_before);
    assert!(package.main_document_part().is_ok());
    drop(package);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn consuming_materialization_rejects_managed_package_before_payload_reads() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"managed consuming document",
        true,
    )));
    let (budget, context) = managed_context(4096);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    let read_bytes_before = source.read_bytes.load(Ordering::SeqCst);

    assert!(matches!(
        package.into_opc_package(),
        Err(OpcError::ManagedPackageMaterialization)
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    assert_eq!(source.read_bytes.load(Ordering::SeqCst), read_bytes_before);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn borrowed_materialization_rejects_stale_source_before_payload_reads() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"stable document",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    source.changed();

    assert!(matches!(
        package.to_opc_package(),
        Err(OpcError::SourceChanged { .. })
    ));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
}

#[test]
fn borrowed_materialization_rejects_source_changed_during_payload_and_keeps_no_partial_result() {
    const DOCUMENT: &[u8] = b"document changes during materialization";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, true);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let source = Arc::new(ChangeDuringPayloadSource::new(bytes, payload_offset));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    source.armed.store(true, Ordering::SeqCst);

    assert!(matches!(
        package.to_opc_package(),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(package.main_document_part().is_err());
}

#[test]
fn borrowed_materialization_rejects_cancellation_before_payload_reads() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"cancelled document",
        false,
    )));
    let (budget, cancellation_source, context) = managed_context_with_cancellation(4096);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let reads_before = source.reads.load(Ordering::SeqCst);
    cancellation_source.cancel();

    assert!(matches!(package.to_opc_package(), Err(OpcError::Cancelled)));
    assert_eq!(source.reads.load(Ordering::SeqCst), reads_before);
    drop(package);
    assert_eq!(budget.used(Resource::Objects), 0);
}

#[test]
fn borrowed_materialization_reports_malformed_payload_without_partial_result() {
    const DOCUMENT: &[u8] = b"payload whose CRC is invalid";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[payload_offset] ^= 0xff;
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    assert!(matches!(
        package.to_opc_package(),
        Err(OpcError::ZipError(_))
    ));
    assert!(package.main_document_part().is_ok());
    assert!(package.main_document_part().unwrap().data().is_err());
}

#[test]
fn borrowed_materialization_preserves_opaque_xml_like_eager_unmarshal() {
    const DOCUMENT: &[u8] = b"<document> <child/></document>";
    let source_bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let source_package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let materialized = source_package.to_opc_package().unwrap();
    let eager = OpcPackage::from_bytes(&source_bytes, SourceLimits::default()).unwrap();

    let mut materialized_bytes = Vec::new();
    materialized.write_to(&mut materialized_bytes).unwrap();
    let mut eager_bytes = Vec::new();
    eager.write_to(&mut eager_bytes).unwrap();
    let materialized_reader = crate::phys_pkg::OwnedPhysPkgReader::from_bytes(
        materialized_bytes,
        SourceLimits::default(),
    )
    .unwrap();
    let eager_reader =
        crate::phys_pkg::OwnedPhysPkgReader::from_bytes(eager_bytes, SourceLimits::default())
            .unwrap();
    assert_eq!(
        materialized_reader
            .read_member("word/document.xml")
            .unwrap(),
        DOCUMENT
    );
    assert_eq!(
        materialized_reader
            .read_member("word/document.xml")
            .unwrap(),
        eager_reader.read_member("word/document.xml").unwrap()
    );
}

#[test]
fn borrowed_signed_materialization_retains_publication_policy() {
    let source = signed_archive(b"<signed/>");
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let materialized = package.to_opc_package().unwrap();
    assert!(materialized.is_signed());

    let mut output = Vec::new();
    assert!(matches!(
        materialized.write_to(&mut output),
        Err(OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(output.is_empty());
}

#[test]
fn one_part_overlay_raw_copies_every_unselected_member_and_reopens() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let source_raw = raw_records(&source_bytes);
    let source = Arc::new(CountingSource::new(source_bytes));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let target = PackURI::new("/word/document.xml").unwrap();
    let mut output = Vec::new();
    package
        .write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec())
        .unwrap();

    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(reopened.get_part(&target).unwrap().blob(), b"<after/>");
    assert_eq!(reopened.part_count(), 2);
    assert_eq!(reopened.non_part_members().len(), 1);
    assert_eq!(reopened.non_part_members()[0].name(), "scratch.bin");
    let output_raw = raw_records(&output);
    assert_eq!(output_raw.len(), source_raw.len());
    for (name, source_record) in source_raw {
        if name == b"word/document.xml" {
            assert_ne!(output_raw[&name].local, source_record.local);
        } else {
            assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
            assert_eq!(
                central_without_local_offset(&output_raw[&name].central),
                central_without_local_offset(&source_record.central),
                "{name:?}"
            );
        }
    }
}

#[test]
fn zip64_one_part_overlay_preserves_raw_unknown_members_and_cold_work() {
    let source_bytes = promote_to_zip64(archive_bytes(root_relationships(), b"<before/>", true));
    let source_raw = raw_records(&source_bytes);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert_eq!(package.cache_diagnostics().cold_loads, 0);

    let target = PackURI::new("/word/document.xml").unwrap();
    let replacement = b"<after zip64=\"true\"/>".to_vec();
    let mut output = Vec::new();
    let mut accounting = OpcOperationAccounting::default();
    package
        .write_part_overlay_to_stream_with_accounting(
            &mut output,
            &target,
            replacement.clone(),
            &mut accounting,
        )
        .unwrap();

    let source_archive = soapberry_zip::ZipArchive::from_slice(&source_bytes).unwrap();
    assert!(source_archive.is_zip64());
    let output_archive = soapberry_zip::ZipArchive::from_slice(&output).unwrap();
    assert!(output_archive.is_zip64());
    assert_eq!(
        accounting.stored_payload_bytes_read(),
        b"<before/>".len() as u64
    );
    assert_eq!(accounting.compressed_deflate_payload_bytes_read(), 0);
    assert_eq!(accounting.output_bytes_accepted(), output.len() as u64);
    assert!(accounting.raw_unchanged_source_bytes_accepted() > 0);

    let output_raw = raw_records(&output);
    assert_eq!(output_raw.len(), source_raw.len());
    for (name, source_record) in source_raw {
        if name.as_slice() == b"word/document.xml" {
            assert_ne!(output_raw[&name].local, source_record.local);
        } else {
            assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
            assert_eq!(
                central_without_local_offset(&output_raw[&name].central),
                central_without_local_offset(&source_record.central),
                "{name:?}"
            );
        }
    }

    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(reopened.get_part(&target).unwrap().blob(), replacement);
    assert_eq!(
        reopened
            .non_part_members()
            .iter()
            .find(|member| member.name() == "scratch.bin")
            .unwrap()
            .name(),
        "scratch.bin"
    );
}

#[test]
fn shared_single_overlay_matches_vec_changed_and_signed_noop_publication() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let target = PackURI::new("/word/document.xml").unwrap();
    let mut vec_output = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_part_overlay_to_stream(&mut vec_output, &target, b"<after/>".to_vec())
    .unwrap();
    let mut shared_output = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_part_overlay_shared_to_stream(
        &mut shared_output,
        &target,
        Arc::new(b"<after/>".to_vec()),
    )
    .unwrap();
    assert_eq!(shared_output, vec_output);

    let signed_bytes = signed_archive(b"<signed/>");
    let signed_target = PackURI::new("/word/document.xml").unwrap();
    let mut vec_noop = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(signed_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_part_overlay_to_stream(&mut vec_noop, &signed_target, b"<signed/>".to_vec())
    .unwrap();
    let mut shared_noop = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(signed_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_part_overlay_shared_to_stream(
        &mut shared_noop,
        &signed_target,
        Arc::new(b"<signed/>".to_vec()),
    )
    .unwrap();
    assert_eq!(vec_noop, signed_bytes);
    assert_eq!(shared_noop, vec_noop);
}

#[test]
fn shared_multi_overlay_matches_vec_and_reopens() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let document = PackURI::new("/word/document.xml").unwrap();
    let orphan = PackURI::new("/custom/orphan.xml").unwrap();
    let mut vec_output = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_part_overlays_to_stream(
        &mut vec_output,
        vec![
            (orphan.clone(), b"<orphan-after/>".to_vec()),
            (document.clone(), b"<document-after/>".to_vec()),
        ],
    )
    .unwrap();

    let mut shared_output = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_part_overlays_shared_to_stream(
        &mut shared_output,
        vec![
            (orphan.clone(), Arc::new(b"<orphan-after/>".to_vec())),
            (document.clone(), Arc::new(b"<document-after/>".to_vec())),
        ],
    )
    .unwrap();
    assert_eq!(shared_output, vec_output);
    let reopened = OpcPackage::from_bytes(&shared_output, SourceLimits::default()).unwrap();
    assert_eq!(
        reopened.get_part(&document).unwrap().blob(),
        b"<document-after/>"
    );
    assert_eq!(
        reopened.get_part(&orphan).unwrap().blob(),
        b"<orphan-after/>"
    );
}

#[test]
fn shared_overlay_preserves_duplicate_and_limit_refusals_before_output() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let document = PackURI::new("/word/document.xml").unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlays_shared_to_stream(
            &mut output,
            vec![
                (document.clone(), Arc::new(b"<first/>".to_vec())),
                (document.clone(), Arc::new(b"<second/>".to_vec())),
            ],
        ),
        Err(OpcError::DuplicatePartName(_))
    ));
    assert!(output.is_empty());

    let limits = SourceLimits::builder()
        .max_part_bytes(10)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert!(matches!(
        package.write_part_overlay_shared_to_stream(
            &mut output,
            &document,
            Arc::new(vec![b'x'; 11]),
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::PartBytes,
            observed: 11,
            maximum: 10,
        })
    ));
    assert!(output.is_empty());
}

#[test]
fn shared_overlay_reports_partial_sink_failure_with_bounded_writes() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        true,
    )));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let target = PackURI::new("/word/document.xml").unwrap();
    let mut sink = BoundedFailingSink {
        accepted: 0,
        limit: 100,
        largest_write: 0,
    };
    let error = package
        .write_part_overlay_shared_to_stream(&mut sink, &target, Arc::new(b"<after/>".to_vec()))
        .unwrap_err();
    match error {
        OpcError::IncompleteOutput { written, .. } => assert_eq!(written, 100),
        other => panic!("unexpected sink error: {other:?}"),
    }
    assert!(sink.largest_write <= SOURCE_PUBLICATION_CHUNK_BYTES);
}

#[test]
fn multi_part_overlay_changes_only_selected_raw_members_and_reopens() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let source_raw = raw_records(&source_bytes);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let document = PackURI::new("/word/document.xml").unwrap();
    let orphan = PackURI::new("/custom/orphan.xml").unwrap();
    let mut output = Vec::new();
    package
        .write_part_overlays_to_stream(
            &mut output,
            vec![
                (orphan.clone(), b"<orphan-after/>".to_vec()),
                (document.clone(), b"<document-after/>".to_vec()),
            ],
        )
        .unwrap();

    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(
        reopened.get_part(&document).unwrap().blob(),
        b"<document-after/>"
    );
    assert_eq!(
        reopened.get_part(&orphan).unwrap().blob(),
        b"<orphan-after/>"
    );
    let output_raw = raw_records(&output);
    assert_eq!(output_raw.len(), source_raw.len());
    for (name, source_record) in source_raw {
        if matches!(name.as_slice(), b"word/document.xml" | b"custom/orphan.xml") {
            assert_ne!(output_raw[&name].local, source_record.local, "{name:?}");
        } else {
            assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
            assert_eq!(
                central_without_local_offset(&output_raw[&name].central),
                central_without_local_offset(&source_record.central),
                "{name:?}"
            );
        }
    }
}

#[test]
fn overlay_mixed_storage_has_semantic_parity_and_managed_cold_work() {
    let source_bytes = mixed_compression_archive();
    let source_raw = raw_records(&source_bytes);
    let document = PackURI::new("/word/document.xml").unwrap();
    let second = PackURI::new("/custom/second.xml").unwrap();
    let orphan = PackURI::new("/custom/orphan.xml").unwrap();
    let document_payload = b"<document>deflated source payload</document>";
    let second_payload = b"<second>deflated source payload</second>";
    let orphan_payload = b"<orphan>stored source payload</orphan>";
    let replacements = || {
        vec![
            (orphan.clone(), orphan_payload.to_vec()),
            (second.clone(), second_payload.to_vec()),
            (
                document.clone(),
                b"<document>changed payload</document>".to_vec(),
            ),
        ]
    };

    let mut unmanaged_output = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap()
    .write_part_overlays_to_stream(&mut unmanaged_output, replacements())
    .unwrap();

    let reopened = OpcPackage::from_bytes(&unmanaged_output, SourceLimits::default()).unwrap();
    assert_eq!(
        reopened.get_part(&document).unwrap().blob(),
        b"<document>changed payload</document>"
    );
    assert_eq!(reopened.get_part(&second).unwrap().blob(), second_payload);
    assert_eq!(reopened.get_part(&orphan).unwrap().blob(), orphan_payload);

    let output_raw = raw_records(&unmanaged_output);
    let source_orphan = source_raw.get(b"custom/orphan.xml".as_slice()).unwrap();
    let output_orphan = output_raw.get(b"custom/orphan.xml".as_slice()).unwrap();
    assert_eq!(output_orphan.local, source_orphan.local);
    assert_eq!(
        central_without_local_offset(&output_orphan.central),
        central_without_local_offset(&source_orphan.central)
    );
    let source_second = source_raw.get(b"custom/second.xml".as_slice()).unwrap();
    let output_second = output_raw.get(b"custom/second.xml".as_slice()).unwrap();
    assert_eq!(output_second.local, source_second.local);
    assert_eq!(
        central_without_local_offset(&output_second.central),
        central_without_local_offset(&source_second.central)
    );
    let source_document = source_raw.get(b"word/document.xml".as_slice()).unwrap();
    let output_document = output_raw.get(b"word/document.xml".as_slice()).unwrap();
    assert_ne!(output_document.local, source_document.local);
    assert_eq!(
        &output_document.central[10..12],
        &source_document.central[10..12]
    );

    let (budget, _cancellation_source, context) =
        managed_context_with_all_resources(64 * 1024, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let mut managed_output = Vec::new();
    SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap()
    .write_part_overlays_to_stream(&mut managed_output, replacements())
    .unwrap();

    assert_eq!(managed_output, unmanaged_output);
    assert_eq!(
        budget.used(Resource::Work),
        (document_payload.len() + second_payload.len() + orphan_payload.len()) as u64
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn multi_part_relationship_removal_has_exact_and_one_over_aggregate_bounds() {
    let exact_each = MAX_SOURCE_RELATIONSHIP_REMOVALS / 2;
    let (source_bytes, first_ids, second_ids) =
        archive_with_relationship_batches(exact_each, exact_each);
    let document = PackURI::new("/word/document.xml").unwrap();
    let orphan = PackURI::new("/custom/orphan.xml").unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    package
        .write_part_overlays_with_external_relationship_removals_to_stream(
            &mut output,
            vec![
                (document.clone(), b"<after/>".to_vec(), first_ids),
                (orphan.clone(), b"<after/>".to_vec(), second_ids),
            ],
        )
        .unwrap();
    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert!(reopened.get_part(&document).unwrap().rels().is_empty());
    assert!(reopened.get_part(&orphan).unwrap().rels().is_empty());

    let (source_bytes, first_ids, second_ids) =
        archive_with_relationship_batches(exact_each, exact_each + 1);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    output.clear();
    assert!(matches!(
        package.write_part_overlays_with_external_relationship_removals_to_stream(
            &mut output,
            vec![
                (document, b"<after/>".to_vec(), first_ids),
                (orphan, b"<after/>".to_vec(), second_ids),
            ],
        ),
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn multi_part_relationship_removal_checks_combined_part_and_archive_totals() {
    let (source_bytes, first_ids, second_ids) = archive_with_relationship_batches(1, 1);
    let source_package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let source_total = source_archive_total(&source_package);
    drop(source_package);
    let large = format!("<after>{}</after>", "x".repeat(4096)).into_bytes();
    let overlays = vec![
        (
            PackURI::new("/word/document.xml").unwrap(),
            large.clone(),
            first_ids.clone(),
        ),
        (
            PackURI::new("/custom/orphan.xml").unwrap(),
            large.clone(),
            second_ids.clone(),
        ),
    ];
    let part_limits = SourceLimits::builder()
        .max_total_part_bytes(19)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        part_limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlays_with_external_relationship_removals_to_stream(
            &mut output,
            overlays.clone(),
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::TotalPartBytes,
            ..
        })
    ));
    assert!(output.is_empty());
    let limits = SourceLimits::builder()
        .max_aggregate_bytes(source_total + 1)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlays_with_external_relationship_removals_to_stream(
            &mut output,
            overlays,
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::AggregateBytes,
            ..
        })
    ));
    assert!(output.is_empty());

    let (source_bytes, first_ids, second_ids) = archive_with_relationship_batches(1, 1);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut unrestricted = Vec::new();
    package
        .write_part_overlays_with_external_relationship_removals_to_stream(
            &mut unrestricted,
            vec![
                (
                    PackURI::new("/word/document.xml").unwrap(),
                    large.clone(),
                    first_ids.clone(),
                ),
                (
                    PackURI::new("/custom/orphan.xml").unwrap(),
                    large.clone(),
                    second_ids.clone(),
                ),
            ],
        )
        .unwrap();
    let published = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(unrestricted)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let published_total = source_archive_total(&published);
    assert!(published_total > source_total);
    drop(published);
    let exact_limits = SourceLimits::builder()
        .max_aggregate_bytes(published_total)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        exact_limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut exact_output = Vec::new();
    package
        .write_part_overlays_with_external_relationship_removals_to_stream(
            &mut exact_output,
            vec![
                (
                    PackURI::new("/word/document.xml").unwrap(),
                    large.clone(),
                    first_ids.clone(),
                ),
                (
                    PackURI::new("/custom/orphan.xml").unwrap(),
                    large.clone(),
                    second_ids.clone(),
                ),
            ],
        )
        .unwrap();
    assert!(!exact_output.is_empty());

    let under_limits = SourceLimits::builder()
        .max_aggregate_bytes(published_total - 1)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        under_limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut under_output = Vec::new();
    assert!(matches!(
        package.write_part_overlays_with_external_relationship_removals_to_stream(
            &mut under_output,
            vec![
                (
                    PackURI::new("/word/document.xml").unwrap(),
                    large.clone(),
                    first_ids,
                ),
                (
                    PackURI::new("/custom/orphan.xml").unwrap(),
                    large,
                    second_ids,
                ),
            ],
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::AggregateBytes,
            ..
        })
    ));
    assert!(under_output.is_empty());
}

#[test]
fn external_relationship_removal_overlay_changes_only_owner_and_rels_member() {
    let source_bytes = archive_with_document_relationships(document_relationships());
    let source_raw = raw_records(&source_bytes);
    let document = PackURI::new("/word/document.xml").unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    package
        .write_part_overlay_with_external_relationship_removals_to_stream(
            &mut output,
            &document,
            b"<after/>".to_vec(),
            vec!["rExternal".to_owned()],
        )
        .unwrap();

    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    let main = reopened.get_part(&document).unwrap();
    assert_eq!(main.blob(), b"<after/>");
    assert!(main.rels().get("rExternal").is_none());
    assert_eq!(
        main.rels()
            .get("rInternal")
            .unwrap()
            .target_partname()
            .unwrap(),
        PackURI::new("/custom/orphan.xml").unwrap()
    );
    let output_raw = raw_records(&output);
    assert_eq!(output_raw.len(), source_raw.len());
    for (name, source_record) in source_raw {
        if matches!(
            name.as_slice(),
            b"word/document.xml" | b"word/_rels/document.xml.rels"
        ) {
            assert_ne!(output_raw[&name].local, source_record.local, "{name:?}");
        } else {
            assert_eq!(output_raw[&name].local, source_record.local, "{name:?}");
            assert_eq!(
                central_without_local_offset(&output_raw[&name].central),
                central_without_local_offset(&source_record.central),
                "{name:?}"
            );
        }
    }
}

#[test]
fn external_relationship_removal_overlay_refuses_ids_and_limits_before_output() {
    let source_bytes = archive_with_document_relationships(document_relationships());
    let document = PackURI::new("/word/document.xml").unwrap();
    for (ids, expected) in [
        (vec!["missing".to_owned()], "missing external relationship"),
        (
            vec!["rExternal".to_owned(), "rExternal".to_owned()],
            "duplicate relationship",
        ),
        (vec!["rInternal".to_owned()], "internal relationship"),
    ] {
        let package = SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(source_bytes.clone())),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let mut output = Vec::new();
        let error = package
            .write_part_overlay_with_external_relationship_removals_to_stream(
                &mut output,
                &document,
                b"<after/>".to_vec(),
                ids,
            )
            .unwrap_err();
        match expected {
            "missing external relationship" => {
                assert!(matches!(error, OpcError::RelationshipNotFound(_)));
            },
            "duplicate relationship" => {
                assert!(matches!(error, OpcError::DuplicateRelationshipId(_)));
            },
            "internal relationship" => {
                assert!(matches!(error, OpcError::InvalidRelationship(_)));
            },
            _ => unreachable!(),
        }
        assert!(output.is_empty());
    }

    let exact_bound_ids = (0..MAX_SOURCE_RELATIONSHIP_REMOVALS)
        .map(|index| format!("missing-{index}"))
        .collect();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlay_with_external_relationship_removals_to_stream(
            &mut output,
            &document,
            b"<after/>".to_vec(),
            exact_bound_ids,
        ),
        Err(OpcError::RelationshipNotFound(_))
    ));
    assert!(output.is_empty());

    let over_bound_ids = (0..=MAX_SOURCE_RELATIONSHIP_REMOVALS)
        .map(|index| format!("missing-{index}"))
        .collect();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlay_with_external_relationship_removals_to_stream(
            &mut output,
            &document,
            b"<after/>".to_vec(),
            over_bound_ids,
        ),
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
    assert!(output.is_empty());

    let limits = SourceLimits::builder()
        .max_part_bytes(10)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlay_with_external_relationship_removals_to_stream(
            &mut output,
            &document,
            vec![b'x'; 11],
            vec!["rExternal".to_owned()],
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::PartBytes,
            observed: 11,
            maximum: 10,
        })
    ));
    assert!(output.is_empty());
}

#[test]
fn multi_part_overlay_checks_set_bounds_duplicates_and_aggregate_limits() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let document = PackURI::new("/word/document.xml").unwrap();
    let orphan = PackURI::new("/custom/orphan.xml").unwrap();

    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlays_to_stream(
            &mut output,
            vec![
                (document.clone(), b"<first/>".to_vec()),
                (document.clone(), b"<second/>".to_vec()),
            ],
        ),
        Err(OpcError::DuplicatePartName(_))
    ));
    assert!(output.is_empty());

    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let oversized_set = (0..=MAX_SOURCE_OVERLAY_PARTS)
        .map(|_| (document.clone(), b"<changed/>".to_vec()))
        .collect();
    assert!(matches!(
        package.write_part_overlays_to_stream(&mut output, oversized_set),
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
    assert!(output.is_empty());

    let limits = SourceLimits::builder()
        .max_part_bytes(20)
        .unwrap()
        .max_total_part_bytes(21)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert!(matches!(
        package.write_part_overlays_to_stream(
            &mut output,
            vec![
                (document, b"<document/>".to_vec()),
                (orphan, b"<orphan-2/>".to_vec()),
            ],
        ),
        Err(OpcError::Limit {
            kind: SourceLimitKind::TotalPartBytes,
            ..
        })
    ));
    assert!(output.is_empty());
}

#[test]
fn multi_part_overlay_all_noop_preserves_signed_source_identity() {
    let source_bytes = signed_archive(b"<signed/>");
    let document = PackURI::new("/word/document.xml").unwrap();
    let signature = PackURI::new("/signature/origin.xml").unwrap();
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    package
        .write_part_overlays_to_stream(
            &mut output,
            vec![
                (document.clone(), b"<signed/>".to_vec()),
                (signature.clone(), b"<origin/>".to_vec()),
            ],
        )
        .unwrap();
    assert_eq!(output, source_bytes);

    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(signed_archive(b"<signed/>"))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    output.clear();
    assert!(matches!(
        package.write_part_overlays_to_stream(
            &mut output,
            vec![
                (document, b"<changed/>".to_vec()),
                (signature, b"<origin/>".to_vec()),
            ],
        ),
        Err(OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(output.is_empty());
}

#[test]
fn one_part_overlay_exact_noop_preserves_every_source_byte() {
    let source_bytes = archive_bytes(root_relationships(), b"malformed but unchanged", true);
    let source = Arc::new(CountingSource::new(source_bytes.clone()));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let target = PackURI::new("/word/document.xml").unwrap();
    let mut output = Vec::new();
    package
        .write_part_overlay_to_stream(&mut output, &target, b"malformed but unchanged".to_vec())
        .unwrap();
    assert_eq!(output, source_bytes);
}

#[test]
fn unmanaged_exact_source_publication_remains_unbudgeted() {
    let source_bytes = archive_bytes(root_relationships(), b"unmanaged exact source", true);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut output = Vec::new();
    artifact.write_to(&mut output).unwrap();

    assert_eq!(output, source_bytes);
    let diagnostics = package.cache_diagnostics();
    assert!(!diagnostics.budget_managed);
    assert_eq!(diagnostics.budget_output_bytes_used, 0);
    assert_eq!(diagnostics.budget_output_bytes_limit, None);
    assert_eq!(diagnostics.budget_reservation_failures, 0);
}

#[test]
fn a_warm_part_read_observes_the_source_twice() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"warm part read",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let name = PackURI::new("/word/document.xml").unwrap();

    // Fill the cache. This read is cold and is measured separately.
    package.part(&name).unwrap().data().unwrap();

    let versions_before = source.versions.load(Ordering::SeqCst);
    let reads_before = source.reads.load(Ordering::SeqCst);
    let hits_before = package.cache_diagnostics().hits;
    package.part(&name).unwrap().data().unwrap();

    assert!(
        package.cache_diagnostics().hits > hits_before,
        "the second read must be a cache hit for this to measure the warm path"
    );
    assert_eq!(
        source.reads.load(Ordering::SeqCst) - reads_before,
        0,
        "a warm part read must not read source bytes"
    );
    // One observation for the retained-catalog lookup and one for the
    // hit's publication check, which also invalidates a stale entry.
    assert_eq!(
        source.versions.load(Ordering::SeqCst) - versions_before,
        2,
        "a warm part read must observe the source exactly twice"
    );
}

#[test]
fn a_cold_part_read_keeps_its_complete_observation_bracket() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"cold part read",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let name = PackURI::new("/word/document.xml").unwrap();

    let versions_before = source.versions.load(Ordering::SeqCst);
    let reads_before = source.reads.load(Ordering::SeqCst);
    package.part(&name).unwrap().data().unwrap();

    assert!(
        source.reads.load(Ordering::SeqCst) > reads_before,
        "a cold part read must read source bytes"
    );
    // The catalog lookup in `part()`, the cold load's opening observation,
    // the post-read fence, and the fence taken after the provisional
    // publication. A removed pre-publication observation once stood
    // between the post-read fence and the publication fence with no source
    // read between them and no side effect of its own.
    assert_eq!(
        source.versions.load(Ordering::SeqCst) - versions_before,
        4,
        "a cold part read must keep its complete observation bracket"
    );
}

#[test]
fn a_stream_monitors_every_positional_read_it_takes() {
    let source = Arc::new(CountingSource::chunked(
        archive_bytes(root_relationships(), b"monitored stream payload", false),
        64,
    ));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let name = PackURI::new("/word/document.xml").unwrap();

    let versions_before = source.versions.load(Ordering::SeqCst);
    let reads_before = source.reads.load(Ordering::SeqCst);
    let mut sink = Vec::new();
    package.part(&name).unwrap().stream_to(&mut sink).unwrap();
    let observations = source.versions.load(Ordering::SeqCst) - versions_before;
    let reads = source.reads.load(Ordering::SeqCst) - reads_before;

    assert!(reads > 1, "the chunked source must split this stream");
    assert!(
        observations >= 2 * reads,
        "a stream in progress must observe the source on both sides of every \
             positional read: {observations} observations for {reads} reads"
    );
}

#[test]
fn a_completed_stream_stops_monitoring_later_positional_reads() {
    let bytes = archive_bytes(root_relationships(), b"bounded monitor scope", false);
    let name = PackURI::new("/word/document.xml").unwrap();

    // A package that never streamed pays only the cold read's own bracket.
    let plain_source = Arc::new(CountingSource::new(bytes.clone()));
    let plain = SourceBackedPackage::from_read_at(
        plain_source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let plain_before = plain_source.versions.load(Ordering::SeqCst);
    let plain_data = plain.part(&name).unwrap().data().unwrap();
    let plain_observations = plain_source.versions.load(Ordering::SeqCst) - plain_before;

    // `stream_to` bypasses the part cache, so the read that follows it on
    // the same package is still a cold read of the same Part.
    let streamed_source = Arc::new(CountingSource::new(bytes));
    let streamed = SourceBackedPackage::from_read_at(
        streamed_source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut sink = Vec::new();
    streamed.part(&name).unwrap().stream_to(&mut sink).unwrap();
    let streamed_before = streamed_source.versions.load(Ordering::SeqCst);
    let streamed_data = streamed.part(&name).unwrap().data().unwrap();
    let streamed_observations = streamed_source.versions.load(Ordering::SeqCst) - streamed_before;

    assert_eq!(streamed_data.as_bytes(), plain_data.as_bytes());
    assert_eq!(sink.as_slice(), plain_data.as_bytes());
    assert_eq!(
        streamed_observations, plain_observations,
        "a completed stream must leave later reads on the observation count \
             an unstreamed package pays"
    );
}

#[test]
fn a_source_change_at_every_observation_of_a_cold_read_is_still_source_changed() {
    // `ChangeOnHitVersionSource` publishes the new version *on* the
    // `version()` call it is armed for, so arming at ordinal k makes the
    // k-th observation of the read the one that sees the change. The cold
    // read takes four: `part()`, the cold-load closure head, the post-read
    // fence, and the fence after the provisional publication. Each must
    // report the same typed error, including ordinal 3, which was once the
    // removed pre-publication observation and is now the publication
    // fence.
    for skip in 0..4usize {
        let source = Arc::new(ChangeOnHitVersionSource::new(archive_bytes(
            root_relationships(),
            b"change under a cold read",
            false,
        )));
        let package = SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None,
        )
        .unwrap();
        let name = PackURI::new("/word/document.xml").unwrap();
        source.arm_after_versions(skip);
        let error = package
            .part(&name)
            .and_then(|view| view.data())
            .expect_err("a changed source must refuse the read");
        assert!(
            matches!(error, OpcError::SourceChanged { .. }),
            "ordinal {skip} reported {error:?}"
        );
    }

    // One ordinal past the read's last observation, the read itself is
    // complete and correct — every fence it took saw the unchanged source —
    // and the change is reported by the next observation, not lost. This
    // ordinal once fell inside the read because the read took one more
    // observation; the mutation is caught either way, only its reporting
    // point moves.
    let source = Arc::new(ChangeOnHitVersionSource::new(archive_bytes(
        root_relationships(),
        b"change under a cold read",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let name = PackURI::new("/word/document.xml").unwrap();
    source.arm_after_versions(4);
    let data = package
        .part(&name)
        .and_then(|view| view.data())
        .expect("the source is unchanged through every fence this read takes");
    assert_eq!(data.as_bytes(), b"change under a cold read");
    let error = package
        .part(&name)
        .and_then(|view| view.data())
        .expect_err("the next observation must report the change");
    assert!(
        matches!(error, OpcError::SourceChanged { .. }),
        "the observation after the read reported {error:?}"
    );
}

#[test]
fn a_nested_monitored_scope_does_not_end_the_scope_around_it() {
    let depth = Arc::new(MonitoredReadDepth::new(0));
    let outer = MonitoredReads::enter(&depth);
    let inner = MonitoredReads::enter(&depth);
    assert_eq!(depth.load(Ordering::SeqCst), 2);
    drop(inner);
    assert_eq!(
        depth.load(Ordering::SeqCst),
        1,
        "an inner scope must not end the monitoring its caller opened"
    );
    drop(outer);
    assert_eq!(depth.load(Ordering::SeqCst), 0);
}

#[test]
fn a_changed_source_outranks_a_missing_part() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"missing part precedence",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let absent = PackURI::new("/word/absent.xml").unwrap();

    assert!(
        matches!(package.part(&absent), Err(OpcError::PartNotFound(_))),
        "a stable source reports the missing Part"
    );

    source.revision.fetch_add(1, Ordering::SeqCst);

    assert!(
        matches!(package.part(&absent), Err(OpcError::SourceChanged { .. })),
        "a changed source outranks the missing Part, which callers that map \
             `PartNotFound` to a non-error depend on"
    );
}

#[test]
fn exact_source_publication_checks_one_fewer_version_per_copy_chunk() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"exact source publication",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let versions_before = source.versions.load(Ordering::SeqCst);
    let mut output = Vec::new();

    package
        .write_part_overlays_to_stream(&mut output, Vec::new())
        .unwrap();

    // This fixture fits in one publication chunk. The initial check,
    // read pre/post checks, sink write/flush checks, and final check are
    // the complete freshness contract; the removed post-read check would
    // make this seven instead of six observations.
    assert_eq!(source.versions.load(Ordering::SeqCst) - versions_before, 6);
}

#[test]
fn one_part_overlay_refuses_unsupported_physical_layout_before_output() {
    let mut source_bytes = b"unsupported ZIP prelude".to_vec();
    source_bytes.extend_from_slice(&archive_bytes(root_relationships(), b"<before/>", true));
    let source = Arc::new(CountingSource::new(source_bytes));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let target = PackURI::new("/word/document.xml").unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec()),
        Err(OpcError::SourceBackedOverlayUnavailable { .. })
    ));
    assert!(output.is_empty());
}

#[test]
fn preservation_index_propagates_typed_limit_allocation_and_io_errors() {
    let mut package = SourceBackedPackage::from_vec(
        archive_bytes(root_relationships(), b"<before/>", true),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        None,
    )
    .unwrap();
    package.limits = SourceLimits::builder()
        .max_archive_metadata_bytes(1)
        .unwrap()
        .build()
        .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_changed_overlays_with_appended(&mut output, &[], &[]),
        Err(OpcError::Limit {
            kind: SourceLimitKind::ArchiveMetadataBytes,
            maximum: 1,
            ..
        })
    ));
    assert!(output.is_empty());

    let source = Arc::new(OverReportingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        true,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    source.overreport.store(true, Ordering::SeqCst);
    output.clear();
    assert!(matches!(
        package.write_changed_overlays_with_appended(&mut output, &[], &[]),
        Err(OpcError::IoError(error))
            if error.kind() == std::io::ErrorKind::InvalidData
    ));
    assert!(output.is_empty());

    let source = Vec::<u8>::new()
        .try_reserve_exact(usize::MAX)
        .expect_err("the test allocation must fail without allocating");
    assert!(matches!(
        map_preservation_error(soapberry_zip::Error::from(
            soapberry_zip::ErrorKind::Allocation {
                resource: "preservation-index test allocation",
                source,
            },
        )),
        OpcError::Allocation {
            resource: "preservation-index test allocation",
            ..
        }
    ));
}

#[test]
fn one_part_overlay_rejects_invalid_xml_and_signed_changes_before_output() {
    let target = PackURI::new("/word/document.xml").unwrap();
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlay_to_stream(&mut output, &target, b"<broken".to_vec()),
        Err(OpcError::XmlPublication { .. })
    ));
    assert!(output.is_empty());

    let signed_bytes = signed_archive(b"<signed/>");
    let signed = Arc::new(CountingSource::new(signed_bytes.clone()));
    let package = SourceBackedPackage::from_read_at(
        signed,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    assert!(matches!(
        package.write_part_overlay_to_stream(&mut output, &target, b"<changed/>".to_vec()),
        Err(OpcError::SignedSourceRequiresExplicitPolicy)
    ));
    assert!(output.is_empty());

    let signed = Arc::new(CountingSource::new(signed_bytes.clone()));
    let package = SourceBackedPackage::from_read_at(
        signed,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    package
        .write_part_overlay_to_stream(&mut output, &target, b"<signed/>".to_vec())
        .unwrap();
    assert_eq!(output, signed_bytes);
}

#[test]
fn one_part_overlay_enforces_replacement_limits_without_output() {
    let limits = SourceLimits::builder()
        .max_part_bytes(10)
        .unwrap()
        .max_total_part_bytes(19)
        .unwrap()
        .build()
        .unwrap();
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        false,
    )));
    let package = SourceBackedPackage::from_read_at(
        source,
        limits,
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let target = PackURI::new("/word/document.xml").unwrap();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlay_to_stream(&mut output, &target, vec![b'x'; 11]),
        Err(OpcError::Limit {
            kind: SourceLimitKind::PartBytes,
            observed: 11,
            maximum: 10
        })
    ));
    assert!(output.is_empty());
}

#[test]
fn one_part_overlay_reports_source_changes_before_and_during_output() {
    let target = PackURI::new("/word/document.xml").unwrap();
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        true,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    source.changed();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec()),
        Err(OpcError::SourceChanged { .. })
    ));
    assert!(output.is_empty());

    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        true,
    )));
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut sink = MutatingSink {
        source,
        bytes: Vec::new(),
        change_after: 1,
        changed: false,
    };
    let error = package
        .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
        .unwrap_err();
    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert!(written > 0);
            assert!(matches!(*source, OpcError::SourceChanged { .. }));
        },
        other => panic!("unexpected source-change error: {other:?}"),
    }
}

#[test]
fn one_part_overlay_bounds_writes_and_reports_partial_sink_failure() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        true,
    )));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let target = PackURI::new("/word/document.xml").unwrap();
    let mut sink = BoundedFailingSink {
        accepted: 0,
        limit: 100,
        largest_write: 0,
    };
    let error = package
        .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
        .unwrap_err();
    match error {
        OpcError::IncompleteOutput { written, .. } => assert_eq!(written, 100),
        other => panic!("unexpected sink error: {other:?}"),
    }
    assert!(sink.largest_write <= SOURCE_PUBLICATION_CHUNK_BYTES);
}

#[test]
fn managed_publication_checks_before_output_and_between_copy_chunks() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let target = PackURI::new("/word/document.xml").unwrap();

    let (budget, cancellation_source, context) =
        managed_context_with_cancellation(source_bytes.len() as u64);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    cancellation_source.cancel();
    let mut output = Vec::new();
    assert!(matches!(
        package.write_part_overlay_to_stream(&mut output, &target, b"<before/>".to_vec()),
        Err(OpcError::Cancelled)
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::Memory), 0);

    let (budget, cancellation_source, context) =
        managed_context_with_cancellation(source_bytes.len() as u64);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut sink = CancelAfterWriteSink {
        cancellation_source,
        bytes: Vec::new(),
        cancelled: false,
    };
    let error = package
        .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
        .unwrap_err();
    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert!(written > 0);
            assert!(matches!(*source, OpcError::Cancelled));
        },
        other => panic!("unexpected managed cancellation error: {other:?}"),
    }
    assert!(!sink.bytes.is_empty());
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::OutputBytes), sink.bytes.len() as u64);
}

#[test]
fn managed_exact_noop_overlay_charges_physical_output_bytes() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let target = PackURI::new("/word/document.xml").unwrap();
    let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut output = Vec::new();
    package
        .write_part_overlay_to_stream(&mut output, &target, b"<before/>".to_vec())
        .unwrap();

    assert_eq!(output, source_bytes);
    assert_eq!(budget.used(Resource::OutputBytes), output.len() as u64);
}

#[test]
fn managed_exact_output_limit_accepts_exact_small_artifact() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let (budget, _cancellation_source, context) =
        managed_context_with_output(source_bytes.len() as u64);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut output = Vec::new();
    artifact.write_to(&mut output).unwrap();

    assert_eq!(output, source_bytes);
    assert_eq!(
        budget.used(Resource::OutputBytes),
        source_bytes.len() as u64
    );
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.budget_reservation_failures, 0);
    assert_eq!(
        diagnostics.budget_output_bytes_used,
        source_bytes.len() as u64
    );
    assert_eq!(
        diagnostics.budget_output_bytes_limit,
        Some(source_bytes.len() as u64)
    );
}

#[test]
fn managed_exact_output_limit_one_under_refuses_before_small_artifact_output() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let (budget, _cancellation_source, context) =
        managed_context_with_output((source_bytes.len() as u64).saturating_sub(1));
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut output = Vec::new();
    let error = artifact.write_to(&mut output).unwrap_err();

    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::OutputBytes
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::OutputBytes), 0);
    let diagnostics = package.cache_diagnostics();
    assert_eq!(diagnostics.budget_reservation_failures, 1);
    assert_eq!(diagnostics.budget_output_bytes_used, 0);
}

#[test]
fn managed_exact_publication_refuses_memory_window_before_output_and_releases() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);

    for with_accounting in [false, true] {
        let (budget, _cancellation_source, context) = managed_context_with_all_resources(
            SOURCE_PUBLICATION_CHUNK_BYTES as u64 - 1,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        );
        let package = SourceBackedPackage::from_read_at(
            Arc::new(CountingSource::new(source_bytes.clone())),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            Some(context),
        )
        .unwrap();
        let artifact = package.source_artifact();
        let mut output = Vec::new();
        let error = if with_accounting {
            let mut accounting = OpcOperationAccounting::default();
            artifact
                .write_to_with_accounting(&mut output, &mut accounting)
                .unwrap_err()
        } else {
            artifact.write_to(&mut output).unwrap_err()
        };

        assert!(matches!(
            error,
            OpcError::Execution(ExecutionError::ResourceLimit(limit))
                if limit.resource == Resource::Memory
        ));
        assert!(output.is_empty());
        assert_eq!(budget.used(Resource::Memory), 0);
    }
}

#[test]
fn managed_large_exact_publication_refuses_mid_stream_with_exact_partial_charge() {
    let source_bytes = large_archive_bytes(root_relationships(), b"<before/>", 150_000);
    assert!(source_bytes.len() > SOURCE_PUBLICATION_CHUNK_BYTES * 2);
    let output_limit = SOURCE_PUBLICATION_CHUNK_BYTES as u64 + 100;
    let (budget, _cancellation_source, context) = managed_context_with_output(output_limit);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut output = Vec::new();
    let error = artifact.write_to(&mut output).unwrap_err();

    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert_eq!(written, SOURCE_PUBLICATION_CHUNK_BYTES as u64);
            match *source {
                OpcError::Execution(ExecutionError::ResourceLimit(limit)) => {
                    assert_eq!(limit.resource, Resource::OutputBytes);
                },
                other => panic!("unexpected mid-stream source error: {other:?}"),
            }
        },
        other => panic!("unexpected mid-stream output error: {other:?}"),
    }
    assert_eq!(output.len(), SOURCE_PUBLICATION_CHUNK_BYTES);
    assert_eq!(
        budget.used(Resource::OutputBytes),
        SOURCE_PUBLICATION_CHUNK_BYTES as u64
    );
    assert_eq!(package.cache_diagnostics().budget_reservation_failures, 1);
}

#[test]
fn managed_output_charge_reaches_hierarchical_parent_and_child_budgets() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let output_bytes = source_bytes.len() as u64;
    let publication_memory = SOURCE_PUBLICATION_CHUNK_BYTES as u64 + 4096;
    let root = Budget::root(
        "opc-source-output-root",
        Limits::new(
            publication_memory,
            u64::MAX,
            output_bytes,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let child = root.child(
        "opc-source-output-child",
        Limits::new(
            publication_memory,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(publication_memory).unwrap(),
        0,
    )
    .unwrap();
    let context = ExecutionContext::new(child.clone(), cancellation, execution_limits);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut output = Vec::new();
    artifact.write_to(&mut output).unwrap();

    assert_eq!(output, source_bytes);
    assert_eq!(root.used(Resource::OutputBytes), output_bytes);
    assert_eq!(child.used(Resource::OutputBytes), output_bytes);
    drop(cancellation_source);
}

#[test]
fn managed_cumulative_output_retry_exhaustion_refuses_second_exact_publication() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let output_bytes = source_bytes.len() as u64;
    let (budget, _cancellation_source, context) = managed_context_with_output(output_bytes);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut first = Vec::new();
    artifact.write_to(&mut first).unwrap();
    let mut retry = Vec::new();
    let error = artifact.write_to(&mut retry).unwrap_err();

    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::OutputBytes
    ));
    assert_eq!(first, source_bytes);
    assert!(retry.is_empty());
    assert_eq!(budget.used(Resource::OutputBytes), output_bytes);
    assert_eq!(package.cache_diagnostics().budget_reservation_failures, 1);
}

#[test]
fn managed_flush_failure_reports_exact_output_charge_as_incomplete() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", false);
    let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut sink = FlushFailingSink { bytes: Vec::new() };
    let error = artifact.write_to(&mut sink).unwrap_err();

    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert_eq!(written, source_bytes.len() as u64);
            assert!(matches!(
                *source,
                OpcError::IoError(error)
                    if error.kind() == std::io::ErrorKind::Other
            ));
        },
        other => panic!("unexpected flush error: {other:?}"),
    }
    assert_eq!(sink.bytes, source_bytes);
    assert_eq!(
        budget.used(Resource::OutputBytes),
        source_bytes.len() as u64
    );
}

#[test]
fn managed_changed_overlay_charges_exact_accepted_output_bytes() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let target = PackURI::new("/word/document.xml").unwrap();
    let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut output = Vec::new();
    package
        .write_part_overlay_to_stream(&mut output, &target, b"<after/>".to_vec())
        .unwrap();

    let reopened = OpcPackage::from_bytes(&output, SourceLimits::default()).unwrap();
    assert_eq!(reopened.get_part(&target).unwrap().blob(), b"<after/>");
    assert_eq!(budget.used(Resource::OutputBytes), output.len() as u64);
}

#[test]
fn managed_exact_noop_refuses_the_first_output_write_at_zero_limit() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        true,
    )));
    let target = PackURI::new("/word/document.xml").unwrap();
    let (budget, _cancellation_source, context) = managed_context_with_output(0);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_part_overlay_to_stream(&mut output, &target, b"<before/>".to_vec())
        .unwrap_err();

    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::OutputBytes
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::OutputBytes), 0);
}

#[test]
fn managed_partial_sink_failure_commits_only_accepted_output_bytes() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        true,
    )));
    let target = PackURI::new("/word/document.xml").unwrap();
    let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut sink = BoundedFailingSink {
        accepted: 0,
        limit: 100,
        largest_write: 0,
    };
    let error = package
        .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
        .unwrap_err();

    match error {
        OpcError::IncompleteOutput { written, .. } => assert_eq!(written, 100),
        other => panic!("unexpected sink error: {other:?}"),
    }
    assert_eq!(budget.used(Resource::OutputBytes), 100);
}

#[test]
fn managed_source_change_after_accepted_output_preserves_partial_accounting() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        true,
    )));
    let target = PackURI::new("/word/document.xml").unwrap();
    let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        source.clone(),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut sink = MutatingSink {
        source,
        bytes: Vec::new(),
        change_after: 0,
        changed: false,
    };
    let error = package
        .write_part_overlay_to_stream(&mut sink, &target, b"<after/>".to_vec())
        .unwrap_err();

    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert_eq!(written, sink.bytes.len() as u64);
            assert!(matches!(*source, OpcError::SourceChanged { .. }));
        },
        other => panic!("unexpected source-change error: {other:?}"),
    }
    assert_eq!(budget.used(Resource::OutputBytes), sink.bytes.len() as u64);
}

#[test]
fn managed_short_interrupted_source_copy_charges_only_accepted_bytes() {
    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut sink = ShortInterruptedSink {
        bytes: Vec::new(),
        maximum_write: 7,
        interrupted: false,
    };
    artifact.write_to(&mut sink).unwrap();

    assert_eq!(sink.bytes, source_bytes);
    assert_eq!(
        budget.used(Resource::OutputBytes),
        source_bytes.len() as u64
    );
    let diagnostics = package.cache_diagnostics();
    assert_eq!(
        diagnostics.budget_output_bytes_used,
        source_bytes.len() as u64
    );
    assert_eq!(diagnostics.budget_output_bytes_limit, Some(u64::MAX));
}

#[test]
fn managed_overreporting_sink_does_not_commit_output_budget() {
    let source = Arc::new(CountingSource::new(archive_bytes(
        root_relationships(),
        b"<before/>",
        false,
    )));
    let (budget, _cancellation_source, context) = managed_context_with_output(u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let artifact = package.source_artifact();
    let mut sink = OverReportingSink {
        calls: 0,
        accepted: 0,
    };

    assert!(matches!(
        artifact.write_to(&mut sink),
        Err(OpcError::IoError(error)) if error.kind() == std::io::ErrorKind::InvalidData
    ));
    assert_eq!(sink.calls, 1);
    assert_eq!(sink.accepted, 0);
    assert_eq!(budget.used(Resource::OutputBytes), 0);
}

#[test]
fn managed_oversized_bypass_releases_comparison_memory_before_exact_publication() {
    struct MemoryObservingSink<'budget> {
        budget: &'budget Budget,
        bytes: Vec<u8>,
        first_memory: Option<u64>,
    }

    impl Write for MemoryObservingSink<'_> {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.first_memory
                .get_or_insert(self.budget.used(Resource::Memory));
            self.bytes.extend_from_slice(bytes);
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }

    let source_bytes = archive_bytes(root_relationships(), b"<before/>", true);
    let target = PackURI::new("/word/document.xml").unwrap();
    let tight_memory = SOURCE_PUBLICATION_CHUNK_BYTES as u64 + b"<before/>".len() as u64;
    let (budget, _cancellation_source, context) =
        managed_context_with_all_resources(tight_memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX);
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(source_bytes.clone())),
        SourceLimits::default(),
        SourceCacheLimits::new(1, 1).unwrap(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut sink = MemoryObservingSink {
        budget: &budget,
        bytes: Vec::new(),
        first_memory: None,
    };

    package
        .write_part_overlay_to_stream(&mut sink, &target, b"<before/>".to_vec())
        .unwrap();

    assert_eq!(sink.bytes, source_bytes);
    assert_eq!(
        sink.first_memory,
        Some(SOURCE_PUBLICATION_CHUNK_BYTES as u64),
        "comparison payload reservation must be released before exact publication"
    );
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[cfg(feature = "performance-diagnostics")]
fn assert_balanced_diagnostic_events(events: &[DiagnosticSnapshot]) {
    assert_eq!(events.len() % 2, 0, "event stream must contain pairs");
    for pair in events.as_chunks::<2>().0 {
        match (&pair[0], &pair[1]) {
            (
                DiagnosticSnapshot::Started { operation: started },
                DiagnosticSnapshot::Finished {
                    operation: finished,
                },
            ) => assert_eq!(started, finished, "lock operation changed in pair"),
            _ => panic!("lock events must alternate Started and Finished"),
        }
    }
}

#[cfg(feature = "performance-diagnostics")]
fn count_started_operation(events: &[DiagnosticSnapshot], wanted: Operation) -> usize {
    events
        .iter()
        .filter(|event| {
            matches!(
                event,
                DiagnosticSnapshot::Started { operation } if *operation == wanted
            )
        })
        .count()
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_uncontended_hit_reports_only_cache_lock() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"diagnostic hit",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    let mut cold_events = Vec::new();
    part.data_with_observer(|event| cold_events.push(event))
        .unwrap();
    assert_balanced_diagnostic_events(&cold_events);
    assert!(count_started_operation(&cold_events, Operation::Cache(CacheState::Lock)) > 0);
    assert!(count_started_operation(&cold_events, Operation::Flight(FlightState::Lock)) > 0);

    let mut hit_events = Vec::new();
    part.data_with_observer(|event| hit_events.push(event))
        .unwrap();
    assert_eq!(
        hit_events,
        vec![
            DiagnosticSnapshot::Started {
                operation: Operation::Cache(CacheState::Lock),
            },
            DiagnosticSnapshot::Finished {
                operation: Operation::Cache(CacheState::Lock),
            },
        ]
    );
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_same_part_loader_and_waiter_are_balanced() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"diagnostic waiter",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let hook = Arc::new(TestPublicationHook::new());
    let _hook_guard = package
        .cache
        .install_test_publication_hook(Arc::clone(&hook));

    let (loader_events, waiter_events) = std::thread::scope(|scope| {
        let package = &package;
        let loader = scope.spawn(move || {
            let mut events = Vec::new();
            let result = package
                .main_document_part()
                .unwrap()
                .data_with_observer(|event| events.push(event));
            (result, events)
        });
        hook.wait_until_entered();
        let waiter = scope.spawn(move || {
            let mut events = Vec::new();
            let result = package
                .main_document_part()
                .unwrap()
                .data_with_observer(|event| events.push(event));
            (result, events)
        });
        hook.wait_until_waiter_joined();
        hook.release();
        let (loader_result, loader_events) = loader.join().unwrap();
        let (waiter_result, waiter_events) = waiter.join().unwrap();
        assert!(loader_result.is_ok());
        assert!(waiter_result.is_ok());
        (loader_events, waiter_events)
    });

    assert_balanced_diagnostic_events(&loader_events);
    assert_balanced_diagnostic_events(&waiter_events);
    assert_eq!(
        count_started_operation(&waiter_events, Operation::Cache(CacheState::Lock)),
        1
    );
    // The waiter blocks in Condvar::wait_timeout after this one direct
    // flight-state lock. Reacquisition and timed waiting must be silent.
    assert_eq!(
        count_started_operation(&waiter_events, Operation::Flight(FlightState::Lock)),
        1
    );
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_cache_and_flight_scopes_are_reported() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"diagnostic scopes",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let mut idle_events = Vec::new();
    let _ = package.cache_diagnostics_with_observer(|event| idle_events.push(event));
    assert_eq!(
        idle_events,
        vec![
            DiagnosticSnapshot::Started {
                operation: Operation::Cache(CacheState::Lock),
            },
            DiagnosticSnapshot::Finished {
                operation: Operation::Cache(CacheState::Lock),
            },
        ]
    );

    let entry_id = package.catalog.parts[0].entry_id;
    let flight = Arc::new(LoadFlight::new(None, None, None));
    package
        .cache
        .state
        .lock()
        .unwrap()
        .flights
        .insert(entry_id, Arc::clone(&flight));
    let mut active_events = Vec::new();
    let diagnostics = package.cache_diagnostics_with_observer(|event| active_events.push(event));
    assert_eq!(diagnostics.in_flight_loads, 1);
    assert_balanced_diagnostic_events(&active_events);
    assert_eq!(
        count_started_operation(&active_events, Operation::Cache(CacheState::Lock)),
        1
    );
    assert_eq!(
        count_started_operation(&active_events, Operation::Flight(FlightState::Lock)),
        1
    );
    package.cache.complete_failure(entry_id, &flight);
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_poison_preserves_error_and_event_pair() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"diagnostic poison",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    std::thread::scope(|scope| {
        let cache = &package.cache;
        let join = scope.spawn(move || {
            let _state = cache.state.lock().unwrap();
            panic!("test diagnostic cache-state poison");
        });
        assert!(join.join().is_err());
    });

    let mut events = Vec::new();
    assert_eq!(
        package.try_cache_diagnostics_with_observer(|event| events.push(event)),
        Err(SourceCacheDiagnosticsError::StatePoisoned)
    );
    assert_eq!(
        events,
        vec![
            DiagnosticSnapshot::Started {
                operation: Operation::Cache(CacheState::Lock),
            },
            DiagnosticSnapshot::Finished {
                operation: Operation::Cache(CacheState::Lock),
            },
        ]
    );
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_failure_retry_and_cancellation_are_balanced() {
    const DOCUMENT: &[u8] = b"diagnostic failure and retry";
    let mut bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    bytes[payload_offset] ^= 0xff;
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(bytes)),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    for _ in 0..2 {
        let mut events = Vec::new();
        assert!(part.data_with_observer(|event| events.push(event)).is_err());
        assert_balanced_diagnostic_events(&events);
        assert!(count_started_operation(&events, Operation::Flight(FlightState::Lock)) > 0);
    }

    let (budget, cancellation_source, context) =
        managed_context_with_cancellation(DOCUMENT.len() as u64);
    let source_bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = source_bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let source = Arc::new(CancelDuringPayloadSource::new(
        source_bytes,
        payload_offset,
        cancellation_source,
    ));
    let package = SourceBackedPackage::from_read_at(
        source,
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        Some(context),
    )
    .unwrap();
    let mut events = Vec::new();
    assert!(matches!(
        package
            .main_document_part()
            .unwrap()
            .data_with_observer(|event| events.push(event)),
        Err(OpcError::Cancelled)
    ));
    assert_balanced_diagnostic_events(&events);
    assert!(count_started_operation(&events, Operation::Flight(FlightState::Lock)) > 0);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_started_panic_has_no_finished_and_cache_recovers() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"started panic",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    let mut started = false;
    let mut events = Vec::new();
    let panic = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let _ = part.data_with_observer(|event| {
            events.push(event);
            if matches!(event, DiagnosticSnapshot::Started { .. }) {
                started = true;
                panic!("test Started callback panic");
            }
        });
    }));
    assert!(panic.is_err());
    assert!(started);
    assert_eq!(events.len(), 1);
    assert!(matches!(events[0], DiagnosticSnapshot::Started { .. }));

    // Started runs before lock acquisition. The observer panic is deferred
    // while the load completes, later notifications are suppressed, and
    // no cache mutex is poisoned; the ordinary path can immediately read
    // the committed payload.
    assert_eq!(part.data().unwrap().as_bytes(), b"started panic");
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_started_flight_panic_does_not_strand_waiter() {
    const DOCUMENT: &[u8] = b"started flight panic with waiter";
    let bytes = archive_bytes(root_relationships(), DOCUMENT, false);
    let payload_offset = bytes
        .windows(DOCUMENT.len())
        .position(|window| window == DOCUMENT)
        .unwrap();
    let source = Arc::new(GatedPayloadSource::new(bytes, payload_offset));
    let package = Arc::new(
        SourceBackedPackage::from_read_at(
            source.clone(),
            SourceLimits::default(),
            SourceCacheLimits::default(),
            SourceReadPolicy::exact(),
            None,
        )
        .expect("gated source package should open"),
    );

    let (loader_tx, loader_rx) = std::sync::mpsc::channel();
    let loader_package = Arc::clone(&package);
    let loader = std::thread::spawn(move || {
        let mut events = Vec::new();
        let panicked = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            loader_package
                .main_document_part()
                .unwrap()
                .data_with_observer(|event| {
                    events.push(event);
                    if matches!(
                        event,
                        DiagnosticSnapshot::Started {
                            operation: Operation::Flight(FlightState::Lock),
                        }
                    ) {
                        panic!("test Started Flight callback panic");
                    }
                })
                .unwrap();
        }))
        .is_err();
        loader_tx
            .send((panicked, events))
            .expect("loader result receiver should remain available");
    });

    // The source gate proves that the loader has already inserted its
    // flight before the observer's first Flight Started callback. This
    // avoids relying on sleeps and keeps every wait in this regression
    // test bounded.
    if !source.wait_until_payload_read(Duration::from_secs(2)) {
        source.release_payload_read();
        panic!("loader did not reach the gated payload read");
    }
    assert!(source.payload_reads.load(Ordering::Acquire) > 0);

    let (waiter_joined_tx, waiter_joined_rx) = std::sync::mpsc::channel();
    let (waiter_tx, waiter_rx) = std::sync::mpsc::channel();
    let waiter_package = Arc::clone(&package);
    let waiter = std::thread::spawn(move || {
        let mut events = Vec::new();
        let result = waiter_package
            .main_document_part()
            .unwrap()
            .data_with_observer(|event| {
                if matches!(
                    event,
                    DiagnosticSnapshot::Finished {
                        operation: Operation::Flight(FlightState::Lock),
                    }
                ) {
                    waiter_joined_tx
                        .send(())
                        .expect("waiter-joined receiver should remain available");
                }
                events.push(event);
            })
            .map(|data| data.as_bytes().to_vec());
        waiter_tx
            .send((result, events))
            .expect("waiter result receiver should remain available");
    });

    // The waiter reports its direct flight-lock Finished boundary while
    // that guard is live and immediately before entering
    // Condvar::wait_timeout. Only then release the loader, guaranteeing
    // that the waiter exercises the same-flight path rather than a cache
    // hit after publication.
    if waiter_joined_rx
        .recv_timeout(Duration::from_secs(2))
        .is_err()
    {
        source.release_payload_read();
        panic!("waiter did not join the active load flight");
    }
    source.release_payload_read();

    let loader_message = loader_rx
        .recv_timeout(Duration::from_secs(3))
        .expect("loader must finish after observer panic");
    let waiter_message = waiter_rx
        .recv_timeout(Duration::from_secs(3))
        .expect("waiter must be notified after loader cleanup");
    loader.join().expect("loader thread should exit normally");
    waiter.join().expect("waiter thread should exit normally");

    assert!(loader_message.0, "loader observer panic must propagate");
    assert_eq!(loader_message.1.len(), 3);
    assert_eq!(
        count_started_operation(&loader_message.1, Operation::Flight(FlightState::Lock)),
        1
    );
    assert!(!loader_message.1.iter().any(|event| matches!(
        event,
        DiagnosticSnapshot::Finished {
            operation: Operation::Flight(FlightState::Lock),
        }
    )));

    assert_eq!(waiter_message.0.unwrap(), DOCUMENT);
    assert_balanced_diagnostic_events(&waiter_message.1);
    assert_eq!(
        count_started_operation(&waiter_message.1, Operation::Cache(CacheState::Lock)),
        1
    );
    // The wait itself and its mutex reacquisitions remain outside the
    // direct-lock event seam.
    assert_eq!(
        count_started_operation(&waiter_message.1, Operation::Flight(FlightState::Lock)),
        1
    );
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    assert!(package.try_cache_diagnostics().is_ok());
    assert_eq!(
        package
            .main_document_part()
            .unwrap()
            .data()
            .unwrap()
            .as_bytes(),
        DOCUMENT
    );
    assert_eq!(package.cache_diagnostics().successful_loads, 1);
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_started_nested_publication_flight_panic_recovers() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"nested publication panic",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    let mut flight_starts = 0;
    let mut events = Vec::new();
    let panic = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let _ = part.data_with_observer(|event| {
            if matches!(
                event,
                DiagnosticSnapshot::Started {
                    operation: Operation::Flight(FlightState::Lock),
                }
            ) {
                flight_starts += 1;
                // The third Flight lock is acquired by publication while
                // the cache-state guard is still held. A direct callback
                // unwind here would poison that guard and strand the
                // loader; the observer must defer it until cleanup.
                if flight_starts == 3 {
                    panic!("test nested publication Started Flight panic");
                }
            }
            events.push(event);
        });
    }));
    assert!(panic.is_err());
    assert_eq!(flight_starts, 3);
    // The third Started callback panics before the caller-owned log can
    // record that event. Its invocation is counted above, while its
    // event and every later notification are intentionally absent.
    assert_eq!(events.len(), 10);
    assert_eq!(
        count_started_operation(&events, Operation::Flight(FlightState::Lock)),
        2
    );
    assert!(matches!(
        events.last(),
        Some(DiagnosticSnapshot::Finished {
            operation: Operation::Cache(CacheState::Lock),
        })
    ));

    // The nested panic was caught before the flight mutex lock and the
    // operation completed publication before resuming it. Both locks are
    // therefore usable and the committed value is immediately a hit.
    assert!(package.try_cache_diagnostics().is_ok());
    assert_eq!(package.cache_diagnostics().in_flight_loads, 0);
    assert_eq!(part.data().unwrap().as_bytes(), b"nested publication panic");
    assert_eq!(package.cache_diagnostics().successful_loads, 1);
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_finished_panic_drops_guard_and_cache_recovers() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"finished panic",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    let panic = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let _ = part.data_with_observer(|event| {
            if matches!(
                event,
                DiagnosticSnapshot::Finished {
                    operation: Operation::Cache(CacheState::Lock),
                }
            ) {
                panic!("test Finished callback panic");
            }
        });
    }));
    assert!(panic.is_err());

    // The helper catches the panic while its guard is live; the public
    // entry point resumes only after dropping that guard. The next load
    // is therefore uncontended and proves observer panic did not poison
    // it.
    let mut events = Vec::new();
    assert_eq!(
        part.data_with_observer(|event| events.push(event))
            .unwrap()
            .as_bytes(),
        b"finished panic"
    );
    assert_balanced_diagnostic_events(&events);
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_finished_panic_drops_flight_guard() {
    let flight = LoadFlight::new(None, None, None);
    let panic = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let mut callback = |event| {
            if matches!(
                event,
                DiagnosticSnapshot::Finished {
                    operation: Operation::Flight(FlightState::Lock),
                }
            ) {
                panic!("test flight Finished callback panic");
            }
        };
        let mut observer = CallbackDiagnosticObserver::new(&mut callback);
        let _ = flight.reservation_with_observer(&mut observer);
        observer.resume_if_panicked();
    }));
    assert!(panic.is_err());

    // The flight mutex is still usable after the panic was resumed.
    assert!(flight.state.lock().is_ok());
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_finished_panic_does_not_abandon_loader_flight() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"flight panic recovery",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();
    let part = package.main_document_part().unwrap();
    let panic = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let _ = part.data_with_observer(|event| {
            if matches!(
                event,
                DiagnosticSnapshot::Finished {
                    operation: Operation::Flight(FlightState::Lock),
                }
            ) {
                panic!("test loader-flight Finished callback panic");
            }
        });
    }));
    assert!(panic.is_err());

    // Finished panic is deferred until the whole read unwinds cleanly;
    // otherwise an abandoned flight would make the next same-Part read
    // wait forever. The package remains usable and the payload committed.
    assert_eq!(part.data().unwrap().as_bytes(), b"flight panic recovery");
    assert_eq!(package.cache_diagnostics().successful_loads, 1);
}

#[cfg(feature = "performance-diagnostics")]
#[test]
fn performance_diagnostics_reentrancy_contract_refuses_before_lock() {
    let package = SourceBackedPackage::from_read_at(
        Arc::new(CountingSource::new(archive_bytes(
            root_relationships(),
            b"reentrancy contract",
            false,
        ))),
        SourceLimits::default(),
        SourceCacheLimits::default(),
        SourceReadPolicy::exact(),
        None,
    )
    .unwrap();

    // Re-entering an observed operation from Finished would deadlock by
    // design because that callback runs with the corresponding guard
    // live. A caller-owned callback can apply a bounded refusal policy;
    // this test models that refusal instead of intentionally hanging.
    let panic = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        let _ = package.cache_diagnostics_with_observer(|event| {
            if matches!(event, DiagnosticSnapshot::Started { .. }) {
                panic!("observer contract refuses reentrant lock acquisition");
            }
        });
    }));
    assert!(panic.is_err());
    assert!(package.try_cache_diagnostics().is_ok());
}
