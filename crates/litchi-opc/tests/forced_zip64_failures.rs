#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "focused ZIP64 failure assertions intentionally fail on fixture errors"
)]

//! Failure-atomicity coverage for Python's small forced-ZIP64 OPC shape.
//!
//! The fixture is generated independently with `zipfile.ZipFile.open(...,
//! force_zip64=true)` and retained with the 0416 evidence corpus.  Its local
//! records reserve ZIP64 sizes while its central records retain ordinary
//! sizes; the descriptor carries the final 64-bit values.  The tests below
//! keep the publication sink caller-owned and exercise only public
//! `SourceBackedPackage` APIs.

use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc,
    atomic::{AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::{OpcError, PackURI, ReadLimits, ReadResource, SourceBackedPackage};

const TARGET_URI: &str = "/word/document.xml";
const SOURCE_FIXTURE: &[u8] = include_bytes!(
    "../../../docs/performance/results/change-0416/corpus/opc-local-only-signed.zip"
);
const ZIP32_MAX: u32 = u32::MAX;
const ZIP_LOCAL_SIGNATURE: u32 = 0x0403_4b50;
const ZIP_CENTRAL_SIGNATURE: u32 = 0x0201_4b50;
const ZIP_DESCRIPTOR_SIGNATURE: u32 = 0x0807_4b50;
const ZIP_EOCD_SIGNATURE: u32 = 0x0605_4b50;

const REPLACEMENT: &[u8] =
    b"<document>replacement ZIP64 overlay fixture with enough bytes for a limit</document>";

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).unwrap()
}

fn u16_at(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(bytes[offset..offset + 2].try_into().unwrap())
}

fn u32_at(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn u64_at(bytes: &[u8], offset: usize) -> u64 {
    u64::from_le_bytes(bytes[offset..offset + 8].try_into().unwrap())
}

fn eocd_offset(bytes: &[u8]) -> usize {
    bytes
        .windows(4)
        .rposition(|window| window == ZIP_EOCD_SIGNATURE.to_le_bytes())
        .expect("the independent OPC fixture must have a ZIP32 EOCD")
}

fn central_metadata_bytes(bytes: &[u8]) -> u64 {
    let eocd = eocd_offset(bytes);
    let count = usize::from(u16_at(bytes, eocd + 10));
    let mut cursor = usize::try_from(u32_at(bytes, eocd + 16)).unwrap();
    let mut total = 0_u64;
    for _ in 0..count {
        assert_eq!(u32_at(bytes, cursor), ZIP_CENTRAL_SIGNATURE);
        let name_len = usize::from(u16_at(bytes, cursor + 28));
        let extra_len = usize::from(u16_at(bytes, cursor + 30));
        let comment_len = usize::from(u16_at(bytes, cursor + 32));
        total += u64::try_from(46 + name_len + extra_len + comment_len).unwrap();
        cursor += 46 + name_len + extra_len + comment_len;
    }
    assert_eq!(cursor, eocd);
    total
}

/// Verify the external fixture's exact local-only ZIP64 grammar before any
/// public OPC operation is attempted.  This keeps the integration cases tied
/// to the producer shape rather than a merely equivalent in-memory archive.
fn assert_local_only_fixture_shape(bytes: &[u8]) {
    let eocd = eocd_offset(bytes);
    let count = usize::from(u16_at(bytes, eocd + 10));
    let mut central = usize::try_from(u32_at(bytes, eocd + 16)).unwrap();
    for _ in 0..count {
        assert_eq!(u32_at(bytes, central), ZIP_CENTRAL_SIGNATURE);
        assert_eq!(u16_at(bytes, central + 6), 45);
        assert_eq!(u16_at(bytes, central + 30), 0);
        assert_ne!(u32_at(bytes, central + 20), ZIP32_MAX);
        assert_ne!(u32_at(bytes, central + 24), ZIP32_MAX);

        let name_len = usize::from(u16_at(bytes, central + 28));
        let extra_len = usize::from(u16_at(bytes, central + 30));
        let comment_len = usize::from(u16_at(bytes, central + 32));
        let compressed = usize::try_from(u32_at(bytes, central + 20)).unwrap();
        let uncompressed = u64::from(u32_at(bytes, central + 24));
        let local = usize::try_from(u32_at(bytes, central + 42)).unwrap();
        assert_eq!(u32_at(bytes, local), ZIP_LOCAL_SIGNATURE);
        assert_eq!(u16_at(bytes, local + 4), 45);
        assert_eq!(u16_at(bytes, local + 6) & 0x0008, 0x0008);
        assert_eq!(u32_at(bytes, local + 18), ZIP32_MAX);
        assert_eq!(u32_at(bytes, local + 22), ZIP32_MAX);

        let local_name_len = usize::from(u16_at(bytes, local + 26));
        let local_extra_len = usize::from(u16_at(bytes, local + 28));
        assert_eq!(local_name_len, name_len);
        assert_eq!(local_extra_len, 20);
        assert_eq!(
            &bytes[local + 30 + local_name_len..local + 34 + local_name_len],
            &[1, 0, 16, 0]
        );
        assert_eq!(
            &bytes[local + 34 + local_name_len..local + 50 + local_name_len],
            &[0; 16]
        );

        let payload_end = local + 30 + local_name_len + local_extra_len + compressed;
        assert_eq!(u32_at(bytes, payload_end), ZIP_DESCRIPTOR_SIGNATURE);
        assert_eq!(
            u64_at(bytes, payload_end + 8),
            u64::from(u32_at(bytes, central + 20))
        );
        assert_eq!(u64_at(bytes, payload_end + 16), uncompressed);
        central += 46 + name_len + extra_len + comment_len;
    }
    assert_eq!(central, eocd);
}

fn fixture() -> &'static [u8] {
    assert_local_only_fixture_shape(SOURCE_FIXTURE);
    SOURCE_FIXTURE
}

fn managed_context(output_bytes: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let memory_bytes = 64 * 1024 * 1024;
    let budget = Budget::root(
        "opc-forced-zip64-failure-matrix",
        Limits::new(
            memory_bytes,
            u64::MAX,
            output_bytes,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (cancellation_source, cancellation) = CancellationSource::pair();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(memory_bytes).unwrap(),
        0,
    )
    .unwrap();
    (
        budget.clone(),
        cancellation_source,
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

#[derive(Debug)]
struct VersionedSource {
    bytes: Arc<Vec<u8>>,
    revision: AtomicU64,
}

impl VersionedSource {
    fn new(bytes: &[u8]) -> Self {
        Self {
            bytes: Arc::new(bytes.to_vec()),
            revision: AtomicU64::new(0),
        }
    }

    fn bump(&self) {
        self.revision.fetch_add(1, Ordering::AcqRel);
    }

    fn bytes_equal(&self, expected: &[u8]) -> bool {
        self.bytes.as_slice() == expected
    }

    fn revision(&self) -> u64 {
        self.revision.load(Ordering::Acquire)
    }
}

impl ReadAt for VersionedSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.bytes.len())
            .map_err(|_| io::Error::other("fixture length does not fit in u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let start = usize::try_from(offset).unwrap_or(usize::MAX);
        let Some(source) = self.bytes.get(start..) else {
            return Ok(0);
        };
        let count = source.len().min(output.len());
        output[..count].copy_from_slice(&source[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x0416_5a49,
            self.revision.load(Ordering::Acquire),
        ))
    }
}

#[derive(Debug)]
struct PartialSink {
    bytes: Vec<u8>,
    remaining: usize,
}

impl PartialSink {
    fn new(remaining: usize) -> Self {
        Self {
            bytes: Vec::new(),
            remaining,
        }
    }
}

impl Write for PartialSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "partial sink"));
        }
        let accepted = self.remaining.min(input.len());
        self.bytes.extend_from_slice(&input[..accepted]);
        self.remaining -= accepted;
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct VersionBumpingSink {
    source: Arc<VersionedSource>,
    bytes: Vec<u8>,
    bumped: bool,
}

impl Write for VersionBumpingSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        self.bytes.extend_from_slice(input);
        if !self.bumped {
            self.bumped = true;
            self.source.bump();
        }
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn forced_zip64_metadata_limit_rejects_source_before_catalog_publication() {
    let source = fixture();
    let metadata_bytes = central_metadata_bytes(source);
    assert!(metadata_bytes > 0);
    let limits = ReadLimits::builder()
        .max_archive_metadata_bytes(metadata_bytes - 1)
        .unwrap()
        .build()
        .unwrap();

    let error = match SourceBackedPackage::from_vec_with_limits(source.to_vec(), limits) {
        Ok(_) => panic!("the one-under metadata limit must reject the fixture"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::ArchiveMetadataBytes,
            actual,
            maximum,
        } if actual == metadata_bytes && maximum == metadata_bytes - 1
    ));
}

#[test]
fn forced_zip64_overlay_part_limit_rejects_before_sink_output() {
    let source = fixture();
    let limits = ReadLimits::builder()
        // The source's largest ordinary Part is 61 bytes; the replacement is
        // deliberately larger so ingress succeeds and overlay preflight fails.
        .max_part_bytes(61)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_vec_with_limits(source.to_vec(), limits).unwrap();
    let mut output = Vec::new();
    let error = package
        .write_part_overlay_to_stream(&mut output, &pack(TARGET_URI), REPLACEMENT.to_vec())
        .unwrap_err();

    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::PartBytes,
            actual,
            maximum: 61,
        } if actual == REPLACEMENT.len() as u64
    ));
    assert!(output.is_empty());
}

#[test]
fn forced_zip64_managed_output_limit_rejects_before_first_sink_write() {
    let source = fixture();
    let (budget, _cancellation_source, context) = managed_context(0);
    let source_owner = Arc::new(VersionedSource::new(source));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source_owner.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let mut output = Vec::new();
    let error = package
        .write_part_overlay_to_stream(&mut output, &pack(TARGET_URI), REPLACEMENT.to_vec())
        .unwrap_err();

    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(limit))
            if limit.resource == Resource::OutputBytes
    ));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::OutputBytes), 0);
    assert!(source_owner.bytes_equal(fixture()));
    assert_eq!(source_owner.revision(), 0);
}

#[test]
fn forced_zip64_managed_cancellation_rejects_before_first_sink_write() {
    let source = fixture();
    let (budget, cancellation_source, context) = managed_context(u64::MAX);
    let source_owner = Arc::new(VersionedSource::new(source));
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source_owner.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    cancellation_source.cancel();

    let mut output = Vec::new();
    let error = package
        .write_part_overlay_to_stream(&mut output, &pack(TARGET_URI), REPLACEMENT.to_vec())
        .unwrap_err();

    assert!(matches!(error, OpcError::Cancelled));
    assert!(output.is_empty());
    assert_eq!(budget.used(Resource::OutputBytes), 0);
    assert!(source_owner.bytes_equal(fixture()));
    assert_eq!(source_owner.revision(), 0);
}

#[test]
fn forced_zip64_overlay_reports_partial_sequential_sink_failure() {
    let source = fixture();
    let source_owner = Arc::new(VersionedSource::new(source));
    let package = SourceBackedPackage::from_read_at(source_owner.clone()).unwrap();
    let mut sink = PartialSink::new(137);
    let error = package
        .write_part_overlay_to_stream(&mut sink, &pack(TARGET_URI), REPLACEMENT.to_vec())
        .unwrap_err();

    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert!(written > 0);
            assert_eq!(written as usize, sink.bytes.len());
            assert!(matches!(
                *source,
                OpcError::IoError(error) if error.kind() == io::ErrorKind::BrokenPipe
            ));
        },
        other => panic!("expected incomplete output after a partial sink failure: {other:?}"),
    }
    assert_eq!(sink.bytes.len(), 137);
    assert!(source_owner.bytes_equal(fixture()));
    assert_eq!(source_owner.revision(), 0);
}

#[test]
fn forced_zip64_overlay_reports_source_version_change_before_output_and_midstream() {
    let source = Arc::new(VersionedSource::new(fixture()));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    source.bump();
    let mut output = Vec::new();
    let error = package
        .write_part_overlay_to_stream(&mut output, &pack(TARGET_URI), REPLACEMENT.to_vec())
        .unwrap_err();
    assert!(matches!(error, OpcError::SourceChanged { .. }));
    assert!(output.is_empty());
    assert!(source.bytes_equal(fixture()));
    assert_eq!(source.revision(), 1);

    let source = Arc::new(VersionedSource::new(fixture()));
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let mut sink = VersionBumpingSink {
        source: Arc::clone(&source),
        bytes: Vec::new(),
        bumped: false,
    };
    let error = package
        .write_part_overlay_to_stream(&mut sink, &pack(TARGET_URI), REPLACEMENT.to_vec())
        .unwrap_err();
    match error {
        OpcError::IncompleteOutput { written, source } => {
            assert!(written > 0);
            assert_eq!(written as usize, sink.bytes.len());
            assert!(matches!(*source, OpcError::SourceChanged { .. }));
        },
        other => panic!("expected incomplete output after a source change: {other:?}"),
    }
    assert!(source.bytes_equal(fixture()));
    assert_eq!(source.revision(), 1);
}
