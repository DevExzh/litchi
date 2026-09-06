#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used,
    reason = "focused transfer assertions intentionally fail on fixture errors"
)]

//! Adversarial coverage for source-authorized compressed Part transfer.
//!
//! The compressed bytes accepted by these tests always come from a ZIP reader
//! issued token.  The tests therefore exercise the public authorization and
//! topology APIs, rather than manufacturing a compressed payload or metadata
//! claim in the test.

use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, Mutex, RwLock,
    atomic::{AtomicBool, AtomicU64, Ordering},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, ReadAt,
    Resource, SourceVersion,
};
use litchi_opc::{
    AuthorizedPrecompressedPart, OpcError, PackURI, ReadLimits, ReadResource, SourceBackedPackage,
    SourceTopologyPlan,
};

const CONTENT_TYPES_NS: &str = "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_REL: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
const SOURCE_PART: &str = "/custom/source.bin";
const SOURCE_CONTENT_TYPE: &str = "application/octet-stream";
const DESTINATION_PART: &str = "/custom/transferred.bin";
const NATIVE_IMAGE_PART: &str = "/ppt/media/image1.png";
const NATIVE_IMAGE_CONTENT_TYPE: &str = "image/png";

const NATIVE_POI_VIDEO_PPTX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/slideshow/EmbeddedVideo.pptx");
const NATIVE_POI_TRAILING_BYTE_PPTX: &[u8] =
    include_bytes!("../../../test-data/poi/test-data/slideshow/bug62513.pptx");

#[derive(Clone, Copy)]
struct Entry<'a> {
    name: &'a [u8],
    data: &'a [u8],
}

fn pack(uri: &str) -> PackURI {
    PackURI::new(uri).expect("test URI must be valid")
}

fn put_u16(bytes: &mut Vec<u8>, value: u16) {
    bytes.extend_from_slice(&value.to_le_bytes());
}

fn put_u32(bytes: &mut Vec<u8>, value: u32) {
    bytes.extend_from_slice(&value.to_le_bytes());
}

fn at_u16(bytes: &[u8], offset: usize) -> u16 {
    u16::from_le_bytes(bytes[offset..offset + 2].try_into().unwrap())
}

fn at_u32(bytes: &[u8], offset: usize) -> u32 {
    u32::from_le_bytes(bytes[offset..offset + 4].try_into().unwrap())
}

fn stored_archive(entries: &[Entry<'_>]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut central = Vec::new();
    for entry in entries {
        let local_offset = u32::try_from(output.len()).unwrap();
        let size = u32::try_from(entry.data.len()).unwrap();
        let crc = soapberry_zip::crc32(entry.data);

        put_u32(&mut output, 0x0403_4b50);
        put_u16(&mut output, 20);
        put_u16(&mut output, 0);
        put_u16(&mut output, 0);
        put_u16(&mut output, 0);
        put_u16(&mut output, 0);
        put_u32(&mut output, crc);
        put_u32(&mut output, size);
        put_u32(&mut output, size);
        put_u16(&mut output, u16::try_from(entry.name.len()).unwrap());
        put_u16(&mut output, 0);
        output.extend_from_slice(entry.name);
        output.extend_from_slice(entry.data);

        put_u32(&mut central, 0x0201_4b50);
        put_u16(&mut central, 20);
        put_u16(&mut central, 20);
        put_u16(&mut central, 0);
        put_u16(&mut central, 0);
        put_u16(&mut central, 0);
        put_u16(&mut central, 0);
        put_u32(&mut central, crc);
        put_u32(&mut central, size);
        put_u32(&mut central, size);
        put_u16(&mut central, u16::try_from(entry.name.len()).unwrap());
        put_u16(&mut central, 0);
        put_u16(&mut central, 0);
        put_u16(&mut central, 0);
        put_u16(&mut central, 0);
        put_u32(&mut central, 0);
        put_u32(&mut central, local_offset);
        central.extend_from_slice(entry.name);
    }

    let central_offset = u32::try_from(output.len()).unwrap();
    let central_size = u32::try_from(central.len()).unwrap();
    output.extend_from_slice(&central);
    put_u32(&mut output, 0x0605_4b50);
    put_u16(&mut output, 0);
    put_u16(&mut output, 0);
    let count = u16::try_from(entries.len()).unwrap();
    put_u16(&mut output, count);
    put_u16(&mut output, count);
    put_u32(&mut output, central_size);
    put_u32(&mut output, central_offset);
    put_u16(&mut output, 0);
    output
}

fn content_types(extra_default: &str) -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="{CONTENT_TYPES_NS}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>{extra_default}</Types>"#
    )
    .into_bytes()
}

fn root_relationships() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="{RELATIONSHIPS_NS}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_REL}" Target="word/document.xml"/></Relationships>"#
    )
    .into_bytes()
}

fn synthetic_source_bytes() -> (Vec<u8>, Vec<u8>) {
    let content_types =
        content_types(r#"<Default Extension="bin" ContentType="application/octet-stream"/>"#);
    let relationships = root_relationships();
    let payload = b"source-authorized compressed transfer payload\n".to_vec();
    let entries = [
        Entry {
            name: b"[Content_Types].xml",
            data: &content_types,
        },
        Entry {
            name: b"_rels/.rels",
            data: &relationships,
        },
        Entry {
            name: b"word/document.xml",
            data: b"<document/>\n",
        },
        Entry {
            name: b"custom/source.bin",
            data: &payload,
        },
    ];
    (stored_archive(&entries), payload)
}

fn synthetic_destination_bytes() -> Vec<u8> {
    let content_types =
        content_types(r#"<Default Extension="bin" ContentType="application/octet-stream"/>"#);
    let relationships = root_relationships();
    let entries = [
        Entry {
            name: b"[Content_Types].xml",
            data: &content_types,
        },
        Entry {
            name: b"_rels/.rels",
            data: &relationships,
        },
        Entry {
            name: b"word/document.xml",
            data: b"<destination/>\n",
        },
    ];
    stored_archive(&entries)
}

fn native_destination_bytes() -> Vec<u8> {
    let content_types = content_types(r#"<Default Extension="png" ContentType="image/png"/>"#);
    let relationships = root_relationships();
    let entries = [
        Entry {
            name: b"[Content_Types].xml",
            data: &content_types,
        },
        Entry {
            name: b"_rels/.rels",
            data: &relationships,
        },
        Entry {
            name: b"word/document.xml",
            data: b"<destination/>\n",
        },
    ];
    stored_archive(&entries)
}

/// Return the central-declared method and the exact compressed range for a
/// normal ZIP32 member.  The native and synthetic fixtures are deliberately
/// small, and the transfer token itself is responsible for ZIP64 support.
fn compressed_member(bytes: &[u8], wanted: &str) -> (u16, Vec<u8>) {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .expect("fixture must contain ZIP32 EOCD");
    let count = usize::from(at_u16(bytes, eocd + 10));
    let mut central = usize::try_from(at_u32(bytes, eocd + 16)).unwrap();
    for _ in 0..count {
        assert_eq!(at_u32(bytes, central), 0x0201_4b50);
        let method = at_u16(bytes, central + 10);
        let compressed_size = usize::try_from(at_u32(bytes, central + 20)).unwrap();
        let name_len = usize::from(at_u16(bytes, central + 28));
        let extra_len = usize::from(at_u16(bytes, central + 30));
        let comment_len = usize::from(at_u16(bytes, central + 32));
        let name_start = central + 46;
        let name = std::str::from_utf8(&bytes[name_start..name_start + name_len]).unwrap();
        if name == wanted {
            let local = usize::try_from(at_u32(bytes, central + 42)).unwrap();
            assert_eq!(at_u32(bytes, local), 0x0403_4b50);
            let local_name_len = usize::from(at_u16(bytes, local + 26));
            let local_extra_len = usize::from(at_u16(bytes, local + 28));
            let start = local + 30 + local_name_len + local_extra_len;
            return (method, bytes[start..start + compressed_size].to_vec());
        }
        central += 46 + name_len + extra_len + comment_len;
    }
    panic!("missing ZIP member {wanted}");
}

fn corrupt_member_crc(mut bytes: Vec<u8>, wanted: &str) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .expect("fixture must contain ZIP32 EOCD");
    let count = usize::from(at_u16(&bytes, eocd + 10));
    let mut central = usize::try_from(at_u32(&bytes, eocd + 16)).unwrap();
    for _ in 0..count {
        assert_eq!(at_u32(&bytes, central), 0x0201_4b50);
        let name_len = usize::from(at_u16(&bytes, central + 28));
        let extra_len = usize::from(at_u16(&bytes, central + 30));
        let comment_len = usize::from(at_u16(&bytes, central + 32));
        let name_start = central + 46;
        let name = std::str::from_utf8(&bytes[name_start..name_start + name_len]).unwrap();
        if name == wanted {
            let crc = at_u32(&bytes, central + 16) ^ 1;
            bytes[central + 16..central + 20].copy_from_slice(&crc.to_le_bytes());
            let local = usize::try_from(at_u32(&bytes, central + 42)).unwrap();
            bytes[local + 14..local + 18].copy_from_slice(&crc.to_le_bytes());
            return bytes;
        }
        central += 46 + name_len + extra_len + comment_len;
    }
    panic!("missing ZIP member {wanted}");
}

#[derive(Debug)]
struct TestSource {
    bytes: RwLock<Vec<u8>>,
    revision: AtomicU64,
    cancel_source: Option<CancellationSource>,
    cancel_next_read: AtomicBool,
}

impl TestSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            bytes: RwLock::new(bytes),
            revision: AtomicU64::new(0),
            cancel_source: None,
            cancel_next_read: AtomicBool::new(false),
        })
    }

    fn with_cancellation(bytes: Vec<u8>, cancel_source: CancellationSource) -> Arc<Self> {
        Arc::new(Self {
            bytes: RwLock::new(bytes),
            revision: AtomicU64::new(0),
            cancel_source: Some(cancel_source),
            cancel_next_read: AtomicBool::new(false),
        })
    }

    fn bump_revision(&self) {
        self.revision.fetch_add(1, Ordering::SeqCst);
    }

    fn cancel_on_next_read(&self) {
        self.cancel_next_read.store(true, Ordering::SeqCst);
    }
}

impl ReadAt for TestSource {
    fn len(&self) -> io::Result<u64> {
        let bytes = self.bytes.read().unwrap();
        u64::try_from(bytes.len()).map_err(|error| io::Error::other(error.to_string()))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        if self.cancel_next_read.swap(false, Ordering::SeqCst) {
            if let Some(source) = &self.cancel_source {
                source.cancel();
            }
        }
        let bytes = self.bytes.read().unwrap();
        let Ok(start) = usize::try_from(offset) else {
            return Ok(0);
        };
        let Some(input) = bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len());
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            0x0431,
            self.revision.load(Ordering::SeqCst),
        ))
    }
}

fn managed_context(memory: u64) -> (Budget, CancellationSource, ExecutionContext) {
    let (cancel_source, cancellation) = CancellationSource::pair();
    let context = managed_context_for_token(memory, cancellation);
    (context.0, cancel_source, context.1)
}

fn managed_context_for_token(
    memory: u64,
    cancellation: litchi_core::CancellationToken,
) -> (Budget, ExecutionContext) {
    let budget = Budget::root(
        "source-part-transfer-test",
        Limits::new(memory, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let max_in_flight = NonZeroU64::new(memory.max(1)).unwrap();
    let execution_limits = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        max_in_flight,
        0,
    )
    .unwrap();
    (
        budget.clone(),
        ExecutionContext::new(budget, cancellation, execution_limits),
    )
}

fn authorized_synthetic_transfer() -> (Arc<TestSource>, AuthorizedPrecompressedPart, Vec<u8>) {
    let (source_bytes, expected) = synthetic_source_bytes();
    let source = TestSource::new(source_bytes);
    let package = SourceBackedPackage::from_read_at(source.clone()).unwrap();
    let view = package.part(&pack(SOURCE_PART)).unwrap();
    let token = view
        .authorize_precompressed(Arc::new(expected.clone()))
        .unwrap();
    drop(package);
    (source, token, expected)
}

fn authorized_managed_synthetic_transfer()
-> (CancellationSource, AuthorizedPrecompressedPart, Vec<u8>) {
    let (source_bytes, expected) = synthetic_source_bytes();
    let source = TestSource::new(source_bytes);
    let (cancel_source, cancel_token) = CancellationSource::pair();
    let (_budget, context) = managed_context_for_token(64 * 1024 * 1024, cancel_token);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source,
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let view = package.part(&pack(SOURCE_PART)).unwrap();
    let token = view
        .authorize_precompressed(Arc::new(expected.clone()))
        .unwrap();
    drop(package);
    (cancel_source, token, expected)
}

fn write_transfer(destination: Vec<u8>, plan: SourceTopologyPlan) -> litchi_opc::Result<Vec<u8>> {
    let package = SourceBackedPackage::from_vec(destination)?;
    let mut output = Vec::new();
    package.write_topology_to_stream(&mut output, plan)?;
    Ok(output)
}

fn add_synthetic_token(
    plan: &mut SourceTopologyPlan,
    token: AuthorizedPrecompressedPart,
) -> litchi_opc::Result<()> {
    plan.try_add_precompressed_part(pack(DESTINATION_PART), SOURCE_CONTENT_TYPE, token)
}

fn is_source_changed(error: &OpcError) -> bool {
    match error {
        OpcError::SourceChanged { .. } => true,
        OpcError::IncompleteOutput { source, .. } => is_source_changed(source),
        _ => false,
    }
}

#[test]
fn cross_source_transfer_reuses_exact_verified_store_member() {
    let (_source, token, expected) = authorized_synthetic_transfer();
    let source_bytes = synthetic_source_bytes().0;
    let (source_method, source_compressed) = compressed_member(&source_bytes, "custom/source.bin");

    let mut plan = SourceTopologyPlan::new();
    add_synthetic_token(&mut plan, token).unwrap();
    let output = write_transfer(synthetic_destination_bytes(), plan).unwrap();

    let (output_method, output_compressed) = compressed_member(&output, "custom/transferred.bin");
    assert_eq!(
        source_method, 0,
        "the synthetic source is an explicit Store member"
    );
    assert_eq!(output_method, source_method);
    assert_eq!(output_compressed, source_compressed);

    let reopened = SourceBackedPackage::from_vec(output).unwrap();
    let data = reopened
        .part(&pack(DESTINATION_PART))
        .unwrap()
        .data()
        .unwrap();
    assert_eq!(data.as_bytes(), expected.as_slice());
}

#[test]
fn native_poi_video_image_transfer_reuses_exact_compressed_member() {
    let source = SourceBackedPackage::from_vec(NATIVE_POI_VIDEO_PPTX.to_vec()).unwrap();
    let view = source.part(&pack(NATIVE_IMAGE_PART)).unwrap();
    let expected = view.data().unwrap().as_bytes().to_vec();
    let token = view
        .authorize_precompressed(Arc::new(expected.clone()))
        .unwrap();
    drop(source);

    let (source_method, source_compressed) =
        compressed_member(NATIVE_POI_VIDEO_PPTX, "ppt/media/image1.png");
    let mut plan = SourceTopologyPlan::new();
    plan.try_add_precompressed_part(
        pack("/custom/native-image.png"),
        NATIVE_IMAGE_CONTENT_TYPE,
        token,
    )
    .unwrap();
    let output = write_transfer(native_destination_bytes(), plan).unwrap();

    let (output_method, output_compressed) = compressed_member(&output, "custom/native-image.png");
    assert_eq!(source_method, 0, "the selected POI PNG is a Stored member");
    assert_eq!(output_method, source_method);
    assert_eq!(output_compressed, source_compressed);

    let reopened = SourceBackedPackage::from_vec(output).unwrap();
    let data = reopened
        .part(&pack("/custom/native-image.png"))
        .unwrap()
        .data()
        .unwrap();
    assert_eq!(data.as_bytes(), expected.as_slice());
}

#[test]
fn native_poi_slide_trailing_byte_refuses_authorized_transfer() {
    let source = SourceBackedPackage::from_vec(NATIVE_POI_TRAILING_BYTE_PPTX.to_vec()).unwrap();
    let view = source.part(&pack("/ppt/media/image2.jpeg")).unwrap();
    let expected = view.data().unwrap().as_bytes().to_vec();
    let error = view
        .authorize_precompressed(Arc::new(expected))
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
}

#[test]
fn source_revision_change_after_view_drop_refuses_before_first_sink_write() {
    let (source, token, _expected) = authorized_synthetic_transfer();
    source.bump_revision();

    let mut plan = SourceTopologyPlan::new();
    add_synthetic_token(&mut plan, token).unwrap();
    let destination = synthetic_destination_bytes();
    let mut output = Vec::new();
    let package = SourceBackedPackage::from_vec(destination).unwrap();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();

    assert!(matches!(error, OpcError::SourceChanged { .. }));
    assert!(
        output.is_empty(),
        "stale source must fail before sink output"
    );
}

#[test]
fn source_revision_change_during_destination_sink_is_typed_and_partial() {
    let (source, token, _expected) = authorized_synthetic_transfer();
    let mut plan = SourceTopologyPlan::new();
    add_synthetic_token(&mut plan, token).unwrap();

    let output = Arc::new(Mutex::new(Vec::new()));
    let sink = MutatingSink {
        output: Arc::clone(&output),
        source: Arc::clone(&source),
        mutated: false,
    };
    let package = SourceBackedPackage::from_vec(synthetic_destination_bytes()).unwrap();
    let error = package.write_topology_to_stream(sink, plan).unwrap_err();

    assert!(
        is_source_changed(&error),
        "error must retain source revision failure"
    );
    let written = match error {
        OpcError::IncompleteOutput { written, .. } => written,
        other => panic!("source mutation after a sink write must be incomplete: {other:?}"),
    };
    assert!(written > 0);
    assert_eq!(
        usize::try_from(written).unwrap(),
        output.lock().unwrap().len()
    );
}

struct MutatingSink {
    output: Arc<Mutex<Vec<u8>>>,
    source: Arc<TestSource>,
    mutated: bool,
}

struct CancellingSink {
    output: Arc<Mutex<Vec<u8>>>,
    cancellation: CancellationSource,
    cancelled: bool,
}

impl Write for CancellingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.output.lock().unwrap().extend_from_slice(bytes);
        if !self.cancelled {
            self.cancellation.cancel();
            self.cancelled = true;
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_cancellation_during_destination_sink_preserves_partial_prefix() {
    let (cancellation, token, _expected) = authorized_managed_synthetic_transfer();
    let mut plan = SourceTopologyPlan::new();
    add_synthetic_token(&mut plan, token).unwrap();
    let output = Arc::new(Mutex::new(Vec::new()));
    let sink = CancellingSink {
        output: Arc::clone(&output),
        cancellation,
        cancelled: false,
    };
    let package = SourceBackedPackage::from_vec(synthetic_destination_bytes()).unwrap();
    let error = package.write_topology_to_stream(sink, plan).unwrap_err();

    let written = match error {
        OpcError::IncompleteOutput { written, source } => {
            assert!(matches!(*source, OpcError::Cancelled));
            written
        },
        other => panic!("source cancellation after a sink write must be incomplete: {other:?}"),
    };
    assert!(written > 0);
    assert_eq!(
        usize::try_from(written).unwrap(),
        output.lock().unwrap().len()
    );
}

impl Write for MutatingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.output.lock().unwrap().extend_from_slice(bytes);
        if !self.mutated {
            self.source.bump_revision();
            self.mutated = true;
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct FinalMutatingSink {
    output: Arc<Mutex<Vec<u8>>>,
    source: Arc<TestSource>,
    expected_len: usize,
    mutated: Arc<AtomicBool>,
}

impl Write for FinalMutatingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let mut output = self.output.lock().unwrap();
        output.extend_from_slice(bytes);
        if output.len() == self.expected_len && !self.mutated.swap(true, Ordering::SeqCst) {
            self.source.bump_revision();
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_revision_change_after_final_sink_write_is_incomplete_with_exact_prefix() {
    let (_control_source, control_token, _expected) = authorized_synthetic_transfer();
    let mut control_plan = SourceTopologyPlan::new();
    add_synthetic_token(&mut control_plan, control_token).unwrap();
    let control_output = write_transfer(synthetic_destination_bytes(), control_plan).unwrap();

    let (source, token, _expected) = authorized_synthetic_transfer();
    let mut plan = SourceTopologyPlan::new();
    add_synthetic_token(&mut plan, token).unwrap();
    let output = Arc::new(Mutex::new(Vec::new()));
    let mutated = Arc::new(AtomicBool::new(false));
    let sink = FinalMutatingSink {
        output: Arc::clone(&output),
        source,
        expected_len: control_output.len(),
        mutated: Arc::clone(&mutated),
    };
    let package = SourceBackedPackage::from_vec(synthetic_destination_bytes()).unwrap();
    let error = package.write_topology_to_stream(sink, plan).unwrap_err();

    let written = match error {
        OpcError::IncompleteOutput { written, source } => {
            assert!(is_source_changed(&source));
            written
        },
        other => panic!("source mutation after the final sink write must be incomplete: {other:?}"),
    };
    assert!(mutated.load(Ordering::SeqCst));
    assert_eq!(written, u64::try_from(control_output.len()).unwrap());
    assert_eq!(
        usize::try_from(written).unwrap(),
        output.lock().unwrap().len()
    );
}

#[test]
fn decoded_mismatch_and_corrupt_crc_are_rejected_before_publication() {
    let (source_bytes, expected) = synthetic_source_bytes();
    let source = SourceBackedPackage::from_vec(source_bytes.clone()).unwrap();
    let view = source.part(&pack(SOURCE_PART)).unwrap();
    let mut wrong = expected.clone();
    wrong[0] ^= 0x40;
    let mismatch = view.authorize_precompressed(Arc::new(wrong)).unwrap_err();
    assert!(matches!(mismatch, OpcError::ZipError(_)));
    drop(source);

    let corrupted = corrupt_member_crc(source_bytes, "custom/source.bin");
    let package = SourceBackedPackage::from_vec(corrupted).unwrap();
    let view = package.part(&pack(SOURCE_PART)).unwrap();
    let error = view
        .authorize_precompressed(Arc::new(expected))
        .unwrap_err();
    assert!(matches!(error, OpcError::ZipError(_)));
}

#[test]
fn cancellation_during_compressed_capture_returns_cancelled_without_token() {
    let (source_bytes, expected) = synthetic_source_bytes();
    let (cancel_source, cancel_token) = CancellationSource::pair();
    let source = TestSource::with_cancellation(source_bytes, cancel_source.clone());
    let (_budget, context) = managed_context_for_token(64 * 1024 * 1024, cancel_token);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source.clone(),
        ReadLimits::default(),
        context,
    )
    .unwrap();
    source.cancel_on_next_read();
    let view = package.part(&pack(SOURCE_PART)).unwrap();
    let error = view
        .authorize_precompressed(Arc::new(expected))
        .unwrap_err();

    assert!(matches!(error, OpcError::Cancelled));
    assert!(cancel_source.is_cancelled());
}

#[test]
fn output_limit_is_checked_before_any_destination_output() {
    let (_source, token, _expected) = authorized_synthetic_transfer();
    let destination = synthetic_destination_bytes();
    let content_types =
        content_types(r#"<Default Extension="bin" ContentType="application/octet-stream"/>"#);
    let relationships = root_relationships();
    let destination_total =
        u64::try_from(content_types.len() + relationships.len() + b"<destination/>\n".len())
            .unwrap();
    let limits = ReadLimits::builder()
        .max_archive_total_bytes(destination_total)
        .unwrap()
        .build()
        .unwrap();
    let package = SourceBackedPackage::from_vec_with_limits(destination, limits).unwrap();
    let mut plan = SourceTopologyPlan::new();
    add_synthetic_token(&mut plan, token).unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();

    assert!(matches!(
        error,
        OpcError::ReadLimit {
            resource: ReadResource::ArchiveTotalBytes,
            ..
        }
    ));
    assert!(output.is_empty());
}

struct FailingSink {
    output: Arc<Mutex<Vec<u8>>>,
    limit: usize,
}

impl Write for FailingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let mut output = self.output.lock().unwrap();
        if output.len() >= self.limit {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "test sink"));
        }
        let accepted = bytes.len().min(self.limit - output.len());
        output.extend_from_slice(&bytes[..accepted]);
        Ok(accepted)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn sequential_sink_failure_reports_incomplete_output_after_partial_prefix() {
    let (_source, token, _expected) = authorized_synthetic_transfer();
    let mut plan = SourceTopologyPlan::new();
    add_synthetic_token(&mut plan, token).unwrap();
    let output = Arc::new(Mutex::new(Vec::new()));
    let sink = FailingSink {
        output: Arc::clone(&output),
        limit: 8,
    };
    let package = SourceBackedPackage::from_vec(synthetic_destination_bytes()).unwrap();
    let error = package.write_topology_to_stream(sink, plan).unwrap_err();

    let written = match error {
        OpcError::IncompleteOutput { written, source } => {
            assert!(matches!(*source, OpcError::IoError(_)));
            written
        },
        other => panic!("expected typed incomplete sink failure, got {other:?}"),
    };
    assert_eq!(written, 8);
    assert_eq!(output.lock().unwrap().len(), 8);
}

#[test]
fn exact_managed_capture_budget_succeeds_and_one_byte_less_refuses() {
    let (source_bytes, expected) = synthetic_source_bytes();
    let source = Arc::new(litchi_core::OwnedSource::new(source_bytes.clone()));
    let (budget, _cancel_source, context) = managed_context(u64::MAX);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        source,
        ReadLimits::default(),
        context,
    )
    .unwrap();
    let view = package.part(&pack(SOURCE_PART)).unwrap();
    let payload = view.data().unwrap();
    let expected = Arc::new(expected);
    let token = view.authorize_precompressed(expected).unwrap();
    let exact = budget.used(Resource::Memory);
    assert!(
        exact > 0,
        "authorization must reserve managed capture memory"
    );
    drop(token);
    drop(payload);
    drop(package);
    assert_eq!(budget.used(Resource::Memory), 0);

    let (_budget, _cancel_source, exact_context) = managed_context(exact);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(litchi_core::OwnedSource::new(source_bytes.clone())),
        ReadLimits::default(),
        exact_context,
    )
    .unwrap();
    let view = package.part(&pack(SOURCE_PART)).unwrap();
    let payload = view.data().unwrap();
    let result = view.authorize_precompressed(Arc::new(payload.as_bytes().to_vec()));
    assert!(
        result.is_ok(),
        "the observed peak must admit an identical operation"
    );
    drop(result);
    drop(payload);
    drop(package);

    let (_budget, _cancel_source, under_context) = managed_context(exact - 1);
    let package = SourceBackedPackage::from_read_at_with_execution_context(
        Arc::new(litchi_core::OwnedSource::new(source_bytes)),
        ReadLimits::default(),
        under_context,
    );
    let package = match package {
        Ok(package) => package,
        Err(error) => {
            assert!(matches!(
                error,
                OpcError::Execution(ExecutionError::ResourceLimit(_))
            ));
            return;
        },
    };
    let view = package.part(&pack(SOURCE_PART)).unwrap();
    let payload = view.data().unwrap();
    let error = view
        .authorize_precompressed(Arc::new(payload.as_bytes().to_vec()))
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::Execution(ExecutionError::ResourceLimit(_))
    ));
}

fn mark_encrypted_with_compatibility_crc_zero(mut bytes: Vec<u8>, wanted: &str) -> Vec<u8> {
    let eocd = bytes
        .windows(4)
        .rposition(|window| window == 0x0605_4b50_u32.to_le_bytes())
        .expect("fixture must contain ZIP32 EOCD");
    let count = usize::from(at_u16(&bytes, eocd + 10));
    let mut central = usize::try_from(at_u32(&bytes, eocd + 16)).unwrap();
    for _ in 0..count {
        assert_eq!(at_u32(&bytes, central), 0x0201_4b50);
        let name_len = usize::from(at_u16(&bytes, central + 28));
        let extra_len = usize::from(at_u16(&bytes, central + 30));
        let comment_len = usize::from(at_u16(&bytes, central + 32));
        let name_start = central + 46;
        let name = std::str::from_utf8(&bytes[name_start..name_start + name_len]).unwrap();
        if name == wanted {
            let flags = at_u16(&bytes, central + 8) | 1;
            bytes[central + 8..central + 10].copy_from_slice(&flags.to_le_bytes());
            bytes[central + 16..central + 20].copy_from_slice(&0_u32.to_le_bytes());
            let local = usize::try_from(at_u32(&bytes, central + 42)).unwrap();
            let local_flags = at_u16(&bytes, local + 6) | 1;
            bytes[local + 6..local + 8].copy_from_slice(&local_flags.to_le_bytes());
            bytes[local + 14..local + 18].copy_from_slice(&0_u32.to_le_bytes());
            return bytes;
        }
        central += 46 + name_len + extra_len + comment_len;
    }
    panic!("missing ZIP member {wanted}");
}

fn signed_synthetic_source_bytes() -> Vec<u8> {
    let content_types =
        content_types(r#"<Default Extension="bin" ContentType="application/octet-stream"/>"#);
    let relationships = root_relationships();
    let entries = [
        Entry {
            name: b"[Content_Types].xml",
            data: &content_types,
        },
        Entry {
            name: b"_rels/.rels",
            data: &relationships,
        },
        Entry {
            name: b"word/document.xml",
            data: b"<signed-source/>\n",
        },
        Entry {
            name: b"custom/source.bin",
            data: b"signed-source-payload\n",
        },
        Entry {
            name: b"_xmlsignatures/origin.sigs",
            data: b"<signature/>\n",
        },
    ];
    stored_archive(&entries)
}

#[test]
fn encrypted_store_with_zero_crc_is_refused_before_token_or_output() {
    let (source_bytes, expected) = synthetic_source_bytes();
    let source_bytes =
        mark_encrypted_with_compatibility_crc_zero(source_bytes, "custom/source.bin");
    let source = SourceBackedPackage::from_vec(source_bytes.clone()).unwrap();
    let view = source.part(&pack(SOURCE_PART)).unwrap();
    let error = view
        .authorize_precompressed(Arc::new(expected))
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SourceBackedOverlayUnavailable { .. }
    ));
    drop(source);

    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack(DESTINATION_PART),
        SOURCE_CONTENT_TYPE,
        b"decoded fallback is still policy-refused".to_vec(),
    )
    .unwrap();
    let package = SourceBackedPackage::from_vec(source_bytes).unwrap();
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
fn rooted_signature_infrastructure_is_refused_before_token_or_output() {
    let source_bytes = signed_synthetic_source_bytes();
    let source = SourceBackedPackage::from_vec(source_bytes.clone()).unwrap();
    let view = source.part(&pack(SOURCE_PART)).unwrap();
    let expected = view.data().unwrap().as_bytes().to_vec();
    let error = view
        .authorize_precompressed(Arc::new(expected))
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SignedSourceRequiresExplicitPolicy
    ));
    drop(source);

    let mut plan = SourceTopologyPlan::new();
    plan.try_add_part(
        pack(DESTINATION_PART),
        SOURCE_CONTENT_TYPE,
        b"decoded signed-source addition".to_vec(),
    )
    .unwrap();
    let package = SourceBackedPackage::from_vec(source_bytes).unwrap();
    let mut output = Vec::new();
    let error = package
        .write_topology_to_stream(&mut output, plan)
        .unwrap_err();
    assert!(matches!(
        error,
        OpcError::SignedSourceRequiresExplicitPolicy
    ));
    assert!(output.is_empty());
}
