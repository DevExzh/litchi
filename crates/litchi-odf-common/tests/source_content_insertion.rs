#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::unwrap_used,
    reason = "These integration fixtures are deliberately small and assertion-driven."
)]

//! Integration coverage for bounded decoded `content.xml` insertion.
//!
//! The tests exercise both ZIP framing and the source transaction.  The
//! source adapter records physical reads so replay can prove that opaque
//! payload verification is optional and that a stale source is rejected in
//! every pass.

use litchi_core::{
    Budget, CancellationSource, CancellationToken, Error, ExecutionContext, ExecutionError,
    ExecutionLimits, Limits as CoreLimits, ReadAt, Resource, SourceVersion,
};
use litchi_odf_common::core::{
    AuthoredXmlFragment, OwnedPackage, SourceBackedPackage, SourceContentInsertionError,
    SourceContentInsertionPlan, SourceContentPublicationError, SourceContentPublicationOptions,
    XmlStreamLimits,
};
use soapberry_zip::{PreservationIndex, ZipArchive};
use std::collections::BTreeMap;
use std::io::{self, Cursor, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::ops::Range;
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicBool, AtomicU64, Ordering},
};
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod as ZipCompressionMethod, ZipWriter};

const MIME: &str = "application/vnd.oasis.opendocument.text";
const CONTENT_MEDIA: &str = "text/xml";
const MANIFEST_NS: &str = "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0";
const SOURCE_CONTENT: &[u8] = br#"<?xml version="1.0"?><office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><office:body><office:text><office:p>source</office:p></office:text></office:body></office:document-content>"#;
const INSERTED_MARKUP: &[u8] = b"<p>inserted</p>";

static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(50_000);

#[derive(Debug, Clone, PartialEq, Eq)]
struct RawMember {
    local: Vec<u8>,
    central_without_offset: Vec<u8>,
}

#[derive(Debug)]
struct BumpState {
    range: Range<u64>,
    armed: bool,
    done: bool,
}

#[derive(Debug)]
struct SourceState {
    bytes: Vec<u8>,
    revision: u64,
    ranges: Vec<Range<u64>>,
    bump: Option<BumpState>,
}

/// A small positional source with deterministic revision and read-range hooks.
#[derive(Debug)]
struct TestSource {
    id: u64,
    state: Mutex<SourceState>,
}

impl TestSource {
    fn new(bytes: Vec<u8>) -> Arc<Self> {
        Arc::new(Self {
            id: NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
            state: Mutex::new(SourceState {
                bytes,
                revision: 0,
                ranges: Vec::new(),
                bump: None,
            }),
        })
    }

    fn bump_revision(&self) {
        let mut state = self.state.lock().unwrap();
        state.revision = state.revision.saturating_add(1);
    }

    fn set_bump_range(&self, range: Range<u64>) {
        self.state.lock().unwrap().bump = Some(BumpState {
            range,
            armed: false,
            done: false,
        });
    }

    fn arm_bump(&self) {
        if let Some(bump) = self.state.lock().unwrap().bump.as_mut() {
            bump.armed = true;
        }
    }

    fn clear_ranges(&self) {
        self.state.lock().unwrap().ranges.clear();
    }

    fn has_read_range(&self, wanted: Range<u64>) -> bool {
        self.state
            .lock()
            .unwrap()
            .ranges
            .iter()
            .any(|range| overlaps(range, &wanted))
    }
}

impl ReadAt for TestSource {
    fn len(&self) -> io::Result<u64> {
        u64::try_from(self.state.lock().unwrap().bytes.len())
            .map_err(|_| io::Error::other("test source length exceeds u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let mut state = self.state.lock().unwrap();
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        let Some(input) = state.bytes.get(start..) else {
            return Ok(0);
        };
        let amount = input.len().min(output.len());
        output[..amount].copy_from_slice(&input[..amount]);
        if amount == 0 {
            return Ok(0);
        }
        let end = offset
            .checked_add(
                u64::try_from(amount).map_err(|_| {
                    io::Error::new(io::ErrorKind::InvalidInput, "read length overflow")
                })?,
            )
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "read range overflow"))?;
        let read_range = offset..end;
        state.ranges.push(read_range.clone());
        let bump_now = state
            .bump
            .as_ref()
            .is_some_and(|bump| bump.armed && !bump.done && overlaps(&read_range, &bump.range));
        if bump_now {
            state.revision = state.revision.saturating_add(1);
            if let Some(bump) = state.bump.as_mut() {
                bump.done = true;
            }
        }
        Ok(amount)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        let state = self.state.lock().unwrap();
        Ok(SourceVersion::new(self.id, state.revision))
    }
}

#[derive(Debug)]
struct ZeroSink;

impl Write for ZeroSink {
    fn write(&mut self, _input: &[u8]) -> io::Result<usize> {
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct ErrorSink;

impl Write for ErrorSink {
    fn write(&mut self, _input: &[u8]) -> io::Result<usize> {
        Err(io::Error::new(
            io::ErrorKind::BrokenPipe,
            "test sink failure",
        ))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct PartialFailSink {
    accepted: usize,
    first: bool,
}

impl Write for PartialFailSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.first {
            self.first = false;
            return Ok(self.accepted.min(input.len()));
        }
        Err(io::Error::new(
            io::ErrorKind::BrokenPipe,
            "test sink failed after a partial write",
        ))
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct MutatingSink {
    source: Arc<TestSource>,
    writes: usize,
    bytes: Vec<u8>,
}

impl Write for MutatingSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        self.writes = self.writes.saturating_add(1);
        self.bytes.extend_from_slice(input);
        // The first two writes are the untouched mimetype local span and
        // payload. The third is the replacement member's local framing; a
        // following callback write must then observe the stale revision.
        if self.writes == 3 {
            self.source.bump_revision();
        }
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[derive(Debug)]
struct ObserveSink {
    source: Arc<TestSource>,
    watched: Range<u64>,
    saw_watched_before_output: Arc<AtomicBool>,
    bytes: Vec<u8>,
}

impl Write for ObserveSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if self.saw_watched_before_output.load(Ordering::Acquire) {
            // Keep the first observation stable after output starts.
        } else if self.source.has_read_range(self.watched.clone()) {
            self.saw_watched_before_output
                .store(true, Ordering::Release);
        }
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn overlaps(left: &Range<u64>, right: &Range<u64>) -> bool {
    left.start < right.end && right.start < left.end
}

fn source_reader(source: &Arc<TestSource>) -> Arc<dyn ReadAt> {
    let source: Arc<dyn ReadAt> = source.clone();
    source
}

fn manifest() -> Vec<u8> {
    format!(
        r#"<?xml version="1.0"?><manifest:manifest xmlns:manifest="{MANIFEST_NS}" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="{MIME}"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="{CONTENT_MEDIA}"/><manifest:file-entry manifest:full-path="Pictures/blob.bin" manifest:media-type="application/octet-stream"/></manifest:manifest>"#
    )
    .into_bytes()
}

fn opaque_media() -> Vec<u8> {
    (0..8192)
        .map(|index| (index as u32).wrapping_mul(37) as u8)
        .collect()
}

fn zip_package(
    content_method: ZipCompressionMethod,
    media_method: ZipCompressionMethod,
) -> Vec<u8> {
    let mut writer = ZipWriter::new(Cursor::new(Vec::new()));
    writer
        .start_file(
            "mimetype",
            SimpleFileOptions::default().compression_method(ZipCompressionMethod::Stored),
        )
        .unwrap();
    writer.write_all(MIME.as_bytes()).unwrap();
    writer
        .start_file(
            "content.xml",
            SimpleFileOptions::default().compression_method(content_method),
        )
        .unwrap();
    writer.write_all(SOURCE_CONTENT).unwrap();
    writer
        .start_file(
            "Pictures/blob.bin",
            SimpleFileOptions::default().compression_method(media_method),
        )
        .unwrap();
    writer.write_all(&opaque_media()).unwrap();
    writer
        .start_file(
            "META-INF/manifest.xml",
            SimpleFileOptions::default().compression_method(ZipCompressionMethod::Deflated),
        )
        .unwrap();
    writer.write_all(&manifest()).unwrap();
    writer.finish().unwrap().into_inner()
}

fn open(bytes: Vec<u8>) -> (Arc<TestSource>, Arc<SourceBackedPackage>) {
    let source = TestSource::new(bytes);
    let package = Arc::new(SourceBackedPackage::from_read_at(source_reader(&source)).unwrap());
    source.clear_ranges();
    (source, package)
}

fn insertion_offset() -> usize {
    SOURCE_CONTENT
        .windows(b"</office:text>".len())
        .position(|window| window == b"</office:text>")
        .unwrap()
}

fn fragment() -> AuthoredXmlFragment {
    AuthoredXmlFragment::markup(INSERTED_MARKUP.to_vec()).unwrap()
}

fn candidate_content() -> Vec<u8> {
    let offset = insertion_offset();
    let mut candidate = Vec::with_capacity(SOURCE_CONTENT.len() + INSERTED_MARKUP.len());
    candidate.extend_from_slice(&SOURCE_CONTENT[..offset]);
    candidate.extend_from_slice(INSERTED_MARKUP);
    candidate.extend_from_slice(&SOURCE_CONTENT[offset..]);
    candidate
}

fn xml_limits() -> XmlStreamLimits {
    XmlStreamLimits::new(SOURCE_CONTENT.len() as u64 + 512, 64, 256, 256, 1024).unwrap()
}

fn raw_members(bytes: &[u8]) -> BTreeMap<Vec<u8>, RawMember> {
    let archive = ZipArchive::from_slice(bytes).unwrap().into_zip_archive();
    let mut buffer = vec![0_u8; soapberry_zip::RECOMMENDED_BUFFER_SIZE];
    let index = PreservationIndex::new(&archive, &mut buffer).unwrap();
    index
        .entries()
        .iter()
        .map(|entry| {
            let name = entry.raw_name_bytes().to_vec();
            let local =
                bytes[entry.local_span().start as usize..entry.local_span().end as usize].to_vec();
            let central_range = entry.central_record();
            let mut central =
                bytes[central_range.start as usize..central_range.end as usize].to_vec();
            central[42..46].fill(0);
            (
                name,
                RawMember {
                    local,
                    central_without_offset: central,
                },
            )
        })
        .collect()
}

fn compressed_range(bytes: &[u8], member: &[u8]) -> Range<u64> {
    let archive = ZipArchive::from_slice(bytes).unwrap();
    archive
        .entries()
        .find_map(|entry| {
            let entry = entry.ok()?;
            if entry.file_path().as_ref() != member {
                return None;
            }
            let (start, end) = archive
                .get_entry(entry.wayfinder())
                .unwrap()
                .compressed_data_range();
            Some(start..end)
        })
        .unwrap()
}

fn plan_for(package: &Arc<SourceBackedPackage>) -> SourceContentInsertionPlan {
    SourceContentInsertionPlan::prepare(
        Arc::clone(package),
        insertion_offset() as u64,
        fragment(),
        xml_limits(),
        &SourceContentPublicationOptions::new(),
    )
    .unwrap()
}

fn execution_limits() -> ExecutionLimits {
    ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(1 << 30).unwrap(),
        0,
    )
    .unwrap()
}

fn execution_context(budget: Budget, token: CancellationToken) -> ExecutionContext {
    ExecutionContext::new(budget, token, execution_limits())
}

#[test]
fn store_and_deflate_preserve_untouched_raw_members_and_xml_prefix_suffix() {
    for content_method in [ZipCompressionMethod::Stored, ZipCompressionMethod::Deflated] {
        let bytes = zip_package(content_method, ZipCompressionMethod::Deflated);
        let before = raw_members(&bytes);
        let (_source, package) = open(bytes);
        let plan = plan_for(&package);
        let mut output = Vec::new();
        let report = plan
            .write_to(&mut output, SourceContentPublicationOptions::new())
            .unwrap();

        let reopened = OwnedPackage::from_bytes(output.clone()).unwrap();
        let content = reopened.get_file("content.xml").unwrap();
        let offset = insertion_offset();
        assert_eq!(&content[..offset], &SOURCE_CONTENT[..offset]);
        assert_eq!(
            &content[offset + INSERTED_MARKUP.len()..],
            &SOURCE_CONTENT[offset..]
        );
        assert_eq!(content, candidate_content());
        assert_eq!(report.bytes(), output.len() as u64);
        let after = raw_members(&output);
        for (name, expected) in before {
            if name.as_slice() != b"content.xml" {
                assert_eq!(
                    after.get(&name),
                    Some(&expected),
                    "untouched member {name:?}"
                );
            }
        }
    }
}

#[test]
fn invalid_offset_token_and_document_are_rejected_before_publication() {
    let bytes = zip_package(ZipCompressionMethod::Stored, ZipCompressionMethod::Stored);
    let (_source, package) = open(bytes);

    let error = SourceContentInsertionPlan::prepare(
        Arc::clone(&package),
        SOURCE_CONTENT.len() as u64 + 1,
        fragment(),
        xml_limits(),
        &SourceContentPublicationOptions::new(),
    )
    .expect_err("offset beyond decoded content must fail");
    assert!(matches!(
        error,
        SourceContentPublicationError::Core(Error::InvalidFormat(reason))
            if reason.contains("offset")
    ));

    let tiny_tokens =
        XmlStreamLimits::new(SOURCE_CONTENT.len() as u64 + 512, 64, 256, 256, 8).unwrap();
    let error = SourceContentInsertionPlan::prepare(
        Arc::clone(&package),
        insertion_offset() as u64,
        fragment(),
        tiny_tokens,
        &SourceContentPublicationOptions::new(),
    )
    .expect_err("a token ceiling below the document root must fail");
    assert!(matches!(error, SourceContentPublicationError::Core(_)));

    let unclosed = AuthoredXmlFragment::start_tag(b"<p>".to_vec()).unwrap();
    let error = SourceContentInsertionPlan::prepare(
        package,
        insertion_offset() as u64,
        unclosed,
        xml_limits(),
        &SourceContentPublicationOptions::new(),
    )
    .expect_err("an unclosed inserted element must fail document validation");
    assert!(matches!(error, SourceContentPublicationError::Core(_)));
}

#[test]
fn source_changes_after_prepare_and_during_measure_and_emit_are_primary() {
    let bytes = zip_package(ZipCompressionMethod::Stored, ZipCompressionMethod::Stored);
    let (source, package) = open(bytes.clone());
    let plan = plan_for(&package);
    source.bump_revision();
    let mut output = Vec::new();
    let error = plan
        .write_to(&mut output, SourceContentPublicationOptions::new())
        .expect_err("stale source after prepare");
    assert!(matches!(
        error,
        SourceContentInsertionError::Publication {
            source: SourceContentPublicationError::SourceChanged { .. },
            ..
        }
    ));
    assert!(output.is_empty());

    let content_range = compressed_range(&bytes, b"content.xml");
    let (measure_source, measure_package) = open(bytes.clone());
    measure_source.set_bump_range(content_range);
    let measure_plan = plan_for(&measure_package);
    measure_source.arm_bump();
    let mut output = Vec::new();
    let error = measure_plan
        .write_to(&mut output, SourceContentPublicationOptions::new())
        .expect_err("source change during replay measurement");
    assert!(matches!(
        error,
        SourceContentInsertionError::Publication {
            source: SourceContentPublicationError::SourceChanged { .. },
            ..
        }
    ));
    assert!(output.is_empty());

    let (emit_source, emit_package) = open(bytes);
    let emit_plan = plan_for(&emit_package);
    let mut sink = MutatingSink {
        source: emit_source,
        writes: 0,
        bytes: Vec::new(),
    };
    let error = emit_plan
        .write_to(&mut sink, SourceContentPublicationOptions::new())
        .expect_err("source change after replay output begins");
    assert!(matches!(
        error,
        SourceContentInsertionError::Publication {
            source: SourceContentPublicationError::SourceChanged { .. },
            ..
        }
    ));
    assert!(!sink.bytes.is_empty());
}

#[test]
fn zero_partial_and_failing_sinks_report_truthful_progress() {
    let bytes = zip_package(ZipCompressionMethod::Stored, ZipCompressionMethod::Stored);
    let (_source, package) = open(bytes);
    let plan = plan_for(&package);

    let error = plan
        .write_to(ZeroSink, SourceContentPublicationOptions::new())
        .expect_err("zero-progress sink");
    assert!(
        matches!(&error, SourceContentInsertionError::Transport {
        publication: SourceContentPublicationError::Sink { source, .. }, ..
    } if source.kind() == io::ErrorKind::WriteZero),
        "{error:?}"
    );
    assert_eq!(error.written(), 0);

    let error = plan
        .write_to(ErrorSink, SourceContentPublicationOptions::new())
        .expect_err("failing sink");
    assert!(
        matches!(&error, SourceContentInsertionError::Transport {
        publication: SourceContentPublicationError::Sink { source, .. }, ..
    } if source.kind() == io::ErrorKind::BrokenPipe),
        "{error:?}"
    );
    assert_eq!(error.written(), 0);

    let error = plan
        .write_to(
            PartialFailSink {
                accepted: 5,
                first: true,
            },
            SourceContentPublicationOptions::new(),
        )
        .expect_err("partial then failing sink");
    assert!(
        matches!(&error, SourceContentInsertionError::Transport {
        publication: SourceContentPublicationError::Sink { source, .. }, ..
    } if source.kind() == io::ErrorKind::BrokenPipe),
        "{error:?}"
    );
    assert_eq!(error.written(), 5);
}

#[test]
fn replacement_and_output_limits_fail_before_output() {
    let bytes = zip_package(ZipCompressionMethod::Deflated, ZipCompressionMethod::Stored);
    let (_source, package) = open(bytes);
    let error = SourceContentInsertionPlan::prepare(
        Arc::clone(&package),
        insertion_offset() as u64,
        fragment(),
        xml_limits(),
        &SourceContentPublicationOptions::new()
            .with_max_replacement_bytes((INSERTED_MARKUP.len() - 1) as u64),
    )
    .expect_err("replacement limit");
    assert!(matches!(
        error,
        SourceContentPublicationError::LimitExceeded { .. }
    ));

    let plan = plan_for(&package);
    let mut output = Vec::new();
    let error = plan
        .write_to(
            &mut output,
            SourceContentPublicationOptions::new().with_max_output_bytes(1),
        )
        .expect_err("output limit");
    assert!(
        matches!(
            &error,
            SourceContentInsertionError::Transport {
                publication: SourceContentPublicationError::LimitExceeded { maximum: 1, .. },
                ..
            }
        ),
        "{error:?}"
    );
    assert_eq!(error.written(), 0);
    assert!(output.is_empty());
}

#[test]
fn cancellation_is_typed_during_prepare_and_publication() {
    let bytes = zip_package(ZipCompressionMethod::Deflated, ZipCompressionMethod::Stored);

    let (cancel, token) = CancellationSource::pair();
    cancel.cancel();
    let (_source, package) = open(bytes.clone());
    let error = SourceContentInsertionPlan::prepare(
        Arc::clone(&package),
        insertion_offset() as u64,
        fragment(),
        xml_limits(),
        &SourceContentPublicationOptions::new().with_cancellation(token),
    )
    .expect_err("pre-cancelled prepare");
    assert!(matches!(
        error,
        SourceContentPublicationError::Cancelled { .. }
    ));

    let (cancel, token) = CancellationSource::pair();
    let (_source, package) = open(bytes.clone());
    let error = SourceContentInsertionPlan::prepare_with_validator(
        Arc::clone(&package),
        insertion_offset() as u64,
        fragment(),
        xml_limits(),
        &SourceContentPublicationOptions::new().with_cancellation(token),
        move |_event, _bindings| {
            cancel.cancel();
            Ok(())
        },
    )
    .expect_err("validator cancellation");
    assert!(matches!(
        error,
        SourceContentPublicationError::Cancelled { .. }
    ));

    let (_source, package) = open(bytes);
    let plan = plan_for(&package);
    let (cancel, token) = CancellationSource::pair();
    cancel.cancel();
    let mut output = Vec::new();
    let error = plan
        .write_to(
            &mut output,
            SourceContentPublicationOptions::new().with_cancellation(token),
        )
        .expect_err("pre-cancelled publication");
    assert!(matches!(
        error,
        SourceContentInsertionError::Publication {
            source: SourceContentPublicationError::Cancelled { .. },
            ..
        }
    ));
    assert!(output.is_empty());
}

#[test]
fn hierarchical_budgets_and_input_accounting_are_enforced() {
    let bytes = zip_package(ZipCompressionMethod::Stored, ZipCompressionMethod::Stored);

    let parent = Budget::root(
        "insertion-parent",
        CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, 0, u64::MAX, u64::MAX),
    );
    let child = parent.child(
        "insertion-child",
        CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, token) = CancellationSource::pair();
    let (_source, package) = open(bytes.clone());
    let error = SourceContentInsertionPlan::prepare(
        Arc::clone(&package),
        insertion_offset() as u64,
        fragment(),
        xml_limits(),
        &SourceContentPublicationOptions::new()
            .with_execution_context(execution_context(child, token)),
    )
    .expect_err("ancestor object budget");
    assert!(matches!(
        error,
        SourceContentPublicationError::Execution {
            source: ExecutionError::ResourceLimit(limit),
            ..
        } if limit.resource == Resource::Objects && limit.scope.as_ref() == "insertion-parent"
    ));

    let input_budget = Budget::root(
        "insertion-input",
        CoreLimits::new(u64::MAX, 1, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
    );
    let (_, token) = CancellationSource::pair();
    let (_source, package) = open(bytes);
    let error = SourceContentInsertionPlan::prepare(
        package,
        insertion_offset() as u64,
        fragment(),
        xml_limits(),
        &SourceContentPublicationOptions::new()
            .with_execution_context(execution_context(input_budget, token)),
    )
    .expect_err("physical source input budget");
    assert!(matches!(
        error,
        SourceContentPublicationError::Execution {
            source: ExecutionError::ResourceLimit(limit),
            ..
        } if limit.resource == Resource::InputBytes
    ));
}

#[test]
fn optional_payload_verification_reads_opaque_member_before_first_output() {
    let bytes = zip_package(
        ZipCompressionMethod::Deflated,
        ZipCompressionMethod::Deflated,
    );
    let media = compressed_range(&bytes, b"Pictures/blob.bin");
    let (source, package) = open(bytes);
    let plan = plan_for(&package);
    source.clear_ranges();
    let saw = Arc::new(AtomicBool::new(false));
    let mut sink = ObserveSink {
        source: Arc::clone(&source),
        watched: media,
        saw_watched_before_output: Arc::clone(&saw),
        bytes: Vec::new(),
    };
    plan.write_to(
        &mut sink,
        SourceContentPublicationOptions::new().with_payload_verification(true),
    )
    .unwrap();
    assert!(saw.load(Ordering::Acquire));
}

#[test]
fn memory_and_work_admission_precede_source_reads() {
    for (resource, limits) in [
        (
            Resource::Memory,
            CoreLimits::new(0, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX),
        ),
        (
            Resource::Work,
            CoreLimits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, 0),
        ),
    ] {
        let (source, package) = open(zip_package(
            ZipCompressionMethod::Stored,
            ZipCompressionMethod::Stored,
        ));
        let (_, token) = CancellationSource::pair();
        let options = SourceContentPublicationOptions::new().with_execution_context(
            execution_context(Budget::root("insertion-admission", limits), token),
        );
        let error = SourceContentInsertionPlan::prepare(
            package,
            insertion_offset() as u64,
            fragment(),
            xml_limits(),
            &options,
        )
        .unwrap_err();
        assert!(matches!(error, SourceContentPublicationError::Execution {
            source: ExecutionError::ResourceLimit(limit), ..
        } if limit.resource == resource));
        assert!(source.state.lock().unwrap().ranges.is_empty());
    }
}

#[test]
fn retained_fragment_capacity_refuses_before_source_reads() {
    const SPARE_CAPACITY: usize = 64 * 1024 * 1024;

    let (source, package) = open(zip_package(
        ZipCompressionMethod::Stored,
        ZipCompressionMethod::Stored,
    ));
    let mut bytes = Vec::with_capacity(SPARE_CAPACITY);
    bytes.extend_from_slice(INSERTED_MARKUP);
    let retained_capacity = u64::try_from(bytes.capacity()).unwrap();
    let fragment = AuthoredXmlFragment::markup(bytes).unwrap();
    let (_, token) = CancellationSource::pair();
    let options = SourceContentPublicationOptions::new().with_execution_context(execution_context(
        Budget::root(
            "insertion-retained-refusal",
            CoreLimits::new(
                retained_capacity - 1,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
                u64::MAX,
            ),
        ),
        token,
    ));

    let error = SourceContentInsertionPlan::prepare(
        package,
        insertion_offset() as u64,
        fragment,
        xml_limits(),
        &options,
    )
    .unwrap_err();
    assert!(matches!(
        error,
        SourceContentPublicationError::Execution {
            source: ExecutionError::ResourceLimit(limit),
            ..
        } if limit.resource == Resource::Memory
    ));
    assert!(source.state.lock().unwrap().ranges.is_empty());
}

#[test]
fn retained_fragment_capacity_stays_charged_until_plan_drop() {
    const SPARE_CAPACITY: usize = 64 * 1024 * 1024;

    let (_source, package) = open(zip_package(
        ZipCompressionMethod::Stored,
        ZipCompressionMethod::Stored,
    ));
    let mut bytes = Vec::with_capacity(SPARE_CAPACITY);
    bytes.extend_from_slice(INSERTED_MARKUP);
    let retained_capacity = u64::try_from(bytes.capacity()).unwrap();
    let fragment = AuthoredXmlFragment::markup(bytes).unwrap();
    let budget = Budget::root(
        "insertion-retained-lifetime",
        CoreLimits::new(
            retained_capacity + 16 * 1024 * 1024,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
            u64::MAX,
        ),
    );
    let (_, token) = CancellationSource::pair();
    let options = SourceContentPublicationOptions::new()
        .with_execution_context(execution_context(budget.clone(), token));

    let plan = SourceContentInsertionPlan::prepare(
        package,
        insertion_offset() as u64,
        fragment,
        xml_limits(),
        &options,
    )
    .expect("retained fragment should be admitted");
    assert_eq!(budget.used(Resource::Memory), retained_capacity);

    drop(plan);
    assert_eq!(budget.used(Resource::Memory), 0);
}

#[test]
fn sink_output_budget_retains_public_execution_cause() {
    let (_source, package) = open(zip_package(
        ZipCompressionMethod::Stored,
        ZipCompressionMethod::Stored,
    ));
    let plan = plan_for(&package);
    let (_, token) = CancellationSource::pair();
    let budget = Budget::root(
        "insertion-output",
        CoreLimits::new(u64::MAX, u64::MAX, 80, u64::MAX, u64::MAX, u64::MAX),
    );
    let options = SourceContentPublicationOptions::new()
        .with_execution_context(execution_context(budget, token));
    let mut output = Vec::new();
    let error = plan.write_to(&mut output, &options).unwrap_err();
    assert!(matches!(&error, SourceContentInsertionError::Transport {
        publication: SourceContentPublicationError::Execution {
            source: ExecutionError::ResourceLimit(limit), ..
        }, ..
    } if limit.resource == Resource::OutputBytes));
    assert_eq!(error.written(), output.len() as u64);
    assert!(!output.is_empty() && output.len() <= 80);
}

#[test]
fn cancellation_after_first_sink_write_retains_typed_cause_and_prefix() {
    struct CancelSink {
        bytes: Vec<u8>,
        cancel: CancellationSource,
    }
    impl Write for CancelSink {
        fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
            self.bytes.extend_from_slice(bytes);
            self.cancel.cancel();
            Ok(bytes.len())
        }
        fn flush(&mut self) -> io::Result<()> {
            Ok(())
        }
    }
    let (_source, package) = open(zip_package(
        ZipCompressionMethod::Stored,
        ZipCompressionMethod::Stored,
    ));
    let plan = plan_for(&package);
    let (cancel, token) = CancellationSource::pair();
    let mut output = CancelSink {
        bytes: Vec::new(),
        cancel,
    };
    let error = plan
        .write_to(
            &mut output,
            SourceContentPublicationOptions::new().with_cancellation(token),
        )
        .unwrap_err();
    assert!(matches!(
        &error,
        SourceContentInsertionError::Transport {
            publication: SourceContentPublicationError::Cancelled { .. },
            ..
        }
    ));
    assert_eq!(error.written(), output.bytes.len() as u64);
    assert!(!output.bytes.is_empty());
}
