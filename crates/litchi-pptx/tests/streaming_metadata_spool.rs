#![allow(
    clippy::expect_used,
    clippy::indexing_slicing,
    clippy::panic,
    clippy::unwrap_used,
    reason = "These integration tests use small, deterministic fault-injection fixtures."
)]

//! Public PPTX coverage for the explicit metadata scratch writer.
//!
//! The byte comparisons in this file are differential checks against the
//! existing ordinary writer.  Reopening goes through the PPTX and OPC
//! readers, so the tests validate the authored graph and every physical
//! member without coupling the assertions to ZIP implementation details.

use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};
use std::sync::{Arc, Mutex};

use litchi_opc::OpcError;
use litchi_opc::phys_pkg::PhysPkgReader;
use litchi_pptx::{
    Error, Package, StreamingPresentationLimits, StreamingPresentationOptions,
    StreamingPresentationScratchLimits, StreamingPresentationWriter, TextBoxSpec,
};

const SCRATCH_BUFFER_BYTES: usize = 11;
const SCRATCH_MAX_BYTES: u64 = 16 * 1024 * 1024;

fn scratch_limits(max_bytes: u64) -> StreamingPresentationScratchLimits {
    StreamingPresentationScratchLimits {
        max_bytes,
        buffer_bytes: SCRATCH_BUFFER_BYTES,
    }
}

fn authored_text(index: usize) -> String {
    format!("authored slide {index}: & <text>")
}

fn authored_title(index: usize) -> String {
    format!("title {index}")
}

fn authored_geometry(index: usize) -> (i64, i64, i64, i64) {
    let index = i64::try_from(index).expect("test index fits EMUs");
    (100_000 + index, 200_000 + index, 3_000_000, 700_000)
}

fn write_slides<W: Write>(
    mut writer: StreamingPresentationWriter<W>,
    slide_count: usize,
) -> Result<W, Error> {
    for index in 0..slide_count {
        let title = authored_title(index);
        let text = authored_text(index);
        let (x, y, width, height) = authored_geometry(index);
        let mut slide = writer.start_slide(Some(&title))?;
        slide.write_text_box(TextBoxSpec::new(&text, x, y, width, height))?;
        writer = slide.finish()?;
    }
    writer.finish()
}

fn ordinary_deck(slide_count: usize, options: StreamingPresentationOptions) -> Vec<u8> {
    let writer = StreamingPresentationWriter::with_options(
        Vec::new(),
        slide_count,
        options,
        StreamingPresentationLimits::default(),
    )
    .expect("ordinary writer construction");
    write_slides(writer, slide_count).expect("ordinary deck")
}

fn spooled_deck(slide_count: usize, options: StreamingPresentationOptions) -> Vec<u8> {
    let writer = StreamingPresentationWriter::with_options_and_metadata_spool(
        Vec::new(),
        slide_count,
        options,
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    )
    .expect("spooled writer construction");
    write_slides(writer, slide_count).expect("spooled deck")
}

fn assert_reopens_all_authored_content(
    bytes: &[u8],
    slide_count: usize,
    options: StreamingPresentationOptions,
) {
    let package = Package::from_bytes(bytes).expect("PPTX must reopen");
    let presentation = package.presentation().expect("presentation graph");
    assert_eq!(
        presentation.slide_count().expect("slide count"),
        slide_count
    );
    assert_eq!(
        presentation.slide_size().expect("slide size"),
        (options.slide_width(), options.slide_height())
    );

    for index in 0..slide_count {
        let slide = presentation
            .slide(index)
            .expect("slide lookup")
            .expect("declared slide exists");
        let title = authored_title(index);
        let text = authored_text(index);
        let (x, y, width, height) = authored_geometry(index);
        let slide_text = slide.text().expect("slide text");
        assert!(
            slide_text.contains(&title),
            "missing title on slide {index}"
        );
        assert!(slide_text.contains(&text), "missing text on slide {index}");

        assert_eq!(slide.shape_count().expect("shape count"), 2);
        let shapes = slide.shapes().expect("slide shapes");
        assert_eq!(
            shapes.shape(0).expect("title shape").text(),
            Some(title.as_str())
        );
        let authored_shape = shapes.shape(1).expect("authored shape");
        assert_eq!(authored_shape.text(), Some(text.as_str()));
        let bounds = authored_shape.bounds().expect("authored geometry");
        assert_eq!(bounds.x(), x);
        assert_eq!(bounds.y(), y);
        assert_eq!(bounds.width(), width);
        assert_eq!(bounds.height(), height);
    }

    let physical = PhysPkgReader::new(bytes).expect("OPC physical reader");
    let names = physical.member_names().expect("physical member names");
    assert!(!names.is_empty());
    for name in &names {
        let _member = physical
            .read_member(name)
            .unwrap_or_else(|error| panic!("physical member {name} must be readable: {error}"));
    }
}

#[test]
fn explicit_metadata_spool_matches_ordinary_bytes_for_default_and_widescreen_decks() {
    for options in [
        StreamingPresentationOptions::default(),
        StreamingPresentationOptions::widescreen(),
    ] {
        for slide_count in [1, 8, 256] {
            let ordinary = ordinary_deck(slide_count, options);
            let spooled = spooled_deck(slide_count, options);
            assert_eq!(
                spooled, ordinary,
                "explicit scratch changed bytes for {slide_count} slides"
            );
            assert_reopens_all_authored_content(&spooled, slide_count, options);
        }
    }
}

#[test]
fn convenience_metadata_spool_matches_the_default_writer() {
    let mut writer = StreamingPresentationWriter::with_metadata_spool(
        Vec::new(),
        1,
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    )
    .expect("convenience spooled writer");
    let title = authored_title(0);
    let text = authored_text(0);
    let (x, y, width, height) = authored_geometry(0);
    let mut slide = writer.start_slide(Some(&title)).expect("start slide");
    slide
        .write_text_box(TextBoxSpec::new(&text, x, y, width, height))
        .expect("write box");
    writer = slide.finish().expect("finish slide");
    let spooled = writer.finish().expect("finish archive");
    assert_eq!(
        spooled,
        ordinary_deck(1, StreamingPresentationOptions::default())
    );
    assert_reopens_all_authored_content(&spooled, 1, StreamingPresentationOptions::default());
}

#[test]
fn explicit_spool_preserves_validation_for_hostile_content_and_slide_order() {
    let output = FaultOutput::with_limit(usize::MAX);
    let output_state = output.clone();
    let writer = StreamingPresentationWriter::with_options_and_metadata_spool(
        output,
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    )
    .expect("spooled writer");
    let before = output_state.len();
    assert!(matches!(
        writer.start_slide(Some("hostile\u{1}")),
        Err(Error::Invalid(_))
    ));
    assert_eq!(output_state.len(), before);

    let mut writer = StreamingPresentationWriter::with_options_and_metadata_spool(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    )
    .expect("spooled writer");
    let mut slide = writer.start_slide(None).expect("start slide");
    let before = slide.output_bytes();
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("bad\u{1}", 0, 0, 1, 1)),
        Err(Error::Invalid(_))
    ));
    assert_eq!(slide.output_bytes(), before);
    assert!(matches!(
        slide.write_text_box(TextBoxSpec::new("bad", -1, 0, 1, 1)),
        Err(Error::Invalid(_))
    ));
    writer = slide.finish().expect("finish empty slide");
    writer.finish().expect("complete one-slide archive");

    let mut missing = StreamingPresentationWriter::with_options_and_metadata_spool(
        Vec::new(),
        2,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    )
    .expect("missing-slide writer");
    let slide = missing.start_slide(None).expect("first slide");
    missing = slide.finish().expect("first slide finish");
    assert!(matches!(missing.finish(), Err(Error::Invalid(_))));

    let mut extra = StreamingPresentationWriter::with_options_and_metadata_spool(
        Vec::new(),
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    )
    .expect("extra-slide writer");
    let slide = extra.start_slide(None).expect("only slide");
    extra = slide.finish().expect("only slide finish");
    assert!(matches!(extra.start_slide(None), Err(Error::Invalid(_))));
}

#[test]
fn explicit_spool_refuses_invalid_configuration_before_output() {
    let mut output = Vec::new();
    let invalid_options = StreamingPresentationOptions::new(1, 1).unwrap_or_default();
    let error = match StreamingPresentationWriter::with_options_and_metadata_spool(
        &mut output,
        0,
        invalid_options,
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    ) {
        Ok(_) => panic!("zero slides must be refused"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Invalid(_)));
    assert!(output.is_empty());

    let error = match StreamingPresentationWriter::with_options_and_metadata_spool(
        &mut output,
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits {
            max_output_bytes: 0,
            ..StreamingPresentationLimits::default()
        },
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    ) {
        Ok(_) => panic!("zero output limit must be refused"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Invalid(_)));
    assert!(output.is_empty());

    let error = match StreamingPresentationWriter::with_options_and_metadata_spool(
        &mut output,
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        StreamingPresentationScratchLimits {
            max_bytes: 0,
            buffer_bytes: SCRATCH_BUFFER_BYTES,
        },
    ) {
        Ok(_) => panic!("zero scratch quota must be refused"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Invalid(_)));
    assert!(output.is_empty());

    let error = match StreamingPresentationWriter::with_options_and_metadata_spool(
        &mut output,
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        StreamingPresentationScratchLimits {
            max_bytes: SCRATCH_MAX_BYTES,
            buffer_bytes: 0,
        },
    ) {
        Ok(_) => panic!("zero scratch buffer must be refused"),
        Err(error) => error,
    };
    assert!(matches!(error, Error::Invalid(_)));
    assert!(output.is_empty());
}

#[test]
fn scratch_quota_failure_is_typed_and_does_not_become_a_generic_pptx_error() {
    let mut output = Vec::new();
    let error = match StreamingPresentationWriter::with_options_and_metadata_spool(
        &mut output,
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
        Cursor::new(Vec::new()),
        scratch_limits(1),
    ) {
        Ok(_) => panic!("one byte cannot hold the generated central metadata"),
        Err(error) => error,
    };
    assert!(has_spool_error(&error));
    assert!(matches!(
        error,
        Error::Opc(OpcError::CentralDirectorySpoolLimitExceeded { .. })
            | Error::Opc(OpcError::IncompleteOutput { .. })
    ));
}

#[test]
fn output_limit_reports_the_exact_accepted_prefix_in_the_spooled_route() {
    let reference = ordinary_deck(1, StreamingPresentationOptions::default());
    let mut output = Vec::new();
    let limits = StreamingPresentationLimits {
        max_output_bytes: u64::try_from(reference.len() - 1).expect("reference fits u64"),
        ..StreamingPresentationLimits::default()
    };
    let writer = StreamingPresentationWriter::with_options_and_metadata_spool(
        &mut output,
        1,
        StreamingPresentationOptions::default(),
        limits,
        Cursor::new(Vec::new()),
        scratch_limits(SCRATCH_MAX_BYTES),
    )
    .expect("one-byte-under runtime limit is admitted by structural preflight");
    let error = match write_slides(writer, 1) {
        Ok(_) => panic!("one byte below the complete archive must fail"),
        Err(error) => error,
    };
    let written = accepted_output_bytes(&error).expect("output-limit progress");
    assert_eq!(
        written,
        u64::try_from(output.len()).expect("output length fits u64")
    );
    assert_eq!(
        written,
        u64::try_from(reference.len() - 1).expect("reference fits u64")
    );
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum StoreFault {
    None,
    Seek,
    Append,
    Replay,
}

#[derive(Debug)]
struct StoreState {
    bytes: Vec<u8>,
    position: u64,
    short_reads: bool,
    short_writes: bool,
    interrupt_reads: usize,
    interrupt_writes: usize,
    fault: StoreFault,
}

#[derive(Clone, Debug)]
struct FaultStore {
    state: Arc<Mutex<StoreState>>,
}

impl FaultStore {
    fn new() -> Self {
        Self {
            state: Arc::new(Mutex::new(StoreState {
                bytes: Vec::new(),
                position: 0,
                short_reads: false,
                short_writes: false,
                interrupt_reads: 0,
                interrupt_writes: 0,
                fault: StoreFault::None,
            })),
        }
    }

    fn with_initial_fault(fault: StoreFault) -> Self {
        let store = Self::new();
        store.state.lock().expect("store lock").fault = fault;
        store
    }

    fn enable_short_transfers(&self) {
        let mut state = self.state.lock().expect("store lock");
        state.short_reads = true;
        state.short_writes = true;
    }

    fn interrupt_once_each_direction(&self) {
        let mut state = self.state.lock().expect("store lock");
        state.interrupt_reads = 1;
        state.interrupt_writes = 1;
    }

    fn fail_next_append(&self) {
        self.state.lock().expect("store lock").fault = StoreFault::Append;
    }

    fn fail_next_replay(&self) {
        self.state.lock().expect("store lock").fault = StoreFault::Replay;
    }

    fn snapshot_bytes(&self) -> Vec<u8> {
        self.state.lock().expect("store lock").bytes.clone()
    }
}

impl Read for FaultStore {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let mut state = self.state.lock().expect("store lock");
        if state.interrupt_reads != 0 {
            state.interrupt_reads -= 1;
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry read"));
        }
        if state.fault == StoreFault::Replay {
            return Err(io::Error::other("replay read fault"));
        }
        let position = usize::try_from(state.position)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "position overflow"))?;
        if position >= state.bytes.len() || buffer.is_empty() {
            return Ok(0);
        }
        let available = state.bytes.len() - position;
        let mut amount = available.min(buffer.len());
        if state.short_reads {
            amount = amount.min(1);
        }
        buffer[..amount].copy_from_slice(&state.bytes[position..position + amount]);
        state.position = state
            .position
            .checked_add(u64::try_from(amount).expect("amount fits u64"))
            .expect("position fits u64");
        Ok(amount)
    }
}

impl Write for FaultStore {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let mut state = self.state.lock().expect("store lock");
        if state.interrupt_writes != 0 {
            state.interrupt_writes -= 1;
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry write"));
        }
        if state.fault == StoreFault::Append {
            return Err(io::Error::other("append write fault"));
        }
        if buffer.is_empty() {
            return Ok(0);
        }
        let position = usize::try_from(state.position)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "position overflow"))?;
        let amount = if state.short_writes {
            buffer.len().min(1)
        } else {
            buffer.len()
        };
        let end = position
            .checked_add(amount)
            .ok_or_else(|| io::Error::new(io::ErrorKind::InvalidInput, "store length overflow"))?;
        if end > state.bytes.len() {
            state.bytes.resize(end, 0);
        }
        state.bytes[position..end].copy_from_slice(&buffer[..amount]);
        state.position = u64::try_from(end).expect("store length fits u64");
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Seek for FaultStore {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        let mut state = self.state.lock().expect("store lock");
        if state.fault == StoreFault::Seek {
            return Err(io::Error::other("initial seek fault"));
        }
        let length = i128::try_from(state.bytes.len()).expect("length fits i128");
        let current = i128::from(state.position);
        let target = match position {
            SeekFrom::Start(value) => i128::from(value),
            SeekFrom::Current(value) => current + i128::from(value),
            SeekFrom::End(value) => length + i128::from(value),
        };
        if target < 0 {
            return Err(io::Error::new(io::ErrorKind::InvalidInput, "negative seek"));
        }
        state.position = u64::try_from(target)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "seek overflow"))?;
        Ok(state.position)
    }
}

#[derive(Clone, Debug)]
struct OutputState {
    bytes: Vec<u8>,
    remaining: usize,
    interrupt_once: bool,
}

#[derive(Clone, Debug)]
struct FaultOutput {
    state: Arc<Mutex<OutputState>>,
}

impl FaultOutput {
    fn with_limit(remaining: usize) -> Self {
        Self {
            state: Arc::new(Mutex::new(OutputState {
                bytes: Vec::new(),
                remaining,
                interrupt_once: false,
            })),
        }
    }

    fn enable_interrupt_once(&self) {
        self.state.lock().expect("output lock").interrupt_once = true;
    }

    fn len(&self) -> usize {
        self.state.lock().expect("output lock").bytes.len()
    }

    fn bytes(&self) -> Vec<u8> {
        self.state.lock().expect("output lock").bytes.clone()
    }
}

impl Write for FaultOutput {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let mut state = self.state.lock().expect("output lock");
        if state.interrupt_once {
            state.interrupt_once = false;
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry output"));
        }
        if state.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::WriteZero, "output fault"));
        }
        let amount = buffer.len().min(state.remaining);
        state.bytes.extend_from_slice(&buffer[..amount]);
        state.remaining -= amount;
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn construct_fault_writer(
    output: FaultOutput,
    store: FaultStore,
) -> Result<StreamingPresentationWriter<FaultOutput>, Error> {
    StreamingPresentationWriter::with_options_and_metadata_spool(
        output,
        1,
        StreamingPresentationOptions::default(),
        StreamingPresentationLimits::default(),
        store,
        scratch_limits(SCRATCH_MAX_BYTES),
    )
}

fn has_spool_error(error: &Error) -> bool {
    match error {
        Error::Opc(OpcError::CentralDirectorySpool { .. })
        | Error::Opc(OpcError::CentralDirectorySpoolLimitExceeded { .. }) => true,
        Error::Opc(OpcError::IncompleteOutput { source, .. }) => has_spool_opc_error(source),
        _ => false,
    }
}

fn has_spool_opc_error(error: &OpcError) -> bool {
    match error {
        OpcError::CentralDirectorySpool { .. }
        | OpcError::CentralDirectorySpoolLimitExceeded { .. } => true,
        OpcError::IncompleteOutput { source, .. } => has_spool_opc_error(source),
        _ => false,
    }
}

fn accepted_output_bytes(error: &Error) -> Option<u64> {
    match error {
        Error::Opc(OpcError::IncompleteOutput { written, .. }) => Some(*written),
        _ => None,
    }
}

#[test]
fn initialization_spool_failure_is_typed_and_leaves_output_empty() {
    let output = FaultOutput::with_limit(usize::MAX);
    let output_state = output.clone();
    let store = FaultStore::with_initial_fault(StoreFault::Seek);
    let error = match construct_fault_writer(output, store) {
        Ok(_) => panic!("initial seek must fail"),
        Err(error) => error,
    };
    assert!(matches!(
        error,
        Error::Opc(OpcError::CentralDirectorySpool { .. })
    ));
    assert_eq!(output_state.len(), 0);
}

#[test]
fn short_and_interrupted_scratch_transfers_preserve_the_deck() {
    let output = FaultOutput::with_limit(usize::MAX);
    let output_state = output.clone();
    let store = FaultStore::new();
    store.enable_short_transfers();
    store.interrupt_once_each_direction();
    let writer = construct_fault_writer(output, store.clone()).expect("fault store construction");
    let output = write_slides(writer, 1).expect("short/interrupted deck");
    let bytes = output.bytes();
    assert!(!bytes.is_empty());
    assert_eq!(output_state.len(), bytes.len());
    assert_reopens_all_authored_content(&bytes, 1, StreamingPresentationOptions::default());
    assert!(!store.snapshot_bytes().is_empty());
}

#[test]
fn appending_spool_failure_is_typed_after_output_progress() {
    let output = FaultOutput::with_limit(usize::MAX);
    let output_state = output.clone();
    let store = FaultStore::new();
    let writer = construct_fault_writer(output, store.clone()).expect("fault store construction");
    let mut slide = writer
        .start_slide(Some("append failure"))
        .expect("start slide");
    slide
        .write_text_box(TextBoxSpec::new("payload", 0, 0, 1, 1))
        .expect("write payload");
    store.fail_next_append();
    let error = match slide.finish() {
        Ok(_) => panic!("append fault must fail finalization"),
        Err(error) => error,
    };
    assert!(has_spool_error(&error));
    let written = accepted_output_bytes(&error).expect("append failure progress");
    assert_eq!(
        written,
        u64::try_from(output_state.len()).expect("output length fits u64")
    );
    assert!(written > 0, "slide bytes precede spool publication");
}

#[test]
fn replay_spool_failure_reports_exactly_the_output_bytes_accepted() {
    let output = FaultOutput::with_limit(usize::MAX);
    let output_state = output.clone();
    let store = FaultStore::new();
    let writer = construct_fault_writer(output, store.clone()).expect("fault store construction");
    let mut slide = writer
        .start_slide(Some("replay failure"))
        .expect("start slide before replay fault");
    slide
        .write_text_box(TextBoxSpec::new("payload", 0, 0, 1, 1))
        .expect("write payload before replay fault");
    let writer = slide.finish().expect("finish slide before replay fault");
    store.fail_next_replay();
    let error = writer
        .finish()
        .expect_err("replay fault must fail archive finish");
    assert!(has_spool_error(&error));
    let written = accepted_output_bytes(&error).expect("replay has incomplete output progress");
    assert_eq!(
        written,
        u64::try_from(output_state.len()).expect("output length fits u64")
    );
}

#[test]
fn output_failure_reports_accepted_bytes_and_preserves_the_typed_progress_wrapper() {
    let output = FaultOutput::with_limit(256);
    let output_state = output.clone();
    let store = FaultStore::new();
    let error = construct_fault_writer(output, store)
        .and_then(|writer| write_slides(writer, 1))
        .expect_err("short output must eventually fail");
    let written = accepted_output_bytes(&error).expect("output failure progress");
    assert_eq!(
        written,
        u64::try_from(output_state.len()).expect("output length fits u64")
    );
    assert!(written > 0);
}

#[test]
fn interrupted_output_is_retryable_in_the_explicit_scratch_path() {
    let output = FaultOutput::with_limit(usize::MAX);
    output.enable_interrupt_once();
    let output_state = output.clone();
    let writer = construct_fault_writer(output, FaultStore::new()).expect("construction");
    let output = write_slides(writer, 1).expect("Interrupted output should be retried");
    let bytes = output.bytes();
    assert_eq!(bytes.len(), output_state.len());
    assert_reopens_all_authored_content(&bytes, 1, StreamingPresentationOptions::default());
}
