#![allow(
    clippy::expect_used,
    clippy::unwrap_used,
    reason = "these tests intentionally panic on an unexpected transaction result"
)]

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionLimits, Limits as CoreLimits, ReadAt,
    Resource, SourceVersion,
};
use litchi_rtf::Document;
use litchi_rtf::tail_append::{
    PlainParagraph, PlainRun, TailAppendError, TailAppendLimits, TailAppendOutputProgress,
    TailAppendPublicationError, TailAppendPublicationLimits, TailAppendSourceEdit,
    TailAppendSourceProof, TailSelector,
};
use std::io::{self, Write};
use std::num::{NonZeroU64, NonZeroUsize};
use std::sync::{
    Arc, Mutex,
    atomic::{AtomicU64, Ordering},
};
use std::{sync::mpsc, thread};

static NEXT_SOURCE_ID: AtomicU64 = AtomicU64::new(50_000);

struct SourceState {
    bytes: Vec<u8>,
    revision: u64,
    max_read: usize,
    fail_after: Option<usize>,
    zero_after: Option<usize>,
    overreport: bool,
    reads: usize,
}

struct TestSource {
    id: u64,
    state: Mutex<SourceState>,
}

impl TestSource {
    fn new(bytes: Vec<u8>, max_read: usize) -> Arc<Self> {
        Arc::new(Self {
            id: NEXT_SOURCE_ID.fetch_add(1, Ordering::Relaxed),
            state: Mutex::new(SourceState {
                bytes,
                revision: 0,
                max_read,
                fail_after: None,
                zero_after: None,
                overreport: false,
                reads: 0,
            }),
        })
    }

    fn mutate(&self) {
        let mut state = self.state.lock().unwrap();
        if let Some(byte) = state.bytes.iter_mut().find(|byte| **byte == b'A') {
            *byte = b'Z';
        } else if let Some(byte) = state.bytes.get_mut(1) {
            *byte = if *byte == b'{' { b' ' } else { b'{' };
        }
        state.revision = state.revision.saturating_add(1);
    }

    fn fail_reads_after(&self, reads: usize) {
        self.state.lock().unwrap().fail_after = Some(reads);
    }

    fn overreport_reads(&self) {
        self.state.lock().unwrap().overreport = true;
    }

    fn zero_reads_from_current(&self) {
        let mut state = self.state.lock().unwrap();
        state.zero_after = Some(state.reads);
    }
}

impl ReadAt for TestSource {
    fn len(&self) -> io::Result<u64> {
        let state = self.state.lock().unwrap();
        u64::try_from(state.bytes.len())
            .map_err(|_| io::Error::other("test source length overflows u64"))
    }

    fn read_at(&self, offset: u64, output: &mut [u8]) -> io::Result<usize> {
        let mut state = self.state.lock().unwrap();
        let read_index = state.reads;
        state.reads = state.reads.saturating_add(1);
        if state.zero_after.is_some_and(|limit| read_index >= limit) {
            return Ok(0);
        }
        if state.fail_after.is_some_and(|limit| read_index >= limit) {
            return Err(io::Error::new(
                io::ErrorKind::BrokenPipe,
                "test source failure",
            ));
        }
        if state.overreport {
            return Ok(output.len().saturating_add(1));
        }
        let start = usize::try_from(offset)
            .map_err(|_| io::Error::new(io::ErrorKind::InvalidInput, "offset overflow"))?;
        let Some(input) = state.bytes.get(start..) else {
            return Ok(0);
        };
        let count = input.len().min(output.len()).min(state.max_read);
        output[..count].copy_from_slice(&input[..count]);
        Ok(count)
    }

    fn version(&self) -> io::Result<SourceVersion> {
        Ok(SourceVersion::new(
            self.id,
            self.state.lock().unwrap().revision,
        ))
    }
}

fn context(
    memory: u64,
    input: u64,
    output: u64,
    work: u64,
) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root(
        "rtf-source-tail-test",
        CoreLimits::new(memory, input, output, 1_000_000, 64, work),
    );
    let (source, token) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).unwrap(),
        NonZeroUsize::new(1).unwrap(),
        NonZeroU64::new(64 * 1024).unwrap(),
        0,
    )
    .unwrap();
    (
        budget.clone(),
        source,
        ExecutionContext::new(budget, token, execution),
    )
}

fn limits() -> TailAppendLimits {
    TailAppendLimits::new(8, 32, 128, 1024, 4096, 16 * 1024)
}

fn publication_limits() -> TailAppendPublicationLimits {
    TailAppendPublicationLimits::new(3, 2)
}

fn fixture() -> (Vec<u8>, Document) {
    let bytes = br#"{\rtf1\ansi A\par B}"#.to_vec();
    let document = Document::from_bytes(&bytes).unwrap();
    (bytes, document)
}

fn make_plan(
    source: Arc<TestSource>,
    document: &Document,
    text: &str,
) -> litchi_rtf::tail_append::TailAppendSourcePublicationPlan {
    let proof = TailAppendSourceProof::from_document_for_source(
        document,
        TailSelector::Body,
        source.as_ref(),
    )
    .unwrap();
    let mut edit = TailAppendSourceEdit::with_limits(source, proof, TailSelector::Body, limits());
    edit.append_text_paragraphs(&[text]).unwrap();
    edit.publication_plan_with_limits(publication_limits())
        .unwrap()
}

#[test]
fn source_plan_is_exact_bounded_and_reopenable_after_document_drop() {
    let (bytes, document) = fixture();
    let source = TestSource::new(bytes.clone(), 1);
    let plan = make_plan(source.clone(), &document, "C");
    drop(document);

    let (_budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    let mut output = Vec::new();
    let report = plan.write_to(&mut output, &execution).unwrap();
    assert_eq!(report.source_bytes(), bytes.len());
    assert_eq!(report.inserted_bytes(), plan.inserted_bytes());
    assert!(report.largest_write() <= 2);
    let reopened = Document::from_bytes(&output).unwrap();
    assert_eq!(reopened.text(), "A\nB\nC\n");
}

#[test]
fn source_plan_noop_preserves_exact_bytes() {
    let (bytes, document) = fixture();
    let source = TestSource::new(bytes.clone(), 2);
    let proof = TailAppendSourceProof::from_document_for_source(
        &document,
        TailSelector::Body,
        source.as_ref(),
    )
    .unwrap();
    let edit = TailAppendSourceEdit::with_limits(source, proof, TailSelector::Body, limits());
    let plan = edit
        .publication_plan_with_limits(publication_limits())
        .unwrap();
    assert!(plan.is_noop());
    let (_budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    let mut output = Vec::new();
    plan.write_to(&mut output, &execution).unwrap();
    assert_eq!(output, bytes);
}

#[test]
fn source_proof_uses_parser_break_tokens_not_raw_suffixes() {
    let escaped = br#"{\rtf1\ansi A \\par}"#.to_vec();
    let escaped_document = Document::from_bytes(&escaped).unwrap();
    let escaped_source = TestSource::new(escaped, 2);
    let escaped_proof = TailAppendSourceProof::from_document_for_source(
        &escaped_document,
        TailSelector::Body,
        escaped_source.as_ref(),
    )
    .unwrap();
    assert!(!escaped_proof.ends_with_par());
    let mut escaped_edit = TailAppendSourceEdit::with_limits(
        escaped_source,
        escaped_proof,
        TailSelector::Body,
        limits(),
    );
    escaped_edit.append_text_paragraphs(&["C"]).unwrap();
    let escaped_plan = escaped_edit
        .publication_plan_with_limits(publication_limits())
        .unwrap();
    let (_budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    let mut escaped_output = Vec::new();
    escaped_plan
        .write_to(&mut escaped_output, &execution)
        .unwrap();
    assert_eq!(
        Document::from_bytes(&escaped_output).unwrap().text(),
        "A \\par\nC\n"
    );

    let real = br#"{\rtf1\ansi A\par\b}"#.to_vec();
    let real_document = Document::from_bytes(&real).unwrap();
    let real_source = TestSource::new(real, 2);
    let real_proof = TailAppendSourceProof::from_document_for_source(
        &real_document,
        TailSelector::Body,
        real_source.as_ref(),
    )
    .unwrap();
    assert!(real_proof.ends_with_par());
    let mut real_edit =
        TailAppendSourceEdit::with_limits(real_source, real_proof, TailSelector::Body, limits());
    real_edit.append_text_paragraphs(&["C"]).unwrap();
    let real_plan = real_edit
        .publication_plan_with_limits(publication_limits())
        .unwrap();
    let mut real_output = Vec::new();
    real_plan.write_to(&mut real_output, &execution).unwrap();
    assert_eq!(Document::from_bytes(&real_output).unwrap().text(), "A\nC\n");
}

#[test]
fn context_aware_source_proof_charges_digest_and_honors_cancellation() {
    let (bytes, document) = fixture();
    let source = TestSource::new(bytes.clone(), 2);
    let (budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    TailAppendSourceProof::from_document_for_source_with_context(
        &document,
        TailSelector::Body,
        source.as_ref(),
        &execution,
    )
    .unwrap();
    assert_eq!(budget.used(Resource::InputBytes), bytes.len() as u64);
    assert_eq!(budget.used(Resource::Work), bytes.len() as u64);

    let source = TestSource::new(bytes, 2);
    let (_budget, cancel, execution) = context(4096, 4096, 4096, 4096);
    cancel.cancel();
    let error = TailAppendSourceProof::from_document_for_source_with_context(
        &document,
        TailSelector::Body,
        source.as_ref(),
        &execution,
    )
    .unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::Execution { written: 0, .. }
    ));
}

#[test]
fn context_aware_source_planning_charges_digest_scan() {
    let (bytes, document) = fixture();
    let source = TestSource::new(bytes.clone(), 2);
    let proof = TailAppendSourceProof::from_document_for_source(
        &document,
        TailSelector::Body,
        source.as_ref(),
    )
    .unwrap();
    let mut edit = TailAppendSourceEdit::with_limits(source, proof, TailSelector::Body, limits());
    edit.append_text_paragraphs(&["C"]).unwrap();
    let (budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    let plan = edit
        .publication_plan_with_limits_and_context(publication_limits(), &execution)
        .unwrap();
    let one_scan = (plan.source_bytes() + plan.inserted_bytes()) as u64;
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::InputBytes), one_scan);
    assert_eq!(budget.used(Resource::Work), one_scan);
}

#[test]
fn source_publication_charges_three_passes_and_exact_window_memory() {
    let (bytes, document) = fixture();
    let source = TestSource::new(bytes, 2);
    let proof = TailAppendSourceProof::from_document_for_source(
        &document,
        TailSelector::Body,
        source.clone().as_ref(),
    )
    .unwrap();
    let mut edit = TailAppendSourceEdit::with_limits(source, proof, TailSelector::Body, limits());
    edit.append_text_paragraphs(&["C"]).unwrap();
    let publication = TailAppendPublicationLimits::new(2, 2);
    let plan = edit.publication_plan_with_limits(publication).unwrap();
    let expected_output = plan.output_bytes() as u64;
    let expected_work = expected_output * 3;
    let memory = (plan
        .publication_limits()
        .max_window_bytes
        .min(plan.source_bytes())
        + plan.inserted_bytes()) as u64;

    let (budget, _cancel, execution) =
        context(memory, expected_work, expected_output, expected_work);
    let mut output = Vec::new();
    plan.write_to(&mut output, &execution).unwrap();
    assert_eq!(budget.used(Resource::Memory), 0);
    assert_eq!(budget.used(Resource::InputBytes), expected_work);
    assert_eq!(budget.used(Resource::Work), expected_work);

    let (_budget, _cancel, execution) = context(
        memory.saturating_sub(1),
        expected_work,
        expected_output,
        expected_work,
    );
    let mut output = Vec::new();
    let error = plan.write_to(&mut output, &execution).unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::Execution { written: 0, .. }
    ));
    assert!(output.is_empty());

    let (_budget, _cancel, execution) = context(
        memory,
        expected_work.saturating_sub(1),
        expected_output,
        expected_work,
    );
    let mut output = Vec::new();
    let error = plan.write_to(&mut output, &execution).unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::Execution { written: 0, .. }
    ));
    assert!(output.is_empty());

    let (_budget, _cancel, execution) = context(
        memory,
        expected_work,
        expected_output,
        expected_work.saturating_sub(1),
    );
    let mut output = Vec::new();
    let error = plan.write_to(&mut output, &execution).unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::Execution { written: 0, .. }
    ));
    assert!(output.is_empty());
}

struct BlockingSink {
    output: Vec<u8>,
    started: mpsc::Sender<()>,
    release: mpsc::Receiver<()>,
    blocked: bool,
}

impl Write for BlockingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if !self.blocked {
            self.blocked = true;
            self.started
                .send(())
                .map_err(|_| io::Error::other("publication test coordinator dropped"))?;
            self.release
                .recv()
                .map_err(|_| io::Error::other("publication test coordinator dropped"))?;
        }
        self.output.extend_from_slice(bytes);
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn concurrent_source_publications_reserve_independent_windows() {
    let (bytes, document) = fixture();
    let source = TestSource::new(bytes, 2);
    let proof = TailAppendSourceProof::from_document_for_source(
        &document,
        TailSelector::Body,
        source.clone().as_ref(),
    )
    .unwrap();
    let mut edit = TailAppendSourceEdit::with_limits(source, proof, TailSelector::Body, limits());
    edit.append_text_paragraphs(&["C"]).unwrap();
    let publication = TailAppendPublicationLimits::new(2, 2);
    let plan = Arc::new(edit.publication_plan_with_limits(publication).unwrap());
    let expected_output = plan.output_bytes() as u64;
    let expected_work = expected_output * 3;
    let memory = (plan
        .publication_limits()
        .max_window_bytes
        .min(plan.source_bytes())
        + plan.inserted_bytes()) as u64;
    let budget = Budget::root(
        "rtf-source-concurrent-test",
        CoreLimits::new(
            memory,
            expected_work * 2,
            expected_output * 2,
            1_000_000,
            64,
            expected_work * 2,
        ),
    );
    let (cancel_source, token) = CancellationSource::pair();
    let execution = ExecutionContext::new(
        budget.clone(),
        token,
        ExecutionLimits::new(
            NonZeroUsize::new(1).unwrap(),
            NonZeroUsize::new(1).unwrap(),
            NonZeroU64::new(64 * 1024).unwrap(),
            0,
        )
        .unwrap(),
    );
    let (started_tx, started_rx) = mpsc::channel();
    let (release_tx, release_rx) = mpsc::channel();
    let first_plan = plan.clone();
    let first_context = execution.clone();
    let first = thread::spawn(move || {
        let mut sink = BlockingSink {
            output: Vec::new(),
            started: started_tx,
            release: release_rx,
            blocked: false,
        };
        first_plan.write_to(&mut sink, &first_context)
    });
    started_rx.recv().unwrap();

    let mut second_output = Vec::new();
    let second_error = plan.write_to(&mut second_output, &execution).unwrap_err();
    assert!(matches!(
        second_error,
        TailAppendPublicationError::Execution { written: 0, .. }
    ));
    assert!(second_output.is_empty());
    release_tx.send(()).unwrap();
    assert!(first.join().unwrap().is_ok());
    cancel_source.cancel();
}

#[test]
fn source_mutation_before_planning_is_refused() {
    let (_bytes, document) = fixture();
    let source = TestSource::new(document.to_bytes().unwrap(), 4);
    let proof = TailAppendSourceProof::from_document_for_source(
        &document,
        TailSelector::Body,
        source.as_ref(),
    )
    .unwrap();
    source.mutate();
    let mut edit = TailAppendSourceEdit::with_limits(source, proof, TailSelector::Body, limits());
    edit.append_text_paragraphs(&["C"]).unwrap();
    let error = edit
        .publication_plan_with_limits(publication_limits())
        .unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::SourceVersionChanged { .. }
            | TailAppendPublicationError::SourceLengthChanged { .. }
            | TailAppendPublicationError::SourceDigestChanged { .. }
    ));
}

struct MutatingSink {
    output: Vec<u8>,
    source: Arc<TestSource>,
    mutated: bool,
}

impl Write for MutatingSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.output.extend_from_slice(bytes);
        if !self.mutated {
            self.mutated = true;
            self.source.mutate();
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_mutation_during_publication_is_incomplete_output() {
    let (_bytes, document) = fixture();
    let source = TestSource::new(document.to_bytes().unwrap(), 2);
    let plan = make_plan(source.clone(), &document, "C");
    let (_budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    let mut sink = MutatingSink {
        output: Vec::new(),
        source,
        mutated: false,
    };
    let error = plan.write_to(&mut sink, &execution).unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::IncompleteOutput {
            progress: TailAppendOutputProgress::CompleteUnverified { .. },
            ..
        }
    ));
    assert!(!sink.output.is_empty());
}

#[test]
fn source_read_failures_are_typed_and_partial_sink_progress_is_preserved() {
    let (_bytes, document) = fixture();
    let source = TestSource::new(document.to_bytes().unwrap(), 2);
    let plan = make_plan(source.clone(), &document, "C");
    source.fail_reads_after(0);
    let (_budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    let mut output = Vec::new();
    let error = plan.write_to(&mut output, &execution).unwrap_err();
    assert!(std::error::Error::source(&error).is_some());
    match error {
        TailAppendPublicationError::Source { error, written } => {
            assert_eq!(written, 0);
            assert_eq!(error.kind(), io::ErrorKind::BrokenPipe);
            assert_eq!(error.to_string(), "test source failure");
        },
        other => panic!("unexpected source error: {other:?}"),
    }

    let source = TestSource::new(document.to_bytes().unwrap(), 2);
    let plan = make_plan(source, &document, "C");
    let mut sink = LimitedSink {
        output: Vec::new(),
        remaining: 4,
    };
    let error = plan.write_to(&mut sink, &execution).unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::IncompleteOutput {
            progress: TailAppendOutputProgress::Prefix { .. },
            ..
        }
    ));
}

struct LimitedSink {
    output: Vec<u8>,
    remaining: usize,
}

impl Write for LimitedSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        if self.remaining == 0 {
            return Err(io::Error::new(io::ErrorKind::BrokenPipe, "limited sink"));
        }
        let count = self.remaining.min(bytes.len());
        self.output.extend_from_slice(&bytes[..count]);
        self.remaining -= count;
        Ok(count)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

struct ZeroAfterProgressSink {
    output: Vec<u8>,
    source: Arc<TestSource>,
    armed: bool,
}

impl Write for ZeroAfterProgressSink {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        self.output.extend_from_slice(bytes);
        if !self.armed {
            self.armed = true;
            self.source.zero_reads_from_current();
        }
        Ok(bytes.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn source_failure_after_progress_preserves_zero_read_error_and_prefix() {
    let (_bytes, document) = fixture();
    let source = TestSource::new(document.to_bytes().unwrap(), 2);
    let plan = make_plan(source.clone(), &document, "C");
    let (_budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    let mut sink = ZeroAfterProgressSink {
        output: Vec::new(),
        source,
        armed: false,
    };
    let error = plan.write_to(&mut sink, &execution).unwrap_err();
    match error {
        TailAppendPublicationError::IncompleteOutput {
            progress: TailAppendOutputProgress::Prefix { accepted, .. },
            source,
        } => {
            assert!(accepted > 0);
            match *source {
                TailAppendPublicationError::Source { error, written } => {
                    assert_eq!(error.kind(), io::ErrorKind::UnexpectedEof);
                    assert!(written > 0);
                },
                other => panic!("unexpected incomplete source cause: {other:?}"),
            }
        },
        other => panic!("unexpected source progress: {other:?}"),
    }
    assert!(!sink.output.is_empty());
}

#[test]
fn short_and_overreported_sources_refuse_proof() {
    let (bytes, document) = fixture();
    let short = TestSource::new(bytes.clone(), 1);
    short.fail_reads_after(1);
    let short_error = TailAppendSourceProof::from_document_for_source(
        &document,
        TailSelector::Body,
        short.as_ref(),
    )
    .unwrap_err();
    assert!(matches!(
        short_error,
        TailAppendPublicationError::Source { .. }
    ));

    let overreported = TestSource::new(bytes, 2);
    overreported.overreport_reads();
    let error = TailAppendSourceProof::from_document_for_source(
        &document,
        TailSelector::Body,
        overreported.as_ref(),
    )
    .unwrap_err();
    assert!(matches!(error, TailAppendPublicationError::Source { .. }));
}

#[test]
fn foreign_revision_source_and_execution_limits_are_rejected_before_output() {
    let (bytes, document) = fixture();
    let source = TestSource::new(bytes.clone(), 2);
    let plan = make_plan(source, &document, "C");
    let foreign = TestSource::new(bytes, 2);
    let (_budget, _cancel, execution) = context(4096, 4096, 4096, 4096);
    let mut output = Vec::new();
    let error = plan
        .write_to_source(foreign.as_ref(), &mut output, &execution)
        .unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::SourceVersionChanged { .. }
    ));
    assert!(output.is_empty());

    let (_budget, cancel, execution) = context(4096, 4096, 4096, 4096);
    cancel.cancel();
    let mut output = Vec::new();
    let error = plan.write_to(&mut output, &execution).unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::Execution { written: 0, .. }
    ));

    let (_budget, _cancel, execution) = context(4096, 4096, 0, 4096);
    let mut output = Vec::new();
    let error = plan.write_to(&mut output, &execution).unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::Execution { written: 0, .. }
    ));

    let (_budget, _cancel, execution) = context(0, 4096, 4096, 4096);
    let mut output = Vec::new();
    let error = plan.write_to(&mut output, &execution).unwrap_err();
    assert!(matches!(
        error,
        TailAppendPublicationError::Execution { written: 0, .. }
    ));
}

#[test]
fn durable_source_plan_round_trips_and_forged_source_never_grants_proof() {
    let (_bytes, document) = fixture();
    let source = TestSource::new(document.to_bytes().unwrap(), 3);
    let plan = make_plan(source, &document, "C");
    let durable = plan.to_durable().unwrap();
    let json = durable.to_deterministic_json().unwrap();
    let parsed =
        litchi_rtf::tail_append::DurableTailAppendPatch::from_deterministic_json(&json, limits())
            .unwrap();
    let target = parsed.apply(&document).unwrap();
    assert_eq!(target.text(), "A\nB\nC\n");
    let restored = parsed.inverse().apply(&target).unwrap();
    assert_eq!(restored.text(), document.text());

    let foreign = Document::from_bytes(br#"{\rtf1\ansi Foreign}"#).unwrap();
    assert!(matches!(
        parsed.apply(&foreign),
        Err(TailAppendError::PatchConflict)
    ));
}

#[test]
fn source_proof_is_private_capability_and_plain_staging_is_atomic() {
    let (_bytes, document) = fixture();
    let source = TestSource::new(document.to_bytes().unwrap(), 2);
    let proof = TailAppendSourceProof::from_document_for_source(
        &document,
        TailSelector::Body,
        source.as_ref(),
    )
    .unwrap();
    let mut edit = TailAppendSourceEdit::with_limits(source, proof, TailSelector::Body, limits());
    let invalid_runs = [PlainRun::new("bad\ntext")];
    let invalid = [PlainParagraph::new(&invalid_runs)];
    assert!(matches!(
        edit.append_paragraphs(&invalid),
        Err(TailAppendError::InvalidText(_))
    ));
    assert_eq!(edit.paragraph_count(), 0);
    edit.append_paragraphs(&[PlainParagraph::new(&[PlainRun::new("C")])])
        .unwrap();
    assert_eq!(edit.paragraph_count(), 1);
}
