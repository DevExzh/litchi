//! Which failure the scalar ODS stream reports when two causes are recorded
//! for the same stream.
//!
//! A caller sink that answers `ErrorKind::Interrupted` once is retried, but the
//! sink adapter remembers the failed write for the rest of the stream. Each
//! test pairs that remembered sink failure with a second cause and pins the
//! error the caller receives.

use std::{
    borrow::Cow,
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, Resource,
};
use litchi_ods::streaming::{
    PublicationFailureKind, StreamingCell, StreamingError, StreamingLimits, stream_scalar_rows_to,
};

type Row = Vec<StreamingCell<'static>>;

fn context_with(limits: Limits) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root("streaming-failure-priority", limits);
    let (cancellation, token) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(u64::MAX).expect("finite byte ceiling"),
        0,
    )
    .expect("valid execution policy");
    let context = ExecutionContext::new(budget.clone(), token, execution);
    (budget, cancellation, context)
}

fn unlimited() -> Limits {
    Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
}

fn limits_with(
    max_rows: usize,
    max_row_xml_bytes: usize,
    max_output_bytes: u64,
) -> StreamingLimits {
    let base = StreamingLimits::default();
    StreamingLimits::new(
        max_rows,
        base.max_cells(),
        base.max_cells_per_row(),
        base.max_text_bytes(),
        max_row_xml_bytes,
        base.max_content_xml_bytes(),
        max_output_bytes,
        base.xml_audit(),
    )
    .expect("checked streaming limits")
}

fn text_row() -> Row {
    vec![StreamingCell::Text(Cow::Borrowed("x"))]
}

/// Fails its first non-empty write with `Interrupted`, then accepts bytes.
/// Optionally cancels the stream during that first write, or stops accepting
/// bytes (`Ok(0)`) while cancelling.
struct InterruptedOnceSink {
    bytes: Vec<u8>,
    interrupted: bool,
    cancel_on_interrupt: Option<CancellationSource>,
}

impl InterruptedOnceSink {
    fn new() -> Self {
        Self {
            bytes: Vec::new(),
            interrupted: false,
            cancel_on_interrupt: None,
        }
    }

    fn cancelling(source: CancellationSource) -> Self {
        Self {
            bytes: Vec::new(),
            interrupted: false,
            cancel_on_interrupt: Some(source),
        }
    }
}

impl Write for InterruptedOnceSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        if !self.interrupted {
            self.interrupted = true;
            if let Some(source) = &self.cancel_on_interrupt {
                source.cancel();
            }
            return Err(io::Error::new(io::ErrorKind::Interrupted, "retry"));
        }
        self.bytes.extend_from_slice(input);
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

/// Accepts nothing and cancels the stream from inside the refused write.
struct ZeroAndCancelSink(CancellationSource);

impl Write for ZeroAndCancelSink {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if !input.is_empty() {
            self.0.cancel();
        }
        Ok(0)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

fn stream<I>(
    sink: &mut InterruptedOnceSink,
    rows: I,
    budget: Limits,
    limits: StreamingLimits,
) -> StreamingError
where
    I: IntoIterator<Item = Row>,
{
    let (_budget, _cancellation, context) = context_with(budget);
    let error = stream_scalar_rows_to(sink, rows, &context, limits)
        .expect_err("the second cause stops the stream");
    assert!(sink.interrupted, "the sink failure was recorded first");
    error
}

fn assert_limit(error: StreamingError, expected: &str, sink: &InterruptedOnceSink) {
    match error {
        StreamingError::LimitExceeded {
            resource, written, ..
        } => {
            assert_eq!(resource, expected);
            assert_eq!(written, sink.bytes.len() as u64);
        },
        other => panic!("expected {expected} limit, got {other:?}"),
    }
}

fn assert_publication(
    error: StreamingError,
    expected: PublicationFailureKind,
    sink: &InterruptedOnceSink,
) {
    match error {
        StreamingError::Publication(error) => {
            assert_eq!(error.kind(), expected);
            assert_eq!(error.written(), sink.bytes.len() as u64);
        },
        other => panic!("expected {expected:?} publication failure, got {other:?}"),
    }
}

#[test]
fn row_count_limit_outranks_an_earlier_sink_failure() {
    let mut sink = InterruptedOnceSink::new();
    let limits = limits_with(1, 4096, u64::MAX);
    let error = stream(&mut sink, vec![text_row(), text_row()], unlimited(), limits);
    assert_limit(error, "physical rows", &sink);
}

#[test]
fn row_xml_limit_outranks_an_earlier_sink_failure() {
    let mut sink = InterruptedOnceSink::new();
    let limits = limits_with(8, 32, u64::MAX);
    let error = stream(&mut sink, vec![text_row()], unlimited(), limits);
    assert_limit(error, "row XML bytes", &sink);
}

#[test]
fn output_byte_limit_outranks_an_earlier_sink_failure() {
    let (_budget, _cancellation, context) = context_with(unlimited());
    let mut complete = Vec::new();
    stream_scalar_rows_to(
        &mut complete,
        vec![text_row()],
        &context,
        StreamingLimits::default(),
    )
    .expect("control stream");

    let mut sink = InterruptedOnceSink::new();
    let limits = limits_with(8, 4096, complete.len() as u64 - 1);
    let error = stream(&mut sink, vec![text_row()], unlimited(), limits);
    assert!(!sink.bytes.is_empty());
    assert_limit(error, "output bytes", &sink);
}

#[test]
fn producer_budget_refusal_outranks_an_earlier_sink_failure() {
    let mut sink = InterruptedOnceSink::new();
    let budget = Limits::new(u64::MAX, u64::MAX, u64::MAX, 3, u64::MAX, u64::MAX);
    let (_budget, _cancellation, context) = context_with(budget);
    let error = stream_scalar_rows_to(
        &mut sink,
        vec![text_row(), text_row()],
        &context,
        StreamingLimits::default(),
    )
    .expect_err("the object budget stops the stream");
    assert!(sink.interrupted, "the sink failure was recorded first");
    match error {
        StreamingError::Execution {
            written,
            error: ExecutionError::ResourceLimit(limit),
        } => {
            assert_eq!(limit.resource, Resource::Objects);
            assert_eq!(written, sink.bytes.len() as u64);
        },
        other => panic!("expected an object budget refusal, got {other:?}"),
    }
}

#[test]
fn cancellation_from_the_row_source_outranks_an_earlier_sink_failure() {
    let (_budget, cancellation, context) = context_with(unlimited());
    let mut sink = InterruptedOnceSink::new();
    let rows = (0..3).map(move |index| {
        if index == 1 {
            cancellation.cancel();
        }
        text_row()
    });
    let error = stream_scalar_rows_to(&mut sink, rows, &context, StreamingLimits::default())
        .expect_err("cancellation stops the stream");
    assert!(sink.interrupted, "the sink failure was recorded first");
    match error {
        StreamingError::Execution {
            written,
            error: ExecutionError::Cancelled,
        } => assert_eq!(written, sink.bytes.len() as u64),
        other => panic!("expected cancellation, got {other:?}"),
    }
}

#[test]
fn cancellation_during_the_interrupted_write_is_reported_as_cancellation() {
    let (_budget, cancellation, context) = context_with(unlimited());
    let mut sink = InterruptedOnceSink::cancelling(cancellation);
    let error = stream_scalar_rows_to(
        &mut sink,
        vec![text_row()],
        &context,
        StreamingLimits::default(),
    )
    .expect_err("cancellation stops the stream");
    assert!(sink.interrupted);
    match error {
        StreamingError::Execution {
            written: 0,
            error: ExecutionError::Cancelled,
        } => assert!(sink.bytes.is_empty()),
        other => panic!("expected cancellation, got {other:?}"),
    }
}

#[test]
fn an_earlier_sink_failure_outranks_an_untyped_producer_refusal() {
    for row in [
        vec![StreamingCell::Number(f64::NAN)],
        vec![StreamingCell::Text(Cow::Borrowed("\u{0}"))],
    ] {
        let mut sink = InterruptedOnceSink::new();
        let error = stream(
            &mut sink,
            vec![text_row(), row],
            unlimited(),
            StreamingLimits::default(),
        );
        assert_publication(error, PublicationFailureKind::Sink, &sink);
    }
}

#[test]
fn a_sink_that_stops_accepting_bytes_is_reported_before_its_own_cancellation() {
    let (_budget, cancellation, context) = context_with(unlimited());
    let mut sink = ZeroAndCancelSink(cancellation);
    let error = stream_scalar_rows_to(
        &mut sink,
        vec![text_row()],
        &context,
        StreamingLimits::default(),
    )
    .expect_err("a sink that accepts nothing stops the stream");
    match error {
        StreamingError::Publication(error) => {
            assert_eq!(error.kind(), PublicationFailureKind::Sink);
            assert_eq!(error.written(), 0);
        },
        other => panic!("expected a sink failure, got {other:?}"),
    }
}
