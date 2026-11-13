//! Which failure the plain-paragraph ODT stream reports when two causes are
//! recorded for the same stream.
//!
//! A caller sink that answers `ErrorKind::Interrupted` once is retried, but the
//! sink adapter remembers the failed write for the rest of the stream. Each
//! test pairs that remembered sink failure with a second cause and pins the
//! error the caller receives.

use std::{
    io::{self, Write},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, Error, ExecutionContext, ExecutionError, ExecutionLimits, Limits,
    Resource,
};
use litchi_odt::streaming::{
    PublicationFailureKind, StreamingError, StreamingLimits, stream_plain_paragraphs_to,
    try_stream_plain_paragraphs_to,
};

fn context_with(limits: Limits) -> (Budget, CancellationSource, ExecutionContext) {
    let budget = Budget::root("odt-streaming-failure-priority", limits);
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
    max_paragraphs: usize,
    max_paragraph_text_bytes: usize,
    max_paragraph_xml_bytes: usize,
    max_output_bytes: u64,
) -> StreamingLimits {
    let base = StreamingLimits::default();
    StreamingLimits::new(
        max_paragraphs,
        max_paragraph_text_bytes,
        base.max_total_text_bytes(),
        max_paragraph_xml_bytes,
        base.max_content_xml_bytes(),
        max_output_bytes,
        base.xml_audit(),
    )
    .expect("checked ODT streaming limits")
}

fn base_limits() -> (usize, usize, usize, u64) {
    let base = StreamingLimits::default();
    (
        base.max_paragraphs(),
        base.max_paragraph_text_bytes(),
        base.max_paragraph_xml_bytes(),
        base.max_output_bytes(),
    )
}

/// Fails its first non-empty write with `Interrupted`, then accepts bytes.
/// Optionally cancels the stream during that first write.
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

fn stream(
    sink: &mut InterruptedOnceSink,
    paragraphs: &[&'static str],
    budget: Limits,
    limits: StreamingLimits,
) -> StreamingError {
    let (_budget, _cancellation, context) = context_with(budget);
    let error = stream_plain_paragraphs_to(sink, paragraphs.to_vec(), &context, limits)
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

#[test]
fn paragraph_count_limit_outranks_an_earlier_sink_failure() {
    let (_, text, xml, output) = base_limits();
    let mut sink = InterruptedOnceSink::new();
    let error = stream(
        &mut sink,
        &["one", "two"],
        unlimited(),
        limits_with(1, text, xml, output),
    );
    assert_limit(error, "paragraphs", &sink);
}

#[test]
fn paragraph_text_limit_outranks_an_earlier_sink_failure() {
    let (paragraphs, _, xml, output) = base_limits();
    let mut sink = InterruptedOnceSink::new();
    let error = stream(
        &mut sink,
        &["one", "longer than four"],
        unlimited(),
        limits_with(paragraphs, 4, xml, output),
    );
    assert_limit(error, "paragraph text bytes", &sink);
}

#[test]
fn paragraph_xml_limit_outranks_an_earlier_sink_failure() {
    let (paragraphs, text, _, output) = base_limits();
    let mut sink = InterruptedOnceSink::new();
    let error = stream(
        &mut sink,
        &["one"],
        unlimited(),
        limits_with(paragraphs, text, 4, output),
    );
    assert_limit(error, "paragraph XML bytes", &sink);
}

#[test]
fn output_byte_limit_outranks_an_earlier_sink_failure() {
    let (paragraphs, text, xml, _) = base_limits();
    let (_budget, _cancellation, context) = context_with(unlimited());
    let mut complete = Vec::new();
    stream_plain_paragraphs_to(
        &mut complete,
        vec!["one"],
        &context,
        StreamingLimits::default(),
    )
    .expect("control stream");

    let mut sink = InterruptedOnceSink::new();
    let error = stream(
        &mut sink,
        &["one"],
        unlimited(),
        limits_with(paragraphs, text, xml, complete.len() as u64 - 1),
    );
    assert!(!sink.bytes.is_empty());
    assert_limit(error, "output bytes", &sink);
}

#[test]
fn producer_budget_refusal_outranks_an_earlier_sink_failure() {
    let mut sink = InterruptedOnceSink::new();
    let budget = Limits::new(u64::MAX, u64::MAX, u64::MAX, 1, u64::MAX, u64::MAX);
    let error = stream(
        &mut sink,
        &["one", "two"],
        budget,
        StreamingLimits::default(),
    );
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
fn cancellation_from_the_paragraph_source_outranks_an_earlier_sink_failure() {
    let (_budget, cancellation, context) = context_with(unlimited());
    let mut sink = InterruptedOnceSink::new();
    let paragraphs = (0..3).map(move |index| {
        if index == 1 {
            cancellation.cancel();
        }
        "text"
    });
    let error =
        stream_plain_paragraphs_to(&mut sink, paragraphs, &context, StreamingLimits::default())
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
    let error =
        stream_plain_paragraphs_to(&mut sink, vec!["one"], &context, StreamingLimits::default())
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
fn invalid_text_outranks_an_earlier_sink_failure() {
    let mut sink = InterruptedOnceSink::new();
    let error = stream(
        &mut sink,
        &["one", "\u{0}"],
        unlimited(),
        StreamingLimits::default(),
    );
    match error {
        StreamingError::Invalid { written, .. } => {
            assert_eq!(written, sink.bytes.len() as u64);
        },
        other => panic!("expected an invalid-text refusal, got {other:?}"),
    }
}

#[test]
fn paragraph_source_error_outranks_an_earlier_sink_failure() {
    let (_budget, _cancellation, context) = context_with(unlimited());
    let mut sink = InterruptedOnceSink::new();
    let paragraphs: Vec<litchi_core::Result<&'static str>> = vec![
        Ok("one"),
        Err(Error::Other("planned source failure".to_string())),
    ];
    let error =
        try_stream_plain_paragraphs_to(&mut sink, paragraphs, &context, StreamingLimits::default())
            .expect_err("the source error stops the stream");
    assert!(sink.interrupted, "the sink failure was recorded first");
    match error {
        StreamingError::Producer { written, source } => {
            assert_eq!(written, sink.bytes.len() as u64);
            assert!(source.to_string().contains("planned source failure"));
        },
        other => panic!("expected a producer failure, got {other:?}"),
    }
}

#[test]
fn a_sink_that_stops_accepting_bytes_is_reported_before_its_own_cancellation() {
    let (_budget, cancellation, context) = context_with(unlimited());
    let mut sink = ZeroAndCancelSink(cancellation);
    let error =
        stream_plain_paragraphs_to(&mut sink, vec!["one"], &context, StreamingLimits::default())
            .expect_err("a sink that accepts nothing stops the stream");
    match error {
        StreamingError::Publication(error) => {
            assert_eq!(error.kind(), PublicationFailureKind::Sink);
            assert_eq!(error.written(), 0);
        },
        other => panic!("expected a sink failure, got {other:?}"),
    }
}
