//! Private differential tests for the ODT paragraph scalar emitter.
//!
//! Retains the original scalar write boundaries as an independent oracle for
//! encoded bytes, local XML refusals, and hierarchical Work accounting.

#![allow(clippy::expect_used, clippy::unwrap_used)]

use super::*;

use litchi_core::{Budget, CancellationSource, ExecutionLimits, Limits};
use std::num::{NonZeroU64, NonZeroUsize};

const REFERENCE_MAX_SPACE_COUNT: usize = 1_000_000;

/// Independent snapshot of the pre-batching scalar write sequence. Each
/// returned element is one call that the old emitter made to
/// `ParagraphWriter::write_bytes`; keeping these boundaries lets threshold
/// tests distinguish a correct batched fast path from a changed limit or Work
/// failure point.
fn scalar_chunks(value: &str) -> Vec<Vec<u8>> {
    if value.is_empty() {
        return vec![b"<text:p/>".to_vec()];
    }

    let mut chunks = vec![b"<text:p>".to_vec()];
    let characters: Vec<char> = value.chars().collect();
    let mut index = 0;
    while index < characters.len() {
        match characters[index] {
            ' ' => {
                let start = index;
                while index < characters.len() && characters[index] == ' ' {
                    index += 1;
                }
                let count = index - start;
                let adjacent_control = (start > 0 && matches!(characters[start - 1], '\t' | '\n'))
                    || (index < characters.len() && matches!(characters[index], '\t' | '\n'));
                let needs_control =
                    count > 1 || start == 0 || index == characters.len() || adjacent_control;
                if !needs_control {
                    chunks.push(b" ".to_vec());
                    continue;
                }

                let mut remaining = count;
                while remaining != 0 {
                    let chunk = remaining.min(REFERENCE_MAX_SPACE_COUNT);
                    chunks.push(b"<text:s text:c=\"".to_vec());
                    chunks.push(chunk.to_string().into_bytes());
                    chunks.push(b"\"/>".to_vec());
                    remaining -= chunk;
                }
            },
            '\t' => {
                chunks.push(b"<text:tab/>".to_vec());
                index += 1;
            },
            '\n' => {
                chunks.push(b"<text:line-break/>".to_vec());
                index += 1;
            },
            '&' => {
                chunks.push(b"&amp;".to_vec());
                index += 1;
            },
            '<' => {
                chunks.push(b"&lt;".to_vec());
                index += 1;
            },
            '>' => {
                chunks.push(b"&gt;".to_vec());
                index += 1;
            },
            '"' => {
                chunks.push(b"&quot;".to_vec());
                index += 1;
            },
            '\'' => {
                chunks.push(b"&apos;".to_vec());
                index += 1;
            },
            character => {
                let mut encoded = [0_u8; 4];
                chunks.push(character.encode_utf8(&mut encoded).as_bytes().to_vec());
                index += 1;
            },
        }
    }
    chunks.push(b"</text:p>".to_vec());
    chunks
}

fn scalar_reference(value: &str) -> Vec<u8> {
    scalar_chunks(value).into_iter().flatten().collect()
}

fn flatten_chunks(chunks: &[Vec<u8>]) -> Vec<u8> {
    chunks.iter().flatten().copied().collect()
}

fn scalar_threshold_cases() -> [(&'static str, String); 3] {
    [
        ("ascii-257", "a".repeat(257)),
        ("multibyte-257", "é".repeat(128) + "a"),
        ("mixed-257", "a".repeat(246) + "é🙂& \t\nx"),
    ]
}

fn scalar_cases() -> Vec<String> {
    vec![
        String::new(),
        "ordinary & < > \" '".to_owned(),
        "  leading  middle   trailing  ".to_owned(),
        "a b".to_owned(),
        "a  b".to_owned(),
        "\t\n".to_owned(),
        "a \t \n b".to_owned(),
        "café🙂漢字".to_owned(),
        "a".repeat(255),
        "a".repeat(256),
        "a".repeat(257),
        "a".repeat(254) + "é",
        "a".repeat(255) + "é",
        "é🙂漢".repeat(87),
    ]
}

fn all_limits() -> Limits {
    Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
}

fn context_with_parent_child(
    parent_limits: Limits,
    child_limits: Limits,
) -> (Budget, Budget, CancellationSource, ExecutionContext) {
    let parent = Budget::root("odt-span-parent", parent_limits);
    let child = parent.child("odt-span-child", child_limits);
    let (cancellation, token) = CancellationSource::pair();
    let execution = ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(u64::MAX).expect("finite byte ceiling"),
        0,
    )
    .expect("valid execution policy");
    let context = ExecutionContext::new(child.clone(), token, execution);
    (parent, child, cancellation, context)
}

#[derive(Debug)]
struct EmitObservation {
    result: litchi_core::Result<()>,
    bytes: usize,
    local_limit: Option<(&'static str, usize, usize)>,
    execution_error: Option<ExecutionError>,
}

fn emit_with_context(
    output: &mut dyn Write,
    value: &str,
    paragraph_limit: usize,
    content_before: usize,
    content_limit: usize,
    context: &ExecutionContext,
) -> EmitObservation {
    let mut writer = ParagraphWriter::new(
        output,
        context,
        paragraph_limit,
        content_before,
        content_limit,
    );
    let result = emit_paragraph(&mut writer, value);
    EmitObservation {
        result,
        bytes: writer.bytes,
        local_limit: writer
            .limit_error
            .map(|error| (error.resource, error.observed, error.limit)),
        execution_error: writer.execution_error,
    }
}

fn emit_reference_chunks(
    output: &mut dyn Write,
    chunks: &[Vec<u8>],
    paragraph_limit: usize,
    content_before: usize,
    content_limit: usize,
    context: &ExecutionContext,
) -> EmitObservation {
    let mut writer = ParagraphWriter::new(
        output,
        context,
        paragraph_limit,
        content_before,
        content_limit,
    );
    let mut result: litchi_core::Result<()> = Ok(());
    for chunk in chunks {
        if let Err(error) = writer.write_bytes(chunk) {
            result = Err(error);
            break;
        }
    }
    EmitObservation {
        result,
        bytes: writer.bytes,
        local_limit: writer
            .limit_error
            .map(|error| (error.resource, error.observed, error.limit)),
        execution_error: writer.execution_error,
    }
}

fn limits_with_work(work: u64) -> Limits {
    Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work)
}

fn run_production(
    value: &str,
    paragraph_limit: usize,
    content_before: usize,
    content_limit: usize,
    parent_limits: Limits,
    child_limits: Limits,
) -> (Budget, Budget, Vec<u8>, EmitObservation) {
    let (parent, child, _cancellation, context) =
        context_with_parent_child(parent_limits, child_limits);
    let mut output = Vec::new();
    let observation = emit_with_context(
        &mut output,
        value,
        paragraph_limit,
        content_before,
        content_limit,
        &context,
    );
    (parent, child, output, observation)
}

fn run_reference(
    chunks: &[Vec<u8>],
    paragraph_limit: usize,
    content_before: usize,
    content_limit: usize,
    parent_limits: Limits,
    child_limits: Limits,
) -> (Budget, Budget, Vec<u8>, EmitObservation) {
    let (parent, child, _cancellation, context) =
        context_with_parent_child(parent_limits, child_limits);
    let mut output = Vec::new();
    let observation = emit_reference_chunks(
        &mut output,
        chunks,
        paragraph_limit,
        content_before,
        content_limit,
        &context,
    );
    (parent, child, output, observation)
}

fn assert_execution_equivalent(actual: &Option<ExecutionError>, expected: &Option<ExecutionError>) {
    match (actual.as_ref(), expected.as_ref()) {
        (None, None) => {},
        (
            Some(ExecutionError::ResourceLimit(actual)),
            Some(ExecutionError::ResourceLimit(expected)),
        ) => {
            assert_eq!(actual.resource, expected.resource);
            assert_eq!(actual.observed, expected.observed);
            assert_eq!(actual.limit, expected.limit);
            assert_eq!(actual.scope, expected.scope);
        },
        (Some(ExecutionError::Cancelled), Some(ExecutionError::Cancelled)) => {},
        (actual, expected) => {
            panic!("execution failure differs: actual={actual:?}, expected={expected:?}")
        },
    }
}

fn assert_observations_equivalent(
    label: &str,
    actual_output: &[u8],
    actual: &EmitObservation,
    expected_output: &[u8],
    expected: &EmitObservation,
) {
    assert_eq!(actual_output, expected_output, "{label}: output differs");
    assert_eq!(
        actual.bytes, expected.bytes,
        "{label}: writer byte count differs"
    );
    assert_eq!(
        actual.local_limit, expected.local_limit,
        "{label}: local limit differs"
    );
    assert_eq!(
        actual.result.is_ok(),
        expected.result.is_ok(),
        "{label}: result status differs"
    );
    assert_execution_equivalent(&actual.execution_error, &expected.execution_error);
}

#[allow(clippy::too_many_arguments)]
fn compare_threshold_run(
    label: &str,
    value: &str,
    chunks: &[Vec<u8>],
    paragraph_limit: usize,
    content_before: usize,
    content_limit: usize,
    parent_limits: Limits,
    child_limits: Limits,
    expected_local_resource: Option<&'static str>,
) {
    let (actual_parent, actual_child, actual_output, actual) = run_production(
        value,
        paragraph_limit,
        content_before,
        content_limit,
        parent_limits,
        child_limits,
    );
    let (expected_parent, expected_child, expected_output, expected) = run_reference(
        chunks,
        paragraph_limit,
        content_before,
        content_limit,
        parent_limits,
        child_limits,
    );
    assert_observations_equivalent(label, &actual_output, &actual, &expected_output, &expected);
    let scalar_output = flatten_chunks(chunks);
    assert!(
        actual_output.len() <= scalar_output.len(),
        "{label}: output grew"
    );
    assert_eq!(
        actual_output,
        &scalar_output[..actual_output.len()],
        "{label}: accepted bytes are not a scalar prefix"
    );
    assert_eq!(
        actual_parent.used(Resource::Work),
        expected_parent.used(Resource::Work),
        "{label}: parent Work differs"
    );
    assert_eq!(
        actual_child.used(Resource::Work),
        expected_child.used(Resource::Work),
        "{label}: child Work differs"
    );
    assert_eq!(
        actual_parent.used(Resource::Work),
        actual_output.len() as u64,
        "{label}: parent Work must equal accepted bytes"
    );
    assert_eq!(
        actual_child.used(Resource::Work),
        actual_output.len() as u64,
        "{label}: child Work must equal accepted bytes"
    );
    if let Some(resource) = expected_local_resource {
        assert_eq!(
            actual.local_limit.map(|limit| limit.0),
            Some(resource),
            "{label}: expected local limit resource"
        );
    }
}

fn assert_limit(
    observation: &EmitObservation,
    expected_resource: &'static str,
    output: &[u8],
    expected_limit: usize,
) {
    let (resource, observed, limit) = observation
        .local_limit
        .expect("local XML ceiling must be recorded");
    assert_eq!(resource, expected_resource);
    assert!(observed > limit);
    assert_eq!(limit, expected_limit);
    assert_eq!(observation.bytes, output.len());
    assert!(observation.result.is_err());
}

fn assert_work_refusal(observation: &EmitObservation, output: &[u8], expected_limit: u64) {
    match observation.execution_error.as_ref() {
        Some(ExecutionError::ResourceLimit(limit)) => {
            assert_eq!(limit.resource, Resource::Work);
            assert!(limit.observed > limit.limit);
            assert_eq!(limit.limit, expected_limit);
        },
        other => panic!("expected Work refusal, got {other:?}"),
    }
    assert_eq!(observation.bytes, output.len());
    assert!(observation.result.is_err());
}

#[test]
fn scalar_reference_matches_paragraph_emitter_at_controls_unicode_and_span_boundaries() {
    for value in scalar_cases() {
        let expected = scalar_reference(&value);
        let (parent, child, _cancellation, context) =
            context_with_parent_child(all_limits(), all_limits());
        let mut output = Vec::new();
        let observation =
            emit_with_context(&mut output, &value, usize::MAX, 0, usize::MAX, &context);
        observation.result.expect("scalar span emits successfully");
        assert_eq!(output, expected, "emitter differs for {value:?}");
        assert_eq!(observation.bytes, expected.len());
        assert_eq!(parent.used(Resource::Work), expected.len() as u64);
        assert_eq!(child.used(Resource::Work), expected.len() as u64);
    }
}

#[test]
fn every_span_threshold_matches_scalar_chunks_for_local_work_and_ties() {
    for (name, value) in scalar_threshold_cases() {
        let chunks = scalar_chunks(&value);
        let encoded = flatten_chunks(&chunks);
        assert_eq!(value.len(), 257, "{name} raw representative length");

        for threshold in 0..=encoded.len() {
            let local_resource = if threshold < encoded.len() {
                Some("paragraph XML bytes")
            } else {
                None
            };
            compare_threshold_run(
                &format!("{name}/paragraph/{threshold}"),
                &value,
                &chunks,
                threshold,
                0,
                usize::MAX,
                all_limits(),
                all_limits(),
                local_resource,
            );

            let local_resource = if threshold < encoded.len() {
                Some("content XML bytes")
            } else {
                None
            };
            compare_threshold_run(
                &format!("{name}/content/{threshold}"),
                &value,
                &chunks,
                usize::MAX,
                7,
                7 + threshold,
                all_limits(),
                all_limits(),
                local_resource,
            );

            let work = threshold as u64;
            compare_threshold_run(
                &format!("{name}/work-child/{threshold}"),
                &value,
                &chunks,
                usize::MAX,
                0,
                usize::MAX,
                all_limits(),
                limits_with_work(work),
                None,
            );
            compare_threshold_run(
                &format!("{name}/work-parent/{threshold}"),
                &value,
                &chunks,
                usize::MAX,
                0,
                usize::MAX,
                limits_with_work(work),
                all_limits(),
                None,
            );

            // Both checks would reject the next scalar chunk at the same
            // threshold. ParagraphWriter checks aggregate content before
            // charging Work, so content is the expected first failure.
            let local_resource = if threshold < encoded.len() {
                Some("content XML bytes")
            } else {
                None
            };
            compare_threshold_run(
                &format!("{name}/paragraph-content-tie/{threshold}"),
                &value,
                &chunks,
                threshold,
                7,
                7 + threshold,
                all_limits(),
                all_limits(),
                local_resource,
            );
            compare_threshold_run(
                &format!("{name}/content-work-child-tie/{threshold}"),
                &value,
                &chunks,
                usize::MAX,
                7,
                7 + threshold,
                all_limits(),
                limits_with_work(work),
                local_resource,
            );
            compare_threshold_run(
                &format!("{name}/content-work-parent-tie/{threshold}"),
                &value,
                &chunks,
                usize::MAX,
                7,
                7 + threshold,
                limits_with_work(work),
                all_limits(),
                local_resource,
            );
        }
    }
}

#[test]
fn paragraph_and_content_xml_ceilings_report_first_failure_and_prefix() {
    for value in scalar_cases() {
        let expected = scalar_reference(&value);
        let content_before = 7;
        let content_exact = content_before + expected.len();

        let (parent, child, _cancellation, context) =
            context_with_parent_child(all_limits(), all_limits());
        let mut exact_output = Vec::new();
        let exact = emit_with_context(
            &mut exact_output,
            &value,
            expected.len(),
            content_before,
            content_exact,
            &context,
        );
        exact.result.expect("exact XML ceilings succeed");
        assert_eq!(exact_output, expected);
        assert_eq!(parent.used(Resource::Work), expected.len() as u64);
        assert_eq!(child.used(Resource::Work), expected.len() as u64);

        let (parent, child, _cancellation, context) =
            context_with_parent_child(all_limits(), all_limits());
        let mut paragraph_output = Vec::new();
        let paragraph_under = emit_with_context(
            &mut paragraph_output,
            &value,
            expected.len() - 1,
            content_before,
            usize::MAX,
            &context,
        );
        assert_limit(
            &paragraph_under,
            "paragraph XML bytes",
            &paragraph_output,
            expected.len() - 1,
        );
        assert_eq!(
            paragraph_output,
            expected[..paragraph_output.len()],
            "paragraph refusal must preserve an accepted prefix"
        );
        assert_eq!(parent.used(Resource::Work), paragraph_output.len() as u64);
        assert_eq!(child.used(Resource::Work), paragraph_output.len() as u64);

        let (parent, child, _cancellation, context) =
            context_with_parent_child(all_limits(), all_limits());
        let mut content_output = Vec::new();
        let content_under = emit_with_context(
            &mut content_output,
            &value,
            expected.len(),
            content_before,
            content_exact - 1,
            &context,
        );
        assert_limit(
            &content_under,
            "content XML bytes",
            &content_output,
            content_exact - 1,
        );
        assert_eq!(
            content_output,
            expected[..content_output.len()],
            "content refusal must preserve an accepted prefix"
        );
        assert_eq!(parent.used(Resource::Work), content_output.len() as u64);
        assert_eq!(child.used(Resource::Work), content_output.len() as u64);

        let (_parent, _child, _cancellation, context) =
            context_with_parent_child(all_limits(), all_limits());
        let mut combined_output = Vec::new();
        let combined = emit_with_context(
            &mut combined_output,
            &value,
            expected.len() - 1,
            content_before,
            content_exact - 1,
            &context,
        );
        assert_limit(
            &combined,
            "content XML bytes",
            &combined_output,
            content_exact - 1,
        );
        assert_eq!(
            combined_output,
            expected[..combined_output.len()],
            "combined ceilings must preserve the first checked failure prefix"
        );
    }
}

#[test]
fn work_budget_exact_and_one_under_preserve_hierarchical_prefix_and_rollback() {
    for value in scalar_cases() {
        let expected = scalar_reference(&value);
        let work = expected.len() as u64;

        let (parent, child, _cancellation, context) = context_with_parent_child(
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, work),
        );
        let mut exact_output = Vec::new();
        let exact = emit_with_context(
            &mut exact_output,
            &value,
            usize::MAX,
            0,
            usize::MAX,
            &context,
        );
        exact.result.expect("exact Work budget succeeds");
        assert_eq!(exact_output, expected);
        assert_eq!(parent.used(Resource::Work), work);
        assert_eq!(child.used(Resource::Work), work);

        let under = work - 1;
        let (parent, child, _cancellation, context) = context_with_parent_child(
            all_limits(),
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, under),
        );
        let mut child_output = Vec::new();
        let child_observation = emit_with_context(
            &mut child_output,
            &value,
            usize::MAX,
            0,
            usize::MAX,
            &context,
        );
        assert_work_refusal(&child_observation, &child_output, under);
        assert_eq!(child_output, expected[..child_output.len()]);
        assert!(child.used(Resource::Work) < work);
        assert_eq!(parent.used(Resource::Work), child.used(Resource::Work));

        let (parent, child, _cancellation, context) = context_with_parent_child(
            Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, under),
            all_limits(),
        );
        let mut parent_output = Vec::new();
        let parent_observation = emit_with_context(
            &mut parent_output,
            &value,
            usize::MAX,
            0,
            usize::MAX,
            &context,
        );
        assert_work_refusal(&parent_observation, &parent_output, under);
        assert_eq!(parent_output, expected[..parent_output.len()]);
        assert!(parent.used(Resource::Work) < work);
        assert_eq!(parent.used(Resource::Work), child.used(Resource::Work));
    }
}

#[derive(Debug)]
struct CancelAfterShortWrite {
    bytes: Vec<u8>,
    max_write: usize,
    cancellation: CancellationSource,
    cancelled: bool,
}

#[derive(Debug)]
struct CancelAtAcceptedBytes {
    bytes: Vec<u8>,
    threshold: usize,
    max_write: usize,
    cancellation: CancellationSource,
    cancelled: bool,
}

impl Write for CancelAtAcceptedBytes {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let remaining = self.threshold.saturating_sub(self.bytes.len()).max(1);
        let amount = bytes.len().min(self.max_write.max(1)).min(remaining);
        self.bytes.extend_from_slice(&bytes[..amount]);
        if self.bytes.len() >= self.threshold && !self.cancelled {
            self.cancelled = true;
            self.cancellation.cancel();
        }
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

impl Write for CancelAfterShortWrite {
    fn write(&mut self, bytes: &[u8]) -> io::Result<usize> {
        let amount = bytes.len().min(self.max_write.max(1));
        self.bytes.extend_from_slice(&bytes[..amount]);
        if !self.cancelled {
            self.cancelled = true;
            self.cancellation.cancel();
        }
        Ok(amount)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}

#[test]
fn pre_cancel_and_short_scratch_write_report_cancellation_without_losing_prefix() {
    let value = "ordinary text with café🙂 and  spaces";
    let expected = scalar_reference(value);

    let (parent, child, cancellation, context) =
        context_with_parent_child(all_limits(), all_limits());
    cancellation.cancel();
    let mut pre_output = Vec::new();
    let pre = emit_with_context(&mut pre_output, value, usize::MAX, 0, usize::MAX, &context);
    assert!(matches!(
        pre.execution_error.as_ref(),
        Some(&ExecutionError::Cancelled)
    ));
    assert!(pre.result.is_err());
    assert!(pre_output.is_empty());
    assert_eq!(pre.bytes, 0);
    assert_eq!(parent.used(Resource::Work), 0);
    assert_eq!(child.used(Resource::Work), 0);

    let (parent, child, cancellation, context) =
        context_with_parent_child(all_limits(), all_limits());
    let mut sink = CancelAfterShortWrite {
        bytes: Vec::new(),
        max_write: 2,
        cancellation,
        cancelled: false,
    };
    let mid = emit_with_context(&mut sink, value, usize::MAX, 0, usize::MAX, &context);
    assert!(matches!(
        mid.execution_error.as_ref(),
        Some(&ExecutionError::Cancelled)
    ));
    assert!(mid.result.is_err());
    assert!(!sink.bytes.is_empty());
    assert!(sink.bytes.len() < expected.len());
    assert_eq!(sink.bytes, expected[..sink.bytes.len()]);
    assert_eq!(mid.bytes, sink.bytes.len());
    assert_eq!(parent.used(Resource::Work), sink.bytes.len() as u64);
    assert_eq!(child.used(Resource::Work), sink.bytes.len() as u64);

    // This deterministic sink cancels after the first completed 256-byte
    // ordinary span. The next span must not start. This checks the batch
    // boundary, without assuming a scalar-equivalent asynchronous timeline.
    let value = "a".repeat(300);
    let expected = scalar_reference(&value);
    let threshold = b"<text:p>".len() + 256;
    let (parent, child, cancellation, context) =
        context_with_parent_child(all_limits(), all_limits());
    let mut ordinary_sink = CancelAtAcceptedBytes {
        bytes: Vec::new(),
        threshold,
        max_write: 17,
        cancellation,
        cancelled: false,
    };
    let ordinary = emit_with_context(
        &mut ordinary_sink,
        &value,
        usize::MAX,
        0,
        usize::MAX,
        &context,
    );
    assert!(matches!(
        ordinary.execution_error.as_ref(),
        Some(&ExecutionError::Cancelled)
    ));
    assert!(ordinary.result.is_err());
    assert_eq!(ordinary_sink.bytes.len(), threshold);
    assert!(ordinary_sink.bytes.len() < expected.len());
    assert_eq!(
        ordinary_sink.bytes,
        expected[..ordinary_sink.bytes.len()],
        "ordinary-span cancellation must preserve a scalar output prefix"
    );
    assert_eq!(ordinary.bytes, ordinary_sink.bytes.len());
    assert_eq!(
        parent.used(Resource::Work),
        ordinary_sink.bytes.len() as u64
    );
    assert_eq!(child.used(Resource::Work), ordinary_sink.bytes.len() as u64);
}
