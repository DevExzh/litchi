//! Public regression checks for bounded scalar text spans.

#![allow(clippy::expect_used, clippy::unwrap_used)]

use std::{
    borrow::Cow,
    io::{self, Cursor, Read, Write},
    num::{NonZeroU64, NonZeroUsize},
};

use litchi_core::{
    Budget, CancellationSource, ExecutionContext, ExecutionError, ExecutionLimits, Limits, Resource,
};
use litchi_ods::streaming::{
    ScalarStreamReport, StreamingCell, StreamingError, StreamingLimits, stream_scalar_rows_to,
};
use zip::ZipArchive;

type Row = Vec<StreamingCell<'static>>;
type Rows = Vec<Row>;

fn all_limits() -> Limits {
    Limits::new(u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX, u64::MAX)
}

fn execution_limits() -> ExecutionLimits {
    ExecutionLimits::new(
        NonZeroUsize::new(1).expect("one worker"),
        NonZeroUsize::new(1).expect("one task"),
        NonZeroU64::new(u64::MAX).expect("finite byte ceiling"),
        0,
    )
    .expect("valid execution policy")
}

fn run_rows(
    rows: &Rows,
    limits: StreamingLimits,
) -> (
    Vec<u8>,
    Result<ScalarStreamReport, StreamingError>,
    Budget,
    Budget,
) {
    run_rows_with_budgets(rows, limits, all_limits(), all_limits())
}

fn run_rows_with_budgets(
    rows: &Rows,
    limits: StreamingLimits,
    parent_limits: Limits,
    child_limits: Limits,
) -> (
    Vec<u8>,
    Result<ScalarStreamReport, StreamingError>,
    Budget,
    Budget,
) {
    let parent = Budget::root("text-span-parent", parent_limits);
    let child = parent.child("text-span-child", child_limits);
    let (_cancellation, token) = CancellationSource::pair();
    let context = ExecutionContext::new(child.clone(), token, execution_limits());
    let mut output = Vec::new();
    let result = stream_scalar_rows_to(&mut output, rows.clone(), &context, limits);
    (output, result, parent, child)
}

fn member(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("successful output is ZIP");
    let mut entry = archive.by_name(name).expect("member exists");
    let mut value = Vec::new();
    entry.read_to_end(&mut value).expect("member reads");
    value
}

fn assert_unfinished_zip(bytes: &[u8]) {
    assert!(
        ZipArchive::new(Cursor::new(bytes)).is_err(),
        "a failed operation must not leave a finalized ZIP archive"
    );
}

fn reference_escape(value: &str) -> String {
    let mut escaped = String::new();
    for character in value.chars() {
        match character {
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            '\r' => escaped.push_str("&#13;"),
            '\n' => escaped.push_str("&#10;"),
            '\t' => escaped.push_str("&#9;"),
            character => escaped.push(character),
        }
    }
    escaped
}

fn reference_row(value: &str) -> String {
    let preserve = if value
        .bytes()
        .any(|byte| matches!(byte, b' ' | b'\t' | b'\n' | b'\r'))
    {
        " xml:space=\"preserve\""
    } else {
        ""
    };
    format!(
        "<table:table-row><table:table-cell office:value-type=\"string\"><text:p{preserve}>{}</text:p></table:table-cell></table:table-row>",
        reference_escape(value)
    )
}

fn authored_row_from_content(content: &[u8]) -> String {
    let content = String::from_utf8(content.to_vec()).expect("content is UTF-8");
    let start = content
        .find("<table:table-row>")
        .expect("content contains row");
    let end = start
        + content[start..]
            .find("</table:table-row>")
            .expect("content closes row")
        + "</table:table-row>".len();
    content[start..end].to_owned()
}

fn make_limits(base: StreamingLimits, max_row_xml_bytes: usize) -> StreamingLimits {
    StreamingLimits::new(
        base.max_rows(),
        base.max_cells(),
        base.max_cells_per_row(),
        base.max_text_bytes(),
        max_row_xml_bytes,
        base.max_content_xml_bytes(),
        base.max_output_bytes(),
        base.xml_audit(),
    )
    .expect("checked streaming limits")
}

fn core_limits_with(resource: Resource, limit: u64) -> Limits {
    let mut values = [u64::MAX; 6];
    let index = match resource {
        Resource::Memory => 0,
        Resource::InputBytes => 1,
        Resource::OutputBytes => 2,
        Resource::Objects => 3,
        Resource::Depth => 4,
        Resource::Work => 5,
        _ => unreachable!("the six core resources are exhaustive"),
    };
    values[index] = limit;
    Limits::new(
        values[0], values[1], values[2], values[3], values[4], values[5],
    )
}

fn assert_execution_resource(error: StreamingError, resource: Resource, output: &[u8]) {
    match error {
        StreamingError::Execution {
            written,
            error: ExecutionError::ResourceLimit(limit),
        } => {
            assert_eq!(limit.resource, resource);
            assert_eq!(written, output.len() as u64);
        },
        other => panic!("expected {resource:?} execution refusal, got {other:?}"),
    }
}

fn assert_local_limit(error: StreamingError, resource: &'static str, output: &[u8]) {
    match error {
        StreamingError::LimitExceeded {
            resource: actual,
            observed,
            limit,
            written,
        } => {
            assert_eq!(actual, resource);
            assert!(observed > limit);
            assert_eq!(written, output.len() as u64);
        },
        other => panic!("expected local {resource} refusal, got {other:?}"),
    }
}

fn assert_cancelled(error: StreamingError, output: &[u8]) {
    match error {
        StreamingError::Execution {
            written,
            error: ExecutionError::Cancelled,
        } => assert_eq!(written, output.len() as u64),
        other => panic!("expected cancellation, got {other:?}"),
    }
}

fn text_case_rows(value: String) -> Rows {
    vec![vec![StreamingCell::Text(Cow::Owned(value))]]
}

#[test]
fn text_output_matches_independent_reference_for_utf8_escapes_empty_and_span_boundaries() {
    let cases = [
        ("utf8", "café🙂漢字".to_owned()),
        ("eight-escapes", "&<>\"'\r\n\t".to_owned()),
        ("empty", String::new()),
        ("span-255", "a".repeat(255)),
        ("span-256", "a".repeat(256)),
        ("span-257", "a".repeat(257)),
    ];
    for (name, value) in cases {
        let rows = text_case_rows(value.clone());
        let (output, result, parent, child) = run_rows(&rows, StreamingLimits::default());
        let report = result.expect(name);
        assert_eq!(report.rows(), 1, "{name}");
        assert_eq!(report.cells(), 1, "{name}");
        assert_eq!(
            authored_row_from_content(&member(&output, "content.xml")),
            reference_row(&value),
            "{name} differs from the independent scalar reference"
        );
        assert_eq!(
            parent.used(Resource::Memory),
            0,
            "{name} parent memory leak"
        );
        assert_eq!(child.used(Resource::Memory), 0, "{name} child memory leak");
    }
}

#[test]
fn row_xml_window_exact_and_one_under_are_checked_for_255_256_257_spans() {
    let base = StreamingLimits::default();
    for length in [255usize, 256, 257] {
        let value = "x".repeat(length);
        let rows = text_case_rows(value.clone());
        let expected = reference_row(&value);
        let exact = make_limits(base, expected.len());
        let (output, result, parent, child) = run_rows(&rows, exact);
        result.expect("exact row window succeeds");
        assert_eq!(
            authored_row_from_content(&member(&output, "content.xml")),
            expected
        );
        assert_eq!(parent.used(Resource::Memory), 0);
        assert_eq!(child.used(Resource::Memory), 0);

        let (output, error, parent, child) = run_rows(&rows, make_limits(base, expected.len() - 1));
        assert_local_limit(
            error.expect_err("one-under row window refuses"),
            "row XML bytes",
            &output,
        );
        assert_unfinished_zip(&output);
        assert_eq!(parent.used(Resource::Memory), 0);
        assert_eq!(child.used(Resource::Memory), 0);
    }
}

#[test]
fn work_budget_exact_and_one_under_preserve_typed_progress_without_compressed_prefix_claims() {
    for length in [255usize, 256, 257] {
        check_work_case(&text_case_rows("x".repeat(length)));
    }
    check_work_case(&text_case_rows("work & < > \" '\r\n\t🙂".repeat(3)));
}

fn check_work_case(rows: &Rows) {
    let (control, result, control_parent, control_child) =
        run_rows(rows, StreamingLimits::default());
    result.expect("unbounded control succeeds");
    let work = control_child.used(Resource::Work);
    assert!(work > 0);
    assert_eq!(control_parent.used(Resource::Work), work);

    let exact_limits = core_limits_with(Resource::Work, work);
    let (output, result, parent, child) =
        run_rows_with_budgets(rows, StreamingLimits::default(), exact_limits, exact_limits);
    result.expect("exact child and parent work limits succeed");
    assert_eq!(output, control);
    assert_eq!(parent.used(Resource::Work), work);
    assert_eq!(child.used(Resource::Work), work);
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);

    let under_limits = core_limits_with(Resource::Work, work - 1);
    let (output, error, parent, child) =
        run_rows_with_budgets(rows, StreamingLimits::default(), all_limits(), under_limits);
    assert_execution_resource(
        error.expect_err("one-under child work budget refuses"),
        Resource::Work,
        &output,
    );
    assert!(child.used(Resource::Work) < work);
    assert_eq!(parent.used(Resource::Work), child.used(Resource::Work));
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);
    assert_unfinished_zip(&output);

    let (output, error, parent, child) =
        run_rows_with_budgets(rows, StreamingLimits::default(), under_limits, all_limits());
    assert_execution_resource(
        error.expect_err("one-under parent work budget refuses"),
        Resource::Work,
        &output,
    );
    assert!(parent.used(Resource::Work) < work);
    assert_eq!(parent.used(Resource::Work), child.used(Resource::Work));
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);
    assert_unfinished_zip(&output);
}

#[test]
fn required_memory_exact_and_one_under_propagates_through_parent_and_child() {
    let rows = text_case_rows("memory".to_owned());
    let required = StreamingLimits::default().required_memory_bytes();
    let exact_limits = core_limits_with(Resource::Memory, required);
    let (_output, result, parent, child) = run_rows_with_budgets(
        &rows,
        StreamingLimits::default(),
        exact_limits,
        exact_limits,
    );
    result.expect("exact memory reservation succeeds");
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);

    let under = core_limits_with(Resource::Memory, required - 1);
    let (output, error, parent, child) =
        run_rows_with_budgets(&rows, StreamingLimits::default(), all_limits(), under);
    assert_execution_resource(
        error.expect_err("one-under child memory reservation refuses"),
        Resource::Memory,
        &output,
    );
    assert!(output.is_empty());
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);

    let (output, error, parent, child) =
        run_rows_with_budgets(&rows, StreamingLimits::default(), under, all_limits());
    assert_execution_resource(
        error.expect_err("one-under parent memory reservation refuses"),
        Resource::Memory,
        &output,
    );
    assert!(output.is_empty());
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);
}

#[test]
fn cancellation_reports_accepted_progress_and_releases_memory_reservation() {
    let rows = text_case_rows("cancel🙂&<>".repeat(8));
    let (control, result, _parent, _child) = run_rows(&rows, StreamingLimits::default());
    result.expect("control succeeds");

    let parent = Budget::root("cancel-parent", all_limits());
    let child = parent.child("cancel-child", all_limits());
    let (source, token) = CancellationSource::pair();
    source.cancel();
    let context = ExecutionContext::new(child.clone(), token, execution_limits());
    let mut output = Vec::new();
    let error = stream_scalar_rows_to(
        &mut output,
        rows.clone(),
        &context,
        StreamingLimits::default(),
    )
    .expect_err("pre-cancelled stream refuses");
    assert_cancelled(error, &output);
    assert!(output.is_empty());
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);

    let parent = Budget::root("cancel-parent", all_limits());
    let child = parent.child("cancel-child", all_limits());
    let (source, token) = CancellationSource::pair();
    let context = ExecutionContext::new(child.clone(), token, execution_limits());
    let mut output = CancelAfterAccepted::new(source, 1);
    let error = stream_scalar_rows_to(
        &mut output,
        rows.clone(),
        &context,
        StreamingLimits::default(),
    )
    .expect_err("mid-stream cancellation refuses");
    assert_cancelled(error, output.bytes());
    assert!(!output.bytes().is_empty());
    assert!(output.bytes().len() < control.len());
    assert_unfinished_zip(output.bytes());
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);

    let parent = Budget::root("cancel-parent", all_limits());
    let child = parent.child("cancel-child", all_limits());
    let (source, token) = CancellationSource::pair();
    let context = ExecutionContext::new(child.clone(), token, execution_limits());
    let mut output = CancelAfterAccepted::new(source, control.len());
    let error = stream_scalar_rows_to(&mut output, rows, &context, StreamingLimits::default())
        .expect_err("final-write cancellation refuses after accepted bytes");
    assert_cancelled(error, output.bytes());
    assert_eq!(output.bytes(), control.as_slice());
    assert_eq!(parent.used(Resource::Memory), 0);
    assert_eq!(child.used(Resource::Memory), 0);
}

struct CancelAfterAccepted {
    bytes: Vec<u8>,
    source: CancellationSource,
    threshold: usize,
}

impl CancelAfterAccepted {
    fn new(source: CancellationSource, threshold: usize) -> Self {
        Self {
            bytes: Vec::new(),
            source,
            threshold,
        }
    }

    fn bytes(&self) -> &[u8] {
        &self.bytes
    }
}

impl Write for CancelAfterAccepted {
    fn write(&mut self, input: &[u8]) -> io::Result<usize> {
        if input.is_empty() {
            return Ok(0);
        }
        self.bytes.extend_from_slice(input);
        if self.bytes.len() >= self.threshold {
            self.source.cancel();
        }
        Ok(input.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
